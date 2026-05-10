"""CATIA DMU agent for automatic between-all analysis outputs."""

import argparse
import csv
import json
import math
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import sys

from PIL import Image, ImageDraw, ImageFilter

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import win32com.client
import win32com.client.dynamic
from pycatia.enumeration.enums import (
    CatClashComputationType,
    CatClashInterferenceType,
    CatConflictComparison,
    CatConflictStatus,
    CatConflictType,
    CatProjectionMode,
)
from pycatia.in_interfaces.viewpoint_3d import ViewPoint3D
from pycatia.navigator_interfaces.groups import Groups
from pycatia.product_structure_interfaces.product import Product
from pycatia.space_analyses_interfaces.clashes import Clashes
from pycatia.space_analyses_interfaces.conflict import Conflict

from catia_agents.user_parameter_agent import (
    as_dynamic_dispatch,
    collect_selectable_targets,
    get_document_product,
    connect_catia,
    safe_com_call,
    safe_com_get,
    safe_text,
)


DMU_WORKBENCH_CANDIDATES = (
    "DMUSpaceAnalysisWorkbench",
    "SPAWorkbench",
)

MODE_BETWEEN_ALL = "between_all_components"
MODE_SELECTION_AGAINST_ALL = "selection_against_all"
MODE_BETWEEN_TWO = "between_two_components"

SUPPORTED_MODES = (
    MODE_BETWEEN_ALL,
    MODE_SELECTION_AGAINST_ALL,
    MODE_BETWEEN_TWO,
)


@dataclass
class ConflictRecord:
    number: int
    product1: str
    product2: str
    conflict_type: str
    value: float
    status: str
    info: str
    keep: str
    image_path: str
    first_product: object
    second_product: object
    first_point: tuple | None
    second_point: tuple | None


@dataclass
class PreparedScope:
    mode: str
    first_group: object | None
    second_group: object | None
    selected_targets: list
    comparison_targets: list
    temporary_group_names: list
    all_targets: list


def parse_args():
    parser = argparse.ArgumentParser(description="Run CATIA DMU clash/contact/clearance analysis.")
    parser.add_argument(
        "--mode",
        default=None,
        choices=list(SUPPORTED_MODES),
        help="Analysis mode. If omitted, choose in PowerShell.",
    )
    parser.add_argument(
        "--clearance_mm",
        default=5.0,
        type=float,
        help="Clearance threshold in millimeters.",
    )
    parser.add_argument(
        "--output_dir",
        default="result/dmu_agent",
        type=str,
        help="Base output folder for the DMU run outputs.",
    )
    parser.add_argument(
        "--show_rows",
        default=12,
        type=int,
        help="How many rows to print in the PowerShell summary table.",
    )
    parser.add_argument(
        "--part_number",
        default=None,
        type=str,
        help="Part number for selection-against-all mode.",
    )
    parser.add_argument(
        "--part_number_a",
        default=None,
        type=str,
        help="First part number for between-two-components mode.",
    )
    parser.add_argument(
        "--part_number_b",
        default=None,
        type=str,
        help="Second part number for between-two-components mode.",
    )
    return parser.parse_args()


def choose_mode_interactive():
    print("")
    print("DMU analysis mode:")
    print("  1. Between all components")
    print("  2. Selection against all")
    print("  3. Between two components")
    while True:
        choice = input("Choose mode [1-3, default 1]: ").strip()
        if choice in ("", "1"):
            return MODE_BETWEEN_ALL
        if choice == "2":
            return MODE_SELECTION_AGAINST_ALL
        if choice == "3":
            return MODE_BETWEEN_TWO
        print("Please enter 1, 2, or 3.")


def wrap_dynamic(com_object):
    return win32com.client.dynamic.DumbDispatch(com_object)


def get_active_document(catia):
    document = safe_com_get(catia, "ActiveDocument")
    if document is None:
        raise RuntimeError("No active CATIA document found. Open the CATProduct first, then run the DMU agent.")
    return document


def get_document_summary(document):
    return {
        "name": safe_text(safe_com_get(document, "Name")) or "unknown",
        "type": safe_text(safe_com_get(document, "Type")) or "unknown",
        "path": safe_text(safe_com_get(document, "FullName")) or "",
    }


def start_dmu_workbench(catia):
    for workbench_name in DMU_WORKBENCH_CANDIDATES:
        safe_com_call(catia, "StartWorkbench", workbench_name)
        current = get_current_workbench_id(catia)
        if current:
            return workbench_name, current
    raise RuntimeError(
        "Could not open a DMU workbench. Tried: {}.".format(", ".join(DMU_WORKBENCH_CANDIDATES))
    )


def get_current_workbench_id(catia):
    workbench_id = safe_com_call(catia, "GetWorkbenchId")
    return safe_text(workbench_id)


def get_clashes_collection(root_product):
    clashes_obj = safe_com_call(root_product, "GetTechnologicalObject", "Clashes")
    if clashes_obj is None:
        raise RuntimeError('CATIA could not retrieve technological object "Clashes".')
    return Clashes(wrap_dynamic(clashes_obj))


def get_groups_collection(root_product):
    groups_obj = safe_com_call(root_product, "GetTechnologicalObject", "Groups")
    if groups_obj is None:
        raise RuntimeError('CATIA could not retrieve technological object "Groups".')
    return Groups(wrap_dynamic(groups_obj))


def get_root_tree_path(root_product, document_summary):
    return safe_text(safe_com_get(root_product, "Name")) or document_summary["name"] or "ROOT_PRODUCT"


def collect_part_targets(root_product, root_tree_path):
    targets = collect_selectable_targets(root_product, root_tree_path, depth=0, include_self=True)
    return [target for target in targets if target.node_kind == "part"]


def normalize_part_number(text):
    return str(text or "").strip().lower()


def choose_target_from_matches(label, matches):
    if len(matches) == 1:
        return matches[0]

    print("")
    print("{} matches:".format(label))
    for index, target in enumerate(matches, start=1):
        print(
            "  {}. {} | instance={} | tree_path={}".format(
                index,
                target.display_name or target.part_number or target.tree_path,
                target.instance_name or "unknown",
                target.tree_path,
            )
        )
    while True:
        raw = input("Choose {} instance [1-{}]: ".format(label.lower(), len(matches))).strip()
        try:
            selected_index = int(raw)
        except Exception:
            selected_index = None
        if selected_index is not None and 1 <= selected_index <= len(matches):
            return matches[selected_index - 1]
        print("Please enter a valid instance number.")


def find_target_by_part_number(targets, part_number, label, exclude_tree_paths=None):
    normalized = normalize_part_number(part_number)
    excluded = set(exclude_tree_paths or [])
    matches = [
        target
        for target in targets
        if normalize_part_number(target.part_number) == normalized and target.tree_path not in excluded
    ]
    if not matches:
        raise RuntimeError("No part instance found for {} part number '{}'.".format(label, part_number))
    return choose_target_from_matches(label, matches)


def prompt_part_number(prompt_text):
    while True:
        value = input(prompt_text).strip()
        if value:
            return value
        print("Please enter a part number.")


def create_group_from_targets(groups_collection, targets, group_name):
    group = groups_collection.add()
    group.name = group_name
    group.extract_mode = 1
    for target in targets:
        group.add_explicit(Product(as_dynamic_dispatch(target.product)))
    return group


def build_scope_from_mode(root_product, document_summary, mode, part_number=None, part_number_a=None, part_number_b=None):
    root_tree_path = get_root_tree_path(root_product, document_summary)
    part_targets = collect_part_targets(root_product, root_tree_path)
    temporary_group_names = []

    if mode == MODE_BETWEEN_ALL:
        return PreparedScope(
            mode=mode,
            first_group=None,
            second_group=None,
            selected_targets=[],
            comparison_targets=part_targets,
            temporary_group_names=temporary_group_names,
            all_targets=part_targets,
        )

    groups_collection = get_groups_collection(root_product)

    if mode == MODE_SELECTION_AGAINST_ALL:
        selected_part_number = part_number or prompt_part_number("Enter part number for selection against all: ")
        selected_target = find_target_by_part_number(part_targets, selected_part_number, "Selection")
        comparison_targets = [target for target in part_targets if target.tree_path != selected_target.tree_path]
        if not comparison_targets:
            raise RuntimeError("No comparison part targets remain after excluding the selected component.")
        first_group = create_group_from_targets(groups_collection, [selected_target], "DMU_SELECTED")
        second_group = create_group_from_targets(groups_collection, comparison_targets, "DMU_ALL_OTHERS")
        temporary_group_names.extend([first_group.name, second_group.name])
        return PreparedScope(
            mode=mode,
            first_group=first_group,
            second_group=second_group,
            selected_targets=[selected_target],
            comparison_targets=comparison_targets,
            temporary_group_names=temporary_group_names,
            all_targets=part_targets,
        )

    if mode == MODE_BETWEEN_TWO:
        selected_part_number_a = part_number_a or prompt_part_number("Enter first part number: ")
        selected_target_a = find_target_by_part_number(part_targets, selected_part_number_a, "Selection A")
        selected_part_number_b = part_number_b or prompt_part_number("Enter second part number: ")
        selected_target_b = find_target_by_part_number(
            part_targets,
            selected_part_number_b,
            "Selection B",
            exclude_tree_paths={selected_target_a.tree_path},
        )
        first_group = create_group_from_targets(groups_collection, [selected_target_a], "DMU_SELECTION_A")
        second_group = create_group_from_targets(groups_collection, [selected_target_b], "DMU_SELECTION_B")
        temporary_group_names.extend([first_group.name, second_group.name])
        return PreparedScope(
            mode=mode,
            first_group=first_group,
            second_group=second_group,
            selected_targets=[selected_target_a, selected_target_b],
            comparison_targets=[],
            temporary_group_names=temporary_group_names,
            all_targets=part_targets,
        )

    raise RuntimeError("Unsupported mode: {}".format(mode))


def resolve_computation_type(mode):
    if mode == MODE_BETWEEN_ALL:
        return int(CatClashComputationType.catClashComputationTypeBetweenAll)
    return int(CatClashComputationType.catClashComputationTypeBetweenTwo)


def run_contact_plus_clash(clashes_collection, mode, first_group=None, second_group=None):
    clash = clashes_collection.add()
    clash.name = "DMU_CONTACT_PLUS_CLASH"
    clash.computation_type = resolve_computation_type(mode)
    clash.interference_type = int(CatClashInterferenceType.catClashInterferenceTypeContact)
    if first_group is not None:
        clash.first_group = first_group
    if second_group is not None:
        clash.second_group = second_group
    clash.compute()
    return clash


def run_clearance(clashes_collection, mode, clearance_mm, first_group=None, second_group=None):
    clash = clashes_collection.add()
    clash.name = "DMU_CLEARANCE"
    clash.computation_type = resolve_computation_type(mode)
    clash.interference_type = int(CatClashInterferenceType.catClashInterferenceTypeClearance)
    clash.clearance = float(clearance_mm)
    if first_group is not None:
        clash.first_group = first_group
    if second_group is not None:
        clash.second_group = second_group
    clash.compute()
    return clash


def map_conflict_type(conflict_type):
    mapping = {
        int(CatConflictType.catConflictTypeClash): "clash",
        int(CatConflictType.catConflictTypeContact): "contact",
        int(CatConflictType.catConflictTypeClearance): "clearance",
    }
    return mapping.get(int(conflict_type), "unknown")


def map_conflict_status(conflict_status):
    try:
        normalized = int(conflict_status)
    except Exception:
        return "Unknown"
    mapping = {
        int(CatConflictStatus.catConflictStatusNotInspected): "Not inspected",
        int(CatConflictStatus.catConflictStatusRelevant): "Relevant",
        int(CatConflictStatus.catConflictStatusIrrelevant): "Irrelevant",
        int(CatConflictStatus.catConflictStatusSolved): "Solved",
    }
    return mapping.get(normalized, "Unknown")


def map_conflict_comparison_info(comparison_info):
    try:
        normalized = int(comparison_info)
    except Exception:
        return "Unknown"
    mapping = {
        int(CatConflictComparison.catConflictComparisonNew): "New",
        int(CatConflictComparison.catConflictComparisonOld): "Old",
        int(CatConflictComparison.catConflictComparisonNo): "No",
    }
    return mapping.get(normalized, "Unknown")


def resolve_instance_name(product_obj):
    if product_obj is None:
        return None
    instance_name = safe_text(safe_com_get(product_obj, "name")) or safe_text(safe_com_get(product_obj, "Name"))
    part_number = safe_text(safe_com_get(product_obj, "part_number")) or safe_text(safe_com_get(product_obj, "PartNumber"))
    definition = safe_text(safe_com_get(product_obj, "definition")) or safe_text(safe_com_get(product_obj, "Definition"))
    if part_number and definition:
        return "{} ({})".format(part_number, definition)
    return part_number or instance_name or definition or "<unknown>"


def slugify(text):
    value = str(text or "item").strip()
    safe = []
    for char in value:
        if char.isalnum() or char in ("-", "_"):
            safe.append(char)
        else:
            safe.append("_")
    return "".join(safe).strip("_") or "item"


def create_output_paths(base_output_dir, document_name):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_name = "{}__{}".format(slugify(Path(document_name).stem if document_name else "catia_document"), timestamp)
    run_root = Path(base_output_dir).resolve() / run_name
    images_dir = run_root / "images"
    images_dir.mkdir(parents=True, exist_ok=True)
    return {
        "root": run_root,
        "images": images_dir,
        "json": run_root / "results.json",
        "csv": run_root / "results.csv",
        "txt": run_root / "results.txt",
    }


def build_info_text(conflict):
    comparison_value = safe_conflict_comparison_info(conflict)
    pieces = [map_conflict_comparison_info(comparison_value)]
    comment = safe_conflict_comment(conflict)
    if comment:
        pieces.append(comment)
    return " | ".join(piece for piece in pieces if piece)


def get_conflict_com_object(conflict):
    return safe_com_get(conflict, "conflict", conflict)


def safe_conflict_type(conflict):
    try:
        return int(conflict.type)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        return int(safe_com_get(raw_conflict, "Type", 0) or 0)


def safe_conflict_status(conflict):
    try:
        return int(conflict.status)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        return safe_com_get(raw_conflict, "Status", None)


def safe_conflict_comparison_info(conflict):
    try:
        return int(conflict.comparison_info)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        return safe_com_get(raw_conflict, "ComparisonInfo", None)


def safe_conflict_comment(conflict):
    try:
        return safe_text(conflict.comment)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        return safe_text(safe_com_get(raw_conflict, "Comment"))


def safe_conflict_value(conflict):
    try:
        return float(conflict.value)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        value = safe_com_get(raw_conflict, "Value", 0.0)
        try:
            return float(value)
        except Exception:
            return 0.0


def safe_conflict_product(conflict, side):
    property_name = "first_product" if side == "first" else "second_product"
    com_name = "FirstProduct" if side == "first" else "SecondProduct"

    try:
        return getattr(conflict, property_name)
    except Exception:
        raw_conflict = get_conflict_com_object(conflict)
        raw_product = safe_com_get(raw_conflict, com_name)
        if raw_product is None:
            return None
        try:
            return Product(raw_product)
        except Exception:
            return raw_product


def safe_conflict_point(conflict, side):
    try:
        wrapped_conflict = Conflict(get_conflict_com_object(conflict))
        if side == "first":
            point = wrapped_conflict.get_first_point_coordinates()
        else:
            point = wrapped_conflict.get_second_point_coordinates()
    except Exception:
        return None

    try:
        values = tuple(float(value) for value in point[:3])
    except Exception:
        return None
    if len(values) != 3:
        return None
    return values


def collect_conflict_records(contact_clash, clearance_clash, images_dir):
    records = []
    number = 1
    sources = (
        (contact_clash, None),
        (clearance_clash, "clearance"),
    )
    for clash_object, required_type in sources:
        conflicts = clash_object.conflicts
        for index in range(1, conflicts.count + 1):
            conflict = conflicts.item(index)
            conflict_type = map_conflict_type(safe_conflict_type(conflict))
            if required_type is not None and conflict_type != required_type:
                continue
            first_product = safe_conflict_product(conflict, "first")
            second_product = safe_conflict_product(conflict, "second")
            first_point = safe_conflict_point(conflict, "first")
            second_point = safe_conflict_point(conflict, "second")
            image_name = "{:03d}_{}_{}_{}.png".format(
                number,
                conflict_type,
                slugify(resolve_instance_name(first_product)),
                slugify(resolve_instance_name(second_product)),
            )
            records.append(
                ConflictRecord(
                    number=number,
                    product1=resolve_instance_name(first_product),
                    product2=resolve_instance_name(second_product),
                    conflict_type=conflict_type,
                    value=safe_conflict_value(conflict),
                    status=map_conflict_status(safe_conflict_status(conflict)),
                    info=build_info_text(conflict),
                    keep="",
                    image_path=str((images_dir / image_name).resolve()),
                    first_product=first_product,
                    second_product=second_product,
                    first_point=first_point,
                    second_point=second_point,
                )
            )
            number += 1
    return records


def select_products_for_capture(document, first_product, second_product):
    selection = safe_com_get(document, "Selection")
    if selection is None:
        raise RuntimeError("CATIA selection API is not available on the active document.")
    safe_com_call(selection, "Clear")
    if first_product is not None:
        safe_com_call(selection, "Add", safe_com_get(first_product, "com_object", first_product))
    if second_product is not None:
        safe_com_call(selection, "Add", safe_com_get(second_product, "com_object", second_product))
    return selection


def set_products_visibility(document, products, show_value):
    selection = safe_com_get(document, "Selection")
    if selection is None:
        raise RuntimeError("CATIA selection API is not available on the active document.")
    if not products:
        return
    safe_com_call(selection, "Clear")
    try:
        for product in products:
            if product is None:
                continue
            safe_com_call(selection, "Add", safe_com_get(product, "com_object", product))
        vis_properties = safe_com_get(selection, "VisProperties")
        if vis_properties is not None:
            safe_com_call(vis_properties, "SetShow", show_value)
    finally:
        safe_com_call(selection, "Clear")


def set_targets_visibility(document, targets, show_value):
    products = [safe_com_get(target, "product", None) for target in targets]
    set_products_visibility(document, products, show_value)


def vector_subtract(left, right):
    return tuple(left[index] - right[index] for index in range(3))


def vector_add(left, right):
    return tuple(left[index] + right[index] for index in range(3))


def vector_scale(vector, scalar):
    return tuple(component * scalar for component in vector)


def vector_dot(left, right):
    return sum(left[index] * right[index] for index in range(3))


def vector_cross(left, right):
    return (
        left[1] * right[2] - left[2] * right[1],
        left[2] * right[0] - left[0] * right[2],
        left[0] * right[1] - left[1] * right[0],
    )


def vector_length(vector):
    return math.sqrt(vector_dot(vector, vector))


def normalize_vector(vector):
    length = vector_length(vector)
    if length <= 1e-9:
        return None
    return tuple(component / length for component in vector)


def midpoint_from_record(record):
    if record.first_point and record.second_point:
        return tuple((record.first_point[index] + record.second_point[index]) / 2.0 for index in range(3))
    return record.first_point or record.second_point


def get_viewpoint(viewer):
    viewpoint = safe_com_get(viewer, "Viewpoint3D")
    if viewpoint is None:
        return None
    return ViewPoint3D(viewpoint)


def capture_viewpoint_state(viewer):
    viewpoint = get_viewpoint(viewer)
    if viewpoint is None:
        return None
    return {
        "origin": tuple(viewpoint.get_origin()),
        "sight": tuple(viewpoint.get_sight_direction()),
        "up": tuple(viewpoint.get_up_direction()),
        "focus_distance": float(viewpoint.focus_distance),
        "projection_mode": int(viewpoint.projection_mode),
        "field_of_view": float(safe_com_get(viewpoint, "field_of_view", 0.0) or 0.0),
        "zoom": float(safe_com_get(viewpoint, "zoom", 1.0) or 1.0),
    }


def restore_viewpoint_state(viewer, state):
    if not state:
        return
    viewpoint = get_viewpoint(viewer)
    if viewpoint is None:
        return
    viewpoint.put_origin(state["origin"])
    viewpoint.put_sight_direction(state["sight"])
    viewpoint.put_up_direction(state["up"])
    viewpoint.focus_distance = state["focus_distance"]
    if state["projection_mode"] == int(CatProjectionMode.catProjectionConic):
        viewpoint.field_of_view = state["field_of_view"]
    elif state["projection_mode"] == int(CatProjectionMode.catProjectionCylindric):
        viewpoint.zoom = state["zoom"]


def focus_viewpoint_on_record(viewer, record):
    midpoint = midpoint_from_record(record)
    if midpoint is None:
        return False

    viewpoint = get_viewpoint(viewer)
    if viewpoint is None:
        return False

    sight = normalize_vector(tuple(viewpoint.get_sight_direction()))
    if sight is None:
        return False

    focus_distance = float(viewpoint.focus_distance)
    new_origin = vector_subtract(midpoint, vector_scale(sight, focus_distance))
    viewpoint.put_origin(new_origin)

    projection_mode = int(viewpoint.projection_mode)
    if projection_mode == int(CatProjectionMode.catProjectionConic):
        current_fov = float(viewpoint.field_of_view)
        viewpoint.field_of_view = max(2.0, current_fov * 0.55)
    elif projection_mode == int(CatProjectionMode.catProjectionCylindric):
        current_zoom = float(viewpoint.zoom)
        viewpoint.zoom = max(current_zoom * 4.0, current_zoom + 0.002)
    return True


def rotate_vector_around_axis(vector, axis, angle_degrees):
    normalized_axis = normalize_vector(axis)
    if normalized_axis is None:
        return vector
    radians = math.radians(angle_degrees)
    cos_theta = math.cos(radians)
    sin_theta = math.sin(radians)
    axis_dot_vector = vector_dot(normalized_axis, vector)
    first = vector_scale(vector, cos_theta)
    second = vector_scale(vector_cross(normalized_axis, vector), sin_theta)
    third = vector_scale(normalized_axis, axis_dot_vector * (1.0 - cos_theta))
    return vector_add(vector_add(first, second), third)


def apply_view_variant(viewer, base_state, record, yaw_degrees=0.0, pitch_degrees=0.0):
    midpoint = midpoint_from_record(record)
    if midpoint is None:
        restore_viewpoint_state(viewer, base_state)
        return False

    viewpoint = get_viewpoint(viewer)
    if viewpoint is None:
        return False

    sight = normalize_vector(base_state["sight"])
    up = normalize_vector(base_state["up"])
    if sight is None or up is None:
        restore_viewpoint_state(viewer, base_state)
        return False

    right = normalize_vector(vector_cross(sight, up))
    if right is None:
        restore_viewpoint_state(viewer, base_state)
        return False

    rotated_sight = sight
    rotated_up = up

    if yaw_degrees:
        rotated_sight = rotate_vector_around_axis(rotated_sight, rotated_up, yaw_degrees)
        right = normalize_vector(vector_cross(rotated_sight, rotated_up)) or right
    if pitch_degrees:
        rotated_sight = rotate_vector_around_axis(rotated_sight, right, pitch_degrees)
        rotated_up = rotate_vector_around_axis(rotated_up, right, pitch_degrees)

    rotated_sight = normalize_vector(rotated_sight) or sight
    rotated_up = normalize_vector(rotated_up) or up
    new_origin = vector_subtract(midpoint, vector_scale(rotated_sight, base_state["focus_distance"]))

    viewpoint.put_sight_direction(rotated_sight)
    viewpoint.put_up_direction(rotated_up)
    viewpoint.put_origin(new_origin)
    viewpoint.focus_distance = base_state["focus_distance"]

    if base_state["projection_mode"] == int(CatProjectionMode.catProjectionConic):
        viewpoint.field_of_view = base_state["field_of_view"]
    elif base_state["projection_mode"] == int(CatProjectionMode.catProjectionCylindric):
        viewpoint.zoom = base_state["zoom"]
    return True


def annotate_single_preview_image(image_path, record, view_label):
    image_file = Path(image_path)
    if not image_file.exists():
        return

    with Image.open(image_file) as source_image:
        canvas = source_image.convert("RGBA").filter(ImageFilter.UnsharpMask(radius=1.8, percent=140, threshold=2))

    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    center_x = canvas.width // 2
    center_y = canvas.height // 2

    draw.ellipse((center_x - 10, center_y - 10, center_x + 10, center_y + 10), outline=(255, 214, 102, 255), width=4)
    draw.line((center_x + 12, center_y - 12, center_x + 90, center_y - 54), fill=(255, 214, 102, 255), width=3)

    label_box = (center_x + 88, center_y - 88, center_x + 330, center_y - 24)
    draw.rounded_rectangle(label_box, radius=10, fill=(18, 24, 36, 220), outline=(255, 214, 102, 255), width=2)
    draw.text((label_box[0] + 12, label_box[1] + 10), "{} {:.3f} mm".format(str(record.conflict_type).upper(), record.value), fill=(255, 245, 210, 255))

    tag_box = (18, 18, 160, 56)
    draw.rounded_rectangle(tag_box, radius=10, fill=(18, 24, 36, 210), outline=(110, 160, 220, 255), width=2)
    draw.text((tag_box[0] + 12, tag_box[1] + 10), view_label, fill=(235, 245, 255, 255))

    composed = Image.alpha_composite(canvas, overlay).convert("RGBA")
    composed.save(image_file, format="PNG", optimize=True)


def compose_preview_montage(image_paths, output_path, record):
    opened = []
    try:
        for image_path in image_paths:
            with Image.open(image_path) as image:
                opened.append(image.convert("RGBA"))
        if not opened:
            return

        tile_width = min(image.width for image in opened)
        tile_height = min(image.height for image in opened)
        tiles = []
        for image in opened:
            tile = image.copy()
            if tile.size != (tile_width, tile_height):
                tile = tile.resize((tile_width, tile_height), Image.Resampling.LANCZOS)
            tiles.append(tile)

        footer_height = 58
        canvas = Image.new("RGBA", (tile_width * 2, tile_height * 2 + footer_height), (14, 18, 26, 255))
        positions = [
            (0, 0),
            (tile_width, 0),
            (0, tile_height),
            (tile_width, tile_height),
        ]
        for tile, position in zip(tiles, positions):
            canvas.paste(tile, position)

        draw = ImageDraw.Draw(canvas)
        draw.line((tile_width, 0, tile_width, tile_height * 2), fill=(70, 78, 96), width=3)
        draw.line((0, tile_height, tile_width * 2, tile_height), fill=(70, 78, 96), width=3)
        footer_top = tile_height * 2
        draw.rectangle((0, footer_top, canvas.width, canvas.height), fill=(16, 20, 28))
        footer_text = "{}  <->  {}    |    {} {:.3f} mm".format(
            record.product1,
            record.product2,
            str(record.conflict_type).upper(),
            record.value,
        )
        draw.text((16, footer_top + 16), footer_text, fill=(235, 235, 235))
        canvas = canvas.filter(ImageFilter.UnsharpMask(radius=1.4, percent=125, threshold=2))
        canvas.save(output_path, format="PNG", optimize=True)
    finally:
        for image in opened:
            image.close()


def capture_conflict_images(catia, document, records, all_targets):
    active_window = safe_com_get(catia, "ActiveWindow")
    viewer = safe_com_get(active_window, "ActiveViewer") if active_window is not None else None
    if viewer is None:
        raise RuntimeError("CATIA active viewer is not available for image capture.")

    selection = safe_com_get(document, "Selection")
    capture_format = win32com.client.constants.catCaptureFormatBMP
    show_attr = win32com.client.constants.catVisPropertyShowAttr
    no_show_attr = win32com.client.constants.catVisPropertyNoShowAttr
    original_view_state = capture_viewpoint_state(viewer)

    try:
        set_targets_visibility(document, all_targets, no_show_attr)
        for record in records:
            set_products_visibility(document, [record.first_product, record.second_product], show_attr)
            select_products_for_capture(document, record.first_product, record.second_product)
            safe_com_call(viewer, "Reframe")
            focus_viewpoint_on_record(viewer, record)
            safe_com_call(viewer, "Update")
            base_record_view_state = capture_viewpoint_state(viewer) or original_view_state
            variant_specs = [
                ("Front", 0.0, 0.0),
                ("Left", 38.0, 0.0),
                ("Right", -38.0, 0.0),
                ("Back", 145.0, 8.0),
            ]
            temp_paths = []
            try:
                for label, yaw_degrees, pitch_degrees in variant_specs:
                    apply_view_variant(viewer, base_record_view_state, record, yaw_degrees=yaw_degrees, pitch_degrees=pitch_degrees)
                    safe_com_call(viewer, "Update")
                    temp_file = Path(tempfile.NamedTemporaryFile(prefix="dmu_view_", suffix=".bmp", delete=False).name)
                    temp_paths.append(temp_file)
                    viewer.CaptureToFile(capture_format, str(temp_file))
                    annotate_single_preview_image(temp_file, record, label)
                compose_preview_montage(temp_paths, record.image_path, record)
            finally:
                restore_viewpoint_state(viewer, base_record_view_state)
                for temp_file in temp_paths:
                    try:
                        temp_file.unlink(missing_ok=True)
                    except Exception:
                        pass
            set_products_visibility(document, [record.first_product, record.second_product], no_show_attr)
    finally:
        restore_viewpoint_state(viewer, original_view_state)
        set_targets_visibility(document, all_targets, show_attr)
        if selection is not None:
            safe_com_call(selection, "Clear")


def write_results_json(json_path, document_summary, clearance_mm, records):
    payload = {
        "document_name": document_summary["name"],
        "document_path": document_summary["path"],
        "analysis_mode": document_summary.get("analysis_mode"),
        "clearance_mm": clearance_mm,
        "results": [
            {
                "No": record.number,
                "Product1": record.product1,
                "Product2": record.product2,
                "Type": record.conflict_type,
                "Value": record.value,
                "Status": record.status,
                "Info": record.info,
                "Keep": record.keep,
                "Image": record.image_path,
            }
            for record in records
        ],
    }
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def write_results_csv(csv_path, records):
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["No", "Product1", "Product2", "Type", "Value", "Status", "Info", "Keep", "Image"])
        for record in records:
            writer.writerow(
                [
                    record.number,
                    record.product1,
                    record.product2,
                    record.conflict_type,
                    record.value,
                    record.status,
                    record.info,
                    record.keep,
                    record.image_path,
                ]
            )


def build_table_rows(records):
    headers = ["No", "Product1", "Product2", "Type", "Value", "Status", "Info", "Keep", "Image"]
    rows = []
    for record in records:
        rows.append(
            [
                str(record.number),
                record.product1,
                record.product2,
                record.conflict_type,
                "{:.3f}".format(record.value),
                record.status,
                record.info,
                record.keep,
                record.image_path,
            ]
        )
    return headers, rows


def write_results_txt(txt_path, records):
    headers, rows = build_table_rows(records)
    widths = [len(header) for header in headers]
    for row in rows:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], len(str(value)))

    def format_row(values):
        return " | ".join(str(value).ljust(widths[index]) for index, value in enumerate(values))

    lines = [
        format_row(headers),
        "-+-".join("-" * width for width in widths),
    ]
    for row in rows:
        lines.append(format_row(row))
    txt_path.write_text("\n".join(lines), encoding="utf-8")


def print_results_table(records, show_rows):
    rows = records[:show_rows]
    headers = ["No", "Product1", "Product2", "Type", "Value", "Status", "Info", "Keep"]
    display_rows = []
    for record in rows:
        display_rows.append(
            [
                str(record.number),
                record.product1,
                record.product2,
                record.conflict_type,
                "{:.3f}".format(record.value),
                record.status,
                record.info,
                record.keep,
            ]
        )

    widths = [len(header) for header in headers]
    for row in display_rows:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], len(str(value)))

    def format_row(values):
        return " | ".join(str(value).ljust(widths[index]) for index, value in enumerate(values))

    print(format_row(headers))
    print("-+-".join("-" * width for width in widths))
    for row in display_rows:
        print(format_row(row))


def cleanup_temporary_clashes(root_product, temporary_clash_names):
    try:
        clashes_collection = get_clashes_collection(root_product)
        for clash_name in temporary_clash_names:
            safe_com_call(clashes_collection.clashes, "Remove", clash_name)
    except Exception:
        pass


def cleanup_temporary_groups(root_product, temporary_group_names):
    try:
        groups_collection = get_groups_collection(root_product)
        for group_name in temporary_group_names:
            safe_com_call(groups_collection.groups, "Remove", group_name)
    except Exception:
        pass


def main():
    args = parse_args()
    mode = args.mode or choose_mode_interactive()
    catia = connect_catia()
    document = get_active_document(catia)
    document_summary = get_document_summary(document)
    document_summary["analysis_mode"] = mode
    requested_workbench, current_workbench = start_dmu_workbench(catia)
    root_product = get_document_product(document)
    scope = build_scope_from_mode(
        root_product,
        document_summary,
        mode,
        part_number=args.part_number,
        part_number_a=args.part_number_a,
        part_number_b=args.part_number_b,
    )
    output_paths = create_output_paths(args.output_dir, document_summary["name"])
    clashes_collection = get_clashes_collection(root_product)
    temporary_clash_names = []

    try:
        contact_clash = run_contact_plus_clash(
            clashes_collection,
            mode,
            first_group=scope.first_group,
            second_group=scope.second_group,
        )
        temporary_clash_names.append(contact_clash.name)

        clearance_clash = run_clearance(
            clashes_collection,
            mode,
            args.clearance_mm,
            first_group=scope.first_group,
            second_group=scope.second_group,
        )
        temporary_clash_names.append(clearance_clash.name)

        records = collect_conflict_records(contact_clash, clearance_clash, output_paths["images"])
        capture_conflict_images(catia, document, records, scope.all_targets)
        write_results_json(output_paths["json"], document_summary, args.clearance_mm, records)
        write_results_csv(output_paths["csv"], records)
        write_results_txt(output_paths["txt"], records)

        print("CATIA DMU between-all analysis complete.")
        print("Active document: {}".format(document_summary["name"]))
        print("Document path: {}".format(document_summary["path"] or "<unknown>"))
        print("Requested workbench: {}".format(requested_workbench))
        print("Current workbench: {}".format(current_workbench or "unknown"))
        print("Mode: {}".format(mode))
        print("Clearance mm: {}".format(args.clearance_mm))
        if scope.selected_targets:
            print(
                "Selected targets: {}".format(
                    ", ".join(target.display_name or target.part_number or target.tree_path for target in scope.selected_targets)
                )
            )
        print("Total result rows: {}".format(len(records)))
        print("JSON output: {}".format(output_paths["json"]))
        print("CSV output: {}".format(output_paths["csv"]))
        print("TXT output: {}".format(output_paths["txt"]))
        print("Images folder: {}".format(output_paths["images"]))
        print("")
        print_results_table(records, args.show_rows)
    finally:
        cleanup_temporary_clashes(root_product, temporary_clash_names)
        cleanup_temporary_groups(root_product, scope.temporary_group_names)


if __name__ == "__main__":
    main()
