# CATIA DMU Preview Tool

This folder contains a small setup guide for the CATIA DMU preview workflow built in this repo.

## What This Tool Does

The DMU tool:

- connects to a running CATIA session
- uses the active `CATProduct` you already opened in CATIA
- runs DMU interference checks
- supports:
  - `between_all_components`
  - `selection_against_all`
  - `between_two_components`
- exports:
  - `results.json`
  - `results.csv`
  - `results.txt`
  - preview images for each result row
- opens a preview UI where you can:
  - browse result rows
  - filter by `contact`, `clash`, `clearance`
  - zoom with mouse wheel
  - pan by dragging

## Main Files

The main files for this DMU workflow are:

- [catia_agents/DMU Agent.py](../catia_agents/DMU%20Agent.py)
- [catia_agents/dmu_ui.py](../catia_agents/dmu_ui.py)

## Dependencies

Python dependencies needed only for this DMU tool:

- `pywin32`
- `pycatia`
- `Pillow`

System dependency:

- `CATIA V5` installed on Windows

CATIA requirement:

- CATIA must be running
- a valid `CATProduct` must already be open
- DMU / Space Analysis functionality must be available in the CATIA installation/license

## Python Version

Recommended:

- Python `3.10`

## Virtual Environment

Create and activate a clean virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install the DMU-only dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install -r github\requirements-dmu.txt
```

If you want to install manually:

```powershell
.\.venv\Scripts\python.exe -m pip install pywin32 pycatia Pillow
```

## CATIA Input Expectations

This tool expects:

- CATIA is already open
- the target assembly is open as the active document
- the active document is a `CATProduct`

Examples of supported runs:

- run against all components in an assembly
- run one selected part number against all other components
- run one typed part number against another typed part number

## How To Run

### Single-file Flow

This is the recommended entry point:

```powershell
.\.venv\Scripts\python.exe "catia_agents\dmu_ui.py"
```

What happens:

1. PowerShell asks startup questions
2. DMU analysis runs
3. Preview UI opens automatically with the latest results

### Viewer Only

If you only want to open the latest preview UI without rerunning analysis:

```powershell
.\.venv\Scripts\python.exe "catia_agents\dmu_ui.py" --viewer-only
```

### Agent Only

If you want only the terminal/analysis export flow:

```powershell
.\.venv\Scripts\python.exe "catia_agents\DMU Agent.py"
```

## Output Files

Each run creates a folder under:

```text
result/dmu_agent/
```

Each run folder contains:

- `results.json`
- `results.csv`
- `results.txt`
- `images/`

## Current Preview Behavior

Preview images currently:

- isolate the relevant result pair
- use multi-angle preview sheets
- include the conflict type and value
- support zoom and pan in the UI

Important note:

- the preview is generated from CATIA viewer captures
- it is not the exact CATIA native `Preview` popup export

## Known Requirements / Limits

- must run on Windows
- must have CATIA available through COM automation
- preview quality can vary depending on assembly complexity and visibility
- duplicate part numbers may require instance choice during selection-based runs
- large products can take longer to compute

## Troubleshooting

### CATIA not found

If the tool cannot connect to CATIA:

- make sure CATIA is running
- make sure you are on Windows
- make sure `pywin32` installed correctly in the active `.venv`

### No active product

If the tool says no active document:

- open the target `CATProduct` in CATIA
- make sure that product window is the active CATIA document

### Part number not found

If selection mode fails:

- confirm the typed part number exists in the assembly
- try the exact CATIA part number shown in the product tree

### Preview image not clear enough

Current improvements already included:

- multi-angle preview image
- higher-quality PNG output
- zoom + pan in the preview UI

## Minimal Setup Summary

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
.\.venv\Scripts\python.exe -m pip install -r github\requirements-dmu.txt
.\.venv\Scripts\python.exe "catia_agents\dmu_ui.py"
```
