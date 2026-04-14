---

# Loom Response Workflow

MATLAB scripts for automated loom stimulus detection, video clip generation,
and escape response analysis in mouse behavioural experiments. Developed as
part of an MSci thesis investigating time-of-day differences in visually
evoked defensive responses (University of Manchester, 2025–2026).

---

## Overview

This repository provides a reproducible pipeline for processing looming assay
data following the Yilmaz & Meister (2013) paradigm. Behavioural recordings
were acquired using a novel below-arena camera configuration, in which a small
indicator light visible in the camera frame provides a frame-accurate signal
of loom onset. This allows each loom event to be automatically detected and
clipped from the raw video without manual timestamp entry.

<img src="https://github.com/user-attachments/assets/4b15870c-ff73-4049-b4f2-8b821ac900cb" width="450" alt="Camera view from below showing animal tracking overlay and stimulus onset indicators"/>

The pipeline does two things:

1. **Automated clip generation**: detects every loom onset from the indicator
   light signal and exports a trimmed .mp4 clip for each event across one or
   more animals.

2. **Escape response analysis**: confirms manually scored escape responses
   using a speed threshold (>30 cm/s), and extracts peak speed and reaction
   time from each confirmed escape. This analysis depends on frame-by-frame
   animal tracking performed using a custom box-tracking algorithm, as described
   in Storchi et al. (2020), which is not included here.

---

## Requirements

- MATLAB R2020a or later
- Image Processing Toolbox
- VideoReader-compatible .avi recordings
- Tracking .mat files (idx, idy coordinates) generated via Storchi et al. (2020)

### Camera and Arena Setup Note
This pipeline was developed for behavioural recordings acquired using an
infrared camera positioned below the arena. Because the infrared camera cannot
detect the dark expanding disc presented on the ceiling monitor, loom onset is
instead identified from two small white indicator boxes displayed simultaneously
on the monitor at stimulus onset, which appear as bright flashes in the camera
frame. The ROI in Loom_signal_coordinate.m must be drawn over one of these
indicator boxes. This pipeline will not function correctly without this or an
equivalent frame-visible onset signal.

---

## Pipeline Overview

```
Raw .avi recording
       |
       v
[1] Loom_signal_coordinate.m
    Define the ROI over the indicator light in the camera frame
       |
       v
[2] detect_loom_onsets_to_excel.m
    Detect loom onset times from ROI brightness, export to Excel
       |
       v
[3] make_loom_clips_from_excel.m
    Generate one .mp4 clip per loom event per animal
       |
       v
    Manual behavioural scoring
    (no response / rearing / freezing / escape)
       |
       v
[4] Speed_Classification.m
    Confirm escape responses via 30 cm/s speed threshold
       |
       v
[5] Reaction_time_and_speed_scoring.m
    Extract reaction time and max speed for confirmed escapes
```

## Script Descriptions

### 1. `Loom_signal_coordinate.m`
Interactive tool for identifying the region of interest (ROI) over the
indicator light box visible in the camera frame. Displays a frame from the
video and allows the user to draw a rectangle over the indicator light.
Saves ROI coordinates as a .mat file for use in loom onset detection.

**Usage:**
```matlab
roiFile = Loom_signal_coordinate(aviPath, roiDir, tShow);
% aviPath : full path to .avi file
% roiDir  : folder to save the ROI .mat file
% tShow   : time at which to display the frame for ROI selection
%           accepts seconds (e.g. 45), 'mm:ss' (e.g. '0:45'), or 'hh:mm:ss'
```

---

### 2. `detect_loom_onsets_to_excel.m`
Detects loom stimulus onset times from ROI brightness across one or more
.avi files and writes results to a versioned Excel table. Each row contains
animal metadata and loom onset timestamps formatted as clip windows ready
for use with the next script. Output files are never overwritten -- each
run produces a new versioned file.

**Usage:**
```matlab
outXlsx = detect_loom_onsets_to_excel(aviFiles, outDir, outName);
% aviFiles : path string or cell array of paths to .avi files
% outDir   : output folder for the Excel file
% outName  : desired output filename (e.g. 'Loom_time_table.xlsx')
```

---

### 3. `make_loom_clips_from_excel.m`
Reads the loom onset Excel table produced by script 2 and generates one
trimmed .mp4 clip per loom event per animal. Videos are matched to Excel
rows using MouseID, Gender, and Experiment number parsed from the filename.
Clips are saved into subfolders named after each source video.

**Usage:**
```matlab
make_loom_clips_from_excel(excelFile, videoRoot, outDir);
% excelFile : path to Excel file from detect_loom_onsets_to_excel.m
% videoRoot : root folder containing source .avi files (searched recursively)
% outDir    : folder to save output .mp4 clips
```

---

### 4. `Speed_Classification.m`
Batch-processes .mp4 clips alongside tracking .mat files to confirm or
reject manually scored escape responses. A trial is confirmed as escape if
peak locomotor speed exceeds 30 cm/s within 5 s of loom onset. The
pixel-to-cm scaling factor must be calibrated for your arena before use
(see USER-EDIT ZONE in script).

**Usage:**
```matlab
xlsxFile = Speed_Classification(mp4Files, matFiles, outTableDir, baseExcelName);
% mp4Files      : cell array of .mp4 clip paths
% matFiles      : cell array of corresponding tracking .mat file paths
%                 (must be in the same order as mp4Files)
% outTableDir   : folder to save output Excel
% baseExcelName : base name for output file (no .xlsx extension)
```

---

### 5. `Reaction_time_and_speed_scoring.m`
Interactive semi-automated tool for extracting reaction time and maximum
escape speed from confirmed escape clips. Displays the tracked video
alongside the locomotor speed trace. The user makes two sequential clicks
on the speed trace: the first marks movement onset (used to compute
reaction time as latency from loom onset), and the second marks the escape
speed peak (used to extract maximum speed). This two-click approach allows
the genuine escape movement to be distinguished from other locomotor events
within the analysis window. Results are saved to Excel after every clip.

### `Reaction_time_and_speed_scoring_demo.m`
Standalone demonstration version of `Reaction_time_and_speed_scoring.m`.
This script is not part of the pipeline. It is provided so that anyone with
the example video file and MATLAB can run the scoring tool immediately,
without any command-line input and without needing the rest of the pipeline.
The clip, folder paths, and all analysis settings are hardcoded. Highlight
all and press Run, no further input is required.

> **Hardcoded clip:** `B_F_Exp4_Loom07_10-39_to_10-52.mp4`
> For the full pipeline version see `Reaction_time_and_speed_scoring.m`.
**Usage:**
```matlab
Reaction_time_and_speed_scoring(searchFolders, outputDir);
% searchFolders : cell array of folder paths containing .mp4 clips
%                 and corresponding tracking .mat files
% outputDir     : folder where output Excel will be saved
```

---

## Example Data

Example clips of each behavioural response category (escape, freezing,
rearing, no response) and general exploration are available here:
[Example Data (Google Drive)](https://drive.google.com/drive/folders/1-eDAcUuFldqSmIJQdIcOjjmpr3t4m-7i?usp=sharing)

---

## Acknowledgements

Animal tracking was performed using scripts developed by Storchi et al. (2020).
