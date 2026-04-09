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

### Design
<img width="2088" height="2016" alt="new_cage_design" src="https://github.com/user-attachments/assets/3aa1856b-e0ea-4531-b1c6-9a95da4378a6" />
![Design of looming assay arena](path_to_image/design.png)

**Figure 1A. Arena design.** Behavioural experiments were conducted in a rectangular Plexiglas arena with three walls rendered opaque using matte white paint and one partially transparent wall for identification and infrared illumination. Visual stimuli were presented from a monitor mounted above the arena. The transparent Plexiglas floor enabled recording from a camera positioned below the arena, while an external red Plexiglas sheet prevented the floor from appearing transparent to the animal. The red shaded region indicates the area in which loom onset was triggered.
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
infrared camera positioned below the arena. Because the infrared camera cannot detect the dark expanding disc 
presented on the ceiling monitor, loom onset is instead identified from two 
small white indicator boxes displayed simultaneously on the monitor at stimulus 
onset, which appear as bright flashes in the camera frame. The ROI in 
Loom_signal_coordinate.m must be drawn over one of these indicator boxes. 
This pipeline will not function correctly without this or an equivalent 
frame-visible onset signal.

## Pipeline Overview

```
Raw .avi recording
       |
       v
[1] Loom_signal_coordinate.m
    Define the ROI over the indicator light in the camera frame
       |
       v
[2] batch2_detect_loom_from_roi_poc_to_excel.m
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
[5] Reaction_time_manual_scoring.m
    Extract peak speed and reaction time for confirmed escapes
```
