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

![Arena view showing loom disc, animal tracking overlay, and stimulus onset 
indicators](example_data/arena_example.png)<img width="1280" height="1024" alt="frame_080_threat" src="https://github.com/user-attachments/assets/4b15870c-ff73-4049-b4f2-8b821ac900cb" />

The pipeline does two things:

1. **Automated clip generation** -- detects every loom onset from the indicator
   light signal and exports a trimmed .mp4 clip for each event across one or
   more animals.

2. **Escape response analysis** -- confirms manually scored escape responses
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
