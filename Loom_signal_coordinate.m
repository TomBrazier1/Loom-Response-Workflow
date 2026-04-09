function roiFile = Loom_signal_coordinate(aviPath, roiDir, tShow)
% Loom_signal_coordinate
%
% Interactive ROI picker for identifying the position of the indicator
% light box in the camera frame. Displays a single frame from the video
% and prompts the user to draw a rectangle tightly around the indicator
% light. The resulting ROI coordinates are saved as a .mat file for use
% in detect_loom_onsets_to_excel.m.
%
% Run this script once per experimental setup. If your camera position or
% arena configuration changes, re-run to update the ROI.
%
% -------------------------------------------------------------------------
% USAGE:
%   roiFile = Loom_signal_coordinate(aviPath, roiDir, tShow)
%
% INPUTS:
%   aviPath : full path to any .avi file from your dataset
%   roiDir  : folder to save the ROI .mat file (created if it does not exist)
%             if omitted, defaults to a 'Regions_of_interest' subfolder
%             next to the video file
%   tShow   : time in the video at which to display the frame for ROI
%             selection. Choose a time when the indicator light is visible.
%             Accepts:
%               - seconds          (e.g. 707)
%               - 'mm:ss'          (e.g. '11:47')
%               - 'hh:mm:ss'       (e.g. '00:11:47')
%             if omitted, defaults to the first frame (0 s)
%
% OUTPUT:
%   roiFile : full path to the saved .mat file containing:
%               roiPos  -- [x y width height] in pixels
%               aviPath -- path to the source video
%               tSec    -- time shown for selection (seconds)
%               tShowStr -- time shown for selection (string)
%
% -------------------------------------------------------------------------
% INSTRUCTIONS:
%   1. Run the function -- a frame from the video will appear
%   2. Draw a rectangle tightly around the indicator light box
%   3. Double-click inside the rectangle to confirm
%   4. The ROI coordinates will be printed to the command window and saved
%   5. Copy the printed [x y w h] values into the USER-EDIT ZONE of
%      detect_loom_onsets_to_excel.m and Reaction_time_and_speed_scoring.m
%
% -------------------------------------------------------------------------
% DEPENDENCIES:
%   MATLAB Image Processing Toolbox (for drawrectangle or imrect)
%
% -------------------------------------------------------------------------
% CITATION:
%   If you use this script please cite:
%   Brazier, T., Allen, A., and Pienaar, A. (2026). Time-of-day differences
%   in visually evoked defensive responses in mice. MSci Thesis,
%   University of Manchester.
% -------------------------------------------------------------------------

clc;

if nargin < 2 || isempty(roiDir)
    roiDir = fullfile(fileparts(aviPath), 'Regions_of_interest');
end

if nargin < 3 || isempty(tShow)
    tSec     = 0;
    tShowStr = '0';
else
    if ischar(tShow) || isstring(tShow)
        tShowStr = char(tShow);
    else
        tShowStr = num2str(tShow);
    end
    tSec = parse_time_to_seconds(tShow);
end

if ~isfile(aviPath)
    error('AVI file not found:\n  %s', aviPath);
end

% Load video and jump to requested frame
v = VideoReader(aviPath);

if v.Duration <= 0
    error('Video duration is invalid.');
end

tSec          = max(0, min(tSec, v.Duration - eps));
v.CurrentTime = tSec;
frame         = readFrame(v);

% Display frame for ROI selection
fig = figure('Color', 'k', 'Name', 'Select Indicator Light ROI');
imshow(frame, []);
title({ ...
    sprintf('Video: %s', get_filename_only(aviPath)), ...
    sprintf('Time shown: %s  (%.3f s)', tShowStr, tSec), ...
    'Draw a rectangle tightly around the indicator light box', ...
    'Double-click inside the rectangle to confirm' ...
    }, 'Color', 'w');

% Draw ROI interactively
roiPos = [];
if exist('drawrectangle', 'file') == 2
    h      = drawrectangle('Color', 'y');
    wait(h);
    roiPos = round(h.Position);   % [x y w h]
else
    h      = imrect; %#ok<IMRECT>
    roiPos = round(wait(h));      % [x y w h]
end

fprintf('\nROI selected (x, y, width, height) = [%d %d %d %d]\n', ...
    roiPos(1), roiPos(2), roiPos(3), roiPos(4));
fprintf('Copy these values into the USER-EDIT ZONE of:\n');
fprintf('  detect_loom_onsets_to_excel.m\n');
fprintf('  Reaction_time_and_speed_scoring.m\n\n');

% Save ROI to file
if ~exist(roiDir, 'dir')
    mkdir(roiDir);
end

base    = get_filename_only(aviPath);
roiFile = fullfile(roiDir, sprintf('%s_indicator_light_roi.mat', base));

save(roiFile, 'roiPos', 'aviPath', 'tSec', 'tShowStr');
fprintf('ROI saved to:\n  %s\n', roiFile);

end

%% ---- Helper: parse time input to seconds ----
function tSec = parse_time_to_seconds(t)
% Accepts numeric seconds or strings in mm:ss or hh:mm:ss format

if isnumeric(t)
    tSec = double(t);
    return;
end

s = strtrim(char(t));

if ~isempty(regexp(s, '^[0-9]*\.?[0-9]+$', 'once'))
    tSec = str2double(s);
    return;
end

parts = regexp(s, ':', 'split');

if numel(parts) == 2
    tSec = str2double(parts{1}) * 60 + str2double(parts{2});
elseif numel(parts) == 3
    tSec = str2double(parts{1}) * 3600 + str2double(parts{2}) * 60 + str2double(parts{3});
else
    error('Could not parse time "%s". Use seconds, mm:ss, or hh:mm:ss.', s);
end
end

%% ---- Helper: extract filename without extension ----
function name = get_filename_only(p)
[~, name, ~] = fileparts(p);
end
