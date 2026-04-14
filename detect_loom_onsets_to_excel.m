function outXlsx = detect_loom_onsets_to_excel(aviFiles, outDir, outName)
% detect_loom_onsets_to_excel
%
% Detects loom stimulus onset times from indicator light ROI brightness
% for one video or a batch of videos, and writes results to a versioned
% Excel table with the following columns:
%   Gender | Experiment | Mouse ID | Loom 1 | Loom 2 | ... (up to max looms)
%
% Each loom cell contains a clip window string: 'mm:ss-mm:ss'
% (pre-loom seconds to post-loom seconds), ready for use with
% make_loom_clips_from_excel.m
%
% Output files are never overwritten. Each run produces a new versioned
% file: _0.1, _0.2, _0.3, etc.
%
% -------------------------------------------------------------------------
% USAGE:
%   outXlsx = detect_loom_onsets_to_excel(aviFiles, outDir, outName)
%
% INPUTS:
%   aviFiles : full path string OR cell array of full paths to .avi files
%   outDir   : folder where the Excel file should be saved
%   outName  : output filename (e.g. 'Loom_time_table.xlsx')
%
% OUTPUT:
%   outXlsx  : full path to the Excel file written
%
% -------------------------------------------------------------------------
% VIDEO NAMING CONVENTION (required for metadata extraction):
%   MouseID_Gender_BatchName_Date.avi
%   e.g. R_F_Acuity1_03_12_2025.avi
%
% -------------------------------------------------------------------------
% DEPENDENCIES:
%   MATLAB Image Processing Toolbox (for rgb2gray)
%
% ------------------------------------------------------------------------

%% ================== USER-EDIT ZONE (EDIT THESE VALUES ONLY) =============

% ROI position over the indicator light box in the camera frame.
% Use Loom_signal_coordinate.m to identify these values for your setup.
% Format: [x, y, width, height] in pixels.
% x and y are the top-left corner of the ROI.
roi = [0 0 15 15];  % <-- REPLACE WITH YOUR OWN VALUES

% Detection parameters -- adjust if looms are being missed or over-detected.
params = struct();
params.startFrame      = 0;      % Frame to begin analysis from (0 = start of video)
params.baselineSeconds = 10;     % Duration of baseline period at start of video (seconds)
params.smoothWinSec    = 0.25;   % Smoothing window applied to ROI brightness signal (seconds)
params.spikeFrac       = 0.50;   % Threshold: fraction between baseline and max brightness
params.useRobustMax    = true;   % Use percentile-based max (recommended; reduces noise sensitivity)
params.robustMaxPrct   = 99.5;   % Percentile used as max signal if useRobustMax is true
params.minOnSec        = 1.00;   % Minimum duration of a valid loom ON event (seconds)
params.minOffGapSec    = 0.75;   % Minimum gap between separate loom events (seconds)

% Clip window around each detected loom onset (used to generate Excel timestamps)
clip.preSec  = 5;   % Seconds before loom onset to include in clip window
clip.postSec = 8;   % Seconds after loom onset to include in clip window

%% ========================================================================

progressEveryNFrames = 500;
outNameDefault = 'Loom_time_table.xlsx';

%% ---- Validate inputs ----
if nargin < 1
    error('You must provide aviFiles (string path or cell array of paths).');
end
if nargin < 2 || isempty(outDir)
    error('You must provide outDir (folder path) as the 2nd input.');
end
if nargin < 3 || isempty(outName)
    outName = outNameDefault;
end

if ischar(aviFiles) || isstring(aviFiles)
    aviFiles = cellstr(aviFiles);
elseif ~iscell(aviFiles)
    error('aviFiles must be a path string, string array, or cell array of strings.');
end

outDir = char(outDir);
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = make_versioned_filename_decimal(outDir, outName);
fprintf('\nWill write NEW Excel file:\n  %s\n', outXlsx);

%% ---- Process each video ----
allRows = {};
maxLoomsAcrossFiles = 0;

for f = 1:numel(aviFiles)
    aviFile = char(aviFiles{f});
    assert(isfile(aviFile), 'File not found: %s', aviFile);

    fprintf('\n[%d/%d] Processing: %s\n', f, numel(aviFiles), aviFile);

    % ---- Parse filename for metadata ----
    [~, baseName, ~] = fileparts(aviFile);
    parts = strsplit(baseName, '_');
    assert(numel(parts) >= 3, ...
        'Filename must follow MouseID_Gender_BatchName_Date format. Got: %s', baseName);

    mouseID  = parts{1};
    gender   = parts{2};
    batchStr = parts{3};

    tok = regexp(batchStr, '(\d+)', 'tokens', 'once');
    assert(~isempty(tok), ...
        'Could not find experiment number in batch string "%s" (e.g. Acuity1).', batchStr);
    experiment = str2double(tok{1});

    % ---- Read video and extract ROI brightness ----
    v = VideoReader(aviFile);
    fps = v.FrameRate;
    nFramesEst = max(1, floor(v.Duration * fps));

    v.CurrentTime = 0;
    frame0 = readFrame(v);
    [H, W, ~] = size(frame0);

    x  = max(1, round(roi(1)));
    y  = max(1, round(roi(2)));
    w  = round(roi(3));
    h  = round(roi(4));
    x2 = min(W, x + w - 1);
    y2 = min(H, y + h - 1);

    roiMean = NaN(nFramesEst, 1);

    v.CurrentTime = 0;
    k = 1;
    while hasFrame(v)
        frame = readFrame(v);
        if size(frame,3) == 3
            gray = rgb2gray(frame);
        else
            gray = frame;
        end
        roiPatch = gray(y:y2, x:x2);
        roiMean(k) = mean(roiPatch(:));
        if mod(k, progressEveryNFrames) == 0
            fprintf('  frame %d\n', k);
        end
        k = k + 1;
    end

    roiMean = roiMean(1:k-1);
    nFrames = numel(roiMean);
    fprintf('  done (%d frames)\n', nFrames);

    % ---- Smooth brightness signal ----
    smoothWin = max(1, round(params.smoothWinSec * fps));
    roiMeanSm = movmean(roiMean, smoothWin);

    % ---- Apply start frame offset ----
    startIdx = max(1, round(params.startFrame));
    if startIdx > nFrames
        error('startFrame (%d) exceeds video length (%d frames).', startIdx, nFrames);
    end
    roiUse = roiMeanSm(startIdx:end);

    % ---- Compute baseline and detection threshold ----
    nBase = max(1, round(params.baselineSeconds * fps));
    endBase = min(numel(roiUse), nBase);
    baselineVal = median(roiUse(1:endBase));

    if params.useRobustMax
        maxSignal = prctile(roiUse, params.robustMaxPrct);
    else
        maxSignal = max(roiUse);
    end
    thr = baselineVal + params.spikeFrac * (maxSignal - baselineVal);

    % ---- Detect loom ON periods ----
    loomOn = roiUse >= thr;

    minOnFrames     = max(1, round(params.minOnSec * fps));
    loomOn          = remove_short_true_runs(loomOn, minOnFrames);

    minOffGapFrames = max(0, round(params.minOffGapSec * fps));
    loomOn          = merge_close_true_runs(loomOn, minOffGapFrames);

    d            = diff([false; loomOn; false]);
    onFrames_rel = find(d == 1);
    onFrames     = onFrames_rel + startIdx - 1;
    onSec        = (onFrames - 1) / fps;

    % ---- Build output row ----
    nLooms = numel(onSec);
    maxLoomsAcrossFiles = max(maxLoomsAcrossFiles, nLooms);

    loomCells = repmat({''}, 1, max(10, nLooms));
    for i = 1:nLooms
        t1 = max(0, onSec(i) - clip.preSec);
        t2 = max(0, onSec(i) + clip.postSec);
        loomCells{i} = sprintf('%s-%s', sec_to_mmss(t1), sec_to_mmss(t2));
    end

    row = cell(1, 3 + numel(loomCells));
    row{1} = gender;
    row{2} = experiment;
    row{3} = mouseID;
    row(4:end) = loomCells;

    allRows(end+1,1) = {row}; %#ok<AGROW>
end

%% ---- Write Excel ----
nColsLoom = max(10, maxLoomsAcrossFiles);

headers = cell(1, 3 + nColsLoom);
headers{1} = 'Gender';
headers{2} = 'Experiment';
headers{3} = 'Mouse ID';
for c = 1:nColsLoom
    headers{3+c} = c;
end

for r = 1:numel(allRows)
    row    = allRows{r};
    needed = (3 + nColsLoom) - numel(row);
    if needed > 0
        row = [row, repmat({''}, 1, needed)];
    elseif needed < 0
        row = row(1:(3 + nColsLoom));
    end
    allRows{r} = row;
end

dataBlock = vertcat(allRows{:});

writecell(headers,   outXlsx, 'FileType','spreadsheet', 'Sheet', 1, 'Range', 'A1');
writecell(dataBlock, outXlsx, 'FileType','spreadsheet', 'Sheet', 1, 'Range', 'A2');

fprintf('\nWrote Excel: %s\n', outXlsx);

end

%% ---- Helper: versioned filename ----
function outFile = make_versioned_filename_decimal(outDir, outName)
[~, nameOnly, ext] = fileparts(char(outName));
if isempty(ext) || ~strcmpi(ext, '.xlsx')
    ext = '.xlsx';
end
v = 1;
while true
    candidate = fullfile(outDir, [nameOnly sprintf('_0.%d', v) ext]);
    if ~isfile(candidate)
        outFile = candidate;
        return;
    end
    v = v + 1;
    if v > 9999
        error('Too many versions exist for "%s".', outName);
    end
end
end

%% ---- Helper: remove short true runs ----
function x = remove_short_true_runs(x, minLen)
d    = diff([false; x; false]);
s    = find(d == 1);
e    = find(d == -1) - 1;
keep = (e - s + 1) >= minLen;
x(:) = false;
for ii = find(keep)'
    x(s(ii):e(ii)) = true;
end
end

%% ---- Helper: merge close true runs ----
function x = merge_close_true_runs(x, maxGap)
d = diff([false; x; false]);
s = find(d == 1);
e = find(d == -1) - 1;
if isempty(s), return; end
newS = s(1); newE = e(1);
outS = []; outE = [];
for ii = 2:numel(s)
    if (s(ii) - newE - 1) <= maxGap
        newE = e(ii);
    else
        outS(end+1,1) = newS; %#ok<AGROW>
        outE(end+1,1) = newE; %#ok<AGROW>
        newS = s(ii); newE = e(ii);
    end
end
outS(end+1,1) = newS;
outE(end+1,1) = newE;
x(:) = false;
for ii = 1:numel(outS)
    x(outS(ii):outE(ii)) = true;
end
end

%% ---- Helper: seconds to mm:ss string ----
function s = sec_to_mmss(t)
t   = max(0, t);
m   = floor(t / 60);
sec = floor(mod(t, 60));
s   = sprintf('%d:%02d', m, sec);
end
