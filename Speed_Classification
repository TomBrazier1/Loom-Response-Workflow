function xlsxFile = Speed_Classification(mp4Files, matFiles, outTableDir, baseExcelName)
% Speed_Classification
%
% Confirms or rejects manually scored escape responses by checking whether
% peak locomotor speed exceeds a defined threshold within a fixed window
% following loom onset. Results are written to a versioned Excel file.
%
% Loom onset is assumed to occur at exactly 5.0 s into each clip, as clips
% are generated with a 5 s pre-loom window by make_loom_clips_from_excel.m.
% A trial is classified as a confirmed escape if speed >= 30 cm/s for at
% least one frame between 5.0 s and 10.0 s post-clip-start.
%
% Speed is calculated from frame-by-frame pixel displacements in the
% tracking .mat files, converted to cm/s using a pixel-to-cm scaling
% factor. This factor must be calibrated for your arena dimensions before
% use (see USER-EDIT ZONE below).
%
% -------------------------------------------------------------------------
% USAGE:
%   xlsxFile = Speed_Classification(mp4Files, matFiles, outTableDir, baseExcelName)
%
% INPUTS:
%   mp4Files      : cell array of full paths to .mp4 clip files
%   matFiles      : cell array of full paths to corresponding tracking .mat
%                   files (must contain idx and idy coordinate vectors)
%                   Files must be provided in the same order as mp4Files.
%   outTableDir   : folder to save the output Excel file
%   baseExcelName : base filename for output Excel (no .xlsx extension)
%
% OUTPUT:
%   xlsxFile : full path to the Excel file written
%
% -------------------------------------------------------------------------
% TRACKING DEPENDENCY:
%   Tracking .mat files must contain idx and idy variables representing
%   frame-by-frame y and x coordinates of the animal centroid, as produced
%   by the box-tracking algorithm described in Storchi et al. (2020).
%
% -------------------------------------------------------------------------
% VIDEO NAMING CONVENTION (required for metadata extraction):
%   MouseID_Gender_BatchName_Date_..._Loom##_....mp4
%   e.g. R_F_Acuity1_Loom01_0-0_to_0-13.mp4
%
% -------------------------------------------------------------------------
% CITATION:
%   If you use this script please cite:
%   Brazier, T. (2026). Time-of-day differences in visually evoked defensive
%   responses in mice. MSci Thesis, University of Manchester.
% -------------------------------------------------------------------------

%% ================== USER-EDIT ZONE (EDIT THESE VALUES ONLY) =============

% Pixel-to-centimetre scaling factor.
% Calibrate this for your arena by measuring a known distance in pixels.
% Default value below was derived from a 30 x 45 cm arena recorded at the
% resolution used in Brazier (2026) -- replace with your own measurement.
cm_per_px = 0.03656269;  % <-- REPLACE WITH YOUR OWN VALUE

% Expected frame rate. A warning is issued if the actual video fps differs.
fps_expected = 15;

% Timing parameters (seconds).
% These assume clips were generated with 5 s pre-loom by make_loom_clips_from_excel.m
loomOnsetSec      = 5.0;   % Time of loom onset within each clip (seconds)
analysisEndSec    = 10.0;  % End of speed analysis window (seconds from clip start)

% Escape speed threshold (cm/s).
% A trial is confirmed as escape if this speed is reached for >= 1 frame.
speedThreshCmPerS = 30;

%% ========================================================================

clc;

if nargin < 3
    error('Usage: xlsxFile = Speed_Classification(mp4Files, matFiles, outTableDir, baseExcelName)');
end
if nargin < 4 || isempty(baseExcelName)
    baseExcelName = 'speed_classification';
end

if ischar(mp4Files) || isstring(mp4Files), mp4Files = cellstr(mp4Files); end
if ischar(matFiles) || isstring(matFiles), matFiles = cellstr(matFiles); end

if ~iscell(mp4Files) || isempty(mp4Files)
    error('mp4Files must be a non-empty cell array of MP4 full paths.');
end
if ~iscell(matFiles) || isempty(matFiles)
    error('matFiles must be a non-empty cell array of MAT full paths.');
end
if numel(mp4Files) ~= numel(matFiles)
    error('mp4Files and matFiles must be the same length. Got %d MP4 and %d MAT.', ...
        numel(mp4Files), numel(matFiles));
end
if ~exist(outTableDir, 'dir')
    mkdir(outTableDir);
end

%% ---- Output table setup ----
nLoomCols = 12;
headers   = [{'Gender','Experiment','Mouse ID'}, ...
              arrayfun(@num2str, 1:nLoomCols, 'UniformOutput', false)];
allRows   = cell(numel(mp4Files), numel(headers));

%% ---- Process each clip ----
for vIdx = 1:numel(mp4Files)

    mp4File  = mp4Files{vIdx};
    trackMat = matFiles{vIdx};

    if ~isfile(mp4File),  error('MP4 file not found:\n  %s', mp4File);  end
    if ~isfile(trackMat), error('MAT file not found:\n  %s', trackMat); end

    fprintf('\n[%d/%d] %s\n', vIdx, numel(mp4Files), mp4File);

    % Parse metadata from filename
    [~, baseName, ~] = fileparts(mp4File);
    [Gender, Experiment, MouseID] = infer_meta_from_filename(baseName);

    % Load video info
    v            = VideoReader(mp4File);
    fps_video    = v.FrameRate;
    nFramesVideo = max(1, floor(v.Duration * fps_video));

    if abs(fps_video - fps_expected) > 0.01
        warning('Video fps is %.3f, not %d as expected. Using actual fps.', ...
            fps_video, fps_expected);
    end

    % Load tracking coordinates
    S = load(trackMat);
    for ii = 1:2
        varName = {'idx','idy'};
        if ~isfield(S, varName{ii})
            error('Tracking MAT missing variable "%s". Fields present: %s', ...
                varName{ii}, strjoin(fieldnames(S), ', '));
        end
    end

    idx = S.idx(:);
    idy = S.idy(:);

    nFramesTrack = numel(idx);
    nFrames      = min(nFramesVideo, nFramesTrack);

    if nFrames ~= nFramesVideo || nFrames ~= nFramesTrack
        warning('Video frames (%d) and tracking frames (%d) differ. Using first %d.', ...
            nFramesVideo, nFramesTrack, nFrames);
    end

    idx = idx(1:nFrames);
    idy = idy(1:nFrames);

    % Compute frame-by-frame speed
    dx          = [NaN; diff(idx)];
    dy          = [NaN; diff(idy)];
    speed_cm_s  = sqrt(dx.^2 + dy.^2) * cm_per_px * fps_video;

    % Define analysis window
    startFrame = floor(loomOnsetSec   * fps_video) + 1;
    endFrame   = floor(analysisEndSec * fps_video);

    if startFrame > nFrames
        warning('Analysis window start exceeds clip length for %s. Marking as No.', mp4File);
        speedResponse = false;
    else
        endFrame = min(endFrame, nFrames);
        if endFrame < startFrame
            warning('Invalid analysis window for %s. Marking as No.', mp4File);
            speedResponse = false;
        else
            spWin       = speed_cm_s(startFrame:endFrame);
            aboveThresh = spWin >= speedThreshCmPerS;
            aboveThresh(isnan(aboveThresh)) = false;
            speedResponse = any(aboveThresh);
        end
    end

    % Build output row
    row    = cell(1, numel(headers));
    row{1} = Gender;
    row{2} = Experiment;
    row{3} = MouseID;
    row{4} = sprintf('%s - %s', format_mmss(loomOnsetSec, 2), ...
                     ternary(speedResponse, 'Yes', 'No'));

    allRows(vIdx, :) = row;

end

%% ---- Write Excel ----
outCell  = [headers; allRows];
xlsxFile = next_versioned_xlsx(outTableDir, baseExcelName);
writecell(outCell, xlsxFile);

fprintf('\nSaved Excel output:\n  %s\n', xlsxFile);

end

%% ---- Helper: conditional output ----
function out = ternary(cond, a, b)
if cond, out = a; else, out = b; end
end

%% ---- Helper: versioned output filename ----
function xlsxPath = next_versioned_xlsx(folder, baseName)
xlsx0 = fullfile(folder, sprintf('%s.xlsx', baseName));
if ~isfile(xlsx0), xlsxPath = xlsx0; return; end
k = 1;
while true
    candidate = fullfile(folder, sprintf('%s_%d.xlsx', baseName, k));
    if ~isfile(candidate), xlsxPath = candidate; return; end
    k = k + 1;
end
end

%% ---- Helper: extract metadata from clip filename ----
function [Gender, Experiment, MouseID] = infer_meta_from_filename(baseName)
Gender     = '';
Experiment = '';
MouseID    = '';

parts = strsplit(baseName, '_');
if ~isempty(parts), MouseID = parts{1}; end

for i = 1:numel(parts)
    if strcmpi(parts{i}, 'F'),  Gender = 'F'; break; end
    if strcmpi(parts{i}, 'M'),  Gender = 'M'; break; end
end

for i = 1:numel(parts)
    m = regexp(parts{i}, 'Acuity(\d+)', 'tokens', 'once', 'ignorecase');
    if ~isempty(m), Experiment = m{1}; break; end
end
end

%% ---- Helper: format seconds as mm:ss.dd string ----
function s = format_mmss(tSec, nDec)
if nargin < 2, nDec = 2; end
if isnan(tSec), s = ''; return; end
scale    = 10^nDec;
tRounded = round(tSec * scale) / scale;
mm       = floor(tRounded / 60);
ss       = tRounded - 60 * mm;
if ss >= 60, mm = mm + 1; ss = ss - 60; end
secFmt = sprintf('%%0%d.%df', 2 + 1 + nDec, nDec);
s      = sprintf('%02d:%s', mm, sprintf(secFmt, ss));
end
