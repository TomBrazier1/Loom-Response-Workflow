function make_loom_clips_from_excel(excelFile, videoRoot, outDir)
% make_loom_clips_from_excel
%
% Reads the loom onset Excel table produced by detect_loom_onsets_to_excel.m
% and generates one trimmed .mp4 clip per loom event per animal.
%
% Clips are saved into subfolders named after each source video:
%   outDir/<VideoBaseName>/MouseID_Gender_Exp#_LoomNN_m-ss_to_m-ss.mp4
%
% Videos are matched to Excel rows using MouseID, Gender, and Experiment
% number parsed from the filename. The function searches videoRoot
% recursively, so videos can be organised into subfolders.
%
% -------------------------------------------------------------------------
% USAGE:
%   make_loom_clips_from_excel(excelFile, videoRoot, outDir)
%
% INPUTS:
%   excelFile  : full path to Excel file from detect_loom_onsets_to_excel.m
%   videoRoot  : root folder containing source .avi files (searched recursively)
%   outDir     : folder where output .mp4 clips will be saved
%
% -------------------------------------------------------------------------
% VIDEO NAMING CONVENTION (required for metadata extraction):
%   MouseID_Gender_BatchName_Date.avi
%   e.g. R_F_Acuity1_03_12_2025.avi
%
% -------------------------------------------------------------------------
% DEPENDENCIES:
%   MATLAB Image Processing Toolbox (for im2uint8)
%
% -------------------------------------------------------------------------
% CITATION:
%   If you use this script please cite:
%   Brazier, T. (2026). Time-of-day differences in visually evoked defensive
%   responses in mice. MSci Thesis, University of Manchester.
% -------------------------------------------------------------------------

%% ---- Validate inputs ----
excelFile = char(excelFile);
videoRoot = char(videoRoot);
outDir    = char(outDir);

assert(isfile(excelFile),          'Excel file not found: %s', excelFile);
assert(exist(videoRoot,'dir') == 7, 'videoRoot folder not found: %s', videoRoot);

if ~exist(outDir, 'dir')
    mkdir(outDir);
end

%% ---- Find all candidate .avi files (recursive search) ----
vidList = dir(fullfile(videoRoot, '**', '*.avi'));
assert(~isempty(vidList), 'No .avi files found under: %s', videoRoot);

V = struct('fullpath', {}, 'basename', {}, 'mouseID', {}, 'gender', {}, 'experiment', {});

for i = 1:numel(vidList)
    fullp = fullfile(vidList(i).folder, vidList(i).name);
    [~, base, ~] = fileparts(fullp);

    parts = strsplit(base, '_');
    if numel(parts) < 3, continue; end

    mouseID  = parts{1};
    gender   = parts{2};
    batchStr = parts{3};

    tok = regexp(batchStr, '(\d+)', 'tokens', 'once');
    if isempty(tok), continue; end

    expNum = str2double(tok{1});
    if ~isfinite(expNum), continue; end

    V(end+1).fullpath  = fullp; %#ok<AGROW>
    V(end).basename    = base;
    V(end).mouseID     = mouseID;
    V(end).gender      = gender;
    V(end).experiment  = expNum;
end

assert(~isempty(V), 'Found .avi files but none matched the expected naming convention.');

%% ---- Read Excel table ----
raw = readcell(excelFile, 'Sheet', 1);

% Remove completely empty rows
keep = true(size(raw,1), 1);
for r = 1:size(raw,1)
    allBlank = true;
    for c = 1:numel(raw(r,:))
        if ~ismissinglike(raw{r,c})
            allBlank = false;
            break;
        end
    end
    keep(r) = ~allBlank;
end
raw = raw(keep, :);

if size(raw,1) < 2
    error('Excel file contains no data rows: %s', excelFile);
end

data  = raw(2:end, :);
nRows = size(data, 1);

fprintf('\nExcel:      %s\n', excelFile);
fprintf('Video root: %s\n', videoRoot);
fprintf('Output dir: %s\n', outDir);
fprintf('Rows to process: %d\n\n', nRows);

%% ---- Process each row ----
for r = 1:nRows
    gender     = data{r,1};
    experiment = data{r,2};
    mouseID    = data{r,3};

    if ismissinglike(gender) || ismissinglike(experiment) || ismissinglike(mouseID)
        fprintf('[Row %d] Skipping -- missing metadata.\n', r);
        continue;
    end

    gender     = char(strtrim(string(gender)));
    mouseID    = char(strtrim(string(mouseID)));
    experiment = double(experiment);

    if ~isfinite(experiment)
        fprintf('[Row %d] Skipping -- invalid experiment number.\n', r);
        continue;
    end

    % Match row to video file
    matchIdx = find(strcmpi({V.mouseID}, mouseID) & ...
                    strcmpi({V.gender},  gender)   & ...
                    [V.experiment] == experiment);

    if isempty(matchIdx)
        fprintf('[Row %d] No video match found for %s | %s | Exp%d\n', ...
            r, mouseID, gender, experiment);
        continue;
    elseif numel(matchIdx) > 1
        fprintf('[Row %d] Warning: multiple video matches found, using first.\n', r);
        matchIdx = matchIdx(1);
    end

    aviFile = V(matchIdx).fullpath;
    [~, videoBase, ~] = fileparts(aviFile);

    outSubDir = fullfile(outDir, videoBase);
    if ~exist(outSubDir, 'dir')
        mkdir(outSubDir);
    end

    fprintf('[Row %d] %s | %s | Exp%d\n  Source: %s\n', ...
        r, mouseID, gender, experiment, aviFile);

    vIn    = VideoReader(aviFile);
    fps    = vIn.FrameRate;
    vidDur = vIn.Duration;

    loomCells = data(r, 4:end);
    loomN     = 0;

    for c = 1:numel(loomCells)
        s = string(loomCells{c});
        if ismissing(s) || strlength(strtrim(s)) == 0
            continue;
        end

        clipStr = char(strtrim(s));
        [tStart, tEnd, ok] = parse_mmss_range(clipStr);
        if ~ok, continue; end

        tStart = max(0, min(tStart, vidDur));
        tEnd   = max(0, min(tEnd,   vidDur));
        if tEnd <= tStart, continue; end

        loomN   = loomN + 1;
        outName = sprintf('%s_%s_Exp%d_Loom%02d_%s_to_%s.mp4', ...
            mouseID, gender, experiment, loomN, ...
            mmss_filename(tStart), mmss_filename(tEnd));

        outFile = fullfile(outSubDir, outName);
        write_video_clip_mp4(vIn, fps, tStart, tEnd, outFile);

        fprintf('  Loom %02d: %s  ->  %s\n', loomN, clipStr, outFile);
    end
end

fprintf('\nDone.\n');
end

%% ---- Helper: check for missing/empty values ----
function tf = ismissinglike(x)
tf = false;
try; if ismissing(x), tf = true; return; end; catch; end
if isempty(x), tf = true; return; end
if ischar(x)   && isempty(strtrim(x)),                         tf = true; return; end
if isstring(x) && (ismissing(x) || strlength(strtrim(x))==0), tf = true; return; end
if isnumeric(x) && all(isnan(x(:))),                           tf = true; return; end
end

%% ---- Helper: parse 'mm:ss-mm:ss' clip string ----
function [tStart, tEnd, ok] = parse_mmss_range(s)
ok = false; tStart = NaN; tEnd = NaN;
parts = strsplit(regexprep(s, '\s+', ''), '-');
if numel(parts) ~= 2, return; end
ta = parse_mmss(parts{1});
tb = parse_mmss(parts{2});
if ~isfinite(ta) || ~isfinite(tb), return; end
tStart = ta; tEnd = tb; ok = true;
end

%% ---- Helper: parse 'mm:ss' to seconds ----
function t = parse_mmss(x)
t   = NaN;
tok = regexp(x, '^(\d+):(\d{2})$', 'tokens', 'once');
if isempty(tok), return; end
t = 60 * str2double(tok{1}) + str2double(tok{2});
end

%% ---- Helper: format seconds as 'm-ss' for filename ----
function s = mmss_filename(t)
t = max(0, t);
s = sprintf('%d-%02d', floor(t/60), floor(mod(t,60)));
end

%% ---- Helper: write trimmed .mp4 clip ----
function write_video_clip_mp4(vIn, fps, tStart, tEnd, outFile)
vw          = VideoWriter(outFile, 'MPEG-4');
vw.FrameRate = fps;
open(vw);

vIn.CurrentTime = tStart;

while hasFrame(vIn) && vIn.CurrentTime < tEnd
    frame = readFrame(vIn);
    if size(frame,3) == 1
        frame = repmat(frame, [1 1 3]);   % convert greyscale to RGB
    end
    if ~isa(frame, 'uint8')
        frame = im2uint8(frame);           % ensure correct bit depth
    end
    writeVideo(vw, frame);
end

close(vw);
end
