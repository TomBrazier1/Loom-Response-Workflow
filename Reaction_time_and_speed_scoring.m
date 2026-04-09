function Reaction_time_and_speed_scoring(searchFolders, outputDir)
% Reaction_time_and_speed_scoring
%
% Interactive tool for semi-automated extraction of reaction time (RT) and
% maximum escape speed from tracked loom response video clips.
%
% For each clip, the user is shown a side-by-side display of the tracked
% video (left) and the locomotor speed trace (right). The user makes two
% sequential clicks on the speed trace:
%
%   Click 1 -- marks the onset of the escape movement (used to compute RT
%              as the latency from loom onset to movement onset)
%   Click 2 -- marks the peak of the escape movement (used to extract
%              maximum speed at that point)
%
% This two-click approach allows the user to distinguish the genuine escape
% peak from other locomotor events within the analysis window, which is
% necessary when multiple speed peaks occur after loom onset.
%
% Results are written to Excel after every clip so no data is lost if the
% session is interrupted.
%
% -------------------------------------------------------------------------
% USAGE:
%   Reaction_time_and_speed_scoring(searchFolders, outputDir)
%
% INPUTS:
%   searchFolders : cell array of folder paths containing .mp4 clips and
%                   corresponding tracking .mat files
%   outputDir     : folder path where the output Excel file will be saved
%
% EXAMPLE:
%   Reaction_time_and_speed_scoring( ...
%       {'/path/to/clips/Week1', '/path/to/clips/Week2'}, ...
%       '/path/to/output')
%
% -------------------------------------------------------------------------
% CONTROLS:
%   Play/Pause  -- spacebar or button
%   Step back   -- rewinds 10 frames
%   Restart     -- returns to frame 1
%   Click 1     -- mark movement onset (reaction time)
%   Click 2     -- mark speed peak (max speed)
%   Confirm     -- save both values and advance to next clip
%   Skip        -- record NaN for both values and advance
%   Redo        -- clear both clicks and start again
%
% -------------------------------------------------------------------------
% TRACKING DEPENDENCY:
%   Tracking .mat files must contain idx and idy variables representing
%   frame-by-frame y and x coordinates of the animal centroid, as produced
%   by the box-tracking algorithm described in Storchi et al. (2020).
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

% Speed threshold line shown on plot for reference (cm/s).
speed_thresh_cm_s = 30;

% Brightness threshold for automatic loom onset detection from indicator light.
% The ROI below is scanned each frame; the first frame exceeding this
% brightness value is taken as loom onset.
brightness_thresh = 65;  % <-- adjust if onset detection is unreliable

% ROI over the indicator light box in the camera frame [x, y, width, height].
% Use Loom_signal_coordinate.m to identify these values for your setup.
roi = [0, 0, 56, 61];  % <-- REPLACE WITH YOUR OWN VALUES

% Trail length (number of frames shown in the tracking overlay trail).
trail_length = 15;

%% ========================================================================

clc; close all;

if nargin < 2
    error(['Usage: Reaction_time_and_speed_scoring(searchFolders, outputDir)\n' ...
           'searchFolders must be a cell array of folder paths.']);
end

if ischar(searchFolders) || isstring(searchFolders)
    searchFolders = cellstr(searchFolders);
end

if ~exist(outputDir, 'dir')
    mkdir(outputDir);
    fprintf('Created output folder: %s\n', outputDir);
end

output_file = fullfile(outputDir, 'RT_and_MaxSpeed_by_Loom.xlsx');

%% ---- Collect clip list ----
clip_list = struct('video_path', {}, 'mat_path', {}, 'mouse_id', {}, 'loom_num', {});

for f = 1:length(searchFolders)
    folder   = char(searchFolders{f});
    mp4_list = dir(fullfile(folder, '*.mp4'));

    for i = 1:length(mp4_list)
        mp4_name = mp4_list(i).name;

        tok_loom = regexp(mp4_name, 'Loom(\d+)', 'tokens');
        if isempty(tok_loom), continue; end
        loom_num = str2double(tok_loom{1}{1});
        if loom_num < 1 || loom_num > 10, continue; end

        parts     = strsplit(mp4_name, '_');
        mouse_tag = parts{1};
        tok_exp   = regexp(mp4_name, '(Exp\d+)', 'tokens');
        if isempty(tok_exp), continue; end
        mouse_id = [mouse_tag '_' tok_exp{1}{1}];

        mat_path = fullfile(folder, [mp4_name '_track_box1_ALLframes.mat']);
        if ~exist(mat_path, 'file'), continue; end

        k = length(clip_list) + 1;
        clip_list(k).video_path = fullfile(folder, mp4_name);
        clip_list(k).mat_path   = mat_path;
        clip_list(k).mouse_id   = mouse_id;
        clip_list(k).loom_num   = loom_num;
    end
end

n_clips = length(clip_list);
fprintf('Found %d clips to score.\n\n', n_clips);
if n_clips == 0
    error('No valid clips found. Check folder paths and file naming.');
end

%% ---- Results store ----
% Each result holds RT and max speed for one clip
results = struct('mouse_id', {}, 'loom_num', {}, ...
                 'reaction_time', {}, 'max_speed', {});

%% ====================================================================
%% MAIN LOOP
%% ====================================================================
for c = 1:n_clips

    video_path = clip_list(c).video_path;
    mat_path   = clip_list(c).mat_path;
    mouse_id   = clip_list(c).mouse_id;
    loom_num   = clip_list(c).loom_num;

    fprintf('=== Clip %d / %d : %s  Loom%d ===\n', c, n_clips, mouse_id, loom_num);

    try
        %% ---- Load tracking data ----
        data  = load(mat_path);
        raw_x = data.idy(:);
        raw_y = data.idx(:);
        N     = length(raw_x);

        v_in         = VideoReader(video_path);
        fps          = v_in.FrameRate;
        img_height   = v_in.Height;
        img_width    = v_in.Width;
        nFramesVideo = max(1, floor(v_in.Duration * fps));
        nFrames      = min(nFramesVideo, N);

        %% ---- Compute speed ----
        x_col      = raw_x(1:nFrames);
        y_col      = raw_y(1:nFrames);
        dx         = [NaN; diff(x_col)];
        dy         = [NaN; diff(y_col)];
        speed_cm_s = sqrt(dx.^2 + dy.^2) * cm_per_px * fps;
        time_s     = (0:nFrames-1)' / fps;

        %% ---- Detect loom onset from indicator light ----
        rx = max(1, roi(1));  ry = max(1, roi(2));
        rw = roi(3);          rh = roi(4);

        onset_frame = NaN;
        v_in.CurrentTime = 0;
        frame_idx = 0;

        while hasFrame(v_in)
            frame     = readFrame(v_in);
            frame_idx = frame_idx + 1;
            if frame_idx > nFrames, break; end
            [H, W, ~] = size(frame);
            rx2 = min(W, rx + rw - 1);
            ry2 = min(H, ry + rh - 1);
            if size(frame,3) == 3
                gray = rgb2gray(frame);
            else
                gray = frame;
            end
            roi_patch  = gray(ry:ry2, rx:rx2);
            brightness = mean(roi_patch(:));
            if brightness > brightness_thresh
                onset_frame = frame_idx;
                break
            end
        end

        if isnan(onset_frame)
            fprintf('  [NO ONSET] Brightness never exceeded threshold of %d.\n', brightness_thresh);
            onset_time_s = NaN;
        else
            onset_time_s = (onset_frame - 1) / fps;
            fprintf('  Onset detected at frame %d (%.3f s)\n', onset_frame, onset_time_s);
        end

        %% ---- Pre-render tracked frames ----
        fprintf('  Pre-rendering %d frames...\n', nFrames);
        tracked_frames = cell(1, nFrames);
        for n = 1:nFrames
            frame = read(v_in, n);
            dot_x = max(1, min(img_width,  round(raw_x(n))));
            dot_y = max(1, min(img_height, round(raw_y(n))));
            trail_start = max(1, n - trail_length);
            trail_x = round(raw_x(trail_start:n));
            trail_y = round(raw_y(trail_start:n));
            for t = 2:length(trail_x)
                x1 = max(1, min(img_width,  trail_x(t-1)));
                y1 = max(1, min(img_height, trail_y(t-1)));
                x2 = max(1, min(img_width,  trail_x(t)));
                y2 = max(1, min(img_height, trail_y(t)));
                frame = local_draw_line(frame, x1, y1, x2, y2, [255 200 0], 2);
            end
            frame = local_draw_circle(frame, dot_x, dot_y, 8, [0 255 0]);
            frame = local_insert_text(frame, sprintf('Frame %d / %d', n, nFrames));
            tracked_frames{n} = frame;
        end

        %% ---- Build figure ----
        fig_title = sprintf('%s  |  Loom%d  |  Clip %d of %d  --  Click 1: RT onset   Click 2: Speed peak', ...
            mouse_id, loom_num, c, n_clips);

        fig = figure('Name', fig_title, ...
                     'NumberTitle', 'off', ...
                     'Color', [0.12 0.12 0.12], ...
                     'Position', [60 60 1440 620], ...
                     'KeyPressFcn', @key_handler);

        %% Left panel -- video
        ax_vid = axes('Parent', fig, ...
                      'Position', [0.01 0.13 0.45 0.84], ...
                      'Color', 'k');
        axis(ax_vid, 'off');
        h_img = imshow(tracked_frames{1}, 'Parent', ax_vid);

        %% Right panel -- speed trace
        ax_spd = axes('Parent', fig, ...
                      'Position', [0.53 0.18 0.44 0.76]);
        hold(ax_spd, 'on');

        valid_speed = speed_cm_s(~isnan(speed_cm_s));
        if isempty(valid_speed)
            ymax = speed_thresh_cm_s * 2;
        else
            ymax = max(speed_thresh_cm_s * 1.5, max(valid_speed) * 1.1);
        end
        ymax = max(ymax, 40);

        % Shaded post-onset window
        if ~isnan(onset_time_s)
            patch(ax_spd, ...
                [onset_time_s, time_s(end), time_s(end), onset_time_s], ...
                [0, 0, ymax, ymax], ...
                [0.85 0.92 1.0], 'EdgeColor', 'none', 'FaceAlpha', 0.45);
        end

        % Speed trace
        plot(ax_spd, time_s, speed_cm_s, ...
            'Color', [0.15 0.15 0.15], 'LineWidth', 1.4);

        % Threshold line
        plot(ax_spd, [time_s(1), time_s(end)], ...
            [speed_thresh_cm_s, speed_thresh_cm_s], ...
            '--', 'Color', [0.85 0.1 0.1], 'LineWidth', 1.2);
        text(ax_spd, time_s(end), speed_thresh_cm_s, ...
            sprintf('  %d cm/s', speed_thresh_cm_s), ...
            'Color', [0.85 0.1 0.1], 'FontSize', 9, ...
            'VerticalAlignment', 'bottom', 'HorizontalAlignment', 'right');

        % Onset marker
        if ~isnan(onset_time_s)
            xline(ax_spd, onset_time_s, ':', ...
                'Color', [0.2 0.2 0.9], 'LineWidth', 1.5, ...
                'Label', 'Loom onset', 'LabelVerticalAlignment', 'top');
        end

        xlabel(ax_spd, 'Time (s)');
        ylabel(ax_spd, 'Speed (cm/s)');
        ylim(ax_spd, [0 ymax]);
        xlim(ax_spd, [time_s(1) time_s(end)]);
        set(ax_spd, 'FontSize', 11, 'Box', 'on');

        % Instructions text on plot
        text(ax_spd, time_s(1) + 0.1, ymax * 0.97, ...
            'Click 1: movement onset (RT)     Click 2: speed peak (max speed)', ...
            'Color', [0.3 0.3 0.3], 'FontSize', 9, ...
            'VerticalAlignment', 'top');

        %% ---- Annotation handles (shared across callbacks) ----
        h_rt_line  = [];  h_rt_dot  = [];
        h_rt_text  = [];  h_rt_arrow = [];
        h_spd_line = [];  h_spd_dot = [];
        h_spd_text = [];

        %% ---- Playback state ----
        state = struct();
        state.frame         = 1;
        state.playing       = false;
        state.confirmed     = false;
        state.skipped       = false;
        % Click 1 -- RT onset
        state.click1_time   = NaN;
        state.reaction_time = NaN;
        % Click 2 -- speed peak
        state.click2_time   = NaN;
        state.max_speed     = NaN;
        state.click_count   = 0;  % tracks whether we are awaiting click 1 or 2
        fig.UserData = state;

        %% ---- Buttons ----
        btn_w = 0.07; btn_h = 0.07; btn_y = 0.02;

        uicontrol('Style','pushbutton','String','Play/Pause', ...
            'Units','normalized','Position',[0.01 btn_y btn_w btn_h], ...
            'Callback',@btn_playpause,'FontSize',10);
        uicontrol('Style','pushbutton','String','Step Back', ...
            'Units','normalized','Position',[0.09 btn_y btn_w btn_h], ...
            'Callback',@btn_stepback,'FontSize',10);
        uicontrol('Style','pushbutton','String','Restart', ...
            'Units','normalized','Position',[0.17 btn_y btn_w btn_h], ...
            'Callback',@btn_restart,'FontSize',10);
        uicontrol('Style','pushbutton','String','Confirm', ...
            'Units','normalized','Position',[0.63 btn_y btn_w btn_h], ...
            'Callback',@btn_confirm,'FontSize',10, ...
            'BackgroundColor',[0.2 0.7 0.3],'ForegroundColor','w');
        uicontrol('Style','pushbutton','String','Skip', ...
            'Units','normalized','Position',[0.71 btn_y btn_w btn_h], ...
            'Callback',@btn_skip,'FontSize',10, ...
            'BackgroundColor',[0.8 0.4 0.1],'ForegroundColor','w');
        uicontrol('Style','pushbutton','String','Redo', ...
            'Units','normalized','Position',[0.79 btn_y btn_w btn_h], ...
            'Callback',@btn_redo,'FontSize',10);

        set(fig, 'WindowButtonDownFcn', @fig_click);
        set(ax_spd, 'ButtonDownFcn',    @axes_click);

        %% ---- Playback loop ----
        while isvalid(fig)
            state = fig.UserData;
            if state.confirmed || state.skipped, break; end

            if state.playing
                state.frame = state.frame + 1;
                if state.frame > nFrames
                    state.frame   = nFrames;
                    state.playing = false;
                end
                fig.UserData = state;
                set(h_img, 'CData', tracked_frames{state.frame});
                drawnow;
                pause(1/fps);
            else
                pause(0.03);
            end
        end

        if ~isvalid(fig), break; end
        state = fig.UserData;

        %% ---- Record result ----
        k = length(results) + 1;
        results(k).mouse_id      = mouse_id;
        results(k).loom_num      = loom_num;

        if state.skipped
            results(k).reaction_time = NaN;
            results(k).max_speed     = NaN;
            fprintf('  Skipped -- NaN recorded for RT and max speed.\n');
        else
            results(k).reaction_time = state.reaction_time;
            results(k).max_speed     = state.max_speed;
            fprintf('  Saved -- RT = %.3f s  |  Max speed = %.2f cm/s\n', ...
                state.reaction_time, state.max_speed);
        end

        close(fig);
        write_excel(results, output_file);

    catch ME
        fprintf('  ERROR on clip %d: %s\n', c, ME.message);
        if exist('fig','var') && isvalid(fig), close(fig); end
    end

end % end main loop

fprintf('\nAll clips scored. Output saved to:\n  %s\n', output_file);

%% ====================================================================
%% NESTED CALLBACK FUNCTIONS
%% ====================================================================

    function btn_playpause(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.playing = ~state.playing;
        fig.UserData  = state;
    end

    function btn_stepback(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.playing = false;
        state.frame   = max(1, state.frame - 10);
        fig.UserData  = state;
        set(h_img, 'CData', tracked_frames{state.frame});
        drawnow;
    end

    function btn_restart(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.playing = false;
        state.frame   = 1;
        fig.UserData  = state;
        set(h_img, 'CData', tracked_frames{1});
        drawnow;
    end

    function btn_confirm(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.confirmed = true;
        state.playing   = false;
        fig.UserData    = state;
    end

    function btn_skip(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.skipped = true;
        state.playing = false;
        fig.UserData  = state;
    end

    function btn_redo(~, ~)
        if ~isvalid(fig), return; end
        state = fig.UserData;
        state.click1_time   = NaN;
        state.click2_time   = NaN;
        state.reaction_time = NaN;
        state.max_speed     = NaN;
        state.click_count   = 0;
        fig.UserData = state;
        clear_annotations();
        drawnow;
    end

    function key_handler(~, evt)
        if strcmp(evt.Key, 'space')
            btn_playpause([], []);
        end
    end

    function fig_click(~, ~)
        % Only act on clicks inside the speed axes
        cp      = get(fig, 'CurrentPoint');
        fig_pos = get(fig, 'Position');
        ax_norm = ax_spd.Position;
        ax_px_x = ax_norm(1) * fig_pos(3);
        ax_px_y = ax_norm(2) * fig_pos(4);
        ax_px_w = ax_norm(3) * fig_pos(3);
        ax_px_h = ax_norm(4) * fig_pos(4);

        if cp(1) < ax_px_x || cp(1) > ax_px_x + ax_px_w || ...
           cp(2) < ax_px_y || cp(2) > ax_px_y + ax_px_h
            return;
        end

        ax_cp     = get(ax_spd, 'CurrentPoint');
        clicked_t = max(time_s(1), min(time_s(end), ax_cp(1,1)));
        process_click(clicked_t);
    end

    function axes_click(~, ~)
        ax_cp     = get(ax_spd, 'CurrentPoint');
        clicked_t = max(time_s(1), min(time_s(end), ax_cp(1,1)));
        process_click(clicked_t);
    end

    function process_click(clicked_t)
        % Routes click 1 to RT annotation, click 2 to max speed annotation
        state = fig.UserData;

        if state.click_count == 0
            % First click -- reaction time onset
            if ~isnan(onset_time_s) && clicked_t > onset_time_s
                rt = clicked_t - onset_time_s;
            else
                rt = NaN;
            end
            state.click1_time   = clicked_t;
            state.reaction_time = rt;
            state.click_count   = 1;
            fig.UserData = state;
            draw_rt_annotation(clicked_t, rt);
            fprintf('  Click 1 recorded -- RT = %.3f s. Now click the speed peak.\n', ...
                rt);

        elseif state.click_count == 1
            % Second click -- max speed peak
            [~, nearest_idx] = min(abs(time_s - clicked_t));
            peak_speed = speed_cm_s(nearest_idx);
            if isnan(peak_speed), peak_speed = 0; end
            state.click2_time = clicked_t;
            state.max_speed   = peak_speed;
            state.click_count = 2;
            fig.UserData = state;
            draw_speed_annotation(clicked_t, peak_speed);
            fprintf('  Click 2 recorded -- Max speed = %.2f cm/s.\n', peak_speed);
        end
        % After two clicks, user presses Confirm or Redo
    end

    function draw_rt_annotation(clicked_t, rt)
        % Green vertical line and label for RT onset click
        if ~isempty(h_rt_line)  && isvalid(h_rt_line),  delete(h_rt_line);  end
        if ~isempty(h_rt_dot)   && isvalid(h_rt_dot),   delete(h_rt_dot);   end
        if ~isempty(h_rt_text)  && isvalid(h_rt_text),  delete(h_rt_text);  end
        if ~isempty(h_rt_arrow) && isvalid(h_rt_arrow), delete(h_rt_arrow); end

        hold(ax_spd, 'on');

        h_rt_line = xline(ax_spd, clicked_t, '-', ...
            'Color', [0.0 0.6 0.3], 'LineWidth', 2.0);

        [~, ni] = min(abs(time_s - clicked_t));
        h_rt_dot = plot(ax_spd, time_s(ni), speed_cm_s(ni), 'o', ...
            'MarkerSize', 9, 'MarkerFaceColor', [0.0 0.6 0.3], ...
            'MarkerEdgeColor', 'w', 'LineWidth', 1.5);

        annotation_y = ymax * 0.12;
        if ~isnan(rt)
            label_str = sprintf('RT = %.3f s', rt);
        else
            label_str = 'Before onset!';
        end

        h_rt_text = text(ax_spd, clicked_t, annotation_y * 2.2, label_str, ...
            'Color', [0.0 0.55 0.25], 'FontSize', 11, 'FontWeight', 'bold', ...
            'HorizontalAlignment', 'center', 'VerticalAlignment', 'bottom', ...
            'BackgroundColor', [0.9 1.0 0.93], 'EdgeColor', [0.0 0.6 0.3], ...
            'Margin', 3);

        if ~isnan(onset_time_s) && ~isnan(rt)
            t_range = time_s(end) - time_s(1);
            ax_pos  = ax_spd.Position;
            x_on    = ax_pos(1) + ax_pos(3) * (onset_time_s - time_s(1)) / t_range;
            x_cl    = ax_pos(1) + ax_pos(3) * (clicked_t    - time_s(1)) / t_range;
            y_arr   = ax_pos(2) + ax_pos(4) * (annotation_y / ymax);
            h_rt_arrow = annotation(fig, 'doublearrow', ...
                'X', [x_on, x_cl], 'Y', [y_arr, y_arr], ...
                'Color', [0.0 0.6 0.3], 'LineWidth', 1.5, 'HeadSize', 8);
        end

        hold(ax_spd, 'off');
        drawnow;
    end

    function draw_speed_annotation(clicked_t, peak_speed)
        % Orange vertical line and label for max speed click
        if ~isempty(h_spd_line) && isvalid(h_spd_line), delete(h_spd_line); end
        if ~isempty(h_spd_dot)  && isvalid(h_spd_dot),  delete(h_spd_dot);  end
        if ~isempty(h_spd_text) && isvalid(h_spd_text), delete(h_spd_text); end

        hold(ax_spd, 'on');

        h_spd_line = xline(ax_spd, clicked_t, '-', ...
            'Color', [0.9 0.5 0.0], 'LineWidth', 2.0);

        [~, ni] = min(abs(time_s - clicked_t));
        h_spd_dot = plot(ax_spd, time_s(ni), speed_cm_s(ni), 'o', ...
            'MarkerSize', 9, 'MarkerFaceColor', [0.9 0.5 0.0], ...
            'MarkerEdgeColor', 'w', 'LineWidth', 1.5);

        h_spd_text = text(ax_spd, clicked_t, peak_speed + ymax * 0.05, ...
            sprintf('Max = %.1f cm/s', peak_speed), ...
            'Color', [0.8 0.4 0.0], 'FontSize', 11, 'FontWeight', 'bold', ...
            'HorizontalAlignment', 'center', 'VerticalAlignment', 'bottom', ...
            'BackgroundColor', [1.0 0.95 0.88], 'EdgeColor', [0.9 0.5 0.0], ...
            'Margin', 3);

        hold(ax_spd, 'off');
        drawnow;
    end

    function clear_annotations()
        handles = {h_rt_line, h_rt_dot, h_rt_text, h_rt_arrow, ...
                   h_spd_line, h_spd_dot, h_spd_text};
        for h = 1:numel(handles)
            if ~isempty(handles{h}) && isvalid(handles{h})
                delete(handles{h});
            end
        end
        h_rt_line  = []; h_rt_dot   = []; h_rt_text  = []; h_rt_arrow = [];
        h_spd_line = []; h_spd_dot  = []; h_spd_text = [];
    end

end % end main function

%% ====================================================================
%% EXCEL WRITER
%% ====================================================================
function write_excel(results, output_file)

all_mouse_ids = unique({results.mouse_id});
all_looms     = 1:10;
n_mice        = length(all_mouse_ids);
n_looms       = length(all_looms);

rt_matrix    = NaN(n_looms, n_mice);
speed_matrix = NaN(n_looms, n_mice);

for r = 1:length(results)
    loom_idx  = results(r).loom_num;
    mouse_idx = find(strcmp(all_mouse_ids, results(r).mouse_id));
    if loom_idx >= 1 && loom_idx <= n_looms
        rt_matrix(loom_idx,    mouse_idx) = results(r).reaction_time;
        speed_matrix(loom_idx, mouse_idx) = results(r).max_speed;
    end
end

% Write RT sheet
write_sheet(output_file, rt_matrix,    all_mouse_ids, all_looms, 'Reaction Time (s)');

% Write max speed sheet
write_sheet(output_file, speed_matrix, all_mouse_ids, all_looms, 'Max Speed (cm per s)');

end

function write_sheet(output_file, data_matrix, all_mouse_ids, all_looms, sheet_name)
n_looms = length(all_looms);

mean_col = NaN(n_looms, 1);
for k = 1:n_looms
    row_vals = data_matrix(k,:);
    if any(~isnan(row_vals))
        mean_col(k) = mean(row_vals(~isnan(row_vals)));
    end
end

full_matrix = [data_matrix, mean_col];
data_cell   = num2cell(full_matrix);
for r = 1:size(data_cell,1)
    for cc = 1:size(data_cell,2)
        if isnan(data_cell{r,cc}), data_cell{r,cc} = ''; end
    end
end

row_labels  = arrayfun(@(k) sprintf('Loom%d', k), all_looms, 'UniformOutput', false);
col_headers = [all_mouse_ids, {'Mean'}];
header_row  = [{'Loom'}, col_headers];
output_cell = [header_row; [row_labels(:), data_cell]];

writecell(output_cell, output_file, 'Sheet', sheet_name);
end

%% ====================================================================
%% LOCAL DRAWING HELPERS
%% ====================================================================
function img = local_draw_circle(img, cx, cy, r, colour)
[H, W, ~] = size(img);
for y = max(1, cy-r) : min(H, cy+r)
    for x = max(1, cx-r) : min(W, cx+r)
        if (x-cx)^2 + (y-cy)^2 <= r^2
            img(y,x,1) = colour(1);
            img(y,x,2) = colour(2);
            img(y,x,3) = colour(3);
        end
    end
end
end

function img = local_draw_line(img, x1, y1, x2, y2, colour, thickness)
[H, W, ~] = size(img);
steps = max(abs(x2-x1), abs(y2-y1));
if steps == 0, return; end
xs   = linspace(x1, x2, steps+1);
ys   = linspace(y1, y2, steps+1);
half = floor(thickness/2);
for k = 1:length(xs)
    for ty = -half:half
        for tx = -half:half
            px = round(xs(k)) + tx;
            py = round(ys(k)) + ty;
            if px >= 1 && px <= W && py >= 1 && py <= H
                img(py,px,1) = colour(1);
                img(py,px,2) = colour(2);
                img(py,px,3) = colour(3);
            end
        end
    end
end
end

function img = local_insert_text(img, txt)
try
    img = insertText(img, [10 10], txt, ...
        'FontSize', 16, 'TextColor', 'white', ...
        'BoxColor', 'black', 'BoxOpacity', 0.5);
catch
    % No Computer Vision Toolbox -- skip silently
end
end
