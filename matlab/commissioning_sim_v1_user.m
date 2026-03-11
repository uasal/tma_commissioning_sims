function [m2_final_compen_array] = commissioning_sim_v1_user()

if ~exist('args', 'var')
    args = [];
end

% Initialize the OpticStudio connection
TheApplication = InitConnection();
if isempty(TheApplication)
    % failed to initialize a connection
    r = [];
else
    try
        [m2_final_compen_array] = BeginApplication(TheApplication);
        CleanupConnection(TheApplication);
    catch err
        CleanupConnection(TheApplication);
        rethrow(err);
    end
end
end

% Initial function to connect to ZOS API
function [m2_final_compen_array]  = BeginApplication(TheApplication)

% test for branch
import ZOSAPI.*;

% Load a sample file
% global TheSystem
TheSystem = TheApplication.PrimarySystem;
TheLDE = TheSystem.LDE;

[ZOS_file, ZOS_file_loc] = uigetfile('*.zos', "MultiSelect", "on"); % misalignment file
fileExtension = '*.zos';
fileList = dir(fullfile(ZOS_file_loc, fileExtension));
disp(fileList)

mc_data = [];
% 
z = 1; % Counter for file #
for z = 1:10 % create loop for going through Monte Carlo files


    close all % close all plots before generating new ones

    ZOS_file_name = fileList(z).name();

    loadFileString = strcat(ZOS_file_loc,'\',ZOS_file_name); % loading file 
     
    display(strcat('loading...',ZOS_file_name));
    TheSystem.LoadFile(System.String.Concat(loadFileString), false);

    %% calculate initial merit function before first step of alignment
    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_mc_2.MF');
    TheSystem.MFE.CalculateMeritFunction();
    init_merit_fun = TheSystem.MFE.GetOperandAt(20).ValueCell.DoubleValue();

    wave_pv = TheSystem.MFE.GetOperandAt(40).ValueCell.DoubleValue();
    init_wave_pv = wave_pv;

    writematrix([init_wave_pv; init_merit_fun], 'initial_metrics_' + string(z) + '.xlsx')


    %% define surfaces in Zemax model
    m2_compen = 15; % Chosing the M2 compensation surface in Zemax
    m2_compen_surf = TheLDE.GetSurfaceAt(m2_compen);
    m2_thickness_compen = 13;
    m2_thickness_compen_surf = TheLDE.GetSurfaceAt(m2_thickness_compen);    

    merit_cell = []; % Initialize a cell of merit function values
    %merit_cell(1) = init_merit_fun; % The first value of the cell will be the original merit function before misalignment
    merit_fun = init_merit_fun; % The first merit function in the loop will be the nominal merit function before misalignment
    
    randMatrix = zeros(1,6); % the first instance of randMatrix will be all zeroes, these are the 6 degrees of freedom that will be fed to the compensation surface

    % Defining M1
    m1_err_mat = zeros(1,5); % set an arbitrary array of zeros as the M1 errors

    m1_err = 7; % define your M1 error surface from Zemax
    m1_err_surf = TheLDE.GetSurfaceAt(m1_err);
    m1_thickness_compen_surf = TheLDE.GetSurfaceAt(5);

    m1_err_mat(1) = m1_err_surf.SurfaceData.Decenter_X_Cell.DoubleValue; % oull original M1 errors from the Zemax file to be read into array
    m1_err_mat(2) = m1_err_surf.SurfaceData.Decenter_Y_Cell.DoubleValue;
    m1_err_mat(4) = m1_err_surf.SurfaceData.TiltAbout_X_Cell.DoubleValue;
    m1_err_mat(5) = m1_err_surf.SurfaceData.TiltAbout_Y_Cell.DoubleValue;
    m1_err_mat(3) = m1_thickness_compen_surf.ThicknessCell.DoubleValue;

    
    %% Initialize arrays for model-based fitting
    
    % focus_fit_x = []; % Initialize array for the perturb Decenter Z value
    % focus_fit_y = []; % Initialize array for corresponding merit function value to the decenter Z value
    % 
    % decen_x_fit_x = [];
    % decen_x_fit_y = [];
    % 
    % decen_y_fit_x = [];
    % decen_y_fit_y = [];
    
    % new_merit_focus = [];
    % new_merit_focus(1) = TheSystem.MFE.CalculateMeritFunction();
    % 
    % Step_scale = 1; % Used in fitting to step and find best value for decenter x/y/z
    
    %% Create a loop to be able to find the local minima of polynomial fit 
    % The overall loop driving the convergence will indicate an iteration once
    % a local minima is found 

    if round(merit_fun, 2) <= 1 % if the initial value is less than 1 mm RMS spot size, then we will set the initial merit value as the value and move on
        final_merit_fun = init_merit_fun;
        merit_fun_spgd = init_merit_fun;
    end
    
    % initializing parameters before model-based correction
    dd = 1;
    Step_arr = [];
    dir_arr = [];
    dir_arr(1) = 1;
    Step_arr(1) = 1;
    merit_cell = [];
    merit_cell(1) = merit_fun;
    time = zeros(1, dd);
    
   while round(merit_fun,2) >= 1 % while loop for driving model-based algorithm

    Step_scale = Step_arr(dd); % step scale array
    ii = 1; % iteration counter
    good_ii = 1; % good iteration counter
    bad_ii = 1; % bad iteration counter
    direction = dir_arr(dd); % direction array
    new_merit_focus = [];
    new_merit_focus(1) = TheSystem.MFE.GetOperandAt(20).ValueCell.DoubleValue(); % grabbing RMS spot size value for first value of focus array
    focus_fit_x = [];
    focus_fit_y = [];
    tic; % timer

    while ii >= 1
         
        % Setting matrix to feed into the evaluation function and calculate
        % a new merit function based on misalignment
        %randMatrix(1) = 0;
        %randMatrix(2) = 0;
        randMatrix(3) = randMatrix(3) + direction*Step_scale; % changing decenter Z which is the third element of the m2 matrix
        randMatrix(4) = 0;
        randMatrix(5) = 0;
        randMatrix(6) = 0;

        % Function to use the random matrix generated and apply to the
        % corresponding degrees of freedom of the M2 compensation surface
        [new_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem); 

        merit_fun = new_merit_fun; % re-defining merit function values, i dont know why but it is

        new_merit_focus(ii + 1) = new_merit_fun; % Append the latest merit function value to the merit function array to look for trend

        if new_merit_focus(end) <= new_merit_focus(end - 1) % GOOD CASE IN DECENTER Z

            disp('Decenter Z improvement')
            good_ii = good_ii + 1;
            
            if good_ii >= 10 % if we have 10 good iterations
                disp('Linear point fit')
                [min_merit_z, min_z_idx] = min(focus_fit_y); % Determine the minimum point of the linear trend
                findZ = focus_fit_x(min_z_idx);
                randMatrix(3) = findZ; % set the Decenter Z value of M2 as the minimum of the linear trend
                break % get outta this loop
            elseif ii > 10 && good_ii < 10 % if there are more than 10 steps of M2 but less than 10 good steps, then we assume polynomial fit
                disp('Polynomial fit')
                focus_poly_fit = polyfit(focus_fit_x, focus_fit_y, 2); % fit a 2 deg. polynomial to find vertex
                findZ = (-focus_poly_fit(2))/ (2*focus_poly_fit(1));
                randMatrix(3) = findZ;
                break % get outta this loop
            end


        elseif new_merit_focus(end) > new_merit_focus(end - 1) % BAD CASE IN DECENTER Z
            
            disp('Decenter Z not improving')
            bad_ii = bad_ii + 1; % Add to the bad fit counter
            Step_scale = 0.5*Step_scale; % change step size and direction so we're not going in some crazy direction for a long time
            direction = -1*direction;

            % If there are more than 10 bad fits in a row, break into
            % this loop
            if good_ii >= 10
                disp('Linear point fit')
                [min_merit_z, min_z_idx] = min(focus_fit_y); % Determine the minimum point of the polynomial and its distance
                findZ = focus_fit_x(min_z_idx);
                randMatrix(3) = findZ;
                break
            elseif ii > 10 && good_ii < 10
                disp('Polynomial fit')
                focus_poly_fit = polyfit(focus_fit_x, focus_fit_y, 2);
                findZ = (-focus_poly_fit(2))/ (2*focus_poly_fit(1));
                randMatrix(3) = findZ;
                break
            end
        end

        focus_fit_x(ii) = randMatrix(3); % Append the random matrix value to the x array for the quadratic fitting
        focus_fit_y(ii) = merit_fun; % Append the merit function value to the y array for the quadratic fitting

        % Plot the current point
        figure(1);
        hold on;
        scatter(focus_fit_x(1:ii), focus_fit_y(1:ii), 'filled', 'r');
        xlabel('M2 Decenter Z Step (mm)')
        ylabel('Avg. RMS Spot Size Radius (mm)')
        grid on
        pause(0.001);

        ii = ii + 1; % counter for the polynomial fitting for finding optimal position of Decenter Z

    end

    [findZ_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem); % evaluate merit value after minimum position for this axis is determined

    if findZ_merit_fun < 1 % if the value is less than 1 mm then break out of the overall loop
        disp('z break')
        merit_fun_spgd = findZ_merit_fun;
        break
    end

    randMatrix(3) = findZ; % set decenter Z position in M2 for further model-based alignment
    
    %% Re-center beam on M3

    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_reaxy.MF') % Load the merit function editor 

    TheLDE = TheSystem.LDE;
    fieldbias_surf = TheLDE.GetSurfaceAt(1); % Field bias surface
    fieldbias_surf.SurfaceData.TiltAbout_X_Cell.MakeSolveVariable(); % set your tilt x/y parameters in your field bias to variable for optimization
    fieldbias_surf.SurfaceData.TiltAbout_Y_Cell.MakeSolveVariable();

    LocalOpt = TheSystem.Tools.OpenLocalOptimization();
    LocalOpt.RunAndWaitForCompletion();
    LocalOpt.Close();

    %% decenter y polynomial fit - same as done before but now doing this in the Decenter Y direction

    kk = 1;
    good_kk = 1;
    Step_scale = Step_arr(dd);
    direction = dir_arr(dd);
    new_decen_y_diff = [];
    new_decen_y_diff(1) = findZ_merit_fun;
    decen_y_fit_x = [];
    decen_y_fit_y = [];
    
    while kk >= 1

        %randMatrix(1) = 0;
        randMatrix(2) = randMatrix(2) + direction * Step_scale;
        randMatrix(3) = findZ;
        randMatrix(4) = 0;
        randMatrix(5) = 0;
        randMatrix(6) = 0;

        % Function to use the random matrix generated and apply to the
        % corresponding degrees of freedom of the M2 compensation surface
        [Y_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem); % Using the matrix calculated by the loop above, evaluate a new merit function to keep track of.
               
        new_decen_y_diff(kk + 1) = Y_merit_fun;

        if new_decen_y_diff(end) <= new_decen_y_diff(end-1)
            disp('Decenter Y improvement')
            good_kk = good_kk + 1;

            % If there are more than 10 bad fits in a row, break into
            % this loop
            if good_kk >= 10
                disp('Linear point fit')
                %y_poly_fit = polyfit(decen_y_fit_x, decen_y_fit_y, 2); 
                [min_merit_y, min_y_idx] = min(decen_y_fit_y); % Determine the minimum point of the polynomial and its distance
                findY = decen_y_fit_x(min_y_idx);
                randMatrix(2) = findY;
                break
            elseif kk > 10 && good_kk < 10
                disp('Polynomial fit')
                y_poly_fit = polyfit(decen_y_fit_x, decen_y_fit_y, 2);
                findY = (-y_poly_fit(2)) / (2*y_poly_fit(1));
                randMatrix(2) = findY;
                break
            end

        elseif new_decen_y_diff(end) > new_decen_y_diff(end-1)
            disp('Decenter Y not improving')
            direction = (-1)*direction;
            Step_scale = 0.5*Step_scale;

            % If there are more than 10 bad fits in a row, break into
            % this loop
            if good_kk >= 10
                disp('Linear point fit')
                %y_poly_fit = polyfit(decen_y_fit_x, decen_y_fit_y, 2); 
                [min_merit_y, min_y_idx] = min(decen_y_fit_y); % Determine the minimum point of the polynomial and its distance
                findY = decen_y_fit_x(min_y_idx);
                randMatrix(2) = findY;
                break
            elseif kk > 10 && good_kk < 10
                disp('Polynomial fit')
                y_poly_fit = polyfit(decen_y_fit_x, decen_y_fit_y, 2);
                findY = (-y_poly_fit(2)) / (2*y_poly_fit(1));
                randMatrix(2) = findY;
                break
            end

        end

        decen_y_fit_x(kk) = randMatrix(2); % Append the random matrix value to the x array for the quadratic fitting
        decen_y_fit_y(kk) = Y_merit_fun; % Append the merit function value to the y array for the quadratic fitting


        % Plot the current point
        figure(2);
        hold on;
        scatter(decen_y_fit_x(1:kk), decen_y_fit_y(1:kk), 'filled', 'r');
        xlabel('M2 Decenter Y Step (mm)')
        ylabel('Avg. RMS Spot Size Radius (mm)')
        grid on
        pause(0.001);

        kk = kk + 1; % counter for the polynomial fitting

    end

    [findY_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem);

    if findY_merit_fun < 1
        disp('y break')
        merit_fun_spgd = findY_merit_fun;
        break
    end

    randMatrix(2) = findY;

    %% load the merit function  - MERIT FUNCTION TO CALCULATE REAX AND REAY AT THE IMAGE PLANE
    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_reaxy.MF') % Load the merit function editor 

    TheLDE = TheSystem.LDE;
    fieldbias_surf = TheLDE.GetSurfaceAt(1); % Field bias surface
    fieldbias_surf.SurfaceData.TiltAbout_X_Cell.MakeSolveVariable(); % set your tilt x/y parameters in your field bias to variable for optimization
    fieldbias_surf.SurfaceData.TiltAbout_Y_Cell.MakeSolveVariable();

    LocalOpt = TheSystem.Tools.OpenLocalOptimization();
    LocalOpt.RunAndWaitForCompletion();
    LocalOpt.Close();
    

    %% decenter x polynomial fit - same as decenter Z and decenter Y calculation

    jj = 1;
    good_jj = 1;
    Step_scale = Step_arr(dd);
    direction = dir_arr(dd);
    new_decen_x_diff = [];
    new_decen_x_diff(1) = findY_merit_fun;
    decen_x_fit_x = [];
    decen_x_fit_y = [];
    
    while jj >= 1

        randMatrix(1) = randMatrix(1) + direction * Step_scale;
        randMatrix(2) = findY;
        randMatrix(3) = findZ;
        randMatrix(4) = 0;
        randMatrix(5) = 0;
        randMatrix(6) = 0;

        % Function to use the random matrix generated and apply to the
        % corresponding degrees of freedom of the M2 compensation surface
        [X_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem); % Using the matrix calculated by the loop above, evaluate a new merit function to keep track of.              
        
        new_decen_x_diff(jj + 1) = X_merit_fun;

        % x_pern = new_decen_x_diff(end-1) / new_decen_x_diff(end); % calculate the 
        % 
        % if x_pern < 0.5
        %     Step_scale = 2*Step_scale;
        %     direction = 1*direction;
        % else
        %     Step_scale = Step_scale + dd;
        % end

        if new_decen_x_diff(end) <= new_decen_x_diff(end-1)
            disp('Decenter X improvement')
            good_jj = good_jj + 1;

            % If there are more than 10 bad fits in a row, break into
            % this loop
            if good_jj >= 10
                disp('Linear point fit')
                [min_merit_x, min_x_idx] = min(decen_x_fit_y); % Determine the minimum point of the polynomial and its distance
                findX = decen_x_fit_x(min_x_idx);
                randMatrix(1) = findX;
                break
            elseif jj > 10 && good_jj < 10
                disp('Polynomial fit')
                x_poly_fit = polyfit(decen_x_fit_x, decen_x_fit_y, 2);
                findX = (-x_poly_fit(2)) / (2*x_poly_fit(1));
                randMatrix(1) = findX;
                break
            end
        else
            disp('Decenter X not improving')
            direction = (-1)*direction;
            %Step_scale = Step_scale;

            % If there are more than 10 bad fits in a row, break into
            % this loop
            if good_jj >= 10
                disp('Linear point fit')
                [min_merit_x, min_x_idx] = min(decen_x_fit_y); % Determine the minimum point of the polynomial and its distance
                findX = decen_x_fit_x(min_x_idx);
                randMatrix(1) = findX;
                break
            elseif jj > 10 && good_jj < 10
                disp('Polynomial fit')
                x_poly_fit = polyfit(decen_x_fit_x, decen_x_fit_y, 2);
                findX = (-x_poly_fit(2)) / (2*x_poly_fit(1));
                randMatrix(1) = findX;
                break
            end

        end

        decen_x_fit_x(jj) = randMatrix(1); % Append the random matrix value to the x array for the quadratic fitting
        decen_x_fit_y(jj) = X_merit_fun; % Append the merit function value to the y array for the quadratic fitting

        % Plot the current point
        figure(3);
        hold on;
        scatter(decen_x_fit_x(1:jj), decen_x_fit_y(1:jj), 'filled', 'r');
        xlabel('M2 Decenter X Step (mm)')
        ylabel('Avg. RMS Spot Size Radius (mm)')
        grid on
        pause(0.001);

        jj = jj + 1; % counter for the polynomial fitting

    end

    [findXY_merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem);
    
    if findXY_merit_fun < 1      
        disp('x break')
        merit_fun_spgd = findXY_merit_fun;
        break
    end

    % findXY_merit_fun = merit_fun;
    merit_cell(dd + 1) = findXY_merit_fun;


    Step_scale = 1; % reset step scale and direction for repeating this entire loop back at decenter Z
    direction = 1;

    merit_fun = findXY_merit_fun;

    randMatrix(1) = findX;

    %% load the merit function  - MERIT FUNCTION TO CALCULATE REAX AND REAY AT THE IMAGE PLANE
    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_reaxy.MF') % Load the merit function editor 

    TheLDE = TheSystem.LDE;
    fieldbias_surf = TheLDE.GetSurfaceAt(1); % Field bias surface
    fieldbias_surf.SurfaceData.TiltAbout_X_Cell.MakeSolveVariable(); % set your tilt x/y parameters in your field bias to variable for optimization
    fieldbias_surf.SurfaceData.TiltAbout_Y_Cell.MakeSolveVariable();

    LocalOpt = TheSystem.Tools.OpenLocalOptimization();
    LocalOpt.RunAndWaitForCompletion();
    LocalOpt.Close();

    % load the merit function  - MERIT FUNCTION TO CALCULATE REAX AND REAY AT THE IMAGE PLANE
    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_mc_2.MF');

    %%

    if merit_cell(end) >= merit_cell(end-1) % if the latest merit function is greater than the previous, then switch direction and increase step size for the next loop
        dir_arr(dd + 1) = (-1)*direction;
        Step_arr(dd + 1) = Step_scale * dd;
    else
        dir_arr(dd + 1) = (1)*direction;
        Step_arr(dd + 1) = 1;
    end

    %% If M1 & M2 Balancing is needed, then break into calculating M1/M2 balanced positions

    if abs(randMatrix(1)) >= 25
        disp('M2 Dx Range of Motion Exceeds 25 mm')
        [m1_bal_x, m1_bal_y, m1_ry, m1_rx, m2_bal_x, m2_bal_y, m2_ry, m2_rx] = m1_m2_balancing(randMatrix(1), randMatrix(2), randMatrix(3), m1_err_mat(1), m1_err_mat(2), m1_err_mat(3));

        randMatrix(1) = m2_bal_x;
        randMatrix(2) = m2_bal_y;
        randMatrix(4) = m2_rx;
        randMatrix(5) = m2_ry;

        m1_err_mat(1) = m1_bal_x;
        m1_err_mat(2) = m1_bal_y;
        m1_err_mat(4) = m1_rx;
        m1_err_mat(5) = m1_ry;

        break
    end

    if abs(randMatrix(2)) >= 25
        disp('M2 Dy Range of Motion Exceeds 25 mm')

        [m1_bal_x, m1_bal_y, m1_ry, m1_rx, m2_bal_x, m2_bal_y, m2_ry, m2_rx] = m1_m2_balancing(randMatrix(1), randMatrix(2), randMatrix(3), m1_err_mat(1), m1_err_mat(2), m1_err_mat(3));

        randMatrix(1) = m2_bal_x;
        randMatrix(2) = m2_bal_y;
        randMatrix(4) = m2_rx;
        randMatrix(5) = m2_ry;

        m1_err_mat(1) = m1_bal_x;
        m1_err_mat(2) = m1_bal_y;
        m1_err_mat(4) = m1_rx;
        m1_err_mat(5) = m1_ry;
        break
    end

    if abs(randMatrix(3)) >= 25
        disp('M2 Dz Range of Motion Exceeds 25 mm')
        [m1_bal_x, m1_bal_y, m1_ry, m1_rx, m2_bal_x, m2_bal_y, m2_ry, m2_rx] = m1_m2_balancing(randMatrix(1), randMatrix(2), randMatrix(3), m1_err_mat(1), m1_err_mat(2), m1_err_mat(3));

        randMatrix(1) = m2_bal_x;
        randMatrix(2) = m2_bal_y;
        randMatrix(4) = m2_rx;
        randMatrix(5) = m2_ry;

        m1_err_mat(1) = m1_bal_x;
        m1_err_mat(2) = m1_bal_y;
        m1_err_mat(4) = m1_rx;
        m1_err_mat(5) = m1_ry;
        break
    end

    %% set up for next loop of model based if needed

    time(dd) = toc;
    sum_t = sum(time);

    dd = dd + 1;

    close all

   end

    % Set all values in the M2 compensation array
    randMatrix(1) = m2_compen_surf.SurfaceData.Decenter_X_Cell.DoubleValue;
    randMatrix(2) = m2_compen_surf.SurfaceData.Decenter_Y_Cell.DoubleValue;
    randMatrix(4) = m2_compen_surf.SurfaceData.TiltAbout_X_Cell.DoubleValue;
    randMatrix(5) = m2_compen_surf.SurfaceData.TiltAbout_Y_Cell.DoubleValue;
    randMatrix(3) = m2_thickness_compen_surf.ThicknessCell.DoubleValue;
    m1_err_mat(1) = m1_err_surf.SurfaceData.Decenter_X_Cell.DoubleValue;
    m1_err_mat(2) = m1_err_surf.SurfaceData.Decenter_Y_Cell.DoubleValue;
    m1_err_mat(4) = m1_err_surf.SurfaceData.TiltAbout_X_Cell.DoubleValue;
    m1_err_mat(5) = m1_err_surf.SurfaceData.TiltAbout_Y_Cell.DoubleValue;
    m1_err_mat(3) = m1_thickness_compen_surf.ThicknessCell.DoubleValue;
    merit_fun_spgd = merit_fun;

    %% load the merit function  - MERIT FUNCTION TO CALCULATE REAX AND REAY on M3
    TheMFE = TheSystem.MFE;
    TheMFE.LoadMeritFunction('spgd_reaxy.MF') % Load the merit function editor 

    TheLDE = TheSystem.LDE;
    fieldbias_surf = TheLDE.GetSurfaceAt(1); % Field bias surface
    fieldbias_surf.SurfaceData.TiltAbout_X_Cell.MakeSolveVariable(); % set your tilt x/y parameters in your field bias to variable for optimization
    fieldbias_surf.SurfaceData.TiltAbout_Y_Cell.MakeSolveVariable();

    LocalOpt = TheSystem.Tools.OpenLocalOptimization();
    LocalOpt.RunAndWaitForCompletion();
    LocalOpt.Close();

    %% begin SPGD
   
    test_trial = spgd(merit_fun_spgd, randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem, merit_cell, z, wave_pv);

    mc_data_new = test_trial;
    mc_data_new = [mc_data; mc_data_new];
    dirLoc = pwd;
    fileNameSeq = System.String.Concat(dirLoc, '\final_alignment_' + string(z) + '.zos');
    TheSystem.SaveAs(fileNameSeq);
    writematrix(mc_data_new, 'spgd_table_' + string(z) + '.xlsx');
    
    z = z + 1;
    
    close all


end


end


%% function for evaluating surface in Zemax

function [merit_fun, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem)

% setting values in Zemax
m2_compen_surf.SurfaceData.Decenter_X_Cell.DoubleValue = randMatrix(1);
m2_compen_surf.SurfaceData.Decenter_Y_Cell.DoubleValue = randMatrix(2);
m2_compen_surf.SurfaceData.TiltAbout_X_Cell.DoubleValue = randMatrix(4);
m2_compen_surf.SurfaceData.TiltAbout_Y_Cell.DoubleValue = randMatrix(5);
m2_compen_surf.SurfaceData.TiltAbout_Z_Cell.DoubleValue = randMatrix(6);
m2_thickness_compen_surf.ThicknessCell.DoubleValue = randMatrix(3);


% load the merit function for calculating average RMS spot size
TheMFE = TheSystem.MFE;
TheMFE.LoadMeritFunction('spgd_mc_2.MF');

% define merit function value as output
TheSystem.MFE.CalculateMeritFunction();
merit_fun = TheSystem.MFE.GetOperandAt(20).ValueCell.DoubleValue();

wave_pv = TheSystem.MFE.GetOperandAt(40).ValueCell.DoubleValue();

wave_f1 = TheSystem.MFE.GetOperandAt(22).ValueCell.DoubleValue();
wave_f2 = TheSystem.MFE.GetOperandAt(23).ValueCell.DoubleValue();
wave_f3 = TheSystem.MFE.GetOperandAt(24).ValueCell.DoubleValue();
wave_f4 = TheSystem.MFE.GetOperandAt(25).ValueCell.DoubleValue();
wave_f5 = TheSystem.MFE.GetOperandAt(26).ValueCell.DoubleValue();
wave_f6 = TheSystem.MFE.GetOperandAt(27).ValueCell.DoubleValue();
wave_f7 = TheSystem.MFE.GetOperandAt(28).ValueCell.DoubleValue();
wave_f8 = TheSystem.MFE.GetOperandAt(29).ValueCell.DoubleValue();
wave_f9 = TheSystem.MFE.GetOperandAt(30).ValueCell.DoubleValue();
wave_f10 = TheSystem.MFE.GetOperandAt(31).ValueCell.DoubleValue();
wave_f11 = TheSystem.MFE.GetOperandAt(32).ValueCell.DoubleValue();
wave_f12 = TheSystem.MFE.GetOperandAt(33).ValueCell.DoubleValue();
wave_f13 = TheSystem.MFE.GetOperandAt(34).ValueCell.DoubleValue();
wave_f14 = TheSystem.MFE.GetOperandAt(35).ValueCell.DoubleValue();
wave_f15 = TheSystem.MFE.GetOperandAt(36).ValueCell.DoubleValue();

wave_cell = [wave_f1; wave_f2; wave_f3; wave_f4; wave_f5; wave_f6; wave_f7; wave_f8; wave_f9; wave_f10; wave_f11; wave_f12; wave_f13; wave_f14; wave_f15];

end


%% function to do SPGD after model-based fitting and correction

function [final_compen] = spgd(merit_fun, randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem, merit_cell, z, wave_pv)


GoodDoF = randMatrix; % initial 'good' trial will be first set of values corresponding to M2 position
BestDoF = GoodDoF; % setting initial 'best' trial as the good trial
% TheLDE = TheSystem.LDE;

merit_cell_spgd = []; % initialize merit function cell
merit_cell_spgd(1) = merit_fun; % initialize first value as initila meritv value
wave_cell = [];
wave_num_cell(1) = wave_pv;
% gain_cell = [];
% gain_cell(1) = merit_fun;

% Set initial parameters for SPGD
gain = merit_fun;
updateCount = 0;
HammerCount = 0;
B_HammerCount = 0;
HammerR = 1;
decimal_M = 7;
improveR = 0.2;
g = 1;
bad_flag = 0;
spgd_table = [];

good_values = [];

best_merit = merit_fun;
goodH = 1;

%% Set the initial perturbation matrix
perturbM = zeros(1,6);
% step_sign = 1;
step_size = 1;
r_scale = 1;
t_scale = 1;
t_scale_1 = 1;
r_scale_2 = 1;
speed = 0.5;
good_signs = sign(randMatrix(1:5));
good_signs(good_signs == 0) = 1;

% timing
t=1;
time = zeros(1,t);
T = [];
it_T = [];

%% Set up SPGD loop - want to loop through algorithm until merit value is less than or equal to 0.03 mm
while g <= 1000

    %store previous merit function value
    prev_merit = merit_cell_spgd(end);
    % % 
        
    % Toggle between different cases to perturb different degrees of freedom
    switch mod(floor(g/10), 3)
    
        case 0
        % Rotation Only
            t_scale = 0;
            t_scale_1 = 0;
            t_scale_2 = 0;
            r_scale = 1;
            r_scale_2 = 1;
            msg = 'Rx/Ry Only';

        case 2
            % Dx/Dy Only
            t_scale = 0;
            t_scale_1 = 1;
            t_scale_2 = 1;
            r_scale = 0;
            r_scale_2 = 0;
            msg = 'Dx/Dy Only';


        case 1
            % Dz Only
            t_scale = 1;
            t_scale_1 = 0;
            t_scale_2 = 0;
            r_scale = 0;
            r_scale_2 = 0;
            msg = 'Dz Only';

    end
        
    % Determining how the M2 compensation values should be calculated based
    % on whether or not there are bad flags or good flags
    if bad_flag == 0  % defined as good case, set the next M2 compensation matrix to equal the last defined improvement matrix plus some perturbation value
        if speed >= 0.9 && speed < 1 % if we hit a plateau, then incorporate signs of matrix in to try and break plateau
            randMatrix(1:5) = good_signs.*(abs(GoodDoF(1:5) + perturbM(1:5)));
        else
            randMatrix(1:5) = GoodDoF(1:5) + perturbM(1:5);
        end

    elseif bad_flag == 1  % defined as bad case, then find a new M2 compensation matrix by using random perturbation matrix and adding to the best defined degree of freedom
        if speed > 0.9 && speed <= 1
            randMatrix(1:5) = good_signs.*abs(BestDoF(1:5) + (step_size).*perturbM(1:5));
        else
            randMatrix(1:5) = BestDoF(1:5) + (step_size).*perturbM(1:5);
        end

   end

    [new_new_merit, wave_pv, wave_cell] = evalSurface(randMatrix, m2_compen_surf, m2_thickness_compen_surf, TheSystem); % calculate new merit value after changing M2 compensation values
    merit_fun = new_new_merit; % again, redefining variables idk

    if merit_fun == 0 % if merit value equals zero, then we assume that the system has vignetted and will re-run the merit function to center beam on M3
        TheMFE = TheSystem.MFE;
        TheMFE.LoadMeritFunction('spgd_reaxy.MF') % Load the merit function editor 
    
        TheLDE = TheSystem.LDE;
        fieldbias_surf = TheLDE.GetSurfaceAt(1); % Field bias surface
        fieldbias_surf.SurfaceData.TiltAbout_X_Cell.MakeSolveVariable(); % set your tilt x/y parameters in your field bias to variable for optimization
        fieldbias_surf.SurfaceData.TiltAbout_Y_Cell.MakeSolveVariable();
    
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();

        TheMFE.LoadMeritFunction('spgd_mc_2.MF'); % re-load merit function for calculating rms spot size

        % Calculating the overall merit function
        TheSystem.MFE.CalculateMeritFunction();
        merit_fun = TheSystem.MFE.GetOperandAt(20).ValueCell.DoubleValue();
    end

    merit_cell_spgd(g + 1) = merit_fun; % Append latest calculated merit function value to cell
    wave_num_cell(g + 1) = wave_pv;
    speed = merit_cell_spgd(end) / merit_cell_spgd(end-1); % calculate the speed by finding the ratio between latest calculated merit value and the current


    if round(merit_cell_spgd(end), decimal_M) >= round(best_merit, decimal_M) % BAD CASE 

        disp('System is not improving')
    
        perturbM = 2*rand(1,6) - 1; % set a new set of pertubation values for the next iteration

        bad_flag = 1;
        %gain_cell(g + 1) = gain;
        step_size = 0.5;

        % with a new pertubation array defined, multiply by new gain values
        % based on degrees of freedom. if we see a plateau, then increase
        % gain
        if speed >= 0.9 && speed < 1 
            % DecGain = 1.2025*((0.03)^3) - 1.3584*((0.03)^2) + 0.9534*(0.03) + 0.3534;
            % ThickGain = 1.2025*((0.03)^3) - 1.3584*((0.03)^2) + 0.9534*(0.03) + 0.3534;
            % % ThickGain = 9.4365*(0.03) + 0.1957;
            % RotGain2 = 0.0226*((0.03)^3) - 0.4153*((0.03)^2) + 1.3987*(0.03) + 0.2183;
            % % RotGain1 = RotGain2;
            % RotGain1 = -0.0501*((0.03)^3) + 0.5401*((0.03)^2) - 0.4716*(0.03) + 4.1099;
            % % % 
            DecGain = 1;
            ThickGain = 1;
            RotGain2 = 0.05;
            RotGain1 = 0.05;
        else
            DecGain = 0.1;
            ThickGain = 0.1;
            RotGain2 = 0.005;
            RotGain1 = 0.005;
        end

        % define new perturbation values
        perturbM(1) = perturbM(1) * DecGain * t_scale_1;
        perturbM(2) = perturbM(2) * DecGain* t_scale_2;
        perturbM(3) = perturbM(3) * ThickGain* t_scale;
        perturbM(4) = perturbM(4) * RotGain1 * r_scale;
        perturbM(5) = perturbM(5) * RotGain2 * r_scale_2;
        
                

    else% if the latest calculated merit function is less than previously calculated, then the system is improving

        disp('System improvement')
        GoodDoF = randMatrix; % set good degree of freedom


        best_merit = min(merit_cell_spgd); % define best merit function value

        if merit_fun == best_merit % if the current merit function value is equal to the best then we store the current M2 position as the "best dof"
            BestDoF = randMatrix; % Store as the best configuration
            disp('best merit fun')
            good_signs = sign(BestDoF(1:5));
            good_signs(good_signs == 0) = 1;
        end

        bad_flag = 0; % set bad flag for next iteration
        goodH = goodH + 1; % update good hit counter

        if goodH < 10 % if the good hit counter is less than 10, then multiply the perturbation matrix by a tenth of the hit counter
            perturbM = perturbM ;
        elseif goodH >= 10 % let the good hit counter reach 10 then reset to 1
            perturbM = perturbM .* 3;
            goodH = 1;
        end

    end

    randMatrix(1:5) % print M2 position to terminal to make sure it is changing sufficiently


    %% end plotting code

    spgd_data = [merit_cell_spgd(end); g; wave_pv(end); randMatrix(1); randMatrix(2); randMatrix(3); randMatrix(4); randMatrix(5); perturbM(1); perturbM(2); perturbM(3); perturbM(4); perturbM(5); wave_cell];
    
    t = t + 1;
    time(t) = toc;
    sum_t = sum(time);
    
    g = g + 1; % update counter
 
    % define the figure
    % define the figure
    figure(4)
    title({'Merit Value vs. Iteration', ['Latest RMS Spot Size Radius (mm): ', num2str(round(merit_cell_spgd(end),2))], ['Latest Avg. WFE PV (waves): ', num2str(wave_pv ,2)], msg});
    grid on

    hold on
    yyaxis left
    xlabel('Iteration')
    ylabel('RMS Spot Size (mm) (log scaled)')
    semilogy(1:g, merit_cell_spgd(1:g), 'b-+')
    % ylim([0 0.1])
    xlim([ 1 length(merit_cell_spgd)])

    yyaxis right
    ylabel('Wavefront PV Error (waves)')
    semilogy(1:g, wave_num_cell(1:g), '--.')
    xlim([1 length(wave_num_cell)])
    hold off
    spgd_table = [spgd_table, spgd_data];

    hold off

 
end

final_compen = spgd_table;

%writematrix(spgd_table, 'spgd_table_0211.xlsx');

end

%% M1 & M2 balancing function

function [m1_bal_x, m1_bal_y, m1_ry, m1_rx, m2_bal_x, m2_bal_y, m2_ry, m2_rx] = m1_m2_balancing(m2_dx, m2_dy, m2_dz, m1_dx, m1_dy, m1_dz)

%%%%%%%%%
% Purpose: Calculate the needed pivot point to balance M1 and M2 about. 
%
% Input: Remaining possible M1 movement in Decenter X/Y/Z and Rotation
% X/Y/Z. The calculated amount that M2 will need to move in Decenter X/Y/Z
% and Rotation X/Y/Z from the model-based fitting algorithm. 
%
% Output: Pivot coordinate to set for the balance coordinate calculation. 
%%%%%%%%%%

%% calculating remaining range of motion for M1 and M2
if m1_dy < 0
    m1_dy_rem = 5.6 + m1_dy;
else
    m1_dy_rem = 5.6 - m1_dy;
end

if m2_dy < 0
    m2_dy_rem = 25 + m1_dy;
else
    m2_dy_rem = 25 - m1_dy;
end

m1_dx_rem = 5.6 - m1_dx; % Subtract the current value in Dx, Dy and Dz from 5.6 mm to get remaining movement in M1. 
%m1_dy_rem = 5.6 - m1_dy;
m1_dz_rem = 5.6 - m1_dz;

m2_dx_rem = 25 - m2_dx; 
%m2_dy_rem = 25 - m2_dy; 
m2_dz_rem = 25 - m2_dz;

% Calculate ratio of movement between M1 and M2
ratio_x = m1_dx_rem / m2_dx_rem; 
ratio_y = m1_dy_rem / m2_dy_rem;
ratio_z = m1_dz_rem / m2_dz_rem;

% Determining when ratios are positive/negative
if ratio_x < 0 
    a_x = 7400 * ratio_x;
    r_x_m1 = 7400 + a_x;
else
    r_x_m2 = 7400 * ratio_x;
    r_x_m1 = -r_x_m2;
    a_x = -7400 + r_x_m1;
end

if ratio_y < 0 
    a_y = 7400 * ratio_y;
    r_y_m1 = 7400 + a_y;
else
    r_y_m2 = 7400 * ratio_y;
    r_y_m1 = -r_y_m2;
    a_y = -7400 + r_y_m1;
end

if ratio_z < 0 
    a_z = 7400 * ratio_z;
    r_z_m1 = 7400 + a_z;
else
    r_z_m2 = 7400 * ratio_z;
    r_z_m1 = -r_z_m2;
    a_z = -7400 + r_z_m1;
end

% % distance from M2 to pivot point
% a_x = 7400 * ratio_x;
% a_y = 7400 * ratio_y;
% a_z = 7400 * ratio_z;
% 
% % distance from M1 to pivot point
% r_x_m1 = 7400 - a_x;
% r_y_m1 = 7400 - a_y;
% r_z_m1 = 7400 - a_z;

p_x = r_x_m1; % define the pivot point in relation to the M1 vertex % take away absolute value in case pivot needs to be outside m1/m2
p_y = r_y_m1;
p_z = r_z_m1;


m2_a_x = a_x; % calculate the adjacent leg 

theta_y = atan(m2_dx / m2_a_x); % calculate the theta relative to M2

m1_dx = tan(theta_y) * p_x; % calculating corresponding M1 translation in x
m1_dy = 0;
m2_dy = 0;
m2_dz = 0;
m1_dz = 0;

m2_dx = m1_dx * ratio_x;
m1_ry = atan(m1_dx / p_x);
m2_ry = atan(m2_dx / m2_a_x);

m1_bal_x = m1_dx;
m2_bal_x = m2_dx;

% m1_dx_rem_new = 5.6 - m1_dx;
% m2_dx_rem_new = 25 - m2_dx;
% 
% new_ratio_x = m1_dx_rem_new / abs(m2_dx_rem_new);
% 
% m2_dist = 7400 * new_ratio_x;
% m1_dist = 7400 - m2_dist;
% 
% % re-calculate a new pivot point
% m2_ry = atan(m2_dx / m2_dist);
% 
% 
% m1_bal_x = tan(m2_ry) * m1_dist;
% m2_bal_x = tan(m2_ry) * m2_dist;
% m1_ry = m2_ry;


m2_a_y = a_y; % calculate the adjacent leg 

theta_x = atan(m2_dy / m2_a_y); % calculate the theta relative to M2

m1_dy = tan(theta_x) * p_y; % calculating corresponding M1 translation in x

m2_dy = m1_dy / ratio_y;
m1_rx = atan(m1_dy / p_y);
m2_rx = atan(m2_dy / m2_a_y);

m1_bal_y = m1_dy;
m2_bal_y = m2_dy;
% 
fprintf('M1 Balance Value in Dx: %f \n', m1_bal_x)
fprintf('M2 Balance Value in Dx: %f \n', m2_bal_x)
fprintf('M1 Balance in Ry: %f \n', m1_ry)
fprintf('M2 Balance in Ry: %f \n', m2_ry)
fprintf('M1 Balance Value in Dy: %f \n', m1_bal_y)
fprintf('M2 Balance Value in Dy: %f \n', m2_bal_y)
fprintf('M1 Balance in Rx: %f \n', m1_rx)
fprintf('M2 Balance in Rx: %f \n', m2_rx)

end

%% Initializing connection to Zemax 

function app = InitConnection()

import System.Reflection.*;

% Find the installed version of OpticStudio.
zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');
NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');
% Note -- uncomment the following line to use a custom NetHelper path
% NetHelper = 'C:\Users\Heejoo\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll';
% This is the path to OpticStudio
NET.addAssembly(NetHelper);

success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
% Note -- uncomment the following line to use a custom initialization path
% success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize('C:\Program Files\OpticStudio\');
if success == 1
    LogMessage(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
else
    app = [];
    return;
end

% Now load the ZOS-API assemblies
NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
NET.addAssembly(AssemblyName('ZOSAPI'));

% Create the initial connection class
TheConnection = ZOSAPI.ZOSAPI_Connection();

% Attempt to create a Standalone connection

% NOTE - if this fails with a message like 'Unable to load one or more of
% the requested types', it is usually caused by try to connect to a 32-bit
% version of OpticStudio from a 64-bit version of MATLAB (or vice-versa).
% This is an issue with how MATLAB interfaces with .NET, and the only
% current workaround is to use 32- or 64-bit versions of both applications.
app = TheConnection.CreateNewApplication();
if isempty(app)
    HandleError('An unknown connection error occurred!');
end
if ~app.IsValidLicenseForAPI
    HandleError('License check failed!');
    app = [];
end

end

function LogMessage(msg)
disp(msg);
end

function HandleError(error)
ME = MException('zosapi:HandleError', error);
throw(ME);
end

function  CleanupConnection(TheApplication)
% Note - this will close down the connection.

% If you want to keep the application open, you should skip this step
% and store the instance somewhere instead.
TheApplication.CloseApplication();
end





