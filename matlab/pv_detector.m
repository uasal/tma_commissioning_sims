%% generating wavefront error PV plots 

% Define the range and number of points for the new grid
num_points = 1000; % Adjust for desired resolution
xq = linspace(min(sheets.FieldX), max(sheets.FieldX), num_points);
yq = linspace(min(sheets.FieldY), max(sheets.FieldY), num_points);

% Create the 2D grid
[X, Y] = meshgrid(xq, yq);

Z = griddata(sheets.FieldX, sheets.FieldY, sheets.after, X, Y, 'cubic');
Z_before = griddata(sheets.FieldX, sheets.FieldY, sheets.before, X, Y, 'cubic');

figure(1);
contourf(X, Y, Z, 'LineStyle', 'None'); % Use contourf for filled contours
a = colorbar; % Add a color bar to indicate z-values
a.Label.String = 'PV WFE (waves)';
%clim([-10 10])
xlabel('X Detector Coordinates (deg.)');
ylabel('Y Detector Coordinates (deg.)');
title({'Peak-to-Valley Wavefront Error Across Detector', 'After Commissioning'});

figure(2);
contourf(X, Y, Z_before, 'LineStyle', 'None'); % Use contourf for filled contours
a = colorbar; % Add a color bar to indicate z-values
a.Label.String = 'PV WFE (waves)';
%clim([-10 10])
xlabel('X Detector Coordinates (deg.)');
ylabel('Y Detector Coordinates (deg.)');
title({'Peak-to-Valley Wavefront Error Across Detector', 'Before Commissioning'});