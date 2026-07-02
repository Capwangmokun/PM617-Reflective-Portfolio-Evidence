% IT Study Workload Analyzer
% This program analyses weekly study workload for PM617 portfolio evidence.

clear; clc; close all;

% Weekly study data for six weeks
weeks = 1:6;

programming = [4 5 6 7 6 8];
reading = [3 4 3 5 4 4];
teamwork = [2 3 2 3 4 3];
videoProduction = [1 2 2 3 3 4];

% Combine all activity data into one matrix
workloadData = [programming; reading; teamwork; videoProduction]';

% Calculate weekly total workload
weeklyTotal = sum(workloadData, 2);

% Calculate summary statistics
totalHours = sum(weeklyTotal);
averageHours = mean(weeklyTotal);
[maxHours, maxWeekIndex] = max(weeklyTotal);
[minHours, minWeekIndex] = min(weeklyTotal);

% Display summary in command window
fprintf('IT Study Workload Analyzer\n');
fprintf('--------------------------\n');
fprintf('Total study hours: %.1f hours\n', totalHours);
fprintf('Average weekly study hours: %.1f hours\n', averageHours);
fprintf('Highest workload: Week %d with %.1f hours\n', maxWeekIndex, maxHours);
fprintf('Lowest workload: Week %d with %.1f hours\n', minWeekIndex, minHours);

% Conditional recommendation
if averageHours >= 18
    recommendation = 'Workload is high. Consider better time management.';
elseif averageHours >= 12
    recommendation = 'Workload is balanced. Continue the current routine.';
else
    recommendation = 'Workload is low. Increase regular study time.';
end

fprintf('Recommendation: %s\n', recommendation);

% Create stacked bar chart for workload types
figure;
bar(weeks, workloadData, 'stacked');
xlabel('Week');
ylabel('Study Hours');
title('Weekly IT Study Workload by Activity Type');
legend('Programming', 'Reading', 'Teamwork', 'Video Production', ...
       'Location', 'northwest');
grid on;

% Create line chart for total weekly workload
figure;
plot(weeks, weeklyTotal, '-o', 'LineWidth', 1.5);
xlabel('Week');
ylabel('Total Study Hours');
title('Total Weekly IT Study Workload');
xticks(weeks);
grid on;

% Save output figures
saveas(1, 'workload_stacked_chart.png');
saveas(2, 'weekly_total_progress.png');
