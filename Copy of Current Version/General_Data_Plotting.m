
% Values 
X = [150, 100, 25];
Y1 = [1.21, .82, .38];
Y2 = [1.27, .86, .41];
Y3 = [2.24,1.58,0.52];

% Set up Plot
figure(1)
hold on
title("Binding Effect with Cortisol")
plot(X, Y1, "bo-", 'LineWidth', 2)
plot(X, Y2, "ko-", 'LineWidth', 2)
plot(X, Y3, "ro--", 'LineWidth', 2)
xlabel('Initial Cortisol Concentration (uM)');
ylabel('Absorbance at 247 nm');
grid on

% perform best fit analysis
fit1 = polyfit(X, Y1,2); % gives me my coefficients of the fitted polynomial equation
fit2 = polyfit(X, Y2,2); % gives me my coefficients of the fitted polynomial equation
fit3 = polyfit(X, Y3,2); % gives me my coefficients of the fitted polynomial equation
% NOTE: Fit gives you the variables as: ax + b = [a,b]
val1 = polyval(fit1, X); % gives me my new 
val2 = polyval(fit2, X); % gives me my new 
val3 = polyval(fit3, X); % gives me my new 

%plot my new best fit line
%plot(X1, val1, 'r--')
%plot(X2, val2, 'r--')

% Legend
legend("Old Solvent with Cortisol (2:8; AcetyNitrile:Toluene)","New Solvent with Cortisol (EthanolAmine)","Cortisol")
hold off


