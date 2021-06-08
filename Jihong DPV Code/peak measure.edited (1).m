order = 4;
iterations = 15;

%% copy and paste the voltammogram's yaxis and run. 
%% corrected_voltammogram: voltammogram after baseline subtraction
%% baseline: baseline
%% pks: peak heights 
%% locs: location of peaks


raw_voltammogram =A;%A=copied y axis


x = 1:length(raw_voltammogram);
raw_voltammogram = transpose(raw_voltammogram);


%%voltammogram = movmean(raw_voltammogram,3);     %%with moving average filter
voltammogram = raw_voltammogram;                  %%without moving average filter

[baseline]=getbaseline(voltammogram,order,iterations);

corrected_voltammogram = voltammogram - baseline;


plot(x,voltammogram,'b-',x,baseline,'r--',x,corrected_voltammogram,'black');

legend('Initial voltammogram','Baseline','Corrected voltammogram');


corrected_voltammogram = transpose(corrected_voltammogram);
baseline = transpose(baseline);

[pks,locs] = findpeaks(corrected_voltammogram);
B=[locs,pks]
maxpeak=find(B==max(pks));
locs(maxpeak-length(pks))
B(maxpeak)
