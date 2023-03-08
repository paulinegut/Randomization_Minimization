%% ########################################################################
% 
% Randomization by Minimization (Pockock and Simon, Taves)
%
% Version 15. March 2021
% Author: Pauline Rose Gut
% #########################################################################


clear; close all; clc;

% Choose the excel file containing the previous allocated patients to the trial
% Each row is a patient
% The first column corresponds to the patient ID
% Each other column corresponds to a variate to be balanced over the treatment groups 
disp('Choose the excel file containing the previous allocated patients to the trial.')
disp('Each row must be a patient. The first column corresponds to the patient ID.')
disp('Each other column corresponds to a variate to be balanced over the treatment groups.')
[Filename,Pathname]=uigetfile('*.xlsx','Allocated patients to the trial');
cd(Pathname)
Allocated_Patients = readtable(Filename);

n = size(Allocated_Patients,1);
clc;
New_Patient.ID = input('New patient ID: ', 's');
New_Patient.SixMWT = input('6MWT of the new patient (in meters): ');
New_Patient.Age = input('Age of the new patient: ');

if n == 0
p = size(Allocated_Patients,1) + 1;
PatientID{1,1} = New_Patient.ID;
Patient6MWT(1,1) = New_Patient.SixMWT;
PatientAge(1,1) = New_Patient.Age;
Allocated_Patients(p,1) = table(PatientID(1,1));
Allocated_Patients(p,2) = table(Patient6MWT(1,1));
Allocated_Patients(p,3) = table(PatientAge(1,1));

% Treatment allocation to the patient through simple randomization
Treat = randi(2,1,1);
SWT{1,1} = 'SWT';
DOT{1,1} = 'DOT';
    if Treat == 1
        Allocated_Patients(p,4) = DOT;
        disp('Patient allocated to DOT')
    else
        Allocated_Patients(p,4) = SWT;
        disp('Patient allocated to SWT')
    end
    
Allocated_Patients.Properties.VariableNames =  {'PatientID', 'x6MWT', 'Age', 'Treatment'};
writetable(Allocated_Patients, Filename, 'WriteVariableNames', true)
disp('Finished!')

else
% Treatment allocation to the patient through minimization
% Calculate the mean 6MWT and the mean age of the patients previsouly
% allocated to the trial
Mean_SixMWT = mean(cell2mat(table2cell(Allocated_Patients(:,2))));
Mean_Age = mean(cell2mat(table2cell(Allocated_Patients(:,3))));

% Computation of the amount of variation among treatments
Smaller_Mean_SixMWT = 0;
Equal_Mean_SixMWT = 0;
Larger_Mean_SixMWT = 0;
Smaller_Mean_Age = 0;
Equal_Mean_Age = 0;
Larger_Mean_Age = 0;

for i = 1:size(Allocated_Patients,1)
if strcmp(table2cell(Allocated_Patients(i,4)), 'DOT')
    Smaller_Mean_SixMWT = Smaller_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) < Mean_SixMWT);
    Equal_Mean_SixMWT = Equal_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) == Mean_SixMWT);
    Larger_Mean_SixMWT = Larger_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) > Mean_SixMWT);

    Smaller_Mean_Age = Smaller_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) < Mean_Age);
    Equal_Mean_Age = Equal_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) == Mean_Age);
    Larger_Mean_Age = Larger_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) > Mean_Age);
    T(1,:) = table(Smaller_Mean_SixMWT, Equal_Mean_SixMWT, Larger_Mean_SixMWT, Smaller_Mean_Age, Equal_Mean_Age, Larger_Mean_Age);
    
end
end

clear Smaller_Mean_SixMWT Equal_Mean_SixMWT Larger_Mean_SixMWT Smaller_Mean_Age Equal_Mean_Age Larger_Mean_Age

Smaller_Mean_SixMWT = 0;
Equal_Mean_SixMWT = 0;
Larger_Mean_SixMWT = 0;
Smaller_Mean_Age = 0;
Equal_Mean_Age = 0;
Larger_Mean_Age = 0;

for i = 1:size(Allocated_Patients,1)
if strcmp(table2cell(Allocated_Patients(i,4)), 'SWT')
    Smaller_Mean_SixMWT = Smaller_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) < Mean_SixMWT);
    Equal_Mean_SixMWT = Equal_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) == Mean_SixMWT);
    Larger_Mean_SixMWT = Larger_Mean_SixMWT + sum(cell2mat(table2cell(Allocated_Patients(i,2))) > Mean_SixMWT);

    Smaller_Mean_Age = Smaller_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) < Mean_Age);
    Equal_Mean_Age = Equal_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) == Mean_Age);
    Larger_Mean_Age = Larger_Mean_Age + sum(cell2mat(table2cell(Allocated_Patients(i,3))) > Mean_Age);
    T(2,:) = table(Smaller_Mean_SixMWT, Equal_Mean_SixMWT, Larger_Mean_SixMWT, Smaller_Mean_Age, Equal_Mean_Age, Larger_Mean_Age);
    
end
end

clear Smaller_Mean_SixMWT Equal_Mean_SixMWT Larger_Mean_SixMWT Smaller_Mean_Age Equal_Mean_Age Larger_Mean_Age



if New_Patient.SixMWT < Mean_SixMWT
    T(3,1) = cell2table(num2cell(1));
    T(4,1) = cell2table(num2cell(sum(T{[1 3],1}, 1)));
    T(5,1) = T(2,1);
    T(6,1) = cell2table(num2cell(abs(diff(T{[4 5],1}, 1))));
    T(7,1) = T(1,1);
    T(8,1) = cell2table(num2cell(sum(T{[2 3],1}, 1)));
    T(9,1) = cell2table(num2cell(abs(diff(T{[7 8], 1}, 1))));
    
elseif New_Patient.SixMWT == Mean_SixMWT
    T(3,2) = cell2table(num2cell(1));
    T(4,2) = cell2table(num2cell(sum(T{[1 3],2}, 1)));
    T(5,2) = T(2,2);
    T(6,2) = cell2table(num2cell(abs(diff(T{[4 5],2}, 1))));
    T(7,2) = T(1,2);
    T(8,2) = cell2table(num2cell(sum(T{[2 3],2}, 1)));
    T(9,2) = cell2table(num2cell(abs(diff(T{[7 8], 2}, 1))));
    
elseif New_Patient.SixMWT > Mean_SixMWT
    T(3,3) = cell2table(num2cell(1));
    T(4,3) = cell2table(num2cell(sum(T{[1 3],3}, 1)));
    T(5,3) = T(2,3);
    T(6,3) = cell2table(num2cell(abs(diff(T{[4 5],3}, 1))));
    T(7,3) = T(1,3);
    T(8,3) = cell2table(num2cell(sum(T{[2 3],3}, 1)));
    T(9,3) = cell2table(num2cell(abs(diff(T{[7 8], 3}, 1))));
end

if New_Patient.Age < Mean_Age
    T(3,4) = cell2table(num2cell(1));
    T(4,4) = cell2table(num2cell(sum(T{[1 3],4}, 1)));
    T(5,4) = T(2,4);
    T(6,4) = cell2table(num2cell(abs(diff(T{[4 5],4}, 1))));
    T(7,4) = T(1,4);
    T(8,4) = cell2table(num2cell(sum(T{[2 3],4}, 1)));
    T(9,4) = cell2table(num2cell(abs(diff(T{[7 8], 4}, 1))));
    
elseif New_Patient.Age == Mean_Age
    T(3,5) = cell2table(num2cell(1));
    T(4,5) = cell2table(num2cell(sum(T{[1 3],5}, 1)));
    T(5,5) = T(2,5);
    T(6,5) = cell2table(num2cell(abs(diff(T{[4 5],5}, 1))));
    T(7,5) = T(1,5);
    T(8,5) = cell2table(num2cell(sum(T{[2 3],5}, 1)));
    T(9,5) = cell2table(num2cell(abs(diff(T{[7 8], 5}, 1))));
    
elseif New_Patient.Age > Mean_Age
    T(3,6) = cell2table(num2cell(1));
    T(4,6) = cell2table(num2cell(sum(T{[1 3],6}, 1)));
    T(5,6) = T(2,6);
    T(6,5) = cell2table(num2cell(abs(diff(T{[4 5],6}, 1))));
    T(7,6) = T(1,6);
    T(8,6) = cell2table(num2cell(sum(T{[2 3],6}, 1)));
    T(9,6) = cell2table(num2cell(abs(diff(T{[7 8], 6}, 1))));
end

T(6,7) = cell2table(num2cell(sum(T{6,:}, 2)));
T(9,7) = cell2table(num2cell(sum(T{9,:}, 2)));


T.Properties.RowNames =  {'DOT group', 'SWT group', 'New patient', 'DOT + new patient', 'SWT', 'Absolute difference 1', 'DOT', 'SWT + new patient', 'Absolute difference 2'};

% Treatment allocation to the patient
p = size(Allocated_Patients,1) + 1;
PatientID{1,1} = New_Patient.ID;
Patient6MWT(1,1) = New_Patient.SixMWT;
PatientAge(1,1) = New_Patient.Age;
Allocated_Patients(p,1) = table(PatientID(1,1));
Allocated_Patients(p,2) = table(Patient6MWT(1,1));
Allocated_Patients(p,3) = table(PatientAge(1,1));

SWT{1,1} = 'SWT';
DOT{1,1} = 'DOT';

if cell2mat(table2cell(T(6,7))) > cell2mat(table2cell(T(9,7)))
    Allocated_Patients(p,4) = SWT;
    disp('Patient allocated to SWT')
    
elseif cell2mat(table2cell(T(6,7))) == cell2mat(table2cell(T(9,7)))
    disp('Simple randomization')
    Treat = randi(2,1,1);
    if Treat == 1
        Allocated_Patients(p,4) = DOT;
        disp('Patient allocated to DOT')
    else
        Allocated_Patients(p,4) = SWT;
        disp('Patient allocated to SWT')
    end
    
elseif cell2mat(table2cell(T(6,7))) < cell2mat(table2cell(T(9,7)))
    Allocated_Patients(p,4) = DOT;
    disp('Patient allocated to DOT')
end


S{1,1} = 'Number of patients in DOT: ';
S{2,1} = 'Number of patients in SWT: ';
S{3,1} = 'Mean 6MWT in DOT: ';
S{4,1} = 'Mean 6MWT in SWT: ';
S{5,1} = 'Mean age in DOT: ';
S{6,1} = 'Mean age in SWT: ';

S{1,2} = sum(strcmp(table2cell(Allocated_Patients(:,4)), 'DOT'),1);
S{2,2} = sum(strcmp(table2cell(Allocated_Patients(:,4)), 'SWT'),1);

for i = 1:size(Allocated_Patients,1)
    if strcmp(table2cell(Allocated_Patients(i,4)), 'DOT')
        Mean_DOT1(i,1) = cell2mat(table2cell(Allocated_Patients(i,2)));
        Mean_DOT2(i,1) = cell2mat(table2cell(Allocated_Patients(i,3)));
    elseif strcmp(table2cell(Allocated_Patients(i,4)), 'SWT')
        Mean_SWT1(i,1) = cell2mat(table2cell(Allocated_Patients(i,2)));
        Mean_SWT2(i,1) = cell2mat(table2cell(Allocated_Patients(i,3)));
    end
end

for i = 1:size(Allocated_Patients,1)
    if ~strcmp(table2cell(Allocated_Patients(i,4)), 'DOT')
        Mean_DOT1(i,1) = 0;
        Mean_DOT2(i,1) = 0;
    elseif ~strcmp(table2cell(Allocated_Patients(i,4)), 'SWT')
        Mean_SWT1(i,1) = 0;
        Mean_SWT2(i,1) = 0;
    end
end

Mean_SixMWT_DOT = mean(nonzeros(Mean_DOT1));
Mean_Age_DOT = mean(nonzeros(Mean_DOT2));
Mean_SixMWT_SWT = mean(nonzeros(Mean_SWT1));
Mean_Age_SWT = mean(nonzeros(Mean_SWT2));


S{3,2} = round(Mean_SixMWT_DOT);
S{4,2} = round(Mean_SixMWT_SWT);
S{5,2} = round(Mean_Age_DOT);
S{6,2} = round(Mean_Age_SWT);

Population_Stat = table(S);
T.Properties.VariableNames([7]) = {'Sum'};

writetable(Allocated_Patients, Filename, 'WriteVariableNames', true)

if strcmp(Filename, 'Randomization_SCI.xlsx')
    writetable(Population_Stat, 'TrialPopStat_SCI.xlsx', 'WriteVariableNames', false)
elseif strcmp(Filename, 'Randomization_MS.xlsx')
    writetable(Population_Stat, 'TrialPopStat_MS.xlsx', 'WriteVariableNames', false)
end

disp('Finished!')
% clearvars -except Allocated_Patients Population_Stat T
end
% system(['set PATH=' Pathname ' && ' Filename]);

