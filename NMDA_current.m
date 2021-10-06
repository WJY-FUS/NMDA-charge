% Clarification
% This code is created by Yu Mingcan from He lab in the Shanghai Institute of Organic Chemistry,CAS
% The former version(2019) was used to the published paper(https://doi.org/10.1016/j.neuron.2019.11.011)
% The latest version(2021.08) is used to this project

% Code Function
% This code is used to auto batch analyze the charge of EPSC 

% Data
% Data were extracted from the abf file and stored as the excel file(xlsx file)

% Principle
% Step 1: Some fragment EPSC data were extracted from the rawdata
% Step 2: Split the fragment EPSC data into some sections and adjust the fragment EPSC curve by each baseline of each sections
% Step 3: Sum up the adjusted EPSC data as the charge

% Protocol
% We offered the test data for this code and you can follow the steps and have a try.
% Step 1: Put all xlsx files in a folder
% Step 2: Run the function
% Step 3: After running, the results output as a xlsx file


% Validation
% Most of results analyzed by this code have been verified by manual analysis in our study project.

% Contribution
% Wang Junying from Wang lab and He lab in the Shanghai Institute of Organic Chemistry offered the rawdata and xlsx files.
% Yu Mingcan from the Shanghai Institute of Organic Chemistry, CAS created the function.
% Wang Junying finished the verification by manual analysis.

clear
% batch load abf file
file_info=dir(fullfile('*.xlsx'));        
struct_file_info=struct2cell(file_info);  
file_name=struct_file_info(1,:);        
[~,nfile]=size(file_name);   
load_count=0; 
for ifile=1:1:nfile           
    if strfind(file_name{ifile},'.xlsx')    
        load_count=load_count+1;
        [All_file{load_count}]=xlsread(file_name{ifile}); 
    end
end
% calculate NMDAR current in each xlsx file
for ixlsx=1:1:nfile; 
Each_xlsx_data=All_file{1,ixlsx}; % extract data from each xlsx file
% Caculate RMSdata to qaulify for the NMDAR-EPSC data
RMSdata=Each_xlsx_data(:,10);
RMSdata=RMSdata(~isnan(RMSdata));
% meanRMSBSL=mode(RMSdata);
% [~,~,RMSBSL]=mode(RMSdata);
% meanRMSBSL=mean(cell2mat(RMSBSL));
meanRMSBSL=median(RMSdata);
RMSAmp=RMSdata-meanRMSBSL;
RMSsquare=arrayfun(@(x)x*x,RMSAmp);
RMS{1,ixlsx}=sqrt(mean(RMSsquare));

% Extract NMDAR-EPSC data from xlsx file
for icolumn=2:2:8
EPSC_data=Each_xlsx_data(:,icolumn);
delete_NaN_data=EPSC_data(~isnan(EPSC_data));
[nrow,~]=size(delete_NaN_data);

% adjust the EPSC curve based on each baseline of each section
adjust_count=1;
for adjust_start=1:5000:5000*fix(nrow/5000);
Each_split_data=delete_NaN_data(adjust_start:adjust_start+4999,:); % split EPSC data into some 500ms sections
% caculate the mode of each section as the baseline of each section

meanBSL=median(Each_split_data);
% meanBSL=mode(Each_split_data);
% [~,~,BSL]=mode(Each_split_data);
% meanBSL=mean(cell2mat(BSL));

% adjust the EPSC curve based on the baseline
Each_adjust_data=bsxfun(@minus,Each_split_data,meanBSL.');

% caculate the charge of each section of each fragment
Each_charge=sum(Each_adjust_data);
fivehudmscharge{icolumn/2,adjust_count,ixlsx}=Each_charge;
adjust_count=adjust_count+1;
end
end

% caculate the charge of each fragment
FragmentMatrix=cellfun(@sum,fivehudmscharge);
FragmentI=FragmentMatrix(1:4,:,ixlsx);
NumI=length(nonzeros(FragmentMatrix(1:4,:,ixlsx)));
CellI{ixlsx}=sum(FragmentI(:))/(NumI*500)*1000;
end

% organize the result table and output as the xlsx file
AllI=cell2mat(CellI);
Filename=struct_file_info(1,:);
Items={'Filename','RMS','I',}';
AllI=num2cell(AllI);
FinalData=[Filename;RMS;AllI];
Result=[Items,FinalData];
%xlswrite('Result.xlsx',Result);
xlswrite('Result.xls',Result);



