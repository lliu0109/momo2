%aaaaa   places I need to change, to make the code work on my local machine

%ddddd   places I should un-comment
%E:\Index_Products\MarketEnvironment\

%sssss   [switch] to speed up the code
%           when 'FLAG_OPTIMIZE_RUNTIME' is true, only Asian & Cliquet options' value and delta are computed;
%           skip (vega, gamma, rho, theta) computations to shorten the run-time.


function [] = DoIACalc( varargin)


%aaaaa
%addpath(genpath('E:\Index_Products\SourceCode\source\'));
addpath(genpath([pwd,'\source\']));                        %enable this line, when run on network share
%addpath(genpath('E:\Index_Products\SourceCode\source\'));   %enable this line, when run on my local machine



%sssss
%when 'FLAG_OPTIMIZE_RUNTIME' is true, only Cliquet options' value
%and delta are computed; skip (vega, gamma, rho, theta) computations to shorten the run-time. 
[dummy Flag_runtime]= xlsread('IA.Process.xls', 'Input', 'D2');
if ( strcmp(Flag_runtime, 'Y') | strcmp(Flag_runtime,'y') )
   FLAG_OPTIMIZE_RUNTIME=true;
else
   FLAG_OPTIMIZE_RUNTIME=false;
end



%htc   read parameters from the Excel workbook that drives this program:   from F:\Index_Products\SourceCode\IA.Process.xls
%htc   Market Environment File
%htc   Company Name/Hedging Block
%htc   Policy Holder Tapes (for current day and last bus. day)
%htc   output results by this program:    Liability Data (for current day and last bus. day)
[dummy MktEnv_directory]=  xlsread('IA.Process.xls', 'Input', 'B11');
[dummy MktEnv_filename]=  xlsread('IA.Process.xls', 'Input', 'B12');
[dummy MktEnv_filename_BOP]=  xlsread('IA.Process.xls', 'Input', 'B13');
[dummy Company_name] = xlsread('IA.Process.xls', 'Input', 'A16:A26');       %htc   A26 is empty
[dummy Block_name] = xlsread('IA.Process.xls', 'Input', 'B16:B26');         %htc   B26 is empty
[dummy Policyholder_tape] = xlsread('IA.Process.xls', 'Input', 'C16:C26');  %htc   C26 is empty
[dummy Policyholder_tape_bop] = xlsread('IA.Process.xls', 'Input', 'D16:D26');  %htc D26 is empty
[dummy Liability_data] = xlsread('IA.Process.xls', 'Input', 'E16:E26');         %htc E26 is empty
[dummy Liability_data_BOP] = xlsread('IA.Process.xls', 'Input', 'F16:F26');     %htc F26 is empty
[dummy Model_Name] = xlsread('IA.Process.xls', 'Input', 'G16:G26');             %htc G26 is empty

number_of_hedging_block = length(Company_name);


%htc  process each company/hedging-block
for i = 1:number_of_hedging_block
    
    if strcmp(Block_name(i), 'P2P')
        DoIAP2PCalc(FLAG_OPTIMIZE_RUNTIME, MktEnv_directory, MktEnv_filename, Policyholder_tape(i), Liability_data(i), Liability_data_BOP(i));
    else
        if strcmp(Block_name(i), 'Cliquet')
            DoIACliquetCalc(FLAG_OPTIMIZE_RUNTIME,MktEnv_directory, MktEnv_filename, MktEnv_filename_BOP, Policyholder_tape(i), Policyholder_tape_bop(i), Liability_data(i), Liability_data_BOP(i), Model_Name(i));
        else
            if strcmp(Block_name(i), 'Asian')
                DoIAAsianCalc(FLAG_OPTIMIZE_RUNTIME, MktEnv_directory, MktEnv_filename, MktEnv_filename_BOP, Policyholder_tape(i), Policyholder_tape_bop(i), Liability_data(i), Liability_data_BOP(i), Model_Name(i));
            end
        end
    end
end
end

%%

function [] = DoIAP2PCalc(FLAG_OPTIMIZE_RUNTIME, MktEnv_directory, MktEnv_filename, Policyholder_tape, Liability_data, Liability_data_BOP)


%==========================================================
%read historic spx-level, current date, prev. hedging date,
%     term, r, q, strike K, sigma(T,K)
%from MktEnv.YYYYMMDD.xls file
%==========================================================

%read historical spx-level
[SPX_level SPX_date ] =xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'SPX', 'A3:B800');
SPX_date_num = datenum(SPX_date, 'mm/dd/yyyy');

%read current spx-level
SPX_current = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B3');

%read current date
[dummy current_date_string] = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B1');
current_date = datenum(current_date_string, 'mm/dd/yyyy');

%read prev. hedging date
[~, bop_date_string] = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B2');
bop_date = datenum(bop_date_string, 'mm/dd/yyyy');

%read the term, r, q
Term_t = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','A2:A15');
Interest_rate = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','B2:B15');
Dividend = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','C2:C15');


%read the strike prices
%read the volatility surface sigma(T,K)
Strike = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Vol_Surface','B2:AX2');
Vol_surface = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Vol_surface','B3:AX16');
Vol_surface_bid = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Vol_surface','B18:AX31');
Vol_surface_ask = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Vol_surface','B33:AX46');



%=====================================================
%read policies from a company/hedging block
%get 'number_of_policies' and parse policy parameters
%=====================================================
[IA_raw_data IA_raw_text] = xlsread(char(Policyholder_tape), 'Sheet1', 'A2:J65535');
[number_of_policies dummy] = size(IA_raw_data);

IA_issue_date = datenum(IA_raw_text(:, 1), 'mm/dd/yyyy');

IA_polnum = IA_raw_text(:, 3);
IA_premium = IA_raw_data(:, 1);
IA_credit_cap = IA_raw_data(:, 3);
IA_credit_floor = IA_raw_data(:, 4);
IA_participation_rate = IA_raw_data(:, 2);
IA_term=IA_raw_data(:, 7);



IA_reset_date = zeros(size(IA_issue_date));
IA_maturity_date = zeros(size(IA_issue_date));
IA_SPX_level=zeros(size(IA_issue_date));
IA_strike_lower = zeros(size(IA_issue_date));
IA_floor_adj= zeros(size(IA_issue_date));
IA_strike_upper = zeros(size(IA_issue_date));
IA_t = zeros(size(IA_issue_date));
IA_interest_rate = zeros(size(IA_issue_date));
IA_dividend = zeros(size(IA_issue_date));
IA_vol_lower = zeros(size(IA_issue_date));
IA_vol_upper = zeros(size(IA_issue_date));
IA_vol_lower_ask = zeros(size(IA_issue_date));
IA_vol_upper_bid = zeros(size(IA_issue_date));
IA_positions_number = zeros(size(IA_issue_date));
IA_issue_date_string = cell(size(IA_issue_date));
IA_reset_date_string = cell(size(IA_issue_date));


for i = 1 : number_of_policies
    flag = mod(year(current_date) - year(IA_issue_date(i)), IA_term(i));    

    if flag ==1 
        IA_maturity_date(i) =  datenum(year(current_date)+1, month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
        IA_reset_date(i) =  datenum(year(current_date)-1, month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
    else
        if datenum(year(current_date), month(IA_issue_date(i)-1), day(IA_issue_date(i)-1)) > current_date
            IA_maturity_date(i) =  datenum(year(current_date), month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
            IA_reset_date(i) =  datenum(year(current_date)-IA_term(i), month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
        else
            IA_maturity_date(i) = datenum(year(current_date)+IA_term(i), month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
            IA_reset_date(i) =  datenum(year(current_date), month(IA_issue_date(i)-1), day(IA_issue_date(i)-1));
        end
    end
    
    IA_issue_date_string(i) = {datestr(IA_issue_date(i), 'mm/dd/yyyy')};
    
    IA_reset_date_string(i) = {datestr(IA_reset_date(i), 'mm/dd/yyyy')};

    IA_SPX_level(i)=SPX_level(SPX_date_num == IA_reset_date(i));                            %S0, the initial S
    
    IA_strike_lower(i) = IA_SPX_level(i)*(1+IA_credit_floor(i)/IA_participation_rate(i));
    
    IA_strike_upper(i) = IA_SPX_level(i) * (1 + IA_credit_cap(i)/IA_participation_rate(i));
    
    IA_positions_number(i) = IA_premium(i) * IA_participation_rate(i) / IA_SPX_level(i);
    
    IA_t(i) = (IA_maturity_date(i) - current_date)/365.25;
    
    IA_interest_rate(i) = interpolation_1d(Term_t, Interest_rate, IA_t(i));
    
    IA_dividend(i) = interpolation_1d(Term_t, Dividend, IA_t(i));
    
    IA_vol_lower(i) = interpolation_2d(Term_t', Strike, Vol_surface, IA_t(i), IA_strike_lower(i));
    
    IA_vol_upper(i) = interpolation_2d(Term_t',Strike, Vol_surface, IA_t(i), IA_strike_upper(i));
    
    IA_vol_lower_ask(i) = interpolation_2d(Term_t', Strike, Vol_surface_ask, IA_t(i), IA_strike_lower(i));
    
    IA_vol_upper_bid(i) = interpolation_2d(Term_t',Strike, Vol_surface_bid, IA_t(i), IA_strike_upper(i));
    
    IA_floor_adj(i)=IA_SPX_level(i)*(IA_credit_floor(i)/IA_participation_rate(i));
end



%===========================================================
%compute the greeks for each policy:  IA_Greeks(policy #,16)
%===========================================================
IA_Greeks = zeros(number_of_policies, 16);
[dummy_lower dummy] = blsprice(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper dummy] = blsprice(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 1) = (dummy_lower+IA_floor_adj);
IA_Greeks(:, 2) = dummy_upper;
[dummy_lower dummy] = blsdelta(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper dummy] = blsdelta(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 3) = dummy_lower;
IA_Greeks(:, 4) = dummy_upper;

[dummy_lower dummy] = blsrho(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper dummy] = blsrho(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 5) = dummy_lower/100;
IA_Greeks(:, 6) = dummy_upper/100;
[dummy_lower] = blsvega(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper] = blsvega(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 7) = dummy_lower/100;
IA_Greeks(:, 8) = dummy_upper/100;
[dummy_lower] = blsgamma(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper] = blsgamma(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 9) = dummy_lower;
IA_Greeks(:, 10) = dummy_upper;

clear dummy_lower dummy_upper dummy

[dummy_lower_up] = blsgamma(SPX_current+5, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_lower_down] = blsgamma(SPX_current-5, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);

[dummy_upper_up] = blsgamma(SPX_current+5, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
[dummy_upper_down] = blsgamma(SPX_current-5, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);

IA_Greeks(:, 11) = (dummy_lower_up - dummy_lower_down)/10;
IA_Greeks(:, 12) = (dummy_upper_up - dummy_upper_down)/10;

[dummy_lower dummy] = blstheta(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower, IA_dividend);
[dummy_upper dummy] = blstheta(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper, IA_dividend);
IA_Greeks(:, 13) = dummy_lower;
IA_Greeks(:, 14) = dummy_upper;

[dummy_lower_ask dummy] = blsprice(SPX_current, IA_strike_lower, IA_interest_rate, IA_t, IA_vol_lower_ask, IA_dividend);
[dummy_upper_bid dummy] = blsprice(SPX_current, IA_strike_upper, IA_interest_rate, IA_t, IA_vol_upper_bid, IA_dividend);
IA_Greeks(:, 15) = dummy_lower_ask;
IA_Greeks(:, 16) = dummy_upper_bid;

clear dummy_lower_up dummy_lower_down dummy_upper_up dummy_upper_down dummy_lower dummy_upper

r_weighted = sum((IA_Greeks(:, 5) - IA_Greeks(:, 6)) .* IA_interest_rate .* IA_positions_number) / sum( (IA_Greeks(:, 5) - IA_Greeks(:, 6)) .* IA_positions_number);


%====================================================
%output the columns A to K in 'IA_Liabilities' worksheet
%====================================================
xlswrite(char(Liability_data), {'Pol_num'}, 'IA_Liabilities','A3');
xlswrite(char(Liability_data), IA_polnum, 'IA_Liabilities','A4');
xlswrite(char(Liability_data), {'issue_date'}, 'IA_Liabilities','B3');
xlswrite(char(Liability_data), IA_issue_date_string, 'IA_Liabilities','B4');
xlswrite(char(Liability_data), {'reset_date'}, 'IA_Liabilities','C3');
xlswrite(char(Liability_data), IA_reset_date_string, 'IA_Liabilities','C4');
xlswrite(char(Liability_data), {'premium'}, 'IA_Liabilities','D3');
xlswrite(char(Liability_data),IA_premium, 'IA_Liabilities','D4');
xlswrite(char(Liability_data), {'# of Positions'}, 'IA_Liabilities','E3');
xlswrite(char(Liability_data),IA_positions_number, 'IA_Liabilities','E4');
xlswrite(char(Liability_data), {'Kl'}, 'IA_Liabilities','F3');
xlswrite(char(Liability_data),IA_strike_lower, 'IA_Liabilities','F4');
xlswrite(char(Liability_data), {'Ku'}, 'IA_Liabilities','G3');
xlswrite(char(Liability_data),IA_strike_upper, 'IA_Liabilities', 'G4');
xlswrite(char(Liability_data), {'r'}, 'IA_Liabilities','H3');
xlswrite(char(Liability_data),IA_interest_rate, 'IA_Liabilities','H4');
xlswrite(char(Liability_data), {'q'}, 'IA_Liabilities','I3');
xlswrite(char(Liability_data),IA_dividend, 'IA_Liabilities','I4');
xlswrite(char(Liability_data), {'sigma_l'}, 'IA_Liabilities','J3');
xlswrite(char(Liability_data),IA_vol_lower, 'IA_Liabilities','J4');
xlswrite(char(Liability_data), {'sigma_u'}, 'IA_Liabilities','K3');
xlswrite(char(Liability_data),IA_vol_upper, 'IA_Liabilities','K4');



%====================================================
%output columns L to AA in 'IA_Liabilities' worksheet
%====================================================
xlswrite(char(Liability_data), {'unit_value_lower', 'unit_value_upper', 'unit_delta_lower', 'unit_delta_upper',...
    'unit_rho_lower', 'unit_rho_upper', 'unit_vega_lower', 'unit_vega_upper', 'unit_gamma_lower',...
    'unit_gamma_upper', 'unit_speed_lower', 'unit_speed_upper', 'unit_theta_lower', 'unit_theta_upper','lower_ask','upper_bid'}, 'IA_Liabilities','L3');
xlswrite(char(Liability_data),IA_Greeks, 'IA_Liabilities','L4'); % can only specify the starting point of the range


%=====================================================
%output columns AE to AM in 'IA_Liabilities' worksheet
%=====================================================
xlswrite(char(Liability_data), {'pol_value', 'pol_delta', 'pol_rho', 'pol_vega_lower','pol_vega_upper', 'pol_gamma', 'pol_speed', 'pol_theta', 'mkt_price'}, 'IA_Liabilities','AE3');

[temp_row temp_col] = size(IA_Greeks);
xlswrite(char(Liability_data),(IA_Greeks(:, 1:2:6) - IA_Greeks(:, 2:2:6)) .*repmat(IA_positions_number, 1, 3), 'IA_Liabilities','AE4');                 %Columns AE to AG
xlswrite(char(Liability_data), IA_Greeks(:, 7:8) .*repmat(IA_positions_number, 1, 2),  'IA_Liabilities','AH4');                                         %columns AH to AI
xlswrite(char(Liability_data),(IA_Greeks(:, 9:2:end) - IA_Greeks(:, 10:2:end)) .*repmat(IA_positions_number, 1, (temp_col-8)/2), 'IA_Liabilities','AJ4');  %columns AJ to AM

IA_Greeks_agg = sum(IA_Greeks .* repmat(IA_positions_number, 1, temp_col));




%=============================================
%output to 'IA_Greeks' worksheet
%=============================================
clear temp_row temp_col

xlswrite(char(Liability_data), {'Date'}, 'IA_Greeks', 'A1');
xlswrite(char(Liability_data), current_date_string, 'IA_Greeks', 'B1');

xlswrite(char(Liability_data), {'Value' 'Delta' 'Rho' 'Vega Lower' 'Vega Upper' 'Gamma' 'Speed' 'theta' 'Mkt Price'}, 'IA_Greeks', 'A1');
xlswrite(char(Liability_data),IA_Greeks_agg(1:2:6) - IA_Greeks_agg(2:2:6), 'IA_Greeks','A2');
xlswrite(char(Liability_data),IA_Greeks_agg(7), 'IA_Greeks','D2');
xlswrite(char(Liability_data),IA_Greeks_agg(8), 'IA_Greeks','E2');
xlswrite(char(Liability_data),IA_Greeks_agg(9:2:end) - IA_Greeks_agg(10:2:end), 'IA_Greeks','F2');


xlswrite(char(Liability_data), {'SPX'; 'r'; 'Sigma lower'; 'Sigma upper'}, 'IA_Greeks', 'A4:A7');
xlswrite(char(Liability_data), SPX_current, 'IA_Greeks', 'B4');
xlswrite(char(Liability_data), r_weighted, 'IA_Greeks', 'B5');
%added


%======================================================================================================================================================================
%If Results from previous business day exists AND we are not optimizing the run-time then
%======================================================================================================================================================================
%sssss
if (   (exist([char(Liability_data_BOP)], 'file') ~= 0) &  (~FLAG_OPTIMIZE_RUNTIME)  )
    
    % first do policy by policy attribution analysis
    
    %Read from Prev. Hedge Date outputs
    [data text] = xlsread(char(Liability_data_BOP), 'IA_liabilities', 'A4:AL65536');
    IA_interest_rate_bop = data(:, 5);
    IA_vol_lower_bop = data(:, 7);
    IA_vol_upper_bop = data(:, 8);
    
    IA_value_bop = data(:, 28);                     %column AE
    IA_delta_bop = data(:, 29);                     %column AF
    IA_rho_bop = data(:, 30);                       %column AG
    IA_vega_lower_bop = data(:, 31);                %column AH
    IA_vega_upper_bop = data(:, 32);                %column AI
    IA_gamma_bop = data(:, 33);
    IA_speed_bop = data(:, 34);
    IA_theta_bop = data(:, 35);
    IA_strike_lower_bop = data(:, 3);
    IA_strike_upper_bop = data(:, 4);
    
    IA_data_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'B4:B7');
    
    IA_aa = zeros(number_of_policies, 13);
    vol_vega_products = zeros(number_of_policies, 4);
    %BOP_Value, Delta Change, Rho Change, Vega Change, Gamma Change,
    %Speed Change, Theta Change,  New Business Ind,
    %Settlement Ind, Settle Value, epsilon, EOP Value
    
    
    % new the EOP policy number, find the correponding BOP policy
    % number
    for i = 1 : number_of_policies
        IA_value_bop_by_pol = IA_value_bop(strcmp(text(:,1),IA_polnum(i)));     %for policy i, find the prev. Hedge date's IA Value 
                    
         
        %if new business else old business
        if isempty(IA_value_bop_by_pol)
            IA_aa(i,13) =  (IA_Greeks(i,1) - IA_Greeks(i, 2))*IA_positions_number(i);
            IA_aa(i,9) = 1;
            
        else
            if size(IA_value_bop_by_pol) == 1
                
                
                IA_aa(i,1) = IA_value_bop_by_pol;
                %delta change
                IA_aa(i,2) = IA_delta_bop(strcmp(text(:,1),IA_polnum(i))) * (SPX_current - IA_data_bop(1));
                %rho change
                IA_aa(i,3) = IA_rho_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_interest_rate(i) - IA_interest_rate_bop(strcmp(text(:,1),IA_polnum(i))))*100;
                %vega change
                %IA_aa(i,4) = IA_vega_lower_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_vol_lower(i) - IA_vol_lower_bop(strcmp(text(:,1),IA_polnum(i))))*100 ...
                %- IA_vega_upper_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_vol_upper(i) - IA_vol_upper_bop(strcmp(text(:,1),IA_polnum(i))))*100;
                IA_aa(i,4) = IA_vega_lower_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_vol_lower(i) - IA_vol_lower_bop(strcmp(text(:,1),IA_polnum(i))))*100;
                
                IA_aa(i,5) = IA_vega_upper_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_vol_upper(i) - IA_vol_upper_bop(strcmp(text(:,1),IA_polnum(i))))*100;
                %gamma change
                IA_aa(i,6) = IA_gamma_bop(strcmp(text(:,1),IA_polnum(i))) * power(SPX_current - IA_data_bop(1),2)/2;
                %Speed change
                IA_aa(i,7) = IA_speed_bop(strcmp(text(:,1),IA_polnum(i))) * power(SPX_current - IA_data_bop(1),3)/6;
                %theta change
                IA_aa(i,8) = IA_theta_bop(strcmp(text(:,1),IA_polnum(i))) * (current_date - bop_date)/365.25;
                %New business Indicator.
                IA_aa(i,9) = (IA_reset_date(i) <= current_date) && (IA_reset_date(i) > bop_date);
                %Settlement Indicator
                IA_aa(i,10) = (IA_reset_date(i) <= current_date) && (IA_reset_date(i) > bop_date);
                %Settlement value/
                spxset = IA_strike_lower(i);
                IA_lower_strike_prev = IA_strike_lower_bop(strcmp(text(:,1),IA_polnum(i)));
                IA_upper_strike_prev = IA_strike_upper_bop(strcmp(text(:,1),IA_polnum(i)));
                
                %column AO (i.e. bop_pol_value)
                %if policy i is renewed, then IA_aa(i,11) is the policy's Settlement Value (column AY)
                ifloor = IA_credit_floor(i)*IA_strike_lower_bop( strcmp(text(:,1),IA_polnum(i))  );
                IA_aa(i,11) = ( min(     max(  (spxset - IA_lower_strike_prev)*IA_participation_rate(i),  ifloor), ...
                                         (IA_upper_strike_prev - IA_lower_strike_prev) ...
                                   )...
                                   / IA_strike_lower_bop( strcmp(text(:,1),IA_polnum(i))  ) * IA_premium(i) ...
                               )* IA_aa(i,10);
                %EOP value
                IA_aa(i,13) = (IA_Greeks(i,1) - IA_Greeks(i, 2))*IA_positions_number(i);
                %epsilon
                IA_aa(i,12) = IA_aa(i,13) - sum(IA_aa(i,1:3))-(IA_aa(i,4) - IA_aa(i,5))-sum(IA_aa(i,6:8))+    (IA_aa(i, 11) - IA_aa(i, 13))*IA_aa(i,10);
                
                vol_vega_products(i, 1) =  IA_vega_lower_bop(strcmp(text(:,1),IA_polnum(i))) * IA_vol_lower(i);
                vol_vega_products(i, 2) =  IA_vega_upper_bop(strcmp(text(:,1),IA_polnum(i))) * IA_vol_upper(i);
                vol_vega_products(i, 3) =  IA_vega_lower_bop(strcmp(text(:,1),IA_polnum(i))) * IA_vol_lower_bop(strcmp(text(:,1),IA_polnum(i)));
                vol_vega_products(i, 4) =  IA_vega_upper_bop(strcmp(text(:,1),IA_polnum(i))) * IA_vol_upper_bop(strcmp(text(:,1),IA_polnum(i)));
                vol_vega_products(i, 5) =  IA_vega_lower_bop(strcmp(text(:,1),IA_polnum(i))) ;
                vol_vega_products(i, 6) =  IA_vega_upper_bop(strcmp(text(:,1),IA_polnum(i))) ;
                
                
            else
                % msg('duplication in policy num in the bopious run date');
            end
        end
        
    end
    
    Vol_weighted_lower = sum(vol_vega_products(:,1)) / sum(vol_vega_products(:,5));
    Vol_weighted_upper = sum(vol_vega_products(:,2)) / sum(vol_vega_products(:,6));
    Vol_weighted_lower_bop = sum(vol_vega_products(:,3)) / sum(vol_vega_products(:,5));
    Vol_weighted_upper_bop = sum(vol_vega_products(:,4)) / sum(vol_vega_products(:,6));
    
    xlswrite(char(Liability_data), Vol_weighted_lower, 'IA_Greeks', 'B6');
    xlswrite(char(Liability_data), Vol_weighted_upper, 'IA_Greeks', 'B7');
    
    
    xlswrite(char(Liability_data), {'bop_pol_value', 'delta_change', 'rho_change', 'vega_lower', 'vega_upper', 'pol_gamma_change', 'pol_speed_change', 'pol_theta_change', 'NB Ind', 'Settlement Ind', 'Settlement Value', 'epsilon','eop_pol_value'}, 'IA_Liabilities','AO3');
    xlswrite(char(Liability_data), IA_aa, 'IA_Liabilities','AO4');
    
    xlswrite(char(Liability_data), bop_date_string, 'Attribution_Analysis', 'B2:C2');
    xlswrite(char(Liability_data), current_date_string, 'Attribution_Analysis', 'D2');
    xlswrite(char(Liability_data), {'BOP Value'; 'Delta'; 'Rho'; 'Vega_lower'; 'Vega_upper';'Gamma'; 'Speed'; 'Theta'; 'Settlement Value'; 'NB value'; 'Epsilon'; 'EOP Value'}, 'Attribution_Analysis', 'A3');
    
    IA_Greeks_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'A2:H2');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(1), 'Attribution_Analysis', 'E3');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(2:end)', 'Attribution_Analysis', 'B4');
    
    xlswrite(char(Liability_data), IA_data_bop, 'Attribution_Analysis', 'C4:C5');
    
    xlswrite(char(Liability_data), SPX_current, 'Attribution_Analysis', 'D4');
    xlswrite(char(Liability_data), r_weighted, 'Attribution_Analysis', 'D5');
    xlswrite(char(Liability_data), Vol_weighted_lower, 'Attribution_Analysis', 'D6');
    xlswrite(char(Liability_data), Vol_weighted_upper, 'Attribution_Analysis', 'D7');
    xlswrite(char(Liability_data), Vol_weighted_lower_bop, 'Attribution_Analysis', 'C6');
    xlswrite(char(Liability_data), Vol_weighted_upper_bop, 'Attribution_Analysis', 'C7');
    
    
    Delta_change = IA_Greeks_bop(2) * (SPX_current - IA_data_bop(1));
    Rho_change = IA_Greeks_bop(3) * (r_weighted - IA_data_bop(2)) * 100;
    %Vega_change = IA_Greeks_bop(4) * (Vol_weighted - IA_data_bop(3)) * 100;
    Vega_lower = IA_Greeks_bop(4) * (Vol_weighted_lower - Vol_weighted_lower_bop) * 100;
    Vega_upper = IA_Greeks_bop(5) * (Vol_weighted_upper - Vol_weighted_upper_bop) * 100;
    Gamma_change = 0.5 * IA_Greeks_bop(6) * (SPX_current - IA_data_bop(1))^2;
    Speed_change = 1/6 * IA_Greeks_bop(7) * (SPX_current - IA_data_bop(1))^3;
    Theta_change = IA_Greeks_bop(8) * (current_date - bop_date)/365.25;
    IA_value =  IA_Greeks_agg(1) - IA_Greeks_agg(2);
    NB_value = sum(IA_aa(:, 13).*IA_aa(:, 9));
    Settle_value = sum(IA_aa(:, 11).*IA_aa(:,10));
    Epsilon = IA_value - IA_Greeks_bop(1) - Delta_change - Rho_change - (Vega_lower-Vega_upper) - Gamma_change - Speed_change-Theta_change+Settle_value - NB_value;
    
    xlswrite(char(Liability_data), Delta_change, 'Attribution_Analysis', 'E4');
    xlswrite(char(Liability_data), Rho_change, 'Attribution_Analysis', 'E5');
    xlswrite(char(Liability_data), Vega_lower, 'Attribution_Analysis', 'E6');
    xlswrite(char(Liability_data), Vega_upper, 'Attribution_Analysis', 'E7');
    xlswrite(char(Liability_data), Gamma_change, 'Attribution_Analysis', 'E8');
    xlswrite(char(Liability_data), Speed_change, 'Attribution_Analysis', 'E9');
    xlswrite(char(Liability_data), Theta_change, 'Attribution_Analysis', 'E10');
    xlswrite(char(Liability_data), NB_value, 'Attribution_Analysis', 'E12');
    xlswrite(char(Liability_data), Settle_value, 'Attribution_Analysis', 'E11');
    xlswrite(char(Liability_data), Epsilon, 'Attribution_Analysis', 'E13');
    xlswrite(char(Liability_data), IA_value, 'Attribution_Analysis', 'E14');
    
    %sum of the policy by policy attribution analysis
    AA_array = zeros(12,1);
    AA_array(1,1) = IA_Greeks_bop(1);%bop
    AA_array(2,1) = sum( IA_aa(:,2));%delta
    AA_array(3,1) = sum( IA_aa(:,3));%rho
    AA_array(4,1) = sum( IA_aa(:,4));%vega lower
    AA_array(5,1) = sum( IA_aa(:,5));%vega upper
    AA_array(6,1) = sum( IA_aa(:,6));%gamma
    AA_array(7,1) = sum( IA_aa(:,7));%speed
    AA_array(8,1) = sum( IA_aa(:,8));%theta
    AA_array(9,1) = Settle_value;
    AA_array(10,1) = NB_value;
    AA_array(12,1) = IA_value;
    AA_array(11,1) = sum( IA_aa(:,12));
    
    xlswrite(char(Liability_data), {'Pol by Pol AA'}, 'Attribution_Analysis', 'G2');
    xlswrite(char(Liability_data), AA_array, 'Attribution_Analysis', 'G3');
    
    
end

end

function [] = DoIACliquetCalc(FLAG_OPTIMIZE_RUNTIME, MktEnv_directory, MktEnv_filename, MktEnv_filename_bop, Policyholder_tape, Policyholder_tape_bop, Liability_data, Liability_data_BOP, Model_Name)

%htc   read market & policy data
[SPX_level, SPX_date_num, SPX_current,current_date, bop_date, Term_t, Interest_rate_forward, ...
 Dividend_forward, Interest_rate_spot, aVolParams, IA_issue_date, IA_issue_date_string, ...
 IA_reset_date, IA_reset_date_string, IA_polnum, IA_premium, IA_participation_rate, ...
 IA_credit_cap, IA_credit_floor, IA_prorata_factor, IA_index_spread, number_of_policies, Num_of_Paths ...
]= Read_Market_Policy_Data(MktEnv_directory, MktEnv_filename, Policyholder_tape, Model_Name);



[aCliquetValue, aCliquetDelta, aCliquetVega, aCliquetGamma, aCliquetRho, aCliquetTheta] = mCliquetValuation(FLAG_OPTIMIZE_RUNTIME, ...
                                                                                                            aVolParams, Interest_rate_forward, Dividend_forward, Term_t, ...
                                                                                                            SPX_current, current_date, SPX_level, SPX_date_num, ...
                                                                                                            IA_reset_date, IA_credit_cap, IA_credit_floor, ...
                                                                                                            (IA_premium .* IA_participation_rate .* IA_prorata_factor), ...
                                                                                                            Num_of_Paths, Model_Name);

                                                                                                        
                                                                                                        
%calculate the average interest rate. For simplicity, calculate the rho
%weighted spot rate. 
IA_spot_rate = zeros(size(IA_credit_floor));

for i = 1:number_of_policies
    Holidays_in_Between = holidays(current_date, datemnth(IA_reset_date(i), 12) );
    Time_to_maturity = wrkdydif(current_date, datemnth(IA_reset_date(i), 12), length(Holidays_in_Between))/252;
    IA_spot_rate(i) = interpolation_1d(Term_t, Interest_rate_spot, Time_to_maturity); 
end

Rate_weighted = sum(IA_spot_rate .* aCliquetRho)/ sum(aCliquetRho);


%htc  output results to workbook:   F:\Index_Products\LiabilityValuation\AGLIACliquet\AGL.IA.Cliquet.Liability.Data.20140102.xls
%htc                           columns A:Q        IA_Liabilities  worksheet
%recalculate the bop cliquet price based on EOP's bates parameters.
xlswrite(char(Liability_data), {'Pol_num'}, 'IA_Liabilities','A3');
xlswrite(char(Liability_data), IA_polnum, 'IA_Liabilities','A4');
xlswrite(char(Liability_data), {'issue_date'}, 'IA_Liabilities','B3');
xlswrite(char(Liability_data), IA_issue_date_string, 'IA_Liabilities','B4');
xlswrite(char(Liability_data), {'reset_date'}, 'IA_Liabilities','C3');
xlswrite(char(Liability_data), IA_reset_date_string, 'IA_Liabilities','C4');
xlswrite(char(Liability_data), {'premium'}, 'IA_Liabilities','D3');
xlswrite(char(Liability_data),IA_premium, 'IA_Liabilities','D4');
xlswrite(char(Liability_data), {'Cap'}, 'IA_Liabilities','E3');
xlswrite(char(Liability_data),IA_credit_cap, 'IA_Liabilities','E4');
xlswrite(char(Liability_data), {'Floor'}, 'IA_Liabilities','F3');
xlswrite(char(Liability_data),IA_credit_floor, 'IA_Liabilities', 'F4');
xlswrite(char(Liability_data), {'rate'}, 'IA_Liabilities','G3');
xlswrite(char(Liability_data),IA_spot_rate, 'IA_Liabilities', 'G4');

xlswrite(char(Liability_data), {'Cliquet Value'}, 'IA_Liabilities','J3');
xlswrite(char(Liability_data),aCliquetValue, 'IA_Liabilities', 'J4');
xlswrite(char(Liability_data), {'Cliquet Delta'}, 'IA_Liabilities','K3');
xlswrite(char(Liability_data),aCliquetDelta, 'IA_Liabilities', 'K4');
xlswrite(char(Liability_data), {'Cliquet Gamma'}, 'IA_Liabilities','L3');
xlswrite(char(Liability_data),aCliquetGamma, 'IA_Liabilities', 'L4');
xlswrite(char(Liability_data), {'Cliquet Vega'}, 'IA_Liabilities','M3');
xlswrite(char(Liability_data),aCliquetVega, 'IA_Liabilities', 'M4');
xlswrite(char(Liability_data), {'Cliquet Rho'}, 'IA_Liabilities','N3');
xlswrite(char(Liability_data),aCliquetRho, 'IA_Liabilities', 'N4');
xlswrite(char(Liability_data), {'Cliquet Theta'}, 'IA_Liabilities','O3');
xlswrite(char(Liability_data),aCliquetTheta, 'IA_Liabilities', 'O4');

%htc  output results to workbook:   F:\Index_Products\LiabilityValuation\AGLIACliquet\AGL.IA.Cliquet.Liability.Data.20140102.xls
%htc                                            IA_Greeks worksheet
xlswrite(char(Liability_data),{'Value'}, 'IA_Greeks', 'A1');
xlswrite(char(Liability_data), sum(aCliquetValue), 'IA_Greeks', 'A2');
xlswrite(char(Liability_data),{'Delta'}, 'IA_Greeks', 'B1');
xlswrite(char(Liability_data), sum(aCliquetDelta), 'IA_Greeks', 'B2');
xlswrite(char(Liability_data),{'Vega'}, 'IA_Greeks', 'C1');
xlswrite(char(Liability_data), sum(aCliquetVega), 'IA_Greeks', 'C2');
xlswrite(char(Liability_data),{'Rho'}, 'IA_Greeks', 'D1');
xlswrite(char(Liability_data), sum(aCliquetRho), 'IA_Greeks', 'D2');
xlswrite(char(Liability_data),{'Gamma'}, 'IA_Greeks', 'E1');
xlswrite(char(Liability_data), sum(aCliquetGamma), 'IA_Greeks', 'E2');
xlswrite(char(Liability_data),{'Theta'}, 'IA_Greeks', 'F1');
xlswrite(char(Liability_data), sum(aCliquetTheta), 'IA_Greeks', 'F2');
xlswrite(char(Liability_data),{'SPX'}, 'IA_Greeks', 'A5');
xlswrite(char(Liability_data), SPX_current, 'IA_Greeks', 'A6');
xlswrite(char(Liability_data),{'r'}, 'IA_Greeks', 'B5');
xlswrite(char(Liability_data), Rate_weighted, 'IA_Greeks', 'B6');


%======================================================================================================================================================================
%htc  If Results from previous business day exists AND we are not optimizing the run-time then
%htc    compute Cliquet Options' prev. day's value,
%htc    compute Cliquet Options' value change due to greeks from prev. business day
%htc    this is columns Q:V in 'IA_Liabilities' worksheet in workbook    F:\Index_Products\LiabilityValuation\AGL IA Cliquet\AGL.IA.Cliquet.Liability.Data.20140102.xls
%======================================================================================================================================================================
%sssss
if (   (exist([char(Liability_data_BOP)], 'file') ~= 0) &  (~FLAG_OPTIMIZE_RUNTIME)  )
    
    [SPX_level_bop, SPX_date_num_bop, SPX_current_bop,current_date_bop, bop_date_bop, Term_t_bop, Interest_rate_forward_bop, Dividend_forward_bop, Interest_rate_spot_bop, aVolParams_bop, IA_issue_date_bop, IA_issue_date_bop_string, ...
        IA_reset_date_bop, IA_reset_date_bop_string, IA_polnum_bop, IA_premium_bop, IA_participation_rate_bop, IA_credit_cap_bop, IA_credit_floor_bop, IA_prorata_factor_bop, IA_index_spread_bop, number_of_policies_bop, Num_of_Paths ] = Read_Market_Policy_Data(MktEnv_directory, MktEnv_filename_bop, Policyholder_tape_bop, Model_Name);
    
    [aCliquetValue_bop_recalc] = mCliquetValuation(FLAG_OPTIMIZE_RUNTIME, ...
        aVolParams, Interest_rate_forward_bop, Dividend_forward_bop, Term_t_bop, SPX_current_bop, current_date_bop, SPX_level_bop, SPX_date_num_bop, ...
        IA_reset_date_bop, IA_credit_cap_bop, IA_credit_floor_bop, (IA_premium_bop .* IA_participation_rate_bop .* IA_prorata_factor_bop), Num_of_Paths, Model_Name );
    
    
    
    % first do policy by policy attribution analysis
    
    [data text] = xlsread(char(Liability_data_BOP), 'IA_liabilities', 'A4:O65536');
    
    IA_rate_bop = data(:, 4);
    IA_value_bop = data(:, 7);
    IA_delta_bop = data(:, 8);
    IA_gamma_bop = data(:, 9);
    IA_rho_bop = data(:, 11);
    IA_theta_bop = data(:, 12);
    
        
    IA_data_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'A6:B6');
    
    IA_aa = zeros(number_of_policies, 11);
    
    %BOP_Value, Delta Change, Rho Change, Vega Change, Gamma Change,
    %Speed Change, Theta Change,  New Business Ind,
    %Settlement Ind, Settle Value, epsilon, EOP Value
    
    
    % new the EOP policy number, find the correponding BOP policy
    % number
   
    % used in Theta attribution
    Holidays_in_between = holidays(bop_date, current_date);
    Buzdays_in_between = wrkdydif(bop_date, current_date, length(Holidays_in_between)) - 1;
    
    for i = 1 : number_of_policies
        IA_value_bop_by_pol = IA_value_bop(strcmp(text(:,1),IA_polnum(i)));
        
       
        if isempty(IA_value_bop_by_pol)
            % new business
            IA_aa(i,11) =  aCliquetValue(i);
            IA_aa(i,7) = 1;
            
        else
            if size(IA_value_bop_by_pol) == 1
                
                
                IA_aa(i,1) = IA_value_bop_by_pol;
                %delta change
                IA_aa(i,2) = IA_delta_bop(strcmp(text(:,1),IA_polnum(i))) * (SPX_current - IA_data_bop(1));
                
                %Vega change
                IA_aa(i,3) = aCliquetValue_bop_recalc(strcmp(text(:,1),IA_polnum(i))) - IA_value_bop_by_pol;
                %gamma change
                IA_aa(i,4) = IA_gamma_bop(strcmp(text(:,1),IA_polnum(i))) * power(SPX_current - IA_data_bop(1),2)/2;
                %rho change
                IA_aa(i,5) = IA_rho_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_spot_rate(i) - IA_rate_bop(strcmp(text(:,1),IA_polnum(i)))) * 100;
                %theta change
                IA_aa(i,6) = IA_theta_bop(strcmp(text(:,1),IA_polnum(i))) * Buzdays_in_between /252;
                
                %New business Indicator.
                IA_aa(i,7) = (IA_reset_date(i) < NextBusinessDate(current_date)) && (IA_reset_date(i) >= current_date);
                %Settlement Indicator
                IA_aa(i,8) = (IA_reset_date(i) < NextBusinessDate(current_date)) && (IA_reset_date(i) >= current_date) && (IA_reset_date(i)~= IA_issue_date(i));
                %Settlement value/
                if IA_aa(i, 8) == 1
                     % add three more SPX level to handle the weekend
                     % settle or weekend and the next Monday is a holiday
                     % situation
                    SPX_level_appd = [SPX_level; SPX_level(end); SPX_level(end); SPX_level(end)];
                    SPX_date_num_appd = [SPX_date_num; SPX_date_num(end)+1; SPX_date_num(end)+2; SPX_date_num(end)+3];
                    IA_aa(i,9) = CliquetPrice(IA_reset_date(i), datenum(year(IA_reset_date(i))-1, month(IA_issue_date(i)), day(IA_issue_date(i))), SPX_level_appd, SPX_date_num_appd , IA_credit_cap(i), IA_credit_floor(i), (IA_premium(i) .* IA_participation_rate(i)), zeros(1,1) );
                end
                %EOP value
                IA_aa(i,11) = aCliquetValue(i);
                %epsilon
                IA_aa(i,10) = IA_aa(i,11) - sum(IA_aa(i, 1:6))+ (IA_aa(i,9)-IA_aa(i,11)) * IA_aa(i,8);
                
            else
                % msg('duplication in policy num in the bopious run date');
            end
        end
        
    end
    
    xlswrite(char(Liability_data), {'bop_pol_value', 'delta_change', 'vega_change', 'gamma_change', 'rho_change', 'theta_change', 'NB Ind', 'Settlement Ind', 'Settlement Value', 'epsilon','eop_pol_value'}, 'IA_Liabilities','Q3');
    xlswrite(char(Liability_data), IA_aa, 'IA_Liabilities','Q4');
    
    xlswrite(char(Liability_data), {datestr(bop_date, 'mm/dd/yyyy')}, 'Attribution_Analysis', 'B2:C2');
    xlswrite(char(Liability_data), {datestr(current_date, 'mm/dd/yyyy')}, 'Attribution_Analysis', 'D2');
    xlswrite(char(Liability_data), {'BOP Value'; 'Delta'; 'Vega'; 'Rho'; 'Gamma'; 'Theta'; 'Settlement Value'; 'NB value'; 'Epsilon'; 'EOP Value'}, 'Attribution_Analysis', 'A3');
    
    IA_Greeks_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'A2:G2');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(1), 'Attribution_Analysis', 'E3');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(2:end)', 'Attribution_Analysis', 'B4');
    
    xlswrite(char(Liability_data), IA_data_bop(1), 'Attribution_Analysis', 'C4');
    
    xlswrite(char(Liability_data), SPX_current, 'Attribution_Analysis', 'D4');
    
    xlswrite(char(Liability_data), IA_data_bop(2), 'Attribution_Analysis', 'C6');
    
    xlswrite(char(Liability_data), Rate_weighted, 'Attribution_Analysis', 'D6'); 
    
    
    Delta_change = IA_Greeks_bop(2) * (SPX_current - IA_data_bop(1));
    Vega_change = sum(IA_aa(:,3));
    Rho_change = IA_Greeks_bop(4) * (Rate_weighted - IA_data_bop(2))*100;
    Gamma_change = 0.5 * IA_Greeks_bop(5) * (SPX_current - IA_data_bop(1))^2;
    Theta_change = sum(IA_aa(:,6));
    IA_value =  sum(IA_aa(:,11));
    NB_value = sum(IA_aa(:, 11).*IA_aa(:, 7));
    Settle_value = sum(IA_aa(:, 9).*IA_aa(:,8));
    Epsilon = IA_value - IA_Greeks_bop(1) - Delta_change - Vega_change - Gamma_change - Rho_change - Theta_change + Settle_value - NB_value;
    
    xlswrite(char(Liability_data), Delta_change, 'Attribution_Analysis', 'E4');
    xlswrite(char(Liability_data), Vega_change, 'Attribution_Analysis', 'E5');
     xlswrite(char(Liability_data), Rho_change, 'Attribution_Analysis', 'E6');
    xlswrite(char(Liability_data), Gamma_change, 'Attribution_Analysis', 'E7');
    xlswrite(char(Liability_data), Theta_change, 'Attribution_Analysis', 'E8');
    xlswrite(char(Liability_data), Settle_value, 'Attribution_Analysis', 'E9');
    xlswrite(char(Liability_data), NB_value, 'Attribution_Analysis', 'E10');
    xlswrite(char(Liability_data), Epsilon, 'Attribution_Analysis', 'E11');
    xlswrite(char(Liability_data), IA_value, 'Attribution_Analysis', 'E12');
    
    %sum of the policy by policy attribution analysis,
    %     AA_array = zeros(9,1);
    %     AA_array(1,1) = IA_Greeks_bop(1);%bop
    %     AA_array(2,1) = sum( IA_aa(:,2));%delta
    %     AA_array(3,1) = sum( IA_aa(:,3));%vega
    %     AA_array(4,1) = sum( IA_aa(:,4));%gamma
    %     AA_array(5,1) = sum( IA_aa(:,5));%theta
    %     AA_array(6,1) = Settle_value;
    %     AA_array(7,1) = NB_value;
    %     AA_array(8,1) = sum( IA_aa(:,9)); % epsilon
    %     AA_array(9,1) = IA_value;
    %
    %
    %     xlswrite(char(Liability_data), {'Pol by Pol AA'}, 'Attribution_Analysis', 'G2');
    %     xlswrite(char(Liability_data), AA_array, 'Attribution_Analysis', 'G3');
    
    
end

end

function [] = DoIAAsianCalc(FLAG_OPTIMIZE_RUNTIME, MktEnv_directory, MktEnv_filename, MktEnv_filename_bop, Policyholder_tape, Policyholder_tape_bop, Liability_data, Liability_data_BOP, Model_Name)

[SPX_level, SPX_date_num, SPX_current,current_date, bop_date, Term_t, Interest_rate_forward, Dividend_forward, Interest_rate_spot, aVolParams, IA_issue_date, IA_issue_date_string, ...
    IA_reset_date, IA_reset_date_string, IA_polnum, IA_premium, IA_participation_rate, IA_credit_cap, IA_credit_floor, IA_prorata_factor, IA_index_spread, number_of_policies, Num_of_Paths] = Read_Market_Policy_Data(MktEnv_directory, MktEnv_filename, Policyholder_tape, Model_Name);



[aAsianValue, aAsianDelta, aAsianVega, aAsianGamma, aAsianRho, aAsianTheta] = mAsianValuation(FLAG_OPTIMIZE_RUNTIME, aVolParams, Interest_rate_forward, Dividend_forward, Term_t, SPX_current, current_date, SPX_level, SPX_date_num, ...
    IA_reset_date, IA_credit_cap, IA_credit_floor, IA_premium, IA_participation_rate , IA_index_spread, Num_of_Paths, Model_Name);

%calculate the average interest rate. For simplicity, calculate the rho
%weighted spot rate. 

IA_spot_rate = zeros(size(IA_credit_floor));

for i = 1:number_of_policies
    Holidays_in_Between = holidays(current_date, datemnth(IA_reset_date(i), 12) );
    Time_to_maturity = wrkdydif(current_date, datemnth(IA_reset_date(i), 12), length(Holidays_in_Between))/252;
    IA_spot_rate(i) = interpolation_1d(Term_t, Interest_rate_spot, Time_to_maturity); 
end

Rate_weighted = sum(IA_spot_rate .* aAsianRho)/ sum(aAsianRho);

%recalculate the bop cliquet price based on EOP's bates parameters.

xlswrite(char(Liability_data), {'Pol_num'}, 'IA_Liabilities','A3');
xlswrite(char(Liability_data), IA_polnum, 'IA_Liabilities','A4');
xlswrite(char(Liability_data), {'issue_date'}, 'IA_Liabilities','B3');
xlswrite(char(Liability_data), IA_issue_date_string, 'IA_Liabilities','B4');
xlswrite(char(Liability_data), {'reset_date'}, 'IA_Liabilities','C3');
xlswrite(char(Liability_data), IA_reset_date_string, 'IA_Liabilities','C4');
xlswrite(char(Liability_data), {'premium'}, 'IA_Liabilities','D3');
xlswrite(char(Liability_data),IA_premium, 'IA_Liabilities','D4');
xlswrite(char(Liability_data), {'Cap'}, 'IA_Liabilities','E3');
xlswrite(char(Liability_data),IA_credit_cap, 'IA_Liabilities','E4');
xlswrite(char(Liability_data), {'Floor'}, 'IA_Liabilities','F3');
xlswrite(char(Liability_data),IA_credit_floor, 'IA_Liabilities', 'F4');
xlswrite(char(Liability_data), {'rate'}, 'IA_Liabilities','G3');
xlswrite(char(Liability_data),IA_spot_rate, 'IA_Liabilities', 'G4');

xlswrite(char(Liability_data), {'Asian Value'}, 'IA_Liabilities','J3');
xlswrite(char(Liability_data),aAsianValue, 'IA_Liabilities', 'J4');
xlswrite(char(Liability_data), {'Asian Delta'}, 'IA_Liabilities','K3');
xlswrite(char(Liability_data),aAsianDelta, 'IA_Liabilities', 'K4');
xlswrite(char(Liability_data), {'Asian Gamma'}, 'IA_Liabilities','L3');
xlswrite(char(Liability_data),aAsianGamma, 'IA_Liabilities', 'L4');
xlswrite(char(Liability_data), {'Asian Vega'}, 'IA_Liabilities','M3');
xlswrite(char(Liability_data),aAsianVega, 'IA_Liabilities', 'M4');
xlswrite(char(Liability_data), {'Asian Rho'}, 'IA_Liabilities','N3');
xlswrite(char(Liability_data),aAsianRho, 'IA_Liabilities', 'N4');
xlswrite(char(Liability_data), {'Asian Theta'}, 'IA_Liabilities','O3');
xlswrite(char(Liability_data),aAsianTheta, 'IA_Liabilities', 'O4');


xlswrite(char(Liability_data),{'Value'}, 'IA_Greeks', 'A1');
xlswrite(char(Liability_data), sum(aAsianValue), 'IA_Greeks', 'A2');
xlswrite(char(Liability_data),{'Delta'}, 'IA_Greeks', 'B1');
xlswrite(char(Liability_data), sum(aAsianDelta), 'IA_Greeks', 'B2');
xlswrite(char(Liability_data),{'Vega'}, 'IA_Greeks', 'C1');
xlswrite(char(Liability_data), sum(aAsianVega), 'IA_Greeks', 'C2');
xlswrite(char(Liability_data),{'Rho'}, 'IA_Greeks', 'D1');
xlswrite(char(Liability_data), sum(aAsianRho), 'IA_Greeks', 'D2');
xlswrite(char(Liability_data),{'Gamma'}, 'IA_Greeks', 'E1');
xlswrite(char(Liability_data), sum(aAsianGamma), 'IA_Greeks', 'E2');
xlswrite(char(Liability_data),{'Theta'}, 'IA_Greeks', 'F1');
xlswrite(char(Liability_data), sum(aAsianTheta), 'IA_Greeks', 'F2');
xlswrite(char(Liability_data),{'SPX'}, 'IA_Greeks', 'A5');
xlswrite(char(Liability_data), SPX_current, 'IA_Greeks', 'A6');
xlswrite(char(Liability_data),{'r'}, 'IA_Greeks', 'B5');
xlswrite(char(Liability_data), Rate_weighted, 'IA_Greeks', 'B6');


%======================================================================================================================================================================
%htc  If Results from previous business day exists AND we are not optimizing the run-time then
%htc    compute Asian Options' prev. day's value,
%htc    compute Asian Options' value change due to greeks from prev. business day
%htc    this is columns Q:V in 'IA_Liabilities' worksheet in workbook    F:\Index_Products\LiabilityValuation\AGL IA Asian\AGL.IA.Asian.Liability.Data.20140102.xls
%======================================================================================================================================================================
%sssss
if(   (exist([char(Liability_data_BOP)], 'file') ~= 0) &  (~FLAG_OPTIMIZE_RUNTIME)  )

    
    [SPX_level_bop, SPX_date_num_bop, SPX_current_bop,current_date_bop, bop_date_bop, Term_t_bop, Interest_rate_forward_bop, Dividend_forward_bop, Interest_rate_spot_bop, aVolParams_bop, IA_issue_date_bop, IA_issue_date_bop_string, ...
        IA_reset_date_bop, IA_reset_date_bop_string, IA_polnum_bop, IA_premium_bop, IA_participation_rate_bop, IA_credit_cap_bop, IA_credit_floor_bop, IA_prorata_factor_bop, IA_index_spread_bop, number_of_policies_bop, Num_of_Paths ] = Read_Market_Policy_Data(MktEnv_directory, MktEnv_filename_bop, Policyholder_tape_bop, Model_Name);
    
    [aAsianValue_bop_recalc] = mAsianValuation(FLAG_OPTIMIZE_RUNTIME, ...
        aVolParams, Interest_rate_forward_bop, Dividend_forward_bop, Term_t_bop, SPX_current_bop, current_date_bop, SPX_level_bop, SPX_date_num_bop, ...
        IA_reset_date_bop, IA_credit_cap_bop, IA_credit_floor_bop, IA_premium_bop, IA_participation_rate_bop, IA_index_spread_bop, Num_of_Paths, Model_Name );
    
    % first do policy by policy attribution analysis
    
    [data text] = xlsread(char(Liability_data_BOP), 'IA_liabilities', 'A4:O65536');
    
    IA_rate_bop = data(:, 4);
    IA_value_bop = data(:, 7);
    IA_delta_bop = data(:, 8);
    IA_gamma_bop = data(:, 9);
    IA_rho_bop = data(:, 11);
    IA_theta_bop = data(:, 12);
    
        
    IA_data_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'A6:B6');
    
    IA_aa = zeros(number_of_policies, 11);
    
    %BOP_Value, Delta Change, Rho Change, Vega Change, Gamma Change,
    %Speed Change, Theta Change,  New Business Ind,
    %Settlement Ind, Settle Value, epsilon, EOP Value
    
    
    % new the EOP policy number, find the correponding BOP policy
    % number
   
    % used in Theta attribution
    Holidays_in_between = holidays(bop_date, current_date);
    Buzdays_in_between = wrkdydif(bop_date, current_date, length(Holidays_in_between)) - 1;
    
    for i = 1 : number_of_policies
        IA_value_bop_by_pol = IA_value_bop(strcmp(text(:,1),IA_polnum(i)));
        
      
        if isempty(IA_value_bop_by_pol)
            % new business
            IA_aa(i,11) =  aAsianValue(i);
            IA_aa(i,7) = 1;
            
        else
            if size(IA_value_bop_by_pol) == 1
                
                
                IA_aa(i,1) = IA_value_bop_by_pol;
                %delta change
                IA_aa(i,2) = IA_delta_bop(strcmp(text(:,1),IA_polnum(i))) * (SPX_current - IA_data_bop(1));
                %Vega change
                IA_aa(i,3) = aAsianValue_bop_recalc(strcmp(text(:,1),IA_polnum(i))) - IA_value_bop_by_pol;
                %gamma change
                IA_aa(i,4) = IA_gamma_bop(strcmp(text(:,1),IA_polnum(i))) * power(SPX_current - IA_data_bop(1),2)/2;
                %rho change
                IA_aa(i,5) = IA_rho_bop(strcmp(text(:,1),IA_polnum(i))) * (IA_spot_rate(i) - IA_rate_bop(strcmp(text(:,1),IA_polnum(i)))) * 100;
                %theta change
                IA_aa(i,6) = IA_theta_bop(strcmp(text(:,1),IA_polnum(i))) * Buzdays_in_between /252;
                %New business Indicator.
                IA_aa(i,7) = (IA_reset_date(i) < NextBusinessDate(current_date)) && (IA_reset_date(i) >= current_date);
                %Settlement Indicator
                IA_aa(i,8) = (IA_reset_date(i) < NextBusinessDate(current_date)) && (IA_reset_date(i) >= current_date) && (IA_reset_date(i)~= IA_issue_date(i));
                %Settlement value/
                if IA_aa(i, 8) == 1
                     % add three more SPX level to handle the weekend
                     % settle or weekend and the next Monday is a holiday
                     % situation
                    SPX_level_appd = [SPX_level; SPX_level(end); SPX_level(end); SPX_level(end)];
                    SPX_date_num_appd = [SPX_date_num; SPX_date_num(end)+1; SPX_date_num(end)+2; SPX_date_num(end)+3];
                    
                    IA_aa(i,9) = AsianPrice(IA_reset_date(i), datenum(year(IA_reset_date(i))-1, month(IA_issue_date(i)), day(IA_issue_date(i))), SPX_level_appd, SPX_date_num_appd , IA_credit_cap(i), IA_credit_floor(i), IA_participation_rate(i), IA_index_spread(i), IA_premium_bop(i) , zeros(1,1),0 ) ;
                    %IA_aa(i,9) = AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_base, aRate ) * Policy_notional(i);
       
                
                end
                %EOP value
                IA_aa(i,11) = aAsianValue(i);
                %epsilon
                IA_aa(i,10) = IA_aa(i,11) - sum(IA_aa(i, 1:6))+ (IA_aa(i,9)-IA_aa(i,11)) * IA_aa(i,8);
                
            else
                % msg('duplication in policy num in the bopious run date');
            end
        end
        
    end
    
    xlswrite(char(Liability_data), {'bop_pol_value', 'delta_change', 'vega_change', 'gamma_change', 'rho_change', 'theta_change', 'NB Ind', 'Settlement Ind', 'Settlement Value', 'epsilon','eop_pol_value'}, 'IA_Liabilities','Q3');
    xlswrite(char(Liability_data), IA_aa, 'IA_Liabilities','Q4');
    
    xlswrite(char(Liability_data), {datestr(bop_date, 'mm/dd/yyyy')}, 'Attribution_Analysis', 'B2:C2');
    xlswrite(char(Liability_data), {datestr(current_date, 'mm/dd/yyyy')}, 'Attribution_Analysis', 'D2');
    xlswrite(char(Liability_data), {'BOP Value'; 'Delta'; 'Vega'; 'Rho'; 'Gamma'; 'Theta'; 'Settlement Value'; 'NB value'; 'Epsilon'; 'EOP Value'}, 'Attribution_Analysis', 'A3');
    
    IA_Greeks_bop = xlsread(char(Liability_data_BOP), 'IA_Greeks', 'A2:G2');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(1), 'Attribution_Analysis', 'E3');
    
    xlswrite(char(Liability_data), IA_Greeks_bop(2:end)', 'Attribution_Analysis', 'B4');
    
    xlswrite(char(Liability_data), IA_data_bop(1), 'Attribution_Analysis', 'C4');
    
    xlswrite(char(Liability_data), SPX_current, 'Attribution_Analysis', 'D4');
    
    xlswrite(char(Liability_data), IA_data_bop(2), 'Attribution_Analysis', 'C6');
    
    xlswrite(char(Liability_data), Rate_weighted, 'Attribution_Analysis', 'D6'); 
    
    
    Delta_change = IA_Greeks_bop(2) * (SPX_current - IA_data_bop(1));
    Vega_change = sum(IA_aa(:,3));
    Rho_change = IA_Greeks_bop(4) * (Rate_weighted - IA_data_bop(2))*100;
    Gamma_change = 0.5 * IA_Greeks_bop(5) * (SPX_current - IA_data_bop(1))^2;
    Theta_change = sum(IA_aa(:,6));
    IA_value =  sum(IA_aa(:,11));
    NB_value = sum(IA_aa(:, 11).*IA_aa(:, 7));
    Settle_value = sum(IA_aa(:, 9).*IA_aa(:,8));
    Epsilon = IA_value - IA_Greeks_bop(1) - Delta_change - Vega_change - Gamma_change - Rho_change - Theta_change + Settle_value - NB_value;
    
    xlswrite(char(Liability_data), Delta_change, 'Attribution_Analysis', 'E4');
    xlswrite(char(Liability_data), Vega_change, 'Attribution_Analysis', 'E5');
     xlswrite(char(Liability_data), Rho_change, 'Attribution_Analysis', 'E6');
    xlswrite(char(Liability_data), Gamma_change, 'Attribution_Analysis', 'E7');
    xlswrite(char(Liability_data), Theta_change, 'Attribution_Analysis', 'E8');
    xlswrite(char(Liability_data), Settle_value, 'Attribution_Analysis', 'E9');
    xlswrite(char(Liability_data), NB_value, 'Attribution_Analysis', 'E10');
    xlswrite(char(Liability_data), Epsilon, 'Attribution_Analysis', 'E11');
    xlswrite(char(Liability_data), IA_value, 'Attribution_Analysis', 'E12');
    
    %sum of the policy by policy attribution analysis,
    %     AA_array = zeros(9,1);
    %     AA_array(1,1) = IA_Greeks_bop(1);%bop
    %     AA_array(2,1) = sum( IA_aa(:,2));%delta
    %     AA_array(3,1) = sum( IA_aa(:,3));%vega
    %     AA_array(4,1) = sum( IA_aa(:,4));%gamma
    %     AA_array(5,1) = sum( IA_aa(:,5));%theta
    %     AA_array(6,1) = Settle_value;
    %     AA_array(7,1) = NB_value;
    %     AA_array(8,1) = sum( IA_aa(:,9)); % epsilon
    %     AA_array(9,1) = IA_value;
    %
    %
    %     xlswrite(char(Liability_data), {'Pol by Pol AA'}, 'Attribution_Analysis', 'G2');
    %     xlswrite(char(Liability_data), AA_array, 'Attribution_Analysis', 'G3');
    
    
end

end



%%
function [SPX_level, SPX_date_num, SPX_current,current_date, bop_date, Term_t, Interest_rate_forward, Dividend_forward, Interest_rate_spot, aVolParams, IA_issue_date, IA_issue_date_string,  ...
    IA_reset_date, IA_reset_date_string, IA_polnum, IA_premium, IA_participation_rate, IA_credit_cap, IA_credit_floor, IA_prorata_factor, IA_index_spread, number_of_policies, Num_of_Paths ] ...
= Read_Market_Policy_Data(MktEnv_directory, MktEnv_filename, Policyholder_tape, Model_Name)

%htc  store historic SPX_level and dates
[SPX_level SPX_date ] =xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'SPX', 'A3:B800');
SPX_date_num = datenum(SPX_date, 'mm/dd/yyyy');

%htc  store current SPX-level
SPX_current = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B3');

[dummy current_date_string] = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B1');

current_date = datenum(current_date_string, 'mm/dd/yyyy');

%htc  store previous hedging date
[dummy bop_date_string] = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B2');

bop_date = datenum(bop_date_string, 'mm/dd/yyyy');

Num_of_Paths = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input','B14');

%htc  store the 7 terms (up to 1 year time) (to option expirations), and
%htc  corresponding interest rate r and q
Term_t = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','A2:A8');
Interest_rate_spot = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','B2:B8');
Dividend = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Int&Dvd','C2:C8');


%htc  store the forward interest rate and dividend corresponding to Term_t
Interest_rate_forward = zeros(size(Interest_rate_spot));
Dividend_forward = zeros(size(Dividend));

Interest_rate_forward(1,1) = Interest_rate_spot(1,1);
Dividend_forward(1,1) = Dividend(1,1);

Interest_rate_forward(2:end, 1) = (Interest_rate_spot(2:end, 1) .* Term_t(2:end, 1) - Interest_rate_spot(1:end-1, 1) .* Term_t(1:end-1, 1)) ./ (Term_t(2:end, 1) - Term_t(1:end-1,1));
Dividend_forward(2:end, 1) = (Dividend(2:end, 1) .* Term_t(2:end, 1) - Dividend(1:end-1, 1) .* Term_t(1:end-1, 1)) ./ (Term_t(2:end, 1) - Term_t(1:end-1,1));

%htc  read Bates or Heston parameters into 'aVolParams'
if strcmp(Model_Name, 'Bates')
    aVolParams = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Bates Params','D3:D10');
else
    if strcmp(Model_Name, 'Heston')
        aVolParams = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Heston Params','D3:D7');
    else
        error('need to be either Bates or Heston model');
    end
end


%htc  IA_raw_text   holds all policies A2:C65535
%htc  IA_raw_data   holds all policies D2:J65535
%htc  read all policies for a company/block such as  F:\Index_Products\PolicyHolderData\AGL IA Cliquet\AGL_IA_PolicyholderData_Cliquet_20140102.xls
%htc  obtain number of policies, policy issue date
[IA_raw_data IA_raw_text] = xlsread(char(Policyholder_tape), 'Sheet1', 'A2:J65535');


[number_of_policies dummy] = size(IA_raw_data);

IA_issue_date = datenum(IA_raw_text(:, 1), 'mm/dd/yyyy');
IA_reset_date = zeros(size(IA_issue_date));

IA_issue_date_string = cell(size(IA_issue_date));
IA_reset_date_string = cell(size(IA_issue_date));

Next_Buz_Date = NextBusinessDate(current_date);

% handle the case that option reset happen in the weekend.
for i = 1 : number_of_policies
    if datenum(year(current_date), month(IA_issue_date(i)), day(IA_issue_date(i))) >= Next_Buz_Date
        IA_maturity_date(i) =  datenum(year(current_date), month(IA_issue_date(i)), day(IA_issue_date(i)));
        IA_reset_date(i) =  datenum(year(current_date)-1, month(IA_issue_date(i)), day(IA_issue_date(i)));
    else
        IA_maturity_date(i) = datenum(year(current_date)+1, month(IA_issue_date(i)), day(IA_issue_date(i)));
        IA_reset_date(i) =  datenum(year(current_date), month(IA_issue_date(i)), day(IA_issue_date(i)));
    end
    
    IA_issue_date_string(i) = {datestr(IA_issue_date(i), 'mm/dd/yyyy')};
    
    IA_reset_date_string(i) = {datestr(IA_reset_date(i), 'mm/dd/yyyy')};
end


IA_polnum = IA_raw_text(:, 3);
IA_premium = IA_raw_data(:, 1) .* IA_raw_data(:,5); %premium amount * prorata factor
IA_participation_rate = IA_raw_data(:, 2);
IA_credit_cap = IA_raw_data(:, 3);
IA_credit_floor = IA_raw_data(:, 4);
IA_prorata_factor = IA_raw_data(:, 5);
IA_index_spread = IA_raw_data(:, 6);
end


%=========================================================================
%mCliquetValuation()
%Note:  flag 'FLAG_OPTIMIZE_RUNTIME' true will speed up the run-time,
%       by only computing Cliequent Options' value and delta
%       and skip computing (vega, gamma, rho, theta) to shorten the run-time
%=========================================================================
function [aCliquetValue, aCliquetDelta, aCliquetVega, aCliquetGamma, aCliquetRho, aCliquetTheta] = mCliquetValuation(FLAG_OPTIMIZE_RUNTIME,iParameters, Interest_rate_forward, Dividend_forward, t, s0, CurrentDate, SPX_Historic_level, SPX_date_num, Policy_reset_date, Policy_cap, Policy_floor, Policy_notional, Num_of_Paths, Model_Name)

%htc set time to 1.25 year and time axis ATimeAxis
%htc where delta t is 1/252 year   aTimeDiff
aProjectionYears = 1.25;

aNumPaths = Num_of_Paths;
aTimeDiff = 1/252;
aNumTimeStep = aProjectionYears/aTimeDiff;

aTimeAxis =  (aTimeDiff : aTimeDiff : aNumTimeStep*aTimeDiff);


rng(2012);


%htc setup forward r aRate[]
%htc       forward q aDividend[]
%prepare forward r and q for monte carlo simulation
aNumoft = length(t);

aRate = zeros(size(aTimeAxis));
aDividend = zeros(size(aTimeAxis));

aIndicator = (aTimeAxis<=t(1));
aRate(aIndicator) = Interest_rate_forward(1);
aDividend(aIndicator) = Dividend_forward(1);

for i =2 : aNumoft
    aIndicator = (aTimeAxis>t(i-1)) .* (aTimeAxis<=t(i));
    aRate(aIndicator == 1) = Interest_rate_forward(i);
    aDividend(aIndicator == 1) = Dividend_forward(i);
end

aIndicator = (aTimeAxis>t(end));
aRate(aIndicator) = Interest_rate_forward(end);
aDividend(aIndicator) = Dividend_forward(end);

%htc setup forward (r-q) aYieldBase[]
aYieldBase = aRate - aDividend;


%htc Setup Heston/Bates Parameters
%htc   Theta is long-term variance
%htc   Sigma is volatility of volatility
%htc   aVolParam.[...]  are the 1x315 (i.e. time step) Bates Parameters
%htc   aVolParamUp.[...]  are the 1x315 (i.e. time step) Bates Parameters shock up sigma by 0.5%
%htc   aVolParamDown.[...]  are the 1x315 (i.e. time step) Bates Parameters shock down sigma by 0.5%
aVolParam.VarZero = iParameters(5);
aVolParam.Theta = [repmat(iParameters(2), 1, aNumTimeStep)];               %htc 1xaNumTimeStep row vector of value Theta
aVolParam.MeanReversionRate = [repmat(iParameters(1),1, aNumTimeStep)];
aVolParam.Sigma = [repmat(iParameters(3),1, aNumTimeStep)];

%htc shock up & shock down Heston/Bates Parameters by 0.5%
% Volatility Parameters with the shock
aVolParamUp.VarZero = (sqrt(iParameters(5))+0.005)^2;
aVolParamUp.Theta = [repmat((sqrt(iParameters(2))+0.005)^2, 1,aProjectionYears/aTimeDiff)];
aVolParamUp.MeanReversionRate = [repmat(iParameters(1),1,aProjectionYears/aTimeDiff)];
aVolParamUp.Sigma = [repmat(iParameters(3),1,aProjectionYears/aTimeDiff)];

aVolParamDown.VarZero = (sqrt(iParameters(5))-0.005)^2;
aVolParamDown.Theta = [repmat((sqrt(iParameters(2))-0.005)^2, 1,aProjectionYears/aTimeDiff)];
aVolParamDown.MeanReversionRate = [repmat(iParameters(1),1,aProjectionYears/aTimeDiff)];
aVolParamDown.Sigma = [repmat(iParameters(3),1,aProjectionYears/aTimeDiff)];

%htc  2x2 correlation matrix between stock & volatility
aCorrMatrix = [1, iParameters(4); iParameters(4), 1];

%htc 2x2 Cholesky Decomposition matrix
%htc for converting 1x2 normally distributed independent random #s ==> 1x2 normally
%htc distributed random numbers that are correlated by rho.
aCholMatrixHeston = chol(aCorrMatrix);

%htc  315x2x50000 (time,2,paths) array of normally distributed random #s
%htc  at every time step, there are 2 independent random #s
aRandn = randn(length(aTimeAxis), 2, aNumPaths);

%htc 315x2x50000 (time, 2, paths)   at every time step, there are 2 correlated random #s.
for aIdx = 1:aNumPaths
    aRandn(1:aNumTimeStep, :, aIdx) = aRandn(1:aNumTimeStep, :, aIdx) * aCholMatrixHeston;
end

%htc  the 315x50000 (time steps, paths) random numbers aEquityRandn and aVarianceRandn
%htc  are correlated by rho
aEquityRandn = reshape(aRandn(:,1, :), [length(aTimeAxis), aNumPaths]);
aVarianceRandn = reshape(aRandn(:, 2, :), [length(aTimeAxis), aNumPaths]);


%htc  setup Bates parameters and create Bates Models
%htc          vol shock up & down by 0.5%
%htc          interest rate r shocked up & down 0.5%
%htc  OR
%htc  Create Heston Models vol shocked up & down by 0.5%
%htc         interest rate r shocked up & down by 0.5%
if strcmp(Model_Name, 'Bates')
    
    aVolParam.JumpIntensity = iParameters(6);
    aVolParam.JumpSizeMean = iParameters(7);
    aVolParam.JumpSizeStd = iParameters(8);
    
    aVolParamUp.JumpIntensity = iParameters(6);
    aVolParamUp.JumpSizeMean = iParameters(7);
    aVolParamUp.JumpSizeStd = iParameters(8);
    
    aVolParamDown.JumpIntensity = iParameters(6);
    aVolParamDown.JumpSizeMean = iParameters(7);
    aVolParamDown.JumpSizeStd = iParameters(8);
    
    
    aJumpSizeRandn = randn(length(aTimeAxis), aNumPaths);
    aJumpRandUnif = random('unif', 0, 1, [length(aTimeAxis), aNumPaths]);       %htc 315x50000 (time step, paths) random # uniformly distributed 0-->1
    aJumpRandp = poissinv(aJumpRandUnif,aVolParam.JumpIntensity * aTimeDiff);
    
    aEquityModel = Bates(aYieldBase, aVolParam, 1);
    aEquityModelVolUp = Bates(aYieldBase, aVolParamUp, 1);
    aEquityModelVolDown = Bates(aYieldBase, aVolParamDown, 1);
    aEquityModelRateUp = Bates(aYieldBase+0.005, aVolParam, 1);
    aEquityModelRateDown = Bates(aYieldBase-0.005, aVolParam, 1);
    
else
    if strcmp(Model_Name, 'Heston')
        aEquityModel = heston(aYieldBase, aVolParam, 1);
        aEquityModelVolUp = heston(aYieldBase, aVolParamUp, 1);
        aEquityModelVolDown = heston(aYieldBase, aVolParamDown, 1);
        aEquityModelRateUp = heston(aYieldBase+0.005, aVolParam, 1);
        aEquityModelRateDown = heston(aYieldBase-0.005, aVolParam, 1);
    else
        error('need to be either Bates or Heston model');
    end
end

aEquityPaths = zeros(aNumPaths, aProjectionYears/aTimeDiff);

%htc  due to Matlab memory constraints, limit the paths to 5,000
%htc  spread into 'aScenSets' batches (i.e. 10 batches)
aSetLength = 5000;
aScenSets = ceil(aNumPaths/aSetLength);


%htc =========================================================
%htc below compute Cliquet Options' value, delta, gamma, theta
%htc =========================================================

%htc  Generate the 50,000 x 315 (paths,time steps) Equity paths 'aEquityPaths'
%htc  according to Bates or Heston Model
%htc  due to memory constraints, we are doing this in 10 batches below (i.e. aScenSets=10)
% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        [aEquityPathsTemp, aVolPaths] = MakePaths(aEquityModel, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsTemp, aVolPaths] = MakePaths(aEquityModel, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        else
            error('need to be either Bates or Heston model');
        end
    end
    %     S_temp = s0 * exp(cumsum(aEquityPaths, 2));
    %
    %     S_temp = [repmat(s0, min(aSetLength, aNumPaths - (i-1)*aSetLength),1) S_temp];
    %
    %     S((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = S_temp;
    
    aEquityPaths((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsTemp;
    
    clear('aEquityPathsTemp', 'aVolPaths');
    
end

%htc  shock equity path up & down by 5 and by 0.5
aS_base = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPaths, 2))];
aS_up_5 = [repmat((s0), aNumPaths,1) (s0+5)*exp(cumsum(aEquityPaths, 2))];
aS_down_5 = [repmat((s0), aNumPaths,1) (s0-5)*exp(cumsum(aEquityPaths, 2))];
aS_up_half = [repmat((s0), aNumPaths,1) (s0+0.5)*exp(cumsum(aEquityPaths, 2))];
aS_down_half = [repmat((s0), aNumPaths,1) (s0-0.5)*exp(cumsum(aEquityPaths, 2))];


%assuming market doesn't move after one day
%htc  SPX_Historic_level  is historic SPX closes
%htc  SPX_Historic_level_theta is historic SPX closes, adding one more day
%htc       (to next business day) and assumed the SPX-level remains the same
NextBuzDate = NextBusinessDate(CurrentDate);
DaystoNextBuzDate =  NextBuzDate - CurrentDate;
SPX_Historic_level_theta = [SPX_Historic_level; ones(DaystoNextBuzDate,1) * s0];
SPX_date_num_theta = [SPX_date_num; CurrentDate+(1:1:DaystoNextBuzDate)'];


%htc  'num_of_options' is the number of policies in a company/hedging block
[num_of_options dummy] = size(Policy_cap);
aCliquetValue = zeros(num_of_options, 1);
aCliquetDelta = zeros(num_of_options, 1);
aCliquetGamma = zeros(num_of_options, 1);
aCliquetRho = zeros(num_of_options, 1);
aCliquetTheta = zeros(num_of_options, 1);
aCliquetVega = zeros(num_of_options, 1);

%htc  loop over # of policies in a company/hedging block
%htc  computes the Cliquet option's value, delta, gamma, theta
for i=1:num_of_options
   
        %htc obtain Cliquet option expiration dates, by shifting the
        %htc policy's reset dates by 12 months forward.
    Option_end_date = datemnth(Policy_reset_date(i), 12);   
    
    %handle the situation that reset happen during the weekend
    if(Option_end_date >=CurrentDate && Policy_reset_date(i) <NextBuzDate)
        aCliquetValue(i) = CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_base, aRate ) * Policy_notional(i);
        aCliquetDelta(i) = CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_up_half, aRate ) ...
                                     * Policy_notional(i) ...
                         - CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_down_half, aRate ) ...
                                     * Policy_notional(i);
                                 
    %sssss   htc skip below if we wish to save run-time.
        if ~FLAG_OPTIMIZE_RUNTIME
              aCliquetGamma(i) = (CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_up_5, aRate ) * Policy_notional(i) ...
                      + CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_down_5, aRate ) * Policy_notional(i) ...
                      - 2*CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_base, aRate ) * Policy_notional(i))/25;
              % normalize the daily theta to yearly theta. Since theta is
              % calculated by business date, multiple the results by 252. 
              aCliquetTheta(i) = (CliquetPrice(NextBuzDate, Policy_reset_date(i), SPX_Historic_level_theta, SPX_date_num_theta, Policy_cap(i), Policy_floor(i), 1, aS_base, aRate ) * Policy_notional(i) ...
                      - aCliquetValue(i))*252;
        end
    end
end

%htc  clear all the equity paths
clear('aS_base', 'aS_up_5', 'aS_down_5', 'aS_up_half', 'aS_down_half', 'aEquityPaths');






%htc ==============================================
%htc below compute Cliquet Options' Vega
%htc ==============================================
%sssss   htc skip below if we wish to save run-time.
if ~FLAG_OPTIMIZE_RUNTIME

aEquityPathsVolUp = zeros(aNumPaths, aProjectionYears/aTimeDiff);
aEquityPathsVolDown = zeros(aNumPaths, aProjectionYears/aTimeDiff);

aSetLength = 5000;

aScenSets = ceil(aNumPaths/aSetLength);

%htc  Generate the 50,000 x 315 (paths,time steps) Equity paths 'aEquityPathsVolUp' & 'aEquityPathsVolDown'
%htc  (i.e. shocked sigma up & down by 0.5%) according to Bates or Heston Model
%htc  due to memory constraints, we are doing this in 10 batches below (i.e. aScenSets=10)
% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        
        [aEquityPathsVolUpTemp, aVolUpPaths] = MakePaths(aEquityModelVolUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
        [aEquityPathsVolDownTemp, aVolDownPaths] = MakePaths(aEquityModelVolDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsVolUpTemp, aVolUpPaths] = MakePaths(aEquityModelVolUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            [aEquityPathsVolDownTemp, aVolDownPaths] = MakePaths(aEquityModelVolDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            
        else
            error('Need to be either Bates or Heston model');
        end
    end
    
    
    aEquityPathsVolUp((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsVolUpTemp;
    aEquityPathsVolDown((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsVolDownTemp;
    
    clear('aEquityPathsVolUpTemp', 'aEquityPathsVolDownTemp', 'aVolUpPaths', 'aVolDownPaths');
    
end



%htc  equity paths with sigma shocked up and down by 0.5%
aS_vol_up = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsVolUp, 2))];
aS_vol_down = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsVolDown, 2))];

%htc  loop over # of policies in a company/hedging block
%htc  computes the Cliquet option's Vega
    for i=1:num_of_options
    
        Option_end_date = datemnth(Policy_reset_date(i), 12);
    
        if(Option_end_date >=CurrentDate && Policy_reset_date(i) < NextBuzDate)
            aCliquetVega(i) = CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_vol_up, aRate ) * Policy_notional(i) ...
                - CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_vol_down, aRate ) * Policy_notional(i);
        end
    end
    
end      %if ~FLAG_OPTIMIZE_RUNTIME   for computing Cliquet Option's Vega



%htc ==============================================
%htc below compute Cliquet Options' Rho
%htc ==============================================
%sssss   htc skip below if we wish to save run-time.
if ~FLAG_OPTIMIZE_RUNTIME

aEquityPathsRateUp = zeros(aNumPaths, aProjectionYears/aTimeDiff);
aEquityPathsRateDown = zeros(aNumPaths, aProjectionYears/aTimeDiff);

aSetLength = 5000;

aScenSets = ceil(aNumPaths/aSetLength);

% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        
        [aEquityPathsRateUpTemp, aVolPaths] = MakePaths(aEquityModelRateUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
        [aEquityPathsRateDownTemp, aVolPaths] = MakePaths(aEquityModelRateDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsRateUpTemp, aVolUpPaths] = MakePaths(aEquityModelRateUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            [aEquityPathsRateDownTemp, aVolDownPaths] = MakePaths(aEquityModelRateDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            
        else
            error('Need to be either Bates or Heston model');
        end
    end
    
    
    aEquityPathsRateUp((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsRateUpTemp;
    aEquityPathsRateDown((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsRateDownTemp;
    
    clear('aEquityPathsRateUpTemp', 'aEquityPathsRateDownTemp', 'aVolPaths', 'aVolPaths');
    
end

aS_rate_up = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsRateUp, 2))];
aS_rate_down = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsRateDown, 2))];


%htc  loop over # of policies in a company/hedging block
%htc  computes the Cliquet option's Rho
    for i=1:num_of_options
    
        Option_end_date = datemnth(Policy_reset_date(i), 12);
    
        if(Option_end_date >=CurrentDate && Policy_reset_date(i) < NextBuzDate)
            aCliquetRho(i) = CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_rate_up, aRate + 0.005 ) * Policy_notional(i) ...
                - CliquetPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), 1, aS_rate_down, aRate - 0.005 ) * Policy_notional(i);
        end
    end
    
end     %if ~FLAG_OPTIMIZE_RUNTIME    compute Cliquet Option's Rho
end











function [Payoff] = CliquetPrice(Current_Date, Option_Start_Date, S_Prev, S_Prev_Dates , Cap, Floor, Notional, iS, Interest_rate )

Month_Ind = (1:1:12);
Monthiversary = datemnth(Option_Start_Date, Month_Ind);
Monthiversary = [Option_Start_Date, Monthiversary];

Observ_Dates = Monthiversary(Monthiversary > Current_Date);

Past_Observ_Dates = Monthiversary(Monthiversary <= Current_Date);

num_past_observ_dates = length(Past_Observ_Dates);
Past_S = zeros(num_past_observ_dates, 1);

for i = 1:num_past_observ_dates
    Past_S(i) = S_Prev(S_Prev_Dates == Past_Observ_Dates(i));
end

Days_To_Observs = zeros(size(Observ_Dates));

for i = 1:length(Observ_Dates)
    Holidays_in_Between = holidays(Current_Date, Observ_Dates(1,i));
    Days_To_Observs(1, i) = wrkdydif(Current_Date, Observ_Dates(1, i), length(Holidays_in_Between));
    
end

if isempty(Days_To_Observs)
    aDiscFactor = 1;
else
    aDiscFactor = exp(-sum(Interest_rate(1:Days_To_Observs(1,end)))/252);
end

[m, n] = size(iS);
Number_of_Paths = m;

S_at_Observ_Dates = [repmat(Past_S', Number_of_Paths, 1), iS(:, Days_To_Observs)];

Capped_Monthly_Return = min((S_at_Observ_Dates(:, 2:end) ./ S_at_Observ_Dates(:, 1:end-1)) - 1, Cap);

Floored_Yearly_Return = max(Floor, sum(Capped_Monthly_Return, 2));

Payoff = Notional * mean(Floored_Yearly_Return) * aDiscFactor;

end

function [Next_Business_Date] = NextBusinessDate(Current_Date )

i = 1;
Holiday_in_Between = holidays(Current_Date, Current_Date + i);

while wrkdydif(Current_Date, Current_Date + i, length(Holiday_in_Between))<2
    i = i+1;
end

Next_Business_Date = Current_Date + i;

end

%=========================================================================
%mAsianValuation()
%Note:  flag 'FLAG_OPTIMIZE_RUNTIME' true will speed up the run-time,
%       by only computing Asian Options' value and delta
%       and skip computing (vega, gamma, rho, theta) to shorten the run-time
%=========================================================================
function [aAsianValue, aAsianDelta, aAsianVega, aAsianGamma, aAsianRho, aAsianTheta] = mAsianValuation(FLAG_OPTIMIZE_RUNTIME, ...
                                        iParameters, Interest_rate_forward, Dividend_forward, ...
                                        t, s0, CurrentDate, SPX_Historic_level, SPX_date_num, ...
                                        Policy_reset_date, Policy_cap, Policy_floor, Policy_notional, Policy_participation_rate, Policy_index_spread, Num_of_Paths, Model_Name)

aProjectionYears = 1.25;

aNumPaths = Num_of_Paths;
aTimeDiff = 1/252;
aNumTimeStep = aProjectionYears/aTimeDiff;

aTimeAxis =  (aTimeDiff : aTimeDiff : aNumTimeStep*aTimeDiff);

rng(2012);


%prepare forward r and q for monte carlo simulation
aNumoft = length(t);

aRate = zeros(size(aTimeAxis));
aDividend = zeros(size(aTimeAxis));

aIndicator = (aTimeAxis<=t(1));
aRate(aIndicator) = Interest_rate_forward(1);
aDividend(aIndicator) = Dividend_forward(1);

for i =2 : aNumoft
    aIndicator = (aTimeAxis>t(i-1)) .* (aTimeAxis<=t(i));
    aRate(aIndicator == 1) = Interest_rate_forward(i);
    aDividend(aIndicator == 1) = Dividend_forward(i);
end

aIndicator = (aTimeAxis>t(end));
aRate(aIndicator) = Interest_rate_forward(end);
aDividend(aIndicator) = Dividend_forward(end);


aYieldBase = aRate - aDividend;


% Volatility Parameters
aVolParam.VarZero = iParameters(5);
aVolParam.Theta = [repmat(iParameters(2), 1, aNumTimeStep)];
aVolParam.MeanReversionRate = [repmat(iParameters(1),1, aNumTimeStep)];
aVolParam.Sigma = [repmat(iParameters(3),1, aNumTimeStep)];

% Volatility Parameters with the shock
aVolParamUp.VarZero = (sqrt(iParameters(5))+0.005)^2;
aVolParamUp.Theta = [repmat((sqrt(iParameters(2))+0.005)^2, 1,aProjectionYears/aTimeDiff)];
aVolParamUp.MeanReversionRate = [repmat(iParameters(1),1,aProjectionYears/aTimeDiff)];
aVolParamUp.Sigma = [repmat(iParameters(3),1,aProjectionYears/aTimeDiff)];

aVolParamDown.VarZero = (sqrt(iParameters(5))-0.005)^2;
aVolParamDown.Theta = [repmat((sqrt(iParameters(2))-0.005)^2, 1,aProjectionYears/aTimeDiff)];
aVolParamDown.MeanReversionRate = [repmat(iParameters(1),1,aProjectionYears/aTimeDiff)];
aVolParamDown.Sigma = [repmat(iParameters(3),1,aProjectionYears/aTimeDiff)];

aCorrMatrix = [1, iParameters(4); iParameters(4), 1];

aCholMatrixHeston = chol(aCorrMatrix);

aRandn = randn(length(aTimeAxis), 2, aNumPaths);

for aIdx = 1:aNumPaths
    aRandn(1:aNumTimeStep, :, aIdx) = aRandn(1:aNumTimeStep, :, aIdx) * aCholMatrixHeston;
end

aEquityRandn = reshape(aRandn(:,1, :), [length(aTimeAxis), aNumPaths]);
aVarianceRandn = reshape(aRandn(:, 2, :), [length(aTimeAxis), aNumPaths]);

if strcmp(Model_Name, 'Bates')
    
    aVolParam.JumpIntensity = iParameters(6);
    aVolParam.JumpSizeMean = iParameters(7);
    aVolParam.JumpSizeStd = iParameters(8);
    
    aVolParamUp.JumpIntensity = iParameters(6);
    aVolParamUp.JumpSizeMean = iParameters(7);
    aVolParamUp.JumpSizeStd = iParameters(8);
    
    aVolParamDown.JumpIntensity = iParameters(6);
    aVolParamDown.JumpSizeMean = iParameters(7);
    aVolParamDown.JumpSizeStd = iParameters(8);
    
    
    aJumpSizeRandn = randn(length(aTimeAxis), aNumPaths);
    aJumpRandUnif = random('unif', 0, 1, [length(aTimeAxis), aNumPaths]);
    aJumpRandp = poissinv(aJumpRandUnif,aVolParam.JumpIntensity * aTimeDiff);
    
    aEquityModel = Bates(aYieldBase, aVolParam, 1);
    aEquityModelVolUp = Bates(aYieldBase, aVolParamUp, 1);
    aEquityModelVolDown = Bates(aYieldBase, aVolParamDown, 1);
    aEquityModelRateUp = Bates(aYieldBase+0.005, aVolParam, 1);
    aEquityModelRateDown = Bates(aYieldBase-0.005, aVolParam, 1);
    
else
    if strcmp(Model_Name, 'Heston')
        aEquityModel = heston(aYieldBase, aVolParam, 1);
        aEquityModelVolUp = heston(aYieldBase, aVolParamUp, 1);
        aEquityModelVolDown = heston(aYieldBase, aVolParamDown, 1);
        aEquityModelRateUp = heston(aYieldBase+0.005, aVolParam, 1);
        aEquityModelRateDown = heston(aYieldBase-0.005, aVolParam, 1);
    else
        error('need to be either Bates or Heston model');
    end
end

aEquityPaths = zeros(aNumPaths, aProjectionYears/aTimeDiff);

aSetLength = 5000;

aScenSets = ceil(aNumPaths/aSetLength);

% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        [aEquityPathsTemp, aVolPaths] = MakePaths(aEquityModel, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsTemp, aVolPaths] = MakePaths(aEquityModel, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        else
            error('need to be either Bates or Heston model');
        end
    end
    %     S_temp = s0 * exp(cumsum(aEquityPaths, 2));
    %
    %     S_temp = [repmat(s0, min(aSetLength, aNumPaths - (i-1)*aSetLength),1) S_temp];
    %
    %     S((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = S_temp;
    
    aEquityPaths((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsTemp;
    
    clear('aEquityPathsTemp', 'aVolPaths');
    
end

aS_base = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPaths, 2))];
aS_up_5 = [repmat((s0), aNumPaths,1) (s0+5)*exp(cumsum(aEquityPaths, 2))];
aS_down_5 = [repmat((s0), aNumPaths,1) (s0-5)*exp(cumsum(aEquityPaths, 2))];
aS_up_half = [repmat((s0), aNumPaths,1) (s0+0.5)*exp(cumsum(aEquityPaths, 2))];
aS_down_half = [repmat((s0), aNumPaths,1) (s0-0.5)*exp(cumsum(aEquityPaths, 2))];

%assuming market doesn't move after one day
NextBuzDate = NextBusinessDate(CurrentDate);
DaystoNextBuzDate =  NextBuzDate - CurrentDate;
SPX_Historic_level_theta = [SPX_Historic_level; ones(DaystoNextBuzDate,1) * s0];
SPX_date_num_theta = [SPX_date_num; CurrentDate+(1:1:DaystoNextBuzDate)'];

[num_of_options dummy] = size(Policy_cap);
aAsianValue = zeros(num_of_options, 1);
aAsianDelta = zeros(num_of_options, 1);
aAsianGamma = zeros(num_of_options, 1);
aAsianRho = zeros(num_of_options, 1);
aAsianTheta = zeros(num_of_options, 1);
aAsianVega = zeros(num_of_options, 1);

for i=1:num_of_options

    Option_end_date = datemnth(Policy_reset_date(i), 12);
    
    %handle the situation that reset happen during the weekend
    if(Option_end_date >=CurrentDate && Policy_reset_date(i) <NextBuzDate)
        aAsianValue(i) = AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_base, aRate ) * Policy_notional(i);
        aAsianDelta(i) = AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_up_half, aRate ) * Policy_notional(i) ...
            - AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_down_half, aRate ) * Policy_notional(i);
 
    %sssss   htc skip below if we wish to save run-time.
        if ~FLAG_OPTIMIZE_RUNTIME
            aAsianGamma(i) = (AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_up_5, aRate ) * Policy_notional(i) ...
                + AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_down_5, aRate ) * Policy_notional(i) ...
                - 2*AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i),  Policy_index_spread(i), 1, aS_base, aRate ) * Policy_notional(i))/25;
            % normalize the daily theta to yearly theta. Since theta is
            % calculated by business date, multiple the results by 252. 
            aAsianTheta(i) = (AsianPrice(NextBuzDate, Policy_reset_date(i), SPX_Historic_level_theta, SPX_date_num_theta, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_base, aRate ) * Policy_notional(i) ...
                - aAsianValue(i))*252;
        end
    end
end

clear('aS_base', 'aS_up_5', 'aS_down_5', 'aS_up_half', 'aS_down_half', 'aEquityPaths');





%htc ==============================================
%htc below compute Asian Options' Vega
%htc ==============================================
%sssss   htc skip below if we wish to save run-time.
if ~FLAG_OPTIMIZE_RUNTIME

aEquityPathsVolUp = zeros(aNumPaths, aProjectionYears/aTimeDiff);
aEquityPathsVolDown = zeros(aNumPaths, aProjectionYears/aTimeDiff);

aSetLength = 5000;

aScenSets = ceil(aNumPaths/aSetLength);

% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        
        [aEquityPathsVolUpTemp, aVolUpPaths] = MakePaths(aEquityModelVolUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
        [aEquityPathsVolDownTemp, aVolDownPaths] = MakePaths(aEquityModelVolDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsVolUpTemp, aVolUpPaths] = MakePaths(aEquityModelVolUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            [aEquityPathsVolDownTemp, aVolDownPaths] = MakePaths(aEquityModelVolDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            
        else
            error('Need to be either Bates or Heston model');
        end
    end
    
    
    aEquityPathsVolUp((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsVolUpTemp;
    aEquityPathsVolDown((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsVolDownTemp;
    
    clear('aEquityPathsVolUpTemp', 'aEquityPathsVolDownTemp', 'aVolUpPaths', 'aVolDownPaths');
    
end

aS_vol_up = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsVolUp, 2))];
aS_vol_down = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsVolDown, 2))];

for i=1:num_of_options
    
    Option_end_date = datemnth(Policy_reset_date(i), 12);
    
    if(Option_end_date >=CurrentDate && Policy_reset_date(i) < NextBuzDate)
        aAsianVega(i) = AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i),Policy_participation_rate(i), Policy_index_spread(i), 1, aS_vol_up, aRate ) * Policy_notional(i) ...
            - AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_vol_down, aRate ) * Policy_notional(i);
    end
end

end      %if ~FLAG_OPTIMIZE_RUNTIME   for computing Asian Option's Vega



%htc ==============================================
%htc below compute Asian Options' Rho
%htc ==============================================
%sssss   htc skip below if we wish to save run-time.
if ~FLAG_OPTIMIZE_RUNTIME

aEquityPathsRateUp = zeros(aNumPaths, aProjectionYears/aTimeDiff);
aEquityPathsRateDown = zeros(aNumPaths, aProjectionYears/aTimeDiff);

aSetLength = 5000;

aScenSets = ceil(aNumPaths/aSetLength);

% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    
    if strcmp(Model_Name, 'Bates')
        
        [aEquityPathsRateUpTemp, aVolPaths] = MakePaths(aEquityModelRateUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
        [aEquityPathsRateDownTemp, aVolPaths] = MakePaths(aEquityModelRateDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))'...
            , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
        
    else
        if strcmp(Model_Name, 'Heston')
            [aEquityPathsRateUpTemp, aVolUpPaths] = MakePaths(aEquityModelRateUp, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            [aEquityPathsRateDownTemp, aVolDownPaths] = MakePaths(aEquityModelRateDown, aTimeAxis, min(aSetLength, aNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, aNumPaths))');
            
            
        else
            error('Need to be either Bates or Heston model');
        end
    end
    
    
    aEquityPathsRateUp((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsRateUpTemp;
    aEquityPathsRateDown((i-1)*aSetLength+1:min(i*aSetLength, aNumPaths), :) = aEquityPathsRateDownTemp;
    
    clear('aEquityPathsRateUpTemp', 'aEquityPathsRateDownTemp', 'aVolPaths', 'aVolPaths');
    
end

aS_rate_up = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsRateUp, 2))];
aS_rate_down = [repmat(s0, aNumPaths,1) s0*exp(cumsum(aEquityPathsRateDown, 2))];

for i=1:num_of_options
    
    Option_end_date = datemnth(Policy_reset_date(i), 12);
    
    if(Option_end_date >=CurrentDate && Policy_reset_date(i) < NextBuzDate)
        aAsianRho(i) = AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_rate_up, aRate + 0.005 ) * Policy_notional(i) ...
            - AsianPrice(CurrentDate, Policy_reset_date(i), SPX_Historic_level, SPX_date_num, Policy_cap(i), Policy_floor(i), Policy_participation_rate(i), Policy_index_spread(i), 1, aS_rate_down, aRate - 0.005 ) * Policy_notional(i);
    end
end
end     %if ~FLAG_OPTIMIZE_RUNTIME    compute Asian Option's Rho

end


function [Payoff] = AsianPrice(Current_Date, Option_Start_Date, S_Prev, S_Prev_Dates , Cap, Floor, Participation_rate, Index_spread, Notional, iS, Interest_rate )

Month_Ind = (1:1:12);
Monthiversary = datemnth(Option_Start_Date, Month_Ind);
Monthiversary = [Option_Start_Date, Monthiversary];

Observ_Dates = Monthiversary(Monthiversary > Current_Date);

Past_Observ_Dates = Monthiversary(Monthiversary <= Current_Date);

num_past_observ_dates = length(Past_Observ_Dates);
Past_S = zeros(num_past_observ_dates, 1);

for i = 1:num_past_observ_dates
    Past_S(i) = S_Prev(S_Prev_Dates == Past_Observ_Dates(i));
end

Days_To_Observs = zeros(size(Observ_Dates));

for i = 1:length(Observ_Dates)
    Holidays_in_Between = holidays(Current_Date, Observ_Dates(1,i));
    Days_To_Observs(1, i) = wrkdydif(Current_Date, Observ_Dates(1, i), length(Holidays_in_Between));
    
end

if isempty(Days_To_Observs)
    aDiscFactor = 1;
else
    aDiscFactor = exp(-sum(Interest_rate(1:Days_To_Observs(1,end)))/252);
end

[m, n] = size(iS);
Number_of_Paths = m;

S_at_Observ_Dates = [repmat(Past_S', Number_of_Paths, 1), iS(:, Days_To_Observs)];

Capped_Return = min(((mean(S_at_Observ_Dates(:, 2:end),2) ./ S_at_Observ_Dates(:, 1)) - 1)*Participation_rate - Index_spread, Cap);

Floored_Return = max(Floor, sum(Capped_Return, 2));

Payoff = Notional * mean(Floored_Return) * aDiscFactor;

end

