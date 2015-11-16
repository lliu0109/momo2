function [] = DoMktEnvCalibration( varargin)
addpath(genpath([pwd,'\source\']));
clear all;
format long;

[dummy MktEnv_directory]=  xlsread('IA.Process.xls', 'Input', 'B11');
[dummy MktEnv_filename]=  xlsread('IA.Process.xls', 'Input', 'B12');
[dummy MktEnv_filename_BOP]=  xlsread('IA.Process.xls', 'Input', 'B13');

[iDateNum iDateStr]=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input', 'B1');
Current_Date = datenum(iDateStr);

SPX=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input', 'B3');
Dividend_initial = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input', 'B4');

%get the option chain data
[num1 text1] = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'SPX Option Data');
Strike = num1(:,1);
Call_bid = num1(:,4);
Call_mid = num1(:,5);
Call_ask = num1(:,6);
Put_bid = num1(:,7);
Put_mid = num1(:,8);
Put_ask = num1(:,9);

%calculate business days to expiry

Days_to_expiry = zeros(size(text1(2:end,4)));

for i = 1:length(text1(2:end,4))
    Holidays_in_Between = holidays(Current_Date, datenum(text1(i+1,4)));
    Days_to_expiry(i,1) = wrkdydif(Current_Date, datenum(text1(i+1,4)), length(Holidays_in_Between));
end

T_to_expiry = Days_to_expiry/252;
% choose the qualified option to be calibration set
[Indicator number_of_expiry expiry_array] = CleanOption(Call_bid, Put_bid, T_to_expiry, 1, 2);
[Indicator_1Y number_of_expiry_1Y expiry_array_1Y] = CleanOption(Call_bid, Put_bid, T_to_expiry, 1, 1);


num3 = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Interest Rates', 'A2:C8' );
Yield_t = num3(:,1);
Yield_rate = num3(:,3);
Interest_rate_at_expiry = interpolation_1d(Yield_t, Yield_rate, expiry_array(end));

Interest_rate_optimal = zeros(number_of_expiry, 1);
Dividend_optimal = zeros(number_of_expiry, 1);
MSE = zeros(number_of_expiry, 1);

C = Call_mid(T_to_expiry == expiry_array(end));
P = Put_mid(T_to_expiry == expiry_array(end));
K = Strike(T_to_expiry == expiry_array(end));
Ind = Indicator(T_to_expiry == expiry_array(end));

% start from the last maturity and going backward
% the upper bound for r is ten times of the initial libor rate
[x] = lsqnonlin(@(x) PutCallParityError(x, C, P ,K ,...
    expiry_array(end), repmat(SPX, size(C)), Ind), ...
    [Interest_rate_at_expiry, Dividend_initial], ...
    [0.0, 0.0],...
    [20*Interest_rate_at_expiry, 0.2],...
    optimset('TolX',1e-8,'TolFun',1e-6,'Display','iter'));

Interest_rate_optimal(end) = max(0,x(1));
Dividend_optimal(end) = max(0,x(2));
[num_of_options dummy] = size(Call_mid(T_to_expiry==expiry_array(end)));
MSE(end) = sqrt(sum(PutCallParityError(x, C, P, K,...
    expiry_array(end), repmat(SPX, size(C)), Ind).^2))/num_of_options;

% set constraints to make sure the interest rate curve is upward slopping

for i=1:number_of_expiry-1
    
    k = number_of_expiry - i;
    Interest_rate_at_expiry = interpolation_1d(Yield_t, Yield_rate, expiry_array(k));
    
    C = Call_mid(T_to_expiry == expiry_array(k));
    P = Put_mid(T_to_expiry == expiry_array(k));
    K = Strike(T_to_expiry == expiry_array(k));
    Ind = Indicator(T_to_expiry == expiry_array(k));
    
    [x] = lsqnonlin(@(x) PutCallParityError(x, C, P, K,...
        expiry_array(k), repmat(SPX, size(C)), Ind), ...
        [Interest_rate_at_expiry, Dividend_initial], ...
        [0.0, 0],...
        [Interest_rate_optimal(k+1), 0.2],...
        optimset('TolX',1e-8,'TolFun',1e-6, 'Display','iter'));
    
    Interest_rate_optimal(k) = max(0,x(1));
    Dividend_optimal(k) = max(0,x(2));
    [num_of_options dummy] = size(Call_mid(T_to_expiry==expiry_array(k)));
    MSE(k) = sqrt(sum(PutCallParityError(x, C, P, K,...
        expiry_array(k), repmat(SPX, size(C)), Ind).^2))/num_of_options;
    
end

% Output Results

aVolAdj = round((SPX-1200)/25)*25;

% calculate volatility surface

Strike_grid = [400 450 500 550 600 650 700 750 800 825 850 875 ...
    900 925 950 975 1000 1025 1050 1075 1100 1125 1150 1175 ...
    1200 1225 1250 1275 1300 1325 1350 1375 1400 1425 1450 1475 ...
    1500 1525 1550 1575 1600 1625 1650 1675 1700 1750 1800 1850 1900]+aVolAdj;


Vol_surface = repmat(zeros(size(Strike_grid)), number_of_expiry, 1);

for i=1:number_of_expiry
    criteria = (T_to_expiry == expiry_array(i).* (Indicator == 1) == 1);  %% choose the option qualified for calibration
    
    vol_skew_call = blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Call_mid(criteria),...
        10, Dividend_optimal(i), [], {'Call'});
    vol_skew_put =  blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Put_mid(criteria),...
        10, Dividend_optimal(i), [], {'Put'});
    
    vol_skew_per_expiry = (vol_skew_call + vol_skew_put)/2;
    
    
    Vol_surface(i, :) = interpolation_1d(Strike(criteria),vol_skew_per_expiry, Strike_grid);
    
    % handle the situation that there are #NA at the end of the surfacepoint outside of the available options. If outside the available option range use the value of the closest one.
    Indnan = isnan(Vol_surface(i,:));
    
    Ind_first = 1;
    
    for j= 1:length(Indnan)
        if Indnan(j) == 1
            Vol_surface(i, j) = Ind_first * Vol_surface(i, find(Indnan==0, 1, 'first')) + (1-Ind_first) * Vol_surface(i, find(Indnan==0, 1, 'last'));
        else
            Ind_first = 0;
        end
    end
end


%build another vol surface with bid price. will results in lower vol


Vol_surface_bid = repmat(zeros(size(Strike_grid)), number_of_expiry, 1);

for i=1:number_of_expiry
    criteria = (T_to_expiry == expiry_array(i).* (Indicator == 1) == 1);  %% choose the option qualified for calibration
    
    vol_skew_call_bid = blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Call_bid(criteria),...
        10, Dividend_optimal(i), [], {'Call'});
    vol_skew_put_bid =  blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Put_bid(criteria),...
        10, Dividend_optimal(i), [], {'Put'});
    
    vol_skew_per_expiry = (vol_skew_call_bid + vol_skew_put_bid)/2;
    
    
    Vol_surface_bid(i, :) = interpolation_1d(Strike(criteria),vol_skew_per_expiry, Strike_grid);
    
    % handle the situation that there are #NA at the end of the surfacepoint outside of the available options. If outside the available option range use the value of the closest one.
    Indnan = isnan(Vol_surface_bid(i,:));
    
    Ind_first = 1;
    
    for j= 1:length(Indnan)
        if Indnan(j) == 1
            Vol_surface_bid(i, j) = Ind_first * Vol_surface_bid(i, find(Indnan==0, 1, 'first')) + (1-Ind_first) * Vol_surface_bid(i, find(Indnan==0, 1, 'last'));
        else
            Ind_first = 0;
        end
    end
end

%build another vol surface with ask price, will results in higher vol


Vol_surface_ask = repmat(zeros(size(Strike_grid)), number_of_expiry, 1);

for i=1:number_of_expiry
    criteria = (T_to_expiry == expiry_array(i).* (Indicator == 1) == 1);  %% choose the option qualified for calibration
    
    vol_skew_call_ask = blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Call_ask(criteria),...
        10, Dividend_optimal(i), [], {'Call'});
    vol_skew_put_ask =  blsimpv(SPX, Strike(criteria),...
        Interest_rate_optimal(i), expiry_array(i),Put_ask(criteria),...
        10, Dividend_optimal(i), [], {'Put'});
    
    vol_skew_per_expiry = (vol_skew_call_ask + vol_skew_put_ask)/2;
    
    
    Vol_surface_ask(i, :) = interpolation_1d(Strike(criteria),vol_skew_per_expiry, Strike_grid);
    
    % handle the situation that there are #NA at the end of the surfacepoint outside of the available options. If outside the available option range use the value of the closest one.
    Indnan = isnan(Vol_surface(i,:));
    
    Ind_first = 1;
    
    for j= 1:length(Indnan)
        if Indnan(j) == 1
            Vol_surface_ask(i, j) = Ind_first * Vol_surface_ask(i, find(Indnan==0, 1, 'first')) + (1-Ind_first) * Vol_surface_ask(i, find(Indnan==0, 1, 'last'));
        else
            Ind_first = 0;
        end
    end
end

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], 't', 'Int&Dvd','A1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], expiry_array, 'Int&Dvd',['A2:A' num2str(number_of_expiry+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], 'r', 'Int&Dvd','B1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], Interest_rate_optimal, 'Int&Dvd',['B2:B' num2str(number_of_expiry+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], 'q', 'Int&Dvd','C1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)],Dividend_optimal, 'Int&Dvd',['C2:C' num2str(number_of_expiry+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'MSE'}, 'Int&Dvd','D1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)],MSE, 'Int&Dvd',['D2:D' num2str(number_of_expiry+1)]);

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], expiry_array, 'Vol_Surface',['A3:A' num2str(number_of_expiry+2)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], Strike_grid, 'Vol_Surface','B2:AX2');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)],Vol_surface, 'Vol_surface',['B3:AX' num2str(number_of_expiry+2)]);


xlswrite([char(MktEnv_directory) char(MktEnv_filename)], expiry_array, 'Vol_Surface',['A18:A' num2str(number_of_expiry+17)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], Strike_grid, 'Vol_Surface','B17:AX17');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)],Vol_surface_bid, 'Vol_surface',['B18:AX' num2str(number_of_expiry+17)]);
%
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], expiry_array, 'Vol_Surface',['A33:A' num2str(number_of_expiry+32)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], Strike_grid, 'Vol_Surface','B32:AX32');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)],Vol_surface_ask, 'Vol_surface',['B33:AX' num2str(number_of_expiry+32)]);


%verify the Call and Put price from the r,q and vol surface

r_array = interpolation_1d(expiry_array, Interest_rate_optimal, T_to_expiry);
q_array = interpolation_1d(expiry_array, Dividend_optimal, T_to_expiry);
vol_array = interpolation_2d(expiry_array, Strike_grid, Vol_surface, T_to_expiry, Strike);

[C_price P_price] = blsprice(SPX, Strike, r_array, T_to_expiry, vol_array, q_array);
number_of_options = length(Strike);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Ind'}, 'SPX Option Data', 'N1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], Indicator, 'SPX Option Data',['N2:N' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'model r'}, 'SPX Option Data', 'O1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], r_array, 'SPX Option Data',['O2:O' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'model q'}, 'SPX Option Data', 'P1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], q_array, 'SPX Option Data',['P2:P' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'model vol'}, 'SPX Option Data', 'Q1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], vol_array, 'SPX Option Data',['Q2:Q' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Call Price'}, 'SPX Option Data', 'R1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], C_price, 'SPX Option Data',['R2:R' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Put Price'}, 'SPX Option Data', 'S1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], P_price, 'SPX Option Data',['S2:S' num2str(number_of_options+1)]);

clear('Call_ask', 'Call_bid', 'Call_mid', 'Put_ask', 'Put_bid', 'Put_mid', 'Vol_surface', 'Vol_surface_bid', 'Vol_surface_ask');


%%

%Do CBOE one year Heston Parameter Calibration

iTgtHestonOptions = [T_to_expiry(Indicator_1Y ==1), Strike(Indicator_1Y==1),P_price(Indicator_1Y==1)];

[aInitialHestonParams]=xlsread([char(MktEnv_directory) char(MktEnv_filename_BOP)], 'Heston Params', 'D3:D7');
[aHestonLowerBnd]=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Heston Params', 'C11:C15');
[aHestonUpperBnd]=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Heston Params', 'D11:D15');


[aObtainedHestonParams resnorm residual] = lsqnonlin(@(x) mHestonFit(x, SPX, Interest_rate_optimal, Dividend_optimal, expiry_array, iTgtHestonOptions),...
    aInitialHestonParams, ...
    aHestonLowerBnd, ...
    aHestonUpperBnd, ...
    optimset('TolX',1e-4,'TolFun',1e-4, 'Display','iter'));  %optimset('TolX',1e-8,'TolFun',1e-6, 'Display','iter'));

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aInitialHestonParams, 'Heston Params', 'C3:C7');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aObtainedHestonParams, 'Heston Params', 'D3:D7');

aOptions = [T_to_expiry, Strike, P_price];
[oHestonDiff, aHestonModelOptionPrices] = mHestonFit(aObtainedHestonParams, SPX, Interest_rate_optimal, Dividend_optimal, expiry_array, aOptions);

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Heston Put Price'}, 'SPX Option Data', 'U1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aHestonModelOptionPrices, 'SPX Option Data',['U2:U' num2str(number_of_options+1)]);
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Heston Price Diff'}, 'SPX Option Data', 'V1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], oHestonDiff, 'SPX Option Data',['V2:V' num2str(number_of_options+1)]);

clear('iTgtHestonOptions', 'aOptions');
% 


% Do CBOE one year Bates Parameter Calibration

iNumofTrials = xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Input', 'B14');

iTgtBatesOptions = [Days_to_expiry(Indicator_1Y ==1), Strike(Indicator_1Y==1),P_price(Indicator_1Y==1)];

[aInitialBatesParams]=xlsread([char(MktEnv_directory) char(MktEnv_filename_BOP)], 'Bates Params', 'D3:D10');
[aBatesLowerBnd]=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Bates Params', 'C14:C21');
[aBatesUpperBnd]=xlsread([char(MktEnv_directory) char(MktEnv_filename)], 'Bates Params', 'D14:D21');

iRandSetting = 2012;

[aObtainedBatesParams resnorm residual] = lsqnonlin(@(x) mBatesFit(x, SPX, Interest_rate_optimal, Dividend_optimal, expiry_array, iTgtBatesOptions, iRandSetting, iNumofTrials),...
    aInitialBatesParams, ...
    aBatesLowerBnd, ...
    aBatesUpperBnd, ...
    optimset('TolX',1e-4,'TolFun',1e-4, 'Display','iter'));  %optimset('TolX',1e-8,'TolFun',1e-6, 'Display','iter'));

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aInitialBatesParams, 'Bates Params', 'C3:C10');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aObtainedBatesParams, 'Bates Params', 'D3:D10');

[aBatesDiff, aBatesModelOptionPrices] = mBatesFit(aObtainedBatesParams, SPX, Interest_rate_optimal, Dividend_optimal, expiry_array, iTgtBatesOptions, iRandSetting, iNumofTrials);

aBatesPutPrices = zeros(size(P_price));
aBatesPutDiff = zeros(size(P_price));

aBatesPutPrices(Indicator_1Y == 1) = aBatesModelOptionPrices;
aBatesPutDiff(Indicator_1Y == 1) = aBatesDiff;

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Bates Put Price'}, 'SPX Option Data', 'X1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aBatesPutPrices, 'SPX Option Data',['X2:X' num2str(number_of_options+1)]);

xlswrite([char(MktEnv_directory) char(MktEnv_filename)], {'Bates Price Diff'}, 'SPX Option Data', 'Y1');
xlswrite([char(MktEnv_directory) char(MktEnv_filename)], aBatesPutDiff, 'SPX Option Data',['Y2:Y' num2str(number_of_options+1)]);

clear('iTgtBatesOptions', 'aBatesDiff', 'aBatesModelOptionPrices', 'aBatesPutPrices', 'aBatesPutDiff');

end
%%


function f = PutCallParityError( x, C, P, K, T, S, Indicator)

r = x(1);
q = x(2);


f = (C + K * exp( -r * T ) - P - S * exp( -q * T)).* Indicator;
end
%%

function [Indicator num_of_expiry_chosen expiry_array_chosen] = CleanOption(C_Bid, P_Bid, T_to_expiry, limit, term)
%first only pick the options that has two consecutive bids that is greater
%than or equal to limit. Assuming C and P are sorted by accending strike

Indicator = zeros(size(C_Bid));

expiry_array = sort(unique(T_to_expiry));
%num_expiry = length(expiry_array);

%find the first expiry date that is long than one year
expiry_end = find(expiry_array>term, 1, 'first') ;
% If term is bigger than or equal to the max of expiry_array the previous
% expression with return empty.  In that case we that the last index to
% be the end of the expiry_array.
if ~length(expiry_end)
    expiry_end = length(expiry_array);
    
end
expiry_begin = find(expiry_array>1/12, 1, 'first');
num_of_expiry_chosen = expiry_end - expiry_begin + 1 ;
expiry_array_chosen = expiry_array(expiry_begin:expiry_end);

for j = expiry_begin:expiry_end
    
    offset = find(T_to_expiry == expiry_array(j), 1, 'first')-1;
    
    num_of_options = length(C_Bid(T_to_expiry == expiry_array(j)));
    
    i = 1;
    while i<num_of_options
        if P_Bid(offset+i)<limit || P_Bid(offset+i+1)<limit
            i = i+1;
        else
            break,
        end
    end
    k_start = offset+i;
    
    i = num_of_options;
    
    while i>1
        if C_Bid(offset+i)<limit || C_Bid(offset+i-1)<limit
            i = i-1;
        else
            break,
        end
    end
    
    k_end = offset+i;
    
    Indicator(k_start:k_end) = 1;
    
end


end

function [oDiff, aModelOptionPrices] = mHestonFit(iModelParameters, s0, Interest_rate, Dividend, t, iTgtOptions)
%iModelParameters include all heston parameters.

aTgtOptionMaturities = iTgtOptions(:, 1);
aTgtOptionStrikes = iTgtOptions(:, 2);
aTgtOptionPrices = iTgtOptions(:, 3);

iRate = interpolation_1d(t, Interest_rate, aTgtOptionMaturities);
iDividend = interpolation_1d(t, Dividend, aTgtOptionMaturities);

aVolParam.VarZero = iModelParameters(5);
aVolParam.Theta = iModelParameters(2);
aVolParam.MeanReversionRate = iModelParameters(1);
aVolParam.Sigma = iModelParameters(3);
aVolParam.Correlation = iModelParameters(4); 

%using the heston close form solution, the yield and boundary indicator are
%not applicable. 
aEquityModel = heston(0, aVolParam, 0);

aModelOptionPrices = EuropeanOptionPricer(aEquityModel, iRate, iDividend, ...
    s0, aTgtOptionStrikes, aTgtOptionMaturities, 0, 10);

oDiff = (aModelOptionPrices - aTgtOptionPrices);

end

%%

function [oDiff, aModelOptionPrices] = mBatesFit(iModelParameters, s0, Interest_rate, Dividend,t, iTgtOptions, iRandSetting, iNumPaths)
%iModelParameters include all heston parameters.

aTgtOptionDaystoMaturities = iTgtOptions(:, 1);
aTgtOptionStrikes = iTgtOptions(:, 2);
aTgtOptionPrices = iTgtOptions(:, 3);

aProjectionYears = 2;

aTimeAxis =  (1/252 : 1/252 : aProjectionYears);
aTimeDiff = mean(diff(aTimeAxis));

% Volatility Parameters
aVolParam.VarZero = iModelParameters(5);
aVolParam.Theta = [repmat(iModelParameters(2), 1, aProjectionYears/aTimeDiff)];
aVolParam.MeanReversionRate = [repmat(iModelParameters(1),1,aProjectionYears/aTimeDiff)];
aVolParam.Sigma = [repmat(iModelParameters(3),1,aProjectionYears/aTimeDiff)];
aVolParam.Correlation = iModelParameters(4);
aVolParam.JumpIntensity = iModelParameters(6);
aVolParam.JumpSizeMean = iModelParameters(7);
aVolParam.JumpSizeStd = iModelParameters(8);

aCorrMatrix = [1,aVolParam.Correlation; aVolParam.Correlation, 1];
aCholMatrixHeston = chol(aCorrMatrix);


rng(iRandSetting);
aRandn = randn(length(aTimeAxis), 2, iNumPaths);

for aIdx = 1:iNumPaths
    aRandn(1:aProjectionYears/aTimeDiff, :, aIdx) = aRandn(1:aProjectionYears/aTimeDiff, :, aIdx) * aCholMatrixHeston;
end

aEquityRandn = reshape(aRandn(:,1, :), [length(aTimeAxis), iNumPaths]);
aVarianceRandn = reshape(aRandn(:, 2, :), [length(aTimeAxis), iNumPaths]);

aJumpSizeRandn = randn(length(aTimeAxis), iNumPaths);
aJumpRandUnif = random('unif', 0, 1, [length(aTimeAxis), iNumPaths]);
aJumpRandp = poissinv(aJumpRandUnif,aVolParam.JumpIntensity * aTimeDiff);

Interest_rate_forward = zeros(size(Interest_rate));
Dividend_forward = zeros(size(Dividend));

Interest_rate_forward(1,1) = Interest_rate(1,1);
Dividend_forward(1,1) = Dividend(1,1);

Interest_rate_forward(2:end, 1) = (Interest_rate(2:end, 1) .* t(2:end, 1) - Interest_rate(1:end-1, 1) .* t(1:end-1, 1)) ./ (t(2:end, 1) - t(1:end-1,1));
Dividend_forward(2:end, 1) = (Dividend(2:end, 1) .* t(2:end, 1) - Dividend(1:end-1, 1) .* t(1:end-1, 1)) ./ (t(2:end, 1) - t(1:end-1,1));

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


%using reflection boundary
aEquityModel = Bates(aYieldBase, aVolParam, 1);

S = zeros(iNumPaths, aProjectionYears/aTimeDiff+1);

aSetLength = 10000;

aScenSets = ceil(iNumPaths/aSetLength);

% do the calculation in 10000 sets to save memory usage.
for i = 1:aScenSets
    [aEquityPaths, aVolPaths] = MakePaths(aEquityModel, aTimeAxis, min(aSetLength, iNumPaths - (i-1)*aSetLength), aEquityRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, iNumPaths))', aVarianceRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, iNumPaths))'...
        , aJumpRandp(:, ((i-1)*aSetLength+1):min(i*aSetLength, iNumPaths))', aJumpSizeRandn(:, ((i-1)*aSetLength+1):min(i*aSetLength, iNumPaths))');
    
    S_temp = s0 * exp(cumsum(aEquityPaths, 2));
    
    S_temp = [repmat(s0, min(aSetLength, iNumPaths - (i-1)*aSetLength),1) S_temp];
    
    S((i-1)*aSetLength+1:min(i*aSetLength, iNumPaths), :) = S_temp;
    
    clear('aEquityPaths', 'aVolPaths', 'S_temp');
    
end


aDiscountFactor = exp(-cumsum(aRate)* aTimeDiff);

aModelOptionPrices = EuropeanOptionPricer(aEquityModel, S, aTgtOptionStrikes, aTgtOptionDaystoMaturities, zeros(size(aTgtOptionStrikes)), aDiscountFactor);

oDiff = (aModelOptionPrices - aTgtOptionPrices);

end

