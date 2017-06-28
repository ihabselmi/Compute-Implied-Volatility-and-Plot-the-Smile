clear all
clc

% Input : Strike Price, Observed Price, Implied Volatility

optionprices = xlsread('OptionPrices.xlsx')

% Set the parameters
r =.01;
S0 = 204.24;
T = 30/252;
sigma = .1220;

for i = 1:size(optionprices,1)
    K = optionprices(i,1);
    price = optionprices(i,2);
    myfunc = @(vol) abs(blsprice(S0, K, r, T, vol)-price)
    implied_vol(i,:) = [fsolve(myfunc, 0.5) blsimpv(S0, K, r, T, price)];
end

out = [optionprices implied_vol]

% write the data 
xlswrite('optionprices_out.xlsx',out)

% plot the curves on separate graphs
figure
subplot(2,1,1);
plot(out(:,1),out(:,3))
title('Yahoo Finance Implied Vol')
xlabel('Strike')
ylabel('Implied Volatility')

subplot(2,1,2);
plot(out(:,1),out(:,4))
title('Calculated Implied Vol')
xlabel('Strike')
ylabel('Implied Volatility')

% or plot them together
figure
plot(out(:,1),out(:,3),out(:,1),out(:,4))

title('Implied Volatilities')
legend('Yahoo Finance','Calculated')


