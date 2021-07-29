# -*- coding: utf-8 -*-
#
# 计算池内产品相关指标
#

import numpy as np
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import warnings
warnings.filterwarnings("ignore")

def ret(asset, start, end):
    '''
    计算各期收益率
    '''
    rets = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', parse_dates=['Date'], index_col='Date', sheet_name=i)
        df = df[(df.index >= start) & (df.index <= end)]
        if len(df) <= 1:
            rets.append('-')
        else:
            returns = df.iloc[-1,1] / df.iloc[0,1] - 1
            rets.append(returns)
    return rets
            
def drawdown(asset, start, end):
    '''
    计算近一年最大回撤
    '''
    drawdown_list = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df = df.sort_index()
        df = df[(df.index >= start) & (df.index <= end)]
        previous_max = df['AccNAV'].cummax()
        drawdowns = (df['AccNAV'] - previous_max) / previous_max
        drawdown_list.append(drawdowns.min())
    return drawdown_list

def annualized_std(asset):
    '''
    计算年化波动率
    '''
    annualized_std_list = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df.index = pd.to_datetime(df.index)
        std = df.iloc[:,0].std()
        date_diff = df.index[-1] - df.index[0]
        t = np.timedelta64(date_diff, 'D').astype(int)
        annualized_std = std * (365 / t) ** 0.5
        annualized_std_list.append(annualized_std)
    return annualized_std_list

def sharpe_ratio(asset, start, end):
    '''
    计算近一年夏普比率
    '''
    sharpe_ratio_list = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df = df[(df.index >= start) & (df.index <= end)]
        ret = df.iloc[-1,1] / df.iloc[0,1] - 1
        std = df['AccNAV'].std()
        sharpe_ratio = (ret - 0.03) / std
        sharpe_ratio_list.append(sharpe_ratio)
    return sharpe_ratio_list

def win_rate(asset):
    '''
    计算周胜率
    '''
    win_rate_list = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df = df.resample('1w', label='left').last()
        if len(df) <= 1:
            win_rate_list.append('-')
        else:
            df['win'] = np.where(df['NAV'] > 0.9995 * df['NAV'].shift(1),1,0)
            win_rate = df['win'].mean()
            win_rate_list.append(win_rate)
    return win_rate_list

def calmar(asset):
    '''
    计算calmar
    '''
    calmar_list = []
    for i in asset:
        df = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        ret = df.iloc[-1,1] / df.iloc[0,1] - 1
        date_diff = df.index[-1] - df.index[0]
        t = np.timedelta64(date_diff, 'D').astype(int)
        annualized_ret = ret * 365 / t
        previous_max = df['AccNAV'].cummax()
        drawdowns = (df['AccNAV'] - previous_max) / previous_max
        calmar = -annualized_ret / drawdowns.min()
        calmar_list.append(calmar)
    return calmar_list

def alpha_ret(asset, path, start, end):
    '''
    计算超额收益
    '''
    alpha_ret_list = []
    for i in asset:
        df1 = pd.read_excel(path, index_col='Date', parse_dates=True)
        df2 = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df = pd.merge(df1, df2, on=['Date'])
        df.drop(['Unnamed: 0'], axis=1)
        df.index = pd.to_datetime(df.index)
        df = df[(df.index >= start) & (df.index <= end)]
        if len(df) <= 1:
            alpha_ret_list.append('-')
        else:
            Index_ret = df['close'].iloc[0] / df['close'].iloc[-1] - 1
            ret = df['AccNAV'].iloc[0] / df['AccNAV'].iloc[-1] - 1
            alpha = ret - Index_ret
            alpha_ret_list.append(alpha)
    return alpha_ret_list

def alpha_win_rate(asset, path):
    '''
    计算超额周胜率
    '''
    alpha_win_rate = []
    for i in asset:
        df1 = pd.read_excel(path, index_col='Date', parse_dates=True)
        df2 = pd.read_excel(r'C:\Users\Steve Liu\Documents\工作\Data\50底层详细历史净值.xlsx', index_col='Date', sheet_name=i, parse_dates=True)
        df = pd.merge(df1, df2, on=['Date'])
        df.drop(['Unnamed: 0'], axis=1)
        df = df.resample('1w', label='right').last()
        if len(df) <= 1:
            alpha_win_rate.append('-')
        else:
            df['index_ret'] = df['close'].pct_change()
            df['ret'] = df['NAV'].pct_change()
            df.dropna(axis=0)
            df['alpha'] = df['ret'] - df['index_ret']
            df['win'] = np.where(df['alpha'] > 0, 1, 0)
            win_rate = df['win'].mean()
            alpha_win_rate.append(win_rate)
    return alpha_win_rate

def mean_median(path, sheet_name, period):
    '''
    计算市场同类产品的同期均值和中位数
    '''
    df = pd.read_excel('./'+path+'.xls', sheet_name=sheet_name)
    df = df[period]
    mu, std = df.mean(), df.std()
    max, min = mu + 3 * std, mu - 3 * std
    mean = df[(df >= min) & (df <= max)].mean()
    median = df[(df >= min) & (df <= max)].median()
    return mean, median


