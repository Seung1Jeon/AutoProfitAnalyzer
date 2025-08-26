import pandas as pd
import numpy as np

def _calculate_and_add_ratio(df, new_metric_name, numerator_name, denominator_name):
    """Helper function to calculate and add a new ratio row."""
    first_col_name = df.columns[0]
    df_indexed = df.set_index(first_col_name)

    if numerator_name in df_indexed.index and denominator_name in df_indexed.index:
        numerator = df_indexed.loc[numerator_name]
        denominator = df_indexed.loc[denominator_name]

        with np.errstate(divide='ignore', invalid='ignore'):
            ratio = (numerator / denominator) * 100
        
        ratio.replace([np.inf, -np.inf], np.nan, inplace=True)
        ratio.fillna(0, inplace=True)

        new_row = pd.DataFrame([{first_col_name: new_metric_name, **ratio.to_dict()}])
        df = pd.concat([df, new_row], ignore_index=True)
    return df

def add_financial_ratios(df_lg):
    """Calculates and adds key financial ratios for the LG dataframe."""
    # Add Net Profit Margin
    df_lg = _calculate_and_add_ratio(df_lg, '매출순수익률', '당기순이익', '매출액')
    
    # Add Gross Profit Margin
    df_lg = _calculate_and_add_ratio(df_lg, '매출총이익률', '매출총이익', '매출액')

    # Add Break-Even Point
    first_col_name = df_lg.columns[0]
    df_lg_indexed = df_lg.set_index(first_col_name)
    required_items = ['판매비와관리비', '매출원가', '매출액']
    if all(item in df_lg_indexed.index for item in required_items):
        sga_expenses = df_lg_indexed.loc['판매비와관리비']
        cost_of_goods = df_lg_indexed.loc['매출원가']
        revenue = df_lg_indexed.loc['매출액']

        with np.errstate(divide='ignore', invalid='ignore'):
            contribution_margin_ratio = 1 - (cost_of_goods / revenue)
            break_even_point = sga_expenses / contribution_margin_ratio
        
        break_even_point.replace([np.inf, -np.inf], np.nan, inplace=True)
        break_even_point.fillna(0, inplace=True)

        new_row = pd.DataFrame([{first_col_name: '손익분기점추정', **break_even_point.to_dict()}])
        df_lg = pd.concat([df_lg, new_row], ignore_index=True)

    return df_lg

def calculate_company_average(dataframes):
    """Calculates the average revenue and operating profit for a list of companies."""
    revenue_data = []
    op_profit_data = []

    for df in dataframes:
        df_indexed = df.set_index(df.columns[0])
        
        if '매출액' in df_indexed.index:
            revenue_series = pd.to_numeric(df_indexed.loc['매출액'], errors='coerce')
            revenue_data.append(revenue_series)
        if '영업이익' in df_indexed.index:
            op_profit_series = pd.to_numeric(df_indexed.loc['영업이익'], errors='coerce')
            op_profit_data.append(op_profit_series)

    if not revenue_data or not op_profit_data:
        print("경고: 3사 평균을 계산하기에 충분한 '매출액' 또는 '영업이익' 데이터를 찾지 못했습니다.")
        return None

    revenue_df = pd.DataFrame(revenue_data)
    op_profit_df = pd.DataFrame(op_profit_data)
    
    avg_revenue_by_year = revenue_df.mean(axis=0)
    avg_op_profit_by_year = op_profit_df.mean(axis=0)

    avg_data = {'구분': ['3사 평균 매출액', '3사 평균 영업이익']}
    all_years = set(revenue_df.columns) | set(op_profit_df.columns)
    for year in sorted(list(all_years)):
        avg_data[str(year)] = [avg_revenue_by_year.get(year), avg_op_profit_by_year.get(year)]
    
    return pd.DataFrame(avg_data)