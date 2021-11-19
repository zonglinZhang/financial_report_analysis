import math

from jqdatasdk import *
from datetime import datetime
import xlwt


def get_style(col, data=0):
    # 危险-red
    red_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    pattern = xlwt.Pattern()  # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = 2  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue,
    # 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow ,
    # almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    red_style.pattern = pattern
    # warn-yellow
    yellow_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    yellow_pattern = xlwt.Pattern()  # Create the Pattern
    yellow_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    yellow_pattern.pattern_fore_colour = 5
    yellow_style.pattern = yellow_pattern
    # common-white
    common_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    common_pattern = xlwt.Pattern()  # Create the Pattern
    common_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    common_pattern.pattern_fore_colour = 1
    common_style.pattern = common_pattern

    # 复合指标
    composite_index_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    composite_index_pattern = xlwt.Pattern()  # Create the Pattern
    composite_index_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    composite_index_pattern.pattern_fore_colour = 17
    composite_index_style.pattern = composite_index_pattern

    # 普通指标
    common_index_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    common_index_pattern = xlwt.Pattern()  # Create the Pattern
    common_index_pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    common_index_pattern.pattern_fore_colour = 5
    common_index_style.pattern = common_index_pattern
    if col == "composite_index":
        return composite_index_style
    if col == "common_index":
        return common_index_style
    elif col == "roe":
        if data < 15:
            return red_style
        elif data < 25:
            return yellow_style
        else:
            return common_style
    elif col == "liability_rate":
        if data > 60:
            return red_style
        else:
            return common_style
    elif col == "assets_grown_rate":
        if data < 0:
            return red_style
        else:
            return common_style
    elif col == "gross_profit_margin":
        if data < 40:
            return red_style
        else:
            return common_style
    elif col == "gross_profit_margin_grown_rate":
        if abs(data) > 20:
            return red_style
        else:
            return common_style
    elif col == "inc_total_revenue_year_on_year":
        if data < 0:
            return red_style
        elif data < 10:
            return yellow_style
        else:
            return common_style
    elif col == "longterm_account_payable":
        if data < 0:
            return red_style
        else:
            return common_style
    elif col == "prepaid_receivables":
        if data < 0:
            return red_style
        else:
            return common_style
    elif col == "account_receivable_contract_assets":
        if data > 15:
            return red_style
        else:
            return common_style
    elif col == "fixed_assets_constru_in_process":
        if data > 40:
            return red_style
        else:
            return common_style
    elif col == "last_fixed_assets_constru_in_process":
        if data > 20:
            return red_style
        else:
            return common_style
    elif col == "good_will":
        if data > 10:
            return red_style
        else:
            return common_style
    elif col == "sale_expense_rate":
        if data > 30:
            return red_style
        elif data > 15:
            return yellow_style
        else:
            return common_style
    elif col == "gross_profit_margin_period_expense":
        if data > 60:
            return red_style
        elif data > 40:
            return yellow_style
        else:
            return common_style
    elif col == "main_profit_rate":
        if data < 15:
            return red_style
        else:
            return common_style
    elif col == "income_main_profit_rate":
        if data < 80:
            return red_style
        else:
            return common_style
    elif col == "net_operate_cash_flow":
        if data < 100:
            return red_style
        else:
            return common_style
    elif col == "fix_intan_other_asset_acqui_cash":
        if data > 100:
            return red_style
        else:
            return common_style
    elif col == "dividend_interest_payment":
        if data < 20 or data > 70:
            return red_style
        else:
            return common_style
    elif col == "investment_assets_rate":
        if data > 10:
            return red_style
        else:
            return common_style
    elif col == "account_receivable_rate":
        if data > 5:
            return red_style
        else:
            return common_style


def init_jqdata(code):
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet("indicator")
    sheet.write(0, 0, "年份")
    for i in range(datetime.now().year - 5, datetime.now().year):
        sheet.write(0, i + 1 - datetime.now().year + 5, i)
    auth('18518904846', 'qianLIN1213')  # ID是申请时所填写的手机号；Password为聚宽官网登录密码
    # 财务指标
    indicator_q = query(
        indicator
    ).filter(
        income.code == code,
    )

    indicator_summary_list = [get_fundamentals(indicator_q, statDate=year) for year in
                              range(datetime.now().year - 5, datetime.now().year)]

    # 资产负债数据汇总
    balance_q = query(
        balance
    ).filter(
        balance.code == code
    )

    balance_summary_list = [get_fundamentals(balance_q, statDate=year) for year in
                            range(datetime.now().year - 5, datetime.now().year)]

    # 合并资产负债表明细
    stk_balance_list = [
        finance.run_query(query(finance.STK_BALANCE_SHEET).filter(finance.STK_BALANCE_SHEET.code == code,
                                                                  finance.STK_BALANCE_SHEET.end_date == str(
                                                                      year) + "-12-31",
                                                                  finance.STK_BALANCE_SHEET.source_id == 321003,
                                                                  finance.STK_BALANCE_SHEET.report_type == 1)) for
        year in
        range(datetime.now().year - 5, datetime.now().year)]
    # 合并利润表明细stk_balance_list
    stk_income_list = [
        finance.run_query(query(finance.STK_INCOME_STATEMENT).filter(finance.STK_INCOME_STATEMENT.code == code,
                                                                     finance.STK_INCOME_STATEMENT.end_date == str(
                                                                         year) + "-12-31",
                                                                     finance.STK_INCOME_STATEMENT.source_id == 321003,
                                                                     finance.STK_INCOME_STATEMENT.report_type == 1)) for
        year in
        range(datetime.now().year - 5, datetime.now().year)]
    # 金融类合并资产负债表
    # finance_balance_list = [
    #     finance.run_query(query(finance.FINANCE_BALANCE_SHEET).filter(finance.FINANCE_BALANCE_SHEET.code == code,
    #                                                                   finance.FINANCE_BALANCE_SHEET.end_date == str(
    #                                                                       year) + "-12-31",
    #                                                                   finance.FINANCE_BALANCE_SHEET.source_id == 321003,
    #                                                                   finance.FINANCE_BALANCE_SHEET.report_type == 1))
    #     for
    #     year in
    #     range(datetime.now().year - 5, datetime.now().year)]
    # for index, data in enumerate(stk_income_list):
    #     print(data["rd_expenses"][0] if len(data["rd_expenses"]) > 0 and data["rd_expenses"][0] is not None else 0)
    # 利润表
    income_q = query(
        income
    ).filter(
        balance.code == code
    )

    income_summary_list = [get_fundamentals(income_q, statDate=year) for year in
                           range(datetime.now().year - 5, datetime.now().year)]

    # 现金流量数据
    cash_q = query(cash_flow).filter(cash_flow.code == code)
    cash_summary_list = [get_fundamentals(cash_q, statDate=year) for year in
                         range(datetime.now().year - 5, datetime.now().year)]
    # 总资产
    sheet.write(1, 0, "总资产(亿)", get_style("common_index"))
    sheet.write(2, 0, "资产增长率", get_style("common_index"))
    # 总负债
    sheet.write(3, 0, "总负债(亿)", get_style("common_index"))
    # 资产负债率
    sheet.write(4, 0, "资产负债率(%) < 60%", get_style("common_index"))
    # 准货币资金 - 有息负债（亿, > 0）
    sheet.write(5, 0, "准货币资金", get_style("common_index"))
    sheet.write(6, 0, "有息负债", get_style("common_index"))
    sheet.write(7, 0, "准货币资金 - 有息负债(亿) > 0", get_style("composite_index"))
    sheet.write(8, 0, "应收预付-应付预收(亿) > 0", get_style("composite_index"))
    sheet.write(9, 0, "(应收账款+合同资产）/总资产,最优秀公司<1%优秀公司<3%大于15%的淘汰掉", get_style("composite_index"))
    sheet.write(10, 0, "固定资产(亿)", get_style("common_index"))
    sheet.write(11, 0, "在建工程(亿)", get_style("common_index"))
    sheet.write(12, 0, "(固定资产+在建工程)/总资产(%)，<40%且稳定或者增幅较小的风险小，短期内增幅较大财务造假可能性大", get_style("composite_index"))
    sheet.write(13, 0, "(固定资产+在建工程)增长幅度(%)", get_style("composite_index"))
    sheet.write(14, 0, "存货占总资产的比例(%)", get_style("composite_index"))
    sheet.write(15, 0, "应收账款占总资产的比例(%) 大于5%并且存货占总资产比例大于15%的公司淘汰", get_style("composite_index"))
    sheet.write(16, 0, "商誉占总资产比例(%) <10", get_style("composite_index"))
    sheet.write(17, 0, "营业总收入(亿)", get_style("common_index"))
    sheet.write(18, 0, "营业收入增长率(%) > 10%", get_style("common_index"))
    sheet.write(19, 0, "毛利率(%) > 40%", get_style("common_index"))
    sheet.write(20, 0, "毛利率波动幅度(%) < 20%", get_style("composite_index"))
    sheet.write(21, 0, "销售费用率(%)，<15%,大于30淘汰", get_style("common_index"))
    sheet.write(22, 0, "期间费用率(%)", get_style("common_index"))
    sheet.write(23, 0, "期间费用率/毛利率 ，<40%,大于60淘汰", get_style("composite_index"))
    sheet.write(24, 0, "营业利润(亿)", get_style("common_index"))
    sheet.write(25, 0, "净利润(亿)", get_style("common_index"))
    sheet.write(26, 0, "主营利润率(%), >15%", get_style("composite_index"))
    sheet.write(27, 0, "主营利润 / 营业利润, > 80 %", get_style("composite_index"))
    sheet.write(28, 0, "平均净利润现金比例，经营活动产生的现金流量净额/净利润 >100%", get_style("composite_index"))
    sheet.write(29, 0, "roe >20%,或者>15%", get_style("common_index"))
    sheet.write(30, 0, "购建固定资产，无形资产或者其他长期资产支付的现金与经营活动产生的现金流量净额比例大于100%或者持续小于3%的淘汰", get_style("composite_index"))
    sheet.write(31, 0, "分配股利，利润或者偿付利息支付的现金与经营活动产生的现金流量净额的比率应该在20%-70%之间", get_style("composite_index"))
    # sheet.write(31, 0,
    #             "投资类资产占总资产比例，<10%投资类资产=（以公允价值计量且其变动计入当期损益的金融资产 +债权投资+可供出售金融资产+其他权益工具投资+其他债权投资+持有至到期投资+其他非流动金融资产+长期股权投资+投资性房地产)",
    #             get_style("composite_index"))
    for index, balance_data in enumerate(balance_summary_list):
        sheet.write(1, index + 1, round(balance_data['total_assets'][0] / 100000000, 2))
        # 资产增长率
        if index == 0:
            sheet.write(2, index + 1, "-")
        else:
            assets_grown_rate = round(
                (balance_data['total_assets'][0] -
                 balance_summary_list[index - 1]['total_assets'][
                     0]) /
                balance_summary_list[index - 1]['total_assets'][
                    0] * 100, 2)
            sheet.write(2, index + 1, assets_grown_rate, get_style("assets_grown_rate", assets_grown_rate))
        sheet.write(3, index + 1, round(balance_data['total_liability'][0] / 100000000, 2))
        liability_rate = round(balance_data['total_liability'][0] /
                               balance_data['total_assets'][0] * 100, 2)
        sheet.write(4, index + 1, liability_rate, get_style("liability_rate", liability_rate))
        # 准货币资金 = 交易性金融资产 + 货币资金
        quasi_monetary_funds = (balance_data['cash_equivalents'][0] if not math.isnan(
            balance_data['cash_equivalents']) else 0) + (
                                   balance_data['trading_assets'][0] if not math.isnan(
                                       balance_data['trading_assets'][0]) else 0)
        # 有息负债=短期借款+一年内到期的非流动负债+长期借款+应付债券+长期应付款
        interest_bearing_liabilities = (balance_data['shortterm_loan'][0] if not math.isnan(
            balance_data['shortterm_loan'][0]) else 0) + \
                                       (balance_data['non_current_liability_in_one_year'][0] if not math.isnan(
                                           balance_data[
                                               'non_current_liability_in_one_year'][0]) else 0) + \
                                       (balance_data['longterm_loan'][0] if not math.isnan(
                                           balance_data['longterm_loan'][0]) else 0) + (
                                           balance_data['bonds_payable'][0] if not math.isnan(
                                               balance_data['bonds_payable'][
                                                   0]) else 0) + \
                                       (balance_data['longterm_account_payable'][0] if not math.isnan(balance_data[
                                                                                                          'longterm_account_payable'][
                                                                                                          0]) else 0)
        sheet.write(5, index + 1, round(quasi_monetary_funds / 100000000, 2))
        sheet.write(6, index + 1, round(interest_bearing_liabilities / 100000000, 2))
        # 准货币资金-有息负债
        sheet.write(7, index + 1, round((quasi_monetary_funds - interest_bearing_liabilities) / 100000000, 2),
                    get_style("longterm_account_payable",
                              round((quasi_monetary_funds - interest_bearing_liabilities) / 100000000, 2)))

        # print(stk_balance_list[index]["contract_liability"])
        # 应付预收=应付票据+应付账款+预收款项+合同负债
        payable_in_advance = (balance_data['notes_payable'][0] if not math.isnan(
            balance_data['notes_payable'][0]) else 0) + (
                                 balance_data['accounts_payable'][0] if not math.isnan(
                                     balance_data['accounts_payable'][0]) else 0) + (
                                 balance_data['advance_peceipts'][0] if not math.isnan(
                                     balance_data['advance_peceipts'][0]) else 0) + (
                                 stk_balance_list[index]["contract_liability"][0] if
                                 stk_balance_list[index]["contract_liability"][0] else 0)
        # 应收预付 = 应收票据+应收账款+应收款项融资+预付款项+合同资产
        prepaid_receivables = (balance_data['bill_receivable'][0] if not math.isnan(
            balance_data['bill_receivable']) else 0) + (
                                  balance_data['account_receivable'][0] if not math.isnan(
                                      balance_data['account_receivable'][0]) else 0) + (
                                  stk_balance_list[index]["receivable_fin"][0] if
                                  stk_balance_list[index]["receivable_fin"][0] else 0) + (
                                  balance_data['advance_payment'][0] if not math.isnan(
                                      balance_data['advance_payment'][0]) else 0) + (
                                  stk_balance_list[index]["contract_assets"][0] if
                                  stk_balance_list[index]["contract_assets"][0] else 0)
        sheet.write(8, index + 1, round((payable_in_advance - prepaid_receivables) / 100000000, 2),
                    get_style("prepaid_receivables", round((payable_in_advance - prepaid_receivables) / 100000000, 2)))
        # (应收账款+合同资产）/总资产
        account_receivable_contract_assets = (balance_data['account_receivable'][0] if not math.isnan(
            balance_data['account_receivable'][0]) else 0) + (
                                                 stk_balance_list[index]["contract_assets"][0] if
                                                 stk_balance_list[index]["contract_assets"][0] else 0)
        sheet.write(9, index + 1, round(account_receivable_contract_assets / balance_data['total_assets'][0] * 100, 2),
                    get_style(data=round(account_receivable_contract_assets / balance_data['total_assets'][0] * 100, 2),
                              col="account_receivable_contract_assets"))
        sheet.write(10, index + 1,
                    round(balance_data["fixed_assets"][0] / 100000000, 2) if not math.isnan(
                        balance_data["fixed_assets"][0]) else 0)
        sheet.write(11, index + 1,
                    round(balance_data["constru_in_process"][0] / 100000000, 2) if not math.isnan(
                        balance_data["constru_in_process"][0]) else 0)
        fixed_assets_constru_in_process = (balance_data["constru_in_process"][0] if not math.isnan(
            balance_data["constru_in_process"][0]) else 0) + (balance_data["fixed_assets"][0] if not math.isnan(
            balance_data["fixed_assets"][0]) else 0)

        sheet.write(12, index + 1,
                    round(fixed_assets_constru_in_process / balance_data['total_assets'][0] * 100, 2),
                    get_style("fixed_assets_constru_in_process",
                              round(fixed_assets_constru_in_process / balance_data['total_assets'][0] * 100, 2)))
        if index == 0:
            sheet.write(13, 1, "-")
        else:
            last_fixed_assets_constru_in_process = (balance_summary_list[index - 1]["constru_in_process"][
                                                        0] if not math.isnan(
                balance_summary_list[index - 1]["constru_in_process"][0]) else 0) + (
                                                       balance_summary_list[index - 1]["fixed_assets"][
                                                           0] if not math.isnan(
                                                           balance_summary_list[index - 1]["fixed_assets"][0]) else 0)

            sheet.write(13, index + 1,
                        round((
                                      fixed_assets_constru_in_process - last_fixed_assets_constru_in_process) / last_fixed_assets_constru_in_process * 100,
                              2) if last_fixed_assets_constru_in_process > 0 else 0,
                        get_style("last_fixed_assets_constru_in_process", round((
                                                                                        fixed_assets_constru_in_process - last_fixed_assets_constru_in_process) / last_fixed_assets_constru_in_process * 100,
                                                                                2) if last_fixed_assets_constru_in_process > 0 else 0))

        sheet.write(14, index + 1,
                    round(balance_data["inventories"][0] / balance_data['total_assets'][0] * 100, 2) if not math.isnan(
                        balance_data["inventories"][0]) else 0)

        sheet.write(15, index + 1,
                    round(balance_data["account_receivable"][0] / balance_data['total_assets'][0] * 100,
                          2) if not math.isnan(
                        balance_data["account_receivable"][0]) else 0, get_style("account_receivable_rate", round(
                balance_data["account_receivable"][0] / balance_data['total_assets'][0] * 100, 2) if not math.isnan(
                balance_data["account_receivable"][0]) else 0))
        sheet.write(16, index + 1,
                    round(balance_data["good_will"][0] / balance_data['total_assets'][0] * 100, 2) if not math.isnan(
                        balance_data["good_will"][0]) else 0, get_style("good_will", round(
                balance_data["good_will"][0] / balance_data['total_assets'][0] * 100, 2)))

    for index, income_data in enumerate(income_summary_list):
        sheet.write(17, index + 1, round(income_data['total_operating_revenue'][0] / 100000000, 2))
        # 销售费用率
        sheet.write(21, index + 1,
                    round(income_data["sale_expense"][0] / income_data['total_operating_revenue'][0] * 100,
                          2) if not math.isnan(income_data["sale_expense"][0]) else 0, get_style("sale_expense_rate",
                                                                                                 round(income_data[
                                                                                                           "sale_expense"][
                                                                                                           0] /
                                                                                                       income_data[
                                                                                                           'total_operating_revenue'][
                                                                                                           0] * 100,
                                                                                                       2) if not math.isnan(
                                                                                                     income_data[
                                                                                                         "sale_expense"][
                                                                                                         0]) else 0))

        # 期间费用率 (管理费用+财务费用+销售费用+研发费用)/营业总收入
        period_expense = (income_data["sale_expense"][0] if not math.isnan(income_data["sale_expense"][0]) else 0) + (
            income_data["administration_expense"][0] if not math.isnan(
                income_data["administration_expense"][0]) else 0) + (
                             income_data["financial_expense"][0] if not math.isnan(
                                 income_data["financial_expense"][0]) else 0)
        period_expense += (
            stk_income_list[index]["rd_expenses"][0] if len(stk_income_list[index]["rd_expenses"] > 0) and
                                                        stk_income_list[index]["rd_expenses"][0] is not None else 0)
        # 期间费用率
        sheet.write(22, index + 1, round(period_expense / income_data['total_operating_revenue'][0] * 100, 2))

        sheet.write(23, index + 1,
                    round((period_expense / income_data['total_operating_revenue'][0] * 100) /
                          indicator_summary_list[index]['gross_profit_margin'][0] * 100,
                          2), get_style("gross_profit_margin_period_expense", round(
                (period_expense / income_data['total_operating_revenue'][0] * 100) /
                indicator_summary_list[index]['gross_profit_margin'][0] * 100,
                2)))
        # 营业利润
        sheet.write(24, index + 1, round(income_data["operating_profit"][0] / 100000000, 2))
        # 净利润
        sheet.write(25, index + 1, round(income_data["net_profit"][0] / 100000000, 2))
        # 主营利润率=主营利润/营业收入 主营利润=营业收入-营业成本-税金及附加-四费
        main_profit = income_data["operating_revenue"][0] - income_data["operating_cost"][0] - period_expense - (
            income_data["operating_tax_surcharges"][0] if not math.isnan(
                income_data["operating_tax_surcharges"][0]) else 0)
        sheet.write(26, index + 1, round(main_profit / income_data["operating_revenue"][0] * 100, 2),
                    get_style("main_profit_rate", round(main_profit / income_data["operating_revenue"][0] * 100, 2)))

        sheet.write(27, index + 1, round(main_profit / income_data["operating_profit"][0] * 100, 2),
                    get_style("income_main_profit_rate",
                              round(main_profit / income_data["operating_profit"][0] * 100, 2)))
        sheet.write(28, index + 1,
                    round(cash_summary_list[index]["net_operate_cash_flow"][0] / income_data["net_profit"][0] * 100, 2),
                    get_style("net_operate_cash_flow", round(
                        cash_summary_list[index]["net_operate_cash_flow"][0] / income_data["net_profit"][0] * 100, 2)))
    for index, indicator_data in enumerate(indicator_summary_list):
        sheet.write(18, index + 1, round(indicator_data['inc_total_revenue_year_on_year'][0], 2),
                    get_style("inc_total_revenue_year_on_year",
                              indicator_data['inc_total_revenue_year_on_year'][0]))
    # 毛利率
    for index, indicator_data in enumerate(indicator_summary_list):
        sheet.write(19, index + 1, round(indicator_data['gross_profit_margin'][0], 2),
                    get_style("gross_profit_margin", indicator_data['gross_profit_margin'][0]))

        if index == 0:
            sheet.write(20, 1, "-")
        else:
            gross_profit_margin_grown_rate = round(
                (indicator_data['gross_profit_margin'][0] - indicator_summary_list[index - 1]['gross_profit_margin'][
                    0]) /
                indicator_summary_list[index - 1]['gross_profit_margin'][0] * 100, 2)

            sheet.write(20, index + 1, gross_profit_margin_grown_rate,
                        get_style("gross_profit_margin_grown_rate", gross_profit_margin_grown_rate))
        # roe
        sheet.write(29, index + 1, indicator_data['roe'][0], get_style("roe", indicator_data['roe'][0]))
    for index, cash_data in enumerate(cash_summary_list):
        sheet.write(30, index + 1, round(
            cash_data["fix_intan_other_asset_acqui_cash"][0] / cash_data["net_operate_cash_flow"][0] * 100,
            2) if not math.isnan(cash_data["fix_intan_other_asset_acqui_cash"][0]) else 0,
                    get_style("fix_intan_other_asset_acqui_cash", round(
                        cash_data["fix_intan_other_asset_acqui_cash"][0] / cash_data["net_operate_cash_flow"][0] * 100,
                        2) if not math.isnan(cash_data["fix_intan_other_asset_acqui_cash"][0]) else 0))

        sheet.write(31, index + 1, round(
            cash_data["dividend_interest_payment"][0] / cash_data["net_operate_cash_flow"][0] * 100,
            2) if not math.isnan(cash_data["dividend_interest_payment"][0]) else 0,
                    get_style("dividend_interest_payment", round(
                        cash_data["dividend_interest_payment"][0] / cash_data["net_operate_cash_flow"][0] * 100,
                        2) if not math.isnan(cash_data["dividend_interest_payment"][0]) else 0))

    # 投资类资产
    # for index, data in enumerate(finance_balance_list):
    #
    #     investment_assets = (data["fairvalue_fianancial_asset"][0] if len(data["fairvalue_fianancial_asset"]) > 0 and
    #                                                                   data["fairvalue_fianancial_asset"][0] else 0) + (
    #                             data["bond_invest"][0] if len(data["bond_invest"]) > 0 and data["bond_invest"][
    #                                 0] else 0) + (
    #                             data["hold_for_sale_assets"][0] if len(data["hold_for_sale_assets"]) > 0 and
    #                                                                data["hold_for_sale_assets"][0] else 0) + (
    #                             data["other_equity_tools_invest"][0] if len(data["other_equity_tools_invest"]) > 0 and
    #                                                                     data["other_equity_tools_invest"][0] else 0) + (
    #                             data["other_bond_invest"][0] if len(data["other_bond_invest"]) > 0 and
    #                                                             data["other_bond_invest"][0] else 0) + (
    #                             data["hold_to_maturity_investments"][0] if len(
    #                                 data["hold_to_maturity_investments"]) > 0 and data["hold_to_maturity_investments"][
    #                                                                            0] else 0)
    #     + (
    #         stk_balance_list[index]["other_non_current_financial_assets"][0] if len(
    #             stk_balance_list[index]["other_non_current_financial_assets"]) > 0 and
    #                                                                                 stk_balance_list[index][
    #                                                                                     "other_non_current_financial_assets"][
    #                                                                                     0] else 0) + (
    #         data["investment_property"][0] if len(data["investment_property"]) > 0 and data["investment_property"][
    #             0] else 0) + (data["longterm_equity_invest"][0] if len(data["longterm_equity_invest"]) > 0 and
    #                                                                data["longterm_equity_invest"][0] else 0)
    #
    #     print(investment_assets)
    #     print(balance_summary_list[index]['total_assets'][0])
    #     sheet.write(31, index + 1, round(investment_assets / balance_summary_list[index]['total_assets'][0] * 100, 2),
    #                 get_style("investment_assets_rate",
    #                           round(investment_assets / balance_summary_list[index]['total_assets'][0] * 100, 2)))
    # for index, balance_data in enumerate(balance_summary_list):
    #     liability_rate = round(balance_data['total_liability'][0] /
    #                            balance_data['total_assets'][0] * 100, 2)
    #     sheet.write(4, index + 1, liability_rate, get_style("liability_rate", liability_rate))
    info = get_security_info(code)
    workbook.save("/Users/zhangzl/Desktop/pythonProject/" + info.display_name + ".xls")


if __name__ == '__main__':
    init_jqdata("002677.XSHE")
