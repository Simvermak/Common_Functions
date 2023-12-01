import json
import pandas as pd
from datetime import datetime, timedelta
from sqlalchemy import text
import sys
import os

sys.path.append("..")
import common.database as database
import ExcelTemplate as xlst


def main():
    '''
    说明：客户与销售行为报告
    状态：
    '''
    print('销售跟进情况汇总（月）')

    today = datetime.today()
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime(r'%Y-%m-%d')

    engine = database.engine('ICCPP_DW_DWS')

    # 用户信息
    sql = f'''
    SELECT
      region3,
      region4,
      realname,
      systemEmail as email
    FROM
      ICCPP_DW_DWS.crm_user
    WHERE
      status=1
      and region2 = '海外营销中心'
      and
        post in ('Sales', 'Sales Support', 'Sales Manager', 'Leader')
    '''
    df_user = pd.read_sql(text(sql), engine.connect())
    print('用户信息', len(df_user))

    # 客户信息
    sql = f'''
    SELECT
      crm_code,
      businessName
    FROM
      ICCPP_DW_DWS.crm_customer
    '''
    df_custom = pd.read_sql(text(sql), engine.connect())
    print('客户信息', len(df_custom))

    # 拜访信息
    sql = f'''
    SELECT
      a.email,
      a.targetIntro,
      a.nickName,
      b.value as visit_date,
      c.value as visit_amount,
      d.value as report_date,
      a.form_key,
      a.form_id,
      a.createTime,
      a.local_time,
      a.crm_code
    FROM
      ICCPP_DW_DWS.crm_form a
    left join ICCPP_DW_DWS.crm_form_data b on
      a.form_id = b.form_id
      and b.name = 'Date of Visit'
    left join ICCPP_DW_DWS.crm_form_data c on
      a.form_id = c.form_id
      and c.name = 'Order Amount Confirmed in the Meeting (CNY)'
    left join ICCPP_DW_DWS.crm_form_data d on
        a.form_id = d.form_id
      and d.name = 'Date of Report'
    WHERE
      a.form_key IN ('7ErjR9Pu', 'epo32S8g', 'MIVvHNoV')
      AND a.local_time < DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-%d 23:59:59') 
      AND a.local_time >= DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-01 00:00:00')
    '''
    df_visit = pd.read_sql(text(sql), engine.connect())
    df_visit['form_key'] = df_visit['form_key'].replace(
        {'7ErjR9Pu': 'Offline', 'epo32S8g': 'Online', 'MIVvHNoV': 'Daily report'})

    # 判断是否为 Offline 或者 Online
    mask = (df_visit['form_key'] == 'Offline') | (df_visit['form_key'] == 'Online')
    daily_report_mask = (df_visit['form_key'] == 'Daily report')

    # 根据报告类型填入不同的超链接
    # 拜访报告超链接
    if mask.any():
        df_visit.loc[mask, 'visit_report'] = (
                    df_visit.loc[mask, 'local_time'].astype(str).str[:10] + '拜访 ' + df_visit.loc[
                mask, 'form_id'].astype(str))
        df_visit.loc[mask, 'visit_report'] = (
                    '=HYPERLINK("[引用附件勿删BF.xlsx]\'' + df_visit.loc[mask, 'visit_report'] + '\'!A1","查看报告")')
    # 工作日报超链接
    if daily_report_mask.any():
        df_visit.loc[daily_report_mask, 'visit_report'] = (
                df_visit.loc[daily_report_mask, 'local_time'].astype(str).str[:10] + '工作日报 ' + df_visit.loc[
            daily_report_mask, 'form_id'].astype(str))
        df_visit.loc[daily_report_mask, 'visit_report'] = ('=HYPERLINK("[引用附件勿删RB.xlsx]\'' + df_visit.loc[
            daily_report_mask, 'visit_report'] + '\'!A1","查看报告")')

    df_visit['visit_amount'] = df_visit['visit_amount'].astype(float)

    # 工作日报
    df_visit_work = df_visit[df_visit['form_key'] == 'Daily report'].copy()
    df_visit_work['visit_date'] = df_visit_work['report_date']
    print('工作日报', len(df_visit_work))

    # 拜访客户
    df_visit_custom = df_visit[df_visit['form_key'] != 'Daily report'].copy()
    print('拜访客户', len(df_visit_custom))
    # df_visit_work['report_date_bj']=pd.to_datetime(df_visit_work['visit_date'])+ pd.DateOffset(days=1)
    # #df_visit_work['report_date_bj']=(df_visit_work['createTime']-df_visit_work['local_time']).dt.days
    # print(df_visit_work['report_date_bj'])

    # 合并成拜访信息
    df_visit = pd.concat([df_visit_work, df_visit_custom], axis=0)
    print('合并拜访', len(df_visit))

    # 用户合并拜访
    df = df_user.merge(df_visit, on='email', how='left')
    # 再合并客户
    df = pd.merge(df, df_custom, how='left', on='crm_code')

    # 指定排序顺序
    custom_order = ['北美大区', '欧洲大区']
    df['has_report'] = df['form_key'].notnull()
    # 按指定顺序排序
    df['order'] = df['region3'].map({v: i for i, v in enumerate(custom_order)})
    df = df.sort_values(by=['order', 'region4', 'has_report', 'realname', 'visit_date'],
                        ascending=[True, True, False, True, False]).drop('order', axis=1)

    # 缺省值处理               
    df.fillna("--", inplace=True)

    df['createTime'] = df['createTime'].astype(str)
    df['local_time'] = df['local_time'].astype(str)

    # 汇总
    file_name = '销售跟进情况汇总（月）'
    root_path = f'report/crm/员工日报与客户拜访月度汇总{yesterday_str}/'
    file_path = f'{root_path}{file_name}{yesterday_str}.xlsx'

    if not os.path.exists(root_path):
        os.makedirs(root_path)

    # 传递给excel
    lo_infos = {}

    # 工作簿名字
    lo_infos['sheet_name'] = file_name

    lo_infos['voopooRows'] = json.loads(df.to_json(orient='records'))
    # 数据时间
    lo_infos['data_time'] = today.strftime(r'%Y-%m-%d %H:%M:%S')
    # 报告日期
    lo_infos['report_date'] = yesterday_str
    lo_infos['report_year'] = yesterday.year
    # 昨天所处月份
    lo_infos['report_month'] = yesterday.month

    xlst.write2(lo_infos, file_path, 'CRM-' + file_name, False)
    print('生成', file_path)
    print()

    # 北美大区
    selected_df = df[df['region3'] == '北美大区'].reset_index(drop=True)
    root_path2 = f'report/crm/北美员工日报与客户拜访月度汇总{yesterday_str}/'
    file_path2 = f'{root_path2}{file_name}{yesterday_str}.xlsx'
    if not os.path.exists(root_path2):
        os.makedirs(root_path2)
    lo_infos['voopooRows'] = json.loads(selected_df.to_json(orient='records'))
    xlst.write2(lo_infos, file_path2, 'CRM-' + file_name, False)
    print('生成', file_path2)
    print()

    return file_path

if __name__ == '__main__':
    main()
