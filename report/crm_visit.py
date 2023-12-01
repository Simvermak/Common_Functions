import pandas as pd
from datetime import datetime, timedelta
from sqlalchemy import text
import numpy as np
import sys
import os
import common.database as database
import ExcelTemplate as xlst

sys.path.append("..")


def main():
    '''
        说明：客户拜访报告
        状态：还没完成
    '''

    print('客户拜访报告汇总')
    today = datetime.today()
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime(r'%Y-%m-%d')

    engine = database.engine('ICCPP_DW_DWS')

    # CRM客户信息与订单出货金额数据
    sql = '''
    SELECT
        cc.crm_code,
        cc.contactsName,
        cc.contactsPost,
        cc.contactsMobile,
        cc.email,
        cc.region3,
        cc.region4 AS area,
        cc.bos_code,
        bcs.no_order_days,
        bcs.month_order_amount,
        bcs.month_shipping_amount,
        bcs.year_order_amount,
        bcs.year_shipping_amount
    FROM
        crm_customer cc
    LEFT JOIN
        bos_customer_state bcs
    ON 
        cc.bos_code = bcs.bos_code
    '''
    amount_df = pd.read_sql(text(sql), engine.connect())
    print('客户信息', len(amount_df))

    # 拜访记录数据
    sql = '''
    SELECT 
      tf.targetIntro AS Customer_Name,
      tf.nickName AS Submitter,
      tf.createTime AS Submit_Time,
      tf.local_time,
      tf.form_key,
      tf.crm_code,
      tf.form_id,
      tf.time_zone,
      tfd.value,
      tfd.name,
      tfd.child_name,
      tf.email AS sales_email
    FROM
        ICCPP_DW_DWS.crm_form tf
    LEFT JOIN
        ICCPP_DW_DWS.crm_form_data tfd 
    ON
        tf.form_id = tfd.form_id
    WHERE
        tf.form_key IN ('7ErjR9Pu', 'epo32S8g')
        AND tf.local_time < DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-%d 23:59:59') 
        AND tf.local_time >= DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-01 00:00:00')
    '''
    visit_df = pd.read_sql(text(sql), engine.connect())
    print('拜访信息', len(visit_df))

    # 评论数据
    sql = '''
    SELECT
        remark,
        form_id
    FROM
        crm_comment
    '''
    remark_df = pd.read_sql(text(sql), engine.connect())
    print('评论信息', len(remark_df))

    def colMerge(row):
        '''
        表单 字段名+子字段名
        '''
        if len(row['child_name']) > 0:
            return row['name'] + "-" + row['child_name']
        return row['name']

    if not visit_df.empty:
        visit_df['name_str'] = visit_df.apply(colMerge, axis=1)
    else:
        visit_df['name_str'] = []

    # 提取标题和对应值
    base_col = ['Customer_Name', 'Submitter', 'Submit_Time', 'local_time', 'crm_code', 'form_key', 'form_id', 'time_zone', 'sales_email']
    visit_df = visit_df.pivot(index=base_col, columns='name_str', values='value').reset_index()

    visit_df.columns = visit_df.columns.str.replace(' ', '_')
    visit_df.rename(
        columns={'Order_Amount_Confirmed_in_the_Meeting_(CNY)': 'Order_Amount_Confirmed_in_the_Meeting',
                 'Customer_Contact_Info-Customer_Contact_Person': 'Customer_Contact_Person',
                 'Customer_Contact_Info-Business_Title': 'Business_Title',
                 'Customer_Contact_Info-Contact': 'Contact', 'form_key': 'visit_form'}, inplace=True)

    visit_df = visit_df.sort_values(by=['Submit_Time', 'Customer_Name'], ascending=(False, True))  # 根据日期与客户排序，便于展示
    merge_df = pd.merge(visit_df, amount_df, how='left', left_on='crm_code', right_on='crm_code')

    # 接单金额为0处理
    merge_df['no_order_days'].fillna('暂未接单', inplace=True)

    # 拜访方式
    merge_df['visit_form'] = merge_df['visit_form'].replace({'7ErjR9Pu': 'Offline', 'epo32S8g': 'Online'})
    if not visit_df.empty:
        merge_df['Submit_Time'] = merge_df['Submit_Time'].astype(str)
        merge_df['local_time'] = merge_df['local_time'].astype(str)
        merge_df['Order_Amount_Confirmed_in_the_Meeting'] = merge_df['Order_Amount_Confirmed_in_the_Meeting'].astype(
            float)

    # 时区换算
    # def process_data(row):
    #     submit_date = row['Submit_Time']
    #     time_zone = row['time_zone']
    #     date_of_visit = row['Date_of_Visit']
    #     if isinstance(time_zone, str):
    #         # 创建目标时区对象
    #         target_offset = timedelta(hours=int(time_zone[3:6]), minutes=int(time_zone[7:9]))
    #         target_tz = timezone(target_offset)
    #
    #         # 北京时间
    #         beijing_offset = timedelta(hours=8)
    #         beijing_tz = timezone(beijing_offset)
    #
    #         # 将时间字符串转换为datetime对象
    #         submit_date_beijing = datetime.strptime(submit_date, '%Y-%m-%d %H:%M:%S')
    #         visit_date_local = datetime.strptime(date_of_visit, '%Y-%m-%d').replace(hour=0, minute=0, second=0)
    #
    #         visit_date_local = visit_date_local.replace(tzinfo=target_tz)  # 将原始时区转换为目标时区
    #
    #         # 将原始时间转换为目标时区的时间
    #         submit_date_local = submit_date_beijing.astimezone(target_tz).replace(tzinfo=None).strftime(
    #             '%Y-%m-%d %H:%M:%S')
    #         visit_date_beijing = visit_date_local.astimezone(beijing_tz).replace(tzinfo=None).strftime('%Y-%m-%d')
    #
    #         return submit_date_local, visit_date_beijing
    #
    # merge_df[['Submit_Date_Local', 'Visit_Date_Beijing']] = merge_df.apply(process_data, axis=1, result_type='expand')

    # 客户联系人联系方式合并
    merge_df['contactsMobile'] = merge_df['contactsMobile'].str.replace("'", "")
    merge_df['combined_contact'] = np.where(merge_df['contactsMobile'].notnull() & merge_df['email'].notnull(),
                                            merge_df['contactsMobile'] + '\r\n' + merge_df['email'],
                                            np.where(merge_df['contactsMobile'].notnull(), merge_df['contactsMobile'],
                                                     merge_df['email']))

    # 将多个评论合并
    remark_df = remark_df.groupby('form_id')['remark'].apply('\r\n'.join).reset_index()

    # 与评论表合并
    merge_df['form_id'] = merge_df['form_id'].astype(str)
    merge_df = pd.merge(merge_df, remark_df, how="left", left_on="form_id", right_on="form_id")

    # 处理缺省值
    merge_df['month_order_amount'].fillna(0, inplace=True)
    merge_df['month_shipping_amount'].fillna(0, inplace=True)
    merge_df['year_order_amount'].fillna(0, inplace=True)
    merge_df['year_shipping_amount'].fillna(0, inplace=True)

    # title字段
    merge_df['collected_time'] = today.strftime('%Y-%m-%d %H:%M:%S')
    merge_df['local_date'] = merge_df['local_time'].str[:10]

    # 超链接
    merge_df['hyperlink'] = '=HYPERLINK("[销售跟进情况汇总（月）"&TEXT(TODAY()-1,"yyyy-mm-dd")&".xlsx]\'销售跟进情况汇总（月）\'!N2","<< 返回跟进情况页")'

    # 工作簿名
    merge_df['sheet_name'] = merge_df['local_date'] + "拜访 " + merge_df['form_id']

    if merge_df.empty:
        new_row = pd.DataFrame({'sheet_name': ['无客户拜访报告']})
        merge_df = pd.concat([merge_df, new_row], ignore_index=True)

    data_list = merge_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
    lo_infos = [data_list[x] for x in range(len(data_list))]

    # 汇总
    root_path = f'report/crm/员工日报与客户拜访月度汇总{yesterday_str}/'
    file_path = f'{root_path}引用附件勿删BF.xlsx'

    if not os.path.exists(root_path):
        os.makedirs(root_path)

    tlp_name = 'CRM-客户拜访报告'
    xlst.write2(lo_infos, file_path, tlp_name, False)
    print('生成', file_path)
    print()

    # 北美大区
    print('生成北美拜访报告')
    selected_df = merge_df[merge_df['region3'] == '北美大区'].reset_index(drop=True)
    if not selected_df.empty:
        na_data_list = selected_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
        lo_infos = [na_data_list[x] for x in range(len(na_data_list))]
        root_path2 = f'report/crm/北美员工日报与客户拜访月度汇总{yesterday_str}/'
        file_path2 = f'{root_path2}引用附件勿删BF.xlsx'
        if not os.path.exists(root_path2):
            os.makedirs(root_path2)

        tlp_name = 'CRM-客户拜访报告'
        xlst.write2(lo_infos, file_path2, tlp_name, False)
        print('生成', file_path2)
        print()

    return file_path

    # # 筛选出昨天的数据
    # print('生成昨日报表')
    # yesterday_df = merge_df[merge_df['local_date'] == yesterday_str].reset_index(drop=True).drop('hyperlink', axis=1)
    # if not yesterday_df.empty:
    #     yesterday_data_list = yesterday_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
    #     lo_infos = [yesterday_data_list[x] for x in range(len(yesterday_data_list))]
    #     file_path2 = f'report/crm/客户拜访日报{yesterday_str}.xlsx'
    #     tlp_name = 'CRM-客户拜访报告'
    #     xlst.write2(lo_infos, file_path2, tlp_name, False)
    #     print('生成', file_path2)
    # else:
    #     print('昨日无拜访报告')
    #


if __name__ == '__main__':
    main()
