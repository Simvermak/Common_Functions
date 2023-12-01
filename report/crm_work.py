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
    说明：工作日报
    状态：还没完成
    '''

    print('工作日报汇总')
    today = datetime.today()
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime(r'%Y-%m-%d')

    engine = database.engine('ICCPP_DW_DWS')

    # 工作日报数据
    sql = '''
    SELECT 
      tf.targetIntro,
      tf.nickName AS Submitter,
      tf.createTime AS Submit_Time,
      tf.local_time,
      tf.form_key,
      tf.form_id,
      tfd.value,
      tfd.name,
      tf.email,
      cu.region2,
      cu.region3,
      cu.region4
    FROM
        ICCPP_DW_DWS.crm_form tf
    LEFT JOIN
        ICCPP_DW_DWS.crm_form_data tfd 
    ON
        tf.form_id = tfd.form_id
    LEFT JOIN
        crm_user cu
    ON 
        tf.email = cu.systemEmail
    WHERE
        tf.form_key = 'MIVvHNoV'
        AND tf.local_time < DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-%d 23:59:59') 
        AND tf.local_time >= DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 1 DAY), '%Y-%m-01 00:00:00')
    '''
    daily_report_df = pd.read_sql(text(sql), engine.connect())
    print('工作日报信息', len(daily_report_df))

    # 提取标题和对应值
    base_col = ['targetIntro', 'Submitter', 'Submit_Time', 'local_time', 'form_key', 'form_id', 'email', 'region2','region3',
                'region4']
    daily_report_df = daily_report_df.pivot(index=base_col, columns='name', values='value')
    daily_report_df = daily_report_df.reset_index()

    daily_report_df.columns = daily_report_df.columns.str.replace(' ', '_')

    # 报告形式
    daily_report_df['form_key'] = daily_report_df['form_key'].replace('MIVvHNoV', 'Daily report')

    daily_report_df['local_time'] = daily_report_df['local_time'].astype(str)

    # title字段
    daily_report_df['local_date'] = daily_report_df['local_time'].str[:10]
    daily_report_df['collected_time'] = today.strftime('%Y-%m-%d %H:%M:%S')
    daily_report_df['hyperlink'] = '=HYPERLINK("[销售跟进情况汇总（月）" & TEXT(TODAY() - 1, "yyyy-mm-dd") & ".xlsx]\'销售跟进情况汇总（月）\'!N2","<< 返回跟进情况页")'

    daily_report_df = daily_report_df.sort_values(by=['local_time'], ascending=False).reset_index(drop=True)

    # 工作簿名
    daily_report_df['sheet_name'] = daily_report_df['local_date'].str[:10] + "工作日报 " + daily_report_df['form_id'].astype(str)

    data_list = daily_report_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
    lo_infos = [data_list[x] for x in range(len(data_list))]

    # 汇总 
    root_path = f'report/crm/员工日报与客户拜访月度汇总{yesterday_str}/'
    file_path = f'{root_path}引用附件勿删RB.xlsx'

    if not os.path.exists(root_path):
        os.makedirs(root_path)

    tlp_name = 'CRM-员工工作日报'
    xlst.write2(lo_infos, file_path, tlp_name, False)
    print('生成', file_path)
    print()

    # # 筛选出昨天的数据
    # print('生成昨日工作日报报表')
    # yesterday_df = daily_report_df[daily_report_df['local_date'] == yesterday_str].drop('hyperlink', axis=1)
    # if not yesterday_df.empty:
    #     yesterday_data_list = yesterday_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
    #     print(yesterday_data_list)
    #     lo_infos = [yesterday_data_list[x] for x in range(len(yesterday_data_list))]
    #     file_path2 = f'report/crm/员工工作报告{yesterday_str}.xlsx'
    #     tlp_name = 'CRM-员工工作日报'
    #     xlst.write2(lo_infos, file_path2, tlp_name, False)
    #     print('生成', file_path2)
    # else:
    #     print('昨日无工作报告')

    # 北美大区
    print('生成北美工作日报')
    selected_df = daily_report_df[daily_report_df['region3'] == '北美大区'].reset_index(drop=True)
    if not selected_df.empty:
        na_data_list = selected_df.to_dict('records')  # 将df转为每个元素为单独一行数据且列名为key的字典的列表
        lo_infos = [na_data_list[x] for x in range(len(na_data_list))]
        root_path2 = f'report/crm/北美员工日报与客户拜访月度汇总{yesterday_str}/'
        file_path2 = f'{root_path2}引用附件勿删RB.xlsx'

        if not os.path.exists(root_path2):
            os.makedirs(root_path2)

        tlp_name = 'CRM-客户拜访报告'
        xlst.write2(lo_infos, file_path2, tlp_name, False)
        print('生成', file_path2)
        print()


    return file_path


if __name__ == '__main__':
    main()
