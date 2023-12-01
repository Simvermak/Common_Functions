import pandas as pd
from datetime import datetime
from sqlalchemy import text
import json
import sys
import os
from w3lib.http import basic_auth_header

sys.path.append("..")
import common.database as database
import ExcelTemplate as xlst


def main():
    '''
    说明: MES工厂库存物料寿命日报表
          BOM 相关部分代码(已注释)仅在处理报废物料时使用，日常推送不需要
    状态:
    '''

    factory = '松岗'
    today_str = datetime.today().strftime(r'%Y-%m-%d')

    # # BOM 料号获取
    # test_engine = database.test_engine('test_xu')
    # # 更新BOM视图
    # sql = '''
    #     create or replace view test_xu.sap_bom_view as
    #     select  
    #     sbh.self_id as ID , 
    #     sbh.WERKS,
    #     sbh.MATNR,
    #     sbb.parent_id as PARENT_ID,
    #     sbh.LOEKZ as 'MARK',
    #     sbh.STLST as 'STATUS',
    #     sbb.DUMPS as 'V_STATUS'
    #     from test_xu.sap_bom_head sbh
    #     left join test_xu.sap_bom_body sbb on sbb.IDNRK = sbh.MATNR
    #     WHERE sbh.WERKS = 1001
    #     AND (sbh.LOEKZ IS NULL OR sbh.LOEKZ = '')
    #     AND sbh.STLST != 3
    #     UNION 
    #     select  
    #     sbh.self_id  as ID , 
    #     sbb.WERKS,
    #     sbb.IDNRK as matnr,
    #     sbb.parent_id as PARENT_ID,
    #     sbh.LOEKZ as 'MARK',
    #     sbh.STLST as 'STATUS',
    #     sbb.DUMPS as 'V_STATUS'
    #     FROM test_xu.sap_bom_body sbb 
    #     left join test_xu.sap_bom_head sbh on sbb.IDNRK = sbh.MATNR and sbb.WERKS = sbh.WERKS
    #     WHERE sbb.WERKS = 1001
    #     AND (sbb.DUMPS IS NULL OR sbb.DUMPS = '')
    # '''
    # database.execute(sql, test_engine)
    # # BOM在用物料料号
    # bom_df = pd.read_sql("SELECT DISTINCT MATNR FROM sap_bom_view", test_engine)

    mes_engine = database.mes_engine('JRDATA')

    sql = f'''
        SELECT
            ProductRoot.ProductName AS product_code,
            Product.ProductDescription AS product_name,
            Product.ProductDescription_1 AS specs,
            center.WorkcenterName,
            center.WorkcenterDescription,
            Cell.CellName,
            WW_RMInventory.Qty,
            Factory.FactoryName,
            WW_RMInventory.EntryDate,
            product.VALIDITY AS material_validity_period,
            product.MINVALIDITY AS remaining_validity_period,
            WW_RMInventory.LotSN
        FROM
            WW_RMInventory
            LEFT JOIN Product ON WW_RMInventory.ProductId = Product.ProductId
            LEFT JOIN ProductRoot ON ProductRoot.ProductRootId = Product.ProductRootId
            LEFT JOIN Workcenter center ON WW_RMInventory.WorkcenterId = center.WorkcenterId
            LEFT JOIN Cell ON WW_RMInventory.CellId = Cell.CellId
            LEFT JOIN Factory ON WW_RMInventory.FactoryId = Factory.FactoryId 
        WHERE
            Product.MINVALIDITY IS NOT NULL 
            AND ( WW_RMInventory.IsIssueToPickingList <> 1 OR IsIssueToPickingList IS NULL ) 
            AND ( ( WW_RMInventory.LotStatus != 0 OR WW_RMInventory.LotStatus != 2 ) OR WW_RMInventory.LotStatus IS NULL ) 
            AND WW_RMInventory.Qty > 0 
            AND WW_RMInventory.EntryDate IS NOT NULL 
            AND ProductRoot.WorkcenterName IS NOT NULL 
            AND WorkcenterDescription LIKE '%{factory}%' 
            AND center.WorkcenterName IN ('1102','1103','1104','1105','1201','1203','1204','1205','1206','1302','1303','1304','1801','1902')
        ORDER BY
            ProductRoot.ProductName DESC,
            WW_RMInventory.EntryDate DESC;
        '''

    df = pd.read_sql(text(sql), mes_engine.connect())

    # 计算天数差 负值为已过期物料
    df['days_until_expiration'] = (df['material_validity_period'] - (datetime.now() - df['EntryDate']).dt.days).astype(int)

    # 仅筛选已过期和处于预警期的物料
    df = df[df['days_until_expiration'] <= df['remaining_validity_period']].reset_index(drop=True)

    # # 筛选BOM在用物料
    # using_df = df.merge(bom_df, left_on='product_code', right_on='MATNR', how='inner')
    # using_df = using_df.drop(columns=['MATNR'])
    # # 获取未在BOM中使用，需报废处理的物料
    # df = pd.concat([df, using_df])
    # df = df[~df.duplicated(keep=False)]

    # 排序 未过期在上,从小到大,过期在下,从大到小
    positive_values = df[df['days_until_expiration'] >= 0]
    negative_values = df[df['days_until_expiration'] < 0]
    positive_values_sorted = positive_values.sort_values(by='days_until_expiration')
    negative_values_sorted = negative_values.sort_values(by='days_until_expiration', ascending=False)
    df = pd.concat([negative_values_sorted,positive_values_sorted ])

    df = df.drop(columns=['remaining_validity_period'])

    df['EntryDate'] = df['EntryDate'].dt.strftime('%Y-%m-%d')

    def best_before_background(days):
        if days < 0:
            return '[[BG:H]]' + str(days)
        else:
            return '[[BG:I]]' + str(days)

    df['days_until_expiration'] = df['days_until_expiration'].apply(best_before_background)

    lo_infos = {'data': json.loads(df.to_json(orient='records')),
                'data_time': datetime.today().strftime(r'%Y-%m-%d 08:00:00'),  # 数仓同步数据时间为八点, 若出现变动需更改
                'report_date': today_str}

    root_path = f'report/mes_inventory/'
    file_path = f'{root_path}{factory}物料有效期报表{today_str}.xlsx'

    if not os.path.exists(root_path):
        os.makedirs(root_path)

    xlst.write2(lo_infos, file_path, 'MES-物料有效期报表', True)

    print('生成', file_path)

    return file_path


if __name__ == '__main__':
    main()
