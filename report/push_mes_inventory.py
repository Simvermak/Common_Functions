import sys 
sys.path.append("..")
import common.db as db
import common.utils as utils
import common.convert as convert

import datetime,time
import os
import platform
import mes_inventory

def main():
    '''
    松岗工厂
    '''

    task = '松岗物料有效期报表'
    system  = 'MES'
    dt = datetime.datetime.now()
    today_str = datetime.datetime.today().strftime(r'%Y-%m-%d')

    # 密码key
    key_name='mes_pwd'
    path= f'report/mes_inventory/{task}{today_str}.xlsx'

    def push():
        '''
        发送文件
        '''

        # 密码
        password = db.read(key_name)
        print('获取密码', password)

        msg = f'{today_str} 相关报告如下：\n文件动态密码：{password}'

        # 发送群 (松岗产品保质期预警)
        # 正式
        wx = utils.weixin('ea12a322-060b-4098-8c58-d32991f359a3')
        # 测试
        # wx = utils.weixin('02bc1475-82ec-4a53-948a-b0eabf3dc86d')

        media_id = wx.upload_file(os.path.abspath(path))
        wx.send_text(msg)
        wx.send_file(media_id)

    def build():
        '''
        生成转换文件
        '''

        # 密码
        password = db.write_pwd(key_name)
        print('今日密码', password)

        print(f'{task}')
        file_path = mes_inventory.main()

        excel = convert.excel(file_path, file_path, password)
        excel.add_pwd()

    # 手动命令
    if len(sys.argv) > 1 and sys.argv[1]:
        argv = sys.argv[1]
        if argv == 'push':
            push()
            utils.task().update_task(system,task)
        elif argv == 'build':
            build()
            utils.task().save_task(system,task)
        else:
            print('未知参数', argv)
        exit()

    if dt.hour == 8 and dt.minute == 0:
        # 创建
        build()
        # 保存
        utils.task().save_task(system, task)
    elif dt.hour == 8:
        # 读取
        utils.task().read_task(system, task)
        # 推送
        push()
        # 更新
        utils.task().update_task(system,task)
if __name__ == '__main__':
    main()
