import sys 
sys.path.append("..")
import common.convert as convert
import common.db as db
import common.utils as utils

import datetime,time
import os
import platform

import crm_sale,crm_work,crm_visit

def main():
    '''
    推送给事业部负责人
    '''

    task = '员工日报与客户拜访月度汇总'
    system  = 'CRM'

    dt = datetime.datetime.now()

    yesterday = dt - datetime.timedelta(days=1)
    yesterday_str = yesterday.strftime(r'%Y-%m-%d')

    # 密码key
    key_name='crm_pwd'

    # crm相关
    root_path = f'report/crm/'
    zip_path= f'员工日报与客户拜访月度汇总{yesterday_str}/'
    zip_name = f"CRM员工日报与客户拜访月度汇总{yesterday_str}.zip"

    def build():
        '''
        生成转换文件
        '''
        # 密码
        password=db.write_pwd(key_name)
        print('写入今日密码',password)

        crm_sale.main()
        crm_work.main()
        crm_visit.main()

        # 为了区分密码，重新压缩
        cmd=f'cd {root_path} && zip --password {password} {zip_name} {zip_path}*.xlsx'
        if(platform.system()=='Windows'):
            cmd = r'cd %s && C:\"Program Files"\WinRAR\WinRAR.exe a -p%s -mezl %s %s' % (root_path,password,zip_name,zip_path)
        os.system(cmd)
        print('生成 员工日报与客户拜访月度汇总')
        time.sleep(1)

    def push():
        '''
        发送文件
        '''
        # 微信配置
        corpid='ww79b6d766f2930869'
        corpsecret='qeWC3QNU_sB74QwaHpA_wEqveCmZPTDtORGvA7lH174'
        agentid='1000141'
        wxa=utils.weixin_app(corpid,corpsecret,agentid)
        token=wxa.get_token()

        # 密码
        password=db.read(key_name)
        print('读取今日密码',password)

        # 上传文件
        media_crm=wxa.get_media(token,os.path.abspath(root_path+zip_name))
        time.sleep(1)

        # 接收人
        #user_id='JianHangJianYuan'
        user_id='JianHangJianYuan|Tim.zhan|Clark.shanguan'
        # 由Clark手动转给|Harry.yang|Nancy.sun

        msg = f'{yesterday_str} 相关报告如下：\n文件动态密码：{password}'
        wxa.send_msg(token, user_id, msg, 'text')
        time.sleep(3)
        wxa.send_msg(token, user_id, media_crm, 'file')

    if len(sys.argv) > 1 and sys.argv[1]:
        argv = sys.argv[1]
        if argv == 'push':
            push()
            utils.task().update_task(system, task)
        elif argv == 'build':
            build()
            utils.task().save_task(system, task)
        else:
            print('未知参数', argv)
        exit()

    if dt.hour == 14 and dt.minute == 29:
        # 创建
        build()
        # 保存
        utils.task().save_task(system, task)
    elif dt.hour == 14 and dt.minute >= 30:
        # 读取
        utils.task().read_task(system, task)
        # 推送
        push()
        # 更新
        utils.task().update_task(system, task)

if __name__ == '__main__':
    main()