import sys
import common.convert as convert
import common.db as db
import common.utils as utils
import datetime, time
import os
import operate_report
import platform

sys.path.append("..")


def main():

    task = '经营运营专项数据指标日报告'
    system  = 'OA'

    dt = datetime.datetime.now()

    yesterday = dt - datetime.timedelta(days=1)
    yesterday_str = yesterday.strftime(r'%Y-%m-%d')

    # 密码key
    key_name = 'operate_report'
    operate_path = f'report/operate/经营运营专项数据指标 日报告 {yesterday_str}.xlsx'

    def push():
        '''
        发送文件
        '''

        # 密码
        password = db.read(key_name)
        print('获取密码', password)

        msg = f'{yesterday_str} 相关报告如下：\n文件动态密码：{password}'

        # 发送群 ()
        wx = utils.weixin('e9a29181-5682-4e98-a457-39885cca2632')
        # 测试
        # wx = utils.weixin('02bc1475-82ec-4a53-948a-b0eabf3dc86d')
        media_id = wx.upload_file(os.path.abspath(operate_path))
        wx.send_text(msg)
        wx.send_file(media_id)

    def build():
        '''
        生成转换文件
        '''

        # 密码
        password = db.write_pwd(key_name)
        print('今日密码', password)

        print('经营运营专项数据指标（日报）')
        file_path = operate_report.main()

        excel = convert.excel(file_path, file_path, password)
        excel.add_pwd()

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

    if dt.hour == 10 and dt.minute == 0:
        # 创建
        build()
        # 保存
        utils.task().save_task(system, task)
    elif dt.hour == 10:
        # 读取
        utils.task().read_task(system, task)
        # 推送
        push()
        # 更新
        utils.task().update_task(system,task)

if __name__ == '__main__':
    main()
