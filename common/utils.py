import os
import glob
from common.idgenerator import options, generator
import requests
import json
from datetime import datetime
from requests_toolbelt import MultipartEncoder
from selenium import webdriver

class files:
    def last_file(excel_folder):
        '''
        获取最新的文件路径
        '''

        # 获取 Excel 文件夹中所有非隐藏的 .xlsx 文件的路径
        excel_files = [f for f in glob.glob(
            excel_folder) if not os.path.basename(f).startswith('~')]

        # 按文件修改时间排序
        excel_files.sort(key=os.path.getmtime)

        # 按文件修改时间排序
        latest_excel_file = excel_files[-1]

        print('最新文件', latest_excel_file)
        return latest_excel_file


class ids:
    options = options.IdGeneratorOptions(worker_id=23)
    idgen = generator.DefaultIdGenerator()
    idgen.set_id_generator(options)

    def snow_flake(self):
        '''
        # 雪花id
        '''
        return self.idgen.next_id()


class weixin:
    def __init__(self, key='69673656-6960-4b4a-972c-caaaffb35fc8'):
        self.url_upload = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key=%s&type=file' % key
        self.url_send = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=%s' % key

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:50.0) Gecko/20100101 Firefox/50.0',
        'Content-Type': 'application/json'
    }

    def upload_file(self, path):
        files = {'file': open(path, 'rb')}
        result = requests.post(url=self.url_upload, files=files)
        info = json.loads(result.text)
        if info['errcode'] == 0:
            return info['media_id']
        else:
            print(info)

    def send_file(self, media_id):
        data = {
            "msgtype": "file",
            "file": {
                "media_id": media_id
            }
        }
        res = requests.post(url=self.url_send, data=json.dumps(
            data), headers=self.headers)
        print('发送文件：'+media_id, res.text)

    def send_text(self, msg):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data = {
            "msgtype": "text",
            "text": {
                # "content": msg+'\n--%s' % now,
                "content": msg,
                # "mentioned_list":["wangqing","@all"],
                # "mentioned_mobile_list":["13800001111","@all"]
            }
        }
        res = requests.post(url=self.url_send, data=json.dumps(
            data), headers=self.headers)
        print('发送消息：'+msg, res.text)


class weixin_app:
    def __init__(self, corpid='ww79b6d766f2930869', corpsecret='x2R7B1DnvmPzePEAfnseZpBj2uCJMfgHYqwP8psD7A0',agentid=1000027):
        self.corpid = corpid
        self.corpsecret = corpsecret
        self.agentid = agentid

    def get_token(self):
        '''
        获取token
        '''
        result = requests.get(
            f'https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={self.corpid}&corpsecret={self.corpsecret}')
        info = json.loads(result.text)
        if info['errcode'] == 0:
            return info['access_token']
        else:
            print(info)

    def get_media(self,access_token,file_path):
        '''
        上传临时素材
        '''
        file_name = os.path.basename(file_path)
        post_file_url = f"https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={access_token}&type=file"
        m = MultipartEncoder(
            fields={
                file_name: (file_name, open(file_path, 'rb'), 'text/plain')
            }
        )
        try:
            result = requests.post(url=post_file_url, data=m, headers={
                                   'Content-Type': m.content_type})
        except Exception as e:
            print(f"error: {e}")
        info = json.loads(result.text)
        if info['errcode'] == 0:
            return info['media_id']
        else:
            print(info)

    def send_msg(self,access_token,user_id, msg,msgtype):
        '''
        发送消息
        '''

        if msgtype=='file':
            contype='media_id'
        else:
            contype='content'

        url = f'https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}'
        data = {
            "touser": user_id,
            "msgtype": msgtype,
            "agentid": self.agentid,
            msgtype: {
                contype: msg
            },
        }
        return requests.post(url, data=json.dumps(data))
    
    class simulate_login():
        def create_chrome_driver(*,headless= False):
            '''
            创建浏览器
            '''
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument('--headless')
                options.add_argument('--log-level=3')
            options.add_experimental_option('excludeSwitches',['enable-automation'])
            options.add_experimental_option('useAutomationExtension',False)
            browser = webdriver.Chrome(options=options)
            browser.execute_cdp_cmd(
                'Page.addScriptToEvaluateOnNewDocument',
                {'source':'Object.defineProperty(navigator,"webdriver",{get:()=> undefined})'}
            )
            return browser