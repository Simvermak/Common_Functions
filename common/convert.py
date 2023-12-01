import platform
if(platform.system()=='Windows'):
  import os
  os.environ['JAVA_HOME'] = r'C:\Program Files\Java\jdk-19'

import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook, License, PdfSaveOptions,PdfSecurityOptions

class excel:

    def __init__(self, excel_path,new_path,password):
        self.excel_path = excel_path
        self.new_path = new_path
        self.password = password

    def to_pdf(self):

        # Load Excel file
        workbook = Workbook(self.excel_path)

        # 加载License文件
        apcelllic = License()
        apcelllic.setLicense('license.xml')

        # pdf保存时的配置
        saveOption = PdfSaveOptions()

        # pdf权限配置
        securityOptions =  PdfSecurityOptions()
        if len(self.password)>0:
          securityOptions.setUserPassword(self.password)
        

        # 计算公式的结果
        # saveOption.setCalculateFormula(True)

        # 设置安全选项
        saveOption.setSecurityOptions(securityOptions)
        
        # 设置一页展示
        saveOption.setOnePagePerSheet(True)

        workbook.save(self.new_path, saveOption)
        # jpype.shutdownJVM()

    def add_pwd(self):

        # Load Excel file
        workbook = Workbook(self.excel_path)

        # 加载License文件
        apcelllic = License()
        apcelllic.setLicense('license.xml')

        workbook.getSettings().setPassword(self.password)

        workbook.save(self.new_path)

        # jpype.shutdownJVM()