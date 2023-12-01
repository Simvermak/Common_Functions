import zipfile
import os


def zip_files(file_paths, zip_path, password):
    with zipfile.ZipFile(zip_path, 'w') as zipf:

        # 将密码转换为字节数组
        password_bytes = password.encode('utf-8')

        # 设置压缩文件密码
        zipf.setpassword(password_bytes)
        
        for file_path in file_paths:
            # 获取文件的目录路径和文件名
            directory_path, file_name = os.path.split(file_path)
            # 将文件添加到压缩文件中，使用指定的文件名
            zipf.write(file_path, arcname=file_name)
