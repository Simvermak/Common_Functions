import shelve
import random

def write(key,val):
    with shelve.open('shelve/data') as db:
        db[key] = val

def read(key):
    with shelve.open('shelve/data') as db:
        val = db[key]
    return val

def write_pwd(key):
    '''
    生成一个6位随机数字密码
    '''
    pwd = ''.join(random.choices('0123456789', k=6))
    write(key,pwd)
    return pwd
