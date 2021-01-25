#coding=utf-8
import json
from jsonpath_rw import parse
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
from Util.handle_json import handle_json

class HandlePrecondition:
    '''
    处理前置条件和依赖key
    '''    
        
    def get_depend_data(self,file_path,pre_data):
        '''
        通过前置条件中的用例编号，获取到依赖的结果集
        '''
        case_id = pre_data.split('>')[0] #将前置条件字符串切割,得到依赖用例编号
        depend_data = handle_json.get_value(file_path,case_id)
        return depend_data

    def get_depend_value(self,file_path,pre_data):
        '''
        在数据集中，拿到依赖key的值
        '''
        depend_data = self.get_depend_data(file_path,pre_data)
        depend_data = json.loads(depend_data)
        depend_rule = pre_data.split('>')[1] #将前置条件字符串切割得到依赖规则
        json_exe = parse(depend_rule) #解析规则
        madle = json_exe.find(depend_data) #根据规则，在数据集中匹配key
        depend_value = [math.value for math in madle][0] 
        return depend_value
    def update_data_value(self,file_path,pre_data,depnd_key,data):
        '''
        更新data数据集中的依赖key值
        '''
        depend_value = self.get_depend_value(file_path,pre_data)
        data = json.loads(data)
        data[depnd_key] = depend_value
        return data

handle_preconditon = HandlePrecondition()

