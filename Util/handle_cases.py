#coding=utf-8
from deepdiff import DeepDiff
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
from Util.handle_json import handle_json

class HandleAssert:
    '''
    对返回结果验证
    '''

    def assert_json_format(self,expect,actual,fuzzy,dic=None):
        '''
        校验json
        param:
            expect:期望值，dict类型
            actual:实际值，dict类型
            fuzzy:是否模糊查询，True/False
            dic:验证预期结果中是否存在某些值，dict类型
        '''
        if isinstance(expect,dict) and isinstance(actual,dict):
            cmp_dict = DeepDiff(expect,actual,ignore_order=True).to_dict()
            if cmp_dict.get('dictionary_item_added') or cmp_dict.get('dictionary_item_removed'):
                return False
            else:
                if dic is not None and not self.assert_dict_value(dic,actual,fuzzy):
                    return False
                else:
                    return True
        return False

    def assert_json_list(self,expect,actual,fuzzy,dic=None):
        '''
        校验list
        param:
            expect:期望值，dict类型
            actual:实际值，dict类型
            fuzzy:是否模糊查询，True/False
            dic:验证预期结果中是否存在某些值，dict类型
        '''
        Flag = False
        if len(actual)==0 and len(expect) == 0:
            return True
        if len(actual):
            for a in actual:
                if self.assert_json_format(expect[0],a,fuzzy,dic) == False:
                    Flag = False
                    break
                Flag = True
        return Flag
    def assert_dict_value(self,expect,actual,fuzzy):
        '''
        判断字典expect在actual中
        param:
            expect:期望值，dict类型
            actual:实际值，dict类型
            fuzzy:是否模糊查询，True/False
        '''
        flag = False
        for key in expect:
            if key in actual:
                if fuzzy:
                    flag = expect[key] in actual[key]
                else:
                    flag = (actual[key]==expect[key])
        return flag

handle_assert = HandleAssert()