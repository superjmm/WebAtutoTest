#coding=utf-8
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import ZzjgManage
from Util.handle_json import handle_json
from Util.handle_excel import handle_excel
from Util.handle_precondition import handle_preconditon
from Util.handle_log import handle_log

log_path = base_path+'/Report/log.txt' #日志路径
logger = logging.getLogger('接口测试日志')
#handle_log.set_log(logger,log_level='DEBUG',log_path=log_path) #设置日志等级及路径
handle_log.set_log(logger,log_level='DEBUG') #设置日志等级及屏幕输出
add_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
delete_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
add_path = base_path+'/Case/zzjg-add.xlsx' #添加测试用例
update_path = base_path+'/Case/zzjg-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/zzjg-delete.xlsx' #删除测试用例
delete_many_path = base_path+'/Case/zzjg-delete-many.xlsx' #批量删除测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 
delete_many_list = handle_excel.get_table_value(delete_many_path) 
zzjg_manage = ""

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global zzjg_manage
    zzjg_manage = ZzjgManage()

class TestZzjg():   
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('组织架构添加')
    def test_add_zzjg(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global zzjg_manage        
        global logger
        global add_path
        global add_row
        add_row += 1  
        n = add_row   
        test_result_col = 11 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = zzjg_manage.add(case)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('组织架构编辑')
    def test_update_zzjg(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global zzjg_manage
        global logger
        global update_path
        global update_row
        update_row += 1   
        n = update_row  
        test_result_col = 10 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = zzjg_manage.update(case)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('组织架构删除')
    def test_delete_zzjg(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global logger
        global zzjg_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        test_result_col = 8 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = zzjg_manage.delete(case)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',delete_many_list)
    @allure.feature('组织架构批量删除')
    def test_deleteMany_zzjg(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global logger
        global zzjg_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        test_result_col = 9 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = zzjg_manage.delete_many(case)
            if result:
                handle_excel.write_value(delete_many_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_many_path,n,test_result_col,'失败')
            assert result