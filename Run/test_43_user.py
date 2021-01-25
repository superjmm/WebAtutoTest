#coding=utf-8
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import UserManage
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
add_path = base_path+'/Case/setup-user-add1.xlsx' #添加测试用例
update_path = base_path+'/Case/setup-user-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/setup-user-delete.xlsx' #删除测试用例
add_ctl_path = base_path+'/Case/setup-user-add2-control.xlsx' #添加测试用例
add_gp_path = base_path+'/Case/setup-user-add3-group.xlsx' #添加测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 
add_ctl_list = handle_excel.get_table_value(add_ctl_path)
add_gp_path = handle_excel.get_table_value(add_gp_path)

user_manage = ''

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global user_manage
    user_manage = UserManage()

class TestUser():    

    @pytest.mark.parametrize('case',add_list)
    @allure.feature('用户添加')
    def test_add(self,case):
        pytest.skip()
        global user_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        test_result_col = 12 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = user_manage.add(case)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('用户编辑')
    def test_update(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global user_manage
        global update_row
        global logger
        update_row += 1  
        n = update_row   
        test_result_col = 15 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = user_manage.update(case)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('热力站删除')
    def test_delete(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global user_manage
        global logger
        global delete_row
        delete_row += 1  
        n = delete_row   
        case_id = case[0]
        is_execute = case[2]
        test_result_col = 7 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = user_manage.delete(case)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result
    
    