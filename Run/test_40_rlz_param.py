#coding=utf-8
import sys
import os
import py
base_path = os.getcwd()
sys.path.append(base_path)

import pytest
import allure
import logging
from Operate.backstage_manage import RlzManage
from Util.handle_excel import handle_excel
from Util.handle_log import handle_log

log_path = base_path+'/Report/log.txt' #日志路径
logger = logging.getLogger('接口测试日志')
#handle_log.set_log(logger,log_level='DEBUG',log_path=log_path) #设置日志等级及路径
handle_log.set_log(logger,log_level='DEBUG') #设置日志等级及屏幕输出
add_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
delete_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_param_path = base_path+'/Case/rlz-param-update.xlsx'
add_param_path = base_path+'/Case/rlz-param-add.xlsx'
add_param_list =  handle_excel.get_table_value(add_param_path)
update_param_list =  handle_excel.get_table_value(update_param_path)
rlz_manage = ''

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global rlz_manage
    rlz_manage = RlzManage()

class TestRlzParam():                
   
    @pytest.mark.parametrize('case',add_param_list)
    @allure.feature('热力站-添加点配置')
    def test_add_param(self,case):
        global rlz_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        #测试结果所在的列
        test_result_col = 34
        #用例编号不能为空
        if case_id is None: 
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        #判断该条用例是否执行
        elif is_execute == 'no' : 
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        # 执行用例
        else: 
            result = rlz_manage.add_param(case)
            if result:
                handle_excel.write_value(add_param_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_param_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',update_param_list)
    @allure.feature('热力站-编辑点配置')
    def test_update_param(self,case):
        global rlz_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        #测试结果所在的列
        test_result_col = 35
        #用例编号不能为空
        if case_id is None: 
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        #判断该条用例是否执行
        elif is_execute == 'no' : 
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        # 执行用例
        else: 
            result = rlz_manage.update_param(case)
            if result:
                handle_excel.write_value(update_param_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_param_path,n,test_result_col,'失败')
            assert result