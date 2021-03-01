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
add_path = base_path+'/Case/rlz-add1.xlsx' 
update_path = base_path+'/Case/rlz-update.xlsx' 
delete_path = base_path+'/Case/rlz-delete.xlsx' 
add_ctl_path = base_path+'/Case/rlz-add2-control.xlsx' 
add_gp_path = base_path+'/Case/rlz-add3-group.xlsx' 
add_gather_path = base_path+'/Case/rlz-param-gather.xlsx'
add_control_path = base_path+'/Case/rlz-param-control.xlsx'
copy_path = base_path+'/Case/rlz-copy-new.xlsx'
update_gather_path = base_path+'/Case/rlz-param-gather-update.xlsx'
update_param_path = base_path+'/Case/rlz-param-update.xlsx'
add_param_path = base_path+'/Case/rlz-param-add.xlsx'
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 
add_ctl_list = handle_excel.get_table_value(add_ctl_path)
add_gp_list = handle_excel.get_table_value(add_gp_path)
add_gather_list = handle_excel.get_table_value(add_gather_path)
add_control_list = handle_excel.get_table_value(add_control_path)
copy_list = handle_excel.get_table_value(copy_path)
update_gather_list = handle_excel.get_table_value(update_gather_path)
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

class TestRlz():                
   
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('热力站添加')
    def test_add(self,case):
        # pytest.skip()
        global rlz_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        test_result_col = 36 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = rlz_manage.add(case)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('热力站编辑')
    def test_update(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global rlz_manage
        global update_row
        global logger
        update_row += 1  
        n = update_row   
        test_result_col = 17 #测试结果所在的列
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
            result = rlz_manage.update(case)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('热力站删除')
    def test_delete(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global rlz_manage
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
            result = rlz_manage.delete(case)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result 
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_ctl_list)
    @allure.feature('热力站-控制柜添加')
    def test_control(self,case):
        # pytest.skip()
        global rlz_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        test_result_col = 17 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = rlz_manage.add_control_box(case)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_gather_list)
    @allure.feature('热力站-采集量')
    def test_add_gather(self,case):
        
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
            result = rlz_manage.add_gather_param(case)
            if result:
                handle_excel.write_value(add_gather_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_gather_path,n,test_result_col,'失败')
            assert result

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

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_control_list)
    @allure.feature('热力站-控制量')
    def test_add_control(self,case):
        pytest.skip()
        global rlz_manage
        global add_row
        global logger
        add_row += 1  
        n = add_row   
        case_id = case[0]
        is_execute = case[2]
        #测试结果所在的列
        test_result_col = 25 
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
            result = rlz_manage.add_control_param(case)
            if result:
                handle_excel.write_value(add_control_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_control_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',copy_list)
    @allure.feature('热力站-复制')
    def test_copy(self,case):
        
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
            result = rlz_manage.copy_to_newrlz(case)
            if result:
                handle_excel.write_value(copy_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(copy_path,n,test_result_col,'失败')
            assert result