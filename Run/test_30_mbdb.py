#coding=utf-8
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import MbdbManage
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
add_path = base_path+'/Case/mbdb-add.xlsx' #添加测试用例
update_path = base_path+'/Case/mbdb-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/mbdb-delete.xlsx' #删除测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 

add_row_param = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_row_param = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
delete_row_param = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
add_path_param = base_path+'/Case/mbdb-param-add.xlsx' #添加测试用例
update_path_param = base_path+'/Case/mbdb-param-update.xlsx' 
delete_path_param = base_path+'/Case/mbdb-param-delete.xlsx' 
add_list_param = handle_excel.get_table_value(add_path_param) 
update_list_param = handle_excel.get_table_value(update_path_param) 
delete_list_param = handle_excel.get_table_value(delete_path_param) 

gather_path = base_path+'/Case/mbdb-param-gather.xlsx'
control_path = base_path+'/Case/mbdb-param-control.xlsx'
gather_list = handle_excel.get_table_value(gather_path)
control_list = handle_excel.get_table_value(control_path)

mbdb_manage = ""

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global mbdb_manage
    mbdb_manage = MbdbManage()

class TestMbdb():   
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('模板点表添加')
    def test_add_mbdb(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global mbdb_manage        
        global logger
        global add_path
        global add_row
        add_row += 1  
        n = add_row   
        case_id,case_name,is_execute,mbdb_name,remark,button,expect_result,test_result= case
        test_result_col = 8 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = mbdb_manage.add(mbdb_name,remark,button)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('模板点表编辑')
    def test_update_mbdb(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global mbdb_manage
        global logger
        global update_path
        global update_row
        update_row += 1   
        n = update_row  
        print(case)
        case_id,case_name,is_execute,mbdb_name_old,mbdb_name_new,remark,button,expect_result,test_result= case
        test_result_col = 9 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = mbdb_manage.update(mbdb_name_old,mbdb_name_new,remark,button)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('模板点表删除')
    def test_delete_mbdb(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global logger
        global mbdb_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        case_id,case_name,is_execute,mbdb_name,is_delete,expect_result,test_result= case
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
            result = mbdb_manage.delete(mbdb_name,is_delete)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',add_list_param)
    @allure.feature('模板点表参量添加')
    def test_add_param(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global mbdb_manage        
        global logger
        global add_path_param
        global add_row_param
        add_row_param += 1  
        n = add_row_param  
        test_result_col = 12 #测试结果所在的列
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
            result = mbdb_manage.add_param(case)
            if result:
                handle_excel.write_value(add_path_param,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path_param,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.skip()
    @pytest.mark.parametrize('case',gather_list)
    @allure.feature('模板点表采集量添加')
    def test_add_param_gather(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global mbdb_manage        
        global logger
        global add_path_param
        global add_row_param
        add_row_param += 1  
        n = add_row_param  
        #测试结果所在的列
        test_result_col = 34 
        case_id = case[0]
        is_execute = case[2]
        #用例编号不能为空
        if case_id is None: 
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        #判断该条用例是否执行
        elif is_execute == 'no' : 
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        # 执行用例
        else: 
            result = mbdb_manage.add_param_gather(case)
            if result:
                handle_excel.write_value(gather_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(gather_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',control_list)
    @allure.feature('模板点表控制量添加')
    def test_add_param_control(self,case):
        '''
        数据驱动测试(Excel)
        '''
        
        global mbdb_manage        
        global logger
        global add_path_param
        global add_row_param
        add_row_param += 1  
        n = add_row_param  
        #测试结果所在的列
        test_result_col = 24 
        case_id = case[0]
        is_execute = case[2]
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
            result = mbdb_manage.add_param_control(case)
            if result:
                handle_excel.write_value(control_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(control_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',update_list_param)
    @allure.feature('模板点表参量配置修改')
    def test_update_param(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global mbdb_manage        
        global logger
        global add_path_param
        global add_row_param
        add_row_param += 1  
        n = add_row_param  
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
            result = mbdb_manage.update_param(case)
            if result:
                handle_excel.write_value(update_path_param,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path_param,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',delete_list_param)
    @allure.feature('模板点表删除')
    def test_delete_param(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global logger
        global mbdb_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        test_result_col = 7 #测试结果所在的列
        case_id = case[0]
        is_execute = case[2]
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = mbdb_manage.delete_param(case)
            if result:
                handle_excel.write_value(delete_path_param,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path_param,n,test_result_col,'失败')
            assert result    