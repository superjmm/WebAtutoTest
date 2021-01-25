#coding=utf-8
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import GrzdbManage
from Util.handle_json import handle_json
from Util.handle_excel import handle_excel
from Util.handle_precondition import handle_preconditon
from Util.handle_log import handle_log

log_path = base_path+'/Report/log.txt' #日志路径
logger = logging.getLogger('接口测试日志')
#handle_log.set_log(logger,log_level='DEBUG',log_path=log_path) #设置日志等级及路径
handle_log.set_log(logger,log_level='DEBUG') #设置日志等级及屏幕输出
add_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_row = 1
delete_row = 1 
add_row_zdxx = 1 
update_row_zdxx = 1
delete_row_zdxx = 1 
add_path = base_path+'/Case/grzdb-add.xlsx' #添加测试用例
update_path = base_path+'/Case/grzdb-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/grzdb-delete.xlsx' #删除测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 

add_path_zdxx = base_path+'/Case/grzdb-xx-add.xlsx' #添加测试用例
update_path_zdxx = base_path+'/Case/grzdb-xx-update.xlsx' #编辑测试用例
delete_path_zdxx = base_path+'/Case/grzdb-xx-delete.xlsx' #删除测试用例
add_list_zdxx = handle_excel.get_table_value(add_path_zdxx) 
update_list_zdxx = handle_excel.get_table_value(update_path_zdxx) 
delete_list_zdxx = handle_excel.get_table_value(delete_path_zdxx) 

grzdb_manage = ""

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global grzdb_manage
    grzdb_manage = GrzdbManage()

class TestGrzdb():   
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('字典添加')
    def test_add_zdl(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global grzdb_manage        
        global logger
        global add_path
        global add_row
        add_row += 1  
        n = add_row   
        case_id,case_name,is_execute,num,code,name,remark,button,expect_result,test_result= case
        test_result_col = 10 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = grzdb_manage.add(num,code,name,remark,button)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('字典编辑')
    def test_update_zdl(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global grzdb_manage
        global logger
        global update_path
        global update_row
        update_row += 1   
        n = update_row  
        case_id,case_name,is_execute,name_old,num,code,name,remark,button,expect_result,test_result= case
        test_result_col = 10 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = grzdb_manage.update(name_old,num,code,name,remark,button)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('字典删除')
    def test_delete_zdl(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global logger
        global grzdb_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        case_id,case_name,is_execute,name,is_delete,expect_result,test_result= case
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
            result = grzdb_manage.delete(name,is_delete)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',add_list_zdxx)
    @allure.feature('字典详细添加')
    def test_add_zdxx(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global grzdb_manage        
        global logger
        global add_path
        global add_row
        add_row += 1  
        n = add_row   
        case_id,case_name,is_execute,num,dict_name,item_name,code,is_default,is_use,remark,button,expect_result,test_result= case
        test_result_col = 13 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = grzdb_manage.add_detail(num,dict_name,item_name,code,is_default,is_use,remark,button)
            if result:
                handle_excel.write_value(add_path_zdxx,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path_zdxx,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',update_list_zdxx)
    @allure.feature('字典详细编辑')
    def test_update_zdxx(self,case):
        '''
        数据驱动测试(Excel)
        '''
        pytest.skip()
        global grzdb_manage
        global logger
        global update_path
        global update_row
        update_row += 1   
        n = update_row  
        case_id,case_name,is_execute,item_name,num,dict_name,item_name_new,code,is_default,is_use,remark,button,expect_result,test_result= case
        test_result_col = 14 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = grzdb_manage.update_detail(item_name,num,dict_name,item_name_new,code,is_default,is_use,remark,button)
            if result:
                handle_excel.write_value(update_path_zdxx,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path_zdxx,n,test_result_col,'失败')
            assert result        
    
    @pytest.mark.parametrize('case',delete_list_zdxx)
    @allure.feature('字典详细删除')
    def test_delete_zdxx(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global logger
        global grzdb_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        case_id,case_name,is_execute,dict_name,item_name,is_delete,expect_result,test_result= case
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
            result = grzdb_manage.delete_detail(dict_name,item_name,is_delete)
            if result:
                handle_excel.write_value(delete_path_zdxx,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path_zdxx,n,test_result_col,'失败')
            assert result