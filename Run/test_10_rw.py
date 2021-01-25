#coding=utf-8
import sys
import os
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import RwManage
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
add_path = base_path+'/Case/rw-add.xlsx' #添加测试用例
update_path = base_path+'/Case/rw-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/rw-delete.xlsx' #删除测试用例
detail_path = base_path+'/Case/rw-detail.xlsx' #详细测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 
detail_list = handle_excel.get_table_value(detail_path) 
rw_manage = ""

@pytest.fixture(scope='class',autouse=True)
def before():
    """
    初始化类
    """    
    global rw_manage
    rw_manage = RwManage()

class TestRw():   
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('热网添加')
    def test_add_rw(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global rw_manage        
        global logger
        global add_path
        global add_row
        add_row += 1  
        n = add_row   
        case_id,case_name,is_execute,rw_name,addr,company,keyword,area,mark,type,button,expect_result,test_result= case
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
            result = rw_manage.add(rw_name,addr,company,keyword,area,mark,type,button)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result
    
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('热网编辑')
    def test_update_rw(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global rw_manage
        global logger
        global update_path
        global update_row
        update_row += 1   
        n = update_row  
        case_id,case_name,is_execute,rw_name_old,rw_name_new,addr,company,keyword,area,mark,type,button,expect_result,test_result= case
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
            result = rw_manage.update(rw_name_old,rw_name_new,addr,company,keyword,area,mark,type,button)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result        
    
    # @pytest.mark.parametrize('case',detail_list)
    # @allure.feature('热网详细')
    # def test_detail(self,case):    
        # pytest.skip()  
        # detail_path = self.detail_path
        # detail_list = self.detail_list
        # rw_manage = self.rw_manage
        # global n  
        # global logger
        # n += 1     
        # case_id,case_name,is_execute,rw_name,expect_result,test_result= case
        # test_result_col = 6 #测试结果所在的列
        # if case_id is None: #用例编号不能为空
        #     logger.info(f'第{n}行，用例编号为空！')
        #     print(f'第{n}行，用例编号为空！')
        #     pytest.skip('用例编号为空')
        
        # elif is_execute == 'no' : #判断该条用例是否执行
        #     logger.info(case_id+'跳过不执行！')
        #     print(case_id+'跳过不执行！')  
        #     pytest.skip('跳过不执行') 
        # else: # 执行用例
        #     result = rw_manage.detail(rw_name)
        #     if result:
        #         handle_excel.write_value(detail_path,n,test_result_col,'通过')
        #     else:
        #         handle_excel.write_value(detail_path,n,test_result_col,'失败')
        #     assert result

    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('热网删除')
    def test_delete_rw(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global logger
        global rw_manage
        global delete_path
        global delete_row
        delete_row += 1 
        n = delete_row    
        case_id,case_name,is_execute,rw_name,is_delete,expect_result,test_result= case
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
            result = rw_manage.delete(rw_name,is_delete)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result