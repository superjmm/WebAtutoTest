#coding=utf-8
import sys
import os

import py
base_path = os.getcwd()
sys.path.append(base_path)
import json
import pytest
import allure
import logging
from Operate.backstage_manage import RyManage
from Util.handle_json import handle_json
from Util.handle_excel import handle_excel
from Util.handle_precondition import handle_preconditon
from Util.handle_log import handle_log

log_path = base_path+'/Report/log.txt' #日志路径
logger = logging.getLogger('接口测试日志')
#handle_log.set_log(logger,log_level='DEBUG',log_path=log_path) #设置日志等级及路径
handle_log.set_log(logger,log_level='DEBUG') #设置日志等级及屏幕输出

add_path = base_path+'/Case/ry-add.xlsx' #添加测试用例
update_path = base_path+'/Case/ry-update.xlsx' #编辑测试用例
delete_path = base_path+'/Case/ry-delete.xlsx' #删除测试用例
#detail_path = base_path+'/Case/ry-detail.xlsx' #详细测试用例
add_list = handle_excel.get_table_value(add_path) 
update_list = handle_excel.get_table_value(update_path) 
delete_list = handle_excel.get_table_value(delete_path) 
#detail_list = handle_excel.get_table_value(detail_path) 
add_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
update_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
delete_row = 1 # 表示是第二行，第一行是表头，从第二行开始才是用例部分
ry_manage = ''

@pytest.fixture(scope='class',autouse=True)
def before():
    global ry_manage
    ry_manage = RyManage()


class TestRy():  
            
    @pytest.mark.parametrize('case',add_list)
    @allure.feature('热源添加')
    def test_add_ry(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global ry_manage        
        global logger
        global add_path
        global add_row
        add_row += 1   
        n = add_row  
        case_id,case_name,is_execute,ry_name,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button,expect_result,test_result= case
        test_result_col = 22 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = ry_manage.add(ry_name,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button)
            if result:
                handle_excel.write_value(add_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(add_path,n,test_result_col,'失败')
            assert result

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',update_list)
    @allure.feature('热源编辑')
    def test_update_ry(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global ry_manage        
        global logger     
        global update_path  
        global update_row 
        update_row += 1  
        n = update_row   
        case_id,case_name,is_execute,ry_name_old,ry_name_new,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button,expect_result,test_result= case
        test_result_col = 23 #测试结果所在的列
        if case_id is None: #用例编号不能为空
            logger.info(f'第{n}行，用例编号为空！')
            print(f'第{n}行，用例编号为空！')
            pytest.skip('用例编号为空')
        
        elif is_execute == 'no' : #判断该条用例是否执行
            logger.info(case_id+'跳过不执行！')
            print(case_id+'跳过不执行！')  
            pytest.skip('跳过不执行') 
        else: # 执行用例
            result = ry_manage.update(ry_name_old,ry_name_new,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button)
            if result:
                handle_excel.write_value(update_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(update_path,n,test_result_col,'失败')
            assert result       

    @pytest.mark.skip()
    @pytest.mark.parametrize('case',delete_list)
    @allure.feature('热源删除')
    def test_delete_ry(self,case):
        '''
        数据驱动测试(Excel)
        '''
        # pytest.skip()
        global logger
        global ry_manage      
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
            result = ry_manage.delete(name,is_delete)
            if result:
                handle_excel.write_value(delete_path,n,test_result_col,'通过')
            else:
                handle_excel.write_value(delete_path,n,test_result_col,'失败')
            assert result
