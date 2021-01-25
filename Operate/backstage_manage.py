#coding=utf-8
import sys
import os
from selenium.webdriver.common import keys
from six import b
base_path = os.getcwd() #获取工程路径
sys.path.append(base_path) #添加进入环境变量
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from Util.handle_ini import handle_ini
from Util.handle_excel import handle_excel
import time
import json

url_ini_path = base_path + '/Config/url.ini' #URL配置文件地址
browser = webdriver.Firefox() #初始化火狐浏览器
class Login:
    def __init__(self):
        url = handle_ini.get_value(url_ini_path,'server','host')
        username = handle_ini.get_value(url_ini_path,'server','username')
        password = handle_ini.get_value(url_ini_path,'server','password')
        browser.get(url)
        self.login(username,password)
    def login(self,username,password):
        flag = False
        try:
            browser.find_element_by_xpath("//input[@placeholder='请输入你的登录账号']").send_keys(username)
            browser.find_element_by_xpath("//input[@placeholder='请输入你的登录密码']").send_keys(password)
            browser.find_element_by_xpath("//span[text()='登录']").click()
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='content']/div[1]/a/div/div").click()
            flag = True
        except:
            raise Exception("登录失败！")
login = Login()
class RwManage: #热网
    def __init__(self):
        #热网管理模块地址
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','rw')
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/a/li/span[text()='热网管理']").click()
        except:
            raise Exception('点击热网管理模块失败！')
        assert "热网管理" in browser.title      
    
    def add(self,rw_name,addr,company,keyword,area,mark,type,button):#添加页面操作
        flag = False      
        try:
            time.sleep(2)
            browser.find_element_by_class_name('el-button--success').click() 
            time.sleep(2)
            if rw_name is not None: #热网名称
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").send_keys(rw_name)
            if addr is not None: #地址
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(addr) 
            if company is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/div/input").click()#所属公司点击出现下拉框
                if keyword is not None:
                    browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword) #所属公司-下拉框-检索关键字
                browser.find_element_by_xpath("//span[text()='"+company+"']").click() #所属公司-下拉框-点击选择下拉项
            if area is not None: #供热面积
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div[1]/div/input").send_keys(area) 
            if '虚拟' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/label[1]/span[1]/span").click() 
            if '真实' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/label[2]/span[1]/span").click() 
            if '环网' == type: #热网类型
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[1]/span[1]/span").click() 
            if '分网' == type: #热网类型
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[2]/span[1]/span").click()
            if 'yes' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[text()='"+rw_name+"']")
                    flag = True
                except:
                    raise Exception(f'保存热网失败！')
            if 'clear' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[1]").click()
                if browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").text == "":
                    flag = True
            if 'no' == button: 
                browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
                try:
                    browser.find_element_by_xpath("//div[@class='dialog-footer']/button[2]")
                except:
                    flag = True
        except:
            raise Exception("添加操作失败！")
        return flag
    def update(self,old,new,addr,company,keyword,area,mark,type,button):#编辑页面操作
        flag = False      
        try:
            time.sleep(2)
            try:
                browser.find_element_by_xpath("//div[text()='"+old+"']/parent::*/parent::*/following-sibling::td[8]/div/button").click()
            except:
                raise Exception("该热网不存在或者点击编辑失败！")
            time.sleep(2)
            if old is not None: 
                try:                    
                    if new is not None:#名称
                        elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div[1]/input")
                        elem.send_keys(" ")#清空前随便输入一点内容，否则清空后再添加内容，文字添加不全（不知道为啥？）
                        elem.clear()
                        elem.send_keys(new)
                    if addr is not None: #地址
                        elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(addr)                     
                    if company is not None:#公司
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/div/input").click()#所属公司点击出现下拉框
                        if keyword is not None:#搜索关键字
                            browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword) #所属公司-下拉框-检索关键字
                        browser.find_element_by_xpath("//span[text()='"+company+"']").click() #所属公司-下拉框-点击选择下拉项
                    if area is not None: #供热面积
                        elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div[1]/div/input")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(area) 
                    if '虚拟' == mark: #热网标示
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/label[1]/span[1]/span").click() 
                    if '真实' == mark: #热网标示
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/label[2]/span[1]/span").click() 
                    if '环网' == type: #热网类型
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[1]/span[1]/span").click() 
                    if '分网' == type: #热网类型
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[2]/span[1]/span").click()
                    if 'yes' == button: 
                        browser.find_element_by_xpath("//span[@class='dialog-footer']/button[2]").click()
                        try:
                            browser.find_element_by_xpath("//div[text()='"+new+"']")
                            flag = True
                        except:
                            raise Exception(f'保存编辑热网失败！')
                    if 'clear' == button:                        
                        browser.find_element_by_xpath("//span[@class='dialog-footer']/button[1]").click()
                        if browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").text == "":
                            flag = True
                            browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
                    if 'no' == button: 
                        browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
                        try:
                            browser.find_element_by_xpath("//div[@class='dialog-footer']/button[2]")
                        except:
                            flag = True
                except:
                    raise Exception("该热网不存在！")
        except:
            raise Exception("编辑操作失败！")
        return flag
    # def detail(self,rw_name):#详细
    #     flag = False
    #     if rw_name is not None:
    #         try:
    #             browser.find_element_by_xpath("//div[text()='"+rw_name+"']/parent::*/parent::*/following-sibling::td[7]/div/button").click()
                
    #             if True:
    #                 flag = True
    #                 browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
    #         except:
    #             raise Exception("该热网不存在，不能查看详细！")
    #     return flag
    def delete(self,rw_name,is_delete):#删除
        flag = False
        if rw_name is not None:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//div[text()='"+rw_name+"']/parent::*/parent::*/following-sibling::td[9]/div/button").click()
                if is_delete == "no":
                    browser.find_element_by_xpath("//span[contains(text(),'取消')]").click()
                    try:
                        browser.find_element_by_xpath("//div[text()='"+rw_name+"']")
                        flag= True
                    except:
                        raise Exception("列表中不存在该热网！")
                if is_delete == "yes":
                    browser.find_element_by_xpath("//span[contains(text(),'确定')]").click()
                    try:
                        browser.find_element_by_xpath("//div[text()='"+rw_name+"']")
                    except:
                        flag = True
                
            except:
                raise Exception('该热网不存在，不能删除！')
        return flag

class RyManage: #热源 
    def __init__(self):
        #热源管理模块地址
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','ry')
        # dict_url=handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','grzdb')
        # browser.get(dict_url)
        # time.sleep(2)
        # browser.get(url)
        try:
            time.sleep(3)
            browser.find_element_by_xpath("//span[text()='热源管理']").click()
        except:
            raise Exception('点击热源模块失败！')
        assert "热源管理" in browser.title      
    
    def add(self,ry_name,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button):#添加页面操作
        flag = False      
        try:
            time.sleep(2)
            browser.find_element_by_class_name('el-button--success').click() 
            time.sleep(2)
            if ry_name is not None: #名称
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").send_keys(ry_name)
            if addr is not None: #地址
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(addr) 
            if area is not None:#供热面积
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/div/input").send_keys(area) 
            if company is not None:
                try:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div/div/input").click()#所属公司点击出现下拉框
                    if keyword is not None:
                        browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword) #所属公司-下拉框-检索关键字
                    browser.find_element_by_xpath("//span[text()='"+company+"']").click() #所属公司-下拉框-点击选择下拉项
                except:
                    raise Exception('所属公司不在下拉列表内，请重新选择！')
            if rw is not None: #所属热网
                try:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/div/input").click()
                    browser.find_element_by_xpath("//span[text()='"+rw+"']").click()
                except:
                    raise Exception('所属热网不在下拉列表内，请重新选择！')
            if output is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/div/input").send_keys(output)
            if dia is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/div/input").send_keys(dia)
            if heat_top is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/div/input").send_keys(heat_top)
            if flow_top is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/div/input").send_keys(flow_top)
            if give_tem_top is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/div/input").send_keys(give_tem_top)
            if back_tem_top is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/div/input").send_keys(back_tem_top)
            if dim is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[12]/div/div[1]/input").send_keys(str(dim))
            if lon is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[13]/div/div[1]/input").send_keys(str(lon))
            if '虚拟' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[14]/div/div/label[1]/span").click() 
            if '真实' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[14]/div/div/label[2]/span").click() 
            if type is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[15]/div/div").click()
                browser.find_element_by_xpath("//span[text()='"+type+"']/parent::li")
                # elem = browser.find_element_by_xpath("//span[text()='"+type+"']/parent::li")
                # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            if 'yes' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[text()='"+ry_name+"']/ancestor::div[@class='el-table__body-wrapper is-scrolling-left']")
                    flag = True
                except:
                    raise Exception(f'保存热源失败！')
            if 'clear' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[1]").click()
                if browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").text == "":
                    flag = True
            if 'no' == button: 
                browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
                try:
                    browser.find_element_by_xpath("//div[@class='dialog-footer']/button[2]")
                except:
                    flag = True
        except:
            raise Exception("添加操作失败！")
        return flag
    def update(self,old,ry_name,addr,area,company,keyword,rw,output,dia,heat_top,flow_top,give_tem_top,back_tem_top,dim,lon,mark,type,button):#添加页面操作
        flag = False      
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+old+"']/parent::*/parent::*/following-sibling::td[17]/div/button")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            time.sleep(2)
            if ry_name is not None: #名称
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(ry_name)
            if addr is not None: #地址
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(addr) 
            if area is not None:#供热面积
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(area) 
            if company is not None:
                try:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div/div/input").click()#所属公司点击出现下拉框
                    if keyword is not None:
                        browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword) #所属公司-下拉框-检索关键字
                    browser.find_element_by_xpath("//span[text()='"+company+"']").click() #所属公司-下拉框-点击选择下拉项
                except:
                    raise Exception('所属公司'+company+'不在下拉列表内，请重新选择！')
            if rw is not None: #所属热网
                try:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div/div/input").click()
                    browser.find_element_by_xpath("//span[text()='"+rw+"']").click()
                except:
                    raise Exception('所属热网'+rw+'不在下拉列表内，请重新选择！')
            if output is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(output)
            if dia is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(dia)
            if heat_top is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(heat_top)
            if flow_top is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(flow_top)
            if give_tem_top is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(give_tem_top)
            if back_tem_top is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/div/input").send_keys(back_tem_top)
            if dim is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[12]/div/div[1]/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(dim)
            if lon is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[13]/div/div[1]/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(lon)
            if '虚拟' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[14]/div/div/label[1]/span").click() 
            if '真实' == mark: #热网标示
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[14]/div/div/label[2]/span").click() 
            if type is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[15]/div/div").click()
                browser.find_element_by_xpath("//span[text()='"+type+"']/parent::li").click()
            if 'yes' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[2]").click()
                try:
                    browser.find_element_by_xpath("//div[text()='"+ry_name+"']")
                    flag = True
                except:
                    raise Exception(f'保存热源失败！')
            if 'clear' == button: 
                browser.find_element_by_xpath("//span[@class='dialog-footer']/button[1]").click()
                if browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").text == "":
                    flag = True
            if 'no' == button: 
                browser.find_element_by_xpath("//button[@class='el-dialog__headerbtn']").click()
                try:
                    browser.find_element_by_xpath("//div[@class='dialog-footer']/button[2]")
                except:
                    flag = True
        except:
            raise Exception("添加操作失败！")
        return flag
    def delete(self,ry_name,is_delete):#删除
        flag = False
        if ry_name is not None:
            try:
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+ry_name+"']/parent::*/parent::*/following-sibling::td[18]/div/button")
                browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                time.sleep(2)
                if is_delete == "no":
                    browser.find_element_by_xpath("//span[contains(text(),'取消')]").click()
                    try:
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+ry_name+"']")
                        flag= True
                    except:
                        raise Exception("列表中不存在该热源！")
                if is_delete == "yes":
                    browser.find_element_by_xpath("//span[contains(text(),'确定')]").click()
                    try:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+ry_name+"']")
                    except:
                        flag = True
                
            except:
                raise Exception('该热源不存在，不能删除！')
        return flag

class BzdbManage: #标准点表
    def __init__(self):
        #标准点表管理模块地址
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','bzdb')
        # dict_url=handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','grzdb')
        # browser.get(dict_url)
        # time.sleep(2)
        # browser.find_element_by_xpath("//span[text()='标准点表管理']").click()
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//span[text()='标准点表管理']").click()
        except:
            raise Exception('点击标准点表模块失败！')
        assert "标准点表管理" in browser.title     
    
    def add(self,type,name,lable_name,unit,group,fix_type,kind,wclx,is_caculate,is_view,is_main,order,button):
        flag = False
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='添加']").click()
        except:
            raise Exception('标准点表-点击添加失败！')
        if name is not None: #and lable_name is not None and order is not None
            time.sleep(2)
            try:
                if '采集量' == type:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if '控制量' == type:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(name)
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/input").send_keys(lable_name.strip())
                if unit is not None:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div").click()
                    # elem1 = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div")
                    # browser.execute_script("arguments[0].click();", elem1)
                    browser.find_element_by_xpath("//span[text()='"+unit+"']").click()
                    # elem2 = browser.find_element_by_xpath("//span[text()='"+unit+"']")
                    # browser.execute_script("arguments[0].click();", elem2)
                if group is not None:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div").click()
                    # elem1 = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div")
                    # browser.execute_script("arguments[0].click();", elem1)
                    browser.find_element_by_xpath("//span[text()='"+group+"']").click()
                    # elem2 = browser.find_element_by_xpath("//span[text()='"+group+"']")
                    # browser.execute_script("arguments[0].click();", elem2)
                if fix_type == '报警量':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if fix_type == '时间值':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                if fix_type == '模拟量':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[3]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[3]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'TX':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'DO':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'CO':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[3]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[3]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'AO':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[4]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[4]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'AI':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[5]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[5]")
                    # browser.execute_script("arguments[0].click();", elem)
                if kind == 'DI':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[6]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[6]")
                    # browser.execute_script("arguments[0].click();", elem)
                if wclx == '公用':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if wclx == '一次侧':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                if wclx == '二次侧':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[3]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[3]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_caculate == 'yes':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_caculate == 'no':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_view == 'yes':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_view == 'no':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_main == 'yes':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                if is_main == 'no':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[12]/div/div/input").send_keys(order)
                if button == 'yes':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[2]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                    #验证保存成功
                    try:
                        flag = True
                    except:
                        raise Exception('标准点表-添加页面-保存信息失败！')
                if button == 'clear':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[2]")
                    # browser.execute_script("arguments[0].click();", elem)
                    #验证重置成功
                    try:
                        flag = True
                    except:
                        raise Exception('标准点表-添加页面-重置操作失败！')

                if button == 'no':
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button[1]").click()
                    # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button[1]")
                    # browser.execute_script("arguments[0].click();", elem)
                    #验证取消成功
                    try:
                        browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button[1]").click()
                        # elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button[1]")
                        # browser.execute_script("arguments[0].click();", elem)
                    except:
                        flag = True
            except:
                raise Exception('标准点表-添加操作失败！')
        return flag    
    def update(self,case):
        flag = False
        name_old = case[2]
        parm_type = case[3]
        name = None
        lable_name = None
        unit = case[6]
        group = None
        fix_type = None
        kind = case[9]        
        wclx = None
        is_caculate = None
        is_view = None
        is_main = None
        order = None
        button = case[15]
        if '控制量' == parm_type:
            try:
                browser.find_element_by_xpath("//span[text()='控制量']/parent::button").click()
            except:
                raise Exception("点击【控制量】失败！")
        try:
            time.sleep(3)
            elem = browser.find_element_by_xpath("//input[@placeholder='请输入参量名称']")
            elem.clear()
            elem.send_keys(name_old)
            elem.send_keys(Keys.ENTER)
        except:
            raise Exception('搜索参量名称失败！')
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+name_old+"']/parent::*/parent::*/following-sibling::td[10]/div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击编辑失败！')
        try:
            if '采集量' == parm_type:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[1]").click()
            if '控制量' == parm_type:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/label[2]").click()
            if name is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").clear()
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(name)
            if lable_name is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/input").clear()
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[3]/div/div/input").send_keys(lable_name)
            if unit is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[4]/div/div").click()
                browser.find_element_by_xpath("//span[text()='"+unit+"']").click()
            if group is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div").click()
                browser.find_element_by_xpath("//span[text()='"+group+"']").click()

            browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div").click()
            browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[5]/div/div").click()
            if fix_type == '报警量':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[1]").click()
            if fix_type == '时间值':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[2]").click()
            if fix_type == '模拟量':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[6]/div/div/label[3]").click()
            
            if kind == 'TX':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[1]")
                browser.execute_script("arguments[0].click();",elem)
            if kind == 'DO':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[2]/")
                browser.execute_script("arguments[0].click();",elem)
            if kind == 'CO':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[3]")
                browser.execute_script("arguments[0].click();",elem)
            if kind == 'AO':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[4]")
                browser.execute_script("arguments[0].click();",elem)
            if kind == 'AI':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[5]")
                browser.execute_script("arguments[0].click();",elem)
            if kind == 'DI':
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[7]/div/div/label[6]")
                browser.execute_script("arguments[0].click();",elem)
              

            if wclx == '公用':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[1]").click()
            if wclx == '一次侧':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[2]").click()
            if wclx == '二次侧':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[8]/div/div/label[3]").click()
            if is_caculate == 'yes':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[1]").click()
            if is_caculate == 'no':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[9]/div/div/label[2]").click()
            if is_view == 'yes':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[1]").click()
            if is_view == 'no':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[10]/div/div/label[2]").click()
            if is_main == 'yes':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[1]").click()
            if is_main == 'no':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[11]/div/div/label[2]").click()
            if order is not None:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[12]/div/div/input").send_keys(order)
            if button == 'yes':
                time.sleep(3)
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[2]").click()
                except:
                    flag = True
                
                # try:
                #     browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+name+"']")
                #     
                # except:
                #     raise Exception('标准点表被误删！')
            if button == 'clear':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/following-sibling::div[1]/span/button[1]").click()
                flag = True
            if button == 'no':
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button").click()
                time.sleep(2)
                try:
                   browser.find_element_by_xpath("//form[@class='el-form ruleForm']/parent::div/preceding-sibling::div[1]/button")
                except:
                    flag = True
        except:
            raise Exception('标准点表-编辑操作失败！')
        return flag
    def delete(self,name,is_delete):
        flag = False
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+name+"']/parent::*/parent::*/following-sibling::td[10]/div/button[2]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            # elem.click()
            
            if 'yes' == is_delete:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='el-message-box__btns']/button[2]").click()
                time.sleep(2)
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+name+"']")
                except:
                    flag = True              
            if 'no' == is_delete:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='el-message-box__btns']/button[1]").click()
                time.sleep(2)
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+name+"']")
                    flag = True
                except:
                    raise Exception('标准点表被误删！')
        except:
            raise Exception('标准点表列表-删除失败！')
        return flag

class DwManage: #单位
    def __init__(self):
    
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','dw')
        # browser.get(url)
        try:
            # browser.find_element_by_xpath("//span[text()='单位管理']").click()
            elem = browser.find_element_by_xpath("//span[text()='单位管理']")
            browser.execute_script("arguments[0].click();",elem)
        except:
            raise Exception('点击单位管理模块失败！')
        # assert "单位管理" in browser.title      
    
    def add(self,num,dw_name,dw,button):#添加页面操作
        flag = False      
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='添加单位']").click()
            time.sleep(2)
            browser.find_element_by_xpath("//input[@placeholder='请输入序号']").send_keys(num)
            browser.find_element_by_xpath("//input[@placeholder='请输入单位类型']").send_keys(dw_name)
            browser.find_element_by_xpath("//input[@placeholder='请输入单位']").send_keys(dw)
            if 'yes' == button: 
                browser.find_element_by_xpath("//span[text()='确认']").click()
                flag = True
                # try:
                #     browser.find_element_by_xpath("//div[text()='"+dw_name+"']/ancestor::div[@class='el-table__body-wrapper is-scrolling-none']")
                #     flag = True
                # except:
                #     raise Exception('添加单位失败！')
            if 'no' == button: 
                browser.find_element_by_xpath("//span[text()='取消']").click()
                flag = True
                # try:
                #     time.sleep(2)
                #     browser.find_element_by_xpath("//span[text()='取消']").click()
                # except:
                #     flag = True
            if 'close' == button: 
                browser.find_element_by_xpath("//i[@class='el-dialog__close el-icon el-icon-close']").click()
                flag = True
                # try:
                #     time.sleep(2)
                #     browser.find_element_by_xpath("//span[text()='取消']").click()
                # except:
                #     flag = True
        except:
            raise Exception("单位-添加操作失败！")
        return flag
    def update(self,dw_name_old,num,dw_name_new,dw,button):#编辑页面操作
        flag = False      
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+dw_name_old+"']/parent::*/parent::*/following-sibling::td[3]/div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            time.sleep(2)
            if dw_name_old is not None: 
                try:                    
                    if num is not None:#序号
                        elem = browser.find_element_by_xpath("//input[@placeholder='请输入序号']")
                        elem.send_keys(" ")#清空前随便输入一点内容，否则清空后再添加内容，文字添加不全（不知道为啥？）
                        elem.clear()
                        elem.send_keys(num)
                    if dw_name_new is not None: #中文名称
                        elem = browser.find_element_by_xpath("//input[@placeholder='请输入中文名称']")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(dw_name_new)                     
                   
                    if dw is not None: #单位符号
                        elem = browser.find_element_by_xpath("//input[@placeholder='请输入单位']")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(dw) 
                   
                    if 'yes' == button: 
                        browser.find_element_by_xpath("//span[text()='确认']").click()
                        try:
                            browser.find_element_by_xpath("//div[text()='"+dw_name_new+"']/ancestor::div[@class='el-table__body-wrapper is-scrolling-none']")
                            flag = True
                        except:
                            raise Exception('没有编辑单位成功！')
                    if 'no' == button: 
                        browser.find_element_by_xpath("//span[text()='取消']").click()
                        try:
                            time.sleep(2)
                            browser.find_element_by_xpath("//span[text()='取消']").click()
                        except:
                            flag = True
                    if 'close' == button: 
                        browser.find_element_by_xpath("//i[@class='el-dialog__close el-icon el-icon-close']").click()
                        try:
                            time.sleep(2)
                            browser.find_element_by_xpath("//span[text()='取消']").click()
                        except:
                            flag = True
                                
                except:
                    raise Exception("单位-编辑操作失败！")
        except:
            raise Exception("该单位不存在，不能编辑！")
        return flag
    
    def delete(self,dw_name,is_delete):#删除
        flag = False
        if dw_name is not None:
            try:
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+dw_name+"']/parent::*/parent::*/following-sibling::td[3]/div/button[2]")
                browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                if is_delete == "no":
                    browser.find_element_by_xpath("//button[@class='el-button el-button--default el-button--small']/span[contains(text(),'取消')]").click()
                    try:
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+dw_name+"']")
                        flag= True
                    except:
                        raise Exception("单位被误删了！")
                if is_delete == "yes":
                    browser.find_element_by_xpath("//span[contains(text(),'确定')]").click()
                    try:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+dw_name+"']")
                    except:
                        flag = True
                
            except:
                raise Exception('该单位不存在，不能删除！')
        return flag  

class ClflManage: #参量分类
    def __init__(self):
       
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','clfl')
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/a/li/span[text()='参量分类管理']").click()
        except:
            raise Exception('点击参量分类管理模块失败！')
        assert "参量分类管理" in browser.title      
    
    def add(self,clfl_num,clfl_name,button):#添加页面操作
        flag = False      
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//button/span[text()='添加参量分类']").click()
            time.sleep(2)
            browser.find_element_by_xpath("//input[@placeholder='请输入序号']").send_keys(clfl_num)
            browser.find_element_by_xpath("//input[@placeholder='请输入参量分类名称']").send_keys(clfl_name)
            
            if 'yes' == button: 
                browser.find_element_by_xpath("//span[text()='确认']").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name+"']")
                    flag = True
                except:
                    raise Exception('没有添加参量分类成功！')
            if 'no' == button: 
                browser.find_element_by_xpath("//span[text()='取消']").click()
                try:
                    browser.find_element_by_xpath("//span[text()='取消']")
                except:
                    flag = True
            if 'close' == button: 
                browser.find_element_by_xpath("//i[@class='el-dialog__close el-icon el-icon-close']").click()
                try:
                    browser.find_element_by_xpath("//span[text()='取消']")
                except:
                    flag = True
        except:
            raise Exception("参量分类-添加操作失败！")
        return flag
    def update(self,clfl_name_old,clfl_num,clfl_name_new,button):#编辑页面操作
        flag = False      
        try:
            time.sleep(2)
            try:
                elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name_old+"']/parent::*/parent::*/following-sibling::td[2]/div/button[1]")
                browser.execute_script("arguments[0].click();", elem)
            except:
                raise Exception("该参量不存在或者点击编辑失败！")
            time.sleep(2)
            if clfl_name_old is not None: 
                try:                    
                    time.sleep(2)
                    if clfl_num is not None:
                        elem = browser.find_element_by_xpath("//input[@placeholder='请输入序号']")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(clfl_num)
                    if clfl_name_new is not None:
                        elem = browser.find_element_by_xpath("//input[@placeholder='请输入参量分类名称']")
                        elem.send_keys(" ")
                        elem.clear()
                        elem.send_keys(clfl_name_new)
                    if 'yes' == button: 
                        browser.find_element_by_xpath("//span[text()='确认']").click()
                        try:
                            time.sleep(2)
                            if clfl_name_new is None:
                                browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name_old+"']")
                            else:
                                browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name_new+"']")
                            flag = True
                        except:
                            raise Exception('没有添加参量分类成功！')
                    if 'no' == button: 
                        browser.find_element_by_xpath("//span[text()='取消']").click()
                        try:
                            time.sleep(2)
                            browser.find_element_by_xpath("//span[text()='取消']")
                        except:
                            flag = True
                    if 'close' == button: 
                        browser.find_element_by_xpath("//i[@class='el-dialog__close el-icon el-icon-close']").click()
                        try:
                            time.sleep(2)
                            browser.find_element_by_xpath("//span[text()='取消']")
                        except:
                            flag = True
                                
                except:
                    raise Exception("该参量不存在，不能编辑！")
        except:
            raise Exception("参量-编辑操作失败！")
        return flag
    
    def delete(self,clfl_name,is_delete):#删除
        flag = False
        if clfl_name is not None:
            try:
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name+"']/parent::*/parent::*/following-sibling::td[2]/div/button[2]")
                browser.execute_script("arguments[0].click();", elem)
                if is_delete == "no":
                    browser.find_element_by_xpath("//span[contains(text(),'取消')]").click()
                    try:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name+"']")
                        flag= True
                    except:
                        raise Exception("参量被误删！")
                if is_delete == "yes":
                    browser.find_element_by_xpath("//span[contains(text(),'确定')]").click()
                    try:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+clfl_name+"']")
                    except:
                        flag = True
                
            except:
                raise Exception('该参量不存在，不能删除！')
        return flag     

class GrzdbManage: #供热字典
    def __init__(self):
       
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','grzdb')
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/a/li/span[text()='供热字典表']").click()
        except:
            raise Exception('点击供热字典表模块失败！')
        assert "供热字典表" in browser.title      
    
    def add(self,num,code,name,remark,button):#添加页面操作
        flag = False
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='字典类']/parent::*/following-sibling::button").click()
            # browser.find_element_by_xpath("//div[@class='btn']/button/i[@class='el-icon-circle-plus-outline']").click()
            flag = True
        except:
            raise Exception('点击添加字典类失败！')
        try:
            if num is not None:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[1]/div/form/div[1]/div/div/input").send_keys(num)
            if code is not None:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[1]/div/form/div[2]/div/div/input").send_keys(code)
            if name is not None:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[1]/div/form/div[3]/div/div/input").send_keys(name)
            if remark is not None:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[1]/div/form/div[4]/div/div/input").send_keys(remark)
            if 'yes' == button:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                try:
                    browser.find_element_by_xpath("").click()
                except:
                    flag = True
            if 'no' == button:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/following-sibling::button").click()
                except:
                    flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/following-sibling::button").click()
                try:
                    browser.find_element_by_xpath("//span[text()='添加字典']/parent::*/following-sibling::button").click()
                except:
                    flag = True
        except:
            raise Exception('字典类-添加操作失败！')
        return flag

    def update(self,name_old,num,code,name,remark,button):#编辑页面操作
        flag = True
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//span[text()='"+name_old+"']/parent::*/following-sibling::div/button[1]").click()
            elem = browser.find_element_by_xpath("//span[text()='"+name_old+"']/parent::*/following-sibling::div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击修改字典类失败！')
        try:
            if num is not None:
                elem = browser.find_element_by_xpath("//input[@placeholder='请输入序号']")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(num)
            if code is not None:
                elem = browser.find_element_by_xpath("//input[@placeholder='请输入字典编码']")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(code)
            if name is not None:
                elem = browser.find_element_by_xpath("//input[@placeholder='请输入字典名']")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(name)
            if remark is not None:
                elem = browser.find_element_by_xpath("//input[@placeholder='请输入备注']")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(remark)
            if 'yes' == button:
                browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                try:
                    browser.find_element_by_xpath("").click()
                except:
                    flag = True
            if 'no' == button:
                browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                except:
                    flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                try:
                    browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                except:
                    flag = True
        except:
            raise Exception('字典类-编辑操作失败！')
        return flag
    
    def delete(self,name,is_delete):
        flag = False
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='"+name+"']/parent::*/following-sibling::div/button[2]").click()
        except:
            raise Exception('点击删除字典类失败！')
        try:
            if is_delete == "no":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[2]/div/div[3]/span/button[1]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='"+name+"']")
                    flag= True
                except:
                    raise Exception("字典类被误删！")
            if is_delete == "yes":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[2]/div/div[3]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='"+name+"']")
                except:
                    flag = True
        except:
            raise Exception('删除字典类操作失败！')
        return flag

    def add_detail(self,num,dict_name,item_name,code,is_default,is_use,remark,button):
        flag = False
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//span[text()='"+dict_name+"']").click()
            elem = browser.find_element_by_xpath("//span[text()='"+dict_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击字典类失败！')
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[contains(text(),'字典详细')]/parent::*/following-sibling::div[4]").click()
        except:
            raise Exception('点击添加字典详细失败！')
        try:
            time.sleep(2)
            if num is not None:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[1]/div/div/input").send_keys(num)
            if item_name is not None:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[2]/div/div/input").send_keys(item_name)
            if code is not None:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[3]/div/div/input").send_keys(code)
            if is_default is not None and '是' == is_default:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[4]/div/div/label[1]").click()
            if is_default is not None and '否' == is_default:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[4]/div/div/label[2]").click()
            if is_use is not None and '是' == is_use:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[5]/div/div/label[1]").click()
            if is_use is not None and '否' == is_use:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[5]/div/div/label[2]").click()
            if remark is not None:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[1]/div/form/div[6]/div/div/input").send_keys(remark)
            if'yes' == button:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                    flag =  True
                except:
                    raise Exception('字典详细列表中不存在新记录！')
            if'no' == button:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                except:
                    flag = True
            if'close' == button:
                browser.find_element_by_xpath("//span[text()='"+dict_name+"-添加']/parent::*/parent::*/button").click()
                try:
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                except:
                    flag = True
        except:
            raise Exception('添加字典详细操作失败！')
        return flag
    
    def update_detail(self,item_name,num,dict_name,item_name_new,code,is_default,is_use,remark,button):
        flag = False
        time.sleep(2)
        try:
            elem = browser.find_element_by_xpath("//span[text()='"+dict_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击字典类失败！')
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']/parent::*/parent::*/following-sibling::td[6]/div/button[1]").click()
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']/parent::*/parent::*/following-sibling::td[6]/div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            flag = True
        except:
            raise Exception('点击字典详细编辑失败！')
        try:
            time.sleep(2)
            if num is not None:
                elem = browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[1]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(num)
            if item_name_new is not None:
                elem = browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[2]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(item_name_new)
            if code is not None:
                elem = browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[3]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(code)
            if is_default is not None and '是' == is_default:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[4]/div/div/label[1]").click()
            if is_default is not None and '否' == is_default:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[4]/div/div/label[2]").click()
            if is_use is not None and '是' == is_use:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[5]/div/div/label[1]").click()
            if is_use is not None and '否' == is_use:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[5]/div/div/label[2]").click()
            if remark is not None:
                elem = browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[1]/div/form/div[6]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(remark)
            if'yes' == button:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                    flag =  True
                except:
                    raise Exception('字典详细列表中不存在新记录！')
            if'no' == button:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                except:
                    flag = True
            if'close' == button:
                browser.find_element_by_xpath("//span[text()='"+item_name+"-编辑']/parent::*/parent::*/button").click()
                try:
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                except:
                    flag = True
        except:
            raise Exception('编辑字典详细操作失败！')
        
        return flag

    def delete_detail(self,dict_name,item_name,is_delete):
        flag = True
        time.sleep(2)
        try:
            elem = browser.find_element_by_xpath("//span[text()='"+dict_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击字典类失败！')
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']/parent::*/parent::*/following-sibling::td[6]/div/button[1]").click()
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']/parent::*/parent::*/following-sibling::td[6]/div/button[2]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            flag = True
        except:
            raise Exception('点击字典详细删除失败！')
        try:
            if is_delete == "no":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[3]/span/button[1]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                    flag =  True
                except:
                    raise Exception("字典详细被误删！")
            if is_delete == "yes":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[3]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-left']/descendant::div[text()='"+item_name+"']")
                except:
                    flag = True
        except:
            raise Exception('删除字典详细操作失败！')
        return flag

class MbdbManage: #模板点表
    def __init__(self):
       
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','mbdb')
        # browser.get(url)
        try:
            time.sleep(3)
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/a/li/span[text()='模板点表管理']").click()
        except:
            raise Exception('点击模板点表管理失败！')
        assert "模板点表管理" in browser.title 
    
    def add(self,name,remark,button):
        flag = False
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='模板列表']/parent::*/following-sibling::button").click()
            flag = True
        except:
            raise Exception('点击添加模板失败！')
        try:
            if name is not None:
                browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/parent::*/following-sibling::div[1]/div/form/div[1]/div/div/input").send_keys(name)
            if remark is not None:
                browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/parent::*/following-sibling::div[1]/div/form/div[2]/div/div/input").send_keys(remark)
            if 'yes' == button:
                browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                try:
                    browser.find_element_by_xpath("//span[text()='"+name+"']")
                    flag = True
                except:
                    raise Exception('模板-列表不存在该模板！')
            if 'no' == button:
                browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/following-sibling::button").click()
                except:
                    flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/following-sibling::button").click()
                try:
                    browser.find_element_by_xpath("//span[text()='添加模板']/parent::*/following-sibling::button").click()
                except:
                    flag = True
        except:
            raise Exception('模板-添加操作失败！')
        return flag
    
    def update(self,name_old,name_new,remark,button):
        flag = False
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//span[text()='"+name_old+"']/parent::*/following-sibling::div/button[1]").click()
            elem = browser.find_element_by_xpath("//span[text()='"+name_old+"']/parent::*/following-sibling::div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击编辑模板失败！')
        try:
            if name_new is not None:
                elem = browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[1]/div/form/div[1]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(name_new)
            if remark is not None:
                elem = browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[1]/div/form/div[2]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(remark)
            if 'yes' == button:
                elem = browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[2]/span/button[2]").click()
                time.sleep(2)
                try:
                    if name_new is None:
                        browser.find_element_by_xpath("//span[text()='"+name_old+"']")
                        flag = True
                    else:
                        browser.find_element_by_xpath("//span[text()='"+name_new+"']")
                        flag = True
                except:
                    raise Exception('模板-列表不存在该模板！')
            if 'no' == button:
                browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/parent::*/following-sibling::div[2]/span/button[1]").click()
                try:
                    browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                except:
                    flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                try:
                    browser.find_element_by_xpath("//span[contains(text(),'编辑')]/parent::*/following-sibling::button").click()
                except:
                    flag = True
        except:
            raise Exception('模板点表-编辑操作失败！')
        return flag  
    
    def delete(self,name,is_delete):
        flag = False
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//span[text()='"+name+"']/parent::*/following-sibling::div/button[2]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            flag = True
        except:
            raise Exception('点击删除模板失败！')
        try:
            if is_delete == "no":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[2]/div/div[3]/span/button[1]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='"+name+"']")
                    flag= True
                except:
                    raise Exception("模板被误删！")
            if is_delete == "yes":
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[2]/div/div[3]/span/button[2]").click()
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='"+name+"']")
                except:
                    flag = True
        except:
            raise Exception('删除模板操作失败！')
        
        return flag           

    def add_param(self,data):
        flag_num = False
        flag_button = False
        flag = False
        mb_name = data[3]
        param_type = data[4]
        wc_type = data[5]
        keyword = data[6]      
        param_list = data[7]
        is_setup = data[8]
        button = data[9]

        #1.先选择模板点表
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//span[text()='"+mb_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击模板失败！')
            
        #2.选择采集量还是控制量
        try:
            time.sleep(2)
            if '采集量' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[1]").click()
            if 'control' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[2]").click()   
        except:
            raise Exception('模板点表-采集量/控制量-选择失败！')
            
        #3.点击添加
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='btn btn2']").click()
        except:
            raise Exception('模板-采集量/控制量-点击添加失败！')

        #4.输入添加内容
        if '采集量' == param_type:
            ##4.1 网侧类型选择
            try:
                time.sleep(2)
                if '公用' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[1]").click()
                if '一次侧' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[2]").click()
                if '二次侧' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[3]").click()
            except:
                raise Exception('采集量-添加-标准参量-点击失败！')
            ##4.2 搜索参量关键字
            if  keyword is not None:
                try:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[3]/input").send_keys(keyword)
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/button").click()
                except:
                    raise Exception('查询失败！')
            ##4.3 选择参量并点击
            time.sleep(2)
            old = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[5]/span").text
            
            for p in param_list:
                try:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[4]/div[text()='"+p+"']").click()
                except:
                    raise Exception('标准参量点击失败！')
            ##4.4 剩余参量数判断
            now = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[5]/span").text
            n = int(old) -int(now)
            if len(param_list) == n:
                flag_num = True
                
            ##4.5 选中参量通讯配置
            if is_setup == 'yes':
                gather_setup_path = base_path + '/Case/mbdb-param-add-setup-gather.xlsx' 
                gather_data = handle_excel.get_table_value(gather_setup_path)
                wc_type = gather_data[3]
                param_name = gather_data[4]
                keyword = gather_data[5]
                device_setup_type = gather_data[6]
                device_setup_item = gather_data[7]
                start_byte = gather_data[8]
                accident_low = gather_data[9]
                accident_high = gather_data[10]
                run_low = gather_data[11]
                run_high = gather_data[12]
                mileage_low = gather_data[13]
                view_order = gather_data[14]
                annotation = gather_data[15]
                alarm_foreign_key = gather_data[16]
                is_alarm = gather_data[17]
                alarm_value = gather_data[18]
                alarm_confirm = gather_data[19]
                is_reverse = gather_data[20]
                data_length = gather_data[21]
                data_type = gather_data[22]
                byte_order = gather_data[23]
                point_addr = gather_data[24]
                trans_group_num = gather_data[25]
                clean_method = gather_data[26]
                extend_field = gather_data[27]            
                try:
                    if '公用' == wc_type:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[1]/div[2]/descendant::span[contains(text(),'"+param_name+"')]").click()
                    if '一次侧' == wc_type:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[2]/descendant::span[contains(text(),'"+param_name+"')]").click()
                    if '二次侧' == wc_type:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[3]/descendant::span[contains(text(),'"+param_name+"')]").click()
                except:
                    raise Exception('选中参量点击失败！')
                try:
                    if device_setup_item is not None and device_setup_item !='':# 设备配置
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div").click()
                        time.sleep(2)
                        browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                        # elem = browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']")
                        # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                        time.sleep(2)
                        browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_item+"']").click()
                    if start_byte is not None:# 开始字节
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[5]/div/div[1]/input").send_keys(start_byte)
                    if accident_low is not None:# 事故低报警
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[6]/div/div[1]/input").send_keys(accident_low)
                    if accident_high is not None:# 事故高报警
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[7]/div/div[1]/input").send_keys(accident_high)
                    if run_low is not None:# 运行低报警
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[8]/div/div[1]/input").send_keys(run_low)                        
                    if run_high is not None:# 运行高报警
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[9]/div/div[1]/input").send_keys(run_high)
                    if mileage_low is not None:# 量程低报警
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[10]/div/div[1]/input").send_keys(mileage_low)
                    if view_order is not None:# 显示顺序
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[11]/div/div[1]/input").send_keys(view_order)
                    if annotation is not None:# 注释
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[12]/div/div[1]/input").send_keys(annotation)
                    if alarm_foreign_key is not None and alarm_foreign_key != '':# 报警外键
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[13]/div/div").click()
                        browser.find_element_by_xpath("//span[text()='"+alarm_foreign_key+"']").click()
                    if is_alarm is not None and is_alarm != '':# 是否报警
                        old = '是'
                        try:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div[@class='el-switch is-checked']")
                        except:
                            old = '否'
                        if old != is_alarm:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div").click()
                    
                    if '0' == alarm_value or 0 == alarm_value: # 报警值0
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[1]").click()
                    if '1' == alarm_value or 1 == alarm_value:# 报警值1
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[2]").click()
                    if alarm_confirm is not None and alarm_confirm != '':# 是否确认报警
                        old = '是'
                        try:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[16]/div/div[@class='el-switch is-checked']")
                        except:
                            old = '否'
                        if old != alarm_confirm:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[16]/div/div").click()
                    if is_reverse is not None:# 是否取反
                        old = '是'
                        try:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div[@class='el-switch is-checked']")
                        except:
                            old = '否'
                        if old != is_reverse:
                            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div").click()
                    if data_length is not None and data_length != '':# 数据长度
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[18]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                    if data_type is not None and data_type != '':# 数据类型
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[19]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                    if byte_order is not None:# 字节序
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[20]/div/div/input").send_keys(byte_order)
                    if point_addr is not None:# 点地址
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[21]/div/div/input").send_keys(point_addr)
                    if trans_group_num is not None:# 传输组数
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[22]/div/div/input").send_keys(trans_group_num)
                    if clean_method is not None:# 上行清洗策略
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[23]/div/div/input").send_keys(clean_method)
                    if extend_field is not None:# 扩展字段
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[24]/div/div/input").send_keys(extend_field)
                except:
                    raise Exception('输入添加内容有误！')
        
        if '控制量' == param_type:
            ##4.1 网侧类型选择
            try:
                time.sleep(2)
                if '公用' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[1]").click()
                if '一次侧' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[2]").click()
                if '二次侧' == wc_type:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[3]").click()
            except:
                raise Exception('采集量-添加-标准参量-点击失败！')
            ##4.2 搜索参量关键字
            if  keyword is not None:
                try:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[3]/input").send_keys(keyword)
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/button").click()
                except:
                    raise Exception('查询失败！')
            ##4.3 选择参量并点击添加
            time.sleep(2)
            old = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[5]/span").text #添加前的剩余数
            
            for p in param_list: #点击参量进行添加
                try:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[4]/div[text()='"+p+"']").click()
                except:
                    raise Exception('标准参量点击失败！')
            ##4.4 剩余参量数判断
            now = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[5]/span").text #添加后的剩余数
            n = int(old) -int(now)
            if len(param_list) == n:
                flag_num = True
                
            ##4.5 选中参量通讯配置
            if is_setup == 'yes':  
                control_setup_path = base_path + '/Case/mbdb-param-add-setup-control.xlsx'                     
                control_data = handle_excel.get_table_value(control_setup_path)
                wc_type = control_data[3]
                param_name = control_data[4]
                keyword = control_data[5]
                device_setup_type = control_data[6]
                device_setup_item = control_data[7]
                start_byte = control_data[8]
                param_correction = control_data[9]
                value_high = control_data[10]
                value_low = control_data[11]
                type_marking = control_data[12]
                view_order = control_data[13]
                annotation = control_data[14]
                data_length = control_data[15]
                data_type = control_data[16]
                clean_method = control_data[17]
                extend_field = control_data[18]
                            
                try: #网侧类型下的配置参量选择
                    if '公用' == wc_type:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[2]/div[2]/div[1]/div[2]/descendant::span[contains(text(),'"+param_name+"')]").click()
                    if '一次侧' == wc_type:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[2]/div[2]/div[2]/descendant::span[contains(text(),'"+param_name+"')]").click()
                    if '二次侧' == wc_type:
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[2]/div[2]/div[3]/descendant::span[contains(text(),'"+param_name+"')]").click()
                except:
                    raise Exception('选中参量点击失败！')
                try: #通讯配置
                    if device_setup_item is not None and device_setup_item !='':# 设备配置
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div").click()
                        time.sleep(5)
                        browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                        time.sleep(5)
                        browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_item+"']").click()
                    if start_byte is not None:# 开始字节
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[5]/div/div[1]/input").send_keys(start_byte)
                    if param_correction is not None:# 参数修正值
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[6]/div/div[1]/input").send_keys(param_correction)
                    if value_high is not None:#  高限
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[7]/div/div[1]/input").send_keys(value_high)
                    if value_low is not None:# 低限
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[8]/div/div[1]/input").send_keys(value_low)                        
                    if type_marking is not None:# 分类标识
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[9]/div/div[1]/input").send_keys(type_marking)
                    if view_order is not None:# 显示顺序
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[10]/div/div[1]/input").send_keys(view_order)
                    if annotation is not None:# 注释
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[11]/div/div[1]/input").send_keys(annotation)
                    if data_length is not None and data_length != '':# 数据长度
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[12]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                    if data_type is not None and data_type != '':# 数据类型
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[13]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                    if clean_method is not None:# 上行清洗策略
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div/input").send_keys(clean_method)
                    if extend_field is not None:# 扩展字段
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/input").send_keys(extend_field)
                except:
                    raise Exception('输入添加内容有误！')
        ##4.6 确定或取消
        if 'yes' == button:
            # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/div/div/button[2]").click()
            browser.find_element_by_xpath("//span[text()='确定']").click()
            flag_button = True
        if 'clear' == button:
            # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/div/div/button[1]").click()
            browser.find_element_by_xpath("//span[text()='取消']").click()
            flag_button = True
        if 'close' == button:
            if '采集量' == type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[1]/button").click()
                flag_button = True
            if '控制量' == type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[1]/button").click()
                flag_button = True        
        if flag_button and flag_num:
            flag = True           
        return flag_button  

    def add_param_gather(self,data):
        flag = False
        mb_name = data[3]    
        is_setup = data[4]
        wc_type = data[5]
        param_name = data[6]
        keyword = data[7]
        device_setup_type = data[8]
        device_setup_item = data[9]
        start_byte = data[10]
        accident_low = data[11]
        accident_high = data[12]
        run_low = data[13]
        run_high = data[14]
        range_low = data[15]
        range_high = data[16]
        view_order = data[17]
        annotation = data[18]
        alarm_foreign_key = data[19]
        is_alarm = data[20]
        alarm_value = data[21]
        alarm_confirm = data[22]
        is_reverse = data[23]
        data_length = data[24]
        data_type = data[25]
        byte_order = data[26]
        point_addr = data[27]
        trans_group_num = data[28]
        clean_method = data[29]
        extend_field = data[30]      
        button = data[31]

        #1.先选择模板点表
        try:
            time.sleep(4)
            elem = browser.find_element_by_xpath("//span[text()='"+mb_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击模板失败！')
            
        #2.选择采集量
        # try:
        #     time.sleep(2)
        #     # browser.find_element_by_xpath("//div[@class='collection clearfix']/div[1]").click() 
        #     elem = browser.find_element_by_xpath("//div[@class='collection clearfix']/div[1]")
        #     browser.execute_script("arguments[0].click();",elem)
        # except:
        #     raise Exception('选择失败！')
            
        #3.点击添加
        try:
            time.sleep(3)
            browser.find_element_by_xpath("//div[@class='btn btn2']").click()
            # elem = browser.find_element_by_xpath("//div[@class='btn btn2']")
            # browser.execute_script("arguments[0].click();",elem)
        except:
            raise Exception('点击添加失败！')

        #4.输入添加内容
        ##4.1 网侧类型选择
        try:
            time.sleep(2)
            if '公用' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[1]").click()
            if '一次侧' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[2]").click()
            if '二次侧' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[2]/div[3]").click()
        except:
            raise Exception('网侧类型选择失败！')
        ##4.2 选择参量并点击
        time.sleep(2)
        try:
            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[1]/div[4]/div[text()='"+param_name+"']").click()
        except:
            raise Exception('标准参量点击失败！')
                
        ##4.3 选中参量通讯配置
        if is_setup == 'yes':      
            try:
                # if '公用' == wc_type:
                #     browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[1]/descendant::span[contains(text(),'"+param_name+"')]").click()
                # if '一次侧' == wc_type:
                #     browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[2]/descendant::span[contains(text(),'"+param_name+"')]").click()
                # if '二次侧' == wc_type:
                #     browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[3]/descendant::span[contains(text(),'"+param_name+"')]").click()
                time.sleep(3)
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[2]/div[2]/div[1]/div/descendant::span[contains(text(),'"+param_name+"')]").click()
            except:
                raise Exception('选中参量点击失败！')
            try:
                if start_byte is not None:# 开始字节
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[5]/div/div[1]/input").send_keys(start_byte)
                if device_setup_item is not None and device_setup_item !='':# 设备配置
                    time.sleep(3)
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div")
                    browser.execute_script("arguments[0].click();",elem)
                    time.sleep(2)
                    browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword)
                    # browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                    time.sleep(2)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_item+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_item+"']").click()
                
                if accident_low is not None:# 事故低报警
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[6]/div/div[1]/input").send_keys(accident_low)
                if accident_high is not None:# 事故高报警
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[7]/div/div[1]/input").send_keys(accident_high)
                if run_low is not None:# 运行低报警
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[8]/div/div[1]/input").send_keys(run_low)                        
                if run_high is not None:# 运行高报警
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[9]/div/div[1]/input").send_keys(run_high)
                if range_low is not None:# 量程低报警
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[10]/div/div[1]/input").send_keys(range_low)
                if view_order is not None:# 显示顺序
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[11]/div/div[1]/input").send_keys(view_order)
                if annotation is not None:# 注释
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[12]/div/div[1]/input").send_keys(annotation)
                if alarm_foreign_key is not None and alarm_foreign_key != '':# 报警外键
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[13]/div/div").click()
                    browser.find_element_by_xpath("//span[text()='"+alarm_foreign_key+"']").click()
                if is_alarm is not None and is_alarm != '':# 是否报警
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != is_alarm:
                        # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div").click()
                        elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div")
                        browser.execute_script("arguments[0].click();",elem)
                        
                if '0' == alarm_value or 0 == alarm_value: # 报警值0
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[1]").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[1]")
                    browser.execute_script("arguments[0].click();",elem)
                if '1' == alarm_value or 1 == alarm_value:# 报警值1
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[2]").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/label[2]")
                    browser.execute_script("arguments[0].click();",elem)
                if alarm_confirm is not None and alarm_confirm != '':# 是否确认报警
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[16]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != alarm_confirm:
                        # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[16]/div/div").click()
                        elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[16]/div/div")
                        browser.execute_script("arguments[0].click();",elem)
                if is_reverse is not None:# 是否取反
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != is_reverse:
                        # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div").click()
                        elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div")
                        browser.execute_script("arguments[0].click();",elem)
                if data_length is not None and data_length != '':# 数据长度
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[18]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[18]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]")
                    browser.execute_script("arguments[0].click();",elem)
                if data_type is not None and data_type != '':# 数据类型
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[19]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                    elem =  browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[19]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]")
                    browser.execute_script("arguments[0].click();",elem)
                if byte_order is not None:# 字节序
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[20]/div/div/input").send_keys(byte_order)
                    # elem = 
                    # browser.execute_script("arguments[0].click();",elem)
                if point_addr is not None:# 点地址
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[21]/div/div/input").send_keys(point_addr)
                    # elem = 
                    # browser.execute_script("arguments[0].click();",elem)
                if trans_group_num is not None:# 传输组数
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[22]/div/div/input").send_keys(trans_group_num)
                    # elem = 
                    # browser.execute_script("arguments[0].click();",elem)
                if clean_method is not None:# 上行清洗策略
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[23]/div/div/input").send_keys(clean_method)
                    # elem = 
                    # browser.execute_script("arguments[0].click();",elem)
                if extend_field is not None:# 扩展字段
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[24]/div/div/input").send_keys(extend_field)

            except:
                raise Exception('输入添加内容有误！')
            ##4.4 确定或取消
            try:
                if 'yes' == button:
                    time.sleep(6)
                    # browser.find_element_by_xpath("//span[text()='确定']").click()
                    elem = browser.find_element_by_xpath("//span[text()='确定']")
                    browser.execute_script("arguments[0].click()",elem)
                    flag = True
                if 'clear' == button:
                    browser.find_element_by_xpath("//span[text()='取消']").click()
                    flag = True
                if 'close' == button:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[3]/div/div[1]/button").click()
                    flag = True
            except:
                raise Exception('点击按钮失败！')
        return flag         
            
    def add_param_control(self,data):
        flag = False
        mb_name = data[3]
        is_setup = data[4]
        wc_type = data[5]
        param_name = data[6]
        keyword = data[7]
        device_setup_type = data[8]
        device_setup_item = data[9]
        start_byte = data[10]
        param_correction = data[11]
        value_high = data[12]
        value_low = data[13]
        type_marking = data[14]
        view_order = data[15]
        annotation = data[16]
        data_length = data[17]
        data_type = data[18]
        clean_method = data[19]
        extend_field = data[20]
        button = data[21]

        #1.先选择模板点表
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//span[text()='"+mb_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击模板失败！')
        
        #2.选择控制量
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='collection clearfix']/div[2]").click()   
        except:
            raise Exception('选择失败！')
            
        #3.点击添加
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='btn btn2']").click()
        except:
            raise Exception('点击添加失败！')

        ##4.1 网侧类型选择
        try:
            time.sleep(2)
            if '公用' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[1]").click()
            if '一次侧' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[2]").click()
            if '二次侧' == wc_type:
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[2]/div[3]").click()
        except:
            raise Exception('采集量-添加-标准参量-点击失败！')
       
        ##4.2 选择参量并点击添加
        time.sleep(2)
        try:
            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[1]/div[4]/div[text()='"+param_name+"']").click()
        except:
            raise Exception('标准参量点击失败！')
                
        ##4.3 选中参量通讯配置
        if is_setup == 'yes':                                
            try: #网侧类型下的配置参量选择
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[2]/div[2]/div[1]/descendant::span[contains(text(),'"+param_name+"')]").click()
            except:
                raise Exception('选中参量点击失败！')
            try: #通讯配置
                if device_setup_item is not None and device_setup_item !='':# 设备配置
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[4]/div/div")
                    browser.execute_script("arguments[0].click();",elem)
                    time.sleep(2)
                    # browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                    browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword)
                    time.sleep(2)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_item+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_item+"']").click()
                if start_byte is not None:# 开始字节
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[5]/div/div[1]/input").send_keys(start_byte)
                if param_correction is not None:# 参数修正值
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[6]/div/div[1]/input").send_keys(param_correction)
                if value_high is not None:#  高限
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[7]/div/div[1]/input").send_keys(value_high)
                if value_low is not None:# 低限
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[8]/div/div[1]/input").send_keys(value_low)                        
                if type_marking is not None:# 分类标识
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[9]/div/div[1]/input").send_keys(type_marking)
                if view_order is not None:# 显示顺序
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[10]/div/div[1]/input").send_keys(view_order)
                if annotation is not None:# 注释
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[11]/div/div[1]/input").send_keys(annotation)
                if data_length is not None and data_length != '':# 数据长度
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[12]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[12]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]")
                    browser.execute_script("arguments[0].click();",elem)
                if data_type is not None and data_type != '':# 数据类型
                    # browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[13]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                    elem = browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[13]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]")
                    browser.execute_script("arguments[0].click();",elem)    
                if clean_method is not None:# 上行清洗策略
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[14]/div/div/input").send_keys(clean_method)
                if extend_field is not None:# 扩展字段
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/form/span/div[15]/div/div/input").send_keys(extend_field)
            except:
                raise Exception('输入添加内容有误！')
            ##4.4 确定或取消
            try:
                if 'yes' == button:
                    time.sleep(6)
                    # browser.find_element_by_xpath("//span[text()='确定']").click()
                    elem = browser.find_element_by_xpath("//span[text()='确定']")
                    browser.execute_script("arguments[0].click();",elem)
                    flag = True
                if 'clear' == button:
                    browser.find_element_by_xpath("//span[text()='取消']").click()
                    flag = True
                if 'close' == button:
                    browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[4]/div/div[1]/button").click()
                    flag = True        
            except:
                raise Exception('点击按钮失败！')          
        return flag

    def update_param(self,data):
        flag_num = False
        flag_button = False
        flag = False
        mb_name = data[3]
        param_type = data[4]
        param_name = data[5]
        setup_list = data[6]
        button = data[7]

        #1.先选择模板
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//span[text()='"+mb_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击模板失败！')
            
        #2.选择采集量还是控制量
        try:
            time.sleep(2)
            if '采集量' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[1]").click()
            if '控制量' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[2]").click()   
        except:
            raise Exception('模板点表-采集量/控制量-选择失败！')
            
        #3.点击编辑
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[text()='"+param_name+"']/ancestor::div[@class='mainTable']/div[1]/div[3]/table/tbody/tr/td[9]/div/button[1]").click()
            elem = browser.find_element_by_xpath("//div[text()='"+param_name+"']/ancestor::div[@class='mainTable']/div[1]/div[3]/table/tbody/tr/td[9]/div/button[1]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('模板-采集量/控制量-点击编辑失败！')

        #4.修改参量配置内容
        if '采集量' == param_type: 
            setup_path = base_path + '/Case/mbdb-param-update-setup-gather.xlsx'  
            setup_data = handle_excel.get_table_value(setup_path)
            param_name = setup_data[4]
            keyword = setup_data[5]
            device_setup_type = setup_data[6]
            device_setup_item = setup_data[7]
            start_byte = setup_data[8]
            accident_low = setup_data[9]
            accident_high = setup_data[10]
            run_low = setup_data[11]
            run_high = setup_data[12]
            mileage_low = setup_data[13]
            view_order = setup_data[14]
            annotation = setup_data[15]
            alarm_foreign_key = setup_data[16]
            is_alarm = setup_data[17]
            alarm_value = setup_data[18]
            alarm_confirm = setup_data[19]
            is_reverse = setup_data[20]
            data_length = setup_data[21]
            data_type = setup_data[22]
            byte_order = setup_data[23]
            point_addr = setup_data[24]
            trans_group_num = setup_data[25]
            clean_method = setup_data[26]
            extend_field = setup_data[27]        
            try:
                if device_setup_item is not None and device_setup_item !='':# 设备配置
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[1]/div/div/div/input").click()
                    time.sleep(2)
                    # browser.find_element_by_xpath("/html/body/div[6]/div[1]/div[1]/ul/li/div/div/descendant::span[text()='"+device_setup_type+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                    # elem = browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']")
                    # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                    time.sleep(2)
                    # browser.find_element_by_xpath("/html/body/div[6]/div[1]/div[1]/ul/li/div/div/descendant::span[text()='"+device_setup_item+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_item+"']").click()
                if start_byte is not None:# 开始字节
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[2]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(start_byte)
                if accident_low is not None:# 事故低报警
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[3]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(accident_low)
                if accident_high is not None:# 事故高报警
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[4]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(accident_high)
                if run_low is not None:# 运行低报警
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[5]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(run_low)                        
                if run_high is not None:# 运行高报警
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[6]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(run_high)
                if mileage_low is not None:# 量程低报警
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[7]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(mileage_low)
                if view_order is not None:# 显示顺序
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[8]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(view_order)
                if annotation is not None:# 注释
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[9]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(annotation)
                if alarm_foreign_key is not None and alarm_foreign_key != '':# 报警外键
                    browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[10]/div/div").click()
                    browser.find_element_by_xpath("//span[text()='"+alarm_foreign_key+"']").click()
                if is_alarm is not None and is_alarm != '':# 是否报警
                    old = '是'
                    try:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[11]/div/div")
                    except:
                        old = '否'
                    if old != is_alarm:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[11]/div/div[@class='el-switch is-checked']").click()
              
                if '0' == alarm_value or 0 == alarm_value:# 报警值0
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[12]/div/div/label[1]").click()
                if '1' == alarm_value or 0 == alarm_value:# 报警值1
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[12]/div/div/label[2]").click()
                if alarm_confirm is not None and alarm_confirm != '':# 是否确认报警
                    old = '是'
                    try:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[13]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != alarm_confirm:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[13]/div/div").click()
                if is_reverse is not None and is_reverse != '':# 是否取反
                    old = '是'
                    try:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[14]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != is_reverse:
                        browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[14]/div/div").click()
                if data_length is not None and data_length != '':# 数据长度
                    browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[15]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                if data_type is not None and data_type != '':# 数据类型
                    browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[16]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                if byte_order is not None:# 字节序
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[17]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(byte_order)
                if point_addr is not None:# 点地址
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[18]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(point_addr)
                if trans_group_num is not None:# 传输组数
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[19]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(trans_group_num)
                if clean_method is not None:# 上行清洗策略
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[20]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(clean_method)
                if extend_field is not None:# 扩展字段
                    elem = browser.find_element_by_xpath(" //span[text()='采集量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[21]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(extend_field)
            except:
                raise Exception('输入添加内容有误！')
        
        if '控制量' == param_type:
            setup_path = base_path + '/Case/mbdb-param-update-setup-control.xlsx'
            setup_data = handle_excel.get_table_value(setup_path)
            param_name = setup_data[4]
            keyword = setup_data[5]
            device_setup_type = setup_data[6]
            device_setup_item = setup_data[7]
            start_byte = setup_data[8]
            param_correction = setup_data[9]
            value_high = setup_data[10]
            value_low = setup_data[11]
            type_marking = setup_data[12]
            view_order = setup_data[13]
            annotation = setup_data[14]
            data_length = setup_data[15]
            data_type = setup_data[16]
            clean_method = setup_data[17]
            extend_field = setup_data[18]
                    
            try: #通讯配置
                if device_setup_item is not None and device_setup_item !='':# 设备配置
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[1]/div/div/div/input").click()
                    # time.sleep(2)
                    # browser.find_element_by_xpath("/html/body/div[6]/div[1]/div[1]/ul/li/div/div/descendant::span[text()='"+device_setup_type+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_type+"']").click()
                    # elem = browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']")
                    # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                    time.sleep(2)
                    # browser.find_element_by_xpath("/html/body/div[6]/div[1]/div[1]/ul/li/div/div/descendant::span[text()='"+device_setup_item+"']").click()
                    browser.find_element_by_xpath("//span[text()='"+device_setup_item+"']").click()
                if start_byte is not None:# 开始字节
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[2]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(start_byte)
                if param_correction is not None:# 参数修正值
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[3]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(param_correction)
                if value_high is not None:#  高限
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[4]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(value_high)
                if value_low is not None:# 低限
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[5]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(value_low)                        
                if type_marking is not None:# 分类标识
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[6]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(type_marking)
                if view_order is not None:# 显示顺序
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[7]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(view_order)
                if annotation is not None:# 注释
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[8]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(annotation)
                if data_length is not None and data_length !='':# 数据长度
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[9]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                if data_type is not None and data_type != '':# 数据类型
                    browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[10]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                if clean_method is not None: # 上行清洗策略
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[11]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(clean_method)
                if extend_field is not None:# 扩展字段
                    elem = browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div[12]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(extend_field)
            except:
                raise Exception('输入添加内容有误！')
        ##4.6 确定或取消
        if 'yes' == button:
            # browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div/div/button[2]").click()
            browser.find_element_by_xpath("//span[text()='确定']").click()
            flag_button = True
        if 'clear' == button:
            # browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div/div/button[1]").click()
            browser.find_element_by_xpath("//span[text()='重置']").click()
            flag_button = True
        if 'close' == button:
            if '采集量' == type:
                browser.find_element_by_xpath("//span[text()='采集量-编辑']/parent::*/following-sibling::button").click()
                flag_button = True
            if '控制量' == type:
                browser.find_element_by_xpath("//span[text()='控制量-编辑']/parent::*/following-sibling::button").click()
                flag_button = True        
        if flag_button and flag_num:
            flag = True           
        return flag_button  
    def delete_param(self,data):
        flag_num = False
        flag_button = False
        flag = False
        mb_name = data[3]
        param_type = data[4]  
        param_name = data[5]
        button = data[6]

        #1.先选择模板
        try:
            time.sleep(2)
            elem = browser.find_element_by_xpath("//span[text()='"+mb_name+"']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击模板失败！')
        
        #2.选择采集量还是控制量
        try:
            time.sleep(2)
            if '采集量' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[1]").click()
            if '控制量' == param_type:
                browser.find_element_by_xpath("//div[@class='collection clearfix']/div[2]").click()   
        except:
            raise Exception('模板点表-采集量/控制量-选择失败！')
            
        #3.点击编辑
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[text()='"+param_name+"']/ancestor::div[@class='mainTable']/div[1]/div[3]/table/tbody/tr/td[9]/div/button[1]").click()
            elem = browser.find_element_by_xpath("//div[text()='"+param_name+"']/ancestor::div[@class='mainTable']/div[1]/div[3]/table/tbody/tr/td[9]/div/button[2]")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('模板-采集量/控制量-点击删除失败！')

        ##4.6 确定或取消
        if 'yes' == button:
            # browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div/div/button[2]").click()
            browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[3]/button[2]").click()
            flag_button = True
        if 'no' == button:
            # browser.find_element_by_xpath("//span[text()='控制量-编辑']/ancestor::div[@class='el-dialog']/div[2]/div/form/div/div/button[1]").click()
            browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[3]/button[1]").click()
            flag_button = True
        if 'close' == button:
            browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[1]/button").click()        
        if flag_button and flag_num:
            flag = True           
        return flag_button  
     
class ZzjgManage: #组织架构
    def __init__(self):
       
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','zzjg')
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/a/li/span[text()='组织架构管理']").click()
        except:
            raise Exception('点击组织架构模块失败！')
        assert "组织架构管理" in browser.title  
    def add(self,data):
        flag = False
        keyword = data[3]
        parent_name = data[4]
        button_add = data[5]
        company_name = data[6]
        annotation = data[7]      
        button = data[8]
        #搜索关键字
        if keyword is not None:
            try:
                time.sleep(2)
                elem= browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div/div/div[2]/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(keyword)
            except:
                raise Exception('搜索失败！')
        #添加及确认操作
        if parent_name is None or parent_name == '': #添加根节点
            time.sleep(2)
            #直接点击添加按钮
            browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div[2]/div[5]/button").click()
            time.sleep(2)
            if 'yes' == button_add:#确认根节点添加
                #点击确认按钮
                try:
                    browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[3]/button[2]").click()
                except:
                    raise Exception('点击确认添加失败！')
                #输入添加内容
                try:
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").send_keys(company_name)
                    browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(annotation)
                except:
                    raise Exception('输入内容出错！')
                #确认框操作
                if 'yes' == button:
                    try:
                        browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[2]").click()
                        flag = True
                    except:
                        raise Exception('点击确认失败！')
                if 'clear' == button:
                    try:
                        browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[1]").click()
                        flag = True
                    except:
                        raise Exception('点击重置失败！')
                if 'close' == button:
                    try:
                        browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                        flag = True 
                    except:
                        raise Exception('点击关闭失败！')

            if 'no' == button_add:#取消根节点添加
                try:
                    browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[3]/button[1]").click()
                    flag= True
                except:
                    raise Exception('点击取消根节点添加失败！')
            if 'close' == button_add:#关闭根节点添加
                try:
                    browser.find_element_by_xpath("//div[@aria-label='提示']/div/div[1]/button").click()
                    flag= True
                except:
                    raise Exception('点击关闭根节点添加失败！')
        else: #添加非根节点
            
            #先选择父组织
            try:
                time.sleep(2)
                # //div[@class='heating-dictionary']/div[1]/div[1]/descendant::span[text()='六分公司'] 备用
                # browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']").click()
                elem = browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']")
                browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            except:
                raise Exception('点击父组织失败！')
            
            #再点击添加按钮
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div[2]/div[5]/button").click()
            except:
                raise Exception('点击添加按钮失败！')            
            #输入添加内容
            try:
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input").send_keys(company_name)
                browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input").send_keys(annotation)
            except:
                raise Exception('输入添加内容失败！')             
            #确认框操作
            if 'yes' == button:
                try:
                    browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[2]").click()
                    flag = True
                except:
                    raise Exception('点击确认失败！')                    
            if 'clear' == button:
                try:
                    browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[1]").click()
                    flag = True
                except:
                    raise Exception('点击重置失败！')
            if 'close' == button:
                try:
                    browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                    flag = True
                except:
                    raise Exception('点击关闭失败！')
        
        return flag     
    
    def update(self,data):
        flag = False
        keyword = data[3]
        parent_name = data[4]
        name_old = data[5]
        name_new = data[6]
        annotation = data[7]
        button = data[8]
        #搜索关键字
        if keyword is not None:
            try:
                time.sleep(2)
                elem= browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div/div/div[2]/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(keyword)
            except:
                raise Exception('搜索失败！')
        
        #有上级公司，点击上级公司    
        if parent_name is not None and parent_name !='':
            try:
                time.sleep(2)    
                browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']").click()
                # elem = browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']")
                # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            except:
                raise Exception('点击上级公司失败！')
        #点击编辑按钮
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+name_old+"']/parent::*/parent::*/following-sibling::td[6]/div/button").click()
            elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+name_old+"']/parent::*/parent::*/following-sibling::td[6]/div/button")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击编辑按钮失败！')
        #输入修改内容
        try:
            time.sleep(2)
            if name_new is not None: 
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[1]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(name_new)
            if annotation is not None:
                elem = browser.find_element_by_xpath("//form[@class='el-form ruleForm']/div[2]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(annotation)
        except:
            raise Exception('输入内容有误！')
        #确认框操作
        try: 
            time.sleep(2)
            if 'yes' == button:
                browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[2]").click()
                flag = True
            if 'clear' == button:
                browser.find_element_by_xpath("//div[@class='el-dialog__footer']/span/button[1]").click()
                flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
        except:
            raise Exception('确认失败！')
        return flag     
    
    def delete(self,data):
        flag = False
        flag = False
        keyword = data[3]
        name = data[4]
        button = data[5]
        #搜索关键字
        if keyword is not None:
            try:
                time.sleep(2)
                elem= browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div/div/div[2]/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(keyword)
            except:
                raise Exception('搜索失败！')
        #执行删除操作
        try:
            #先选中要删除的组织名
            time.sleep(2)
            browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+name+"']").click()
        except:
            raise Exception('该组织名不存在！')
        #然后点击删除按钮
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+name+"']/following-sibling::div").click()       
        except:
            raise Exception('点击删除按钮失败！')
        #确认是否删除
        try:
            if 'yes' == button:
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button").click()
                flag = True
            if 'no' == button:
                # browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button").click()
                flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/following-sibling::button").click()
                flag = True
        except:
            raise Exception('确认失败！')
        return flag         

    def delete_many(self,data):
        flag = False
        keyword = data[3]
        parent_name = data[4]
        name_list = data[5]
        button = data[6]
        #搜索关键字
        if keyword is not None:
            try:
                time.sleep(2)
                elem= browser.find_element_by_xpath("//div[@class='heating-dictionary']/div[1]/div/div/div[2]/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(keyword)
            except:
                raise Exception('搜索失败！')
        
        #有上级公司，点击上级公司    
        if parent_name is not None and parent_name !='':
            try:
                time.sleep(2)    
                browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']").click()
                # elem = browser.find_element_by_xpath("//span[@class='custom-tree-node']/descendant::span[text()='"+parent_name+"']")
                # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
            except:
                raise Exception('点击上级公司失败！')
        #勾选要删除的组织
        try:
            for name in name_list:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+name+"']/parent::*/parent::*/preceding-sibling::td/div").click()
                # elem = browser.find_element_by_xpath("//div[@class='el-table__body-wrapper is-scrolling-none']/descendant::div[text()='"+name+"']/parent::*/parent::*/preceding-sibling::td/div")
                # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('勾选组织失败！')
        #点击【批量删除】
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='批量删除']").click()
        except:
            raise Exception('点击批量删除失败！')
        #确认框操作
        try:
            if 'yes' == button:
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button[2]").click()
                flag = True
            if 'no' == button:
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button[1]").click()
                flag = True
            if 'close' == button:
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/following-sibling::button").click()
                flag = True
        except:
            raise Exception('确认框操作失败！')
        return flag     

class RlzManage: #热力站
    def __init__(self):
       
        # url = handle_ini.get_value(url_ini_path,'server','host') + handle_ini.get_value(url_ini_path,'module','rlz')
        # browser.get(url)
        try:
            browser.find_element_by_xpath("//span[text()='热力站管理']").click()
        except:
            raise Exception('点击热力站管理模块失败！')
        assert "热力站管理" in browser.title      
    def add(self,data):
        flag = False
        #基本信息
        name = data[3]
        addr = data[4]
        company = data[5]
        keyword = data[6]
        build_date = data[7]
        rebuild_date = data[8]
        heating_area = data[9]
        online_area = data[10]
        longitude = data[11]
        dimensionality = data[12]
        contact_person =  data[13]
        contact_phone = data[14]
        status = data[15]
        #热力站信息
        ry = data[16]
        build_type = data[17]
        is_lifewater = data[18]
        is_mainstation = data[19]
        insulation_construction = data[20]
        give_heating_method = data[21]
        manage_method = data[22]
        pipeline_layout = data[23]
        charge_method = data[24]
        station_type = data[25]
        station_terrain = data[26]
        distance = data[27]
        protocol_type = data[28]
        #配置信息
        is_add_ctl = data[29]
        is_add_gp = data[30]
        is_setup_pro = data[31]
        is_setup_comm = data[32]
        
        #控制柜信息
        control_path = base_path+'/Case/rlz-add2-control.xlsx' #控制柜数据
        control_data = handle_excel.get_table_value(control_path)
        #机组列表
        group_path = base_path+'/Case/rlz-add3-group.xlsx' #机组数据
        group_data = handle_excel.get_table_value(group_path)
        
        #采集量模板配置
        pro_path = base_path+'/Case/rlz-add4-param.xlsx' #协议配置数据(参量配置)
        pro_data = handle_excel.get_table_value(pro_path)
        #采集量通讯配置
        gather_set_path= base_path+'/Case/rlz-add5-setup-gather.xlsx' #采集量配置
        ctl_set_path = base_path+'/Case/rlz-add6-setup-control.xlsx'#控制量配置
        #1.点击添加
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//span[text()='添加']").click()
            #由于分辨率导致按钮显示不出来，先使用JavaScript点击
            elem = browser.find_element_by_xpath("//span[text()='添加']")
            browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
        except:
            raise Exception('点击添加按钮失败！')

        #2.基本信息 //div[@class='steps-content']/form[1]
        try:
            
            #站点名称
            if name is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[1]/div/div/input").send_keys(name)
            
            #所属分公司
            if company is not None and company != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[3]/div/div").click()
                #搜索关键字
                if keyword is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::input[@placeholder='检索关键字']").send_keys(keyword) 
                #选择公司
                
                time.sleep(3)
                # browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']").click() 
                elem = browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']")
                browser.execute_script("arguments[0].click();", elem)
            #建站日期
            if build_date is not None and build_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[4]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #改造日期
            if rebuild_date is not None and rebuild_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[5]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #站点地址
            if addr is not None:
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[2]/div/div/input")
                elem.send_keys(addr)
            #供热面积
            if heating_area is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[6]/div/div/div/input").send_keys(heating_area)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[6]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(heating_area))
                elem.click()
            #在网面积
            if online_area is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[7]/div/div/div/input").send_keys(online_area)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[7]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(online_area)
                elem.click()
            #经度坐标
            if longitude is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[8]/div/div/div/input").send_keys(longitude)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[8]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(longitude))
            #维度坐标
            if dimensionality is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[9]/div/div/div/input").send_keys(dimensionality)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[9]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(dimensionality))
                
            #负责人
            if contact_person is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[10]/div/div/input").send_keys(contact_person)
            # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[10]/div/div/input")
            #联系电话
            if contact_phone is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[11]/div/div/input").send_keys(contact_phone)
            # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[11]/div/div/input")
            #状态
            
            if '正常' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[1]").click()    
            if '冻结' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[2]").click()    
            #点击【下一步】
            time.sleep(2)
            # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[13]/div/button").click()
            elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[13]/div/button")
            browser.execute_script("arguments[0].click();",elem)
        except:
            raise Exception('输入基本信息有误！')
        #3.热力站信息 //div[@class='steps-content']/form[2]
        try:
            
            #所属热源
            if ry is not None and ry != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[1]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+ry+"']").click()
            #建筑类型
            if build_type is not None and build_type != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[2]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+build_type+"']").click()
            #生活水
            try:
                default = '是'
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_lifewater != default:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div").click()
            #重点站
            try:
                default = '是'
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_mainstation != default:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div").click()
            #保温结构
            if '节能' == insulation_construction:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[1]").click()
            if '非节能' == insulation_construction:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[2]").click()
            #供热方式
            if '汽水' == give_heating_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[1]").click()
            if '水水' == give_heating_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[2]").click()
            #管理方式
            if '场营' == manage_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[1]").click()
            if '自营' == manage_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[2]").click()
            #管路布置
            if '串联供热' == pipeline_layout:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[1]").click()
            if '分户供热' == pipeline_layout:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[2]").click()
            #收费方式
            if '集体' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[1]").click()
            if '计量' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[2]").click()
            if '到户' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[3]").click()
            if '面积' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[4]").click()
            #站点类型
            if '人工站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[1]").click()
            if '监控站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[2]").click()
            if '管线监测点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[3]").click()
            if '监测站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[4]").click()
            #热力站地势
            if station_terrain is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[11]/div/div/input").send_keys(station_terrain)
            #离热源距离
            if distance is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[12]/div/div/div/input").send_keys(distance)
            #协议类型
            if protocol_type is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[13]/div/div/div/input").send_keys(protocol_type)
            #下一步
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[14]/div/button[2]").click()
        except:
            raise Exception('输入热力站信息有误！')
        #4.控制柜信息 //div[@class='steps-content']/div[1]
        if 'yes' == is_add_ctl:            
            try:
                #点击添加控制柜
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div/i").click()
                #遍历控制柜数据表
                n =0
                for box in control_data:
                    if box[3] == name and box[2]== 'yes':#选择热力站对应的控制柜
                        box_name = box[4]
                        box_ip = box[5]
                        box_port = box[6]
                        server_ip = box[7]
                        server_port = box[8]
                        card_num = box[9]
                        code = box[10]
                        protocol = box[11]
                        method = box[12]
                        #再次点击添加
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/div[2]").click()
                        n +=1
                        #激活控制柜界面
                        time.sleep(3)
                        browser.find_element_by_xpath("//div[@class='steps-content']/div/div/div[1]/div["+str(n)+"]").click()

                        #控制柜名称
                        if box_name is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[2]/div/div/input").send_keys(box_name)
                        #IP地址
                        if box_ip is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[3]/div/div/input").send_keys(box_ip)
                        #端口号
                        if box_port is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[4]/div/div/div/input").send_keys(box_port)
                        #通讯服务器IP
                        if server_ip is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[5]/div/div/input").send_keys(server_ip)
                        #通讯服务器端口号
                        if server_port is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[6]/div/div/div/input").send_keys(server_port)
                        #传输卡号
                        if card_num is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[7]/div/div/div/input").send_keys(card_num)
                        #控制器编码
                        if code is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[8]/div/div/div/input").send_keys(code)
                        #通讯协议
                        time.sleep(2)
                        if '天时PLC' == protocol is not None:                            
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[9]/div/div/label[1]").click()
                        if '天时474' == protocol is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[9]/div/div/label[2]").click()
                        if '天时448' == protocol is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[9]/div/div/label[3]").click()
                        #通讯方式
                        time.sleep(2)
                        if 'ADSL' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[1]").click()
                        if '光纤传输' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[2]").click()
                        if '无线传输' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[3]").click()
                        if 'GPRS' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[4]").click()
                        if 'CDMA' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[5]").click()
                        if '3G猫' == method is not None:
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[1]/form["+str(n)+"]/div[10]/div/div/label[6]").click()
            #下一步
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[2]/button[2]").click()
            except:
                raise Exception('输入控制柜信息有误！')
        else:
            try:
                time.sleep(2)
                #跳过
                browser.find_element_by_xpath("//div[@class='steps-content']/div[1]/div[2]/button[2]").click()  
            except:
                raise Exception('点击跳过失败！')
        #5.机组信息 //div[@class='steps-content']/div[2]
        if 'yes' == is_add_gp:
            try:   
                m = 0   
                ctl_num = '0'     
                for group in group_data:
                    if group[2] == 'yes':
                        control_num = str(group[4])
                        group_num = group[5]
                        group_name = group[6]
                        heating_meter_dia = group[7]
                        pipeline_dia = group[8]
                        heating_meter_num = group[9]
                        group_heating_area = group[10]
                        heating_meter_type = group[11]
                        develop = group[12]
                        install_location = group[13]
                        if ctl_num != control_num:
                            m = 0
                            ctl_num = control_num
                        
                        #点击控制柜
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='control-left']/div["+control_num+"]").click()
                        try:
                            #点击添加机组(大按钮)
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[2]/div[2]/div/i").click()                            
                        except:
                            time.sleep(2)
                        #添加机组 
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/div/div[last()]").click()
                        m +=1
                        
                        #激活机组页面
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/div/div["+str(m)+"]").click()
                        #机组编号
                        if group_num is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[1]/div/div/input").send_keys(group_num)
                        #机组名称
                        if group_name is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[2]/div/div/input").send_keys(group_name)
                        #热表口径
                        if heating_meter_dia is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[3]/div/div/div/input").send_keys(heating_meter_dia)
                        #管网口径
                        if pipeline_dia is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[4]/div/div/input").send_keys(pipeline_dia)
                        #热表编号
                        if heating_meter_num is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[5]/div/div/div/input").send_keys(heating_meter_num)
                        #供暖面积
                        if group_heating_area is not None:
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[6]/div/div/div/input").send_keys(str(group_heating_area))
                        #热表类型
                        time.sleep(2)
                        if '采暖' == heating_meter_type:
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[7]/div/div/label[1]").click()
                        if '生活水' == heating_meter_type:
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[7]/div/div/label[2]").click()
                        #分支
                        time.sleep(2)
                        if '有' == develop:
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[8]/div/div/label[1]").click()
                        if '无' == develop:
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[8]/div/div/label[2]").click()
                        #安装位置
                        time.sleep(2)
                        if install_location is not None:
                            browser.find_element_by_xpath("//div[@class='control-right']/div["+control_num+"]/div/form["+str(m)+"]/div[9]/div/div/textarea").send_keys(install_location)
                #下一步
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/div[2]/div[3]/button[2]").click()
            except:
                raise Exception('输入机组信息有误！')
        else:
            try:
                time.sleep(3)
                #跳过
                browser.find_element_by_xpath("//div[@class='steps-content']/div[2]/div[3]/button[2]").click()
                # elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[2]/div[3]/button[2]")
                # browser.execute_script("arguments[0].click();",elem)
            except:
                raise Exception('点击跳过失败！')
        #6.协议配置(参量配置) 
        # //div[@class='steps-content']/div[3]
        # 定位机柜：//div[@class='steps-content']/div[3]/div[2]/div[机柜位置]
        # 定位机组：//div[@class='steps-content']/div[3]/div[2]/div[机柜位置]/div[机组位置]/div/div[2]/button
        if 'yes' == is_setup_pro:
            time.sleep(2)
            for pro in pro_data:
                if pro[1] == 'yes':
                    ctl_num = str(int(pro[2])+1) #机柜位置
                    gr_num = str(int(pro[3])+1) #机组位置
                    mb_name = pro[4]#模板名称
                    param_type = pro[5]#参量类型
                    button = pro[6]
                    try:
                        # 点击添加按钮
                        time.sleep(2)
                        browser.find_element_by_xpath("//div[@class='steps-content']/div[3]/div[2]/div["+ctl_num+"]/div["+gr_num+"]/div/div[2]/button").click()
                        # elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[3]/div[2]/div["+ctl_num+"]/div["+gr_num+"]/div/div[2]/button")
                        # browser.execute_script("arguments[0].click();",elem)
                        # 选择某个模板下的采集量或控制量
                        time.sleep(2)
                        browser.find_element_by_xpath("//span[text()='添加协议']/ancestor::div[@class='el-dialog']/descendant::span[text()='"+mb_name+"']/parent::*/following-sibling::div[1]/descendant::span[text()='"+param_type+"']").click()
                        time.sleep(2)
                        if 'yes' == button:
                            browser.find_element_by_xpath("//span[text()='添加协议']/ancestor::div[@class='el-dialog']/div[3]/span/button[2]").click()
                        if 'no' == button:
                            browser.find_element_by_xpath("//span[text()='添加协议']/ancestor::div[@class='el-dialog']/div[3]/span/button[1]").click()
                    except:
                        raise Exception('协议配置有误！')
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/div[3]/div[3]/button[2]").click()
        
            #7.通讯配置(参量配置) //div[@class='steps-content']/div[4]
            #机柜：//div[@class='steps-content']/div[4]/div[1]/div[2]/div[机柜位置]
            #机组：//div[@class='steps-content']/div[4]/div[1]/div[2]/div[机柜位置]/div[机组位置+1]
            #采集量：//div[@class='steps-content']/div[4]/div[1]/div[2]/div[机柜位置]/div[机组位置+1]/descendant::span[text()='采集量']
            #网侧类型：//div[@class='steps-content']/div[4]/div[1]/div[2]/div[机柜位置]/div[机组位置+1]/descendant::span[text()='公用']
            #参量：//div[@class='steps-content']/div[4]/div[1]/div[2]/div[机柜位置]/div[机组位置+1]/descendant::span[text()='参量名']
            if 'yes' == is_setup_comm:
                gather_set = handle_excel.get_table_value(gather_set_path)
                control_set = handle_excel.get_table_value(ctl_set_path)
                #采集量及协议配置
                for ga in gather_set:
                    if ga[0] == 'yes':
                        ctl_num = str(ga[1])
                        grp_num = str(ga[2]+1)
                        wc_type = ga[3]
                        param_tag = ga[5]
                        keyword = ga[6]
                        device_setup_type = ga[7]
                        device_setup_item = ga[8]
                        start_byte = ga[9]
                        accident_low = ga[10]
                        accident_high = ga[11]
                        run_low = ga[12]
                        run_high = ga[13]
                        range_low = ga[14]
                        range_high = ga[15]
                        view_order = ga[16]
                        annotation = ga[17]
                        alarm_foreign_key = ga[18]
                        is_alarm = ga[19]
                        alarm_value = ga[20]
                        alarm_confirm = ga[21]
                        is_reverse = ga[22]
                        data_length = ga[23]
                        data_type = ga[24]
                        byte_order = ga[25]
                        point_addr = ga[26]
                        trans_group_num = ga[27]
                        clean_method = ga[28]
                        extend_field = ga[29]
                 
                        try:
                            #点击机柜
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]").click()
                            #点击机组
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]").click()
                            #点击参量类型
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='采集量']").click()
                            #点击网侧类型
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='"+wc_type+"']").click()
                            #点击参量名
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='"+param_tag+"']").click()                
                        except:
                            raise Exception('点击参量名失败！')
                        try:
                            if device_setup_item is not None and device_setup_item !='':# 设备配置
                                
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[4]/div/div/div/input").click()
                                if keyword is not None:
                                    time.sleep(2)
                                    browser.find_element_by_xpath("//body/div[last()]/descendant::input[@placeholder='检索关键字']").send_keys(keyword)
                                    time.sleep(2)
                                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                                else:
                                    time.sleep(2)
                                    try:
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                                    except:
                                        time.sleep(2)
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_type+"']").click()                                
                                        time.sleep(2)
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                            if start_byte is not None:# 开始字节
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[5]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(start_byte)
                            if accident_low is not None:# 事故低报警
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[6]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(accident_low)
                            if accident_high is not None:# 事故高报警
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[7]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(accident_high)
                            if run_low is not None:# 运行低报警
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[8]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(run_low)                        
                            if run_high is not None:# 运行高报警
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[9]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(run_high)
                            if range_low is not None:# 量程低报警
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[10]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(range_low)
                            if view_order is not None:# 显示顺序
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[11]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(view_order)
                            if annotation is not None:# 注释
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[12]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(annotation)
                            if alarm_foreign_key is not None and alarm_foreign_key != '':# 报警外键  or 报警类型
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[13]/div/div/div").click()
                                browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+alarm_foreign_key+"']").click()
                            if is_alarm is not None and is_alarm != '':# 是否报警
                                time.sleep(2)
                                old = '是'
                                try:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[14]/div/div")
                                except:
                                    old = '否'
                                if old != is_alarm:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[14]/div/div[@class='el-switch is-checked']").click()
                            time.sleep(2)
                            if '0' == alarm_value or 0 == alarm_value:# 报警值0
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[15]/div/div/label[1]").click()
                            if '1' == alarm_value or 1 == alarm_value:# 报警值1
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[15]/div/div/label[2]").click()
                            if alarm_confirm is not None and alarm_confirm != '':# 是否确认报警
                                time.sleep(2)
                                old = '是'
                                try:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[16]/div/div[@class='el-switch is-checked']")
                                except:
                                    old = '否'
                                if old != alarm_confirm:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[16]/div/div").click()
                            if is_reverse is not None and is_reverse != '':# 是否取反
                                time.sleep(2)
                                old = '是'
                                try:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[17]/div/div[@class='el-switch is-checked']")
                                except:
                                    old = '否'
                                if old != is_reverse:
                                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[17]/div/div").click()
                            if data_length is not None and data_length != '':# 数据长度
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[18]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                            if data_type is not None and data_type != '':# 数据类型
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[19]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                            if byte_order is not None:# 字节序
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[20]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(byte_order)
                            if point_addr is not None:# 点地址
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[21]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(point_addr)
                            if trans_group_num is not None:# 传输组数
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[22]/div/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(trans_group_num)
                            if clean_method is not None:# 上行清洗策略
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[23]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(clean_method)
                            if extend_field is not None:# 扩展字段
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[2]/form/div[24]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(extend_field)
                        except:
                            raise Exception('采集量配置修改有误！')
                for cl in control_set:
                    if cl[0] == 'yes':
                        ctl_num = str(cl[1])
                        grp_num = str(cl[2]+1)           
                        wc_type = cl[3]
                        param_name = cl[4]
                        keyword = cl[5]
                        device_setup_type = cl[6]
                        device_setup_item = cl[7]
                        start_byte = cl[8]
                        param_correction = cl[9]
                        value_high = cl[10]
                        value_low = cl[11]
                        type_marking = cl[12]
                        view_order = cl[13]
                        annotation = cl[14]
                        data_length = cl[15]
                        data_type = cl[16]
                        clean_method = cl[17]
                        extend_field = cl[18]
                        try:
                            time.sleep(2)
                            #点击机柜
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]").click()
                            #点击机组
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]").click()
                            #点击参量类型
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='控制量']").click()
                            #点击网侧类型
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='"+wc_type+"']").click()
                            #点击参量名
                            time.sleep(2)
                            browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[1]/div[2]/div["+ctl_num+"]/div["+grp_num+"]/descendant::span[text()='"+param_name+"']").click()                
                        except:
                            raise Exception('点击参量名失败！')        
                        try: 
                            if device_setup_item is not None and device_setup_item !='':# 设备配置
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[4]/div/div/div/input").click()
                                if keyword is not None:
                                    time.sleep(2)
                                    browser.find_element_by_xpath("//body/div[last()]/descendant::input[@placeholder='检索关键字']").send_keys(keyword)
                                    time.sleep(2)
                                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                                else:
                                    time.sleep(2)
                                    try:
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                                    except:
                                        time.sleep(2)
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_type+"']").click()                                
                                        time.sleep(2)
                                        browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                            if start_byte is not None:# 开始字节
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[5]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(start_byte)
                            if param_correction is not None:# 参数修正值
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[6]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(param_correction)
                            if value_high is not None:#  高限
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[7]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(value_high)
                            if value_low is not None:# 低限
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[8]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(value_low)                        
                            if type_marking is not None:# 分类标识
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[9]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(type_marking)
                            if view_order is not None:# 显示顺序
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[10]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(view_order)
                            if annotation is not None:# 注释
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[11]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(annotation)
                            if data_length is not None and data_length !='':# 数据长度
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[12]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                            if data_type is not None and data_type != '':# 数据类型
                                time.sleep(2)
                                browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[13]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                            if clean_method is not None: # 上行清洗策略
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[14]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(clean_method)
                            if extend_field is not None:# 扩展字段
                                time.sleep(2)
                                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[2]/div[2]/div/div/div[5]/form/div[15]/div/div/input")
                                elem.send_keys(" ")
                                elem.clear()
                                elem.send_keys(extend_field)
                        except:
                            raise Exception('输入添加内容有误！')    
                #下一步
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[3]/button[2]").click()
                except:
                    raise Exception('点击下一步失败！')
            else:
                try:            
                    #跳过通讯配置
                    time.sleep(2)
                    browser.find_element_by_xpath("//div[@class='steps-content']/div[4]/div[3]/button[3]").click()
                except:
                    raise Exception('点击跳过按钮失败！')
        else:
            try:#跳过协议配置
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/div[3]/div[3]/button[3]").click()
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/div[3]/div[3]/button[3]")
                browser.execute_script("arguments[0].click();",elem)
                flag = True
            except:
                raise Exception('点击跳过失败!')
        
        #8.完成
        try:
            time.sleep(2)
            # browser.find_element_by_xpath("//div[@class='steps-content']/div[5]/button").click()
            # flag = True
        except:
            raise Exception('点击确定有误！')
        return flag     
    def update(self,data):
        flag = False
        #基本信息
        name_old = data[3]
        name_new = data[4]
        addr = data[5]
        company = data[6]
        keyword = data[7]
        build_date = data[8]
        rebuild_date = data[9]
        heating_area = data[10]
        online_area = data[11]
        longitude = data[12]
        dimensionality = data[13]
        contact_person =  data[14]
        contact_phone = data[15]
        status = data[16]
        #热力站信息
        ry = data[17]
        build_type = data[18]
        is_lifewater = data[19]
        is_mainstation = data[20]
        insulation_construction = data[21]
        give_heating_method = data[22]
        manage_method = data[23]
        pipeline_layout = data[24]
        charge_method = data[25]
        station_type = data[26]
        station_terrain = data[27]
        distance = data[28]
        protocol_type = data[29]
        button = data[30]
        #点击编辑
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+name_old+"']/parent::*/parent::*/following-sibling::td[12]").click()
        except:
            raise Exception('点击添加按钮失败！')

        #2.基本信息 
        try:            
            #站点名称
            if name_new is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[1]/div/div/input").send_keys(name)
            #站点地址
            if addr is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[2]/div/div/input").send_keys(addr)
            #所属分公司
            if company is not None and company != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[3]/div/div").click()
                #搜索关键字
                if keyword is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::input[@placeholder='检索关键字']").send_keys(keyword) 
                #选择公司
                
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']").click() 
                # elem = browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']")
                # browser.execute_script("arguments[0].click();", elem)
            #建站日期
            if build_date is not None and build_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[4]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #改造日期
            if rebuild_date is not None and rebuild_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[5]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #供热面积
            if heating_area is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[6]/div/div/div/input").send_keys(heating_area)
            #在网面积
            if online_area is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[7]/div/div/div/input").send_keys(online_area)
            #经度坐标
            if longitude is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[8]/div/div/div/input").send_keys(longitude)
            #维度坐标
            if dimensionality is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[9]/div/div/div/input").send_keys(dimensionality)
            #负责人
            if contact_person is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[10]/div/div/input").send_keys(contact_person)
            #联系电话
            if contact_phone is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[11]/div/div/input").send_keys(contact_phone)
            #状态
            if '正常' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[1]").click()
            if '冻结' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[2]").click()    
            #点击【下一步】
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[13]/div/button").click()
        except:
            raise Exception('输入基本信息有误！')
        #2.热力站信息 
        try:
            #所属热源
            if ry is not None and ry != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[1]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+ry+"']").click()
            #建筑类型
            if build_type is not None and build_type != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[2]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+build_type+"']").click()
            #生活水
            try:
                default = '是'
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_lifewater != default:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div").click()
            #重点站
            try:
                default = '是'
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_mainstation != default:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div").click()
            #保温结构
            if '节能' == insulation_construction:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[1]").click()
            if '非节能' == insulation_construction:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[2]").click()
            #供热方式
            if '汽水' == give_heating_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[1]").click()
            if '水水' == give_heating_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[2]").click()
            #管理方式
            if '场营' == manage_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[1]").click()
            if '自营' == manage_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[2]").click()
            #管路布置
            if '串联供热' == pipeline_layout:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[1]").click()
            if '分户供热' == pipeline_layout:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[2]").click()
            #收费方式
            if '集体' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[1]").click()
            if '计量' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[2]").click()
            if '到户' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[3]").click()
            if '面积' == charge_method:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[4]").click()
            #站点类型
            if '人工站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[1]").click()
            if '监控站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[2]").click()
            if '管线监测点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[3]").click()
            if '监测站点' == station_type:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[4]").click()
            #热力站地势
            if station_terrain is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[11]/div/div/input").send_keys(station_terrain)
            #离热源距离
            if distance is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[12]/div/div/div/input").send_keys(distance)
            #协议类型
            if protocol_type is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[13]/div/div/div/input").send_keys(protocol_type)
            #下一步
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[14]/div/button[2]").click()
        except:
            raise Exception('输入热力站信息有误！')
        #3.完成
        try:
            if 'yes' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/div/button").click()
                flag = True
        except:
            raise Exception('点击确定有误！')
        return flag     
    def delete(self,data):
        flag = False
        name = data[3]
        button = data[4]
        #点击详细
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+name+"']/parent::*/parent::*/following-sibling::td[11]").click()
        except:
            raise Exception('点击详细按钮失败！')
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='删除']").click()
        except:
            raise Exception('点击删除按钮失败！')
        try:
            if 'close' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/following-sibling::button").click()
                flag = True
            if 'yes' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button[2]").click()
                flag = True
            if 'no' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='提示']/parent::*/parent::*/following-sibling::div[2]/button[1]").click()
                flag = True
        except:
            raise Exception('确认框操作失败！')
        return flag    
    def add_control_box(self,box):#控制柜
        flag = False
        flag_ctl = False
        flag_group = False
        try:
            rlz_name = box[3]
            box_name = box[4]
            box_ip = box[5]
            box_port = box[6]
            server_ip = box[7]
            server_port = box[8]
            card_num = box[9]
            code = box[10]
            protocol = box[11]
            method = box[12]
            button = box[13]
            is_add_group = box[14]
            #点击控制柜
            time.sleep(3)
            browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+rlz_name+"']/parent::*/parent::*/following-sibling::td[13]").click()
            # 控制柜：//div[@class='control-cabinet control-cabinet']/div[1]
            # 机组：//div[@class='control-cabinet control-cabinet']/div[2]
            #点击添加
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[1]").click()
            #控制柜名称
            if box_name is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[1]/div/div/input").send_keys(box_name)
            #IP地址
            if box_ip is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[2]/div/div/input").send_keys(box_ip)
            #端口号
            if box_port is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[3]/div/div/input").send_keys(box_port)
            #通讯服务器IP
            if server_ip is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[4]/div/div/input").send_keys(server_ip)
            #通讯服务器端口号
            if server_port is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[5]/div/div/input").send_keys(server_port)
            #传输卡号
            if card_num is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[6]/div/div/input").send_keys(card_num)
            #控制器编码
            if code is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[7]/div/div/input").send_keys(code)
            #通讯协议
            time.sleep(2)
            if '天时PLC' == protocol is not None:                            
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[8]/div/div/label[1]").click()
            if '天时474' == protocol is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[8]/div/div/label[2]").click()
            if '天时448' == protocol is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[8]/div/div/label[3]").click()
            #通讯方式
            time.sleep(2)
            if 'ADSL' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[1]").click()
            if '光纤传输' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[2]").click()
            if '无线传输' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[3]").click()
            if 'GPRS' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[4]").click()
            if 'CDMA' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[5]").click()
            if '3G猫' == method is not None:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/form/div[9]/div/div/label[6]").click()
        except:
            raise Exception('添加控制柜操作失败！')
        #点击确定
        try:
            if 'yes' == button:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/span/button[1]").click()
                flag_ctl = True
            if 'clear' == button:
                browser.find_element_by_xpath("//div[@class='control-cabinet control-cabinet']/div[1]/div[2]/span/button[2]").click()
                flag_ctl = True
        except:
            raise Exception('点击添加按钮失败！')
        #添加机组
        if 'yes' == is_add_group:
            #机组列表
            group_path = base_path+'/Case/rlz-add3-group.xlsx' #机组数据
            group_data = handle_excel.get_table_value(group_path)
            for group in group_data:
                if 'yes' == group[2]:
                    control_num = str(group[4])
                    group_num = group[5]
                    group_name = group[6]
                    heating_meter_dia = group[7]
                    pipeline_dia = group[8]
                    heating_meter_num = group[9]
                    group_heating_area = group[10]
                    heating_meter_type = group[11]
                    develop = group[12]
                    install_location = group[13]
                    #点击添加
                     
                    #输入添加内容
                    #机组编号
                    if group_num is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[1]/div/div/input").send_keys(group_num)
                    #机组名称
                    if group_name is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[2]/div/div/input").send_keys(group_name)
                    #热表口径
                    if heating_meter_dia is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[3]/div/div/input").send_keys(heating_meter_dia)
                    #管网口径
                    if pipeline_dia is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[4]/div/div/input").send_keys(pipeline_dia)
                    #热表编号
                    if heating_meter_num is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[5]/div/div/input").send_keys(heating_meter_num)
                    #供暖面积
                    if group_heating_area is not None:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[6]/div/div/input").send_keys(group_heating_area)
                    #热表类型
                    time.sleep(2)
                    if '采暖' == heating_meter_type:
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[7]/div/div/label[1]").click()
                    if '生活水' == heating_meter_type:
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[7]/div/div/label[2]").click()
                    #分支
                    time.sleep(2)
                    if '有' == develop:
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[8]/div/div/label[1]").click()
                    if '无' == develop:
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[8]/div/div/label[2]").click()
                    #安装位置
                    time.sleep(2)
                    if install_location is not None:
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[2]/form/div[9]/div/div/textarea").send_keys(install_location)            
                    if 'yes' == button:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[3]/span/button[1]").click()
                        flag_group = True
                    if 'clear' == button:
                        time.sleep(2)
                        browser.find_element_by_xpath("//body/div[last()-1]/div/div[3]/span/button[2]").click()
                        flag_group = True
        if flag_ctl and flag_group:
            flag = True
        return flag     

    def add_gather_param(self,data):#采集量
            flag = False
            rlz_name = data[3]
            control_box = data[4]
            group = data[5]
            wc_type = data[6]
            param_name = data[7]
            keyword = data[8]      
            device_setup_type = data[9]
            device_setup_item = data[10]
            start_byte = data[11]
            accident_low = data[12]
            accident_high = data[13]
            run_low = data[14]
            run_high = data[15]
            range_high = data[16]
            range_low = data[17]            
            view_order = data[18]
            annotation = data[19]
            alarm_foreign_key = data[20]
            is_alarm = data[21]
            alarm_value = data[22]
            alarm_confirm = data[23]
            is_reverse = data[24]
            data_length = data[25]
            data_type = data[26]
            byte_order = data[27]
            point_addr = data[28]
            trans_group_num = data[29]
            clean_method = data[30]
            extend_field = data[31]     
            button = data[32]
            #1.先点击【采集量】
            try:
                time.sleep(3)
                browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+rlz_name+"']/parent::*/parent::*/following-sibling::td[15]").click()
            except:
                raise Exception('点击采集量失败！')
            
            #2.点击【添加】
            try:
                time.sleep(3)
                # browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::span[text()='添加']").click() 
                elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::span[text()='添加']")
                browser.execute_script("arguments[0].click();",elem)  
            except:
                raise Exception('点击添加失败！')

            #3.输入添加内容
            ##3.1 网侧类型选择
            try:
                
                if '公用' == wc_type:
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='network-type clearfix']/div[1]").click()
                if '一次侧' == wc_type:
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='network-type clearfix']/div[2]").click()
                if '二次侧' == wc_type:
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='network-type clearfix']/div[3]").click()
            except:
                raise Exception('点击网侧类型失败！')
    
            ##3.3 选择参量并点击
            time.sleep(2)
            try:
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='search-content']/div[text()='"+param_name+"']").click()
            except:
                raise Exception('点击标准参量失败！')
                            
            ##3.4 选中参量通讯配置     
            try:
                # if '公用' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                # if '一次侧' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                # if '二次侧' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                time.sleep(3)
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
            except:
                raise Exception('选中参量点击失败！')
            #3.5协议通讯配置
            try:
                # 设备配置
                if device_setup_item is not None and device_setup_item !='':
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[4]/div/div").click()
                    time.sleep(2)
                    # browser.find_element_by_xpath("//input[@placeholder='检索关键字']").send_keys(keyword)
                    elem = browser.find_element_by_xpath("//body/div[last()]/descendant::input[@placeholder='检索关键字']")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(keyword)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']").click()
                    # elem = browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']")
                    # browser.execute_script("arguments[0].click();", elem) #使用JavaScript进行点击
                    time.sleep(3)
                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                #控制柜
                if control_box is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[5]/div/div/div[1]").click()
                    time.sleep(2)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item hover']/span[text()='"+control_box+"']").click()
                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+control_box+"']").click()
                #机组
                if group is not None:
                    group = group.split(',')
                    time.sleep(3)
                    for g in group:
                        try:
                            browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[6]/div/div/descendant::span[text()='"+g+"']/parent::label[@class='el-checkbox is-checked']")
                        except:
                            browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[6]/div/div/descendant::span[text()='"+g+"']").click()
                # 开始字节
                if start_byte is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[7]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(start_byte)
                # 事故低报警
                if accident_low is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[8]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(accident_low)
                # 事故高报警
                if accident_high is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[9]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(accident_high)
                # 运行低报警
                if run_low is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[10]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(run_low)                        
                # 运行高报警
                if run_high is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[11]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(run_high)
                # 量程低报警
                if range_low is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[12]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(range_low)
                # 量程高报警
                if range_high is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[13]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(range_high)
                # 显示顺序
                if view_order is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[14]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(view_order)
                # 注释
                if annotation is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[15]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(annotation)
                # 报警外键
                if alarm_foreign_key is not None and alarm_foreign_key != '':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[16]/div/div").click()
                    browser.find_element_by_xpath("//span[text()='"+alarm_foreign_key+"']").click()
                # 是否报警
                if is_alarm is not None and is_alarm != '':
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[17]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != is_alarm:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[17]/div/div").click()
                
                # 报警值0
                if '0' == alarm_value or 0 == alarm_value: 
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[18]/div/div/label[1]").click()
                # 报警值1
                if '1' == alarm_value or 1 == alarm_value:
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[18]/div/div/label[2]").click()
                # 是否确认报警
                if alarm_confirm is not None and alarm_confirm != '':
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[19]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != alarm_confirm:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[19]/div/div").click()
                # 是否取反
                if is_reverse is not None:
                    old = '是'
                    try:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[20]/div/div[@class='el-switch is-checked']")
                    except:
                        old = '否'
                    if old != is_reverse:
                        browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[20]/div/div").click()
                # 数据长度
                if data_length is not None and data_length != '':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[21]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                # 数据类型
                if data_type is not None and data_type != '':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[22]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                # 字节序
                if byte_order is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[23]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(byte_order)
                # 点地址
                if point_addr is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[24]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(point_addr)
                # 传输组数
                if trans_group_num is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[25]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(trans_group_num)
                # 上行清洗策略
                if clean_method is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[26]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(clean_method)
                # 扩展字段
                if extend_field is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[27]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(extend_field)
            except:
                raise Exception('输入添加内容有误！')
            
            #4.确定或取消
            if 'yes' == button:
                time.sleep(2)
                # browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::span[text()='保存']").click()
                elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::span[text()='保存']")
                browser.execute_script("arguments[0].click();",elem)
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='采集量-添加']/parent::*/following-sibling::button").click()
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'采集量维护')]/parent::*/following-sibling::button").click()
                flag = True
            if 'clear' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][1]/descendant::span[text()='重置']").click()
                time.sleep(2)
                # browser.find_element_by_xpath("//span[contains(text(),'采集量维护')]/parent::*/following-sibling::button").click()
                flag= True
            if 'close' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='采集量-添加']/parent::*/parent::*/button").click()
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'采集量维护')]/parent::*/following-sibling::button").click()
                flag = True      
            return flag 
    
    def update_gather_param(self,data):#采集量
            flag = False
            rlz_name = data[3]
            tag_name = data[8].strip()
            point_addr = data[9].strip()
           
            #1.搜索热力站
            try:
                time.sleep(3)
                elem = browser.find_element_by_xpath("//input[@placeholder='输入站点名称或者汉字']")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(rlz_name)
                elem.send_keys(Keys.ENTER)
            except:
                raise Exception('搜索热力站失败！')  
            #2.点击【采集量】
            try:
                time.sleep(3)
                browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+rlz_name+"']/parent::*/parent::*/following-sibling::td[15]").click()
            except:
                raise Exception('点击采集量失败！')
            
            #3.搜索
            try:
                time.sleep(3)
                elem = browser.find_element_by_xpath("//span[contains(text(),'采集量维护')]/parent::*/parent::*/following-sibling::div/div/div[1]/div[2]/div[1]/div[@class='inputName'][2]/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(tag_name)
                elem.send_keys(Keys.ENTER)
                elem.send_keys(Keys.ENTER)
            except:
                raise Exception('搜索标签名失败！')
            #4.点击【编辑】
            try:
                time.sleep(5)
                browser.find_element_by_xpath("//div[@class='el-table__fixed-body-wrapper']/descendant::div[text()='"+tag_name+"']/parent::*/parent::*/following-sibling::td[6]").click()
                # elem = browser.find_element_by_xpath("//div[@class='el-table__fixed-body-wrapper']/descendant::div[text()='"+tag_name+"']/parent::*/parent::*/following-sibling::td[6]")
                # browser.execute_script("arguments[0].click();", elem)
            except:
                raise Exception('点击编辑失败！')

            #5.修改点地址
            try:
                time.sleep(2)
                elem = browser.find_element_by_xpath("//section/div/div[4]/div[2]/div/div[2]/div/form/div[18]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(point_addr)
            except:
                raise Exception('修改点地址失败！')
            #6.保存
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//section/div/div[4]/div[2]/div/div[2]/div/form/div[22]/div/button[2]").click()
                flag = True  
            except:
                raise Exception("点击保存失败！")  
            try:
                #关闭添加窗口
                # elem = browser.find_element_by_xpath("//span[text()='采集量-编辑']/parent::*/following-sibling::button")
                # browser.execute_script("arguments[0].click();",elem)
                #关闭采集量维护窗口
                time.sleep(2)
                elem = browser.find_element_by_xpath("//span[contains(text(),'采集量维护')]/parent::*/following-sibling::button")
                browser.execute_script("arguments[0].click();",elem)
            except:
                raise Exception("关闭窗口失败！") 
            return flag 
          
    def add_control_param(self,data):#控制量
            flag = False
            rlz_name = data[3]
            control_box = data[4]
            group = data[5]
            wc_type = data[6]
            param_name = data[7]
            keyword = data[8]      
            device_setup_type = data[9]
            device_setup_item = data[10]
            start_byte = data[11]
            param_correction = data[12]
            value_high = data[13]
            value_low = data[14]
            type_marking = data[15]
            view_order = data[16]
            annotation = data[17]
            data_length = data[18]
            data_type = data[19]
            clean_method = data[20]
            extend_field = data[21]
            button = data[22]
            
            #1.点击【控制量】
            try:
                time.sleep(3)
                browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+rlz_name+"']/parent::*/parent::*/following-sibling::td[16]").click()
            except:
                raise Exception('点击控制量失败！')
            
            #2.点击添加
            try:
                time.sleep(3)
                # browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/descendant::span[text()='添加']").click()
                elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/descendant::span[text()='添加']")
                browser.execute_script("arguments[0].click();", elem)
            except:
                raise Exception('点击添加失败！')
            
            #3.输入添加内容
            ##3.1 网侧类型选择
            try:
                time.sleep(2)
                if '公用' == wc_type:
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='network-type clearfix']/div[1]").click()
                if '一次侧' == wc_type:
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='network-type clearfix']/div[2]").click()
                if '二次侧' == wc_type:
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='network-type clearfix']/div[3]").click()
            except:
                raise Exception('点击网侧类型失败！')
            ##3.2 搜索参量关键字
            # if  keyword is not None:
            #     try:
            #         browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='el-input el-input--suffix']/input").send_keys(keyword)
            #     except:
            #         raise Exception('查询失败！')
            ##3.3 选择参量并点击
            time.sleep(2)
            try:
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='search-content']/div[text()='"+param_name+"']").click()
            except:
                raise Exception('点击标准参量失败！')
                            
            ##3.4 选中参量通讯配置     
            try:
                # if '公用' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                # if '一次侧' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                # if '二次侧' == wc_type:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']").click()
                time.sleep(3)
                elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='center']/descendant::span[text()='"+param_name+"']")
                browser.execute_script("arguments[0].click();",elem)
            except:
                raise Exception('选中参量点击失败！')
            #3.5协议通讯配置
            try: 
                # 设备配置
                if device_setup_item is not None and device_setup_item !='':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[4]/div/div").click()
                    time.sleep(3)
                    elem = browser.find_element_by_xpath("//body/div[last()]/descendant::input[@placeholder='检索关键字']")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(keyword)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_type+"']").click()
                    time.sleep(5)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item options selected hover']/descendant::span[text()='"+device_setup_item+"']").click()
                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+device_setup_item+"']").click()
                #控制柜
                if control_box is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[5]/div/div/div[1]").click()
                    time.sleep(2)
                    # browser.find_element_by_xpath("//li[@class='el-select-dropdown__item hover']/span[text()='"+control_box+"']").click()
                    browser.find_element_by_xpath("//body/div[last()]/descendant::span[text()='"+control_box+"']").click()
                #机组
                if group is not None:
                    group = group.split(',')
                    for g in group:
                        time.sleep(2)
                        try:
                            browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[6]/div/div/descendant::span[text()='"+g+"']/parent::label[@class='el-checkbox is-checked']")
                        except:
                            browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[6]/div/div/descendant::span[text()='"+g+"']").click()
                    
                # 开始字节
                if start_byte is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[7]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(start_byte)
                # 参数修正值
                if param_correction is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[8]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(param_correction)
                #  高限
                if value_high is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[9]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(value_high)
                # 低限
                if value_low is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[10]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(value_low)                        
                # 分类标识
                if type_marking is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[11]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(type_marking)
                # 显示顺序
                # if view_order is not None:
                #     browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[12]/div/div[1]/input").send_keys(view_order)
                # 注释
                if annotation is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[12]/div/div[1]/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(annotation)
                # 数据长度
                if data_length is not None and data_length != '':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[13]/div/div/label/descendant::span[contains(text(),'"+data_length+"')]").click()
                # 数据类型
                if data_type is not None and data_type != '':
                    browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[14]/div/div/label/descendant::span[contains(text(),'"+data_type+"')]").click()
                # 上行清洗策略
                if clean_method is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[15]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(clean_method)
                # 扩展字段
                if extend_field is not None:
                    elem = browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/div[2]/descendant::div[@class='right']/div[2]/div[1]/div/div/form/div[16]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(extend_field)
            except:
                raise Exception('输入添加内容有误！')
            #4.确定或取消
            if 'yes' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/descendant::span[text()='确定']").click()
                time.sleep(3)
                browser.find_element_by_xpath("//span[contains(text(),'控制量维护')]/parent::*/following-sibling::button").click()
                flag = True
            if 'clear' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//section/div/div[@class='collection-capacity-alert'][2]/descendant::span[text()='重置']").click()
                flag= True
            if 'close' == button:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='控制量-添加']/parent::*/parent::*/button").click()
                flag = True        
    
            return flag
          
    def copy_to_newrlz(self,data):#复制
        flag = False
        #复制源
        source_name = data[3]
        ctl_box = json.loads(data[4])
        param_type = data[5].split(",")
        #基本信息
        name = data[6]
        addr = data[7]
        company = data[8]
        keyword = data[9]
        build_date = data[10]
        rebuild_date = data[11]
        heating_area = data[12]
        online_area = data[13]
        longitude = data[14]
        dimensionality = data[15]
        contact_person =  data[16]
        contact_phone = data[17]
        status = data[18]
        #热力站信息
        ry = data[19]
        build_type = data[20]
        is_lifewater = data[21]
        is_mainstation = data[22]
        insulation_construction = data[23]
        give_heating_method = data[24]
        manage_method = data[25]
        pipeline_layout = data[26]
        charge_method = data[27]
        station_type = data[28]
        station_terrain = data[29]
        distance = data[30]
        protocol_type = data[31]
        button = data[32]
        #1.先点击【复制】
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='el-table__fixed-right']/descendant::div[text()='"+source_name+"']/parent::*/parent::*/following-sibling::td[17]").click()
        except:
            raise Exception('点击复制失败！')
        
        #2.复制子页面操作
        ##2.1 复制源
        try:
            time.sleep(3)
            for ctl in ctl_box.keys():
                ### 勾选控制柜
                browser.find_element_by_xpath("//span[text()='"+ctl+"']").click()
                groups = ctl_box.get(ctl)
                ## 勾选机组
                for group in groups:
                    browser.find_element_by_xpath("//span[text()='"+group+"']").click()
        except:
            raise Exception("勾选控制柜和机组失败！")
        ### 勾选参量类型
        try:
            time.sleep(2)
            for type in param_type:
                if '采集量' == type:
                    try:
                        browser.find_element_by_xpath("//label[text()='参量方式：']/following-sibling::div/div/label[1][@class='el-checkbox is-checked']")
                    except:
                        browser.find_element_by_xpath("//label[text()='参量方式：']/following-sibling::div/div/label[1]").click()
                if '控制量' == type:
                    try:
                        browser.find_element_by_xpath("//label[text()='参量方式：']/following-sibling::div/div/label[2][@class='el-checkbox is-checked']")
                    except:
                        browser.find_element_by_xpath("//label[text()='参量方式：']/following-sibling::div/div/label[2]").click()
        except:
            raise Exception("勾选参量类型失败！")
        ### 复制方式为【新建站点】
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//span[text()='新建站点']").click()
        except:
            raise Exception("勾选新建站点失败")
        ##2.2 输入新站信息        
        ### 基本信息 
        try:
            time.sleep(3)
            #站点名称
            if name is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[1]/div/div/input").send_keys(name)
            #站点地址
            if addr is not None:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[2]/div/div/input").send_keys(addr)
            #所属分公司
            if company is not None and company != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[3]/div/div").click()
                #搜索关键字
                if keyword is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::input[@placeholder='检索关键字']").send_keys(keyword) 
                #选择公司
                
                time.sleep(3)
                # browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']").click() 
                elem = browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper select-option']/descendant::span[text()='"+company+"']")
                browser.execute_script("arguments[0].click();", elem)
            #建站日期
            if build_date is not None and build_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[4]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #改造日期
            if rebuild_date is not None and rebuild_date !='':
                time.sleep(2)
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[5]/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(build_date)
            #供热面积
            if heating_area is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[6]/div/div/div/input").send_keys(str(heating_area))
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[6]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(heating_area))
                elem.click()
            #在网面积
            if online_area is not None:
                time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[7]/div/div/div/input").send_keys(str(online_area))
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[7]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(online_area))

            #经度坐标
            if longitude is not None:
                # time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[8]/div/div/div/input").send_keys(str(longitude))
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[8]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(longitude))
            #维度坐标
            if dimensionality is not None:
                # time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[9]/div/div/div/input").send_keys(str(dimensionality))
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[9]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(dimensionality))
            #负责人
            if contact_person is not None:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[10]/div/div/input").send_keys(contact_person)
            #联系电话
            if contact_phone is not None:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[11]/div/div/input").send_keys(contact_phone)
            #状态
            if '正常' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[1]").click()
            if '冻结' == status:
                browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[12]/div/div/label[2]").click()    
            #点击【下一步】
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/form[1]/div[13]/div/button").click()
        except:
            raise Exception('输入基本信息有误！')
        ### 热力站信息 
        try:
            
            #所属热源
            if ry is not None and ry != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[1]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+ry+"']").click()
            #建筑类型
            if build_type is not None and build_type != '':
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[2]/div/div/div/input").click()
                time.sleep(2)
                browser.find_element_by_xpath("//body/div[@class='el-select-dropdown el-popper']/descendant::span[text()='"+build_type+"']").click()
            #生活水
            try:
                default = '是'
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_lifewater != default:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[3]/div/div").click()
            #重点站
            try:
                default = '是'
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div[@class='el-switch is-checked']")
            except:
                default = '否'
            if is_mainstation != default:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[4]/div/div").click()
            #保温结构
            if '节能' == insulation_construction:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[1]").click()
            if '非节能' == insulation_construction:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[5]/div/div/label[2]").click()
            #供热方式
            if '汽水' == give_heating_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[1]").click()
            if '水水' == give_heating_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[6]/div/div/label[2]").click()
            #管理方式
            if '场营' == manage_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[1]").click()
            if '自营' == manage_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[7]/div/div/label[2]").click()
            #管路布置
            if '串联供热' == pipeline_layout:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[1]").click()
            if '分户供热' == pipeline_layout:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[8]/div/div/label[2]").click()
            #收费方式
            if '集体' == charge_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[1]").click()
            if '计量' == charge_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[2]").click()
            if '到户' == charge_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[3]").click()
            if '面积' == charge_method:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[9]/div/div/label[4]").click()
            #站点类型
            if '人工站点' == station_type:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[1]").click()
            if '监控站点' == station_type:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[2]").click()
            if '管线监测点' == station_type:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[3]").click()
            if '监测站点' == station_type:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[10]/div/div/label[4]").click()
            #热力站地势
            if station_terrain is not None:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[11]/div/div/input").send_keys(station_terrain)
            #离热源距离
            if distance is not None:
                # time.sleep(2)
                # browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[12]/div/div/div/input").send_keys((distance))
                elem = browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[12]/div/div/div/input")
                elem.send_keys(' ')
                elem.clear()
                elem.send_keys(str(distance))

            #协议类型
            if protocol_type is not None:
                # time.sleep(2)
                browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[13]/div/div/div/input").send_keys(protocol_type)
            #下一步
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/form[2]/div[14]/div/button[2]").click()
        except:
            raise Exception('输入热力站信息有误！')
        #3. 完成
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='steps-content']/div/button").click()
            flag = True
        except:
            raise Exception('点击确定失败！')
        return flag

class UserManage:#用户
    def __init__(self):
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[1]").click()
        except:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/div/span[text()='系统配置']").click()
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[1]").click()
        assert "用户管理" in browser.title  
    def add(self,data):
        flag = False
        username = data[3]
        truename = data[4]
        phone = data[5]
        position = data[6]
        password = data[7]
        company = data[8]
        role = data[9]
        button = data[10]

        #点击添加按钮
        try:
            browser.find_element_by_xpath("//span[text()='添加']").click()
        except:
            raise Exception('点击添加按钮失败！')
        #输入用户信息
        try:
            browser.find_element_by_xpath("//div[text()='用户信息']").click()
            #用户名
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[1]/div/div/input").send_keys(username)
            #真实姓名
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[2]/div/div/input").send_keys(truename)
            #电话
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[3]/div/div/input").send_keys(phone)
            #职位
            if position is not None:
                browser.find_element_by_xpath("//div[@id='pane-first']/form/div[4]/div/div/input").send_keys(position)
            #密码
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[5]/div/div/input").send_keys(password)   
        except:
            raise Exception('输入用户信息失败！')
        #所属组织机构
        browser.find_element_by_xpath("//div[text()='所属组织机构']").click()
        for com in company:
            try:
                browser.find_element_by_xpath("//span[text()='"+com+"']").click()
            except:
                raise Exception('点击勾选公司失败！')
        #分配角色
        browser.find_element_by_xpath("//div[text()='分配角色']").click()
        for ro in role:
            try:
                browser.find_element_by_xpath("//span[text()='"+ro+"']").click()
            except:
                raise Exception('点击勾选角色失败！')
        #点击提交
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def update(self,data):
        flag = False
        username_old = data[3]
        username_new = data[4]
        truename = data[5]
        phone = data[6]
        position = data[7]
        company_old = data[8]
        company_new = data[9]
        role_old = data[10]
        role_new = data[11]
        button = data[12]

        #点击编辑按钮
        try:
            browser.find_element_by_xpath("//div[text()='"+username_old+"']/parent::*/following-sibling::td[5]/div/button[1]").click()
        except:
            raise Exception('点击编辑按钮失败！')
        #输入用户信息
        try:
            #点击用户信息选项卡
            browser.find_element_by_xpath("//div[text()='用户信息']").click() 
            #用户名
            if username_new is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[1]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(username_new)
            #真实姓名
            if truename is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[2]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(truename)
            #电话
            if phone is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[3]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(phone)
            #职位
            if position is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[4]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(position)
        except:
            raise Exception('输入用户信息失败！')
        #所属组织机构
        browser.find_element_by_xpath("//div[text()='所属组织机构']").click()
        for com in company_old:
            try:
                browser.find_element_by_xpath("//span[text()='"+com+"']").click()
            except:
                raise Exception('点击勾选公司失败！')
        for com in company_new:
            try:
                browser.find_element_by_xpath("//span[text()='"+com+"']").click()
            except:
                raise Exception('点击勾选公司失败！')
        #分配角色
        browser.find_element_by_xpath("//div[text()='分配角色']").click()
        for ro in role_old:
            try:
                browser.find_element_by_xpath("//span[text()='"+ro+"']").click()
            except:
                raise Exception('点击勾选角色失败！')
        for ro in role_new:
            try:
                browser.find_element_by_xpath("//span[text()='"+ro+"']").click()
            except:
                raise Exception('点击勾选角色失败！')
        #点击提交
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def delete(self,data):
        flag = False
        username = data[3]
        button = data[4]
        try:
            browser.find_element_by_xpath("//div[text()='"+username+"']/parent::*/following-sibling::td[5]/div/button[2]").click()
        except:
            raise Exception('点击删除按钮失败！')
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击确定失败！')
        if 'no' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击取消失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

class RoleManage:#角色
    def __init__(self):
        try:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[2]").click()
        except:
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/div/span[text()='系统配置']").click()
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[2]").click()
        assert "角色管理" in browser.title  
    def add(self,data):
        flag = False
        name = data[3]
        code = data[4]
        describe = data[5]
        operate = data[6]
        data = data[7]
        button = data[10]

        #点击添加按钮
        try:
            browser.find_element_by_xpath("//span[text()='添加']").click()
        except:
            raise Exception('点击添加按钮失败！')
        #角色数据
        try:
            browser.find_element_by_xpath("//div[text()='角色数据']").click()
            #角色名称
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[1]/div/div/input").send_keys(name)
            #角色编码
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[2]/div/div/input").send_keys(code)
            #角色描述
            browser.find_element_by_xpath("//div[@id='pane-first']/form/div[3]/div/div/input").send_keys(describe)
        except:
            raise Exception('输入角色数据失败！')
        #功能权限
        browser.find_element_by_xpath("//div[text()='功能权限']").click()
        for op in operate:
            try:
                browser.find_element_by_xpath("//span[text()='"+op+"']").click()
            except:
                raise Exception('点击勾选功能权限失败！')
        #数据权限
        browser.find_element_by_xpath("//div[text()='数据权限']").click()
        for da in data:
            try:
                browser.find_element_by_xpath("//span[text()='"+da+"']").click()
            except:
                raise Exception('点击勾选数据权限失败！')
        #点击提交
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def update(self,data):
        flag = False
        name_old = data[3]
        name_new = data[4]
        code = data[5]
        describe = data[6]
        operate_old = data[7]
        operate_new = data[8]
        data_old = data[9]
        data_new = data[10]
        button = data[11]

        #点击编辑按钮
        try:
            browser.find_element_by_xpath("//div[text()='"+name_old+"']/parent::*/following-sibling::td[3]/div/button[1]").click()
        except:
            raise Exception('点击编辑按钮失败！')
        #角色数据
        try:
            #点击角色数据选项卡
            browser.find_element_by_xpath("//div[text()='角色数据']").click() 
            #角色名称
            if name_new is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[1]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(name_new)
            #角色编码
            if code is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[2]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(code)
            #角色描述
            if describe is not None:
                elem = browser.find_element_by_xpath("//div[@id='pane-first']/form/div[3]/div/div/input")
                elem.send_keys(" ")
                elem.clear()
                elem.send_keys(describe)
        except:
            raise Exception('输入角色数据失败！')
        #功能权限
        browser.find_element_by_xpath("//div[text()='功能权限']").click()
        for op in operate_old:
            try:
                browser.find_element_by_xpath("//span[text()='"+op+"']").click()
            except:
                raise Exception('点击勾选功能权限失败！')
        for op in operate_new:
            try:
                browser.find_element_by_xpath("//span[text()='"+op+"']").click()
            except:
                raise Exception('点击勾选功能权限失败！')
        #数据权限
        browser.find_element_by_xpath("//div[text()='数据权限']").click()
        for data in data_old:
            try:
                browser.find_element_by_xpath("//span[text()='"+data+"']").click()
            except:
                raise Exception('点击勾选数据权限失败！')
        for data in data_new:
            try:
                browser.find_element_by_xpath("//span[text()='"+data+"']").click()
            except:
                raise Exception('点击勾选数据权限失败！')
        #点击提交
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def delete(self,data):
        flag = False
        name = data[3]
        button = data[4]
        try:
            browser.find_element_by_xpath("//div[text()='"+name+"']/parent::*/following-sibling::td[3]/div/button[2]").click()
        except:
            raise Exception('点击删除按钮失败！')
        if 'yes' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击确定失败！')
        if 'no' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击取消失败！')
        if 'close' == button:
            try:
                browser.find_element_by_xpath("").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

class OperateManage:#权限
    def __init__(self):
        try:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[3]").click()
        except:
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/div/span[text()='系统配置']").click()
            time.sleep(2)
            browser.find_element_by_xpath("//div[@class='menu-wrapper nest-menu']/li/ul/div[3]").click()
        assert "权限管理" in browser.title  
    def add(self,data):
        flag = False
        parent_list = data[3]
        type = data[4]
        name = data[5]
        mark = data[6]
        data3 = data[7]
        data4 = data[8]
        button = data[9]

        #点击添加按钮       
        if parent_list is None:
            
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='添加一级节点']").click()
            except:
                raise Exception('点击添加一级节点失败！')
        else:
            parent = parent_list.split('-')
            num = len(parent)
            if num >1:
                for p in parent:#点开父节点
                    if num >1:
                        try:
                            time.sleep(2)
                            browser.find_element_by_xpath("//span[contains(text(),'"+p+"')]/preceding-sibling::div").click()
                        except:
                            raise Exception('点击父节点失败！')
                    num -= 1
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'"+parent[-1]+"')]/parent::*/parent::td/following-sibling::td[5]/div/button[1]").click()
            except:
                raise Exception('点击新增子项按钮失败！')
        #添加页面操作
        try:
            #菜单类型
            if '菜单' == type:
                time.sleep(2)
                browser.find_element_by_xpath("//form/div[1]/div/div/label[1]").click()
                #菜单名称
                time.sleep(2)
                browser.find_element_by_xpath("//form/div[2]/div/div/input").send_keys(name)
                #标识
                if mark is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[3]/div/div/input").send_keys(mark)
                #菜单路由
                if data3 is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[4]/div/div/div[1]/div/input").send_keys(data3)
                #菜单图标            
                if data4 is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[5]/div/div/div[1]/div/input").send_keys(data4)
            if '按钮' == type:
                time.sleep(2)
                browser.find_element_by_xpath("//form/div[1]/div/div/label[2]").click()
                #按钮名称
                time.sleep(2)
                browser.find_element_by_xpath("//form/div[2]/div/div/input").send_keys(name)
                #标识
                if mark is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[3]/div/div/input").send_keys(mark)
                #授权标识
                if data3 is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[4]/div/div/input").send_keys(data3)
                #排序编号            
                if data4 is not None:
                    time.sleep(2)
                    browser.find_element_by_xpath("//form/div[5]/div/div/div/input").send_keys(data4)
            
        except:
            raise Exception('新增失败！')
        #点击提交
        if 'yes' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'no' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='取消']").click()
                flag = True
            except:
                raise Exception('点击取消按钮失败！')
        if 'close' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def update(self,data):
        flag = False
        parent_list = data[3]
        type_old = data[4]
        type_new = data[5]
        name_old = data[6]
        name_new = data[7]
        mark = data[8]
        route = data[9]
        icon = data[10]
        button = data[11]

        #点击编辑按钮
        if parent_list is not None:
            parent = parent_list.split('-')
            for p in parent:#逐个点开父节点
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[contains(text(),'"+p+"')]/preceding-sibling::div").click()
                except:
                    raise Exception('点击父节点失败！')
        try:
            if '菜单' == type_old:
                time.sleep(3)
                browser.find_element_by_xpath("//span[contains(text(),'"+name_old+"')]/parent::*/parent::td/following-sibling::td[5]/div/button[2]").click()
            if '按钮' == type_old:
                time.sleep(3)
                browser.find_element_by_xpath("//span[contains(text(),'"+name_old+"')]/parent::*/parent::td/following-sibling::td[5]/div/button[1]").click()
        except:
            raise Exception('点击编辑按钮失败！')
        
        #修改页面操作
        time.sleep(2)
        try:
            #菜单类型
            if '菜单' == type_new:
                browser.find_element_by_xpath("//form/div[1]/div/div/label[1]").click()
                #菜单名称
                if name_new is not None:                
                    elem = browser.find_element_by_xpath("//form/div[2]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(name_new)
                #标识
                if mark is not None:
                    elem = browser.find_element_by_xpath("//form/div[3]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(mark)
                #菜单路由
                if route is not None:
                    elem = browser.find_element_by_xpath("//form/div[4]/div/div/div[1]/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(route)
                #菜单图标
                if icon is not None:
                    elem = browser.find_element_by_xpath("//form/div[5]/div/div/div[1]/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(icon)
            if '按钮' == type_new:        
                browser.find_element_by_xpath("//form/div[1]/div/div/label[2]").click()
                #按钮名称
                if name_new is not None:                
                    elem = browser.find_element_by_xpath("//form/div[2]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(name_new)
                #标识
                if mark is not None:
                    elem = browser.find_element_by_xpath("//form/div[3]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(mark)
                #授权标识
                if route is not None:
                    elem = browser.find_element_by_xpath("//form/div[4]/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(route)
                #排序编号
                if icon is not None:
                    elem = browser.find_element_by_xpath("//form/div[5]/div/div/div/input")
                    elem.send_keys(" ")
                    elem.clear()
                    elem.send_keys(icon)
        except:
            raise Exception('修改失败！')
        #点击提交
        if 'yes' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='提交']").click()
                flag = True
            except:
                raise Exception('点击提交按钮失败！')
        if 'no' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[text()='取消']").click()
                flag = True
            except:
                raise Exception('点击取消按钮失败！')
        if 'close' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//div[@class='el-dialog__header']/button").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag

    def delete(self,data):
        flag = False
        parent_list = data[3]
        type = data[4]
        name = data[5]
        button = data[6]
        if parent_list is not None:
            parent = parent_list.split('-')
            for p in parent:#逐个点开父节点
                try:
                    time.sleep(2)
                    browser.find_element_by_xpath("//span[contains(text(),'"+p+"')]/preceding-sibling::div").click()
                except:
                    raise Exception('点击父节点失败！')
        try:
            if '菜单' == type:
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'"+name+"')]/parent::*/parent::td/following-sibling::td[5]/div/button[3]").click()
            if '按钮' == type:
                time.sleep(3)
                browser.find_element_by_xpath("//span[contains(text(),'"+name+"')]/parent::*/parent::td/following-sibling::td[5]/div/button[2]").click()
        except:
            raise Exception('点击删除按钮失败！')
        if 'yes' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'确定')]").click()
                flag = True
            except:
                raise Exception('点击确定失败！')
        if 'no' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//span[contains(text(),'取消')]").click()
                flag = True
            except:
                raise Exception('点击取消失败！')
        if 'close' == button:
            try:
                time.sleep(2)
                browser.find_element_by_xpath("//button[@class='el-message-box__headerbtn']").click()
                flag = True
            except:
                raise Exception('点击关闭失败！')
        return flag
