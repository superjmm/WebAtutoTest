import pytest
import allure
import os

if __name__ == "__main__":   

    # pytest.main(['Run/test_00_dw.py','-s','--html=Report/report.html '])
    # pytest.main(['Run/test_10_bzdb.py','-s','--html=Report/report.html '])
    # pytest.main(['Run/test_20_ry.py','-s','--html=Report/report.html '])
    # pytest.main(['Run/test_30_mbdb.py','-s','--html=Report/report.html '])
    # pytest.main(['Run/test_40_rlz_param.py','-s','--html=Report/report.html '])
    pytest.main(['Run/test_40_rlz_copy.py','-s','--html=Report/report.html '])
    # 
    # 执行pytest测试，生成 Allure 报告需要的数据存在 Report/ 目录
    #pytest.main(['Run/test_cases.py','--alluredir','./allure-result'])
    # 使用allure ，生成测试报告

    # os.system('allure generate ./allure-result/ -o ./allure-report/ --clean') 
    # os.system('allure serve allure-result') 
