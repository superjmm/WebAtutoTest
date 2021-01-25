#coding=utf-8
import logging

class HandleLog:
    '''
    处理程序执行过程的日志
    '''
    def set_log(self,logger,log_level='DEBUG',log_path=None):
        '''
        设置日志输出方式
        Param:
            logger: 日志的实例，logger=logging.getLogger('日志')
            log_level='DEBUG':日志输出等级,DEBUG/INFO/WARNING/ERROR/CRITICAL
            log_path：日志文件路径，None是屏幕输出
        ''' 
        #设置日志输出等级
        if log_level == 'DEBUG':
            log_level = logging.DEBUG
        if log_level == 'INFO':
            log_level = logging.INFO
        if log_level =='WARNING':
            log_level = logging.WARNING
        if log_level =='ERROR':
            log_level = logging.ERROR
        if log_level =='CRITICAL':
            log_level = logging.CRITICAL
        logger.setLevel(log_level)
       
        #定义日志的输出格式
        formatter = logging.Formatter('%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s')
        
        if log_path is None:#屏幕输出日志
            ch = logging.StreamHandler()
            ch.setFormatter(formatter)
            logger.addHandler(ch)
        else:#存储到日志文件            
            fh = logging.FileHandler(log_path,encoding='utf-8')
            fh.setFormatter(formatter)
            logger.addHandler(fh)     

handle_log = HandleLog()