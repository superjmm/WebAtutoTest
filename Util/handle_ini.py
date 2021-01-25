#coding = utf-8
import configparser

class HandleIni:
    def load_ini(self,file_path):
        cf = None
        try:
            cf = configparser.ConfigParser()
            cf.read(file_path,encoding='utf-8-sig')
        except:
            print('文件路径不正确！'+file_path)   
        return cf
    def get_value(self,file_path,node,key):
        '''
        获取文件中某个键值
        '''
        cf = self.load_ini(file_path)
        data = None
        try:
            data = cf.get(node,key)
        except:
             print('没有获取到值！')
        return data
handle_ini = HandleIni()
