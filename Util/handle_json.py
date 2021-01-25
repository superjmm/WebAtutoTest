#coding=utf-8
import json

class HandleJson:
    def read_file(self,file_path):
        '''
        读取json文件
        :param:
            file_path:文件路径+文件名
        '''
        data = {}
        try:
            with open(file_path,encoding='UTF-8') as f:
                data = json.load(f)
        except:
            print('文件读取失败！'+file_path)
        return data

    def write_file(self,file_path,data):
        '''
        往json文件写入数据
        :param：
            file_path:文件路径+文件名
            data：json格式数据
        '''
        try:
            file_data = self.read_file(file_path)
            if file_data=='':
                file_data = {}
            file_data.update(data)
            with open(file_path,'w') as f:
                f.write(json.dumps(file_data))
        except:
            print('文件写入失败！'+file_path)

    def get_value(self,file_path,key):
        '''
        获取json文件中某个key的值
        param：
            file_path:文件路径+文件名
            data：key名
        '''
        data = self.read_file(file_path)
        return data.get(key)
        
handle_json = HandleJson()