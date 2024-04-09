import os
import data_processor

def main():
    script_directory = os.path.dirname(__file__)
    file_path = os.path.realpath(script_directory + '/123.xlsx')
    data_processor.process_excel(os.path.normpath(file_path))
    print('success')
# 运行代码的入口
if __name__ == '__main__':
    main()
