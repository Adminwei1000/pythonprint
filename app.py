import socket
import threading
import os
import urllib.request
import win32print
import win32api
import win32con
import json
import subprocess
import datetime

import sys
def start_tcp_server(host='0.0.0.0', port=12345):


    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
        server_socket.bind((host, port))
        server_socket.listen(5)
        print(f"TCP服务器正在监听 {host}:{port}")
        log_error("star=================开始打印信息=================")

        while True:

            client_socket, client_address = server_socket.accept()
            log_error(f"已连接客户端 {client_address}")
            client_thread = threading.Thread(target=handle_client, args=(client_socket,))
            client_thread.start()
def print_pdfs(pdf_path,sys_currentprinter,print_type):
    # 获取.exe文件的路径
    exe_path = sys.executable if getattr(sys, 'frozen', False) else __file__
    log_error(f"获取.exe文件的路径:{exe_path}")
    # 获取.exe文件所在的目录
    exe_dir = os.path.dirname(exe_path)
    log_error(f"获取.exe文件所在的目录:{exe_dir}")

    # 获取.exe文件所在目录的绝对路径
    absolute_path = os.path.abspath(exe_dir)
    log_error(f"获取.exe文件所在目录的绝对路径:{absolute_path}")

    # pdf_paths = f"{absolute_path}\\{pdf_path}"
    log_error(f"准备打印文件位置:{pdf_path}")

    # pdf_paths=f"D:\\python\\pythonprint\\{pdf_path}"
    # jjstatus=win32api.ShellExecute(0, "print", pdf_path, f'/d:"{currentprinter}"', ".", 0)
    try:
        # 使用Windows的print命令打印PDF文件
        message1="自定义打印机"
        if not sys_currentprinter:
            sys_currentprinter = win32print.GetDefaultPrinter()  # 获取默认打印机名称
            message1 = "系统采用默认打印机"

        GHOSTSCRIPT_PATH = 'D:\\python\\pythonprint\\GHOSTSCRIPT\\bin\\gswin32.exe'
        jjstatus=0
        print_type_msg=""

        # print_type=66
        # jjstatus=42

        if print_type=='1':
            print_type_msg="模式1"
            GSPRINT_PATH = 'D:\\python\\pythonprint\\GSPRINT\\gsprint.exe'
            jjstatus = win32api.ShellExecute(0, 'open', GSPRINT_PATH,
                                             '-ghostscript "' + GHOSTSCRIPT_PATH + '" -printer "' + sys_currentprinter + '" "' + pdf_path + '"',
                                             '.', 0)
        if print_type=='2':
            print_type_msg = "模式2"
            # 构建Ghostscript的命令行参数
            gs_params = f'-dSAFER -dBATCH -dNOPAUSE -sDEVICE=mswinpr2 -sOutputFile="%printer%{sys_currentprinter}" "{pdf_path}" -c quit'
            # 执行Ghostscript命令，并尝试将日志输出到文件
            jjstatus = win32api.ShellExecute(0, 'open', GHOSTSCRIPT_PATH,
                                             f'-q -dNOPAUSE -dBATCH  {gs_params}', '.',
                                             win32con.SW_HIDE)


        if jjstatus > 32:
            message=f"{print_type_msg}命令已成功发送至系统"
        else:
            message = f"{print_type_msg}命令发送失败"
        response = {
            'status': 'success',
            'message': message,
            'data': {"message1":message1,
                     'jjstatus':jjstatus,
                     'pdf_path':pdf_path,
                     'absolute_path':absolute_path
                     }

        }

        log_error('=================返回打印信息=================end')

        return response

        # return True
    except subprocess.CalledProcessError as e:
        log_error(f"打印PDF失败: {e}")
        return False
def list_all_pdf_files_with_subdirs(directory):
    # 确保目录存在
    if not os.path.exists(directory):
        log_error(f"目录 {directory} 不存在。")
        return {}

    # 初始化返回的字典
    result = {}

    # 获取目录中所有文件和文件夹的列表
    files_and_folders = os.listdir(directory)

    # 过滤出所有以.pdf结尾的文件名
    pdf_files = [file for file in files_and_folders if file.lower().endswith('.pdf')]

    # 获取所有子目录的路径
    sub_directories = [os.path.join(directory, sub_dir) for sub_dir in files_and_folders if
                       os.path.isdir(os.path.join(directory, sub_dir))]
    # 递归地读取所有子目录中的PDF文件
    for sub_dir in sub_directories:
        sub_dir_name = os.path.basename(sub_dir)
        file_path=f"{directory}/{sub_dir_name}"
        files_and_folders = os.listdir(file_path)
        log_error(f'files_and_folders:{files_and_folders}')
        result[sub_dir] = files_and_folders
    # 将当前目录的PDF文件添加到结果中
    if directory not in result:
        result[directory] = []
    result[directory] += pdf_files

    return result

def delete_file(directory, file_name):
    # 确保目录存在
    if not os.path.exists(directory):
        response = {"status": "success", "message":f"目录 {directory} 不存在。"}
        return response
    # 获取目录中所有文件和文件夹的列表
    files_and_folders = os.listdir(directory)
    # 检查文件是否存在
    if file_name in files_and_folders:
        file_path = os.path.join(directory, file_name)
        os.remove(file_path)
        response = {"status": "success", "message": f"文件 {file_name} 已删除。"}
        log_error(f"文件 {file_name} 已删除。")
    else:
        response = {"status": "success", "message": f"文件 {file_name} 不存在。"}
    return response
def delete_all_pdf_files(directory):
    # 确保目录存在
    if not os.path.exists(directory):
        log_error(f"目录 {directory} 不存在。")
        return

    # 获取目录中所有文件和文件夹的列表
    files_and_folders = os.listdir(directory)

    # 过滤出所有以.pdf结尾的文件名
    pdf_files = [file for file in files_and_folders if file.lower().endswith('.pdf')]

    # 删除所有pdf文件
    for pdf_file in pdf_files:
        file_path = os.path.join(directory, pdf_file)
        os.remove(file_path)

def print_api():
    currentprinter = win32print.GetDefaultPrinter()  # 获取默认打印机名称

    # 你可以将数据传递给模板
    printers = win32print.EnumPrinters(2)
    printers_list = printers
    # 用于存储整理后的数据
    formatted_data = []
    # 遍历原始数据列表
    for printer_info in printers_list:
        # 提取必要的数据
        printer_dict = {'name': printer_info[2]}
        # 将字典添加到格式化数据列表中
        formatted_data.append(printer_dict)
    # response = {'currentprinter': currentprinter, 'machineList': formatted_data}
    # response = {'currentprinter': currentprinter, 'machineList': json.dumps(printers, ensure_ascii=False)}
    response = {'currentprinter': currentprinter, 'machineList': formatted_data}
    return response
def handle_client(client_socket):

    with client_socket:

        while True:
            data = client_socket.recv(1024).decode('utf-8')
            if not data:
                print(f"客户端 {client_socket.getpeername()} 断开连接")
                break

            #     参数处理
            params = data.split("&")
            request_params = {}
            for param in params:
                if '=' in param:
                    key, value = param.split('=', 1)
                    request_params[key] = value
            form_type = request_params['form_type']


            if form_type.lower() == "print":
                # 打印
                pdf_url = request_params['pdf_url']
                print_type = request_params['print_type']
                log_error(f'开始远程下载文件:{pdf_url}')
                # 1.进行下载文件
                base_download_directory = 'D:\python\pythonprint\dist\pdf'
                pdf_path = download_pdf(pdf_url,base_download_directory)
                if pdf_path is not None:
                    sys_currentprinter = request_params['sys_currentprinter']
                    # 2.进行打印
                    response_result = print_pdfs(pdf_path, sys_currentprinter,print_type)
                    response = {"status": "success", "message": "PDF已打印", "data": {'currentprinter': response_result}}
                else:
                    response = {"status": "success", "message": "下载PDF失败,本次打印失败"}

            elif form_type.lower() == "api_data":
                # 获取打印机
                resultData=print_api()
                response = {"status": "success","data":resultData, "message": "获取成功"}

            elif form_type.lower() == "get_pdf_list":
                # 读取pdf目录下所有文件名
                pdf_files_with_subdirs = list_all_pdf_files_with_subdirs('pdf')
                response = {"status": "success", "message": "PDF列表已获取", "data": pdf_files_with_subdirs}
            elif form_type.lower() == "del_all":
                # 清空pdf目录下所有文件
                pdf_mulu = request_params['pdf_mulu']
                delete_all_pdf_files(pdf_mulu)
                response = {"status": "success", "message": "PDF目录已清空"}

            elif form_type.lower() == "del_file":
                # 删除指定目录下指定文件
                directory = request_params['directory']
                file_name = request_params['file_name']
                response=delete_file(directory, file_name)
            else:
                response = {"status": "error", "message": "无效的form_type参数"}
            # 发送JSON格式的响应
            client_socket.sendall(json.dumps(response).encode('utf-8'))

def generate_unique_filename(url):
    # 假设这个函数能够根据URL生成唯一的文件名
    return url.split('/')[-1]

def log_error(message):
    log_dir = 'D:/python/pythonprint/dist/logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file_path = os.path.join(log_dir, 'download_errors.log')
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file_path, 'a', encoding='utf-8') as log_file:  # 指定编码为 utf-8
        log_file.write(f"{current_time} - {message}\n")

def get_date_folder(base_dir):
    date_folder = datetime.datetime.now().strftime('%Y-%m-%d')
    return os.path.join(base_dir, date_folder)
def download_pdf(url,base_download_dir):
    date_folder = get_date_folder(base_download_dir)
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)

    file_name = generate_unique_filename(url)
    file_path = os.path.join(date_folder, file_name)

    try:
        response = urllib.request.urlopen(url)
        if response.getcode() == 200:
            with open(file_path, 'wb') as f:
                f.write(response.read())

            log_error(f'下载成功，文件保存位于目录: {date_folder}')
            log_error(f'下载成功，文件保存位于: {date_folder}\{file_name}')
            log_error(f'下载成功，返回文件信息提供打印: {file_path}')
            return file_path
        else:
            error_message = f'下载失败，服务器响应码: {response.getcode()} for URL: {url}'
            print(error_message)
            log_error(error_message)
            return None
    except urllib.error.HTTPError as e:
        error_message = f'HTTP错误: {e.code} - {e.reason} for URL: {url}'
        print(error_message)
        log_error(error_message)
        return None
    except urllib.error.URLError as e:
        error_message = f'URL错误: {e.reason} for URL: {url}'
        print(error_message)
        log_error(error_message)
        return None
    except Exception as e:
        error_message = f'下载PDF失败: {str(e)} for URL: {url}'
        print(error_message)
        log_error(error_message)
        return None

    # date_folder = datetime.datetime.now().strftime('%Y-%m-%d')
    # pdf_dir = os.path.join('pdf', date_folder)
    #
    # if not os.path.exists(pdf_dir):
    #     os.makedirs(pdf_dir)
    #
    # file_name = generate_unique_filename(url)
    # file_path = os.path.join(pdf_dir, file_name)
    #
    # try:
    #     # 尝试下载PDF
    #     response = urllib.request.urlopen(url)
    #     if response.getcode() == 200:
    #         with open(file_path, 'wb') as f:
    #             f.write(response.read())
    #         print(f'下载成功，文件保存位于: {pdf_dir}')
    #         log_error(f'下载成功，文件保存位于: {pdf_dir}')
    #         return file_path
    #     else:
    #         error_message = f'下载失败，服务器响应码: {response.getcode()}'
    #         print(error_message)
    #         log_error(error_message)
    #         return None
    # except Exception as e:
    #     error_message = f"下载PDF失败: {e}"
    #     print(error_message)
    #     log_error(error_message)
    #     return None




    # # 获取当前日期作为文件夹名称
    # date_folder = datetime.datetime.now().strftime('%Y-%m-%d')
    # pdf_dir = os.path.join('pdf', date_folder)
    #
    # if not os.path.exists(pdf_dir):
    #     os.makedirs(pdf_dir)
    # # 生成唯一的文件名
    # file_name = generate_unique_filename(url)
    # file_path = os.path.join(pdf_dir, file_name)
    # try:
    #     # 下载PDF
    #     urllib.request.urlretrieve(url, file_path)
    #     print(f'下载成功文件保存位于:{pdf_dir}')
    #     return file_path
    # except Exception as e:
    #     print(f"下载PDF失败: {e}")
    #     return None
def generate_unique_filename(url):
    # 使用URL的哈希值作为文件名，确保唯一性
    import hashlib
    return hashlib.md5(url.encode('utf-8')).hexdigest() + '.pdf'

if __name__ == "__main__":
    import threading
    start_tcp_server()
