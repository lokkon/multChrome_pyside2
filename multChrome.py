from PySide2.QtWidgets import QApplication, QMessageBox, QAbstractItemView, QFileDialog
from PySide2.QtUiTools import QUiLoader
from PySide2.QtGui import QIcon
import win32process as win32p
import json
import os
import base64
from icon import img

class Draw():
    def __init__(self):

        self.ui = QUiLoader().load('ui\multChrome.ui')
        self.jsonfile = 'users.json'
        self.ui.add_user_button.clicked.connect(self.add_user)
        self.ui.open_button.clicked.connect(self.openChrome)
        self.ui.remove_button.clicked.connect(self.remove_user)
        self.ui.open_folder_button.clicked.connect(self.open_userfolder)
        self.ui.open_url_button.clicked.connect(self.open_site)
        self.ui.chrome_path_button.clicked.connect(self.chrome_findpath)
        self.ui.user_lb.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.ui.user_lb.doubleClicked.connect(self.openChrome)

        #生成ico
        tmp = open("tmpw.ico", "wb+")
        tmp.write(base64.b64decode(img))
        tmp.close()
        self.ui.setWindowIcon(QIcon("tmpw.ico"))
        os.remove("tmpw.ico")

    def readData(self):     #加载用户配置文件
        try:
            with open(self.jsonfile, 'r', encoding='utf-8') as f_obj:
                json_data = json.load(f_obj)
                #原本json文件只有一个列表存储用户名单，新版本json改为字典，加入chrome路径参数
                #判断如果json文件是列表，就重新格式化为字典
                if type(json_data) == list:
                    json_data = {"user_list": json_data, "chrome_path": "default"}
                    with open(self.jsonfile, 'w') as f_obj:
                        json.dump(json_data, f_obj)
                else:
                    self.chrome_path = json_data['chrome_path']
                    user_list = json_data['user_list']
                    for user in user_list:
                        self.ui.user_lb.addItem(user)
        #如果没有json文件（首次运行），则初始化json文件
        except FileNotFoundError:
            json_data = {"user_list": [], "chrome_path": "default"}
            with open(self.jsonfile, 'w') as f_obj:
                json.dump(json_data, f_obj)

    def add_user(self):     #添加新用户
        new_user = self.ui.add_user_entry.text()
        try:
            if new_user.encode('utf-8').isalnum():
                with open(self.jsonfile, 'r', encoding='utf-8') as f_obj:
                    json_data = json.load(f_obj)
                    if len(json_data['user_list']) == 0:
                        user_list = []
                        user_list.append(new_user)
                        json_data['user_list'] = user_list
                        self.ui.user_lb.addItem(new_user)
                    else:
                        user_list = json_data['user_list']
                        print(user_list)
                        if new_user in user_list:
                            QMessageBox.warning(self.ui, '错误', '该用户已存在！')
                        else:
                            user_list.append(new_user)
                            json_data['user_list'] = user_list
                            self.ui.user_lb.addItem(new_user)
                with open(self.jsonfile, 'w') as f_obj:
                    json.dump(json_data, f_obj)
            else:
                QMessageBox.warning(self.ui, '格式错误', '用户名仅限数字及字母')
        except FileNotFoundError:
            with open(self.jsonfile, 'w') as f_obj:
                json_data = {"user_list": [], "chrome_path": "default"}
                user_list.append(new_user)
                json_data['user_list'] = user_list
                with open(self.jsonfile, 'w') as f_obj:
                    json.dump(json_data, f_obj)
                self.ui.user_lb.addItem(new_user)

    def remove_user(self):      #移除用户
        selected_users = self.ui.user_lb.selectedItems()  # 提取选中
        users_num = self.ui.user_lb.count()
        with open(self.jsonfile, 'r', encoding='utf-8') as f_obj:
            json_data = json.load(f_obj)
            user_list = json_data['user_list']
        for i in range(users_num):
            for j in selected_users:
                try:
                    if self.ui.user_lb.item(i).text() == j.text():
                        self.ui.user_lb.takeItem(i)
                        user_list.remove(j.text())
                        json_data['user_list'] = user_list
                        print('removed')
                        with open(self.jsonfile, 'w') as f_obj:
                            json.dump(json_data, f_obj)
                except:
                    continue

    def openChrome(self):       #打开浏览器
        users = self.ui.user_lb.selectedItems()
        for user in list(users):
            try:
                if self.chrome_path == "default":
                    try:
                        self.chrome_path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                        chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text()
                        win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None,
                                                   None, win32p.STARTUPINFO())
                    except:
                        self.chrome_path = "C:\Program Files\Google\Chrome\Application\chrome.exe"
                        chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text()
                        win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None,
                                                   None, win32p.STARTUPINFO())
                else:
                    chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text()
                    win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None,
                                         None, win32p.STARTUPINFO())
            except:
                QMessageBox.warning(self.ui, '错误', '请检查Chrome路径')
                break

    def open_site(self):        #打开指定url
        users = self.ui.user_lb.selectedItems()
        site_url = self.ui.open_url_entry.text()
        for user in list(users):
            try:
                if self.chrome_path == "default":
                    try:
                        self.chrome_path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                        chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text() + site_url
                        win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None, None,
                                                   win32p.STARTUPINFO())
                    except:
                        self.chrome_path = "C:\Program Files\Google\Chrome\Application\chrome.exe"
                        chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text() + site_url
                        win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None, None,
                                                   win32p.STARTUPINFO())
                else:
                    chrome_url = self.chrome_path + '  --profile-directory="%s" ' % user.text() + site_url
                    win32p.CreateProcess(None, chrome_url, None, None, 0, win32p.CREATE_NO_WINDOW, None, None,
                                         win32p.STARTUPINFO())
            except:
                QMessageBox.warning(self.ui, '错误', '请检查Chrome路径')
                break

    def open_userfolder(self):      #打开用户文件夹
        users = self.ui.user_lb.selectedItems()
        for user in users:
            appdata = os.getenv("LOCALAPPDATA")
            user_data = appdata + r'\Google\Chrome\User Data' + '\\' + user.text()
            # print(user_data)
            if os.path.exists(user_data):
                os.startfile(user_data)
            else:
                QMessageBox.warning(self.ui, '错误', '用户 ' + user.text() + ' 文件不存在！')

    def chrome_findpath(self):      #设置chrome路径
        self.chrome_path, filetype = QFileDialog.getOpenFileName(
            self.ui,  # 父窗口对象
            "选择chrome",  # 标题
            r"c:\\",  # 起始目录
            "exe文件 (*.exe)"  # 选择类型过滤项，过滤内容在括号中
        )
        with open(self.jsonfile, 'r', encoding='utf-8') as f_obj:
            json_file = json.load(f_obj)
            json_file['chrome_path'] = self.chrome_path
        with open(self.jsonfile, 'w') as f_obj:
            json.dump(json_file, f_obj)

if __name__ == "__main__":
    app = QApplication([])
    draw = Draw()
    draw.ui.show()
    draw.readData()
    app.exec_()