import ctypes
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow

# 定义Windows API的常量和函数
winapi = ctypes.windll.shell32
winapi.SetCurrentProcessExplicitAppUserModelID("myappid")

if __name__ == '__main__':
    app = QApplication([])

    # 创建主窗口
    window = QMainWindow()

    # 设置任务栏图标
    window.setWindowIcon(QIcon('icon/app.jpg'))

    # 设置窗口标题
    window.setWindowTitle("My Application")

    # 设置窗口尺寸
    window.resize(800, 600)

    # 显示窗口
    window.show()

    # 运行应用程序
    app.exec_()