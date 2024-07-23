import ctypes
from datetime import datetime
import time

import xlwt
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QUrl
from PyQt5.QtWidgets import QApplication, QMainWindow, QButtonGroup, QComboBox, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QBrush, QColor, QFont, QDesktopServices, QIcon

from UI import Ui_MainWindow

from PyQt5 import QtGui
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtGui import QDesktopServices

import cv2
import os
import numpy as np

import yaml

from YAML.parser import get_config
from tools import update_center_points, draw_info
from yolov5_62.tools.inferer import Yolov5

# 任务栏图标
winapi = ctypes.windll.shell32
winapi.SetCurrentProcessExplicitAppUserModelID("myappid")


class HyperLink(object):
    '''
    超链接控制
    '''

    def __init__(self, MyMainWindow, cfg):
        MyMainWindow.label_csdn.setOpenExternalLinks(True)  # 允许打开外部链接
        MyMainWindow.label_csdn.setCursor(Qt.PointingHandCursor)  # 更改光标样式
        # 设置网址超链接
        url_blog = QUrl("https://blog.csdn.net/qq_28949847/article/details/132412611")
        MyMainWindow.label_csdn.setToolTip(url_blog.toString())
        # 博客点击
        MyMainWindow.label_csdn.mousePressEvent = lambda event: self.open_url(
            url_blog) if event.button() == Qt.LeftButton else None

        MyMainWindow.label_mbd.setOpenExternalLinks(True)  # 允许打开外部链接
        MyMainWindow.label_mbd.setCursor(Qt.PointingHandCursor)  # 更改光标样式
        # 设置网址超链接
        url_mbd = QUrl("https://mbd.pub/o/author-cGeYm2xm/work")
        MyMainWindow.label_mbd.setToolTip(url_blog.toString())
        # 博客点击
        MyMainWindow.label_mbd.mousePressEvent = lambda event: self.open_url(
            url_mbd) if event.button() == Qt.LeftButton else None

        MyMainWindow.label_tb.setOpenExternalLinks(True)  # 允许打开外部链接
        MyMainWindow.label_tb.setCursor(Qt.PointingHandCursor)  # 更改光标样式
        # 设置网址超链接
        url_tb = QUrl("https://mbd.pub/o/author-cGeYm2xm/work")
        MyMainWindow.label_tb.setToolTip(url_blog.toString())
        # 博客点击
        MyMainWindow.label_tb.mousePressEvent = lambda event: self.open_url(
            url_tb) if event.button() == Qt.LeftButton else None

    # 打开超链接
    def open_url(self, url):
        QDesktopServices.openUrl(url)

    # 设置label不可见
    def label_invalid(self):
        MyMainWindow.label_csdn.setVisible(False)
        MyMainWindow.label_mbd.setVisible(False)
        MyMainWindow.label_tb.setVisible(False)


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, cfg=None):
        super().__init__()
        self.results = None
        self.result_img_name = None
        self.setupUi(self)
        # 加载超链接
        HyperLink(self, cfg)
        if cfg.CONFIG.MODE != "mbd":
            self.label_csdn.setVisible(False)
            self.label_mbd.setVisible(False)
            self.label_tb.setVisible(False)
        # 修改table的宽度
        self.update_table_width()
        self.start_type = None
        self.img = None
        self.img_path = None
        self.video = None
        self.video_path = None
        # 绘制了识别信息的frame
        self.img_show = None
        # 是否结束识别的线程
        self.start_end = False
        self.sign = True

        self.worker_thread = None
        self.result_info = None

        # 获取当前工程文件位置
        self.ProjectPath = os.getcwd()
        self.selected_text = 'All target'
        run_time = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        # 保存所有的输出文件
        self.output_dir = os.path.join(self.ProjectPath, 'output')
        if not os.path.exists(self.output_dir):
            os.mkdir(self.output_dir)
        result_time_path = os.path.join(self.output_dir, run_time)
        os.mkdir(result_time_path)
        # 保存txt内容
        self.result_txt = os.path.join(result_time_path, 'result.txt')
        with open(self.result_txt, 'w') as result_file:
            result_file.write(str(['序号', '图片名称', '录入时间', '识别结果', '目标数目', '用时', '保存路径'])[1:-1])
            result_file.write('\n')
        # 保存绘制好的图片结果
        self.result_img_path = os.path.join(result_time_path, 'img_result')
        os.mkdir(self.result_img_path)
        # 默认选择为所有目标
        self.comboBox_value = 'All target'

        self.number = 1
        self.RowLength = 0
        self.consum_time = 0
        self.input_time = 0

        # 打开图片
        self.pushButton_img.clicked.connect(self.open_img)
        # 打开文件夹
        self.pushButton_dir.clicked.connect(self.open_dir)
        # 打开视频
        self.pushButton_video.clicked.connect(self.open_video)

        # 绑定开始运行
        self.pushButton_start.clicked.connect(self.start)
        # 导出数据
        self.pushButton_export.clicked.connect(self.write_files)

        self.comboBox.activated.connect(self.onComboBoxActivated)
        self.comboBox.mousePressEvent = self.handle_mouse_press

        # 表格点击事件绑定
        self.tableWidget_info.cellClicked.connect(self.cell_clicked)

    def update_table_width(self):
        # 设置每列宽度
        column_widths = [50, 220, 120, 200, 80, 80, 140]
        for column, width in enumerate(column_widths):
            self.tableWidget_info.setColumnWidth(column, width)

    # 连接单元格点击事件
    def cell_clicked(self, row, column):
        result_info = {}
        # 判断此行是否有值
        if self.tableWidget_info.item(row, 1) is None:
            return
        # 图片路径
        self.img_path = self.tableWidget_info.item(row, 1).text()
        # 识别结果
        self.results = eval(self.tableWidget_info.item(row, 3).text())
        # 保存路径
        self.result_img_name = self.tableWidget_info.item(row, 6).text()
        self.img_show = cv2.imdecode(np.fromfile(self.result_img_name, dtype=np.uint8), -1)
        box = self.results[0][2]
        score = self.results[0][1]
        cls_name = self.results[0][0]
        result_info['label_xmin_v'] = int(box[0])
        result_info['label_ymin_v'] = int(box[1])
        result_info['label_xmax_v'] = int(box[2])
        result_info['label_ymax_v'] = int(box[3])
        result_info['score'] = score
        result_info['cls_name'] = cls_name
        self.get_comboBox_value(self.results)
        self.show_all(self.img_show, result_info)

    def handle_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self.sign = False
            # 清空列表
            self.comboBox.clear()
            if type(self.comboBox_value) == str:
                self.comboBox_value = [self.comboBox_value]
            self.comboBox.addItems(self.comboBox_value)
        QComboBox.mousePressEvent(self.comboBox, event)

    def onComboBoxActivated(self):
        self.sign = True
        self.selected_text = self.comboBox.currentText()
        result_info = {}
        lst_info = []
        # 所有目标，默认显示结果中的第一个
        if self.selected_text == 'All target':
            box = self.results[0][2]
            score = self.results[0][1]
            cls_name = self.results[0][0]
            lst_info = self.results
        else:
            for bbox in self.results:
                box = bbox[2]
                cls_name = bbox[0]
                score = bbox[1]
                lst_info = [[cls_name, score, box]]
                if self.selected_text == cls_name:
                    break
        result_info['label_xmin_v'] = int(box[0])
        result_info['label_ymin_v'] = int(box[1])
        result_info['label_xmax_v'] = int(box[2])
        result_info['label_ymax_v'] = int(box[3])
        result_info['score'] = score
        result_info['cls_name'] = cls_name

        self.img = cv2.imdecode(np.fromfile(self.img_path, dtype=np.uint8), cv2.IMREAD_COLOR)

        self.img_show = draw_info(self.img, lst_info)
        self.show_all(self.img_show, result_info)

    def resize_with_padding(self, image, target_width, target_height, padding_value=None):
        # 原始图像大小
        if padding_value is None:
            padding_value = [125, 107, 95]
        original_height, original_width = image.shape[:2]

        # 计算宽高比例
        width_ratio = target_width / original_width
        height_ratio = target_height / original_height

        # 确定调整后的图像大小和填充大小
        if width_ratio < height_ratio:
            new_width = target_width
            new_height = int(original_height * width_ratio)
            top = (target_height - new_height) // 2
            bottom = target_height - new_height - top
            left, right = 0, 0
        else:
            new_width = int(original_width * height_ratio)
            new_height = target_height
            left = (target_width - new_width) // 2
            right = target_width - new_width - left
            top, bottom = 0, 0

        # 调整图像大小并进行固定值填充
        resized_image = cv2.resize(image, (new_width, new_height))
        padded_image = cv2.copyMakeBorder(resized_image, top, bottom, left, right, cv2.BORDER_CONSTANT,
                                          value=padding_value)

        return padded_image

    def show_frame(self, img):
        self.update()  # 刷新界面
        if img is not None:
            # 尺寸适配
            shrink = self.resize_with_padding(img, self.label_img.width(), self.label_img.height())
            shrink = cv2.cvtColor(shrink, cv2.COLOR_BGR2RGB)
            QtImg = QtGui.QImage(shrink[:], shrink.shape[1], shrink.shape[0], shrink.shape[1] * 3,
                                 QtGui.QImage.Format_RGB888)
            self.label_img.setPixmap(QtGui.QPixmap.fromImage(QtImg))

    def open_img(self):
        try:
            # 选择文件  ;;All Files (*)
            self.img_path, filetype = QFileDialog.getOpenFileName(None, "选择文件", self.ProjectPath,
                                                                  "JPEG Image (*.jpg);;PNG Image (*.png);;JFIF Image (*.jfif)")
            if self.img_path == "":  # 未选择文件
                self.start_type = None
                return 0
            self.img_name = os.path.basename(self.img_path)
            # 显示相对应的文字
            self.label_img_path.setText(" " + self.img_path)
            self.label_dir_path.setText(" Select img folder")
            self.label_video_path.setText(" Select video")

            self.start_type = 'img'
            # 读取中文路径下图片
            self.img = cv2.imdecode(np.fromfile(self.img_path, dtype=np.uint8), cv2.IMREAD_COLOR)

            # 显示原图
            self.show_frame(self.img)
        except Exception as e:
            print(e)

    def open_dir(self):
        try:
            self.img_path_dir = QFileDialog.getExistingDirectory(None, "选择文件夹")
            if self.img_path_dir == '':
                self.start_type = None
                return 0

            self.start_type = 'dir'
            # 显示相对应的文字
            self.label_dir_path.setText(" " + self.img_path_dir)
            self.label_img_path.setText(" 选择图片文件")
            self.label_video_path.setText(" 选择视频文件")

            image_files = [file for file in os.listdir(self.img_path_dir) if file.lower().endswith(
                ('.bmp', '.dib', '.png', '.jpg', '.jpeg', '.pbm', '.pgm', '.ppm', '.tif', '.tiff'))]
            if image_files:
                self.img_path = os.path.join(self.img_path_dir, image_files[0])
            else:
                print('没有图片')
            self.img = cv2.imdecode(np.fromfile(self.img_path, dtype=np.uint8), cv2.IMREAD_COLOR)
            # 显示原图
            self.show_frame(self.img)
        except Exception as e:
            print(e)

    def open_video(self):
        try:
            # 选择文件
            self.video_path, filetype = QFileDialog.getOpenFileName(None, "选择文件", self.ProjectPath,
                                                                    "Video Files (*.mp4 *.avi)")
            if self.video_path == "":  # 未选择文件
                self.start_type = None
                return 0
            self.start_type = 'video'

            # 显示相对应的文字
            self.label_video_path.setText(" " + self.video_path)
            self.label_dir_path.setText(" 选择图片文件夹")
            self.label_img_path.setText(" 选择图片文件")

            cap = cv2.VideoCapture(self.video_path)
            # 读取第一帧
            ret, self.img = cap.read()
            # 显示原图
            self.show_frame(self.img)
        except Exception as e:
            print(e)

    def show_all(self, img, info):
        '''
        展示所有的信息
        '''
        self.show_frame(img)
        self.show_info(info)

    def start(self):
        try:
            if self.start_type == 'video':
                if self.pushButton_start.text() == '开始运行 >':
                    # 开始识别标志
                    self.start_end = False
                    # 开启线程，否则界面会卡死
                    self.worker_thread = WorkerThread(self.video_path, self)
                    self.worker_thread.result_ready.connect(self.show_all)
                    self.worker_thread.start()
                    # 修改文本为结束识别
                    self.pushButton_start.setText("结束运行 >")

                elif self.pushButton_start.text() == '结束运行 >':
                    # 修改文本为开始运行
                    self.pushButton_start.setText("开始运行 >")
                    # 结束识别
                    self.start_end = True

            elif self.start_type == 'img':
                _, result_info = self.predict_img(self.img)
                self.show_all(self.img_show, result_info)

            elif self.start_type == 'dir':
                if self.pushButton_start.text() == '开始运行 >':
                    # 开始识别标志
                    self.start_end = False
                    # 开启线程，否则界面会卡死
                    self.worker_thread = WorkerThread(None, self)
                    self.worker_thread.result_ready.connect(self.show_all)
                    self.worker_thread.start()
                    # 修改文本为结束识别
                    self.pushButton_start.setText("结束运行 >")
                elif self.pushButton_start.text() == '结束运行 >':
                    # 修改文本为开始运行
                    self.pushButton_start.setText("开始运行 >")
                    # 结束识别
                    self.start_end = True

        except Exception as e:
            print(e)

    def predict_img(self, img):
        result_info = {}
        t1 = time.time()
        self.result_img_name = os.path.join(self.result_img_path, self.img_name)
        self.results = y5.infer(img, conf_thres=conf_thres, classes=classes)
        self.consum_time = str(round(time.time() - t1, 2)) + 's'

        if len(self.results) == 0:
            return self.results, result_info
        self.input_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(self.result_txt, 'a+') as result_file:
            result_file.write(
                str([self.number, self.img_path, self.input_time, self.results, len(self.results), self.consum_time,
                     self.result_img_name])[1:-1])
            result_file.write('\n')
        # 显示识别信息
        self.show_table()
        self.number += 1

        # 获取下拉列表值
        self.get_comboBox_value(self.results)

        box = self.results[0][2]
        score = self.results[0][1]
        cls_name = self.results[0][0]

        # 画结果
        self.img_show = draw_info(img, self.results)
        # 保存结果图片
        cv2.imencode('.jpg', self.img_show)[1].tofile(self.result_img_name)

        # 类别
        result_info['cls_name'] = cls_name
        # 置信度
        result_info['score'] = round(score, 2)
        # 位置
        result_info['label_xmin_v'] = int(box[0])
        result_info['label_ymin_v'] = int(box[1])
        result_info['label_xmax_v'] = int(box[2])
        result_info['label_ymax_v'] = int(box[3])
        return self.results, result_info

    def get_comboBox_value(self, results):
        '''
        获取当前所有的类别和ID，点击下拉列表时，使用
        '''
        # 默认第一个是 所有目标
        lst = ["All target"]
        for bbox in results:
            cls_name = bbox[0]
            if cls_name not in lst:
                lst.append(str(cls_name))
        self.comboBox_value = lst

    def show_info(self, result):
        try:

            cls_name = result['cls_name']
            if len(cls_name) > 5:
                # 当字符串太长时，显示不完整
                lst_cls_name = cls_name.split('_')
                cls_name = lst_cls_name[0][:3] + '...' + lst_cls_name[-1]
            self.label_class.setText(str(cls_name))
            self.label_score.setText(str(result['score']))
            self.label_xmin_v.setText(str(result['label_xmin_v']))
            self.label_ymin_v.setText(str(result['label_ymin_v']))
            self.label_xmax_v.setText(str(result['label_xmax_v']))
            self.label_ymax_v.setText(str(result['label_ymax_v']))
            self.update()  # 刷新界面
        except Exception as e:
            print(e)

    def show_table(self):
        try:
            # 显示表格
            self.RowLength = self.RowLength + 1
            self.tableWidget_info.setRowCount(self.RowLength)
            for column, content in enumerate(
                    [self.number, self.img_path, self.input_time, self.results, len(self.results), self.consum_time,
                     self.result_img_name]):
                row = self.RowLength - 1
                item = QtWidgets.QTableWidgetItem(str(content))
                # 居中
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                # 设置字体颜色
                item.setForeground(QColor.fromRgb(197, 223, 250))
                self.tableWidget_info.setItem(row, column, item)
            # 滚动到底部
            self.tableWidget_info.scrollToBottom()
        except Exception as e:
            print('error:', e)

    def write_files(self):
        path, filetype = QFileDialog.getSaveFileName(None, "另存为", self.ProjectPath,
                                                     "Excel 工作簿(*.xls);;CSV (逗号分隔)(*.csv)")
        with open(self.result_txt, 'r') as f:
            lst_txt = f.readlines()
            data = [list(eval(x.replace('\n', ''))) for x in lst_txt]

        if path == "":  # 未选择
            return
        if filetype == 'Excel 工作簿(*.xls)':
            self.writexls(data, path)
        elif filetype == 'CSV (逗号分隔)(*.csv)':
            self.writecsv(data, path)

    def writexls(self, DATA, path):
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Data')
        for i, Data in enumerate(DATA):
            for j, data in enumerate(Data):
                ws.write(i, j, str(data))
        wb.save(path)
        QMessageBox.information(None, "成功", "数据已保存！", QMessageBox.Yes)

    def writecsv(self, DATA, path):
        try:
            f = open(path, 'w', encoding='utf8')
            for data in DATA:
                f.write(','.join('%s' % dat for dat in data) + '\n')
            f.close()
        except Exception as e:
            print(e)
        QMessageBox.information(None, "成功", "数据已保存！", QMessageBox.Yes)


class WorkerThread(QThread):
    '''
    识别视频进程
    '''
    result_ready = pyqtSignal(np.ndarray, dict)

    def __init__(self, path, main_window):
        super().__init__()
        self.path = path
        self.img = None
        self.main_window = main_window

    def run(self):
        if self.main_window.start_type == 'video':
            try:
                cap = cv2.VideoCapture(self.path)
                if not cap.isOpened():
                    raise ValueError("Unable to open video file or cam")
                if self.path == 0:
                    self.main_window.img_path = 'camera'
                else:
                    self.main_window.img_path = self.path
                frame_num = 0
                video_name = os.path.basename(self.path)
                while True:
                    self.main_window.img_name = video_name + '_' + str(frame_num) + '.jpg'

                    if not self.main_window.sign:
                        time.sleep(1)
                        continue

                    ret, self.img = cap.read()
                    frame_num += 1
                    if not ret or self.main_window.start_type is None or self.main_window.start_end is True:
                        # 修改文本为开始运行
                        self.main_window.pushButton_start.setText("开始运行 >")
                        break

                    results, result_info = self.main_window.predict_img(self.img)
                    if len(results) == 0:
                        continue

                    # 将数据传递回渲染
                    self.result_ready.emit(self.main_window.img_show, result_info)

            except Exception as e:
                print(e)
        if self.main_window.start_type == 'dir':
            for self.main_window.img_name in os.listdir(self.main_window.img_path_dir):
                if self.main_window.img_name.lower().endswith(
                        ('.bmp', '.dib', '.png', '.jpg', '.jpeg', '.pbm', '.pgm', '.ppm', '.tif', '.tiff')):
                    self.main_window.img_path = os.path.join(self.main_window.img_path_dir, self.main_window.img_name)
                    self.main_window.img = cv2.imdecode(np.fromfile(self.main_window.img_path, dtype=np.uint8), -1)
                    results, result_info = self.main_window.predict_img(self.main_window.img)
                    if len(results) == 0:
                        continue
                    # 将数据传递回渲染
                    self.result_ready.emit(self.main_window.img_show, result_info)
            else:
                # 修改文本为开始运行
                self.main_window.pushButton_start.setText("开始运行 >")


if __name__ == "__main__":
    path_cfg = 'config/configs.yaml'
    cfg = get_config()
    cfg.merge_from_file(path_cfg)
    if cfg.CONFIG.MODEL == 'YOLOV5':
        cfg_model = cfg.YOLOV5
    else:
        cfg_model = cfg.YOLOV5
    weights = cfg_model.WEIGHT
    conf_thres = float(cfg_model.CONF)
    classes = eval(cfg_model.CLASSES)

    device = cfg.CONFIG.DEVICE

    y5 = Yolov5(weights=weights, device=device)

    # 创建QApplication实例
    app = QApplication([])
    # 创建自定义的主窗口对象
    window = MyMainWindow(cfg)
    # 显示窗口
    window.show()
    # 运行应用程序
    app.exec_()
