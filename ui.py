from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QCheckBox, QSpinBox, QLabel, QFileDialog, QStatusBar, QSizePolicy, QMessageBox, QComboBox
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QPixmap, QScreen
from logic import AttendanceLogic
from PySide6.QtWidgets import QApplication
import random
import configparser
import os  # 用于处理文件路径
import sys  # 用于获取应用程序对象
import img


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logic = AttendanceLogic()
        self.timer = QTimer()  # 初始化 QTimer
        self.flash_count = 0
        self.final_result = []
        self.is_animating = False
        self.current_file = ""  # 当前选中的文件名称
        self.config_file = "config.ini"  # 配置文件路径
        self.setup_ui()
        self.connect_signals()
        self.center_on_screen()  # 启动时将窗口居中

    def setup_ui(self):
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()

        x = screen_geometry.width() / 1.8
        y = screen_geometry.height() / 1.8

        self.setWindowTitle("随机点名系统")
        # self.setGeometry(100, 100, 1600, 900)  # 初始窗口大小
        self.setGeometry(100, 100, x, y)
        self.setMinimumSize(800, 600)

        # 主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # 控制区域
        control_widget = QWidget()
        control_layout = QHBoxLayout()
        self.import_btn = QPushButton("导入名单")
        self.allow_repeat = QCheckBox("允许重复")
        self.num = QLabel("抽取数量:")
        self.start_btn = QPushButton("开始点名")

        self.combo_box = QComboBox()
        for i in range(1, 10):
            self.combo_box.addItem(str(i))
        self.combo_box.setCurrentIndex(0)

        # 将控件添加到控制区域
        control_layout.addWidget(self.import_btn)
        control_layout.addWidget(self.num)
        control_layout.addWidget(self.combo_box)
        control_layout.addWidget(self.allow_repeat)
        control_layout.addWidget(self.start_btn)
        control_widget.setLayout(control_layout)

        # 中央显示区域
        self.result_container = QWidget()
        self.result_layout = QGridLayout()
        self.result_container.setLayout(self.result_layout)
        self.result_container.setStyleSheet("""
            QWidget {
                background-color: rgba(255, 255, 255, 200);
                border-radius: 10px;
                padding: 10px;
            }
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #333;
            }
        """)
        # 设置显示区域尺寸策略为 Expanding
        self.result_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # 图片控件（显示在中央区域）
        self.image_label = QLabel()
        self.image_label.setPixmap(QPixmap(":/img/src/1.jpg"))
        self.image_label.setScaledContents(True)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setStyleSheet("background-color: #EEE; border: 1px solid #CCC;")
        self.image_label.setFixedSize(180, 180)  # 设置图片控件大小

        # 使用 QVBoxLayout 将结果显示和图片控件放在一起
        center_widget = QWidget()
        center_layout = QVBoxLayout()
        center_layout.addWidget(self.result_container, 1)  # 结果显示区域
        center_layout.addWidget(self.image_label, 0, Qt.AlignCenter)  # 图片控件
        center_layout.setSpacing(10)  # 设置控件间距
        center_widget.setLayout(center_layout)

        # 将控制区域和中央区域添加到主布局
        main_layout.addWidget(control_widget)
        main_layout.addWidget(center_widget, 1)
        main_layout.setSpacing(10)  # 设置控件间距
        main_layout.setContentsMargins(10, 10, 10, 10)  # 设置布局边距

        # 设置主布局
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # 状态栏（显示当前文件）
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("当前文件：未选择")

        self.load_last_file()
        self.load_window_size()

    def connect_signals(self):
        self.import_btn.clicked.connect(self.handle_import)
        self.start_btn.clicked.connect(self.handle_start)
        self.timer.timeout.connect(self.flash_result)  # 连接 QTimer 的超时信号

    def handle_import(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择学生名单", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.current_file = os.path.splitext(os.path.basename(file_name))[0]  # 获取文件名（不含后缀）
            self.status_bar.showMessage(f"当前文件：{self.current_file}")
            self.logic.load_students(file_name)
            self.save_last_file(file_name)

    def handle_start(self):
        if self.is_animating:
            return

        # count = self.spin_box.value()
        count = int(self.combo_box.currentText())
        allow_repeat = self.allow_repeat.isChecked()
        result = self.logic.random_select(count, allow_repeat)

        if isinstance(result, list) and len(result) > 0:
            self.final_result = result
        else:
            self.final_result = result

        # 初始化动画参数
        self.is_animating = True
        self.flash_count = 0
        self.timer.start(100)  # 每100ms更新一次

    def flash_result(self):
        font_size = int(self.width() / 30)  # 根据窗口宽度动态调整字体大小
        if self.flash_count < 20:  # 100ms * 30 = 3秒
            # 生成随机显示内容
            if self.logic.students:
                # count = self.spin_box.value()
                count = int(self.combo_box.currentText())
                random_names = random.choices(self.logic.students, k=count)
                self.update_result_display(random_names)

            # 随机颜色效果
            colors = ["#FF5722", "#2196F3", "#4CAF50", "#9C27B0"]
            self.result_container.setStyleSheet(f"""
                QWidget {{
                    background-color: rgba(255, 255, 255, 200);
                    border-radius: 10px;
                    padding: 10px;
                }}
                QLabel {{
                    color: {random.choice(colors)};
                    font-size: {font_size}px;
                    font-weight: bold;
                }}
            """)

            self.flash_count += 1
        else:
            self.timer.stop()
            self.update_result_display(self.final_result)
            self.result_container.setStyleSheet(f"""
                QWidget {{
                    background-color: rgba(255, 255, 255, 200);
                    border-radius: 10px;
                    padding: 10px;
                }}
                QLabel {{
                    color: #333;
                    font-size: {font_size}px;
                    font-weight: bold;
                }}
            """)
            self.is_animating = False

    def update_result_display(self, names):
        # 清空当前显示
        for i in reversed(range(self.result_layout.count())):
            self.result_layout.itemAt(i).widget().setParent(None)

        count = len(names)
        if count == 1:
            # 如果只抽取一个人，将名字放在正中心
            label = QLabel(names[0])
            label.setAlignment(Qt.AlignCenter)
            self.result_layout.addWidget(label, 1, 1, Qt.AlignCenter)  # 放在网格的中心位置
        else:
            # 如果抽取两人及以上，按 3x3 网格布局
            row, col = 0, 0
            for i, name in enumerate(names):
                label = QLabel(name)
                label.setAlignment(Qt.AlignCenter)
                self.result_layout.addWidget(label, row, col, Qt.AlignCenter)
                col += 1
                if col == 3:  # 每行显示三人
                    row += 1
                    col = 0

            # 如果抽取人数不足 9 人，填充空标签
            total_cells = 9
            for i in range(len(names), total_cells):
                spacer = QLabel("")
                spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                self.result_layout.addWidget(spacer, row, col, Qt.AlignCenter)
                col += 1
                if col == 3:  # 每行显示三人
                    row += 1
                    col = 0

    def resizeEvent(self, event):
        # 动态调整字体大小以适应窗口缩放
        font_size = int(self.width() / 30)  # 根据窗口宽度动态调整字体大小
        font_size2 = int(self.height() / 53)

        font = QFont()
        font.setPointSize(font_size2)
        self.import_btn.setFont(font)
        self.allow_repeat.setFont(font)
        self.num.setFont(font)
        self.start_btn.setFont(font)
        self.combo_box.setFont(font)
        self.status_bar.setFont(font)

        self.result_container.setStyleSheet(f"""
            QWidget {{
                background-color: rgba(255, 255, 255, 200);
                border-radius: 10px;
                padding: 10px;
            }}
            QLabel {{
                font-size: {font_size}px;
                font-weight: bold;
                color: #333;
            }}
        """)
        super().resizeEvent(event)
        self.save_window_size()

    def center_on_screen(self):
        # 获取屏幕的几何信息
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()

        # 计算窗口居中位置
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2

        # 设置窗口位置
        self.move(x, y)

    def save_last_file(self, file_path):
        """保存上次选择的文件路径"""
        config = configparser.ConfigParser()
        config.read(self.config_file)  # 读取配置文件
        if not config.has_section('Settings'):
            config.add_section('Settings')
        config.set('Settings', 'LastFilePath', file_path)  # 保存文件路径
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)

    def save_window_size(self):
        """保存当前窗口大小"""
        config = configparser.ConfigParser()
        config.read(self.config_file)  # 读取配置文件
        if not config.has_section('Settings'):
            config.add_section('Settings')
        config.set('Settings', 'WindowWidth', str(self.width()))  # 保存窗口宽度
        config.set('Settings', 'WindowHeight', str(self.height()))  # 保存窗口高度
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)

    def load_last_file(self):
        """加载上次选择的文件路径"""
        config = configparser.ConfigParser()
        config.read(self.config_file)  # 读取配置文件
        if config.has_section('Settings'):
            last_file_path = config.get('Settings', 'LastFilePath', fallback="")
            if last_file_path and os.path.exists(last_file_path):
                self.logic.load_students(last_file_path)
                self.current_file = os.path.splitext(os.path.basename(last_file_path))[0]
                if hasattr(self, 'status_bar'):
                    self.status_bar.showMessage(f"当前文件：{self.current_file}")

    def load_window_size(self):
        """加载上次的窗口大小"""
        config = configparser.ConfigParser()
        config.read(self.config_file)  # 读取配置文件
        if config.has_section('Settings'):
            # window_width = config.getint('Settings', 'WindowWidth', fallback=1600)  # 默认宽度
            # window_height = config.getint('Settings', 'WindowHeight', fallback=900)  # 默认高度
            window_width = config.getint('Settings', 'WindowWidth')
            window_height = config.getint('Settings', 'WindowHeight')
            self.resize(window_width, window_height)  # 设置窗口大小