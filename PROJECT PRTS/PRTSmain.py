# PRTSmain.py

import sys
import os
import time
import platform
import psutil
import GPUtil
import socket
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QFrame, QSizePolicy, QPushButton, QGraphicsOpacityEffect
)
from PySide6.QtCore import Qt, QTimer, QPoint, QThread, Signal
from PySide6.QtGui import QFont, QPixmap, QColor, QFontDatabase, QPainter, QBrush, QPolygon, QFontMetrics

# 常量配置
NOVECENTO_FONT = "Novecento Wide"  # 已安装字体名
BENDER_FONT = "Bender"             # 已安装字体名
CHINESE_FONT = "FZQuenyaSongS-R-GB"  # 方正准雅宋字体名（需已安装）

IMG_DIR = r"C:\Users\24177\Desktop\PROJECT PRTS"
NET_ON = os.path.join(IMG_DIR, "NET-ON.png")
NET_OFF = os.path.join(IMG_DIR, "NET-OFF.png")
SPLASH_IMG = os.path.join(IMG_DIR, "62c55f42d02be5ae409df87cde30f1d.jpg")
class SplashScreen(QWidget):
    def __init__(self, pixmap_path, duration=1800, fade_duration=800, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background: transparent;")
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        pixmap = QPixmap(pixmap_path)
        self.label.setPixmap(pixmap.scaled(520, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        layout = QVBoxLayout(self)
        layout.addStretch()
        layout.addWidget(self.label, alignment=Qt.AlignCenter)
        layout.addStretch()
        self.setFixedSize(520, 400)
        self.duration = duration
        self.fade_duration = fade_duration
        self._opacity_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self._opacity_effect)
        self._opacity_effect.setOpacity(1.0)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.start_fade)
        self._timer.start(self.duration)
        self._fade_timer = QTimer(self)
        self._fade_timer.timeout.connect(self._fade_step)
        self._fade_steps = 20
        self._fade_count = 0
        self._fade_opacity_step = 1.0 / self._fade_steps
        self._fade_interval = self.fade_duration // self._fade_steps
        self._on_finish = None

    def start_fade(self):
        self._fade_count = 0
        self._fade_timer.start(self._fade_interval)

    def _fade_step(self):
        opacity = max(0.0, 1.0 - self._fade_count * self._fade_opacity_step)
        self._opacity_effect.setOpacity(opacity)
        self._fade_count += 1
        if self._fade_count > self._fade_steps:
            self._fade_timer.stop()
            self.hide()
            if self._on_finish:
                self._on_finish()

    def set_on_finish(self, callback):
        self._on_finish = callback

class PortMonitorBar(QWidget):
    """独立的端口监听横栏 - 新层"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(200, 300)  # 纵向布局，适合显示5个栏目
        
        # 端口监听相关
        self._active_ports = []
        self._current_display_index = 0
        self._scan_interval = 1500  # 每1.5秒扫描一次端口（更快）
        self._display_columns = 5  # 显示5个栏目
        
        self.init_ui()
        self.position_window()
        
        # 定时器用于端口扫描
        self._scan_timer = QTimer(self)
        self._scan_timer.timeout.connect(self._scan_ports)
        self._scan_timer.start(self._scan_interval)
        
        # 定时器用于滚动显示端口信息（更快的流动速度）
        self._scroll_timer = QTimer(self)
        self._scroll_timer.timeout.connect(self._scroll_ports)
        self._scroll_timer.start(2000)  # 每2秒切换一次显示（更快）
        
        # 立即执行一次扫描
        self._scan_ports()
        
    def init_ui(self):
        """初始化界面"""
        layout = QVBoxLayout(self)  # 改为纵向布局
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        # 创建5个端口显示标签
        self.port_labels = []
        for i in range(self._display_columns):
            label = QLabel(f"端口{i+1}: 扫描中...")
            label.setFont(QFont(BENDER_FONT, 11, QFont.Bold))
            label.setStyleSheet(f"""
                color: #FFFFFF;
                background: rgba(35, 39, 46, 0.9);
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
                font-family: '{BENDER_FONT}', '{CHINESE_FONT}', Arial, sans-serif;
                min-width: 160px;
            """)
            label.setAlignment(Qt.AlignCenter)
            label.setFixedHeight(45)
            self.port_labels.append(label)
            layout.addWidget(label)
        
        # 设置整体样式
        self.setStyleSheet("""
            QWidget {
                background: transparent;
            }
        """)
        
    def position_window(self):
        """将窗口定位到屏幕右上角"""
        screen = QApplication.primaryScreen().geometry()
        x = screen.width() - self.width() - 20  # 距离右边缘20px
        y = 20  # 距离顶部20px
        self.move(x, y)
        
    def _scan_ports(self):
        """扫描活跃端口 - 优化版本"""
        try:
            listening_ports = set()
            
            # 方法1: 使用psutil获取网络连接（主要方法）
            try:
                connections = psutil.net_connections(kind='inet')
                for conn in connections:
                    if conn.status == 'LISTEN' and conn.laddr:
                        port = conn.laddr.port
                        if 1 <= port <= 65535:
                            listening_ports.add(port)
                print(f"psutil检测到 {len(listening_ports)} 个监听端口")
            except Exception as e:
                print(f"psutil方法失败: {e}")
            
            # 方法2: 使用netstat命令（补充方法）
            try:
                import subprocess
                import re
                
                if platform.system().lower() == "windows":
                    result = subprocess.run(['netstat', '-an'], capture_output=True, text=True, timeout=2)
                else:
                    result = subprocess.run(['netstat', '-tuln'], capture_output=True, text=True, timeout=2)
                
                if result.returncode == 0:
                    lines = result.stdout.split('\n')
                    netstat_ports = set()
                    for line in lines:
                        if 'LISTENING' in line or 'LISTEN' in line:
                            # 匹配端口号
                            match = re.search(r':(\d+)\s', line)
                            if match:
                                port = int(match.group(1))
                                if 1 <= port <= 65535:
                                    netstat_ports.add(port)
                    
                    # 合并结果
                    listening_ports.update(netstat_ports)
                    print(f"netstat补充检测到 {len(netstat_ports)} 个端口")
            except Exception as e:
                print(f"netstat方法失败: {e}")
            
            # 方法3: 主动扫描常用端口（验证方法）
            try:
                import socket
                # 扩展的常用端口列表
                common_ports = [
                    21, 22, 23, 25, 53, 80, 110, 143, 443, 993, 995, 3389,
                    5432, 3306, 6379, 27017, 8080, 8443, 3000, 5000, 8000, 9000,
                    1433, 1521, 5984, 9200, 9300, 11211, 2181, 9092, 5672, 15672
                ]
                
                active_ports = set()
                for port in common_ports:
                    if self._is_port_listening(port):
                        active_ports.add(port)
                
                # 合并结果
                listening_ports.update(active_ports)
                print(f"socket验证检测到 {len(active_ports)} 个活跃端口")
            except Exception as e:
                print(f"socket扫描失败: {e}")
            
            # 转换为列表并排序
            self._active_ports = sorted(list(listening_ports))
            
            # 限制显示数量，优先显示常用端口
            if len(self._active_ports) > 30:
                # 优先保留常用端口
                priority_ports = [21, 22, 23, 25, 53, 80, 110, 143, 443, 993, 995, 3389, 5432, 3306, 6379, 27017, 8080, 8443]
                priority_found = [p for p in self._active_ports if p in priority_ports]
                other_ports = [p for p in self._active_ports if p not in priority_ports]
                self._active_ports = priority_found + other_ports[:30-len(priority_found)]
            
            self._update_display()
                
        except Exception as e:
            print(f"端口扫描总错误: {e}")
            # 显示错误信息
            for i, label in enumerate(self.port_labels):
                label.setText(f"端口 {i+1}\n扫描错误")
    
    def _is_port_listening(self, port):
        """检测端口是否真正在监听 - 优化版本"""
        try:
            import socket
            
            # 尝试多种地址
            addresses = ['127.0.0.1', 'localhost', '0.0.0.0']
            
            for addr in addresses:
                try:
                    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    sock.settimeout(0.03)  # 更短的超时时间
                    result = sock.connect_ex((addr, port))
                    sock.close()
                    if result == 0:
                        return True
                except:
                    continue
            
            return False
        except:
            return False
    
    def _update_display(self):
        """更新显示内容"""
        if not self._active_ports:
            for i, label in enumerate(self.port_labels):
                label.setText(f"端口 {i+1}\n无活跃")
        else:
            # 显示当前索引开始的5个端口
            for i, label in enumerate(self.port_labels):
                port_index = (self._current_display_index + i) % len(self._active_ports)
                if port_index < len(self._active_ports):
                    current_port = self._active_ports[port_index]
                    port_info = self._get_port_info(current_port)
                    display_text = f"端口 {current_port}\n{port_info}"
                    
                    # 如果信息太长，截断显示
                    if len(port_info) > 12:
                        port_info = port_info[:9] + "..."
                        display_text = f"端口 {current_port}\n{port_info}"
                    
                    label.setText(display_text)
                else:
                    label.setText(f"端口 {i+1}\n无")
            
    def _get_port_info(self, port):
        """获取端口信息 - 扩展版本"""
        common_ports = {
            # 基础服务
            21: "FTP", 22: "SSH", 23: "Telnet", 25: "SMTP", 53: "DNS",
            # Web服务
            80: "HTTP", 443: "HTTPS", 8080: "HTTP-Alt", 8443: "HTTPS-Alt",
            # 邮件服务
            110: "POP3", 143: "IMAP", 993: "IMAPS", 995: "POP3S",
            # 数据库
            3306: "MySQL", 5432: "PostgreSQL", 27017: "MongoDB", 6379: "Redis",
            1433: "SQL Server", 1521: "Oracle", 5984: "CouchDB",
            # 搜索引擎
            9200: "Elasticsearch", 9300: "Elasticsearch",
            # 缓存
            11211: "Memcached",
            # 消息队列
            2181: "Zookeeper", 9092: "Kafka", 5672: "RabbitMQ", 15672: "RabbitMQ-Web",
            # 远程服务
            3389: "RDP", 5900: "VNC", 22: "SSH",
            # 开发服务
            3000: "Node.js", 5000: "Flask", 8000: "Django", 9000: "SonarQube"
        }
        return common_ports.get(port, "未知")
            
    def _scroll_ports(self):
        """滚动显示端口信息"""
        if len(self._active_ports) > self._display_columns:
            self._current_display_index = (self._current_display_index + 1) % len(self._active_ports)
            self._update_display()
            
    def mousePressEvent(self, event):
        """鼠标按下事件 - 手动刷新端口扫描"""
        if event.button() == Qt.LeftButton:
            self._scan_ports()
        super().mousePressEvent(event)
        
    def mouseDoubleClickEvent(self, event):
        """双击事件 - 重置显示索引"""
        self._current_display_index = 0
        self._update_display()
        super().mouseDoubleClickEvent(event)
    
    def close_monitor(self):
        """关闭端口监听栏"""
        self._scan_timer.stop()
        self._scroll_timer.stop()
        self.close()

class ArknightsMonitor(QWidget):
    def __init__(self):
        super().__init__()
        # 拖拽相关
        self._drag_active = False
        self._drag_pos = None
        # 端口监听栏引用
        self.port_monitor = None
        # 字体动态加载（如有本地ttf/otf）
        # QFontDatabase.addApplicationFont("C:/path/to/Bender.ttf")
        # QFontDatabase.addApplicationFont("C:/path/to/NovecentoWide.ttf")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet(f"""
            QWidget {{
                background: #23272E;
                border-radius: 0;
            }}
            QLabel {{
                color: #F8F8F8;
                font-family: '{NOVECENTO_FONT}', '{CHINESE_FONT}', Arial, sans-serif;
                background: transparent;
            }}
            QFrame#slant_card {{
                background: transparent;
                border-radius: 0;
                margin-bottom: 10px;
                border: none;
            }}
            QFrame#line {{
                background: #FFB400;
                max-height: 3px;
                min-height: 3px;
                border-radius: 0;
            }}
            QPushButton {{
                background: transparent;
                color: #FFB400;
                border: none;
                border-radius: 0;
                font-size: 26px;
                font-family: {NOVECENTO_FONT}, '{CHINESE_FONT}', Arial, sans-serif;
                padding: 0 10px;
            }}
            QPushButton:hover {{
                background: #FFB400;
                color: #23272E;
            }}
        """)
        self.last_net = psutil.net_io_counters()
        self.last_time = time.time()
        self.init_ui()
        psutil.cpu_percent(interval=0.1)  # 预热
        self.update_status()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_status)
        self.timer.start(1000)
        # 跑马灯相关
        self._marquee_text = ""
        self._marquee_pos = 0
        self._marquee_timer = QTimer(self)
        self._marquee_timer.timeout.connect(self._update_marquee)
        self._marquee_timer.start(120)  # 可调整速度

    def showEvent(self, event):
        """窗口显示事件"""
        super().showEvent(event)

    def hideEvent(self, event):
        """窗口隐藏事件"""
        super().hideEvent(event)

    def changeEvent(self, event):
        """窗口状态变化事件"""
        super().changeEvent(event)


    def _on_exit_clicked(self):
        """退出按钮点击处理"""
        # 关闭端口监听栏
        if hasattr(self, 'port_monitor') and self.port_monitor:
            self.port_monitor.close_monitor()
        # 关闭主界面
        self.close()

    # 拖拽功能
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = True
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            # 采集当前窗口信息
            geo = self.geometry()
            self._last_click_info = {
                'title': self.windowTitle() or "PRTS Monitor",
                'size': f"{geo.width()}x{geo.height()}",
                'pos': f"({geo.x()},{geo.y()})",
                'active': '是' if self.isActiveWindow() else '否'
            }
            # 记录鼠标按下操作（已移除独立操作栏）
            event.accept()
    def mouseMoveEvent(self, event):
        if self._drag_active and event.buttons() == Qt.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()
    def mouseReleaseEvent(self, event):
        self._drag_active = False
        # 记录鼠标释放操作（已移除独立操作栏）
        event.accept()

    def init_ui(self):
        # 极简透明布局
        self.setAttribute(Qt.WA_TranslucentBackground)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(14)
        # 顶部标题+网络图标
        top_bar = QHBoxLayout()
        logo = QLabel("PRTS")
        logo_font = QFont(NOVECENTO_FONT, 28, QFont.Bold)
        logo.setFont(logo_font)
        logo.setStyleSheet("color: #FFFFFF; background: transparent;")
        logo.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        top_bar.addWidget(logo)
        top_bar.addStretch()
        
        # 退出按钮
        exit_btn = QPushButton("×")
        exit_btn.setFont(QFont(NOVECENTO_FONT, 22, QFont.Bold))
        exit_btn.setStyleSheet(f"color: #FF4C4C; background: transparent; border: none; font-family: '{NOVECENTO_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
        exit_btn.setFixedSize(36, 36)
        exit_btn.clicked.connect(self._on_exit_clicked)
        top_bar.addWidget(exit_btn)
        # 动态设置网络图片高度为logo字体高度的100%
        font_metrics = QFontMetrics(logo_font)
        logo_height = font_metrics.height()
        net_img_h = max(32, int(logo_height * 1.0))
        self.net_label = QLabel()
        self.net_label.setPixmap(QPixmap(NET_OFF).scaled(net_img_h, net_img_h, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.net_label.setStyleSheet("background: transparent;")
        self._net_img_h = net_img_h  # 供后续动态缩放用
        top_bar.addWidget(self.net_label)
        main_layout.addLayout(top_bar)
        # 监控数据区块
        def card(label, value, obj_name):
            h = QHBoxLayout()
            h.setContentsMargins(0, 0, 0, 0)
            l1 = QLabel(label)
            l1.setFont(QFont(NOVECENTO_FONT, 16, QFont.Bold))
            l1.setStyleSheet(f"color: #FFFFFF; background: transparent; font-family: '{NOVECENTO_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
            l2 = QLabel(value)
            l2.setFont(QFont(BENDER_FONT, 22, QFont.Bold))
            l2.setStyleSheet(f"color: #FFFFFF; background: transparent; font-family: '{BENDER_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
            l2.setWordWrap(True)
            l2.setMinimumWidth(60)
            l2.setMaximumWidth(180)
            l2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            setattr(self, obj_name, l2)
            h.addWidget(l1)
            h.addStretch()
            h.addWidget(l2)
            return h
        main_layout.addLayout(card("CPU", "0%", "cpu_label"))
        main_layout.addLayout(card("GPU", "0%", "gpu_label"))
        main_layout.addLayout(card("MEM", "0%", "mem_label"))
        main_layout.addLayout(card("Disk", "0%", "disk_label"))
        main_layout.addLayout(card("Net", "0 KB/s", "net_speed_label"))
        main_layout.addLayout(card("IP", "0.0.0.0", "ip_label"))
        main_layout.addLayout(card("Uptime", "0h0m", "uptime_label"))
        main_layout.addStretch()
        # 信息采集栏
        self.info_bar = QLabel()
        self.info_bar.setFont(QFont(BENDER_FONT, 12))
        self.info_bar.setStyleSheet(f"color: #FFB400; background: #23272E; border-radius: 6px; padding: 6px 10px; font-family: '{BENDER_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
        self.info_bar.setWordWrap(False)  # 禁止换行
        self.info_bar.setText("")
        self.info_bar.setFixedHeight(2 * 28)  # 2行高度，28像素/行可根据字体微调
        main_layout.addWidget(self.info_bar)
        # 剪切板内容采集栏
        self.clipboard_bar = QLabel()
        self.clipboard_bar.setFont(QFont(BENDER_FONT, 12))
        self.clipboard_bar.setStyleSheet(f"color: #00CFFF; background: #23272E; border-radius: 6px; padding: 6px 10px; font-family: '{BENDER_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
        self.clipboard_bar.setWordWrap(True)
        self.clipboard_bar.setText("")
        main_layout.addWidget(self.clipboard_bar)

        # 新增：网页信息采集栏
        self.webinfo_bar = QLabel()
        self.webinfo_bar.setFont(QFont(BENDER_FONT, 12))
        self.webinfo_bar.setStyleSheet(f"color: #00FF88; background: #23272E; border-radius: 6px; padding: 6px 10px; font-family: '{BENDER_FONT}', '{CHINESE_FONT}', Arial, sans-serif;")
        self.webinfo_bar.setWordWrap(True)
        self.webinfo_bar.setText("")
        main_layout.addWidget(self.webinfo_bar)


    def get_hwinfo(self):
        info = {}
        uname = platform.uname()
        info["OS"] = platform.platform()
        # CPU型号
        try:
            import cpuinfo
            cpu_name = cpuinfo.get_cpu_info()['brand_raw']
        except Exception:
            cpu_name = uname.processor or os.environ.get('PROCESSOR_IDENTIFIER', '') or "Unknown"
        info["CPU Model"] = cpu_name
        info["CPU Cores"] = str(psutil.cpu_count(logical=False))
        info["CPU Threads"] = str(psutil.cpu_count(logical=True))
        info["Total Memory"] = f"{round(psutil.virtual_memory().total / (1024**3), 2)} GB"
        info["Machine Type"] = uname.machine
        info["Host Name"] = uname.node
        info["Python Version"] = platform.python_version()
        # 主板信息
        info["Mainboard"] = getattr(uname, 'version', 'Unknown')
        # 显卡信息
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                info["GPU"] = gpus[0].name
            else:
                info["GPU"] = "N/A"
        except Exception:
            info["GPU"] = "N/A"
        # 硬盘信息
        try:
            disks = psutil.disk_partitions()
            diskinfo = []
            for d in disks:
                try:
                    usage = psutil.disk_usage(d.mountpoint)
                    diskinfo.append(f"{d.device} {usage.total//(1024**3)}GB")
                except Exception:
                    continue
            info["Disk"] = ", ".join(diskinfo)
        except Exception:
            info["Disk"] = "N/A"
        return info

    def update_status(self):
        # CPU
        try:
            cpu = psutil.cpu_percent(interval=0)
            self.cpu_label.setText(f"{cpu:.1f}%")
        except Exception:
            self.cpu_label.setText("N/A")
        # 内存
        try:
            mem = psutil.virtual_memory().percent
            self.mem_label.setText(f"{mem:.1f}%")
        except Exception:
            self.mem_label.setText("N/A")
        # GPU
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu = gpus[0]
                gpu_load = getattr(gpu, 'load', None)
                gpu_load_str = f"{gpu_load * 100:.1f}%" if gpu_load is not None else "N/A"
                freq_str = ""
                clock = getattr(gpu, 'clock', None)
                if clock and isinstance(clock, (int, float)) and clock > 0:
                    freq_str = f" @ {clock:.0f}MHz"
                else:
                    try:
                        import subprocess
                        result = subprocess.check_output(
                            ["nvidia-smi", "--query-gpu=clocks.sm", "--format=csv,noheader,nounits"],
                            encoding="utf-8", stderr=subprocess.DEVNULL
                        )
                        freq_val = result.strip().split('\n')[0]
                        if freq_val.isdigit():
                            freq_str = f" @ {freq_val}MHz"
                    except Exception:
                        pass
                self.gpu_label.setText(f"{gpu_load_str}{freq_str}")
            else:
                self.gpu_label.setText("N/A")
        except Exception:
            self.gpu_label.setText("N/A")
        # 硬盘
        try:
            disk = psutil.disk_usage('/').percent
            self.disk_label.setText(f"{disk:.1f}%")
        except Exception:
            self.disk_label.setText("N/A")
        # 网络速度
        try:
            now_net = psutil.net_io_counters()
            now_time = time.time()
            sent = now_net.bytes_sent - getattr(self, 'last_net', now_net).bytes_sent
            recv = now_net.bytes_recv - getattr(self, 'last_net', now_net).bytes_recv
            duration = now_time - getattr(self, 'last_time', now_time)
            if duration > 0:
                up_speed = sent / duration / 1024
                down_speed = recv / duration / 1024
                self.net_speed_label.setText(f"↑{up_speed:.1f}KB/s ↓{down_speed:.1f}KB/s")
            self.last_net = now_net
            self.last_time = now_time
        except Exception:
            self.net_speed_label.setText("N/A")
        # IP
        try:
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            self.ip_label.setText(f"{ip}")
        except Exception:
            self.ip_label.setText("N/A")
        # Uptime
        try:
            uptime = int(time.time() - psutil.boot_time())
            hours = uptime // 3600
            minutes = (uptime % 3600) // 60
            self.uptime_label.setText(f"{hours}h{minutes}m")
        except Exception:
            self.uptime_label.setText("N/A")
        # 网络状态图标
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=1)
            self.net_label.setPixmap(QPixmap(NET_ON).scaled(self._net_img_h, self._net_img_h, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception:
            self.net_label.setPixmap(QPixmap(NET_OFF).scaled(self._net_img_h, self._net_img_h, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        # 信息采集栏
        # 复杂网络状态分析
        net_analysis = []
        # 本地IP
        try:
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            net_analysis.append(f"IP:{local_ip}")
        except Exception:
            net_analysis.append("IP:未知")
        # 默认网关
        try:
            gws = psutil.net_if_stats()
            # 只取第一个up的接口
            gw_name = next((k for k, v in gws.items() if v.isup), None)
            net_analysis.append(f"网卡:{gw_name if gw_name else '未知'}")
        except Exception:
            net_analysis.append("网卡:未知")
        # DNS可用性
        dns_status = []
        for dnsip, name in [("8.8.8.8", "DNS1"), ("114.114.114.114", "DNS2")]:
            try:
                socket.create_connection((dnsip, 53), timeout=1)
                dns_status.append(f"{name}:可用")
            except Exception:
                dns_status.append(f"{name}:异常")
        net_analysis.append(",".join(dns_status))
        # ping延迟
        try:
            import subprocess
            import platform as pf
            ping_host = "www.baidu.com"
            if pf.system().lower() == "windows":
                ping_cmd = ["ping", "-n", "1", "-w", "1000", ping_host]
            else:
                ping_cmd = ["ping", "-c", "1", "-W", "1", ping_host]
            result = subprocess.run(ping_cmd, capture_output=True, text=True)
            import re
            match = re.search(r"平均 = (\d+)ms|time[=<]([\d\.]+)ms", result.stdout)
            if match:
                delay = match.group(1) or match.group(2)
                net_analysis.append(f"延迟:{delay}ms")
            else:
                net_analysis.append("延迟:超时")
        except Exception:
            net_analysis.append("延迟:未知")
        # 网络类型
        try:
            nics = psutil.net_if_addrs()
            net_type = "未知"
            for nic in nics:
                if "wi-fi" in nic.lower() or "wlan" in nic.lower():
                    net_type = "无线"
                    break
                elif "eth" in nic.lower() or "以太网" in nic.lower():
                    net_type = "有线"
            net_analysis.append(f"类型:{net_type}")
        except Exception:
            net_analysis.append("类型:未知")
        net_status = "网络: " + ", ".join(net_analysis)
        # 仅USB等其他接口状态检测（不显示网卡）
        iface_lines = [net_status]
        try:
            for part in psutil.disk_partitions():
                # Windows下removable设备通常为U盘、移动硬盘
                if 'removable' in part.opts.lower() or part.fstype == '':
                    try:
                        usage = psutil.disk_usage(part.mountpoint)
                        iface_lines.append(f"USB[{part.device}]: {usage.total//(1024**3)}GB 已挂载")
                    except Exception:
                        iface_lines.append(f"USB[{part.device}]: 已挂载")
        except Exception:
            iface_lines.append("USB检测失败")
        # 只取前两栏内容
        info_str = " | ".join(iface_lines[:2])
        # 跑马灯内容更新逻辑
        if not hasattr(self, '_marquee_text') or self._marquee_text != info_str:
            self._marquee_text = info_str
            self._marquee_pos = 0
        # 只显示一部分，剩余部分由定时器滚动
        self._set_marquee_text()

        # 新增：网页信息采集栏内容
        try:
            webinfo = self._get_browser_active_title()
            self.webinfo_bar.setText(webinfo)
        except Exception:
            self.webinfo_bar.setText("网页信息采集失败")
    def _get_browser_active_title(self):
        """获取主流浏览器的活动窗口标题（仅支持Windows，需pywin32）"""
        try:
            import win32gui
            import win32process
        except ImportError:
            return "请安装pywin32以启用网页信息采集"
        browser_names = ["chrome.exe", "msedge.exe", "firefox.exe", "opera.exe", "iexplore.exe", "safari.exe"]
        def enum_windows_callback(hwnd, result):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                try:
                    tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                    p = psutil.Process(pid)
                    name = p.name().lower()
                    if name in browser_names:
                        result.append((name, win32gui.GetWindowText(hwnd)))
                except Exception:
                    pass
        result = []
        win32gui.EnumWindows(enum_windows_callback, result)
        if result:
            # 只取第一个浏览器窗口
            name, title = result[0]
            return f"当前网页窗口: {name} | {title}"
        else:
            return "未检测到浏览器活动窗口"
    def _set_marquee_text(self):
        # 跑马灯显示宽度（字符数，实际可根据label宽度动态调整）
        width = 48
        text = self._marquee_text
        # 固定宽度，超出部分滚动
        self.info_bar.setFixedWidth(width * 10)  # 10像素/字符，实际可微调
        if len(text) <= width:
            self.info_bar.setText(text)
        else:
            pos = self._marquee_pos % (len(text) + 8)  # 8为空格补偿，循环更自然
            show = (text + "        ") * 2
            self.info_bar.setText(show[pos:pos+width])

    def _update_marquee(self):
        if hasattr(self, '_marquee_text') and len(self._marquee_text) > 0:
            if len(self._marquee_text) > 48:
                self._marquee_pos += 1
                self._set_marquee_text()

        # 剪切板内容采集（仅内容变化时刷新）
        try:
            clipboard = QApplication.clipboard()
            clip_text = clipboard.text()
            if not hasattr(self, '_last_clipboard') or self._last_clipboard != clip_text:
                self._last_clipboard = clip_text
                if clip_text:
                    show_text = clip_text[:100] + ("..." if len(clip_text) > 100 else "")
                    self.clipboard_bar.setText(f"剪切板### {show_text}")
                else:
                    self.clipboard_bar.setText("剪切板### (空)")
        except Exception:
            self.clipboard_bar.setText("剪切板### 采集失败")

class SlantCard(QFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setMinimumHeight(44)
        self.setStyleSheet("""
            QFrame#slant_card {
                background: transparent;
                border-radius: 0;
                margin-bottom: 10px;
                border: none;
            }
        """)

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        # 左上斜切角
        painter.setBrush(QBrush(QColor('#FFB400')))
        painter.setPen(Qt.NoPen)
        points = [
            self.rect().topLeft(),
            self.rect().topLeft() + QPoint(30, 0),
            self.rect().topLeft() + QPoint(0, 12)
        ]
        painter.drawPolygon(QPolygon(points))
        # 右下斜切角
        painter.setBrush(QBrush(QColor('#00CFFF')))
        points = [
            self.rect().bottomRight(),
            self.rect().bottomRight() + QPoint(-30, 0),
            self.rect().bottomRight() + QPoint(0, -12)
        ]
        painter.drawPolygon(QPolygon(points))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    splash = SplashScreen(SPLASH_IMG, duration=1800, fade_duration=800)
    window = ArknightsMonitor()
    window.setMinimumSize(400, 540)
    window.resize(520, 660)
    window.setWindowOpacity(0.0)
    
    # 创建独立的端口监听栏
    port_monitor = PortMonitorBar()
    window.port_monitor = port_monitor  # 设置引用
    
    def show_main():
        window.show()
        port_monitor.show()  # 显示端口监听栏
        # 渐变显示主界面
        effect = QGraphicsOpacityEffect(window)
        window.setGraphicsEffect(effect)
        effect.setOpacity(0.0)
        window.setWindowOpacity(1.0)
        fade_steps = 20
        fade_interval = 30
        def fade_in_step(step=[0]):
            op = min(1.0, step[0] / fade_steps)
            effect.setOpacity(op)
            if op < 1.0:
                QTimer.singleShot(fade_interval, lambda: fade_in_step([step[0]+1]))
            else:
                window.setGraphicsEffect(None)
        fade_in_step()
    splash.set_on_finish(show_main)
    splash.show()
    sys.exit(app.exec())
