import asyncio
import socket
import threading
import tkinter as tk
from tkinter import ttk
from xknx import XKNX
from xknx.io import ConnectionConfig, GatewayScanner, ConnectionType  # 添加ConnectionType导入
from xknx.dpt import DPTBinary
from xknx.telegram import GroupAddress, Telegram
from xknx.telegram.apci import GroupValueWrite
import logging
import time
import re
import netifaces

# 配置日志记录
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger("KNXController")


class KNXControllerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KNX控制器")
        self.root.geometry("600x550")
        self.root.resizable(True, True)

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 获取本地IP地址
        self.local_ips = self.get_local_ips()

        # 创建UI
        self.create_ui()

        # 初始化变量
        self.gateways = []
        self.selected_gateway = None
        self.selected_local_ip = self.local_ips[0] if self.local_ips else None
        self.scan_running = False
        self.scan_progress = 0
        self.scan_start_time = 0

    def get_local_ips(self):
        """获取所有本地IP地址"""
        ips = []
        try:
            # 使用netifaces获取所有网络接口信息
            interfaces = netifaces.interfaces()
            for interface in interfaces:
                ifaddresses = netifaces.ifaddresses(interface)
                if netifaces.AF_INET in ifaddresses:
                    for link in ifaddresses[netifaces.AF_INET]:
                        ip = link['addr']
                        if ip != "127.0.0.1":  # 排除回环地址
                            ips.append(ip)
        except Exception as e:
            logger.error(f"获取IP地址失败: {e}")
            ips = ["192.168.0.24"]  # 默认值
        return ips

    def create_ui(self):
        """创建用户界面"""
        # 网络适配器选择
        adapter_frame = ttk.LabelFrame(self.main_frame, text="网络适配器", padding="10")
        adapter_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(adapter_frame, text="选择本地IP:").pack(side=tk.LEFT, padx=(0, 10))

        self.ip_var = tk.StringVar()
        self.ip_combo = ttk.Combobox(
            adapter_frame,
            textvariable=self.ip_var,
            values=self.local_ips,
            state="readonly",
            width=20
        )
        self.ip_combo.pack(side=tk.LEFT, padx=(0, 20))
        if self.local_ips:
            self.ip_combo.current(0)
        self.ip_combo.bind("<<ComboboxSelected>>", self.on_ip_selected)

        # 扫描按钮
        self.scan_button = ttk.Button(
            adapter_frame,
            text="扫描KNX路由器",
            command=self.start_scan,
            width=15
        )
        self.scan_button.pack(side=tk.LEFT)

        # 路由器选择
        gateway_frame = ttk.LabelFrame(self.main_frame, text="KNX路由器", padding="10")
        gateway_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(gateway_frame, text="选择路由器:").pack(side=tk.LEFT, padx=(0, 10))

        self.gateway_var = tk.StringVar()
        self.gateway_combo = ttk.Combobox(
            gateway_frame,
            textvariable=self.gateway_var,
            state="readonly",
            width=40
        )
        self.gateway_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.gateway_combo.bind("<<ComboboxSelected>>", self.on_gateway_selected)

        # 手动输入框
        manual_frame = ttk.Frame(gateway_frame)
        manual_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        ttk.Label(manual_frame, text="或手动输入:").pack(side=tk.LEFT, padx=(0, 5))

        self.manual_ip_var = tk.StringVar()
        self.manual_ip_entry = ttk.Entry(manual_frame, textvariable=self.manual_ip_var, width=15)
        self.manual_ip_entry.pack(side=tk.LEFT, padx=(0, 5))
        self.manual_ip_entry.insert(0, "192.168.0.11")  # 默认路由器IP
        # 添加文本变化监听
        self.manual_ip_var.trace_add("write", self.validate_manual_input)

        ttk.Label(manual_frame, text=":").pack(side=tk.LEFT)

        self.manual_port_var = tk.StringVar()
        self.manual_port_entry = ttk.Entry(manual_frame, textvariable=self.manual_port_var, width=5)
        self.manual_port_entry.pack(side=tk.LEFT)
        self.manual_port_entry.insert(0, "3671")  # 默认端口
        # 添加文本变化监听
        self.manual_port_var.trace_add("write", self.validate_manual_input)

        # 进度条区域
        progress_frame = ttk.Frame(self.main_frame)
        progress_frame.pack(fill=tk.X, pady=(10, 15))

        self.progress_label = ttk.Label(progress_frame, text="扫描进度:")
        self.progress_label.pack(side=tk.LEFT, padx=(0, 10))

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, expand=True)

        self.progress_value = ttk.Label(progress_frame, text="0%")
        self.progress_value.pack(side=tk.RIGHT, padx=(10, 0))

        # 默认隐藏进度条
        self.progress_label.pack_forget()
        self.progress_bar.pack_forget()
        self.progress_value.pack_forget()

        # 命令发送区域
        command_frame = ttk.LabelFrame(self.main_frame, text="发送命令", padding="10")
        command_frame.pack(fill=tk.BOTH, expand=True)

        # 组地址输入
        group_frame = ttk.Frame(command_frame)
        group_frame.pack(fill=tk.X, pady=5)

        ttk.Label(group_frame, text="组地址:").pack(side=tk.LEFT, padx=(0, 10))

        self.group_var = tk.StringVar()
        self.group_entry = ttk.Entry(group_frame, textvariable=self.group_var, width=20)
        self.group_entry.pack(side=tk.LEFT, padx=(0, 20))
        self.group_entry.insert(0, "0/2/7")  # 默认组地址

        # 值输入
        value_frame = ttk.Frame(command_frame)
        value_frame.pack(fill=tk.X, pady=5)

        ttk.Label(value_frame, text="值:").pack(side=tk.LEFT, padx=(0, 10))

        self.value_var = tk.StringVar()
        self.value_entry = ttk.Entry(value_frame, textvariable=self.value_var, width=10)
        self.value_entry.pack(side=tk.LEFT, padx=(0, 20))
        self.value_entry.insert(0, "1")  # 默认值

        # 发送按钮
        self.send_button = ttk.Button(
            command_frame,
            text="发送命令",
            command=self.send_command,
            state=tk.DISABLED
        )
        self.send_button.pack(pady=10)

        # 日志区域
        log_frame = ttk.Frame(command_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        ttk.Label(log_frame, text="操作日志:").pack(anchor=tk.W)

        self.log_text = tk.Text(log_frame, height=6, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 初始化后验证一次手动输入
        self.validate_manual_input()

    def validate_manual_input(self, *args):
        """验证手动输入的路由器IP和端口是否有效"""
        ip = self.manual_ip_var.get().strip()
        port = self.manual_port_var.get().strip()

        # 验证IP地址格式
        ip_valid = False
        if re.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$', ip):
            parts = ip.split('.')
            if all(0 <= int(part) <= 255 for part in parts):
                ip_valid = True

        # 验证端口格式
        port_valid = False
        if port.isdigit():
            port_num = int(port)
            if 1 <= port_num <= 65535:
                port_valid = True

        # 如果输入有效，启用发送按钮
        if ip_valid and port_valid:
            self.send_button.config(state=tk.NORMAL)
            self.log_message("手动输入有效，可以发送命令")
        else:
            self.send_button.config(state=tk.DISABLED)
            self.log_message("请检查手动输入的路由器IP和端口")

    def log_message(self, message):
        """向日志区域添加消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 滚动到底部
        self.log_text.config(state=tk.DISABLED)

    def on_ip_selected(self, event=None):
        """当选择本地IP时"""
        self.selected_local_ip = self.ip_var.get()
        self.log_message(f"已选择本地IP: {self.selected_local_ip}")
        self.gateway_combo.set('')  # 清空路由器选择
        self.send_button.config(state=tk.DISABLED)

    def start_scan(self):
        """启动扫描线程"""
        if not self.selected_local_ip:
            self.log_message("请先选择本地IP地址")
            return

        if self.scan_running:
            self.log_message("扫描正在进行中，请稍候...")
            return

        self.log_message(f"开始扫描网络中的KNX路由器(使用{self.selected_local_ip})...")
        self.scan_button.config(state=tk.DISABLED)
        self.scan_running = True

        # 显示进度条
        self.progress_label.pack(side=tk.LEFT, padx=(0, 10))
        self.progress_bar.pack(fill=tk.X, expand=True)
        self.progress_value.pack(side=tk.RIGHT, padx=(10, 0))

        # 重置进度
        self.scan_progress = 0
        self.scan_start_time = time.time()
        self.progress_var.set(0)
        self.progress_value.config(text="0%")

        # 启动进度更新
        self.update_progress()

        # 在新线程中运行扫描
        threading.Thread(
            target=self.scan_network,
            args=(self.selected_local_ip,),  # 传递选择的本地IP
            daemon=True
        ).start()

    def update_progress(self):
        """更新扫描进度条"""
        if not self.scan_running:
            return

        # 计算已用时间
        elapsed = time.time() - self.scan_start_time

        # 计算进度 (假设扫描总时间为15秒)
        self.scan_progress = min(100, int(elapsed / 15 * 100))

        # 更新UI
        self.progress_var.set(self.scan_progress)
        self.progress_value.config(text=f"{self.scan_progress}%")

        # 每100毫秒更新一次
        if self.scan_progress < 100:
            self.root.after(100, self.update_progress)

    def scan_network(self, local_ip):
        """扫描网络中的KNX路由器"""

        async def scan():
            try:
                # 创建连接配置，强制使用指定的本地IP
                connection_config = ConnectionConfig(
                    local_ip=local_ip,
                    gateway_ip=None,  # 扫描时不需要指定网关IP
                    gateway_port=None,
                    auto_reconnect=False,
                    auto_reconnect_wait=3,
                    connection_type=ConnectionType.TUNNELING,  # 使用ConnectionType枚举
                )

                # 创建XKNX实例，使用指定的连接配置
                async with XKNX(connection_config=connection_config) as xknx:
                    # 创建网关扫描器
                    gateway_scanner = GatewayScanner(xknx)

                    # 强制使用选择的接口
                    gateway_scanner.local_ip = local_ip

                    # 禁用自动接口选择
                    gateway_scanner.bind_to_multicast = False

                    # 设置组播地址
                    gateway_scanner.multicast_group = "224.0.23.12"

                    # 开始扫描
                    self.log_message("正在扫描... (这可能需要10-15秒)")

                    # 根据版本选择正确的方法
                    if hasattr(gateway_scanner, 'scan'):
                        await gateway_scanner.scan()
                    elif hasattr(gateway_scanner, 'start'):
                        await gateway_scanner.start()
                    else:
                        self.log_message("错误: 不支持的扫描方法")
                        return

                    # 等待扫描结果 (增加扫描时间)
                    await asyncio.sleep(15)

                    # 获取扫描结果
                    found_gateways = gateway_scanner.found_gateways

                    # 修复错误：正确处理HPAI对象
                    self.gateways = []
                    for gw in found_gateways:
                        # 检查是否是GatewayDescriptor对象
                        if hasattr(gw, 'name'):
                            # GatewayDescriptor对象有name属性
                            gateway_info = {
                                "name": gw.name,
                                "ip": gw.ip_addr,
                                "port": gw.port
                            }
                        else:
                            # HPAI对象没有name属性
                            gateway_info = {
                                "name": f"路由器 {gw.ip_addr}:{gw.port}",
                                "ip": gw.ip_addr,
                                "port": gw.port
                            }
                        self.gateways.append(gateway_info)

                    # 更新UI
                    self.root.after(0, self.update_gateway_list)

                    self.log_message(f"扫描完成，找到 {len(self.gateways)} 个路由器")
            except Exception as e:
                # 修复错误：使用默认参数传递异常对象
                self.root.after(0, lambda e=e: self.log_message(f"扫描错误: {str(e)}"))
            finally:
                self.root.after(0, lambda: self.scan_button.config(state=tk.NORMAL))
                self.scan_running = False
                # 隐藏进度条
                self.root.after(0, self.hide_progress_bar)

        # 创建新的事件循环
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(scan())
        loop.close()

    def hide_progress_bar(self):
        """隐藏进度条"""
        self.progress_label.pack_forget()
        self.progress_bar.pack_forget()
        self.progress_value.pack_forget()

    def update_gateway_list(self):
        """更新路由器下拉列表"""
        gateway_names = [f"{gw['name']} ({gw['ip']}:{gw['port']})" for gw in self.gateways]
        self.gateway_combo.config(values=gateway_names)

        if gateway_names:
            self.gateway_combo.current(0)
            self.on_gateway_selected()
            self.log_message("请从下拉列表中选择一个路由器")
        else:
            self.log_message("未找到任何KNX路由器")
            # 允许手动输入
            self.send_button.config(state=tk.NORMAL)
            self.selected_gateway = {
                "ip": self.manual_ip_var.get(),
                "port": int(self.manual_port_var.get())
            }
            self.log_message("已使用手动输入的IP和端口")

    def on_gateway_selected(self, event=None):
        """当选择路由器时"""
        selected_index = self.gateway_combo.current()
        if selected_index >= 0 and selected_index < len(self.gateways):
            self.selected_gateway = self.gateways[selected_index]
            self.log_message(f"已选择路由器: {self.selected_gateway['name']}")
            self.send_button.config(state=tk.NORMAL)
        else:
            self.selected_gateway = None
            self.send_button.config(state=tk.DISABLED)

    def send_command(self):
        """发送KNX命令"""
        if not self.selected_local_ip:
            self.log_message("错误: 请先选择本地IP地址")
            return

        # 如果没有扫描到路由器，使用手动输入的值
        if not self.selected_gateway:
            try:
                # 验证IP地址格式
                ip = self.manual_ip_var.get().strip()
                if not re.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$', ip):
                    self.log_message("错误: 请输入有效的IP地址")
                    return

                # 验证端口
                port_str = self.manual_port_var.get().strip()
                if not port_str.isdigit():
                    self.log_message("错误: 端口必须是数字")
                    return

                port = int(port_str)
                if port < 1 or port > 65535:
                    self.log_message("错误: 端口必须在1-65535范围内")
                    return

                self.selected_gateway = {
                    "ip": ip,
                    "port": port
                }
                self.log_message(f"使用手动输入的路由器: {ip}:{port}")
            except Exception as e:
                self.log_message(f"错误: {str(e)}")
                return

        group_address = self.group_var.get().strip()
        value_str = self.value_var.get().strip()

        if not group_address:
            self.log_message("错误: 请输入组地址")
            return

        try:
            # 转换值
            value = int(value_str)
        except ValueError:
            self.log_message("错误: 值必须是整数")
            return

        # 在新线程中发送命令
        threading.Thread(
            target=self.send_knx_command,
            args=(group_address, value),
            daemon=True
        ).start()

    def send_knx_command(self, group_address, value):
        """实际发送KNX命令"""

        async def send():
            try:
                # 配置连接参数
                connection_config = ConnectionConfig(
                    local_ip=self.selected_local_ip,
                    gateway_ip=self.selected_gateway["ip"],
                    gateway_port=self.selected_gateway["port"],
                )

                self.log_message(f"正在连接到 {self.selected_gateway['ip']}:{self.selected_gateway['port']}...")

                # 创建XKNX实例
                async with XKNX(connection_config=connection_config) as xknx:
                    self.log_message("连接成功")

                    # 创建目标组地址
                    ga = GroupAddress(group_address)

                    # 创建有效载荷
                    payload = GroupValueWrite(value=DPTBinary(value))

                    # 创建并发送Telegram
                    telegram = Telegram(
                        destination_address=ga,
                        payload=payload
                    )
                    await xknx.telegrams.put(telegram)
                    self.log_message(f"命令已发送到 {group_address}: 值={value}")

                    # 等待命令完成
                    await asyncio.sleep(1)

                self.log_message("已断开连接")

            except Exception as e:
                self.log_message(f"错误: {str(e)}")

        # 创建新的事件循环
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(send())
        loop.close()


if __name__ == "__main__":
    root = tk.Tk()
    app = KNXControllerApp(root)
    root.mainloop()