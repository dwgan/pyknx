import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import csv
from datetime import datetime
import threading
import queue
import time


class NFCReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NFC卡号读取器")
        self.root.geometry("900x550")

        self.serial_port = None
        self.csv_file = None
        self.csv_writer = None
        self.data_queue = queue.Queue()
        self.current_permission = 0  # 默认权限等级为0
        self.running = True  # 控制读取线程的运行状态
        self.seen_card_ids = set()  # 用于存储已见过的卡号

        # 创建UI组件
        self.create_widgets()

        # 每100ms检查一次数据队列
        self.root.after(100, self.process_queue)

    def create_widgets(self):
        # 串口配置面板
        config_frame = ttk.LabelFrame(self.root, text="串口配置")
        config_frame.pack(fill="x", padx=10, pady=5, ipadx=5, ipady=5)

        ttk.Label(config_frame, text="串口号:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.port_combobox = ttk.Combobox(config_frame, width=15)
        self.port_combobox.grid(row=0, column=1, padx=5, pady=5)
        self.refresh_ports()

        ttk.Label(config_frame, text="波特率:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.baud_entry = ttk.Entry(config_frame, width=10)
        self.baud_entry.grid(row=0, column=3, padx=5, pady=5)
        self.baud_entry.insert(0, "115200")  # 默认波特率

        self.connect_btn = ttk.Button(config_frame, text="连接", command=self.toggle_connection)
        self.connect_btn.grid(row=0, column=4, padx=5, pady=5)

        # 刷新串口按钮
        self.refresh_btn = ttk.Button(config_frame, text="刷新", command=self.refresh_ports)
        self.refresh_btn.grid(row=0, column=5, padx=5, pady=5)

        # 权限控制面板
        perm_frame = ttk.LabelFrame(self.root, text="权限设置")
        perm_frame.pack(fill="x", padx=10, pady=5, ipadx=5, ipady=5)

        ttk.Label(perm_frame, text="当前权限等级:").pack(side="left", padx=5, pady=5)

        self.perm_var = tk.StringVar(value=f"{self.current_permission} (普通)")
        ttk.Label(perm_frame, textvariable=self.perm_var, font=("Arial", 10, "bold")).pack(side="left", padx=5, pady=5)

        self.toggle_perm_btn = ttk.Button(perm_frame, text="切换为高级权限",
                                          command=self.toggle_permission,
                                          state="disabled")
        self.toggle_perm_btn.pack(side="left", padx=20, pady=5)

        # 数据记录控制
        record_frame = ttk.LabelFrame(self.root, text="数据记录")
        record_frame.pack(fill="x", padx=10, pady=5, ipadx=5, ipady=5)

        ttk.Label(record_frame, text="文件名:").pack(side="left", padx=5, pady=5)
        self.filename_entry = ttk.Entry(record_frame, width=30)
        self.filename_entry.pack(side="left", padx=5, pady=5)
        self.filename_entry.insert(0, "nfc_data.csv")  # 默认文件名

        self.record_btn = ttk.Button(record_frame, text="开始记录",
                                     command=self.toggle_recording, state="disabled")
        self.record_btn.pack(side="left", padx=5, pady=5)

        # 导出格式提示
        export_frame = ttk.Frame(record_frame)
        export_frame.pack(side="left", padx=20)
        ttk.Label(export_frame, text="导出格式:").pack(side="left")
        self.encoding_var = tk.StringVar(value="GBK")
        encodings = ["GBK", "UTF-8 with BOM", "UTF-8"]
        self.encoding_combobox = ttk.Combobox(export_frame, width=15, textvariable=self.encoding_var)
        self.encoding_combobox['values'] = encodings
        self.encoding_combobox.pack(side="left", padx=5)

        # 状态显示
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom", padx=0, pady=0)

        # 数据显示表格 (根据新要求调整列顺序)
        data_frame = ttk.LabelFrame(self.root, text="读取数据")
        data_frame.pack(fill="both", expand=True, padx=10, pady=5, ipadx=5, ipady=5)

        # 新列顺序: time, employee_name, employee_id, card_id, permission
        columns = ("time", "employee_name", "employee_id", "card_id", "permission")
        self.tree = ttk.Treeview(data_frame, columns=columns, show="headings")

        # 设置表头和列宽
        self.tree.heading("time", text="时间")
        self.tree.heading("employee_name", text="员工姓名")
        self.tree.heading("employee_id", text="员工号")
        self.tree.heading("card_id", text="卡号")
        self.tree.heading("permission", text="权限等级")

        self.tree.column("time", width=200)
        self.tree.column("employee_name", width=120)
        self.tree.column("employee_id", width=100)
        self.tree.column("card_id", width=120)
        self.tree.column("permission", width=100)

        # 添加垂直和水平滚动条
        vsb = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # 表格布局
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # 配置网格行列权重
        data_frame.grid_rowconfigure(0, weight=1)
        data_frame.grid_columnconfigure(0, weight=1)

    def toggle_permission(self):
        """切换权限等级"""
        self.current_permission = 1 - self.current_permission  # 在0和1之间切换

        # 更新权限显示
        perm_text = f"{self.current_permission} "
        perm_text += "(高级)" if self.current_permission == 1 else "(普通)"
        self.perm_var.set(perm_text)

        # 更新按钮文本
        btn_text = "切换为普通权限" if self.current_permission == 1 else "切换为高级权限"
        self.toggle_perm_btn.config(text=btn_text)

        self.status_var.set(f"权限已切换为: {perm_text}")

    def refresh_ports(self):
        """刷新可用的串口列表"""
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combobox["values"] = ports
        if ports:
            self.port_combobox.current(0)

    def toggle_connection(self):
        """切换串口连接状态"""
        if self.serial_port and self.serial_port.is_open:
            self.close_serial()
            self.connect_btn.config(text="连接")
            self.record_btn.config(state="disabled")
            self.toggle_perm_btn.config(state="disabled")
        else:
            self.open_serial()
            self.toggle_perm_btn.config(state="normal")

    def open_serial(self):
        """打开串口连接"""
        port = self.port_combobox.get()
        baud_rate = self.baud_entry.get()

        if not port:
            messagebox.showerror("错误", "请选择串口号")
            return

        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=int(baud_rate),
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            self.status_var.set(f"已连接 {port}@{baud_rate}")
            self.connect_btn.config(text="断开")
            self.record_btn.config(state="normal")

            # 启动读取线程
            self.running = True
            self.serial_thread = threading.Thread(target=self.read_serial, daemon=True)
            self.serial_thread.start()
        except Exception as e:
            messagebox.showerror("连接错误", str(e))

    def close_serial(self):
        """关闭串口连接"""
        if self.serial_port and self.serial_port.is_open:
            self.running = False  # 通知线程停止
            if hasattr(self, 'serial_thread') and self.serial_thread.is_alive():
                self.serial_thread.join(timeout=1.0)  # 等待线程结束
            self.serial_port.close()
            self.status_var.set("连接已断开")
            self.record_btn.config(state="disabled")
            self.toggle_perm_btn.config(state="disabled")
            if self.csv_file:
                self.stop_recording()
            # 清空已见过的卡号集合
            self.seen_card_ids.clear()

    def toggle_recording(self):
        """切换数据记录状态"""
        if self.csv_file:
            self.stop_recording()
            self.record_btn.config(text="开始记录")
        else:
            self.start_recording()
            self.record_btn.config(text="停止记录")

    def start_recording(self):
        """开始记录到CSV文件"""
        filename = self.filename_entry.get()
        if not filename:
            messagebox.showerror("错误", "请输入文件名")
            return

        try:
            # 根据选择的编码格式创建文件
            encoding = self.encoding_var.get()
            file_encoding = "GBK" if encoding == "GBK" else "utf-8"

            # 检查是否需要添加BOM标记
            if encoding == "UTF-8 with BOM":
                file_encoding = "utf-8-sig"

            self.csv_file = open(filename, "a", newline="", encoding=file_encoding)
            self.csv_writer = csv.writer(self.csv_file)

            # 如果文件为空，写入标题（按新顺序）
            if self.csv_file.tell() == 0:
                self.csv_writer.writerow(["时间戳", "员工姓名", "员工号", "卡号", "权限等级"])

            self.status_var.set(f"正在记录到: {filename} ({encoding}编码)")
        except Exception as e:
            messagebox.showerror("文件错误", str(e))
            self.csv_file = None
            self.csv_writer = None

    def stop_recording(self):
        """停止记录并关闭文件"""
        if self.csv_file:
            try:
                self.csv_file.close()
                self.status_var.set("记录已停止")
            except Exception as e:
                messagebox.showerror("错误", f"关闭文件时出错: {str(e)}")
            finally:
                self.csv_file = None
                self.csv_writer = None

    def read_serial(self):
        """从串口读取数据的线程函数"""
        buffer = bytearray()
        try:
            while self.running and self.serial_port and self.serial_port.is_open:
                # 读取串口数据
                if self.serial_port.in_waiting > 0:
                    data = self.serial_port.read(self.serial_port.in_waiting)
                    buffer.extend(data)
                    
                    # 检查是否有完整的数据包（4字节卡号 + \r\n）
                    while b'\r\n' in buffer:
                        # 找到第一个回车换行符的位置
                        end_index = buffer.index(b'\r\n')
                        
                        # 提取卡号数据（应该是4字节）
                        card_bytes = buffer[:end_index]
                        
                        # 移除已处理的数据
                        buffer = buffer[end_index+2:]
                        
                        # 检查卡号长度
                        if len(card_bytes) == 4:
                            # 将4字节转换为16进制字符串
                            card_id = "".join(f"{b:02X}" for b in card_bytes)

                            # 获取当前时间
                            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M")

                            # 通过队列安全地传递数据给主线程
                            self.data_queue.put((timestamp, "", "", card_id, self.current_permission))
                        else:
                            # 长度不是4字节，可能是错误数据
                            self.data_queue.put(("ERROR", "", "", f"无效数据长度: {len(card_bytes)}字节", ""))
                
                # 短暂休眠，避免过度占用CPU
                time.sleep(0.01)
                
        except serial.SerialException as e:
            if self.serial_port and self.serial_port.is_open:
                self.data_queue.put(("ERROR", "", "", f"串口错误: {str(e)}", ""))
            self.close_serial()
        except Exception as e:
            self.data_queue.put(("ERROR", "", "", f"未知错误: {str(e)}", ""))

    def process_queue(self):
        """处理从串口线程接收到的数据"""
        try:
            while True:
                data = self.data_queue.get_nowait()
                if data[0] == "ERROR":
                    messagebox.showerror("错误", data[3])
                else:
                    timestamp, employee_name, employee_id, card_id, permission = data
                    
                    # 检查卡号是否重复
                    if card_id in self.seen_card_ids:
                        # 卡号重复，显示错误信息
                        messagebox.showerror("重复卡号", f"卡号 {card_id} 已存在，未添加到列表中")
                        self.status_var.set(f"检测到重复卡号: {card_id}")
                    else:
                        # 卡号未重复，添加到已见集合
                        self.seen_card_ids.add(card_id)
                        
                        # 添加权限文本描述
                        perm_text = f"{permission} ({'高级' if permission == 1 else '普通'})"

                        # 添加到表格显示（使用新列顺序）
                        self.tree.insert("", "end", values=(timestamp, employee_name, employee_id, card_id, perm_text))

                        # 滚动到底部
                        self.tree.yview_moveto(1)

                        # 如果正在记录，写入CSV文件（权限只存储数字，不存储文本描述）
                        if self.csv_writer:
                            self.csv_writer.writerow([timestamp, employee_name, employee_id, card_id, permission])
                        
                        # 更新状态栏
                        self.status_var.set(f"已添加卡号: {card_id}")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def on_closing(self):
        """窗口关闭时的清理操作"""
        self.running = False  # 通知线程停止
        if self.serial_port and self.serial_port.is_open:
            self.close_serial()
        if self.csv_file:
            self.stop_recording()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = NFCReaderApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
