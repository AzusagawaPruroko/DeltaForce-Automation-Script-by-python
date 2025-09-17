#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
自动点击程序
使用大漠插件实现自动化操作

作者: P1nKM41D
版本: 1.0.0
"""

import multiprocessing
import os
import sys
import ctypes
import time
import keyboard
import threading
import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox
from ttkthemes import ThemedTk
from win32com.client import Dispatch
import base64
import hashlib
import secrets
import requests
import json
from datetime import datetime
from pathlib import Path
import ntplib
import pyautogui
import random

# WinAPI相关导入
import win32api
import win32con

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import padding
import wmi

# 设置Tcl/Tk路径
if getattr(sys, 'frozen', False):
    # 如果是打包后的exe
    application_path = os.path.dirname(sys.executable)
    os.environ['TCL_LIBRARY'] = os.path.join(application_path, 'tcl', 'tcl8.6')
    os.environ['TK_LIBRARY'] = os.path.join(application_path, 'tcl', 'tk8.6')

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 导入UAC提权模块
from uac_power import run_as_admin

def is_admin():
    """检查是否具有管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def 注册大漠_简(注册码='', 附加码=''):
    """大漠插件免注册实现"""
    print('正在初始化')
    # 通过调用DmReg.dll注册大漠 这样不会把dm.dll写到系统中，从而实现免注册
    patch = ctypes.windll.LoadLibrary(os.path.dirname(__file__) + '/DmReg.dll')
    patch.SetDllPathW(os.path.dirname(__file__) + '/dm.dll', 0)
    dm_主对象 = Dispatch('dm.dmsoft')  # 创建对象1
    dm_主对象2 = Dispatch('dm.dmsoft')  # 创建对象2
    ver = dm_主对象.ver()
    print('免注册调用初始化成功 版本号为:', ver)

    # 注册大漠VIP
    if ver != '':
        reg1 = dm_主对象.reg(注册码, 附加码)
        reg2 = dm_主对象2.reg(注册码, 附加码)
        if reg1 == 1 and reg2 == 1:
            print("大漠vip注册成功")
            return dm_主对象, dm_主对象2
        else:
            print(f"大漠注册失败,错误代码: 主对象={reg1}, 副对象={reg2}")
            return None, None

class 满仓助手:
    def __init__(self, root):
        self.root = root
        self.root.attributes('-topmost', True)
        self.root.update()

        # 添加窗口焦点事件处理
        self.root.bind('<FocusOut>', self.on_focus_out)
        self.root.bind('<FocusIn>', self.on_focus_in)

        # 初始化运行状态
        self.running = False
        self.paused = False
        self.current_equipment = 1  # 当前配装位置
        self.stable_delay = 100  # 稳定延迟默认值(ms)
        self.restart_delay = 500  # 重启延迟默认值(ms)
        self.window_hidden = False  # 添加窗口隐藏状态标记
        self.配装价格识别错误计数 = 0  # 添加配装价格识别错误计数器
        self.恢复操作计数 = 0  # 新增：恢复操作计数器
        self.dangshi_var = tk.BooleanVar(value=True)
        
        # 错误库相关
        self.error_prices = set()  # 存储错误价格
        self.error_library_loaded = False  # 错误库是否已加载
        self.error_library_path = ""  # 错误库文件路径
        
        # 测试模式相关
        self.test_mode = False  # 测试模式状态

        # 设置全局样式
        self.style = ttk.Style()
        self.style.configure("TLabel", font=('微软雅黑', 9))
        self.style.configure("TButton", font=('微软雅黑', 9))
        self.style.configure("TLabelframe.Label", font=('微软雅黑', 9))
        self.style.configure("TEntry", font=('微软雅黑', 9))
        self.style.configure("Heading.TLabel", font=('微软雅黑', 12, 'bold'))
        self.style.configure("Status.TLabel", font=('微软雅黑', 9, 'bold'))
        self.style.configure("License.TLabel", font=('微软雅黑', 10))
        self.style.configure("TRadiobutton", font=('微软雅黑', 10))

        # 配置抗锯齿
        try:
            if os.name == 'nt':
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        # 显示免责声明
        if not self.show_disclaimer():
            print("用户不同意，退出程序")
            self.root.quit()
            sys.exit(0)
            return

        try:
            # 用户同意后，初始化大漠插件
            self.dm, self.dm2 = 注册大漠_简()
            if not self.dm or not self.dm2:
                messagebox.showerror("错误", "初始化失败，请检查插件是否正确安装")
                self.root.quit()
                return

            # 初始化界面
            self.init_ui()
        except Exception as e:
            messagebox.showerror("错误", f"程序初始化失败: {str(e)}")
            self.root.quit()

    def init_ui(self):
        """初始化用户界面"""
        self.root.deiconify()
        self.root.title("抢装备助手")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        self.root.attributes('-topmost', True)

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部信息栏
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 8))

        title_label = ttk.Label(header_frame, text="抢装备助手", style="Heading.TLabel")
        title_label.pack(side=tk.LEFT)

        # 状态显示
        self.status_frame = ttk.Frame(header_frame)
        self.status_frame.pack(side=tk.RIGHT)

        ttk.Label(self.status_frame, text="状态: ").pack(side=tk.LEFT)
        self.status_label = ttk.Label(self.status_frame, text="未运行", foreground="red", style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT)

        # 创建左右分栏
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧控制面板
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))

        # 价格区间设置
        price_frame = ttk.LabelFrame(left_frame, text="参数设置", padding=8)
        price_frame.pack(fill=tk.X, pady=(0, 8))

        price_grid = ttk.Frame(price_frame)
        price_grid.pack(fill=tk.X)

        # 1. 装备类型选择
        ttk.Label(price_grid, text="装备类型:").grid(row=0, column=0, padx=3, pady=3, sticky='w')
        self.equipment_type_var = tk.StringVar(value="头")
        equipment_type_combo = ttk.Combobox(price_grid, textvariable=self.equipment_type_var, values=["头", "甲"], width=5, state="readonly")
        equipment_type_combo.grid(row=0, column=1, padx=3, pady=3, sticky='w')

        # 2. 价格相关
        ttk.Label(price_grid, text="购入区间min:").grid(row=1, column=0, padx=3, pady=(10,3), sticky='w')
        self.min_price_entry = ttk.Entry(price_grid, width=8)
        self.min_price_entry.insert(0, "100000")
        self.min_price_entry.grid(row=1, column=1, padx=3, pady=(10,3), sticky='w')

        ttk.Label(price_grid, text="购入区间max:").grid(row=2, column=0, padx=3, pady=3, sticky='w')
        self.max_price_entry = ttk.Entry(price_grid, width=8)
        self.max_price_entry.insert(0, "390000")
        self.max_price_entry.grid(row=2, column=1, padx=3, pady=3, sticky='w')

        # 3. 延迟设置
        ttk.Label(price_grid, text="识别延迟(ms):").grid(row=3, column=0, padx=3, pady=(10,3), sticky='w')
        self.step2_delay_entry = ttk.Entry(price_grid, width=8)
        self.step2_delay_entry.insert(0, "100")
        self.step2_delay_entry.grid(row=3, column=1, padx=3, pady=(10,3), sticky='w')

        # 新增：识别相似度设置（百分比）
        ttk.Label(price_grid, text="识别相似度(%):").grid(row=7, column=0, padx=3, pady=(10,3), sticky='w')
        self.similarity_entry = ttk.Entry(price_grid, width=8)
        self.similarity_entry.insert(0, "85")
        self.similarity_entry.grid(row=7, column=1, padx=3, pady=(10,3), sticky='w')

        # 4. 功能开关
        self.test_mode_var = tk.BooleanVar(value=False)
        test_mode_check = ttk.Checkbutton(price_grid, text="测试模式(不点击购买)", variable=self.test_mode_var)
        test_mode_check.grid(row=4, column=0, columnspan=2, padx=3, pady=(10,3), sticky='w')

        # 5. 错误库管理
        error_lib_frame = ttk.Frame(price_grid)
        error_lib_frame.grid(row=5, column=0, columnspan=2, padx=3, pady=(10,3), sticky='ew')
        
        self.select_error_lib_btn = ttk.Button(error_lib_frame, text="选择错误库", command=self.select_error_library)
        self.select_error_lib_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.load_error_lib_btn = ttk.Button(error_lib_frame, text="加载错误库", command=self.load_error_library)
        self.load_error_lib_btn.pack(side=tk.LEFT)
        
        # 保存配置按钮
        self.save_config_btn = ttk.Button(error_lib_frame, text="保存配置", command=self.save_config)
        self.save_config_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 错误库路径显示
        self.error_lib_path_label = ttk.Label(price_grid, text="未选择错误库文件", font=('微软雅黑', 8), foreground='gray')
        self.error_lib_path_label.grid(row=6, column=0, columnspan=2, padx=3, pady=(0, 3), sticky='w')


        # 右侧日志区域
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # 日志区域
        log_frame = ttk.LabelFrame(right_frame, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_area = scrolledtext.ScrolledText(log_frame, height=15, font=('Consolas', 10),
                                                bg="white", fg="black")
        self.log_area.pack(fill=tk.BOTH, expand=True)

        # 按钮区
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(8, 0))

        self.reset_btn = ttk.Button(button_frame, text="重置设置", command=self.reset)
        self.reset_btn.pack(side=tk.LEFT, padx=3, expand=True, fill=tk.X)



        # 添加快捷键提示
        hotkey_frame = ttk.Frame(main_frame)
        hotkey_frame.pack(fill=tk.X, pady=(5, 0))
        
        hotkey_label = ttk.Label(hotkey_frame, text="按 F9/Home 启动/暂停", font=('微软雅黑', 9))
        hotkey_label.pack(side=tk.LEFT)

        # 添加版权信息
        copyright_frame = ttk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X, pady=(5, 0))

        copyright_text = "© 2025 P1nKM41D 版权所有 | 版本代号:装备 | 软件版本:v2.0.0"
        copyright_label = ttk.Label(copyright_frame, text=copyright_text, 
                                  font=('微软雅黑', 8), foreground='gray')
        copyright_label.pack(side=tk.RIGHT, padx=5)

        # 初始日志
        self.log("程序已准备就绪，按F9开始/停止")
        
        # 显示程序路径
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe
            base_path = os.path.dirname(sys.executable)
            self.log(f"📁 程序路径: {base_path}")
        else:
            # 如果是Python脚本
            base_path = os.path.dirname(os.path.abspath(__file__))
            self.log(f"📁 程序路径: {base_path}")
        
        # 加载保存的所有配置
        self.load_config()
            
        self.log("💡 提示：点击'选择错误库'按钮选择错误库文件")
        self.log("💡 提示：勾选'测试模式'可以测试识别功能而不实际购买")
        self.log("💡 提示：点击'保存配置'按钮可以手动保存当前设置")

    def _get_price_ocr_similarity(self):
        """获取当前OCR识别的相似度百分比"""
        try:
            similarity_percent = float(self.similarity_entry.get())
            if 0 <= similarity_percent <= 100:
                return similarity_percent / 100.0
            else:
                # 默认相似度，防止OCR失败
                return 0.85 # 默认相似度
        except ValueError:
            # 默认相似度，防止OCR失败
            return 0.85 # 默认相似度

    def 识别价格(self):
        """识别价格"""
        try:
            # 检查程序是否应该停止
            if not self.running:
                return None
                
            # 设置字库
            self.dm.SetDict(0, "ziku3.txt")
            
            # 执行OCR识别
            识别结果 = self.dm.OcrEx(326, 954, 418, 977, "cdcecf-303030", self._get_price_ocr_similarity())
            
            if not 识别结果:
                self.log("❌ 配装价格识别失败：未识别到任何内容")
                return None
            
            try:
                价格数字 = []
                for 部分 in 识别结果.split('|'):
                    if '$' in 部分:
                        数字 = 部分.split('$')[0]
                        if 数字.isdigit():
                            价格数字.append(数字)
                if not 价格数字:
                    self.log("❌ 配装价格识别失败：未找到有效数字")
                    return None
                    
                price = int(''.join(价格数字))
                self.log(f"识别价格: {price}")
                return price
                
            except ValueError as e:
                self.log(f"❌ 配装价格解析错误: {str(e)}")
                return None
                
        except Exception as e:
            self.log(f"配装价格识别错误: {e}")
            return None



    def main_loop(self):
        """主循环"""
        while self.running:
            # 每次循环开始都检查停止状态
            if not self.running:
                self.log("🛑 主循环检测到停止信号，立即退出")
                break
            if not self.paused:
                try:
                    # 检查是否应该停止
                    if not self.running:
                        break
                    # 获取价格区间、配装子弹数
                    min_price = float(self.min_price_entry.get())
                    max_price = float(self.max_price_entry.get())


                    # 新的购买逻辑：先点击两个区域，然后识别价格
                    if self.dangshi_var.get():
                        # 第一步：点击101,766,414,832内任意一处
                        click_x1 = random.randint(101, 414)
                        click_y1 = random.randint(766, 832)
                        self.dm.MoveTo(click_x1, click_y1)
                        time.sleep(0.01)
                        if not self.running:
                            break
                        self.dm.LeftClick()
                        time.sleep(0.01)
                        
                        if not self.running:
                            break
                            
                        # 第二步：点击107,884,297,975内任意一处
                        click_x2 = random.randint(107, 297)
                        click_y2 = random.randint(884, 975)
                        self.dm.MoveTo(click_x2, click_y2)
                        time.sleep(0.01)
                        if not self.running:
                            break
                        self.dm.LeftClick()
                        
                        # 使用用户自定义的第二步延迟
                        try:
                            step2_delay = float(self.step2_delay_entry.get()) / 1000.0  # 转换为秒
                            if step2_delay < 0:
                                step2_delay = 0.1
                        except ValueError:
                            step2_delay = 0.1
                        time.sleep(step2_delay)
                        
                        if not self.running:
                            break
                            
                        # 第三步：识别326,954,418,977区域内的价格
                        price = self.识别价格()
                        if not self.running:
                            break
                        if price is None:
                            self.配装价格识别错误计数 += 1
                            self.log(f"❌ 价格识别失败，连续错误次数: {self.配装价格识别错误计数}/5")
                            if self.配装价格识别错误计数 >= 5:
                                self.log("⚠️ 价格连续5次识别错误，执行恢复操作")
                                self.恢复操作计数 += 1
                                for i in range(5):
                                    if not self.running:
                                        self.log("🛑 恢复操作中检测到停止信号")
                                        break
                                    pyautogui.press('esc')
                                    time.sleep(0.3)
                                if not self.running:
                                    break
                                screen_width = self.root.winfo_screenwidth()
                                screen_height = self.root.winfo_screenheight()
                                center_x = screen_width // 2
                                center_y = screen_height // 2
                                self.dm.MoveTo(center_x, center_y)
                                time.sleep(0.1)
                                self.dm.LeftClick()
                                time.sleep(0.5)
                                if not self.running:
                                    break
                                self.配装价格识别错误计数 = 0
                                if self.恢复操作计数 >= 4:
                                    self.log("🔄 恢复操作已达3次，触发大战场重置...")
                                    self.大战场重置()
                                    self.恢复操作计数 = 0
                                    break
                                self.log("进入第四阶段...")
                                self.执行第四阶段()
                                break
                            continue
                            
                        self.配装价格识别错误计数 = 0
                        
                        # 检查价格是否在错误库中
                        if self.is_price_in_error_library(price):
                            self.log(f"⚠️ 价格 {price} 在错误库中，跳过")
                            if not self.running:
                                break
                            time.sleep(0.01)
                            continue
                        
                        # 判断价格是否在用户设定区间内
                        if min_price <= price <= max_price:
                            self.log(f"✅ 找到合适价格！价格: {price}")
                            if not self.running:
                                return
                                
                            # 检查是否为测试模式
                            if self.test_mode_var.get():
                                self.log("🧪 测试模式：跳过购买，继续下一轮")
                                time.sleep(0.1)
                                if not self.running:
                                    break
                                continue
                                
                            # 第四步：点击1613,843,1751,877内任意一处完成购买
                            buy_x = random.randint(1613, 1751)
                            buy_y = random.randint(843, 877)
                            self.dm.MoveTo(buy_x, buy_y)
                            time.sleep(0.01)
                            if not self.running:
                                break
                            self.dm.LeftClick()
                            time.sleep(0.01)
                            if not self.running:
                                break
                            self.dm.LeftClick()
                            
                            self.log("✅ 已点击购买！")
                            time.sleep(0.1)
                            if not self.running:
                                break
                            self.log("继续下一轮购买...")
                            # 不停止程序，继续循环
                            continue
                        else:
                            #self.log(f"价格 {price} 不在区间 {min_price}-{max_price}")
                            if not self.running:
                                break
                            time.sleep(0.01)
                            continue
                except Exception as e:
                    self.log(f"主循环错误: {e}")
                    time.sleep(0.5)
            time.sleep(0.05)  # 短暂休眠，避免CPU占用过高
        self.log("🛑 主循环已结束，程序已停止")

    def toggle_run(self):
        """切换运行状态"""
        if not self.running:
            self.start()
        else:
            self.stop()

    def start(self):
        """开始运行"""
        self.running = True
        self.status_label.config(text="运行中", foreground="green")
        self.log("🚀 程序已启动")
        
        # 隐藏窗口
        self.hide_window()
        
        # 启动主循环线程
        threading.Thread(target=self.main_loop, daemon=True).start()

    def stop(self):
        """停止运行"""
        self.running = False
        self.paused = False
        self.status_label.config(text="未运行", foreground="red")
        self.log("🛑 程序已停止")
        
        # 显示窗口
        self.show_window()
        
        self.root.deiconify()

    def reset(self):
        """重置设置"""
        self.stop()
        self.min_price_entry.delete(0, tk.END)
        self.min_price_entry.insert(0, "100000")
        self.max_price_entry.delete(0, tk.END)
        self.max_price_entry.insert(0, "390000")
        
        # 重置第二步延迟为100ms
        self.step2_delay_entry.delete(0, tk.END)
        self.step2_delay_entry.insert(0, "100")
        
        # 新增：重置OCR相似度为85%
        self.similarity_entry.delete(0, tk.END)
        self.similarity_entry.insert(0, "85")
        
        # 重置装备类型为头
        self.equipment_type_var.set("头")
        # 重置配装价格识别错误计数器
        self.配装价格识别错误计数 = 0
        self.恢复操作计数 = 0 # 重置恢复操作计数
        
        # 重置测试模式
        self.test_mode_var.set(False)
        
        # 重置错误库
        self.error_prices.clear()
        self.error_library_loaded = False
        self.error_library_path = ""
        self.error_lib_path_label.config(text="未选择错误库文件", foreground='gray')
        
        # 保存重置后的配置
        self.save_config()
        
        self.log("程序已重置")

    def select_error_library(self):
        """选择错误库文件"""
        try:
            from tkinter import filedialog
            
            # 打开文件选择对话框
            file_path = filedialog.askopenfilename(
                title="选择错误库文件",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialdir=os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
            )
            
            if file_path:
                self.error_library_path = file_path
                # 更新路径显示标签
                if len(file_path) > 50:
                    display_path = "..." + file_path[-47:]
                else:
                    display_path = file_path
                self.error_lib_path_label.config(text=display_path, foreground='black')
                self.log(f"✅ 已选择错误库文件: {file_path}")
                
                # 保存所有配置到配置文件
                self.save_config()
            else:
                self.log("❌ 未选择任何文件")
                
        except Exception as e:
            self.log(f"❌ 选择错误库文件失败: {str(e)}")
            messagebox.showerror("错误", f"选择错误库文件失败: {str(e)}")

    def save_config(self):
        """保存所有配置到配置文件"""
        try:
            config_path = os.path.join(
                os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)),
                "config.json"
            )
            
            config_data = {
                "error_library_path": self.error_library_path,
                "equipment_type": self.equipment_type_var.get(),
                "min_price": self.min_price_entry.get(),
                "max_price": self.max_price_entry.get(),
                "step2_delay": self.step2_delay_entry.get(),
                "test_mode": self.test_mode_var.get(),
                # 新增：OCR相似度百分比
                "ocr_similarity_percent": self.similarity_entry.get()
            }
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
                
            self.log("✅ 所有配置已保存到配置文件")
            
        except Exception as e:
            self.log(f"⚠️ 保存配置文件失败: {str(e)}")

    def load_config(self):
        """从配置文件加载所有配置"""
        try:
            config_path = os.path.join(
                os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)),
                "config.json"
            )
            
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                # 加载错误库路径
                if "error_library_path" in config_data:
                    self.error_library_path = config_data["error_library_path"]
                    # 更新路径显示标签
                    if len(self.error_library_path) > 50:
                        display_path = "..." + self.error_library_path[-47:]
                    else:
                        display_path = self.error_library_path
                    self.error_lib_path_label.config(text=display_path, foreground='black')
                    self.log(f"✅ 已加载保存的错误库路径: {self.error_library_path}")
                
                # 加载装备类型
                if "equipment_type" in config_data:
                    self.equipment_type_var.set(config_data["equipment_type"])
                    self.log(f"✅ 已加载装备类型设置: {config_data['equipment_type']}")
                
                # 加载价格区间
                if "min_price" in config_data:
                    self.min_price_entry.delete(0, tk.END)
                    self.min_price_entry.insert(0, config_data["min_price"])
                    self.log(f"✅ 已加载最低价格设置: {config_data['min_price']}")
                
                if "max_price" in config_data:
                    self.max_price_entry.delete(0, tk.END)
                    self.max_price_entry.insert(0, config_data["max_price"])
                    self.log(f"✅ 已加载最高价格设置: {config_data['max_price']}")
                
                # 加载第二步延迟
                if "step2_delay" in config_data:
                    self.step2_delay_entry.delete(0, tk.END)
                    self.step2_delay_entry.insert(0, config_data["step2_delay"])
                    self.log(f"✅ 已加载第二步延迟设置: {config_data['step2_delay']}ms")
                
                # 新增：加载OCR相似度百分比
                if "ocr_similarity_percent" in config_data:
                    self.similarity_entry.delete(0, tk.END)
                    self.similarity_entry.insert(0, str(config_data["ocr_similarity_percent"]))
                    self.log(f"✅ 已加载识别相似度设置: {config_data['ocr_similarity_percent']}%")
                
                # 加载测试模式
                if "test_mode" in config_data:
                    self.test_mode_var.set(config_data["test_mode"])
                    self.log(f"✅ 已加载测试模式设置: {'开启' if config_data['test_mode'] else '关闭'}")
                    
        except Exception as e:
            self.log(f"⚠️ 加载配置文件失败: {str(e)}")

    def load_error_library(self):
        """加载错误库文件"""
        try:
            # 检查是否已选择错误库文件
            if not self.error_library_path:
                self.log("❌ 请先选择错误库文件")
                messagebox.showwarning("警告", "请先点击'选择错误库'按钮选择错误库文件")
                return
            
            # 检查文件是否存在
            if not os.path.exists(self.error_library_path):
                self.log(f"❌ 错误库文件不存在: {self.error_library_path}")
                messagebox.showwarning("警告", f"错误库文件不存在: {self.error_library_path}")
                return
            
            # 清空之前的错误库
            self.error_prices.clear()
            
            # 读取错误库文件
            with open(self.error_library_path, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if line and not line.startswith('#'):  # 忽略空行和注释行
                        try:
                            price = int(line)
                            self.error_prices.add(price)
                        except ValueError:
                            self.log(f"⚠️ 第{line_num}行价格格式错误: {line}")
            
            self.error_library_loaded = True
            self.log(f"✅ 错误库加载成功，共加载 {len(self.error_prices)} 个错误价格")
            
            # 显示加载的错误价格（限制显示数量）
            if self.error_prices:
                sample_prices = list(self.error_prices)[:10]  # 只显示前10个
                self.log(f"错误价格示例: {', '.join(map(str, sample_prices))}")
                if len(self.error_prices) > 10:
                    self.log(f"... 还有 {len(self.error_prices) - 10} 个价格")
                    
        except Exception as e:
            self.log(f"❌ 加载错误库失败: {str(e)}")
            messagebox.showerror("错误", f"加载错误库失败: {str(e)}")
            self.error_library_loaded = False

    def is_price_in_error_library(self, price):
        """检查价格是否在错误库中"""
        if not self.error_library_loaded:
            return False
        return price in self.error_prices

    def log(self, message):
        """记录日志"""
        try:
            print(f"日志: {message}")
            
            if not hasattr(self, 'log_area') or self.log_area is None:
                print(message)
                return
                
            timestamp = time.strftime("[%H:%M:%S]")
            if self.log_area and self.log_area.winfo_exists():
                self.log_area.insert(tk.END, f"{timestamp} {message}\n")
                self.log_area.see(tk.END)
                self.log_area.update()
            
            try:
                if self.root and self.root.winfo_exists():
                    self.root.update()
            except:
                pass
        except Exception as e:
            print(f"日志记录错误: {str(e)}")
            print(f"原始消息: {message}")

    def show_disclaimer(self):
        """显示免责声明"""
        disclaimer_text = """
    免责声明:

    1. 本软件提供自动化辅助功能，用户在使用过程中需遵守相关法律法规。
    2. 软件许可使用权仅授予被许可人，未经授权不得转让、出租或出借本软件。
    3. 使用本软件可能违反特定服务或游戏的用户协议，由此产生的一切风险和后果由用户自行承担。
    4. 开发者不对使用本软件导致的任何直接或间接损失负责。
    5. 本软件不会收集用户个人隐私数据。
    6. 使用本软件视为同意本免责声明的全部条款，如不同意请勿使用本软件。
    7. 最终解释权归开发者所有。

    是否同意上述条款并继续使用？
        """

        disclaimer_window = tk.Toplevel(self.root)
        disclaimer_window.title("免责声明")
        disclaimer_window.geometry("520x420")
        disclaimer_window.resizable(False, False)
        disclaimer_window.attributes('-topmost', True)
        disclaimer_window.configure(bg="white")

        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(icon_path):
                disclaimer_window.iconbitmap(default=icon_path)
            else:
                print(f"图标文件不存在: {icon_path}")
        except Exception as e:
            print(f"设置免责窗口图标失败: {str(e)}")

        disclaimer_window.update_idletasks()
        width = disclaimer_window.winfo_width()
        height = disclaimer_window.winfo_height()
        x = (disclaimer_window.winfo_screenwidth() // 2) - (width // 2)
        y = (disclaimer_window.winfo_screenheight() // 2) - (height // 2)
        disclaimer_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        title_label = tk.Label(disclaimer_window, text="使用须知", font=('微软雅黑', 14, 'bold'),
                             bg="white", fg="#333333")
        title_label.pack(pady=15)

        text_frame = tk.Frame(disclaimer_window, bg="white", padx=20, pady=10)
        text_frame.pack(fill=tk.BOTH, expand=True)

        text_area = scrolledtext.ScrolledText(text_frame, width=50, height=10, wrap=tk.WORD,
                                            font=('微软雅黑', 10), bg="white", fg="black")
        text_area.pack(fill=tk.BOTH, expand=True)
        text_area.insert(tk.END, disclaimer_text)
        text_area.config(state=tk.DISABLED)

        user_response = tk.BooleanVar(value=False)

        button_frame = tk.Frame(disclaimer_window, bg="white", padx=20, pady=15)
        button_frame.pack(fill=tk.X)

        def on_agree():
            user_response.set(True)
            disclaimer_window.destroy()

        def on_disagree():
            user_response.set(False)
            disclaimer_window.destroy()

        agree_btn = tk.Button(button_frame, text="同意并继续", command=on_agree,
                            bg="#4CAF50", fg="white", font=('微软雅黑', 10),
                            relief=tk.FLAT, padx=15, pady=8)
        agree_btn.pack(side=tk.LEFT, padx=(70, 10))

        disagree_btn = tk.Button(button_frame, text="不同意并退出", command=on_disagree,
                               bg="#f44336", fg="white", font=('微软雅黑', 10),
                               relief=tk.FLAT, padx=15, pady=8)
        disagree_btn.pack(side=tk.RIGHT, padx=(10, 70))

        disclaimer_window.protocol("WM_DELETE_WINDOW", on_disagree)

        self.root.wait_window(disclaimer_window)
        return user_response.get()

    def on_focus_out(self, event):
        """当窗口失去焦点时触发"""
        if self.running and not self.paused:
            self.root.attributes('-topmost', True)
            self.root.lift()
            self.root.update()

    def on_focus_in(self, event):
        """当窗口获得焦点时触发"""
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.update()



    def 识别开始游戏(self):
        """识别开始游戏"""
        try:
            # 设置字库
            self.dm.SetDict(0, "ziku4.txt")
            
            # 执行OCR识别
            识别结果 = self.dm.OcrEx(146, 36, 237, 66, "e8e9e9-505050", 0.8)
            
            if not 识别结果:
                return False
                
            return "开始游戏" in 识别结果
                
        except Exception as e:
            self.log(f"开始游戏识别错误: {e}")
            return False
    

    def 识别配装字样(self):
        """识别配装字样"""
        try:
            # 设置字库
            self.dm.SetDict(0, "ziku4.txt")
            
            # 执行OCR识别
            识别结果 = self.dm.OcrEx(1517, 953, 1596, 980, "1b2526-404040", 0.85)
            
            if not 识别结果:
                return False
                
            return "配装" in 识别结果
                
        except Exception as e:
            self.log(f"配装字样识别错误: {e}")
            return False

    def 执行第四阶段(self):
        """执行第四阶段操作"""
        self.log("开始执行第四阶段...")
        
        # 检查停止状态
        if not self.running:
            self.log("🛑 第四阶段开始前检测到停止信号")
            return
        
        # 点击开始游戏
        self.dm.MoveTo(192, 54)
        time.sleep(0.1)
        if not self.running:
            return
        self.dm.LeftClick()
        time.sleep(1)
        
        if not self.running:
            return
        
        # 识别是否有配装字样
        if self.识别配装字样():
            self.log("检测到配装字样，直接点击配装按钮")
            # 点击配装按钮
            self.dm.MoveTo(1557, 966)
            time.sleep(0.1)
            if not self.running:
                return
            self.dm.LeftClick()
            time.sleep(1)
        else:
            self.log("未检测到配装字样，执行正常流程")
            # 点击地图
            self.dm.MoveTo(1636, 882)
            time.sleep(0.1)
            if not self.running:
                return
            self.dm.LeftClick()
            time.sleep(1)
            # 新增：识别零号大坝
            self.dm.SetDict(0, "ziku4.txt")
            零号大坝结果 = self.dm.OcrEx(760, 184, 890, 241, "fefeff-404040", 0.6)
            self.log(f"零号大坝识别结果: {零号大坝结果}")
            if 零号大坝结果 and "零号大坝" in 零号大坝结果:
                self.log("检测到零号大坝，直接进入常规大坝")
                self.dm.MoveTo(827, 262)
                self.dm.LeftClick()
                time.sleep(0.5)
            else:
                self.log("未检测到零号大坝，按ESC后进入常规大坝")
                pyautogui.press("esc")
                time.sleep(0.5)
                self.dm.MoveTo(827, 262)
                self.dm.LeftClick()
                time.sleep(0.5)
            
            if not self.running:
                return
                
            # 点击开始行动
            self.dm.MoveTo(1658, 639)
            time.sleep(0.1)
            if not self.running:
                return
            self.dm.LeftClick()
            time.sleep(1)
        
        if not self.running:
            return
        
        # 点击进入配装
        self.dm.MoveTo(383, 1050)
        #根据装备类型选择点击位置
        equipment_type = self.equipment_type_var.get()
        if equipment_type == "头":
            # 如果选头则点击236,410
            enterpz_x = 236
            enterpz_y = 410
        elif equipment_type == "甲":
            # 如果是甲则点击236,527
            enterpz_x = 236
            enterpz_y = 527
        else:
            # 默认点击头的位置
            enterpz_x = 236
            enterpz_y = 410
        self.dm.MoveTo(enterpz_x, enterpz_y)
        time.sleep(0.1)
        if not self.running:
            return
        self.dm.LeftClick()
        time.sleep(0.5)
        self.dm.MoveTo(433, 975)
        time.sleep(0.1)
        self.dm.LeftClick()
        time.sleep(0.5)
        self.dm.MoveTo(433, 975)
        time.sleep(0.1)
        self.dm.LeftClick()
        time.sleep(0.5)
        
        # 添加鼠标滚轮向下滑动操作
        if not self.running:
            return
        #self.log("执行鼠标滚轮向下滑动...")
        self.dm.WheelDown()  # 向下滚动
        time.sleep(0.5)  # 等待0.5秒
        
        if not self.running:
            return
        
        # 重启程序（改为只重启主循环线程，不重置对象）
        self.log("重启程序...")
        self.running = False
        self.paused = False
        self.stop()
        # self.root.after(2000, self.safe_restart)  # 注释掉原有重启
        # 新增：直接重启主循环线程
        def restart_main_loop():
            self.running = True
            self.paused = False
            self.status_label.config(text="运行中", foreground="green")
            self.log("🚀 程序已重新启动")
            self.hide_window()
            threading.Thread(target=self.main_loop, daemon=True).start()
        self.root.after(2000, restart_main_loop)
        
        # 第四阶段结束检查点
        if not self.running:
            self.log("🛑 第四阶段执行完毕，程序已停止")
            return



    def hide_window(self):
        """隐藏窗口"""
        if not self.window_hidden:
            self.root.withdraw()  # 隐藏窗口
            self.window_hidden = True
            # 设置窗口样式为工具窗口，这样在Alt+Tab中不会显示
            self.root.attributes('-toolwindow', True)
            # 移除窗口的标题栏
            self.root.overrideredirect(True)
            # 设置窗口透明度为0
            self.root.attributes('-alpha', 0)

    def show_window(self):
        """显示窗口"""
        if self.window_hidden:
            self.root.deiconify()  # 显示窗口
            self.window_hidden = False
            # 恢复窗口样式
            self.root.attributes('-toolwindow', False)
            # 恢复标题栏
            self.root.overrideredirect(False)
            # 恢复窗口透明度
            self.root.attributes('-alpha', 1)
            # 确保窗口置顶
            self.root.attributes('-topmost', True)
            self.root.lift()
            self.root.update()



    def 大战场重置(self):
        """大战场重置流程，完成后进入第四阶段"""
        try:
            self.log("🔄 开始大战场重置流程...")
            max_attempts = 15
            for attempt in range(max_attempts):
                if not self.running:
                    self.log("程序已停止，退出大战场重置...")
                    return
                self.dm.SetDict(0, "ziku4.txt")
                keyboard.send('esc')
                time.sleep(1)
                if not self.running:
                    self.log("程序已停止，退出大战场重置...")
                    return
                # 检测烽火地带
                烽火结果 = self.dm.OcrEx(129, 327, 217, 350, "bec1c2-404040", 0.8)
                if 烽火结果 and "烽火地带" in 烽火结果:
                    self.log("✅ 检测到烽火地带，开始大战场重置...")
                    break
                if attempt == max_attempts - 1:
                    self.log("❌ 大战场重置失败：未检测到烽火地带")
                    return
                time.sleep(1)
            # 点击266,487
            if not self.running:
                return
            self.dm.MoveTo(266, 487)
            time.sleep(0.1)
            self.dm.LeftClick()
            time.sleep(3)
            # 按一次ESC
            if not self.running:
                return
            keyboard.send('esc')
            time.sleep(1)
            # 点击263,288
            if not self.running:
                return
            self.dm.MoveTo(263, 288)
            time.sleep(0.1)
            self.dm.LeftClick()
            time.sleep(2)
            self.log("✅ 大战场重置完成，重新进入配装页...")
            # 进入第四阶段
            self.执行第四阶段()
            self.恢复操作计数 = 0 # 重置恢复操作计数
        except Exception as e:
            self.log(f"❌ 大战场重置过程发生错误: {str(e)}")

def main():
    try:
        if not is_admin():
            run_as_admin()
            return

        root = tk.Tk()
        root.withdraw()

        root.title("满仓助手")
        root.geometry("600x500")  # 修改窗口大小为600x500
        root.minsize(600, 500)  # 修改最小窗口大小为600x500
        root.attributes('-topmost', True)

        root.option_add('*Font', '微软雅黑 12')

        try:
            app = 满仓助手(root)

            if not root.winfo_exists():
                return

            def toggle_function():
                try:
                    if not app.running:
                        app.log("🔄 通过快捷键启动程序")
                        app.start()
                    else:
                        app.log("🛑 通过快捷键立即停止程序")
                        # 立即设置停止标志
                        app.running = False
                        app.paused = False
                        # 立即停止程序
                        app.stop()
                        # 强制显示窗口
                        app.show_window()
                        app.root.deiconify()
                        app.root.lift()
                        app.root.focus_force()
                except Exception as e:
                    app.log(f"❌ 热键操作错误: {str(e)}")

            def safe_toggle():
                try:
                    toggle_function()
                except Exception as e:
                    print(f"热键错误: {str(e)}")

            # 添加F9和Home热键 - 立即停止功能
            keyboard.add_hotkey('F9', safe_toggle)
            keyboard.add_hotkey('home', safe_toggle)
            app.log("✨ 程序已就绪，F9或Home键可立即停止程序")

            root.mainloop()
        except Exception as e:
            messagebox.showerror("错误", f"程序初始化失败: {str(e)}")
        finally:
            try:
                keyboard.remove_hotkey('F9')
                keyboard.remove_hotkey('home')
            except:
                pass

    except Exception as e:
        print(f"程序启动失败: {str(e)}")
        print("请确保已正确安装Python和Tcl/Tk")
        print(f"当前Python路径: {sys.executable}")
        print(f"当前Tcl路径: {os.environ.get('TCL_LIBRARY', '未设置')}")
        print(f"当前Tk路径: {os.environ.get('TK_LIBRARY', '未设置')}")
        input("按回车键退出...")
        sys.exit(1)

if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()