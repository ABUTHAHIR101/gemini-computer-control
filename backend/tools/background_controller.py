"""
后台窗口控制器
使用 pywinauto 和 Win32 API 实现后台窗口操作，不需要窗口焦点
支持后台截图、后台鼠标点击、后台键盘输入
"""
from typing import Dict, Any, Optional, Tuple, List
import logging
import base64
import io
import time
import ctypes
from ctypes import wintypes

try:
    import win32gui
    import win32con
    import win32api
    import win32ui
    from PIL import Image
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    from pywinauto.findwindows import find_windows
    HAS_PYWINAUTO = True
except ImportError:
    HAS_PYWINAUTO = False

logger = logging.getLogger(__name__)

# Windows API 常量
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_MBUTTONDOWN = 0x0207
WM_MBUTTONUP = 0x0208
WM_MOUSEMOVE = 0x0200
WM_MOUSEWHEEL = 0x020A
WM_CHAR = 0x0102
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101

# 虚拟键码映射
VK_MAP = {
    'enter': 0x0D,
    'return': 0x0D,
    'tab': 0x09,
    'space': 0x20,
    'backspace': 0x08,
    'delete': 0x2E,
    'escape': 0x1B,
    'esc': 0x1B,
    'ctrl': 0x11,
    'control': 0x11,
    'alt': 0x12,
    'shift': 0x10,
    'win': 0x5B,
    'up': 0x26,
    'down': 0x28,
    'left': 0x25,
    'right': 0x27,
    'home': 0x24,
    'end': 0x23,
    'pageup': 0x21,
    'pagedown': 0x22,
    'insert': 0x2D,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73,
    'f5': 0x74, 'f6': 0x75, 'f7': 0x76, 'f8': 0x77,
    'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
    'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
    'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
    'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
    'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
}


def make_lparam(x: int, y: int) -> int:
    """创建 LPARAM 值（低16位是x，高16位是y）"""
    return (y << 16) | (x & 0xFFFF)


class BackgroundController:
    """
    后台窗口控制器
    使用 Win32 API 实现后台操作，不干扰用户的前台操作
    """
    
    def __init__(self, window_title: Optional[str] = None, window_class: Optional[str] = None):
        """
        初始化后台控制器
        
        Args:
            window_title: 目标窗口标题（支持部分匹配）
            window_class: 目标窗口类名
        """
        if not HAS_WIN32:
            raise ImportError("需要安装 pywin32: pip install pywin32")
        
        self.window_title = window_title
        self.window_class = window_class
        self.hwnd: Optional[int] = None
        self.window_width = 0
        self.window_height = 0
        
        # 如果指定了窗口，尝试查找
        if window_title or window_class:
            self.find_window(window_title, window_class)
        
        logger.info(f"后台控制器初始化完成, pywin32={HAS_WIN32}, pywinauto={HAS_PYWINAUTO}")
    
    def find_window(self, title: Optional[str] = None, class_name: Optional[str] = None) -> bool:
        """
        查找目标窗口
        
        Args:
            title: 窗口标题（支持部分匹配）
            class_name: 窗口类名
            
        Returns:
            是否找到窗口
        """
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                window_class = win32gui.GetClassName(hwnd)
                
                title_match = True
                class_match = True
                
                if title:
                    title_match = title.lower() in window_text.lower()
                if class_name:
                    class_match = class_name.lower() in window_class.lower()
                
                if title_match and class_match:
                    results.append((hwnd, window_text, window_class))
            return True
        
        results = []
        win32gui.EnumWindows(enum_windows_callback, results)
        
        if results:
            self.hwnd = results[0][0]
            self.window_title = results[0][1]
            self.window_class = results[0][2]
            
            # 获取窗口大小
            rect = win32gui.GetClientRect(self.hwnd)
            self.window_width = rect[2] - rect[0]
            self.window_height = rect[3] - rect[1]
            
            logger.info(f"找到窗口: {self.window_title} ({self.window_class}), "
                       f"尺寸: {self.window_width}x{self.window_height}")
            return True
        
        logger.warning(f"未找到匹配的窗口: title={title}, class={class_name}")
        return False
    
    def list_windows(self) -> List[Dict[str, Any]]:
        """
        列出所有可见窗口
        
        Returns:
            窗口信息列表
        """
        windows = []
        
        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # 只显示有标题的窗口
                    class_name = win32gui.GetClassName(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    windows.append({
                        "hwnd": hwnd,
                        "title": title,
                        "class": class_name,
                        "rect": {
                            "left": rect[0],
                            "top": rect[1],
                            "right": rect[2],
                            "bottom": rect[3],
                            "width": rect[2] - rect[0],
                            "height": rect[3] - rect[1]
                        }
                    })
            return True
        
        win32gui.EnumWindows(enum_callback, None)
        return windows
    
    def set_target_window(self, hwnd: int) -> bool:
        """
        设置目标窗口句柄
        
        Args:
            hwnd: 窗口句柄
            
        Returns:
            是否成功
        """
        if not win32gui.IsWindow(hwnd):
            logger.error(f"无效的窗口句柄: {hwnd}")
            return False
        
        self.hwnd = hwnd
        self.window_title = win32gui.GetWindowText(hwnd)
        self.window_class = win32gui.GetClassName(hwnd)
        
        rect = win32gui.GetClientRect(hwnd)
        self.window_width = rect[2] - rect[0]
        self.window_height = rect[3] - rect[1]
        
        logger.info(f"设置目标窗口: {self.window_title}, 尺寸: {self.window_width}x{self.window_height}")
        return True
    
    def take_screenshot(self) -> Dict[str, Any]:
        """
        截取目标窗口截图（后台截图，不需要窗口在前台）
        
        Returns:
            包含base64编码截图的字典
        """
        if not self.hwnd:
            return {
                "success": False,
                "error": "未设置目标窗口"
            }
        
        try:
            # 1. 尝试强制窗口重绘，防止后台挂起不渲染
            win32gui.InvalidateRect(self.hwnd, None, True)
            win32gui.UpdateWindow(self.hwnd)
            
            # 获取窗口尺寸
            left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
            width = right - left
            height = bottom - top
            
            # 创建设备上下文
            hwnd_dc = win32gui.GetWindowDC(self.hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            
            # 创建位图
            bitmap = win32ui.CreateBitmap()
            bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(bitmap)
            
            # 使用 PrintWindow 进行后台截图
            # PW_RENDERFULLCONTENT = 2 可以捕获 DWM 合成的内容
            result = ctypes.windll.user32.PrintWindow(self.hwnd, save_dc.GetSafeHdc(), 2)
            
            if not result:
                # 如果 PrintWindow 失败，尝试 BitBlt
                save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)
            
            # 获取位图数据
            bmpinfo = bitmap.GetInfo()
            bmpstr = bitmap.GetBitmapBits(True)
            
            # 转换为 PIL Image
            img = Image.frombuffer(
                'RGB',
                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                bmpstr, 'raw', 'BGRX', 0, 1
            )
            
            # 清理资源
            win32gui.DeleteObject(bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(self.hwnd, hwnd_dc)
            
            # 转换为 base64
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            screenshot_base64 = f"data:image/png;base64,{base64.b64encode(buffered.getvalue()).decode()}"
            
            # 更新窗口尺寸
            self.window_width = width
            self.window_height = height
            
            logger.info(f"后台截图成功: {self.window_title}, 尺寸: {width}x{height}")
            
            return {
                "success": True,
                "screenshot": screenshot_base64,
                "window_title": self.window_title,
                "width": width,
                "height": height
            }
            
        except Exception as e:
            logger.error(f"后台截图失败: {str(e)}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def mouse_click(self, x: int, y: int, button: str = "left") -> Dict[str, Any]:
        """
        后台鼠标点击（不需要窗口焦点）
        
        Args:
            x: 点击位置的x坐标（相对于窗口客户区）
            y: 点击位置的y坐标（相对于窗口客户区）
            button: 鼠标按键 ("left", "middle", "right")
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            lparam = make_lparam(x, y)
            
            if button == "left":
                win32gui.PostMessage(self.hwnd, WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                time.sleep(0.05)
                win32gui.PostMessage(self.hwnd, WM_LBUTTONUP, 0, lparam)
            elif button == "right":
                win32gui.PostMessage(self.hwnd, WM_RBUTTONDOWN, win32con.MK_RBUTTON, lparam)
                time.sleep(0.05)
                win32gui.PostMessage(self.hwnd, WM_RBUTTONUP, 0, lparam)
            elif button == "middle":
                win32gui.PostMessage(self.hwnd, WM_MBUTTONDOWN, win32con.MK_MBUTTON, lparam)
                time.sleep(0.05)
                win32gui.PostMessage(self.hwnd, WM_MBUTTONUP, 0, lparam)
            
            logger.info(f"后台点击: ({x}, {y}), 按键: {button}")
            
            return {
                "status": "success",
                "action": "mouse_click",
                "x": x,
                "y": y,
                "button": button,
                "message": f"已在坐标({x}, {y})执行{button}键后台点击"
            }
            
        except Exception as e:
            logger.error(f"后台点击失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def mouse_double_click(self, x: int, y: int, button: str = "left") -> Dict[str, Any]:
        """
        后台鼠标双击
        
        Args:
            x: 点击位置的x坐标
            y: 点击位置的y坐标
            button: 鼠标按键
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            lparam = make_lparam(x, y)
            
            if button == "left":
                win32gui.PostMessage(self.hwnd, WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lparam)
                time.sleep(0.05)
                win32gui.PostMessage(self.hwnd, WM_LBUTTONUP, 0, lparam)
            
            logger.info(f"后台双击: ({x}, {y})")
            
            return {
                "status": "success",
                "action": "mouse_double_click",
                "x": x,
                "y": y,
                "button": button,
                "message": f"已在坐标({x}, {y})执行后台双击"
            }
            
        except Exception as e:
            logger.error(f"后台双击失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def mouse_move(self, x: int, y: int) -> Dict[str, Any]:
        """
        后台鼠标移动
        
        Args:
            x: 目标x坐标
            y: 目标y坐标
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            lparam = make_lparam(x, y)
            win32gui.PostMessage(self.hwnd, WM_MOUSEMOVE, 0, lparam)
            
            logger.info(f"后台鼠标移动: ({x}, {y})")
            
            return {
                "status": "success",
                "action": "mouse_hover",
                "x": x,
                "y": y,
                "message": f"已将鼠标移动到坐标({x}, {y})"
            }
            
        except Exception as e:
            logger.error(f"后台鼠标移动失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def mouse_drag(self, start_x: int, start_y: int, end_x: int, end_y: int, 
                   button: str = "left", steps: int = 10) -> Dict[str, Any]:
        """
        后台鼠标拖动
        
        Args:
            start_x: 起点x坐标
            start_y: 起点y坐标
            end_x: 终点x坐标
            end_y: 终点y坐标
            button: 鼠标按键
            steps: 拖动步数
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            # 鼠标按下
            start_lparam = make_lparam(start_x, start_y)
            if button == "left":
                win32gui.PostMessage(self.hwnd, WM_LBUTTONDOWN, win32con.MK_LBUTTON, start_lparam)
            
            time.sleep(0.05)
            
            # 逐步移动
            for i in range(1, steps + 1):
                x = start_x + (end_x - start_x) * i // steps
                y = start_y + (end_y - start_y) * i // steps
                lparam = make_lparam(x, y)
                win32gui.PostMessage(self.hwnd, WM_MOUSEMOVE, win32con.MK_LBUTTON, lparam)
                time.sleep(0.02)
            
            # 鼠标释放
            end_lparam = make_lparam(end_x, end_y)
            if button == "left":
                win32gui.PostMessage(self.hwnd, WM_LBUTTONUP, 0, end_lparam)
            
            logger.info(f"后台拖动: ({start_x}, {start_y}) -> ({end_x}, {end_y})")
            
            return {
                "status": "success",
                "action": "mouse_drag",
                "start_x": start_x,
                "start_y": start_y,
                "end_x": end_x,
                "end_y": end_y,
                "button": button,
                "message": f"已从({start_x}, {start_y})拖动到({end_x}, {end_y})"
            }
            
        except Exception as e:
            logger.error(f"后台拖动失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def mouse_scroll(self, scroll_x: int = 0, scroll_y: int = 0, x: int = 0, y: int = 0) -> Dict[str, Any]:
        """
        后台鼠标滚轮
        
        Args:
            scroll_x: 水平滚动量
            scroll_y: 垂直滚动量（正数向下，负数向上）
            x: 滚动位置x坐标
            y: 滚动位置y坐标
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            # WM_MOUSEWHEEL: wParam 高位是滚动量，低位是按键状态
            # 滚动量: 正数向上，负数向下（与我们的约定相反，需要取反）
            wheel_delta = -scroll_y * 120  # 120 是标准滚动单位
            wparam = (wheel_delta << 16) & 0xFFFF0000
            lparam = make_lparam(x, y)
            
            win32gui.PostMessage(self.hwnd, WM_MOUSEWHEEL, wparam, lparam)
            
            logger.info(f"后台滚轮: scroll_y={scroll_y}")
            
            return {
                "status": "success",
                "action": "mouse_scroll",
                "scroll_x": scroll_x,
                "scroll_y": scroll_y,
                "message": f"已执行后台滚轮操作，滚动量: ({scroll_x}, {scroll_y})"
            }
            
        except Exception as e:
            logger.error(f"后台滚轮失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def keyboard_type(self, text: str, clear_existing: bool = False) -> Dict[str, Any]:
        """
        后台键盘输入文本
        
        Args:
            text: 要输入的文本
            clear_existing: 是否先清除现有文本
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            # 如果需要清除现有文本，使用 WM_SETTEXT 或全选删除
            if clear_existing:
                # 方法1: 尝试使用 EM_SETSEL + WM_CLEAR（适用于 Edit 控件）
                EM_SETSEL = 0x00B1
                WM_CLEAR = 0x0303
                win32gui.SendMessageTimeout(self.hwnd, EM_SETSEL, 0, -1, win32con.SMTO_ABORTIFHUNG, 100)  # 全选
                time.sleep(0.02)
                win32gui.SendMessageTimeout(self.hwnd, WM_CLEAR, 0, 0, win32con.SMTO_ABORTIFHUNG, 100)  # 清除
                time.sleep(0.02)
            
            # 逐字符发送
            for char in text:
                # 使用 WM_CHAR 发送字符
                win32gui.PostMessage(self.hwnd, WM_CHAR, ord(char), 0)
                time.sleep(0.01)
            
            logger.info(f"后台输入文本: {text[:50]}{'...' if len(text) > 50 else ''}")
            
            return {
                "status": "success",
                "action": "keyboard_type",
                "text": text,
                "text_length": len(text),
                "clear_existing": clear_existing,
                "message": f"已输入文本，长度{len(text)}字符"
            }
            
        except Exception as e:
            logger.error(f"后台输入文本失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def keyboard_press(self, keys: List[str]) -> Dict[str, Any]:
        """
        后台按键操作
        注意：对于现代浏览器（Chrome、Edge等），后台按键可能不完全工作
        
        Args:
            keys: 按键列表（组合键时多个元素）
            
        Returns:
            执行结果
        """
        if not self.hwnd:
            return {"status": "error", "error": "未设置目标窗口"}
        
        try:
            keys_str = "+".join(keys)
            
            # 特殊处理常见的组合键
            if len(keys) == 2:
                key1, key2 = keys[0].lower(), keys[1].lower()
                
                # Ctrl+A: 全选 - 使用 EM_SETSEL
                if key1 == 'ctrl' and key2 == 'a':
                    EM_SETSEL = 0x00B1
                    win32gui.SendMessage(self.hwnd, EM_SETSEL, 0, -1)
                    logger.info(f"后台按键 (EM_SETSEL): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Ctrl+C: 复制 - 使用 WM_COPY
                if key1 == 'ctrl' and key2 == 'c':
                    WM_COPY = 0x0301
                    win32gui.SendMessage(self.hwnd, WM_COPY, 0, 0)
                    logger.info(f"后台按键 (WM_COPY): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Ctrl+V: 粘贴 - 使用 WM_PASTE
                if key1 == 'ctrl' and key2 == 'v':
                    WM_PASTE = 0x0302
                    win32gui.SendMessage(self.hwnd, WM_PASTE, 0, 0)
                    logger.info(f"后台按键 (WM_PASTE): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Ctrl+X: 剪切 - 使用 WM_CUT
                if key1 == 'ctrl' and key2 == 'x':
                    WM_CUT = 0x0300
                    win32gui.SendMessage(self.hwnd, WM_CUT, 0, 0)
                    logger.info(f"后台按键 (WM_CUT): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
            
            # 单个按键
            if len(keys) == 1:
                key = keys[0].lower()
                
                # Delete 键
                if key == 'delete':
                    WM_CLEAR = 0x0303
                    win32gui.SendMessage(self.hwnd, WM_CLEAR, 0, 0)
                    logger.info(f"后台按键 (WM_CLEAR): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Enter 键 - 使用 WM_KEYDOWN + WM_KEYUP
                if key in ('enter', 'return'):
                    VK_RETURN = 0x0D
                    # 构造 lParam：scan code 为 0x1C (Enter 的扫描码)
                    lparam_down = 1 | (0x1C << 16)  # repeat=1, scan=0x1C
                    lparam_up = 1 | (0x1C << 16) | 0xC0000000  # 加上 release 标志
                    
                    win32gui.PostMessage(self.hwnd, WM_KEYDOWN, VK_RETURN, lparam_down)
                    time.sleep(0.05)
                    win32gui.PostMessage(self.hwnd, WM_KEYUP, VK_RETURN, lparam_up)
                    
                    logger.info(f"后台按键 (WM_KEYDOWN/UP Enter): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Tab 键 - 使用 WM_KEYDOWN + WM_KEYUP
                if key == 'tab':
                    VK_TAB = 0x09
                    # 构造 lParam：scan code 为 0x0F (Tab 的扫描码)
                    lparam_down = 1 | (0x0F << 16)
                    lparam_up = 1 | (0x0F << 16) | 0xC0000000
                    
                    win32gui.PostMessage(self.hwnd, WM_KEYDOWN, VK_TAB, lparam_down)
                    time.sleep(0.05)
                    win32gui.PostMessage(self.hwnd, WM_KEYUP, VK_TAB, lparam_up)
                    
                    logger.info(f"后台按键 (WM_KEYDOWN/UP Tab): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Escape 键
                if key in ('escape', 'esc'):
                    VK_ESCAPE = 0x1B
                    lparam_down = 1 | (0x01 << 16)
                    lparam_up = 1 | (0x01 << 16) | 0xC0000000
                    
                    win32gui.PostMessage(self.hwnd, WM_KEYDOWN, VK_ESCAPE, lparam_down)
                    time.sleep(0.05)
                    win32gui.PostMessage(self.hwnd, WM_KEYUP, VK_ESCAPE, lparam_up)
                    
                    logger.info(f"后台按键 (WM_KEYDOWN/UP Escape): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
                
                # Backspace 键
                if key == 'backspace':
                    VK_BACK = 0x08
                    lparam_down = 1 | (0x0E << 16)
                    lparam_up = 1 | (0x0E << 16) | 0xC0000000
                    
                    win32gui.PostMessage(self.hwnd, WM_KEYDOWN, VK_BACK, lparam_down)
                    time.sleep(0.05)
                    win32gui.PostMessage(self.hwnd, WM_KEYUP, VK_BACK, lparam_up)
                    
                    logger.info(f"后台按键 (WM_KEYDOWN/UP Backspace): {keys_str}")
                    return {
                        "status": "success",
                        "action": "keyboard_press",
                        "keys": keys,
                        "keys_str": keys_str,
                        "message": f"已执行按键操作：{keys_str}"
                    }
            
            # 通用按键处理（使用 WM_KEYDOWN/WM_KEYUP）
            # 获取虚拟键码
            vk_codes = []
            for key in keys:
                key_lower = key.lower()
                if key_lower in VK_MAP:
                    vk_codes.append(VK_MAP[key_lower])
                elif len(key) == 1:
                    vk_codes.append(ord(key.upper()))
                else:
                    logger.warning(f"未知按键: {key}")
                    continue
            
            if not vk_codes:
                return {"status": "error", "error": "没有有效的按键"}
            
            # 构造 lParam（包含扫描码等信息）
            def make_key_lparam(vk: int, is_up: bool = False) -> int:
                scan_code = win32api.MapVirtualKey(vk, 0) if hasattr(win32api, 'MapVirtualKey') else 0
                lparam = 1  # repeat count
                lparam |= (scan_code & 0xFF) << 16  # scan code
                if is_up:
                    lparam |= 0xC0000000  # transition state and previous key state
                return lparam
            
            # 发送修饰键按下
            for vk in vk_codes[:-1]:
                lparam = make_key_lparam(vk, False)
                win32gui.PostMessage(self.hwnd, WM_KEYDOWN, vk, lparam)
                time.sleep(0.01)
            
            # 发送主键
            if vk_codes:
                main_vk = vk_codes[-1]
                lparam_down = make_key_lparam(main_vk, False)
                lparam_up = make_key_lparam(main_vk, True)
                win32gui.PostMessage(self.hwnd, WM_KEYDOWN, main_vk, lparam_down)
                time.sleep(0.03)
                win32gui.PostMessage(self.hwnd, WM_KEYUP, main_vk, lparam_up)
            
            # 释放修饰键
            for vk in reversed(vk_codes[:-1]):
                lparam = make_key_lparam(vk, True)
                win32gui.PostMessage(self.hwnd, WM_KEYUP, vk, lparam)
                time.sleep(0.01)
            
            logger.info(f"后台按键: {keys_str}")
            
            return {
                "status": "success",
                "action": "keyboard_press",
                "keys": keys,
                "keys_str": keys_str,
                "message": f"已执行按键操作：{keys_str}"
            }
            
        except Exception as e:
            logger.error(f"后台按键失败: {str(e)}")
            return {"status": "error", "error": str(e)}
    
    def clear_text(self) -> Dict[str, Any]:
        """
        清除当前输入框文本
        
        Returns:
            执行结果
        """
        return self.keyboard_press(['ctrl', 'a']) and self.keyboard_press(['delete'])
    
    def click_and_type(self, x: int, y: int, text: str = "", clear_existing: bool = True) -> Dict[str, Any]:
        """
        点击后输入文本（组合操作）
        注意：点击操作可能会激活窗口，这是某些应用的固有行为
        
        Args:
            x: 点击位置x坐标
            y: 点击位置y坐标
            text: 要输入的文本
            clear_existing: 是否清除现有文本
            
        Returns:
            执行结果
        """
        # 记录当前前台窗口
        foreground_hwnd = win32gui.GetForegroundWindow()
        
        # 先点击
        click_result = self.mouse_click(x, y, "left")
        if click_result.get("status") != "success":
            return click_result
        
        time.sleep(0.15)
        
        # 如果需要清除现有文本
        if clear_existing:
            # 使用 EM_SETSEL 全选
            EM_SETSEL = 0x00B1
            WM_CLEAR = 0x0303
            win32gui.SendMessage(self.hwnd, EM_SETSEL, 0, -1)
            time.sleep(0.02)
            win32gui.SendMessage(self.hwnd, WM_CLEAR, 0, 0)
            time.sleep(0.02)
        
        # 输入文本
        if text:
            for char in text:
                win32gui.PostMessage(self.hwnd, WM_CHAR, ord(char), 0)
                time.sleep(0.01)
        
        # 尝试恢复之前的前台窗口（可选，可能失败）
        try:
            if foreground_hwnd and foreground_hwnd != self.hwnd:
                # 注意：SetForegroundWindow 可能因为权限问题失败
                pass  # 暂不恢复，因为这可能导致更多问题
        except:
            pass
        
        return {
            "status": "success",
            "action": "click_and_type",
            "x": x,
            "y": y,
            "text": text,
            "text_length": len(text),
            "clear_existing": clear_existing,
            "message": f"已在坐标({x}, {y})点击并输入文本"
        }
    
    def get_window_info(self) -> Dict[str, Any]:
        """
        获取当前窗口信息
        
        Returns:
            窗口信息字典
        """
        if not self.hwnd:
            return {
                "success": False,
                "error": "未设置目标窗口"
            }
        
        return {
            "success": True,
            "hwnd": self.hwnd,
            "title": self.window_title,
            "class": self.window_class,
            "width": self.window_width,
            "height": self.window_height
        }
    
    def bring_to_front(self) -> bool:
        """
        将窗口带到前台（可选操作）
        
        Returns:
            是否成功
        """
        if not self.hwnd:
            return False
        
        try:
            # 如果窗口最小化，恢复它
            if win32gui.IsIconic(self.hwnd):
                win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            
            # 将窗口带到前台
            win32gui.SetForegroundWindow(self.hwnd)
            return True
            
        except Exception as e:
            logger.error(f"无法将窗口带到前台: {str(e)}")
            return False


class BackgroundComputerController:
    """
    后台电脑控制器
    提供与 RealComputerController 类似的接口，但操作在后台进行
    """
    
    def __init__(self, window_title: Optional[str] = None):
        """
        初始化后台控制器
        
        Args:
            window_title: 目标窗口标题
        """
        self.controller = BackgroundController(window_title)
        self.screen_width = self.controller.window_width
        self.screen_height = self.controller.window_height
        logger.info(f"后台电脑控制器初始化，目标窗口: {window_title}")
    
    def set_target(self, window_title: Optional[str] = None, hwnd: Optional[int] = None) -> bool:
        """
        设置目标窗口
        
        Args:
            window_title: 窗口标题
            hwnd: 窗口句柄
            
        Returns:
            是否成功
        """
        if hwnd:
            result = self.controller.set_target_window(hwnd)
        else:
            result = self.controller.find_window(window_title)
        
        if result:
            self.screen_width = self.controller.window_width
            self.screen_height = self.controller.window_height
        
        return result
    
    def list_windows(self) -> List[Dict[str, Any]]:
        """列出所有窗口"""
        return self.controller.list_windows()
    
    async def take_screenshot(self) -> Dict[str, Any]:
        """
        截取目标窗口截图（后台）
        
        Returns:
            包含base64编码截图的字典
        """
        result = self.controller.take_screenshot()
        
        if result.get("success"):
            # 更新尺寸
            self.screen_width = result.get("width", self.screen_width)
            self.screen_height = result.get("height", self.screen_height)
            
            return {
                "success": True,
                "screenshot": result["screenshot"],
                "url": result.get("window_title", "Background"),
                "width": self.screen_width,
                "height": self.screen_height
            }
        
        return result
    
    async def execute_action(self, action: Dict[str, Any]) -> Dict[str, Any]:
        """
        执行后台操作
        
        Args:
            action: 操作数据
            
        Returns:
            执行结果
        """
        action_type = action.get('action') or action.get('function_name', '')
        args = {k: v for k, v in action.items() if k not in ['action', 'function_name']}
        
        logger.info(f"执行后台操作: {action_type}")
        
        try:
            if action_type == "mouse_click":
                result = self.controller.mouse_click(
                    args.get('x', 0),
                    args.get('y', 0),
                    args.get('button', 'left')
                )
            elif action_type == "mouse_double_click":
                result = self.controller.mouse_double_click(
                    args.get('x', 0),
                    args.get('y', 0),
                    args.get('button', 'left')
                )
            elif action_type == "mouse_hover":
                result = self.controller.mouse_move(
                    args.get('x', 0),
                    args.get('y', 0)
                )
            elif action_type == "mouse_drag":
                result = self.controller.mouse_drag(
                    args.get('start_x', 0),
                    args.get('start_y', 0),
                    args.get('end_x', 0),
                    args.get('end_y', 0),
                    args.get('button', 'left')
                )
            elif action_type == "mouse_scroll":
                result = self.controller.mouse_scroll(
                    args.get('scroll_x', 0),
                    args.get('scroll_y', 0)
                )
            elif action_type == "keyboard_type":
                result = self.controller.keyboard_type(
                    args.get('text', ''),
                    args.get('clear_existing', False)
                )
            elif action_type == "keyboard_press":
                result = self.controller.keyboard_press(
                    args.get('keys', [])
                )
            elif action_type == "clear_text":
                result = self.controller.keyboard_press(['ctrl', 'a'])
                time.sleep(0.05)
                result = self.controller.keyboard_press(['delete'])
            elif action_type == "click_and_type":
                result = self.controller.click_and_type(
                    args.get('x', 0),
                    args.get('y', 0),
                    args.get('text', ''),
                    args.get('clear_existing', True)
                )
            elif action_type == "wait":
                seconds = max(1, min(30, args.get('seconds', 1)))
                time.sleep(seconds)
                result = {
                    "status": "success",
                    "action": "wait",
                    "seconds": seconds,
                    "message": f"已等待 {seconds} 秒"
                }
            elif action_type == "task_complete":
                result = {
                    "status": "completed",
                    "action": "task_complete",
                    "summary": args.get('summary', ''),
                    "success": args.get('success', True),
                    "message": "任务已完成"
                }
            else:
                result = {
                    "status": "error",
                    "error": f"未知的操作类型: {action_type}"
                }
            
            return {
                "success": result.get("status") != "error",
                "message": result.get("message", ""),
                "action": action_type,
                **result
            }
            
        except Exception as e:
            logger.error(f"执行后台操作失败: {str(e)}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_screen_info(self) -> Dict[str, int]:
        """获取窗口尺寸信息"""
        return {
            "width": self.screen_width,
            "height": self.screen_height
        }
    
    def get_window_info(self) -> Dict[str, Any]:
        """获取窗口详细信息"""
        return self.controller.get_window_info()