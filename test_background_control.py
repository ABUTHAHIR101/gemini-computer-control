"""
后台控制功能测试脚本
测试 pywinauto/Win32 API 实现的后台窗口控制功能
"""
import sys
import os
import time

# 添加 backend 目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

def test_import():
    """测试模块导入"""
    print("=" * 60)
    print("测试 1: 模块导入")
    print("=" * 60)
    
    try:
        from tools.background_controller import (
            BackgroundController,
            BackgroundComputerController,
            HAS_WIN32,
            HAS_PYWINAUTO
        )
        print(f"✓ 模块导入成功")
        print(f"  - pywin32 可用: {HAS_WIN32}")
        print(f"  - pywinauto 可用: {HAS_PYWINAUTO}")
        
        if not HAS_WIN32:
            print("✗ 需要安装 pywin32: pip install pywin32")
            return False
        
        return True
    except ImportError as e:
        print(f"✗ 模块导入失败: {e}")
        return False

def test_list_windows():
    """测试列出窗口功能"""
    print("\n" + "=" * 60)
    print("测试 2: 列出所有可见窗口")
    print("=" * 60)
    
    try:
        from tools.background_controller import BackgroundController
        controller = BackgroundController()
        windows = controller.list_windows()
        
        print(f"✓ 找到 {len(windows)} 个可见窗口")
        print("\n前 10 个窗口:")
        for i, win in enumerate(windows[:10]):
            print(f"  [{i+1}] hwnd={win['hwnd']}, title='{win['title'][:40]}...'" if len(win['title']) > 40 else f"  [{i+1}] hwnd={win['hwnd']}, title='{win['title']}'")
        
        return windows
    except Exception as e:
        print(f"✗ 列出窗口失败: {e}")
        return []

def test_find_window(title_keyword: str):
    """测试查找窗口功能"""
    print("\n" + "=" * 60)
    print(f"测试 3: 查找包含 '{title_keyword}' 的窗口")
    print("=" * 60)
    
    try:
        from tools.background_controller import BackgroundController
        controller = BackgroundController()
        
        found = controller.find_window(title=title_keyword)
        
        if found:
            print(f"✓ 找到窗口:")
            print(f"  - 句柄: {controller.hwnd}")
            print(f"  - 标题: {controller.window_title}")
            print(f"  - 类名: {controller.window_class}")
            print(f"  - 尺寸: {controller.window_width}x{controller.window_height}")
            return controller
        else:
            print(f"✗ 未找到包含 '{title_keyword}' 的窗口")
            return None
    except Exception as e:
        print(f"✗ 查找窗口失败: {e}")
        return None

def test_screenshot(controller):
    """测试后台截图功能"""
    print("\n" + "=" * 60)
    print("测试 4: 后台截图")
    print("=" * 60)
    
    try:
        result = controller.take_screenshot()
        
        if result.get('success'):
            screenshot_data = result.get('screenshot', '')
            # 计算截图大小（base64 编码后的大小）
            size_kb = len(screenshot_data) / 1024
            print(f"✓ 后台截图成功:")
            print(f"  - 窗口: {result.get('window_title')}")
            print(f"  - 尺寸: {result.get('width')}x{result.get('height')}")
            print(f"  - 数据大小: {size_kb:.1f} KB")
            
            # 保存截图到文件
            if screenshot_data.startswith('data:image/png;base64,'):
                import base64
                img_data = base64.b64decode(screenshot_data.split(',')[1])
                with open('test_background_screenshot.png', 'wb') as f:
                    f.write(img_data)
                print(f"  - 已保存到: test_background_screenshot.png")
            
            return True
        else:
            print(f"✗ 截图失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"✗ 截图异常: {e}")
        return False

def test_mouse_click(controller, x: int, y: int):
    """测试后台鼠标点击"""
    print("\n" + "=" * 60)
    print(f"测试 5: 后台鼠标点击 ({x}, {y})")
    print("=" * 60)
    
    try:
        result = controller.mouse_click(x, y, "left")
        
        if result.get('status') == 'success':
            print(f"✓ 后台点击成功:")
            print(f"  - 位置: ({x}, {y})")
            print(f"  - 消息: {result.get('message')}")
            return True
        else:
            print(f"✗ 点击失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"✗ 点击异常: {e}")
        return False

def test_keyboard_type(controller, text: str):
    """测试后台键盘输入"""
    print("\n" + "=" * 60)
    print(f"测试 6: 后台键盘输入 '{text}'")
    print("=" * 60)
    
    try:
        result = controller.keyboard_type(text)
        
        if result.get('status') == 'success':
            print(f"✓ 后台输入成功:")
            print(f"  - 文本: {text}")
            print(f"  - 长度: {len(text)} 字符")
            return True
        else:
            print(f"✗ 输入失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"✗ 输入异常: {e}")
        return False

def test_keyboard_press(controller, keys: list):
    """测试后台按键"""
    print("\n" + "=" * 60)
    print(f"测试 7: 后台按键 {'+'.join(keys)}")
    print("=" * 60)
    
    try:
        result = controller.keyboard_press(keys)
        
        if result.get('status') == 'success':
            print(f"✓ 后台按键成功:")
            print(f"  - 按键: {result.get('keys_str')}")
            return True
        else:
            print(f"✗ 按键失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"✗ 按键异常: {e}")
        return False

def test_background_computer_controller():
    """测试 BackgroundComputerController 类"""
    print("\n" + "=" * 60)
    print("测试 8: BackgroundComputerController 集成测试")
    print("=" * 60)
    
    try:
        from tools.background_controller import BackgroundComputerController
        import asyncio
        
        # 查找一个记事本窗口进行测试
        controller = BackgroundComputerController()
        windows = controller.list_windows()
        
        # 尝试找一个记事本窗口
        notepad_win = None
        for win in windows:
            if 'notepad' in win['title'].lower() or '记事本' in win['title']:
                notepad_win = win
                break
        
        if notepad_win:
            print(f"✓ 找到记事本窗口: {notepad_win['title']}")
            controller.set_target(hwnd=notepad_win['hwnd'])
            
            # 测试异步截图
            async def async_test():
                result = await controller.take_screenshot()
                return result
            
            result = asyncio.run(async_test())
            if result.get('success'):
                print(f"✓ 异步截图成功，尺寸: {result.get('width')}x{result.get('height')}")
            
            # 测试异步执行动作
            async def async_action_test():
                action = {
                    'action': 'keyboard_type',
                    'text': 'Hello from background!'
                }
                result = await controller.execute_action(action)
                return result
            
            print("✓ BackgroundComputerController 测试完成")
            return True
        else:
            print("! 未找到记事本窗口，跳过集成测试")
            print("  提示: 打开一个记事本窗口后再运行测试")
            return True
            
    except Exception as e:
        print(f"✗ 集成测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    print("\n" + "=" * 60)
    print("后台控制功能测试")
    print("=" * 60)
    print("此测试验证 pywinauto/Win32 API 后台窗口控制功能")
    print("后台控制允许在不干扰用户操作的情况下控制窗口\n")
    
    # 测试 1: 模块导入
    if not test_import():
        print("\n测试中止: 模块导入失败")
        return
    
    # 测试 2: 列出窗口
    windows = test_list_windows()
    if not windows:
        print("\n测试中止: 无法列出窗口")
        return
    
    # 测试 3: 查找窗口（尝试查找记事本或其他常见窗口）
    test_targets = ['notepad', '记事本', 'chrome', 'edge', 'vscode']
    controller = None
    
    for target in test_targets:
        controller = test_find_window(target)
        if controller:
            break
    
    if not controller:
        print("\n! 未找到常见窗口，将使用第一个可见窗口")
        from tools.background_controller import BackgroundController
        controller = BackgroundController()
        if windows:
            controller.set_target_window(windows[0]['hwnd'])
            print(f"  使用窗口: {windows[0]['title']}")
    
    # 测试 4: 后台截图
    if controller and controller.hwnd:
        test_screenshot(controller)
    
    # 测试 8: 集成测试
    test_background_computer_controller()
    
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)
    print("\n注意事项:")
    print("1. 后台点击和键盘输入测试需要目标窗口支持消息驱动")
    print("2. 某些现代应用（如 Chrome、Electron）可能不完全支持后台操作")
    print("3. 建议使用记事本等传统 Win32 应用测试完整功能")
    print("4. 如需测试点击和输入，请打开记事本并修改测试代码")

if __name__ == '__main__':
    main()