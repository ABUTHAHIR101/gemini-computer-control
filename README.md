# 🤖 Gemini Computer Use

基于 **Gemini 3/Flash** 系列模型构建的智能计算机控制系统。本项目提供了一个美观、统一的控制台，支持通过自然语言指令操作浏览器、物理桌面及后台窗口。

![Dashboard Preview](frontend/assets/logo.png)

## 🌟 核心特性

- 🎨 **统一控制台**：全新的 MD3 + Tailwind CSS 仪表盘，单页面管理所有控制模式。
- 🎭 **浏览器自动化**：集成 Playwright，支持无头/有头模式，实时标签页管理。
- 🖥️ **物理桌面控制**：直接操控 OS 层级，支持中文输入，具备安全停止机制。
- 🌙 **后台窗口控制**：基于 Win32 API，在不干扰用户当前工作的前提下操控后台应用。
- 🤖 **智能 Agent**：支持“单步调试”与“自动循环”模式，AI 会自主决策并执行多步任务。
- 🧠 **思维可视化**：实时展示 AI 的推理过程（Reasoning）与执行轨迹（Timeline）。
- 🖼️ **高密度 UI**：针对 100% 缩放优化的紧凑型设计，所有关键信息一屏尽览。

## 🏗️ 项目架构

### 后端技术栈 (Python 3.8+)
- **Flask**: 提供 RESTful API 与前端交互。
- **Google GenAI**: 核心 AI 模型调用，支持函数调用（Function Calling）。
- **Playwright**: 负责无头浏览器环境下的截图与精准操作执行。
- **PyAutoGUI**: 用于物理桌面的鼠标键盘模拟（支持中文输入增强）。
- **pywin32 (Windows)**: 通过 Win32 API 实现后台窗口句柄操作与消息发送。
- **SSE (Server-Sent Events)**: 实现 Agent 执行过程中的实时截图与状态推送。

### 核心后端 API
- `POST /analyze`: 基础图像分析，返回建议的操作指令。
- `POST /agent/start`: 启动 Agent 任务（单步或自动模式）。
- `GET /agent/events/<session_id>`: SSE 流端点，实时监听执行过程。
- `POST /playwright/launch`: 启动隔离的浏览器环境。
- `GET /background/windows`: 获取系统中可操作的窗口句柄列表。

## 🛠️ AI 工具系统实现

系统通过「归一化坐标系统 (0-1000)」解决了屏幕分辨率适配问题。AI 在分析时不需要关注实际像素，而是返回百分比位置。后端自动根据当前的屏幕/窗口/浏览器视口尺寸进行坐标转换。

### 内置工具箱 (11 种核心工具)
- **Mouse**: `mouse_click`, `mouse_double_click`, `mouse_hover`, `mouse_drag`, `mouse_scroll`
- **Keyboard**: `keyboard_type` (通过剪贴板支持中文), `keyboard_press` (支持组合键), `clear_text`
- **Composite**: `click_and_type` (组合操作提高效率)
- **Control**: `wait` (智能等待UI刷新), `task_complete` (任务状态反馈)

## ⚠️ 注意事项

1. **分辨率适配**：系统采用 0-1000 归一化坐标，支持任意屏幕尺寸。
2. **安全停机**：在物理桌面模式下，若遇到紧急情况，请将鼠标快速移动至**屏幕四个角落**即可强制停止。
3. **兼容性**：后台模式主要针对传统 Win32 应用（如记事本、专业软件），现代 Electron 应用兼容性有限。

## 📜 许可证
[MIT License](LICENSE)
