
# AutoGLM UI 自动化工具

**首先，向 AutoGLM 项目致以最诚挚的敬意！**  
本工具基于强大的 [AutoGLM](https://github.com/zai-org/Open-AutoGLM/) 项目构建，感谢 AutoGLM 团队提供的优秀开源框架，让手机 UI 自动化变得如此简单高效。

> 请先查看 AutoGLM 官方仓库，了解底层依赖的安装和使用方式（必须先完成这些步骤，本工具才能正常运行）：  
> 👉 https://github.com/zai-org/Open-AutoGLM/

AutoGLM 提供了核心的手机控制能力（iOS/Android），本工具在其基础上封装了图形化界面、测试用例管理、执行监控和 Excel 报告导出等功能，让测试工作更直观、高效。

## 项目介绍

**AutoGLM UI Automation Controller** 是一款桌面图形化工具，专为 AutoGLM 设计，帮助用户更方便地进行手机 UI 自动化测试。

<img width="2560" height="1438" alt="ScreenShot_2025-12-26_175722_671" src="https://github.com/user-attachments/assets/643780f6-dceb-4845-83e0-9006da6fb0a2" />

### 主要功能
- 支持 iOS 和 Android 设备控制
- 通过文本框或上传 Excel/CSV 文件批量执行测试用例
- 严格模式（禁止滑动、模糊匹配，仅精确点击可见元素）
- 实时执行状态监控（用例步骤通过/失败/未执行）
- 测试结束后一键导出专业 Excel 报告（含概览数据、用例明细）

### 适用场景
- 功能回归测试
- UI 自动化脚本快速验证
- 批量用例执行与结果统计
- 演示或分享自动化测试流程

## 安装与运行

### 1. 安装 AutoGLM 核心依赖（必须先完成）
请严格按照官方教程操作：  
https://github.com/zai-org/Open-AutoGLM/

主要步骤包括：
- 安装 WebDriverAgent（iOS）
- 配置 ADB（Android）
- 获取并配置大模型 API Key（推荐智谱 AutoGLM）

### 2. 克隆并运行本项目
```bash
git clone https://github.com/MillerAllen98/AutoGLM-UI-Automation-Tool.git
cd AutoGLM-UI-Automation-Tool
```

### 3. 安装 Python 依赖
```bash
pip install pandas openpyxl
```

（其他依赖已由 AutoGLM 提供）

### 4. 运行程序
在项目根目录执行：
```bash
python autoglm_IDE.py
```

程序启动后：
1. 选择平台（iOS / Android）
2. 选择引擎（默认 ZhipuAI-AutoGLM）
3. 输入你的 API Key
4. 勾选“严格模式”（推荐开启）
5. 输入测试步骤或上传用例文件
6. 点击 **开始执行 (RUN)**

### 5. 导出报告
执行完成后，点击 **导出报告 (EXPORT)**，选择保存路径，即可生成 Excel 报告。

## 致谢
再次感谢 [AutoGLM](https://github.com/zai-org/Open-AutoGLM/) 团队，没有你们的优秀工作，这个工具不可能诞生。

欢迎 Star ⭐ 和 Fork，期待你的反馈与贡献！

— MillerAllen
