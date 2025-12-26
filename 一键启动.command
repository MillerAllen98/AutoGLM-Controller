# 文件名：一键启动AutoGLM.py   （保存到桌面或任何地方，双击运行）
import os
import subprocess
import sys
import gradio as gr

# 自动查找 Open-AutoGLM 项目路径（支持常见位置）
possible_paths = [
    "/Users/millerallen/Open-AutoGLM",
    os.path.expanduser("~/Open-AutoGLM"),
    os.path.expanduser("~/Desktop/Open-AutoGLM"),
    "/Users/millerallen/Desktop/Open-AutoGLM",
]

project_path = None
for p in possible_paths:
    if os.path.exists(os.path.join(p, "main.py")):
        project_path = p
        break

if not project_path:
    print("没找到 Open-AutoGLM 项目！请手动把这个文件拖到项目文件夹里")
    input("按回车退出...")
    sys.exit()

os.chdir(project_path)
print(f"已进入项目目录：{project_path}")

# 激活虚拟环境（兼容zsh和bash）
activate_script = os.path.join(project_path, "venv", "bin", "activate_this.py")
if os.path.exists(activate_script):
    exec(open(activate_script).read(), {'__file__': activate_script})

# 配置
configs = {
    "智谱 BigModel": {"base_url": "https://open.bigmodel.cn/api/paas/v4", "model": "autoglm-phone", "api_key": "sk-你的智谱Key"},
    "ModelScope": {"base_url": "https://api-inference.modelscope.cn/v1", "model": "ZhipuAI/AutoGLM-Phone-9B", "api_key": ""}
}

def run(task, platform="智谱 BigModel", key=""):
    api_key = key.strip() or configs[platform]["api_key"]
    if not api_key or not task.strip():
        return "请填写完整！"
    
    cmd = f'python main.py --base-url {configs[platform]["base_url"]} --model {configs[platform]["model"]} --apikey "{api_key}" "{task.strip()}"'
    
    applescript = f'''
    tell application "Terminal"
        activate
        do script "cd \\"{project_path}\\" && source venv/bin/activate && clear && echo \'AutoGLM 正在执行任务\' && echo \'任务：{task.strip()}\' && echo \'----------------------------------------\' && {cmd}"
    end tell
    '''
    subprocess.run(["osascript", "-e", applescript])
    return f"已弹出终端执行：{task.strip()}"

with gr.Blocks() as app:
    gr.Markdown("# AutoGLM 一键控制台（双击即用版）")
    task = gr.Textbox(label="输入指令", placeholder="例：打开美团搜索附近的火锅店", lines=3)
    platform = gr.Dropdown(["智谱 BigModel", "ModelScope"], value="智谱 BigModel", label="平台")
    key = gr.Textbox(label="API Key", value=configs["智谱 BigModel"]["api_key"])
    btn = gr.Button("立刻弹出终端执行！", variant="primary")
    status = gr.Textbox(label="状态")
    
    btn.click(run, inputs=[task, platform, key], outputs=status)
    platform.change(lambda p: configs[p]["api_key"], platform, key)

app.launch(inbrowser=True, share=False)
