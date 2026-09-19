import subprocess, sys

# 运行所有脚本-相对路径，即所有脚本和设置在同一文件夹
# run all 3 scripts - relative path, means when all scripts and configs are in one same folder
for script in ["02_email_attachment_collector.py", "03_clean_unzip_trash.py", "04_upload.py"]:
    subprocess.run([sys.executable, script], check=True)
