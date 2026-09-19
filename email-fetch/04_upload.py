import shutil
import os
import json

# 读取配置文件
# load the config files
config_path = "upload_path.json"
with open(config_path, "r", encoding="utf-8") as f:
    config = json.load(f)

# 遍历每个任务
# iterate over each task
for task in config["tasks"]:
    print(task["name"])
    src_dir = task["src_dir"]
    dst_dir = task["dst_dir"]
  
    # 遍历源文件夹里的所有内容并移动
    # iterate over each path
    for item in os.listdir(src_dir):
        src_path = os.path.join(src_dir, item)
        dst_path = os.path.join(dst_dir, item)
        shutil.move(src_path, dst_path)

print("move done")
