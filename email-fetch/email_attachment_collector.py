"""
READ_ME / 使用说明
1) profiles.json中的SAVE_DIR字段必须事先存在
   The SAVE_DIR field in profiles.json must point to a directory that already exists.

2) profiles.json中构建 DASL 的四个参数不能全为空
   The four parameters used to build the DASL filter in profiles.json cannot all be empty.

3) days_back其实没用上,事实上代码里写死了“今天”
   The days_back parameter is currently unused; the code is hard-coded to retrieve emails from today.

4) allowed_ext / skip_inline / mark_as_read 的功能还没写进代码
   The functionality of allowed_ext, skip_inline, and mark_as_read has not yet been implemented.
"""

import os
import win32com.client
import json
from pathlib import Path
import pyzipper
from pathlib import Path
from datetime import date

# 连接到 Outlook 应用程序的 MAPI 命名空间
# Connect to the MAPI namespace of the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 读取配置文件
# Read the configuration file
SUBJECT   = "urn:schemas:httpmail:subject"
FROMNAME  = "urn:schemas:httpmail:fromname"
FROMEMAIL = "urn:schemas:httpmail:fromemail"
BODY      = "urn:schemas:httpmail:textdescription"
READ      = "urn:schemas:httpmail:read"
PROFILES_PATH = Path("profiles.json")
profiles = json.loads(PROFILES_PATH.read_text(encoding="utf-8"))

# 构建 DASL 过滤语句的函数，过滤前四个参数：
# subject_keyword、sender_keyword、body_keyword、unread_only
# Build a DASL filter string based on the first four parameters:
# subject_keyword, sender_keyword, body_keyword, and unread_only
def esc(s):
    # DASL 字符串字面量里的单引号要写成两个
    # Single quotes in DASL string literals must be escaped as two single quotes
    return s.replace("'", "''")

def build_dasl(cfg):
    conds = []

    if cfg.get("subject_keyword"):
        conds.append(f"\"{SUBJECT}\" LIKE '%{esc(cfg['subject_keyword'])}%'")

    if cfg.get("sender_keyword"):
        kw = esc(cfg["sender_keyword"])
        conds.append(f"(\"{FROMNAME}\" LIKE '%{kw}%' OR \"{FROMEMAIL}\" LIKE '%{kw}%')")

    if cfg.get("body_keyword"):
        conds.append(f"\"{BODY}\" LIKE '%{esc(cfg['body_keyword'])}%'")

    if cfg.get("unread_only"):
        conds.append(f"\"{READ}\" = 0")

    return "@SQL=" + " AND ".join(conds) if conds else None


# 遍历每个配置项
# Iterate through each configuration profile
for profile in profiles:
    print(profile["name"])
    FOLDER_PATHS = profile["folder_path"]
    SAVE_DIR = profile["save_dir"]

    # 处理“folder_path”部分
    # Process the "folder_path" configuration
    folder = outlook.Folders(FOLDER_PATHS[0])
    for part in FOLDER_PATHS[1:]:
        folder = folder.Folders(part)

    items = folder.Items

    # 筛选出今天的邮件
    # Filter emails received today
    today = date.today().strftime("%m/%d/%Y")
    items.Sort("[ReceivedTime]", True)

    items = items.Restrict(f"[ReceivedTime] >= '{today} 00:00'")

    # 筛选出符合条件的邮件
    # Filter emails that match the configured criteria
    items = items.Restrict(build_dasl(profile))
    
    # 打印符合条件的邮件数量
    # Print the number of matching emails
    print(items.Count)
    for msg in items:
        print(msg.ReceivedTime, "|", repr(msg.Subject), "|", msg.Attachments.Count)

    # 下载附件
    # Download email attachments
    for msg in items:
        for att in msg.Attachments:
            saved = Path(SAVE_DIR) / att.FileName
            att.SaveAsFile(str(saved))
            print(msg.ReceivedTime, msg.Subject, att.FileName)

            # 处理 AES 加密的 zip 文件
            # Process AES-encrypted ZIP files
            if profile.get("extract_zip") and saved.suffix.lower() == ".zip":
                out_dir = saved.with_suffix("")
                for pw in profile.get("password") or [None]:
                    try:
                        with pyzipper.AESZipFile(saved) as z:
                            z.extractall(
                                path=out_dir,
                                pwd=pw.encode() if pw else None
                            )
                            print(f"  unzipping successfully -> {out_dir.name}/")
                            break
                    except Exception as e:
                        last = e
                else:
                    print(f"  unzipping failed: {type(last).__name__}: {last}")
