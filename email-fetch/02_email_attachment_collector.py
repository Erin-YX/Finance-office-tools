"""
READ_ME
1) profiles.json中的SAVE_DIR字段必须事先存在
   The SAVE_DIR field in profiles.json must point to a folder that already exists.
   
   
2) profiles.json中构建 DASL 的四个参数不能全为空
   At least one of the four DASL filtering parameters in profiles.json
   ("subject_keyword", "sender_keyword", "body_keyword", "unread_only")
   must be provided; they cannot all be empty.
   
3) days_back其实没用上,事实上代码里写死了”今天“
   The days_back parameter is not currently used. The code currently uses
   a hard-coded date rule based on today's date (with a special case for
   the "Customer Conversion Blotter s2bx" profile).
   
4) mark_as_read 的功能还没写进代码
   The mark_as_read functionality has not yet been implemented.
"""

import os
import win32com.client
import json
from pathlib import Path
import pyzipper
from pathlib import Path
from datetime import date
from datetime import date, timedelta

# 01 连接到 Outlook 应用程序的 MAPI 名空间
# 01 Connect to Outlook and access its MAPI namespace
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 02 读取配置文件,可以开发不同业务用途的配置文件
# 02 Load the configuration file. Different configuration files can be created for different business purposes.
PROFILES_PATH = Path("profiles_recon.json")
profiles = json.loads(PROFILES_PATH.read_text(encoding="utf-8"))

# 03 构建 DASL 过滤语句的函数
# 03 Build the DASL filter query.
#     The first four parameters are used to filter emails:
#     "subject_keyword", "sender_keyword", "body_keyword", and "unread_only"

# Define DASL field names
SUBJECT   = "urn:schemas:httpmail:subject"
FROMNAME  = "urn:schemas:httpmail:fromname"
FROMEMAIL = "urn:schemas:httpmail:fromemail"
BODY      = "urn:schemas:httpmail:textdescription"
READ      = "urn:schemas:httpmail:read"

def esc(s):
    return s.replace("'", "''")# Single quotes in DASL string literals must be escaped by doubling them

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

# 04 遍历每个配置项，先筛选邮件，再筛选附件
# 04 Iterate through each profile.
#     For each profile, filter the emails first, then filter and process their attachments.
for profile in profiles:
    # Start filtering emails
    print(profile["name"])
    FOLDER_PATHS = profile["folder_path"]
    SAVE_DIR = profile["save_dir"]

    # Process the folder_path and navigate through the Outlook folder hierarchy
    folder = outlook.Folders(FOLDER_PATHS[0])
    for part in FOLDER_PATHS[1:]:
        folder = folder.Folders(part)

    items = folder.Items

    # Filter emails received today.
    target_date = date.today()
    today = target_date.strftime("%m/%d/%Y")
    items.Sort("[ReceivedTime]", True)
    items = items.Restrict(f"[ReceivedTime] >= '{today} 00:00'")

    # Filter emails that match the configured criteria
    items = items.Restrict(build_dasl(profile))
    
    # Print the number of matching emails
    print(items.Count)
    for msg in items:
        print(msg.ReceivedTime, "|", repr(msg.Subject), "|", msg.Attachments.Count)
 
    # Start filtering attachments
    ALLOWED_EXT = profile.get("allowed_ext")  # None means no file-extension restriction
    SKIP_INLINE = profile.get("skip_inline", False)
    MARK_AS_READ = profile.get("mark_as_read", False)
    # Filter attachments
    for msg in items:
        for att in msg.Attachments:
            
            # Skip inline attachments, such as signature images
            if SKIP_INLINE:
                try:
                    # PR_ATTACHMENT_HIDDEN: inline attachments are typically marked as hidden
                    is_hidden = att.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
                    )
                except Exception:
                    is_hidden = False
                if is_hidden:
                    continue

            # Download only attachments with the specified file extensions
            ext = Path(att.FileName).suffix.lower()
            if ALLOWED_EXT and ext not in [e.lower() for e in ALLOWED_EXT]:
                continue

            # Save the attachment
            saved = Path(SAVE_DIR) / att.FileName
            att.SaveAsFile(str(saved))
            print(msg.ReceivedTime, msg.Subject, att.FileName)

            # If the attachment is a ZIP file, extract it.
            #     Encrypted ZIP files are also supported when the correct password is provided.
            if profile.get("extract_zip") and saved.suffix.lower() == ".zip":
                out_dir = saved.with_suffix("")
                for pw in profile.get("password") or [None]:
                    try:
                         with pyzipper.AESZipFile(saved) as z:
                            z.extractall(path=out_dir, pwd=pw.encode() if pw else None)
                            print(f"  unzipping successfully -> {out_dir.name}/")
                            break
                    except Exception as e:
                        last = e
                else:
                    print(f"  unzipping failed: {type(last).__name__}: {last}")
