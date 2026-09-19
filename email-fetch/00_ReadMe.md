A modular workflow for collecting email attachments, cleaning up temporary files, and uploading processed files.

一个模块化的邮件附件处理流程，用于**收集邮件附件 → 清理临时文件 → 上传处理后的文件**。

## 📁 Project Structure

```text
00_ReadMe.md
│   # 本说明文档 | Project introduction and usage guide
│
├── 01_RunMe.py
│   # 流水线入口：按顺序调用 02 → 03 → 04
│   # Pipeline entry point: runs Steps 1 → 2 → 3 in sequence
│
├── 02_email_attachment_collector.py
│   # 步骤 1：收集邮件附件
│   # Step 1: Collect email attachments
│
│   └── profiles.json
│       # 邮件收集配置：定义邮件搜索与附件收集规则
│       # Configuration for email search and attachment collection rules
│
├── 03_clean_unzip_trash.py
│   # 步骤 2：清理解压过程中产生的临时文件
│   # Step 2: Clean up temporary files generated during extraction
│
│   └── clean_config.json
│       # 清理配置：定义清理模式与目标路径
│       # Configuration for cleanup modes and target paths
│
└── 04_upload.py
    # 步骤 3：上传处理后的文件
    # Step 3: Upload processed files

    └── upload_path.json
        # 上传配置：定义文件的来源路径与目标路径
        # Configuration for source and destination paths
```

## 🔄 Workflow

```text
Email
  ↓
02_email_attachment_collector.py
  ↓
Collect Attachments and unzip the zip ones
  ↓
03_clean_unzip_trash.py
  ↓
Clean Temporary / Extracted Files
  ↓
04_upload.py
  ↓
Upload Processed Files
```

The workflow is designed as separate modules so that each step can be tested, modified, or reused independently.

整个流程采用模块化设计，每个步骤都可以独立测试、修改和复用。

## ⚙️ Configuration

The workflow keeps operational settings in separate `.json` configuration files, so the Python scripts do not need to be modified for routine changes.

配置文件与程序逻辑分离，日常使用时通常只需要修改 `.json` 配置，而无需修改 Python 代码。

* `profiles.json` — Email search and attachment collection settings
* `clean_config.json` — Cleanup modes and target paths
* `upload_path.json` — Source and destination paths for file upload

## ▶️ Run

For the full workflow, run:

```bash
python 01_RunMe.py
```

You can also run each module separately when testing or troubleshooting individual steps.
