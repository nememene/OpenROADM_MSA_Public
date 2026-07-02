# 硬盘文件夹大小监控工具

扫描指定目录下所有文件夹的大小，导出 Excel 报表。报表按文件夹大小从大到小排序，并包含路径、父目录和层级包含关系。

## 功能

- 递归统计每个文件夹的总大小（包含子文件夹内文件）
- 输出 Excel，字段包括：
  - 排名
  - 路径
  - 文件夹名称
  - 大小（字节 / 可读格式）
  - 父目录
  - 层级深度
  - 包含关系（如 `C: > Users > Documents`）
- 自动跳过无权限访问的文件/目录

## 安装

**首次使用请先检测 Python 环境：**

- Windows：双击 `检测Python环境.bat`（会自动检测、安装依赖，无 Python 时尝试用 winget 安装）
- Linux / macOS：运行 `./检测Python环境.sh`

或手动安装：

```bash
pip install -r requirements.txt
```

图形界面使用 Python 自带的 `tkinter`。若双击后无法打开窗口：

- **Windows**：安装 Python 时勾选 `tcl/tk and IDLE`
- **Linux**：执行 `sudo apt install python3-tk`（或对应发行版的 tk 包）

## 使用

### 图形界面（双击启动）

安装依赖后，双击以下任一文件即可打开对话框，输入要扫描的文件夹地址：

- `启动文件夹大小监控.bat`（推荐，Windows）
- `folder_size_monitor.pyw`（Windows，无黑色命令行窗口）
- `启动文件夹大小监控.sh`（Linux / macOS）
- `folder_size_monitor_gui.py`（各平台均可运行）

对话框支持：

- 手动输入文件夹路径
- 点击「浏览...」选择文件夹
- 自定义 Excel 输出位置
- 扫描完成后自动提示并打开结果所在目录

### 命令行

```bash
# 扫描当前目录
python folder_size_monitor.py

# 扫描指定磁盘或目录
python folder_size_monitor.py /home

# 指定输出文件
python folder_size_monitor.py D:\ -o D:\folder_report.xlsx
```

## 输出说明

- **文件夹大小** 工作表：所有文件夹明细，按大小降序排列
- **扫描信息** 工作表：扫描根目录、文件夹总数、总占用空间、生成时间
