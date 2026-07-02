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

```bash
pip install -r requirements.txt
```

## 使用

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
