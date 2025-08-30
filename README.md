# WPS Excel 图片修复工具

**将WPS Office的`=_xlfn.DISPIMG()`公式转换为原生Excel图片对象**

## 📋 项目背景

### 问题来源
WPS Office创建的Excel文件使用特殊的`=_xlfn.DISPIMG("图片ID")`公式来嵌入图片，这种格式在Microsoft Excel中会出现以下问题：

- 图片无法正常显示
- 公式计算错误
- 文件兼容性差
- 图片位置混乱

## 🔧 技术原理

### 文件结构分析
Excel文件(.xlsx)本质上是一个ZIP压缩包，包含：
```
test.xlsx/
├── xl/
│   ├── worksheets/          # 工作表数据
│   ├── cellimages.xml       # WPS图片定义
│   ├── _rels/
│   │   └── cellimages.xml.rels  # 图片关系映射
│   └── media/              # 实际图片文件
└── [Content_Types].xml
```

### 核心转换逻辑
1. **识别阶段**：扫描所有包含`=_xlfn.DISPIMG()`的单元格
2. **映射阶段**：解析cellimages.xml.rels获取图片路径映射
3. **提取阶段**：从xlsx中提取原始图片数据
4. **计算阶段**：根据单元格尺寸计算最佳显示尺寸
5. **替换阶段**：删除公式，插入原生Excel图片对象

### 通用尺寸计算
```python
# 基于Excel标准像素密度的通用计算
# 去除表格名判断，适用于所有工作表
cell_width_px = max(int(column_width * 7.5), 60)
cell_height_px = max(int(row_height * 4.5), 40)
```

### 等比例缩放算法
```python
# 保持原始比例计算
original_ratio = image_width / image_height
max_width = cell_width_px * 0.9  # 留10%边距
max_height = cell_height_px * 0.9

scale_factor = min(max_width/image_width, max_height/image_height)
final_width = int(image_width * scale_factor)
final_height = int(image_height * scale_factor)
```

## 🚀 功能特性

### ✅ 已完成特性
- **拖拽操作**：支持文件拖拽到程序图标
- **进度显示**：实时显示修复进度和状态
- **自动定位**：图片精确放置在原始单元格
- **比例保持**：智能等比例缩放，避免变形
- **批量处理**：支持多个工作表同时处理
- **自动保存**：生成新文件，保留原始文件
- **自动打开**：修复完成后自动打开文件

### 🎯 精确修复
- **工作表识别**：根据工作表名称智能选择尺寸计算规则
- **像素级定位**：使用Excel的EMU(English Metric Unit)系统
- **居中显示**：计算最佳偏移量，确保图片居中
- **安全边界**：防止图片超出单元格边界


## 🛠️ 使用方法

### 方式一：拖拽使用（推荐）
1. 下载 `WPS_Excel_Repair_Tool.exe`
2. 将需要修复的.xlsx文件拖拽到程序图标上
3. 等待进度窗口完成修复
4. 修复完成后自动打开新文件

### 方式二：命令行使用
```bash
python wps_repair_dragdrop.py your_file.xlsx
```

### 方式三：代码集成
```python
from wps_excel_fixer_precise_safe import PreciseSafeWPSExcelFixer

fixer = PreciseSafeWPSExcelFixer('input.xlsx')
fixed_file = fixer.fix_excel_file_precise_safe('output.xlsx')
```

## 📊 性能表现

| 文件大小 | 工作表数量 | 图片数量 | 修复时间 |
|---------|------------|----------|----------|
| 8.7MB   | 16         | 100+     | ~15秒    |
| 2MB     | 5          | 20       | ~3秒     |
| 500KB   | 2          | 5        | ~1秒     |

**优化措施：**
- 延迟加载：按需加载工作表
- 进度显示：避免用户误以为程序卡死
- 内存管理：及时释放大文件内存

## 🔍 调试与故障排除

### 常见问题
1. **文件损坏警告**：确保使用最新版本
2. **图片位置偏移**：检查单元格尺寸计算
3. **图片变形**：确认等比例缩放算法

### 日志输出
程序会在修复过程中输出详细日志：
```
=== 精确安全WPS图片修复工具 ===
正在分析所有工作表中的DISPIMG单元格...
总共发现 45 个DISPIMG公式需要修复



**注意**：此工具专门针对WPS Office创建的Excel文件，Microsoft Excel原生创建的文件不需要此工具。
