# 🌡️ Testo 184 温度计数据读取汇总工具

跨平台桌面应用，用于逐个读取 Testo 184 系列 USB 温度计数据，并汇总导出为 Excel 文件。

## ✨ 功能特点

- **📝 测点管理** — 输入测点编号，与设备 SN 号一一对应
- **🔌 逐个读取** — 每次插入一个设备，读取完成后提示插入下一个
- **👀 数据预览** — 检测到设备后先预览数据，确认后再读取
- **💾 本地存储** — 所有数据保存在本地 SQLite 数据库
- **📊 Excel 导出** — 包含汇总数据、设备概览、统计摘要三个工作表
- **📋 历史会话** — 可回看和管理历史测试数据
- **🖥️ 跨平台** — macOS 和 Windows 均可运行

## 📋 支持型号

Testo 184 T1 / T2 / T3 / T4 / H1 / G1

## 🔧 安装

### 前置条件

- **Python 3.8+**
- **pip**（Python 包管理器）

### macOS 启动

```bash
chmod +x start_mac.sh
./start_mac.sh
```

### Windows 启动

双击 `start_win.bat` 即可

### 手动启动

```bash
pip install -r requirements.txt
python app.py
```

启动后在浏览器打开 **http://localhost:5000**

## 📖 使用流程

### 第 1 步：配置测点

1. 输入本次测试的会话名称（可选）
2. 设置测点数量，点击「生成测点」
3. 可编辑每个测点编号

### 第 2 步：逐个读取设备

1. 点击「开始读取设备」进入读取模式
2. 插入第一个 Testo 184 温度计
3. 点击「扫描设备」— 系统自动识别设备并预览数据
4. 确认后点击「确认读取此设备」
5. 看到 ✅ 完成提示后，拔出当前设备，插入下一个
6. 重复直到所有设备读取完成

### 第 3 步：数据汇总

- 查看所有测点的 SN 号和数据量
- 确认数据无误后进入导出

### 第 4 步：导出 Excel

- 点击「导出 Excel 文件」下载汇总数据
- Excel 包含三个工作表：
  - **汇总数据** — 所有设备数据合并，含测点号、SN号、温度、湿度、报警
  - **设备概览** — 各设备基本信息、数据量、时间范围
  - **统计摘要** — 每个设备的最小/最大/平均温度

## 🗂️ 项目结构

```
Testo温度计数据汇总/
├── app.py                  # Flask 后端（API + 设备检测 + 数据库）
├── requirements.txt        # Python 依赖
├── start_mac.sh           # macOS 启动脚本
├── start_win.bat          # Windows 启动脚本
├── testo_data.db          # SQLite 数据库（运行后自动生成）
├── exports/               # 导出的 Excel 文件
├── templates/
│   └── index.html         # 前端页面
├── static/
│   ├── css/style.css      # 样式
│   └── js/app.js          # 前端逻辑
└── README.md              # 本说明
```

## ❓ 常见问题

**Q: 插入 USB 后扫描不到设备？**
- 确认 Finder/资源管理器中能看到设备卷标（TESTO 184）
- 尝试先弹出再重新插入设备
- macOS 设备路径：/Volumes/TESTO 184

**Q: Windows 上双击 bat 文件没有反应？**
- 确保已安装 Python 并勾选了 "Add to PATH"
- 尝试右键 → 以管理员身份运行

**Q: 数据存在哪里？**
- 所有数据保存在项目目录下的 `testo_data.db`（SQLite 数据库）
- 导出的 Excel 文件在 `exports/` 目录下

**Q: 如何清空历史数据重新开始？**
- 在「历史会话」中可以删除旧会话
- 或直接删除 `testo_data.db` 文件

## 🔒 隐私说明

- 完全本地运行，数据不上传到任何服务器
- 不依赖云服务或外部 API
- 所有数据保存在你的电脑上

## 📝 技术栈

- **后端**: Python + Flask
- **前端**: HTML + CSS + JavaScript（原生，无需构建）
- **数据库**: SQLite
- **Excel**: openpyxl
- **设备检测**: 跨平台 USB 卷扫描
