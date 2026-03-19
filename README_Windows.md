# 养老金融通报数据自动化处理系统

## Windows 打包说明

### 快速开始（推荐）

1. **获取项目文件夹**
   将整个 `pension_financial_system` 文件夹拷贝到你的 Windows 电脑

2. **双击打包脚本**
   双击运行 `一键打包_Windows.bat`

3. **等待打包完成**
   首次打包需要3-5分钟，脚本会自动安装依赖和打包

4. **找到 EXE 文件**
   打包完成后，EXE 文件位于 `dist` 文件夹中

---

### 详细说明

#### 系统要求

- Windows 10 或更高版本
- Python 3.10+（可选，脚本会自动检查）

#### 文件夹结构

```
pension_financial_system/
├── 一键打包_Windows.bat    ← 双击这个！
├── main.py                 # 主程序入口
├── requirements.txt         # 依赖列表
├── core/                   # 核心模块（勿修改）
├── utils/                  # 工具模块（勿修改）
├── gui/                    # 界面模块（勿修改）
└── dist/                   # 打包输出目录
    └── 养老金融通报数据处理系统.exe  ← 生成的EXE
```

#### 打包过程

1. 脚本自动检测 Python 环境
2. 自动安装所需依赖（openpyxl, pandas, numpy）
3. 自动安装 PyInstaller 打包工具
4. 执行打包命令生成 EXE

#### 常见问题

**Q: 提示"找不到Python"**
A: 脚本会自动提示下载 Python，请从 https://www.python.org/downloads/ 下载安装

**Q: 打包失败怎么办**
A: 请将错误信息截图发给我，我帮你解决

**Q: 生成的 EXE 在哪里**
A: 在 `dist` 文件夹中，文件名是 `养老金融通报数据处理系统.exe`

---

### 使用生成的 EXE

1. 将 Excel 模板文件（首季通报数据模板_新表样适配版.xlsx）放到 EXE 同一目录
2. 双击 EXE 启动程序
3. 按照界面提示操作即可

### 卸载

直接删除文件夹和 EXE 文件即可，无残留

---

## 技术支持

如遇问题，请联系开发者
