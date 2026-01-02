# 共享表格

多人实时共享编辑 Excel 表格的 Web 应用，专为极简的内部网络环境部署设计。

## 功能特性

- 多人实时协作编辑表格
- Excel 文件导入导出
- 历史记录与版本恢复
- 访问权限管理（密钥保护）
- 管理后台统一管理

## 技术栈

### 后端
- FastAPI - Web 框架
- aiosqlite - 异步数据库
- openpyxl - Excel 处理
- WebSocket - 实时通信

### 前端
- 原生 HTML/CSS/JavaScript
- SheetJS (xlsx) - 表格处理

## 部署与运行

### Windows 环境部署（推荐）

#### 方式一：在线快速启动
1. 确保已安装 Python 3.8+
2. 双击运行 `run.bat`
3. 自动安装依赖并启动服务

#### 方式二：离线环境部署（适合内部网络）

**准备阶段（在线环境）：**
1. 安装 Anaconda 或 Miniconda
2. 双击运行 `setup_env.bat`，自动完成：
   - 创建独立的 conda 环境（`d:\env\spreadsheet`）
   - 下载所有依赖包到本地（`d:\env\offline_packages`）
   - 配置国内镜像源

**移植到离线环境：**
1. 复制整个项目文件夹
2. 复制 `d:\env\spreadsheet` 环境目录
3. 复制 `d:\env\offline_packages` 离线包目录

**离线环境运行：**
1. 双击运行 `run.bat`
2. 脚本自动识别本地环境并启动

### Linux/macOS 环境

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 配置环境变量（可选）
cp .env.example .env
# 编辑 .env 文件修改配置

# 3. 启动服务
python run.py
```

### 配置说明

`.env` 文件配置项：
```env
ADMIN_KEY=admin123456      # 管理员密钥
SERVER_PORT=8000           # 服务器端口
SYSTEM_NAME=共享表格       # 系统名称
HISTORY_COUNT=20            # 历史记录数量
SYNC_INTERVAL=3000          # 同步间隔（毫秒）
IP_WHITELIST=21.,127.0.0.1 # IP白名单
AUTHOR_INFO=未知作者        # 作者信息
```

### 访问地址

启动后自动打开浏览器：
- 管理后台: `http://localhost:8000/admin/{ADMIN_KEY}`
- 表格访问: `http://localhost:8000/sheet/{密钥}`

局域网内其他设备可访问 `http://{服务器IP}:8000/...`

## 项目结构

```
spreadsheet/
├── backend/           # 后端服务
│   ├── main.py       # FastAPI 主应用
│   ├── database.py   # 数据库操作
│   ├── excel_handler.py  # Excel 处理
│   ├── websocket_manager.py  # WebSocket 管理
│   └── models.py     # 数据模型
├── frontend/         # 前端页面
│   ├── admin.html    # 管理后台
│   └── editor.html   # 表格编辑器
├── data/            # 数据存储
└── run.py           # 启动脚本
```

## 许可证

MIT
