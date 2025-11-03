# 创业孵化基地信息管理系统

## 项目简介
基于Flask开发的轻量级信息管理系统，支持创业团队/在孵企业在线提交数据、管理员后台管理与数据导出，适用于中小型孵化基地日常信息收集与统计。


## 源码获取
1. 克隆仓库：
   ```bash
   git clone https://github.com/Tgh66/manager.git
   cd manager
   ```

2. 解压项目文件：
   ```bash
   unzip D281__TGH.zip -d project
   cd project
   ```


## Linux服务器部署步骤

### 1. 环境准备
```bash
# 安装Python及依赖工具
sudo apt update && sudo apt install -y python3 python3-venv python3-pip unzip

# 进入项目目录（解压后的project文件夹）
cd project
```

### 2. 依赖安装
```bash
# 创建并激活虚拟环境
python3 -m venv venv
source venv/bin/activate

# 安装依赖包
pip install flask openpyxl pillow gunicorn
```

### 3. 数据库初始化
```bash
# 初始化SQLite数据库（自动创建用户表和管理员表）
python3 -c "from app import init_db; init_db()"
```

### 4. 启动服务
#### 开发调试（临时运行）
```bash
export FLASK_APP=app.py
flask run --host=0.0.0.0 --port=5000
```

#### 生产环境（推荐）
```bash
# 后台启动（4个进程，绑定8000端口）
nohup gunicorn -w 4 -b 0.0.0.0:8000 "app:app" &
```


## 系统访问
- **用户端**：`http://服务器IP:端口`（首次使用需注册，房间号为唯一标识）
- **管理员端**：`http://服务器IP:端口/admin/login`  
  默认账号：`admin`，默认密码：`123`（建议首次登录后修改）


## 注意事项
1. 生产环境必须修改`app.py`中的`app.secret_key`为随机安全字符串（如：`openssl rand -hex 16`生成）。
2. 确保`uploads`和`excel_files`目录有读写权限（代码会自动创建，无需手动操作）。
3. 如需停止服务：`pkill gunicorn`。
4. 数据备份：定期复制`users.db`（数据库）和`excel_files`（表单数据）目录。


## 功能说明
- **用户功能**：表单填写、历史记录查询、数据导出、上次提交数据回填。
- **管理员功能**：查看所有用户提交记录、单条/批量数据导出、用户管理。