# 飞书 PPT 助手机器人

一个部署在 Render 上的飞书机器人，帮助用户生成 PPT 内容。

## 部署步骤

### 1. 创建 Render 服务

1. 登录 [Render Dashboard](https://dashboard.render.com)
2. 点击 **New → Web Service**
3. 连接你的 GitHub/GitLab 仓库
4. 配置如下：
   - **Name**: feishu-ppt-bot (或你喜欢的名字)
   - **Region**: Singapore (离中国最近)
   - **Branch**: main
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`

### 2. 配置环境变量

在 Render Dashboard → 你的服务 → Environment 中添加：

```
FEISHU_APP_ID=cli_a925251e4b611bc6
FEISHU_APP_SECRET=你的app_secret
```

### 3. 配置飞书 Webhook

1. 登录 [飞书开发者平台](https://open.feishu.cn/app)
2. 进入你的应用 → 事件与回调
3. 配置请求地址：
   ```
   https://你的服务名.onrender.com/webhook
   ```
4. 订阅以下事件：
   - `im.message.receive_v1` (接收消息)
   - `im.chat.access_event.bot_p2p_chat_entered_v1` (用户进入私聊)

### 4. 发布应用

1. 飞书开发者平台 → 版本管理与发布
2. 创建版本并发布
3. 在飞书里搜索你的机器人，开始使用！

## 测试

部署完成后，访问以下地址检查服务状态：
- 首页: `https://你的服务名.onrender.com/`
- 健康检查: `https://你的服务名.onrender.com/health`

## 本地开发

```bash
# 安装依赖
pip install -r requirements.txt

# 设置环境变量
export FEISHU_APP_ID=cli_xxx
export FEISHU_APP_SECRET=xxx

# 运行
python app.py
```

## 文件说明

- `app.py` - 主程序
- `requirements.txt` - Python 依赖
- `.gitignore` - Git 忽略文件（已包含 .env）
