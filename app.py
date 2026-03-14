from flask import Flask, request, jsonify
import requests
import os
import json
import time

app = Flask(__name__)
print("🚀 PPT Bot Started - Version 7 (Workflow)")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")
KIMI_API_KEY = os.environ.get("KIMI_API_KEY")
OPENCLAW_WEBHOOK_URL = os.environ.get("OPENCLAW_WEBHOOK_URL")  # 可选，用于通知主机器人

print(f"APP_ID: {FEISHU_APP_ID}")
print(f"KIMI_API_KEY: {'***SET***' if KIMI_API_KEY else '***NOT SET***'}")

# 用户会话状态存储
user_sessions = {}

# 步骤定义
STEP_TOPIC = "topic"           # 等待输入主题
STEP_OUTLINE = "outline"       # 已生成大纲，等待审核
STEP_DETAIL = "detail"         # 已生成详细内容，等待确认
STEP_GENERATING = "generating" # 正在生成PPT
STEP_COMPLETE = "complete"     # 完成


def get_tenant_token():
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    res = requests.post(url, json={
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    })
    data = res.json()
    if data.get("code") != 0:
        return None
    return data.get("tenant_access_token")


def send_message(token, receive_id, receive_id_type, text):
    """发送消息"""
    url = "https://open.feishu.cn/open-apis/im/v1/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = {
        "receive_id": receive_id,
        "msg_type": "text",
        "content": json.dumps({"text": text}, ensure_ascii=False)
    }
    res = requests.post(url, headers=headers, json=body, params={"receive_id_type": receive_id_type})
    return res.json().get("code") == 0


def call_kimi(prompt, system_prompt=None):
    """调用 Kimi API"""
    if not KIMI_API_KEY:
        return "⚠️ Kimi API Key 未配置"
    
    url = "https://api.moonshot.cn/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {KIMI_API_KEY}",
        "Content-Type": "application/json"
    }
    
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})
    
    data = {
        "model": "moonshot-v1-8k",
        "messages": messages,
        "temperature": 0.7
    }
    
    try:
        res = requests.post(url, headers=headers, json=data, timeout=60)
        result = res.json()
        if "choices" in result and len(result["choices"]) > 0:
            return result["choices"][0]["message"]["content"]
        return f"AI 生成失败: {result}"
    except Exception as e:
        return f"调用出错: {str(e)}"


def generate_outline(topic):
    """生成大纲"""
    prompt = f"""请为"{topic}"生成一份专业的PPT大纲。

要求：
1. 包含封面、目录、内容页（3-5页）、结束页
2. 每页给出：标题 + 3-5个要点
3. 内容专业、有深度
4. 用Markdown格式，清晰易读

请直接输出PPT大纲："""
    
    return call_kimi(prompt, "你是一个专业的PPT制作专家，擅长生成结构清晰的PPT大纲。")


def generate_detail_content(outline, topic):
    """根据大纲生成详细内容"""
    prompt = f"""请根据以下大纲，为"{topic}"生成详细的PPT内容。

大纲：
{outline}

要求：
1. 为每一页生成完整的演讲稿内容
2. 每页内容适合演讲3-5分钟
3. 包含关键数据、案例或论据
4. 语言专业、流畅
5. 用Markdown格式，每页之间用---分隔

请生成详细内容："""
    
    return call_kimi(prompt, "你是一个专业的PPT内容撰写专家，擅长将大纲扩展为完整的演讲内容。")


def notify_openclaw(user_id, topic, content, content_type="outline"):
    """通知主机器人审核（可选）"""
    if not OPENCLAW_WEBHOOK_URL:
        return False
    
    try:
        payload = {
            "user_id": user_id,
            "topic": topic,
            "content": content,
            "type": content_type,
            "timestamp": int(time.time())
        }
        requests.post(OPENCLAW_WEBHOOK_URL, json=payload, timeout=5)
        return True
    except:
        return False


def handle_message(user_id, text, chat_id, chat_type):
    """处理用户消息，返回回复内容和下一步操作"""
    
    # 获取或创建会话
    if user_id not in user_sessions:
        user_sessions[user_id] = {
            "step": STEP_TOPIC,
            "topic": None,
            "outline": None,
            "detail": None,
            "chat_id": chat_id,
            "chat_type": chat_type
        }
    
    session = user_sessions[user_id]
    text_lower = text.lower().strip()
    
    # 帮助命令
    if text_lower in ["帮助", "help", "?"]:
        return """🤖 PPT助手使用指南

【工作流程】
1️⃣ 发送主题 → 生成大纲
2️⃣ 回复"确认"或"修改意见" → 审核大纲
3️⃣ 回复"确认" → 生成详细内容
4️⃣ 回复"确认" → 生成PPT文件

【可用命令】
• 直接发送主题 → 开始制作
• "确认" → 进入下一步
• "重新生成" → 重新生成当前步骤
• "取消" → 重新开始
• "帮助" → 查看指南

请发送主题开始！""", None
    
    # 取消命令
    if text_lower in ["取消", "cancel", "重新开始"]:
        user_sessions[user_id] = {
            "step": STEP_TOPIC,
            "topic": None,
            "outline": None,
            "detail": None,
            "chat_id": chat_id,
            "chat_type": chat_type
        }
        return "已重置，请发送新的PPT主题", None
    
    # 步骤1: 等待主题
    if session["step"] == STEP_TOPIC:
        if len(text) < 2:
            return "主题太短了，请详细描述一下你的PPT主题", None
        
        session["topic"] = text
        session["step"] = STEP_OUTLINE
        
        # 生成大纲
        outline = generate_outline(text)
        session["outline"] = outline
        
        # 通知主机器人审核（可选）
        notify_openclaw(user_id, text, outline, "outline")
        
        reply = f"""📋 已生成大纲：《{text}》

{outline}

---
✅ 请审核以上大纲

回复：
• "确认" → 生成详细内容
• "修改：xxx" → 告诉我修改意见，我会重新生成
• "重新生成" → 重新生成大纲
• "取消" → 重新开始"""
        
        return reply, None
    
    # 步骤2: 大纲审核
    if session["step"] == STEP_OUTLINE:
        if text_lower in ["确认", "确定", "ok", "yes", "通过"]:
            session["step"] = STEP_DETAIL
            
            # 生成详细内容
            detail = generate_detail_content(session["outline"], session["topic"])
            session["detail"] = detail
            
            # 通知主机器人
            notify_openclaw(user_id, session["topic"], detail, "detail")
            
            reply = f"""📝 已生成详细内容：《{session['topic']}》

{detail[:2000]}...

（内容较长，以上为预览）

---
✅ 请审核以上内容

回复：
• "确认" → 开始生成PPT文件
• "重新生成" → 重新生成详细内容
• "取消" → 重新开始"""
            
            return reply, None
        
        elif text_lower in ["重新生成", "再来一次", "regenerate"]:
            # 重新生成大纲
            outline = generate_outline(session["topic"])
            session["outline"] = outline
            
            reply = f"""📋 重新生成的大纲：《{session['topic']}》

{outline}

---
✅ 请审核以上大纲

回复"确认"继续，或"修改"告诉我意见"""
            return reply, None
        
        elif text_lower.startswith("修改") or text_lower.startswith("优化"):
            # 根据修改意见重新生成
            feedback = text[2:].strip() if text_lower.startswith("修改") else text[2:].strip()
            prompt = f"""请根据以下修改意见，重新生成PPT大纲。

原主题：{session['topic']}
修改意见：{feedback}

请重新生成大纲："""
            
            outline = call_kimi(prompt, "你是一个专业的PPT制作专家。")
            session["outline"] = outline
            
            reply = f"""📋 根据你的意见修改后的大纲：

{outline}

---
回复"确认"继续，或继续提修改意见"""
            return reply, None
        
        else:
            return "请回复'确认'通过大纲，或'修改:xxx'告诉我修改意见", None
    
    # 步骤3: 详细内容审核
    if session["step"] == STEP_DETAIL:
        if text_lower in ["确认", "确定", "ok", "yes", "通过"]:
            session["step"] = STEP_GENERATING
            
            reply = f"""🎨 正在生成PPT文件：《{session['topic']}》

请稍等，正在：
1. 优化排版格式
2. 生成PPT文件
3. 上传文件

（预计需要10-30秒）"""
            
            return reply, "generate_ppt"
        
        elif text_lower in ["重新生成", "再来一次"]:
            # 重新生成详细内容
            detail = generate_detail_content(session["outline"], session["topic"])
            session["detail"] = detail
            
            reply = f"""📝 重新生成的详细内容：

{detail[:1500]}...

回复"确认"继续生成PPT，或"取消"重新开始"""
            return reply, None
        
        else:
            return "请回复'确认'生成PPT，或'重新生成'重新创建内容", None
    
    # 步骤4: 生成中
    if session["step"] == STEP_GENERATING:
        return "正在生成PPT中，请稍等...", None
    
    # 默认
    return "请发送PPT主题开始制作", None


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    # 处理飞书验证
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})
    
    header = data.get("header", {})
    event_type = header.get("event_type", "")
    event = data.get("event", {})
    
    # 用户进入私聊
    if event_type == "im.chat.access_event.bot_p2p_chat_entered_v1":
        sender_id = event.get("operator_id", {}).get("open_id")
        token = get_tenant_token()
        if token and sender_id:
            welcome = """👋 你好！我是AI PPT助手 🤖

我可以帮你：
• AI生成专业大纲
• 生成详细演讲内容  
• 制作完整PPT文件

【使用方法】
直接发送主题，例如：
"成都写字楼市场分析"

我会：
1️⃣ 生成大纲 → 请你审核
2️⃣ 生成内容 → 请你确认
3️⃣ 生成PPT → 发送给你

发送"帮助"查看详细指南！"""
            send_message(token, sender_id, "open_id", welcome)
        return jsonify({"code": 0}), 200
    
    # 处理消息
    if event_type == "im.message.receive_v1":
        message = event.get("message", {})
        chat_type = message.get("chat_type")
        chat_id = message.get("chat_id")
        sender = message.get("sender", {})
        sender_id = sender.get("sender_id", {}).get("open_id")

        # 获取消息内容
        content_str = message.get("content", "{}")
        try:
            content = json.loads(content_str)
            user_text = content.get("text", "").strip()
        except:
            user_text = ""

        print(f"[{sender_id}]: {user_text}")

        # 处理消息
        reply_text, action = handle_message(sender_id, user_text, chat_id, chat_type)

        # 发送回复
        token = get_tenant_token()
        if token:
            if chat_type == "p2p":
                send_message(token, sender_id, "open_id", reply_text)
            else:
                send_message(token, chat_id, "chat_id", reply_text)
            
            # 如果需要生成PPT（异步处理）
            if action == "generate_ppt":
                # 这里会调用本地API生成PPT
                # 暂时先发送提示
                time.sleep(2)
                send_message(token, sender_id, "open_id", 
                    "⚠️ PPT生成功能需要连接本地服务。\n\n目前只能提供文字版内容。\n\n如需完整PPT，请复制以上内容，使用本地PPT工具生成。")
                
                # 重置会话
                if sender_id in user_sessions:
                    user_sessions[sender_id]["step"] = STEP_TOPIC
        
        return jsonify({"code": 0}), 200
    
    return jsonify({"code": 0}), 200


@app.route("/")
def home():
    return "✅ PPT Bot (Workflow) is running!"


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


# 本地API接口（用于生成PPT）
@app.route("/api/generate-ppt", methods=["POST"])
def api_generate_ppt():
    """接收本地服务调用，生成PPT"""
    data = request.json
    user_id = data.get("user_id")
    topic = data.get("topic")
    detail = data.get("detail")
    
    # 这里会调用本地PPT生成工具
    # 返回PPT文件URL或base64
    
    return jsonify({
        "code": 0,
        "message": "PPT生成任务已接收",
        "task_id": f"ppt_{int(time.time())}"
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
