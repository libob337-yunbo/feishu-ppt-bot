from flask import Flask, request, jsonify
import requests
import os
import json

app = Flask(__name__)
print("🚀 PPT Bot Started - Version 6 (Kimi AI Powered)")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")
KIMI_API_KEY = os.environ.get("KIMI_API_KEY")

print(f"APP_ID: {FEISHU_APP_ID}")
print(f"APP_SECRET: {'***SET***' if FEISHU_APP_SECRET else '***NOT SET***'}")
print(f"KIMI_API_KEY: {'***SET***' if KIMI_API_KEY else '***NOT SET***'}")

# 用户会话状态存储
user_sessions = {}


def get_tenant_token():
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"

    res = requests.post(url, json={
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    })

    data = res.json()

    if data.get("code") != 0:
        print("ERROR getting token:", data.get("msg"))
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

    res = requests.post(
        url, 
        headers=headers, 
        json=body, 
        params={"receive_id_type": receive_id_type}
    )

    result = res.json()

    if result.get("code") == 0:
        print("✅ Message sent!")
        return True
    else:
        print(f"❌ Failed: {result.get('msg')}")
        return False


def call_kimi(prompt):
    """调用 Kimi API 生成内容"""
    if not KIMI_API_KEY:
        return "⚠️ Kimi API Key 未配置"
    
    url = "https://api.moonshot.cn/v1/chat/completions"
    
    headers = {
        "Authorization": f"Bearer {KIMI_API_KEY}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": "moonshot-v1-8k",
        "messages": [
            {"role": "system", "content": "你是一个专业的PPT制作助手，擅长根据用户需求生成结构清晰、内容专业的PPT大纲。请用中文回复，使用Markdown格式。"},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7
    }
    
    try:
        res = requests.post(url, headers=headers, json=data, timeout=30)
        result = res.json()
        
        if "choices" in result and len(result["choices"]) > 0:
            return result["choices"][0]["message"]["content"]
        else:
            print(f"Kimi API Error: {result}")
            return "抱歉，AI 生成失败，请重试"
    except Exception as e:
        print(f"Kimi API Exception: {e}")
        return f"调用 AI 出错: {str(e)}"


def generate_ppt_with_ai(topic):
    """使用 Kimi AI 生成 PPT 大纲"""
    prompt = f"""请为"{topic}"生成一份专业的PPT大纲。

要求：
1. 包含封面、目录、内容页、结束页
2. 内容要专业、有深度
3. 每页给出标题和要点
4. 适合商务演示风格
5. 用Markdown格式输出

请直接输出PPT大纲："""

    return call_kimi(prompt)


def handle_message(user_id, text):
    """处理用户消息"""
    
    # 获取或创建用户会话
    if user_id not in user_sessions:
        user_sessions[user_id] = {"step": "welcome"}
    
    session = user_sessions[user_id]
    
    # 简单命令处理
    text_lower = text.lower().strip()
    
    if text_lower in ["帮助", "help", "?"]:
        return """🤖 PPT助手使用指南

我可以帮你：
• 生成PPT大纲 - 直接发送主题
• AI智能生成 - 发送"AI:主题"

示例：
"成都写字楼市场分析"
"AI:新能源汽车行业报告"
"年度工作总结"

直接发送主题即可开始！"""
    
    # AI 模式
    if text_lower.startswith("ai:") or text_lower.startswith("ai "):
        topic = text[3:].strip()
        if topic:
            return generate_ppt_with_ai(topic)
        else:
            return "请提供主题，例如：AI:成都写字楼市场分析"
    
    # 普通模式 - 直接生成大纲
    if len(text) > 2:
        return generate_ppt_with_ai(text)
    
    return "请告诉我你需要什么主题的PPT？"


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    # 处理飞书验证请求
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})
    
    # 获取事件类型
    header = data.get("header", {})
    event_type = header.get("event_type", "")
    event = data.get("event", {})
    
    # 处理用户进入私聊
    if event_type == "im.chat.access_event.bot_p2p_chat_entered_v1":
        sender_id = event.get("operator_id", {}).get("open_id")
        
        token = get_tenant_token()
        if token and sender_id:
            welcome = """👋 你好！我是AI PPT助手 🤖

我可以调用 Kimi AI 为你生成专业PPT大纲！

使用方法：
• 直接发送主题 → AI生成大纲
• 发送"帮助" → 查看使用指南

示例：
"成都写字楼市场分析"
"新能源汽车行业报告"
"年度工作总结"

请直接发送主题开始！"""
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

        print(f"Message from {sender_id}: {user_text}")

        # 生成回复
        reply_text = handle_message(sender_id, user_text)

        # 发送回复
        token = get_tenant_token()
        if token:
            if chat_type == "p2p":
                send_message(token, sender_id, "open_id", reply_text)
            else:
                send_message(token, chat_id, "chat_id", reply_text)
        
        return jsonify({"code": 0}), 200
    
    return jsonify({"code": 0}), 200


@app.route("/")
def home():
    return "✅ PPT Bot (Kimi AI Powered) is running!"


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
