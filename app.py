from flask import Flask, request, jsonify
import requests
import os
import json

app = Flask(__name__)
print("🚀 PPT Bot Started - Version 5 (AI Powered)")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")
OPENCLAW_API_KEY = os.environ.get("OPENCLAW_API_KEY")  # 可选，用于 AI 调用

print(f"APP_ID: {FEISHU_APP_ID}")
print(f"APP_SECRET: {'***SET***' if FEISHU_APP_SECRET else '***NOT SET***'}")

# 简单的 PPT 模板库
PPT_TEMPLATES = {
    "商务": "商务专业风格，蓝色主调，适合企业介绍、项目汇报",
    "科技": "科技感风格，深色背景，适合技术产品、创新项目",
    "简约": "极简风格，留白充足，适合概念展示、设计提案",
    "活力": "活力风格，色彩丰富，适合市场活动、团队建设",
}

# 用户会话状态存储（简单版，重启会丢失）
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

    print(f"Sending message to: {receive_id}")
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


def generate_ppt_outline(topic, style="商务"):
    """生成 PPT 大纲"""
    
    # 根据主题生成结构化大纲
    outline = f"""📊 《{topic}》PPT 大纲

**一、封面页**
- 标题：{topic}
- 副标题：专业分析与解决方案
- 风格：{PPT_TEMPLATES.get(style, PPT_TEMPLATES['商务'])}

**二、目录页**
1. 背景介绍
2. 核心内容
3. 数据分析
4. 解决方案
5. 总结展望

**三、内容详情**

【第1部分】背景介绍
- 行业现状概述
- 市场机遇分析
- 项目/产品定位

【第2部分】核心内容
- 关键概念解析
- 核心优势展示
- 差异化特点

【第3部分】数据分析
- 关键数据指标
- 趋势图表展示
- 对比分析

【第4部分】解决方案
- 具体实施步骤
- 预期效果
- 风险与应对

【第5部分】总结展望
- 核心要点回顾
- 未来发展规划
- 行动呼吁

**四、结束页**
- 感谢语
- 联系方式
- 二维码（可选）

---
💡 接下来你可以：
• 回复"生成完整PPT" - 生成详细内容
• 回复"换风格" - 选择其他模板风格
• 回复具体修改意见 - 调整大纲内容"""

    return outline


def analyze_user_intent(text):
    """分析用户意图"""
    text = text.lower()
    
    # 检查是否是主题/需求描述
    if any(word in text for word in ["ppt", "幻灯片", "大纲", "帮我做", "生成", "制作"]):
        return "create_ppt", text.replace("ppt", "").replace("幻灯片", "").strip()
    
    # 检查是否是风格选择
    if any(word in text for word in ["风格", "模板", "样式", "主题"]):
        for style in PPT_TEMPLATES.keys():
            if style in text:
                return "select_style", style
        return "list_styles", None
    
    # 检查是否是生成完整内容
    if any(word in text for word in ["完整", "详细", "全部", "生成"]):
        return "generate_full", None
    
    # 默认：获取主题
    return "get_topic", text


def handle_message(user_id, text):
    """处理用户消息，返回回复内容"""
    
    # 获取或创建用户会话
    if user_id not in user_sessions:
        user_sessions[user_id] = {"step": "welcome", "topic": None, "style": "商务"}
    
    session = user_sessions[user_id]
    intent, data = analyze_user_intent(text)
    
    print(f"User: {user_id}, Intent: {intent}, Data: {data}")
    
    # 欢迎状态
    if session["step"] == "welcome":
        if intent == "create_ppt" and data:
            session["topic"] = data
            session["step"] = "has_topic"
            return generate_ppt_outline(data, session["style"])
        else:
            return """你好！我是PPT助手 🤖

我可以帮你快速生成 PPT 大纲和内容。

请告诉我：
• 你需要什么主题的 PPT？
  例如："成都写字楼市场分析"、"产品介绍"、"年度总结"

或直接发送主题，我会立即生成大纲！"""
    
    # 已有主题
    if session["step"] == "has_topic":
        if intent == "select_style" and data:
            session["style"] = data
            return f"已切换至【{data}】风格！\n\n" + generate_ppt_outline(session["topic"], data)
        
        elif intent == "list_styles":
            styles = "\n".join([f"• {k} - {v}" for k, v in PPT_TEMPLATES.items()])
            return f"可选风格：\n{styles}\n\n请回复风格名称切换"
        
        elif intent == "create_ppt" and data:
            session["topic"] = data
            return generate_ppt_outline(data, session["style"])
        
        else:
            # 把用户输入当作新的主题
            session["topic"] = text
            return generate_ppt_outline(text, session["style"])
    
    return "收到！请告诉我你需要什么主题的 PPT？"


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
            welcome = """👋 你好！我是PPT助手 🤖

我可以帮你：
• 生成PPT大纲
• 设计内容结构  
• 推荐模板风格

请直接发送主题，例如：
"成都写字楼市场分析"
"产品介绍PPT"
"年度工作总结"

我会立即为你生成专业大纲！"""
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
    return "✅ PPT Bot (AI Powered) is running!"


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
