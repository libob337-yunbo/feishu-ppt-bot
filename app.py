from flask import Flask, request, jsonify
import requests
import os
import json

app = Flask(__name__)
print("🚀 PPT Bot Started - Version 3")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

print(f"APP_ID: {FEISHU_APP_ID}")
print(f"APP_SECRET: {'***SET***' if FEISHU_APP_SECRET else '***NOT SET***'}")


def get_tenant_token():
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"

    res = requests.post(url, json={
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    })

    data = res.json()
    print("TOKEN RESPONSE:", data)

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

    print(f"Sending message to: {receive_id} (type: {receive_id_type})")
    res = requests.post(
        url, 
        headers=headers, 
        json=body, 
        params={"receive_id_type": receive_id_type}
    )

    result = res.json()
    print("SEND RESULT:", json.dumps(result, ensure_ascii=False, indent=2))

    if result.get("code") == 0:
        print("✅ Message sent!")
        return True
    else:
        print(f"❌ Failed: {result.get('msg')}")
        return False


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    print("=" * 50)
    print("WEBHOOK RECEIVED:", json.dumps(data, ensure_ascii=False, indent=2))
    print("=" * 50)

    # 处理飞书验证请求
    if "challenge" in data:
        print("Challenge received:", data["challenge"])
        return jsonify({"challenge": data["challenge"]})
    
    # 获取事件类型
    event = data.get("event", {})
    event_type = event.get("type", "")
    
    # 处理用户进入私聊事件
    if event_type == "im.chat.access_event.bot_p2p_chat_entered_v1":
        sender_id = event.get("operator", {}).get("operator_id", {}).get("open_id")
        print(f"User entered chat: {sender_id}")
        
        token = get_tenant_token()
        if token and sender_id:
            welcome_text = """你好！我是PPT助手 🤖

我可以帮你：
• 生成PPT大纲
• 制作PPT内容  
• 优化PPT结构

请告诉我你需要什么主题的PPT？"""
            send_message(token, sender_id, "open_id", welcome_text)
        return jsonify({"code": 0}), 200
    
    # 处理消息事件
    if event_type == "im.message.receive_v1":
        message = event.get("message", {})

        chat_type = message.get("chat_type")
        chat_id = message.get("chat_id")
        sender = message.get("sender", {})
        sender_id = sender.get("sender_id", {}).get("open_id")

        print(f"Chat: {chat_type}, ChatID: {chat_id}, Sender: {sender_id}")

        # 获取消息内容
        content_str = message.get("content", "{}")
        try:
            content = json.loads(content_str)
            user_text = content.get("text", "")
        except:
            user_text = ""
        print(f"Message: {user_text}")

        # 获取token
        token = get_tenant_token()
        if not token:
            print("ERROR: No token")
            return jsonify({"code": 0}), 200

        # 判断回复对象
        if chat_type == "p2p":
            receive_id = sender_id
            receive_id_type = "open_id"
        else:
            receive_id = chat_id
            receive_id_type = "chat_id"

        # 回复
        reply_text = f"你好！我是PPT助手 🤖\n\n你发送了：{user_text}\n\n请告诉我你需要什么主题的PPT？"
        send_message(token, receive_id, receive_id_type, reply_text)
        return jsonify({"code": 0}), 200
    
    # 其他事件
    print(f"Unhandled event: {event_type}")
    return jsonify({"code": 0}), 200


@app.route("/")
def home():
    return "✅ PPT Bot is running!"


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
