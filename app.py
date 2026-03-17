# AI 驱动的 PPT 助手 - 新架构
from flask import Flask, request, jsonify
import requests
import os
import json
import time
import threading
from ppt_generator import generate_ppt_file

app = Flask(__name__)
print("🚀 AI PPT Bot Started")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")
KIMI_API_KEY = os.environ.get("KIMI_API_KEY")

# 对话历史存储
conversations = {}
MAX_HISTORY = 200

# 系统提示词
SYSTEM_PROMPT = """你是专业的PPT制作助手，擅长：
1. 根据用户需求设计PPT结构和内容
2. 生成专业、有深度的商业PPT
3. 灵活调整，满足用户的个性化需求
4. 保持对话流畅，像专业顾问一样交流

工作流程：
1. 了解用户需求和偏好
2. 设计大纲并征求确认
3. 生成详细内容并展示预览
4. 根据反馈调整优化
5. 生成最终PPT文件

原则：
- 主动询问，确保理解用户需求
- 提供选项，让用户做决策
- 随时接受修改，不限制修改次数
- 保持专业但友好的语气

当前状态：{state}
"""

# 用户状态存储
user_states = {}

def get_conversation_key(chat_id, user_id):
    return f"{chat_id}:{user_id}"

def get_conversation(conv_key):
    if conv_key not in conversations:
        conversations[conv_key] = []
    return conversations[conv_key]

def add_message(conv_key, role, content):
    conv = get_conversation(conv_key)
    conv.append({"role": role, "content": content, "time": time.time()})
    # 保留最近 MAX_HISTORY 条
    if len(conv) > MAX_HISTORY:
        conversations[conv_key] = conv[-MAX_HISTORY:]

def get_state(conv_key):
    return user_states.get(conv_key, {
        "topic": None,
        "outline": None,
        "detail": None,
        "ppt_path": None
    })

def update_state(conv_key, **kwargs):
    state = get_state(conv_key)
    state.update(kwargs)
    user_states[conv_key] = state

def call_kimi(conv_key, user_message, state):
    """调用 Kimi API 进行对话"""
    if not KIMI_API_KEY:
        return "⚠️ Kimi API Key 未配置"
    
    # 构建消息历史
    messages = [{"role": "system", "content": SYSTEM_PROMPT.format(state=json.dumps(state, ensure_ascii=False))}]
    
    # 添加历史对话
    conv = get_conversation(conv_key)
    for msg in conv:
        messages.append({"role": msg["role"], "content": msg["content"]})
    
    # 添加当前消息
    messages.append({"role": "user", "content": user_message})
    
    url = "https://api.moonshot.cn/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {KIMI_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "moonshot/kimi-k2.5",
        "messages": messages,
        "temperature": 0.7
    }
    
    try:
        res = requests.post(url, headers=headers, json=data, timeout=120)
        result = res.json()
        if "choices" in result and len(result["choices"]) > 0:
            return result["choices"][0]["message"]["content"]
        else:
            return f"AI 响应错误: {result}"
    except Exception as e:
        return f"调用出错: {str(e)}"

def handle_ai_response(conv_key, ai_response, chat_id, token):
    """处理 AI 响应，执行相应操作"""
    # 检查 AI 响应中是否包含特殊指令
    # 例如：[GENERATE_PPT] 表示要生成PPT文件
    
    if "[GENERATE_PPT]" in ai_response:
        # 提取实际回复内容（去掉指令标记）
        reply = ai_response.replace("[GENERATE_PPT]", "").strip()
        # 发送回复
        send_message(token, chat_id, "chat_id", reply)
        # 异步生成PPT
        def generate():
            state = get_state(conv_key)
            if state.get("outline"):
                ppt_path = generate_ppt_file(
                    state.get("topic", "PPT"),
                    state.get("outline", ""),
                    state.get("detail", ""),
                    output_dir="/tmp"
                )
                file_key = upload_file(token, ppt_path)
                if file_key:
                    send_file(token, chat_id, "chat_id", file_key)
                    update_state(conv_key, ppt_path=ppt_path)
        threading.Thread(target=generate).start()
        return True
    
    return False

# 飞书相关函数（保持原有实现）
def get_tenant_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    try:
        res = requests.post(url, json={
            "app_id": FEISHU_APP_ID,
            "app_secret": FEISHU_APP_SECRET
        }, timeout=10)
        data = res.json()
        if data.get("code") == 0:
            return data.get("tenant_access_token")
    except:
        pass
    return None

def send_message(token, receive_id, receive_id_type, text):
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
    try:
        requests.post(url, headers=headers, json=body, 
                     params={"receive_id_type": receive_id_type}, timeout=10)
    except:
        pass

def upload_file(token, file_path):
    url = "https://open.feishu.cn/open-apis/im/v1/files"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        with open(file_path, 'rb') as f:
            files = {'file': (os.path.basename(file_path), f, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')}
            data = {'file_type': 'pptx', 'file_name': os.path.basename(file_path)}
            res = requests.post(url, headers=headers, files=files, data=data, timeout=30)
        result = res.json()
        if result.get("code") == 0:
            return result.get("data", {}).get("file_key")
    except:
        pass
    return None

def send_file(token, receive_id, receive_id_type, file_key):
    url = "https://open.feishu.cn/open-apis/im/v1/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    body = {
        "receive_id": receive_id,
        "msg_type": "file",
        "content": json.dumps({"file_key": file_key}, ensure_ascii=False)
    }
    try:
        res = requests.post(url, headers=headers, json=body, 
                          params={"receive_id_type": receive_id_type}, timeout=10)
        return res.json().get("code") == 0
    except:
        return False

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    
    # 处理飞书验证
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})
    
    header = data.get("header", {})
    event_type = header.get("event_type", "")
    event = data.get("event", {})
    
    if event_type != "im.message.receive_v1":
        return jsonify({"code": 0}), 200
    
    message = event.get("message", {})
    chat_type = message.get("chat_type")
    chat_id = message.get("chat_id")
    
    # 群聊需要被@才处理
    if chat_type != "p2p":
        content_str = message.get("content", "{}")
        try:
            content = json.loads(content_str)
            user_text = content.get("text", "")
        except:
            user_text = ""
        
        # 检查是否被@（需要根据实际情况调整）
        if "@" not in user_text:
            return jsonify({"code": 0}), 200
        
        # 去掉@内容
        user_text = user_text.split(" ", 1)[-1].strip() if " " in user_text else ""
    else:
        content_str = message.get("content", "{}")
        try:
            content = json.loads(content_str)
            user_text = content.get("text", "").strip()
        except:
            user_text = ""
    
    # 获取发送者ID
    sender_id = message.get("sender", {}).get("sender_id", {}).get("open_id")
    if not sender_id:
        sender_id = event.get("sender", {}).get("sender_id", {}).get("open_id")
    if not sender_id and chat_type == "p2p":
        sender_id = chat_id
    
    if not chat_id or not sender_id:
        return jsonify({"code": 0}), 200
    
    conv_key = get_conversation_key(chat_id, sender_id)
    
    # 记录用户消息
    add_message(conv_key, "user", user_text)
    
    # 获取当前状态
    state = get_state(conv_key)
    
    # 调用 AI
    token = get_tenant_token()
    if not token:
        return jsonify({"code": 0}), 200
    
    ai_response = call_kimi(conv_key, user_text, state)
    
    # 记录 AI 回复
    add_message(conv_key, "assistant", ai_response)
    
    # 处理 AI 响应（检查是否需要执行操作）
    handled = handle_ai_response(conv_key, ai_response, chat_id, token)
    
    # 如果没有特殊操作，直接发送回复
    if not handled:
        send_message(token, chat_id, "chat_id", ai_response)
    
    return jsonify({"code": 0}), 200

@app.route("/")
def home():
    return "✅ AI PPT Bot is running!"

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "conversations": len(conversations),
        "version": "AI-1.0"
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
