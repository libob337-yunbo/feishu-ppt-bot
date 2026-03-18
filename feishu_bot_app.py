# 飞书机器人 Webhook 处理示例（安全版本）
import os
from flask import Flask, request, jsonify
import requests
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

app = Flask(__name__)

# 从环境变量读取敏感信息（不再硬编码）
APP_ID = os.environ.get('FEISHU_APP_ID')
APP_SECRET = os.environ.get('FEISHU_APP_SECRET')
VERIFICATION_TOKEN = os.environ.get('FEISHU_VERIFICATION_TOKEN')

# 获取 tenant_access_token 的缓存
token_cache = {
    'token': None,
    'expire_time': 0
}

def get_tenant_access_token():
    """获取飞书 tenant_access_token"""
    import time
    
    # 检查缓存
    if token_cache['token'] and time.time() < token_cache['expire_time']:
        return token_cache['token']
    
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    data = {
        "app_id": APP_ID,
        "app_secret": APP_SECRET
    }
    
    resp = requests.post(url, headers=headers, json=data)
    result = resp.json()
    
    if result.get("code") == 0:
        token_cache['token'] = result["tenant_access_token"]
        token_cache['expire_time'] = time.time() + result["expire"] - 60  # 提前60秒过期
        return token_cache['token']
    else:
        print(f"获取 token 失败: {result}")
        return None

def reply_message(message_id, content):
    """回复消息"""
    token = get_tenant_access_token()
    if not token:
        return False
    
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/reply"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    data = {
        "content": json.dumps({"text": content}),
        "msg_type": "text"
    }
    
    resp = requests.post(url, headers=headers, json=data)
    result = resp.json()
    
    if result.get("code") != 0:
        print(f"回复消息失败: {result}")
        return False
    return True

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    print(f"收到请求: {data}")
    
    # 1. 处理飞书验证请求（首次配置 webhook 时会发送）
    if data.get('type') == 'url_verification':
        return jsonify({"challenge": data.get('challenge')})
    
    # 2. 验证 token（可选，用于确认请求来自飞书）
    token = data.get('token')
    if VERIFICATION_TOKEN and token != VERIFICATION_TOKEN:
        print(f"Token 验证失败: {token}")
        return jsonify({"code": 403, "msg": "Invalid token"}), 403
    
    # 3. 处理消息事件
    event = data.get('event', {})
    if event.get('type') == 'im.message.receive_v1':
        message = event.get('message', {})
        message_id = message.get('message_id')
        content = json.loads(message.get('content', '{}'))
        text = content.get('text', '')
        
        print(f"收到消息: {text}")
        
        # 回复消息
        reply_text = f"你说了: {text}"
        reply_message(message_id, reply_text)
    
    return jsonify({"code": 0, "msg": "success"})

if __name__ == '__main__':
    # 检查环境变量
    if not APP_ID or not APP_SECRET:
        print("错误: 请设置环境变量 FEISHU_APP_ID 和 FEISHU_APP_SECRET")
        print("示例: export FEISHU_APP_ID=cli_xxx")
        exit(1)
    
    app.run(host='0.0.0.0', port=10000)
