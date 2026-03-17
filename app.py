from flask import Flask, request, jsonify
import requests
import os
import json
import time
import threading
import hashlib
from ppt_generator import generate_ppt_file

app = Flask(__name__)
print("🚀 PPT Bot Started - Version 9 (Fixed)")

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")
KIMI_API_KEY = os.environ.get("KIMI_API_KEY")

print(f"APP_ID: {FEISHU_APP_ID}")
print(f"KIMI_API_KEY: {'***SET***' if KIMI_API_KEY else '***NOT SET***'}")

# 用户会话状态存储 - 使用 chat_id:user_id 作为 key
user_sessions = {}

# 消息去重 - 记录已处理的消息 ID
processed_messages = set()

# 步骤定义
STEP_TOPIC = "topic"
STEP_OUTLINE = "outline"
STEP_DETAIL = "detail"
STEP_GENERATING = "generating"
STEP_COMPLETE = "complete"


def get_session_key(chat_id, user_id):
    """生成会话 key"""
    return f"{chat_id}:{user_id}"


def get_tenant_token():
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    try:
        res = requests.post(url, json={
            "app_id": FEISHU_APP_ID,
            "app_secret": FEISHU_APP_SECRET
        }, timeout=10)
        data = res.json()
        if data.get("code") != 0:
            print(f"获取token失败: {data}")
            return None
        return data.get("tenant_access_token")
    except Exception as e:
        print(f"获取token异常: {e}")
        return None


def send_message(token, receive_id, receive_id_type, text):
    """发送文本消息"""
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
        res = requests.post(url, headers=headers, json=body, 
                          params={"receive_id_type": receive_id_type}, timeout=10)
        result = res.json()
        if result.get("code") != 0:
            print(f"发送消息失败: {result}")
        return result.get("code") == 0
    except Exception as e:
        print(f"发送消息异常: {e}")
        return False


def upload_file(token, file_path):
    """上传文件到飞书，返回 file_key"""
    url = "https://open.feishu.cn/open-apis/im/v1/files"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    
    try:
        with open(file_path, 'rb') as f:
            files = {
                'file': (os.path.basename(file_path), f, 
                        'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            }
            data = {
                'file_type': 'pptx',
                'file_name': os.path.basename(file_path)
            }
            res = requests.post(url, headers=headers, files=files, data=data, timeout=30)
        
        result = res.json()
        if result.get("code") == 0:
            return result.get("data", {}).get("file_key")
        else:
            print(f"上传文件失败: {result}")
            return None
    except Exception as e:
        print(f"上传文件异常: {e}")
        return None


def send_file(token, receive_id, receive_id_type, file_key):
    """发送文件消息"""
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
        result = res.json()
        if result.get("code") != 0:
            print(f"发送文件消息失败: {result}")
        return result.get("code") == 0
    except Exception as e:
        print(f"发送文件异常: {e}")
        return False


def call_kimi_async(prompt, system_prompt, callback):
    """异步调用 Kimi API"""
    def _call():
        if not KIMI_API_KEY:
            callback("⚠️ Kimi API Key 未配置", None)
            return

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
            res = requests.post(url, headers=headers, json=data, timeout=120)
            result = res.json()
            if "choices" in result and len(result["choices"]) > 0:
                content = result["choices"][0]["message"]["content"]
                callback(content, None)
            else:
                callback(None, f"AI 生成失败: {result}")
        except Exception as e:
            callback(None, f"调用出错: {str(e)}")
    
    thread = threading.Thread(target=_call)
    thread.start()


def generate_outline(topic, callback):
    """生成大纲（异步）"""
    prompt = f"""请为"{topic}"生成一份专业的PPT大纲。

要求：
1. 包含封面、目录、内容页（5-8页）、结束页
2. 每页给出：标题 + 3-5个要点
3. 内容专业、有深度、有数据支撑
4. 用Markdown格式，清晰易读
5. 标题用 ## 开头，要点用 • 开头

请直接输出PPT大纲："""

    call_kimi_async(prompt, "你是一个专业的PPT制作专家，擅长生成结构清晰的PPT大纲。", callback)


def generate_detail_content(outline, topic, callback):
    """根据大纲生成详细内容（异步）"""
    prompt = f"""请根据以下大纲，为"{topic}"生成详细的PPT内容。

大纲：
{outline}

要求：
1. 为每一页生成完整的演讲稿内容
2. 每页内容适合演讲3-5分钟
3. 包含关键数据、案例或论据
4. 语言专业、流畅
5. 用Markdown格式，每页之间用---分隔
6. 每页标题用 ## 开头

请生成详细内容："""

    call_kimi_async(prompt, "你是一个专业的PPT内容撰写专家，擅长将大纲扩展为完整的演讲内容。", callback)


def handle_message(session_key, user_id, text, chat_id, chat_type, token):
    """处理用户消息"""

    # 获取或创建会话
    if session_key not in user_sessions:
        user_sessions[session_key] = {
            "step": STEP_TOPIC,
            "topic": None,
            "outline": None,
            "detail": None,
            "chat_id": chat_id,
            "chat_type": chat_type,
            "user_id": user_id
        }

    session = user_sessions[session_key]
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
        user_sessions[session_key] = {
            "step": STEP_TOPIC,
            "topic": None,
            "outline": None,
            "detail": None,
            "chat_id": chat_id,
            "chat_type": chat_type,
            "user_id": user_id
        }
        return "已重置，请发送新的PPT主题", None

    # 步骤1: 等待主题
    if session["step"] == STEP_TOPIC:
        if len(text) < 2:
            return "主题太短了，请详细描述一下你的PPT主题", None

        session["topic"] = text
        session["step"] = STEP_OUTLINE

        # 异步生成大纲
        def on_outline_done(content, error):
            if error:
                send_message(token, chat_id, "chat_id", 
                           f"❌ 生成大纲失败: {error}")
                session["step"] = STEP_TOPIC
                return
            
            session["outline"] = content
            reply = f"""📋 已生成大纲：《{text}》

{content}

---
✅ 请审核以上大纲

回复：
• "确认" → 生成详细内容
• "修改：xxx" → 告诉我修改意见，我会重新生成
• "重新生成" → 重新生成大纲
• "取消" → 重新开始"""
            send_message(token, chat_id, "chat_id", reply)

        generate_outline(text, on_outline_done)
        return "🤔 正在生成大纲，请稍等...", None

    # 步骤2: 大纲审核
    if session["step"] == STEP_OUTLINE:
        if text_lower in ["确认", "确定", "ok", "yes", "通过"]:
            session["step"] = STEP_DETAIL

            # 异步生成详细内容
            def on_detail_done(content, error):
                if error:
                    send_message(token, chat_id, "chat_id",
                               f"❌ 生成内容失败: {error}")
                    session["step"] = STEP_OUTLINE
                    return

                session["detail"] = content
                reply = f"""📝 已生成详细内容：《{session['topic']}》

{content[:2000]}...

（内容较长，以上为预览）

---
✅ 请审核以上内容

回复：
• "确认" → 开始生成PPT文件
• "重新生成" → 重新生成详细内容
• "取消" → 重新开始"""
                send_message(token, chat_id, "chat_id", reply)

            generate_detail_content(session["outline"], session["topic"], on_detail_done)
            return "🤔 正在生成详细内容，请稍等...", None

        elif text_lower in ["重新生成", "再来一次", "regenerate"]:
            session["step"] = STEP_TOPIC
            
            def on_regenerate_done(content, error):
                if error:
                    send_message(token, chat_id, "chat_id",
                               f"❌ 重新生成失败: {error}")
                    return
                
                session["outline"] = content
                session["step"] = STEP_OUTLINE
                reply = f"""📋 重新生成的大纲：《{session['topic']}》

{content}

---
✅ 请审核以上大纲

回复"确认"继续，或"修改"告诉我意见"""
                send_message(token, chat_id, "chat_id", reply)

            generate_outline(session["topic"], on_regenerate_done)
            return "🤔 正在重新生成大纲，请稍等...", None

        elif text_lower.startswith("修改") or text_lower.startswith("优化"):
            feedback = text[2:].strip() if len(text) > 2 else ""
            prompt = f"""请根据以下修改意见，重新生成PPT大纲。

原主题：{session['topic']}
修改意见：{feedback}

请重新生成大纲："""

            def on_modify_done(content, error):
                if error:
                    send_message(token, chat_id, "chat_id",
                               f"❌ 修改失败: {error}")
                    return
                
                session["outline"] = content
                reply = f"""📋 根据你的意见修改后的大纲：

{content}

---
回复"确认"继续，或继续提修改意见"""
                send_message(token, chat_id, "chat_id", reply)

            call_kimi_async(prompt, "你是一个专业的PPT制作专家。", on_modify_done)
            return "🤔 正在根据意见修改，请稍等...", None

        else:
            return '请回复"确认"通过大纲，或"修改：xxx"告诉我修改意见', None

    # 步骤3: 详细内容审核
    if session["step"] == STEP_DETAIL:
        if text_lower in ["确认", "确定", "ok", "yes", "通过"]:
            session["step"] = STEP_GENERATING

            def generate_and_send():
                try:
                    topic = session["topic"]
                    outline = session["outline"]
                    detail = session.get("detail", "")

                    send_message(token, chat_id, "chat_id",
                               f"🎨 正在生成PPT文件：《{topic}》\n\n请稍等，预计需要15-30秒...")

                    print(f"开始生成PPT: {topic}")
                    ppt_path = generate_ppt_file(topic, outline, detail, output_dir="/tmp")
                    print(f"PPT已生成: {ppt_path}")

                    file_key = upload_file(token, ppt_path)
                    if not file_key:
                        send_message(token, chat_id, "chat_id",
                                   "❌ 文件上传失败，请稍后重试")
                        session["step"] = STEP_DETAIL
                        return

                    print(f"文件已上传: {file_key}")
                    success = send_file(token, chat_id, "chat_id", file_key)

                    if success:
                        try:
                            os.remove(ppt_path)
                        except:
                            pass

                        user_sessions[session_key] = {
                            "step": STEP_TOPIC,
                            "topic": None,
                            "outline": None,
                            "detail": None,
                            "chat_id": chat_id,
                            "chat_type": chat_type,
                            "user_id": user_id
                        }

                        send_message(token, chat_id, "chat_id",
                                   f"✅ PPT《{topic}》已生成并发送！\n\n如需制作新的PPT，请直接发送主题。")
                    else:
                        send_message(token, chat_id, "chat_id",
                                   "❌ 文件发送失败，请稍后重试")
                        session["step"] = STEP_DETAIL

                except Exception as e:
                    print(f"生成PPT出错: {e}")
                    send_message(token, chat_id, "chat_id",
                               f"❌ 生成PPT时出错: {str(e)}\n请稍后重试或联系管理员")
                    session["step"] = STEP_DETAIL

            thread = threading.Thread(target=generate_and_send)
            thread.start()
            return None, None

        elif text_lower in ["重新生成", "再来一次"]:
            session["step"] = STEP_OUTLINE
            
            def on_regenerate_detail(content, error):
                if error:
                    send_message(token, chat_id, "chat_id",
                               f"❌ 重新生成失败: {error}")
                    return

                session["detail"] = content
                session["step"] = STEP_DETAIL
                reply = f"""📝 重新生成的详细内容：

{content[:1500]}...

回复"确认"继续生成PPT，或"取消"重新开始"""
                send_message(token, chat_id, "chat_id", reply)

            generate_detail_content(session["outline"], session["topic"], on_regenerate_detail)
            return "🤔 正在重新生成详细内容，请稍等...", None

        else:
            return '请回复"确认"生成PPT，或"重新生成"重新创建内容', None

    # 步骤4: 生成中
    if session["step"] == STEP_GENERATING:
        return "正在生成PPT中，请稍等...", None

    # 默认
    return "请发送PPT主题开始制作", None


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    print(f"收到请求: {json.dumps(data, ensure_ascii=False)[:500]}")

    # 处理飞书验证
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    header = data.get("header", {})
    event_type = header.get("event_type", "")
    event = data.get("event", {})
    
    # 消息去重 - 使用 message_id
    message_id = header.get("event_id", "")
    if message_id in processed_messages:
        print(f"重复消息，跳过: {message_id}")
        return jsonify({"code": 0}), 200
    processed_messages.add(message_id)
    
    # 限制缓存大小
    if len(processed_messages) > 1000:
        processed_messages.clear()

    # 处理消息
    if event_type == "im.message.receive_v1":
        message = event.get("message", {})
        chat_type = message.get("chat_type")
        chat_id = message.get("chat_id")
        
        # 尝试多种方式获取发送者ID
        sender = message.get("sender", {})
        sender_id = sender.get("sender_id", {}).get("open_id")
        
        # 如果 message 中没有，尝试从 event 根级别获取
        if not sender_id:
            sender_id = event.get("sender", {}).get("sender_id", {}).get("open_id")
        
        # 再尝试从 operator_id 获取（某些事件类型）
        if not sender_id:
            sender_id = event.get("operator_id", {}).get("open_id")
        
        # 备用方案：私聊模式下，如果还是获取不到，使用 chat_id 作为 sender_id
        if not sender_id and chat_type == "p2p" and chat_id:
            sender_id = chat_id
            print(f"使用 chat_id 作为 sender_id: {sender_id}")
        
        # 获取消息内容
        content_str = message.get("content", "{}")
        try:
            content = json.loads(content_str)
            user_text = content.get("text", "").strip()
        except:
            user_text = ""

        print(f"[{chat_id}:{sender_id}]: {user_text}")
        
        # 检查必要参数
        if not chat_id or not sender_id:
            print(f"缺少参数: chat_id={chat_id}, sender_id={sender_id}")
            return jsonify({"code": 0}), 200

        # 生成会话 key
        session_key = get_session_key(chat_id, sender_id)
        
        # 获取 token
        token = get_tenant_token()
        if not token:
            return jsonify({"code": 0}), 200

        # 处理消息
        reply_text, action = handle_message(session_key, sender_id, user_text, chat_id, chat_type, token)
        
        # 发送回复
        if reply_text:
            if chat_type == "p2p":
                # 私聊模式下使用 chat_id 作为接收者（chat_id 在私聊中就是会话ID）
                send_message(token, chat_id, "chat_id", reply_text)
            else:
                send_message(token, chat_id, "chat_id", reply_text)

        return jsonify({"code": 0}), 200

    return jsonify({"code": 0}), 200


@app.route("/")
def home():
    return "✅ PPT Bot (Fixed Version 9) is running!"


@app.route("/health")
def health():
    return jsonify({
        "status": "ok", 
        "version": "9.0",
        "sessions": len(user_sessions),
        "processed_messages": len(processed_messages)
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
