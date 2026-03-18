#!/bin/bash
# 修复 PPT 助手配置路径问题

echo "🔧 修复 PPT 助手配置路径..."
echo "=============================="

PPT_DIR="$HOME/.openclaw-ppt"

# 创建正确的 .openclaw 目录结构
mkdir -p "$PPT_DIR/.openclaw"

# 移动配置文件到正确位置
if [ -f "$PPT_DIR/openclaw.json" ]; then
    mv "$PPT_DIR/openclaw.json" "$PPT_DIR/.openclaw/openclaw.json"
    echo "✅ 配置文件已移动到正确位置"
fi

# 同样处理简化配置
if [ -f "$PPT_DIR/openclaw-nolaunch.json" ]; then
    mv "$PPT_DIR/openclaw-nolaunch.json" "$PPT_DIR/.openclaw/openclaw.json"
    echo "✅ 简化配置已应用"
fi

echo ""
echo "🚀 现在启动 PPT 助手..."
echo ""

# 前台启动测试
cd "$PPT_DIR"
export HOME="$PPT_DIR"

echo "📂 HOME 目录: $HOME"
echo "📄 配置文件: $HOME/.openclaw/openclaw.json"
echo ""

if [ -f "$HOME/.openclaw/openclaw.json" ]; then
    echo "✅ 配置文件存在"
    echo "🚀 启动 Gateway..."
    echo ""
    # 前台运行，可以看到所有输出
    openclaw gateway run 2>&1
else
    echo "❌ 配置文件不存在，检查路径:"
    ls -la "$PPT_DIR/"
    ls -la "$PPT_DIR/.openclaw/" 2>/dev/null || echo "   .openclaw 目录不存在"
fi
