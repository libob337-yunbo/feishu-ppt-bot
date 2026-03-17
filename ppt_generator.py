# PPT Generator Module
# 使用 python-pptx 生成商务简约风格 PPT

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os
import json
import re
from datetime import datetime

# 商务简约配色方案
COLORS = {
    'primary': RGBColor(0x1E, 0x3A, 0x5F),      # 深蓝 #1E3A5F
    'secondary': RGBColor(0x2E, 0x5C, 0x8A),    # 中蓝 #2E5C8A
    'accent': RGBColor(0x4A, 0x90, 0xA4),       # 青蓝 #4A90A4
    'text_dark': RGBColor(0x33, 0x33, 0x33),    # 深灰 #333333
    'text_light': RGBColor(0x66, 0x66, 0x66),   # 浅灰 #666666
    'white': RGBColor(0xFF, 0xFF, 0xFF),        # 白色
    'bg_light': RGBColor(0xF5, 0xF7, 0xFA),     # 浅灰背景
}

class PPTGenerator:
    def __init__(self):
        self.prs = Presentation()
        # 设置幻灯片尺寸为 16:9
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        
    def add_title_slide(self, title, subtitle=""):
        """添加封面页"""
        slide_layout = self.prs.slide_layouts[6]  # 空白布局
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 添加背景色块
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
            Inches(13.333), Inches(7.5)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLORS['primary']
        shape.line.fill.background()
        
        # 添加装饰线条
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1), Inches(4.5), 
            Inches(2), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = COLORS['accent']
        line.line.fill.background()
        
        # 添加标题
        title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(1.5))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.font.name = "Microsoft YaHei"
        
        # 添加副标题
        if subtitle:
            sub_box = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(11), Inches(1))
            tf = sub_box.text_frame
            p = tf.paragraphs[0]
            p.text = subtitle
            p.font.size = Pt(20)
            p.font.color.rgb = COLORS['accent']
            p.font.name = "Microsoft YaHei"
            
        # 添加日期
        date_box = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(4), Inches(0.5))
        tf = date_box.text_frame
        p = tf.paragraphs[0]
        p.text = datetime.now().strftime("%Y年%m月%d日")
        p.font.size = Pt(12)
        p.font.color.rgb = COLORS['white']
        p.font.name = "Microsoft YaHei"
        
    def add_content_slide(self, title, content_list):
        """添加内容页"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 添加顶部色条
        header = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(0.15)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = COLORS['primary']
        header.line.fill.background()
        
        # 添加标题
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(12), Inches(0.8))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = COLORS['primary']
        p.font.name = "Microsoft YaHei"
        
        # 添加内容
        content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(12), Inches(5.5))
        tf = content_box.text_frame
        tf.word_wrap = True
        
        for i, item in enumerate(content_list):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            # 处理列表项
            if item.startswith('•') or item.startswith('-') or item.startswith('*'):
                p.text = "    " + item[1:].strip()
                p.level = 0
            elif re.match(r'^\d+[\.\)]\s', item):
                p.text = item
                p.level = 0
            else:
                p.text = "    • " + item
                p.level = 0
                
            p.font.size = Pt(16)
            p.font.color.rgb = COLORS['text_dark']
            p.font.name = "Microsoft YaHei"
            p.space_before = Pt(12)
            p.space_after = Pt(6)
            
    def add_section_slide(self, section_title):
        """添加章节分隔页"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 背景
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLORS['secondary']
        bg.line.fill.background()
        
        # 章节标题
        title_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1.5))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = section_title
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.font.name = "Microsoft YaHei"
        p.alignment = PP_ALIGN.CENTER
        
    def add_chart_slide(self, title, chart_data):
        """添加图表页（简化版，使用表格展示数据）"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 顶部色条
        header = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(0.15)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = COLORS['primary']
        header.line.fill.background()
        
        # 标题
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(12), Inches(0.8))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = COLORS['primary']
        p.font.name = "Microsoft YaHei"
        
        # 添加数据展示（使用文本框模拟图表说明）
        if chart_data:
            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(12), Inches(5.5))
            tf = content_box.text_frame
            tf.word_wrap = True
            
            for i, data in enumerate(chart_data):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = f"{data}"
                p.font.size = Pt(16)
                p.font.color.rgb = COLORS['text_dark']
                p.font.name = "Microsoft YaHei"
                p.space_before = Pt(8)
                
    def add_end_slide(self, message="谢谢观看"):
        """添加结束页"""
        slide_layout = self.prs.slide_layouts[6]
        slide = self.prs.slides.add_slide(slide_layout)
        
        # 背景
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLORS['primary']
        bg.line.fill.background()
        
        # 结束语
        msg_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1.5))
        tf = msg_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = message
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = COLORS['white']
        p.font.name = "Microsoft YaHei"
        p.alignment = PP_ALIGN.CENTER
        
    def parse_outline(self, outline_text):
        """解析大纲文本，提取页面结构"""
        pages = []
        current_page = None
        
        lines = outline_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 检测标题（# 或数字开头）
            if line.startswith('#') or re.match(r'^\d+[\.\)]', line):
                if current_page:
                    pages.append(current_page)
                title = re.sub(r'^[#\d\.\)\s]+', '', line)
                current_page = {'title': title, 'content': []}
            elif current_page and (line.startswith('•') or line.startswith('-') or line.startswith('*')):
                current_page['content'].append(line)
            elif current_page and len(line) > 0:
                current_page['content'].append(line)
                
        if current_page:
            pages.append(current_page)
            
        return pages
        
    def generate_from_outline(self, topic, outline_text, detail_text=""):
        """根据大纲生成完整PPT"""
        # 封面
        self.add_title_slide(topic, "AI 智能生成")
        
        # 解析大纲
        pages = self.parse_outline(outline_text)
        
        # 目录页
        if len(pages) > 0:
            self.add_section_slide("目录")
            toc_slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            
            # 顶部色条
            header = toc_slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                Inches(13.333), Inches(0.15)
            )
            header.fill.solid()
            header.fill.fore_color.rgb = COLORS['primary']
            header.line.fill.background()
            
            title_box = toc_slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(12), Inches(0.8))
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.text = "目录"
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = COLORS['primary']
            p.font.name = "Microsoft YaHei"
            
            content_box = toc_slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(12), Inches(5.5))
            tf = content_box.text_frame
            tf.word_wrap = True
            
            for i, page in enumerate(pages[:6]):  # 最多显示6个章节
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = f"{i+1}. {page['title']}"
                p.font.size = Pt(18)
                p.font.color.rgb = COLORS['text_dark']
                p.font.name = "Microsoft YaHei"
                p.space_before = Pt(16)
        
        # 内容页
        for i, page in enumerate(pages):
            if i % 3 == 0 and i > 0:  # 每3页添加一个章节分隔
                self.add_section_slide(page['title'])
            self.add_content_slide(page['title'], page['content'])
            
        # 结束页
        self.add_end_slide()
        
    def save(self, filename):
        """保存PPT文件"""
        self.prs.save(filename)
        return filename


def generate_ppt_file(topic, outline, detail="", output_dir="/tmp"):
    """生成PPT文件的便捷函数"""
    generator = PPTGenerator()
    generator.generate_from_outline(topic, outline, detail)
    
    filename = f"{output_dir}/PPT_{topic[:20]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    generator.save(filename)
    return filename


if __name__ == "__main__":
    # 测试
    test_outline = """
# 成都写字楼市场分析

## 市场概况
• 成都写字楼总存量超过500万平方米
• 主要分布在高新区、天府新区、锦江区
• 2024年整体空置率约18%

## 租金走势
• 平均租金约80-120元/㎡/月
• 核心商务区租金保持稳定
• 新兴区域租金呈下降趋势

## 投资建议
• 关注TOD项目周边写字楼
• 优选地铁沿线物业
• 考虑产业聚集效应
"""
    
    filename = generate_ppt_file("成都写字楼市场分析", test_outline)
    print(f"PPT已生成: {filename}")
