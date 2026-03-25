    def add_kpi_slide(self, title, kpis, layout='grid'):
        """添加KPI展示页
        
        Args:
            title: 页面标题
            kpis: KPI列表，每个KPI为字典 {'name': '名称', 'value': '值', 'unit': '单位', 'change': '变化'}
            layout: 布局方式 'grid'(2x2网格) 或 'large'(单个大KPI)
        """
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
        
        if layout == 'large' and kpis:
            # 单个大KPI布局
            kpi = kpis[0]
            
            # 大数字
            value_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
            tf = value_box.text_frame
            p = tf.paragraphs[0]
            p.text = f"{kpi['value']}{kpi.get('unit', '')}"
            p.font.size = Pt(72)
            p.font.bold = True
            p.font.color.rgb = COLORS['primary']
            p.font.name = "Microsoft YaHei"
            p.alignment = PP_ALIGN.CENTER
            
            # KPI名称
            name_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11), Inches(0.8))
            tf = name_box.text_frame
            p = tf.paragraphs[0]
            p.text = kpi['name']
            p.font.size = Pt(24)
            p.font.color.rgb = COLORS['text_dark']
            p.font.name = "Microsoft YaHei"
            p.alignment = PP_ALIGN.CENTER
            
            # 变化值
            if 'change' in kpi:
                change_box = slide.shapes.add_textbox(Inches(1), Inches(5.3), Inches(11), Inches(0.6))
                tf = change_box.text_frame
                p = tf.paragraphs[0]
                p.text = kpi['change']
                p.font.size = Pt(18)
                p.font.color.rgb = COLORS['accent'] if '+' in kpi['change'] else COLORS['secondary']
                p.font.name = "Microsoft YaHei"
                p.alignment = PP_ALIGN.CENTER
        else:
            # 2x2网格布局
            positions = [
                (Inches(0.8), Inches(1.5)),
                (Inches(6.8), Inches(1.5)),
                (Inches(0.8), Inches(4.2)),
                (Inches(6.8), Inches(4.2)),
            ]
            
            for i, kpi in enumerate(kpis[:4]):
                if i >= len(positions):
                    break
                    
                x, y = positions[i]
                
                # KPI卡片背景
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, x, y,
                    Inches(5.5), Inches(2.5)
                )
                card.fill.solid()
                card.fill.fore_color.rgb = COLORS['bg_light']
                card.line.fill.background()
                
                # KPI值
                value_box = slide.shapes.add_textbox(x + Inches(0.3), y + Inches(0.4), Inches(5), Inches(1))
                tf = value_box.text_frame
                p = tf.paragraphs[0]
                p.text = f"{kpi['value']}{kpi.get('unit', '')}"
                p.font.size = Pt(36)
                p.font.bold = True
                p.font.color.rgb = COLORS['primary']
                p.font.name = "Microsoft YaHei"
                
                # KPI名称
                name_box = slide.shapes.add_textbox(x + Inches(0.3), y + Inches(1.4), Inches(5), Inches(0.5))
                tf = name_box.text_frame
                p = tf.paragraphs[0]
                p.text = kpi['name']
                p.font.size = Pt(14)
                p.font.color.rgb = COLORS['text_dark']
                p.font.name = "Microsoft YaHei"
                
                # 变化值
                if 'change' in kpi:
                    change_box = slide.shapes.add_textbox(x + Inches(0.3), y + Inches(1.9), Inches(5), Inches(0.4))
                    tf = change_box.text_frame
                    p = tf.paragraphs[0]
                    p.text = kpi['change']
                    p.font.size = Pt(12)
                    p.font.color.rgb = COLORS['accent'] if '+' in str(kpi['change']) else COLORS['secondary']
                    p.font.name = "Microsoft YaHei"

