import os
import sys
import uno
from com.sun.star.beans import PropertyValue

# ====================== 颜色配置 ======================
def rgb_to_bgr(rgb):
    """将RGB颜色转换为BGR格式（LibreOffice使用）"""
    r = (rgb >> 16) & 0xFF
    g = (rgb >> 8) & 0xFF
    b = rgb & 0xFF
    return (b << 16) | (g << 8) | r

RGB_COLORS = [
    0xFF3333,  # 浅红
    0x33CC33,  # 浅绿
    0x3333FF,  # 浅蓝
    0xFF9933,  # 浅橙
    0x9933FF,  # 浅紫
    0x33CCCC,  # 浅青
    0xFF3399,  # 浅玫红
    0x999933   # 浅橄榄绿
]

COLORS = [rgb_to_bgr(c) for c in RGB_COLORS]
COLOR_NAMES = ["浅红", "浅绿", "浅蓝", "浅橙", "浅紫", "浅青", "浅玫红", "浅橄榄绿"]

class LineColorizer:
    def __init__(self, doc, target_lines):
        """
        初始化行着色器
        doc: 文档对象
        target_lines: 目标行列表
        """
        self.doc = doc
        self.text = doc.getText()
        self.target_lines = target_lines
        self.total_lines = len(target_lines)
        self.current_target_index = 0
        self.colored_count = 0
        self.search_desc = None
        
        print(f"\n🎯 目标行列表 (共{self.total_lines}行):")
        for i, line in enumerate(target_lines[:10]):
            print(f"  行{i+1}: {line[:50]}")
    
    def init_search(self):
        """初始化搜索描述符"""
        if not self.search_desc:
            self.search_desc = self.doc.createSearchDescriptor()
            self.search_desc.SearchCaseSensitive = True
            self.search_desc.SearchWords = False
            self.search_desc.SearchRegularExpression = False
    
    def colorize_text_range_simple(self, text_range):
        """
        为找到的文本范围着色 - 最简单的实现
        直接使用找到的文本范围设置属性
        """
        try:
            # 直接设置找到的文本范围的颜色属性
            color_idx = self.current_target_index % len(COLORS)
            text_range.CharColor = COLORS[color_idx]
            
            print(f"  ✅ 行{self.current_target_index+1} 已着色为 {COLOR_NAMES[color_idx]}")
            
            self.colored_count += 1
            self.current_target_index += 1
            return True
            
        except Exception as e:
            print(f"  ❌ 着色失败: {str(e)}")
            return False
    
    def safe_get_string(self, element):
        """安全地获取元素的字符串内容"""
        try:
            if hasattr(element, 'getString'):
                return element.getString()
            return ""
        except:
            return ""
    
    def process_paragraph(self, paragraph, source="段落"):
        """处理段落"""
        if self.current_target_index >= self.total_lines:
            return False
        
        try:
            # 获取段落文本
            para_text = self.safe_get_string(paragraph)
            if not para_text:
                return False
            
            print(f"\n📄 [{source}] 处理: {para_text[:50]}...")
            
            # 在当前段落中查找目标行
            processed = False
            while self.current_target_index < self.total_lines:
                target = self.target_lines[self.current_target_index]
                
                if not target.strip():
                    print(f"  ⏭️ 行{self.current_target_index+1} 是空行，跳过")
                    self.current_target_index += 1
                    processed = True
                    continue
                
                # 在段落中查找目标行
                pos = para_text.find(target)
                if pos >= 0:
                    # 验证是完整行
                    before_ok = (pos == 0 or para_text[pos-1] in '\n\r')
                    after_ok = (pos + len(target) >= len(para_text) or 
                               para_text[pos + len(target)] in '\n\r')
                    
                    if before_ok and after_ok:
                        print(f"  🔍 找到目标行 {self.current_target_index+1} 在位置 {pos}")
                        
                        # 创建文本范围
                        start = paragraph.getStart()
                        cursor = self.text.createTextCursorByRange(start)
                        
                        # 移动到目标位置
                        for _ in range(pos):
                            cursor.goRight(1, False)
                        
                        # 创建结束位置
                        end_cursor = self.text.createTextCursorByRange(cursor.getStart())
                        for _ in range(len(target)):
                            end_cursor.goRight(1, False)
                        
                        # 创建范围并着色
                        range_cursor = self.text.createTextCursorByRange(cursor.getStart())
                        range_cursor.gotoRange(end_cursor.getStart(), True)
                        
                        color_idx = self.current_target_index % len(COLORS)
                        range_cursor.CharColor = COLORS[color_idx]
                        
                        print(f"  ✅ 行{self.current_target_index+1} 已着色为 {COLOR_NAMES[color_idx]}")
                        
                        self.colored_count += 1
                        self.current_target_index += 1
                        processed = True
                    else:
                        print(f"  ⚠️ 找到部分匹配，跳过")
                        self.current_target_index += 1
                        processed = True
                else:
                    # 当前段落中没有找到，退出循环
                    break
            
            return processed
            
        except Exception as e:
            print(f"❌ 处理{source}出错: {str(e)}")
            return False
    
    def process_table_cell(self, cell):
        """处理表格单元格"""
        if self.current_target_index >= self.total_lines:
            return False
        
        try:
            # 获取单元格文本
            cell_text = self.safe_get_string(cell)
            if not cell_text:
                return False
            
            print(f"\n  📋 处理单元格: {cell_text[:50]}...")
            
            # 在单元格中查找目标行
            processed = False
            while self.current_target_index < self.total_lines:
                target = self.target_lines[self.current_target_index]
                
                if not target.strip():
                    print(f"    ⏭️ 行{self.current_target_index+1} 是空行，跳过")
                    self.current_target_index += 1
                    processed = True
                    continue
                
                pos = cell_text.find(target)
                if pos >= 0:
                    print(f"    🔍 找到目标行 {self.current_target_index+1} 在位置 {pos}")
                    
                    # 创建文本范围
                    start = cell.getStart()
                    cursor = self.text.createTextCursorByRange(start)
                    
                    # 移动到目标位置
                    for _ in range(pos):
                        cursor.goRight(1, False)
                    
                    # 创建结束位置
                    end_cursor = self.text.createTextCursorByRange(cursor.getStart())
                    for _ in range(len(target)):
                        end_cursor.goRight(1, False)
                    
                    # 创建范围并着色
                    range_cursor = self.text.createTextCursorByRange(cursor.getStart())
                    range_cursor.gotoRange(end_cursor.getStart(), True)
                    
                    color_idx = self.current_target_index % len(COLORS)
                    range_cursor.CharColor = COLORS[color_idx]
                    
                    print(f"    ✅ 行{self.current_target_index+1} 已着色为 {COLOR_NAMES[color_idx]}")
                    
                    self.colored_count += 1
                    self.current_target_index += 1
                    processed = True
                else:
                    break
            
            return processed
                    
        except Exception as e:
            print(f"  ❌ 处理单元格出错: {str(e)}")
            return False
    
    def process_table(self, table):
        """处理表格"""
        if self.current_target_index >= self.total_lines:
            return
        
        print(f"\n📊 处理表格:")
        try:
            rows = table.getRows()
            row_count = rows.getCount()
            columns = table.getColumns()
            col_count = columns.getCount()
            
            print(f"  表格大小: {row_count}行 x {col_count}列")
            
            for row_idx in range(row_count):
                for col_idx in range(col_count):
                    try:
                        cell = table.getCellByPosition(col_idx, row_idx)
                        if cell:
                            self.process_table_cell(cell)
                            if self.current_target_index >= self.total_lines:
                                return
                    except Exception as e:
                        print(f"  无法访问单元格 [{row_idx},{col_idx}]: {str(e)}")
                        
        except Exception as e:
            print(f"  ❌ 处理表格出错: {str(e)}")
    
    def process_header_footer(self, page_style):
        """处理页眉页脚"""
        if self.current_target_index >= self.total_lines:
            return
        
        try:
            if hasattr(page_style, 'getHeaderText'):
                header_text = page_style.getHeaderText()
                if header_text:
                    print(f"\n📌 处理页眉:")
                    paragraphs = header_text.createEnumeration()
                    while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
                        para = paragraphs.nextElement()
                        self.process_paragraph(para, "页眉")
            
            if hasattr(page_style, 'getFooterText'):
                footer_text = page_style.getFooterText()
                if footer_text:
                    print(f"\n📌 处理页脚:")
                    paragraphs = footer_text.createEnumeration()
                    while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
                        para = paragraphs.nextElement()
                        self.process_paragraph(para, "页脚")
                        
        except Exception as e:
            print(f"  ⚠️ 处理页眉页脚出错: {str(e)}")
    
    def process_footnote(self, footnote):
        """处理脚注"""
        if self.current_target_index >= self.total_lines:
            return
        
        try:
            if hasattr(footnote, 'getText'):
                footnote_text = footnote.getText()
                if footnote_text:
                    print(f"\n📝 处理脚注:")
                    paragraphs = footnote_text.createEnumeration()
                    while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
                        para = paragraphs.nextElement()
                        self.process_paragraph(para, "脚注")
        except Exception as e:
            print(f"  ⚠️ 处理脚注出错: {str(e)}")
    
    def process_endnote(self, endnote):
        """处理尾注"""
        if self.current_target_index >= self.total_lines:
            return
        
        try:
            if hasattr(endnote, 'getText'):
                endnote_text = endnote.getText()
                if endnote_text:
                    print(f"\n📝 处理尾注:")
                    paragraphs = endnote_text.createEnumeration()
                    while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
                        para = paragraphs.nextElement()
                        self.process_paragraph(para, "尾注")
        except Exception as e:
            print(f"  ⚠️ 处理尾注出错: {str(e)}")
    
    def process_text_frame(self, frame):
        """处理文本框"""
        if self.current_target_index >= self.total_lines:
            return
        
        try:
            if hasattr(frame, 'getText'):
                frame_text = frame.getText()
                if frame_text:
                    print(f"\n🖼️ 处理文本框:")
                    paragraphs = frame_text.createEnumeration()
                    while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
                        para = paragraphs.nextElement()
                        self.process_paragraph(para, "文本框")
        except Exception as e:
            print(f"  ❌ 处理文本框出错: {str(e)}")
    
    def search_with_find_api_simple(self):
        """
        使用LibreOffice的查找API搜索剩余行 - 最简单的实现
        直接对找到的文本范围着色
        """
        if self.current_target_index >= self.total_lines:
            return
        
        print(f"\n🔍 使用查找API搜索剩余行 (从行{self.current_target_index+1}开始)...")
        
        # 初始化查找
        self.init_search()
        
        # 从文档开头开始查找
        start_point = self.text.getStart()
        
        while self.current_target_index < self.total_lines:
            target = self.target_lines[self.current_target_index]
            
            if not target.strip():
                print(f"  ⏭️ 行{self.current_target_index+1} 是空行，跳过")
                self.current_target_index += 1
                continue
            
            print(f"\n  🔍 搜索: {target[:50]}")
            
            # 设置搜索字符串
            self.search_desc.setSearchString(target)
            
            # 从当前点开始搜索
            found = self.doc.findNext(start_point, self.search_desc)
            
            if found:
                print(f"  ✅ 找到行{self.current_target_index+1}")
                print(f"     找到内容: {found.getString()[:50]}")
                
                # 直接对找到的文本范围着色 - 最简单的方法
                if self.colorize_text_range_simple(found):
                    # 更新搜索起点到找到位置之后
                    start_point = found.getEnd()
                else:
                    # 着色失败，移动到下一个位置继续搜索
                    start_point = found.getEnd()
                    self.current_target_index += 1
            else:
                print(f"  ❌ 无法找到行{self.current_target_index+1}")
                self.current_target_index += 1
                
                # 重置搜索起点到文档开头，尝试下一个目标
                start_point = self.text.getStart()
    
    def process_document(self):
        """处理整个文档的所有内容"""
        print("\n🔍 开始遍历文档所有内容...")
        
        # 1. 处理所有普通段落
        print("\n📝 处理普通段落:")
        paragraphs = self.text.createEnumeration()
        para_count = 0
        while paragraphs.hasMoreElements() and self.current_target_index < self.total_lines:
            para = paragraphs.nextElement()
            para_count += 1
            self.process_paragraph(para, "普通段落")
        print(f"  共处理 {para_count} 个段落")
        
        if self.current_target_index >= self.total_lines:
            return
        
        # 2. 处理所有表格
        try:
            if hasattr(self.doc, 'getTextTables'):
                tables = self.doc.getTextTables()
                if tables and hasattr(tables, 'getCount') and tables.getCount() > 0:
                    print(f"\n📊 发现 {tables.getCount()} 个表格")
                    for i in range(tables.getCount()):
                        try:
                            table = tables.getByIndex(i)
                            self.process_table(table)
                            if self.current_target_index >= self.total_lines:
                                return
                        except Exception as e:
                            print(f"  无法访问表格 {i}: {str(e)}")
        except Exception as e:
            print(f"  ⚠️ 处理表格集合出错: {str(e)}")
        
        if self.current_target_index >= self.total_lines:
            return
        
        # 3. 处理所有文本框
        try:
            if hasattr(self.doc, 'getTextFrames'):
                frames = self.doc.getTextFrames()
                if frames and hasattr(frames, 'getCount') and frames.getCount() > 0:
                    print(f"\n🖼️ 发现 {frames.getCount()} 个文本框")
                    for i in range(frames.getCount()):
                        try:
                            frame = frames.getByIndex(i)
                            self.process_text_frame(frame)
                            if self.current_target_index >= self.total_lines:
                                return
                        except Exception as e:
                            print(f"  无法访问文本框 {i}: {str(e)}")
        except Exception as e:
            print(f"  ⚠️ 处理文本框集合出错: {str(e)}")
        
        if self.current_target_index >= self.total_lines:
            return
        
        # 4. 处理页眉页脚
        try:
            if hasattr(self.doc, 'getStyleFamilies'):
                style_families = self.doc.getStyleFamilies()
                if style_families and hasattr(style_families, 'hasByName') and style_families.hasByName('PageStyles'):
                    page_styles = style_families.getByName('PageStyles')
                    if page_styles:
                        style_names = page_styles.getElementNames()
                        for style_name in style_names:
                            page_style = page_styles.getByName(style_name)
                            self.process_header_footer(page_style)
                            if self.current_target_index >= self.total_lines:
                                return
        except Exception as e:
            print(f"  ⚠️ 处理页眉页脚出错: {str(e)}")
        
        if self.current_target_index >= self.total_lines:
            return
        
        # 5. 处理脚注
        try:
            if hasattr(self.doc, 'getFootnotes'):
                footnotes = self.doc.getFootnotes()
                if footnotes and hasattr(footnotes, 'getCount') and footnotes.getCount() > 0:
                    print(f"\n📝 发现 {footnotes.getCount()} 个脚注")
                    for i in range(footnotes.getCount()):
                        try:
                            footnote = footnotes.getByIndex(i)
                            self.process_footnote(footnote)
                            if self.current_target_index >= self.total_lines:
                                return
                        except Exception as e:
                            print(f"  无法访问脚注 {i}: {str(e)}")
        except Exception as e:
            print(f"  ⚠️ 处理脚注出错: {str(e)}")
        
        if self.current_target_index >= self.total_lines:
            return
        
        # 6. 处理尾注
        try:
            if hasattr(self.doc, 'getEndnotes'):
                endnotes = self.doc.getEndnotes()
                if endnotes and hasattr(endnotes, 'getCount') and endnotes.getCount() > 0:
                    print(f"\n📝 发现 {endnotes.getCount()} 个尾注")
                    for i in range(endnotes.getCount()):
                        try:
                            endnote = endnotes.getByIndex(i)
                            self.process_endnote(endnote)
                            if self.current_target_index >= self.total_lines:
                                return
                        except Exception as e:
                            print(f"  无法访问尾注 {i}: {str(e)}")
        except Exception as e:
            print(f"  ⚠️ 处理尾注出错: {str(e)}")
        
        # 7. 如果还有未处理的行，使用最简单的查找API
        if self.current_target_index < self.total_lines:
            self.search_with_find_api_simple()
    
    def get_results(self):
        return self.colored_count, self.current_target_index

def process_doc_ultimate(input_path, output_path):
    # 连接LibreOffice
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        print("❌ 无法连接LibreOffice服务，请先运行：")
        print("   libreoffice --headless --norestore --invisible --accept='socket,host=localhost,port=2002;urp;' &")
        sys.exit(1)

    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 打开文档
    input_abs = os.path.abspath(input_path)
    in_url = uno.systemPathToFileUrl(input_abs)
    try:
        doc = desktop.loadComponentFromURL(in_url, "_blank", 0, ())
    except Exception as e:
        print(f"❌ 打开文档失败：{str(e)}")
        sys.exit(1)

    # 获取文档中的所有文本行
    text = doc.getText()
    full_text = text.getString()
    doc_lines = full_text.split('\n')
    
    print(f"\n📄 文档包含 {len(doc_lines)} 行文本")
    
    # 创建着色器并处理文档
    colorizer = LineColorizer(doc, doc_lines)
    colorizer.process_document()
    
    colored_count, processed_count = colorizer.get_results()

    # 保存DOCX
    output_abs = os.path.abspath(output_path)
    out_url = uno.systemPathToFileUrl(output_abs)
    try:
        doc.storeAsURL(out_url, ())
        print(f"\n✅ DOCX保存成功：{output_abs}")
    except Exception as e:
        print(f"❌ DOCX保存失败：{str(e)}")
        sys.exit(1)

    # 导出PDF
    try:
        pdf_path = os.path.splitext(output_abs)[0] + ".pdf"
        pdf_url = uno.systemPathToFileUrl(pdf_path)
        
        filter_data = []
        f = PropertyValue()
        f.Name = "FilterName"
        f.Value = "writer_pdf_Export"
        filter_data.append(f)
        
        doc.storeToURL(pdf_url, tuple(filter_data))
        print(f"✅ PDF导出成功：{pdf_path}")
    except Exception as e:
        print(f"⚠️ PDF导出失败：{str(e)}")

    # 关闭文档
    doc.close(True)
    print(f"\n📊 统计 | 总行数：{len(doc_lines)} | 成功上色行数：{colored_count}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("📋 正确用法：")
        print(f"   python3 {sys.argv[0]} 输入.docx 输出.docx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    if not os.path.exists(input_file):
        print(f"❌ 文件不存在：{input_file}")
        sys.exit(1)

    process_doc_ultimate(input_file, output_file)