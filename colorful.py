import os
import sys
import time
import uno
import zipfile
import subprocess
import shutil
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
        """初始化行着色器"""
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
        """为找到的文本范围着色"""
        try:
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
            para_text = self.safe_get_string(paragraph)
            if not para_text:
                return False
            
            print(f"\n📄 [{source}] 处理: {para_text[:50]}...")
            
            processed = False
            while self.current_target_index < self.total_lines:
                target = self.target_lines[self.current_target_index]
                
                if not target.strip():
                    print(f"  ⏭️ 行{self.current_target_index+1} 是空行，跳过")
                    self.current_target_index += 1
                    processed = True
                    continue
                
                pos = para_text.find(target)
                if pos >= 0:
                    before_ok = (pos == 0 or para_text[pos-1] in '\n\r')
                    after_ok = (pos + len(target) >= len(para_text) or 
                               para_text[pos + len(target)] in '\n\r')
                    
                    if before_ok and after_ok:
                        print(f"  🔍 找到目标行 {self.current_target_index+1} 在位置 {pos}")
                        
                        start = paragraph.getStart()
                        cursor = self.text.createTextCursorByRange(start)
                        
                        for _ in range(pos):
                            cursor.goRight(1, False)
                        
                        end_cursor = self.text.createTextCursorByRange(cursor.getStart())
                        for _ in range(len(target)):
                            end_cursor.goRight(1, False)
                        
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
            cell_text = self.safe_get_string(cell)
            if not cell_text:
                return False
            
            print(f"\n  📋 处理单元格: {cell_text[:50]}...")
            
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
                    
                    start = cell.getStart()
                    cursor = self.text.createTextCursorByRange(start)
                    
                    for _ in range(pos):
                        cursor.goRight(1, False)
                    
                    end_cursor = self.text.createTextCursorByRange(cursor.getStart())
                    for _ in range(len(target)):
                        end_cursor.goRight(1, False)
                    
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
        """使用查找API搜索剩余行"""
        if self.current_target_index >= self.total_lines:
            return
        
        print(f"\n🔍 使用查找API搜索剩余行 (从行{self.current_target_index+1}开始)...")
        
        self.init_search()
        start_point = self.text.getStart()
        
        while self.current_target_index < self.total_lines:
            target = self.target_lines[self.current_target_index]
            
            if not target.strip():
                print(f"  ⏭️ 行{self.current_target_index+1} 是空行，跳过")
                self.current_target_index += 1
                continue
            
            print(f"\n  🔍 搜索: {target[:50]}")
            
            self.search_desc.setSearchString(target)
            found = self.doc.findNext(start_point, self.search_desc)
            
            if found:
                print(f"  ✅ 找到行{self.current_target_index+1}")
                print(f"     找到内容: {found.getString()[:50]}")
                
                if self.colorize_text_range_simple(found):
                    start_point = found.getEnd()
                else:
                    start_point = found.getEnd()
                    self.current_target_index += 1
            else:
                print(f"  ❌ 无法找到行{self.current_target_index+1}")
                self.current_target_index += 1
                start_point = self.text.getStart()
    
    def process_document(self):
        """处理整个文档"""
        print("\n🔍 开始遍历文档所有内容...")
        
        # 处理普通段落
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
        
        # 处理表格
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
        
        # 处理文本框
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
        
        # 处理页眉页脚
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
        
        # 处理脚注
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
        
        # 处理尾注
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
        
        # 使用查找API处理剩余行
        if self.current_target_index < self.total_lines:
            self.search_with_find_api_simple()
    
    def get_results(self):
        return self.colored_count, self.current_target_index

def run_command(cmd):
    """执行系统命令并返回结果（兼容无stdout/stderr情况）"""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        return True, result.stdout + result.stderr
    except subprocess.CalledProcessError as e:
        return False, f"命令执行失败(码:{e.returncode}): {e.stderr}"
    except Exception as e:
        return False, f"命令执行异常: {str(e)}"

def clean_libreoffice_marks(docx_path):
    """清理DOCX文件中的LibreOffice私有标记，确保符合MS Word 2007规范"""
    if not os.path.exists(docx_path) or os.path.getsize(docx_path) < 100:
        print(f"⚠️ DOCX文件无效，跳过清理")
        return False
    
    print(f"\n🧹 清理LibreOffice私有标记，适配MS Word 2007格式...")
    temp_dir = f"{docx_path}.clean.tmp"
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # 解压DOCX文件（兼容压缩格式）
        with zipfile.ZipFile(docx_path, 'r', zipfile.ZIP_DEFLATED) as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 清理核心XML文件中的LibreOffice标记
        clean_files = [
            os.path.join(temp_dir, 'word', 'document.xml'),
            os.path.join(temp_dir, 'word', 'styles.xml'),
            os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels'),
            os.path.join(temp_dir, '[Content_Types].xml')
        ]
        
        for file_path in clean_files:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # 移除LibreOffice私有命名空间和标记
                content = content.replace('xmlns:lo=', 'xmlns:tmp_lo=')
                content = content.replace('lo:', 'tmp_lo:')
                content = content.replace('xmlns:office=', 'xmlns:tmp_office=')
                content = content.replace('office:', 'tmp_office:')
                content = content.replace('xmlns:odf=', 'xmlns:tmp_odf=')
                content = content.replace('odf:', 'tmp_odf:')
                
                # 修复Word 2007不兼容的属性值
                content = content.replace('w:color="auto"', 'w:color="#000000"')
                content = content.replace('w:highlight="auto"', 'w:highlight="#FFFFFF"')
                content = content.replace('w:val="auto"', 'w:val="normal"')
                
                with open(file_path, 'w', encoding='utf-8', errors='ignore') as f:
                    f.write(content)
        
        # 重新打包为标准DOCX（覆盖原文件）
        os.remove(docx_path)
        with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_abs = os.path.join(root, file)
                    file_rel = os.path.relpath(file_abs, temp_dir)
                    zipf.write(file_abs, file_rel)
        
        print(f"✅ DOCX标记清理完成，文件大小：{os.path.getsize(docx_path)/1024:.1f}KB")
        return True
        
    except Exception as e:
        print(f"⚠️ DOCX清理失败: {str(e)}")
        return False
    finally:
        # 强制清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

def convert_odt_to_word2007_docx(odt_path, docx_path):
    """
    强制转换为MS Word 2007 DOCX格式
    核心改进：直接指定输出文件名，放弃重命名，使用LibreOffice原生--outfile参数
    """
    if not os.path.exists(odt_path) or os.path.getsize(odt_path) < 100:
        print(f"❌ 源ODT文件无效：{odt_path}")
        return False
    
    # 确保输出目录存在，删除已存在的目标文件（避免占用）
    output_dir = os.path.dirname(docx_path)
    os.makedirs(output_dir, exist_ok=True)
    if os.path.exists(docx_path):
        os.remove(docx_path)
    # 临时转换文件（防止LibreOffice自动加后缀）
    temp_convert = os.path.join(output_dir, f"temp_rowdye_{os.getpid()}.docx")
    if os.path.exists(temp_convert):
        os.remove(temp_convert)
    
    print(f"\n🔄 强制转换为 MS Word 2007 (.docx) 格式...")
    print(f"   源文件：{odt_path}")
    print(f"   目标文件：{docx_path}")
    
    # 核心命令：直接指定输出文件，无重命名，兼容所有LibreOffice版本
    cmd = (
        f"libreoffice --headless --norestore --invisible --nofirststartwizard "
        f"--convert-to docx:'Microsoft Word 2007/2010/2013 (.docx)' "
        f"--outdir {output_dir} "
        f"{odt_path} "
        f"&& mv {os.path.join(output_dir, os.path.basename(os.path.splitext(odt_path)[0]) + '.docx')} {temp_convert}"
    )
    
    # 执行转换
    success, output = run_command(cmd)
    if success and os.path.exists(temp_convert) and os.path.getsize(temp_convert) > 100:
        # 移动临时文件到目标路径
        shutil.move(temp_convert, docx_path)
        # 清理LibreOffice标记
        clean_libreoffice_marks(docx_path)
        # 最终校验
        if os.path.exists(docx_path) and os.path.getsize(docx_path) > 100:
            print(f"✅ MS Word 2007 DOCX转换成功：{docx_path}")
            return True
        else:
            print(f"❌ 转换后文件无效（空文件）")
    else:
        print(f"❌ 主转换命令失败：{output[:200]}")
        # 兜底方案1：简化参数转换
        print("\n🔄 兜底方案1：使用最简参数转换...")
        cmd_simple = (
            f"libreoffice --headless --convert-to docx {odt_path} --outdir {output_dir} "
            f"&& mv {os.path.join(output_dir, os.path.basename(os.path.splitext(odt_path)[0]) + '.docx')} {docx_path}"
        )
        success_s, output_s = run_command(cmd_simple)
        if success_s and os.path.exists(docx_path) and os.path.getsize(docx_path) > 100:
            clean_libreoffice_marks(docx_path)
            print(f"✅ 兜底方案1转换成功")
            return True
        
        # 兜底方案2：先转DOC再转DOCX（最高兼容性）
        print("\n🔄 兜底方案2：DOC中转再转DOCX（最高兼容性）...")
        temp_doc = os.path.join(output_dir, f"temp_rowdye_{os.getpid()}.doc")
        cmd_doc = f"libreoffice --headless --convert-to doc {odt_path} --outdir {output_dir}"
        run_command(cmd_doc)
        if os.path.exists(temp_doc):
            cmd_doc2docx = f"libreoffice --headless --convert-to docx {temp_doc} --outdir {output_dir} && mv {os.path.splitext(temp_doc)[0] + '.docx'} {docx_path}"
            success_d, output_d = run_command(cmd_doc2docx)
            if success_d and os.path.exists(docx_path) and os.path.getsize(docx_path) > 100:
                clean_libreoffice_marks(docx_path)
                os.remove(temp_doc)
                print(f"✅ 兜底方案2转换成功")
                return True
            os.remove(temp_doc)
    
    # 最终清理临时文件
    if os.path.exists(temp_convert):
        os.remove(temp_convert)
    print(f"❌ 所有转换方案均失败")
    return False

def convert_odt_to_pdf(odt_path, pdf_path):
    """转换ODT到PDF（稳定版）"""
    if not os.path.exists(odt_path):
        print(f"❌ 源ODT文件不存在，跳过PDF转换")
        return False
    
    print(f"\n🔄 转换为PDF格式...")
    output_dir = os.path.dirname(pdf_path)
    os.makedirs(output_dir, exist_ok=True)
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    
    cmd = (
        f"libreoffice --headless --norestore --invisible "
        f"--convert-to pdf:writer_pdf_Export:{{\"Quality\":\"90\",\"UseLosslessCompression\":\"true\"}} "
        f"--outdir {output_dir} {odt_path}"
    )
    success, output = run_command(cmd)
    # 校验PDF文件
    converted_pdf = os.path.splitext(odt_path)[0] + ".pdf"
    if success and os.path.exists(converted_pdf) and os.path.getsize(converted_pdf) > 100:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        shutil.move(converted_pdf, pdf_path)
        print(f"✅ PDF转换成功：{pdf_path}")
        return True
    else:
        print(f"⚠️ PDF转换失败：{output[:200]}")
        return False

def process_doc_ultimate(input_path, output_path):
    """主处理函数 - 强制MS Word 2007格式转换+多层兜底"""
    # 路径规范化（解决绝对/相对路径问题）
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)
    
    # 临时ODT文件（放在输出目录，避免跨路径问题）
    temp_odt = os.path.join(output_dir, f"rowdye_temp_{os.getpid()}.odt")
    # 确保临时文件无残留
    for f in [temp_odt, temp_odt.replace('.odt', '.docx'), output_path]:
        if os.path.exists(f):
            os.remove(f)
    
    # 连接LibreOffice UNO服务
    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
    except Exception as e:
        print("❌ 无法初始化UNO上下文：")
        print(f"   错误详情: {str(e)}")
        print(f"   请安装：sudo apt install libreoffice-python3 python3-uno")
        sys.exit(1)
        
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        print("❌ 无法连接LibreOffice服务，请先运行：")
        print("   libreoffice --headless --norestore --invisible --accept='socket,host=localhost,port=2002;urp;' &")
        sys.exit(1)

    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 打开原始文档（只读模式，避免文件锁）
    in_url = uno.systemPathToFileUrl(input_path)
    doc = None
    open_props = [PropertyValue(Name="ReadOnly", Value=True)]
    try:
        doc = desktop.loadComponentFromURL(in_url, "_blank", 0, tuple(open_props))
        if not doc:
            raise Exception("无法加载文档对象（空对象）")
    except Exception as e:
        print(f"❌ 打开文档失败：{str(e)}")
        sys.exit(1)

    colored_count = 0
    try:
        # 获取文档文本行并着色
        text = doc.getText()
        full_text = text.getString()
        doc_lines = full_text.split('\n')
        doc_lines = [line.rstrip('\r') for line in doc_lines]  # 统一换行符
        
        print(f"\n📄 文档包含 {len(doc_lines)} 行文本")
        if len(doc_lines) == 0:
            raise Exception("文档无文本内容")
        
        colorizer = LineColorizer(doc, doc_lines)
        colorizer.process_document()
        colored_count, _ = colorizer.get_results()

        # 第一步：保存为ODT格式（原生格式，确保数据完整）
        odt_url = uno.systemPathToFileUrl(temp_odt)
        doc.storeAsURL(odt_url, ())
        time.sleep(2)  # 等待写入完成
        
        # 校验ODT文件
        if not os.path.exists(temp_odt) or os.path.getsize(temp_odt) < 100:
            raise Exception(f"ODT保存失败，文件无效：{temp_odt}")
        print(f"✅ 临时ODT保存成功：{temp_odt}（大小：{os.path.getsize(temp_odt)/1024:.1f}KB）")

    except Exception as e:
        print(f"❌ 文档处理/保存失败：{str(e)}")
        if doc:
            doc.close(True)
        sys.exit(1)
    finally:
        # 强制关闭文档，释放所有文件锁
        if doc:
            try:
                doc.close(True)
            except:
                pass
        # 等待LibreOffice释放文件锁（关键）
        time.sleep(3)

    # 第二步：强制转换为MS Word 2007 DOCX格式（核心修复）
    docx_success = convert_odt_to_word2007_docx(temp_odt, output_path)
    
    # 第三步：转换PDF（不依赖DOCX结果）
    pdf_path = os.path.splitext(output_path)[0] + ".pdf"
    convert_odt_to_pdf(temp_odt, pdf_path)
    
    # 强制清理临时文件（无论成败）
    if os.path.exists(temp_odt):
        os.remove(temp_odt)
    # 清理所有临时文件
    for f in os.listdir(output_dir):
        if f.startswith("temp_rowdye_") or f.startswith("rowdye_temp_"):
            os.remove(os.path.join(output_dir, f))
    print(f"🗑️ 临时文件已全部清理")

    # 输出统计信息
    print(f"\n📊 处理统计 | 文档总行数：{len(doc_lines)} | 成功上色行数：{colored_count}")
    
    # 最终校验DOCX文件
    if docx_success and os.path.exists(output_path) and os.path.getsize(output_path) > 100:
        print(f"\n🎉 全部处理完成！生成的DOCX文件可直接用MS Word 2007+打开")
    else:
        print(f"\n❌ DOCX文件生成失败！请检查LibreOffice安装是否完整")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("📋 RowDye - DOCX行级着色工具 正确用法：")
        print(f"   python3 {sys.argv[0]} 输入.docx 输出.docx")
        print(f"   示例：python3 {sys.argv[0]} input.docx output.docx")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    # 基础验证
    if not os.path.exists(input_file):
        print(f"❌ 输入文件不存在：{input_file}")
        sys.exit(1)
    if not input_file.lower().endswith('.docx'):
        print(f"⚠️ 警告：输入文件不是DOCX格式，可能存在兼容性问题")
    if os.path.dirname(output_file) and not os.path.exists(os.path.dirname(output_file)):
        print(f"⚠️ 输出目录不存在，将自动创建：{os.path.dirname(output_file)}")
    
    # 执行主处理
    process_doc_ultimate(input_file, output_file)
