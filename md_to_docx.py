import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor
import markdown
from bs4 import BeautifulSoup

def add_code_block(doc, code_text):
    p = doc.add_paragraph()
    run = p.add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(10)
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), "DDDDDD")
    p._element.get_or_add_pPr().append(shading_elm)

def convert_md_to_docx(md_file):
    with open(md_file, 'r', encoding='utf-8') as f:
        md_content = f.read()

    html = markdown.markdown(md_content, extensions=['fenced_code', 'codehilite'])
    soup = BeautifulSoup(html, 'html.parser')

    doc = Document()
    md_dir = os.path.dirname(md_file)

    for elem in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'pre']):
        if elem.name == 'pre':
            code = elem.find('code')
            if code:
                add_code_block(doc, code.text)
        elif elem.name == 'p':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            level = int(elem.name[1])
            p = doc.add_heading(level=level)

        if elem.name != 'pre':
            for content in elem.contents:
                if content.name == 'img':
                    img_src = content['src']
                    img_path = os.path.join(md_dir, img_src)
                    if os.path.exists(img_path):
                        doc.add_picture(img_path, width=Inches(6))
                    else:
                        p.add_run(f"[图片未找到: {img_src}]")
                elif content.name == 'code':
                    run = p.add_run(content.text)
                    run.font.color.rgb = RGBColor(0, 0, 255)
                else:
                    p.add_run(str(content))

    docx_file = md_file.replace('.md', '.docx')
    doc.save(docx_file)
    print(f"已将 {md_file} 转换为 {docx_file}")

md_files = [f for f in os.listdir('.') if f.endswith('.md')]

for md_file in md_files:
    convert_md_to_docx(md_file)
