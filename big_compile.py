'''
Created on Mar 10, 2021

@author: austin
'''
import docx
from docx import Document
from functions.format_preserver import format_preserver2

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def big_compile2(doc_names, section_names, output_name):
    the_big_one = Document()
    for section in range(len(section_names)):
        the_big_one.add_paragraph(section_names[section],'Heading 1')
        for name in doc_names:
            doc = Document(name)
            delete_paragraph(doc.paragraphs[0])
            while len(doc.paragraphs) > 0:
                if doc.paragraphs[0].style.name == 'Heading 1':
                    break
                else:
                    format_preserver2(the_big_one, doc.paragraphs[0])
                    delete_paragraph(doc.paragraphs[0])
            doc.save(name)
    the_big_one.save(output_name)
    return

sections = ['Shells','Links','Impacts','Alts','Answers To']
x = [r"C:\Users\austi\Documents\setcolproj\processed once\7.docx", r"C:\Users\austi\Documents\setcolproj\processed once\5.docx"]
y = r"C:\Users\austi\Documents\setcolproj\processed once\new new test.docx"
big_compile2(x, sections, y)