#it randomly didnt work without both of these imports idk why
import docx
from docx import Document
#adds a paragraph to a document while preserving its formatting
def format_preserver2(output_doc_name, paragraph):
    stylename_set = set()
#creates a set of the names of the document's base styles to refer to. I use a 
#set because it's faster than checking a list for large amounts of data. I don't
#know how to map this to a better data type
    for style in output_doc_name.styles:
        stylename_set.add(style.name)
    output_para = output_doc_name.add_paragraph()
#iterates through the paragraph run by run, since that's how styles are
#distributed
    for run in paragraph.runs:
#stores the run so it can be edited
        output_run = output_para.add_run(run.text)
#checks if the run's style is in our big set of the document's styles. if not,
#adds that style to the document and puts its name in the set
        if run.style.name not in stylename_set:
            output_doc_name.styles.add_style(run.style.name, run.style.type)
            stylename_set.add(run.style.name)
#sets up default, non-style-based formatting. I don't think this needs to be
#here but i havent tried to take it out
        output_run.bold = run.bold
        output_run.font.highlight_color = run.font.highlight_color
        output_run.italic = run.italic
        output_run.font.size = run.font.size
        output_run.underline = run.underline
        output_run.font.color.rgb = run.font.color.rgb
#sets style of run in new doc
        output_run.style = run.style.name
#fixes overall paragraph styles
    if paragraph.style.name not in stylename_set:
        output_doc_name.styles.add_style(paragraph.style.name, paragraph.style.type)
        stylename_set.add(paragraph.style.name)
    output_para.style = paragraph.style.name
#fixes alignment
    output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
    return
