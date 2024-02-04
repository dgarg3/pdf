import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib import styles
import textwrap
from reportlab.pdfbase.pdfmetrics import stringWidth
import os

def header(c,text):
    text = f"Feedback Form {text}"
    y_position = 750
    text_width = c.stringWidth(text, "Helvetica-Bold", 18)
    #text_width1 = text_width + 10
    text_width1 = 180
    text_height = 12
    x_position = (letter[0] - text_width) / 2
    # c.drawString(x_position, y_position, text)
    c.drawString(x_position, y_position, text)
    c.line(x_position, y_position - 2, x_position + text_width1, y_position - 2)

def e_name(c,text):
    text = f"Employee Name:    {text}"
    y_position = 725
    #text_width = c.stringWidth(text, "Helvetica", 14)
    #text_width1 = text_width + 40
    text_height = 12
    #x_position = (letter[0] - text_width) / 2
    # c.drawString(x_position, y_position, text)
    x_position = 150
    c.drawString(x_position, y_position, text)
    #c.line(x_position, y_position - 2, x_position + text_width1, y_position - 2)
def e_id(c,text):
    text = f"Employee Id:     {text}"
    y_position = 700
    #text_width = c.stringWidth(text, "Helvetica", 14)
    #text_width1 = text_width + 40
    text_height = 12
    #x_position = (letter[0] - text_width) / 2
    # c.drawString(x_position, y_position, text)
    x_position = 150
    c.drawString(x_position, y_position, text)
    #c.line(x_position, y_position - 2, x_position + text_width1, y_position - 2)

def m_name(c,text):
    text = f"Manager Name:     {text}"
    y_position = 675
    #text_width = c.stringWidth(text, "Helvetica", 14)
    #text_width1 = text_width + 40
    text_height = 12
    #x_position = (letter[0] - text_width) / 2
    # c.drawString(x_position, y_position, text)
    x_position = 150
    c.drawString(x_position, y_position, text)
    #c.line(x_position, y_position - 2, x_position + text_width1, y_position - 2)


def wrap_to_fixed_width(text, width):
    wrapped_lines = textwrap.wrap(text, width=width)
    wrapped_text = '\n'.join(wrapped_lines)
    return wrapped_text


def add_fback(c, data):
    # Set the line spacing
    line_height = 25
    # Determine the starting position for the lines
    x = 150
    #y = c._pagesize[1] - 100
    y = 650
    for index,row in data.iterrows():

        ##initializing the y axis

        ##logic to add additional page
        feeline_text = row['feedback']
        s = wrap_to_fixed_width(feeline_text, 70)
        if y < 150:
            c.showPage()
            y= 750


        #c.setFont(*style)
        line_text = f"Feedback By:  {row['feedbackby']}"
        c.drawString(x, y, line_text)
        y -= line_height

        x_rect = x-10
        y_rect = y+ 10
        padding = 5
        textobject = c.beginText(x,y)

        for line in s.splitlines(False):
            textobject.textLine(line.rstrip())
        adjustment = len(s.splitlines(False))*13
        c.drawText(textobject)
        c.rect(x_rect - padding, y_rect + padding, 360 + 15 * padding, -adjustment - 5 * padding,stroke=1, fill=0)
        y -= line_height +adjustment





def create_pdf(output_file,headertext,empname,mname,df):
    ##directory_path = "/Users/dhirajgarg/python-brain/pdf/output" ## Use your own path
    file_path = os.path.join(directory_path, output_file )

    c = canvas.Canvas(file_path, pagesize=letter)

    font_name, font_size = "Helvetica-Bold", 18
    c.setFont(font_name, font_size)

    header(c, headertext)
    font_name, font_size = "Helvetica", 12
    c.setFont(font_name, font_size)
    print(type(empname))
    e_name(c,empname)
    e_id(c,empname)
    m_name(c, mname)
    add_fback(c, df)
    c.save()





####### Read Excel file
excel_file_path = '"test.xlsx"'
df = pd.read_excel("test.xlsx")
# Write dynamic sections
group_columns = ['feedbackfor','year','Managenamr']
# grouped = df.groupby(group_columns)
# for name,group in grouped:
#     #print(group.drop(columns=group_columns))
#     output_pdf_file = f"{name[0]}_{name[1]}.pdf"
#
#     create_pdf(output_pdf_file, 'abc', 'abc', 'abc', group)
df_g = df.groupby(group_columns)
for name, group in df_g:

    header_1 = str(name[1])
    emp_name_1 = name[0]
    mgr_name_1 = name[2]
    output_pdf_file = f"{emp_name_1}_{header_1}.pdf"
    create_pdf(output_pdf_file, header_1,emp_name_1,mgr_name_1, group)



# for (group_name_1, group_name_2,group_name_3), group in df.groupby(group_columns):
#     # Print group information
#     output_pdf_file = f"{group_name_1}_{group_name_2}.pdf"
#     other_columns = group.drop(columns=group_columns)
#     create_pdf(output_pdf_file, group_name_2, group_name_1, group_name_3, other_columns)


    # Iterate through other columns apart from the group columns
    # for column_name, column_data in group_data.iteritems():
    #     if column_name not in group_columns:
    #         print(f"Column: {column_name}")
    #         print(column_data)

# for (group_name_1, group_name_2,group_name_3), group_data in grouped_data:
#
#     create_pdf(output_pdf_file,group_name_2,group_name_1,group_name_3,group_data)
# for (group_name_1, group_name_2,group_name_3),group in df.groupby([df.columns[0], df.columns[1],df.columns[2]]):
#     output_pdf_file = f"{group_name_1}_{group_name_2}.pdf"
#     #print(group_name_1)
#     create_pdf(output_pdf_file, group_name_2, group_name_1, group_name_3, group)





