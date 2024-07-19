现在我有一张Excel表格和模板PPT文件，A列到G列，第一行是标题行（A到E每一列都是标题行，F和G在第一行合并了单元格），现在需要用Python提取Excel中第一行(标题行)(A列到G列)和第二行内容行组成小表格，插入到PPT的第二页(遵循该页的字体模板)，第二页PPT标题采用当前行的B列的内容，第一行(标题行)再和第三行内容行组成小表哥，插入到PPT的第三页，第三页PPT标题采用当前行的B列的内容,以此类推，直到第32行。

``` python
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

# 读取Excel文件
excel_path = 'path_to_your_excel_file.xlsx'
df = pd.read_excel(excel_path, header=0)

# 读取模板PPT文件
template_ppt_path = 'path_to_your_template_ppt.pptx'
ppt = Presentation(template_ppt_path)
slide_layout = ppt.slide_layouts[1]  # 使用适当的布局

# 提取标题行
headers = df.columns.tolist()

# 遍历每一行，创建小表格并插入到PPT中
for index, row in df.iloc[1:32].iterrows():  # 从第二行开始到第32行
    slide = ppt.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = row['B']  # 将B列内容作为PPT标题

    # 创建表格
    table = slide.shapes.add_table(rows=2, cols=len(headers), left=Inches(1), top=Inches(1.5), width=Inches(8), height=Inches(1.0)).table
    
    # 设置标题行
    for col_idx, header in enumerate(headers):
        table.cell(0, col_idx).text = header
    
    # 设置数据行
    for col_idx, value in enumerate(row):
        table.cell(1, col_idx).text = str(value)

# 保存PPT文件
output_ppt_path = 'output_presentation.pptx'
ppt.save(output_ppt_path)
```
