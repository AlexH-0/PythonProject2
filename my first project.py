# Excel 数据批量生成 PPT 工具
#一个基于 Python 的自动化工具，支持从 Excel 读取姓名和成绩数据，自动匹配文件夹中的图片，批量生成 PPT 演示文稿。
#功能
"""- 自动识别 Excel 表头方向（行表头 / 竖列表头）
- 支持按行或按列输入数据位置（如 `C` 列或 `1` 行）
- 按姓名自动匹配文件夹中的图片，忽略后缀名差异
- 自动处理图片缺失情况，生成标注提示
- 支持自定义模板版式"""
## 使用方法

#报错了别找我，因为我也不会改

# PythonProject2
#作用：读取excel数据和图片，写入ppt
import os,openpyxl
from pptx import *
class homework():
    def jd(self, path1, path2, path3):
        workbook = openpyxl.load_workbook(path1)
        sheet = workbook.active
        name_input = input("请输入姓名所在行/列序号：") or "C"
        score_input = input("请输入成绩所在行/列序号：") or "D"
    # 读取姓名
        if name_input.isdigit():
            names = [cell.value for cell in sheet[int(name_input)]]
        else:
            names = [cell.value for cell in sheet[name_input.upper()]]
        names = [n for n in names if n]
    # 读取成绩
        if score_input.isdigit():
            scores = [cell.value for cell in sheet[int(score_input)]]
        else:
            scores = [cell.value for cell in sheet[score_input.upper()]]
        scores = [s for s in scores if s]
    # 图片匹配
        def find_image(folder, name):
            for f in os.listdir(folder):
                if f.split(".")[0] == name:
                    return os.path.join(folder, f)
            return None
    # 新建PPT
        pptx_file = Presentation(path3)
        layout = pptx_file.slide_layouts[8]
        missing=[]
        for i, name in enumerate(names):
            img_path = find_image(path2, name)
            slide = pptx_file.slides.add_slide(layout)
            if img_path:
            # 写入姓名
                slide.placeholders[0].text = name
            # 添加图片
                ph = slide.placeholders[1]
                left, top, width, height = ph.left, ph.top, ph.width, ph.height
                sp = ph._element
                sp.getparent().remove(sp)
                slide.shapes.add_picture(img_path, left, top, width, height)
            # 写入成绩
                slide.placeholders[2].text = str(scores[i])
            else:
                missing.append(name)
                slide.placeholders[2].text = f"缺失图片的同学：\n姓名：{name},成绩：{scores[i]}\n"
        pptx_file.save(path3)
        print("保存完毕")
