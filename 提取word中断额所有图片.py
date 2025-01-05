import os
from docx import Document
from PIL import Image
from io import BytesIO

def extract_images_by_paragraph(docx_path, output_folder):
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)

    # 打开 Word 文档
    doc = Document(docx_path)
    image_counter = 1  # 初始化图片计数器

    # 定义命名空间
    nsmap = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }

    # 遍历段落，按顺序提取图片
    for paragraph_index, paragraph in enumerate(doc.paragraphs):
        print(f"分析段落 {paragraph_index + 1}: {paragraph.text}")

        # 遍历段落中的运行（Run）
        for run in paragraph.runs:
            # 检查运行中是否包含图片
            drawing_elements = run.element.findall(".//w:drawing", namespaces=nsmap)
            if not drawing_elements:
                continue

            for drawing in drawing_elements:
                # 提取嵌入的图片关系 ID
                blip_elements = drawing.findall(".//a:blip", namespaces=nsmap)
                for blip in blip_elements:
                    embed_id = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if embed_id and embed_id in doc.part.rels:
                        image_part = doc.part.rels[embed_id].target_part
                        image_data = image_part.blob

                        # 确定图片格式
                        image = Image.open(BytesIO(image_data))
                        image_format = image.format.lower()

                        # 保存图片
                        image_filename = f"paragraph_{paragraph_index + 1}_image_{image_counter}.{image_format}"
                        image_path = os.path.join(output_folder, image_filename)
                        image.save(image_path)

                        print(f"已保存图片: {image_path}")
                        image_counter += 1

# 示例使用
docx_file = "测试测试0.docx"  # 替换为你的文件路径
output_dir = "tmp_img"
extract_images_by_paragraph(docx_file, output_dir)
