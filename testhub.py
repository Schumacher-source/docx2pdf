import os
from pdf2image import convert_from_path
from PIL import Image


def pdf_to_png(pdf_path, output_dir):
    # 将PDF转换为PNG图片
    images = convert_from_path(pdf_path)
    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]

    for i, image in enumerate(images):
        image.save(f"{output_dir}/{base_filename}_{i}.png", "PNG")


def png_to_pdf(png_dir, output_pdf):
    # 将PNG图片转换为PDF
    image_files = [f"{png_dir}/{file}" for file in os.listdir(png_dir) if file.lower().endswith(".png")]
    image_files.sort()  # 确保按顺序处理图像

    images = []
    for image_file in image_files:
        image = Image.open(image_file)
        images.append(image.convert("RGB"))

    images[0].save(output_pdf, save_all=True, append_images=images[1:], optimize=False)


# 示例用法
pdf_path = "input.pdf"
png_output_dir = "png_output"
converted_pdf_path = "output.pdf"

# 创建输出目录
os.makedirs(png_output_dir, exist_ok=True)

# 将PDF转换为PNG
pdf_to_png(pdf_path, png_output_dir)

# 将PNG转换为PDF
png_to_pdf(png_output_dir, converted_pdf_path)

print("转换完成！")
