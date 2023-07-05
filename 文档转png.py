import os
import tempfile
from docx2pdf import convert
from PIL import Image
import fitz

def convert_docx_to_pdf(docx_path, pdf_path):
    convert(docx_path, pdf_path)

def convert_pdf_to_images(pdf_path, dpi=300):
    doc = fitz.open(pdf_path)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        mode = "RGB" if pix.n - pix.alpha < 4 else "RGBA"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        images.append(img)
    return images

def combine_images(images, output_path):
    widths, heights = zip(*(i.size for i in images))
    total_height = sum(heights)
    max_width = max(widths)

    new_image = Image.new('RGB', (max_width, total_height))
    y_offset = 0
    for image in images:
        new_image.paste(image, (0, y_offset))
        y_offset += image.size[1]

    new_image.save(output_path, "PNG")

def main():
    # 获取脚本执行目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print("脚本执行目录:", script_dir)

    # 查找脚本所在目录下的所有docx文件
    docx_files = [f for f in os.listdir(script_dir) if f.endswith(".docx")]

    for docx_file in docx_files:
        docx_path = os.path.join(script_dir, docx_file)
        pdf_path = os.path.join(script_dir, f"{os.path.splitext(docx_file)[0]}.pdf")
        png_file = os.path.join(script_dir, f"{os.path.splitext(docx_file)[0]}.png")

        # 转换docx为pdf
        convert_docx_to_pdf(docx_path, pdf_path)

        # 将pdf转换为images（设置dpi为300）
        images = convert_pdf_to_images(pdf_path, dpi=300)

        # 合并多个images为一张长图
        combine_images(images, png_file)

        print(f"转换完成：{docx_file} -> {os.path.basename(png_file)}")

if __name__ == "__main__":
    main()
