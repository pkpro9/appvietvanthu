import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import json
import os
import re
from datetime import datetime

def get_latest_sequential_number():
    if os.path.exists('sequential_number.json'):
        with open('sequential_number.json', 'r') as file:
            data = json.load(file)
            return data.get("latest_number", 1)
    return 1

def save_latest_sequential_number(number):
    with open('sequential_number.json', 'w') as file:
        json.dump({"latest_number": number}, file)

def fill_form(template_path, output_path, fields):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in fields.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
                run = paragraph.runs[0]
                run.font.name = 'Times New Roman'
                run.font.size = Pt(13)
    doc.save(output_path)

def sanitize_filename(filename):
    filename = re.sub(r'[/*?"<>|:\n]', '', filename)
    return filename.upper()

def main():
    st.title("Ứng dụng Điền Tờ Trình")

    st.sidebar.title("Cài đặt")
    template_path = st.sidebar.text_input("Đường dẫn file mẫu", "template.docx")
    output_directory = st.sidebar.text_input("Thư mục lưu file kết quả", "./")

    sequential_number = get_latest_sequential_number()
    st.text(f"Số thứ tự gần đây nhất: {sequential_number}")
    
    entry_1 = st.text_input("Số tờ trình", str(sequential_number))
    content_5 = st.text_area("Nội dung tờ trình")
    content_6 = st.text_area("Kính gửi")
    content_7 = st.text_area("Thực trạng")
    content_8 = st.text_area("Nguyên nhân/Diễn giải")
    content_9 = st.text_area("Giải pháp đề xuất")
    content_10 = st.text_area("Khoa Xét nghiệm kính trình")

    if st.button("Tạo Tờ Trình"):
        fields = {
            "(1)": entry_1,
            "(2)": str(datetime.now().day),
            "(3)": str(datetime.now().month),
            "(4)": str(datetime.now().year),
            "(5)": content_5,
            "(6)": content_6,
            "(7)": content_7,
            "(8)": content_8,
            "(9)": content_9,
            "(10)": content_10,
        }

        output_filename = sanitize_filename(f"{fields['(1)']}_to_trinh.docx")
        output_path = os.path.join(output_directory, output_filename)

        try:
            fill_form(template_path, output_path, fields)
            save_latest_sequential_number(int(entry_1) + 1)
            st.success(f"Tờ trình đã được tạo tại: {output_path}")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")

if __name__ == "__main__":
    main()
