import pandas as pd
from openai import OpenAI
import streamlit as st
import os
import zipfile
from io import BytesIO

# Custom CSS for RTL and Persian fonts
st.markdown("""
    <style>
        body {
            font-family: 'B Nazanin', Tahoma, Arial, sans-serif;
            direction: rtl;
            text-align: right;
        }
        h1, h2, h3, h4, h5, h6 {
            font-weight: bold;
            text-align: right;
        }
        p {
            text-align: right;
        }
    </style>
    """, unsafe_allow_html=True)

# Streamlit app title
st.markdown("<h1 style='text-align: right;'>سوالات شخصی‌سازی شده برای دانش‌آموزان</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>لطفاً اطلاعات مورد نیاز را وارد کنید.</p>", unsafe_allow_html=True)

# Input fields for API key and temperature
api_key = st.text_input("API Key OpenAI", type="password")
temp = st.slider("دمای مدل (Temperature)", min_value=0.1, max_value=1.0, value=0.5, step=0.1)

# Upload Excel file
uploaded_file = st.file_uploader("فایل اکسل اطلاعات دانش‌آموزان را آپلود کنید", type=["xlsx"])

# Function to delete files
def delete_files(files):
    for file in files:
        if os.path.exists(file):
            os.remove(file)

# Function to save HTML files
def save_text_with_math_to_html_in_drive(text, output_file):
    html_template = f"""
    <!DOCTYPE html>
    <html lang="fa" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>سوالات شخصی‌سازی شده</title>
        <script type="text/javascript" async
            src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js">
        </script>
        <style>
            body {{ font-family: B Nazanin, sans-serif; line-height: 1.6; }}
            h1, h2, h3, h4, h5, h6 {{ font-weight: bold; }}
            ul, ol {{ margin: 0; padding: 0 1.5em; }}
            p {{ margin: 0.5em 0; }}
            .math {{ text-align: center; }}
        </style>
    </head>
    <body>
        {text}
    </body>
    </html>
    """
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(html_template)

# Function to create a zip file
def create_zip(files):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file in files:
            zip_file.write(file, os.path.basename(file))
    zip_buffer.seek(0)
    return zip_buffer

# Main function
def generate_questions(api_key, temp, uploaded_file):
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, sheet_name=None)

        os.environ["GITHUB_TOKEN"] = api_key
        token = os.environ["GITHUB_TOKEN"]
        endpoint = "https://models.inference.ai.azure.com"
        model_name = "gpt-4o"

        client = OpenAI(
        base_url=endpoint,
        api_key=token,
        )

        # List to store generated HTML files
        html_files = []

        # Progress bar
        progress_bar = st.progress(0)
        total_students = len(data['Sheet2'])
        current_student = 0

        # Placeholder for student name
        student_name_placeholder = st.empty()

        sys_msg = f'''
            وظیفه ی تو طراحی سوالات متناسب سطح دانش آموز با توجه به اطلاعات وارد شده از او توسط معلم است.
            درس مد نظر برای طراحی سوالات: {data['Sheet1']['نام درس']}
            عنوان مبحث مورد نظر برای طراحی سوالات: {data['Sheet1']['نام مبحث']}
            معلم میانگین نمرات امتحان دانش آموز و توضیحاتی در مورد خصوصیات دانش آموز را برای تو ارسال می کند. سپس تو باید سوالاتی مختص دانش آموز با توجه به سطح او مطرح کنی.
            تعداد سوالی که برای دانش آموز باید طرح کنی: {data['Sheet1']['تعداد سوال به ازای هر دانش آموز']}
            ابتدا سطح دانش آموز را مشخص کن و توضیح مختصری در مورد اینکه چطور وضعیت خود را بهبود دهد، بنویس. سپس سوالات را مورد به مورد ارائه بده و سطح هر سوال که طرح می کنی را در سه سطح دشوار، متوسط یا آسان مشخص کن.
            پس از طراحی سوالات، در پایان نیز پاسخ هر کدام از سوالات را با توضیحات صمیمانه و با لحنی طنز، برای دانش آموز به صورتی که دانش آموز کاملا متوجه درس شود بنویس.

            خروجی حتماً با استفاده از تگ های html در که بایستی در تگ body قرار بگیرند نوشته شود.
            خروجی مستقیماً در صفحه ی وب نمایش داده می شود، پس از تگ های اچ تی ام ال برای فرمول ها و سایر قسمت ها استفاده کن.

            خروجی خطاب به دانش آموز نوشته شود.

            سایر توضیحات در مورد طرح سوال که باید مد نظر باشد : {data['Sheet1']['توضیحات کمکی در مورد نحوه ی طراحی سوال']}
            '''
    

        for _, row in data['Sheet2'].iterrows():
            current_student += 1
            student_name = row['نام و نام خانوادگی دانش آموز']
            # Display student name and then clear it
            student_name_placeholder.markdown(f"<p style='text-align: right;'>در حال تولید سوالات برای دانش‌آموز: <strong>{student_name}</strong></p>", unsafe_allow_html=True)

            user_msg = f'''
            نام دانش آموز: {row['نام و نام خانوادگی دانش آموز']}
            میانگین نمره ی دانش آموز از 20 نمره: {row['میانگین نمره ی دانش آموز از 20 نمره']}
            توضیحات در مورد عملکرد و ویژگی های دانش آموز: {row['توضیحات']}
            '''

            response = client.chat.completions.create(
                messages=[
                    {'role': 'system', 'content': sys_msg},
                    {'role': 'user', 'content': user_msg}
                ],
                temperature=temp,
                model=model_name
                                        )

            response = response.choices[0].message.content

            # Save HTML file
            student_name = row["نام و نام خانوادگی دانش آموز"]
            output_file = f"{student_name}.html"
            save_text_with_math_to_html_in_drive(response, output_file)
            html_files.append(output_file)

            # Update progress bar
            progress_bar.progress(current_student / total_students)

            # Clear student name for the next one
            student_name_placeholder.empty()

        # Create a zip file
        st.success("تولید سوالات با موفقیت انجام شد!")
        zip_buffer = create_zip(html_files)
        
        # Download button for the zip file
        if st.download_button(
            label="دانلود تمام سوالات به صورت فایل ZIP",
            data=zip_buffer,
            file_name="students_questions.zip",
            mime="application/zip"
        ):
            # Delete files after download
            delete_files(html_files)
            st.info("فایل‌ها از سرور حذف شدند.")

# Run the app
if api_key and uploaded_file:
    if st.button("تولید سوالات"):
        generate_questions(api_key, temp, uploaded_file)