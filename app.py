# import zipfile
# import os
# from docx import Document
# import json

# ZIP_PATH = '원짐 PT 평촌본점.zip'  # ← 네 zip 파일 이름으로 변경
# EXTRACT_DIR = 'unzipped'

# # 1. zip 압축 풀기
# with zipfile.ZipFile(ZIP_PATH, 'r') as zip_ref:
#     zip_ref.extractall(EXTRACT_DIR)

# # 2. .docx 찾기
# docx_files = [f for f in os.listdir(EXTRACT_DIR) if f.endswith('.docx')]
# if not docx_files:
#     raise Exception("ZIP 파일 안에 .docx 파일이 없습니다.")

# docx_path = os.path.join(EXTRACT_DIR, docx_files[0])
# doc = Document(docx_path)


# # 3. 표 추출 함수
# def extract_tables(doc):
#     tables = []
#     for table in doc.tables:
#         rows = []
#         for row in table.rows:
#             cells = [cell.text.strip() for cell in row.cells]
#             rows.append(cells)
#         tables.append(rows)

#     parsed_data = {}
#     for item in tables[0]:
#         if len(item) == 2:
#             key = item[0].replace('\n', ' ').strip() # 줄바꿈 문자를 공백으로 바꾸고 앞뒤 공백 제거
#             value = item[1].strip() # 값의 앞뒤 공백 제거

#             # '※ 포스팅 가이드'처럼 정보가 없는 중복 키는 건너뜀
#             if key == '※ 포스팅 가이드' and value == '※ 포스팅 가이드':
#                 continue
            
#             parsed_data[key] = value
    
#     print(parsed_data)
#     return parsed_data

# # def convert_to_json(table):
# #     headers = table[0]
# #     return [dict(zip(headers, row)) for row in table[1:]]

# # # 4. 실행
# tables = extract_tables(doc)
# # for idx, table in enumerate(tables):
# #     print(f'\n📄 테이블 {idx + 1}')
# #     print(json.dumps(convert_to_json(table), ensure_ascii=False, indent=2))


# import openai
# import os

# # 환경 변수에서 API 키 로드
# openai.api_key = ""

# if openai.api_key is None:
#     print("OPENAI_API_KEY 환경 변수가 설정되지 않았습니다. API 키를 설정해주세요.")
# else:
    
#     prompt = f"""
# You are an AI assistant specialized in generating engaging and SEO-optimized Naver blog posts. Your task is to write a blog review based on the provided input data.

# Here are the requirements for the blog post:

# 1.  **Audience & Tone:** Write a friendly and engaging review, suitable for an "after-use" experience.
# 2.  **Mobile Optimization:** Ensure the content is **extremely optimized for mobile viewing, prioritizing maximum readability on small screens.**
#     * **Line Length:** Every line must be very short. **Even if a sentence is long, break it into multiple lines using a newline character (`\n`) after every few words or a natural phrase break.** Aim for lines that are typically no more than 10-15 characters wide (Korean characters), ensuring a very punchy, easy-to-scan format.
#     * **Paragraph Separation:** **After each main paragraph, you MUST insert two blank lines (i.e., three newline characters: `\n\n\n`) to create a very distinct visual separation.** This ensures that when printed, there are clearly two empty lines between paragraphs.
# 3.  **Introduction:** Start the blog post with a welcoming greeting.
# 4.  **SEO & Structure:**
#     * Write the post with Naver blog SEO best practices in mind.
#     * Divide the main content into 4-5 distinct paragraphs.
#     * **Paragraph Content Guidance:** Each main paragraph should focus on a specific aspect of the experience, similar to these examples:
#         * Paragraph 1: Initial impressions (e.g., interior, facilities, cleanliness, modern feel).
#         * Paragraph 2: Instructor's quality (e.g., friendliness, personalized program design, attention to detail, expectation of results).
#         * Paragraph 3-5: Continue with other positive aspects derived from the input data, such as program effectiveness, atmosphere, or specific benefits.
# 5.  **Word Count:** The total word count must be at least 450 characters (Korean characters).
# 6.  **Keywords & Company Name:** Incorporate the provided keyword "{tables['키워드']}" and company name "{tables['상호 (상품/서비스)']}" naturally throughout the text.
# 7.  **Title Suggestion:** Suggest one title following this format: `[~는 + keyword + company name]`.
# 8.  **Hashtag Recommendation:** Recommend at least 10 relevant hashtags.
# 9.  **Closing:** Conclude the post by including the phrase "감사합니다" (Thank you).



#         **Input Data (Parsed from document):**
#         {str(tables['포스팅 참고 내용'])}

#         **Expected Output Language:** Korean
#     """
#     completion = openai.chat.completions.create(
#         model="gpt-4.1-nano",  # 또는 "gpt-4", "gpt-4o" 등
#         messages=[
#             {"role": "system", "content": "You are a helpful assistant."},
#             {"role": "user", "content": prompt}
#         ],
#         max_tokens=1000,  # 생성할 최대 토큰 수
#         temperature=0.7, # 창의성 조절 (0.0: 보수적, 1.0: 창의적)
#     )

#     print(completion.choices[0].message.content)

import zipfile
import os
from docx import Document
import openai
import json
from dotenv import load_dotenv
import pyzipper

load_dotenv()

ZIP_FOLDER = 'documents'
RESULT_FOLDER = 'result'

# OpenAI API 키 설정
openai.api_key = os.getenv("OPENAI_API_KEY")

# result 폴더 생성
os.makedirs(RESULT_FOLDER, exist_ok=True)

# 1. 문서에서 테이블 데이터 추출 함수
def extract_tables(doc):
    tables = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)
        tables.append(rows)

    parsed_data = {}
    for item in tables[0]:
        if len(item) == 2:
            key = item[0].replace('\n', ' ').strip()
            value = item[1].strip()
            if key == '※ 포스팅 가이드' and value == '※ 포스팅 가이드':
                continue
            parsed_data[key] = value
    return parsed_data

# 2. 각 zip 파일 반복 처리
for zip_filename in os.listdir(ZIP_FOLDER):
    if not zip_filename.endswith('.zip'):
        continue

    zip_path = os.path.join(ZIP_FOLDER, zip_filename)
    zip_name_without_ext = os.path.splitext(zip_filename)[0]
    extract_path = os.path.join(RESULT_FOLDER, zip_name_without_ext)
    os.makedirs(extract_path, exist_ok=True)

    try:
        with pyzipper.AESZipFile(zip_path, 'r') as zf:
            for name in zf.namelist():
                # 폴더면 건너뜀
                if name.endswith('/'):
                    continue

                # 한글 인코딩 복구
                decoded_name = name.encode('cp437').decode('euc-kr')
                target_path = os.path.join(extract_path, decoded_name)

                # ✅ 상위 디렉토리 경로 생성
                os.makedirs(os.path.dirname(target_path), exist_ok=True)

                # ✅ 파일 쓰기
                with open(target_path, 'wb') as f:
                    f.write(zf.read(name))

        # ✅ 압축 해제 완료 후 zip 삭제
        os.remove(zip_path)

        print(f"✅ {zip_filename} → 압축 해제 및 삭제 완료")

    except Exception as e:
        print(f"❌ {zip_filename} 처리 중 에러: {e}")

     # .docx 찾기
    docx_files = [f for f in os.listdir(extract_path) if f.endswith('.docx')]
    if not docx_files:
        print(f"❌ {zip_filename} 안에 .docx 파일이 없습니다.")
        continue
    docx_path = os.path.join(extract_path, docx_files[0])
    doc = Document(docx_path)

    tables = extract_tables(doc)

    # OpenAI 프롬프트 작성
    prompt = f"""
You are an AI assistant specialized in generating engaging, natural-sounding, and SEO-optimized Naver blog posts. Your task is to write a blog review based on the provided input data.

You must analyze the **business name (상호)** and **keyword (키워드)** to understand whether the subject is a **restaurant, fitness center, beauty shop, educational service**, or other type of service. Based on that analysis, tailor your writing tone, structure, and content to match the business type **naturally and persuasively**.

---

Here are the requirements for the blog post:

1.  **Audience & Tone:**  
    Write a friendly, realistic, and engaging review as if you personally used the service or visited the location. Use a natural tone that fits a typical Naver blog review.

2.  **Mobile Optimization:**  
    Ensure the content is **extremely optimized for mobile viewing**, prioritizing maximum readability on small screens.
    
    - **Line Length:** Every line must be very short.  
      Even if a sentence is long, break it into multiple lines using newline characters (`\n`) after every few words or natural phrase breaks.  
      Each line should be **no more than 10–15 Korean characters** wide.
    
    - **Paragraph Separation:**  
      After each main paragraph, insert **two blank lines** (i.e., three newline characters: `\n\n\n`) to create clear visual spacing.

3.  **Introduction:**  
    Begin with a welcoming and personal greeting. Mention what made you interested in this place (if applicable).

4.  **SEO & Structure:**  
    - Use Naver blog SEO best practices.  
    - Divide the content into **4–5 distinct paragraphs**, each highlighting a specific experience or impression.
    - Paragraph suggestions (depending on business type):
        - For restaurants: interior, taste, cleanliness, service, menu variety, location.
        - For fitness: cleanliness, facility quality, trainer experience, personalization, atmosphere.
        - For beauty salons: hygiene, service quality, technician skill, pricing, results.
        - For education: teaching style, clarity, curriculum, staff responsiveness, outcomes.

5.  **Word Count:**  
    The **main body content must be at least 500 Korean characters or more** (not including title or hashtags).  
    Prioritize meaningful, descriptive writing over filler content.

6.  **Keywords & Company Name Usage:**  
    Naturally incorporate the following:
    - Keyword: "{tables.get('키워드', '')}"  
    - Company Name: "{tables.get('상호 (상품/서비스)', '')}"  
    Use them **multiple times throughout the post**, especially in the introduction and key impressions.

7.  **Title Suggestion:**  
    Suggest one blog post title in the format:  
    `[~는 + keyword + company name]`  
    The title must be attractive and clickable.

8.  **Hashtag Recommendation:**  
    Recommend at least 10 relevant, popular hashtags that relate to the business and blog content.

9.  **Closing:**  
    Conclude with a warm, personal phrase and **must** include “감사합니다”.

---

**Input Data (Parsed from document):**
{tables.get('포스팅 참고 내용', '')}

**Output Language:** Korean

"""

    try:
        completion = openai.chat.completions.create(
            model="gpt-4.1-nano",  # 또는 "gpt-4", "gpt-4.1-nano"
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1500,
            temperature=0.7
        )

        result_text = completion.choices[0].message.content

        # 결과 텍스트 파일로 저장
        output_path = os.path.join(extract_path, 'result.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(result_text)

        print(f"✅ 완료: {zip_filename} → {output_path}")

    except Exception as e:
        print(f"❌ 오류 발생: {zip_filename}")
        print(e)
