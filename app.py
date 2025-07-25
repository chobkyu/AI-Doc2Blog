

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
        # os.remove(zip_path)

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
    with open("prompt_template.txt", "r", encoding="utf-8") as f:
        prompt_template = f.read()

    # ✅ 실제 데이터 삽입
    prompt = prompt_template.format(
        keyword=tables.get("키워드", ""),
        company_name=tables.get("상호 (상품/서비스)", ""),
        reference=tables.get("포스팅 참고 내용", "")
    )

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
