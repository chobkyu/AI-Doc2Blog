import os
import zipfile
from docx import Document

ZIP_FOLDER = 'documents'   # ZIP 파일들이 있는 폴더
RESULT_FOLDER = 'result'   # 압축 해제 결과를 저장할 폴더

# 결과 폴더 없으면 생성
os.makedirs(RESULT_FOLDER, exist_ok=True)

# 1. documents 폴더 내 모든 zip 파일 순회
for zip_filename in os.listdir(ZIP_FOLDER):
    if not zip_filename.endswith('.zip'):
        continue  # zip 파일만 처리

    zip_path = os.path.join(ZIP_FOLDER, zip_filename)

    # 2. 압축 해제 폴더 이름 → result/zip이름(확장자 제외)
    extract_dir_name = os.path.splitext(zip_filename)[0]
    extract_dir_path = os.path.join(RESULT_FOLDER, extract_dir_name)

    # 3. zip 압축 풀기
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir_path)

    # 4. 압축 해제된 폴더에서 .docx 파일 찾기
    docx_files = [f for f in os.listdir(extract_dir_path) if f.endswith('.docx')]
    if not docx_files:
        print(f"❌ {zip_filename} 안에 .docx 파일이 없습니다.")
        continue

    docx_path = os.path.join(extract_dir_path, docx_files[0])

    # 5. 문서 열기 및 내용 출력 (또는 원하는 작업 수행)
    doc = Document(docx_path)
    print(f"✅ {zip_filename} → {docx_files[0]} 내용:")

    for para in doc.paragraphs:
        print(para.text)

    print("-" * 40)
