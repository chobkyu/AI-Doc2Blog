import os
import pyzipper
from docx import Document

ZIP_FOLDER = 'documents'
RESULT_FOLDER = 'result'

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
    print(tables)