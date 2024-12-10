import subprocess
import os


def convert_ui_to_python(ui_file: str, output_file: str):
    # pyuic5 명령을 실행할 경로 지정
    if not os.path.exists(ui_file):
        raise FileNotFoundError(f"UI 파일이 존재하지 않습니다: {ui_file}")

    # pyuic5 명령을 실행
    try:
        subprocess.run(['pyuic5', ui_file, '-o', output_file], check=True)
        print(f"{ui_file} 파일이 성공적으로 {output_file}로 변환되었습니다.")
    except subprocess.CalledProcessError as e:
        print(f"UI 파일 변환 중 오류가 발생했습니다: {e}")
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
    except Exception as e:
        print(f"예상치 못한 오류가 발생했습니다: {e}")


# 사용 예시
ui_file = 'ExceltoCSV.ui'  # 변환할 UI 파일
output_file = 'mainUI.py'  # 출력할 Python 파일 이름

convert_ui_to_python(ui_file, output_file)
