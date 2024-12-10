import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from mainUI import Ui_MainWindow  # 변환된 UI 파일을 임포트

# 변환
# pyuic5 -o mainUI.py ExceltoCSV.ui

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()  # UI 객체 생성
        self.ui.setupUi(self)  # UI 설정

        ####################################################
        ## 버튼 클릭 이벤트
        self.ui.btn_csv_search.clicked.connect(self.select_csv_file)        # csv 파일 열기
        self.ui.btn_savefile_search.clicked.connect(self.select_save_path)  # 저장 경로 선택
        self.ui.btn_convert.clicked.connect(self.convert_to_csv)            # 변환

        self.last_folder = os.path.expanduser('~\\Downloads')  # 처음 기본 경로는 다운로드 폴더

    def select_csv_file(self):
        # 파일 선택 다이얼로그 열기 (엑셀 파일 필터)
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx; *.xls)")

        if file_path:
            self.ui.lineEdit_csvpath.setText(file_path)  # 선택한 경로를 lineEdit에 설정

            # 선택한 폴더를 기억해서 다음에 사용할 수 있도록 설정
            self.last_folder = os.path.dirname(file_path)  # 선택한 파일의 경로에서 디렉토리만 추출

    # 파일 저장 다이얼로그 열기
    def select_save_path(self):
        # 파일 저장 경로를 설정할 수 있는 다이얼로그 열기
        save_path = QFileDialog.getExistingDirectory(self, "폴더 선택", "")  # 파일이 아닌 폴더만 선택

        if save_path:
            self.ui.lineEdit_savepath.setText(save_path)  # 선택한 경로를 lineEdit에 설정
    def convert_to_csv(self):
        # lineEdit_csvpath에서 엑셀 파일 경로 가져오기
        excel_path = self.ui.lineEdit_csvpath.text()

        # lineEdit_savepath에서 저장할 경로 가져오기
        save_path = self.ui.lineEdit_savepath.text() + ".csv"

        if not excel_path or not save_path:
            QMessageBox.warning(self, "경고", "엑셀 파일 경로와 저장 경로를 모두 입력해주세요.")
            return

        # 경로가 존재하는지 확인
        save_dir = os.path.dirname(save_path)
        if not os.path.exists(save_dir):
            QMessageBox.warning(self, "경고", "저장 경로가 존재하지 않습니다.")
            return

        # print(save_path)
        # print(os.access(excel_path, os.F_OK))
        # print(os.access(excel_path, os.R_OK))
        # print(os.access(excel_path, os.W_OK))
        # print(os.access(excel_path, os.X_OK))

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(excel_path)

            # CSV 파일로 저장 (덮어쓰기가 허용되도록 함)
            df.to_csv(save_path, index=False, encoding='utf-8')

            # 성공 메시지
            QMessageBox.information(self, "성공", f"파일이 성공적으로 저장되었습니다: {save_path}")
        
        # 예외처리
        except FileNotFoundError:
            QMessageBox.critical(self, "파일 오류", f"지정한 엑셀 파일을 찾을 수 없습니다: {excel_path}")
        except PermissionError:
            QMessageBox.critical(self, "권한 오류", f"파일을 저장할 수 없습니다. 권한을 확인하세요: {save_path}")
        except ValueError:
            QMessageBox.critical(self, "값 오류", "엑셀 파일에서 잘못된 값을 읽어올 수 없습니다. 파일을 확인해주세요.")
        except pd.errors.EmptyDataError:
            QMessageBox.critical(self, "빈 데이터 오류", "엑셀 파일에 데이터가 없습니다. 파일을 확인해주세요.")
        except Exception as e:
            # 다른 일반적인 오류 처리
            QMessageBox.critical(self, "오류", f"파일 변환 중 오류가 발생했습니다: {e}")


if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()
