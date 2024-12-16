import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from mainUI import Ui_MainWindow  # 변환된 UI 파일을 임포트
from datetime import datetime

# 변환
# pyuic5 -o mainUI.py ExceltoCSV.ui

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.save_path = None
        self.ui = Ui_MainWindow()  # UI 객체 생성
        self.ui.setupUi(self)  # UI 설정

        # 선택한 파일 경로들을 저장할 리스트와 저장 경로를 저장할 변수
        self.selected_files = []
        self.save_folder = ""

        ####################################################
        ## 버튼 클릭 이벤트
        self.ui.btn_csv_search.clicked.connect(self.select_csv_file)        # csv 파일 열기
        self.ui.btn_savefile_search.clicked.connect(self.select_save_path)  # 저장 경로 선택
        self.ui.btn_convert.clicked.connect(self.convert_to_csv)            # 변환

        self.last_folder = os.path.expanduser('~\\Downloads')  # 처음 기본 경로는 다운로드 폴더

    def select_csv_file(self):
        # 파일 선택 다이얼로그 열기 (엑셀 파일 필터)
        #file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx; *.xls)")

        # 파일 선택 다이얼로그 열기 (엑셀 파일 필터) - 여러 개의 파일 선택
        file_paths, _ = QFileDialog.getOpenFileNames(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx; *.xls)")

        if file_paths:
            self.selected_files = file_paths  # 선택된 파일 경로들을 리스트에 저장
            file_names = "\n".join([os.path.basename(f) for f in file_paths])
            self.ui.lineEdit_csvpath.setText(file_names)  # 파일 이름들을 lineEdit에 표시

    # 파일 저장 다이얼로그 열기
    def select_save_path(self):
        # 파일 저장 경로를 설정할 수 있는 다이얼로그 열기
        self.save_path = QFileDialog.getExistingDirectory(self, "폴더 선택", "")  # 파일이 아닌 폴더만 선택

        if self.save_path:
            self.ui.lineEdit_savepath.setText(self.save_path)  # 선택한 경로를 lineEdit에 설정

    def convert_to_csv(self):


        # 저장 폴더가 존재하는지 확인
        if not os.path.exists(self.save_path):
            QMessageBox.warning(self, "경고", "저장 경로가 존재하지 않습니다.")
            print(self.save_path)
            return


        cnt = 0

        for excel_path in self.selected_files:
            try:
                # 엑셀 파일 읽기
                df = pd.read_excel(excel_path)

                # 첫 번째 열을 날짜로 변환 후 원하는 형식으로 변환
                df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0]).dt.strftime('%Y%m%d%H%M')

                # 파일 이름에서 확장자 제외하고 오늘 날짜 추가
                file_name = os.path.basename(excel_path)
                base_name, _ = os.path.splitext(file_name)

                # 오늘 날짜를 "YYYYMMDD" 형식으로 가져오기
                today = datetime.today().strftime('%Y%m%d')

                # 저장할 파일 경로 생성
                save_file_path = self.save_path + os.path.join(self.save_folder, f"//{today}_{base_name}.csv")
                print(f"저장 경로: {save_file_path}")
                # 경로 구분자를 자동으로 맞춰주기 (Windows의 \를 /로 변경)
                save_file_path = save_file_path.replace("\\", "/")

                # CSV 파일로 저장 (헤더를 제외하고 저장)
                # df.to_csv(save_file_path, index=False, header=False, encoding='utf-8')
                df.to_csv(save_file_path, index=False, header=False, encoding='utf-8-sig')

                cnt += 1


            except FileNotFoundError:
                QMessageBox.critical(self, "파일 오류", f"지정한 엑셀 파일을 찾을 수 없습니다: {excel_path}")
            except PermissionError:
                QMessageBox.critical(self, "권한 오류", f"파일을 저장할 수 없습니다. 권한을 확인하세요: {save_file_path}")
            except ValueError:
                QMessageBox.critical(self, "값 오류", "엑셀 파일에서 잘못된 값을 읽어올 수 없습니다. 파일을 확인해주세요.")
            except pd.errors.EmptyDataError:
                QMessageBox.critical(self, "빈 데이터 오류", "엑셀 파일에 데이터가 없습니다. 파일을 확인해주세요.")
            except Exception as e:
                # 다른 일반적인 오류 처리
                QMessageBox.critical(self, "오류", f"파일 변환 중 오류가 발생했습니다: {e}")

        # 모든 파일 변환 후 성공 메시지
        if cnt > 0:
            QMessageBox.information(self, "성공", f"{cnt} 개 파일이 성공적으로 저장되었습니다")
        else:
            QMessageBox.warning(self, "경고", "저장된 파일이 없습니다. 변환이 실패한 파일이 있을 수 있습니다.")

if __name__ == '__main__':
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()
