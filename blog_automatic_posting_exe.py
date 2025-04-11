from PyQt5.QtWidgets import *
from PyQt5 import uic
import sys
import os
import blog_automatic_posting

if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS  # PyInstaller 임시 폴더에서 리소스를 찾기
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 개발 중에는 현재 디렉토리에서 찾기
    
UI_PATH = "bkig_automatic_posting.ui"

class MainDialog(QDialog):
    def __init__(self):
        super().__init__(None)
        uic.loadUi(os.path.join(BASE_DIR, UI_PATH), self)
        self.start.clicked.connect(self.run_main)

    def run_main(self):
        try:
            coupang_id = self.coupg_id.text()
            coupang_pw = self.coupang_pw.text()
            blog_id = self.blog_id.text()
            blog_pw = self.blog_pw.text()
            blog_write_page = self.write_page_url.text()
            search_word = self.search_word.text()
            post_num = self.post_num.text()
            api_key = self.api_key.text()

            if not post_num.isdigit() or int(post_num) == 0:
                raise ValueError("글 개수에는 0이 아닌 숫자만 들어갈 수 있습니다.")

            # 실제 실행
            blog_automatic_posting.main(
                coupang_id, coupang_pw, blog_id, blog_pw,
                blog_write_page, search_word, post_num, api_key
            )

        except Exception as e:
            QMessageBox.critical(self, "에러", str(e))  # GUI 에러 메시지 출력


if __name__ == "__main__" :
    QApplication.setStyle("fusion")
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    main_dialog = MainDialog() 

    #프로그램 화면을 보여주는 코드
    main_dialog.show()
    sys.exit(app.exec_())