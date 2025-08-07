"""
DB 산출물 생성 GUI 애플리케이션
메인 진입점
"""

from gui.main_window import DBSpecGeneratorApp

if __name__ == "__main__":
    app = DBSpecGeneratorApp()
    app.run()