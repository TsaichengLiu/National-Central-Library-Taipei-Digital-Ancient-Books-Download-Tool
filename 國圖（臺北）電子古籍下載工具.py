"""
國圖（臺北）電子古籍下載工具

該工具包括以下主要功能：
1. 鍵入臺北國圖的書目地址（如https://rbook.ncl.edu.tw/NCLSearch/Search/SearchDetail?item=60b8b835ce744667b2b8d2b6aa2022f2fDUyMDg20.TVP7cHyLivhxmwEBqM2mscMvqL6VGwmVhRFgj85TA8k&page=&whereString=&sourceWhereString=&SourceID=0&HasImage=）。
2. 程序將打開網頁，你需要完成臺北國圖給出的圖形驗證碼。
3. 完成之後程序會自動跳轉，你得選擇保存圖像的位置。
4. 你須設置翻頁間隔和圖像質量。設定翻頁間隔是因爲頁面加載新圖像時會出現黑屏，程序此時抓取會得到空圖。所以你須根據網速等因素選擇合適的翻頁間隔。
5. 程序會自動抓取圖像並保存到指定位置，嗣後自動翻頁以抓取下一頁的圖像。
6. 你得隨時暫停或繼續抓取過程。

使用說明：
1. 確保已安裝 PyQt5 和 PyQtWebEngine 庫。
2. 運行此腳本。
3. 在“輸入網址”字段中輸入目標網頁的 URL。
4. 點擊“打開網頁”按鈕，程序將自動打開網頁並點擊進入古籍影像瀏覽按鈕。
5. 網頁打開後，手動完成驗證碼驗證。
6. 驗證碼完成後，選擇保存圖像的位置。
7. 設置頁面加載間隔（秒）和圖片質量（0.1-1.0）。
8. 點擊“開始抓取”按鈕開始抓取圖像。
9. 可使用“暫停”和“繼續”按鈕來控制抓取過程。

Tại-Sinh
2024

"""


import sys
import os
import base64
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QFileDialog, QVBoxLayout, QWidget, QSpinBox, QLabel, QDoubleSpinBox, QMessageBox
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QTimer, pyqtSlot

class WebScraperApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.browser = None
        self.save_path = None
        self.is_paused = False
        self.image_count = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.scrape_images)

    def initUI(self):
        self.setWindowTitle('國圖（臺北）電子古籍下載工具')

        layout = QVBoxLayout()

        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText('輸入網址')
        layout.addWidget(self.url_input)

        self.open_button = QPushButton('打開網頁', self)
        self.open_button.clicked.connect(self.open_webpage)
        layout.addWidget(self.open_button)

        self.save_path_button = QPushButton('選擇保存位置', self)
        self.save_path_button.clicked.connect(self.choose_save_path)
        layout.addWidget(self.save_path_button)

        self.interval_label = QLabel('設置頁面加載間隔（秒）:', self)
        layout.addWidget(self.interval_label)

        self.interval_input = QSpinBox(self)
        self.interval_input.setRange(1, 60)
        self.interval_input.setValue(5)
        layout.addWidget(self.interval_input)

        self.quality_label = QLabel('設置圖片質量（0.1-1.0）:', self)
        layout.addWidget(self.quality_label)

        self.quality_input = QDoubleSpinBox(self)
        self.quality_input.setRange(0.1, 1.0)
        self.quality_input.setSingleStep(0.1)
        self.quality_input.setValue(1.0)
        layout.addWidget(self.quality_input)

        self.start_button = QPushButton('開始抓取', self)
        self.start_button.clicked.connect(self.start_scraping)
        self.start_button.setEnabled(False)
        layout.addWidget(self.start_button)

        self.pause_button = QPushButton('暫停', self)
        self.pause_button.clicked.connect(self.pause_scraping)
        self.pause_button.setEnabled(False)
        layout.addWidget(self.pause_button)

        self.resume_button = QPushButton('繼續', self)
        self.resume_button.clicked.connect(self.resume_scraping)
        self.resume_button.setEnabled(False)
        layout.addWidget(self.resume_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def open_webpage(self):
        url = self.url_input.text()
        self.browser = QWebEngineView()
        self.browser.load(QUrl(url))
        self.browser.show()

        # 等待页面加载完成后自动点击“进入古籍影像浏览”按钮
        self.browser.page().loadFinished.connect(self.click_view_button)

    def click_view_button(self):
        # 尝试点击“进入古籍影像浏览”按钮
        script = """
        var button = document.querySelector('.btn.btn-default.bg-malachite-green.text-white.mr-5px');
        if (button) {
            button.click();
        }
        """
        self.browser.page().runJavaScript(script)

        # 等待页面跳转到包含古籍图片的页面
        QTimer.singleShot(30000, self.enable_save_path_button)

    def enable_save_path_button(self):
        self.save_path_button.setEnabled(True)

    def choose_save_path(self):
        self.save_path = QFileDialog.getExistingDirectory(self, '選擇保存目錄')
        if self.save_path:
            self.start_button.setEnabled(True)
            self.pause_button.setEnabled(True)
            self.resume_button.setEnabled(True)

    def start_scraping(self):
        self.is_paused = False
        interval = self.interval_input.value() * 1000
        self.timer.start(interval)

    def pause_scraping(self):
        self.is_paused = True
        self.timer.stop()

    def resume_scraping(self):
        self.is_paused = False
        interval = self.interval_input.value() * 1000
        self.timer.start(interval)

    def scrape_images(self):
        if self.is_paused:
            return

        quality = self.quality_input.value()

        # 执行JavaScript代码获取图片的base64编码数据和页面编号
        script = f"""
        (function() {{
            var img = document.querySelector('#ImageDisplay');
            var pageNum = document.querySelector('#sel-content-no').selectedOptions[0].textContent;
            var totalPages = document.querySelector('span.total-page').textContent.trim().replace('/ ', '');
            if (img) {{
                console.log('Found image element:', img);
                var canvas = document.createElement('canvas');
                canvas.width = img.naturalWidth;
                canvas.height = img.naturalHeight;
                var ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, img.naturalWidth, img.naturalHeight);
                var dataUrl = canvas.toDataURL('image/png').substring(22);  // 使用PNG格式
                console.log('Data URL:', dataUrl);
                return {{dataUrl: dataUrl, pageNum: pageNum, totalPages: totalPages}};
            }} else {{
                console.log('Image element not found');
                return null;
            }}
        }})();
        """
        self.browser.page().runJavaScript(script, self.save_image)

    @pyqtSlot(dict)
    def save_image(self, result):
        if result and result.get("dataUrl"):
            img_base64 = result["dataUrl"]
            page_num = result["pageNum"].strip()
            total_pages = result["totalPages"].strip()
            img_data = base64.b64decode(img_base64)
            img_name = os.path.join(self.save_path, f'第{page_num}頁.png')
            with open(img_name, 'wb') as f:
                f.write(img_data)
            self.image_count += 1

            if page_num == total_pages:
                self.timer.stop()
                QMessageBox.information(self, "完成", "所有頁面已成功抓取。")
            else:
                # 点击“下一页”按钮
                next_script = """
                var nextButton = document.getElementById('AftT');
                if (nextButton) {
                    nextButton.click();
                }
                """
                QTimer.singleShot(2000, lambda: self.browser.page().runJavaScript(next_script))  # 延迟点击下一页按钮
        else:
            print("Failed to get image data")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WebScraperApp()
    ex.show()
    sys.exit(app.exec_())
