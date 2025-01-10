import ctypes
import os
import sys
import time

import pyperclip
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit, QPushButton, QListWidget, QHBoxLayout, \
    QMessageBox

# 핀번호 파일 경로
PINS_FILE = 'pins.txt'

# Windows API 함수 로드
user32 = ctypes.windll.user32

# 가상 키 코드
VK_CONTROL = 0x11  # Ctrl 키
VK_SHIFT = 0x10  # Shift 키
VK_V = 0x56  # 'V' 키 (붙여넣기)
VK_J = 0x4A  # 'J' 키
ENTER_KEY = 0x0D  # Enter 키


def press_key(key):
    user32.keybd_event(key, 0, 0, 0)  # 키 누름


def release_key(key):
    user32.keybd_event(key, 0, 2, 0)  # 키 떼기


class PinManager(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('에그머니 자동입력기')  # 제목명 수정
        self.setMinimumSize(330, 400)  # 최소 크기 설정

        # 레이아웃 설정
        self.layout = QVBoxLayout()

        # 핀번호 입력 텍스트박스
        self.pin_input = QTextEdit(self)
        self.pin_input.setPlaceholderText(
            '핀번호를 추가하려면 이곳에 입력해주세요.\n숫자만 일치한다면 어떠한 형태로 입력해도 됩니다.\n예시:\n12345-12345-12345-12345\n12345-12345-12345-67890')
        self.pin_input.setAcceptRichText(False)  # 서식이 포함된 텍스트를 거부
        self.pin_input.setAutoFormatting(QTextEdit.AutoAll)  # 자동 서식 적용 비활성화
        self.layout.addWidget(self.pin_input)

        # 핀번호 목록 표시 (단일 선택과 다중 선택이 모두 가능하도록 설정)
        self.pin_list = QListWidget(self)
        self.pin_list.setSelectionMode(QListWidget.ExtendedSelection)  # 클릭은 단일 선택, 드래그는 다중 선택
        self.layout.addWidget(self.pin_list)

        # 버튼 레이아웃
        button_layout = QHBoxLayout()

        # 핀번호 추가 버튼
        self.add_button = QPushButton('핀번호 추가', self)
        self.add_button.clicked.connect(self.add_pin)
        button_layout.addWidget(self.add_button)

        # 핀번호 삭제 버튼
        self.delete_button = QPushButton('핀번호 삭제', self)
        self.delete_button.clicked.connect(self.delete_pin)
        button_layout.addWidget(self.delete_button)

        # 핀번호 자동 입력 버튼
        self.auto_input_button = QPushButton('핀번호 자동 입력', self)
        self.auto_input_button.clicked.connect(self.auto_input_pin)
        button_layout.addWidget(self.auto_input_button)

        self.layout.addLayout(button_layout)

        self.setLayout(self.layout)

        # 핀번호 목록을 파일에서 로드
        self.load_pins()

    def load_pins(self):
        """pins.txt 파일에서 핀번호를 읽어서 목록에 표시"""
        if os.path.exists(PINS_FILE):
            with open(PINS_FILE, 'r') as file:
                pins = file.readlines()
            for pin in pins:
                formatted_pin = self.format_pin(pin.strip())
                self.pin_list.addItem(formatted_pin)

    def add_pin(self):
        """입력된 핀번호를 목록에 추가하고 파일에 저장"""
        pin_input = self.pin_input.toPlainText().strip()

        # 입력값에서 숫자만 추출
        pin_input = ''.join(filter(str.isdigit, pin_input))

        if len(pin_input) % 20 != 0:
            # 길이가 20의 배수가 아닌 경우
            self.show_error_message('길이가 잘못된 핀번호가 존재합니다.')
        else:
            # 20자씩 잘라서 핀번호로 추가
            for i in range(0, len(pin_input), 20):
                pin = pin_input[i:i + 20]
                formatted_pin = self.format_pin(pin)
                self.pin_list.addItem(formatted_pin)

            self.save_pins()
            self.pin_input.clear()

    def delete_pin(self):
        """선택된 핀번호를 삭제하고 파일에 저장"""
        selected_items = self.pin_list.selectedItems()
        if selected_items:
            for item in selected_items:
                self.pin_list.takeItem(self.pin_list.row(item))
            self.save_pins()

    def save_pins(self):
        """핀번호 목록을 pins.txt 파일에 저장"""
        with open(PINS_FILE, 'w') as file:
            for index in range(self.pin_list.count()):
                pin = self.pin_list.item(index).text().replace('-', '')  # '-' 제거
                file.write(pin + '\n')

    def format_pin(self, pin):
        """핀번호를 5자마다 '-'를 추가하여 보기 좋게 포맷"""
        return '-'.join([pin[i:i + 5] for i in range(0, len(pin), 5)])

    def show_error_message(self, message):
        """에러 메시지를 표시"""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.setWindowTitle("에러")
        msg.exec_()

    def auto_input_pin(self):
        """HAOPLAY 창 활성화 및 자바스크립트 입력"""
        try:
            if self.pin_list.count() < 1:
                return

            # 1️⃣ HAOPLAY 창 핸들 찾기
            haoplay_hwnd = user32.FindWindowW(None, "HAOPLAY")
            if not haoplay_hwnd:
                print("❌ HAOPLAY 창을 찾을 수 없습니다.")
                return

            # 2️⃣ "Chrome_WidgetWin_0" 컨트롤 핸들 찾기 (웹뷰 컨트롤)
            webview_hwnd = user32.FindWindowExW(haoplay_hwnd, 0, "Chrome_WidgetWin_0", None)
            if not webview_hwnd:
                print("❌ 웹뷰 컨트롤을 찾을 수 없습니다.")
                return

            print(f"✅ 웹뷰 컨트롤 핸들 찾음: {webview_hwnd}")

            # 3️⃣ 창 활성화 (child_hWnd로 변경)
            user32.SetForegroundWindow(webview_hwnd)  # 웹뷰 컨트롤을 최상위로 활성화
            time.sleep(0.5)  # 안정성을 위해 대기

            open_debug_window()
            time.sleep(1)

            # 목록에서 최대 5개의 핀번호를 가져옴
            pins_to_inject = [self.pin_list.item(i).text() for i in range(min(5, self.pin_list.count()))]

            # 4️⃣ 핀번호를 입력박스에 추가
            add_pin_input_box(len(pins_to_inject))
            time.sleep(0.2)

            # 5️⃣ 핀번호들 자바스크립트를 통해 입력
            inject_pin_codes(pins_to_inject)

            # 모두 동의
            click_all_agree()

            # 제출
            submit()

        except Exception as e:
            self.show_error_message(f"자동 입력에 실패했습니다: {e}")


def open_debug_window():
    # Ctrl + Shift 누름
    press_key(VK_CONTROL)
    press_key(VK_SHIFT)
    time.sleep(0.1)

    # 'J' 누름
    press_key(VK_J)
    time.sleep(0.1)
    release_key(VK_J)

    time.sleep(0.1)

    # Ctrl + Shift 떼기
    release_key(VK_SHIFT)
    release_key(VK_CONTROL)

    print("✅ Ctrl + Shift + J 입력 완료.")


def add_pin_input_box(num):
    if num < 1:
        return

    # 5️⃣ 자바스크립트 코드 준비
    javascript_code = f'''
        while(document.querySelector("input[name='pyo_cnt']").value < {num})
            PinBoxInsert('pyo_cnt');
    '''
    paste_javascript_code(javascript_code)


def inject_pin_codes(pins: list[str]):
    if pins is None or len(pins) == 0:
        return

    arr = [item for pin in pins for item in pin.split("-")]
    arr_text = f"['{"', '".join(arr)}']"
    # 5️⃣ 자바스크립트 코드 준비
    javascript_code = '''
        let i = 0;
        arr = arr_text;
        document.querySelectorAll("#pinno").forEach(obj => {
            obj.querySelectorAll("input").forEach(input => {
                input.value = arr[i++];
            })
        })
    '''.replace("arr_text", arr_text)
    paste_javascript_code(javascript_code)


def click_all_agree():
    javascript_code = 'document.querySelector("#all-agree").click()'
    paste_javascript_code(javascript_code)


def submit():
    javascript_code = 'goSubmit(document.form)'
    paste_javascript_code(javascript_code)


def paste_javascript_code(javascript_code):
    # 6️⃣ 클립보드에 자바스크립트 코드 복사 (pyperclip 사용)
    pyperclip.copy(javascript_code)
    print("✅ 자바스크립트 코드 클립보드에 복사 완료.")

    # 7️⃣ Ctrl + V 키 입력 (붙여넣기)
    press_key(VK_CONTROL)
    time.sleep(0.1)
    press_key(VK_V)
    time.sleep(0.1)
    release_key(VK_V)
    time.sleep(0.1)
    release_key(VK_CONTROL)
    time.sleep(0.1)

    # 8️⃣ Enter 키 입력 (자바스크립트 실행)
    press_key(ENTER_KEY)
    time.sleep(0.1)
    release_key(ENTER_KEY)

    print("✅ 자바스크립트 실행 완료.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PinManager()
    window.show()
    sys.exit(app.exec_())
