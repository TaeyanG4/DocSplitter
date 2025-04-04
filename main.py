import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import copy

from PyQt5 import QtWidgets, QtCore

def extract_hwpx_xml(file_path):
    """
    HWPX/hwtx 파일은 ZIP 압축 파일 형식입니다.
    여기서는 주 내용이 저장된 XML 파일(예: 'Contents/section0.xml')을 파싱하여 XML 트리를 반환합니다.
    """
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            xml_path = 'Contents/section0.xml'
            if xml_path in z.namelist():
                with z.open(xml_path) as f:
                    tree = ET.parse(f)
                return tree, xml_path
            else:
                raise FileNotFoundError(f"'{xml_path}' 파일을 찾을 수 없습니다. 파일 구조를 확인해주세요.")
    except Exception as e:
        raise Exception(f"파일 열기/파싱 중 오류 발생: {e}")

def split_by_template(tree, template, skip_pages):
    """
    XML 트리의 자식 요소들을 페이지로 간주하고,
    첫 skip_pages 만큼은 무시한 후, 각 페이지의 텍스트에 template 문자열이 나타나면
    새로운 구간의 시작으로 판단하여 분리합니다.
    """
    root = tree.getroot()
    pages = list(root)
    pages = pages[skip_pages:]  # 첫 skip_pages 페이지 무시

    sections = []
    current_section = []

    for page in pages:
        page_text = ''.join(page.itertext())
        if template in page_text:
            if current_section:
                sections.append(current_section)
            # template이 포함된 페이지도 새 구간의 시작에 포함
            current_section = [page]
        else:
            if current_section:
                current_section.append(page)
    if current_section:
        sections.append(current_section)
    return sections

def create_section_tree(original_tree, section_pages):
    """
    원본 트리의 루트 속성을 유지하며, 분리된 해당 구간(section_pages)만 deep copy하여 새로운 XML 트리를 생성합니다.
    """
    root = original_tree.getroot()
    new_root = ET.Element(root.tag, root.attrib)
    for page in section_pages:
        new_root.append(copy.deepcopy(page))
    return ET.ElementTree(new_root)

def create_section_hwpx(original_file, section_tree, xml_path, output_file):
    """
    원본 HWPX 파일의 전체 구조(이미지, 표, 폰트 등)를 그대로 복사하고,
    지정된 xml_path(주 내용 XML 파일)를 새 구간 XML 내용으로 대체하여 새로운 HWPX 파일을 생성합니다.
    """
    with zipfile.ZipFile(original_file, 'r') as zin:
        with zipfile.ZipFile(output_file, 'w') as zout:
            for item in zin.infolist():
                if item.filename == xml_path:
                    xml_bytes = ET.tostring(section_tree.getroot(), encoding="utf-8", method="xml")
                    xml_content = b'<?xml version="1.0" encoding="utf-8"?>\n' + xml_bytes
                    zout.writestr(item, xml_content)
                else:
                    file_data = zin.read(item.filename)
                    zout.writestr(item, file_data)

def process_file(input_file, output_dir, template, skip_pages, log_callback):
    """
    설정값에 따라 HWPX 파일을 분리 처리합니다.
    진행 사항은 log_callback 함수를 통해 전달합니다.
    """
    try:
        log_callback("XML 파일 추출 중...")
        tree, xml_path = extract_hwpx_xml(input_file)
    except Exception as e:
        log_callback(str(e))
        return

    try:
        log_callback("분리 기준에 따라 페이지 분리 중...")
        sections = split_by_template(tree, template, skip_pages)
    except Exception as e:
        log_callback(f"페이지 분리 중 오류 발생: {e}")
        return

    if not sections:
        log_callback(f"분리 기준 '{template}'을(를) 포함하는 페이지를 찾지 못했습니다. 설정을 확인해주세요.")
        return

    os.makedirs(output_dir, exist_ok=True)
    log_callback(f"총 {len(sections)}개의 구간이 발견되었습니다.")

    for i, section_pages in enumerate(sections, start=1):
        log_callback(f"구간 {i}번 처리 중...")
        section_tree = create_section_tree(tree, section_pages)
        output_file = os.path.join(output_dir, f'section_{i}.hwpx')
        try:
            create_section_hwpx(input_file, section_tree, xml_path, output_file)
            log_callback(f"구간 {i}번이 '{output_file}' 파일로 저장되었습니다.")
        except Exception as e:
            log_callback(f"구간 {i}번 저장 중 오류 발생: {e}")

def merge_hwpx_files(input_folder, output_folder, log_callback):
    """
    선택한 폴더 내의 모든 HWPX/hwtx 파일을 하나의 파일로 병합합니다.
    각 파일에서 'Contents/section0.xml'을 추출하여 모든 페이지를 순서대로 합친 후,
    기본 파일의 구조를 복사하여 새로운 HWPX 파일을 생성합니다.
    """
    # 입력 폴더에서 HWPX/hwtx 파일 찾기
    files = []
    for file in os.listdir(input_folder):
        if file.lower().endswith(('.hwpx', '.hwtx')):
            files.append(os.path.join(input_folder, file))
    if not files:
        log_callback("선택된 폴더에서 HWPX/hwtx 파일을 찾지 못했습니다.")
        return
    files.sort()
    log_callback(f"총 {len(files)}개의 파일을 병합합니다.")

    try:
        base_tree, xml_path = extract_hwpx_xml(files[0])
    except Exception as e:
        log_callback(f"기본 파일의 XML 추출에 실패했습니다: {e}")
        return

    merged_pages = []
    for file in files:
        log_callback(f"{file} 파일에서 페이지 추출 중...")
        try:
            tree, _ = extract_hwpx_xml(file)
            root = tree.getroot()
            pages = list(root)
            merged_pages.extend(copy.deepcopy(pages))
        except Exception as e:
            log_callback(f"{file} 파일에서 XML 추출 실패: {e}")

    # 기본 파일의 루트 속성을 이용하여 새로운 XML 트리 생성
    new_root = ET.Element(base_tree.getroot().tag, base_tree.getroot().attrib)
    for page in merged_pages:
        new_root.append(page)
    merged_tree = ET.ElementTree(new_root)

    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, "merged.hwpx")
    try:
        create_section_hwpx(files[0], merged_tree, xml_path, output_file)
        log_callback(f"병합된 파일이 '{output_file}' 에 저장되었습니다.")
    except Exception as e:
        log_callback(f"병합 파일 생성 중 오류 발생: {e}")

# =======================
# GUI - 분리하기 탭
# =======================
class SplitTab(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # 입력 파일 선택 영역
        file_layout = QtWidgets.QHBoxLayout()
        self.input_line = QtWidgets.QLineEdit()
        self.input_line.setPlaceholderText("분리할 HWPX/hwtx 파일 경로")
        browse_btn = QtWidgets.QPushButton("파일 선택")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(QtWidgets.QLabel("입력 파일:"))
        file_layout.addWidget(self.input_line)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)

        # 출력 폴더 선택 영역
        out_layout = QtWidgets.QHBoxLayout()
        self.output_line = QtWidgets.QLineEdit("output_sections")
        out_btn = QtWidgets.QPushButton("폴더 선택")
        out_btn.clicked.connect(self.select_output_dir)
        out_layout.addWidget(QtWidgets.QLabel("출력 폴더:"))
        out_layout.addWidget(self.output_line)
        out_layout.addWidget(out_btn)
        layout.addLayout(out_layout)

        # 분리 기준(template) 및 무시할 페이지 수 설정 영역
        setting_layout = QtWidgets.QHBoxLayout()
        self.template_line = QtWidgets.QLineEdit("[별지 제7호서식]")
        self.skip_spin = QtWidgets.QSpinBox()
        self.skip_spin.setRange(0, 100)
        self.skip_spin.setValue(3)
        setting_layout.addWidget(QtWidgets.QLabel("분리 기준:"))
        setting_layout.addWidget(self.template_line)
        setting_layout.addSpacing(20)
        setting_layout.addWidget(QtWidgets.QLabel("첫 페이지 무시 수:"))
        setting_layout.addWidget(self.skip_spin)
        layout.addLayout(setting_layout)

        # 실행 버튼
        self.start_btn = QtWidgets.QPushButton("실행")
        self.start_btn.clicked.connect(self.start_processing)
        layout.addWidget(self.start_btn)

        # 진행 및 로그 출력 창
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(QtWidgets.QLabel("진행 로그:"))
        layout.addWidget(self.log_text)

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def browse_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "HWPX/hwtx 파일 선택", "", "HWPX Files (*.hwpx *.hwtx);;All Files (*)"
        )
        if file_path:
            self.input_line.setText(file_path)

    def select_output_dir(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "출력 폴더 선택", "")
        if folder:
            self.output_line.setText(folder)

    def start_processing(self):
        input_file = self.input_line.text().strip()
        output_dir = self.output_line.text().strip()
        template = self.template_line.text().strip()
        skip_pages = self.skip_spin.value()

        if not input_file or not os.path.isfile(input_file):
            self.log("유효한 입력 파일을 선택해주세요.")
            return

        if not output_dir:
            self.log("출력 폴더를 지정해주세요.")
            return

        self.log("분리 처리를 시작합니다...")
        QtCore.QTimer.singleShot(100, lambda: process_file(input_file, output_dir, template, skip_pages, self.log))

# =======================
# GUI - 합치기 탭
# =======================
class MergeTab(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # 합칠 파일이 있는 폴더 선택 영역
        input_layout = QtWidgets.QHBoxLayout()
        self.merge_input_line = QtWidgets.QLineEdit()
        self.merge_input_line.setPlaceholderText("합칠 HWPX/hwtx 파일이 있는 폴더 경로")
        input_btn = QtWidgets.QPushButton("폴더 선택")
        input_btn.clicked.connect(self.select_input_folder)
        input_layout.addWidget(QtWidgets.QLabel("입력 폴더:"))
        input_layout.addWidget(self.merge_input_line)
        input_layout.addWidget(input_btn)
        layout.addLayout(input_layout)

        # 출력 폴더 선택 영역 (병합 파일 저장 경로)
        output_layout = QtWidgets.QHBoxLayout()
        self.merge_output_line = QtWidgets.QLineEdit("merged_output")
        output_btn = QtWidgets.QPushButton("폴더 선택")
        output_btn.clicked.connect(self.select_output_folder)
        output_layout.addWidget(QtWidgets.QLabel("출력 폴더:"))
        output_layout.addWidget(self.merge_output_line)
        output_layout.addWidget(output_btn)
        layout.addLayout(output_layout)

        # 실행 버튼
        self.merge_btn = QtWidgets.QPushButton("합치기 실행")
        self.merge_btn.clicked.connect(self.start_merging)
        layout.addWidget(self.merge_btn)

        # 진행 및 로그 출력 창
        self.log_text = QtWidgets.QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(QtWidgets.QLabel("진행 로그:"))
        layout.addWidget(self.log_text)

    def log(self, message):
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def select_input_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "합칠 파일이 있는 폴더 선택", "")
        if folder:
            self.merge_input_line.setText(folder)

    def select_output_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "출력 폴더 선택", "")
        if folder:
            self.merge_output_line.setText(folder)

    def start_merging(self):
        input_folder = self.merge_input_line.text().strip()
        output_folder = self.merge_output_line.text().strip()

        if not input_folder or not os.path.isdir(input_folder):
            self.log("유효한 입력 폴더를 선택해주세요.")
            return

        if not output_folder:
            self.log("출력 폴더를 지정해주세요.")
            return

        self.log("병합 처리를 시작합니다...")
        QtCore.QTimer.singleShot(100, lambda: merge_hwpx_files(input_folder, output_folder, self.log))

# =======================
# 메인 윈도우 (탭 구성)
# =======================
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWPX 문서 분리 및 합치기 도구")
        self.setMinimumSize(600, 400)
        self.init_ui()

    def init_ui(self):
        tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(tabs)

        self.split_tab = SplitTab()
        self.merge_tab = MergeTab()

        tabs.addTab(self.split_tab, "분리하기")
        tabs.addTab(self.merge_tab, "합치기")

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec_())
