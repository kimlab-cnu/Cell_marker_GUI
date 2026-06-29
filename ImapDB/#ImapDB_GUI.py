import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt

import pandas as pd

# --- 이 스크립트(.py)가 들어있는 폴더를 기준 폴더로 잡는다 ---
# 모든 .ui / .xlsx / .jpg 파일을 이 .py 와 "같은 폴더"에 넣어두면
# 윈도우/맥 어디서 실행하든 알아서 찾는다. (경로를 직접 고칠 필요 없음)
BASE = os.path.dirname(os.path.abspath(__file__))
def P(filename):
    return os.path.join(BASE, filename)

# 파일명을 대소문자 무시하고 폴더에서 찾아준다 (맥은 대소문자를 구분하므로)
def resolve_path(filename):
    target = os.path.join(BASE, filename)
    if os.path.exists(target):
        return target
    low = filename.lower()
    try:
        for f in os.listdir(BASE):
            if f.lower() == low:
                return os.path.join(BASE, f)
    except Exception:
        pass
    return target  # 못 찾으면 원래 경로 반환(에러 메시지에 기대한 파일명이 보이도록)


# --- 패널 라벨 매핑: panel_labels.xlsx 가 있으면 시트 이름 대신 기능 라벨을 보여준다 ---
def _norm_key(s):
    return str(s).replace('\ufeff', '').replace('\u200b', '').replace('\xa0', ' ').strip().lower()

def load_panel_labels():
    mapping = {}
    path = resolve_path('panel_labels.xlsx')
    if os.path.exists(path):
        try:
            m = pd.read_excel(path)
            for _, r in m.iterrows():
                f = _norm_key(r.get('file', ''))
                s = _norm_key(r.get('sheet', ''))
                lab = str(r.get('label', '')).strip()
                if f and s and lab and lab.lower() != 'nan':
                    mapping[(f, s)] = lab
        except Exception as e:
            print('panel_labels load error:', e)
    return mapping

PANEL_LABELS = load_panel_labels()

def panel_label(code_file, sheet_name):
    # 매핑에 라벨이 있으면 그 라벨, 없으면 원래 시트 이름을 그대로 사용
    key = (_norm_key(os.path.basename(code_file)), _norm_key(sheet_name))
    return PANEL_LABELS.get(key, sheet_name)


# --- 패널 복잡도(항체 파라미터 수) 추정: 복잡도 필터의 기준값 ---
import re as _re
_MARKER_RE = _re.compile(
    r'CD\d+[a-zA-Z]*|HLA-?DR|CCR\d+|CXCR\d+|FoxP3|Ki-?67|PD-?1|PD-?L1|TIM-?3|'
    r'LAG-?3|TIGIT|TCR\w*|Ig[MDGAE]|NKG2[A-Z]|CD45RA|CD45RO|CTLA-?4|ICOS', _re.I)

def count_panel_params(df):
    """로드된 패널 시트(df)에서 서로 다른 항체 마커(파라미터) 수를 추정한다."""
    try:
        text = ' '.join(str(v).replace('\xa0', ' ') for v in df.values.flatten()
                        if isinstance(v, str))
        text += ' ' + ' '.join(str(c) for c in df.columns)
        ms = set(m.group(0).upper().replace(' ', '') for m in _MARKER_RE.finditer(text))
        return len(ms)
    except Exception:
        return 0

Lym_file = pd.read_excel(P("Lym_input.xlsx"))
marker_data = Lym_file['Cell Type_1']
marker_data1 = marker_data.values.tolist()
marker_data2 = pd.Series(marker_data1).dropna().tolist()
# C2 (Cell Type) placeholder prompt
marker_data2.insert(0, 'Cell Type')
new_list = []
for v in marker_data2:
    if v not in new_list:
        new_list.append(v)
Lym_word = new_list

My_file = pd.read_excel(P("Myeloid_input.xlsx"))
my_data = My_file['Cell Type_1']
my_data1 = my_data.values.tolist()
my_data2 = pd.Series(my_data1).dropna().tolist()
# C2 (Cell Type) placeholder prompt
my_data2.insert(0, 'Cell Type')
my_list = []
for v in my_data2:
    if v not in my_list:
        my_list.append(v)
My_word = my_list
All_file = pd.read_excel(P("All_input.xlsx"))

form_home = uic.loadUiType(P("Intro_Real_fin.ui"))[0]
form_immune = uic.loadUiType(P("Immune_Dialog.ui"))[0]
form_CyTOF = uic.loadUiType(P("CyTOF_Dialog.ui"))[0]
form_FACS = uic.loadUiType(P("FACS_Dialog.ui"))[0]


# --- 화면에 보이는 표(QTableWidget)를 그대로 엑셀(.xlsx)로 내보내는 공용 함수 ---
def export_table_to_excel(parent, table, default_name="export.xlsx"):
    rows = table.rowCount()
    cols = table.columnCount()
    if rows == 0 or cols == 0:
        QMessageBox.information(parent, "Export", "내보낼 데이터가 없습니다. 먼저 마커/패널을 조회하세요.")
        return

    # 표 맨 윗줄의 컬럼 이름 가져오기 (없으면 Column 1, 2 ...)
    headers = []
    for c in range(cols):
        h = table.horizontalHeaderItem(c)
        headers.append(h.text() if h is not None else "Column %d" % (c + 1))

    # 표 내용 읽기
    data = []
    for r in range(rows):
        row_values = []
        for c in range(cols):
            it = table.item(r, c)
            row_values.append(it.text() if it is not None else "")
        data.append(row_values)

    df = pd.DataFrame(data, columns=headers)

    # 저장 위치/파일명 선택 창
    path, _ = QFileDialog.getSaveFileName(parent, "Save as Excel", default_name, "Excel Files (*.xlsx)")
    if not path:
        return
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    try:
        df.to_excel(path, index=False)
        QMessageBox.information(parent, "Export", "저장 완료:\n%s" % path)
    except Exception as e:
        QMessageBox.critical(parent, "Export error", str(e))


class Main(QMainWindow, QWidget, form_home):
    def __init__(self):
        super().__init__()
        self.initUI()


    def initUI(self):
        KimLab = QLabel(self)
        KimLab.setPixmap(QPixmap(P("KimLab_Logo.jpg")))
        KimLab.setGeometry(345, 450, 220, 80)

        KimLab.setScaledContents(True)
        KimLab.setAlignment(Qt.AlignCenter)

        self.setupUi(self)
        self.Immune_btn.clicked.connect(self.btn_second)
        self.CyTOF_btn.clicked.connect(self.btn_CyTOF)
        self.FACS_btn.clicked.connect(self.btn_FACS)
        self.show()

    def btn_second(self): #Immune cell
        self.close()
        self.second = HomeWindow()
        self.second.exec()
        self.show()


    def btn_CyTOF(self): #CyTOF panel
        self.close()
        self.CyTOF = CyTOFWindow()
        self.CyTOF.exec()
        self.show()

    def btn_FACS(self): #FACS panel
        self.close()
        self.FACS = FACSWindow()
        self.FACS.exec()
        self.show()


class HomeWindow(QDialog, QWidget, form_immune):
    def __init__(self):
        super(HomeWindow, self).__init__()
        self.initUI()
        self.show()
        self.setWindowTitle("Immune Cell Marker")

        self.C1.currentIndexChanged.connect(self.c1_changed)
        self.C2.currentIndexChanged.connect(self.c2_changed)
        self.Done_btn.clicked.connect(self.accept)

        # --- Export to Excel 버튼: Markers 표의 오른쪽 아래 구석에 배치 ---
        self.Export_btn = QPushButton("Export to Excel", self)
        self.Export_btn.adjustSize()
        g = self.Show_Marker.geometry()
        bw = self.Export_btn.width()
        bh = self.Export_btn.height()
        self.Export_btn.move(g.x() + g.width() - bw - 6,
                             self.height() - bh - 10)
        self.Export_btn.clicked.connect(
            lambda: export_table_to_excel(self, self.Show_Marker, "Immune_markers.xlsx"))
        self.Export_btn.raise_()
        self.Export_btn.show()

        # --- 검색창: 계층 드롭다운 없이도 cell type/subtype/marker 를 바로 찾기 ---
        self.Search_label = QLabel("Search:", self)
        self.Search_box = QLineEdit(self)
        self.Search_box.setPlaceholderText("e.g. Treg, exhausted, FoxP3, CD25 ...")
        self.Search_btn = QPushButton("Search", self)
        self.Search_box.returnPressed.connect(self._run_search)
        self.Search_btn.clicked.connect(self._run_search)
        for w in (self.Search_label, self.Search_box, self.Search_btn):
            w.show()
        self._position_search()

        # --- 창을 크게 + 크기 조절 가능하게 (.ui의 고정 크기 제한 해제) ---
        self.setMaximumSize(16777215, 16777215)
        self.setMinimumSize(820, 600)
        self.resize(1100, 820)
        # 표 스크롤바는 필요할 때 자동으로 보이도록
        self.Show_Marker.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.Show_Marker.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # Markers 표(및 그 그룹박스)가 창 크기를 따라가도록 처음 위치 기억
        self._gb = self.Show_Marker.parentWidget()
        self._gb_origin = (self._gb.x(), self._gb.y())
        self._table_origin = (self.Show_Marker.x(), self.Show_Marker.y())
        self._layout_table()

    # --- 세 단계 선택 이름을 Figure 2 범례와 맞춤: Lineage / Cell Type / Subtype ---
    # "First / Second / Third cell type" 라벨과 "Thrid" 오타는 Immune_Dialog.ui 안에 있다.
    # 아래 코드가 실행 시점에 라벨 글자를 검사해서 자동으로 바꿔준다(오타 포함).
    def _relabel_levels(self):
        mapping = [
            (("first",), "Lineage"),
            (("second",), "Cell Type"),
            (("third", "thrid"), "Subtype"),
        ]
        for lbl in self.findChildren(QLabel):
            text = lbl.text().strip().lower()
            for keys, new_text in mapping:
                if any(k in text for k in keys):
                    lbl.setText(new_text)
                    break

    # 창 크기에 맞춰 Markers 그룹박스/표/Export 버튼 위치·크기를 다시 잡는다
    def _layout_table(self):
        if not hasattr(self, '_table_origin') or not hasattr(self, 'Export_btn'):
            return
        margin = 16
        gb = self._gb
        if gb is not self:
            # 표가 "Markers" 그룹박스 안에 있는 경우: 그룹박스를 먼저 창에 맞춰 늘림
            gx, gy = self._gb_origin
            gw = max(200, self.width() - gx - margin)
            gh = max(160, self.height() - gy - margin)
            gb.setGeometry(gx, gy, gw, gh)
            avail_w, avail_h = gb.width(), gb.height()
        else:
            avail_w, avail_h = self.width(), self.height()
        # 표를 (부모 기준) 가득 채움
        tx, ty = self._table_origin
        tw = max(150, avail_w - tx - margin)
        th = max(120, avail_h - ty - margin - 8)
        self.Show_Marker.setGeometry(tx, ty, tw, th)
        # Export 버튼은 창 오른쪽 아래
        bw = self.Export_btn.width()
        bh = self.Export_btn.height()
        self.Export_btn.move(self.width() - bw - margin, self.height() - bh - 10)
        self.Export_btn.raise_()
        self._position_search()

    # 검색창을 Home 버튼 왼쪽(같은 높이)에 배치해 겹치지 않게 한다.
    # 위치를 옮기고 싶으면 box_w(검색창 폭)나 gap 값만 바꾸면 됩니다.
    def _position_search(self):
        if not hasattr(self, 'Search_box'):
            return
        margin = 16
        gap = 10
        box_w = 240
        self.Search_btn.adjustSize()
        self.Search_label.adjustSize()
        btn_w = self.Search_btn.width()
        lbl_w = self.Search_label.width()
        h = max(self.Search_box.sizeHint().height(), 24)

        # Home 버튼이 있으면 그 왼쪽을 오른쪽 한계로, 같은 y에 정렬
        if hasattr(self, 'Home_btn') and self.Home_btn is not None:
            hb = self.Home_btn.geometry()
            right_limit = hb.x() - gap
            top = hb.y() + max(0, (hb.height() - h) // 2)
        else:
            right_limit = self.width() - margin
            top = 10

        x_btn = right_limit - btn_w
        x_box = x_btn - 6 - box_w
        x_lbl = x_box - 6 - lbl_w
        # 폭이 모자라면 검색창을 줄여서라도 왼쪽 한계 안에 들어오게
        if x_lbl < margin:
            shift = margin - x_lbl
            box_w = max(120, box_w - shift)
            x_btn = right_limit - btn_w
            x_box = x_btn - 6 - box_w
            x_lbl = x_box - 6 - lbl_w
        self.Search_label.move(max(margin, x_lbl), top + 3)
        self.Search_box.setGeometry(max(margin + lbl_w + 6, x_box), top, box_w, h)
        self.Search_btn.move(x_btn, top)
        for w in (self.Search_label, self.Search_box, self.Search_btn):
            w.raise_()

    def resizeEvent(self, event):
        super(HomeWindow, self).resizeEvent(event)
        self._layout_table()

    def c1_changed(self, index):
        selected_option = self.C1.currentText()
        self.C2.clear()
        if selected_option == "Lymphocyte":
            self.C2.addItems(Lym_word)
        elif selected_option == "Myeloid cell":
            self.C2.addItems(My_word)
        # C1 = Lineage 선택. 기존 기본값("Cell type")과 새 표기("Lineage") 모두 허용
        elif selected_option in ("Lineage", "Cell type"):
            self.C2.addItem("Cell Type")

    def c2_changed(self, index):
        selected_option1 = self.C1.currentText()
        selected_option = self.C2.currentText()

        plan = Lym_file[Lym_file['Cell Type_1'] == selected_option]
        plan1 = plan.values.tolist()
        plan2 = pd.Series(plan1).dropna().tolist()
        plan2.insert(0, 'Cell Type')

        want = plan['Cell Type_2']
        want1 = want.values.tolist()
        want2 = pd.Series(want1).dropna().tolist()
        # C3 = Subtype 선택 -> 안내 문구를 "Subtype"으로
        want2.insert(0, 'Subtype')

        new_list1 = []
        for v in want2:
            if v not in new_list1:
                new_list1.append(v)
        List = new_list1

        planB = My_file[My_file['Cell Type_1'] == selected_option]
        planB1 = planB.values.tolist()
        planB2 = pd.Series(planB1).dropna().tolist()
        planB2.insert(0, 'Cell Type')

        wantB = planB['Cell Type_2']
        wantB1 = wantB.values.tolist()
        wantB2 = pd.Series(wantB1).dropna().tolist()
        # C3 = Subtype 선택 -> 안내 문구를 "Subtype"으로
        wantB2.insert(0, 'Subtype')

        my_list1 = []
        for v in wantB2:
            if v not in my_list1:
                my_list1.append(v)
        ListB = my_list1

        self.C3.clear()
        if selected_option1 == "Lymphocyte":
            self.C3.addItems(List)
        elif selected_option1 == "Myeloid cell":
            self.C3.addItems(ListB)

    def accept(self):
        selected_option1 = self.C2.currentText()
        selected_option2 = self.C3.currentText()
        marker = All_file[(All_file['Cell Type_1'] == selected_option1) & (All_file['Cell Type_2'] == selected_option2)]
        self._render_dataframe(marker)

    # 표 그리기 공통 함수 (드롭다운 선택과 검색이 함께 사용)
    def _render_dataframe(self, marker):
        self.Show_Marker.setFont(QFont("Arial", 13))   # 표 안 글씨 크기
        self.Show_Marker.setRowCount(marker.shape[0])
        self.Show_Marker.setColumnCount(marker.shape[1])
        # 표 맨 윗줄에 컬럼 이름 표시 (Lineage / Cell Type / Subtype / Marker / Location / Defining level)
        self.Show_Marker.setHorizontalHeaderLabels([str(c) for c in marker.columns])
        for row in range(marker.shape[0]):
           for col in range(marker.shape[1]):
                item = QTableWidgetItem(str(marker.iloc[row, col]))
                self.Show_Marker.setItem(row, col, item)
        self.Show_Marker.resizeColumnsToContents()

    # --- 검색: cell type / subtype / marker 어디든 일치하면 표시 (계층 몰라도 찾기 가능) ---
    def _run_search(self):
        q = self.Search_box.text()
        q = str(q).replace('\xa0', ' ').strip().lower()
        if not q:
            return
        cols = [c for c in ['Lineage', 'Cell Type_1', 'Cell Type_2',
                            'Marker', 'Defining level', 'Location'] if c in All_file.columns]
        def row_matches(r):
            for c in cols:
                v = str(r[c]).replace('\xa0', ' ').lower()
                if q in v:
                    return True
            return False
        hits = All_file[All_file.apply(row_matches, axis=1)]
        self._render_dataframe(hits)


    def initUI(self):
        self.setupUi(self)
        # Lineage / Cell Type / Subtype 라벨 적용 (+ "Thrid" 오타 수정)
        self._relabel_levels()
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


class CyTOFWindow(QDialog, QWidget, form_CyTOF):
    def __init__(self):
        super(CyTOFWindow, self).__init__()
        self.initUI()
        self.show()
        self.resize(1200, 800)
        self.df_list = []

        horizon = QHBoxLayout()
        horizon.addWidget(self.sheet, 4)
        # 복잡도 필터: 보유 장비가 감당 가능한 채널 수 기준으로 패널을 거른다
        self.Complexity = QComboBox()
        self.Complexity.addItem("All panels", None)
        self.Complexity.addItem("\u2264 20 parameters", 20)
        self.Complexity.addItem("\u2264 30 parameters", 30)
        self.Complexity.addItem("\u2264 40 parameters", 40)
        self.Complexity.currentIndexChanged.connect(self._on_complexity_changed)
        horizon.addWidget(self.Complexity, 3)
        horizon.addWidget(self.Com_btn,2)
        horizon.addWidget(self.Ex_btn,2)
        self._combo = self.sheet
        self._visible_idx = []
        self._current_base = ''
        self._is_total = False
        self.Total_btn = QPushButton("Total")
        self.Total_btn.clicked.connect(lambda: self._open_total([
            (resolve_path('Com_CyTOF_panel.xlsx'), 'Com'),
            (resolve_path('Ex_CyTOF_panel.xlsx'), 'Exp'),
        ]))
        horizon.addWidget(self.Total_btn, 2)
        self.Export_btn = QPushButton("Export to Excel")
        self.Export_btn.clicked.connect(
            lambda: export_table_to_excel(self, self.panel, "CyTOF_panel.xlsx"))
        horizon.addWidget(self.Export_btn, 2)
        horizon.addWidget(self.Home_btn,1)

        vertical = QVBoxLayout()
        vertical.addWidget(self.label)
        vertical.addLayout(horizon)
        vertical.addWidget(self.panel)
        self.setLayout(vertical)
        self.Com_btn.clicked.connect(self.Com_open)
        self.Ex_btn.clicked.connect(self.Ex_open)
        self.sheet.currentIndexChanged[int].connect(self.cmbChanged)
        # macOS에서 다이얼로그가 버튼 하나를 자동으로 '기본 버튼(파란색)'으로
        # 칠하는 것을 막는다. 이제 클릭한 버튼에만 강조가 생긴다.
        for _b in self.findChildren(QPushButton):
            _b.setAutoDefault(False)
            _b.setDefault(False)

    def Com_open(self):
        self._open_panel(resolve_path('Com_CyTOF_panel.xlsx'))

    def Ex_open(self):
        self._open_panel(resolve_path('Ex_CyTOF_panel.xlsx'))

    def _open_panel(self, file_path):
        self._current_base = os.path.basename(file_path)
        self._is_total = False
        self.df_list = self.loadData(file_path)
        # 각 시트의 파라미터 수를 미리 계산해 둔다(필터 기준)
        self._counts = [count_panel_params(d) for d in self.df_list]
        self._refresh_sheet_list()

    # 복잡도 필터를 반영해 드롭다운 목록을 다시 채운다
    def _refresh_sheet_list(self):
        if self._is_total:
            return
        thr = self.Complexity.currentData()
        self._combo.blockSignals(True)
        self._combo.clear()
        self._visible_idx = []
        for i, d in enumerate(self.df_list):
            n = self._counts[i] if hasattr(self, '_counts') else count_panel_params(d)
            if thr is not None and n > thr:
                continue
            lab = panel_label(self._current_base, d.name)
            # 라벨에 마커/파라미터 수가 없으면 복잡도를 보조 표기
            if n > 0 and ('marker' not in lab.lower() and 'param' not in lab.lower()):
                lab = "%s  (%dp)" % (lab, n)
            self._combo.addItem(lab)
            self._visible_idx.append(i)
        self._combo.blockSignals(False)
        if self._visible_idx:
            self._combo.setCurrentIndex(0)
            self.initTableWidget(self._visible_idx[0])
        else:
            self.panel.clearContents()
            self.panel.setRowCount(0)

    def _on_complexity_changed(self, _=None):
        if getattr(self, '_is_total', False):
            return
        if self.df_list:
            self._refresh_sheet_list()

    # 상용 + 실험 패널을 모두 하나의 표에 이어붙여 보여준다 ("Total")
    def _combine_all(self, specs):
        frames = []
        for path, tag in specs:
            base = os.path.basename(path)
            for d in self.loadData(path):
                d2 = d.copy()
                d2.insert(0, 'Panel', "[%s] %s" % (tag, panel_label(base, d.name)))  # 출처+기능 라벨
                frames.append(d2)
        if not frames:
            return None
        merged = pd.concat(frames, ignore_index=True, sort=False).fillna('')
        merged.name = 'Total'
        return merged

    def _open_total(self, specs):
        merged = self._combine_all(specs)
        self._is_total = True
        self.sheet.blockSignals(True)
        self.sheet.clear()
        if merged is None:
            self.sheet.blockSignals(False)
            self.df_list = []
            self._visible_idx = []
            self.panel.clear()
            return
        self.df_list = [merged]
        self._visible_idx = [0]
        self.sheet.addItem('Total (all panels)')
        self.sheet.blockSignals(False)
        self.sheet.setCurrentIndex(0)
        self.initTableWidget(0)

    def cmbChanged(self, pos):
        # 콤보의 표시 위치를 원래 df_list 인덱스로 변환(필터로 일부만 보일 수 있음)
        if pos is None or pos < 0 or pos >= len(self._visible_idx):
            return
        self.initTableWidget(self._visible_idx[pos])

    def loadData(self, file_name):
        df_list = []
        try:
            wb = pd.ExcelFile(file_name)
        except Exception as e:
            QMessageBox.critical(
                self, "File error",
                "엑셀 파일을 열 수 없습니다:\n%s\n\n%s" % (file_name, e))
            return df_list
        with wb:
            for i, sn in enumerate(wb.sheet_names):
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    df.name = sn
                    df_list.append(df)
        return df_list

    def initTableWidget(self, id):
        # 잘못된 인덱스(-1 등)나 빈 목록이면 아무것도 안 하고 종료 (강제종료 방지)
        if id is None or id < 0 or id >= len(self.df_list):
            return
        self.panel.clear()
        self.panel.setFont(QFont("Arial", 13))      # CyTOF 표 글씨 크기
        df = self.df_list[id]
        col = len(df.keys())
        self.panel.setColumnCount(col)
        self.panel.setHorizontalHeaderLabels(df.keys())

        row = len(df.index)
        self.panel.setRowCount(row)
        self.writeTableWidget(id, df, row, col)

    def writeTableWidget(self, id, df, row, col):
        for r in range(row):
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r, c]))
                self.panel.setItem(r, c, item)
        self.panel.resizeColumnsToContents()

    def initUI(self):
        self.setupUi(self)
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


class FACSWindow(QDialog, QWidget, form_FACS):
    def __init__(self):
        super(FACSWindow, self).__init__()
        self.initUI()
        self.show()
        self.resize(1200, 800)  # 위젯 사이즈
        self.df_list = []

        horizon = QHBoxLayout()
        horizon.addWidget(self.sheet1, 4)
        # 복잡도 필터: 보유 flow cytometer 의 색(채널) 수용량 기준으로 패널을 거른다
        self.Complexity = QComboBox()
        self.Complexity.addItem("All panels", None)
        self.Complexity.addItem("\u2264 10-color", 10)
        self.Complexity.addItem("\u2264 18-color", 18)
        self.Complexity.addItem("\u2264 28-color", 28)
        self.Complexity.currentIndexChanged.connect(self._on_complexity_changed)
        horizon.addWidget(self.Complexity, 3)
        horizon.addWidget(self.Com_btn,2)
        horizon.addWidget(self.Ex_btn,2)
        self._combo = self.sheet1
        self._visible_idx = []
        self._current_base = ''
        self._is_total = False
        self.Total_btn = QPushButton("Total")
        self.Total_btn.clicked.connect(lambda: self._open_total([
            (resolve_path('Com_FACS_Panel.xlsx'), 'Com'),
            (resolve_path('Ex_FACS_panel.xlsx'), 'Exp'),
        ]))
        horizon.addWidget(self.Total_btn, 2)
        self.Export_btn = QPushButton("Export to Excel")
        self.Export_btn.clicked.connect(
            lambda: export_table_to_excel(self, self.panel1, "FACS_panel.xlsx"))
        horizon.addWidget(self.Export_btn, 2)
        horizon.addWidget(self.Home_btn,1)

        vertical = QVBoxLayout()

        vertical.addWidget(self.label1)
        vertical.addLayout(horizon)
        vertical.addWidget(self.panel1)

        self.setLayout(vertical)

        self.Com_btn.clicked.connect(self.Com_open)
        self.Ex_btn.clicked.connect(self.Ex_open)

        self.sheet1.currentIndexChanged[int].connect(self.cmbChanged)
        # macOS에서 다이얼로그가 버튼 하나를 자동으로 '기본 버튼(파란색)'으로
        # 칠하는 것을 막는다. 이제 클릭한 버튼에만 강조가 생긴다.
        for _b in self.findChildren(QPushButton):
            _b.setAutoDefault(False)
            _b.setDefault(False)

    def Com_open(self):
        self._open_panel(resolve_path('Com_FACS_Panel.xlsx'))

    def Ex_open(self):
        self._open_panel(resolve_path('Ex_FACS_panel.xlsx'))

    def _open_panel(self, file_path):
        self._current_base = os.path.basename(file_path)
        self._is_total = False
        self.df_list = self.loadData(file_path)
        self._counts = [count_panel_params(d) for d in self.df_list]
        self._refresh_sheet_list()

    def _refresh_sheet_list(self):
        if self._is_total:
            return
        thr = self.Complexity.currentData()
        self._combo.blockSignals(True)
        self._combo.clear()
        self._visible_idx = []
        for i, d in enumerate(self.df_list):
            n = self._counts[i] if hasattr(self, '_counts') else count_panel_params(d)
            if thr is not None and n > thr:
                continue
            lab = panel_label(self._current_base, d.name)
            if n > 0 and ('marker' not in lab.lower() and 'param' not in lab.lower()):
                lab = "%s  (%dp)" % (lab, n)
            self._combo.addItem(lab)
            self._visible_idx.append(i)
        self._combo.blockSignals(False)
        if self._visible_idx:
            self._combo.setCurrentIndex(0)
            self.initTableWidget(self._visible_idx[0])
        else:
            self.panel1.clearContents()
            self.panel1.setRowCount(0)

    def _on_complexity_changed(self, _=None):
        if getattr(self, '_is_total', False):
            return
        if self.df_list:
            self._refresh_sheet_list()

    # 상용 + 실험 패널을 모두 하나의 표에 이어붙여 보여준다 ("Total")
    def _combine_all(self, specs):
        frames = []
        for path, tag in specs:
            base = os.path.basename(path)
            for d in self.loadData(path):
                d2 = d.copy()
                d2.insert(0, 'Panel', "[%s] %s" % (tag, panel_label(base, d.name)))
                frames.append(d2)
        if not frames:
            return None
        merged = pd.concat(frames, ignore_index=True, sort=False).fillna('')
        merged.name = 'Total'
        return merged

    def _open_total(self, specs):
        merged = self._combine_all(specs)
        self._is_total = True
        self.sheet1.blockSignals(True)
        self.sheet1.clear()
        if merged is None:
            self.sheet1.blockSignals(False)
            self.df_list = []
            self._visible_idx = []
            self.panel1.clear()
            return
        self.df_list = [merged]
        self._visible_idx = [0]
        self.sheet1.addItem('Total (all panels)')
        self.sheet1.blockSignals(False)
        self.sheet1.setCurrentIndex(0)
        self.initTableWidget(0)

    def cmbChanged(self, pos):
        if pos is None or pos < 0 or pos >= len(self._visible_idx):
            return
        self.initTableWidget(self._visible_idx[pos])

    def loadData(self, file_name):
        df_list = []
        try:
            wb = pd.ExcelFile(file_name)
        except Exception as e:
            QMessageBox.critical(
                self, "File error",
                "엑셀 파일을 열 수 없습니다:\n%s\n\n%s" % (file_name, e))
            return df_list
        with wb:
            for i, sn in enumerate(wb.sheet_names):
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    df.name = sn
                    df_list.append(df)
        return df_list

    def initTableWidget(self, id):
        # 잘못된 인덱스(-1 등)나 빈 목록이면 아무것도 안 하고 종료 (강제종료 방지)
        if id is None or id < 0 or id >= len(self.df_list):
            return
        self.panel1.clear()
        self.panel1.setFont(QFont("Arial", 13))     # FACS 표 글씨 크기
        df = self.df_list[id]
        col = len(df.keys())
        self.panel1.setColumnCount(col)
        self.panel1.setHorizontalHeaderLabels(df.keys())

        row = len(df.index)
        self.panel1.setRowCount(row)
        self.writeTableWidget(id, df, row, col)

    def writeTableWidget(self, id, df, row, col):
        for r in range(row):
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r, c]))
                self.panel1.setItem(r, c, item)
        self.panel1.resizeColumnsToContents()

    def initUI(self):
        self.setupUi(self)
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # --- 전체 글씨 크기 ---
    # 아래 숫자(13)를 키우면 프로그램 전체 글씨가 커진다. (예: 15, 16, 18 ...)
    app.setFont(QFont("Arial", 13))
    m = Main()
    m.show()
    sys.exit(app.exec_())
