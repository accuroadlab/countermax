# -*- coding: utf-8 -*-
"""
환경설정 (프로 플러스) — all‑in‑one patch8
요약
- 차종관리: 차종 설정 / 카운터 단축키 설정(좌·직·우, 추가/수정/삭제) / 카운터 템플릿 관리
- 조사관리: 조사정보(용도/진행상태/SN 자동/등록일/발주처/기간/설명),
            조사시간 설정(구간 누적 추가 + 자동생성 + 행 추가/삭제),
            조사차종 설정(프로젝트 참조 + 로컬편집),
            조사지점 설정(좌: 지점 리스트, 우: 방향→그룹, 카운터, 시트 미리보기)
- 시트 미리보기: 탭=방향(그룹), 열=시간대 + 현재 조사차종(차종명) 컬럼들
데이터 파일: env_data_plus_allinone.json
"""
import os, sys, json, datetime as dt
from PyQt5 import QtWidgets, QtCore, QtGui

APP_VER = "1.1.6-patch8"

# ----- NAS / Survey 데이터 루트 설정 -----
# 외부/내부에서 브라우저로 접근할 때 사용하는 HTTP 루트 (참고용 상수)
EXTERNAL_HTTP_ROOT = "http://accuroad.synology.me:5096/Survey"
INTERNAL_HTTP_ROOT = "http://192.168.35.239:5096/Survey"

# 실제 파일 입출력은 윈도우에서 인식하는 경로(드라이브/UNC)를 사용합니다.
# 1순위: 환경변수 COUNTERMAX_DATA_ROOT
# 2순위: 아래 SMB_CANDIDATES 중 실제로 존재하는 경로
# 3순위: 이 env_settings 파일이 위치한 폴더
import os as _os

LOCAL_ROOT = _os.path.dirname(__file__)

SMB_CANDIDATES = [
    r"K:\Survey",                    # WebDAV나 SMB를 K:로 연결한 경우
    r"Z:\Survey",                    # 예비 드라이브 문자
    r"Y:\Survey",                    # 예비 드라이브 문자

    # WebDAV UNC 후보 (드라이브 매핑 없이도 접근 가능할 때)
    r"\\accuroad.synology.me@5096\DavWWWRoot\Survey",
    r"\\192.168.35.239@5096\DavWWWRoot\Survey",
    # SMB 공유 후보
    r"\\192.168.35.239\Survey",
]

def detect_data_root():
    env = _os.environ.get("COUNTERMAX_DATA_ROOT")
    if env and _os.path.isdir(env):
        return env
    for cand in SMB_CANDIDATES:
        try:
            if _os.path.isdir(cand):
                return cand
        except Exception:
            pass
    return LOCAL_ROOT

DATA_ROOT = detect_data_root()
DATA_PATH = _os.path.join(DATA_ROOT, "env_data_plus_allinone.json")

# ----- Projects / DAT 파일 저장 경로 -----
# 과업(SN_...) 단위로 DAT를 저장하는 루트 폴더:
#   <DATA_ROOT>/Projects/SN_xxxxxxxxxxxxxxxx/WN_xxx_YYYYMMDD_USERID.dat
PROJECTS_ROOT = _os.path.join(DATA_ROOT, "Projects")


def ensure_project_dir(survey_no: str) -> str:
    """
    과업번호(SN_...) 기준으로 Projects/SN_xxx 폴더를 생성하고 경로를 반환합니다.
    """
    if not survey_no:
        raise ValueError("survey_no(SN_...)가 비어 있습니다.")
    folder = _os.path.join(PROJECTS_ROOT, survey_no)
    _os.makedirs(folder, exist_ok=True)
    return folder


def format_korean_datetime(value) -> str:
    """
    datetime 값을 'YYYY-MM-DD 오전/오후 HH:MM:SS' 형식 문자열로 변환합니다.
    이미 문자열이면 그대로 반환합니다.
    """
    if isinstance(value, str):
        return value
    if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
        value = dt.datetime(value.year, value.month, value.day)
    if not isinstance(value, dt.datetime):
        raise TypeError("datetime 또는 날짜 문자열이어야 합니다.")
    ampm = "오전" if value.hour < 12 else "오후"
    hour12 = value.hour % 12
    if hour12 == 0:
        hour12 = 12
    return f"{value.strftime('%Y-%m-%d')} {ampm} {hour12:02d}:{value.minute:02d}:{value.second:02d}"


def write_survey_dat(
    survey_no: str,   # SN_...
    work_no: str,     # WN_...
    user_id: str,     # 로그인 ID (예: hkw3316)
    user_name: str,   # 사용자 이름
    rec_date,         # 조사일 (date, datetime, 또는 'YYYY-MM-DD' 문자열)
    start_time,       # 시작시간 (datetime 또는 이미 포맷된 문자열)
    last_time,        # 종료시간 (datetime 또는 이미 포맷된 문자열)
    seq_no: int = 0,
    state: int = 1,
    deleted: int = 0,
    matrices=None,    # 계수 데이터 (2차원 리스트들의 리스트)
):
    """
    조사/지점 작업 1건에 대한 DAT 파일을 생성합니다.

    파일 경로:
        Projects/<SURVEY_NO>/<WORK_NO>_YYYYMMDD_<USER_ID>.dat

    헤더 형식은 기존 예시 DAT와 동일합니다.
    matrices 는 2차원 리스트들의 리스트입니다.
        예: [ mat0, mat1, ... ]
        matN 은 rows x cols 크기의 int 값 리스트입니다.
    """
    # 조사일 문자열 처리
    if isinstance(rec_date, (dt.date, dt.datetime)):
        rec_date_str = rec_date.strftime("%Y-%m-%d")
    else:
        rec_date_str = str(rec_date)
    date_for_name = rec_date_str.replace("-", "")

    # DOC_NO = WORK_NO_YYYYMMDD_USERID
    doc_no = f"{work_no}_{date_for_name}_{user_id}"

    folder = ensure_project_dir(survey_no)
    filename = _os.path.join(folder, doc_no + ".dat")

    # 헤더 작성
    with open(filename, "w", encoding="cp949", newline="\r\n") as f:
        f.write("[INFO]\r\n")
        f.write(f"DOC_NO={doc_no}\r\n")
        f.write(f"SURVEY_NO={survey_no}\r\n")
        f.write(f"WORK_NO={work_no}\r\n")
        f.write(f"USER_ID={user_id}\r\n")
        f.write(f"USER_NM={user_name}\r\n")
        f.write(f"REC_DATE={rec_date_str}\r\n")
        f.write(f"SEQ_NO={int(seq_no)}\r\n")
        f.write(f"START_TIME={format_korean_datetime(start_time)}\r\n")
        f.write(f"LAST_TIME={format_korean_datetime(last_time)}\r\n")
        f.write(f"STATE={int(state)}\r\n")
        f.write(f"DELETED={int(deleted)}\r\n")
        f.write("\r\n")

        # 계수 데이터 섹션 작성
        if matrices:
            for section_index, mat in enumerate(matrices):
                f.write(f"[{section_index}]\r\n")
                for r, row in enumerate(mat):
                    for c, val in enumerate(row):
                        try:
                            ival = int(val)
                        except Exception:
                            ival = 0
                        f.write(f"{r},{c}={ival}\r\n")
                f.write("\r\n")

    return filename



# ---------------- I/O ----------------
def load_data():
    if os.path.exists(DATA_PATH):
        try:
            with open(DATA_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "projects": [{
            "name": "기본 작업",
            "vehicle_set": [
                {"번호": 1, "차종명": "승용"},
                {"번호": 2, "차종명": "소형버스"},
                {"번호": 3, "차종명": "대형버스"},
                {"번호": 4, "차종명": "소형화물"},
                {"번호": 5, "차종명": "중형화물"},
                {"번호": 6, "차종명": "대형화물"},
            ],
            "hotkeys": {"좌": {}, "직": {}, "우": {}},
            "templates": []
        }],
        "surveys": []
    }

def _hotkeys_db_path():
    """NAS/공유 루트(DATA_ROOT) 기준으로 hotkeys_db.json 저장 위치를 결정합니다.

    - 기본: <DATA_ROOT>/survey/hotkeys_db.json  (survey 폴더가 실제로 존재하는 경우)
    - 대안: <DATA_ROOT>/hotkeys_db.json        (DATA_ROOT가 이미 Survey 루트인 경우 등)
    """
    try:
        root = DATA_ROOT
    except Exception:
        root = os.path.dirname(os.path.abspath(DATA_PATH))

    cand_survey = os.path.join(root, "survey")
    if os.path.isdir(cand_survey):
        folder = cand_survey
    else:
        folder = root

    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass

    return os.path.join(folder, "hotkeys_db.json")


def export_hotkeys_db(data):
    """계수프로그램이 읽을 hotkeys_db.json 생성(과업/지점/방향/차종/단축키 포함)."""
    out = {
        "version": 1,
        "exported_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        # 계수쪽에서 그대로 참조할 수 있게 전체 프로젝트/과업 구조를 그대로 포함
        "projects": (data or {}).get("projects", []) or [],
        "surveys": (data or {}).get("surveys", []) or [],
    }
    path = _hotkeys_db_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    return path


def save_data(data):
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # 계수프로그램 연동용 DB도 함께 갱신
    try:
        export_hotkeys_db(data)
    except Exception:
        pass
class WorkVehicleTab(QtWidgets.QWidget):
    def __init__(self, work):
        super().__init__(work); self.work = work
        v = QtWidgets.QVBoxLayout(self)
        self.tbl = QtWidgets.QTableWidget(0,4)
        self.tbl.setHorizontalHeaderLabels(["번호","차종구분","설명",""])
        self.tbl.setColumnHidden(3, False)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.horizontalHeader().setStretchLastSection(True)
        v.addWidget(self.tbl,1)
        hb = QtWidgets.QHBoxLayout()
        hb.addStretch(1)  # ← 왼쪽 여백, 버튼을 오른쪽으로 밀기
        self.b_add = QtWidgets.QPushButton("추가")
        self.b_edit = QtWidgets.QPushButton("수정")
        self.b_del = QtWidgets.QPushButton("삭제")
        for b in (self.b_add, self.b_edit, self.b_del):
            b.setFixedWidth(40)       # ← 버튼 가로 크기 축소
            hb.addWidget(b)
        v.addLayout(hb)
        self.b_add.clicked.connect(self.add_row)
        self.b_edit.clicked.connect(self.persist)
        self.b_del.clicked.connect(self.del_row)
        self.tbl.itemChanged.connect(lambda *_: self.persist())

    def load(self):
        p = self.work.current_project()
        self.tbl.blockSignals(True); self.tbl.setRowCount(0)
        if p:
            for row in p.get("vehicle_set",[]):
                r=self.tbl.rowCount(); self.tbl.insertRow(r)
                num = row.get("번호", r+1)
                name = row.get("차종구분", row.get("차종명",""))
                desc = row.get("설명", "")
                self.tbl.setItem(r,0,QtWidgets.QTableWidgetItem(str(num)))
                self.tbl.item(r,0).setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                self.tbl.item(r,0).setFlags(self.tbl.item(r,0).flags() & ~QtCore.Qt.ItemIsEditable)
                self.tbl.setItem(r,1,QtWidgets.QTableWidgetItem(name))
                self.tbl.item(r,1).setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(desc))
                self.tbl.item(r,2).setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
                
        self.tbl.blockSignals(False)
        try:
            self.tbl.setColumnWidth(0, 60)  # 번호 열 축소
        except Exception:
            pass

    def rows(self):
        out=[]
        for r in range(self.tbl.rowCount()):
            num_item = self.tbl.item(r,0)
            name_item= self.tbl.item(r,1)
            desc_item= self.tbl.item(r,2)
            num = num_item.text() if num_item else str(r+1)
            name= name_item.text() if name_item else ""
            desc= desc_item.text() if desc_item else ""
            if name:
                try:
                    num_val = int(num) if str(num).isdigit() else r+1
                except Exception:
                    num_val = r+1
                out.append({"번호": num_val, "차종구분": name, "차종명": name, "설명": desc})
        return out

    def persist(self):
        p = self.work.current_project()
        if not p: return
        p["vehicle_set"] = self.rows()
        save_data(self.work.data)

    def add_row(self):
        r=self.tbl.rowCount(); self.tbl.insertRow(r)
        it0=QtWidgets.QTableWidgetItem(str(r+1)); it0.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); it0.setFlags(it0.flags() & ~QtCore.Qt.ItemIsEditable); self.tbl.setItem(r,0,it0);it1=QtWidgets.QTableWidgetItem(""); it1.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,1,it1);it2=QtWidgets.QTableWidgetItem(""); it2.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,2,it2);it3=QtWidgets.QTableWidgetItem(""); it3.setFlags(it3.flags() & ~QtCore.Qt.ItemIsEditable); self.tbl.setItem(r,3,it3)
        
        self.persist()

    def del_row(self):
        r=self.tbl.currentRow()
        if r>=0:
            self.tbl.removeRow(r)
            self.persist(); self.load()



# ---------------- 차종관리: 카운터 단축키 ----------------
class WorkHotkeyTab(QtWidgets.QWidget):
    """
    (좌/직/우 제거) 전역 카운터 시트 관리 + 단축키 설정
    저장 구조:
        project["hotkey_sheets_global"] = [
            {"name":"1", "items":[{"순번","차종명","단축키"}, ...]},
            ...
        ]
    호환(옵션):
        - 첫 번째 시트를 project["hotkey_items_global"], project["hotkeys_global"] 로 미러링
        - 초기 생성은 vehicle_set 기반
    """
    def __init__(self, work):
        super().__init__(work); self.work = work
        root = QtWidgets.QHBoxLayout(self)

        # Center: 카운터 시트
        center = QtWidgets.QVBoxLayout()
        center.addWidget(QtWidgets.QLabel("카운터"))
        self.counter_list = QtWidgets.QListWidget()
        self.counter_list.setFixedWidth(200)   # ← 카운터 리스트 가로 폭 조절
        center.addWidget(self.counter_list, 1)
        cbtn = QtWidgets.QHBoxLayout()
        self.c_add = QtWidgets.QPushButton("추가")
        self.c_del = QtWidgets.QPushButton("삭제")
        # 카운터 추가/삭제 버튼을 차종유형 버튼과 비슷한 크기로 통일
        for b in (self.c_add, self.c_del):
            b.setMinimumWidth(60)
            b.setFixedHeight(24)
        cbtn.addWidget(self.c_add)
        cbtn.addWidget(self.c_del)
        cbtn.addStretch(1)
        center.addLayout(cbtn)
        root.addLayout(center, 1)

        # Right: 단축키 설정
        right = QtWidgets.QVBoxLayout()
        right.addWidget(QtWidgets.QLabel("단축키"))
        self.tbl = QtWidgets.QTableWidget(0,4)
        # 열 순서: 순번 / 차종명 / 단축키
        self.tbl.setHorizontalHeaderLabels(["순번","차종명","단축키",""])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setColumnHidden(0, False)
        # 단축키 열 가로폭 수정
        try:
            self.tbl.setColumnWidth(0, 50)
            self.tbl.setColumnWidth(1, 150)
            self.tbl.setColumnWidth(2, 50)
        except Exception:
            pass
        right.addWidget(self.tbl, 1)

        # 단축키 편집 UI + 초기화 (입력창 폭 축소 + 좌측 정렬, 초기화 버튼은 우측 끝)
        key_row = QtWidgets.QHBoxLayout()
        key_row.addWidget(QtWidgets.QLabel("단축키"))
        try:
            self.key_edit = QtWidgets.QKeySequenceEdit()
        except Exception:
            self.key_edit = QtWidgets.QLineEdit()
        # 입력창 가로 크기 축소 및 좌측 정렬: 고정 폭으로 두고 stretch 는 뒤에 배치
        self.key_edit.setFixedWidth(160)
        key_row.addWidget(self.key_edit)
        self.btn_apply_key = QtWidgets.QPushButton("적용")
        key_row.addWidget(self.btn_apply_key)
        key_row.addStretch(1)
        # 초기화 버튼을 같은 줄의 우측 끝에 배치
        self.btn_reset = QtWidgets.QPushButton("초기화")
        key_row.addWidget(self.btn_reset)
        right.addLayout(key_row)

        root.addLayout(right, 3)

        # Signals
        self.counter_list.currentRowChanged.connect(self.load_counter)
        self.counter_list.itemDoubleClicked.connect(lambda *_: self.rename_counter())
        self.tbl.currentCellChanged.connect(self.sync_key_editor)
        self.tbl.itemChanged.connect(lambda *_: self.persist_current())
        self.c_add.clicked.connect(self.add_counter)
        self.c_del.clicked.connect(self.del_counter)
        self.btn_apply_key.clicked.connect(self.apply_hotkey)
        self.btn_reset.clicked.connect(self.reset_hotkeys)

        # Init
        self.reload_counters()

    # ---------- Data helpers ----------
    def _project(self):
        return self.work.current_project()

    def _sheets(self):
        p = self._project()
        if p is None: return None
        return p.setdefault("hotkey_sheets_global", [])

    def _ensure_first(self):
        p = self._project()
        if p is None: return
        sheets = self._sheets()
        if sheets:
            return
        # initialize from vehicle_set
        vset = p.get("vehicle_set", [])
        items = [{"순번": i+1, "차종명": row.get("차종명",""), "단축키": ""} for i, row in enumerate(vset)]
        sheets.append({"name":"1","items":items})
        save_data(self.work.data)

    # ---------- Loading ----------
    def load_dir(self, *_):
        self.reload_counters()

    def reload_counters(self, preferred_index=None):
        self._ensure_first()
        self.counter_list.blockSignals(True)
        self.counter_list.clear()
        sheets = self._sheets() or []
        for s in sheets:
            self.counter_list.addItem(str(s.get("name","1")))
        self.counter_list.blockSignals(False)
        if self.counter_list.count():
            # 기본은 첫 번째, preferred_index가 있으면 그 위치를 우선 선택
            if preferred_index is None:
                preferred_index = 0
            preferred_index = max(0, min(preferred_index, self.counter_list.count() - 1))
            self.counter_list.setCurrentRow(preferred_index)
        else:
            # If no project yet, do nothing
            pass

    def load_counter(self, *_):
        self.tbl.blockSignals(True)
        self.tbl.setRowCount(0)
        sheets = self._sheets() or []
        i = self.counter_list.currentRow()
        if 0 <= i < len(sheets):
            for rec in sheets[i].get("items", []):
                r = self.tbl.rowCount(); self.tbl.insertRow(r)
                # 0열: 순번, 1열: 차종명, 2열: 단축키
                self.tbl.setItem(r,0,QtWidgets.QTableWidgetItem(str(rec.get("순번", r+1))))
                self.tbl.setItem(r,1,QtWidgets.QTableWidgetItem(rec.get("차종명","")))
                self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(rec.get("단축키","")))
                for c in (0,1,2):
                    it = self.tbl.item(r,c)
                    if it: it.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)

        self.tbl.blockSignals(False)
        try:
            self.tbl.setColumnWidth(0, 38)  # 순번 열 약간 축소
        except Exception:
            pass
        try:
            self.tbl.setColumnWidth(1, 80)  # 차종명 열 기본 폭
        except Exception:
            pass
        self.sync_key_editor(self.tbl.currentRow(), 0, -1, -1)
    # ---------- Persist ----------
    def persist_current(self):
        p = self._project()
        if p is None: return
        sheets = self._sheets()
        i = self.counter_list.currentRow()
        if sheets is None or not (0 <= i < len(sheets)): return
        items = []
        for r in range(self.tbl.rowCount()):
            # 0열: 순번(표시용), 1열: 차종명, 2열: 단축키
            name_item = self.tbl.item(r,1)
            key_item  = self.tbl.item(r,2)
            name = name_item.text().strip() if name_item else ""
            key  = key_item.text().strip() if key_item else ""
            if name:
                items.append({"순번": r+1, "차종명": name, "단축키": key})
        sheets[i]["items"] = items
        # mirrors
        if i == 0:
            p["hotkey_items_global"] = items
            p["hotkeys_global"] = {rec["차종명"]: rec["단축키"] for rec in items}
        save_data(self.work.data)

    # ---------- Counter ops ----------
    def add_counter(self):
        p = self._project()
        if p is None: return
        sheets = self._sheets()
        next_no = 1 + len(sheets or [])
        vset = p.get("vehicle_set", [])
        items = [{"순번": i+1, "차종명": row.get("차종명",""), "단축키": ""} for i, row in enumerate(vset)]
        sheets.append({"name": str(next_no), "items": items})
        save_data(self.work.data)
        self.reload_counters()
        self.counter_list.setCurrentRow(self.counter_list.count()-1)

    def rename_counter(self):
        sheets = self._sheets() or []
        i = self.counter_list.currentRow()
        if not (0 <= i < len(sheets)): return
        cur = sheets[i].get("name","")
        text, ok = QtWidgets.QInputDialog.getText(self, "카운터 이름 수정", "이름:", text=str(cur))
        if ok and text.strip():
            sheets[i]["name"] = text.strip()
            save_data(self.work.data)
            self.reload_counters()
            self.counter_list.setCurrentRow(i)

    def del_counter(self):
        sheets = self._sheets() or []
        i = self.counter_list.currentRow()
        if not (0 <= i < len(sheets)): return
        if QtWidgets.QMessageBox.question(self,"삭제 확인","현재 카운터를 삭제할까요?")==QtWidgets.QMessageBox.Yes:
            sheets.pop(i)
            # re-number default names 1..n if numeric
            for idx, s in enumerate(sheets, 1):
                if str(s.get("name","")).isdigit():
                    s["name"] = str(idx)
            save_data(self.work.data)
            # 삭제된 항목 바로 위(또는 동일 위치)를 선택하도록 요청
            preferred = i - 1
            self.reload_counters(preferred_index=preferred)

    # ---------- Key editor ----------
    def sync_key_editor(self, cur_row, cur_col, prev_row, prev_col):
        try:
            key = self.tbl.item(cur_row, 2).text() if cur_row is not None and cur_row >= 0 else ""
        except Exception:
            key = ""
        if hasattr(self, "key_edit") and hasattr(self.key_edit, "setKeySequence"):
            from PyQt5 import QtGui
            self.key_edit.setKeySequence(QtGui.QKeySequence(key))
        else:
            self.key_edit.setText(key)

    def apply_hotkey(self):
        r = self.tbl.currentRow()
        if r < 0:
            QtWidgets.QMessageBox.information(self,"알림","행을 먼저 선택하세요.")
            return
        if hasattr(self.key_edit, "keySequence"):
            key = self.key_edit.keySequence().toString()
        else:
            key = self.key_edit.text()
        self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(key))
        self.persist_current()

    def reset_hotkeys(self):
        if QtWidgets.QMessageBox.question(self,"초기화","현재 카운터의 단축키를 모두 비울까요?")!=QtWidgets.QMessageBox.Yes:
            return
        for r in range(self.tbl.rowCount()):
            self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(""))
        
        self.persist_current()

# ---------------- 차종관리: 템플릿 ----------------
class WorkTemplateTab(QtWidgets.QWidget):
    def __init__(self, work):
        super().__init__(work)
        self.work = work

        # 전체 레이아웃: 좌측 템플릿 / 우측 구성
        root = QtWidgets.QGridLayout(self)

        # ----- 왼쪽: 입력그룹 템플릿 리스트 -----
        root.addWidget(QtWidgets.QLabel("입력그룹 템플릿"), 0, 0)
        self.lst = QtWidgets.QListWidget()
        # 과업명 리스트 폭 300px 고정
        self.lst.setFixedWidth(200)
        self.lst.setMinimumWidth(220)
        self.lst.setMaximumWidth(220)
        root.addWidget(self.lst, 1, 0, 1, 1)

        tpl_btns = QtWidgets.QHBoxLayout()
        self.b_add = QtWidgets.QPushButton("추가")
        self.b_del = QtWidgets.QPushButton("삭제")
        for b in (self.b_add, self.b_del):
            tpl_btns.addWidget(b)
        root.addLayout(tpl_btns, 2, 0)

        # ----- 오른쪽: 방향번호 / 입력그룹(탭) / 카운터 -----
        # 오른쪽 전체 패널을 wrapper로 감싸 상단 정렬
        right_wrapper = QtWidgets.QWidget()
        right_wrapper_layout = QtWidgets.QVBoxLayout(right_wrapper)
        right_wrapper_layout.setContentsMargins(0, 0, 0, 0)
        right_wrapper_layout.setAlignment(QtCore.Qt.AlignTop)

        right = QtWidgets.QGridLayout()
        # 방향번호 / 화살표 / 입력그룹(탭)+카운터 사이 가로 여백 최소화 및 우측 영역 확장
        right.setHorizontalSpacing(2)
        right.setColumnStretch(0, 0)  # 방향번호 열
        right.setColumnStretch(1, 0)  # 화살표 열
        right.setColumnStretch(2, 1)  # 입력그룹(탭)+카운터 열을 가장 넓게
        right_wrapper_layout.addLayout(right)
        root.addWidget(right_wrapper, 0, 1, 3, 1)

        # 0행: 방향수 + 생성/자동생성 (좌측 정렬, 버튼 작게)
        form = QtWidgets.QHBoxLayout()
        # "입력 그룹 구성" 라벨 제거, 방향수 행을 좌측 정렬
        form.addWidget(QtWidgets.QLabel("방향수"))
        self.spin = QtWidgets.QSpinBox()
        self.spin.setRange(1, 40)
        self.spin.setValue(2)
        # 방향수 스핀박스 가로 크기 2배 확대
        try:
            w = self.spin.sizeHint().width()
            self.spin.setFixedWidth(int(w * 2))
        except Exception:
            pass
        form.addWidget(self.spin)
        self.b_make = QtWidgets.QPushButton("생성")
        self.b_auto = QtWidgets.QPushButton("자동생성")
        self.b_make.setFixedWidth(60)
        self.b_auto.setFixedWidth(80)
        form.addWidget(self.b_make)
        form.addWidget(self.b_auto)
        form.addStretch(1)
        right.addLayout(form, 0, 0, 1, 3)

        # 1행: 방향번호 / 입력 그룹(탭) 라벨
        lbl_dir = QtWidgets.QLabel("방향 번호")
        lbl_tabs = QtWidgets.QLabel("입력 그룹(탭)")
        right.addWidget(lbl_dir, 1, 0)
        right.addWidget(lbl_tabs, 1, 2)

        # 2~3행: 방향번호 리스트(세로 전체), 화살표, 입력그룹(탭) + 카운터
        # 방향번호 리스트: row 2~3 전체 사용 (카운터 영역까지 존재)
        self.dir = QtWidgets.QListWidget()
        self.dir.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        # 방향번호 리스트: 가로 폭을 기존보다 2배 정도 넓게
        self.dir.setFixedWidth(100)
        try:
            self.dir.setMinimumHeight(250)
        except Exception:
            pass
        right.addWidget(self.dir, 2, 0, 2, 1)

        # 화살표 버튼: 방향번호와 동일 높이(rowSpan=2)
        vbtn = QtWidgets.QVBoxLayout()
        # 위/아래 스트레치로 버튼을 세로 중앙 배치
        vbtn.addStretch(1)
        self.to = QtWidgets.QPushButton("▶")
        self.fr = QtWidgets.QPushButton("◀")
        vbtn.addWidget(self.to)
        vbtn.addWidget(self.fr)
        vbtn.addStretch(1)
        # 화살표 버튼을 열의 오른쪽으로 정렬하여 입력그룹 탭 쪽에 가깝게 배치
        right.addLayout(vbtn, 2, 1, 2, 1, alignment=QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # 입력 그룹(탭): 위쪽 영역(row 2)
        self.tabs = QtWidgets.QListWidget()
        # 가로 방향으로는 레이아웃 공간을 꽉 채우도록 설정
        self.tabs.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        # 탭 영역 세로 크기 조정 (너무 커지지 않도록 상한 설정)
        try:
            self.tabs.setMaximumHeight(100)
            self.tabs.setMinimumHeight(100)
        except Exception:
            pass
        right.addWidget(self.tabs, 2, 2, 1, 1)


        # 카운터: 입력 그룹(탭) 아래(row 3, col 2)
        counter = QtWidgets.QVBoxLayout()
        lbl_counter = QtWidgets.QLabel("카운터")
        counter.addWidget(lbl_counter)
        self.tbl = QtWidgets.QTableWidget(0, 3)
        # 가로는 확장, 세로는 고정 높이 사용
        self.tbl.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.tbl.setFixedHeight(250)   # 세로 크기 조정 (입력그룹 탭을 줄이고 카운터는 크게)
        self.tbl.setHorizontalHeaderLabels(["카운터", "방향", "표시명"])
        self.tbl.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignHCenter)
        self.tbl.verticalHeader().setVisible(False)
        # 카운터 템플릿 테이블은 카운터/방향은 잠그고, 표시명만 수정 가능하도록 설정
        self.tbl.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
        )
        # 카운터 템플릿 테이블 폰트 크기 축소 (입력그룹 시트와 유사하게)
        try:
            f = self.tbl.font()
            if f.pointSize() > 0:
                f.setPointSize(max(8, f.pointSize() - 2))
                self.tbl.setFont(f)
                self.tbl.horizontalHeader().setFont(f)
        except Exception:
            pass
        # 행 높이를 약간 줄여서 방향 열 세로 크기를 10 정도 축소
        try:
            vh = self.tbl.verticalHeader()
            h = vh.defaultSectionSize()
            if h > 12:
                vh.setDefaultSectionSize(max(10, h - 10))
        except Exception:
            pass
        # 카운터/번호/표시명 열 구성:
        # - 0열(카운터), 1열(방향)은 가운데 정렬
        # - 2열(표시명)은 가로 사이즈를 넓게 사용
        try:
            for r in range(self.tbl.rowCount()):
                it_counter = self.tbl.item(r, 0)
                if it_counter:
                    it_counter.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                it_dir = self.tbl.item(r, 1)
                if it_dir:
                    it_dir.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        except Exception:
            pass
        try:
            base_w = self.tbl.columnWidth(2) or 100
            self.tbl.setColumnWidth(2, int(base_w * 2))
        except Exception:
            pass
        # 방향 열(1열)의 가로 사이즈를 5px 정도 축소
        try:
            w_dir = self.tbl.columnWidth(1) or 40
            self.tbl.setColumnWidth(1, max(10, w_dir - 5))
        except Exception:
            pass
        # 열 가로크기를 각각 1/2로 축소
        try:
            for c in range(3):
                w = self.tbl.columnWidth(c)
                if w > 0:
                    self.tbl.setColumnWidth(c, max(20, int(w * 0.5)))
        except Exception:
            pass
        # 카운터 열(0열) 가로 크기를 110으로 설정
        try:
            self.tbl.setColumnWidth(0, 110)
        except Exception:
            pass
        # 전체 폭을 살짝 줄여서 공간 확보
        try:
            cw = self.tbl.sizeHint().width()
            if cw > 20:
                pass  # fixed width 제거
        except Exception:
            pass
        counter.addWidget(self.tbl)

        # 카운터 시트: 선택된 행을 삭제할 수 있는 버튼 추가
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_counter_del = QtWidgets.QPushButton("선택 삭제")
        btn_row.addWidget(self.btn_counter_del)
        counter.addLayout(btn_row)

        self.tbl.itemChanged.connect(lambda *_: self.persist_current())

        # 카운터 레이아웃을 바로 아래 행에 배치하고, 행 비율을 조정하여 상단으로 끌어올림
        right.addLayout(counter, 3, 2, 1, 1)

        # 위/아래 영역 비율 (입력 그룹 영역을 줄이고 카운터 영역을 키움)
        right.setRowStretch(2, 1)   # 탭 영역
        right.setRowStretch(3, 3)   # 카운터 영역

        # 시그널 연결
        self.b_make.clicked.connect(self.make_dirs)
        self.b_auto.clicked.connect(self.make_dirs)
        self.to.clicked.connect(self.add_group)
        self.fr.clicked.connect(self.remove_group)
        self.b_add.clicked.connect(self.add_tpl)
        self.b_del.clicked.connect(self.del_tpl)
        self.lst.currentRowChanged.connect(self.load_tpl)
        # 카운터 시트 - 선택된 행 삭제 버튼
        self.btn_counter_del.clicked.connect(self.del_counter_row)
        # 템플릿 항목 더블클릭 시 이름 수정 팝업
        self.lst.itemDoubleClicked.connect(self.ren_tpl)
        # 입력 그룹(탭) 더블클릭 시 단축키 설정 팝업
        self.tabs.itemDoubleClicked.connect(self._on_group_tab_double_clicked)
    def current_project(self): return self.work.current_project()

    def load(self, preferred_index=None):
        p=self.current_project(); self.lst.blockSignals(True); self.lst.clear()
        if p:
            for t in p.get("templates",[]): self.lst.addItem(t.get("name","(무제)"))
        self.lst.blockSignals(False)
        if self.lst.count():
            if preferred_index is None:
                preferred_index = 0
            preferred_index = max(0, min(preferred_index, self.lst.count() - 1))
            self.lst.setCurrentRow(preferred_index)
        else:
            self.clear_right()

    def clear_right(self):
        self.dir.clear(); self.tabs.clear(); self.tbl.setRowCount(0)

    def _rebuild_dir_list(self):
        """현재 방향수와 입력그룹(탭)에 맞추어 남은 방향번호 리스트를 다시 생성."""
        import re as _re
        self.dir.clear()
        try:
            total = self.spin.value()
        except Exception:
            total = 0
        all_dirs = set(range(1, total + 1))
        used = set()
        for i in range(self.tabs.count()):
            label = self.tabs.item(i).text()
            for tok in _re.findall(r"\d+", str(label)):
                try:
                    used.add(int(tok))
                except Exception:
                    pass
        remain = sorted(all_dirs - used)
        for n in remain:
            self.dir.addItem(str(n))

    def make_dirs(self):
        # 방향번호 리스트를 새로 생성하되, 이미 입력그룹(탭)에 포함된 방향은 제외
        self._rebuild_dir_list()

    def add_group(self):
        sels=[it.text() for it in self.dir.selectedItems()]
        if not sels: 
            return
        try:
            nums=sorted(int(x) for x in sels)
            label="-".join(str(n) for n in nums)
        except Exception:
            label="-".join(sels)
        self.tabs.addItem(label)
        # 선택한 방향번호는 사용 처리 → 남은 방향만 리스트에 표시
        self._rebuild_dir_list()
        self.persist_current()

    def remove_group(self):
        # 선택된 그룹(탭)을 제거한 뒤, 남은 그룹 정보를 기준으로 방향번호 리스트를 재계산
        for it in self.tabs.selectedItems():
            self.tabs.takeItem(self.tabs.row(it))
        self._rebuild_dir_list()
        self.persist_current()


    def _on_group_tab_double_clicked(self, item):
        """차종관리 > 카운터 템플릿 관리: 입력그룹(탭) 단축키 시트를 설정."""
        if item is None:
            return
        label = item.text() or ""
        import re as _re
        dirs = []
        for part in _re.split(r"[^0-9]+", label):
            if part.isdigit():
                try:
                    dirs.append(int(part))
                except Exception:
                    pass
        dirs = sorted(set(dirs))
        if not dirs:
            QtWidgets.QMessageBox.information(self, "알림", "이 탭에서 방향번호를 찾을 수 없습니다.")
            return

        p = self.current_project()
        if p is None:
            QtWidgets.QMessageBox.information(self, "알림", "먼저 차종 유형을 선택하세요.")
            return

        sheets = p.get("hotkey_sheets_global", []) or []
        sheet_names = [str(s.get("name", "") or "") for s in sheets]
        if not sheet_names:
            QtWidgets.QMessageBox.information(self, "알림", "현재 차종 유형에 등록된 카운터 단축키 시트가 없습니다.")
            return

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("입력그룹 단축키 설정 (템플릿)")
        vbox = QtWidgets.QVBoxLayout(dlg)

        # 차종유형 라벨 (템플릿은 현재 프로젝트 고정)
        form_top = QtWidgets.QFormLayout()
        form_top.addRow("차종유형:", QtWidgets.QLabel(p.get("name", "") or "(이름 없음)"))
        vbox.addLayout(form_top)

        # 현재 카운터 테이블에서 방향별로 이미 지정된 시트 이름을 수집
        current_by_dir = {}
        for row in range(self.tbl.rowCount()):
            d_item = self.tbl.item(row, 1)
            n_item = self.tbl.item(row, 0)
            if not d_item:
                continue
            d_text = d_item.text().strip()
            if not d_text:
                continue
            if n_item:
                current_by_dir.setdefault(d_text, n_item.text().strip())

        # 방향별 콤보박스
        dir_layout = QtWidgets.QFormLayout()
        dir_combos = {}

        def _refresh_preview_all():
            # 전체 미리보기 영역 초기화
            while preview_layout.count():
                item = preview_layout.takeAt(0)
                w = item.widget()
                if w is not None:
                    w.deleteLater()

            # 각 방향별로 그룹박스를 만들어 선택된 시트의 단축키를 표시
            for d in dirs:
                combo = dir_combos.get(d)
                group_box = QtWidgets.QGroupBox(f"{d}번 방향")
                gb_layout = QtWidgets.QVBoxLayout(group_box)
                gb_layout.setContentsMargins(4, 4, 4, 4)
                gb_layout.setSpacing(2)

                if combo is None:
                    gb_layout.addWidget(QtWidgets.QLabel("방향 콤보박스를 찾을 수 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                sheet_name = combo.currentText().strip()
                if not sheet_name:
                    gb_layout.addWidget(QtWidgets.QLabel("선택된 단축키 시트가 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                target = None
                for s in sheets:
                    if str(s.get("name", "") or "") == sheet_name:
                        target = s
                        break

                if target is None:
                    gb_layout.addWidget(QtWidgets.QLabel(f"'{sheet_name}' 시트를 찾을 수 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                items = target.get("items", []) or []
                if not items:
                    gb_layout.addWidget(QtWidgets.QLabel("등록된 단축키가 없습니다."))
                else:
                    for rec in items:
                        label_txt = str(rec.get("차종명", ""))
                        key_txt = str(rec.get("단축키", ""))
                        row = QtWidgets.QHBoxLayout()
                        lab = QtWidgets.QLabel(f"{label_txt} 키")
                        edit = QtWidgets.QLineEdit(key_txt)
                        edit.setReadOnly(True)
                        edit.setMaximumWidth(80)
                        row.addWidget(lab)
                        row.addWidget(edit)
                        gb_layout.addLayout(row)

                preview_layout.addWidget(group_box)

        def _on_dir_combo_changed(_idx, d):
            # 콤보박스 변경 시 전체 방향 미리보기를 다시 그린다.
            _refresh_preview_all()

        for d in dirs:
            combo = QtWidgets.QComboBox(dlg)
            combo.addItem("")  # 공백 선택
            for n in sheet_names:
                combo.addItem(n)
            key = str(d)
            base = current_by_dir.get(key, "")
            if base and base in sheet_names:
                combo.setCurrentIndex(sheet_names.index(base) + 1)
            dir_combos[d] = combo
            combo.currentIndexChanged.connect(lambda idx, dd=d: _on_dir_combo_changed(idx, dd))
            dir_layout.addRow(f"{d}번 방향", combo)

        vbox.addLayout(dir_layout)

        # 미리보기 영역
        # 그룹박스 사용 시 제목과 내용이 겹쳐 보이는 현상이 있어
        # 제목 라벨 + 스크롤 영역 조합으로 변경.
        # 미리보기 영역: QGroupBox + QScrollArea 조합으로 구성하여
        # 제목과 내용이 자연스럽게 배치되고, 항목이 많을 경우 스크롤로 확인할 수 있도록 한다.
        preview_group = QtWidgets.QGroupBox("선택된 단축키 시트 미리보기")
        preview_group.setStyleSheet(
            "QGroupBox { margin-top: 12px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 8px; }"
        )
        group_layout = QtWidgets.QVBoxLayout(preview_group)
        group_layout.setContentsMargins(4, 4, 4, 4)
        group_layout.setSpacing(2)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        preview_container = QtWidgets.QWidget()
        preview_layout = QtWidgets.QVBoxLayout(preview_container)
        preview_layout.setAlignment(QtCore.Qt.AlignTop)
        preview_layout.setContentsMargins(4, 4, 4, 4)
        preview_layout.setSpacing(2)
        scroll.setWidget(preview_container)
        group_layout.addWidget(scroll)

        vbox.addWidget(preview_group)

        # 초기 미리보기: 모든 방향 기준
        _refresh_preview_all()


        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok = QtWidgets.QPushButton("OK")
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        vbox.addLayout(btn_row)

        def _on_ok():
            # 선택된 시트를 카운터 테이블에 반영
            for d in dirs:
                sheet_name = dir_combos[d].currentText().strip()
                if not sheet_name:
                    continue
                dir_str = str(d)
                found = False
                for row in range(self.tbl.rowCount()):
                    d_item = self.tbl.item(row, 1)
                    if d_item and d_item.text().strip() == dir_str:
                        name_item = QtWidgets.QTableWidgetItem(sheet_name)
                        name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                        self.tbl.setItem(row, 0, name_item)
                        found = True
                # 해당 방향번호에 대한 카운터 행이 없으면 새로 추가
                if not found:
                    r = self.tbl.rowCount()
                    self.tbl.insertRow(r)
                    name_item = QtWidgets.QTableWidgetItem(sheet_name)
                    name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.tbl.setItem(r, 0, name_item)
                    dir_item = QtWidgets.QTableWidgetItem(dir_str)
                    dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
                    self.tbl.setItem(r, 1, dir_item)
                    self.tbl.setItem(r, 2, QtWidgets.QTableWidgetItem(""))
            self.persist_current()
            QtWidgets.QMessageBox.information(self, "완료", "입력그룹 단축키 설정이 완료되었습니다.")
            dlg.accept()

        btn_ok.clicked.connect(_on_ok)
        btn_cancel.clicked.connect(dlg.reject)

        dlg.exec_()

    def add_counter(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        # 0열: 카운터 이름(숨김 처리되지만 내부적으로 유지, 편집 불가)
        name_item = QtWidgets.QTableWidgetItem(f"카운터{r+1}")
        name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 0, name_item)
        # 1열: 번호(가운데 정렬, 편집 불가)
        dir_item = QtWidgets.QTableWidgetItem("1")
        dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        # 방향(번호) 열은 읽기 전용
        dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 1, dir_item)
        # 2열: 표시명(사용자 수정 가능)
        label_item = QtWidgets.QTableWidgetItem("")
        self.tbl.setItem(r, 2, label_item)
        self.persist_current()


    def del_counter(self):
        r=self.tbl.currentRow()
        if r>=0: 
            self.tbl.removeRow(r)
            self.persist_current()

    def add_tpl(self):
        name, ok = QtWidgets.QInputDialog.getText(self,"템플릿 추가","이름:")
        if not ok or not name.strip(): return
        p=self.current_project(); 
        p.setdefault("templates",[]).append({"name":name.strip(),"dirs":[],"counters":[]})
        save_data(self.work.data); self.load()

    def ren_tpl(self, *_):
        i=self.lst.currentRow(); 
        if i<0: return
        it=self.lst.item(i); name, ok = QtWidgets.QInputDialog.getText(self,"이름 변경","이름:",text=it.text())
        if ok and name.strip():
            p=self.current_project(); p["templates"][i]["name"]=name.strip(); save_data(self.work.data); self.load(); self.lst.setCurrentRow(i)

    def del_tpl(self):
        i=self.lst.currentRow(); 
        if i<0: return
        if QtWidgets.QMessageBox.question(self,"삭제 확인","삭제할까요?")==QtWidgets.QMessageBox.Yes:
            p=self.current_project(); p["templates"].pop(i); save_data(self.work.data); self.load(preferred_index=i-1)


    def persist_current(self):
        """현재 선택된 템플릿의 입력그룹(탭)과 카운터 구성을 데이터에 저장."""
        p = self.current_project()
        i = self.lst.currentRow()
        if not p or i < 0:
            return
        templates = p.setdefault("templates", [])
        if not (0 <= i < len(templates)):
            return
        t = templates[i]
        # 입력그룹(탭)
        dirs = []
        for idx in range(self.tabs.count()):
            text = self.tabs.item(idx).text()
            if text:
                dirs.append(str(text))
        # 카운터 테이블
        counters = []
        for r in range(self.tbl.rowCount()):
            name_item = self.tbl.item(r, 0)
            dir_item  = self.tbl.item(r, 1)
            label_item= self.tbl.item(r, 2)
            name  = name_item.text().strip() if name_item else ""
            dtext = dir_item.text().strip() if dir_item else ""
            label = label_item.text().strip() if label_item else ""
            if not (name or dtext or label):
                continue
            try:
                dval = int(dtext) if dtext else None
            except Exception:
                dval = None
            entry = {"name": name, "dir": dval, "label": label}
            counters.append(entry)
        t["dirs"] = dirs
        t["counters"] = counters
        save_data(self.work.data)

    def del_counter_row(self):
        """카운터 시트에서 선택된 행(들)을 삭제한다."""
        rows = sorted({i.row() for i in self.tbl.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            if 0 <= r < self.tbl.rowCount():
                self.tbl.removeRow(r)
        self.persist_current()
    def load_tpl(self,*_):
        p=self.current_project(); i=self.lst.currentRow(); self.clear_right()
        if not p or i<0: 
            return
        t=p.get("templates",[])[i]
        for g in t.get("dirs",[]):
            self.tabs.addItem(str(g))
        for c in t.get("counters", []):
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            # 0열: 카운터 이름(숨김 컬럼, 편집 불가)
            name_item = QtWidgets.QTableWidgetItem(c.get("name", ""))
            name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 0, name_item)
            # 1열: 번호(가운데 정렬, 편집 불가)
            dir_text = str(c.get("dir", "") if c.get("dir", "") is not None else "")
            dir_item = QtWidgets.QTableWidgetItem(dir_text)
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 1, dir_item)
            # 2열: 표시명 (사용자 수정 가능)
            label_item = QtWidgets.QTableWidgetItem(c.get("label", ""))
            self.tbl.setItem(r, 2, label_item)
        # 템플릿을 불러온 후, 남은 방향번호 리스트를 다시 계산
        self._rebuild_dir_list()

# ---------------- 차종관리 페이지 ----------------
class WorkManagerPage(QtWidgets.QWidget):
    changed = QtCore.pyqtSignal()
    def __init__(self, data, parent=None):
        super().__init__(parent); self.data = data
        h = QtWidgets.QHBoxLayout(self)
        left = QtWidgets.QVBoxLayout()
        self.lst = QtWidgets.QListWidget()
        left.addWidget(QtWidgets.QLabel("차종 유형"))
        self.lst.setFixedWidth(250)
        self.lst.setMinimumWidth(220)
        self.lst.setMaximumWidth(220)
        left.addWidget(self.lst, 1)
        try: self.lst.setMaximumWidth(200)
        except Exception: pass
        hb = QtWidgets.QHBoxLayout(); 
        self.b_add = QtWidgets.QPushButton("추가"); self.b_del=QtWidgets.QPushButton("삭제")
        for b in (self.b_add,self.b_del): hb.addWidget(b)
        left.addLayout(hb); h.addLayout(left,1)

        self.tabs = QtWidgets.QTabWidget()
        self.tab_vehicle = WorkVehicleTab(self)
        self.tab_hotkey  = WorkHotkeyTab(self)
        self.tab_tpl     = WorkTemplateTab(self)
        self.tabs.addTab(self.tab_vehicle,"차종 설정")
        self.tabs.addTab(self.tab_hotkey,"카운터 단축키 설정")
        self.tabs.addTab(self.tab_tpl,"카운터 템플릿 관리")
        h.addWidget(self.tabs,3)

        self.b_add.clicked.connect(self.add_project)
        self.b_del.clicked.connect(self.del_project)
        self.lst.currentRowChanged.connect(self._load_right)
        self.lst.itemDoubleClicked.connect(self.ren_project)
        self.reload()

    def projects(self): return self.data["projects"]
    def current_project(self):
        i=self.lst.currentRow(); 
        return self.projects()[i] if 0<=i<len(self.projects()) else None

    def reload(self, preferred_index=None):
        self.lst.clear(); 
        for p in self.projects(): self.lst.addItem(p["name"])
        if self.lst.count():
            if preferred_index is None:
                preferred_index = 0
            preferred_index = max(0, min(preferred_index, self.lst.count()-1))
            self.lst.setCurrentRow(preferred_index)

    def _load_right(self,*_):
        self.tab_vehicle.load(); self.tab_hotkey.load_dir(); self.tab_tpl.load()
        self.changed.emit()

    def add_project(self):
        name, ok = QtWidgets.QInputDialog.getText(self,"새 작업","이름:")
        if ok and name.strip():
            # 새 항목을 현재 순서 맨 아래에 추가하고, 추가된 위치를 그대로 선택
            self.projects().append({
                "name": name.strip(),
                "vehicle_set": [],
                "hotkeys": {"좌": {}, "직": {}, "우": {}},
                "templates": []
            })
            save_data(self.data)
            preferred = len(self.projects()) - 1  # 방금 추가된 인덱스
            self.reload(preferred_index=preferred)

    def ren_project(self):
        it = self.lst.currentItem()
        if not it:
            return
        cur_index = self.lst.currentRow()
        name, ok = QtWidgets.QInputDialog.getText(self,"수정","이름:",text=it.text())
        if ok and name.strip():
            self.projects()[cur_index]["name"] = name.strip()
            save_data(self.data)
            # 이름만 바뀌고 위치는 그대로 유지하도록 현재 인덱스를 다시 선택
            self.reload(preferred_index=cur_index)

    def del_project(self):
        i=self.lst.currentRow(); 
        if i<0: return
        if QtWidgets.QMessageBox.question(self,"삭제 확인","삭제할까요?")==QtWidgets.QMessageBox.Yes:
            self.projects().pop(i); save_data(self.data); 
            # keep selection at the position of the deleted item (or last item if deleted last)
            preferred = max(0, min(i, len(self.projects())-1))
            self.reload(preferred_index=preferred)

# ---------------- 조사관리 탭들 ----------------
def next_sn():
    """
    조사번호: SN_YYYYMMDDhhmmssff
    예) 2025년 12월 5일 14시 25분 23초 15밀리초 -> SN_2025120514252315
    """
    now = dt.datetime.now()
    ms2 = int(now.microsecond / 10000)  # 0~99 (약 10ms 단위)
    return "SN_" + now.strftime("%Y%m%d%H%M%S") + f"{ms2:02d}"


class SurveyInfoTab(QtWidgets.QWidget):
    def __init__(self, page):
        super().__init__(page)
        f = QtWidgets.QFormLayout(self)
        self.cb_purpose = QtWidgets.QComboBox(); self.cb_purpose.addItems(["일반 조사용(모든작업자 노출)","관리자용(관리자만 노출)"])
        self.cb_state = QtWidgets.QComboBox(); self.cb_state.addItems(["대기","진행","완료"])
        self.ed_name = QtWidgets.QLineEdit()
        self.ed_sn   = QtWidgets.QLineEdit()
        # 조사번호는 자동 생성되므로 사용자 편집 불가
        self.ed_sn.setReadOnly(True)
        self.ed_reg  = QtWidgets.QDateEdit(QtCore.QDate.currentDate()); self.ed_reg.setCalendarPopup(True)
        self.ed_client = QtWidgets.QLineEdit()
        self.ed_p1  = QtWidgets.QDateEdit(QtCore.QDate.currentDate()); self.ed_p1.setCalendarPopup(True)
        self.ed_p2  = QtWidgets.QDateEdit(QtCore.QDate.currentDate()); self.ed_p2.setCalendarPopup(True)
        # 조사기간 날짜 콤보박스(시작/종료) 가로 크기 2배로 확대
        try:
            w = self.ed_p1.sizeHint().width()
            self.ed_p1.setFixedWidth(int(w * 2))
            self.ed_p2.setFixedWidth(int(w * 2))
        except Exception:
            pass
        self.ed_desc= QtWidgets.QPlainTextEdit()
        f.addRow("용도", self.cb_purpose)
        f.addRow("진행상태", self.cb_state)
        f.addRow("조사명", self.ed_name)
        f.addRow("조사번호", self.ed_sn)
        f.addRow("등록일", self.ed_reg)
        f.addRow("발주처", self.ed_client)
        h=QtWidgets.QHBoxLayout(); h.addWidget(self.ed_p1); h.addWidget(QtWidgets.QLabel(" ~ ")); h.addWidget(self.ed_p2); h.addStretch(1); f.addRow("조사기간", h)
        f.addRow("설명", self.ed_desc)

        # 진행상태 색상 초기 설정
        try:
            self.cb_state.currentTextChanged.connect(self.update_state_color)
        except Exception:
            pass
        self.update_state_color(self.cb_state.currentText())

    def get(self):
        return {
            "purpose": self.cb_purpose.currentText(),
            "state": self.cb_state.currentText(),
            "name": self.ed_name.text().strip(),
            "sn": self.ed_sn.text().strip() or next_sn(),
            "reg_date": self.ed_reg.date().toString("yyyy-MM-dd"),
            "client": self.ed_client.text().strip(),
            "period": [self.ed_p1.date().toString("yyyy-MM-dd"), self.ed_p2.date().toString("yyyy-MM-dd")],
            "desc": self.ed_desc.toPlainText()
        }

    def set(self, d):
        if not isinstance(d, dict): d={}
        self.cb_purpose.setCurrentText(d.get("purpose","일반 조사용(모든작업자 노출)"))
        self.cb_state.setCurrentText(d.get("state","대기"))
        self.update_state_color(self.cb_state.currentText())
        self.ed_name.setText(d.get("name",""))
        self.ed_sn.setText(d.get("sn",""))
        qd=QtCore.QDate.fromString(d.get("reg_date",""),"yyyy-MM-dd")
        if qd.isValid(): self.ed_reg.setDate(qd)
        self.ed_client.setText(d.get("client",""))
        per=d.get("period")
        if isinstance(per,(list,tuple)) and len(per)==2:
            a=QtCore.QDate.fromString(str(per[0]),"yyyy-MM-dd")
            b=QtCore.QDate.fromString(str(per[1]),"yyyy-MM-dd")
            if a.isValid(): self.ed_p1.setDate(a)
            if b.isValid(): self.ed_p2.setDate(b)
        self.ed_desc.setPlainText(d.get("desc",""))


    def update_state_color(self, text):
        """진행상태 콤보박스 글자 색상 변경"""
        # 대기: 검정, 진행: 파랑, 완료: 연회색
        color_map = {"대기": "black", "진행": "blue", "완료": "lightgray"}
        color = color_map.get(text, "black")
        # QComboBox 전체 텍스트 색 변경
        self.cb_state.setStyleSheet("QComboBox { color: %s; }" % color)

class SurveyTimeTab(QtWidgets.QWidget):
    def __init__(self, page):
        super().__init__(page)
        v=QtWidgets.QVBoxLayout(self)
        top=QtWidgets.QHBoxLayout()
        self.t1=QtWidgets.QTimeEdit(QtCore.QTime(7,0)); self.t1.setDisplayFormat("HH:mm")
        self.t2=QtWidgets.QTimeEdit(QtCore.QTime(9,0)); self.t2.setDisplayFormat("HH:mm")
        self.step=QtWidgets.QSpinBox(); self.step.setRange(1,60); self.step.setValue(15)

        # 시간/종료/간격 입력 위젯 가로 크기 2배로 조정
        try:
            w = self.t1.sizeHint().width()
            self.t1.setFixedWidth(int(w*2))
            self.t2.setFixedWidth(int(w*2))
            sw = self.step.sizeHint().width()
            self.step.setFixedWidth(int(sw*2))
        except Exception:
            pass
        self.b_auto=QtWidgets.QPushButton("자동생성")
        self.b_reset=QtWidgets.QPushButton("초기화")
        top.addWidget(QtWidgets.QLabel("시작")); top.addWidget(self.t1)
        top.addWidget(QtWidgets.QLabel("종료")); top.addWidget(self.t2)
        top.addWidget(QtWidgets.QLabel("간격(분)")); top.addWidget(self.step)
        top.addStretch(1); top.addWidget(self.b_auto); top.addWidget(self.b_reset)
        v.addLayout(top)
        hb=QtWidgets.QHBoxLayout();
        self.b_del=QtWidgets.QPushButton("행 삭제"); hb.addWidget(self.b_del); hb.addStretch(1)
        v.addLayout(hb)
        self.tbl=QtWidgets.QTableWidget( 0,4); self.tbl.setHorizontalHeaderLabels(["번호","시작","종료"]); self.tbl.setHorizontalHeaderLabels(["번호","시작","종료",""]); self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.verticalHeader().setVisible(False)
        try:
            self.tbl.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignHCenter)
        except Exception:
            pass
        # 셀 직접 편집 금지 (자동생성된 시간대만 사용)
        self.tbl.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        # 번호/시간/종료 열 가로간격 축소 & 가운데 정렬
        try:
            self.tbl.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignHCenter)
            self.tbl.setColumnWidth(0, 50)
            self.tbl.setColumnWidth(1, 70)
            self.tbl.setColumnWidth(2, 70)
        except Exception:
            pass
        v.addWidget(self.tbl,1)
        self._ranges=[] # list of (QTime,QTime,step)
        self.b_auto.clicked.connect(self.generate)
        self.b_reset.clicked.connect(self.reset)
        self.b_del.clicked.connect(self.del_row)

    def add_row(self):
        r=self.tbl.rowCount(); self.tbl.insertRow(r)
        it0=QtWidgets.QTableWidgetItem(str(r+1)); it0.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,0,it0);it1=QtWidgets.QTableWidgetItem(self.t1.time().toString("HH:mm")); it1.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,1,it1);it2=QtWidgets.QTableWidgetItem(self.t2.time().toString("HH:mm")); it2.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,2,it2);it3=QtWidgets.QTableWidgetItem(""); it3.setFlags(it3.flags() & ~QtCore.Qt.ItemIsEditable); self.tbl.setItem(r,3,it3)

    def del_row(self):
        rows=sorted({i.row() for i in self.tbl.selectedIndexes()}, reverse=True)
        for r in rows: self.tbl.removeRow(r)
        for i in range(self.tbl.rowCount()):
            self.tbl.setItem(i,0,QtWidgets.QTableWidgetItem(str(i+1)))


    def reset(self):
        """조사시간 설정 초기화: 범위와 테이블을 모두 비움"""
        if QtWidgets.QMessageBox.question(self, "초기화", "조사시간 설정을 모두 비울까요?") != QtWidgets.QMessageBox.Yes:
            return
        self._ranges.clear()
        self.tbl.setRowCount(0)
    def add_range(self):
        a=self.t1.time(); b=self.t2.time(); s=self.step.value()
        if a>=b: return
        key=(a.toString('HH:mm'), b.toString('HH:mm'), s)
        exists={(x[0].toString('HH:mm'), x[1].toString('HH:mm'), x[2]) for x in self._ranges}
        if key in exists:
            QtWidgets.QMessageBox.information(self,"알림","이미 동일한 구간이 추가되어 있습니다."); return
        self._ranges.append((a,b,s))
        self._refresh_ranges()

    def _refresh_ranges(self):
        self.tbl.setRowCount(0)
        for i,(a,b,s) in enumerate(self._ranges,1):
            self.tbl.insertRow(self.tbl.rowCount())
            self.tbl.setItem(self.tbl.rowCount()-1,0,QtWidgets.QTableWidgetItem(str(i)));
            self.tbl.item(self.tbl.rowCount()-1,0).setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
            self.tbl.setItem(self.tbl.rowCount()-1,1,QtWidgets.QTableWidgetItem(a.toString('HH:mm')))
            self.tbl.setItem(self.tbl.rowCount()-1,2,QtWidgets.QTableWidgetItem(b.toString('HH:mm')))

    @staticmethod
    def _segment(a:QtCore.QTime, b:QtCore.QTime, step:int):
        import datetime as _dt
        t=_dt.datetime(2000,1,1,a.hour(),a.minute())
        e=_dt.datetime(2000,1,1,b.hour(),b.minute())
        out=[]
        while t<e:
            t2=t+_dt.timedelta(minutes=step)
            if t2>e: break
            out.append((t.strftime("%H:%M"), t2.strftime("%H:%M")))
            t=t2
        return out

    def generate(self):
        """현재 시작/종료/간격 설정을 기준으로 시간대를 자동생성하여 기존 목록 뒤에 이어붙입니다."""
        segs = self._segment(self.t1.time(), self.t2.time(), self.step.value())
        # 기존 행 수를 유지한 채 뒤에 추가
        for s, e in segs:
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            it0 = QtWidgets.QTableWidgetItem(str(r + 1))
            it0.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it0.setFlags(it0.flags() & ~QtCore.Qt.ItemIsEditable)
            it1 = QtWidgets.QTableWidgetItem(s)
            it1.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it1.setFlags(it1.flags() & ~QtCore.Qt.ItemIsEditable)
            it2 = QtWidgets.QTableWidgetItem(e)
            it2.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it2.setFlags(it2.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 0, it0)
            self.tbl.setItem(r, 1, it1)
            self.tbl.setItem(r, 2, it2)


    def slots(self):
        """현재 테이블에 표시된 시간대(번호/시작/종료)를 그대로 반환합니다."""
        return self.get()

    def get(self):
        out=[]
        for r in range(self.tbl.rowCount()):
            a=self.tbl.item(r,1).text() if self.tbl.item(r,1) else ""
            b=self.tbl.item(r,2).text() if self.tbl.item(r,2) else ""
            out.append({"번호":r+1,"시작":a,"종료":b})
        return out

    def set(self, rows):
        self.tbl.setRowCount(0)
        for i,row in enumerate(rows or []):
            self.tbl.insertRow(i)
            it0 = QtWidgets.QTableWidgetItem(str(row.get("번호", i + 1)))
            it0.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it0.setFlags(it0.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(i, 0, it0)

            it1 = QtWidgets.QTableWidgetItem(row.get("시작", ""))
            it1.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it1.setFlags(it1.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(i, 1, it1)

            it2 = QtWidgets.QTableWidgetItem(row.get("종료", ""))
            it2.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            it2.setFlags(it2.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(i, 2, it2)

class SurveyVehicleTab(QtWidgets.QWidget):
    def __init__(self, page):
        super().__init__(page); self.page = page
        v=QtWidgets.QVBoxLayout(self)
        top=QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("차종 선택"))
        self.cb=QtWidgets.QComboBox(); top.addWidget(self.cb,1)
        self.b_reload=QtWidgets.QPushButton("재읽기"); top.addWidget(self.b_reload)
        v.addLayout(top)
        opt=QtWidgets.QHBoxLayout()
        self.chk_local=QtWidgets.QCheckBox("조사 로컬 편집(원본 덮어쓰기 아님)")
        opt.addWidget(self.chk_local); opt.addStretch(1)
        self.b_copy=QtWidgets.QPushButton("프로젝트에서 복사"); self.b_add=QtWidgets.QPushButton("행 추가"); self.b_del=QtWidgets.QPushButton("행 삭제")
        for b in (self.b_copy,self.b_add,self.b_del): opt.addWidget(b)
        v.addLayout(opt)
        self.tbl=QtWidgets.QTableWidget( 0,4); self.tbl.setHorizontalHeaderLabels(["차종명","순번","단축키",""]); self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.setHorizontalHeaderLabels(["순번","차종명","단축키",""])
        self.tbl.verticalHeader().setVisible(False)
        try:
            self.tbl.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignHCenter)
        except Exception:
            pass
        # 열 가로폭 고정: 순번/차종명/단축키/빈열
        try:
            self.tbl.setColumnWidth(0, 50)
            self.tbl.setColumnWidth(1, 150)
            self.tbl.setColumnWidth(2, 120)  # 단축키 열 넓게
        except Exception:
            pass
        v.addWidget(self.tbl,1)
        self.b_reload.clicked.connect(self.refresh_combo); self.cb.currentIndexChanged.connect(self.load_from_proj)
        self.b_copy.clicked.connect(self.copy_from_proj); self.b_add.clicked.connect(self.add_row); self.b_del.clicked.connect(self.del_row)
        self.refresh_combo()

    def refresh_combo(self):
        self.cb.blockSignals(True); self.cb.clear()
        for p in self.page.data["projects"]: self.cb.addItem(p["name"])
        self.cb.blockSignals(False); self.load_from_proj()

    def proj_rows(self):
        i=self.cb.currentIndex(); 
        return self.page.data["projects"][i].get("vehicle_set",[]) if i>=0 else []

    def load_from_proj(self):
        if self.chk_local.isChecked():
            return
        self.tbl.setRowCount(0)
        for i, row in enumerate(self.proj_rows(), 1):
            r = self.tbl.rowCount()
            self.tbl.insertRow(r)
            self.tbl.setItem(r, 0, QtWidgets.QTableWidgetItem(str(i)))
            self.tbl.setItem(r, 1, QtWidgets.QTableWidgetItem(row.get("차종명", "")))
            self.tbl.setItem(r, 2, QtWidgets.QTableWidgetItem(""))
            for c in (0, 1, 2):
                it = self.tbl.item(r, c)
                if it:
                    it.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        try:
            # 조사차종 설정 시트 열 폭 고정: 순번 / 차종명 / 단축키
            self.tbl.setColumnWidth(0, 60)
            self.tbl.setColumnWidth(1, 150)
            self.tbl.setColumnWidth(2, 60)
        except Exception:
            pass



    def copy_from_proj(self):
        self.tbl.setRowCount(0)
        for i,row in enumerate(self.proj_rows(),1):
            r=self.tbl.rowCount(); self.tbl.insertRow(r)
            self.tbl.setItem(r,0,QtWidgets.QTableWidgetItem(str(i)))
            self.tbl.setItem(r,1,QtWidgets.QTableWidgetItem(row.get("차종명","")))
            self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(""))
            for c in (0,1,2):
                it = self.tbl.item(r,c)
                if it: it.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
            
            try:
                # 조사차종 설정 시트 열 폭 고정: 순번 / 차종명 / 단축키
                self.tbl.setColumnWidth(0, 60)
                self.tbl.setColumnWidth(1, 150)
                self.tbl.setColumnWidth(2, 60)
            except Exception:
                pass

        

    def add_row(self):
        r=self.tbl.rowCount(); self.tbl.insertRow(r)
        it0=QtWidgets.QTableWidgetItem(str(r+1)); it0.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); it0.setFlags(it0.flags() & ~QtCore.Qt.ItemIsEditable); self.tbl.setItem(r,0,it0);it1=QtWidgets.QTableWidgetItem(""); it1.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,1,it1);it2=QtWidgets.QTableWidgetItem(""); it2.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter); self.tbl.setItem(r,2,it2);it3=QtWidgets.QTableWidgetItem(""); it3.setFlags(it3.flags() & ~QtCore.Qt.ItemIsEditable); self.tbl.setItem(r,3,it3)
        

    def del_row(self):
        r=self.tbl.currentRow(); 
        if r>=0: self.tbl.removeRow(r)

    def get(self):
        out=[]
        for r in range(self.tbl.rowCount()):
            out.append({"순번":r+1,
                        "차종명": self.tbl.item(r,1).text() if self.tbl.item(r,1) else "",
                        "단축키": self.tbl.item(r,2).text() if self.tbl.item(r,2) else ""})
        return {"작업참조": self.cb.currentText(), "로컬편집": self.chk_local.isChecked(), "차종목록": out}

    def set(self, d):
        self.refresh_combo()
        if not d: return
        self.chk_local.setChecked(bool(d.get("로컬편집", False)))
        if d.get("작업참조"): self.cb.setCurrentText(d["작업참조"])
        rows=d.get("차종목록")
        if rows:
            self.tbl.setRowCount(0)
            for i,row in enumerate(rows,1):
                r=self.tbl.rowCount(); self.tbl.insertRow(r)
                self.tbl.setItem(r,0,QtWidgets.QTableWidgetItem(str(i)))
                self.tbl.setItem(r,1,QtWidgets.QTableWidgetItem(row.get("차종명","")))
                self.tbl.setItem(r,2,QtWidgets.QTableWidgetItem(row.get("단축키","")))
                for c in (0,1,2):
                    it=self.tbl.item(r,c)
                    if it: it.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)

                
            try:
                # 조사차종 설정 시트 열 폭 고정: 순번 / 차종명 / 단축키
                self.tbl.setColumnWidth(0, 60)
                self.tbl.setColumnWidth(1, 150)
                self.tbl.setColumnWidth(2, 60)
            except Exception:
                pass


class SurveySitesTab(QtWidgets.QWidget):
    def __init__(self, page):
        super().__init__(page)
        self.page = page
        # row_key -> {"groups": [...], "counters": [...]}
        self.site_data = {}

        root = QtWidgets.QHBoxLayout(self)

        # -------- 좌측: 지점정보 --------
        left = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("지점정보")
        left.addWidget(title)

        # 순번 / 지번 / 지점명 / 작업번호 / 방향수 / 상태 / (빈 열)
        self.tbl = QtWidgets.QTableWidget(0, 7)
        self.tbl.setHorizontalHeaderLabels(["순번", "지번", "지점명", "작업번호", "방향수", "상태", ""])
        self.tbl.verticalHeader().setVisible(False)
        # 지점정보 리스트 행 높이를 기존보다 약 3px 줄여서 표시
        try:
            vh = self.tbl.verticalHeader()
            vh.setDefaultSectionSize(max(0, vh.defaultSectionSize() - 3))
        except Exception:
            pass
        # 표선이 보이지 않도록 설정 (스크린샷 스타일)
        self.tbl.setShowGrid(False)
        try:
            self.tbl.setColumnWidth(0, 38)   # 순번
            self.tbl.setColumnWidth(1, 48)   # 지번
            self.tbl.setColumnWidth(2, 148)  # 지점명
            self.tbl.setColumnWidth(3, 56)   # 작업번호
            self.tbl.setColumnWidth(4, 46)   # 방향수
            self.tbl.setColumnWidth(5, 46)   # 상태
            self.tbl.setColumnWidth(6, 6)   # 상태 우측 빈 슬롯
        except Exception:
            pass
        # 지점정보 테이블 최소 가로 크기 고정 (컬럼 합계 기준)
        try:
            self.tbl.setMinimumWidth(420)
        except Exception:
            pass
        # 행 전체 선택 + 한번 더 클릭 또는 더블클릭 시 편집 가능
        self.tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tbl.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.tbl.setEditTriggers(
            QtWidgets.QAbstractItemView.SelectedClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
            | QtWidgets.QAbstractItemView.DoubleClicked
        )
        self.tbl.horizontalHeader().setStretchLastSection(True)
        left.addWidget(self.tbl, 1)

        btn_row = QtWidgets.QHBoxLayout()
        self.b_top = QtWidgets.QPushButton("↑↑")
        self.b_up = QtWidgets.QPushButton("▲")
        self.b_dn = QtWidgets.QPushButton("▼")
        self.b_apply_in = QtWidgets.QPushButton("입력 등록")
        self.b_add = QtWidgets.QPushButton("추가")
        self.b_del = QtWidgets.QPushButton("삭제")
        for b in (self.b_top, self.b_up, self.b_dn, self.b_apply_in, self.b_add, self.b_del):
            btn_row.addWidget(b)
        btn_row.addStretch(1)
        left.addLayout(btn_row)

        root.addLayout(left, 8)

        # -------- 우측: 카운터 설정 --------
        right = QtWidgets.QVBoxLayout()
        gb = QtWidgets.QGroupBox("카운터 설정")
        gb.setMaximumWidth(260)
        g = QtWidgets.QGridLayout(gb)
        g.setColumnMinimumWidth(1, 0)
        g.setColumnMinimumWidth(2, 5)  # 화살표 컬럼(2번)의 가로 여백 조정용
        g.setColumnMinimumWidth(2, 0)
        g.setHorizontalSpacing(0)

        # 0행: 방향수 + 적용/자동생성/설정 저장
        g.addWidget(QtWidgets.QLabel("방향수"), 0, 0)
        self.spin = QtWidgets.QSpinBox()
        self.spin.setRange(1, 40)
        self.spin.setValue(2)
        # 방향수 콤보박스가 잘 보이도록 적당한 고정 폭으로 설정
        try:
            w = self.spin.sizeHint().width()
            self.spin.setFixedWidth(max(60, int(w * 0.9)))
        except Exception:
            self.spin.setFixedWidth(65)
        g.addWidget(self.spin, 0, 1)

        # 자동생성 / 템플릿 버튼: 우측 정렬 (저장 버튼 제거)
        self.b_auto = QtWidgets.QPushButton("생성")
        self.b_tpl = QtWidgets.QPushButton("템플릿")

        try:
            bw_auto = self.b_auto.sizeHint().width()
            bw_tpl = self.b_tpl.sizeHint().width()
            self.b_auto.setFixedWidth(int(bw_auto * 1.0))
            self.b_tpl.setFixedWidth(int(bw_tpl * 1.0))
        except Exception:
            pass

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(self.b_auto)
        btn_row.addWidget(self.b_tpl)
        g.addLayout(btn_row, 0, 3, 1, 2)

        # 1행: 라벨들
        g.addWidget(QtWidgets.QLabel("방향번호"), 1, 0)
        g.addWidget(QtWidgets.QLabel("입력 그룹(탭)"), 1, 2, 1, 2)

        # 2~5행: 방향번호 리스트 / 화살표 / 입력 그룹(탭) / 카운터
        self.dir = QtWidgets.QListWidget()
        self.dir.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.dir.setFixedWidth(50)
        self.dir.setMinimumHeight(290)
        g.addWidget(self.dir, 2, 0, 5, 1, alignment=QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

        vbtn = QtWidgets.QVBoxLayout()
        vbtn.setContentsMargins(0, 2, 0, 1)  # 화살표 좌우 여백 조정
        vbtn.addStretch(1)
        self.to = QtWidgets.QPushButton("▶")
        self.fr = QtWidgets.QPushButton("◀")
        vbtn.addWidget(self.to)
        vbtn.addWidget(self.fr)
        vbtn.addStretch(1)
        g.addLayout(vbtn, 2, 1, 4, 1, alignment=QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

        # 오른쪽 위: 입력 그룹(탭)
        self.tabs = QtWidgets.QListWidget()
        self.tabs.setFixedWidth(150)
        # 세로 높이를 더 줄여서(약 100 정도) 하단 카운터 시트 공간을 확보
        try:
            h = self.tabs.sizeHint().height()
            if h > 100:
                self.tabs.setMinimumHeight(max(40, h - 100))
                self.tabs.setMaximumHeight(max(40, h - 100))
            else:
                self.tabs.setMinimumHeight(40)
                self.tabs.setMaximumHeight(40)
        except Exception:
            pass
        g.addWidget(self.tabs, 2, 2, 2, 4)

        # 오른쪽 아래: 카운터 테이블
        g.addWidget(QtWidgets.QLabel("카운터"), 4, 2, 1, 2)
        self.ct = QtWidgets.QTableWidget(0, 4)
        self.ct.setMinimumWidth(110)
        # 세로 높이도 약 10 축소
        try:
            h = self.ct.sizeHint().height()
            if h > 10:
                self.ct.setMinimumHeight(70)
        except Exception:
            pass
        # 카운터 시트 높이는 현재 레이아웃 비율에만 맡긴다 (별도 증가 없음)
        # 열 구성: 카운터 / 번호 / 표시명 / (빈칸)
        self.ct.setHorizontalHeaderLabels(["카운터", "방향", "표시명", ""])
        self.ct.verticalHeader().setVisible(False)
        # 카운터 테이블 폰트 크기 축소 (입력그룹 시트와 유사하게)
        try:
            f = self.ct.font()
            if f.pointSize() > 0:
                f.setPointSize(max(8, f.pointSize() - 3))
                self.ct.setFont(f)
                self.ct.horizontalHeader().setFont(f)
        except Exception:
            pass
        # 행 높이도 줄여서 세로 간격을 축소
        try:
            vh = self.ct.verticalHeader()
            h = vh.defaultSectionSize()
            if h > 14:
                vh.setDefaultSectionSize(max(10, h - 6))
        except Exception:
            pass
        # 헤더: 마지막 열(빈칸)은 최대한 좁게 사용 (남는 공간을 채우지 않도록)
        header = self.ct.horizontalHeader()
        header.setStretchLastSection(False)
        try:
            header.setSectionResizeMode(3, QtWidgets.QHeaderView.Fixed)
        except Exception:
            pass
                # 카운터 시트 열 가로폭을 1/2 수준으로 축소 (컴팩트하게)
        try:
            self.ct.setColumnWidth(0, 60)   # 카운터
            self.ct.setColumnWidth(1, 40)   # 방향
            self.ct.setColumnWidth(2, 120)  # 표시명
            self.ct.setColumnWidth(3, 1)    # 빈칸 폭 최대한 축소
        except Exception:
            pass
        try:
            self.ct.setColumnWidth(0, 80)
            self.ct.setColumnWidth(0, 80)
        except Exception:
            pass
        g.addWidget(self.ct, 5, 2, 2, 4)

        # 카운터 셀 더블클릭 시 단축키 정보 팝업
        self.ct.cellDoubleClicked.connect(self._on_counter_cell_double_clicked)
        # 표시명 등 카운터시트 편집 즉시 저장(현재 지점 row에 반영)
        try:
            self.ct.itemChanged.connect(lambda *_: self._apply_to_current_row(False))
        except Exception:
            pass

        # 카운터 하단 추가/삭제 버튼 제거 (레이아웃에 추가하지 않음)
        hb = QtWidgets.QHBoxLayout()
        self.c_add = QtWidgets.QPushButton("추가")
        self.c_del = QtWidgets.QPushButton("삭제")
        # hb.addWidget(self.c_add)
        # hb.addWidget(self.c_del)
        # hb.addStretch(1)
        # g.addLayout(hb, 7, 3, 1, 2)

        right.addWidget(gb, 1)

        self.b_preview = QtWidgets.QPushButton("시트 미리보기")
        right.addWidget(self.b_preview, 0, alignment=QtCore.Qt.AlignRight)

        root.addLayout(right, 1)

        # 레이아웃 비율 조정
        # 입력그룹(탭) 영역을 조금 줄이고, 하단 카운터 시트 영역을 늘린다.
        g.setRowStretch(2, 2)
        g.setRowStretch(5, 3)

        # -------- 시그널 연결 --------
        self.tbl.currentCellChanged.connect(self._on_select_row)
        self.b_auto.clicked.connect(self.gen_dirs)
        self.to.clicked.connect(self.to_tabs)
        self.fr.clicked.connect(self.from_tabs)
        self.c_add.clicked.connect(self.add_counter)
        self.c_del.clicked.connect(self.del_counter)
        self.b_preview.clicked.connect(self.preview)
        # 입력 그룹(탭) 더블클릭 시 단축키 설정 팝업
        self.tabs.itemDoubleClicked.connect(self._on_group_tab_double_clicked)
        self.b_add.clicked.connect(self.add_site)
        self.b_del.clicked.connect(self.del_site)
        self.b_apply_in.clicked.connect(self.apply_to_selected)
        self.b_top.clicked.connect(self.move_site_top)
        self.b_up.clicked.connect(lambda: self.move_site(-1))
        self.b_dn.clicked.connect(lambda: self.move_site(1))
        # 템플릿에서 입력그룹/카운터 불러오기
        self.b_tpl.clicked.connect(self.apply_template_from_vehicle)
        # 방향수/그룹/카운터 설정이 바뀔 때마다 지점정보에 즉시 반영
        self.spin.valueChanged.connect(lambda *_: self._apply_to_current_row(False))

    # ---------- 내부 헬퍼 ----------
    def _row_key(self, row=None):
        if row is None:
            row = self.tbl.currentRow()
        return str(row)

    def _apply_state_color(self, item):
        """지점정보 상태 셀 색상 지정: 대기=검정, 진행=파랑, 완료=연회색"""
        if not item:
            return
        txt = item.text().strip()
        color_map = {
            "대기": QtGui.QColor("black"),
            "진행": QtGui.QColor("blue"),
            "완료": QtGui.QColor("lightgray"),
        }
        item.setForeground(color_map.get(txt, QtGui.QColor("black")))


    def _next_job_no(self):
        """
        작업번호를 'WN_YYMMDDhhmmssff' 형식으로 생성.
        예: 2025년 12월 4일 14시 04분 30초 13 = WN_25120414043013
        """
        now = dt.datetime.now()
        # 마이크로초를 1/100 단위(두 자리수)로 변환
        ms2 = int(now.microsecond / 10000)  # 0~99
        return "WN_" + now.strftime("%y%m%d%H%M%S") + f"{ms2:02d}"

    def _collect_ui(self):

        # 현재 UI에서 그룹/카운터/방향 리스트/방향수까지 모두 수집
        groups = [self.tabs.item(i).text() for i in range(self.tabs.count())]
        counters = []
        for r in range(self.ct.rowCount()):
            def _text(col):
                item = self.ct.item(r, col)
                return item.text().strip() if item else ""
            counters.append({
                "name": _text(0),
                "dir": _text(1),
                "label": _text(2),
            })
        dirs = [self.dir.item(i).text() for i in range(self.dir.count())]

        # 이미 저장된 지점 데이터가 있다면, 그 안의 group_projects(그룹별 차종유형 선택값)를 그대로 가져온다.
        try:
            row_key = self._row_key()
            prev = self.site_data.get(row_key, {}) if hasattr(self, "site_data") else {}
            group_projects = prev.get("group_projects", {}) or {}
        except Exception:
            group_projects = {}

        cfg = {"groups": groups, "counters": counters, "dirs": dirs, "spin": self.spin.value()}
        if group_projects:
            cfg["group_projects"] = group_projects
        return cfg



    # ---------- 카운터 / 단축키 연동 ----------
    def _current_project(self):
        """조사차종 설정 탭에서 현재 선택된 차종 유형(Project)을 반환."""
        veh_tab = getattr(self.page, "veh", None)
        if veh_tab is None:
            return None
        try:
            proj_index = veh_tab.cb.currentIndex()
        except Exception:
            return None
        projects = (self.page.data or {}).get("projects", [])
        if 0 <= proj_index < len(projects):
            return projects[proj_index]
        return None

    def _hotkey_sheets(self):
        """현재 프로젝트의 카운터 단축키 시트 목록을 반환."""
        proj = self._current_project()
        if proj is None:
            return []
        return proj.get("hotkey_sheets_global", []) or []

    
    def _on_group_tab_double_clicked(self, item):
        """입력그룹(탭) 더블클릭 시, 해당 그룹에 포함된 방향들의 단축키 시트를 일괄 설정."""
        if item is None:
            return

        label = item.text() or ""
        import re as _re
        dirs = []
        for part in _re.split(r"[^0-9]+", label):
            if part.isdigit():
                try:
                    dirs.append(int(part))
                except Exception:
                    pass
        dirs = sorted(set(dirs))
        if not dirs:
            QtWidgets.QMessageBox.information(self, "알림", "이 탭에서 방향번호를 찾을 수 없습니다.")
            return

        data = getattr(self.page, "data", None) or {}
        projects = data.get("projects", [])
        if not projects:
            QtWidgets.QMessageBox.information(self, "알림", "먼저 차종관리에서 차종유형과 카운터 단축키를 설정하세요.")
            return

        # 기본 프로젝트 인덱스:
        # 1순위: 현재 지점(site_data)에 저장된 그룹별 차종유형 선택값
        # 2순위: 조사차종 탭의 현재 선택값
                # 기본 프로젝트 인덱스:
        # 1순위: 현재 지점(site_data)에 저장된 그룹별 차종유형 선택값
        # 2순위: 이미 이 그룹의 방향들에 연결된 시트 이름에서 추정
        # 3순위: 조사차종 탭의 현재 선택값
        proj_index = None
        try:
            row_key = self._row_key()
            cfg_existing = self.site_data.get(row_key, {}) or {}
            group_projects = cfg_existing.get("group_projects", {}) or {}
            if label in group_projects:
                saved_idx = int(group_projects.get(label, 0))
                if 0 <= saved_idx < len(projects):
                    proj_index = saved_idx
        except Exception:
            proj_index = None

        # 저장된 group_projects 에 값이 없으면, 현재 카운터 테이블에 설정된 시트 이름을 이용해 추정
        if proj_index is None:
            try:
                # 이 그룹(label)에 포함된 방향들 중, 이미 시트가 지정된 방향을 찾는다.
                existing_sheet = None
                for row in range(self.ct.rowCount()):
                    d_item = self.ct.item(row, 1)
                    n_item = self.ct.item(row, 0)
                    if not d_item or not n_item:
                        continue
                    try:
                        d_val = int(d_item.text().strip())
                    except Exception:
                        continue
                    if d_val not in dirs:
                        continue
                    sheet_name = n_item.text().strip()
                    if not sheet_name:
                        continue
                    existing_sheet = sheet_name
                    break

                # 찾은 시트 이름이 있으면, 그 시트를 포함하고 있는 프로젝트 인덱스를 찾는다.
                if existing_sheet and projects:
                    for idx, proj in enumerate(projects):
                        sheets_local = proj.get("hotkey_sheets_global", []) or []
                        for s in sheets_local:
                            if str(s.get("name", "") or "").strip() == existing_sheet:
                                proj_index = idx
                                break
                        if proj_index is not None:
                            break
            except Exception:
                proj_index = None

        # 그래도 proj_index 를 찾지 못하면 조사차종 탭의 현재 선택값 사용
        if proj_index is None:
            veh_tab = getattr(self.page, "veh", None)
            if veh_tab is not None:
                try:
                    idx = veh_tab.cb.currentIndex()
                    if 0 <= idx < len(projects):
                        proj_index = idx
                except Exception:
                    proj_index = 0

        if proj_index is None:
            proj_index = 0
        if proj_index == 0:
            veh_tab = getattr(self.page, "veh", None)
            if veh_tab is not None:
                try:
                    idx = veh_tab.cb.currentIndex()
                    if 0 <= idx < len(projects):
                        proj_index = idx
                except Exception:
                    pass

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("입력그룹 단축키 설정 (지점별)")
        vbox = QtWidgets.QVBoxLayout(dlg)

        # 차종유형 콤보박스
        form_top = QtWidgets.QFormLayout()
        cb_proj = QtWidgets.QComboBox(dlg)
        for p in projects:
            cb_proj.addItem(p.get("name", "") or "(이름 없음)")
        cb_proj.setCurrentIndex(proj_index)
        form_top.addRow("차종유형:", cb_proj)
        vbox.addLayout(form_top)

        # 현재 카운터 테이블에서 방향별로 이미 지정된 시트 이름을 수집
        current_by_dir = {}
        for row in range(self.ct.rowCount()):
            d_item = self.ct.item(row, 1)
            n_item = self.ct.item(row, 0)
            if not d_item:
                continue
            d_text = d_item.text().strip()
            if not d_text:
                continue
            if n_item:
                current_by_dir.setdefault(d_text, n_item.text().strip())

        # 방향별 콤보박스
        dir_layout = QtWidgets.QFormLayout()
        dir_combos = {}

        def _sheet_names_for_project(p_index):
            if not (0 <= p_index < len(projects)):
                return []
            sheets = projects[p_index].get("hotkey_sheets_global", []) or []
            return [str(s.get("name", "") or "") for s in sheets]

        def _reload_dir_combos():
            names = _sheet_names_for_project(cb_proj.currentIndex())
            for d, combo in dir_combos.items():
                prev = combo.currentText().strip()
                combo.blockSignals(True)
                combo.clear()
                combo.addItem("")  # 비움 선택
                for n in names:
                    combo.addItem(n)
                # 기본 선택: 기존에 저장된 값이 있으면 우선
                key = str(d)
                base = current_by_dir.get(key, "")
                target = base or prev
                if target and target in names:
                    combo.setCurrentIndex(names.index(target) + 1)
                else:
                    combo.setCurrentIndex(0)
                combo.blockSignals(False)

        def _refresh_preview_all():
            # 미리보기 레이아웃 초기화
            while preview_layout.count():
                item = preview_layout.takeAt(0)
                w = item.widget()
                if w is not None:
                    w.deleteLater()

            # 현재 선택된 프로젝트의 시트 목록
            p_index = cb_proj.currentIndex()
            if not (0 <= p_index < len(projects)):
                sheets = []
            else:
                sheets = projects[p_index].get("hotkey_sheets_global", []) or []

            # 각 방향별로 그룹박스를 추가
            for d in dirs:
                combo = dir_combos.get(d)
                group_box = QtWidgets.QGroupBox(f"{d}번 방향")
                gb_layout = QtWidgets.QVBoxLayout(group_box)
                gb_layout.setContentsMargins(4, 4, 4, 4)
                gb_layout.setSpacing(2)

                if combo is None:
                    gb_layout.addWidget(QtWidgets.QLabel("방향 콤보박스를 찾을 수 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                sheet_name = combo.currentText().strip()
                if not sheet_name:
                    gb_layout.addWidget(QtWidgets.QLabel("선택된 단축키 시트가 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                target = None
                for s in sheets:
                    if str(s.get("name", "") or "") == sheet_name:
                        target = s
                        break

                if target is None:
                    gb_layout.addWidget(QtWidgets.QLabel(f"'{sheet_name}' 시트를 찾을 수 없습니다."))
                    preview_layout.addWidget(group_box)
                    continue

                items = target.get("items", []) or []
                if not items:
                    gb_layout.addWidget(QtWidgets.QLabel("등록된 단축키가 없습니다."))
                else:
                    for rec in items:
                        label_txt = str(rec.get("차종명", ""))
                        key_txt = str(rec.get("단축키", ""))
                        row = QtWidgets.QHBoxLayout()
                        lab = QtWidgets.QLabel(f"{label_txt} 키")
                        edit = QtWidgets.QLineEdit(key_txt)
                        edit.setReadOnly(True)
                        edit.setMaximumWidth(80)
                        row.addWidget(lab)
                        row.addWidget(edit)
                        gb_layout.addLayout(row)

                preview_layout.addWidget(group_box)

        def _on_dir_combo_changed(_index, d):
            # 콤보박스 변경 시 전체 방향 미리보기를 다시 그린다.
            _refresh_preview_all()
        for d in dirs:
            combo = QtWidgets.QComboBox(dlg)
            dir_combos[d] = combo
            dir_layout.addRow(f"{d}번 방향", combo)
            combo.currentIndexChanged.connect(lambda idx, dd=d: _on_dir_combo_changed(idx, dd))

        vbox.addLayout(dir_layout)

        # 미리보기 영역
        # 그룹박스를 쓰면 제목과 내용이 겹쳐 보이는 경우가 있어,
        # 제목 라벨 + 스크롤 영역 조합으로 교체했다.
        # 미리보기 영역: QGroupBox + QScrollArea 조합으로 구성하여
        # 제목과 내용이 자연스럽게 배치되고, 항목이 많을 경우 스크롤로 확인할 수 있도록 한다.
        preview_group = QtWidgets.QGroupBox("선택된 단축키 시트 미리보기")
        preview_group.setStyleSheet(
            "QGroupBox { margin-top: 12px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 8px; }"
        )
        group_layout = QtWidgets.QVBoxLayout(preview_group)
        group_layout.setContentsMargins(4, 4, 4, 4)
        group_layout.setSpacing(2)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        preview_container = QtWidgets.QWidget()
        preview_layout = QtWidgets.QVBoxLayout(preview_container)
        preview_layout.setAlignment(QtCore.Qt.AlignTop)
        preview_layout.setContentsMargins(4, 4, 4, 4)
        preview_layout.setSpacing(2)
        scroll.setWidget(preview_container)
        group_layout.addWidget(scroll)

        vbox.addWidget(preview_group)

        cb_proj.currentIndexChanged.connect(lambda *_: (_reload_dir_combos(), _refresh_preview_all()))

        # 초기 콤보박스 채우기
        _reload_dir_combos()
        # 모든 방향에 대한 미리보기 갱신
        _refresh_preview_all()


        # 버튼 행
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok = QtWidgets.QPushButton("OK")
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        vbox.addLayout(btn_row)

        def _on_ok():
            # 선택된 시트를 카운터 테이블에 반영
            for d in dirs:
                sheet_name = dir_combos[d].currentText().strip()
                if not sheet_name:
                    continue
                dir_str = str(d)
                found = False
                for row in range(self.ct.rowCount()):
                    d_item = self.ct.item(row, 1)
                    if d_item and d_item.text().strip() == dir_str:
                        name_item = QtWidgets.QTableWidgetItem(sheet_name)
                        name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                        # 카운터 이름 열은 사용자 편집 불가
                        name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
                        self.ct.setItem(row, 0, name_item)
                        found = True
                # 해당 방향번호에 대한 카운터 행이 없으면 새로 추가
                if not found:
                    r = self.ct.rowCount()
                    self.ct.insertRow(r)
                    name_item = QtWidgets.QTableWidgetItem(sheet_name)
                    name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
                    self.ct.setItem(r, 0, name_item)
                    dir_item = QtWidgets.QTableWidgetItem(dir_str)
                    dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    # 방향 열은 읽기 전용
                    dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
                    self.ct.setItem(r, 1, dir_item)
                    self.ct.setItem(r, 2, QtWidgets.QTableWidgetItem(""))
            # 현재 지점 설정을 저장
            self._apply_to_current_row(False)
            # 이 입력그룹(탭)에 사용된 차종유형(프로젝트 인덱스)을 함께 저장
            try:
                row_key = self._row_key()
                cfg_saved = self.site_data.get(row_key, {}) or {}
                group_projects = cfg_saved.get("group_projects", {}) or {}
                group_projects[label] = cb_proj.currentIndex()
                cfg_saved["group_projects"] = group_projects
                self.site_data[row_key] = cfg_saved
            except Exception:
                pass
            QtWidgets.QMessageBox.information(self, "완료", "입력그룹 단축키 설정이 완료되었습니다.")
            dlg.accept()

        btn_ok.clicked.connect(_on_ok)
        btn_cancel.clicked.connect(dlg.reject)

        dlg.exec_()

    def _on_counter_cell_double_clicked(self, row, column):
        """카운터 열을 더블클릭하면 해당 카운터 항목의 단축키 구성을 팝업으로 보여준다."""
        # 카운터 열(0번)이 아닌 경우는 무시
        if column != 0:
            return
        name_item = self.ct.item(row, 0)
        dir_item = self.ct.item(row, 1)
        if not name_item or not dir_item:
            return
        counter_name = name_item.text().strip()
        dir_text = dir_item.text().strip() or "?"

        # 이 지점/방향에 대해 실제로 사용해야 할 프로젝트(차종유형) 인덱스를 계산
        data = getattr(self.page, "data", {}) if hasattr(self.page, "data") else {}
        projects = data.get("projects", []) if isinstance(data, dict) else []

        proj_index = None
        # 1순위: site_data 에 저장된 group_projects 에서 방향별 프로젝트 인덱스 조회
        try:
            row_key = self._row_key()
            cfg = self.site_data.get(row_key, {}) or {}
            group_projects = cfg.get("group_projects", {}) or {}
            import re as _re
            try:
                dir_val = int(dir_text)
            except Exception:
                dir_val = None
            if dir_val is not None:
                for grp_label, p_idx in group_projects.items():
                    for part in _re.split(r"[^0-9]+", str(grp_label)):
                        if not part.isdigit():
                            continue
                        try:
                            d_val = int(part)
                        except Exception:
                            continue
                        if d_val == dir_val:
                            try:
                                proj_index = int(p_idx)
                            except Exception:
                                proj_index = None
                            break
                    if proj_index is not None:
                        break
        except Exception:
            proj_index = None

        # 2순위: 조사차종 탭의 현재 선택값 사용
        if proj_index is None and projects:
            try:
                veh_tab = getattr(self.page, "veh", None)
                if veh_tab is not None and hasattr(veh_tab, "cb"):
                    idx = veh_tab.cb.currentIndex()
                    if 0 <= idx < len(projects):
                        proj_index = idx
            except Exception:
                proj_index = None

        # 3순위: _hotkey_sheets() 가 사용하는 현재 프로젝트
        if proj_index is not None and 0 <= proj_index < len(projects):
            sheets = projects[proj_index].get("hotkey_sheets_global", []) or []
        else:
            sheets = self._hotkey_sheets()

        if not sheets:
            QtWidgets.QMessageBox.information(self, "알림", "현재 차종유형에 설정된 카운터 단축키 정보가 없습니다.")
            return

        # 1) 시트 name 과 정확히 일치하는 항목을 우선 검색
        target = None
        for s in sheets:
            if str(s.get("name", "")).strip() == counter_name:
                target = s
                break

        # 2) 없으면 counter_name 안의 숫자를 이용해 인덱스로 추정 ("항목1", "카운터2" 등)
        if target is None:
            import re as _re
            m = _re.search(r"(\d+)", counter_name)
            if m:
                idx = int(m.group(1)) - 1
                if 0 <= idx < len(sheets):
                    target = sheets[idx]

        # 3) 그래도 못 찾으면, 같은 방향번호를 가진 다른 카운터 행에서 시트 이름을 대신 찾는다.
        if target is None:
            for r in range(self.ct.rowCount()):
                if r == row:
                    continue
                alt_dir_item = self.ct.item(r, 1)
                alt_name_item = self.ct.item(r, 0)
                if not alt_dir_item or not alt_name_item:
                    continue
                if (alt_dir_item.text() or "").strip() != dir_text:
                    continue
                alt_name = (alt_name_item.text() or "").strip()
                if not alt_name:
                    continue
                for s in sheets:
                    if str(s.get("name", "")).strip() == alt_name:
                        target = s
                        break
                if target is not None:
                    break

        # 4) 마지막 시도: 모든 차종유형의 단축키 시트에서 같은 이름을 검색
        if target is None and projects:
            try:
                import re as _re
                try:
                    dir_val = int(dir_text)
                except Exception:
                    dir_val = None
                for p_idx, p in enumerate(projects):
                    for s in p.get("hotkey_sheets_global", []) or []:
                        if str(s.get("name", "")).strip() == counter_name:
                            target = s
                            # 찾은 경우, site_data 의 group_projects 에도 보강 저장
                            try:
                                row_key = self._row_key()
                                cfg_saved = self.site_data.get(row_key, {}) or {}
                                groups = cfg_saved.get("groups", []) or []
                                group_projects = cfg_saved.get("group_projects", {}) or {}
                                if dir_val is not None:
                                    for grp_label in groups:
                                        for part in _re.split(r"[^0-9]+", str(grp_label)):
                                            if part.isdigit() and int(part) == dir_val:
                                                group_projects[grp_label] = p_idx
                                                cfg_saved["group_projects"] = group_projects
                                                self.site_data[row_key] = cfg_saved
                                                break
                                        else:
                                            continue
                                        break
                            except Exception:
                                pass
                            break
                    if target is not None:
                        break
            except Exception:
                pass

        if target is None:
            QtWidgets.QMessageBox.information(
                self,
                "알림",
                f"'{counter_name}' 항목에 연결된 단축키 시트를 찾을 수 없습니다.\n"
                "입력그룹(탭)의 단축키 설정에서 먼저 시트를 지정해 주세요.",
            )
            return

        items = target.get("items", [])
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(f"단축키 설정 ({dir_text}번 방향)")
        v = QtWidgets.QVBoxLayout(dlg)

        # 상단: 방향 정보
        dir_label = QtWidgets.QLabel(f"{dir_text}번 방향")
        dir_label.setAlignment(QtCore.Qt.AlignCenter)
        font = dir_label.font()
        font.setBold(True)
        dir_label.setFont(font)
        v.addWidget(dir_label)

        # 각 차종별 단축키 정보를 나열 (읽기 전용)
        for rec in items:
            row_layout = QtWidgets.QHBoxLayout()
            name = rec.get("차종명", "")
            key = rec.get("단축키", "")
            lbl = QtWidgets.QLabel(f"{name} 키")
            edit = QtWidgets.QLineEdit(key)
            edit.setReadOnly(True)
            edit.setMaximumWidth(100)
            row_layout.addWidget(lbl)
            row_layout.addWidget(edit)
            row_layout.addStretch(1)
            v.addLayout(row_layout)

        if not items:
            v.addWidget(QtWidgets.QLabel("등록된 단축키 정보가 없습니다."))

        # 닫기 버튼
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QtWidgets.QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_close)
        v.addLayout(btn_row)

        dlg.exec_()
    def apply_template_from_vehicle(self):
        """조사차종 설정에서 선택된 차종 유형의 입력그룹 템플릿을 불러와 적용."""
        # SurveyManagerPage 내의 조사차종 탭(SurveyVehicleTab)에서 현재 프로젝트 인덱스 확인
        veh_tab = getattr(self.page, "veh", None)
        if veh_tab is None:
            QtWidgets.QMessageBox.information(self, "알림", "조사차종 설정 탭 정보를 찾을 수 없습니다.")
            return

        try:
            proj_index = veh_tab.cb.currentIndex()
        except Exception:
            QtWidgets.QMessageBox.information(self, "알림", "조사차종 설정 콤보박스를 찾을 수 없습니다.")
            return

        projects = (self.page.data or {}).get("projects", [])
        if not (0 <= proj_index < len(projects)):
            QtWidgets.QMessageBox.information(self, "알림", "먼저 조사차종 설정에서 차종 유형을 선택하세요.")
            return

        proj = projects[proj_index]
        templates = proj.get("templates", [])
        if not templates:
            QtWidgets.QMessageBox.information(self, "알림", "선택된 차종 유형에 등록된 입력그룹 템플릿이 없습니다.")
            return

        # 템플릿 선택 팝업
        names = [t.get("name", "(무제)") or "(무제)" for t in templates]
        name, ok = QtWidgets.QInputDialog.getItem(self, "템플릿 선택", "입력그룹 템플릿을 선택하세요:", names, 0, False)
        if not ok:
            return
        try:
            idx = names.index(name)
        except ValueError:
            QtWidgets.QMessageBox.information(self, "알림", "선택한 템플릿을 찾을 수 없습니다.")
            return

        tpl = templates[idx]
        dirs_list = tpl.get("dirs", []) or []
        counters = tpl.get("counters", []) or []

        # 방향수 계산: 템플릿의 그룹/카운터에서 사용된 방향번호의 최댓값
        dir_nums = set()
        for c in counters:
            try:
                d = int(str(c.get("dir", "")).strip())
                dir_nums.add(d)
            except Exception:
                pass
        for label in dirs_list:
            parts = str(label).split("-")
            for p in parts:
                try:
                    dir_nums.add(int(p))
                except Exception:
                    pass
        if dir_nums:
            max_dir = max(dir_nums)
        else:
            max_dir = max(self.spin.value(), 1)

        # 방향수 및 방향번호 리스트 재구성
        self.spin.setValue(max_dir)
        self.dir.clear()
        for i in range(1, self.spin.value() + 1):
            self.dir.addItem(str(i))

        # 입력그룹(탭) 구성 적용
        self.tabs.clear()
        for gname in dirs_list:
            self.tabs.addItem(str(gname))

        # 카운터 테이블 구성 적용
        self.ct.setRowCount(0)
        for c in counters:
            r = self.ct.rowCount()
            self.ct.insertRow(r)
            name_item = QtWidgets.QTableWidgetItem(c.get("name", ""))
            name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            # 카운터 이름 열은 사용자 편집 불가
            name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
            dir_item = QtWidgets.QTableWidgetItem(str(c.get("dir", "")))
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            # 방향 열은 읽기 전용
            dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
            label_item = QtWidgets.QTableWidgetItem(c.get("label", ""))
            self.ct.setItem(r, 0, name_item)
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.ct.setItem(r, 1, dir_item)
            self.ct.setItem(r, 2, label_item)

        # 현재 지점행에 설정 즉시 반영
        self._apply_to_current_row(False)

    def _load_ui(self, row):
        key = self._row_key(row)
        cfg = self.site_data.get(key, {})
        # 방향수 / 방향번호 리스트
        spin_val = cfg.get("spin")
        if isinstance(spin_val, int) and 1 <= spin_val <= 40:
            self.spin.setValue(spin_val)
        dirs = cfg.get("dirs") or []
        self.dir.clear()
        if dirs:
            for d in dirs:
                self.dir.addItem(str(d))
        else:
            # 저장된 방향 리스트가 없으면 현재 스핀 값 기준으로 1..N 생성
            for i in range(1, self.spin.value() + 1):
                self.dir.addItem(str(i))

        # 입력 그룹(탭)
        self.tabs.clear()
        for gname in cfg.get("groups", []):
            self.tabs.addItem(gname)

        # 카운터 테이블
        self.ct.setRowCount(0)
        for c in cfg.get("counters", []):
            r = self.ct.rowCount()
            self.ct.insertRow(r)
            name_item = QtWidgets.QTableWidgetItem(c.get("name", ""))
            name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            # 카운터 이름 열은 사용자 편집 불가
            name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
            dir_item = QtWidgets.QTableWidgetItem(str(c.get("dir", "")))
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            # 방향 열은 읽기 전용
            dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
            label_item = QtWidgets.QTableWidgetItem(c.get("label", ""))
            self.ct.setItem(r, 0, name_item)
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.ct.setItem(r, 1, dir_item)
            self.ct.setItem(r, 2, label_item)

    def _on_select_row(self, cur_row, cur_col, prev_row, prev_col):
        if prev_row is not None and prev_row >= 0:
            self.site_data[self._row_key(prev_row)] = self._collect_ui()
        if cur_row is not None and cur_row >= 0:
            self._load_ui(cur_row)

    # ---------- 방향/그룹/카운터 UI ----------
    def gen_dirs(self):
        """
        방향번호 생성:
        - 스핀박스 값(1~N)을 기준으로 전체 방향번호를 만든 뒤
        - 이미 입력그룹(탭)에 포함된 방향은 제외하여 남은 방향만 리스트에 표시합니다.
        """
        import re as _re

        all_dirs = set(range(1, self.spin.value() + 1))
        used_dirs = set()
        # 탭에 이미 사용된 방향번호 수집
        for i in range(self.tabs.count()):
            label = self.tabs.item(i).text()
            for part in _re.split(r"[^0-9]+", label):
                if part.isdigit():
                    used_dirs.add(int(part))

        remain = sorted(d for d in all_dirs if d not in used_dirs)
        self.dir.clear()
        for d in remain:
            self.dir.addItem(str(d))
        # 방향 리스트가 바뀌면 현재 지점의 설정도 갱신
        self._apply_to_current_row(False)

    def to_tabs(self):
        """선택된 방향번호들을 하나의 입력그룹(탭)으로 추가하고, 해당 방향은 방향번호 리스트에서 제거."""
        items = self.dir.selectedItems()
        if not items:
            return
        sels = [it.text() for it in items]
        try:
            nums = sorted(int(x) for x in sels)
            label = "-".join(str(n) for n in nums)
        except Exception:
            label = "-".join(sels)
        # 입력그룹 탭 추가
        self.tabs.addItem(label)

        # 추가된 방향번호는 좌측 방향번호 리스트에서 제거(중복 방지)
        rows = sorted({self.dir.row(it) for it in items}, reverse=True)
        for r in rows:
            self.dir.takeItem(r)

        # 설정 반영
        self._apply_to_current_row(False)

    def from_tabs(self):
        """선택된 입력그룹(탭)을 제거하고, 그 안에 포함된 방향번호를 다시 방향번호 리스트에 복원."""
        import re as _re

        items = self.tabs.selectedItems()
        if not items:
            return

        # 탭에서 복원해야 할 방향번호 수집
        restore_dirs = set()
        for it in items:
            label = it.text()
            for part in _re.split(r"[^0-9]+", label):
                if part.isdigit():
                    restore_dirs.add(int(part))

        # 선택된 탭 제거
        for it in items:
            self.tabs.takeItem(self.tabs.row(it))

        # 이미 리스트에 있는 방향번호와 중복되지 않도록 체크
        existing = set()
        for i in range(self.dir.count()):
            try:
                existing.add(int(self.dir.item(i).text()))
            except Exception:
                pass

        # 숫자 기준 오름차순 위치에 맞게 복원
        for d in sorted(restore_dirs):
            if d in existing:
                continue
            pos = self.dir.count()
            for i in range(self.dir.count()):
                try:
                    cur = int(self.dir.item(i).text())
                    if d < cur:
                        pos = i
                        break
                except Exception:
                    continue
            self.dir.insertItem(pos, str(d))

        # 설정 반영
        self._apply_to_current_row(False)

    def add_counter(self):
        r = self.ct.rowCount()
        self.ct.insertRow(r)
        name_item = QtWidgets.QTableWidgetItem(f"카운터{r+1}")
        name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        name_item.setFlags(name_item.flags() & ~QtCore.Qt.ItemIsEditable)
        dir_item = QtWidgets.QTableWidgetItem("1")
        dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        # 방향 열은 읽기 전용
        dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
        label_item = QtWidgets.QTableWidgetItem("")
        self.ct.setItem(r, 0, name_item)
        dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.ct.setItem(r, 1, dir_item)
        self.ct.setItem(r, 2, label_item)
        # 카운터 구성이 바뀌면 바로 저장
        self._apply_to_current_row(False)

    def del_counter(self):
        r = self.ct.currentRow()
        if r >= 0:
            self.ct.removeRow(r)
            # 카운터 구성이 바뀌면 바로 저장
            self._apply_to_current_row(False)

    # ---------- 지점 행 추가/삭제 ----------
    def add_site(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        # 순번: 가운데 정렬 (사용자 편집 불가)
        num_item = QtWidgets.QTableWidgetItem(str(r + 1))
        num_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        num_item.setFlags(num_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 0, num_item)

        # 지번: 가운데 정렬
        jibun_item = QtWidgets.QTableWidgetItem("")
        jibun_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.tbl.setItem(r, 1, jibun_item)

        # 지점명: 가운데 정렬
        name_item = QtWidgets.QTableWidgetItem("")
        name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.tbl.setItem(r, 2, name_item)

        # 작업번호: 전체 과업에서 중복되지 않게 자동 부여 (가운데 정렬, 편집 불가)
        job_no = self._next_job_no()
        job_item = QtWidgets.QTableWidgetItem(job_no)
        job_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        job_item.setFlags(job_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 3, job_item)

        # 방향수: 가운데 정렬 (사용자 편집 불가)
        dir_item = QtWidgets.QTableWidgetItem(str(self.spin.value()))
        dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 4, dir_item)

        # 상태: 조사 정보 탭의 진행상태를 반영
        state_text = ""
        try:
            state_text = self.page.info.cb_state.currentText()
        except Exception:
            state_text = "대기"
        if not state_text:
            state_text = "대기"
        state_item = QtWidgets.QTableWidgetItem(state_text)
        state_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        state_item.setFlags(state_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self._apply_state_color(state_item)
        self.tbl.setItem(r, 5, state_item)

        # 빈 슬롯 열(6)은 비워둠
        self.site_data[self._row_key(r)] = {"groups": [], "counters": []}

    def del_site(self):
        """선택된 지점 행(여러 개 선택 가능)을 삭제하고 순번/설정을 재정렬."""
        rows = self.get()
        if not rows:
            return

        # 선택된 행 인덱스 수집 (없으면 현재 행만 사용)
        sel_rows = sorted({idx.row() for idx in self.tbl.selectedIndexes()})
        if not sel_rows and self.tbl.currentRow() >= 0:
            sel_rows = [self.tbl.currentRow()]
        if not sel_rows:
            return

        # 뒤에서부터 삭제해야 인덱스가 꼬이지 않음
        for r in reversed(sel_rows):
            if 0 <= r < len(rows):
                del rows[r]

        # 지점정보/카운터 구성을 다시 세팅하여 site_data까지 함께 정리
        self.set(rows)

    def _apply_to_current_row(self, show_message=False):
        """현재 선택된 지점 행에 카운터 설정을 즉시 반영하고,
        '입력 그룹(탭)'에 실제로 포함된 방향번호만을 기준으로 방향수를 계산합니다.
        또한, 그룹별로 선택한 차종유형(프로젝트 인덱스) 정보(group_projects)는 기존 값을 유지합니다."""
        r = self.tbl.currentRow()
        if r < 0:
            return
        cfg = self._collect_ui()
        row_key = self._row_key(r)
        prev = self.site_data.get(row_key, {}) or {}
        # 기존에 저장된 group_projects(그룹별 차종유형 선택값)는 유지
        if "group_projects" in prev:
            cfg["group_projects"] = prev.get("group_projects", {})
        self.site_data[row_key] = cfg

        # 방향수 = 입력그룹(탭)에 포함된 모든 방향번호의 "종류 수"
        # (남은 방향번호 리스트는 포함하지 않음)
        import re as _re

        dir_numbers = set()
        for label in cfg.get("groups", []):
            for part in _re.split(r"[^0-9]+", label):
                if part.isdigit():
                    try:
                        dir_numbers.add(int(part))
                    except Exception:
                        pass

        # 입력그룹이 하나도 없다면, 스핀박스 값(기본 방향수)을 그대로 사용
        cnt = len(dir_numbers) if dir_numbers else self.spin.value()

        dir_item = QtWidgets.QTableWidgetItem(str(cnt))
        dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        # 방향수는 항상 자동 계산되므로 편집 불가능하게 고정
        dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
        self.tbl.setItem(r, 4, dir_item)
        if show_message:
            QtWidgets.QMessageBox.information(self, "적용", "선택 지점에 입력그룹/카운터 구성이 저장되었습니다.")
    def apply_to_selected(self):
        """[버튼용] 현재 선택된 지점에 카운터 설정을 저장 (알림 표시)."""
        if self.tbl.currentRow() < 0:
            QtWidgets.QMessageBox.information(self, "알림", "좌측 지점을 먼저 선택하세요.")
            return
        self._apply_to_current_row(show_message=True)

    def move_site(self, direction):
        """지점 행을 위/아래로 이동."""
        r = self.tbl.currentRow()
        if r < 0:
            return
        new_r = r + direction
        if not (0 <= new_r < self.tbl.rowCount()):
            return
        # 현재 행의 카운터 설정을 먼저 저장
        self._apply_to_current_row(show_message=False)
        # 전체 지점정보 + 카운터 설정을 가져와 순서를 바꾸고 다시 세팅
        rows = self.get()
        rows[r], rows[new_r] = rows[new_r], rows[r]
        self.set(rows)
        # 새 위치 행을 선택 & 포커스 이동
        try:
            self.tbl.setCurrentCell(new_r, 0)
        except Exception:
            try:
                self.tbl.selectRow(new_r)
            except Exception:
                pass


    def move_site_top(self):
        """선택된 지점 행을 최상단(1번 행)으로 이동."""
        r = self.tbl.currentRow()
        if r <= 0:
            return
        # 현재 행의 카운터 설정을 먼저 저장
        self._apply_to_current_row(show_message=False)
        rows = self.get()
        if not (0 <= r < len(rows)):
            return
        row = rows.pop(r)
        rows.insert(0, row)
        self.set(rows)
        try:
            self.tbl.setCurrentCell(0, 0)
        except Exception:
            try:
                self.tbl.selectRow(0)
            except Exception:
                pass

    # ---------- 저장/복원 ----------
    def get(self):
        """지점정보 + 카운터 설정까지 모두 저장"""
        # 현재 선택된 지점의 최신 편집 내용을 site_data에 반영
        try:
            self._apply_to_current_row(False)
        except Exception:
            pass
        out = []
        for r in range(self.tbl.rowCount()):
            def _text(col):
                item = self.tbl.item(r, col)
                return item.text().strip() if item else ""
            key = self._row_key(r)
            cfg = self.site_data.get(key, {"groups": [], "counters": []})
            rec = {
                "순번": r + 1,
                "지번": _text(1),
                "지점명": _text(2),
                "작업번호": _text(3),
                "방향수": _text(4),
                "상태": _text(5),
                "groups": cfg.get("groups", []),
                "counters": cfg.get("counters", []),
            }
            out.append(rec)
        return out

    def set(self, rows):
        """저장된 지점정보 복원"""
        self.tbl.setRowCount(0)
        self.site_data = {}
        rows = rows or []
        for r, row in enumerate(rows):
            self.tbl.insertRow(self.tbl.rowCount())
            # 순번 (편집 불가)
            num_item = QtWidgets.QTableWidgetItem(str(r + 1))
            num_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            num_item.setFlags(num_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 0, num_item)
            # 지번 (가운데 정렬) / 지점명 / 작업번호 / 방향수 / 상태
            jibun_item = QtWidgets.QTableWidgetItem(str(row.get("지번", "")))
            jibun_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.tbl.setItem(r, 1, jibun_item)
            name_item = QtWidgets.QTableWidgetItem(str(row.get("지점명", "")))
            name_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.tbl.setItem(r, 2, name_item)
            job_item = QtWidgets.QTableWidgetItem(str(row.get("작업번호", "")))
            job_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            job_item.setFlags(job_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 3, job_item)
            dir_item = QtWidgets.QTableWidgetItem(str(row.get("방향수", "")))
            dir_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            dir_item.setFlags(dir_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.tbl.setItem(r, 4, dir_item)
            state_item = QtWidgets.QTableWidgetItem(str(row.get("상태", "")))
            state_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            state_item.setFlags(state_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self._apply_state_color(state_item)
            self.tbl.setItem(r, 5, state_item)

            # 카운터 구성 복원용 메모리
            self.site_data[self._row_key(r)] = {
                "groups": row.get("groups", []),
                "counters": row.get("counters", []),
            }

    # ---------- 시트 미리보기 ----------
    def preview(self):
        """지점정보 시트 미리보기 (방향별).
        각 방향에 대해, 해당 방향에 연결된 단축키 시트가 있으면
        그 시트의 차종명을 열 제목으로 사용합니다.
        (없으면 프로젝트 기본 차종명을 사용)
        """
        # 시간대는 조사시간 탭에서 가져옴
        try:
            times = self.page.time.get()
        except Exception:
            times = []

        # 현재 선택된 지점의 그룹/카운터 정보
        cfg = {}
        r = self.tbl.currentRow()
        if r >= 0:
            cfg = self.site_data.get(self._row_key(r), {})
            groups = cfg.get("groups") or ["1"]
            counter_cfg = cfg.get("counters") or []
        else:
            groups = ["1"]
            counter_cfg = []

        # 방향별로 연결된 단축키 시트 이름 매핑: {방향번호: 시트이름}
        dir_to_sheet = {}
        for rec in counter_cfg:
            try:
                d_str = str(rec.get("dir", "")).strip()
                sheet_name = str(rec.get("name", "")).strip()
            except Exception:
                continue
            if not d_str or not sheet_name:
                continue
            if not d_str.isdigit():
                continue
            try:
                d_val = int(d_str)
            except Exception:
                continue
            if sheet_name:
                dir_to_sheet[d_val] = sheet_name

        # 그룹 문자열(예: "4-5-6", "7-8-9")에서 실제 사용하는 방향번호만 추출
        import re as _re
        dir_numbers = set()
        for label in groups:
            for part in _re.split(r"[^0-9]+", str(label)):
                if part.isdigit():
                    try:
                        dir_numbers.add(int(part))
                    except Exception:
                        pass

        if dir_numbers:
            dir_list = sorted(dir_numbers)
        else:
            # 그룹 정보가 비어 있는 경우 1번 방향만 존재한다고 가정
            dir_list = [1]

        # 프로젝트/차종/단축키 시트 정보
        data = getattr(self.page, "data", {}) if hasattr(self.page, "data") else {}
        projects = data.get("projects", []) if isinstance(data, dict) else []

        # 현재 지점의 group_projects(그룹별 차종유형 인덱스)에서 방향별 프로젝트 인덱스를 계산
        dir_to_proj_index = {}
        group_projects = cfg.get("group_projects", {}) or {}
        for grp_label, p_idx in group_projects.items():
            try:
                p_idx_int = int(p_idx)
            except Exception:
                continue
            for part in _re.split(r"[^0-9]+", str(grp_label)):
                if part.isdigit():
                    try:
                        d_val = int(part)
                    except Exception:
                        continue
                    dir_to_proj_index[d_val] = p_idx_int

        def _default_cols_for_direction(d):
            """해당 방향의 기본 차종 열 목록을 반환 (단축키 시트가 없을 때 사용)."""
            if not projects:
                return ["차종1", "차종2"]
            proj_idx = dir_to_proj_index.get(d, None)
            # 저장된 값이 없으면 조사차종 탭의 현재 선택값을 사용
            if proj_idx is None:
                try:
                    veh_tab = getattr(self.page, "veh", None)
                    if veh_tab is not None and hasattr(veh_tab, "cb"):
                        idx = veh_tab.cb.currentIndex()
                        if 0 <= idx < len(projects):
                            proj_idx = idx
                except Exception:
                    proj_idx = None
            if proj_idx is None:
                proj_idx = 0
            if not (0 <= proj_idx < len(projects)):
                proj_idx = 0
            proj_local = projects[proj_idx]
            vehicle_rows = proj_local.get("vehicle_set") or []
            cols = [v.get("차종명", "") for v in vehicle_rows]
            return cols or ["차종1", "차종2"]

        def _cols_for_direction(d):
            """해당 방향에 연결된 단축키 시트가 있으면 그 시트의 차종명을 사용하고,
            그렇지 않으면 group_projects/조사차종 설정에 따른 기본 차종명을 사용한다."""
            sheet_name = dir_to_sheet.get(d)
            if sheet_name and projects:
                # 모든 프로젝트에서 시트를 찾아본다.
                for proj_local in projects:
                    sheets = proj_local.get("hotkey_sheets_global") or []
                    for s in sheets:
                        if str(s.get("name", "")).strip() == sheet_name:
                            items = s.get("items") or []
                            cols = [str(it.get("차종명", "")) for it in items]
                            if cols:
                                return cols
            # 없으면 프로젝트 기본 차종명을 사용
            return _default_cols_for_direction(d)

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("시트 보기")

        lay = QtWidgets.QVBoxLayout(dlg)
        title_label = QtWidgets.QLabel("지점정보 시트 미리보기 (방향별)")
        title_label.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        font = title_label.font()
        font.setBold(True)
        title_label.setFont(font)
        lay.addWidget(title_label)

        tabs = QtWidgets.QTabWidget()
        lay.addWidget(tabs, 1)

        # 각 "입력그룹"이 아니라 실제 방향번호별로 시트를 구성
        for d in dir_list:
            cols = _cols_for_direction(d)
            tbl = QtWidgets.QTableWidget(len(times), 1 + len(cols) + 1)
            tbl.setHorizontalHeaderLabels(["시간대"] + cols + [""])
            tbl.horizontalHeader().setStretchLastSection(True)
            tbl.verticalHeader().setVisible(False)

            for i, row in enumerate(times):
                label = f"{row.get('시작','')}~{row.get('종료','')}"
                it = QtWidgets.QTableWidgetItem(label)
                it.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
                tbl.setItem(i, 0, it)

                for c in range(len(cols)):
                    num = QtWidgets.QTableWidgetItem("0")
                    num.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    tbl.setItem(i, 1 + c, num)

                blank = QtWidgets.QTableWidgetItem("")
                blank.setFlags(blank.flags() & ~QtCore.Qt.ItemIsEditable)
                tbl.setItem(i, 1 + len(cols), blank)

            tabs.addTab(tbl, f"{d}번 방향")

        btn = QtWidgets.QPushButton("닫기")
        lay.addWidget(btn, 0, alignment=QtCore.Qt.AlignRight)
        btn.clicked.connect(dlg.accept)
        dlg.resize(1100, 680)
        dlg.exec_()



class SurveyManagerPage(QtWidgets.QWidget):
    def __init__(self, data, parent=None):
        super().__init__(parent); self.data=data
        h=QtWidgets.QHBoxLayout(self)
        left=QtWidgets.QVBoxLayout(); self.lst=QtWidgets.QListWidget();
        self.lst.setFixedWidth(220)
        self.lst.setMinimumWidth(220)
        self.lst.setMaximumWidth(220)
        left.addWidget(self.lst,1)
        # Add header label above the list as requested ("과업명")
        try:
            left.insertWidget(0, QtWidgets.QLabel("과업명"))
        except Exception:
            pass
        hb=QtWidgets.QHBoxLayout(); self.b_new=QtWidgets.QPushButton("새조사"); self.b_dup=QtWidgets.QPushButton("복제"); self.b_del=QtWidgets.QPushButton("삭제")
        for b in (self.b_new,self.b_dup,self.b_del): hb.addWidget(b)
        left.addLayout(hb); h.addLayout(left,1)
        self.tabs=QtWidgets.QTabWidget()
        self.info=SurveyInfoTab(self); self.time=SurveyTimeTab(self); self.veh=SurveyVehicleTab(self); self.sites=SurveySitesTab(self)
        self.tabs.addTab(self.info,"조사 정보"); self.tabs.addTab(self.time,"조사시간 설정"); self.tabs.addTab(self.veh,"조사차종 설정"); self.tabs.addTab(self.sites,"조사지점 설정")
        box=QtWidgets.QVBoxLayout(); box.addWidget(self.tabs,1); b=QtWidgets.QHBoxLayout(); b.addStretch(1); self.b_reg=QtWidgets.QPushButton("등록"); self.b_save=QtWidgets.QPushButton("저장"); b.addWidget(self.b_reg); b.addWidget(self.b_save); box.addLayout(b); h.addLayout(box,3)
        self.b_new.clicked.connect(self.add_survey); self.b_dup.clicked.connect(self.dup_survey); self.b_del.clicked.connect(self.del_survey); self.b_save.clicked.connect(self.persist_current); self.lst.currentRowChanged.connect(self.load_current); self.lst.itemDoubleClicked.connect(self.rename_survey)
        self.reload_list()

    def surveys(self): return self.data["surveys"]
    def reload_list(self, preferred_index=None):
        self.lst.clear(); 
        for s in self.surveys():
            info = s.get("info", {})
            name = info.get("name","(제목 없음)")
            state = info.get("state","대기")
            item = QtWidgets.QListWidgetItem(name)
            # 과업명 리스트 진행상태 색상: 대기=검정, 진행=파랑, 완료=연회색
            color_map = {"대기": QtGui.QColor("black"), "진행": QtGui.QColor("blue"), "완료": QtGui.QColor("lightgray")}
            item.setForeground(color_map.get(state, QtGui.QColor("black")))
            self.lst.addItem(item)
        if self.lst.count():
            if preferred_index is None:
                preferred_index = 0
            preferred_index = max(0, min(preferred_index, self.lst.count()-1))
            self.lst.setCurrentRow(preferred_index)

    def current(self):
        i=self.lst.currentRow()
        return self.surveys()[i] if 0<=i<len(self.surveys()) else None

    def rename_survey(self, item=None):
        if item is None:
            item = self.lst.currentItem()
        if not item:
            return
        cur_index = self.lst.currentRow()
        if not (0 <= cur_index < len(self.surveys())):
            return
        info = self.surveys()[cur_index].setdefault("info", {})
        old_name = info.get("name", item.text())
        name, ok = QtWidgets.QInputDialog.getText(self, "수정", "이름:", text=old_name)
        if ok and name.strip():
            info["name"] = name.strip()
            save_data(self.data)
            self.reload_list(preferred_index=cur_index)

    def add_survey(self):
        d={"info":{"purpose":"일반 조사용(모든작업자 노출)","state":"대기","name":"새 조사","sn":next_sn(),"reg_date":dt.date.today().strftime("%Y-%m-%d"),"client":"", "period":[dt.date.today().strftime("%Y-%m-%d"),dt.date.today().strftime("%Y-%m-%d")], "desc":""}, "times":[], "vehicle":{}, "sites":[]}
        self.surveys().append(d); save_data(self.data); self.reload_list(); self.lst.setCurrentRow(self.lst.count()-1)

    def dup_survey(self):
        cur = self.current()
        if not cur:
            return
        import copy
        new_surv = copy.deepcopy(cur)

        # --- 작업번호 재부여 ---
        # env 전체에서 사용 중인 작업번호 수집
        nums = []
        data = self.data if isinstance(self.data, dict) else {}
        for surv in data.get("surveys", []):
            for row in surv.get("sites", []):
                jid = row.get("작업번호", "")
                if isinstance(jid, str) and jid.startswith("WN_"):
                    tail = jid[3:]
                    if tail.isdigit():
                        nums.append(int(tail))

        # 새 번호 시작값
        n = max(nums) + 1 if nums else 1

        # 복제된 과업의 지점들에 대해 새 작업번호 부여
        for row in new_surv.get("sites", []):
            row["작업번호"] = f"WN_{n:04d}"
            n += 1

        # env에 추가 및 저장
        self.surveys().append(new_surv)
        save_data(self.data)

        # 리스트 갱신 후, 방금 복제된 항목 선택
        self.reload_list(preferred_index=self.lst.count() - 1)



    def del_survey(self):
        i=self.lst.currentRow(); 
        if i<0: return
        if QtWidgets.QMessageBox.question(self,"삭제 확인","삭제할까요?")==QtWidgets.QMessageBox.Yes:
            self.surveys().pop(i); save_data(self.data); 
            preferred = max(0, min(i, len(self.surveys())-1))
            self.reload_list(preferred_index=preferred)

    def persist_current(self):
        i=self.lst.currentRow(); 
        if i<0: return
        d=self.surveys()[i]
        d["info"]=self.info.get(); d["times"]=self.time.get(); d["vehicle"]=self.veh.get(); d["sites"]=self.sites.get()
        save_data(self.data); self.reload_list()

    def load_current(self,*_):
        d=self.current()
        if not d:
            self.info.set({}); self.time.set([]); self.veh.set({}); self.sites.set([]); return
        self.info.set(d.get("info")); self.time.set(d.get("times")); self.veh.set(d.get("vehicle")); self.sites.set(d.get("sites"))

# ---------------- 메인 다이얼로그 ----------------


# ---------------- 사용자 관리 ----------------
class UserEditDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.setWindowTitle("사용자 편집")
        self._data = None
        layout = QtWidgets.QFormLayout(self)

        self.ed_name = QtWidgets.QLineEdit()
        self.ed_id = QtWidgets.QLineEdit()
        self.ed_pw = QtWidgets.QLineEdit()
        self.ed_pw.setEchoMode(QtWidgets.QLineEdit.Password)

        self.cb_role = QtWidgets.QComboBox()
        self.cb_role.addItems(["일반사용자", "관리자"])

        self.cb_status = QtWidgets.QComboBox()
        self.cb_status.addItems(["사용", "정지"])

        today = QtCore.QDate.currentDate()
        self.dt_reg = QtWidgets.QDateEdit(today)
        self.dt_reg.setCalendarPopup(True)
        self.dt_reg.setDisplayFormat("yyyy-MM-dd")

        self.dt_start = QtWidgets.QDateEdit(today)
        self.dt_start.setCalendarPopup(True)
        self.dt_start.setDisplayFormat("yyyy-MM-dd")

        self.dt_end = QtWidgets.QDateEdit(today.addYears(1))
        self.dt_end.setCalendarPopup(True)
        self.dt_end.setDisplayFormat("yyyy-MM-dd")

        self.ed_extra = QtWidgets.QLineEdit()

        layout.addRow("이름", self.ed_name)
        layout.addRow("아이디", self.ed_id)
        layout.addRow("비밀번호", self.ed_pw)
        layout.addRow("권한", self.cb_role)
        layout.addRow("등록일자", self.dt_reg)
        layout.addRow("상태", self.cb_status)
        layout.addRow("시작일", self.dt_start)
        layout.addRow("종료일", self.dt_end)
        layout.addRow("부가정보", self.ed_extra)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addRow(btns)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        if record:
            self.ed_name.setText(record.get("이름", ""))
            self.ed_id.setText(record.get("아이디", ""))
            self.ed_pw.setText(record.get("비밀번호", ""))
            role = record.get("권한") or "일반사용자"
            idx = self.cb_role.findText(role)
            if idx >= 0:
                self.cb_role.setCurrentIndex(idx)
            status = record.get("상태") or "사용"
            idx = self.cb_status.findText(status)
            if idx >= 0:
                self.cb_status.setCurrentIndex(idx)
            for widget, key in [
                (self.dt_reg, "등록일자"),
                (self.dt_start, "시작일"),
                (self.dt_end, "종료일"),
            ]:
                val = record.get(key)
                try:
                    if isinstance(val, str) and val:
                        d = QtCore.QDate.fromString(val, "yyyy-MM-dd")
                        if d.isValid():
                            widget.setDate(d)
                except Exception:
                    pass
            self.ed_extra.setText(record.get("부가정보", ""))

    def get_data(self):
        return self._data

    def accept(self):
        name = self.ed_name.text().strip()
        user_id = self.ed_id.text().strip()
        if not user_id:
            QtWidgets.QMessageBox.warning(self, "확인", "아이디를 입력해주세요.")
            return
        reg = self.dt_reg.date().toString("yyyy-MM-dd")
        start = self.dt_start.date().toString("yyyy-MM-dd")
        end = self.dt_end.date().toString("yyyy-MM-dd")
        self._data = {
            "이름": name,
            "아이디": user_id,
            "비밀번호": self.ed_pw.text(),
            "권한": self.cb_role.currentText(),
            "등록일자": reg,
            "상태": self.cb_status.currentText(),
            "시작일": start,
            "종료일": end,
            "부가정보": self.ed_extra.text().strip(),
        }
        super().accept()


class UserManagerPage(QtWidgets.QWidget):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.data = data
        v = QtWidgets.QVBoxLayout(self)

        lbl = QtWidgets.QLabel(
            "프로그램 사용자 계정을 관리합니다.\n"
            "- 권한이 '관리자'인 계정만 관리자 모드를 사용할 수 있습니다.\n"
            "- 상태가 '사용'인 계정만 로그인 가능합니다."
        )
        lbl.setWordWrap(True)
        v.addWidget(lbl)

        self.table = QtWidgets.QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(
            ["번호", "이름", "아이디", "비밀번호", "권한", "등록일자", "상태", "사용기간", "부가정보", ""]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        # 번호 열은 가로 폭을 더 작게, 그리고 다른 열과 독립적으로 조정
        try:
            header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
            self.table.setColumnWidth(0, 40)
        except Exception:
            pass
        # 행 번호(좌측 숫자)는 숨김
        try:
            self.table.verticalHeader().setVisible(False)
        except Exception:
            pass
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        v.addWidget(self.table, 1)

        hb = QtWidgets.QHBoxLayout()
        self.chk_all = QtWidgets.QCheckBox("전체보기")
        self.chk_all.setChecked(True)
        self.chk_all.toggled.connect(self.reload)
        hb.addWidget(self.chk_all)
        hb.addStretch(1)
        btn_add = QtWidgets.QPushButton("추가")
        btn_edit = QtWidgets.QPushButton("수정")
        btn_del = QtWidgets.QPushButton("삭제")
        hb.addWidget(btn_add)
        hb.addWidget(btn_edit)
        hb.addWidget(btn_del)
        v.addLayout(hb)

        btn_add.clicked.connect(self.add_user)
        btn_edit.clicked.connect(self.edit_user)
        btn_del.clicked.connect(self.del_user)

        self.row_map = []
        self.reload()

    def _users(self):
        return self.data.get("users", [])

    def reload(self):
        users = self._users()
        show_all = self.chk_all.isChecked()
        self.table.setRowCount(0)
        self.row_map = []

        for idx, rec in enumerate(users):
            if (not show_all) and rec.get("상태", "사용") != "사용":
                continue
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.row_map.append(idx)

            rec.setdefault("번호", idx + 1)
            num_item = QtWidgets.QTableWidgetItem(str(rec.get("번호", idx + 1)))
            num_item.setTextAlignment(QtCore.Qt.AlignCenter)
            num_item.setFlags(num_item.flags() & ~QtCore.Qt.ItemIsEditable)
            self.table.setItem(row, 0, num_item)

            def _item(text, align_center=False):
                it = QtWidgets.QTableWidgetItem(text or "")
                if align_center:
                    it.setTextAlignment(QtCore.Qt.AlignCenter)
                it.setFlags(it.flags() & ~QtCore.Qt.ItemIsEditable)
                return it

            self.table.setItem(row, 1, _item(rec.get("이름", ""), True))
            self.table.setItem(row, 2, _item(rec.get("아이디", ""), True))
            pw_display = "*****" if rec.get("비밀번호") else ""
            self.table.setItem(row, 3, _item(pw_display, True))
            self.table.setItem(row, 4, _item(rec.get("권한", ""), True))
            self.table.setItem(row, 5, _item(rec.get("등록일자", ""), True))
            self.table.setItem(row, 6, _item(rec.get("상태", ""), True))

            start = rec.get("시작일", "")
            end = rec.get("종료일", "")
            period = ""
            if start or end:
                period = f"{start} ~ {end}".strip()
            self.table.setItem(row, 7, _item(period, True))
            self.table.setItem(row, 8, _item(rec.get("부가정보", "")))
            self.table.setItem(row, 9, _item(""))

            # 관리자 행은 빨간색으로 강조
            if rec.get("권한") == "관리자":
                for col in range(1, 9):
                    item = self.table.item(row, col)
                    if item:
                        item.setForeground(QtGui.QBrush(QtGui.QColor("red")))

        self.table.resizeColumnsToContents()

    def current_index(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.row_map):
            return None
        return self.row_map[row]

    def add_user(self):
        dlg = UserEditDialog(self)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            if data is None:
                return
            users = self._users()
            users.append(data)
            self.data["users"] = users
            save_data(self.data)
            self.reload()

    def edit_user(self):
        idx = self.current_index()
        if idx is None:
            QtWidgets.QMessageBox.information(self, "알림", "수정할 사용자를 선택하세요.")
            return
        users = self._users()
        rec = dict(users[idx])
        dlg = UserEditDialog(self, rec)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            data = dlg.get_data()
            if data is None:
                return
            users[idx] = data
            self.data["users"] = users
            save_data(self.data)
            self.reload()
            # 다시 같은 행 선택
            if self.row_map:
                try:
                    row = self.row_map.index(idx)
                    self.table.selectRow(row)
                except ValueError:
                    pass

    def del_user(self):
        idx = self.current_index()
        if idx is None:
            QtWidgets.QMessageBox.information(self, "알림", "삭제할 사용자를 선택하세요.")
            return
        users = self._users()
        rec = users[idx]
        if QtWidgets.QMessageBox.question(
            self, "확인", f"선택한 사용자 '{rec.get('아이디','')}' 를 삭제하시겠습니까?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        ) != QtWidgets.QMessageBox.Yes:
            return
        del users[idx]
        # 번호 재정렬
        for i, u in enumerate(users):
            u["번호"] = i + 1
        self.data["users"] = users
        save_data(self.data)
        self.reload()

# ---------------- 메인 다이얼로그 ----------------


class Main(QtWidgets.QDialog):
    def __init__(self, data):
        super().__init__(); self.setWindowTitle(f"환경설정 (프로 플러스 – {APP_VER})"); self.resize(810,560)
        v=QtWidgets.QVBoxLayout(self); self.tabs=QtWidgets.QTabWidget(); v.addWidget(self.tabs,1)
        self.pg_work = WorkManagerPage(data, self); self.tabs.addTab(self.pg_work, "차종 관리")
        self.pg_survey = SurveyManagerPage(data, self); self.tabs.addTab(self.pg_survey, "조사 관리")
        self.pg_user = UserManagerPage(data, self); self.tabs.addTab(self.pg_user, "사용자 관리")
        hb=QtWidgets.QHBoxLayout(); hb.addStretch(1); b_save=QtWidgets.QPushButton("저장"); b_close=QtWidgets.QPushButton("닫기"); hb.addWidget(b_save); hb.addWidget(b_close); v.addLayout(hb)
        b_save.clicked.connect(lambda *_:(
            getattr(self.pg_survey, "persist_current", lambda: None)(),
            save_data(data),
            QtWidgets.QMessageBox.information(self,"저장","저장되었습니다.")
        ))
        b_close.clicked.connect(self.close)
        self.pg_work.changed.connect(lambda *_: self.pg_survey.veh.refresh_combo())

def main():
    app=QtWidgets.QApplication(sys.argv)

    # ----- Global 80% scale (fonts, paddings) -----
    scale = 0.8
    font = app.font()
    fs = font.pointSizeF() if hasattr(font, "pointSizeF") else float(font.pointSize())
    if fs <= 0:
        fs = float(font.pointSize())
    if fs > 0:
        font.setPointSizeF(fs * scale)
        app.setFont(font)
    # Optional: lighten paddings to match scale visually
    app.setStyleSheet(f"""
        * {{ font-size: {max(7, int(10*scale))}pt; }}
        QTabBar::tab {{ padding: {int(8*scale)}px {int(12*scale)}px; }}
        QHeaderView::section {{ padding: {int(6*scale)}px; }}
        QPushButton {{ padding: {int(6*scale)}px {int(10*scale)}px; }}
        QAbstractItemView {{ font-size: {max(7, int(10*scale))}pt; }}
    """)
# -----------------------------------------------
    data=load_data()
    w=Main(data); w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()