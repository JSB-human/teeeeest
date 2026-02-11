"""
한글(HWP) 문서를 제어하기 위한 컨트롤러 모듈
win32com을 이용하여 한글 프로그램을 자동화합니다.
"""

import os
import logging
import win32com.client
import win32gui
import win32con
import time
import pythoncom
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any, Tuple, Union

logger = logging.getLogger("hwp-controller")


class HwpController:
    """한글 문서를 제어하는 클래스"""

    def __init__(self):
        """한글 애플리케이션 인스턴스를 초기화합니다."""
        self.hwp: Any = None
        self.visible = True
        self.is_hwp_running = False
        self.current_document_path = None

    def connect(
        self, visible: bool = True, register_security_module: bool = True
    ) -> bool:
        """
        한글 프로그램에 연결합니다.

        Args:
            visible (bool): 한글 창을 화면에 표시할지 여부
            register_security_module (bool): 보안 모듈을 등록할지 여부

        Returns:
            bool: 연결 성공 여부
        """
        try:
            # COM 초기화
            pythoncom.CoInitialize()

            # GetActiveObject 시도
            try:
                self.hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
                logger.info("GetActiveObject 성공 - 기존 HWP 인스턴스에 연결됨")
            except Exception as e:
                logger.warning(f"GetActiveObject 실패: {e}")
                # Dispatch는 새 창을 열 수 있음 - HWP의 한계
                self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                logger.info("Dispatch로 HWP에 연결됨 (새 창이 열렸을 수 있음)")

            # 보안 모듈 등록 (파일 경로 체크 보안 경고창 방지)
            if register_security_module:
                try:
                    # 보안 모듈 DLL 경로 - 프로젝트 루트 기준 상대 경로로 설정
                    current_file_dir = os.path.dirname(os.path.abspath(__file__))
                    project_root = os.path.dirname(os.path.dirname(current_file_dir))
                    module_path = os.path.join(
                        project_root,
                        "security_module",
                        "FilePathCheckerModuleExample.dll",
                    )

                    if os.path.exists(module_path):
                        self.hwp.RegisterModule(
                            "FilePathCheckerModuleExample", module_path
                        )
                        logger.info(f"보안 모듈이 등록되었습니다: {module_path}")
                    else:
                        logger.warning(
                            f"보안 모듈 파일을 찾을 수 없습니다: {module_path}"
                        )
                except Exception as e:
                    logger.error(f"보안 모듈 등록 실패 (무시하고 계속 진행): {e}")

            self.visible = visible
            try:
                self.hwp.XHwpWindows.Item(0).Visible = visible
            except Exception as e:
                logger.warning(f"Visible 설정 실패: {e}")

            self.is_hwp_running = True
            return True
        except Exception as e:
            logger.error(f"한글 프로그램 연결 실패: {e}")
            return False

    def disconnect(self) -> bool:
        """
        한글 프로그램 연결을 종료합니다.

        Returns:
            bool: 종료 성공 여부
        """
        try:
            if self.is_hwp_running:
                # HwpObject를 해제합니다
                self.hwp = None
                self.is_hwp_running = False

            return True
        except Exception as e:
            logger.error(f"한글 프로그램 종료 실패: {e}")
            return False

    def set_message_box_mode(self, mode: int = 0x00020000) -> bool:
        """
        메시지 박스 표시 모드를 설정합니다.

        Args:
            mode (int): 메시지 박스 모드
                - 0x00000000: 기본값 (모든 메시지 박스 표시)
                - 0x00010000: 메시지 박스 표시 안함 (확인 버튼 자동 클릭)
                - 0x00020000: 메시지 박스 표시 안함 (취소 버튼 자동 클릭)
                - 0x00100000: 메시지 박스 표시 안함 (저장 안함 선택)

        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running or self.hwp is None:
                return False
            self.hwp.SetMessageBoxMode(mode)
            return True
        except Exception as e:
            logger.error(f"메시지 박스 모드 설정 실패: {e}")
            return False

    def close_document(self, save: bool = False, suppress_dialog: bool = True) -> bool:
        """
        현재 문서를 닫습니다.

        Args:
            save (bool): 저장 후 닫을지 여부
            suppress_dialog (bool): 저장 확인 대화상자 표시 안함 (기본값: True)

        Returns:
            bool: 닫기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 대화상자 표시 안함 설정
            if suppress_dialog:
                if save:
                    # 저장하고 닫기: 확인 버튼 자동 클릭
                    self.hwp.SetMessageBoxMode(0x00010000)
                else:
                    # 저장 안하고 닫기: 저장 안함 선택
                    self.hwp.SetMessageBoxMode(0x00100000)

            if save:
                self.hwp.HAction.Run("FileSave")

            result = self.hwp.HAction.Run("FileClose")
            self.current_document_path = None

            # 메시지 박스 모드 복원
            if suppress_dialog:
                self.hwp.SetMessageBoxMode(0x00000000)

            return bool(result)
        except Exception as e:
            logger.error(f"문서 닫기 실패: {e}")
            # 메시지 박스 모드 복원 시도
            try:
                self.hwp.SetMessageBoxMode(0x00000000)
            except Exception as e_restore:
                logger.debug(f"SetMessageBoxMode 복원 실패 (무시): {e_restore}")
            return False

    def close_all_documents(
        self, save: bool = False, suppress_dialog: bool = True
    ) -> bool:
        """
        모든 문서를 닫습니다.

        Args:
            save (bool): 저장 후 닫을지 여부
            suppress_dialog (bool): 저장 확인 대화상자 표시 안함 (기본값: True)

        Returns:
            bool: 닫기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 대화상자 표시 안함 설정
            if suppress_dialog:
                if save:
                    self.hwp.SetMessageBoxMode(0x00010000)
                else:
                    self.hwp.SetMessageBoxMode(0x00100000)

            if save:
                self.hwp.HAction.Run("FileSaveAll")

            result = self.hwp.HAction.Run("FileCloseAll")
            self.current_document_path = None

            # 메시지 박스 모드 복원
            if suppress_dialog:
                self.hwp.SetMessageBoxMode(0x00000000)

            return bool(result)
        except Exception as e:
            logger.error(f"모든 문서 닫기 실패: {e}")
            try:
                self.hwp.SetMessageBoxMode(0x00000000)
            except Exception as e_restore:
                logger.debug(f"SetMessageBoxMode 복원 실패 (무시): {e_restore}")
            return False

    def create_new_document(self) -> bool:
        """
        새 문서를 생성합니다.

        Returns:
            bool: 생성 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()

            self.hwp.Run("FileNew")
            self.current_document_path = None
            return True
        except Exception as e:
            logger.error(f"새 문서 생성 실패: {e}")
            return False

    def get_open_documents(self) -> Tuple[bool, List[Dict[str, Any]]]:
        """
        열려있는 문서 목록을 반환합니다.

        Returns:
            Tuple[bool, List[Dict]]: (성공 여부, 문서 목록)
            각 문서는 {"index": int, "path": str, "is_current": bool} 형태
        """
        try:
            if not self.is_hwp_running:
                return False, []

            documents = []

            # XHwpWindows를 사용하여 열린 윈도우(문서) 목록 가져오기
            try:
                windows = self.hwp.XHwpWindows
                window_count = windows.Count
            except Exception as e:
                logger.debug(f"XHwpWindows 접근 실패, XHwpDocuments로 폴백: {e}")
                window_count = self.hwp.XHwpDocuments.Count
                windows = self.hwp.XHwpDocuments

            # 현재 활성 문서 인덱스
            current_idx = None
            try:
                current_idx = self.hwp.CurDocIndex
            except Exception as e:
                logger.debug(f"CurDocIndex 조회 실패 (무시): {e}")

            for i in range(window_count):
                try:
                    doc = windows.Item(i)
                    doc_path = ""
                    try:
                        doc_path = doc.Path if doc.Path else "(새 문서)"
                    except Exception as e:
                        logger.debug(f"문서 경로 조회 실패: {e}")
                        doc_path = "(새 문서)"

                    is_current = (
                        (i == current_idx) if current_idx is not None else (i == 0)
                    )
                    documents.append(
                        {"index": i, "path": doc_path, "is_current": is_current}
                    )
                except Exception as e:
                    documents.append(
                        {"index": i, "path": f"(오류: {e})", "is_current": False}
                    )

            return True, documents
        except Exception as e:
            logger.error(f"문서 목록 조회 실패: {e}")
            return False, []

    def switch_document(self, index: int) -> Tuple[bool, str]:
        """
        특정 인덱스의 문서로 전환합니다.

        Args:
            index (int): 문서 인덱스

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            doc_count = self.hwp.XHwpDocuments.Count
            if index < 0 or index >= doc_count:
                return False, f"유효하지 않은 인덱스입니다. (0~{doc_count - 1})"

            doc = self.hwp.XHwpDocuments.Item(index)

            # 여러 방법 시도
            try:
                doc.SetActive_OnlyStrongHold()
            except Exception as e1:
                logger.debug(f"SetActive_OnlyStrongHold 실패: {e1}")
                try:
                    doc.SetActive()
                except Exception as e2:
                    logger.debug(f"SetActive 실패, HAction 사용: {e2}")
                    self.hwp.HAction.Run("MoveDocBegin")
                    for _ in range(index):
                        self.hwp.HAction.Run("WindowNext")

            doc_path = doc.Path if doc.Path else "(새 문서)"
            return True, f"문서 전환 완료: {doc_path}"
        except Exception as e:
            return False, f"문서 전환 실패: {e}"

    def get_all_hwp_instances(self) -> Tuple[bool, List[Dict[str, Any]]]:
        """
        Running Object Table에서 모든 HWP 인스턴스를 찾습니다.

        Returns:
            Tuple[bool, List[Dict]]: (성공 여부, 인스턴스 목록)
            각 인스턴스는 {"index": int, "hwnd": int, "title": str, "is_current": bool} 형태
        """
        try:
            instances = []
            current_hwnd = None

            # 현재 연결된 HWP의 윈도우 핸들
            if self.hwp:
                try:
                    current_hwnd = self.hwp.XHwpWindows.Item(0).WindowHandle
                except Exception as e:
                    logger.debug(f"현재 WindowHandle 조회 실패 (무시): {e}")

            # 모든 HWP 윈도우 찾기
            def enum_hwp_windows(hwnd, results):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "HwpFrame" or (
                        class_name is not None and "Hwp" in class_name
                    ):
                        title = win32gui.GetWindowText(hwnd)
                        if title:  # 제목이 있는 창만
                            results.append(
                                {"hwnd": hwnd, "title": title, "class": class_name}
                            )
                except Exception as e:
                    logger.debug(f"창 정보 조회 실패 hwnd={hwnd}: {e}")
                return True

            hwp_windows = []
            win32gui.EnumWindows(enum_hwp_windows, hwp_windows)

            for i, win in enumerate(hwp_windows):
                instances.append(
                    {
                        "index": i,
                        "hwnd": win["hwnd"],
                        "title": win["title"],
                        "is_current": win["hwnd"] == current_hwnd
                        if current_hwnd
                        else False,
                    }
                )

            return True, instances
        except Exception as e:
            logger.error(f"HWP 인스턴스 목록 조회 실패: {e}")
            return False, []

    def connect_to_hwp_instance(self, hwnd: int) -> Tuple[bool, str]:
        """
        특정 HWP 윈도우에 연결합니다.

        Args:
            hwnd: 윈도우 핸들

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            title = win32gui.GetWindowText(hwnd)

            # 기존 연결 해제
            self.hwp = None
            self.is_hwp_running = False

            # 해당 윈도우를 최상위로 가져오기
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                return False, f"창 활성화 실패: {e}"

            time.sleep(0.5)

            # COM 재초기화
            try:
                pythoncom.CoInitialize()
            except Exception as e:
                logger.debug(f"CoInitialize: {e}")  # 이미 초기화된 경우

            # 방법 1: GetActiveObject 시도
            try:
                self.hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
                self.is_hwp_running = True
                logger.info(f"GetActiveObject 성공: {title}")
                return True, f"HWP 인스턴스에 연결됨: {title}"
            except Exception as e:
                logger.warning(f"GetActiveObject 실패: {e}")

            # 방법 2: Dispatch로 연결 (활성화된 HWP에 연결됨)
            try:
                self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                self.is_hwp_running = True
                logger.info(f"Dispatch로 연결됨")
                # Dispatch 후 현재 문서 경로로 확인
                try:
                    current_path = self.hwp.Path
                    return (
                        True,
                        f"HWP에 연결됨: {title} (문서: {current_path or '새 문서'})",
                    )
                except Exception as e:
                    logger.debug(f"Path 가져오기 실패: {e}")
                    return True, f"HWP에 연결됨: {title}"
            except Exception as e:
                logger.error(f"Dispatch 실패: {e}")
                return False, f"연결 실패: {e}"

        except Exception as e:
            return False, f"연결 실패: {e}"

    def close_hwp_window(self, hwnd: int) -> Tuple[bool, str]:
        """
        HWP 윈도우를 닫습니다 (WM_CLOSE 메시지 전송).

        Args:
            hwnd: 윈도우 핸들

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            title = win32gui.GetWindowText(hwnd)
            WM_CLOSE = 0x0010
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            return True, f"창 닫기 요청: {title}"
        except Exception as e:
            return False, f"창 닫기 실패: {e}"

    def open_document(self, file_path: str) -> bool:
        """
        문서를 엽니다.

        Args:
            file_path (str): 열 문서의 경로

        Returns:
            bool: 열기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()

            abs_path = os.path.abspath(file_path)
            logger.debug(f"Opening document: {abs_path}")
            logger.debug(f"File exists: {os.path.exists(abs_path)}")

            # Use HAction with FileOpen for reliable file opening
            pset = self.hwp.HParameterSet.HFileOpenSave
            self.hwp.HAction.GetDefault("FileOpen", pset.HSet)
            pset.filename = abs_path
            pset.Format = "HWP"
            result = self.hwp.HAction.Execute("FileOpen", pset.HSet)
            logger.debug(f"FileOpen result: {result}")
            if result:
                self.current_document_path = abs_path
            return result
        except Exception as e:
            logger.error(f"문서 열기 실패: {e}")
            import traceback

            logger.error(traceback.format_exc())
            return False

    def save_document(self, file_path: Optional[str] = None) -> bool:
        """
        문서를 저장합니다.

        Args:
            file_path (str, optional): 저장할 경로. None이면 현재 경로에 저장.

        Returns:
            bool: 저장 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            if file_path:
                abs_path = os.path.abspath(file_path)
                # 파일 형식과 경로 모두 지정하여 저장
                self.hwp.SaveAs(abs_path, "HWP", "")
                self.current_document_path = abs_path
            else:
                if self.current_document_path:
                    self.hwp.Save()
                else:
                    # 저장 대화 상자 표시 (파라미터 없이 호출)
                    self.hwp.SaveAs()
                    # 대화 상자에서 사용자가 선택한 경로를 알 수 없으므로 None 유지

            return True
        except Exception as e:
            logger.error(f"문서 저장 실패: {e}")
            return False

    def insert_text(self, text: str, preserve_linebreaks: bool = True) -> bool:
        """
        현재 커서 위치에 텍스트를 삽입합니다.

        Args:
            text (str): 삽입할 텍스트
            preserve_linebreaks (bool): 줄바꿈 유지 여부

        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            if preserve_linebreaks and "\n" in text:
                # 줄바꿈이 포함된 경우 줄 단위로 처리
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    if i > 0:  # 첫 줄이 아니면 줄바꿈 추가
                        self.insert_paragraph()
                    if line.strip():  # 빈 줄이 아니면 텍스트 삽입
                        self._insert_text_direct(line)
                return True
            else:
                # 줄바꿈이 없거나 유지하지 않는 경우 한 번에 처리
                return self._insert_text_direct(text)
        except Exception as e:
            logger.error(f"텍스트 삽입 실패: {e}")
            return False

    def _set_table_cursor(self) -> bool:
        """
        표 안에서 커서 위치를 제어하는 내부 메서드입니다.
        현재 셀을 선택하고 취소하여 커서를 셀 안에 위치시킵니다.

        Returns:
            bool: 성공 여부
        """
        try:
            # 현재 셀 선택
            self.hwp.Run("TableSelCell")
            # 선택 취소 (커서는 셀 안에 위치)
            self.hwp.Run("Cancel")
            # 셀 내부로 커서 이동을 확실히
            self.hwp.Run("CharRight")
            self.hwp.Run("CharLeft")
            return True
        except Exception as e:
            logger.debug(f"셀 내부 커서 이동 실패: {e}")
            return False

    def _insert_text_direct(self, text: str) -> bool:
        """
        텍스트를 직접 삽입하는 내부 메서드입니다.

        Args:
            text (str): 삽입할 텍스트

        Returns:
            bool: 삽입 성공 여부
        """
        try:
            # 텍스트 삽입을 위한 액션 초기화
            self.hwp.HAction.GetDefault(
                "InsertText", self.hwp.HParameterSet.HInsertText.HSet
            )
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute(
                "InsertText", self.hwp.HParameterSet.HInsertText.HSet
            )
            return True
        except Exception as e:
            logger.error(f"텍스트 직접 삽입 실패: {e}")
            return False

    def set_font(
        self,
        font_name: Optional[str],
        font_size: Optional[int],
        bold: bool = False,
        italic: bool = False,
        select_previous_text: bool = False,
    ) -> bool:
        """
        글꼴 속성을 설정합니다. 현재 위치에서 다음에 입력할 텍스트에 적용됩니다.

        Args:
            font_name (str): 글꼴 이름
            font_size (int): 글꼴 크기
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부

        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 새로운 구현: set_font_style 메서드 사용
            return self.set_font_style(
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=False,
                select_previous_text=select_previous_text,
            )
        except Exception as e:
            logger.error(f"글꼴 설정 실패: {e}")
            return False

    def set_font_style(
        self,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        strike: bool = False,
        color: int = 0,
        select_previous_text: bool = False,
    ) -> bool:
        """
        현재 선택된 텍스트의 글꼴 스타일을 설정합니다.

        Args:
            color: 글자 색상 (0xBBGGRR 형식, 예: 빨강 0x0000FF, 초록 0x00FF00)
            strike: 취소선 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            if select_previous_text:
                self.select_last_text()

            self.hwp.HAction.GetDefault(
                "CharShape", self.hwp.HParameterSet.HCharShape.HSet
            )

            if font_name:
                self.hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameLatin = font_name

            if font_size:
                self.hwp.HParameterSet.HCharShape.Height = font_size * 100

            self.hwp.HParameterSet.HCharShape.Bold = bold
            self.hwp.HParameterSet.HCharShape.Italic = italic
            self.hwp.HParameterSet.HCharShape.UnderlineType = 1 if underline else 0
            self.hwp.HParameterSet.HCharShape.StrikeOutType = 1 if strike else 0
            self.hwp.HParameterSet.HCharShape.TextColor = color

            self.hwp.HAction.Execute(
                "CharShape", self.hwp.HParameterSet.HCharShape.HSet
            )
            return True
        except Exception as e:
            logger.error(f"글꼴 스타일 설정 실패: {e}")
            return False

    def clear_cell_content(self) -> bool:
        """현재 셀의 내용을 완전히 삭제합니다."""
        try:
            self.hwp.Run("TableSelCell")
            self.hwp.HAction.Run("SelectAll")
            try:
                self.hwp.HAction.Run("EditCut")
            except Exception:
                self.hwp.Run("Delete")
            self.hwp.Run("Cancel")
            self._set_table_cursor()
            return True
        except Exception as e:
            logger.error(f"셀 내용 삭제 실패: {e}")
            return False

    def insert_diff_text(self, old_text: str, new_text: str) -> bool:
        """한 셀 안에 이전 텍스트(빨간 취소선)와 새 텍스트(초록 굵게)를 함께 삽입합니다."""
        try:
            if not old_text or old_text == "(빈 셀)":
                old_text = ""

            # 1. 이전 텍스트 (빨강 + 취소선)
            if old_text:
                self.set_font_style(color=0x0000FF, strike=True, bold=False)
                self._insert_text_direct(old_text)
                self.set_font_style(color=0, strike=False, bold=False)
                self._insert_text_direct(" → ")

            # 2. 새 텍스트 (초록 + 굵게)
            self.set_font_style(color=0x00FF00, strike=False, bold=True)
            self._insert_text_direct(new_text)

            # 3. 스타일 복구
            self.set_font_style(color=0, strike=False, bold=False)
            return True
        except Exception as e:
            logger.error(f"Diff 텍스트 삽입 실패: {e}")
            return False

    def _get_current_position(self):
        """현재 커서(캐럿) 위치 정보를 가져옵니다.

        현재 연결된 HWP 인스턴스에서 GetPos()를 호출해
        (position_type, list_id, para_id, char_pos)의 형태 또는
        (list_id, para_id, char_pos) 형태의 값을 반환합니다.
        """
        try:
            if not self.is_hwp_running or not self.hwp:
                logger.debug("HWP가 실행 중이 아니거나 객체가 없습니다. (GetPos 호출 불가)")
                return None
            
            # COM 객체 유효성 체크를 위해 간단한 속성 조회
            try:
                _ = self.hwp.Path
            except Exception:
                logger.warning("HWP COM 객체 연결이 끊긴 것 같습니다. 재연결을 시도하세요.")
                return None

            pos = self.hwp.GetPos()
            # 디버깅용: 실제 반환값 확인
            logger.debug(f"RAW GetPos: {pos}")
            return pos

        except Exception as e:
            logger.error(f"GetPos() 호출 중 치명적 오류: {e}")
            return None

    def _set_position(self, pos):
        """커서 위치를 지정된 위치로 변경합니다."""
        try:
            if pos:
                self.hwp.SetPos(*pos)
            return True
        except Exception as e:
            logger.debug(f"SetPos 실패: {e}")
            return False

    def get_cursor_pos(self):
        """현재 커서 위치를 딕셔너리 형태로 반환합니다.

        Returns:
            dict | None: {"list_id": int, "para_id": int, "char_pos": int}
        """
        try:
            pos = self._get_current_position()
            if pos is None:
                return None

            # GetPos 반환값이 3개(리스트, 문단, 문자오프셋)일 수도 있고
            # 4개(위치 유형 포함)일 수도 있으므로, 항상 마지막 3개를 사용
            list_id, para_id, char_pos = pos[-3:]

            logger.debug(
                f"get_cursor_pos RAW POS: {pos} => {list_id}, {para_id}, {char_pos}"
            )

            return {
                "list_id": list_id,
                "para_id": para_id,
                "char_pos": char_pos,
            }
        except Exception as e:
            logger.debug(f"get_cursor_pos 실패: {e}")
            return None

    def insert_table(self, rows: int, cols: int) -> bool:
        """
        현재 커서 위치에 표를 삽입합니다.

        Args:
            rows (int): 행 수
            cols (int): 열 수

        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            self.hwp.HAction.GetDefault(
                "TableCreate", self.hwp.HParameterSet.HTableCreation.HSet
            )
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = (
                0  # 0: 단에 맞춤, 1: 절대값
            )
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 0: 자동, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.WidthValue = (
                0  # 단에 맞춤이므로 무시됨
            )
            self.hwp.HParameterSet.HTableCreation.HeightValue = 1000  # 셀 높이(hwpunit)

            # 각 열의 너비를 설정 (모두 동일하게)
            # PageWidth 대신 고정 값 사용
            col_width = 8000 // cols  # 전체 너비를 열 수로 나눔
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_width)

            self.hwp.HAction.Execute(
                "TableCreate", self.hwp.HParameterSet.HTableCreation.HSet
            )
            return True
        except Exception as e:
            logger.error(f"표 삽입 실패: {e}")
            return False

    def insert_image(self, image_path: str, width: int = 0, height: int = 0) -> bool:
        """
        현재 커서 위치에 이미지를 삽입합니다.

        Args:
            image_path (str): 이미지 파일 경로
            width (int): 이미지 너비(0이면 원본 크기)
            height (int): 이미지 높이(0이면 원본 크기)

        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            abs_path = os.path.abspath(image_path)
            if not os.path.exists(abs_path):
                logger.error(f"이미지 파일을 찾을 수 없습니다: {abs_path}")
                return False

            self.hwp.HAction.GetDefault(
                "InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet
            )
            self.hwp.HParameterSet.HInsertPicture.FileName = abs_path
            self.hwp.HParameterSet.HInsertPicture.Width = width
            self.hwp.HParameterSet.HInsertPicture.Height = height
            self.hwp.HParameterSet.HInsertPicture.Embed = 1  # 0: 링크, 1: 파일 포함
            self.hwp.HAction.Execute(
                "InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet
            )
            return True
        except Exception as e:
            logger.error(f"이미지 삽입 실패: {e}")
            return False

    def undo(self, count: int = 1) -> Tuple[bool, str]:
        """
        실행 취소(Undo)를 수행합니다.

        Args:
            count (int): 취소할 횟수 (기본값: 1)

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            success_count = 0
            for _ in range(count):
                result = self.hwp.HAction.Run("Undo")
                if result:
                    success_count += 1
                else:
                    break

            if success_count == count:
                return True, f"실행 취소 {success_count}회 완료"
            elif success_count > 0:
                return True, f"실행 취소 {success_count}회 완료 (요청: {count}회)"
            else:
                return False, "실행 취소할 항목이 없습니다."
        except Exception as e:
            return False, f"실행 취소 실패: {e}"

    def redo(self, count: int = 1) -> Tuple[bool, str]:
        """
        다시 실행(Redo)을 수행합니다.

        Args:
            count (int): 다시 실행할 횟수 (기본값: 1)

        Returns:
            Tuple[bool, str]: (성공 여부, 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 실행되지 않았습니다."

            success_count = 0
            for _ in range(count):
                result = self.hwp.HAction.Run("Redo")
                if result:
                    success_count += 1
                else:
                    break

            if success_count == count:
                return True, f"다시 실행 {success_count}회 완료"
            elif success_count > 0:
                return True, f"다시 실행 {success_count}회 완료 (요청: {count}회)"
            else:
                return False, "다시 실행할 항목이 없습니다."
        except Exception as e:
            return False, f"다시 실행 실패: {e}"

    def find_text(self, text: str) -> bool:
        """
        문서에서 텍스트를 찾습니다.

        Args:
            text (str): 찾을 텍스트

        Returns:
            bool: 찾기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # HAction으로 찾기
            pset = self.hwp.HParameterSet.HFindReplace
            self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = text
            pset.FindRegExp = 0
            pset.IgnoreMessage = 1
            pset.Direction = 0  # 0: forward
            result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)
            return bool(result)
        except Exception as e:
            logger.error(f"텍스트 찾기 실패: {e}")
            return False

    def replace_text(
        self, find_text: str, replace_text: str, replace_all: bool = True
    ) -> bool:
        """
        문서에서 텍스트를 찾아 바꿉니다.

        Args:
            find_text (str): 찾을 텍스트
            replace_text (str): 바꿀 텍스트
            replace_all (bool): 모두 바꾸기 여부

        Returns:
            bool: 바꾸기 성공 여부 (예외가 없으면 성공으로 간주)
        """
        try:
            if not self.is_hwp_running:
                return False

            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            pset = self.hwp.HParameterSet.HFindReplace
            self.hwp.HAction.GetDefault("AllReplace", pset.HSet)
            pset.FindString = find_text
            pset.ReplaceString = replace_text
            pset.FindRegExp = 0
            pset.IgnoreMessage = 1

            # Note: HWP COM API의 AllReplace는 성공해도 False를 반환함
            # 예외가 발생하지 않으면 성공으로 간주
            self.hwp.HAction.Execute("AllReplace", pset.HSet)
            return True
        except Exception as e:
            logger.error(f"텍스트 바꾸기 실패: {e}")
            return False

    def get_text(self) -> str:
        """현재 문서의 전체 텍스트를 가져옵니다.

        1차 시도: HWP COM API의 GetTextFile("TEXT", "") 사용
        2차 시도: 실패 시 전체 선택 + 클립보드를 통해 텍스트 추출 (fallback)

        Returns:
            str: 문서 텍스트 (실패 시 빈 문자열)
        """
        try:
            if not self.is_hwp_running:
                return ""

            # 1차 시도: 표준 API
            try:
                text = self.hwp.GetTextFile("TEXT", "")
                if text:
                    return text
            except Exception as e:
                logger.debug(f"텍스트 가져오기 실패(GetTextFile): {e}")

            # 2차 시도: 클립보드 기반 fallback (윈도우에 실제 키 입력 보내기)
            try:
                import win32clipboard
                import win32api
                import win32con

                # 한글 창을 전면으로 가져오기
                try:
                    hwnd = self.hwp.XHwpWindows.Item(0).WindowHandle
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.2)
                except Exception as e_hwnd:
                    logger.debug(f"윈도우 활성화 실패(무시하고 진행): {e_hwnd}")

                # Ctrl+A, Ctrl+C 실제 키 이벤트 전송
                def send_ctrl_combo(vk: int):
                    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                    win32api.keybd_event(vk, 0, 0, 0)
                    time.sleep(0.05)
                    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(
                        win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0
                    )

                # 전체 선택
                send_ctrl_combo(ord("A"))
                time.sleep(0.1)
                # 복사
                send_ctrl_combo(ord("C"))
                time.sleep(0.1)

                # 클립보드에서 텍스트 읽기
                win32clipboard.OpenClipboard()
                try:
                    text = win32clipboard.GetClipboardData(
                        win32clipboard.CF_UNICODETEXT
                    )
                finally:
                    win32clipboard.CloseClipboard()

                return text or ""
            except Exception as e_fb:
                logger.error(f"클립보드 기반 텍스트 추출 실패: {e_fb}")
                return ""
        except Exception as e:
            logger.error(f"텍스트 가져오기 실패(알 수 없는 오류): {e}")
            return ""

    def set_page_setup(
        self,
        orientation: str = "portrait",
        margin_left: int = 1000,
        margin_right: int = 1000,
        margin_top: int = 1000,
        margin_bottom: int = 1000,
    ) -> bool:
        """
        페이지 설정을 변경합니다.

        Args:
            orientation (str): 용지 방향 ('portrait' 또는 'landscape')
            margin_left (int): 왼쪽 여백(hwpunit)
            margin_right (int): 오른쪽 여백(hwpunit)
            margin_top (int): 위쪽 여백(hwpunit)
            margin_bottom (int): 아래쪽 여백(hwpunit)

        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 매크로 명령 사용
            orient_val = 0 if orientation.lower() == "portrait" else 1

            # 페이지 설정 매크로
            result = self.hwp.Run(
                f"PageSetup3 {orient_val} {margin_left} {margin_right} {margin_top} {margin_bottom}"
            )
            return bool(result)
        except Exception as e:
            logger.error(f"페이지 설정 실패: {e}")
            return False

    def insert_paragraph(self) -> bool:
        """
        새 단락을 삽입합니다.

        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            self.hwp.HAction.Run("BreakPara")
            return True
        except Exception as e:
            logger.error(f"단락 삽입 실패: {e}")
            return False

    def select_all(self) -> bool:
        """
        문서 전체를 선택합니다.

        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            self.hwp.Run("SelectAll")
            return True
        except Exception as e:
            logger.error(f"전체 선택 실패: {e}")
            return False

    def fill_cell_field(self, field_name: str, value: str, n: int = 1) -> bool:
        """
        동일한 이름의 셀필드 중 n번째에만 값을 채웁니다.
        위키독스 예제: https://wikidocs.net/261646

        Args:
            field_name (str): 필드 이름
            value (str): 채울 값
            n (int): 몇 번째 필드에 값을 채울지 (1부터 시작)

        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 1. 필드 목록 가져오기
            # HGO_GetFieldList은 현재 문서에 있는 모든 필드 목록을 가져옵니다.
            self.hwp.HAction.GetDefault(
                "HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet
            )
            self.hwp.HAction.Execute(
                "HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet
            )

            # 2. 필드 이름이 동일한 모든 셀필드 찾기
            field_list = []
            field_count = self.hwp.HParameterSet.HGo.FieldList.Count

            for i in range(field_count):
                field_info = self.hwp.HParameterSet.HGo.FieldList.Item(i)
                if field_info.FieldName == field_name:
                    field_list.append((field_info.FieldName, i))

            # 3. n번째 필드가 존재하는지 확인 (인덱스는 0부터 시작하므로 n-1)
            if len(field_list) < n:
                logger.warning(
                    f"해당 이름의 필드가 충분히 없습니다. 필요: {n}, 존재: {len(field_list)}"
                )
                return False

            # 4. n번째 필드의 위치로 이동
            target_field_idx = field_list[n - 1][1]

            # HGo_SetFieldText를 사용하여 해당 필드 위치로 이동한 후 텍스트 설정
            self.hwp.HAction.GetDefault(
                "HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet
            )
            self.hwp.HParameterSet.HGo.HSet.SetItem("FieldIdx", target_field_idx)
            self.hwp.HParameterSet.HGo.HSet.SetItem("Text", value)
            self.hwp.HAction.Execute(
                "HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet
            )

            return True
        except Exception as e:
            logger.error(f"셀필드 값 채우기 실패: {e}")
            return False

    def select_last_text(self) -> bool:
        """
        현재 단락의 마지막으로 입력된 텍스트를 선택합니다.

        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 현재 위치 저장
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return False

            # 현재 단락의 시작으로 이동
            self.hwp.Run("MoveLineStart")
            start_pos = self.hwp.GetPos()

            # 이전 위치로 돌아가서 선택 영역 생성
            self.hwp.SetPos(*start_pos)
            self.hwp.SelectText(start_pos, current_pos)

            return True
        except Exception as e:
            logger.error(f"텍스트 선택 실패: {e}")
            return False

    def fill_cell_next_to_label(
        self,
        label: str,
        value: str,
        direction: str = "right",
        occurrence: int = 1,
        mode: str = "replace",
    ) -> Tuple[bool, str]:
        """
        표에서 레이블을 찾아 옆 셀에 값을 입력합니다.

        Args:
            label (str): 찾을 레이블 텍스트 (예: "성명")
            value (str): 입력할 값 (예: "장예준")
            direction (str): 이동 방향 - "right"(오른쪽), "down"(아래), "left"(왼쪽), "up"(위)
                - 레이블이 왼쪽에 있고 값이 오른쪽에 있으면 "right" 사용
                - 레이블이 위에 있고 값이 아래에 있으면 "down" 사용
            occurrence (int): 동일 레이블 중 몇 번째를 사용할지 (1부터 시작, 기본값: 1)
            mode (str): 입력 모드 - "replace"(기존 내용 삭제 후 입력), "prepend"(앞에 추가), "append"(뒤에 추가)

        Returns:
            Tuple[bool, str]: (성공 여부, 결과 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            # 1. 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 2. 레이블 찾기 (occurrence 횟수만큼 반복)
            found = False
            for i in range(occurrence):
                pset = self.hwp.HParameterSet.HFindReplace
                self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
                pset.FindString = label
                pset.FindRegExp = 0
                pset.IgnoreMessage = 1
                pset.Direction = 0  # forward
                result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

                if not result:
                    if i == 0:
                        return False, f"레이블 '{label}'을(를) 찾을 수 없습니다."
                    else:
                        return (
                            False,
                            f"레이블 '{label}'의 {occurrence}번째 항목을 찾을 수 없습니다. (총 {i}개 발견)",
                        )
                found = True

            if not found:
                return False, f"레이블 '{label}'을(를) 찾을 수 없습니다."

            # 3. 현재 셀(레이블 셀) 전체 선택 후 해제 - 커서 위치 확정
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")

            # 4. 지정된 방향으로 옆 셀로 이동
            direction_lower = direction.lower()
            if direction_lower == "right":
                self.hwp.HAction.Run("TableRightCell")
            elif direction_lower == "left":
                self.hwp.HAction.Run("TableLeftCell")
            elif direction_lower == "down":
                self.hwp.HAction.Run("MoveDown")
            elif direction_lower == "up":
                self.hwp.HAction.Run("TableUpperCell")
            else:
                return (
                    False,
                    f"잘못된 방향입니다: {direction}. 'right', 'left', 'down', 'up' 중 하나를 사용하세요.",
                )

            # 5. mode에 따라 값 입력
            mode_lower = mode.lower()
            if mode_lower == "replace":
                # 셀 전체 내용 선택 후 잘라내기
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("EditCut")
                self._insert_text_direct(value)
            elif mode_lower == "prepend":
                # 셀 시작으로 이동 후 입력
                self.hwp.HAction.Run("MoveSelCellBegin")
                self.hwp.HAction.Run("Cancel")
                self._insert_text_direct(value)
            elif mode_lower == "append":
                # 셀 끝으로 이동: 전체 선택 후 오른쪽으로 이동하면 끝으로 감
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("Cancel")
                self.hwp.HAction.Run("MoveLineEnd")
                self._insert_text_direct(value)
            else:
                return (
                    False,
                    f"잘못된 mode입니다: {mode}. 'replace', 'prepend', 'append' 중 하나를 사용하세요.",
                )

            return True, f"'{label}' 옆 셀에 '{value}' 입력 완료"

        except Exception as e:
            logger.error(f"셀 채우기 실패: {e}")
            return False, f"셀 채우기 실패: {str(e)}"

    def fill_table_cell(self, row: int, col: int, text: str) -> bool:
        """현재 커서가 위치한 표에서 지정한 셀(row, col)에 텍스트를 채운다.

        Args:
            row: 행 번호 (1부터 시작)
            col: 열 번호 (1부터 시작)
            text: 입력할 텍스트

        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 현재 위치 저장
            original_pos = self.hwp.GetPos()

            # 1. 표의 1행 1열로 확실히 이동
            if not self._move_to_top_left_cell():
                return False

            # 2. 지정된 행/열로 이동
            for _ in range(max(0, row - 1)):
                self.hwp.Run("TableLowerCell")
            for _ in range(max(0, col - 1)):
                self.hwp.Run("TableRightCell")

            # 셀 내용 삭제 후 텍스트 입력
            # 셀 전체 선택 → 내용 잘라내기(EditCut)로 완전히 비운 뒤 새 텍스트 입력
            self.hwp.Run("TableSelCell")
            self.hwp.HAction.Run("SelectAll")
            try:
                self.hwp.HAction.Run("EditCut")
            except Exception:
                # EditCut이 실패하면 Delete로 폴백
                self.hwp.Run("Delete")
            self._set_table_cursor()
            self._insert_text_direct(text)
            self.hwp.Run("Cancel")

            # 위치 복원 (실패해도 무시)
            try:
                if original_pos:
                    self.hwp.SetPos(*original_pos)
            except Exception:
                pass

            return True
        except Exception as e:
            logger.error(f"단일 셀 채우기 실패: {e}")
            return False

    def fill_cells_from_dict(
        self, label_value_map: Dict[str, str], direction: str = "right"
    ) -> Dict[str, Tuple[bool, str]]:
        """
        여러 레이블에 대해 옆 셀에 값을 입력합니다.

        Args:
            label_value_map (Dict[str, str]): 레이블과 값의 매핑 (예: {"성명": "장예준", "기업명": "mutual"})
            direction (str): 이동 방향 - "right", "down", "left", "up" (기본값: "right")

        Returns:
            Dict[str, Tuple[bool, str]]: 각 레이블에 대한 (성공 여부, 결과 메시지) 딕셔너리
        """
        results = {}

        for label, value in label_value_map.items():
            success, message = self.fill_cell_next_to_label(label, value, direction)
            results[label] = (success, message)

        return results

    def is_cursor_in_table(self) -> bool:
        """현재 커서가 표 안에 있는지 대략 판별한다.

        아이디어:
        - 표 안에서는 `TableSelCell` 액션이 정상적으로 동작해서 셀 선택이 된다.
        - 표 밖에서는 아무 변화가 없거나 예외가 날 수 있다.
        - 이 함수를 "가벼운 프로빙" 용도로만 쓰고, 실패해도 치명적이지 않게 처리한다.
        """
        try:
            if not self.is_hwp_running or not self.hwp:
                return False

            # 현재 위치 백업
            original_pos = self.hwp.GetPos()

            # 셀 선택 시도
            result = self.hwp.HAction.Run("TableSelCell")

            # 선택 해제
            self.hwp.HAction.Run("Cancel")

            # 위치 복원
            if original_pos:
                try:
                    self.hwp.SetPos(*original_pos)
                except Exception as e_setpos:
                    logger.debug(
                        f"is_cursor_in_table: 위치 복원 실패(무시): {e_setpos}"
                    )

            # HAction.Run 반환값이 환경에 따라 다를 수 있으므로,
            # 일단 예외 없이 실행되었다면 True 쪽으로 간주하고,
            # 필요하면 나중에 더 정교하게 튜닝한다.
            return bool(result)
        except Exception as e:
            logger.debug(f"is_cursor_in_table 체크 중 오류: {e}")
            return False

    def _move_to_top_left_cell(self) -> bool:
        """현재 커서가 있는 표의 가장 첫 번째 셀(1행 1열)로 이동합니다.
        드래그(셀 블록) 상태에서도 안전하게 작동합니다.
        """
        try:
            # 1. 셀 선택 모드 진입 (이미 드래그 중이면 해당 영역 유지, 아니면 현재 셀 선택)
            self.hwp.Run("TableSelCell")
            # 2. 현재 열의 맨 위행으로 이동
            self.hwp.Run("TableRowBegin")
            # 3. 현재 행의 맨 앞열로 이동
            self.hwp.Run("TableColBegin")
            # 4. 선택 모드 해제 (커서는 이제 1행 1열에 위치)
            self.hwp.Run("Cancel")
            return True
        except Exception as e:
            logger.error(f"표 첫 번째 셀로 이동 실패: {e}")
            return False

    def get_current_table_as_text(self) -> str:
        """현재 커서가 있는 표의 전체 내용을 탭과 개행으로 구분된 텍스트로 가져온다.
        AI 서버의 plan_table 입력용으로 사용됩니다.
        """
        try:
            if not self.is_hwp_running:
                return ""

            # 표 전체 선택
            self.hwp.Run("TableSelCell")
            self.hwp.Run("TableSelTable")

            # 클립보드로 복사
            self.hwp.HAction.Run("Copy")
            self.hwp.Run("Cancel")

            import win32clipboard

            win32clipboard.OpenClipboard()
            try:
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()

            return text or ""
        except Exception as e:
            logger.error(f"표 텍스트 추출 실패: {e}")
            return ""

    def get_table_cell_text(self, row: int, col: int) -> str:
        """현재 표의 특정 셀(row, col)의 텍스트를 가져온다.

        Args:
            row: 행 번호 (1부터 시작)
            col: 열 번호 (1부터 시작)

        Returns:
            str: 셀 텍스트 내용
        """
        try:
            if not self.is_hwp_running:
                return ""

            # 1. 표의 1행 1열로 확실히 이동
            if not self._move_to_top_left_cell():
                return ""

            # 2. 지정된 행/열로 이동
            for _ in range(max(0, row - 1)):
                self.hwp.Run("TableLowerCell")
            for _ in range(max(0, col - 1)):
                self.hwp.Run("TableRightCell")

            # 셀 선택 후 텍스트 가져오기
            self.hwp.HAction.Run("TableSelCell")
            text = self._get_cell_text_by_clipboard()
            self.hwp.HAction.Run("Cancel")

            return text
        except Exception as e:
            logger.error(f"표 셀 텍스트 가져오기 실패: {e}")
            return ""

    def merge_table_cells(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ) -> bool:
        """현재 표의 지정된 범위의 셀들을 병합한다.

        Args:
            start_row: 시작 행 (1부터 시작)
            start_col: 시작 열 (1부터 시작)
            end_row: 종료 행 (1부터 시작)
            end_col: 종료 열 (1부터 시작)

        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 1. 표의 1행 1열로 확실히 이동
            if not self._move_to_top_left_cell():
                return False

            # 2. 시작 셀로 이동
            for _ in range(max(0, start_row - 1)):
                self.hwp.Run("TableLowerCell")
            for _ in range(max(0, start_col - 1)):
                self.hwp.Run("TableRightCell")

            # 선택 시작 (F5 한 번: 셀 선택, F5 두 번: 다중 셀 선택 시작)
            self.hwp.Run("TableCellBlock")  # F5
            self.hwp.Run("TableCellBlock")  # F5 again for multi-selection

            # 종료 셀까지 확장
            row_diff = end_row - start_row
            col_diff = end_col - start_col

            for _ in range(max(0, row_diff)):
                self.hwp.Run("TableLowerCell")
            for _ in range(max(0, col_diff)):
                self.hwp.Run("TableRightCell")

            # 병합 실행
            result = self.hwp.Run("TableMergeCell")
            self.hwp.Run("Cancel")

            return bool(result)
        except Exception as e:
            logger.error(f"표 셀 병합 실패: {e}")
            return False

    def fill_table_with_data(
        self,
        data: List[List[str]],
        start_row: int = 1,
        start_col: int = 1,
        has_header: bool = False,
    ) -> bool:
        """현재 커서 위치의 표에 데이터를 채운다.

        Args:
            data (List[List[str]]): 채울 데이터 2차원 리스트 (행 x 열)
            start_row (int): 시작 행 번호 (1부터 시작)
            start_col (int): 시작 열 번호 (1부터 시작)
            has_header (bool): 첫 번째 행을 헤더로 처리할지 여부

        Returns:
            bool: 작업 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 현재 위치 저장 (실패해도 진행)
            original_pos = None
            try:
                original_pos = self.hwp.GetPos()
            except Exception as e_pos:
                logger.debug(f"fill_table_with_data: 현재 위치 백업 실패(무시): {e_pos}")

            # 1. 표의 1행 1열로 확실히 이동
            if not self._move_to_top_left_cell():
                return False

            # 2. 시작 위치(start_row, start_col)로 이동
            for _ in range(start_row - 1):
                self.hwp.Run("TableLowerCell")

            for _ in range(start_col - 1):
                self.hwp.Run("TableRightCell")

            # 데이터 채우기
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    # 셀 선택 및 내용 삭제
                    self.clear_cell_content()
                    self._set_table_cursor()

                    # 셀에 값 입력
                    if has_header and row_idx == 0:
                        self.set_font_style(bold=True)
                        self.hwp.HAction.GetDefault(
                            "InsertText", self.hwp.HParameterSet.HInsertText.HSet
                        )
                        self.hwp.HParameterSet.HInsertText.Text = cell_value
                        self.hwp.HAction.Execute(
                            "InsertText", self.hwp.HParameterSet.HInsertText.HSet
                        )
                        self.set_font_style(bold=False)
                    else:
                        self.hwp.HAction.GetDefault(
                            "InsertText", self.hwp.HParameterSet.HInsertText.HSet
                        )
                        self.hwp.HParameterSet.HInsertText.Text = cell_value
                        self.hwp.HAction.Execute(
                            "InsertText", self.hwp.HParameterSet.HInsertText.HSet
                        )

                    # 다음 셀로 이동 (마지막 셀은 이동하지 않음)
                    if col_idx < len(row_data) - 1:
                        self.hwp.Run("TableRightCell")

                # 다음 행으로 이동 (마지막 행이 아닌 경우)
                if row_idx < len(data) - 1:
                    for _ in range(len(row_data) - 1):
                        self.hwp.Run("TableLeftCell")
                    self.hwp.Run("TableLowerCell")

            # 표 밖으로 커서 이동
            self.hwp.Run("TableSelCell")  # 현재 셀 선택
            self.hwp.Run("Cancel")  # 선택 취소
            self.hwp.Run("MoveDown")  # 아래로 이동

            # 위치를 완전히 되돌릴 필요는 없지만, 필요하면 아래 코드 활성화
            # if original_pos:
            #     self.hwp.SetPos(*original_pos)

            return True

        except Exception as e:
            logger.error(f"표 데이터 채우기 실패: {e}")
            return False

    def increment_date_column_in_current_table(
        self, days: int = 1, date_col: int = 1
    ) -> bool:
        """현재 커서가 위치한 표에서 특정 열(날짜 열)의 날짜만 +days 만큼 증가시킨다.

        이 메서드는 텍스트 기반 파싱 대신, 실제 표 셀을 순회하면서
        날짜가 들어있는 셀만 직접 수정한다.

        Args:
            days (int): 증가시킬 일 수 (기본 1)
            date_col (int): 날짜가 위치한 열 번호 (1부터 시작)

        Returns:
            bool: 작업 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False

            # 날짜 파싱 포맷 후보
            date_formats = [
                "%Y. %m. %d",
                "%Y.%m.%d",
                "%Y-%m-%d",
                "%Y/%m/%d",
            ]

            def try_parse_date(s: str):
                s_stripped = str(s).strip()
                for fmt in date_formats:
                    try:
                        return datetime.strptime(s_stripped, fmt)
                    except ValueError:
                        continue
                return None

            # 현재 위치 저장
            original_pos = self.hwp.GetPos()

            # 1. 표의 1행 1열로 확실히 이동
            if not self._move_to_top_left_cell():
                return False

            # 2. 날짜 열로 이동 (1-based)
            for _ in range(max(0, date_col - 1)):
                self.hwp.Run("TableRightCell")

            while True:
                # 현재 셀 텍스트 얻기
                self.hwp.HAction.Run("TableSelCell")
                cell_text = self._get_cell_text_by_clipboard()
                self.hwp.HAction.Run("Cancel")

                dt = try_parse_date(cell_text)
                if dt is not None:
                    new_dt = dt + timedelta(days=days)
                    if "/" in cell_text:
                        out = new_dt.strftime("%Y/%m/%d")
                    elif "-" in cell_text:
                        out = new_dt.strftime("%Y-%m-%d")
                    else:
                        out = new_dt.strftime("%Y. %m. %d")

                    # 셀 내용 교체
                    self.hwp.HAction.Run("TableSelCell")
                    self.hwp.HAction.Run("Delete")
                    self._insert_text_direct(out)
                    self.hwp.HAction.Run("Cancel")

                # 다음 행으로 이동 시도
                try:
                    self.hwp.HAction.Run("TableLowerCell")
                except Exception:
                    break

                # 표 범위를 벗어나면 예외가 나거나, TableSelCell이 실패할 수 있음
                try:
                    self.hwp.HAction.Run("TableSelCell")
                    self.hwp.HAction.Run("Cancel")
                except Exception:
                    break

            # 위치 복원 시도 (실패해도 치명적이지 않으므로 무시)
            try:
                if original_pos:
                    self.hwp.SetPos(*original_pos)
            except Exception:
                pass

            return True
        except Exception as e:
            logger.error(f"표 날짜 열 증가 실패: {e}")
            return False

    def _move_direction(self, direction: str) -> bool:
        """
        지정된 방향으로 셀 이동.

        Args:
            direction: 이동 방향 ("right", "left", "down", "up")

        Returns:
            bool: 성공 여부
        """
        move_actions = {
            "right": "TableRightCell",
            "left": "TableLeftCell",
            "down": "TableLowerCell",
            "up": "TableUpperCell",
        }
        action = move_actions.get(direction.lower())
        if action:
            self.hwp.HAction.Run(action)
            return True
        return False

    def _get_cell_text_by_clipboard(self) -> str:
        """
        현재 선택된 셀의 텍스트를 클립보드를 통해 가져옵니다.
        (내부 헬퍼 함수 - 셀이 이미 선택된 상태에서 호출)
        """
        import win32clipboard

        # SelectAll로 셀 내용 전체 선택 후 복사
        self.hwp.HAction.Run("SelectAll")
        self.hwp.HAction.Run("Copy")
        self.hwp.HAction.Run("Cancel")

        # 클립보드에서 텍스트 읽기
        win32clipboard.OpenClipboard()
        try:
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        except Exception as e:
            logger.debug(f"클립보드 읽기 실패: {e}")
            text = ""
        finally:
            win32clipboard.CloseClipboard()

        return text.strip() if text else "(빈 셀)"

    def navigate_and_get_cell(self, direction: str) -> Tuple[bool, str, str]:
        """
        지정된 방향으로 이동하고 현재 셀의 내용을 반환합니다.

        Args:
            direction: 이동 방향 ("right", "left", "down", "up")

        Returns:
            Tuple[bool, str, str]: (성공 여부, 방향, 셀 내용)
        """
        try:
            if not self.is_hwp_running:
                return False, direction, "HWP가 연결되어 있지 않습니다."

            # 현재 셀 위치 확정 후 이동 (기존 _find_labels_recursive와 동일한 로직)
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")
            self._move_direction(direction)

            # 이동 후 셀 선택하고 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            text = self._get_cell_text_by_clipboard()

            return True, direction, text
        except Exception as e:
            return False, direction, f"네비게이션 실패: {str(e)}"

    def get_table_view(self, depth: int = 1) -> Tuple[bool, Dict[str, Any]]:
        """
        현재 위치 기준으로 주변 셀들의 내용을 가져옵니다.

        Args:
            depth: 탐색 깊이 (1이면 상하좌우 1칸씩, 2면 2칸씩...)

        Returns:
            Tuple[bool, Dict]: (성공 여부, 셀 내용 딕셔너리)
            딕셔너리 구조:
            {
                "center": "현재 셀 내용",
                "up_1": "위 1칸",
                "up_2": "위 2칸",
                "down_1": "아래 1칸",
                "left_1": "왼쪽 1칸",
                "right_1": "오른쪽 1칸",
                ...
            }
        """
        try:
            if not self.is_hwp_running:
                return False, {"error": "HWP가 연결되어 있지 않습니다."}

            result = {}

            # 현재 셀 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            result["center"] = self._get_cell_text_by_clipboard()
            self.hwp.HAction.Run("Cancel")

            # 각 방향으로 탐색
            directions = [
                ("up", "TableUpperCell"),
                ("down", "TableLowerCell"),
                ("left", "TableLeftCell"),
                ("right", "TableRightCell"),
            ]

            for dir_name, action in directions:
                # 현재 위치로 돌아오기 위해 반대 방향 저장
                opposite = {
                    "up": "TableLowerCell",
                    "down": "TableUpperCell",
                    "left": "TableRightCell",
                    "right": "TableLeftCell",
                }

                for d in range(1, depth + 1):
                    # 이동
                    self.hwp.HAction.Run(action)
                    self.hwp.HAction.Run("TableSelCell")
                    cell_text = self._get_cell_text_by_clipboard()
                    result[f"{dir_name}_{d}"] = cell_text
                    self.hwp.HAction.Run("Cancel")

                # 원래 위치로 복귀
                for _ in range(depth):
                    self.hwp.HAction.Run(opposite[dir_name])

            return True, result
        except Exception as e:
            return False, {"error": f"테이블 뷰 가져오기 실패: {str(e)}"}

    def find_and_get_cell(self, text: str) -> Tuple[bool, str]:
        """
        텍스트를 찾고 해당 셀의 내용을 반환합니다.

        Args:
            text: 찾을 텍스트

        Returns:
            Tuple[bool, str]: (성공 여부, 셀 내용 또는 에러 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            # 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 텍스트 찾기
            pset = self.hwp.HParameterSet.HFindReplace
            self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
            pset.FindString = text
            pset.FindRegExp = 0
            pset.IgnoreMessage = 1
            pset.Direction = 0
            result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

            if not result:
                return False, f"'{text}'을(를) 찾을 수 없습니다."

            # 찾은 후 셀 선택하고 내용 가져오기
            self.hwp.HAction.Run("TableSelCell")
            cell_text = self._get_cell_text_by_clipboard()

            return True, cell_text
        except Exception as e:
            return False, f"찾기 실패: {str(e)}"

    def _find_labels_recursive(
        self, path: List[str], depth: int = 0
    ) -> Tuple[bool, int]:
        """
        경로의 레이블들을 순차적으로 찾는 재귀 함수.
        방향 키워드(<left>, <right>, <up>, <down>)도 지원합니다.

        Args:
            path: 찾을 레이블 경로 (예: ["대표자", "<down>", "<right>"])
            depth: 현재 깊이 (인덱스)

        Returns:
            Tuple[bool, int]: (성공 여부, 찾은 depth)
        """
        # 기저 조건: 모든 항목 처리 완료
        if depth >= len(path):
            return True, depth

        item = path[depth]

        # 방향 키워드 처리: <left>, <right>, <up>, <down>
        if item.startswith("<") and item.endswith(">"):
            direction = item[1:-1].lower()  # "<down>" -> "down"
            if direction in ["left", "right", "up", "down"]:
                # 현재 셀 위치 확정 후 이동
                self.hwp.HAction.Run("TableSelCell")
                self.hwp.HAction.Run("Cancel")
                self._move_direction(direction)
                # 재귀: 다음 항목 처리
                return self._find_labels_recursive(path, depth + 1)
            else:
                return False, depth  # 잘못된 방향 키워드

        # 일반 레이블 찾기
        pset = self.hwp.HParameterSet.HFindReplace
        self.hwp.HAction.GetDefault("RepeatFind", pset.HSet)
        pset.FindString = item
        pset.FindRegExp = 0
        pset.IgnoreMessage = 1
        pset.Direction = 0  # forward
        result = self.hwp.HAction.Execute("RepeatFind", pset.HSet)

        if not result:
            return False, depth

        # 재귀: 다음 항목 처리
        return self._find_labels_recursive(path, depth + 1)

    def fill_cell_by_path(
        self,
        path: List[str],
        value: str,
        direction: str = "right",
        mode: str = "replace",
    ) -> Tuple[bool, str]:
        """
        경로를 따라 레이블과 방향 키워드를 순차적으로 처리하여 셀에 값을 입력합니다.

        **방향 키워드 지원:** <left>, <right>, <up>, <down>
        - 레이블 찾기와 방향 이동을 조합하여 복잡한 표 구조도 정확하게 탐색

        **사용 예시:**
        - path=["대표자", "<down>"] → 대표자 찾고 → 아래로 이동 → direction 방향으로 값 입력
        - path=["대표자", "<down>", "<right>"] → 대표자 → 아래 → 오른쪽 → 값 입력
        - path=["투자"] → 투자 찾고 → direction 방향으로 값 입력 (단위 셀에 prepend)

        Args:
            path: 레이블과 방향 키워드의 경로 (예: ["대표자", "<down>"])
            value: 입력할 값
            direction: 마지막 항목 처리 후 이동 방향 ("right", "down", "left", "up")
            mode: 입력 모드 ("replace", "prepend", "append")

        Returns:
            Tuple[bool, str]: (성공 여부, 결과 메시지)
        """
        try:
            if not self.is_hwp_running:
                return False, "HWP가 연결되어 있지 않습니다."

            if not path or len(path) == 0:
                return False, "경로가 비어있습니다."

            # 1. 문서 처음으로 이동
            self.hwp.HAction.Run("MoveDocBegin")

            # 2. 재귀적으로 경로의 모든 레이블 찾기
            found, found_depth = self._find_labels_recursive(path)
            if not found:
                if found_depth == 0:
                    return False, f"첫 번째 레이블 '{path[0]}'을(를) 찾을 수 없습니다."
                else:
                    found_path = " > ".join(path[:found_depth])
                    missing_label = path[found_depth]
                    return (
                        False,
                        f"'{found_path}' 이후에 '{missing_label}'을(를) 찾을 수 없습니다.",
                    )

            # 3. 현재 셀 선택 후 해제 - 커서 위치 확정
            self.hwp.HAction.Run("TableSelCell")
            self.hwp.HAction.Run("Cancel")

            # 4. 마지막 항목이 방향 키워드가 아닌 경우에만 direction으로 추가 이동
            last_item = path[-1] if path else ""
            is_last_direction = last_item.startswith("<") and last_item.endswith(">")

            if not is_last_direction:
                direction_lower = direction.lower()
                if direction_lower == "right":
                    self.hwp.HAction.Run("TableRightCell")
                elif direction_lower == "left":
                    self.hwp.HAction.Run("TableLeftCell")
                elif direction_lower == "down":
                    self.hwp.HAction.Run("TableLowerCell")
                elif direction_lower == "up":
                    self.hwp.HAction.Run("TableUpperCell")

            # 5. mode에 따라 값 입력
            mode_lower = mode.lower()
            if mode_lower == "replace":
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("EditCut")
                self._insert_text_direct(value)
            elif mode_lower == "prepend":
                self.hwp.HAction.Run("MoveSelCellBegin")
                self.hwp.HAction.Run("Cancel")
                self._insert_text_direct(value)
            elif mode_lower == "append":
                # 셀 끝으로 이동: 전체 선택 후 오른쪽으로 이동하면 끝으로 감
                self.hwp.HAction.Run("SelectAll")
                self.hwp.HAction.Run("Cancel")
                self.hwp.HAction.Run("MoveLineEnd")
                self._insert_text_direct(value)
            else:
                return (
                    False,
                    f"잘못된 mode입니다: {mode}. 'replace', 'prepend', 'append' 중 하나를 사용하세요.",
                )

            path_str = " > ".join(path)
            return True, f"'{path_str}' 경로의 셀에 '{value}' 입력 완료"

        except Exception as e:
            return False, f"셀 채우기 실패: {str(e)}"

    def fill_cells_by_path_batch(
        self,
        path_value_map: Dict[str, str],
        direction: str = "right",
        mode: str = "replace",
    ) -> Dict[str, Tuple[bool, str]]:
        """
        여러 경로에 대해 값을 일괄 입력합니다.

        Args:
            path_value_map: 경로(문자열)와 값의 매핑
                - 경로는 " > " 또는 "/"로 구분 (예: "대표자 > 총 인원" 또는 "대표자/총 인원")
            direction: 이동 방향 ("right", "down", "left", "up")
            mode: 입력 모드 ("replace", "prepend", "append")

        Returns:
            Dict[str, Tuple[bool, str]]: 각 경로에 대한 (성공 여부, 결과 메시지)
        """
        results = {}

        for path_str, value in path_value_map.items():
            # 경로 문자열을 리스트로 변환
            if " > " in path_str:
                path = [p.strip() for p in path_str.split(" > ")]
            elif "/" in path_str:
                path = [p.strip() for p in path_str.split("/")]
            else:
                path = [path_str]

            success, message = self.fill_cell_by_path(path, value, direction, mode)
            results[path_str] = (success, message)

        return results
