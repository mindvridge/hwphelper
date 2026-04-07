"""pyhwpx / win32com COM 자동화 래퍼 — 한/글 프로그램 직접 제어.

사용법::

    with HwpController(visible=False) as hwp:
        hwp.open("template.hwp")
        # ... 작업
        hwp.save_as("output.hwp")
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import structlog

logger = structlog.get_logger()

# COM 가용 여부 플래그 — 테스트에서 skip 판단용
HAS_HWP = False
try:
    import win32com.client as win32  # type: ignore[import-untyped]

    HAS_HWP = True
except ImportError:
    win32 = None  # type: ignore[assignment]


class HwpController:
    """한/글 COM 자동화 컨트롤러.

    컨텍스트 매니저로 사용::

        with HwpController() as hwp:
            hwp.open("문서.hwp")
            ...
    """

    def __init__(self, visible: bool = False, security_module: bool = True) -> None:
        self._hwp: Any = None
        self._visible = visible
        self._security_module = security_module
        self._file_path: str | None = None

    # ------------------------------------------------------------------
    # 라이프사이클
    # ------------------------------------------------------------------

    def connect(self) -> None:
        """한/글 COM 객체를 생성하고 초기 설정을 적용한다."""
        if self._hwp is not None:
            return

        self._via_pyhwpx = False
        try:
            from pyhwpx import Hwp
            # pyhwpx Hwp()는 내부적으로 register_module()을 호출하므로
            # 별도 보안 모듈 등록이 불필요하다.
            self._hwp = Hwp(visible=self._visible, register_module=self._security_module)
            self._via_pyhwpx = True
            logger.info("한/글 COM 연결 (pyhwpx)")
        except Exception:
            if win32 is None:
                raise RuntimeError("pywin32가 설치되어 있지 않습니다.")
            try:
                self._hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                self._hwp = win32.Dispatch("HWPFrame.HwpObject")
            if not self._visible:
                self._hwp.XHwpWindows.Active = False
            logger.info("한/글 COM 연결 (win32com)")

            # win32com 직접 연결 시에만 보안 설정 필요
            self._suppress_security_popups()
            if self._security_module:
                self._register_security_module()

        # 모든 경로 공통: 팝업 자동 처리
        self._set_default_msgbox_mode()

    def _set_default_msgbox_mode(self) -> None:
        """모든 팝업을 자동 처리하도록 설정한다.

        한/글 SetMessageBoxMode 플래그 (pyhwpx 문서 기준):
          0x00001 — OK만 있는 팝업 → OK 자동
          0x00010 — OK/Cancel 팝업 → OK 자동
          0x01000 — Yes/No/Cancel 팝업 → Yes 자동
          0x10000 — Yes/No 팝업 → Yes 자동

        '문서 끝 도달' 찾기 팝업(Yes/No)은 Yes 선택 시 무한 루프 위험이
        있으나, 우리 코드에서 FindReplace를 직접 호출하지 않으므로 안전.
        pyhwpx의 find/replace 메서드는 내부에서 0x2FFF1로 전환 후 복원함.
        """
        try:
            self._hwp.SetMessageBoxMode(0x11011)
        except Exception:
            pass

    def _suppress_security_popups(self) -> None:
        """win32com 직접 연결 시 보안 관련 추가 설정을 적용한다."""
        self._set_default_msgbox_mode()

        try:
            # 파일 접근 보안 경고 비활성화
            self._hwp.XHwpDocuments.XHwpOptions.SetFileAccessControl(0)
        except Exception:
            pass

    def _register_security_module(self) -> None:
        """한/글 보안 모듈을 등록하여 자동화 접근을 허용한다."""
        try:
            # 레지스트리에 DLL이 없으면 먼저 등록
            self._ensure_security_dll_registered()
            self._hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            logger.debug("보안 모듈 등록 완료")
        except Exception:
            logger.warning("보안 모듈 등록 실패 — 파일 접근 팝업이 뜰 수 있습니다")

    @staticmethod
    def _ensure_security_dll_registered() -> None:
        """FilePathCheckerModule.dll이 레지스트리에 없으면 자동 등록한다."""
        import importlib.util

        try:
            from winreg import (
                ConnectRegistry, HKEY_CURRENT_USER,
                OpenKey, KEY_READ, QueryValueEx, CloseKey,
            )
        except ImportError:
            return

        reg = ConnectRegistry(None, HKEY_CURRENT_USER)
        reg_path = r"Software\HNC\HwpAutomation\Modules"
        try:
            key = OpenKey(reg, reg_path, 0, KEY_READ)
            val, _ = QueryValueEx(key, "FilePathCheckerModule")
            CloseKey(key)
            if val and os.path.exists(val):
                return  # 이미 등록됨
        except Exception:
            pass

        # pyhwpx 패키지에서 DLL 경로 찾기
        spec = importlib.util.find_spec("pyhwpx")
        if spec is None or spec.origin is None:
            return
        dll_path = os.path.join(os.path.dirname(spec.origin), "FilePathCheckerModule.dll")
        if not os.path.exists(dll_path):
            return

        try:
            from winreg import CreateKeyEx, KEY_WRITE, SetValueEx, REG_SZ
            for rp in [r"Software\HNC\HwpAutomation\Modules",
                       r"Software\Hnc\HwpUserAction\Modules"]:
                key = CreateKeyEx(reg, rp, 0, KEY_WRITE)
                SetValueEx(key, "FilePathCheckerModule", 0, REG_SZ, dll_path)
                CloseKey(key)
            logger.info("보안 모듈 DLL 레지스트리 등록 완료", path=dll_path)
        except Exception:
            pass

    @property
    def hwp(self) -> Any:
        """내부 COM 객체를 반환한다. 연결되어 있지 않으면 자동 연결."""
        if self._hwp is None:
            self.connect()
        return self._hwp

    @property
    def file_path(self) -> str | None:
        """현재 열린 문서 경로."""
        return self._file_path

    # ------------------------------------------------------------------
    # 문서 열기 / 저장 / 닫기
    # ------------------------------------------------------------------

    def open(self, file_path: str) -> None:
        """HWP/HWPX 문서를 연다. 상대 경로는 절대 경로로 자동 변환."""
        abs_path = os.path.abspath(file_path)
        if not Path(abs_path).exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {abs_path}")

        # win32com 직접 연결 시에만 팝업 재설정 (pyhwpx는 자체 처리)
        if not self._via_pyhwpx:
            self._suppress_security_popups()
        self.hwp.Open(abs_path, "HWP", "forceopen:true")
        self._file_path = abs_path
        logger.info("문서 열기 완료", path=abs_path)

    def save(self, file_path: str | None = None) -> None:
        """현재 문서를 저장한다. file_path 미지정 시 원본 경로에 덮어쓴다."""
        if file_path:
            self.save_as(file_path)
        else:
            self.hwp.Save()
            logger.info("문서 저장 완료", path=self._file_path)

    def save_as(self, file_path: str, fmt: str = "hwp") -> None:
        """다른 이름으로 저장한다.

        Parameters
        ----------
        file_path : str
            저장 경로.
        fmt : str
            포맷 — "hwp", "hwpx", "pdf", "html", "txt".
        """
        abs_path = os.path.abspath(file_path)
        Path(abs_path).parent.mkdir(parents=True, exist_ok=True)

        format_map = {
            "hwp": "HWP",
            "hwpx": "HWPX",
            "pdf": "PDF",
            "html": "HTML",
            "txt": "TEXT",
        }
        save_fmt = format_map.get(fmt.lower(), "HWP")

        self.hwp.SaveAs(abs_path, save_fmt)
        self._file_path = abs_path
        logger.info("문서 다른 이름으로 저장", path=abs_path, format=save_fmt)

    def export_pdf(self, output_path: str) -> None:
        """현재 문서를 PDF로 내보낸다."""
        abs_path = os.path.abspath(output_path)
        Path(abs_path).parent.mkdir(parents=True, exist_ok=True)

        self.hwp.SaveAs(abs_path, "PDF")
        logger.info("PDF 내보내기 완료", path=abs_path)

    def close(self) -> None:
        """현재 문서를 닫는다 (한/글 프로그램은 유지)."""
        try:
            self.hwp.Clear(option=1)  # 저장 안 함
        except Exception:
            logger.debug("문서 닫기 중 오류 (무시)")
        self._file_path = None
        logger.info("문서 닫기 완료")

    def quit(self) -> None:
        """한/글 프로그램을 종료한다."""
        if self._hwp is None:
            return
        try:
            self._hwp.Clear(option=1)
            self._hwp.Quit()
        except Exception:
            logger.debug("한/글 종료 중 오류 (무시)")
        finally:
            self._hwp = None
            self._file_path = None
            logger.info("한/글 종료 완료")

    # ------------------------------------------------------------------
    # 컨텍스트 매니저
    # ------------------------------------------------------------------

    def __enter__(self) -> HwpController:
        self.connect()
        return self

    def __exit__(self, exc_type: type | None, exc_val: Exception | None, exc_tb: Any) -> None:
        self.quit()

    # ------------------------------------------------------------------
    # 내부 유틸리티 (하위 모듈에서 사용)
    # ------------------------------------------------------------------

    def run_action(self, action_id: str, pset: dict[str, Any] | None = None) -> Any:
        """한/글 Action을 실행한다."""
        act = self.hwp.CreateAction(action_id)
        param = act.CreateSet()
        if pset:
            for k, v in pset.items():
                param.SetItem(k, v)
        act.Execute(param)
        return param

    def get_pos(self) -> tuple[int, int, int]:
        """현재 커서 위치를 반환한다. (list, para, pos)"""
        return self.hwp.GetPos()

    def set_pos(self, list_id: int, para: int, pos: int) -> None:
        """커서 위치를 설정한다."""
        self.hwp.SetPos(list_id, para, pos)

    def get_text(self) -> tuple[int, str]:
        """현재 커서 위치의 텍스트를 가져온다."""
        return self.hwp.GetText()

    def move_to_field(self, field_name: str) -> bool:
        """누름틀 필드로 이동한다."""
        try:
            self.hwp.MoveToField(field_name, True, True, True)
            return True
        except Exception:
            return False

    def get_char_shape(self) -> dict[str, Any]:
        """현재 커서 위치의 글자 모양을 가져온다."""
        act = self.hwp.CreateAction("CharShape")
        param = act.CreateSet()
        act.GetDefault(param)
        return {
            "font_name": param.Item("FaceNameUser"),
            "font_size": param.Item("Height") / 100,  # 1/100pt → pt
            "bold": bool(param.Item("Bold")),
            "italic": bool(param.Item("Italic")),
            "char_spacing": param.Item("LetterSpacing"),
            "text_color": param.Item("TextColor"),
        }

    def set_char_shape(self, shape: dict[str, Any]) -> None:
        """글자 모양을 설정한다."""
        act = self.hwp.CreateAction("CharShape")
        param = act.CreateSet()
        act.GetDefault(param)
        if "font_name" in shape:
            param.SetItem("FaceNameUser", shape["font_name"])
            param.SetItem("FaceNameSymbol", shape["font_name"])
            param.SetItem("FaceNameOther", shape["font_name"])
            param.SetItem("FaceNameJapanese", shape["font_name"])
            param.SetItem("FaceNameHanja", shape["font_name"])
            param.SetItem("FaceNameLatin", shape["font_name"])
            param.SetItem("FaceNameHangul", shape["font_name"])
        if "font_size" in shape:
            param.SetItem("Height", int(shape["font_size"] * 100))
        if "bold" in shape:
            param.SetItem("Bold", shape["bold"])
        if "italic" in shape:
            param.SetItem("Italic", shape["italic"])
        if "char_spacing" in shape:
            param.SetItem("LetterSpacing", shape["char_spacing"])
        if "text_color" in shape:
            param.SetItem("TextColor", shape["text_color"])
        act.Execute(param)

    def get_para_shape(self) -> dict[str, Any]:
        """현재 커서 위치의 문단 모양을 가져온다."""
        act = self.hwp.CreateAction("ParaShape")
        param = act.CreateSet()
        act.GetDefault(param)
        return {
            "alignment": param.Item("Alignment"),
            "line_spacing_type": param.Item("LineSpacingType"),
            "line_spacing": param.Item("LineSpacing"),
        }
