"""pyhwpx FindCtrl 팝업 억제 패치.

pyhwpx의 find_ctrl() 메서드를 수정하여 COM FindCtrl() 호출 전에
SetMessageBoxMode(0x2FFF1)을 설정한다. (모드 복원 없음 — 비동기 팝업 대응)
나머지 13곳의 self.hwp.FindCtrl() 직접 호출도 self.find_ctrl() 경유로 변경.

사용법:
    python scripts/patch_pyhwpx.py          # 패치 적용
    python scripts/patch_pyhwpx.py --check  # 패치 상태 확인
"""

from __future__ import annotations

import importlib.util
import os
import sys

ORIGINAL = """\
    def find_ctrl(self):
        \"\"\"컨트롤 선택하기\"\"\"
        return self.hwp.FindCtrl()"""

PATCHED = """\
    def find_ctrl(self):
        \"\"\"컨트롤 선택하기 (FindCtrl 팝업 자동 억제)\"\"\"
        self.hwp.SetMessageBoxMode(Mode=0x2FFF1)
        return self.hwp.FindCtrl()"""

ORIGINAL_DIRECT = "self.hwp.FindCtrl()"
PATCHED_DIRECT = "self.find_ctrl()"


def find_pyhwpx_core() -> str | None:
    spec = importlib.util.find_spec("pyhwpx")
    if spec is None or spec.origin is None:
        return None
    return os.path.join(os.path.dirname(spec.origin), "core.py")


def check(path: str) -> bool:
    with open(path, encoding="utf-8") as f:
        content = f.read()
    return "FindCtrl 팝업 자동 억제" in content


def patch(path: str) -> bool:
    with open(path, encoding="utf-8") as f:
        content = f.read()

    if "FindCtrl 팝업 자동 억제" in content:
        print("이미 패치 적용됨")
        return True

    if ORIGINAL not in content:
        print("ERROR: 원본 find_ctrl 메서드를 찾을 수 없음 — pyhwpx 버전 불일치?")
        return False

    # 1) find_ctrl 메서드 교체
    content = content.replace(ORIGINAL, PATCHED)

    # 2) self.hwp.FindCtrl() 직접 호출을 self.find_ctrl()로 변경
    content = content.replace(ORIGINAL_DIRECT, PATCHED_DIRECT)

    # 3) find_ctrl 내부에서 재귀 호출 방지 — 다시 COM 직접 호출로 복원
    content = content.replace(
        "        self.hwp.SetMessageBoxMode(Mode=0x2FFF1)\n        return self.find_ctrl()",
        "        self.hwp.SetMessageBoxMode(Mode=0x2FFF1)\n        return self.hwp.FindCtrl()",
    )

    # 4) pyc 캐시 삭제
    cache_dir = os.path.join(os.path.dirname(path), "__pycache__")
    if os.path.isdir(cache_dir):
        for f in os.listdir(cache_dir):
            if f.startswith("core"):
                os.remove(os.path.join(cache_dir, f))
                print(f"캐시 삭제: {f}")

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"패치 완료: {path}")
    return True


if __name__ == "__main__":
    core_path = find_pyhwpx_core()
    if core_path is None:
        print("pyhwpx를 찾을 수 없습니다")
        sys.exit(1)

    if "--check" in sys.argv:
        ok = check(core_path)
        print(f"패치 상태: {'적용됨' if ok else '미적용'}")
        sys.exit(0 if ok else 1)
    else:
        ok = patch(core_path)
        sys.exit(0 if ok else 1)
