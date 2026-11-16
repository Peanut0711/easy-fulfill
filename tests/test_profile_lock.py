"""같은 브라우저 프로필(쿠팡·우체국)을 쓰는 작업이 동시에 실행되지 않는지 검증한다."""

import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

HOLDER = r"""
import sys, time
sys.path.insert(0, sys.argv[1])
from pathlib import Path
from shared.profile_lock import hold_profile_lock
hold_profile_lock(Path(sys.argv[2]), "테스트")
print("HELD", flush=True)
time.sleep(float(sys.argv[3]))
"""


def remove_tree_when_released(path, timeout=10.0):
    """임시 폴더를 지운다. 강제 종료한 자식 프로세스의 잠금 파일 핸들을 Windows가 늦게 풀면
    WinError 32(다른 프로세스가 사용 중)가 나므로, 풀릴 때까지 잠시 다시 시도한다.
    제한 시간 안에 풀리지 않으면 검증 결과를 가리지 않도록 OS 임시 폴더에 남겨 둔다.
    """
    deadline = time.monotonic() + timeout
    while True:
        try:
            shutil.rmtree(path)
            return
        except FileNotFoundError:
            return
        except PermissionError:
            if time.monotonic() >= deadline:
                return
            time.sleep(0.1)


class ProfileLockTests(unittest.TestCase):
    def setUp(self):
        from shared.profile_lock import release_profile_lock
        directory = Path(tempfile.mkdtemp())
        self.addCleanup(remove_tree_when_released, directory)
        self.profile = directory / "coupang-browser-profile"
        self.profile.mkdir()
        self.addCleanup(release_profile_lock, self.profile)

    def start_holder(self, seconds=20):
        process = subprocess.Popen(
            [sys.executable, "-c", HOLDER, str(ROOT), str(self.profile), str(seconds)],
            stdout=subprocess.PIPE, text=True)
        self.addCleanup(lambda: (process.kill(), process.wait(timeout=10), process.stdout.close()))
        self.assertEqual(process.stdout.readline().strip(), "HELD")
        return process

    def test_second_process_is_refused_while_first_holds_profile(self):
        from shared.profile_lock import ProfileBusyError, hold_profile_lock
        self.start_holder()
        with self.assertRaises(ProfileBusyError) as raised:
            hold_profile_lock(self.profile, "쿠팡 WING")
        self.assertIn("쿠팡 WING", str(raised.exception))

    def test_same_process_can_reacquire_and_lock_is_released_on_exit(self):
        from shared.profile_lock import hold_profile_lock
        holder = self.start_holder(seconds=0.2)
        holder.wait(timeout=10)
        hold_profile_lock(self.profile, "쿠팡 WING")
        hold_profile_lock(self.profile, "쿠팡 WING")  # 같은 프로세스 재진입(업로드 재시도 등)

    def test_clear_session_refuses_while_profile_is_in_use(self):
        import product_detail.coupang_session as coupang_session
        import product_detail.coupang_cdn_upload as cdn
        self.start_holder()
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        code = (
            "import sys; sys.path.insert(0, sys.argv[1]);"
            "from pathlib import Path;"
            "import product_detail.coupang_cdn_upload as cdn;"
            "import product_detail.coupang_session as s;"
            "cdn.PROFILE_DIR = Path(sys.argv[2]); s.COUPANG_PROFILE_DIR = Path(sys.argv[2]);"
            "sys.argv = ['coupang_session', '--clear']; s.main()"
        )
        completed = subprocess.run([sys.executable, "-c", code, str(ROOT), str(self.profile)],
                                   capture_output=True, text=True, encoding="utf-8", env=env, timeout=60)
        self.assertNotEqual(completed.returncode, 0, completed.stdout + completed.stderr)
        self.assertTrue(self.profile.exists())
        self.assertIn("사용 중", completed.stdout + completed.stderr)


if __name__ == "__main__":
    unittest.main()
