#!/usr/bin/env python3
"""
바코드 라벨 생성기 메인 실행 파일
"""

import sys
import os
from pathlib import Path

# 현재 디렉토리를 Python 경로에 추가
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

# src 패키지를 모듈로 실행
if __name__ == "__main__":
    # -m 옵션으로 src.main 모듈 실행
    import subprocess
    result = subprocess.run([sys.executable, "-m", "src.main"], cwd=current_dir)
    sys.exit(result.returncode)
