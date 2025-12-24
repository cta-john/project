"""
KB증권 이미지 인식 테스트
이미지가 실제로 인식되는지 확인
"""

import os
import sys
sys.path.append(r"C:\HTS_Automation")

from hts_utils import imglocation, load_config
import cv2
import numpy as np
from PIL import ImageGrab

def test_image_recognition():
    """이미지 인식 테스트"""

    print("=" * 60)
    print("KB증권 이미지 인식 테스트")
    print("=" * 60)

    # 설정 로드
    config = load_config()
    kb_config = config['KB']
    image_folder = kb_config['이미지폴더']

    print(f"\n이미지 폴더: {image_folder}")
    print("\n테스트할 이미지 목록:")

    # 테스트할 이미지 목록
    test_images = [
        ("아이디_탭_활성화.png", 0.7),
        ("아이디_탭_비활성화.png", 0.7),
        ("조회전용_체크됨.png", 0.8),
        ("조회전용_체크안됨.png", 0.8),
        ("로그인_버튼.png", 0.7),
        ("조회전용안내.png", 0.7),
        ("hts_logo.png", 0.7),
    ]

    print("\n⚠️  KB HTS를 실행하고 로그인 화면을 띄운 상태에서 테스트하세요!\n")
    input("준비되면 Enter를 누르세요...")

    # 현재 화면 캡처
    print("\n현재 화면 캡처 중...")
    screenshot = ImageGrab.grab()
    screenshot_np = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
    print(f"화면 크기: {screenshot.size}")

    # 각 이미지 테스트
    for img_name, confidence in test_images:
        img_path = os.path.join(image_folder, img_name)

        print(f"\n{'=' * 60}")
        print(f"테스트: {img_name}")
        print(f"경로: {img_path}")

        # 파일 존재 확인
        if not os.path.exists(img_path):
            print(f"  ❌ 파일이 존재하지 않습니다!")
            continue

        # 이미지 크기 확인
        img = cv2.imread(img_path)
        if img is None:
            print(f"  ❌ 이미지를 읽을 수 없습니다!")
            continue

        print(f"  이미지 크기: {img.shape[1]}x{img.shape[0]} (width x height)")

        # 인식 테스트 (여러 confidence 값으로)
        for conf in [0.5, 0.6, 0.7, 0.8, 0.9]:
            result = imglocation(img_path, confidence=conf)
            if result:
                print(f"  ✅ Confidence {conf}: 발견! 위치 = {result}")
            else:
                print(f"  ❌ Confidence {conf}: 찾지 못함")

    print("\n" + "=" * 60)
    print("테스트 완료!")
    print("=" * 60)

    # 듀얼모니터 확인
    print("\n🖥️  모니터 정보:")
    try:
        import win32api
        monitors = win32api.EnumDisplayMonitors()
        print(f"  모니터 개수: {len(monitors)}")
        for i, monitor in enumerate(monitors):
            print(f"  모니터 {i+1}: {monitor}")
    except:
        print("  모니터 정보를 가져올 수 없습니다")

    # DPI 정보
    print("\n📐 DPI/배율 정보:")
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        print(f"  화면 해상도: {screen_width}x{screen_height}")
    except:
        print("  DPI 정보를 가져올 수 없습니다")


if __name__ == "__main__":
    test_image_recognition()
