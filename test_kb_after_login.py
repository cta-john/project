"""
KB증권 자동화 테스트 - 로그인 이후 단계만 (8~13단계)
사전 조건: KB HTS에 이미 로그인되어 메인 화면에 있어야 함
"""

import os
import sys
import time
from datetime import datetime

# hts_utils 임포트
sys.path.append(r"C:\HTS_Automation")
from hts_utils import (
    setup_logging,
    load_config,
    imglocation,
    wait_for_image,
    click_at_image,
    save_error_screenshot
)

import pyautogui
import pyperclip


def test_kb_download():
    """KB증권 데이터 다운로드 테스트 (로그인 이후)"""

    # 로깅 설정
    logger = setup_logging()
    logger.info("=" * 50)
    logger.info("KB증권 데이터 다운로드 테스트 시작 (8~13단계)")
    logger.info("=" * 50)

    try:
        # 설정 파일 로드
        print("\n[준비] 설정 파일 로드 중...")
        config = load_config()

        kb_config = config['KB']

        # 저장 폴더 생성
        save_dir = r"C:\Users\kiwoom\Documents\종투사증거금취합"
        os.makedirs(save_dir, exist_ok=True)

        print(f"✅ 이미지 폴더: {kb_config['이미지폴더']}")
        print(f"✅ 저장 폴더: {save_dir}")

        # 8단계: 화면번호 0191 입력
        print("\n[8단계] 화면번호 0191 입력 중...")
        logger.info("8단계: 화면번호 0191 입력")

        # 화면번호 입력창 이미지로 찾아서 클릭
        screen_number_field = os.path.join(kb_config['이미지폴더'], "screen_number_field.png")

        if click_at_image(screen_number_field, timeout=5, confidence=0.5, logger=logger):
            logger.info("✅ 화면번호 입력창 클릭 성공")
            print("✅ 화면번호 입력창 클릭 성공")
            time.sleep(1)  # 클릭 후 충분히 대기

            # 기존 내용 지우기 (혹시 있을 수 있으니)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)

            # 0191 입력 - pyperclip 사용 (더 안정적)
            pyperclip.copy("0191")
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'v')
            logger.info("0191 붙여넣기 완료")
            print("✅ 0191 붙여넣기 완료")
            time.sleep(0.5)

            # Enter 누르기
            pyautogui.press('enter')
            logger.info("Enter 키 입력 완료")
            print("✅ Enter 키 입력 완료")

            time.sleep(3)  # 화면 전환 대기
        else:
            logger.error("❌ 화면번호 입력창을 찾지 못했습니다!")
            print("❌ [8단계 실패] 화면번호 입력창을 찾지 못했습니다!")
            save_error_screenshot("KB_화면번호입력창찾기실패")
            return False

        # 9단계: 전체받기 버튼 클릭
        print("\n[9단계] 전체받기 버튼 클릭 중...")
        logger.info("9단계: 전체받기 버튼 클릭")
        download_button = os.path.join(kb_config['이미지폴더'], "download_all_button.png")

        if click_at_image(download_button, timeout=10, confidence=0.5, logger=logger):
            logger.info("✅ 전체받기 버튼 클릭 성공")
            print("✅ 전체받기 버튼 클릭 완료")
        else:
            logger.error("❌ 전체받기 버튼을 찾지 못했습니다!")
            print("❌ [9단계 실패] 전체받기 버튼을 찾지 못했습니다!")
            save_error_screenshot("KB_전체받기실패")
            return False

        # 10단계: 데이터 로딩 대기
        print("\n[10단계] 데이터 로딩 대기 중... (7초)")
        logger.info("10단계: 데이터 로딩 대기 (7초)...")
        time.sleep(7)
        print("✅ 데이터 로딩 대기 완료")

        # 11단계: 테이블 우클릭 및 파일 저장
        print("\n[11단계] 테이블 우클릭 중...")
        logger.info("11단계: 테이블 우클릭 및 파일 저장")

        # 테이블 영역 이미지로 찾아서 우클릭
        table_area = os.path.join(kb_config['이미지폴더'], "table_area.png")

        table_pos = imglocation(table_area, confidence=0.5)
        if table_pos:
            logger.info(f"✅ 테이블 영역 발견: {table_pos}")
            table_x, table_y = table_pos

            # 테이블 클릭
            pyautogui.click(table_x, table_y)
            time.sleep(0.5)

            # 우클릭
            pyautogui.rightClick(table_x, table_y)
            logger.info("테이블 우클릭 완료")
            print("✅ 테이블 우클릭 완료")
            time.sleep(1)
        else:
            logger.error("❌ 테이블 영역을 찾지 못했습니다!")
            print("❌ [11단계 실패] 테이블 영역을 찾지 못했습니다!")
            save_error_screenshot("KB_테이블영역찾기실패")
            return False

        # 12단계: 우클릭 메뉴에서 "파일로 내보내기" 선택
        print("\n[12단계] 우클릭 메뉴에서 '파일로 내보내기' 선택 중...")
        logger.info("12단계: 파일로 내보내기 메뉴 선택")

        # 우클릭 메뉴가 뜨는 시간 충분히 대기
        time.sleep(1.5)

        # 우클릭 메뉴 확인 (선택 사항 - 이미지 있으면 확인)
        context_menu = os.path.join(kb_config['이미지폴더'], "context_menu_export.png")

        # 이미지로 메뉴 확인 (있으면)
        if os.path.exists(context_menu):
            if wait_for_image(context_menu, timeout=3, confidence=0.5, logger=logger):
                logger.info("✅ 우클릭 메뉴 확인됨")
                print("✅ 우클릭 메뉴 확인됨")
            else:
                logger.warning("⚠️  우클릭 메뉴 이미지를 찾지 못함 - 계속 진행")
                print("⚠️  우클릭 메뉴 이미지를 찾지 못함 - 계속 진행")

        # 키보드로 메뉴 이동 (5번 down으로 변경 - 조정 가능)
        print("  → 메뉴 항목으로 이동 중... (down 키 5번)")
        for i in range(5):  # "파일로 내보내기"까지 이동 (횟수 조정 가능)
            pyautogui.press('down')
            logger.info(f"  down 키 {i+1}번째")
            time.sleep(0.2)

        # 우측 화살표로 서브메뉴 열기
        time.sleep(0.3)
        pyautogui.press('right')
        logger.info("파일로 내보내기 서브메뉴 열기 (right 키)")
        print("✅ 파일로 내보내기 선택 (서브메뉴 열기)")
        time.sleep(1)

        # 13단계: CSV 내보내기 선택
        print("\n[13단계] CSV 내보내기 선택 중...")
        logger.info("13단계: CSV 내보내기 선택")

        # CSV 메뉴 확인 (선택 사항 - 이미지 있으면 확인)
        csv_menu = os.path.join(kb_config['이미지폴더'], "csv_export_option.png")

        if os.path.exists(csv_menu):
            if wait_for_image(csv_menu, timeout=3, confidence=0.5, logger=logger):
                logger.info("✅ CSV 메뉴 확인됨")
                print("✅ CSV 메뉴 확인됨")

        # CSV 선택
        pyautogui.press('down')  # 첫 번째 항목으로 이동
        time.sleep(0.1)
        pyautogui.press('enter')
        logger.info("CSV 내보내기 선택 완료")
        print("✅ CSV 내보내기 선택 완료")

        time.sleep(2)

        # 14단계: 파일 저장 다이얼로그 확인
        print("\n[14단계] 파일 저장 다이얼로그 확인 중...")
        logger.info("14단계: 파일 저장 다이얼로그 대기")

        # 저장 다이얼로그 확인 (선택 사항 - 이미지 있으면 확인)
        save_dialog = os.path.join(kb_config['이미지폴더'], "save_dialog.png")

        if os.path.exists(save_dialog):
            if wait_for_image(save_dialog, timeout=5, confidence=0.5, logger=logger):
                logger.info("✅ 파일 저장 다이얼로그 확인됨")
                print("✅ 파일 저장 다이얼로그 확인됨")
            else:
                logger.warning("⚠️  저장 다이얼로그 이미지를 찾지 못함 - 계속 진행")
                print("⚠️  저장 다이얼로그 이미지를 찾지 못함 - 계속 진행")
        else:
            # 이미지 없으면 그냥 대기
            time.sleep(1)
            print("✅ 파일 저장 다이얼로그 대기 완료 (이미지 확인 생략)")

        # 15단계: 파일 저장
        print("\n[15단계] 파일 저장 중...")
        logger.info("15단계: 파일 저장 경로 및 파일명 입력")

        # 파일명 생성
        today = datetime.now()
        filename = f"{today.year}-{today.month:02d}-{today.day:02d}_KB.csv"
        save_path = os.path.join(save_dir, filename)

        logger.info(f"저장 경로: {save_path}")

        # 파일 경로 입력
        pyperclip.copy(save_path)
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)

        # 저장 버튼 클릭
        pyautogui.press('enter')
        logger.info("파일 저장 명령 완료")

        time.sleep(2)

        # 파일 저장 확인
        if os.path.exists(save_path):
            file_size = os.path.getsize(save_path)
            logger.info(f"✅ 파일 저장 성공: {save_path} ({file_size} bytes)")
            print(f"✅ 파일 저장 성공: {filename} ({file_size} bytes)")
            print("\n" + "=" * 60)
            print("✅✅✅ 데이터 다운로드 테스트 완료!")
            print("=" * 60)
            return True
        else:
            logger.error(f"❌ 파일 저장 실패: {save_path}")
            print(f"❌ [15단계 실패] 파일 저장 실패: {save_path}")
            save_error_screenshot("KB_파일저장실패")
            return False

    except Exception as e:
        logger.error(f"오류 발생: {e}")
        import traceback
        logger.error(traceback.format_exc())
        save_error_screenshot("KB_예외발생")
        return False


if __name__ == "__main__":
    print("=" * 60)
    print("KB증권 데이터 다운로드 테스트 (8단계~15단계)")
    print("=" * 60)
    print()
    print("⚠️  사전 조건:")
    print("  - KB HTS에 이미 로그인되어 있어야 함")
    print("  - 메인 화면에 있어야 함")
    print()
    print("테스트 범위:")
    print("  [8단계] 화면번호 0191 입력")
    print("  [9단계] 전체받기 버튼 클릭")
    print("  [10단계] 데이터 로딩 대기 (7초)")
    print("  [11단계] 테이블 우클릭")
    print("  [12단계] 우클릭 메뉴에서 '파일로 내보내기' 선택")
    print("  [13단계] CSV 내보내기 선택")
    print("  [14단계] 파일 저장 다이얼로그 확인")
    print("  [15단계] 파일 저장 및 확인")
    print()
    print("주의사항:")
    print("1. 테스트 중 마우스/키보드를 사용하지 마세요")
    print("2. 로그는 C:\\HTS_Automation\\logs에 저장됩니다")
    print()

    input("KB HTS에 로그인 완료 후 Enter 키를 누르세요...")

    success = test_kb_download()

    print()
    print("=" * 60)
    if success:
        print("✅✅✅ 데이터 다운로드 테스트 성공!")
        print("CSV 파일 다운로드까지 정상 완료되었습니다.")
    else:
        print("❌ 데이터 다운로드 테스트 실패! 로그를 확인하세요.")
    print("=" * 60)
