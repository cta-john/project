"""
KB증권 자동화 테스트 스크립트
"""

import os
import sys
import time
from datetime import datetime
from subprocess import Popen, PIPE

# hts_utils 임포트
sys.path.append(r"C:\HTS_Automation")
from hts_utils import (
    setup_logging,
    load_config,
    load_account_info,
    imglocation,
    wait_for_image,
    click_at_image,
    press_keyboard,
    enter_text_fast,
    minimize_all_windows,
    get_image_path,
    save_error_screenshot
)

import pyautogui
import pyperclip


def test_kb_automation():
    """KB증권 자동화 테스트"""

    # 로깅 설정
    logger = setup_logging()
    logger.info("=" * 50)
    logger.info("KB증권 자동화 테스트 시작")
    logger.info("=" * 50)

    try:
        # 설정 파일 로드
        logger.info("설정 파일 로드 중...")
        config = load_config()
        account_info = load_account_info()

        kb_config = config['KB']
        kb_account = account_info['KB']

        logger.info(f"HTS 경로: {kb_config['실행경로']}")
        logger.info(f"이미지 폴더: {kb_config['이미지폴더']}")

        # 저장 폴더 생성
        save_dir = r"C:\Users\kiwoom\Documents\종투사증거금취합"
        os.makedirs(save_dir, exist_ok=True)
        logger.info(f"저장 폴더: {save_dir}")

        # 1단계: HTS 실행
        logger.info("1단계: KB HTS 실행")
        minimize_all_windows()

        proc = Popen([kb_config['실행경로']], stdin=PIPE)
        logger.info(f"HTS 실행 완료 (PID: {proc.pid})")
        time.sleep(5)

        # 2단계: 로그인 화면 대기
        logger.info("2단계: 로그인 화면 대기...")
        # 로그인 창이 뜰 때까지 대기 (hts_logo나 특정 요소)
        time.sleep(3)

        # 3단계: 아이디 탭 확인 및 클릭
        logger.info("3단계: 아이디 탭 확인 및 클릭")

        # 아이디 탭이 이미 활성화되어 있는지 확인
        id_tab_active = os.path.join(kb_config['이미지폴더'], "아이디_탭_활성화.png")
        id_tab_inactive = os.path.join(kb_config['이미지폴더'], "아이디_탭_비활성화.png")

        # 활성화 상태 확인
        if imglocation(id_tab_active, confidence=0.7):
            logger.info("아이디 탭이 이미 활성화되어 있음 - 스킵")
        elif imglocation(id_tab_inactive, confidence=0.7):
            logger.info("아이디 탭이 비활성화되어 있음 - 클릭 시도")
            if click_at_image(id_tab_inactive, timeout=5, logger=logger):
                logger.info("아이디 탭 클릭 성공")
            else:
                logger.warning("아이디 탭 클릭 실패 - 좌표로 시도")
                pyautogui.click(936, 75)
        else:
            logger.warning("아이디 탭을 찾지 못함 - 좌표로 시도")
            pyautogui.click(936, 75)

        time.sleep(1)

        # 3-1단계: 조회전용 체크박스 확인 및 체크
        logger.info("3-1단계: 조회전용 체크박스 확인")

        # 체크 상태 확인
        check_checked = os.path.join(kb_config['이미지폴더'], "조회전용_체크됨.png")
        check_unchecked = os.path.join(kb_config['이미지폴더'], "조회전용_체크안됨.png")

        if imglocation(check_checked, confidence=0.8):
            logger.info("조회전용 체크박스가 이미 체크되어 있음 - 스킵")
        elif imglocation(check_unchecked, confidence=0.8):
            logger.info("조회전용 체크박스가 체크 안 됨 - 체크 시도")
            if click_at_image(check_unchecked, timeout=5, logger=logger):
                logger.info("조회전용 체크박스 체크 성공")
            else:
                logger.warning("조회전용 체크박스를 찾지 못함")
        else:
            logger.warning("조회전용 체크박스 이미지를 찾지 못함")

        time.sleep(0.5)

        # 4단계: ID 입력
        logger.info("4단계: ID 입력")
        # ID 입력창 클릭 (대략적인 위치)
        pyautogui.click(837, 173)
        time.sleep(0.5)

        # ID 입력
        enter_text_fast(kb_account['id'])
        logger.info(f"ID 입력 완료")
        time.sleep(0.5)

        # 5단계: PW 입력
        logger.info("5단계: PW 입력")
        pyautogui.press('tab')  # 비밀번호 입력창으로 이동
        time.sleep(0.3)

        # PW 입력
        enter_text_fast(kb_account['password'])
        logger.info("PW 입력 완료")
        time.sleep(0.5)

        # 6단계: 로그인 버튼 클릭
        logger.info("6단계: 로그인 버튼 클릭")
        login_button = os.path.join(kb_config['이미지폴더'], "로그인_버튼.png")

        if click_at_image(login_button, timeout=5, logger=logger):
            logger.info("로그인 버튼 클릭 성공")
        else:
            logger.warning("로그인 버튼 이미지를 찾지 못함 - Enter 시도")
            pyautogui.press('enter')

        time.sleep(2)

        # 7단계: 조회전용안내 팝업 대기
        logger.info("7단계: 조회전용안내 팝업 대기...")
        popup_img = os.path.join(kb_config['이미지폴더'], "조회전용안내.png")

        if wait_for_image(popup_img, timeout=30, logger=logger):
            logger.info("조회전용안내 팝업 발견!")
            time.sleep(1)

            # 예(Y) 버튼 클릭
            yes_button = os.path.join(kb_config['이미지폴더'], "예_버튼.png")
            if click_at_image(yes_button, timeout=5, logger=logger):
                logger.info("예(Y) 버튼 클릭 성공")
            else:
                logger.warning("예(Y) 버튼을 찾지 못함 - Enter 키 시도")
                pyautogui.press('enter')

            time.sleep(2)
        else:
            logger.error("조회전용안내 팝업을 찾지 못했습니다!")
            save_error_screenshot("KB_로그인실패")
            return False

        # 8단계: 메인 화면 진입 확인
        logger.info("8단계: 메인 화면 진입 확인...")
        logo_img = os.path.join(kb_config['이미지폴더'], "hts_logo.png")

        if wait_for_image(logo_img, timeout=20, logger=logger):
            logger.info("메인 화면 진입 확인!")
        else:
            logger.error("메인 화면 진입 실패!")
            save_error_screenshot("KB_메인화면실패")
            return False

        time.sleep(2)

        # 9단계: 화면번호 0191 입력
        logger.info("9단계: 화면번호 0191 입력")

        # 화면번호 입력창 클릭 (좌측 상단 - 좌표는 임시, 이미지 인식으로 변경 가능)
        # 임시로 좌표 사용 (나중에 이미지로 대체)
        pyautogui.click(60, 47)  # 대략적인 위치
        time.sleep(0.5)

        # 0191 입력
        enter_text_fast("0191")
        time.sleep(0.3)
        pyautogui.press('enter')
        logger.info("0191 입력 완료")

        time.sleep(3)

        # 10단계: 전체받기 버튼 클릭
        logger.info("10단계: 전체받기 버튼 클릭")
        download_button = os.path.join(kb_config['이미지폴더'], "전체받기_버튼.png")

        if click_at_image(download_button, timeout=10, logger=logger):
            logger.info("전체받기 버튼 클릭 성공")
        else:
            logger.error("전체받기 버튼을 찾지 못했습니다!")
            save_error_screenshot("KB_전체받기실패")
            return False

        # 11단계: 데이터 로딩 대기 (다음 버튼 비활성화 확인)
        logger.info("11단계: 데이터 로딩 대기 (10초)...")
        time.sleep(10)  # 실제로는 다음 버튼 비활성화 이미지 확인

        # 12단계: 테이블 우클릭 및 파일 저장
        logger.info("12단계: 테이블 우클릭 및 파일 저장")

        # 테이블 중앙 클릭 (대략적인 위치)
        table_x, table_y = 300, 300  # 임시 좌표
        pyautogui.click(table_x, table_y)
        time.sleep(0.5)

        # 우클릭
        pyautogui.rightClick(table_x, table_y)
        logger.info("테이블 우클릭 완료")
        time.sleep(1)

        # 파일로 내보내기 메뉴로 이동
        logger.info("파일로 내보내기 선택 중...")
        # 키보드로 메뉴 이동 (더 안정적)
        for _ in range(3):  # 다이렉트 보기까지 3번 아래
            pyautogui.press('down')
            time.sleep(0.1)

        # 파일로 내보내기 선택 (우측 화살표)
        pyautogui.press('right')
        time.sleep(0.3)

        # CSV 내보내기 선택
        pyautogui.press('down')  # 첫 번째 항목
        time.sleep(0.1)
        pyautogui.press('enter')
        logger.info("CSV 내보내기 선택 완료")

        time.sleep(2)

        # 13단계: 파일 저장 대화상자
        logger.info("13단계: 파일 저장 경로 및 파일명 입력")

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
        logger.info("파일 저장 완료!")

        time.sleep(2)

        # 14단계: 파일 저장 확인
        if os.path.exists(save_path):
            file_size = os.path.getsize(save_path)
            logger.info(f"✅ 파일 저장 성공: {save_path} ({file_size} bytes)")
            return True
        else:
            logger.error(f"❌ 파일 저장 실패: {save_path}")
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
    print("KB증권 자동화 테스트")
    print("=" * 60)
    print()
    print("주의사항:")
    print("1. 테스트 중 마우스/키보드를 사용하지 마세요")
    print("2. 화면 해상도를 변경하지 마세요")
    print("3. 로그는 C:\\HTS_Automation\\logs에 저장됩니다")
    print()

    input("준비되면 Enter 키를 누르세요...")

    success = test_kb_automation()

    print()
    print("=" * 60)
    if success:
        print("✅ 테스트 성공!")
    else:
        print("❌ 테스트 실패! 로그를 확인하세요.")
    print("=" * 60)
