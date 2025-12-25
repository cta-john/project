"""
KB증권 자동화 테스트 스크립트
"""

import os
import sys
import time
import ctypes
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
        print("\n[준비] 설정 파일 로드 중...")
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
        print("\n[1단계] KB HTS 실행 중...")
        logger.info("1단계: KB HTS 실행")
        minimize_all_windows()

        proc = Popen([kb_config['실행경로']], stdin=PIPE)
        logger.info(f"HTS 실행 완료 (PID: {proc.pid})")
        print(f"✅ HTS 실행 완료 (PID: {proc.pid})")
        time.sleep(5)

        # 2단계: 로그인 화면 대기
        print("\n[2단계] 로그인 화면 대기 중...")
        logger.info("2단계: 로그인 화면 대기...")
        time.sleep(3)
        print("✅ 로그인 화면 대기 완료")

        # 3단계: ID 입력
        print("\n[3단계] ID 입력 중...")
        logger.info("3단계: ID 입력")

        # ID 입력창 이미지로 찾아서 클릭 (모니터 환경 독립적)
        id_input_field = os.path.join(kb_config['이미지폴더'], "id_input_field.png")

        if click_at_image(id_input_field, timeout=5, confidence=0.5, logger=logger):
            logger.info("✅ ID 입력창 클릭 성공")
            time.sleep(0.5)

            # ID 입력
            enter_text_fast(kb_account['id'])
            logger.info(f"ID 입력 완료")
            print(f"✅ ID 입력 완료")
            time.sleep(0.5)
        else:
            logger.error("❌ ID 입력창을 찾지 못했습니다!")
            print("❌ [3단계 실패] ID 입력창을 찾지 못했습니다!")
            save_error_screenshot("KB_ID입력창찾기실패")
            return False

        # 4단계: PW 입력
        print("\n[4단계] PW 입력 중...")
        logger.info("4단계: PW 입력")

        # Tab 키로 비밀번호 입력창으로 이동
        pyautogui.press('tab')
        time.sleep(0.3)

        # PW 입력
        enter_text_fast(kb_account['password'])
        logger.info("✅ PW 입력 완료 (Tab 키 사용)")
        print("✅ PW 입력 완료")
        time.sleep(0.5)

        # 5단계: 로그인 버튼 클릭
        print("\n[5단계] 로그인 버튼 클릭 중...")
        logger.info("5단계: 로그인 버튼 클릭")
        login_button = os.path.join(kb_config['이미지폴더'], "login_button.png")

        if click_at_image(login_button, timeout=5, confidence=0.5, logger=logger):
            logger.info("✅ 로그인 버튼 클릭 성공")
            print("✅ 로그인 버튼 클릭 완료")
        else:
            logger.warning("로그인 버튼 이미지를 찾지 못함 - Enter 시도")
            print("⚠️  로그인 버튼 이미지를 찾지 못함 - Enter 키로 시도")
            pyautogui.press('enter')

        time.sleep(2)

        # 6단계: 조회전용안내 팝업 대기
        print("\n[6단계] 조회전용안내 팝업 대기 중...")
        logger.info("6단계: 조회전용안내 팝업 대기...")
        popup_img = os.path.join(kb_config['이미지폴더'], "readonly_notice.png")

        if wait_for_image(popup_img, timeout=30, confidence=0.5, logger=logger):
            logger.info("✅ 조회전용안내 팝업 발견!")
            print("✅ 조회전용안내 팝업 발견")
            time.sleep(1)

            # 예(Y) 버튼 클릭
            yes_button = os.path.join(kb_config['이미지폴더'], "yes_button.png")
            if click_at_image(yes_button, timeout=5, confidence=0.5, logger=logger):
                logger.info("✅ 예(Y) 버튼 클릭 성공")
                print("✅ 예(Y) 버튼 클릭 완료")
            else:
                logger.warning("예(Y) 버튼을 찾지 못함 - Enter 키 시도")
                print("⚠️  예(Y) 버튼을 찾지 못함 - Enter 키로 시도")
                pyautogui.press('enter')

            time.sleep(2)
        else:
            logger.error("❌ 조회전용안내 팝업을 찾지 못했습니다!")
            print("❌ [6단계 실패] 조회전용안내 팝업을 찾지 못했습니다!")
            save_error_screenshot("KB_로그인실패")
            return False

        # 7단계: 메인 화면 진입 확인
        print("\n[7단계] 메인 화면 진입 확인 중...")
        logger.info("7단계: 메인 화면 진입 확인...")
        logo_img = os.path.join(kb_config['이미지폴더'], "hts_logo.png")

        if wait_for_image(logo_img, timeout=20, confidence=0.5, logger=logger):
            logger.info("✅ 메인 화면 진입 성공!")
            print("✅ 메인 화면 진입 성공!")
        else:
            logger.error("❌ 메인 화면 진입 실패!")
            print("❌ [7단계 실패] 메인 화면 진입 실패!")
            save_error_screenshot("KB_메인화면실패")
            return False

        time.sleep(2)

        # 8단계: 화면번호 0191 입력
        print("\n[8단계] 화면번호 0191 입력 중...")
        logger.info("8단계: 화면번호 0191 입력")

        # 화면번호 입력창 이미지로 찾아서 클릭 (모니터 환경 독립적)
        screen_number_field = os.path.join(kb_config['이미지폴더'], "screen_number_field.png")

        if click_at_image(screen_number_field, timeout=5, confidence=0.5, logger=logger):
            logger.info("✅ 화면번호 입력창 클릭 성공")
            time.sleep(0.5)

            # 0191 입력
            enter_text_fast("0191")
            time.sleep(0.3)
            pyautogui.press('enter')
            logger.info("0191 입력 완료")
            print("✅ 화면번호 0191 입력 완료")

            time.sleep(3)
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

        # 테이블 영역 이미지로 찾아서 우클릭 (모니터 환경 독립적)
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

        # 우클릭 메뉴 확인 (선택 사항 - 이미지 있으면 확인)
        context_menu = os.path.join(kb_config['이미지폴더'], "context_menu_export.png")

        # 메뉴가 뜨는 시간 대기
        time.sleep(0.5)

        # 이미지로 메뉴 확인 (있으면)
        if os.path.exists(context_menu):
            if wait_for_image(context_menu, timeout=3, confidence=0.5, logger=logger):
                logger.info("✅ 우클릭 메뉴 확인됨")
                print("✅ 우클릭 메뉴 확인됨")
            else:
                logger.warning("⚠️  우클릭 메뉴 이미지를 찾지 못함 - 계속 진행")
                print("⚠️  우클릭 메뉴 이미지를 찾지 못함 - 계속 진행")

        # 키보드로 메뉴 이동
        for _ in range(3):  # "파일로 내보내기"까지 이동
            pyautogui.press('down')
            time.sleep(0.1)

        # 우측 화살표로 서브메뉴 열기
        pyautogui.press('right')
        logger.info("파일로 내보내기 서브메뉴 열기")
        print("✅ 파일로 내보내기 선택")
        time.sleep(0.5)

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
            print("✅✅✅ 전체 자동화 테스트 완료!")
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


def is_admin():
    """관리자 권한으로 실행 중인지 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    """관리자 권한으로 스크립트 재실행"""
    try:
        if not is_admin():
            print("⚠️  관리자 권한이 필요합니다. 관리자 모드로 재실행합니다...")
            print()

            # UAC 프롬프트를 통해 관리자 권한으로 재실행
            script_path = os.path.abspath(sys.argv[0])

            ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                sys.executable,
                f'"{script_path}"',
                None,
                1  # SW_SHOWNORMAL
            )
            sys.exit(0)
        else:
            print("✅ 관리자 권한으로 실행 중")
            print()
    except Exception as e:
        print(f"❌ 관리자 권한 승격 실패: {e}")
        print("PowerShell을 관리자 모드로 실행한 후 다시 시도하세요.")
        sys.exit(1)


if __name__ == "__main__":
    # 관리자 권한 확인 및 자동 승격 (가장 먼저 실행)
    run_as_admin()

    print("=" * 60)
    print("KB증권 전체 자동화 테스트 (1단계~15단계)")
    print("=" * 60)
    print()
    print("테스트 범위:")
    print("  [1단계] HTS 실행")
    print("  [2단계] 로그인 화면 대기")
    print("  [3단계] ID 입력")
    print("  [4단계] PW 입력")
    print("  [5단계] 로그인 버튼 클릭")
    print("  [6단계] 조회전용안내 팝업 처리")
    print("  [7단계] 메인 화면 진입 확인")
    print("  [8단계] 화면번호 0191 입력")
    print("  [9단계] 전체받기 버튼 클릭")
    print("  [10단계] 데이터 로딩 대기 (7초)")
    print("  [11단계] 테이블 우클릭")
    print("  [12단계] 우클릭 메뉴에서 '파일로 내보내기' 선택")
    print("  [13단계] CSV 내보내기 선택")
    print("  [14단계] 파일 저장 다이얼로그 확인")
    print("  [15단계] 파일 저장 및 확인")
    print()
    print("⚠️  사전 준비:")
    print("  - HTS 기본 세팅: 아이디 탭 활성화")
    print("  - HTS 기본 세팅: 조회전용 체크")
    print()
    print("주의사항:")
    print("1. 테스트 중 마우스/키보드를 사용하지 마세요")
    print("2. 로그는 C:\\HTS_Automation\\logs에 저장됩니다")
    print("3. 이미지 기반 인식으로 모니터 환경 변경 시에도 작동합니다")
    print()

    input("준비되면 Enter 키를 누르세요...")

    success = test_kb_automation()

    print()
    print("=" * 60)
    if success:
        print("✅✅✅ 전체 자동화 테스트 성공!")
        print("CSV 파일 다운로드까지 정상 완료되었습니다.")
    else:
        print("❌ 자동화 테스트 실패! 로그를 확인하세요.")
    print("=" * 60)
