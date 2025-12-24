"""
HTS 자동화 유틸리티 모듈
- 로깅 시스템
- 공통 함수 (이미지 인식, 마우스/키보드 조작, 대기 등)
- 에러 처리
"""

import time
import logging
import os
import json
from datetime import datetime
from pathlib import Path

import pyautogui
import pyperclip
import cv2
import numpy as np
from PIL import ImageGrab


# ================================================================================
# 로깅 설정
# ================================================================================

def setup_logging(log_dir="C:\\HTS_Automation\\logs"):
    """로깅 시스템 초기화"""
    # 로그 디렉토리 생성
    Path(log_dir).mkdir(parents=True, exist_ok=True)

    # 로그 파일명 (날짜별)
    log_filename = os.path.join(log_dir, f"{datetime.now().strftime('%Y-%m-%d')}.log")

    # 로거 설정
    logger = logging.getLogger('HTS_Automation')
    logger.setLevel(logging.DEBUG)

    # 기존 핸들러 제거 (중복 방지)
    logger.handlers = []

    # 파일 핸들러
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    file_handler.setFormatter(file_formatter)

    # 콘솔 핸들러
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    console_handler.setFormatter(console_formatter)

    # 핸들러 추가
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


def save_error_screenshot(증권사명, log_dir="C:\\HTS_Automation\\logs"):
    """에러 발생 시 스크린샷 저장"""
    try:
        Path(log_dir).mkdir(parents=True, exist_ok=True)
        screenshot = ImageGrab.grab()
        screenshot_path = os.path.join(
            log_dir,
            f"{증권사명}_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )
        screenshot.save(screenshot_path)
        return screenshot_path
    except Exception as e:
        return None


# ================================================================================
# 설정 파일 로드
# ================================================================================

def load_config(config_path="C:\\HTS_Automation\\config\\hts_config.json"):
    """HTS 설정 파일 로드"""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_account_info(info_path="C:\\HTS_Automation\\config\\hts_info.json"):
    """계정 정보 파일 로드"""
    with open(info_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ================================================================================
# 이미지 인식 함수
# ================================================================================

def imglocation(path, confidence=0.7):
    """
    화면에서 이미지를 찾아 중심 좌표 반환

    Args:
        path: 찾을 이미지 파일 경로
        confidence: 매칭 신뢰도 (0.0 ~ 1.0)

    Returns:
        tuple: (x, y) 좌표 또는 None
    """
    try:
        img = cv2.imread(path)
        if img is None:
            return None

        img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        screenshot = ImageGrab.grab()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

        result = cv2.matchTemplate(screenshot, img_gray, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        if max_val >= confidence:
            h, w = img_gray.shape
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            return (center_x, center_y)

        return None
    except Exception as e:
        return None


def imglocation_multi(image_paths, confidence=0.7):
    """
    여러 이미지 중 하나라도 찾으면 반환 (버전 호환성)

    Args:
        image_paths: 이미지 경로 리스트
        confidence: 매칭 신뢰도

    Returns:
        tuple: (x, y) 좌표 또는 None
    """
    for img_path in image_paths:
        result = imglocation(img_path, confidence)
        if result:
            return result
    return None


def wait_for_image(image_path, timeout=30, confidence=0.7, check_interval=0.5, logger=None):
    """
    이미지가 화면에 나타날 때까지 대기

    Args:
        image_path: 찾을 이미지 경로
        timeout: 최대 대기 시간 (초)
        confidence: 매칭 신뢰도
        check_interval: 체크 간격 (초)
        logger: 로거 객체

    Returns:
        tuple: (x, y) 좌표 또는 None (타임아웃)
    """
    start_time = time.time()

    while time.time() - start_time < timeout:
        result = imglocation(image_path, confidence)
        if result:
            if logger:
                logger.debug(f"이미지 발견: {os.path.basename(image_path)} at {result}")
            return result
        time.sleep(check_interval)

    if logger:
        logger.warning(f"이미지 타임아웃: {os.path.basename(image_path)} ({timeout}초)")
    return None


# ================================================================================
# 마우스/키보드 조작 함수
# ================================================================================

def click_at_coord(x, y, delay=0.1):
    """좌표 클릭"""
    pyautogui.moveTo(x, y)
    time.sleep(delay)
    pyautogui.click()


def click_at_image(image_path, timeout=10, confidence=0.7, logger=None):
    """이미지를 찾아서 클릭"""
    coord = wait_for_image(image_path, timeout, confidence, logger=logger)
    if coord:
        click_at_coord(coord[0], coord[1])
        return True
    return False


def press_keyboard(text, delay=0.1):
    """
    키보드 입력 (한 글자씩 천천히)

    Args:
        text: 입력할 텍스트
        delay: 글자 간 딜레이 (초)
    """
    for char in str(text):
        pyperclip.copy(char)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(delay)


def enter_text_fast(text):
    """빠른 텍스트 입력 (클립보드 사용)"""
    pyperclip.copy(str(text))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)


# ================================================================================
# 대기 함수 (sleep 최적화)
# ================================================================================

def smart_wait(image_path=None, max_wait=10, min_wait=1, check_interval=0.3, confidence=0.7):
    """
    스마트 대기: 이미지가 나타나면 즉시 반환, 아니면 max_wait까지 대기

    Args:
        image_path: 대기할 이미지 (None이면 min_wait만큼 고정 대기)
        max_wait: 최대 대기 시간
        min_wait: 최소 대기 시간 (이미지 없을 때)
        check_interval: 체크 간격
        confidence: 이미지 매칭 신뢰도

    Returns:
        bool: 이미지를 찾았는지 여부
    """
    if image_path is None:
        time.sleep(min_wait)
        return False

    start = time.time()
    while time.time() - start < max_wait:
        if imglocation(image_path, confidence):
            # 최소 대기 시간 보장
            elapsed = time.time() - start
            if elapsed < min_wait:
                time.sleep(min_wait - elapsed)
            return True
        time.sleep(check_interval)

    return False


# ================================================================================
# HTS 제어 함수
# ================================================================================

def minimize_all_windows():
    """모든 창 최소화 (바탕화면 보기)"""
    pyautogui.hotkey('win', 'd')
    time.sleep(0.5)


def get_image_path(증권사명, 이미지명, base_path="C:\\HTS_Automation\\images"):
    """
    이미지 파일 경로 생성

    Args:
        증권사명: 증권사 이름 (예: "키움")
        이미지명: 이미지 파일명 (예: "로그인_완료.png")
        base_path: 이미지 기본 경로

    Returns:
        str: 전체 이미지 경로
    """
    return os.path.join(base_path, 증권사명, 이미지명)


# ================================================================================
# 기타 유틸리티
# ================================================================================

def create_directory(path):
    """디렉토리 생성 (존재하면 무시)"""
    Path(path).mkdir(parents=True, exist_ok=True)


def get_today_str():
    """오늘 날짜 문자열 반환 (YYYY.M.D 형식)"""
    today = datetime.now()
    return f"{today.year}.{today.month}.{today.day}"


def get_file_search_pattern(증권사명, 파일종류):
    """
    파일 검색 패턴 생성

    Args:
        증권사명: 증권사 이름
        파일종류: 파일 종류 (예: "증거금20종목")

    Returns:
        str: 검색 패턴 (예: "증거금20종목키움_2025.12.24*.csv")
    """
    today = get_today_str()
    return f"{파일종류}{증권사명}_{today}*"
