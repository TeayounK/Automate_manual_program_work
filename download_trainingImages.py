# file: automate_imageviewer_pywinauto.py
from pywinauto import Application, timings, Desktop
from pywinauto.keyboard import send_keys
import time
import logging
import os
import sys
import pyautogui
import pandas as pd

import re
import psutil

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(m' \
'essage)s')

SAVE_PATH = "C:\FamilySearch\ImageViewer\TESTFILE" #TrainingLastImage"
SAVE_PATH_CHECK = "C:\FamilySearch\ImageViewer\TESTFILE" # TrainingImages"
EXE_PATH = r"C:\FamilySearch\ImageViewer\ImageViewer.exe"  
USERNAME = "****"
PASSWORD = "****"

def click_confirmation_dialog(timeout=10):
    """
    로그인 후 안내창의 '확인' 버튼 자동 클릭
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            windows = Desktop(backend="uia").windows()
            for win in windows:
                title = win.window_text()
                if "" in title:
                    logging.info(f"안내창 발견: {title}")
                    try:
                        # 버튼 탐색: element_info 사용
                        for ctrl in win.descendants():
                            ctrl_type = ctrl.element_info.control_type
                            text = ctrl.window_text()
                            if ctrl_type == "Button" and text in ["OK"]:
                                logging.info(f"Found button: {text}")
                                ctrl.click_input()
                                logging.info("Click Sucessful")
                                return True
                    except Exception as e:
                        logging.warning("Fail to find button: %s", e)
        except Exception as e:
            logging.warning("Searching Window Failed: %s", e)
        time.sleep(0.5)

    logging.warning("Failed to find Window")
    return False


def click_cancel(main_win):
    button_controls = main_win.descendants(control_type="Button")

    for button in button_controls:
        label_text = button.window_text().strip()
        if "cancel" in label_text.lower():
            logging.info(f"매칭된 버튼: -{label_text}-")
            button.click_input()
            break


def submit_token_from_embedded_viewer(main_win, token_id):
    logging.info("토큰 입력창 탐색 중...")

    try:
        # 모든 Edit 필드를 가져옴
        edit_controls = main_win.descendants(control_type="Edit")
       
        # 라벨 텍스트 기준으로 찾기
        for edit in edit_controls:
            label_text = edit.window_text().strip()
            if "film number" in label_text.lower():
                logging.info(f"매칭된 입력창: {label_text}")
                edit.set_edit_text(token_id)
                break
        else:
            logging.error("토큰 입력창을 찾을 수 없음")
            return False

        # OK 버튼 찾기
        button_controls = main_win.descendants(control_type="Button")
       
        for button in button_controls:
            label_text = button.window_text().strip()
            if "ok" in label_text.lower():
                logging.info(f"매칭된 버튼: -{label_text}-")
                button.click_input()
                break
        else:
            logging.error("토큰 입력창을 찾을 수 없음")
            return False
        return True


    except Exception as e:
        logging.error("토큰 제출 실패: %s", e)
        return False


def submit_index_number(main_win, index_id):
    logging.info("인덱스 입력창 탐색 중...")

    try:
        # 모든 Edit 필드를 가져옴
        edit_controls = main_win.descendants(control_type="Edit")
       
        # 라벨 텍스트 기준으로 찾기
        for edit in edit_controls:
            label_text = edit.window_text().strip()
            if "enter image index" in label_text.lower():
                logging.info(f"매칭된 입력창: {label_text}")
                edit.set_edit_text(index_id)
                break
        else:
            logging.error("토큰 입력창을 찾을 수 없음")
            return False

        # OK 버튼 찾기
        button_controls = main_win.descendants(control_type="Button")
       
        for button in button_controls:
            label_text = button.window_text().strip()
            if "ok" in label_text.lower():
                logging.info(f"매칭된 버튼: -{label_text}-")
                button.click_input()
                break
        else:
            logging.error("토큰 입력창을 찾을 수 없음")
            return False
        return True
   

    except Exception as e:
        logging.error("토큰 제출 실패: %s", e)
        return False

def debug_save_dialog(main_win):
    print("[INFO] 저장창 컨트롤 구조 탐색 시작...\n")
    for i, ctrl in enumerate(main_win.descendants()):
        print(f"[{i}] Type: {ctrl.element_info.control_type:<12} | "
              f"Text: '{ctrl.window_text()}' | "
              f"AutomationId: {ctrl.element_info.automation_id} | "
              f"ClassName: {ctrl.element_info.class_name}")
       

def save_image(main_win,found_image,SAVE_PATH,file_name):
    logging.info("Saving an image...")
    if found_image:
        button_controls = main_win.descendants(control_type="Button")
       
        # 라벨 텍스트 기준으로 찾기
        for button in button_controls:
            label_text = button.window_text().strip()
            if "save" == label_text.lower():
                logging.info(f"매칭된 입력창: {label_text}")
                button.click_input()
                break    
       
        # saving the image
        full_path = "\\".join([SAVE_PATH,file_name])+".jpg"
        edit_controls = main_win.descendants(control_type="Edit")
        for edit in edit_controls:
            label = edit.window_text().strip().lower()
            if "file name" in label or "파일 이름" in label or edit.element_info.automation_id == "1148":
                if not os.path.exists(full_path):
                    edit.set_focus()
                    send_keys('^a')  # 전체 선택
                    send_keys('{BACKSPACE}')  # 삭제
                    send_keys(full_path)  # 전체 경로 입력
                    time.sleep(0.5)

                    # 4. 저장 (Enter)
                    send_keys('{ENTER}')
                    logging.info(f"[INFO] 저장 요청 완료: {full_path}")
                else:
                    pyautogui.press('esc')
                    break

        time.sleep(0.5)


def login():
    try:
        logging.info("Start ImageViewer...")
        app = Application(backend="uia").start(EXE_PATH)  # 또는 connect(path=...) if already running
        # 앱 메인 윈도우 찾기 (타이틀에 따라 조정)
        dlg = app.window(title_re="FamilySearch Login")  
        dlg.wait("visible enabled ready", timeout=15)
        logging.info("Found it: %s", dlg)
        # 로그인 대화상자/패널 찾기 (구조에 따라 조정)
        # Inspect.exe로 실제 컨트롤 이름을 확인하고 아래를 수정하세요.
        # 예시1: 자동 노출된 Edit 컨트롤 사용
        try:
            username_edit = dlg.child_window(control_type="Edit", found_index=0)
            password_edit = dlg.child_window(control_type="Edit", found_index=1)
            login_btn = dlg.child_window(control_type="Button", found_index=0)
            login_success = dlg.child_window(auto_id="btnOk", control_type="Button")
        except Exception:
            # 대체: 인덱스 기반 (비권장, UI 변경 시 깨짐)
            username_edit = dlg.child_window(control_type="Edit", found_index=0)
            password_edit = dlg.child_window(control_type="Edit", found_index=1)
            login_btn = dlg.child_window(control_type="Button", found_index=0)

        logging.info("Putting ID...")
        username_edit.wait("visible enabled", timeout=10)
        username_edit.set_edit_text(USERNAME)

        logging.info("Putting Password...")
        password_edit.wait("visible enabled", timeout=10)
        password_edit.set_edit_text(PASSWORD)

        logging.info("Login Button Click...")
        login_btn.wait("visible enabled", timeout=10)
        login_btn.click_input()

        # 로그인 성공을 기다리기
        timings.wait_until_passes(15, 1, lambda: dlg.child_window(title_re=".*Gallery|.*Home.*").exists())

        logging.info("Login Successful")
       
        # Click confirmation button
        click_confirmation_dialog()
        # send_keys('{ENTER}') 이거로도 됨ㅋㅋ
        logging.info("Ready for work!")
        return app
    except Exception as e:
        logging.exception("실패: %s", e)
        sys.exit(1)
        return None

def save_pages(first,last,main_win,token_id):
    if first.isdigit and last.isdigit():
        # get the page numbers
        print("page diff: ",int(last) - int(first))
        if abs(int(last) - int(first)) > 3:
            pages = [first,str(int(first)+1),str(int(last)-1),last]
        else:
            pages = [first,last]
        pages = [last] # TODO: remove this after you doen with downloading last images
        # for each page numbers
        for page in pages:
            # try:
           
            main_win.menu_select("File->Goto Image")

            file_name = "_".join([token_id,page])
           
            found_image = submit_index_number(main_win,page)
            save_image(main_win,found_image,SAVE_PATH,file_name)
            # except Exception as e:
            #     logging.exception(page,"page doesn't work: ", e)
            #     send_keys('{ENTER}')
            #     pass




def main():
    skipped = []
    existing_imgs = set([im.split("_")[0] for im in os.listdir(SAVE_PATH_CHECK)])
    existing_imgs = existing_imgs.union(set([im.split("_")[0] for im in os.listdir(SAVE_PATH)]))
   
    # login into the image viewer
    app = login()

    # load genetic record data
    genetic_record = pd.read_excel("Wed1_29_Final_fix.xlsx")
   
    try:
        # downloading images
        # Set main window
        main_win = app.window(title_re="FamilySearch Image Viewer:*")
        proc = psutil.Process(main_win.process_id())
        for idx, row in genetic_record.iterrows():
            NGID = str(row["25 Natural Group Id"])
            first_page = row["First Person Name (as written)"]
            last_page = row["Last Person Name (as written)"]

            first_page_ = re.search("(\d+)-",first_page)
            last_page_ = re.search("(\d+)-",last_page)

            if first_page_ and last_page_:
                first = first_page_.group(1)
                last = last_page_.group(1)
            else:
                continue

            if NGID.isdigit() and NGID not in existing_imgs and int(last) < 800 and NGID not in ["108293318","107519280","107519282","107519284","107524437"]:
                               
                try:
                    # 메뉴 클릭하기
                    main_win.menu_select("File->Distribution Store Viewer")
                    submit_token_from_embedded_viewer(main_win, token_id=NGID)
                    main_win.wait("visible enabled",timeout=1)
                except Exception as e:
                    logging.exception(NGID,"NGID doesn't work: ", e)
                    send_keys('{ENTER}')
                    click_cancel(main_win)
                    time.sleep(30)
                    continue

               
                logging.info(f"NGID: {NGID}")
                logging.info(f"First: {first}")
                logging.info(f"Last: {last}")
                while True:
                    cpu = proc.cpu_percent(interval=1)
                    if cpu < 0.2:  # assume idle if below 2%
                        break
                    print(f"Waiting... CPU {cpu:.1f}%")
                time.sleep(2)
                try:
                    save_pages(first,last,main_win,NGID)
                except:
                    send_keys('{ENTER}')
                    skipped.append(idx)
                    continue


    except Exception as e:
        logging.exception("ddd 실패: ", e)
        print("skipped: ",skipped)
        sys.exit(1)
   
    print("skipped: ",skipped)

if __name__ == "__main__":
    main()
