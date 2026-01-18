import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random

def get_naver_related_keywords(search_keyword):
    # 1. 크롬 드라이버 설정 (undetected_chromedriver)
    # 크롬 드라이버를 설정한다. 
    # 옵션 값을 설정한다. 
    options = uc.ChromeOptions()
    # options.add_argument('--headless') # 필요 시 주석 해제
    
    driver = uc.Chrome(options=options)
    
    try:
        # 2. 네이버 메인 페이지 접속
        # 네이버 메인 페이지에 접속한다.
        # 드라이버가 대기한다. 
        driver.get("https://www.naver.com")
        
        # 검색창이 뜰 때까지 대기
        wait = WebDriverWait(driver, 10)
        search_input = wait.until(EC.element_to_be_clickable((By.ID, "query")))
        
        # 3. [변경 후] 키워드 사람처럼 입력 (글자별 랜덤 딜레이)
        search_input.click() # 입력 전 클릭
        print(f"'{search_keyword}' 입력 중: ", end="", flush=True)
        
        for char in search_keyword:
            search_input.send_keys(char)
            # 0.1초 ~ 0.3초 사이에서 무작위로 지연 시간 발생
            typing_speed = random.uniform(0.1, 0.3)
            time.sleep(typing_speed)
            print(char, end="", flush=True)
        print("\n입력 완료!")
        
        # 입력 완료 후 자동완성 목록이 갱신될 수 있도록 잠시 대기
        time.sleep(1.0)
        
        # 4. 연관 검색어(자동완성) 추출
        # 입력 완료 후 자동완성 목록이 나타날 때까지 최대 5초 대기
        try:
            print("연관 검색어 로딩 대기 중 (최대 5초)...")

            wait = WebDriverWait(driver, 5)
            wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "li._item[data-keyword]")
                )
            )

            items = driver.find_elements(By.CSS_SELECTOR, "li._item[data-keyword]")
            keywords_list = [
                item.get_attribute("data-keyword")
                for item in items
                if item.get_attribute("data-keyword")
            ]

            if keywords_list:
                safe_name = search_keyword.replace(" ", "_")
                file_name = f"related_{safe_name}.txt"

                with open(file_name, "w", encoding="utf-8") as f:
                    for kw in keywords_list:
                        f.write(kw + "\n")

                print(f"--- 추출 완료 ---")
                print(f"파일명: {file_name}")
                print(f"추출된 키워드: {keywords_list}")
            else:
                print("연관 검색어가 표시되지 않았습니다.")

        except Exception:
            print("⏱️ 5초 동안 연관 검색어가 나타나지 않았습니다.")

            
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
    
    finally:
        # 6. 종료 에러(WinError 6) 방지를 위한 개선된 종료 로직
        try:
            if driver:
                print("\n브라우저를 안전하게 종료합니다...")
                driver.quit()   # close() 하지 말고 quit()만
        except Exception:
            pass
        finally:
            driver = None  # __del__에서 또 quit() 시도하는 걸 최대한 방지  

if __name__ == "__main__":
    user_input = input("네이버에 입력할 키워드를 입력하세요: ")
    if user_input:
        get_naver_related_keywords(user_input)
    else:
        print("입력된 키워드가 없습니다.")