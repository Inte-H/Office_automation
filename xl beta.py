import datetime
import os

import xlwings as xw
from xlwings.constants import AutoFillType
from selenium import webdriver
from selenium.webdriver.support.select import Select
from datetime import date
import re
import time
import pandas as pd

download = 'C:/Users/NAH01080005/Downloads/req_mng_print.xls'


def get_cell(workbook, column_alphabet, down=0) -> str:
    cell = column_alphabet + str(last_row(workbook, column_alphabet) + down)
    return cell


def last_row(workbook, column_alphabet) -> int:
    '''

    :param workbook: xlwings WorkBook 객체
    :param column_alphabet: 마지막 행을 구할 열의 알파벳
    :return: 맨 마지막 행 번수를 리턴
    '''
    sheet = workbook.sheets[0]
    origin_cell = sheet.range(column_alphabet + '2').end('down').row
    return origin_cell


def apply_formula(workbook, column_alphabets, append_num) -> None:
    '''

    :param workbook: xlwings WorkBook 객체
    :param column_alphabets: 자동서식채우기를 할 알파벳 열
    :param append_num: 채울 갯수
    '''
    for column_alphabet in column_alphabets:
        start = get_cell(workbook, column_alphabet)
        end = get_cell(workbook, column_alphabet, append_num)
        workbook.sheets[0].range(start).api.AutoFill(workbook.sheets[0].range(start + ':' + end).api,
                                                     AutoFillType.xlFillDefault)


def check_req_and_delete():
    if os.path.isfile(download):
        xw.Book(download).close()
        os.remove(download)


today = date.today().strftime('%y.%m.%d')
yesterday = (date.today() - datetime.timedelta(days=1)).strftime('%y.%m.%d')
friday = (date.today() - datetime.timedelta(days=3)).strftime('%y.%m.%d')

# 오늘이 월요일이면 시작일을 3일전으로
try:
    manage = xw.Book(
        'C:/Users/NAH01080005/Documents/NATEON BIZ 받은 파일/21년 납품자료 관리대장(ADT캡스)_' + today)
except:
    if date.today().weekday() == 0:
        manage = xw.Book(
            'C:/Users/NAH01080005/Documents/NATEON BIZ 받은 파일/21년 납품자료 관리대장(ADT캡스)_' + friday)
    else:
        manage = xw.Book(
            'C:/Users/NAH01080005/Documents/NATEON BIZ 받은 파일/21년 납품자료 관리대장(ADT캡스)_' + yesterday)

manage_sheet = manage.sheets[0].range('A2').options(
    pd.DataFrame, expand='table').value
columns_list = manage.sheets[0].range('A2').expand('right').value
# 빈 데이터프레임 만들기
order = pd.DataFrame(columns=columns_list)
manager = manage.sheets[3].range('A1').options(
    pd.DataFrame, expand='table').value

check_req_and_delete()

# 로그인
options = webdriver.ChromeOptions()
options.add_argument('start-maximized')
driver = webdriver.Chrome(
    'C:/Users/NAH01080005/Downloads/chromedriver.exe', options=options)
driver.get('link')
driver.implicitly_wait(10)
driver.find_element_by_id('inp_id').send_keys('')
driver.find_element_by_id('inp_pw').send_keys('')
driver.find_element_by_id('btn-login').click()
driver.implicitly_wait(5)

frame = driver.find_element_by_xpath(
    '/html/body/div[4]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/div/div[3]/iframe')
driver.switch_to.frame(frame)
main_frame = driver.find_element_by_xpath('//*[@id="main"]')
driver.switch_to.frame(main_frame)

start_date = driver.find_element_by_xpath('//*[@id="iptFR_DATE"]')
start_date.clear()
# 오늘이 월요일이면 시작일을 3일전으로
if date.today().weekday() == 0:
    friday = (date.today() - datetime.timedelta(days=3)).strftime('%Y-%m-%d')
    start_date.send_keys(friday)
# 시작일을 하루 전으로
else:
    yesterday = (date.today() - datetime.timedelta(days=1)
                 ).strftime('%Y-%m-%d')
    start_date.send_keys(yesterday)

# 에이디티캡스 필터링
driver.find_element_by_xpath(
    '//*[@id="ContentHeader"]/table/tbody[1]/tr[3]/td[2]/table/tbody/tr/td/a[1]').click()
driver.switch_to.window(driver.window_handles[1])
driver.find_element_by_xpath('//*[@id="Content_entry"]/table/tbody/tr/td/table/tbody/tr[1]/td/input').send_keys(
    '에이디티캡스')
driver.find_element_by_xpath(
    '//*[@id="Content_entry"]/table/tbody/tr/td/table/tbody/tr[1]/td/a').click()
iframe = driver.find_element_by_xpath(
    '//*[@id="Content_entry"]/table/tbody/tr/td/table/tbody/tr[3]/td/iframe')
driver.switch_to.frame(iframe)
driver.find_element_by_xpath(
    '/html/body/form/table/tbody/tr/td/table/tbody/tr[1]/td[1]/a').click()

driver.switch_to.window(driver.window_handles[0])
driver.switch_to.frame(frame)
driver.switch_to.frame(main_frame)

# 부서보기
view_department = driver.find_element_by_xpath(
    '//*[@id="tbody_control"]/tr[1]/td[2]/input[1]')
view_department.click()

while True:
    driver.switch_to.default_content()
    driver.switch_to.frame(frame)
    driver.switch_to.frame(main_frame)
    driver.find_element_by_xpath('//*[@id="gBtn1"]/a[1]/span').click()
    time.sleep(1)
    bottom_frame = driver.find_element_by_xpath('//*[@id="req_mng_mainFrm"]')
    driver.switch_to.frame(bottom_frame)
    elements = driver.find_elements_by_css_selector('#iframe_tbl > tbody > tr')
    ea = len(elements)
    # 발주 있는지 체크
    if ea == 1:
        print('발주가 없습니다')
        time.sleep(30)

    else:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)
        driver.switch_to.frame(main_frame)
        driver.find_element_by_xpath('//*[@id="gBtn1"]/a[2]/span').click()
        time.sleep(3)
        req = xw.Book(download).sheets(1).range(
            'A1').options(pd.DataFrame, expand='table').value
        order[['고객사명', '고객', '주문일자', '상품코드', '수량']
              ] = req[['사업장명', '수령자', '의뢰일자', '자재코드', '수량']]
        count = 0
        codes = []
        stats = []
        driver.switch_to.frame(bottom_frame)
        for element in elements:
            if count % 2 == 0:
                filtering = re.findall(
                    '\d+-\d+', element.find_element_by_xpath('td[10]/span/font/a').text)
                order_number = ''.join(filtering)
                codes.append(order_number)
                stat = element.find_element_by_xpath(
                    'td[11]/a[2]/font/span[1]').text.strip('()')
                stats.append(stat)

            count += 1

        # 주문번호에 해당하는 고객의 이전 주소들을 불러와서 이전 주소를 입력하거나 새로운 주소 입력
        req['처리번호'] = order['주문번호'] = codes
        order['비고'] = stats
        last = ''
        index = 0
        recep_cnt = 0
        for code in codes:
            target = order[order['주문번호'] == code]
            stat = target['비고'].values[0]
            if stat is not None:
                if stat.find('반품') == 0:
                    order.iloc[index, 18] = '반품처리'
                    order.iloc[index, 8] *= -1

                elif stat.find('교환') == 0:
                    order.iloc[index, 8] = 0

            front_num = re.findall('\d.', code)
            front_num = ''.join(front_num)
            if last != front_num:
                last = front_num
                customer = target['고객'].values[0]
                revenue = req[req['수령자'] == customer]['매출가'] * \
                    req[req['수령자'] == customer]['수량']
                try:
                    registerd_addr = manager[manager['고객명']
                                             == customer]['주소'].values[0]
                except IndexError:
                    print('등록되지 않은 고객입니다')
                    registerd_addr = '등록되지 않음'

                memo = req[req['처리번호'] == code]['전체주문사유'].values[0]
                print('1.', customer, '주소:', registerd_addr)
                addrs = list(
                    set(manage_sheet[manage_sheet['고객'] == customer]['주소'].tolist()))
                try:
                    print('2. 메모 : ' + memo)
                except TypeError:
                    print('메모 없음')
                i = 3
                for addr in addrs:
                    print('%d. %s' % (i, addr))
                    i += 1
                print('%d. 직접입력' % i)
                val = int(input())
                if val == 1:
                    addr = registerd_addr
                elif val == 2:
                    addr = memo
                elif val <= len(addrs) + 2:
                    addr = addrs[val - 3]
                else:
                    addr = input('입력: ')

                order.iloc[index, 17] = addr
                mem = addr

                try:
                    is_memo = memo.find('경동')
                except AttributeError:
                    is_memo = -1

                if addr.find('경동택배') >= 0 and is_memo >= 0:
                    if revenue.sum() < 100000:
                        print('화물발주를 위한 금액이 부족합니다')

                else:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(frame)
                    driver.switch_to.frame(main_frame)
                    if recep_cnt > 0:
                        driver.find_element_by_xpath(
                            '//*[@id="tbody_control"]/tr[1]/td[2]/input[1]').click()
                    Select(driver.find_element_by_xpath(
                        '//*[@id="ContentHeader"]/table/tbody[1]/tr[2]/td[5]/select')).select_by_visible_text('발주No')
                    blank = driver.find_element_by_xpath(
                        '//*[@id="ContentHeader"]/table/tbody[1]/tr[2]/td[6]/input')
                    blank.clear()
                    blank.send_keys(front_num)
                    driver.find_element_by_xpath(
                        '//*[@id="gBtn1"]/a[1]/span').click()
                    time.sleep(1)
                    driver.find_element_by_xpath(
                        '/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[1]/td[1]/a').click()
                    if stat:
                        driver.find_element_by_xpath(
                            '/html/body/form/div[3]/a[7]').click()
                        driver.switch_to.window(driver.window_handles[1])
                        if stat.find('반품') == 0:
                            driver.find_element_by_xpath(
                                '//*[@id="input_002"]')

                        elif stat.find('교환') == 0:
                            driver.find_element_by_xpath(
                                '//*[@id="input_003"]')

                        window = driver.switch_to.alert
                        window.accept()
                        window = driver.switch_to.alert
                        window.accept()
                        driver.switch_to.window(driver.window_handles[0])
                        driver.switch_to.frame(frame)
                        driver.switch_to.frame(main_frame)

                    # 자동발주(화물 아닐경우)
                    elif addr.find('경동택배') == -1 and is_memo == -1:
                        driver.find_element_by_xpath(
                            '/html/body/form/div[3]/a[3]').click()
                        driver.implicitly_wait(3)
                        time.sleep(2)
                        driver.switch_to.default_content()
                        driver.switch_to.frame(driver.find_element_by_xpath(
                            '/html/body/div[4]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/div/div[3]/iframe'))
                        driver.switch_to.frame(
                            driver.find_element_by_xpath('//*[@id="main"]'))
                        driver.find_element_by_xpath(
                            '//*[@id="Content_entry"]/div[2]/form/table/tbody/tr[4]/td/table/tbody/tr/td[1]/input').click()
                        time.sleep(1)
                        window = driver.switch_to.alert
                        print(front_num, window.text)
                        window.accept()
                        recep_cnt += 1

            else:
                order.iloc[index, 17] = mem
            index += 1

        # 서식 채워넣기
        manage.sheets[0].range(get_cell(manage, 'A', 1)).options(
            index=False, header=False).value = order
        columns = 'BEGHJKLMNOQTU'
        apply_formula(manage, columns, len(order))
        break

check_req_and_delete()

manage.save(
    'C:/Users/NAH01080005/Documents/NATEON BIZ 받은 파일/21년 납품자료 관리대장(ADT캡스)_' + today)
driver.quit()
