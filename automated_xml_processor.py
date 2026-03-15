
import subprocess
import os
import requests
import zipfile
import io
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import time
import xml.etree.ElementTree as ET
import pandas as pd

def get_edge_version():
    """获取本地Microsoft Edge浏览器的版本号。"""
    try:
        command = r'(Get-Item "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").VersionInfo.ProductVersion'
        result = subprocess.run(['powershell', '-Command', command], capture_output=True, text=True, check=True)
        return result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"获取Edge版本失败: {e}")
        return None

def setup_edgedriver():
    """自动检测、下载并设置msedgedriver，返回驱动路径。"""
    driver_path = os.path.join(os.getcwd(), 'msedgedriver.exe')
    local_version = get_edge_version()
    if not local_version:
        return None # 无法获取本地版本，无法继续

    # 检查本地驱动版本是否匹配
    if os.path.exists(driver_path):
        try:
            # 尝试执行本地驱动获取版本信息
            result = subprocess.run([driver_path, '--version'], capture_output=True, text=True, check=True)
            # msedgedriver --version 输出格式: Microsoft Edge WebDriver 123.0.2420.5 (ef1b2267b5...)
            local_driver_version = result.stdout.split(' ')[3]
            if local_driver_version == local_version:
                print(f"本地驱动版本 ({local_driver_version}) 与浏览器版本 ({local_version}) 匹配，无需下载。")
                return driver_path
            else:
                print(f"本地驱动版本 ({local_driver_version}) 与浏览器版本 ({local_version}) 不匹配，准备更新...")
        except Exception as e:
            print(f"检查本地驱动版本失败: {e}，准备重新下载...")

    print("开始查找并下载匹配的WebDriver...")
    try:
        # 使用已确认可用的URL格式
        driver_zip_url = f"https://msedgedriver.microsoft.com/{local_version}/edgedriver_win64.zip"
        print(f"尝试从以下链接下载: {driver_zip_url}")

        # 下载并解压
        zip_response = requests.get(driver_zip_url, timeout=60)
        zip_response.raise_for_status()
        
        with zipfile.ZipFile(io.BytesIO(zip_response.content)) as z:
            print("正在解压...")
            with z.open('msedgedriver.exe') as source, open(driver_path, 'wb') as target:
                target.write(source.read())
        
        print(f"msedgedriver.exe 已成功下载并放置在: {driver_path}")
        return driver_path

    except requests.exceptions.RequestException as e:
        print(f"下载或访问时发生网络错误: {e}")
    except Exception as e:
        print(f"处理WebDriver时发生未知错误: {e}")
    
    return None

def get_usd_rate_by_date(driver_path, input_date, headless=False):
    """使用指定的msedgedriver来获取汇率。"""
    service = Service(executable_path=driver_path)
    edge_options = Options()
    if headless:
        edge_options.add_argument('--headless')
        
    driver = webdriver.Edge(service=service, options=edge_options)
    
    try:
        date_obj = datetime.strptime(input_date, '%Y-%m-%d')
        first_day = date_obj.replace(day=1).strftime('%Y-%m-%d')
        last_day = (date_obj.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        last_day = last_day.strftime('%Y-%m-%d')

        driver.get('https://www.safe.gov.cn/safe/rmbhlzjj/index.html')
        iframes = driver.find_elements('tag name', 'iframe')
        if len(iframes) > 0:
            driver.switch_to.frame(iframes[0])
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#startDateId')))
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#endDateId')))
        driver.find_element(By.CSS_SELECTOR, '#startDateId').clear()
        driver.find_element(By.CSS_SELECTOR, '#startDateId').send_keys(first_day)
        driver.find_element(By.CSS_SELECTOR, '#endDateId').clear()
        driver.find_element(By.CSS_SELECTOR, '#endDateId').send_keys(last_day)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#RMBQuery > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr > td:nth-child(2) > span.form-table-label > input[type=button]'))).click()
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#InfoTable')))
        time.sleep(2)
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', id='InfoTable')
        if table:
            rows = table.find_all('tr')
            if rows:
                last_row = rows[-1]
                tds = last_row.find_all('td')
                if len(tds) >= 2:
                    return tds[1].get_text(strip=True)
    finally:
        driver.quit()
    return None

def unit_map(fd_unit):
    if fd_unit == '035':
        return '千克'
    elif fd_unit == '011':
        return '件'
    else:
        return fd_unit

# --- 主程序逻辑 ---
def main():
    print("--- 开始执行自动化脚本 ---")
    # 1. 自动设置/更新驱动
    driver_executable_path = setup_edgedriver()
    if not driver_executable_path:
        print("\n错误: WebDriver 设置失败，无法继续执行。")
        return

    # 2. 执行核心业务逻辑
    try:
        tree = ET.parse('8910001049082-20260209-20260215--20260228214022_解密后.xml')
        root = tree.getroot()

        data = []
        lj_date_raw = root.find('Dec').find('DecHead').find('lj_date').text
        lj_date = f'{lj_date_raw[:4]}-{lj_date_raw[4:6]}-{lj_date_raw[6:]}'
        
        print("\n正在获取汇率...")
        rate = get_usd_rate_by_date(driver_executable_path, lj_date, headless=True)
        if not rate:
            print("未能获取到汇率，后续本币金额将为空。")
        else:
            print(f"成功获取到汇率: {rate}")

        for dec in root.findall('Dec'):
            bgd_no = dec.find('DecHead').find('bgd_no').text
            ht_no = dec.find('DecHead').find('ht_no').text
            lj_date_raw = dec.find('DecHead').find('lj_date').text
            lj_date = f'{lj_date_raw[:4]}-{lj_date_raw[4:6]}-{lj_date_raw[6:]}'
            for dec_list in dec.find('DecLists').findall('DecList'):
                cm_name = dec_list.find('cm_name').text
                fd_qnt = dec_list.find('Fd_qnt').text
                fd_unit = unit_map(dec_list.find('Fd_unit').text)
                yb_amt = dec_list.find('yb_amt').text
                try:
                    rmb_amt = round(float(rate)/100 * float(yb_amt), 2) if rate else None
                except (ValueError, TypeError):
                    rmb_amt = None
                data.append({
                    '报关单号': bgd_no,
                    '合同号': ht_no,
                    '出口日期': lj_date,
                    '项目': cm_name,
                    '数量': fd_qnt,
                    '单位': fd_unit,
                    '原币金额': yb_amt,
                    '记账汇率': rate,
                    '本币金额': rmb_amt
                })

        df = pd.DataFrame(data)
        df.to_excel('报关单清单提取结果.xlsx', index=False)
        print('\n已成功导出到 报关单清单提取结果.xlsx')

    except FileNotFoundError as e:
        print(f"错误: {e}")
    except Exception as e:
        print(f"处理过程中发生未知错误: {e}")
    finally:
        print("\n--- 脚本执行结束 ---")

if __name__ == '__main__':
    main()
