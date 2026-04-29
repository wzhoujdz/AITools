import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

NAME = "你的用户名"
PASSWORD = "你的密码"
LOGIN_URL = "https://系统网址.com"
PRODUCT_NAME = "MOM"                         # 要搜索的产品简称
DOWNLOAD_DIR = os.path.expanduser("~/Downloads")  # 浏览器下载路径
LOAD_TIME = 120                               # 等待PDF加载的时间（秒），视网速调整

def ensure_download_dir(path):
    """如果下载目录不存在则自动创建"""
    if not os.path.exists(path):
        os.makedirs(path)

def wait_for_loading_to_finish(driver, timeout=30):
    """
    等待页面加载动画（el-loading）消失。
    如果超时则强制用JavaScript移除遮罩，避免卡住。
    """
    try:
        WebDriverWait(driver, timeout).until(
            EC.any_of(
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-spinner")),
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-mask"))
            )
        )
        WebDriverWait(driver, timeout).until_not(
            EC.any_of(
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-spinner")),
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-mask"))
            )
        )
    except TimeoutException:
        # 超时后强制移除遮罩层，防止界面被锁死
        try:
            driver.execute_script("""
                const masks = document.querySelectorAll('.el-loading-mask, .el-loading-spinner');
                masks.forEach(mask => mask.remove());
            """)
        except:
            pass
    except:
        pass

def search_product(driver, product_name):
    """在搜索框中输入产品名称并点击查询按钮"""
    product_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="请输入产品简称"]'))
    )
    product_input.clear()
    product_input.send_keys(product_name)

    query_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="front_template"]/div[1]/form/div[5]/div[2]/div/button[2]'))
    )
    query_btn.click()
    print("点击查询")
    wait_for_loading_to_finish(driver)

def main():
    # 准备下载目录
    ensure_download_dir(DOWNLOAD_DIR)

    # 配置 Chrome 选项，设置下载路径和禁止弹窗
    options = Options()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-site-isolation-trials")
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)

    try:
        # 1. 打开登录页面
        driver.get(LOGIN_URL)
        driver.maximize_window()
        time.sleep(1)

        # 2. 点击登录按钮进入表单
        login_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div'))
        )
        login_btn.click()

        # 3. 等待用户名密码输入框出现
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@placeholder="请输入用户名"]'))
        )
        username_input = driver.find_element(By.XPATH, '//input[@placeholder="请输入用户名"]')
        password_input = driver.find_element(By.XPATH, '//input[@placeholder="请输入您的登录密码"]')
        username_input.send_keys(NAME)
        password_input.send_keys(PASSWORD)

        # 4. 点击立即登录
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(span, "立即登录")]'))
        )
        login_button.click()

        # 5. 点击“统计报表”菜单
        report_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '#app > div > div.menu > div.menu_content > ul > li:nth-child(5)'))
        )
        report_menu.click()

        # 等待报表菜单可能出现的加载遮罩消失
        try:
            loading_mask = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-mask"))
            )
            WebDriverWait(driver, 20).until(EC.staleness_of(loading_mask))
        except:
            pass

        # 6. 进入“业绩报告”页面（使用 JS 点击避免被其他元素遮挡）
        report_item = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//li[text()="业绩报告"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", report_item)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", report_item)

        # 等待页面表格加载完成
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="front_template"]/div[2]/div[1]/div[2]/table/thead/tr/th[2]/div'))
        )
        time.sleep(5)  # 给页面足够时间稳定

        # 7. 筛选产品类型：点击“平层基金”和“子基金”
        type1_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[text()="平层基金" and @class="el-checkbox-button__inner"]'))
        )
        type1_btn.click()
        type2_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//label[.//span[text()="子基金"]]'))
        )
        type2_btn.click()

        # 搜索指定产品
        search_product(driver, PRODUCT_NAME)

        # 记录列表页URL，用于下载后返回
        list_page_url = driver.current_url
        current_page = 1
        download_count = 0

        # 循环处理每一页
        while True:
            table_xpath = '//*[@id="front_template"]/div[2]/div[1]/div[3]/div/div[1]/div/table'
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, table_xpath)))

            # 获取当前页总行数（通过最后一行 tr 的 rowIndex）
            total_rows = 0
            try:
                last_row = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, f'{table_xpath}/tbody/tr[last()]'))
                )
                total_rows = int(last_row.get_attribute('rowIndex')) + 1
            except Exception as e:
                print(f"获取总行数失败: {e}")

            # 找出所有包含产品名称的行（第二列包含关键词 MOM 的）
            mom_rows = []
            for i in range(1, total_rows + 1):
                cell_xpath = f'{table_xpath}/tbody/tr[{i}]/td[2]'
                try:
                    cell = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, cell_xpath))
                    )
                    if PRODUCT_NAME.upper() in cell.text.upper():
                        mom_rows.append(i)
                except Exception as e:
                    print(f"检查第{i}行时出错: {e}")

            # 逐一下载每个报告的PDF
            for i in mom_rows:
                try:
                    report_btn_xpath = f'{table_xpath}/tbody/tr[{i}]/td[14]/div/button'
                    wait_for_loading_to_finish(driver)
                    report_btn = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, report_btn_xpath))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", report_btn)
                    time.sleep(1)
                    report_btn.click()
                    print(f"已点击第{i}个报告按钮")

                    # 等待“导出PDF”按钮出现并点击
                    export_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '//button[contains(@class, "el-button--danger") and contains(., "导出PDF")]'))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_btn)
                    time.sleep(LOAD_TIME)  # 等待PDF预览加载完成
                    export_btn.click()
                    print(f"已触发第{i}个PDF下载")
                    download_count += 1

                    # 返回列表页（尝试多次，防止弹窗未及时关闭）
                    for attempt in range(3):
                        try:
                            back_btn = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable(
                                    (By.XPATH, '//*[@id="front_template"]/div[4]/div/div/div/header/button/i'))
                            )
                            time.sleep(5)  # 等待页面稳定
                            back_btn.click()
                            print(f"成功返回列表（第{i}个报告）")
                            break
                        except Exception as e:
                            print(f"返回按钮点击失败 (第{attempt+1}次尝试): {e}")
                            # 若失败则重新导航到列表页
                            driver.get(list_page_url)
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, '//table/tbody/tr[1]/td[14]/div/button'))
                            )
                            break
                    else:
                        print("多次尝试返回均失败，终止下载")
                        break

                except Exception as e:
                    print(f"处理第{i}个报告时出错: {e}")
                    if "timed out" in str(e).lower() or "no such element" in str(e).lower():
                        print(f"当前页可能没有第{i}个报告，跳过")
                        break
                    continue

            # 翻页：尝试点击下一页
            wait_for_loading_to_finish(driver)
            next_page_num = current_page + 1
            next_page_btn_xpath = f'//*[@id="front_template"]/div[3]/ul/li[{next_page_num}]'
            try:
                next_page_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, next_page_btn_xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", next_page_btn)
                wait_for_loading_to_finish(driver)
                next_page_btn.click()
                wait_for_loading_to_finish(driver)
                # 等待新页面表格加载
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="front_template"]/div[2]/div[1]/div[2]/table/thead/tr/th[2]/div'))
                )
                current_page = next_page_num
                print(f"已翻至第{current_page}页")
            except Exception as e:
                print(f"无法翻到下一页，已到最后一页或出错: {e}")
                break

        print(f"任务完成，共下载 {download_count} 个PDF文件，保存在 {DOWNLOAD_DIR}")

    except Exception as e:
        print(f"操作失败：{e}")
        # 出错时截图并保存页面源码，方便调试
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        driver.save_screenshot(f"error_{timestamp}.png")
        with open(f"page_source_{timestamp}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("已保存错误截图和页面源码")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()