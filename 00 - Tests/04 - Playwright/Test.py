from playwright.sync_api import sync_playwright
import time


with sync_playwright() as pw:
    browser = pw.chromium.launch(executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe', headless=False)

    page = browser.new_page()

    page.goto('https://bosch.wescale.com/app/#/dashboard')

    time.sleep(5)

    page.locator('//*[@id="i0116"]').fill('mao8ct@bosch.com')
    page.locator('//*[@id="idSIButton9"]').click()

    time.sleep(20)
    browser.close()