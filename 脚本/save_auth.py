# save_auth.py (保存登录态，运行一次即可)
from playwright.sync_api import sync_playwright
import time

# 填入你电脑上谷歌浏览器的实际路径
CHROME_PATH = "C:\Program Files\Google\Chrome\Application\chrome.exe" 
TARGET_URL = "https://shopee.menglar.com/workbench/home"

with sync_playwright() as p:
    # 使用你本地的 Chrome 内核，且必须是 non-headless（有头模式）
    browser = p.chromium.launch(
        executable_path=CHROME_PATH, 
        headless=False,
        args=["--start-maximized"] # 最大化窗口，方便你操作
    )
    context = browser.new_context(viewport={"width": 1920, "height": 1080})
    page = context.new_page()
    
    page.goto(TARGET_URL)
    print(">>> 请在打开的浏览器中手动完成登录和验证码！")
    print(">>> 登录成功，能看到商品数据后，回到这里按回车键...")
    
    input() # 等待你按回车
    
    # 保存登录状态
    context.storage_state(path="auth.json")
    print(">>> 登录态已保存为 auth.json！")
    browser.close()
