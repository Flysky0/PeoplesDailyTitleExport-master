from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager


def CreateEdgeDriverService():
    try:
        service = EdgeService(
            EdgeChromiumDriverManager().install(), verbose=True)
    except Exception as e:
        service = EdgeService(verbose=True)
        print(f'EdgeDriver更新失败,错误信息:{e.__class__.__name__}')
    # 以下代码片段会创建新的 EdgeDriverService 并启用详细日志输出：
    # service = EdgeService(verbose=True)
    # driver = webdriver.Edge(service=service)

    edge_options = webdriver.EdgeOptions()
    edge_options.add_experimental_option(
        "excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option('useAutomationExtension', False)
    edge_options.add_argument('lang=zh-CN,zh,zh-TW,en-US,en')
    edge_options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.53 Safari/537.36 Edg/103.0.1264.37')
    # 'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36')
    # 就是这一行告诉edge去掉了webdriver痕迹
    edge_options.add_argument("disable-blink-features=AutomationControlled")

    driver = webdriver.Edge(service=service, options=edge_options)

    return driver
