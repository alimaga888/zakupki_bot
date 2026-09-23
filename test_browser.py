from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False
    )

    context = browser.new_context(
        ignore_https_errors=True
    )

    page = context.new_page()

    page.goto(
        "https://zakupki.gov.ru/",
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("Страница открыта:")
    print(page.title())

    page.wait_for_timeout(10000)

    browser.close()