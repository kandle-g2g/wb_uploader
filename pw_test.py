from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        "/tmp/pw_test_profile", channel="chrome", headless=False, timeout=20000)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto("https://example.com")
    print("OK:", page.title())
    ctx.close()