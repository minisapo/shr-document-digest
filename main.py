"""
SmartHR文書配付ダイジェスト自動化スクリプト
ログイン → 文書配付一覧のスクレイピング → 最新文書のダイジェスト生成 → 出力 → ログアウト
"""
import csv
import ctypes
import io
import os
import sys
from datetime import date
from pathlib import Path
from urllib.parse import urljoin

import anthropic
import pdfplumber
from dotenv import load_dotenv
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

load_dotenv()

SMARTHR_URL = os.environ["SMARTHR_URL"]
SMARTHR_HOME_URL = SMARTHR_URL.replace("/login", "/home")
SMARTHR_LOGIN_ID = os.environ["SMARTHR_LOGIN_ID"]
SMARTHR_PASSWORD = os.environ["SMARTHR_PASSWORD"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]

DIGEST_MODEL = "claude-sonnet-4-6"
CSV_COLUMNS = ["依頼ステータス", "書類名", "書類ステータス", "受信日", "URL"]


def get_desktop_path() -> Path:
    """OneDriveでリダイレクトされている場合も考慮し、実際のデスクトップフォルダのパスを取得する"""
    CSIDL_DESKTOPDIRECTORY = 0x0010
    SHGFP_TYPE_CURRENT = 0
    buf = ctypes.create_unicode_buffer(1024)
    ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_DESKTOPDIRECTORY, 0, SHGFP_TYPE_CURRENT, buf)
    return Path(buf.value)


def login(page: Page) -> None:
    """SmartHRのログイン画面から、社員番号（またはメールアドレス）とパスワードでログインする"""
    page.goto(SMARTHR_URL)
    page.locator('input[name="id"]').fill(SMARTHR_LOGIN_ID)
    page.locator('input[name="password"]').fill(SMARTHR_PASSWORD)
    page.locator('button[type="submit"]').click()

    try:
        page.wait_for_url("**/home", timeout=15000)
    except PlaywrightTimeoutError:
        error_alert = page.get_by_role("alert")
        if error_alert.count() > 0:
            raise RuntimeError(f"SmartHRログインに失敗しました: {error_alert.first.inner_text()}")
        raise RuntimeError("SmartHRログインに失敗しました（想定外の画面遷移です）")

    print("SmartHRへのログインに成功しました。")


def go_to_document_distribution(page: Page) -> None:
    """左メニューの「文書配付」から受信ボックス一覧へ遷移する"""
    page.get_by_role("link", name="文書配付").click()
    page.wait_for_url("**/inbox", timeout=15000)
    print("文書配付の受信ボックスを開きました。")


def scrape_inbox(page: Page) -> list[dict]:
    """受信ボックス一覧をスクレイピングし、文書ごとの行データ（一番上が最新）を返す"""
    base_url = page.url
    rows = page.locator("table tbody tr")
    records = []
    for i in range(rows.count()):
        row = rows.nth(i)
        cells = row.locator("td")
        request_status = cells.nth(0).inner_text().strip()
        received_date = cells.nth(3).inner_text().strip()
        doc_links = cells.nth(1).locator("a")
        doc_statuses = cells.nth(2).locator("li")
        for j in range(doc_links.count()):
            link = doc_links.nth(j)
            status_text = doc_statuses.nth(j).inner_text().strip() if j < doc_statuses.count() else ""
            records.append({
                "依頼ステータス": request_status,
                "書類名": link.inner_text().strip(),
                "書類ステータス": status_text,
                "受信日": received_date,
                "URL": urljoin(base_url, link.get_attribute("href")),
            })
    print(f"受信ボックスから{len(records)}件の文書を取得しました。")
    return records


def save_inbox_csv(records: list[dict], output_path: Path) -> None:
    """受信ボックス一覧をCSVに保存する（Excelでの文字化け防止のためutf-8-sigで出力）"""
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(records)
    print(f"文書配付一覧を保存しました: {output_path}")


def extract_latest_document_text(page: Page, document_url: str) -> str:
    """文書ページへ遷移し、PDFをダウンロードしてテキストを抽出する"""
    page.goto(document_url)
    download_href = page.get_by_role("link", name="ダウンロード").get_attribute("href")
    pdf_url = urljoin(page.url, download_href)

    response = page.request.get(pdf_url)
    if not response.ok:
        raise RuntimeError(f"文書PDFのダウンロードに失敗しました（status={response.status}）")

    text_parts = []
    with pdfplumber.open(io.BytesIO(response.body())) as pdf:
        for pdf_page in pdf.pages:
            text_parts.append(pdf_page.extract_text() or "")
    return "\n".join(text_parts)


def generate_digest(document_text: str) -> str:
    """Anthropic APIで文書内容のダイジェスト（要約）を生成する"""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model=DIGEST_MODEL,
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": (
                "以下はSmartHRの文書配付で届いた書類の内容です。"
                "要点を箇条書きで日本語で分かりやすく要約してください。\n\n" + document_text
            ),
        }],
    )
    return message.content[0].text


def logout(page: Page) -> None:
    """SmartHRからログアウトする"""
    page.goto(SMARTHR_HOME_URL)
    page.get_by_role("button", name="アカウント:", exact=False).click()
    page.get_by_role("button", name="ログアウト").click()
    page.wait_for_url("**/login", timeout=15000)
    print("SmartHRからログアウトしました。")


def main() -> None:
    desktop = get_desktop_path()
    today = date.today().strftime("%Y%m%d")
    csv_path = desktop / f"文書配付一覧_{today}.csv"
    digest_path = desktop / f"文書ダイジェスト_{today}.txt"

    with sync_playwright() as playwright:
        # Chrome本体を起動し、ウィンドウ最大化で表示する
        browser = playwright.chromium.launch(
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        try:
            login(page)
            go_to_document_distribution(page)
            records = scrape_inbox(page)
            save_inbox_csv(records, csv_path)

            if not records:
                print("受信ボックスに文書がないため、ダイジェスト生成をスキップします。")
            else:
                latest = records[0]
                document_text = extract_latest_document_text(page, latest["URL"])
                digest = generate_digest(document_text)
                digest_path.write_text(
                    f"■ 対象文書: {latest['書類名']}（受信日: {latest['受信日']}）\n\n{digest}\n",
                    encoding="utf-8",
                )
                print(f"文書ダイジェストを保存しました: {digest_path}")

            logout(page)
        except RuntimeError as error:
            print(f"エラー: {error}")
            browser.close()
            sys.exit(1)

        browser.close()


if __name__ == "__main__":
    main()
