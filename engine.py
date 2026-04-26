import re
import json
import os
import requests
import anthropic
from datetime import datetime
from playwright.async_api import async_playwright

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ============================================================
# SalesOS GAS設定
# デプロイしたGASのURLを入れてください
# ============================================================
GAS_URL = os.environ.get("GAS_URL", "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec")

def generate_tracking_url(company_name, document_url):
    """GASを呼び出して追跡URLを生成する"""
    try:
        response = requests.get(GAS_URL, params={
            "action": "generate",
            "company": company_name,
            "docUrl": document_url,
        }, timeout=10)
        data = response.json()
        return data.get("trackingUrl", document_url)
    except Exception as e:
        print(f"[追跡URL生成失敗] {company_name}: {e}")
        return document_url  # 失敗した場合は元のURLをそのまま使う

FIND_FORM_PROMPT = """
あなたはWebサイトのHTML解析の専門家です。
以下のHTMLから問い合わせフォームへのリンクを探してください。

探すべきリンクの例：
- 「お問い合わせ」「Contact」「問合せ」「ご相談」「資料請求」
- 「協力会社の方へ」「お取引先の方」などの専用ページ

レスポンスは必ずJSON形式で返してください：
{
  "form_url": "フォームページのURL（見つからない場合はnull）",
  "is_partner_form": true/false（協力会社専用フォームかどうか）,
  "blocked": true/false（営業禁止・自動送信禁止の記載があるか）,
  "blocked_reason": "ブロック理由（blockedがtrueの場合）"
}

HTML:
"""

FILL_FORM_PROMPT = """
あなたはWebフォームの入力専門家です。
以下のフォームHTMLと送信者情報をもとに、各フィールドへの入力値を教えてください。

送信者情報:
{sender_json}

フォームHTML:
{form_html}

各inputのname属性またはid属性をキーに、入力値をバリューとしたJSONで返してください。
selectやtextareaも含めてください。
送信ボタンは含めないでください。
必ずJSONのみ返し、説明文は不要です。
例: {{"name": "山田太郎", "email": "test@example.com", "message": "お問い合わせ内容"}}
"""

import os

class FormSender:
    def __init__(self, test_mode=True):
        self.test_mode = test_mode
        self.playwright = None
        self.browser = None
        self.client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

    async def init(self):
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=True)

    async def close(self):
        await self.browser.close()
        await self.playwright.stop()

    async def process(self, company_name, top_url, sender, document_url=""):
        result = {
            'company': company_name,
            'top_url': top_url,
            'status': '',
            'reason': '',
            'form_url': '',
            'timestamp': datetime.now().strftime('%Y/%m/%d %H:%M'),
        }
        try:
            # 追跡URLを生成して本文に埋め込む
            if document_url and "{TRACKING_URL}" in sender.get("message", ""):
                tracking_url = generate_tracking_url(company_name, document_url)
                sender = dict(sender)  # コピーして元データを変更しない
                sender["message"] = sender["message"].replace("{TRACKING_URL}", tracking_url)
                result["tracking_url"] = tracking_url
                print(f"[追跡URL生成] {company_name}: {tracking_url}")

            page = await self.browser.new_page()
            await page.goto(top_url, timeout=15000)
            html = await page.content()

            # Step1: フォームリンクを探す
            analysis = self._call_claude(FIND_FORM_PROMPT + html[:8000])
            data = json.loads(analysis)

            if data.get('blocked'):
                result['status'] = '⚠️ スキップ'
                result['reason'] = f"送信禁止: {data.get('blocked_reason', '営業禁止の記載あり')}"
                await page.close()
                return result

            form_url = data.get('form_url')
            if not form_url:
                result['status'] = '❌ 失敗'
                result['reason'] = 'フォームリンクが見つかりませんでした'
                await page.close()
                return result

            # 相対URLを絶対URLに変換
            if form_url.startswith('/'):
                from urllib.parse import urlparse
                parsed = urlparse(top_url)
                form_url = f"{parsed.scheme}://{parsed.netloc}{form_url}"

            result['form_url'] = form_url
            await page.goto(form_url, timeout=15000)

            # Step2: CAPTCHAチェック
            content = await page.content()
            if 'recaptcha' in content.lower() or 'hcaptcha' in content.lower():
                result['status'] = '❌ 失敗'
                result['reason'] = 'CAPTCHA検知'
                await page.close()
                return result

            # Step3: PDFチェック
            if form_url.endswith('.pdf'):
                result['status'] = '❌ 失敗'
                result['reason'] = 'PDFフォーム（対応不可）'
                await page.close()
                return result

            # Step4: フォーム入力値をAIで判定
            form_html = await page.inner_html('body')
            prompt = FILL_FORM_PROMPT.format(
                sender_json=json.dumps(sender, ensure_ascii=False),
                form_html=form_html[:6000]
            )
            fill_data_str = self._call_claude(prompt)
            fill_data = json.loads(fill_data_str)

            # Step5: フォームに入力
            for key, value in fill_data.items():
                try:
                    locator = page.locator(f'[name="{key}"], [id="{key}"]').first
                    tag = await locator.evaluate('el => el.tagName.toLowerCase()')
                    if tag == 'select':
                        await locator.select_option(label=str(value))
                    else:
                        await locator.fill(str(value))
                except:
                    pass

            # Step6: テストモードは送信しない
            if self.test_mode:
                result['status'] = '✅ 送信完了（テストモード・未送信）'
                result['reason'] = 'テストモードのため実際の送信はしていません'
            else:
                # 送信ボタンをクリック
                submit = page.locator('button[type="submit"], input[type="submit"]').first
                await submit.click()
                await page.wait_for_timeout(2000)
                result['status'] = '✅ 送信完了'

            await page.close()

        except Exception as e:
            result['status'] = '❌ 失敗'
            result['reason'] = f'エラー: {str(e)[:100]}'

        return result

    def _call_claude(self, prompt):
        message = self.client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
