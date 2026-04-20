import os
import json
from datetime import datetime
from playwright.async_api import async_playwright
import anthropic

FIND_FORM_PROMPT = """
あなたはWebサイトのHTML解析の専門家です。
以下のHTMLから問い合わせフォームへのリンクを探してください。

探すべきリンクの例：
- 「お問い合わせ」「Contact」「問合せ」「ご相談」「資料請求」
- 「協力会社の方へ」「お取引先の方」などの専用ページ

レスポンスは必ずJSON形式のみで返してください：
{
  "form_url": "フォームページのURL（見つからない場合はnull）",
  "is_partner_form": true,
  "blocked": false,
  "blocked_reason": ""
}

HTML:
"""

FILL_FORM_PROMPT = """
あなたはWebフォームの入力専門家です。
以下のフォームHTMLと送信者情報をもとに、各フィールドへの入力値と
チェックボックスの選択を決定してください。

送信者情報:
{sender_json}

問い合わせ種別ヒント:
{inquiry_type}

フォームHTML:
{form_html}

以下のJSON形式のみで返してください。説明文は不要です：
{{
  "text_fields": {{
    "フィールドのnameまたはid": "入力値"
  }},
  "checkboxes": [
    {{
      "name": "チェックボックスのnameまたはid",
      "value": "valueの値",
      "should_check": true,
      "reason": "チェックする理由"
    }}
  ],
  "selects": {{
    "selectのnameまたはid": "選択するoption値またはラベル"
  }},
  "missing_fields": [
    {{
      "field_name": "入力できなかったフィールド名",
      "reason": "入力できなかった理由"
    }}
  ]
}}

チェックボックスの判断基準：
- 「個人情報」「同意」「プライバシー」「利用規約」→ 必ずtrue
- 問い合わせ種別は「問い合わせ種別ヒント」に近いものをtrue
- 「営業」「勧誘」「広告」などはfalse

missing_fieldsには、選択肢が不明・特殊すぎて入力できなかった項目を記録してください。
フリガナ・カナ・よみがなフィールドは送信者名カナを使って入力してください。
会社名カナも同様にカタカナで入力してください。
"""


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
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    async def process(self, company_name, top_url, sender):
        result = {
            'company': company_name,
            'top_url': top_url,
            'status': '',
            'reason': '',
            'missing_fields': '-',
            'form_url': '',
            'timestamp': datetime.now().strftime('%Y/%m/%d %H:%M'),
        }
        try:
            page = await self.browser.new_page()
            await page.goto(top_url, timeout=15000)
            html = await page.content()

            # Step1: フォームリンクをAIで探す
            analysis = self._call_claude(FIND_FORM_PROMPT + html[:8000])
            analysis = self._extract_json(analysis)
            data = json.loads(analysis)

            if data.get('blocked'):
                result['status'] = 'スキップ'
                result['reason'] = data.get('blocked_reason', '営業禁止の記載あり')
                await page.close()
                return result

            form_url = data.get('form_url')
            if not form_url:
                result['status'] = '失敗'
                result['reason'] = 'フォームリンクが見つかりませんでした'
                await page.close()
                return result

            if form_url.startswith('/'):
                from urllib.parse import urlparse
                parsed = urlparse(top_url)
                form_url = f"{parsed.scheme}://{parsed.netloc}{form_url}"

            result['form_url'] = form_url
            await page.goto(form_url, timeout=15000)

            content = await page.content()
            if 'recaptcha' in content.lower() or 'hcaptcha' in content.lower():
                result['status'] = '失敗'
                result['reason'] = 'CAPTCHA検知'
                await page.close()
                return result

            if form_url.lower().endswith('.pdf'):
                result['status'] = '失敗'
                result['reason'] = 'PDFフォーム（対応不可）'
                await page.close()
                return result

            # Step4: AIでフォーム入力値を判定
            form_html = await page.inner_html('body')
            prompt = FILL_FORM_PROMPT.format(
                sender_json=json.dumps(sender, ensure_ascii=False),
                inquiry_type=sender.get('inquiry_type', '協力会社として取引のご案内'),
                form_html=form_html[:6000]
            )
            fill_data_str = self._call_claude(prompt)
            fill_data_str = self._extract_json(fill_data_str)
            fill_data = json.loads(fill_data_str)

            # 入力できなかった項目を記録
            missing = fill_data.get('missing_fields', [])
            if missing:
                result['missing_fields'] = '、'.join([f"{m['field_name']}（{m['reason']}）" for m in missing])

            # テキストフィールド入力
            for key, value in fill_data.get('text_fields', {}).items():
                try:
                    locator = page.locator(f'[name="{key}"], [id="{key}"]').first
                    await locator.fill(str(value))
                except Exception:
                    pass

            # チェックボックス操作
            for cb in fill_data.get('checkboxes', []):
                try:
                    name = cb.get('name', '')
                    value = cb.get('value', '')
                    should_check = cb.get('should_check', False)
                    if value:
                        locator = page.locator(f'input[type="checkbox"][name="{name}"][value="{value}"]').first
                    else:
                        locator = page.locator(f'input[type="checkbox"][name="{name}"], input[type="checkbox"][id="{name}"]').first
                    is_checked = await locator.is_checked()
                    if should_check and not is_checked:
                        await locator.check()
                    elif not should_check and is_checked:
                        await locator.uncheck()
                except Exception:
                    pass

            # セレクトボックス操作
            for key, value in fill_data.get('selects', {}).items():
                try:
                    locator = page.locator(f'select[name="{key}"], select[id="{key}"]').first
                    try:
                        await locator.select_option(value=str(value))
                    except Exception:
                        await locator.select_option(label=str(value))
                except Exception:
                    pass

            # 送信
            if self.test_mode:
                if missing:
                    result['status'] = '部分送信（テスト）'
                    result['reason'] = '一部入力できない項目があります'
                else:
                    result['status'] = '送信完了（テスト）'
                    result['reason'] = 'テストモードのため実際の送信はしていません'
            else:
                submit = page.locator('button[type="submit"], input[type="submit"]').first
                await submit.click()
                await page.wait_for_timeout(2000)
                if missing:
                    result['status'] = '部分送信'
                    result['reason'] = '一部入力できない項目があります'
                else:
                    result['status'] = '送信完了'

            await page.close()

        except Exception as e:
            result['status'] = '失敗'
            result['reason'] = f'エラー: {str(e)[:100]}'

        return result

    def _extract_json(self, text):
        text = text.strip()
        if text.startswith("```"):
            parts = text.split("```")
            if len(parts) >= 2:
                text = parts[1]
                if text.startswith("json"):
                    text = text[4:]
        return text.strip()

    def _call_claude(self, prompt):
        message = self.client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
