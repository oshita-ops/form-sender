import os
import json
from datetime import datetime
import anthropic
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

FIND_FORM_PROMPT = """
あなたはWebサイトのHTML解析の専門家です。
以下のHTMLから問い合わせフォームへのリンクを探してください。

探すべきリンクの例：
- 「お問い合わせ」「Contact」「問合せ」「ご相談」「資料請求」
- 「協力会社の方へ」「お取引先の方」などの専用ページ

また、ページ内に営業禁止・自動送信禁止の記載がないかも確認してください。

レスポンスは必ずJSON形式のみで返してください：
{
  "form_url": "フォームページのURL（見つからない場合はnull）",
  "blocked": false,
  "blocked_reason": ""
}

ベースURL: {base_url}
HTML:
"""

FILL_FORM_PROMPT = """
あなたはWebフォームの入力専門家です。
以下のフォームHTMLと送信者情報をもとに、送信に必要な情報を分析してください。

送信者情報:
{sender_json}

問い合わせ種別ヒント:
{inquiry_type}

フォームHTML:
{form_html}

以下のJSON形式のみで返してください。説明文は不要です：
{{
  "action": "フォームのaction URL（なければnull）",
  "method": "post または get",
  "fields": {{
    "フィールドのname属性": "入力値"
  }},
  "missing_fields": [
    {{
      "field_name": "入力できなかったフィールド名",
      "reason": "入力できなかった理由"
    }}
  ],
  "has_captcha": false,
  "has_js_required": false,
  "notes": "特記事項があれば"
}}

チェックボックスの判断基準：
- 「個人情報」「同意」「プライバシー」「利用規約」→ チェックありの値をfieldsに含める
- 問い合わせ種別は「問い合わせ種別ヒント」に近いものを選択
- 「営業」「勧誘」はfalseまたは選択しない

フリガナ・カナ・よみがな フィールドは送信者名カナを使って入力。
会社名カナも同様にカタカナで入力。
missing_fieldsには選択肢が特殊すぎて判断できなかった項目を記録。
"""


class FormSender:
    def __init__(self, test_mode=True):
        self.test_mode = test_mode
        self.client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })

    async def init(self):
        pass  # requestsは初期化不要

    async def close(self):
        self.session.close()

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
            # Step1: トップページ取得
            resp = self.session.get(top_url, timeout=15, allow_redirects=True)
            resp.encoding = resp.apparent_encoding
            html = resp.text

            # Step2: フォームリンクをAIで探す
            prompt = FIND_FORM_PROMPT.replace('{base_url}', top_url) + html[:8000]
            analysis = self._call_claude(prompt)
            analysis = self._extract_json(analysis)
            data = json.loads(analysis)

            if data.get('blocked'):
                result['status'] = 'スキップ'
                result['reason'] = data.get('blocked_reason', '営業禁止の記載あり')
                return result

            form_url = data.get('form_url')
            if not form_url:
                result['status'] = '失敗'
                result['reason'] = 'フォームリンクが見つかりませんでした'
                return result

            # 相対URLを絶対URLに変換
            form_url = urljoin(top_url, form_url)
            result['form_url'] = form_url

            # Step3: フォームページ取得
            resp2 = self.session.get(form_url, timeout=15, allow_redirects=True)
            resp2.encoding = resp2.apparent_encoding
            form_html = resp2.text

            # CAPTCHAチェック
            if 'recaptcha' in form_html.lower() or 'hcaptcha' in form_html.lower():
                result['status'] = '失敗'
                result['reason'] = 'CAPTCHA検知'
                return result

            # PDFチェック
            if form_url.lower().endswith('.pdf'):
                result['status'] = '失敗'
                result['reason'] = 'PDFフォーム（対応不可）'
                return result

            # Step4: AIでフォーム入力値を判定
            prompt2 = FILL_FORM_PROMPT.format(
                sender_json=json.dumps(sender, ensure_ascii=False),
                inquiry_type=sender.get('inquiry_type', '協力会社として取引のご案内'),
                form_html=form_html[:6000]
            )
            fill_str = self._call_claude(prompt2)
            fill_str = self._extract_json(fill_str)
            fill_data = json.loads(fill_str)

            # JavaScriptが必須の場合は失敗
            if fill_data.get('has_captcha'):
                result['status'] = '失敗'
                result['reason'] = 'CAPTCHA検知（フォーム内）'
                return result

            if fill_data.get('has_js_required'):
                result['status'] = '失敗'
                result['reason'] = 'JavaScript必須フォーム（対応不可）'
                return result

            # 入力できなかった項目を記録
            missing = fill_data.get('missing_fields', [])
            if missing:
                result['missing_fields'] = '、'.join(
                    [f"{m['field_name']}（{m['reason']}）" for m in missing]
                )

            # Step5: フォーム送信
            fields = fill_data.get('fields', {})
            action = fill_data.get('action') or form_url
            action = urljoin(form_url, action)
            method = (fill_data.get('method') or 'post').lower()

            if self.test_mode:
                # テストモード：送信せずに入力内容をログ
                if missing:
                    result['status'] = '部分送信（テスト）'
                    result['reason'] = '一部入力できない項目があります'
                else:
                    result['status'] = '送信完了（テスト）'
                    result['reason'] = f'入力項目数: {len(fields)}件 ／ テストのため未送信'
            else:
                # 本番モード：実際に送信
                if method == 'get':
                    send_resp = self.session.get(action, params=fields, timeout=15)
                else:
                    send_resp = self.session.post(action, data=fields, timeout=15)

                if send_resp.status_code in [200, 201, 302]:
                    if missing:
                        result['status'] = '部分送信'
                        result['reason'] = '一部入力できない項目があります'
                    else:
                        result['status'] = '送信完了'
                else:
                    result['status'] = '失敗'
                    result['reason'] = f'送信エラー（HTTPステータス: {send_resp.status_code}）'

        except requests.exceptions.Timeout:
            result['status'] = '失敗'
            result['reason'] = 'タイムアウト（サイトへの接続に時間がかかりすぎ）'
        except requests.exceptions.ConnectionError:
            result['status'] = '失敗'
            result['reason'] = 'サイトに接続できませんでした'
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
            model="claude-sonnet-4-5",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
