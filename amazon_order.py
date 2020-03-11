import sys
import os
import subprocess
import re

from robobrowser import RoboBrowser

# 認証の情報は環境変数から取得
AMAZON_EMAIL = os.environ['AMAZON_EMAIL']
AMAZON_PASSWORD = os.environ['AMAZON_PASSWORD']

# RoboBrowserオブジェクトを作成する
browser = RoboBrowser(
    parser="html.parser",
    user_agent='Mozilla/5.0 (Windows NT 10.0; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'
)


def main():
    # 注文履歴のぺージを開く
    print("Navigaiting...", file=sys.stderr)

    browser.open('https://www.amazon.co.jp/gp/css/order-history')
    # サインインページにリダイレクトされていることを確認
    assert 'Amazonログイン' in browser.parsed.title.string

    # name="signIn"というサインインフォームを埋める
    # フォームのname属性の値はブラウザーの開発者ツールで確認できる
    form = browser.get_form(attrs={'name': 'signIn'})
    form['email'] = AMAZON_EMAIL  # name="email"という入力ボックスを埋める
    form['password'] = AMAZON_PASSWORD
    # フォームを送信。正常にログインするにはRefererヘッダーとAccept-Languageヘッダーが必要
    print('Singning in...', file=sys.stderr)
    browser.submit_form(form, headers={
        'Referer': browser.url,
        'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
    })

    # ログインに失敗する場合は、次の行のコメントを外してHTMLのソースを確認するといい。
    # print(browser.parsed.pretitify())

    # ページャーをたどる
    # while True:
    #   assert '注文履歴' in browser.parsed.title.string(
    #      )

    print(browser.parsed.title.string)
    print(browser.find('form'))
    two_step_authentication = ['oathtool', '--totp', '--base32',
                               'GDD6NZCJ3FPZVIRCQQQHLTYDDZXMWKNSHXIKFQDPPBD36CBRHDNQ']
    two_step_key = re.findall(
        r'\d+', subprocess.check_output(two_step_authentication).decode('UTF-8'))[0]


if __name__ == '__main__':
    main()
