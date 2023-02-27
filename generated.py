from playwright.sync_api import (
    Playwright,
    sync_playwright,
    expect,
    TimeoutError as PlaywrightTimeoutError,
)
from datetime import datetime as dt
import yaml
import time
import datetime

# YAMLファイルを読み込む
with open("config.yaml", "r", encoding="utf-8") as f:
    data = yaml.safe_load(f)

#############################クラスエリア################################
class Issue:
    def __init__(self, issue: dict):
        self.name = issue["name"]
        self.number = issue["number"]
        self.order_quantity = issue["order_quantity"]
        self.limit_ordered_status = issue["limit_ordered_status"]
        self.market_ordered_status = issue["market_ordered_status"]
        self.limit_ordered_number = issue["limit_ordered_number"]
        self.market_ordered_number = issue["market_ordered_number"]


class Account:
    def __init__(self, account: dict):
        self.shitencode = account["shitencode"]
        self.kouzabangou = account["kouzabangou"]
        self.loginpassword = account["loginpassword"]
        self.torihikipassword = account["torihikipassword"]


#############################関数エリア################################


def data_update_save(data, issues):
    # YAMLファイルから読み込んだデータと、Issueオブジェクトのリストを受け取り、
    # Issueオブジェクトの各属性を使って、YAMLファイルのデータを更新します。
    # 最後に、更新されたデータをYAMLファイルに書き込みます。
    for i in range(0, len(issues)):

        # 値を更新する
        # data['issues'][i]['order_quantity'] = issues.order_quantity
        data["issues"][i]["limit_ordered_status"] = issues[i].limit_ordered_status
        data["issues"][i]["market_ordered_status"] = issues[i].market_ordered_status
        data["issues"][i]["limit_ordered_number"] = issues[i].limit_ordered_number
        data["issues"][i]["market_ordered_number"] = issues[i].market_ordered_number
        data["issues"][i]["name"] = issues[i].name

    # YAMLファイルに書き込む
    with open("config.yaml", "w", encoding="utf-8") as f:
        # with open('config.yaml', 'w', encoding='cp932') as f:
        yaml.dump(data, f, allow_unicode=True)


def is_within_timeframe(*timeframes) -> bool:
    """
    指定された時間帯のいずれかの範囲内にあるかどうかを判定する関数
    """
    now = dt.now().time()
    for start, end in timeframes:
        start_time = datetime.time(
            hour=start.hour, minute=start.minute, second=start.second
        )
        end_time = datetime.time(hour=end.hour, minute=end.minute, second=end.second)
        if start_time <= now <= end_time:
            return True
    return False


def is_market_available() -> bool:
    """
    営業日の指定された時間帯のみ実行される関数
    """
    # 営業日の時間帯
    business_day_timeframes = [
        (datetime.time(0, 0), datetime.time(2, 0)),
        (datetime.time(5, 0), datetime.time(11, 30)),
        (datetime.time(11, 35), datetime.time(15, 30)),
        (datetime.time(17, 0), datetime.time(20, 15)),
        (datetime.time(20, 20), datetime.time(23, 59)),
    ]

    if is_within_timeframe(*business_day_timeframes):
        print("稼働時間内です")
        return True

    # 土日の時間帯
    weekend_timeframes = [
        (datetime.time(0, 0), datetime.time(2, 0)),
        (datetime.time(5, 0), datetime.time(24, 0)),
    ]

    if is_within_timeframe(*weekend_timeframes):
        print("土日の稼働時間内です")
        return True

    print("稼働時間外です")
    return False


def thursday_night_function() -> True:
    """
    木曜の夜18時00分から23時59分までのみ実行する関数
    """
    # 木曜の曜日番号は3
    if datetime.datetime.today().weekday() == 3:
        # 18時00分のtimeオブジェクト
        start_time = datetime.time(hour=18, minute=0, second=0)
        # 23時59分のtimeオブジェクト
        end_time = datetime.time(hour=23, minute=59, second=59)
        # 現在の時刻が範囲内にあるかどうかを判定する
        if start_time <= datetime.datetime.now().time() <= end_time:
            # 実行する処理
            print("木曜の夜18時00分から23時59分の範囲内です: Trueを返します")
            return True
    print("木曜の夜18時00分から23時59分ではありません: Falseを返します")
    return False


def run(playwright: Playwright, data) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    # タイムアウトを0に
    context.set_default_timeout(0)
    page = context.new_page()

    def is_element_found_with_text(page, text):
        try:
            element = page.locator(':has-text("{text}")'.format(text=text))
            if element.is_visible():
                return True
        except:
            pass
        return False

    def nikko_login(page, account):

        try:
            page.goto("https://trade.smbcnikko.co.jp/Etc/1/webtoppage/")
            page.locator("#padInput0").click()
            page.locator("#padInput0").fill(str(account.shitencode))
            page.locator("#padInput1").click()
            page.locator("#padInput1").fill(str(account.kouzabangou))
            page.locator("#padInput2").click()
            page.locator("#padInput2").fill(str(account.loginpassword))
            page.get_by_role("button", name="ログイン").click()
        except:
            page.screenshot(path="login_error_screenshot.png")
            return False

        return True

    def check_availability(page, issue_number: str):
        try:
            # page.pause()
            page.locator(
                "body > div:nth-child(4) > table:nth-child(1) > tbody > tr > td:nth-child(1) > a > img"
            ).click()
            page.locator(
                "body > div:nth-child(5) > table:nth-child(1) > tbody > tr > td:nth-child(3) > a > img"
            ).click()

        except Exception as e:
            # print(e)
            pass

        try:
            lnk_Toriatsukaimeigaraichiran = page.get_by_role(
                "cell",
                name="取扱銘柄一覧 新規買付 新規売付 建玉一覧（返済・現引・現渡） 注文約定一覧・取消・訂正 信用取引担保状況 担保状況詳細 代用証券一覧 保証金振替指示 取引注意銘柄 株式約定通知メール 信用電子交付書面一覧 株価通知メール 取引のご案内 信用取引の利用申込",
                exact=True,
            ).get_by_role("link", name="取扱銘柄一覧")
            if lnk_Toriatsukaimeigaraichiran.is_visible():
                lnk_Toriatsukaimeigaraichiran.click()

        except PlaywrightTimeoutError:
            page.screenshot(path="toriatsukaimeigara_error_screenshot.png")
            return False

        # 取引銘柄ページに移動できなかった場合にFalseを返す
        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 取扱銘柄一覧 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：取引銘柄ページ")

            else:
                print(dt.now().isoformat() + "  :  " + "ページ位置：取引銘柄ページに移動できませんでした")
                return False

        # 銘柄の選択
        try:
            # page.pause()
            page.locator('input[name="searchmeig"]').fill(issue_number)
        except PlaywrightTimeoutError:
            print(dt.now().isoformat() + "  :  " + "タイムアウトエラーが発生しました。スクリーンショットを保存します。")
            page.screenshot(path=f"{issue_number}_error_screenshot.png")
            return False
        page.get_by_role("row", name="検索", exact=True).get_by_role(
            "button", name="検索"
        ).click()

        # 存在判定
        element = page.locator('a:has-text("一般新規売")')
        if element.is_visible():
            print(dt.now().isoformat() + "  :  " + "一般信用新規売りが見つかりました")
            return True
        else:
            print(dt.now().isoformat() + "  :  " + "一般信用新規売りが見つかりませんでした")
            return False

    def check_ordered(page, issues):
        # 同意 出ていればチェックをする
        try:
            chk = page.get_by_role("checkbox")
            if chk.is_visible():
                chk.click()
            button = page.get_by_role("button", name="同意します(先へ進む)")
            if button.is_visible():
                button.click()
        except:
            pass
        # TODO: オーダーされた注文を管理する
        page.get_by_role("row", name="トップ 投資情報 お取引 残高の確認 商品ガイド 各種お手続き").get_by_role(
            "link", name="お取引"
        ).click()
        # 約定一覧ページに移動
        try:
            lnk_Yakujouichiran = page.get_by_role(
                "cell",
                name="取扱銘柄一覧 新規買付 新規売付 建玉一覧（返済・現引・現渡） 注文約定一覧・取消・訂正 信用取引担保状況 担保状況詳細 代用証券一覧 保証金振替指示 取引注意銘柄 株式約定通知メール 信用電子交付書面一覧 株価通知メール 取引のご案内 信用取引の利用申込",
                exact=True,
            ).get_by_role("link", name="注文約定一覧・取消・訂正")
            if lnk_Yakujouichiran.is_visible():
                lnk_Yakujouichiran.click()

        except PlaywrightTimeoutError:
            page.screenshot(path="yakujouichiran_error_screenshot.png")
            return False

        # 約定一覧ページに移動できなかった場合にFalseを返す
        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 注文約定一覧・取消・訂正 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 注文約定一覧・取消・訂正 -")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 注文約定一覧・取消・訂正 -に移動できませんでした"
                )
                return False
        tbl = page.locator(
            '//*[@id="printzone"]/div[2]/table/tbody/tr/td/table[1]/tbody/tr'
        )

        for i in range(1, tbl.count()):
            issue_number = tbl.nth(i).inner_text().split("\t")[0].split("／")[0]
            if tbl.nth(i).inner_text().split("\t")[3].split("\n")[1] == "成行":
                order_type = "成行"
            else:
                order_type = "指値"
            order_status = tbl.nth(i).inner_text().split("\t")[6]
            order_number = tbl.nth(i).inner_text().split("\t")[7]

            active_order_statuses = ["注文中", "取消中", "訂正中", "訂正済", "注文済"]

            # 注文状況を取得する
            for issue in issues:
                # 記録するべきorder_statusの場合
                if order_status in active_order_statuses:
                    # order_numberが無い場合は共に記録、ある場合は、order_numberが一致するときのみ記録
                    if str(issue.number) == issue_number:
                        if order_type == "成行":
                            if issue.market_ordered_number is None:
                                issue.market_ordered_number = order_number
                                issue.market_ordered_status = order_status
                            else:
                                if issue.market_ordered_number == order_number:
                                    issue.market_ordered_status = order_status
                        else:
                            if issue.limit_ordered_number is None:
                                issue.limit_ordered_number = order_number
                                issue.limit_ordered_status = order_status
                            else:
                                if issue.limit_ordered_number == order_number:
                                    issue.limit_ordered_status = order_status
        return issues

    def place_limit_order(page, suryou: int, account, order_type: str):
        # TODO: 指値注文を入れる
        element = page.locator('a:has-text("一般新規売")')
        if element.is_visible():
            element.click()
        else:
            print(dt.now().isoformat() + "  :  " + "一般信用新規売りが見つかりませんでした")
            return False

        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )

        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 新規売り注文 注文入力 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 新規売り注文 注文入力")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 新規売り注文 注文入力に移動できませんでした"
                )
                return False

        # 数量の入力
        textbox_amount = page.locator('//*[@id="isuryo"]/table/tbody/tr[1]/td[1]/input')
        textbox_amount.fill(str(suryou))
        print(dt.now().isoformat() + "  :  " + "数量:  " + str(suryou))

        if order_type == "成行":
            page.get_by_label("成行").check()
            page.get_by_text("今週中", exact=True).click()

        elif order_type == "指値":

            # 制限値幅上限の取得
            str_sashinejougen = page.locator(
                '//*[@id="itanka"]/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[3]/span[2]'
            ).inner_text()

            str_sashinejougen = str(str_sashinejougen.replace(",", ""))

            # 指値の入力
            textbox_sashine_kakaku = page.locator('input[name="kakaku"]')
            textbox_sashine_kakaku.click()
            textbox_sashine_kakaku.fill(str_sashinejougen)
            print(dt.now().isoformat() + "  :  " + "指値:  " + str(str_sashinejougen))
        else:
            print("不明なorder_typeです")
            return False

        page.get_by_text("特定口座").click()
        page.get_by_label("今週中").check()

        # 同意 出ていればチェックをする
        try:
            chk_element = page.get_by_label(
                "私は、「日興イージートレード信用取引の契約締結前交付書面（インターネット取引）」、「金融商品取引法により禁止されている取引」、「注意事項」、「空売り価格規制の注意事項」の記載内容を確認・理解しました。",
            )
            if chk_element.is_visible():
                chk_element.check()
        except:
            pass
        page.get_by_role("button", name="注文内容を確認する").click()

        # 同意 出ていればチェックをする
        try:
            chk_element = page.get_by_role("checkbox")
            if chk_element.is_visible():
                chk_element.check()
            btn_element = page.get_by_role("button", name="同意します(先へ進む)")
            if btn_element.is_visible():
                btn_element.click()
        except:
            pass

        # ページ遷移　確認
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 新規売り注文 内容確認 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 新規売り注文 内容確認 -")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 新規売り注文 内容確認 -に移動できませんでした"
                )
                return False

        # 取引パスワードの入力
        page.locator("#padInput").click()
        page.locator("#padInput").fill(account.torihikipassword)
        btn_element = page.get_by_role("button", name="注文する")
        if btn_element.is_visible():
            btn_element.click()

        # ページ遷移　確認
        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 新規売り注文　受付完了 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 新規売り注文 受付完了 -")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 新規売り注文 受付完了 -に移動できませんでした"
                )
                return False

        # page.pause()
        return True

    def change_orders_to_market(account):
        # TODO: 指値注文を成行注文に変更する
        # 同意 出ていればチェックをする
        #  オーダーされた注文をチェックしていく
        page.get_by_role("row", name="トップ 投資情報 お取引 残高の確認 商品ガイド 各種お手続き").get_by_role(
            "link", name="お取引"
        ).click()

        try:
            chk = page.get_by_role("checkbox")
            if chk.is_visible():
                chk.click()
            button = page.get_by_role("button", name="同意します(先へ進む)")
            if button.is_visible():
                button.click()
        except:
            pass

        # 約定一覧ページに移動
        try:
            lnk_Yakujouichiran = page.get_by_role(
                "cell",
                name="取扱銘柄一覧 新規買付 新規売付 建玉一覧（返済・現引・現渡） 注文約定一覧・取消・訂正 信用取引担保状況 担保状況詳細 代用証券一覧 保証金振替指示 取引注意銘柄 株式約定通知メール 信用電子交付書面一覧 株価通知メール 取引のご案内 信用取引の利用申込",
                exact=True,
            ).get_by_role("link", name="注文約定一覧・取消・訂正")
            if lnk_Yakujouichiran.is_visible():
                lnk_Yakujouichiran.click()

        except PlaywrightTimeoutError:
            page.screenshot(path="yakujouichiran_error_screenshot.png")
            return False

        # 約定一覧ページに移動できなかった場合にFalseを返す
        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 注文約定一覧・取消・訂正 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 注文約定一覧・取消・訂正 -")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 注文約定一覧・取消・訂正 -に移動できませんでした"
                )
                return False
        tbl = page.locator(
            '//*[@id="printzone"]/div[2]/table/tbody/tr/td/table[1]/tbody/tr'
        )

        for i in range(1, tbl.count()):
            issue_number = tbl.nth(i).inner_text().split("\t")[0].split("／")[0]
            if tbl.nth(i).inner_text().split("\t")[3].split("\n")[1] == "成行":
                order_type = "成行"
            else:
                order_type = "指値"
            order_status = tbl.nth(i).inner_text().split("\t")[6]
            order_number = tbl.nth(i).inner_text().split("\t")[7]

            # 成行きに変更するの訂正済みと注文済みのみ
            active_order_statuses = ["注文済"]

            # 注文状況を取得する
            for issue in issues:
                # 注文訂正するべきorder_statusの場合
                # if order_status in active_order_statuses:
                if True:
                    # order_numberが指値にあり、成行に無い場合、注文の訂正を行う
                    if str(issue.number) == issue_number:
                        if order_type == "指値":
                            if issue.market_ordered_number is None:
                                print("指値注文の成行注文への訂正処理を行います")
                                tbl.nth(i).get_by_role(
                                    "link", name="訂正", exact=True
                                ).click()
                                place_market_order(account)
                            else:
                                print("既に成行注文が存在するため、指値の変更を行いません")

    def place_market_order(account):
        ######################################
        # ページ遷移　確認
        title_element = page.locator(
            '//*[@id="printzone"]/div[1]/table/tbody/tr/td[2]/h1'
        )
        if title_element.is_visible():
            if title_element.inner_text() == "信用取引 - 新規売り注文訂正　訂正入力 -":
                print(dt.now().isoformat() + "  :  " + "ページ位置：信用取引 - 新規売り注文訂正　訂正入力 -")

            else:
                print(
                    dt.now().isoformat()
                    + "  :  "
                    + "ページ位置：信用取引 - 新規売り注文訂正　訂正入力 -に移動できませんでした"
                )
                return False
        try:
            page.get_by_label("成行").check()
            page.get_by_role("button", name="訂正内容を確認する").click()
            page.locator("#padInput").click()
            page.locator("#padInput").fill(account.torihikipassword)
            page.get_by_role("button", name="この内容で訂正").click()
            page.get_by_role("link", name="注文約定一覧・取消・訂正").click()
        except:
            print(dt.now().isoformat() + "  :  " + "新規売り注文訂正に失敗しました")
            return False

        return True

    #############################コードエリア################################
    if is_market_available():
        print("処理を開始します")
    else:
        print("処理を中断します")
        exit()

    issues = []
    for issue in data.get("issues"):
        issues.append(Issue(issue))

    account = Account(data.get("accounts"))

    # ログインする
    nikko_login(page, account)

    # オーダー状況の確認を行う
    issues = check_ordered(page, issues)
    # オーダーデータをアップデートして保存する
    data_update_save(data, issues)
    for issue in issues:
        print(dt.now().isoformat() + "  :  " + issue.name + "をチェック中")
        # issue_number = issue['number']

        if check_availability(page, str(issue.number)):
            # 購入処理
            if issue.limit_ordered_status:
                print("既に指値注文が存在するため、注文しません。")
            else:
                res = place_limit_order(page, issue.order_quantity, account)
                # オーダー状況の確認を行う
                issues = check_ordered(page, issues)
        else:
            print("------------------------")
            # 次の銘柄へ
            continue

    data_update_save(data, issues)

    # 木曜の夜18時00分から23時59分までのみ実行
    if thursday_night_function():
        # 成行=>指値
        place_market_order(account)

    page.close()

    # ---------------------
    context.close()
    browser.close()
    exit()


with sync_playwright() as playwright:
    run(playwright, data)
