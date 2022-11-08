# -*- coding: utf-8 -*-

from flask import Flask
from time import sleep
from bs4 import BeautifulSoup
from flask import render_template, request, redirect, url_for
import requests, openpyxl, datetime

app = Flask(__name__)

PAGE = 5 

@app.route('/')
def index():
    return render_template(
        'index.html'
    )

@app.route('/', methods=['POST'])
def options():
    job_type = request.form.get('job')
    skill_list = request.form.getlist('skill')
    if len(skill_list) >= 2:
        skill_list.remove('index')
    sex_type = request.form.get('sex')
    age_group = request.form.get('age')
    pref = request.form.get('pref')
    idv = request.form.get('id')
    webm = request.form.get('webmeet')
    score = request.form.get('score')
    job_url = job_ui(job_type)
    skill_url = skill_ui(skill_list)
    sex_url = sex_ui(sex_type)
    age_url = age_ui(age_group)
    pref_url = pref_ui(pref)
    id_url = id_ui(idv)
    webm_url = webm_ui(webm)
    score_url = score_ui(score)
    value_list = [
        job_type, skill_list, sex_type, age_group, pref, idv, webm, score
    ]
    url_list = [
        job_url, skill_url, sex_url, age_url, pref_url, id_url, webm_url, score_url
    ]
    url = makeURL(url_list)
    excel(url, value_list)
    return render_template(
        'data.html',
        job=job_type,
        skill_list=skill_list,
        sex=sex_type,
        age=age_group,
        prefecture=pref,
        idv=idv,
        webm=webm,
        score=score,
        list=value_list
    )

def makeURL(rawList):
    filter_Obj = filter(None, rawList)
    filtered_list = list(filter_Obj)
    keepIndex0 = filtered_list[0]
    behindIndex1 = filtered_list[1:]
    option_url = keepIndex0 + '?' + '&'.join(behindIndex1)
    return option_url

def excel(searchURL, value_list):
    d_list = []
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "スクレイピング結果"
    sheet["A1"].value = "ユーザー名"
    sheet["B1"].value = "職種"
    sheet["C1"].value = "受注実績"
    sheet["D1"].value = "評価"
    sheet["E1"].value = "時間単価"
    sheet["F1"].value = "稼働可能時間/週"
    sheet["G1"].value = "ユーザーページへのリンク"
    sheet["I1"].value = "絞り込んだ項目"
    sheet.merge_cells('I1:J1')
    sheet["I2"].value = "職種"
    sheet["I3"].value = "スキル"
    sheet["I4"].value = "性別"
    sheet["I5"].value = "年齢層"
    sheet["I6"].value = "都道府県"
    sheet["I7"].value = "本人確認の有無"
    sheet["I8"].value = "Web会議可能か"
    sheet["I9"].value = "評価"
    sheet["J2"].value = value_list[0]
    skill_list = ', '.join(value_list[1])
    sheet["J3"].value = skill_list
    sheet["J4"].value = value_list[2]
    sheet["J5"].value = value_list[3]
    sheet["J6"].value = value_list[4]
    sheet["J7"].value = value_list[5]
    sheet["J8"].value = value_list[6]
    sheet["J9"].value = value_list[7]
    fill1 = openpyxl.styles.PatternFill(patternType='solid', fgColor='BAD1C2', bgColor='BAD1C2')
    fill2 = openpyxl.styles.PatternFill(patternType='solid', fgColor='F6F6C9', bgColor='F6F6C9')
    sheet["A1"].fill = sheet["B1"].fill = sheet["C1"].fill = sheet["D1"].fill = sheet["E1"].fill = sheet["F1"].fill = sheet["G1"].fill = fill1
    sheet["I1"].fill = fill2
    for i in range(1, PAGE+1):
        page_num = '&page=' + str(i)
        add_urls = searchURL + page_num
        base_url = 'https://crowdworks.jp/public/employees/'
        base_userURL = 'https://crowdworks.jp'
        url = base_url + add_urls
        target_url = url.format(1)
        r = requests.get(target_url)
        soup = BeautifulSoup(r.text, 'html.parser')
        contents = soup.find_all('div', class_="member_item")
        sleep(1)
        for content in contents:
            user_name = content.find('span', class_="username").text
            h2 = content.select(".item_title a") # href部分取得
            user_url = base_userURL + h2[0].get('href') # ユーザーの個別url部分を取得 ex)'/public/employees/0000000' <str型>
            user_occupation = content.find('span', class_="user_occupation").text
            misc = content.find('div', class_="misc")
            count = content.find('span', class_="count").text
            score = content.find('span', class_="score").text
            score_list = misc.find_all('ul', class_="cw-list_inline")
            score_data1 = score_list[0]
            score_data2 = score_list[1]
            wage_data = score_data1.find('li', class_="")
            wage_data.find('span', {"class": "data_label"}).extract()
            wage = wage_data.text
            work_time_data = score_data2.find('li', class_="")
            work_time_data.find('span', {"class": "data_label"}).extract()
            work_time = work_time_data.text
            d = [
                user_name,
                user_occupation,
                count,
                score,
                wage,
                work_time,
                user_url
            ]
            d_list.append(d)
    write_list_2d(sheet, d_list, 2, 1)
    sheet.freeze_panes = "A2"
    time = now()
    wb.save( time + value_list[0] + '.xlsx')
    wb.close()
    return redirect(url_for('index'))

def write_list_2d(sheet, l_2d, start_row, start_col):
    for y, row in enumerate(l_2d):
        for x, cell in enumerate(row):
            sheet.cell(row=start_row + y, column=start_col + x, value=l_2d[y][x])

def now():
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    a = now.strftime('%Y%m%d ')
    return a

def job_ui(value):
    if value == "ITエンジニア(全般)":
        url_parts = "ogroup/1"
    if value == "システムエンジニア(SE)":
        url_parts = "occupation/1"
    if value == "プログラマ(PG)":
        url_parts = "occupation/4"
    if value == "プログラマ(スマートフォン)":
        url_parts = "occupation/3"
    if value == "Androidアプリエンジニア":
        url_parts = "occupation/53"
    if value == "AIエンジニア":
        url_parts = "occupation/98"
    if value == "ITコンサルタント":
        url_parts = "occupation/48"
    if value == "セキュリティエンジニア":
        url_parts = "occupation/99"
    if value == "ネットワークエンジニア":
        url_parts = "occupation/5"
    if value == "サーバーエンジニア・インフラエンジニア":
        url_parts = "occupation/6"
    if value == "データベースエンジニア":
        url_parts = "occupation/49"
    if value == "デスクトップアプリ・業務アプリ開発者":
        url_parts = "occupation/50"
    if value == "テスター":
        url_parts = "occupation/51"
    if value == "プロジェクトマネージャー(PM)":
        url_parts = "occupation/15"
    if value == "その他エンジニア":
        url_parts = "occupation/7"
    return url_parts

def sex_ui(value):
    if value == "男性":
        url_parts = "sex=male"
    if value == "女性":
        url_parts = "sex=female"
    if value == "":
        url_parts = None
    return url_parts

def skill_ui(value_list):
    if value_list[0] == 'index':
        return None
    base_skills = [
        "PHP",
        "JavaScript",
        "Java",
        "Python",
        "MySQL",
        "HTML",
        "CSS",
        "AWS",
        "Linux",
        "jQuery"
    ]
    value_set = set(value_list)
    base_set = set(base_skills)
    matched_list = list(value_set & base_set)
    base_skill_urls = "skill_id="
    base_first = matched_list[0]
    if base_first == 'HTML':
        first = "3"
        base_skill_urls = base_skill_urls + first
    if base_first == 'CSS':
        first = "5"
        base_skill_urls = base_skill_urls + first
    if base_first == 'JavaScript':
        first = "8"
        base_skill_urls = base_skill_urls + first
    if base_first == 'jQuery':
        first = "9"
        base_skill_urls = base_skill_urls + first
    if base_first == 'Linux':
        first = "14"
        base_skill_urls = base_skill_urls + first
    if base_first == 'Python':
        first = "21"
        base_skill_urls = base_skill_urls + first
    if base_first == 'PHP':
        first = "24"
        base_skill_urls = base_skill_urls + first
    if base_first == 'Java':
        first = "40"
        base_skill_urls = base_skill_urls + first
    if base_first == 'MySQL':
        first = "62"
        base_skill_urls = base_skill_urls + first
    if base_skill_urls[0] == 'AWS':
        first = "1038"
        base_skill_urls = base_skill_urls + first
    add_part = ''
    for i in matched_list:
        temp = i
        if temp == base_first:
            continue
        if temp == 'HTML':
            afters = "%2C3"
            add_part = add_part + afters
        if temp == 'CSS':
            afters = "%2C5"
            add_part = add_part + afters
        if temp == 'JavaScript':
            afters = "%2C8"
            add_part = add_part + afters
        if temp == 'jQuery':
            afters = "%2C9"
            add_part = add_part + afters
        if temp == 'Linux':
            afters = "%2C14"
            add_part = add_part + afters
        if temp == 'Python':
            afters = "%2C21"
            add_part = add_part + afters
        if temp == 'PHP':
            afters = "%2C24"
            add_part = add_part + afters
        if temp == 'Java':
            afters = "%2C40"
            add_part = add_part + afters
        if temp == 'MySQL':
            afters = "%2C62"
            add_part = add_part + afters
        if temp == 'AWS':
            afters = "%2C1038"
            add_part = add_part + afters
    url_parts = base_skill_urls + add_part
    return url_parts

def age_ui(value):
    if value == "10代":
        url_parts = "age=10"
    if value == "20代":
        url_parts = "age=20"
    if value == "30代":
        url_parts = "age=30"
    if value == "40代":
        url_parts = "age=40"
    if value == "50代":
        url_parts = "age=50"
    if value == "60歳以上":
        url_parts = "age=60"
    if value == "":
        url_parts = None
    return url_parts

def pref_ui(value):
    if value == "":
        url_parts = None
    if value == "北海道":
        url_parts = "prefecture_id=1"
    if value == "青森県":
        url_parts = "prefecture_id=2"
    if value == "岩手県":
        url_parts = "prefecture_id=3"
    if value == "宮城県":
        url_parts = "prefecture_id=4"
    if value == "秋田県":
        url_parts = "prefecture_id=5"
    if value == "山形県":
        url_parts = "prefecture_id=6"
    if value == "福島県":
        url_parts = "prefecture_id=7"
    if value == "茨城県":
        url_parts = "prefecture_id=8"
    if value == "栃木県":
        url_parts = "prefecture_id=9"
    if value == "群馬県":
        url_parts = "prefecture_id=10"
    if value == "埼玉県":
        url_parts = "prefecture_id=11"
    if value == "千葉県":
        url_parts = "prefecture_id=12"
    if value == "東京都":
        url_parts = "prefecture_id=13"
    if value == "神奈川県":
        url_parts = "prefecture_id=14"
    if value == "新潟県":
        url_parts = "prefecture_id=15"
    if value == "富山県":
        url_parts = "prefecture_id=16"
    if value == "石川県":
        url_parts = "prefecture_id=17"
    if value == "福井県":
        url_parts = "prefecture_id=18"
    if value == "山梨県":
        url_parts = "prefecture_id=19"
    if value == "長野県":
        url_parts = "prefecture_id=20"
    if value == "岐阜県":
        url_parts = "prefecture_id=21"
    if value == "静岡県":
        url_parts = "prefecture_id=22"
    if value == "愛知県":
        url_parts = "prefecture_id=23"
    if value == "三重県":
        url_parts = "prefecture_id=24"
    if value == "滋賀県":
        url_parts = "prefecture_id=25"
    if value == "京都府":
        url_parts = "prefecture_id=26"
    if value == "大阪府":
        url_parts = "prefecture_id=27"
    if value == "兵庫県":
        url_parts = "prefecture_id=28"
    if value == "奈良県":
        url_parts = "prefecture_id=29"
    if value == "和歌山県":
        url_parts = "prefecture_id=30"
    if value == "鳥取県":
        url_parts = "prefecture_id=31"
    if value == "島根県":
        url_parts = "prefecture_id=32"
    if value == "岡山県":
        url_parts = "prefecture_id=33"
    if value == "広島県":
        url_parts = "prefecture_id=34"
    if value == "山口県":
        url_parts = "prefecture_id=35"
    if value == "徳島県":
        url_parts = "prefecture_id=36"
    if value == "香川県":
        url_parts = "prefecture_id=37"
    if value == "愛媛県":
        url_parts = "prefecture_id=38"
    if value == "高知県":
        url_parts = "prefecture_id=39"
    if value == "福岡県":
        url_parts = "prefecture_id=40"
    if value == "佐賀県":
        url_parts = "prefecture_id=41"
    if value == "長崎県":
        url_parts = "prefecture_id=42"
    if value == "熊本県":
        url_parts = "prefecture_id=43"
    if value == "大分県":
        url_parts = "prefecture_id=44"
    if value == "宮崎県":
        url_parts = "prefecture_id=45"
    if value == "鹿児島県":
        url_parts = "prefecture_id=46"
    if value == "沖縄県":
        url_parts = "prefecture_id=47"
    return url_parts

def id_ui(value):
    if value == "指定なし":
        url_parts = None
    if value == "済み":
        url_parts = "identity_verified=true"
    return url_parts

def webm_ui(value):
    if value == "指定なし":
        url_parts = None
    if value == "可能":
        url_parts = "web_meeting=true"
    return url_parts

def score_ui(value):
    if value == "":
        url_parts = None
    if value == "3.0以上":
        url_parts = "average_score_as_employee=3"
    if value == "4.0以上":
        url_parts = "average_score_as_empsloyee=4"
    if value == "5.0":
        url_parts = "average_score_as_employee=5"
    return url_parts


if __name__ == '__main__':
    app.run()