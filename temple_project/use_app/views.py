from datetime import date
import datetime
from django.shortcuts import render, HttpResponse, HttpResponseRedirect, render_to_response, redirect
from django.http import JsonResponse
from .forms import homeform, peopleform, activity_form, choose_form, login_form, fix_peopleform

from .models import Home, People_data, activity_data, history_data
#try
from django.contrib.auth.forms import UserCreationForm
from django.contrib import auth
from django.urls import reverse
from mailmerge import MailMerge
from django.contrib.auth import logout

import os
import json
from django.contrib.auth.decorators import login_required

import comtypes.client
import time

from docxtpl import DocxTemplate
from docx.enum.section import WD_ORIENT
from docx import Document
import csv

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

from use_app import LunarSolarConverter


@login_required
def x_try(request):
    x = {}
    x["z"] = [{
        'people': [
            '柯星雯 本命 己亥 年 一十 月 二八 號   生行庚 WAIT 歲 ',
            '楊逸凡 本命 丁酉 年 一十 月 二七 號   生行庚 WAIT 歲 ',
            '楊雅嵐 本命 戊戌 年 一十 月 二九 號   生行庚 WAIT 歲 '
        ],
        'address':
        '048544412',
        'one_people':
        '楊雅嵐'
    }, {
        'people': [
            '羅弘寧 本命 戊寅 年 一十 月 二二 號   生行庚 WAIT 歲 ',
            '李冰茜 本命 庚寅 年 一十 月 四 號   生行庚 WAIT 歲 ',
            ' 蔣原杰 本命 乙酉 年 七 月 二八 號   生行庚 WAIT 歲 '
        ],
        'address':
        '048548888',
        'one_people':
        '蔣原杰'
    }, {
        'people': [
            '許 雅喜 本命 丁酉 年 二 月 一八 號   生行庚 WAIT 歲 ',
            '雅文 本命 庚子 年 一十 月 二八 號   生行庚 WAIT 歲 ',
            '陳家銘 本命 癸巳 年 一十 月 二八 號   生行庚 WAIT 歲 '
        ],
        'address':
        '077632214',
        'one_people':
        '陳家銘'
    }]
    #{"bye":[[[{"table_name":"點光明燈者"},[{"table_data":"陳閔致、曹美雲、蕭孟勳、劉美惠、柯星雯、楊逸凡、楊雅嵐"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"陳恭宜、陳依光、曹志嘉、謝純鑫"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"許雅喜、李雅婷、雅文"}]]]]}
    try:

        for c in x["z"]:
            c["address"] = Home.objects.get(home_phone=c["address"]).address

        tpl = DocxTemplate(
            r"C:\Users\asd19\Downloads\try_git\temple_project\files\files\mode1.docx"
        )

        if date.today().month >= 10:
            x["year"] = twelve(int(date.today().year) + 1)
        else:
            x["year"] = twelve(date.today().year)
        x["title"] = "標題 未處理"
        tpl.render(x)
        tpl.save(r"C:\Users\asd19\Downloads\tryw.docx")
        os.system(r"C:\Users\asd19\Downloads\tryw.docx")
    except Exception as e:
        cc = e

    return render(request, "try.html", locals())



def data_up(request):
    get_id = request.GET.get("id", None)
    history_data.objects.filter(id=get_id).update(history=request.GET.get("data",None))
    data = {"OK":"已經更新"}
    return JsonResponse(data)

def old(request,pk):
    x_form = choose_form(request.POST or None)
    x_max = Home.objects.all().count()
    use_id = pk
    the_object = history_data.objects.get(pk=pk).history
  #  [{&quot;048544412&quot;:&quot;柯星雯,楊逸凡,楊雅嵐,測試名,第煙火&quot;},{&quot;048548888&quot;:&quot;陳政姍,羅弘寧,李冰茜,蔣原杰,測試2,測試1,測試3&quot;},{&quot;077632214&quot;:&quot;雅文,陳家銘&quot;},{&quot;048789301&quot;:&quot;葉宇軒,郭靜怡&quot;}]
    return render(request, "old_activity.html", locals())

def new(request):
    x_form = choose_form(request.POST or None)
    x_max = Home.objects.all().count()
    return render(request, "join_activity.html", locals())


def csv_add(request):

    if (request.method == "POST"):
        homes = request.FILES.get('home')
        people = request.FILES.get('people')

        if homes == "" or people == "":
            error = "請一次輸入兩個檔案"
            return render(request, "up_date.html", locals())

        if homes.name.split(".")[-1] != "csv" and people.name.split(
                ".")[-1] != "csv":
            error = "請輸入csv檔"
            return render(request, "up_date.html", locals())

        fobj = open(os.path.join(BASE_DIR, "people", homes.name), 'wb')
        for line in homes.chunks():
            fobj.write(line)
        fobj.close()

        fobj = open(os.path.join(BASE_DIR, "people", people.name), 'wb')
        for line in people.chunks():
            fobj.write(line)
        fobj.close()

        use_path = os.path.join(BASE_DIR, "people")
        home_path = os.path.join(use_path, homes.name)
        people_path = os.path.join(use_path, people.name)
        with open(home_path, newline='') as csvfile:

            homes = csv.reader(csvfile)

            one = 0

            for row in homes:

                if one != 0:
                    if not Home.objects.filter(home_phone=row[0]).exists():
                        Home.objects.create(home_phone=row[0], address=row[1])

                    home_id = Home.objects.get(home_phone=row[0]).pk
                    with open(people_path, newline='') as peoplefile:
                        peoples = csv.reader(peoplefile)
                        for people in peoples:
                            if people[0] != "信眾名字":
                                print("e")
                                if row[0] == people[4]:
                                    people[1] = people[1].replace("/", "-")
                                    if people[3] == "男":
                                        people[3] = "male"
                                    else:
                                        people[3] = "female"
                                    if not People_data.objects.filter(
                                            home_id=home_id,
                                            name=people[0]).exists():
                                        People_data.objects.create(
                                            name=people[0],
                                            birthday=people[1],
                                            time=people[2],
                                            gender=people[3],
                                            home_id=home_id)
                one = 1

    return render(request, "up_date.html", locals())


def home_page(request, pk):

    return render(request, "index.html", locals())


def validate_get_table(request):
    use_file = request.GET.get("use_file", None)
    get_activity_ID = activity_data.objects.get(id=use_file)
    data = {"reslut": get_activity_ID.table_name}
    return JsonResponse(data)


def validate_get_Home(request):
    start = request.GET.get("start", None)
    end = request.GET.get("end", None)
    the_data = Home.objects.all()[int(start):int(end)]
    get_allname_array = []
    for i in range(the_data.count()):
        get_allname_array.append(the_data[i].home_phone)

    data = {"reslut": ' '.join(get_allname_array)}
    return JsonResponse(data)


def validate_people_all_date(request):
    get_phone = request.GET.get("phone", None)
    Get_home_id = Home.objects.get(home_phone=get_phone).id
    the_data = People_data.objects.filter(home_id=Get_home_id)
    get_allname_array = []
    for i in range(len(the_data)):
        date = the_data[i].birthday
        output = the_data[i].name + " 本命 " + twelve(
            date.year) + " 年 " + time_chinese(
                date.month) + " 月 " + time_chinese(
                    date.day) + " 號 " + "  生行庚 " + time_chinese(
                        year(date)) + " 歲 "
        get_allname_array.append(the_data[i].name + "|" + output + "|" + "F")

    data = {"reslut": '㊣'.join(get_allname_array)}
    return JsonResponse(data)


def year(x):

    time = date.today()
    ex = LunarSolarConverter.Solar(time.year, time.month, time.day)
    true_time = LunarSolarConverter.LunarSolarConverter.SolarToLunar(ex, ex)

    old = int(true_time.lunarYear) - int(x.year)
    if x.month > true_time.lunarMonth and x.day > true_time.lunarDay:
        old += 1
    else:
        old -= 1
    return abs(old)


def time_chinese(x):
    use = "一 二 三 四 五 六 七 八 九 十".split(" ")
    answer = ""
    if x >= 10:
        y = x // 10 - 1
        answer = use[y]
        x = x % 10


#    print(answer + use[x-1])
    return answer + use[x - 1]


def twelve(x):
    sky = "甲、乙、丙、丁、戊、己、庚、辛、壬、癸".split("、")
    land = "子、丑、寅、卯、辰、巳、午、未、申、酉、戌、亥".split("、")
    x = x - 1911
    the_land = land[(x % 12) - 1]
    the_sky = sky[(x - 2) % 10 - 1]
    return the_sky + the_land


def hour_string(x):
    if x > 23 or x <= 1:
        return "子時"
    elif x > 1 and x <= 3:
        return "丑時"
    elif x > 3 and x <= 5:
        return "寅時"
    elif x > 5 and x <= 7:
        return "卯時"
    elif x > 7 and x <= 9:
        return "辰時"
    elif x > 9 and x <= 11:
        return "巳時"
    elif x > 11 and x <= 13:
        return "午時"
    elif x > 13 and x <= 15:
        return "未時"
    elif x > 15 and x <= 17:
        return "申時"
    elif x > 17 and x <= 19:
        return "酉時"
    elif x > 19 and x <= 21:
        return "戌時"
    elif x > 21 and x <= 23:
        return "亥時"

    return "wait"


def logout(request):
    auth.logout(request)
    return redirect('/')


def login(request):
    form = login_form(request.POST or None)
    if request.method == 'POST':  # 如果是 <login.html> 按登入鈕傳送

        name = request.POST['user_name']  # 取得表單傳送的帳號、密碼
        password = request.POST['password']
        user = auth.authenticate(username=name, password=password)  # 使用者驗證
        message = name
        if user is not None:  # 若驗證成功，以 auth.login(request,user) 登入
            if user.is_active:
                auth.login(request, user)
                return redirect('/')  # 登入成功產生一個 Session，重導到<index.html>
                message = '登入成功!'
            else:
                message = '帳號尚未啟用!'
        else:
            message = '登入失敗!'
            return render(request, "login.html", locals())
    return render(request, "login.html", locals())


def register(request):
    if request.method == 'POST':
        form = UserCreationForm(request.POST)
        if form.is_valid():
            user = form.save()
            register_ok = True
            return render(request, "index.html", locals())
    else:
        form = UserCreationForm()

    return render(request, "register.html", locals())


def validate_people_data(request):
    old_name = request.GET.get("old_name", None)
    new_name = request.GET.get("new_name", None)  # 只判斷有沒有重複名字
    new_birthday = request.GET.get("new_birthday", None)
    new_gender = request.GET.get("new_gender", None)
    home_id = request.GET.get("home_id", None)
    time = request.GET.get("time", None)
    if old_name != new_name:
        data = {"is_taken": False, "error_message": "要更改的名字已經被註冊過了"}
    else:
        if new_birthday == "":
            People_data.objects.filter(home_id=home_id,
                                       name=old_name).update(name=new_name,
                                                             gender=new_gender,
                                                             time=time)
        else:
            People_data.objects.filter(home_id=home_id, name=old_name).update(
                name=new_name,
                birthday=new_birthday,
                gender=new_gender,
                time=time)
        data = {'is_taken': True, "result": "更改成功"}

    return JsonResponse(data)


def validate_date(request):
    find_data = Home.objects.filter(
        home_phone__contains=request.GET.get("find_value", None))
    find_format = []
    for i in range(len(find_data)):

        find_format.append(find_data[i].home_phone + "/" +
                           find_data[i].address + "/" + str(find_data[i].pk))
    date = {"find_format": find_format}
    return JsonResponse(date)


def validate_people_del(request):
    del_name = request.GET.get('del_name', None)
    home_id = request.GET.get('home_id', None)
    People_data.objects.filter(home_id=home_id, name=del_name).delete()
    date = {'is_taken': True, "result": "刪除成功"}
    return JsonResponse(date)


def validate_del(request):
    del_phone = request.GET.get('phone', None)
    the_home_id = Home.objects.get(home_phone=str(del_phone)).id

    Home.objects.filter(home_phone=str(del_phone)).delete()

    People_data.objects.filter(home_id=the_home_id).delete()

    date = {'is_taken': True, "result": "刪除成功"}
    return JsonResponse(date)


def data_save(request):

    data= {"ee":"成功"}
    try:
        get = str(request.GET.get("data", None))
        use_time_name =str( time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        history_data.objects.create(history=get,name=use_time_name)
        data["id"] = history_data.objects.get(name=use_time_name).pk
    except Exception as e:
        data= {"ee":e}
   
    return JsonResponse(data)

def validate_username(request):
    old_phone = request.GET.get('old_phone', None)
    new_phone = request.GET.get('new_phone', None)
    old_address = request.GET.get("old_address", None)
    new_address = request.GET.get("new_address", None)
    try:
        if old_phone != new_phone:
            if Home.objects.filter(home_phone=new_phone).exists():
                data = {"is_taken": False, "error_message": "新的電話已經被註冊過了"}
            else:

                Home.objects.filter(home_phone=str(old_phone)).update(
                    address=str(new_address), home_phone=str(new_phone))
                # update_column.update(address=str(new_address),home_phone=str(new_phone))
                #        update_column.save()
                data = {
                    'is_taken':
                    True,
                    "result":
                    old_phone + " 更改成 " + new_phone + "  與  " + old_address +
                    "更改成" + new_address
                }
        else:
            Home.objects.filter(home_phone=str(old_phone)).update(
                address=str(new_address), home_phone=str(new_phone))
            data = {
                'is_taken':
                True,
                "result":
                "家庭電話為" + old_phone + "的地址由" + old_address + "更改成" +
                new_address
            }
    except Exception as e:
        data = {'is_taken': True, "result": e}

    return JsonResponse(data)


def validate_file_other(request):

    get_file = request.GET.get('file_string', None)
    get_name_id = request.GET.get('name_id', None)
    get_name = request.GET.get('name', None)
    old_name = request.GET.get('old_name', None)

    if old_name != get_name:
        if activity_data.objects.filter(name=get_name).exists():
            data = {"result": "新的代碼重複"}
            return JsonResponse(data)

    if get_file == "":
        activity_data.objects.filter(name=old_name).update(
            name=get_name, table_name=get_name_id)
    else:
        activity_data.objects.filter(name=old_name).update(
            name=get_name, table_name=get_name_id, use_file=get_file)
    data = {"result": "更改成功"}
    return JsonResponse(data)


from django.views.decorators.csrf import csrf_exempt


@csrf_exempt
def validate_file(request):
    if request.method == 'POST':
        the_file = request.FILES.get('file')
        data = {"ok": "檔案更改成功", "result1": the_file.size}
        cc = open(r"C:\Users\asd19\temple_project\files\files\\" +
                  the_file.name, 'wb')  # 伺服器建立上傳同名的檔案
        for line in the_file.chunks():  # 分塊拿上傳資料
            cc.write(line)
        cc.close()
        return JsonResponse(data)


def validate_remove_file(request):
    del_file = request.GET.get('remove_name_id', None)

    activity_data.objects.filter(id=del_file).delete()
    date = {"result": "刪除成功"}
    return JsonResponse(date)


@login_required(login_url='/use_login')
def home_form(request):

    form = homeform(request.POST or None)

    title_one = "家庭電話"
    title_two = "家庭地址"
    get_x = "家庭資料"
    get_all_data = Home.objects.all()  # 表單資料
    load_js = "home"

    if form.is_valid():
        process_string = request.POST['phone'].replace("-", "")
        try:
            if (Home.objects.get(home_phone=process_string)):
                x_bug = "已經有註冊過的家庭電話"
        except:
            Home.objects.create(address=request.POST["address"],
                                home_phone=process_string)
            messages = "已送出"

    context = locals()
    return render(request, "form.html", context)


@login_required(login_url='/use_login')
def activityform(request):
    form = activity_form(request.POST or None)

    title_one = "活動名稱"
    title_two = "活動欄位"
    get_x = "活動資料"
    get_all_data = activity_data.objects.all()  # 表單資料

    if request.method == "POST":
        try:
            if form.is_valid():
                if request.POST["use_table"] != "" and request.POST[
                        "name"] != "":

                    if activity_data.objects.filter(
                            name=request.POST["name"]).exists():
                        x_bug = "此名稱被註冊過了"
                    else:
                        messages = "已送出"  # request.FILES.get("x_file")
                        activity_data.objects.create(
                            name=request.POST["name"],
                            use_file=request.FILES.get("x_file"),
                            table_name=request.POST["use_table"])
                        form = activity_form()
                else:
                    x_bug = "請輸入全部資料"

        except Exception as e:
            x_bug = e

    context = locals()
    return render(request, "word_form.html", context)


@login_required(login_url='/use_login')
def join_activity(request):
    historys = history_data.objects.all().order_by("-id")
    return render(request, "choose.html", locals())


import docx


def process_word(get_word_path, get_home_id):  # 現在只得到香客的名稱，要其他資料 要用查資料庫
    return get_word_path
    all_word = docx.Document()  #儲存所有的檔案
    all_replace_value = {}
    try:

        use_word = MailMerge(get_word_path)

        for i in range(len(get_home_id)):
            get_people_data = People_data.objects.get(name=get_home_id)
            all_replace_value.setdefault(
                str(i + 1) + "_name", str(get_people_data.name))
            all_replace_value.setdefault(
                str(i + 1) + "_birthday",
                str(get_people_data.birthday.strftime('%Y年%m月%d天')))

            today = date.today()
            all_replace_value.setdefault(
                str(i + 1) + "_year",
                str(today.year - get_people_data.birthday.year -
                    ((today.month, today.day) <
                     (get_people_data.birthday.month,
                      get_people_data.birthday.day))))

        use_word.merge(**all_replace_value)

        use_word.write("C:\\Users\\asd19\\Downloads\\OKOK.docx")
        os.system('C:\\Users\\asd19\\Downloads\\OKOK.docx')

    except Exception as e:
        return e

    return "ok"  # all_replace_value


def process_haveno_blank(get_list):

    for i in range(len(get_list)):
        if get_list[i] == "":
            return False

    return True


def home_del(request, pk, people_id):
    People_data.objects.filter(home_id=pk, pk=people_id).delete()
    return HttpResponseRedirect(reverse('home', kwargs={'pk': pk}))


@login_required(login_url='/use_login')
def people_form(request, pk):
    form = peopleform(request.POST or None)
    fix_form = fix_peopleform(None)

    people_all = People_data.objects.filter(home_id=pk)

    x_try = Home.objects.get(pk=pk).home_phone

    title_one = "香客名字"
    title_two = "香客生日"
    title_three = "香客性別"
    get_x = "此家庭香客"

    if request.method == "POST" and request.POST.getlist('name'):

        get_all_name = request.POST.getlist('name')
        get_all_birthday = request.POST.getlist('birthday')
        get_all_gender = request.POST.getlist('gender')
        get_all_time = request.POST.getlist('time')

        if process_haveno_blank(get_all_birthday) and process_haveno_blank(
                get_all_name):
            use_bug = ""
            for i in range(len(get_all_name)):

                x_bug = ""

                if (People_data.objects.filter(home_id=pk).filter(
                        name=get_all_name[i]).count() == 0):
                    People_data.objects.create(name=get_all_name[i].replace(
                        " ", ""),
                                               birthday=get_all_birthday[i],
                                               gender=get_all_gender[i],
                                               time=get_all_time[i],
                                               home_id=pk)
                else:
                    use_bug += get_all_name[i] + " "
            form = peopleform(None)
            if use_bug != "":
                use_bug = "名字重複的名單有:" + use_bug
        else:
            x_bug = "請輸入全部欄位"

    context = locals()
    return render(request, "people_add.html", context)


import uuid


def validate_submit(request):

    try:
        x = {}
        x["z"] = json.loads(request.GET.get("all_data", None))
        for c in x["z"]:
            c["address"] = Home.objects.get(home_phone=c["address"]).address
        if date.today().month >= 10:
            x["year"] = twelve(int(date.today().year) + 1)
        else:
            x["year"] = twelve(date.today().year)
        x["title"] = request.GET.get("title", None)
        print(x["title"])
        if request.GET.get("title", None) == "祈求值年太歲星君解除沖剋文疏":
            tpl = DocxTemplate(
                r"C:\Users\asd19\Downloads\try_git\temple_project\files\files\mode1.docx"
            )
        else:
            tpl = DocxTemplate(
                r"C:\Users\asd19\Downloads\try_git\temple_project\files\files\mode2.docx"
            )

        tpl.render(x)

        comtypes.CoInitialize()  #轉pdf
        word = comtypes.client.CreateObject('Word.Application')

        file_location = os.path.dirname(
            os.path.dirname(os.path.abspath(__file__)))
        find_folder = os.path.join(file_location, "output")
        find_yes_no = os.path.exists(find_folder)

        if not find_yes_no:
            os.makedirs(find_folder)

        while (True):
            random_string = str(uuid.uuid4())
            find_x = os.path.join(find_folder, random_string + ".docx")
            if not os.path.exists(find_x):
                tpl.save(find_x)
                doc = word.Documents.Open(find_x)
                find_y = os.path.join(find_folder,
                                      str(uuid.uuid4()) + "_to_pdf")
                doc.SaveAs(find_y, FileFormat=17)
                os.system(find_y + ".pdf")
                break

        doc.Close()
        word.Quit()
        data = {"result": "已經送出"}

    except Exception as e:
        data = {"result": e}
    return JsonResponse(data)


def index(request):

    return render(request, "index.html", locals())
