from datetime import date
import datetime
from django.shortcuts import render, HttpResponse, HttpResponseRedirect, render_to_response, redirect
from django.http import JsonResponse
from .forms import homeform, peopleform, activity_form, choose_form, login_form,fix_peopleform

from .models import Home, People_data, activity_data

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

class Lunar:
    def __init__(self, lunarYear, lunarMonth, lunarDay, isleap):
        self.isleap = isleap
        self.lunarDay = lunarDay
        self.lunarMonth = lunarMonth
        self.lunarYear = lunarYear

class Solar:
    def __init__(self, solarYear, solarMonth, solarDay):
        self.solarDay = solarDay
        self.solarMonth = solarMonth
        self.solarYear = solarYear
def GetBitInt(data, length, shift):
    return (data & (((1 << length) - 1) << shift)) >> shift
def SolarToInt(y, m, d):
    m = (m + 9) % 12
    y -= m // 10
    return 365 * y + y // 4 - y // 100 + y // 400 + (m * 306 + 5) // 10 + (d - 1)
class LunarSolarConverter:
    lunar_month_days = [1887, 0x1694, 0x16aa, 0x4ad5, 0xab6, 0xc4b7, 0x4ae, 0xa56, 0xb52a,
                        0x1d2a, 0xd54, 0x75aa, 0x156a, 0x1096d, 0x95c, 0x14ae, 0xaa4d, 0x1a4c, 0x1b2a, 0x8d55,
                        0xad4, 0x135a, 0x495d,
                        0x95c, 0xd49b, 0x149a, 0x1a4a, 0xbaa5, 0x16a8, 0x1ad4, 0x52da, 0x12b6, 0xe937, 0x92e,
                        0x1496, 0xb64b, 0xd4a,
                        0xda8, 0x95b5, 0x56c, 0x12ae, 0x492f, 0x92e, 0xcc96, 0x1a94, 0x1d4a, 0xada9, 0xb5a, 0x56c,
                        0x726e, 0x125c,
                        0xf92d, 0x192a, 0x1a94, 0xdb4a, 0x16aa, 0xad4, 0x955b, 0x4ba, 0x125a, 0x592b, 0x152a,
                        0xf695, 0xd94, 0x16aa,
                        0xaab5, 0x9b4, 0x14b6, 0x6a57, 0xa56, 0x1152a, 0x1d2a, 0xd54, 0xd5aa, 0x156a, 0x96c,
                        0x94ae, 0x14ae, 0xa4c,
                        0x7d26, 0x1b2a, 0xeb55, 0xad4, 0x12da, 0xa95d, 0x95a, 0x149a, 0x9a4d, 0x1a4a, 0x11aa5,
                        0x16a8, 0x16d4,
                        0xd2da, 0x12b6, 0x936, 0x9497, 0x1496, 0x1564b, 0xd4a, 0xda8, 0xd5b4, 0x156c, 0x12ae,
                        0xa92f, 0x92e, 0xc96,
                        0x6d4a, 0x1d4a, 0x10d65, 0xb58, 0x156c, 0xb26d, 0x125c, 0x192c, 0x9a95, 0x1a94, 0x1b4a,
                        0x4b55, 0xad4,
                        0xf55b, 0x4ba, 0x125a, 0xb92b, 0x152a, 0x1694, 0x96aa, 0x15aa, 0x12ab5, 0x974, 0x14b6,
                        0xca57, 0xa56, 0x1526,
                        0x8e95, 0xd54, 0x15aa, 0x49b5, 0x96c, 0xd4ae, 0x149c, 0x1a4c, 0xbd26, 0x1aa6, 0xb54,
                        0x6d6a, 0x12da, 0x1695d,
                        0x95a, 0x149a, 0xda4b, 0x1a4a, 0x1aa4, 0xbb54, 0x16b4, 0xada, 0x495b, 0x936, 0xf497,
                        0x1496, 0x154a, 0xb6a5,
                        0xda4, 0x15b4, 0x6ab6, 0x126e, 0x1092f, 0x92e, 0xc96, 0xcd4a, 0x1d4a, 0xd64, 0x956c,
                        0x155c, 0x125c, 0x792e,
                        0x192c, 0xfa95, 0x1a94, 0x1b4a, 0xab55, 0xad4, 0x14da, 0x8a5d, 0xa5a, 0x1152b, 0x152a,
                        0x1694, 0xd6aa,
                        0x15aa, 0xab4, 0x94ba, 0x14b6, 0xa56, 0x7527, 0xd26, 0xee53, 0xd54, 0x15aa, 0xa9b5, 0x96c,
                        0x14ae, 0x8a4e,
                        0x1a4c, 0x11d26, 0x1aa4, 0x1b54, 0xcd6a, 0xada, 0x95c, 0x949d, 0x149a, 0x1a2a, 0x5b25,
                        0x1aa4, 0xfb52,
                        0x16b4, 0xaba, 0xa95b, 0x936, 0x1496, 0x9a4b, 0x154a, 0x136a5, 0xda4, 0x15ac]

    solar_1_1 = [1887, 0xec04c, 0xec23f, 0xec435, 0xec649, 0xec83e, 0xeca51, 0xecc46, 0xece3a,
                 0xed04d, 0xed242, 0xed436, 0xed64a, 0xed83f, 0xeda53, 0xedc48, 0xede3d, 0xee050, 0xee244, 0xee439,
                 0xee64d,
                 0xee842, 0xeea36, 0xeec4a, 0xeee3e, 0xef052, 0xef246, 0xef43a, 0xef64e, 0xef843, 0xefa37, 0xefc4b,
                 0xefe41,
                 0xf0054, 0xf0248, 0xf043c, 0xf0650, 0xf0845, 0xf0a38, 0xf0c4d, 0xf0e42, 0xf1037, 0xf124a, 0xf143e,
                 0xf1651,
                 0xf1846, 0xf1a3a, 0xf1c4e, 0xf1e44, 0xf2038, 0xf224b, 0xf243f, 0xf2653, 0xf2848, 0xf2a3b, 0xf2c4f,
                 0xf2e45,
                 0xf3039, 0xf324d, 0xf3442, 0xf3636, 0xf384a, 0xf3a3d, 0xf3c51, 0xf3e46, 0xf403b, 0xf424e, 0xf4443,
                 0xf4638,
                 0xf484c, 0xf4a3f, 0xf4c52, 0xf4e48, 0xf503c, 0xf524f, 0xf5445, 0xf5639, 0xf584d, 0xf5a42, 0xf5c35,
                 0xf5e49,
                 0xf603e, 0xf6251, 0xf6446, 0xf663b, 0xf684f, 0xf6a43, 0xf6c37, 0xf6e4b, 0xf703f, 0xf7252, 0xf7447,
                 0xf763c,
                 0xf7850, 0xf7a45, 0xf7c39, 0xf7e4d, 0xf8042, 0xf8254, 0xf8449, 0xf863d, 0xf8851, 0xf8a46, 0xf8c3b,
                 0xf8e4f,
                 0xf9044, 0xf9237, 0xf944a, 0xf963f, 0xf9853, 0xf9a47, 0xf9c3c, 0xf9e50, 0xfa045, 0xfa238, 0xfa44c,
                 0xfa641,
                 0xfa836, 0xfaa49, 0xfac3d, 0xfae52, 0xfb047, 0xfb23a, 0xfb44e, 0xfb643, 0xfb837, 0xfba4a, 0xfbc3f,
                 0xfbe53,
                 0xfc048, 0xfc23c, 0xfc450, 0xfc645, 0xfc839, 0xfca4c, 0xfcc41, 0xfce36, 0xfd04a, 0xfd23d, 0xfd451,
                 0xfd646,
                 0xfd83a, 0xfda4d, 0xfdc43, 0xfde37, 0xfe04b, 0xfe23f, 0xfe453, 0xfe648, 0xfe83c, 0xfea4f, 0xfec44,
                 0xfee38,
                 0xff04c, 0xff241, 0xff436, 0xff64a, 0xff83e, 0xffa51, 0xffc46, 0xffe3a, 0x10004e, 0x100242,
                 0x100437,
                 0x10064b, 0x100841, 0x100a53, 0x100c48, 0x100e3c, 0x10104f, 0x101244, 0x101438, 0x10164c,
                 0x101842, 0x101a35,
                 0x101c49, 0x101e3d, 0x102051, 0x102245, 0x10243a, 0x10264e, 0x102843, 0x102a37, 0x102c4b,
                 0x102e3f, 0x103053,
                 0x103247, 0x10343b, 0x10364f, 0x103845, 0x103a38, 0x103c4c, 0x103e42, 0x104036, 0x104249,
                 0x10443d, 0x104651,
                 0x104846, 0x104a3a, 0x104c4e, 0x104e43, 0x105038, 0x10524a, 0x10543e, 0x105652, 0x105847,
                 0x105a3b, 0x105c4f,
                 0x105e45, 0x106039, 0x10624c, 0x106441, 0x106635, 0x106849, 0x106a3d, 0x106c51, 0x106e47,
                 0x10703c, 0x10724f,
                 0x107444, 0x107638, 0x10784c, 0x107a3f, 0x107c53, 0x107e48]

def SolarToLunar(self, solar):

    lunar = Lunar(0, 0, 0, False)
    index = solar.solarYear - LunarSolarConverter.solar_1_1[0]
    data = (solar.solarYear << 9) | (solar.solarMonth << 5) | solar.solarDay
    if LunarSolarConverter.solar_1_1[index] > data:
        index -= 1

    solar11 = LunarSolarConverter.solar_1_1[index]
    y = GetBitInt(solar11, 12, 9)
    m = GetBitInt(solar11, 4, 5)
    d = GetBitInt(solar11, 5, 0)
    offset = SolarToInt(solar.solarYear, solar.solarMonth, solar.solarDay) - SolarToInt(y, m, d)

    days = LunarSolarConverter.lunar_month_days[index]
    leap = GetBitInt(days, 4, 13)

    lunarY = index + LunarSolarConverter.solar_1_1[0]
    lunarM = 1
    offset += 1

    for i in range(0, 13):

        dm = GetBitInt(days, 1, 12 - i) == 1 and 30 or 29
        if offset > dm:

            lunarM += 1
            offset -= dm

        else:

            break

    lunarD = int(offset)
    lunar.lunarYear = lunarY
    lunar.lunarMonth = lunarM
    lunar.isleap = False
    if leap != 0 and lunarM > leap:

        lunar.lunarMonth = lunarM - 1
        if lunarM == leap + 1:
            lunar.isleap = True

    lunar.lunarDay = lunarD
    return lunar

    def __init__(self):
        pass



@login_required
def x_try(request):
    x={}
    x["z"]=[{'people': ['柯星雯 本命 己亥 年 一十 月 二八 號   生行庚 WAIT 歲 ', '楊逸凡 本命 丁酉 年 一十 月 二七 號   生行庚 WAIT 歲 ', '楊雅嵐 本命 戊戌 年 一十 月 二九 號   生行庚 WAIT 歲 '], 'address': '048544412', 'one_people': '楊雅嵐'}, {'people': ['羅弘寧 本命 戊寅 年 一十 月 二二 號   生行庚 WAIT 歲 ', '李冰茜 本命 庚寅 年 一十 月 四 號   生行庚 WAIT 歲 ', ' 蔣原杰 本命 乙酉 年 七 月 二八 號   生行庚 WAIT 歲 '], 'address': '048548888', 'one_people': '蔣原杰'}, {'people': ['許 雅喜 本命 丁酉 年 二 月 一八 號   生行庚 WAIT 歲 ', '雅文 本命 庚子 年 一十 月 二八 號   生行庚 WAIT 歲 ', '陳家銘 本命 癸巳 年 一十 月 二八 號   生行庚 WAIT 歲 '], 'address': '077632214', 'one_people': '陳家銘'}]
    #{"bye":[[[{"table_name":"點光明燈者"},[{"table_data":"陳閔致、曹美雲、蕭孟勳、劉美惠、柯星雯、楊逸凡、楊雅嵐"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"陳恭宜、陳依光、曹志嘉、謝純鑫"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"許雅喜、李雅婷、雅文"}]]]]}
    try:
        for c in x["z"]:
            c["address"] = Home.objects.get(home_phone=c["address"]).address
    
        tpl = DocxTemplate(r"C:\Users\asd19\Downloads\try_git\temple_project\files\files\mode1.docx")

        if date.today().month >=10 :
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

def new(request):
    x_form = choose_form(request.POST or None)
    x_max = Home.objects.all().count()
    return render(request,"join_activity.html", locals())

def csv_add(request):
   
    if(request.method == "POST"):
        homes = request.FILES.get('home')
        people = request.FILES.get('people')

        if homes =="" or people =="":
            error = "請一次輸入兩個檔案"
            return render(request,"up_date.html",locals())

        if homes.name.split(".")[-1] != "csv" and people.name.split(".")[-1] != "csv" :
            error = "請輸入csv檔"
            return render(request,"up_date.html",locals())


        fobj = open(os.path.join(BASE_DIR,"people",homes.name), 'wb')
        for line in homes.chunks():
            fobj.write(line)
        fobj.close()

        fobj = open(os.path.join(BASE_DIR,"people",people.name), 'wb')
        for line in people.chunks():
            fobj.write(line)
        fobj.close()
        
        use_path = os.path.join(BASE_DIR,"people")
        home_path=os.path.join(use_path, homes.name)
        people_path = os.path.join(use_path, people.name)
        with open(home_path, newline='') as csvfile:
            
            homes = csv.reader(csvfile)
               
            one = 0
               
            for row in homes:
                    
                if one != 0:
                    if not Home.objects.filter(home_phone=row[0]).exists():
                        Home.objects.create(home_phone=row[0],address=row[1])                        
                    
                    home_id = Home.objects.get(home_phone=row[0]).pk                        
                    with open(people_path, newline='') as peoplefile:
                        peoples = csv.reader(peoplefile)
                        for people in peoples:                                               
                            if people[0] !="信眾名字" :
                                print("e")                         
                                if row[0] == people[4]:
                                    people[1] = people[1].replace("/", "-")
                                    if people[3] == "男":
                                        people[3]= "male"
                                    else:
                                        people[3]="female"
                                    if not People_data.objects.filter(home_id = home_id,name=people[0]).exists():                                       
                                        People_data.objects.create(name=people[0],birthday= people[1],time=people[2],gender=people[3],home_id=home_id)
                one = 1
                    

    return render(request,"up_date.html",locals())


def home_page(request,pk):

    return render(request,"index.html",locals())


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
        output = the_data[i].name + " 本命 " +twelve(date.year)  + " 年 " + time_chinese(date.month) + " 月 "  + time_chinese(date.day) +" 號 " + "  生行庚 " +year(date) +" 歲 "
        get_allname_array.append(the_data[i].name + "|" + output + "|" + "F")

    data = {"reslut": '㊣'.join(get_allname_array)}
    return JsonResponse(data)


def year(x):
    time = date.today()
    ex = Solar(time.year,time.month,time.day)
    true_time = SolarToLunar(ex,ex)

    old = true_time.lunarYear - x.year
    if x.month > true_time.lunarMonth and x.day >true_time.lunarDay:
        old +=1
    else:
        old -=1
    print(old)
    return str(old)


def time_chinese(x):
    use = "一 二 三 四 五 六 七 八 九 十".split(" ")    
    answer = ""
    if x >=10:            
        y =x//10-1
        answer = use[y]
        x = x % 10   
#    print(answer + use[x-1])
    return answer + use[x-1]

def twelve (x):
    sky="甲、乙、丙、丁、戊、己、庚、辛、壬、癸".split("、")
    land = "子、丑、寅、卯、辰、巳、午、未、申、酉、戌、亥".split("、")
    x = x - 1911
    the_land =land[(x % 12)-1]
    the_sky =sky[(x-2) % 10-1]
    return the_sky + the_land

def hour_string(x):
    if x > 23 or x<=1:
        return "子時"
    elif x > 1 and x<=3:
        return "丑時"
    elif x > 3 and x<=5:
        return "寅時"
    elif x > 5 and x<=7:
        return "卯時"
    elif x > 7 and x<=9:
        return "辰時"
    elif x > 9 and x<=11:
        return "巳時"
    elif x > 11 and x<=13:
        return "午時"
    elif x > 13 and x<=15:
        return "未時"
    elif x > 15 and x<= 17:
        return "申時"
    elif x > 17 and x<= 19:
        return "酉時"
    elif x > 19 and x<=21:
        return "戌時"
    elif x > 21 and x<=23:
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
            People_data.objects.filter(home_id=home_id,name=old_name).update(name=new_name, gender=new_gender,time=time)
        else:
            People_data.objects.filter(home_id=home_id, name=old_name).update(
                name=new_name, birthday=new_birthday, gender=new_gender,time=time)
        data = {'is_taken': True, "result": "更改成功"}

    return JsonResponse(data)


def validate_date(request):
    find_data = Home.objects.filter(
        home_phone__contains=request.GET.get("find_value", None))
    find_format = []
    for i in range(len(find_data)):
       
        find_format.append(find_data[i].home_phone + "/" +
                           find_data[i].address + "/" +str(find_data[i].pk))
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
 
    context = locals()
    return render(request, "choose.html", context)


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


def home_del(request,pk,people_id):
    People_data.objects.filter(home_id=pk,pk=people_id).delete()
    return HttpResponseRedirect(reverse('home',  kwargs={'pk':pk}));

    

@login_required(login_url='/use_login')
def people_form(request,pk):
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

        if process_haveno_blank(get_all_birthday) and process_haveno_blank(get_all_name):
            use_bug = ""
            for i in range(len(get_all_name)):
          
                x_bug = ""

                if (People_data.objects.filter(home_id=pk).filter(name=get_all_name[i]).count() == 0):
                    People_data.objects.create(
                            name=get_all_name[i].replace(" ", ""),
                            birthday=get_all_birthday[i],
                            gender=get_all_gender[i],time = get_all_time[i],
                            home_id=pk)
                else:
                    use_bug += get_all_name[i] + " "
            form = peopleform(None)
            if use_bug !="":
                use_bug = "名字重複的名單有:" + use_bug 
        else:
            x_bug = "請輸入全部欄位"

    context = locals()
    return render(request, "people_add.html", context)

import uuid
def validate_submit(request):



    try:
        tpl = DocxTemplate(r"C:\Users\asd19\Downloads\try_git\temple_project\files\files\mode1.docx")
        x = json.loads(request.GET.get("all_data", None))
       # x["title"] = request.GET.get("title", None)
       # x["year"] = "未處理"
        print(x)
        # for w in range(len(x["bye"])):
        #     for t in range(len(x["bye"][w])):
        #         num = 0 
        #         for m in range(len(x["bye"][w][t][1])):
        #             get_string =x["bye"][w][t][1][m]["table_data"]
        #             num += len(get_string.split("、"))
        #         x["bye"][w][t].insert(1,{"people":num})
        # tpl.render(x)
       


        # comtypes.CoInitialize()  #轉pdf
        # word = comtypes.client.CreateObject('Word.Application')

        # file_location = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # find_folder = os.path.join(file_location, "output")
        # find_yes_no = os.path.exists(find_folder)

        # if not find_yes_no:
        #     os.makedirs(find_folder) 

        # while(True):
        #     random_string =str(uuid.uuid4())
        #     find_x =  os.path.join(find_folder,random_string +".docx") 
        #     if not os.path.exists(find_x):
        #         tpl.save(find_x)
        #         doc = word.Documents.Open(find_x)
        #         find_y = os.path.join(find_folder,  str(uuid.uuid4()) +"_to_pdf")
        #         doc.SaveAs(find_y,FileFormat=17)
        #         os.system(find_y +".pdf")
        #         break
          
        
        # doc.Close()
        # word.Quit()
        data = {"result": "已經送出"}


#        data = {"result": str(x)}
    except Exception as e:
        data = {"result": e}
    return JsonResponse(data)






def index(request):

    return render(request, "index.html", locals())
