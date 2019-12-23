from datetime import date
import datetime
from django.shortcuts import render, HttpResponse, HttpResponseRedirect, render_to_response, redirect
from django.http import JsonResponse
from .forms import homeform, peopleform, find_home, activity_form, choose_form, login_form

from .models import Home, People_data, activity_data

from django.contrib.auth.forms import UserCreationForm
from django.contrib import auth
from django.urls import reverse
from mailmerge import MailMerge
from django.contrib.auth import logout

import os
import json
from django.contrib.auth.decorators import login_required

import os
import comtypes.client

from docxtpl import DocxTemplate
from docx.enum.section import WD_ORIENT
from docx import Document


@login_required
def x_try(request):
    file_location = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    find_folder = os.path.join(file_location, "output")

    x={'bye': [[[{'table_name': '點光明燈者'}, [{'table_data': '陳閔致、曹美雲'}]], [{'table_name': '安奉太歲星君者'}, [{'table_data': '楊逸凡 ── 2017年10月27號18時'}, {'table_data': '劉美惠 ── 2012年01月28號04時'}, {'table_data': '蕭孟勳 ── 2019年01月01號10時'}, {'table_data': '曹美雲 ── 2019年12月09號05時'}]]], [[{'table_name': '安奉太歲星君者'}, [{'table_data': '張珊財 ── 1989年03月15號22時'}, {'table_data': '謝純鑫 ── 2013年04月15號14時'}]], [{'table_name': '安奉財神燈者'}, [{'table_data': '曹志嘉 ── 2011年07月04號20時'}, {'table_data': '陳依光 ── 2007年09月11號06時'}, {'table_data': '陳恭宜 ── 2020年10月27號18時'}]]], [[{'table_name': '點光明燈者'}, [{'table_data': '李雅婷、雅文、陳家銘'}]], [{'table_name': '安奉文昌燈者'}, [{'table_data': '雅文 ── 2020年10月28號21時'}, {'table_data': '李雅婷 ── 2016年12月29號04時'}, {'table_data': '許雅喜 ── 2017年02月18號17時'}]]]], 'title': '各種燈', 'today': '2019-12-23'}
    #{"bye":[[[{"table_name":"點光明燈者"},[{"table_data":"陳閔致、曹美雲、蕭孟勳、劉美惠、柯星雯、楊逸凡、楊雅嵐"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"陳恭宜、陳依光、曹志嘉、謝純鑫"}]]],[[{"table_name":"點光明燈者"},[{"table_data":"許雅喜、李雅婷、雅文"}]]]]}
    try:
        tpl = DocxTemplate(r"C:\Users\asd19\Downloads\廟口案子\祈安植福文疏.docx")
        y = x
        day = str(date.today())
        y["today"] = day        
        d=y["bye"][2][1][0]
        for w in range(len(y["bye"])):
            
            for t in range(len(y["bye"][w])):
                num = 0 
                for m in range(len(y["bye"][w][t][1])):
                    get_string =y["bye"][w][t][1][m]["table_data"]
                    num += len(get_string.split("、"))
                y["bye"][w][t].insert(1,{"people":num})
                                            
      #      y["bye"][w][0].insert(1,{"people":num})

   #     d = y
  #      d=y["bye"][2][1][1]
        tpl.render(y)
        tpl.save(r"C:\Users\asd19\Downloads\廟口案子\tryw.docx")
    except Exception as e:
        x = e
        
    return render(request, "try.html", locals())




def download(request):
    file = open('crm/models.py', 'rb')
    response = HttpResponse(file)
    response['Content-Type'] = 'application/octet-stream' #設定頭資訊，告訴瀏覽器這是個檔案
    response['Content-Disposition'] = 'attachment;filename="models.py"'
    return response

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
    for i in range(the_data.count()):
        get_allname_array.append(the_data[i].name + "|" +
                                 the_data[i].birthday.strftime('%Y年%m月%d號%H時'))

    data = {"reslut": '㊣'.join(get_allname_array)}
    return JsonResponse(data)


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

    if old_name != new_name:
        if People_data.objects.filter(home_id=home_id, name=new_name).exists():
            data = {"is_taken": False, "error_message": "要更改的名字已經被註冊過了"}
        else:
            if new_birthday == "":
                People_data.objects.filter(
                    home_id=home_id, name=old_name).update(name=new_name,
                                                           gender=new_gender)
            else:
                People_data.objects.filter(home_id=home_id,
                                           name=old_name).update(
                                               name=new_name,
                                               birthday=new_birthday,
                                               gender=new_gender)

            data = {'is_taken': True, "result": "更改成功"}
    else:
        if new_birthday == "":
            People_data.objects.filter(home_id=home_id,
                                       name=old_name).update(name=new_name,
                                                             gender=new_gender)
        else:
            People_data.objects.filter(home_id=home_id, name=old_name).update(
                name=new_name, birthday=new_birthday, gender=new_gender)
            data = {'is_taken': True, "result": "更改成功"}

    return JsonResponse(data)


def validate_date(request):
    find_data = Home.objects.filter(
        home_phone__contains=request.GET.get("find_value", None))
    find_format = []
    for i in range(len(find_data)):
        find_format.append(find_data[i].home_phone + "/" +
                           find_data[i].address)
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


def index(request):
    return render(request, "index.html", locals())


@login_required(login_url='/use_login')
def join_activity(request):
    x_form = choose_form(request.POST or None)
    x_max = Home.objects.all().count()
    context = locals()
    return render(request, "join_activity.html", context)


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


@login_required(login_url='/use_login')
def people_form(request):
    form = peopleform(request.POST or None)
    find_home_form = find_home(request.POST or None)

    title_one = "香客名字"
    title_two = "香客生日"
    title_three = "香客性別"
    get_x = "此家庭香客"

    try:
        get = Home.objects.get(home_phone=request.POST["homephone"])
        get_all_data = People_data.objects.filter(home_id=get.id)  # 表單資料
    except:
        pass

    try:
        if request.method == "POST" and request.POST["homephone"] != "":
            try:
                get_find_home = Home.objects.get(
                    home_phone=request.POST["homephone"])

                x_try = request.POST["homephone"]

                request.session['get_home_id'] = get_find_home.id

                request.session['get_home_num'] = request.POST["homephone"]

                one_two = True
            except Exception as e:
                x_bug = "搜尋不到此家庭"

    except:
        pass

    if request.method == "POST" and request.POST.getlist('name'):

        get_all_name = request.POST.getlist('name')
        get_all_birthday = request.POST.getlist('birthday')
        get_all_gender = request.POST.getlist('gender')

        if process_haveno_blank(get_all_birthday) and process_haveno_blank(
                get_all_gender) and process_haveno_blank(get_all_name):

            for i in range(len(get_all_name)):
                try:

                    x_bug = ""

                    if get_all_name[i] != "" and (People_data.objects.filter(
                            home_id=request.session['get_home_id']).filter(
                                name=get_all_name[i]).count() == 0):
                        People_data.objects.create(
                            name=get_all_name[i].replace(" ", ""),
                            birthday=get_all_birthday[i],
                            gender=get_all_gender[i],
                            home_id=request.session['get_home_id'])
                    else:
                        if get_all_name[i] == "":
                            x_bug += "禁止空白\n"
                        else:
                            x_bug += get_all_name[i] + "的名字有重複 \n"

                    if x_bug == "":
                        x_bug = "已經送出"
                    else:
                        if len(get_all_name) >= 2:
                            x_bug += "其他正確資料已經送出"

                except Exception as e:
                    x_bug = ""  # e    # len(get_all_gender)

        else:
            one_two = True
            x_try = request.session['get_home_num']
            x_bug = "請輸入全部欄位"

    context = locals()
    return render(request, "people_add.html", context)


def validate_submit(request):

    use_file = request.GET.get("use_file", None)

    get_activity_ID = ""
    get_activity_ID = activity_data.objects.get(id=use_file)

    # data = {"result": request.GET.get("all_data", None)}
    # return JsonResponse(data)
    try:
        tpl = DocxTemplate(get_activity_ID.use_file)
        #tpl = DocxTemplate(r'C:\Users\asd19\Downloads\one.docx')
        x = request.GET.get("all_data", None)
        x = json.loads(x)
        x["title"] = request.GET.get("title", None)
        day = str(date.today())
        x["today"] = day    
        #data = {"result": str(x)}
        #return JsonResponse(data)

        for w in range(len(x["bye"])):
            for t in range(len(x["bye"][w])):
                num = 0 
                for m in range(len(x["bye"][w][t][1])):
                    get_string =x["bye"][w][t][1][m]["table_data"]
                    num += len(get_string.split("、"))
                x["bye"][w][t].insert(1,{"people":num})
        tpl.render(x)
       


        comtypes.CoInitialize()  #轉pdf
        word = comtypes.client.CreateObject('Word.Application')

        file_location = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        find_folder = os.path.join(file_location, "output")
        find_yes_no = os.path.exists(find_folder)

        if not find_yes_no:
            os.makedirs(find_folder)
        
        find_num = 0
        while(True):
            find_x =  os.path.join(find_folder,"file"+ str(find_num) +".docx") 
            if not os.path.exists(find_x):
                tpl.save(find_x)
                doc = word.Documents.Open(find_x)
                find_y = os.path.join(find_folder,"file"+ str(find_num) +"_to_pdf") 
                doc.SaveAs(find_y,FileFormat=17)
                os.system(find_y +".pdf")
                break
            find_num += 1
          
        
        doc.Close()
        word.Quit()
        data = {"result": "已經送出"}


#        data = {"result": str(x)}
    except Exception as e:
        data = {"result": e}
    return JsonResponse(data)
