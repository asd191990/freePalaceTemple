from django import forms
from django.forms import widgets
from .models import Home, People_data, activity_data

from django.utils.translation import gettext_lazy as _
from django.db import models
# from django.forms import ModelChoiceField
from django.core import validators


class login_form(forms.Form):

    user_name = forms.CharField(label='請輸入帳號', )
    password = forms.CharField(label='請輸入密碼', widget=forms.PasswordInput)


class choose_form(forms.Form):
    activity_ID = forms.ModelChoiceField(label="請選擇活動名稱",
                                         initial=1,
                                         queryset=activity_data.objects.all())


class activity_form(forms.Form):
    name = forms.CharField(label="活動名稱", required=False)
    x_file = forms.FileField(label="檔案上傳", required=False)
    use_table = forms.CharField(
        label="請輸入要使用的欄位格式,欄位格式之間已、符號分隔。欄位格式為 欄位名稱_是否連接出生年月日。 範:光明燈_F ",
        required=False)
    # required=False

    # class Meta:
    #     model = activity_data
    #     fields = '__all__'


class homeform(forms.Form):
    address = forms.CharField(label="地址")
    phone = forms.CharField(label="家庭電話",
                            validators=[
                                validators.RegexValidator(
                                    "(^\d{2}-?\d{3}-?\d{4}$)|(^\d{10}$)",
                                    message='請輸入正確格式的電話號碼！')
                            ])


class peopleform(forms.Form):
    name = forms.CharField(required=False, label="輸入香客名稱", max_length=20)
    birthday_y = forms.CharField(label="年",
                                 widget=forms.TextInput(attrs={"size": "1mv"}))
    birthday_m = forms.CharField(label="月",
                                 widget=forms.TextInput(attrs={"size": "1mv"}))
    birthday_d = forms.CharField(label="日",
                                 widget=forms.TextInput(attrs={"size": "1mv"}))

    time = forms.ChoiceField(
        label='時辰',
        required=False,
        choices=(('子', '子'), ('丑', '丑'), ('寅', '寅'), ('卯', '卯'), ('辰', '辰'),
                 ('巳', '巳'), ('午', '午'), ('未', '未'), ('申', '申'), ('酉', '酉'),
                 ('戌', '戌'), ('亥', '亥'), ('吉', '吉')),
        initial="子",
        widget=forms.widgets.Select())
    animal = forms.ChoiceField(
        label='生肖',
        required=False,
        choices=(('鼠', '鼠'), ('牛', '牛'), ('虎', '虎'), ('兔', '兔'), ('龍', '龍'),
                 ('蛇', '蛇'), ('馬', '馬'), ('羊', '羊'), ('猴', '猴'), ('雞', '雞'),
                 ('狗', '狗'), ('豬', '豬'), ('吉', '吉')),
        initial="子",
        widget=forms.widgets.Select())

    gender = forms.ChoiceField(label='性別',
                               required=False,
                               choices=(('male', '男'), ('female', '女')),
                               initial="男",
                               widget=forms.widgets.Select())


class fix_peopleform(forms.Form):
    x_name = forms.CharField(required=False, label="輸入香客名稱", max_length=20)
    x_birthday = forms.CharField(required=False, label="輸入民國年", max_length=20)
    x_month = forms.CharField(required=False, label="輸入月", max_length=20)
    x_day = forms.CharField(required=False, label="輸入日", max_length=20)
    x_time = forms.ChoiceField(label='時辰',
                               required=False,
                               choices=(('子', '子'), ('丑', '丑'), ('寅', '寅'),
                                        ('卯', '卯'), ('辰', '辰'), ('巳', '巳'),
                                        ('午', '午'), ('未', '未'), ('申', '申'),
                                        ('酉', '酉'), ('戌', '戌'), ('亥', '亥')),
                               initial="子",
                               widget=forms.widgets.Select())
    x_gender = forms.ChoiceField(label='性別',
                                 required=False,
                                 choices=(('male', '男'), ('female', '女')),
                                 initial="男",
                                 widget=forms.widgets.Select())

    # homephone = forms.CharField(
    #     label="輸入家庭電話"
    # )


# queryset=Home.objects.all().values_list(
#            'id', 'home_phone')
