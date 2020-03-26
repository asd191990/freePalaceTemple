from django.db import models
#from django import timezone
# Create your models here.


class Day(models.Model):
    date_name = models.CharField(max_length=12)
    class Meta:
        verbose_name_plural = '法會日期'
    def __str__(self):
        return self.date_name

class every_day(models.Model):
    Day_date = models.ForeignKey(
		Day,
		related_name='name',
		null=True,
		blank=True,
		on_delete=models.CASCADE
    )
    date  = models.CharField(max_length=12)
    one_lights = models.TextField(blank=True)
    two_lights = models.TextField(blank=True)
    three_lights = models.TextField(blank=True)
    four_lights = models.TextField(blank=True)
    five_lights = models.TextField(blank=True)
    class Meta:
        verbose_name_plural = '燈的紀錄'
    def __str__(self):
        return self.date




class activity_data(models.Model):
    name = models.CharField(verbose_name="活動名稱", max_length=50)
    use_file = models.FileField(verbose_name="上傳檔案", upload_to="files")
    table_name= models.TextField(verbose_name="請輸入要使用的欄位名稱,並已、符號分隔")


    class Meta:
        verbose_name_plural = '活動資料庫'
    def __str__(self):
        return self.name

class history_data(models.Model):
    history = models.TextField(max_length =1000)
    name = models.CharField(max_length = 20)
    class Meta:
        verbose_name_plural = '歷史紀錄'
    def __str__(self):
        return self.name



class Home(models.Model):
    address = models.CharField(verbose_name='地址', max_length=100)
    home_phone = models.CharField(verbose_name='家庭電話', max_length=100)

    class Meta:
        verbose_name_plural = '家庭資料庫'


class People_data(models.Model):

    name = models.CharField(verbose_name='輸入姓名', max_length=10)

    birthday = models.CharField(
        verbose_name='西元生日',max_length=20)  # default = timezone.now
    time = models.CharField(max_length=5)
    gender = models.CharField(verbose_name='性別', max_length=32, choices=(
        ('male', '男'), ('female', '女')), default="男")

    home_id = models.CharField(max_length=10)

    class Meta:
        verbose_name_plural = '香客資料庫'
