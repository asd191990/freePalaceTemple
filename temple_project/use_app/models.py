from django.db import models
#from django import timezone
# Create your models here.


class activity_data(models.Model):
    name = models.CharField(verbose_name="活動名稱", max_length=50)
    use_file = models.FileField(verbose_name="上傳檔案", upload_to="files")
    table_name= models.TextField(verbose_name="請輸入要使用的欄位名稱,並已、符號分隔")
    

    class Meta:
        verbose_name_plural = '活動資料庫'
    def __str__(self):
        return self.name


class Home(models.Model):
    address = models.CharField(verbose_name='地址', max_length=100)
    home_phone = models.CharField(verbose_name='家庭電話', max_length=100)

    class Meta:
        verbose_name_plural = '家庭資料庫'


class People_data(models.Model):

    name = models.CharField(verbose_name='輸入姓名', max_length=10)

    birthday = models.DateTimeField(
        verbose_name='西元生日(ex:2012-01-01)')  # default = timezone.now

    gender = models.CharField(verbose_name='性別', max_length=32, choices=(
        ('male', '男'), ('female', '女')), default="男")

    home_id = models.CharField(max_length=10)

    class Meta:
        verbose_name_plural = '香客資料庫'
