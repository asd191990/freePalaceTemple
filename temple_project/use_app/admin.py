from django.contrib import admin

from .models import Home, People_data, activity_data,history_data

# Register your models here.

admin.site.site_header = '後臺管理系統'
admin.site.site_title = '後臺管理'
admin.site.index_title = '鄉廟資料庫 管理'


class set_history(admin.ModelAdmin):
    list_display = [field.name for field in history_data._meta.fields]

    class Meat:
        ordering = ['order_date']
admin.site.register(history_data,set_history)
class set_home(admin.ModelAdmin):
    list_display = [field.name for field in Home._meta.fields]

    class Meat:
        ordering = ['order_date']


class set_people_data(admin.ModelAdmin):
    list_display = [field.name for field in People_data._meta.fields]


class set_activity(admin.ModelAdmin):
    list_display = [field.name for field in activity_data._meta.fields]

    class Meat:
        ordering = ['order_date']


admin.site.register(Home, set_home)

admin.site.register(activity_data, set_activity)

admin.site.register(People_data, set_people_data)
