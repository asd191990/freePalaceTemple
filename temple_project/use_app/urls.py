"""temple_project URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, re_path


from . import views
urlpatterns = [
    path('', views.index),
    path('people_register', views.people_form),
    path('activity_register', views.activityform),
    path('home_register', views.home_form),
    path('join_activity', views.join_activity),
    path("register",views.register),
    path("use_login",views.login),
    path("logout",views.logout),
    path('try',views.x_try), 
    path('ajax/validate_username', views.validate_username, name='validate_username'),
    path('ajax/validate_del', views.validate_del, name='validate_del'),
    path('ajax/validate_date', views.validate_date, name='validate_date'),
    path('ajax/validate_people_data', views.validate_people_data, name='validate_people_data'),
    path('ajax/validate_people_del', views.validate_people_del, name='validate_people_del'),
    path('ajax/validate_get_Home', views.validate_get_Home, name='validate_get_Home'),
    path('ajax/validate_get_people', views.validate_get_Home, name='validate_get_Home'),
    path('ajax/validate_people_all_date', views.validate_people_all_date, name='validate_people_all_date'),
    path('ajax/validate_submit', views.validate_submit, name='validate_submit'),
    path('ajax/validate_file', views.validate_file, name='validate_file'),
    path('ajax/validate_remove_file', views.validate_remove_file, name='validate_remove_file'),    
    path('ajax/validate_get_table', views.validate_get_table, name='validate_get_table'),
    path('ajax/validate_file_other', views.validate_file_other, name='validate_file_other'),
    
    re_path(r".",views.index)    
]
