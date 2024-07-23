# complaint/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.redirect_to_home),  # Redirect root URL to /home/
    path('home/', views.home_page, name='home'),
    path('superuser/signup/', views.signup_view, name='signup'),
    path('user/login/', views.login_view, name='login'),  
    path('superuser/login/',views.login1_view,name="login1"),
    path('superuser/services/', views.admin_page, name='admin'),
    path('user/services/', views.user_page, name='user'),
    path('user/complaint/', views.complaint_form, name='complaint_form'),
    path('user/DGR/', views.generate_word, name='generate_word'),
]
