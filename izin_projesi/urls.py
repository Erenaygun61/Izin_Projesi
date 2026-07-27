from django.contrib import admin
from django.contrib.auth import views as auth_views
from .views import IzinLoginView
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    path("giris/", IzinLoginView.as_view(template_name="registration/login.html"), name="login"),
    path("cikis/", auth_views.LogoutView.as_view(), name="logout"),
    path("", include("portal.urls")),
]
