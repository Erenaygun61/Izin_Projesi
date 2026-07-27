"""
Django settings for İzin_Projesi.
Personel izin portal?, ana maa? program?n?n veritaban?n? kullan?r.
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
MAAS_PROJESI_DB = Path(os.environ.get(
    "MAAS_PROJESI_DB",
    r"C:\Users\erenaygun\Desktop\yeni_maas_projesi\db.sqlite3",
))

SECRET_KEY = os.environ.get("DJANGO_SECRET_KEY", "django-insecure-izin-portali-local-key")
DEBUG = os.environ.get("DJANGO_DEBUG", "1").strip().lower() in {"1", "true", "yes", "on"}
ALLOWED_HOSTS = [h.strip() for h in os.environ.get("DJANGO_ALLOWED_HOSTS", "localhost,127.0.0.1,testserver").split(",") if h.strip()]

SESSION_COOKIE_NAME = os.environ.get("DJANGO_SESSION_COOKIE_NAME", "izin_sessionid")
CSRF_COOKIE_NAME = os.environ.get("DJANGO_CSRF_COOKIE_NAME", "izin_csrftoken")
CSRF_FAILURE_VIEW = "izin_projesi.views.csrf_failure"

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "portal",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "izin_projesi.urls"
TEMPLATES = [{"BACKEND": "django.template.backends.django.DjangoTemplates", "DIRS": [BASE_DIR / "templates"], "APP_DIRS": True, "OPTIONS": {"context_processors": ["django.template.context_processors.request", "django.contrib.auth.context_processors.auth", "django.contrib.messages.context_processors.messages", "portal.context_processors.aktif_vekaletler"]}}]
WSGI_APPLICATION = "izin_projesi.wsgi.application"
DATABASES = {"default": {"ENGINE": "django.db.backends.sqlite3", "NAME": MAAS_PROJESI_DB}}
AUTH_PASSWORD_VALIDATORS = [{"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"}, {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"}, {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"}, {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"}]
LANGUAGE_CODE = "tr-tr"
TIME_ZONE = "Europe/Istanbul"
USE_I18N = True
USE_TZ = True
STATIC_URL = "static/"
STATICFILES_DIRS = [BASE_DIR / "static"]
LOGIN_URL = "/giris/"
LOGIN_REDIRECT_URL = "/"
LOGOUT_REDIRECT_URL = "/giris/"
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

