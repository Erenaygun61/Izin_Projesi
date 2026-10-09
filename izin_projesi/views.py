from django.contrib import messages
from django.contrib.auth.views import LoginView
from django.http import JsonResponse
from django.shortcuts import redirect
from django.urls import reverse
from django.utils.decorators import method_decorator
from django.views.decorators.cache import never_cache
from django.views.decorators.csrf import ensure_csrf_cookie


@method_decorator([never_cache, ensure_csrf_cookie], name="dispatch")
class IzinLoginView(LoginView):
    """Giriş sayfasında her açılışta taze CSRF çerezi üretir."""

    pass


def csrf_failure(request, reason=""):
    """Bayat veya karışmış CSRF tokenlarında kullanıcıyı 403 sayfasına düşürme."""

    match = request.resolver_match
    if match and (match.url_name or "").startswith("personel_aktarim_"):
        return JsonResponse({"error": "Oturum guvenlik bilginiz yenilendi. Sayfayi yenileyip tekrar deneyin."}, status=403)
    messages.warning(
        request,
        "Oturum güvenlik bilginiz yenilendi. Lütfen tekrar giriş yapın.",
    )
    login_url = reverse("login")
    next_url = request.POST.get("next") or request.GET.get("next") or request.get_full_path()
    if next_url and next_url != login_url:
        return redirect(f"{login_url}?next={next_url}")
    return redirect(login_url)
