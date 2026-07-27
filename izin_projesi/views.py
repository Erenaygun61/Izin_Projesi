from django.contrib import messages
from django.contrib.auth.views import LoginView
from django.shortcuts import redirect
from django.urls import reverse
from django.utils.decorators import method_decorator
from django.views.decorators.cache import never_cache
from django.views.decorators.csrf import ensure_csrf_cookie


@method_decorator([never_cache, ensure_csrf_cookie], name="dispatch")
class IzinLoginView(LoginView):
    """Login sayfas?nda her a??l??ta taze CSRF ?erezi ?retir."""

    pass


def csrf_failure(request, reason=""):
    """Bayat/kar??m?? CSRF tokenlar?nda kullan?c?y? 403 sayfas?na d???rme."""

    messages.warning(
        request,
        "Oturum g?venlik bilginiz yenilendi. L?tfen tekrar giri? yap?n.",
    )
    login_url = reverse("login")
    next_url = request.POST.get("next") or request.GET.get("next") or request.get_full_path()
    if next_url and next_url != login_url:
        return redirect(f"{login_url}?next={next_url}")
    return redirect(login_url)
