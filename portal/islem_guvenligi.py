from functools import wraps

from django.contrib import messages
from django.db import OperationalError
from django.http import JsonResponse
from django.shortcuts import redirect


def veritabani_yogunlugunu_yakala(view):
    """Catch contention only after the atomic mutation has rolled back."""
    @wraps(view)
    def wrapped(request, *args, **kwargs):
        try:
            return view(request, *args, **kwargs)
        except OperationalError as error:
            if not any(token in str(error).lower() for token in ("database is locked", "database table is locked", "database is busy")):
                raise
            message = "Başka bir işlem devam ediyor. Değişiklik uygulanmadı; güncel ekranı açıp tekrar deneyin."
            if 'application/json' in request.headers.get('Accept', ''):
                return JsonResponse({'error': message}, status=409)
            messages.warning(request, message)
            return redirect('dashboard')
    return wrapped
