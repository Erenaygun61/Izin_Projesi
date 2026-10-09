from django.contrib import admin
from .models import IzinTalebi

@admin.register(IzinTalebi)
class IzinTalebiAdmin(admin.ModelAdmin):
    list_display = ("personel_id", "izin_turu", "baslangic_tarihi", "bitis_tarihi", "kullanilan_gun", "durum", "olusturma_zamani")
    list_filter = ("durum", "izin_turu")
    search_fields = ("personel_id", "aciklama", "adres")
    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False
