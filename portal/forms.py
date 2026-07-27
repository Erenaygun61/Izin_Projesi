from decimal import Decimal

from django import forms
from django.utils import timezone
from .models import IzinTalebi, IzinTuru


class DateInput(forms.DateInput):
    input_type = "date"


class TcSorguForm(forms.Form):
    tc_no = forms.CharField(
        label="T.C. Kimlik No",
        min_length=11,
        max_length=11,
        widget=forms.TextInput(attrs={"class": "tc-input", "placeholder": "11 haneli T.C. Kimlik No", "inputmode": "numeric"}),
    )

    def clean_tc_no(self):
        tc = self.cleaned_data["tc_no"].strip()
        if not tc.isdigit():
            raise forms.ValidationError("T.C. Kimlik No sadece rakamlardan olu\u015fmal\u0131d\u0131r.")
        return tc


def izin_sebebi_metni(izin_turu_ad):
    ad = (izin_turu_ad or "").lower()
    if "y\u0131ll\u0131k" in ad:
        return "Y\u0131ll\u0131k izin kullanaca\u011f\u0131mdan dolay\u0131"
    if "mazeret" in ad:
        return "Mazeret izni kullanaca\u011f\u0131mdan dolay\u0131"
    if "rapor" in ad:
        return "Rapor izni kullanaca\u011f\u0131mdan dolay\u0131"
    if "refakat" in ad:
        return "Refakat izni kullanaca\u011f\u0131mdan dolay\u0131"
    if "sendika" in ad:
        return "Sendika izni kullanaca\u011f\u0131mdan dolay\u0131"
    if "\u00fccretsiz" in ad:
        return "\u00dccretsiz izin kullanaca\u011f\u0131mdan dolay\u0131"
    if "yol" in ad:
        return "Yol izni kullanaca\u011f\u0131mdan dolay\u0131"
    return "\u0130zin kullanaca\u011f\u0131mdan dolay\u0131"


class IzinEvrakForm(forms.ModelForm):
    evrakin_birimi = forms.CharField(label="Evrak\u0131n Birimi", required=False, initial="\u0130nsan Kaynaklar\u0131 \u015eube M\u00fcd\u00fcrl\u00fc\u011f\u00fc")
    dosya_kodu = forms.CharField(label="Dosya Kodu", required=False, initial="903.05.01")
    konu = forms.CharField(label="Konu", required=False, initial="\u0130zin \u0130ste\u011fi")
    tc_no = forms.CharField(label="T.C. Kimlik No", required=False, disabled=True)
    calisan_tipi = forms.CharField(label="\u00c7al\u0131\u015fan Tipi", required=False, disabled=True)
    sicil_no = forms.CharField(label="Sicil No", required=False, disabled=True)
    cep_telefonu = forms.CharField(label="Cep Telefonu", required=False)

    class Meta:
        model = IzinTalebi
        fields = [
            "evrak_no", "cep_telefonu", "izin_turu",
            "ait_oldugu_yil", "izin_hakki", "kullanilan_gun", "kalan_izin",
            "baslangic_tarihi", "bitis_tarihi", "goreve_baslama_tarihi",
        ]
        labels = {
            "evrak_no": "\u0130zin Evrak No",
            "izin_turu": "\u0130zin T\u00fcr\u00fc (*)",
            "ait_oldugu_yil": "Ait Oldu\u011fu Y\u0131l",
            "izin_hakki": "\u0130zin Hakk\u0131 (G\u00fcn)",
            "kullanilan_gun": "\u0130zin S\u00fcresi (G\u00fcn) (*)",
            "kalan_izin": "Kalan \u0130zin (G\u00fcn)",
            "baslangic_tarihi": "\u0130zin Ba\u015flang\u0131\u00e7 Tarihi (*)",
            "bitis_tarihi": "\u0130zin Biti\u015f Tarihi (*)",
        }
        widgets = {
            "baslangic_tarihi": DateInput(),
            "bitis_tarihi": DateInput(),
            "goreve_baslama_tarihi": DateInput(),
        }

    def __init__(self, *args, personel=None, haklar=None, calisan_tipi="Memur", imzacilar=None, next_evrak_no="", toplam_izin=None, taslak=False, **kwargs):
        super().__init__(*args, **kwargs)
        self.taslak = taslak
        self.fields["goreve_baslama_tarihi"].label = "G\u00f6reve Ba\u015flama Tarihi (*)"
        self.fields["izin_turu"].queryset = IzinTuru.objects.filter(aktif_mi=True).order_by("ad")
        yillik = self.fields["izin_turu"].queryset.filter(yillik_izin_mi=True).first()
        if yillik and not self.initial.get("izin_turu"):
            self.fields["izin_turu"].initial = yillik.id
        haklar = list(haklar or [])
        yil_choices = []
        kalanli_yillar = [h for h in haklar if getattr(h, "kalan", 0) and h.kalan > 0]
        for h in haklar:
            yil_choices.append((str(h.izin_yili), str(h.izin_yili)))
        yillar = {str(h.izin_yili) for h in haklar}
        if "2025" in yillar and "2026" in yillar:
            yil_choices.insert(0, ("2025-2026", "2025-2026"))
        self.fields["ait_oldugu_yil"].choices = yil_choices
        if yil_choices:
            default_hak = kalanli_yillar[-1] if kalanli_yillar else haklar[0]
            self.fields["ait_oldugu_yil"].initial = str(default_hak.izin_yili)
            if "2025" in yillar and "2026" in yillar:
                self.fields["ait_oldugu_yil"].initial = "2025-2026"
        if personel:
            self.fields["tc_no"].initial = personel.tc_no
            self.fields["sicil_no"].initial = personel.sicil_no
            self.fields["cep_telefonu"].initial = getattr(self.instance, "cep_telefonu", "") or personel.cep_tel
            self.fields["calisan_tipi"].initial = calisan_tipi
        if not self.instance.pk:
            self.fields["evrak_no"].initial = next_evrak_no
        sistem_alanlari = {"evrakin_birimi", "evrak_no", "dosya_kodu", "konu", "tc_no", "calisan_tipi", "sicil_no"}
        for name in sistem_alanlari:
            if name in self.fields:
                self.fields[name].disabled = True
                self.fields[name].widget.attrs["data_system"] = "1"
        # ModelForm, taslak evrakta kaydedilmemiş (None) alanı kendi
        # ``initial`` sözlüğünde tutar. Yalnızca field.initial vermek bu
        # durumda ekrana boş değer basılmasına yol açıyordu. Kaynak izin
        # bakiyesini boş evraklara form initial'ı üzerinden de aktar.
        if toplam_izin is not None and self.initial.get("izin_hakki") in (None, ""):
            self.initial["izin_hakki"] = toplam_izin
            self.fields["izin_hakki"].initial = toplam_izin
        self.fields["kalan_izin"].initial = ""
        ordered = [
            "evrakin_birimi", "evrak_no", "dosya_kodu", "konu",
            "tc_no", "calisan_tipi", "sicil_no", "cep_telefonu", "izin_turu",
            "ait_oldugu_yil", "izin_hakki", "kullanilan_gun", "kalan_izin", "baslangic_tarihi", "bitis_tarihi", "goreve_baslama_tarihi",
        ]
        self.order_fields(ordered)
        for name, field in self.fields.items():
            field.widget.attrs.setdefault("class", "form-control")
            if name in {"izin_hakki", "kullanilan_gun", "kalan_izin"}:
                field.widget.attrs["data-balance"] = name
        self.fields["kalan_izin"].widget.attrs["readonly"] = "readonly"
        self.fields["kalan_izin"].widget.attrs["tabindex"] = "-1"
        for name in {"izin_turu", "ait_oldugu_yil", "kullanilan_gun", "baslangic_tarihi", "bitis_tarihi", "goreve_baslama_tarihi"}:
            self.fields[name].required = not taslak

    def clean(self):
        cleaned = super().clean()
        baslangic = cleaned.get("baslangic_tarihi")
        bitis = cleaned.get("bitis_tarihi")
        gun = cleaned.get("kullanilan_gun")
        hak = cleaned.get("izin_hakki")
        if baslangic and bitis and bitis < baslangic:
            self.add_error("bitis_tarihi", "Biti\u015f tarihi ba\u015flang\u0131\u00e7 tarihinden \u00f6nce olamaz.")
        if gun is not None and gun <= 0:
            self.add_error("kullanilan_gun", "\u0130zin s\u00fcresi 0'dan b\u00fcy\u00fck olmal\u0131d\u0131r.")
        if hak is not None and hak < 0:
            self.add_error("izin_hakki", "\u0130zin hakk\u0131 negatif olamaz.")
        if hak is not None and gun is not None:
            if gun > hak:
                self.add_error("kullanilan_gun", "\u0130zin s\u00fcresi izin hakk\u0131ndan fazla olamaz.")
            cleaned["kalan_izin"] = max(Decimal("0.00"), hak - gun)
        return cleaned


IzinTalebiForm = IzinEvrakForm

