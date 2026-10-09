from decimal import Decimal

import json

from django import forms
from django.db.models import Q
from django.utils import timezone
from .models import IzinTalebi, IzinTuru


class DateInput(forms.DateInput):
    input_type = "date"

    def __init__(self, attrs=None, format=None):
        # Native date controls require ISO values, regardless of locale.
        super().__init__(attrs=attrs, format=format or "%Y-%m-%d")


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


def izin_gunu_form_degeri(value):
    if value in (None, ""):
        return ""
    try:
        number = Decimal(str(value))
    except Exception:
        return value
    return format(number.normalize(), "f")


class IzinEvrakForm(forms.ModelForm):
    takip_no = forms.CharField(label="Takip No", required=False, disabled=True)
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

    def __init__(self, *args, personel=None, haklar=None, tur_haklari=None, calisan_tipi="Memur", imzacilar=None, next_evrak_no="", toplam_izin=None, taslak=False, **kwargs):
        super().__init__(*args, **kwargs)
        mevcut_evrak = bool(self.instance and self.instance.pk)
        bos_taslak = bool(
            mevcut_evrak
            and getattr(self.instance, "durum", "") == "TASLAK"
            and not self.instance.izin_turu_id
            and self.instance.izin_hakki in (None, "")
            and self.instance.kullanilan_gun in (None, "")
        )
        self.taslak = taslak
        self.fields["goreve_baslama_tarihi"].label = "G\u00f6reve Ba\u015flama Tarihi (*)"
        izin_turu_filtre = Q(aktif_mi=True)
        if mevcut_evrak and self.instance.izin_turu_id:
            izin_turu_filtre |= Q(pk=self.instance.izin_turu_id)
        self.fields["izin_turu"].queryset = IzinTuru.objects.filter(izin_turu_filtre).order_by("ad")
        yillik = self.fields["izin_turu"].queryset.filter(yillik_izin_mi=True).first()
        if not mevcut_evrak and yillik and not self.initial.get("izin_turu"):
            self.fields["izin_turu"].initial = yillik.id
        # Yıllık izin bakiyeleri eski yıldan yeni yıla kullanılır.
        haklar = sorted(list(haklar or []), key=lambda hak: (hak.izin_yili, hak.pk or 0))
        self._yillik_haklar = haklar
        self._tur_haklari = {str(key): sorted(value, key=lambda hak: (hak.izin_yili, hak.pk or 0)) for key, value in (tur_haklari or {}).items()}
        kalanli_yillar = [
            hak for hak in haklar
            if Decimal(getattr(hak, "kalan", 0) or 0) > Decimal("0")
        ]
        yil_choices = [(str(timezone.localdate().year), str(timezone.localdate().year))]
        gorulen_secimler = {yil_choices[0][0]}
        # Tek yıllar ile eski yıldan başlayan bütün olası kapsamlara yer
        # verilir; değer kullanıcı tarafından değil otomatik atanır.
        for index, hak in enumerate(kalanli_yillar):
            tek_yil = str(hak.izin_yili)
            if tek_yil not in gorulen_secimler:
                yil_choices.append((tek_yil, tek_yil))
                gorulen_secimler.add(tek_yil)
            kapsam = "-".join(str(item.izin_yili) for item in kalanli_yillar[:index + 1])
            if kapsam not in gorulen_secimler:
                yil_choices.append((kapsam, kapsam))
                gorulen_secimler.add(kapsam)
        kayitli_yil = str(self.instance.ait_oldugu_yil or "").strip() if mevcut_evrak else ""
        for tur_haklar in self._tur_haklari.values():
            for index, hak in enumerate(tur_haklar):
                for yil in [str(hak.izin_yili), "-".join(dict.fromkeys(str(item.izin_yili) for item in tur_haklar[:index + 1]))]:
                    if yil not in gorulen_secimler:
                        yil_choices.append((yil, yil))
                        gorulen_secimler.add(yil)
        if kayitli_yil and kayitli_yil not in {str(value) for value, _label in yil_choices}:
            yil_choices.insert(0, (kayitli_yil, kayitli_yil))
        self.fields["ait_oldugu_yil"].choices = yil_choices
        if not mevcut_evrak or bos_taslak:
            # İlk açılışta en eski kullanılabilir yıllık izin hakkı gelir.
            if kalanli_yillar:
                ilk_hak = kalanli_yillar[0]
                toplam_hak = Decimal(toplam_izin or 0)
                if toplam_hak <= 0:
                    toplam_hak = sum((Decimal(hak.kalan or 0) for hak in kalanli_yillar), Decimal("0"))
                self.initial["ait_oldugu_yil"] = str(ilk_hak.izin_yili)
                self.fields["ait_oldugu_yil"].initial = str(ilk_hak.izin_yili)
                self.initial["izin_hakki"] = izin_gunu_form_degeri(toplam_hak)
                self.fields["izin_hakki"].initial = izin_gunu_form_degeri(toplam_hak)
                if yillik and not self.initial.get("izin_turu"):
                    self.initial["izin_turu"] = yillik.id
                    self.fields["izin_turu"].initial = yillik.id
            else:
                self.initial["ait_oldugu_yil"] = str(timezone.localdate().year)
                self.fields["ait_oldugu_yil"].initial = str(timezone.localdate().year)
        if personel:
            self.fields["tc_no"].initial = personel.tc_no
            self.fields["sicil_no"].initial = personel.sicil_no
            if not mevcut_evrak:
                self.fields["cep_telefonu"].initial = personel.cep_tel
            self.fields["calisan_tipi"].initial = calisan_tipi
        if not self.instance.pk:
            self.fields["evrak_no"].initial = next_evrak_no
        self.fields["takip_no"].initial = self.instance.takip_no
        self.fields["evrak_no"].label = "Resmi Evrak No"
        self.fields["evrak_no"].widget.attrs["placeholder"] = "Son imzada verilecek"
        sistem_alanlari = {"evrakin_birimi", "takip_no", "evrak_no", "dosya_kodu", "konu", "tc_no", "calisan_tipi", "sicil_no"}
        for name in sistem_alanlari:
            if name in self.fields:
                self.fields[name].disabled = True
                self.fields[name].widget.attrs["data_system"] = "1"
        if not mevcut_evrak or bos_taslak:
            self.fields["kalan_izin"].initial = ""
        else:
            # Reddedilen/iade edilen evrak yeniden açıldığında bütün girilmiş
            # değerler korunur; kullanıcı yalnızca isterse bunları değiştirir.
            kayitli_degerler = {
                "evrak_no": self.instance.evrak_no,
                "cep_telefonu": self.instance.cep_telefonu,
                "izin_turu": self.instance.izin_turu_id,
                "ait_oldugu_yil": self.instance.ait_oldugu_yil,
                "izin_hakki": izin_gunu_form_degeri(self.instance.izin_hakki),
                "kullanilan_gun": izin_gunu_form_degeri(self.instance.kullanilan_gun),
                "kalan_izin": izin_gunu_form_degeri(self.instance.kalan_izin),
                "baslangic_tarihi": self.instance.baslangic_tarihi,
                "bitis_tarihi": self.instance.bitis_tarihi,
                "goreve_baslama_tarihi": self.instance.goreve_baslama_tarihi,
            }
            for name, value in kayitli_degerler.items():
                if value is not None:
                    self.initial[name] = value
                    self.fields[name].initial = value
        ordered = [
            "evrakin_birimi", "takip_no", "evrak_no", "dosya_kodu", "konu",
            "tc_no", "calisan_tipi", "sicil_no", "cep_telefonu", "izin_turu",
            "ait_oldugu_yil", "izin_hakki", "kullanilan_gun", "kalan_izin", "baslangic_tarihi", "bitis_tarihi", "goreve_baslama_tarihi",
        ]
        self.order_fields(ordered)
        for name, field in self.fields.items():
            field.widget.attrs.setdefault("class", "form-control")
            if name in {"izin_hakki", "kullanilan_gun", "kalan_izin"}:
                field.widget.attrs["data-balance"] = name
        self.fields["izin_turu"].widget.attrs["data-yillik-izin-idleri"] = json.dumps([
            str(izin.pk)
            for izin in self.fields["izin_turu"].queryset.filter(yillik_izin_mi=True)
        ])
        self.fields["izin_hakki"].widget.attrs["data-yillik-haklar"] = json.dumps([
            {"yil": str(hak.izin_yili), "kalan": str(hak.kalan or 0)}
            for hak in kalanli_yillar
        ])
        self.fields["izin_hakki"].widget.attrs["data-guncel-yil"] = str(timezone.localdate().year)
        self.tur_bakiyeleri = {
            str(tur.pk): [{"yil": str(hak.izin_yili), "kalan": str(max(Decimal("0"), Decimal(hak.kalan or 0)))} for hak in (haklar if tur.yillik_izin_mi else self._tur_haklari.get(str(tur.pk), []))]
            for tur in self.fields["izin_turu"].queryset
        }
        self.fields["izin_hakki"].widget.attrs["data-otomatik-yeni"] = "1" if (not mevcut_evrak or bos_taslak) else "0"
        self.fields["ait_oldugu_yil"].widget.attrs["class"] += " otomatik-yil"
        self.fields["ait_oldugu_yil"].widget.attrs["tabindex"] = "-1"
        self.fields["ait_oldugu_yil"].widget.attrs["aria-readonly"] = "true"
        self.fields["ait_oldugu_yil"].widget.attrs["readonly"] = "readonly"
        self.fields["kalan_izin"].widget.attrs["readonly"] = "readonly"
        self.fields["kalan_izin"].widget.attrs["tabindex"] = "-1"
        for name in {"izin_turu", "kullanilan_gun", "baslangic_tarihi", "bitis_tarihi", "goreve_baslama_tarihi"}:
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
        izin_turu = cleaned.get("izin_turu")
        secili_haklar = self._yillik_haklar if izin_turu and izin_turu.yillik_izin_mi else self._tur_haklari.get(str(getattr(izin_turu, "pk", "")), [])
        if izin_turu and izin_turu.yillik_izin_mi and not self._yillik_haklar:
            self.add_error(
                "izin_hakki",
                "İzin hakkınız bulunmadı. Yanlışlık olduğunu düşünüyorsanız lütfen İK ile iletişime geçiniz.",
            )
        if izin_turu and secili_haklar and gun is not None:
            kalan_gun = Decimal(gun)
            kapsanan_haklar = []
            for izin_hakki in secili_haklar:
                bakiye = Decimal(getattr(izin_hakki, "kalan", 0) or 0)
                if bakiye <= 0:
                    continue
                kapsanan_haklar.append(izin_hakki)
                kalan_gun -= bakiye
                if kalan_gun <= 0:
                    break
            if kalan_gun > 0:
                self.add_error(
                    "kullanilan_gun",
                    "Seçilen izin türü için yeterli bakiye bulunamadı. Lütfen İK ile iletişime geçiniz.",
                )
            elif kapsanan_haklar:
                toplam_hak = sum((max(Decimal("0"), Decimal(izin_hakki.kalan or 0)) for izin_hakki in secili_haklar), Decimal("0"))
                cleaned["ait_oldugu_yil"] = "-".join(str(izin_hakki.izin_yili) for izin_hakki in kapsanan_haklar)
                cleaned["izin_hakki"] = toplam_hak
                cleaned["kalan_izin"] = max(Decimal("0.00"), toplam_hak - Decimal(gun))
                hak = toplam_hak
        elif izin_turu and not izin_turu.yillik_izin_mi:
            # Diğer izin türlerinde yıl güncel yıldan gelir; izin hakkı
            # gerekiyorsa kullanıcı tarafından elle girilebilir.
            cleaned["ait_oldugu_yil"] = str(timezone.localdate().year)
        if hak is not None and gun is not None:
            if gun > hak:
                self.add_error("kullanilan_gun", "\u0130zin s\u00fcresi izin hakk\u0131ndan fazla olamaz.")
            cleaned["kalan_izin"] = max(Decimal("0.00"), hak - gun)
        elif hak is None:
            cleaned["kalan_izin"] = None
        return cleaned


IzinTalebiForm = IzinEvrakForm

