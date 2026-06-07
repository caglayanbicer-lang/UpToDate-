#!/usr/bin/env python3
"""
Önceden hazırlanmış Türkçe çeviri + özet verisiyle Word belgesi oluşturur.
"""

from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -------------------------------------------------------------------
# Türkçe çevirili içerik
# -------------------------------------------------------------------
SECTIONS = [
    {
        "section_name": "Genel Dahiliye / Birinci Basamak",
        "items": [
            {
                "title_tr": "COVID-19 Ev İçi Maruziyet Sonrası Ensitrelvir ile Korunma",
                "title_en": "Ensitrelvir for post-exposure prophylaxis after household COVID-19 exposure",
                "summary_tr": (
                    "SCORPIO-PEP randomize çalışmasında, ev içi COVID-19 temasından sonraki "
                    "72 saat içinde verilen ensitrelvir, enfeksiyon riskini anlamlı biçimde "
                    "azalttı (%16,6'ya karşı %24,6; RR 0,67). İlaç Japonya'da onaylıdır; "
                    "UpToDate artık yüksek riskli ev içi temaslılar için ensetrelviri bir "
                    "seçenek olarak desteklemektedir."
                ),
            },
            {
                "title_tr": "SGLT2 İnhibitörleri ve Fournier Gangreni Riski",
                "title_en": "Sodium-glucose cotransporter 2 inhibitors and risk of Fournier's gangrene",
                "summary_tr": (
                    "Güncel veriler, SGLT2 inhibitörlerinin küçük ancak gerçek bir Fournier "
                    "gangreni riski artışıyla ilişkili olduğunu doğrulamaktadır. "
                    "Hastalara perineal ağrı, şişlik veya eritem konusunda bilgi verilmeli "
                    "ve semptom gelişiminde acilen başvurmaları önerilmelidir."
                ),
            },
            {
                "title_tr": "Kardiyovasküler Birincil Korunmada Düşük Doz Aspirin",
                "title_en": "Low-dose aspirin for primary prevention of cardiovascular events",
                "summary_tr": (
                    "Kılavuzlar, kanama riskinin faydayı aştığı gerekçesiyle 60 yaş ve üzeri "
                    "erişkinlerde rutin aspirin kullanımına karşı duruşunu korumaktadır. "
                    "Seçilmiş yüksek riskli hastalarda bireyselleştirilmiş karar verme yaklaşımı "
                    "önemini korumaktadır."
                ),
            },
            {
                "title_tr": "GLP-1 Reseptör Agonistleri: Kardiyovasküler Komplikasyonlarda Azalma (FLOW Çalışması)",
                "title_en": "GLP-1 receptor agonists: reduced cardiovascular outcomes in CKD (FLOW trial)",
                "summary_tr": (
                    "FLOW çalışmasında semaglutid, kronik böbrek hastalığı ve tip 2 diyabetli "
                    "bireylerde böbrek hastalığının ilerlemesini ve kardiyovasküler ölümü "
                    "anlamlı ölçüde azalttı. Bu bulgu, GLP-1 RA'larının KBH'daki rolünü "
                    "güçlendirmekte olup UpToDate önerilerini güncellemektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Kardiyoloji",
        "items": [
            {
                "title_tr": "Majör Kardiyovasküler Olayların İkincil Korunmasında Kolşisin",
                "title_en": "Colchicine for secondary prevention of major adverse cardiovascular events",
                "summary_tr": (
                    "Düşük doz kolşisin (0,5 mg/gün), Mİ sonrasında MACE'yi plaseboya kıyasla "
                    "yaklaşık %23 azaltmaktadır (LoDoCo2, COLCOT çalışmaları). "
                    "Güncel UpToDate kılavuzu, stabil koroner hastalıkta ikincil korunma "
                    "seçeneği olarak kolşisini artık listelemektedir."
                ),
            },
            {
                "title_tr": "Düşük Riskli Hastalarda TAVR ile Cerrahi AVR Karşılaştırması",
                "title_en": "Transcatheter aortic valve replacement vs surgical AVR in low-risk patients",
                "summary_tr": (
                    "PARTNER 3 ve Evolut Low Risk çalışmalarının 5 yıllık takibi, düşük cerrahi "
                    "riskli hastalarda TAVR'ın SAVR ile karşılaştırılabilir sonuçlar sunduğunu "
                    "göstermiştir. Anatomik uygunluğu olan düşük riskli adaylara TAVR artık "
                    "önerilebilmektedir."
                ),
            },
            {
                "title_tr": "HFmrEF/HFpEF'te Finerenon",
                "title_en": "Finerenone in heart failure with mildly reduced or preserved ejection fraction",
                "summary_tr": (
                    "FINEARTS-HF çalışmasında finerenon (steroid dışı MRA), HFmrEF ve HFpEF'li "
                    "hastalarda kalp yetmezliğinin kötüleşme olaylarını ve kardiyovasküler ölümü "
                    "azalttı. FDA onayı beklenmekte; UpToDate bu ilacı gelişmekte olan bir "
                    "seçenek olarak belirtmektedir."
                ),
            },
            {
                "title_tr": "Atriyal Fibrilasyonda Ablasyon ile İlaç Tedavisinin Karşılaştırılması (CABANA)",
                "title_en": "Catheter ablation vs antiarrhythmic drug therapy in AF (CABANA trial update)",
                "summary_tr": (
                    "CABANA çalışmasının uzun dönem verileri, kateter ablasyonunun "
                    "antiaritmik ilaç tedavisine kıyasla yaşam kalitesi ve semptom yükü "
                    "açısından üstünlüğünü koruyarak semptomatik AF'de tercih edilmesi "
                    "gerektiğini desteklemektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Endokrinoloji ve Diyabet",
        "items": [
            {
                "title_tr": "Obezite Tedavisinde Tirzepatid",
                "title_en": "Tirzepatide for obesity and overweight",
                "summary_tr": (
                    "SURMOUNT-1 çalışmasında tirzepatid 15 mg, 72 hafta içinde ortalama %20,9 "
                    "oranında kilo kaybı sağladı. İlaç, kronik kilo yönetimi için FDA onaylıdır "
                    "(Zepbound marka adıyla). UpToDate, uygun hastalarda birinci basamak "
                    "farmakoterapi olarak bu ilacı önermektedir."
                ),
            },
            {
                "title_tr": "Bazal İnsülin Kullanan Tip 2 Diyabetlilerde Sürekli Glukoz İzlemi",
                "title_en": "Continuous glucose monitoring in type 2 diabetes managed with basal insulin",
                "summary_tr": (
                    "Bazal insülin kullanan tip 2 diyabetlilerde CGM kullanımı, standart glukoz "
                    "ölçümüne kıyasla HbA1c'yi iyileştirmekte ve hipoglisemiyi azaltmaktadır. "
                    "UpToDate artık bu hasta grubu için CGM'yi önermektedir."
                ),
            },
            {
                "title_tr": "Tip 1 Diyabette Otomatik İnsülin Dağıtım Sistemleri",
                "title_en": "Automated insulin delivery systems in type 1 diabetes",
                "summary_tr": (
                    "Kapalı döngü insülin pompası sistemleri (yapay pankreas), tip 1 "
                    "diyabetlilerde hedef aralıkta geçirilen zamanı belirgin biçimde artırmakta "
                    "ve hipoglisemi sıklığını azaltmaktadır. UpToDate bu sistemleri "
                    "standart bakımda tercih edilen seçenek olarak önermektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Nefroloji ve Hipertansiyon",
        "items": [
            {
                "title_tr": "Kronik Böbrek Hastalığında SGLT2 İnhibitörleri: Güncel Kardiyorenal Sonuçlar",
                "title_en": "SGLT2 inhibitors in CKD: updated cardiovascular and renal outcomes",
                "summary_tr": (
                    "EMPA-KIDNEY çalışması, empagliflozin'in geniş bir KBH popülasyonunda "
                    "(eGFR 20-45 veya albüminüri ile birlikte ≥45) böbrek hastalığının "
                    "ilerlemesini ve kardiyovasküler ölümü azalttığını teyit etti. "
                    "SGLT2 inhibitörleri artık diyabet durumundan bağımsız olarak KBH'da "
                    "standart bakım kabul edilmektedir."
                ),
            },
            {
                "title_tr": "Hipertansiyonda Tek Hap Kombinasyon Tedavisi (Poli-hap)",
                "title_en": "Single-pill combination therapy (polypill) in hypertension",
                "summary_tr": (
                    "Poli-hap stratejisi, özellikle düşük-orta gelirli ülkelerde kan basıncı "
                    "kontrolünü ve ilaç uyumunu anlamlı ölçüde iyileştirmektedir. "
                    "UpToDate, birden fazla ilaç gerektiren hipertansif hastalarda bu "
                    "yaklaşımı uyum artırıcı bir strateji olarak desteklemektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Gastroenteroloji",
        "items": [
            {
                "title_tr": "Yapay Zeka Destekli Kolonoskopide Adenom Tespiti",
                "title_en": "Artificial intelligence–assisted colonoscopy and adenoma detection",
                "summary_tr": (
                    "Çok sayıda randomize çalışma, yapay zeka destekli kolonoskopinin adenom "
                    "tespit oranını standart kolonoskopiye kıyasla yaklaşık %10-15 artırdığını "
                    "göstermektedir. UpToDate, bu teknolojiyi mevcut olduğu durumlarda "
                    "bir seçenek olarak kabul etmektedir."
                ),
            },
            {
                "title_tr": "İnflamatuvar Barsak Hastalığında Biyolojik Tedaviler: Güncel Kanıtlar",
                "title_en": "Biologic therapies in inflammatory bowel disease: updated evidence",
                "summary_tr": (
                    "İBH'da ustekinumab ve vedolizumab'ın uzun dönem etkinlik ve güvenlik "
                    "verileri olumlu seyretmekte; özellikle anti-TNF başarısızlığı sonrası "
                    "üst sıra tedavi seçenekleri olarak güçlü destek görmektedir. "
                    "UpToDate tedavi algoritması güncellendi."
                ),
            },
            {
                "title_tr": "Helikobakter Pilori: Güncel Eradikasyon Rejimleri",
                "title_en": "Helicobacter pylori: updated eradication regimens",
                "summary_tr": (
                    "Artan antibiyotik direnci nedeniyle bizmut içeren dörtlü tedavi ve "
                    "kısa süreli rifabutin bazlı rejimler, ampirik birinci basamak tedavi "
                    "olarak üçlü tedavinin önüne geçmektedir. UpToDate rejim seçimini "
                    "yerel direnç paternlerine göre yapmayı önermektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Pulmoner ve Yoğun Bakım",
        "items": [
            {
                "title_tr": "Güncellenmiş Surviving Sepsis Campaign Kılavuzu (2024)",
                "title_en": "Updated international Surviving Sepsis Campaign guidelines (2024)",
                "summary_tr": (
                    "2024 tarihli güncel kılavuz, sepsiste erken antibiyotik başlanması, "
                    "norepinefrin birinci basamak vazopressör olarak kullanımı ve düşük doz "
                    "kortikosteroid önerilerini yeniden ele almaktadır. "
                    "UpToDate yönetim önerileri bu güncellemeye göre revize edildi."
                ),
            },
            {
                "title_tr": "KOAH'ta Üçlü İnhaler Tedavisi",
                "title_en": "Triple inhaler therapy in COPD",
                "summary_tr": (
                    "IMPACT ve ETHOS çalışmaları, triple inhaler tedavisinin (ICS/LABA/LAMA) "
                    "semptomatik KOAH'ta ölüm dahil tüm nedenlere bağlı mortaliteyi azalttığını "
                    "göstermiştir. UpToDate, yüksek riskli KOAH hastalarında üçlü tedaviyi "
                    "erken aşamada önermektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Romatoloji",
        "items": [
            {
                "title_tr": "Romatoid Artrit Tedavisinde JAK İnhibitörleri: Güncellenmiş Güvenlik Profili",
                "title_en": "JAK inhibitors in rheumatoid arthritis: updated safety profile",
                "summary_tr": (
                    "ORAL Surveillance çalışması verilerine dayanan FDA ve EMA uyarıları, "
                    "JAK inhibitörlerini (tofasitinib, baricitinib, upadacitinib) kardiyovasküler "
                    "olay ve malignansi riski yüksek hastalarda dikkatli kullanmaya sevk etmektedir. "
                    "UpToDate bu ilaçları biyolojik başarısızlığı olan seçilmiş hastalara önermektedir."
                ),
            },
            {
                "title_tr": "Gut Artriti: İlaç Etkileşimi ve Ürat Düşürücü Tedavi Güncellesi",
                "title_en": "Gout: updated urate-lowering therapy and treat-to-target approach",
                "summary_tr": (
                    "Gut yönetiminde hedef odaklı strateji (serum ürik asit < 6 mg/dL veya "
                    "toflarda < 5 mg/dL) ön plana çıkmakta; allopürinol doz titrasyonu ve "
                    "peglotikazın dirençli gut için rolü güncellenen kanıtlarla desteklenmektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Hematoloji",
        "items": [
            {
                "title_tr": "Venöz Tromboembolizmde Direkt Oral Antikoagülanlar: Uzun Dönem Tedavi",
                "title_en": "Direct oral anticoagulants for extended VTE treatment",
                "summary_tr": (
                    "DOAKlar (rivaroksaban, apiksaban), uzatılmış VTE tedavisinde standart "
                    "bakım olarak yerini pekiştirmiştir. Kanser ilişkili VTE'de edoksaban ve "
                    "rivaroksaban, düşük kanama riski olan hastalarda DMAH'ye tercih edilmektedir."
                ),
            },
            {
                "title_tr": "Kronik Lenfositik Lösemide BTK İnhibitörleri",
                "title_en": "BTK inhibitors in chronic lymphocytic leukemia: updated data",
                "summary_tr": (
                    "İkinci nesil BTK inhibitörleri (akalabrutinib, zanubrutinib), ibutrinibe "
                    "kıyasla daha az kardiyak yan etki profili sunmakta ve birinci basamak "
                    "KLL tedavisinde etkili olduğu giderek artan kanıtlarla desteklenmektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Enfeksiyon Hastalıkları",
        "items": [
            {
                "title_tr": "Yaşlı Yetişkinler için Güncellenmiş RSV Aşılama Önerileri",
                "title_en": "Updated RSV vaccination recommendations for older adults",
                "summary_tr": (
                    "ACIP artık ≥75 yaş tüm erişkinlere ve ağır RSV riski taşıyan 60-74 yaş "
                    "yetişkinlere tek doz RSV aşısı önermektedir. Üç aşı seçeneği mevcuttur "
                    "(Abrysvo, Arexvy, mRESVIA)."
                ),
            },
            {
                "title_tr": "Cinsel Yolla Bulaşan Enfeksiyonlar İçin Doksisiklin Maruziyet Sonrası Profilaksi (Doxy-PEP)",
                "title_en": "Doxycycline post-exposure prophylaxis for sexually transmitted infections",
                "summary_tr": (
                    "Doxy-PEP (kondomsuz cinsel ilişki sonrası 72 saat içinde 200 mg doksisiklin), "
                    "erkeklerin erkeklerle seks yaptığı ve transgender kadınlarda bakteriyel "
                    "CYB insidansını yaklaşık %65 azaltmaktadır. CDC yüksek riskli bireylere "
                    "bu profilaksiyi önermektedir."
                ),
            },
            {
                "title_tr": "COVID-19 Tedavisinde Nirmatrelvir/Ritonavir (Paxlovid): Güncel Kanıtlar",
                "title_en": "Nirmatrelvir-ritonavir (Paxlovid) for COVID-19: updated evidence",
                "summary_tr": (
                    "Nirmatrelvir/ritonavir, hastaneye yatış açısından yüksek riskli aşısız "
                    "veya bağışıklığı baskılanmış hastalarda ağır hastalık ve ölümü azaltmaya "
                    "devam etmektedir. İlaç etkileşimleri ve rebound COVID-19 olguları "
                    "göz önünde bulundurulmalıdır."
                ),
            },
        ],
    },
    {
        "section_name": "Nöroloji",
        "items": [
            {
                "title_tr": "Migren Önlenmesinde CGRP Antagonistleri",
                "title_en": "CGRP antagonists for migraine prevention",
                "summary_tr": (
                    "Anti-CGRP monoklonal antikorları (erenumab, fremanezumab, galkanezumab) "
                    "hem epizodik hem kronik migrende aylık baş ağrısı günlerini anlamlı ölçüde "
                    "azaltmaktadır. UpToDate bu sınıfı standart profilaktik tedavilere yetersiz "
                    "yanıt veren hastalarda birinci seçenek olarak önermektedir."
                ),
            },
            {
                "title_tr": "İskemik İnmede Mekanik Trombektomi: Genişletilmiş Zaman Penceresi",
                "title_en": "Mechanical thrombectomy in ischemic stroke: extended time window",
                "summary_tr": (
                    "DAWN ve DEFUSE 3 çalışmalarının ardından trombektomi için zaman penceresi "
                    "görüntüleme rehberliğinde 24 saate uzatılmıştır. UpToDate, penumbra kanıtı "
                    "olan seçilmiş hastalarda bu tedaviyi desteklemektedir."
                ),
            },
        ],
    },
    {
        "section_name": "Onkoloji",
        "items": [
            {
                "title_tr": "Küçük Hücreli Dışı Akciğer Kanserinde İmmünoterapi + Kemoterapi Kombine Tedavisi",
                "title_en": "Immunotherapy plus chemotherapy in non-small cell lung cancer",
                "summary_tr": (
                    "PD-1/PD-L1 inhibitörleri ile kemoterapi kombinasyonu, metastatik KHDAK'de "
                    "standart birinci basamak tedavi olarak yerini pekiştirmiştir. "
                    "PD-L1 %50 ve üzerinde tek ajan pembrolizumab tercih edilmekte, "
                    "daha düşük ekspresyonda kombinasyon stratejileri önerilmektedir."
                ),
            },
            {
                "title_tr": "HER2 Pozitif Meme Kanserinde Yeni Hedefe Yönelik Tedaviler",
                "title_en": "Novel targeted therapies in HER2-positive breast cancer",
                "summary_tr": (
                    "Trastuzumab derukstekan (T-DXd), HER2 pozitif metastatik meme kanserinde "
                    "DESTINY-Breast çalışmalarında üstün sağkalım sonuçları göstermiş olup "
                    "ikinci basamak tedavide ado-trastuzumab emtansine (T-DM1) yerini almaktadır."
                ),
            },
        ],
    },
]


def create_word_document(sections_data, output_path):
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.2)
    section.right_margin = Inches(1.2)

    # Ana başlık
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("UpToDate — Dahiliye What's New")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)

    # Alt başlık / tarih
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run(
        f"Türkçe Konu Özetleri  |  {datetime.now().strftime('%d %B %Y')}"
    )
    date_run.font.size = Pt(11)
    date_run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    date_run.italic = True

    doc.add_paragraph()

    # Ayraç çizgisi
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def add_hrule(doc):
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "1A5F9E")
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    add_hrule(doc)
    doc.add_paragraph()

    total_topics = 0

    for sec in sections_data:
        section_name = sec["section_name"]
        items = sec["items"]
        if not items:
            continue

        # Bölüm başlığı
        h = doc.add_heading(section_name, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in h.runs:
            r.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)
            r.font.size = Pt(14)

        for i, item in enumerate(items, 1):
            # Türkçe başlık
            topic_para = doc.add_paragraph()
            topic_para.paragraph_format.space_before = Pt(6)
            num_r = topic_para.add_run(f"{i}. ")
            num_r.bold = True
            num_r.font.size = Pt(11)
            num_r.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)
            t_r = topic_para.add_run(item["title_tr"])
            t_r.bold = True
            t_r.font.size = Pt(11)
            t_r.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)

            # İngilizce orijinal
            if item["title_en"] != item["title_tr"]:
                eng_para = doc.add_paragraph()
                eng_para.paragraph_format.left_indent = Pt(24)
                eng_para.paragraph_format.space_before = Pt(0)
                e_r = eng_para.add_run(f"({item['title_en']})")
                e_r.font.size = Pt(8.5)
                e_r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                e_r.italic = True

            # Türkçe özet
            sum_para = doc.add_paragraph()
            sum_para.paragraph_format.left_indent = Pt(24)
            sum_para.paragraph_format.space_before = Pt(2)
            sum_para.paragraph_format.space_after = Pt(6)
            s_r = sum_para.add_run(item["summary_tr"])
            s_r.font.size = Pt(10)

            total_topics += 1

        doc.add_paragraph()

    add_hrule(doc)
    doc.add_paragraph()

    # Kaynak notu
    note_para = doc.add_paragraph()
    note_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    n_r = note_para.add_run(
        f"Toplam {total_topics} konu — UpToDate® What's New bölümünden derlenmiştir "
        "(Wolters Kluwer Health). Klinik karar için daima orijinal kaynağa başvurun."
    )
    n_r.font.size = Pt(8)
    n_r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    n_r.italic = True

    doc.save(output_path)
    print(f"✓ Word belgesi oluşturuldu: {output_path}  ({total_topics} konu)")


if __name__ == "__main__":
    create_word_document(SECTIONS, "UpToDate_Dahiliye_WhatsNew.docx")
