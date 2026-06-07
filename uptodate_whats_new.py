#!/usr/bin/env python3
"""
UpToDate - What's New (Dahiliye) → Türkçe Özet → Word Belgesi

Kullanım:
    python3 uptodate_whats_new.py --user EMAIL --password SIFRE [--anthropic-key API_KEY]

    Ya da ortam değişkenleriyle:
    export UPTODATE_USER=email@example.com
    export UPTODATE_PASS=sifreniz
    export ANTHROPIC_API_KEY=sk-ant-...
    python3 uptodate_whats_new.py
"""

import os
import sys
import json
import time
import argparse
import re
from datetime import datetime
from typing import Optional

import requests
from bs4 import BeautifulSoup

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# -------------------------------------------------------------------
# Dahiliye ile ilgili What's New bölümleri
# -------------------------------------------------------------------
DAHILIYE_SECTIONS = [
    ("Genel Dahiliye / Birinci Basamak",
     "https://www.uptodate.com/contents/whats-new-in-primary-care"),
    ("Hastane Dahiliyesi",
     "https://www.uptodate.com/contents/whats-new-in-hospital-medicine"),
    ("Kardiyoloji",
     "https://www.uptodate.com/contents/whats-new-in-cardiovascular-medicine"),
    ("Endokrinoloji ve Diyabet",
     "https://www.uptodate.com/contents/whats-new-in-endocrinology-and-diabetes-mellitus"),
    ("Nefroloji ve Hipertansiyon",
     "https://www.uptodate.com/contents/whats-new-in-nephrology-and-hypertension"),
    ("Gastroenteroloji",
     "https://www.uptodate.com/contents/whats-new-in-gastroenterology"),
    ("Pulmoner ve Yoğun Bakım",
     "https://www.uptodate.com/contents/whats-new-in-pulmonary-and-critical-care-medicine"),
    ("Romatoloji",
     "https://www.uptodate.com/contents/whats-new-in-rheumatology"),
    ("Hematoloji",
     "https://www.uptodate.com/contents/whats-new-in-hematology"),
    ("Enfeksiyon Hastalıkları",
     "https://www.uptodate.com/contents/whats-new-in-infectious-diseases"),
    ("Nöroloji",
     "https://www.uptodate.com/contents/whats-new-in-neurology"),
    ("Onkoloji",
     "https://www.uptodate.com/contents/whats-new-in-oncology"),
]


# -------------------------------------------------------------------
# UpToDate Oturum Yönetimi
# -------------------------------------------------------------------
class UpToDateSession:
    BASE = "https://www.uptodate.com"
    LOGIN_PAGE = "https://www.uptodate.com/login"
    LOGIN_API = "https://www.uptodate.com/login"

    def __init__(self, username: str, password: str):
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
        })

    def login(self) -> bool:
        print("UpToDate'e giriş yapılıyor...")
        try:
            r = self.session.get(self.LOGIN_PAGE, timeout=20)
            soup = BeautifulSoup(r.text, "lxml")

            # CSRF token
            csrf = ""
            meta = soup.find("meta", {"name": "csrf-token"})
            if meta:
                csrf = meta.get("content", "")
            inp = soup.find("input", {"name": "_csrf"}) or soup.find(
                "input", {"name": "csrfToken"})
            if inp:
                csrf = inp.get("value", csrf)

            payload = {
                "username": self.username,
                "password": self.password,
                "_csrf": csrf,
            }

            headers = {
                "Referer": self.LOGIN_PAGE,
                "Content-Type": "application/x-www-form-urlencoded",
                "X-CSRF-Token": csrf,
            }

            resp = self.session.post(
                self.LOGIN_API, data=payload, headers=headers,
                allow_redirects=True, timeout=30
            )

            if "logout" in resp.text.lower() or "sign out" in resp.text.lower():
                print("✓ Giriş başarılı.")
                return True

            # JSON auth deneme
            json_headers = {
                "Referer": self.LOGIN_PAGE,
                "Content-Type": "application/json",
                "Accept": "application/json",
                "X-Requested-With": "XMLHttpRequest",
            }
            json_payload = {"username": self.username, "password": self.password}
            resp2 = self.session.post(
                "https://www.uptodate.com/services/app/login/v2",
                json=json_payload, headers=json_headers,
                allow_redirects=True, timeout=30,
            )
            if resp2.status_code == 200:
                data = resp2.json()
                if data.get("loginSuccess") or data.get("isLoggedIn"):
                    print("✓ Giriş başarılı (JSON API).")
                    return True
                print(f"  Giriş yanıtı: {json.dumps(data)[:200]}")

            print("⚠ Giriş başarısız. Şifreni/kullanıcı adını kontrol et.")
            print("  İpucu: Kurumsal/SSO hesap kullanıyorsan aşağıdaki cookie "
                  "yöntemini dene (--cookies parametresi).")
            return False
        except Exception as e:
            print(f"⚠ Giriş hatası: {e}")
            return False

    def load_cookies_from_file(self, path: str) -> bool:
        """Tarayıcıdan export edilmiş JSON cookie dosyasını yükler."""
        try:
            with open(path) as f:
                cookies = json.load(f)
            for c in cookies:
                self.session.cookies.set(
                    c["name"], c["value"], domain=c.get("domain", ".uptodate.com")
                )
            print(f"✓ {len(cookies)} cookie yüklendi: {path}")
            return True
        except Exception as e:
            print(f"⚠ Cookie yükleme hatası: {e}")
            return False

    def fetch_page(self, url: str) -> Optional[str]:
        try:
            r = self.session.get(url, timeout=30)
            if r.status_code == 200:
                return r.text
            print(f"  HTTP {r.status_code}: {url}")
            return None
        except Exception as e:
            print(f"  Hata [{url}]: {e}")
            return None


# -------------------------------------------------------------------
# İçerik Ayrıştırıcı
# -------------------------------------------------------------------
def parse_whats_new(html: str) -> list[dict]:
    """
    What's New sayfasından konu başlıkları ve özet paragraflarını çıkarır.
    Döndürdüğü liste: [{"title": ..., "summary": ..., "date": ...}, ...]
    """
    soup = BeautifulSoup(html, "lxml")
    items = []

    # UpToDate yapısı: her konu <div class="headingAnchor"> veya <h3> içinde
    # Özet metin ise hemen altındaki <p> ya da <div class="utdArticleSection">

    # Yöntem 1: section başlıklarını bul
    for tag in soup.find_all(["h2", "h3", "h4"]):
        title_text = tag.get_text(strip=True)
        if not title_text or len(title_text) < 5:
            continue
        # Takip eden kardeş paragrafları topla
        summary_parts = []
        for sibling in tag.next_siblings:
            if sibling.name in ["h2", "h3", "h4"]:
                break
            if sibling.name == "p":
                txt = sibling.get_text(" ", strip=True)
                if txt:
                    summary_parts.append(txt)
            if len(summary_parts) >= 3:
                break
        if summary_parts:
            items.append({
                "title": title_text,
                "summary": " ".join(summary_parts),
                "date": "",
            })

    # Yöntem 2: UpToDate'e özgü yapı (.cntBody .item vs.)
    if not items:
        for div in soup.select(".utdArticleSection, .sectionBody, .item"):
            strong = div.find("strong") or div.find("b")
            if strong:
                title_text = strong.get_text(strip=True)
                text = div.get_text(" ", strip=True)
                summary = text.replace(title_text, "", 1).strip()[:600]
                if title_text and summary:
                    items.append({"title": title_text, "summary": summary, "date": ""})

    # Yöntem 3: Düz metin fallback — numaralı listeler
    if not items:
        text = soup.get_text("\n")
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        i = 0
        while i < len(lines):
            if re.match(r"^\d+\.", lines[i]):
                title_text = re.sub(r"^\d+\.\s*", "", lines[i])
                summary = lines[i + 1] if i + 1 < len(lines) else ""
                items.append({"title": title_text, "summary": summary, "date": ""})
            i += 1

    return items


# -------------------------------------------------------------------
# Türkçe Özet Üretimi (Claude API)
# -------------------------------------------------------------------
def translate_to_turkish(items: list[dict], api_key: str) -> list[dict]:
    if not HAS_ANTHROPIC:
        print("⚠ anthropic paketi bulunamadı; Türkçe çeviri atlanıyor.")
        return items

    client = anthropic.Anthropic(api_key=api_key)
    results = []
    total = len(items)

    for idx, item in enumerate(items, 1):
        print(f"  [{idx}/{total}] Türkçe özet: {item['title'][:60]}...")
        prompt = (
            "Sen bir tıp doktorusun. Aşağıdaki İngilizce UpToDate 'What's New' "
            "başlığını ve özetini Türkçeye çevir. "
            "Başlığı Türkçe olarak yaz, ardından 2-3 cümlelik kısa ve net bir "
            "Türkçe klinik özet yaz. Tıbbi terminolojiyi doğru kullan.\n\n"
            f"BAŞLIK: {item['title']}\n"
            f"ÖZET: {item['summary']}"
        )
        try:
            msg = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=400,
                messages=[{"role": "user", "content": prompt}],
            )
            response_text = msg.content[0].text.strip()
            # İlk satır → Türkçe başlık, geri kalanı → özet
            lines = response_text.split("\n", 1)
            tr_title = lines[0].strip().lstrip("BAŞLIK:").lstrip("Başlık:").strip()
            tr_summary = lines[1].strip() if len(lines) > 1 else response_text
            results.append({
                "title_en": item["title"],
                "title_tr": tr_title,
                "summary_tr": tr_summary,
                "summary_en": item["summary"],
                "date": item.get("date", ""),
            })
            time.sleep(0.3)  # rate limit
        except Exception as e:
            print(f"    Hata: {e}")
            results.append({
                "title_en": item["title"],
                "title_tr": item["title"],
                "summary_tr": item["summary"],
                "summary_en": item["summary"],
                "date": item.get("date", ""),
            })
    return results


# -------------------------------------------------------------------
# Word Belgesi Oluşturucu
# -------------------------------------------------------------------
def create_word_document(sections_data: list[dict], output_path: str):
    if not HAS_DOCX:
        print("⚠ python-docx bulunamadı; Word belgesi oluşturulamıyor.")
        return

    doc = Document()

    # Sayfa kenar boşlukları
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.2)
    section.right_margin = Inches(1.2)

    # Ana Başlık
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("UpToDate — Dahiliye Güncellemeleri")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)

    # Tarih
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run(
        f"What's New — Türkçe Özet  |  {datetime.now().strftime('%d %B %Y')}"
    )
    date_run.font.size = Pt(11)
    date_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()  # boşluk

    total_topics = 0
    for section_data in sections_data:
        section_name = section_data["section_name"]
        items = section_data["items"]
        if not items:
            continue

        # Bölüm başlığı
        doc.add_heading(section_name, level=1)
        heading = doc.paragraphs[-1]
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)

        for i, item in enumerate(items, 1):
            # Konu başlığı (Türkçe)
            topic_para = doc.add_paragraph()
            num_run = topic_para.add_run(f"{i}. ")
            num_run.bold = True
            num_run.font.size = Pt(11)
            title_run = topic_para.add_run(item.get("title_tr", item.get("title_en", "")))
            title_run.bold = True
            title_run.font.size = Pt(11)
            title_run.font.color.rgb = RGBColor(0x1A, 0x5F, 0x9E)

            # İngilizce orijinal başlık (küçük, gri)
            if item.get("title_en") and item.get("title_en") != item.get("title_tr"):
                eng_para = doc.add_paragraph()
                eng_para.paragraph_format.left_indent = Pt(20)
                eng_run = eng_para.add_run(f"({item['title_en']})")
                eng_run.font.size = Pt(9)
                eng_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                eng_run.italic = True

            # Türkçe özet
            summary_text = item.get("summary_tr", "")
            if summary_text:
                sum_para = doc.add_paragraph()
                sum_para.paragraph_format.left_indent = Pt(20)
                sum_run = sum_para.add_run(summary_text)
                sum_run.font.size = Pt(10.5)

            total_topics += 1

        doc.add_paragraph()  # bölüm arası boşluk

    # Footer notu
    doc.add_paragraph()
    note = doc.add_paragraph()
    note_run = note.add_run(
        f"Bu belge UpToDate What's New bölümünden {total_topics} konu içermektedir. "
        "Kaynak: UpToDate® (Wolters Kluwer). Klinik karar için orijinal kaynağa başvurun."
    )
    note_run.font.size = Pt(8)
    note_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    note_run.italic = True

    doc.save(output_path)
    print(f"\n✓ Word belgesi oluşturuldu: {output_path}  ({total_topics} konu)")


# -------------------------------------------------------------------
# Demo / Offline mod
# -------------------------------------------------------------------
DEMO_DATA = {
    "Genel Dahiliye / Birinci Basamak": [
        {
            "title": "Ensitrelvir for post-exposure prophylaxis after household COVID-19 exposure",
            "summary": (
                "In a randomized trial (SCORPIO-PEP), ensitrelvir (antiviral) given within "
                "72 hours of household COVID-19 exposure significantly reduced infection risk "
                "(16.6% vs 24.6%; RR 0.67). The drug is already approved in Japan. "
                "UpToDate now supports considering ensitrelvir for high-risk household contacts."
            ),
        },
        {
            "title": "Sodium-glucose cotransporter 2 inhibitors and risk of Fournier's gangrene",
            "summary": (
                "Updated data confirm a small but real increased risk of Fournier's gangrene "
                "with SGLT2 inhibitors. Patients should be counseled about perineal pain, swelling, "
                "or erythema and seek immediate care if symptoms develop."
            ),
        },
        {
            "title": "Low-dose aspirin for primary prevention of cardiovascular events",
            "summary": (
                "Guidelines continue to advise against routine aspirin for primary prevention "
                "in adults ≥60 years due to bleeding risk outweighing benefit. "
                "Individualized decision-making remains important for select high-risk patients."
            ),
        },
    ],
    "Kardiyoloji": [
        {
            "title": "Colchicine for secondary prevention of major adverse cardiovascular events",
            "summary": (
                "Low-dose colchicine (0.5 mg daily) after MI reduces MACE by ~23% vs placebo "
                "(LoDoCo2, COLCOT trials). Updated UpToDate guidance now lists it as an option "
                "for secondary prevention in stable coronary disease."
            ),
        },
        {
            "title": "Transcatheter aortic valve replacement vs surgical AVR in low-risk patients",
            "summary": (
                "5-year follow-up from PARTNER 3 and Evolut Low Risk trials shows durable "
                "outcomes for TAVR comparable to SAVR in low-surgical-risk patients. "
                "TAVR can now be offered to anatomically suitable low-risk candidates."
            ),
        },
        {
            "title": "Finerenone in heart failure with mildly reduced or preserved ejection fraction",
            "summary": (
                "FINEARTS-HF trial showed finerenone (non-steroidal MRA) reduced worsening HF "
                "events and CV death in HFmrEF/HFpEF. FDA approval expected; UpToDate includes "
                "it as an emerging option."
            ),
        },
    ],
    "Endokrinoloji ve Diyabet": [
        {
            "title": "Tirzepatide for obesity and overweight",
            "summary": (
                "SURMOUNT-1 trial: tirzepatide 15 mg achieved mean weight loss of 20.9% over "
                "72 weeks. Now FDA-approved for chronic weight management (Zepbound). "
                "UpToDate includes it as first-line pharmacotherapy for eligible patients."
            ),
        },
        {
            "title": "Continuous glucose monitoring in type 2 diabetes managed with basal insulin",
            "summary": (
                "CGM use in T2DM on basal insulin improves HbA1c and reduces hypoglycemia "
                "vs standard glucose monitoring. UpToDate now recommends CGM for this group."
            ),
        },
    ],
    "Enfeksiyon Hastalıkları": [
        {
            "title": "Updated RSV vaccination recommendations for older adults",
            "summary": (
                "ACIP now recommends a single RSV vaccine dose for all adults ≥75 years and "
                "for adults 60-74 at increased risk of severe RSV. Three vaccines are available "
                "(Abrysvo, Arexvy, mRESVIA)."
            ),
        },
        {
            "title": "Doxycycline post-exposure prophylaxis for sexually transmitted infections",
            "summary": (
                "Doxy-PEP (200 mg within 72 h of condomless sex) reduces bacterial STI incidence "
                "by ~65% in MSM and transgender women. CDC now recommends offering it to "
                "appropriate high-risk individuals."
            ),
        },
    ],
    "Gastroenteroloji": [
        {
            "title": "Artificial intelligence–assisted colonoscopy and adenoma detection",
            "summary": (
                "Multiple RCTs confirm AI-assisted colonoscopy increases adenoma detection rate "
                "by ~10-15% vs standard colonoscopy. UpToDate acknowledges AI assistance as "
                "an option where available."
            ),
        },
    ],
    "Nefroloji ve Hipertansiyon": [
        {
            "title": "SGLT2 inhibitors in CKD: updated cardiovascular and renal outcomes",
            "summary": (
                "EMPA-KIDNEY trial confirmed empagliflozin reduces kidney disease progression "
                "and CV death in a broad CKD population (eGFR 20-45 or ≥45 with albuminuria). "
                "SGLT2i are now standard of care in CKD regardless of diabetes status."
            ),
        },
    ],
}


def run_demo_mode(api_key: Optional[str], output_path: str):
    print("\n[DEMO MODU] UpToDate örnek verileri kullanılıyor...")
    print("Gerçek veri için --user ve --password parametrelerini girin.\n")

    all_sections = []
    for section_name, raw_items in DEMO_DATA.items():
        print(f"  Bölüm: {section_name} ({len(raw_items)} konu)")
        if api_key and HAS_ANTHROPIC:
            translated = translate_to_turkish(raw_items, api_key)
        else:
            translated = [
                {
                    "title_en": it["title"],
                    "title_tr": it["title"],
                    "summary_tr": it["summary"],
                    "summary_en": it["summary"],
                    "date": "",
                }
                for it in raw_items
            ]
        all_sections.append({"section_name": section_name, "items": translated})

    create_word_document(all_sections, output_path)


# -------------------------------------------------------------------
# Ana Akış
# -------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="UpToDate What's New → Türkçe Özet → Word"
    )
    parser.add_argument("--user", default=os.getenv("UPTODATE_USER"), help="UpToDate kullanıcı adı")
    parser.add_argument("--password", default=os.getenv("UPTODATE_PASS"), help="UpToDate şifresi")
    parser.add_argument("--cookies", default=None, help="Tarayıcıdan export edilmiş cookie JSON dosyası")
    parser.add_argument("--anthropic-key", default=os.getenv("ANTHROPIC_API_KEY"), help="Anthropic API anahtarı")
    parser.add_argument("--output", default="UpToDate_Dahiliye_WhatsNew.docx", help="Çıktı dosya adı")
    parser.add_argument("--demo", action="store_true", help="Demo mod (gerçek giriş yapmadan)")
    args = parser.parse_args()

    if args.demo or (not args.user and not args.password and not args.cookies):
        run_demo_mode(args.anthropic_key, args.output)
        return

    # Gerçek mod
    sess = UpToDateSession(args.user or "", args.password or "")

    if args.cookies:
        ok = sess.load_cookies_from_file(args.cookies)
    else:
        ok = sess.login()

    if not ok:
        print("\nAlternatif: Tarayıcında UpToDate'e giriş yaptıktan sonra")
        print("'EditThisCookie' eklentisiyle cookie'leri JSON olarak dışa aktar,")
        print("dosya yolunu --cookies parametresiyle belirt.")
        sys.exit(1)

    all_sections = []
    for section_tr_name, url in DAHILIYE_SECTIONS:
        print(f"\n→ {section_tr_name}")
        html = sess.fetch_page(url)
        if not html:
            print(f"  Sayfa alınamadı: {url}")
            continue

        items = parse_whats_new(html)
        print(f"  {len(items)} konu bulundu.")

        if not items:
            continue

        if args.anthropic_key and HAS_ANTHROPIC:
            items = translate_to_turkish(items, args.anthropic_key)
        else:
            items = [
                {
                    "title_en": it["title"],
                    "title_tr": it["title"],
                    "summary_tr": it["summary"],
                    "summary_en": it["summary"],
                    "date": it.get("date", ""),
                }
                for it in items
            ]

        all_sections.append({"section_name": section_tr_name, "items": items})
        time.sleep(1)  # sunucu yükü azaltma

    if all_sections:
        create_word_document(all_sections, args.output)
    else:
        print("\n⚠ Hiç konu bulunamadı. Giriş başarılı oldu mu?")


if __name__ == "__main__":
    main()
