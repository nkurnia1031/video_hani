#!/usr/bin/env python3
"""Build editable SVG scene boards, PNG previews, a Word board, and the web board."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
from html import escape
from xml.etree import ElementTree as ET

import cairosvg
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
SCENES_DIR = ROOT / "scenes"
PREVIEW_DIR = ROOT / "previews"
ASSETS_DIR = ROOT / "assets"
SVG_NS = "http://www.w3.org/2000/svg"
ET.register_namespace("", SVG_NS)
W, H = 1920, 1080

INK = "#263238"
BLUE = "#407BFF"
BLUE_DARK = "#2859BE"
BLUE_PALE = "#E6F0FC"
PAPER = "#F8F7F2"
CORAL = "#ED8067"
MINT = "#B9D8CA"
LILAC = "#DADFFC"
WHITE = "#FFFFFF"
MUTED = "#697985"
PLUM = "#3D2645"
SAGE = "#7FB89A"
TERRACOTTA = "#E4572E"
MUSTARD = "#F2B134"
PEACH = "#FFC4A3"
CREAM = "#FFF6EC"
ESPRESSO = "#2A1F2D"
CASE_TINTS = ["#FFE3D3", "#DCEFE4", "#FCE9C2"]
PALETTE = [PLUM, SAGE, TERRACOTTA, MUSTARD, PEACH, CREAM, ESPRESSO]
PALETTE_NAMES = ["deep plum", "sage", "terracotta", "mustard", "peach", "cream", "espresso"]

ASSETS = {
    "presenter": {"file": ASSETS_DIR / "People-creating-robot_Rafiki.svg", "title": "People Creating Robot · Rafiki", "url": "https://storyset.com/illustration/people-creating-robot/rafiki", "license": "Storyset free with attribution"},
    "feed": {"file": REPO / "Feed-rafiki.svg", "title": "Feed", "url": "https://storyset.com/illustration/feed/rafiki"},
    "texting": {"file": REPO / "Texting-rafiki.svg", "title": "Texting", "url": "https://storyset.com/illustration/texting/rafiki"},
    "student_stress": {"file": ASSETS_DIR / "Student-stress_Rafiki.svg", "title": "Student Stress · Rafiki", "url": "https://storyset.com/illustration/student-stress/rafiki"},
    "feeling_blue": {"file": REPO / "Feeling Blue-rafiki.svg", "title": "Feeling Blue · Rafiki", "url": "https://storyset.com/illustration/feeling-blue/rafiki"},
    "bullying": {"file": ASSETS_DIR / "Bullying_Rafiki.svg", "title": "Bullying · Rafiki", "url": "https://storyset.com/illustration/bullying/rafiki"},
    "couple": {"file": REPO / "Couple stress-rafiki.svg", "title": "Couple Stress", "url": "https://storyset.com/illustration/couple-stress/rafiki"},
    "problem": {"file": REPO / "Problem solving-rafiki.svg", "title": "Problem Solving", "url": "https://storyset.com/illustration/problem-solving/rafiki"},
    "document": {"file": REPO / "Online document-rafiki.svg", "title": "Online Document", "url": "https://storyset.com/illustration/online-document/rafiki"},
    "research": {"file": REPO / "Research paper-rafiki.svg", "title": "Research Paper", "url": "https://storyset.com/illustration/research-paper/rafiki"},
    "todo": {"file": ASSETS_DIR / "To-do-list-rafiki.svg", "title": "To do list", "url": "https://storyset.com/illustration/to-do-list/rafiki"},
    "social": {"file": REPO / "Social interaction-rafiki.svg", "title": "Social Interaction", "url": "https://storyset.com/illustration/social-interaction/rafiki"},
    "coworkers": {"file": ASSETS_DIR / "Co-workers_Rafiki.svg", "title": "Co-workers · Rafiki", "url": "https://storyset.com/illustration/co-workers/rafiki"},
}

SCENE1_COMMENT_SOURCES = [
    {
        "id": "avoidant", "quote": "pengalaman gue pacaran sama orang avoidant sih",
        "lines": ["pengalaman gue pacaran", "sama orang avoidant sih"], "sequence": 1,
        "alias": "Naya",
        "platform": "Reddit · r/r4rindonesia", "url": "https://www.reddit.com/r/r4rindonesia/comments/1oi543a/dealing_with_avoidant_partner_with_possible_ptsd/",
        "posted": "2025-10-28", "accessed": "2026-11-26", "kind": "ASLI",
    },
    {
        "id": "toxic", "quote": "gua ngerasa lg “toxic” bgt.",
        "lines": ["gua ngerasa lg “toxic” bgt."], "sequence": 2,
        "alias": "Bagas",
        "platform": "Reddit · r/indonesia", "url": "https://www.reddit.com/r/indonesia/comments/n73anm/",
        "posted": "2021-05-08", "accessed": "2026-11-26", "kind": "ASLI",
    },
    {
        "id": "depresi", "quote": "depresi dikira sedih dll",
        "lines": ["depresi dikira sedih dll"], "sequence": 3,
        "alias": "Rani",
        "platform": "Reddit · r/indonesiabebas", "url": "https://www.reddit.com/r/indonesiabebas/comments/1wje4q0/jangan_sepelekan_stress/",
        "posted": "2026-09-18", "accessed": "2026-11-26", "kind": "ASLI",
    },
    {
        "id": "narcissistic", "quote": "Nggak sesuai ekspektasi dikit, langsung dibilang narcissistic.",
        "lines": ["Nggak sesuai ekspektasi dikit,", "langsung dibilang narcissistic."], "sequence": 4,
        "alias": "Dito",
        "platform": "Tiruan · tidak berasal dari platform atau akun nyata", "url": "",
        "posted": "Tiruan, ditulis 2026-11-26", "accessed": "2026-11-26", "kind": "TIRUAN (bukan komentar asli)",
    },
]

SCENE2_COMMENTS = [
    {
        "id": "comment-01", "handle": "@raraa", "accent": PLUM,
        "lines": [
            [("Pacar gue tiap ada masalah suka ngilang.", INK, False)],
            [("Gue chat malam, pagi baru dibales", INK, False)],
            [("‘ketiduran’. Fix ", INK, False), ("avoidant", BLUE_DARK, True), (" sih.", INK, False)],
        ],
    },
    {
        "id": "comment-02", "handle": "@anakrantau", "accent": MUSTARD,
        "lines": [
            [("Gue kemarin dimarahin temen.", INK, False)],
            [("Padahal dia yang salah, malah marah", INK, False)],
            [("ke gue. ", INK, False), ("Toxic", CORAL, True), (" banget, asli.", INK, False)],
        ],
    },
    {
        "id": "comment-03", "handle": "@halokak", "accent": PEACH,
        "lines": [
            [("Ih kak, itu ", INK, False), ("gaslight", BLUE_DARK, True), (" dah kayaknya.", INK, False)],
            [("Padahal kakaknya cuma cerita soal", INK, False)],
            [("apa yang dia rasain.", INK, False)],
        ],
    },
    {
        "id": "comment-04", "handle": "@ncy", "accent": PLUM,
        "lines": [
            [("Kak, dia murung terus akhir-akhir ini.", INK, False)],
            [("Kayaknya ", INK, False), ("depresi", CORAL, True), (" deh.", INK, False)],
        ],
    },
    {
        "id": "comment-05", "handle": "@vanyaa", "accent": PEACH,
        "lines": [
            [("Baru kenal seminggu udah dikirimin", INK, False)],
            [("hadiah tiap hari. ", INK, False), ("Love bombing", CORAL, True), (" nggak sih?", INK, False)],
        ],
    },
    {
        "id": "comment-06", "handle": "@dindap", "accent": PLUM,
        "lines": [
            [("Gue bilang butuh waktu sendiri, malah", INK, False)],
            [("dibilang nggak punya ", INK, False), ("boundaries.", BLUE_DARK, True), (" Lah?", INK, False)],
        ],
    },
    {
        "id": "comment-07", "handle": "@mikaela", "accent": MUSTARD,
        "lines": [
            [("Ih kak, dia cuma beda pendapat, kok", INK, False)],
            [("langsung dibilang ", INK, False), ("narcissist?", BLUE_DARK, True)],
        ],
    },
    {
        "id": "comment-08", "handle": "@firaaja", "accent": PEACH,
        "lines": [
            [("Nggak bales chat seharian dibilang", INK, False)],
            [
                ("ghosting.", CORAL, True),
                (" Kadang emang lagi sibuk, bestie.", INK, False),
            ],
        ],
    },
    {
        "id": "comment-09", "handle": "@rereaja", "accent": PEACH,
        "lines": [
            [("Baru putus seminggu, dibilang ", INK, False), ("trauma.", BLUE_DARK, True)],
            [("Padahal gue cuma lagi sedih, wkwk.", INK, False)],
        ],
    },
    {
        "id": "comment-10", "handle": "@eka", "accent": PLUM,
        "lines": [
            [("Ih, liburan aja dibilang ", INK, False), ("healing.", CORAL, True)],
            [("Istilahnya dipakai buat semua hal.", INK, False)],
        ],
    },
]

SCENES = [
    {
        "no": 1, "slug": "the-hook", "time": "00:00–00:30", "title": "The Hook",
        "visual": "Live presenter in medium close-up, surrounded by fast social-comment bubbles; the frame then zooms toward a smartphone.",
        "screen": "Are we using psychology terms correctly?",
        "motion": "Komentar masuk cepat mengitari presenter. Presenter menunjuk kamera; garis zoom mengantar ke layar smartphone.",
        "voice": "Guys, pernah nggak sih, cuma karena seseorang nggak membalas chat, kita langsung bilang, 'Dia avoidant!' Teman nggak jadi datang, kita sebut toxic. Lagi sedih beberapa hari, langsung merasa depresi. Atau ketika seseorang nggak sesuai ekspektasi, tiba-tiba kita bilang, 'That's narcissistic!' Wait. Are we actually understanding psychology, or are we just throwing psychological terms everywhere? Welcome to Therapy Speak Gone Wrong. Let's talk about it!",
        "assets": [], "roles": ["Frame presenter live dan antarmuka smartphone dibuat sebagai vektor SVG yang bisa diedit."],
        "credits": [], "background": "Warm paper, a blue footage frame, and small coral comment accents.",
    },
    {
        "no": 2, "slug": "what-is-therapy-speak", "time": "00:30–01:05", "title": "What Is Therapy Speak?",
        "visual": "Wide cream living-room frame: TV and wall art on the left, a comment feed on an espresso phone at center, and the reclining man on a sage sofa with colored shelves at right.",
        "screen": "02A: intact comments with only the term highlighted. 02B: #Toxic · #Narcissist · #Gaslighting · #Trauma · #Avoidant · #Depressed. 02C: Understanding psychology = Good! · Misusing psychological terms = Problematic.",
        "motion": "02A · 00:30–00:45: comments scroll; the gaslighting comment pops out and zooms on the left as the VO names the term. 02B · 00:45–00:55: selected terms lift into a full-screen hashtag stack during the benefits of learning psychology. 02C · 00:55–01:05: the final problem statement arrives on the VO's ‘But here's the problem’ and holds.",
        "voice": "Di era social media, istilah psikologi semakin sering kita dengar. Mulai dari toxic relationship, gaslighting, trauma, sampai narcissistic behavior. Fenomena ini sering disebut sebagai therapy speak, yaitu penggunaan istilah yang berasal dari dunia psikologi atau terapi dalam percakapan sehari-hari. Actually, there's nothing wrong with learning psychological terms. Justru, memahami psikologi bisa membantu kita mengenali emosi, memahami perilaku, dan membangun hubungan yang lebih sehat. But here's the problem: ketika istilah-istilah tersebut digunakan tanpa memahami maknanya, kita bisa salah menilai diri sendiri maupun orang lain.",
        "assets": ["feed"], "roles": ["Feed-rafiki.svg: karakter dan sofa dipisahkan dari seluruh perangkat latar; ukurannya disesuaikan dengan komposisi ruang yang lebih lebar.", "Rak buku, bunga, TV, lukisan, dan latar ruang biru adalah vektor editable.", "Scene-02-komponen.svg: enam komentar tambahan sebagai grup SVG terpisah untuk animasi scroll."],
        "credits": ["feed"], "background": "02A: cream room, espresso phone, plum UI details, and restrained warm accents. 02B: deep plum full-screen word cloud in Caveat handwritten font, with cream base text and selected terracotta, mustard, and peach terms. 02C: centered sage and terracotta statements.",
        "frames": [
            {"code": "02A", "slug": "02a-comment-scroll", "title": "Phone Zoom + Comment Scroll", "time": "00:30–00:45"},
            {"code": "02B", "slug": "02b-hashtag-cloud", "title": "Hashtag Cloud Transition", "time": "00:45–00:55"},
            {"code": "02C", "slug": "02c-on-screen-text", "title": "Understanding Psychology · On-Screen Text", "time": "00:55–01:05"},
        ],
    },
    {
        "no": 3, "slug": "normal-emotions-become-labels", "time": "01:05–01:55", "title": "When Normal Emotions Become Labels",
        "visual": "Three separate full-screen examples: an unanswered message, a worried student at a study desk as a crossed-out 100 changes to a red 45 remedial score, and a disagreement. A final frame brings the examples together behind centered on-screen text.",
        "screen": "They're avoidant! · I'm depressed! (target tidak tercapai) · You're gaslighting me!\nOne behavior ≠ A psychological diagnosis.",
        "motion": "03A–03C: tiap contoh tampil sendiri; bubble terracotta muncul lebih dulu, lalu koreksi sage. 03B memakai ilustrasi Raka yang sama dengan 4A dan penanda target gagal yang terpisah. 03D: ketiga contoh bertemu dan kalimat penutup tampil di tengah.",
        "voice": "Contoh pertama. Teman kita belum membalas chat selama beberapa jam. Lalu kita langsung menyimpulkan bahwa dia punya avoidant attachment. Padahal, mungkin dia sedang sibuk, kelelahan, atau memang belum sempat membalas. One situation doesn't define someone's entire personality. Contoh kedua. Kita merasa sedih setelah gagal mencapai target. Kemudian langsung menyebut diri sendiri mengalami depresi. Padahal, sedih merupakan emosi manusia yang normal. Kesedihan setelah mengalami kegagalan tidak otomatis menunjukkan adanya gangguan depresi. Contoh ketiga. Seseorang tidak setuju dengan pendapat kita, lalu langsung dituduh melakukan gaslighting. Padahal, disagreement atau perbedaan pendapat tidak otomatis termasuk gaslighting. Kita perlu melihat konteks, pola perilaku, dan bagaimana tindakan tersebut terjadi.",
        "assets": ["texting", "student_stress", "couple"], "roles": ["Texting-rafiki.svg: pesan belum dibalas di 03A.", "Student-stress_Rafiki.svg: Raka dari pose 4A dipakai kembali di 03B; penanda target gagal menjadi overlay animasi terpisah.", "Couple stress-rafiki.svg: konteks perbedaan pendapat."],
        "credits": ["texting", "student_stress", "couple"], "background": "Case 1 peach tint #FFE3D3, Case 2 sage tint #DCEFE4, and Case 3 mustard tint #FCE9C2; terracotta labels and sage corrections use the PDF's contrast rules.",
        "frames": [
            {"code": "03A", "slug": "03a-case-01-avoidant", "title": "Case 1 · Unanswered Message", "time": "01:05–01:20"},
            {"code": "03B", "slug": "03b-case-02-depressed", "title": "Case 2 · Sadness After Failure", "time": "01:20–01:35"},
            {"code": "03C", "slug": "03c-case-03-gaslighting", "title": "Case 3 · Disagreement", "time": "01:35–01:50"},
            {"code": "03D", "slug": "03d-combined-on-screen-text", "title": "Combined Examples · On-Screen Text", "time": "01:50–01:55"},
        ],
    },
    {
        "no": 4, "slug": "why-it-matters", "time": "01:55–02:35", "title": "Why Does It Matter?",
        "visual": "Four full-screen scenes chain the consequences of careless labels: self-misunderstanding, conflict, stigma, then a turn toward evidence and context.",
        "screen": "04D (akhir): Psychology is more than viral terms. It needs context, evidence, and careful understanding.",
        "motion": "04A: emosi berubah menjadi self-label. 04B: label berpindah ke teman dan meretakkan hubungan. 04C: satu label membesar menjadi stigma; kartu konteks mengitari orangnya. 04D: label terurai menjadi konteks, bukti, dan pemahaman cermat; teks layar tampil di akhir.",
        "voice": "Now, why should we care? First, misusing psychological terms can make us misunderstand our own emotions. Normal feelings like sadness, fear, or disappointment can feel like something is seriously wrong with us. Second, it can create unnecessary conflict. Imagine labeling someone as narcissistic just because they disagree with you. And third, it can reinforce stigma. Mental health conditions are complex experiences, not simply labels we can attach to someone based on one behavior. Remember, psychology is more than just a collection of viral terms. It is a scientific discipline that requires context, evidence, and careful understanding.",
        "assets": ["feeling_blue", "couple", "bullying", "research"],
        "roles": ["04A · Feeling Blue: emosi sedih menjadi self-label.", "04B · Couple Stress: beda pendapat berubah menjadi label personal dan jarak.", "04C · Bullying: stigma divisualkan sebagai pelabelan yang mengucilkan seseorang; kartu konteks menjaga pengalaman tetap terlihat kompleks.", "04D · Research Paper: menutup rantai dengan bukti dan konteks; teks layar muncul terakhir."],
        "credits": ["feeling_blue", "couple", "bullying", "research"],
        "background": "04A peach, 04B mustard, 04C sage, and 04D cream; terracotta marks careless labels, sage carries context, and plum anchors the closing statement.",
        "frames": [
            {"code": "04A", "slug": "04a-emotions-misunderstood", "title": "Normal Emotions Misread", "time": "01:55–02:08"},
            {"code": "04B", "slug": "04b-label-causes-conflict", "title": "A Label Turns Disagreement Into Conflict", "time": "02:08–02:15"},
            {"code": "04C", "slug": "04c-one-label-becomes-stigma", "title": "One Label Hides a Complex Experience", "time": "02:15–02:25"},
            {"code": "04D", "slug": "04d-context-evidence-onscreen-text", "title": "Context, Evidence · On-Screen Text", "time": "02:25–02:35"},
        ],
    },
    {
        "no": 5, "slug": "ethics", "time": "02:35–03:20", "title": "The Ethical Side of Psychology",
        "visual": "Satu komposisi animasi layar penuh: dokumen digital dan riset di kiri, empat kartu pasal berurutan di kanan, lalu pita dasar ilmiah dan kehati-hatian menutup komposisi.",
        "screen": "Pasal 9 · 15 · 28 · 31 · Scientific basis + professional care",
        "motion": "Dokumen terbuka dengan gerak lembaran, empat kartu pasal masuk bertahap dengan aksen penanda, lalu pita dasar ilmiah menyapu sebagai penutup. Tidak ada judul produksi, timestamp, atau footer di dalam frame.",
        "voice": "That's also why ethical responsibility matters when we talk about psychology. Dalam Kode Etik Psikologi Indonesia, ada beberapa prinsip yang relevan dengan fenomena ini. Pasal 9 menekankan bahwa keputusan psikolog dan ilmuwan psikologi harus didasarkan pada pengetahuan ilmiah dan sikap profesional yang telah teruji. Artinya, psychological claims should not be based on assumptions or viral trends alone. Pasal 15 membahas penghindaran dampak buruk. Jadi, informasi psikologi perlu disampaikan dengan mempertimbangkan kemungkinan dampaknya terhadap orang lain. Pasal 28 mengingatkan bahwa pernyataan publik harus disampaikan secara bijaksana, jujur, teliti, dan hati-hati, serta berpedoman pada dasar ilmiah. Sementara itu, Pasal 31 menekankan bahwa pernyataan melalui media harus konsisten dengan kode etik dan didasarkan pada pengetahuan serta praktik psikologi yang tepat. So, sharing psychology is not just about being relatable. It's also about being responsible.",
        "assets": ["document", "research"], "roles": ["Online document-rafiki.svg: halaman digital.", "Research paper-rafiki.svg: konteks dasar ilmiah.", "Empat kartu pasal dibuat sebagai vektor."],
        "credits": ["document", "research"], "background": "Warm white document page on a pale-lilac digital canvas.",
    },
    {
        "no": 6, "slug": "use-therapy-speak-wisely", "time": "03:20–03:50", "title": "How to Use Therapy Speak Wisely",
        "visual": "Satu komposisi animasi layar penuh: ilustrasi checklist di kiri dan tiga langkah praktis di kanan, diakhiri ajakan membiarkan emosi manusia hadir.",
        "screen": "Understand before you label · Describe, don’t diagnose · Learn, don’t self-diagnose · Allow normal human emotions.",
        "motion": "Checklist bergerak masuk dan tiga langkah tersusun berurutan. Frasa ‘Allow normal human emotions.’ menjadi teks terakhir di layar. Frame hanya berisi ilustrasi, langkah, ikon, dan teks animasi.",
        "voice": "So, how can we use psychological terms more wisely? Here's your three-step reminder. Number one, understand before you label. Cari tahu makna suatu istilah sebelum menggunakannya. Number two, describe behavior, not diagnose people. Daripada langsung menyebut seseorang narcissist, lebih baik jelaskan perilaku spesifik yang membuat kita merasa tidak nyaman. Number three, know your limits. Konten media sosial bisa menjadi sumber edukasi awal, but it cannot replace professional assessment. And most importantly, allow yourself and others to experience normal human emotions without immediately turning them into disorders.",
        "assets": ["todo"], "roles": ["To-do list-rafiki.svg: ilustrasi checklist; tiga langkah dan ikon pendamping disusun sebagai vektor."],
        "credits": ["todo"], "background": "Light mint backdrop with blue checklist rows.",
    },
    {
        "no": 7, "slug": "closing", "time": "03:50–04:00", "title": "Closing",
        "visual": "Komposisi penutup layar penuh: istilah viral memudar di dekat kelompok, empat ekspresi emosi tetap berwarna, lalu pesan penutup memenuhi sisi kanan.",
        "screen": "THINK BEFORE YOU LABEL.\nStay curious. Stay kind!",
        "motion": "Istilah viral meredup, empat ekspresi menyala bergantian, kemudian ‘THINK BEFORE YOU LABEL.’ dan ‘Stay curious. Stay kind!’ masuk sebagai teks akhir. Tidak ada slot produksi atau identitas pembuat di artwork.",
        "voice": "Psychology should help us understand people, not reduce them to labels. So, before you say it, make sure you understand it. Because not every emotion is a disorder, and not every behavior needs a psychological label. Think before you label. Stay curious, stay kind!",
        "assets": ["social"], "roles": ["Social interaction-rafiki.svg: kelompok dengan ekspresi beragam.", "Empat ikon emosi dan end card dibuat sebagai vektor."],
        "credits": ["social"], "background": "Soft peach paper, muted fading labels, and a clear final card.",
    },
]

# Revised visual scope from the attached script. Dialogue and voice-over stay verbatim.
SCENES[0].update({
    "visual": "Satu komposisi utuh berlatar plum menampilkan perempuan dari grup Character pada ilustrasi People Creating Robot · Rafiki, diisolasi sebagai close-up setengah badan di tengah. Robot, latar, panel, dan properti luar dihilangkan. Empat kartu komentar publik tampil bergantian di sisi kiri dan kanan dengan nama samaran serta avatar generik.",
    "screen": "Are we using psychology terms correctly?",
    "motion": "Satu scene berlanjut tanpa pemotongan keyframe: empat screenshot komentar publik masuk berurutan di kedua sisi presenter, mengikuti dialog ‘Dia avoidant!’ → ‘toxic’ → ‘depresi’ → ‘That's narcissistic!’. Presenter tetap stabil tanpa pointer atau frame. Nama pada kartu adalah pseudonim; avatar tetap generik dan tidak ada logo aplikasi. Pertanyaan penutup muncul sebelum transisi menuju scene smartphone. Kutipan dan tanggal akses tercatat pada audit sumber Scene 1.",
    "assets": ["presenter"],
    "roles": ["People Creating Robot · Rafiki: hanya objek perempuan dari grup Character dipakai. Robot, panel, dan latar sumber tidak dimasukkan. Crop menampilkan setengah badan; enam path pakaian biru diwarnai Cream. Tablet kecil tetap di tangannya agar pose sumber tersambung alami. Dari 41 drawable dalam objek Character, enam terpotong oleh crop dan enam diwarnai ulang; 29/41 (70,7%) tetap utuh.", "Empat kartu komentar memakai potongan publik. Kartu memakai nama samaran Naya, Bagas, Rani, dan Dito; nama akun sumber, foto profil, serta logo platform tidak ditampilkan. Tiga komentar berlabel ASLI dengan tautan; komentar narcissistic diberi label TIRUAN (bukan komentar asli) pada audit sumber.", "Presenter, empat kartu komentar, dan pertanyaan penutup adalah grup SVG editable. Tidak ada pointer atau bingkai presenter."],
    "credits": ["presenter"],
    "background": "Deep Plum #3D2645 polos. Pakaian presenter memakai Cream #FFF6EC, kulit mengikuti warna ilustrasi sumber, dan kartu komentar memakai Cream, Espresso, serta Peach.",
    "frames": [
        {"code": "01", "slug": "01-the-hook", "title": "The Hook · One Continuous Scene", "time": "00:00–00:30"},
    ],
})
SCENES[2].update({
    "visual": "Tiga case terpisah, masing-masing memakai satu ilustrasi dan satu latar warna sendiri: pesan belum dibalas; Raka dari frame 4A sedih setelah gagal mencapai target; dua orang berbeda pendapat. Koreksi sage muncul saat VO sampai ‘Padahal...’. Frame terakhir menampilkan kalimat penutup di tengah.",
    "screen": "03A · They're avoidant! · Mungkin sibuk, kelelahan, atau belum sempat.\n03B · I'm depressed! · Penanda target tidak tercapai · Sedih setelah gagal itu emosi yang normal.\n03C · You're gaslighting me! · Beda pendapat bukan otomatis gaslighting.\n03D · One behavior ≠ A psychological diagnosis.",
    "motion": "Tiga case tampil sendiri-sendiri. Di tiap case, bubble label terracotta muncul dahulu; saat VO mencapai ‘Padahal...’, kotak koreksi sage masuk. 03B memakai Raka dari frame 4A dan mempertahankan penanda target gagal sebagai layer terpisah. Frame 03D menyatukan komposisi dan menampilkan teks penutup.",
    "roles": ["Texting · Rafiki: teman mengirim pesan yang hanya dibaca, tanpa balasan.", "Co-workers · Rafiki: Raka yang sama dengan 4B/4A dipakai dengan pose stres; suasana meja belajar dan kartu target gagal tetap terpisah sebagai layer komposisi.", "Couple Stress · Rafiki: dua orang tidak sependapat. Semua asset credit dan komposisi tetap berada di luar artwork frame."],
    "assets": ["texting", "student_stress", "couple", "coworkers"],
    "credits": ["texting", "student_stress", "couple", "coworkers"],
    "background": "Case 1 peach muda #FFE3D3, Case 2 sage muda #DCEFE4, Case 3 mustard muda #FCE9C2. Label terracotta dengan teks putih; koreksi sage dengan teks espresso. Maksimal tiga warna dominan per case.",
    "frames": [
        {"code": "03A", "slug": "03a-case-01-avoidant", "title": "Case 1 · Unanswered Message", "time": "01:05–01:20", "voice": "Contoh pertama. Teman kita belum membalas chat selama beberapa jam. Lalu kita langsung menyimpulkan bahwa dia punya avoidant attachment. Padahal, mungkin dia sedang sibuk, kelelahan, atau memang belum sempat membalas. One situation doesn't define someone's entire personality."},
        {"code": "03B", "slug": "03b-case-02-depressed", "title": "Case 2 · Sadness After Failure", "time": "01:20–01:35", "voice": "Contoh kedua. Kita merasa sedih setelah gagal mencapai target. Kemudian langsung menyebut diri sendiri mengalami depresi. Padahal, sedih merupakan emosi manusia yang normal. Kesedihan setelah mengalami kegagalan tidak otomatis menunjukkan adanya gangguan depresi."},
        {"code": "03C", "slug": "03c-case-03-gaslighting", "title": "Case 3 · Disagreement", "time": "01:35–01:50", "voice": "Contoh ketiga. Seseorang tidak setuju dengan pendapat kita, lalu langsung dituduh melakukan gaslighting. Padahal, disagreement atau perbedaan pendapat tidak otomatis termasuk gaslighting. Kita perlu melihat konteks, pola perilaku, dan bagaimana tindakan tersebut terjadi."},
        {"code": "03D", "slug": "03d-combined-on-screen-text", "title": "One Behavior · Closing On-Screen Text", "time": "01:50–01:55", "voice": ""},
    ],
})
SCENES[3].update({
    "time": "01:55–02:35",
    "visual": "Satu cerita sebab-akibat Raka dan Dina bergerak dalam lima beat: Raka salah memahami emosinya → konflik saat rapat → ia menyimpulkan Dina narsistik tanpa bukti → sebuah post menyebarkan stigma dan Raka mengeluarkan Dina dari grup → Dina sendirian tanpa Raka, ditutup teks layar.",
    "screen": "4A · Aku depresi. Ada yang salah banget sama aku.\n4B · Raka: Kamu selalu bikin aku merasa salah! · Dina: Aku cuma kasih masukan. Jangan baper terus, deh.\n4B2 · Dia narsistik. · Cap: NARSISTIK.\n4C · Hati-hati sama Dina, dia narsistik. · Biar yang lain aman. · Dina dikeluarkan dari grup; post menyebar menjadi notifikasi.\n4D · Don't reduce people to labels.",
    "motion": "Transisi objek tetap mengalir: bubble self-label menjadi awan pikiran dan debat → berubah menjadi cap NARSISTIK → cap menyebar sebagai post dan rantai notifikasi → berakhir pada Dina sendirian dengan bubble ‘...’ dan teks layar. Raka tidak muncul di frame 4D.",
    "voice": "Now, why should we care? First, misusing psychological terms can make us misunderstand our own emotions. Normal feelings like sadness, fear, or disappointment can feel like something is seriously wrong with us. Second, it can create unnecessary conflict. Imagine labeling someone as narcissistic just because they disagree with you. And third, it can reinforce stigma. Mental health conditions are complex experiences, not simply labels we can attach to someone based on one behavior. Remember, psychology is more than just a collection of viral terms. It is a scientific discipline that requires context, evidence, and careful understanding.",
    "assets": ["coworkers"],
    "roles": ["4A memakai basis kantor Co-workers · Rafiki dari 4B dengan objek Dina dihilangkan; Raka dan latar tetap dari ilustrasi yang sama. 4B, 4B2, dan 4C memakai ilustrasi dasar utuh dengan Raka dan Dina bersama. Label, kartu target, post, notifikasi, dan teks menjadi overlay terpisah.", "4B dan 4B2 · Co-workers · Rafiki: karakter Raka dan Dina di rapat serta cap NARSISTIK. Bubble dialog 4B tanpa panah/garis penghubung; cap tetap kesimpulan Raka, bukan diagnosis.", "4C · Post peringatan dan rantai notifikasi dibuat sebagai layer SVG editable di sisi kanan komposisi.", "4D · Co-workers · Rafiki Dina tetap menjadi satu-satunya karakter; wash cream, area meja yang ditutup, bubble ‘...’, notifikasi tak terjawab, bayangan label memudar, dan teks penutup adalah layer komposisi terpisah."],
    "credits": ["coworkers"],
    "background": "Palet Opsi B. 4A dan 4C memakai cream, plum, dan terracotta; 4B dan 4B2 mempertahankan palette rapat; 4D memudarkan ruang menjadi cream dan plum. Terracotta hanya menandai label/kesalahan, sage dipakai untuk tanda pemahaman.",
    "frames": [
        {"code": "4A", "slug": "4a-self-misunderstanding", "title": "Raka Mislabels His Own Sadness", "time": "01:55–02:07", "voice": "Now, why should we care? First, misusing psychological terms can make us misunderstand our own emotions. Normal feelings like sadness, fear, or disappointment can feel like something is seriously wrong with us."},
        {"code": "4B", "slug": "4b-conflict-in-the-meeting", "title": "Misjudgment Becomes Conflict", "time": "02:07–02:12", "voice": "Second, it can create unnecessary conflict."},
        {"code": "4B2", "slug": "4b2-narcissistic-label", "title": "One Disagreement Becomes a Label", "time": "02:12–02:17", "voice": "Imagine labeling someone as narcissistic just because they disagree with you."},
        {"code": "4C", "slug": "4c-stigma-and-removal", "title": "Stigma Spreads Through the Group", "time": "02:17–02:27", "voice": "And third, it can reinforce stigma. Mental health conditions are complex experiences, not simply labels we can attach to someone based on one behavior."},
        {"code": "4D", "slug": "4d-dina-alone-closing-text", "title": "Dina Alone · Closing On-Screen Text", "time": "02:27–02:35", "voice": "Remember, psychology is more than just a collection of viral terms. It is a scientific discipline that requires context, evidence, and careful understanding."},
    ],
})


def tag(name: str) -> str:
    return f"{{{SVG_NS}}}{name}"


def frame_duration(frame):
    start, end = frame["time"].split("–")
    start_min, start_sec = map(int, start.split(":"))
    end_min, end_sec = map(int, end.split(":"))
    return end_min * 60 + end_sec - (start_min * 60 + start_sec)


def add(parent, name: str, **attrs):
    return ET.SubElement(parent, tag(name), {k.replace("_", "-"): str(v) for k, v in attrs.items()})


def rect(parent, x, y, w, h, fill, rx=0, stroke=None, sw=1, opacity=None):
    attrs = {"x": x, "y": y, "width": w, "height": h, "fill": fill, "rx": rx}
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw})
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "rect", **attrs)


def circle(parent, cx, cy, r, fill, opacity=None, stroke=None, sw=1):
    attrs = {"cx": cx, "cy": cy, "r": r, "fill": fill}
    if opacity is not None:
        attrs["opacity"] = opacity
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw})
    return add(parent, "circle", **attrs)


def line(parent, x1, y1, x2, y2, stroke, sw=2, dash=None, opacity=None):
    attrs = {"x1": x1, "y1": y1, "x2": x2, "y2": y2, "stroke": stroke, "stroke-width": sw, "stroke-linecap": "round"}
    if dash:
        attrs["stroke-dasharray"] = dash
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "line", **attrs)


def path(parent, d, fill="none", stroke=None, sw=2, opacity=None):
    attrs = {"d": d, "fill": fill}
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw, "stroke-linecap": "round", "stroke-linejoin": "round"})
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "path", **attrs)


def txt(parent, x, y, value, size=24, fill=INK, weight=400, anchor="start", letter=0, opacity=None):
    node = add(parent, "text", x=x, y=y, fill=fill, font_family="Arial, Helvetica, sans-serif", font_size=size,
               font_weight=weight, text_anchor=anchor, letter_spacing=letter)
    if opacity is not None:
        node.set("opacity", str(opacity))
    for i, row in enumerate(str(value).split("\n")):
        span = add(node, "tspan", x=x, dy=0 if i == 0 else size * 1.25)
        span.text = row
    return node


def inline_txt(parent, x, y, pieces, size=18, anchor="start"):
    node = add(parent, "text", x=x, y=y, fill=INK, font_family="Arial, Helvetica, sans-serif", font_size=size,
               font_weight=500, text_anchor=anchor)
    for value, color, bold in pieces:
        color = ESPRESSO if color == INK else TERRACOTTA if color in (BLUE, BLUE_DARK, CORAL) else color
        span = add(node, "tspan", fill=color, font_weight=700 if bold else 500)
        span.text = value
    return node


def draw_social_comment(parent, x, y, w, h, comment_id, handle, lines, accent, size=17):
    group = add(parent, "g", id=comment_id, role="group", aria_label=f"Comment by {handle}")
    rect(group, x, y, w, h, CREAM, rx=17, stroke=PEACH, sw=2)
    circle(group, x + 27, y + 27, 15, accent)
    txt(group, x + 27, y + 32, handle[1:2].upper(), 12, WHITE if accent in (PLUM, TERRACOTTA) else ESPRESSO, 700, "middle")
    txt(group, x + 52, y + 30, handle, 15, ESPRESSO, 700)
    for row, pieces in enumerate(lines):
        inline_txt(group, x + 52, y + 66 + row * 23, pieces, size)
    return group


def pill(parent, x, y, w, h, label, fill=WHITE, color=INK, size=22, stroke=None, weight=700, opacity=None):
    rect(parent, x, y, w, h, fill, rx=h / 2, stroke=stroke, sw=2, opacity=opacity)
    txt(parent, x + w / 2, y + h * 0.66, label, size, color, weight, "middle", opacity=opacity)


def card(parent, x, y, w, h, fill=WHITE, border="#E3E8ED", radius=22):
    rect(parent, x + 7, y + 10, w, h, ESPRESSO, rx=radius, opacity=0.12)
    return rect(parent, x, y, w, h, fill, rx=radius, stroke=border, sw=2)


def omit_source_drawables(group, indices):
    drawable_tags = {"path", "rect", "circle", "ellipse", "polygon", "polyline", "line"}
    omitted = set(indices)
    drawable_index = 0

    def visit(parent):
        nonlocal drawable_index
        for child in list(parent):
            if child.tag.rsplit("}", 1)[-1] in drawable_tags:
                if drawable_index in omitted:
                    parent.remove(child)
                drawable_index += 1
            else:
                visit(child)

    visit(group)


def embed_asset(parent, asset_key, x, y, w, h, omit_ids=(), include_ids=None, color_map=None,
                view_box=None, sad_expression=False, raka_expression=False, color_map_policy=None,
                child_omit_indices=None, student_typing_expression=False, color_map_by_group=None,
                drawable_omit_indices=None):
    info = ASSETS[asset_key]
    source = ET.parse(info["file"]).getroot()
    vb = view_box or source.get("viewBox", "0 0 500 500")
    source_label = "Storyset Rafiki · attribution required" if asset_key == "presenter" else "Storyset Rafiki"
    nested = add(parent, "svg", x=x, y=y, width=w, height=h, viewBox=vb, preserveAspectRatio="xMidYMid meet", overflow="hidden", aria_label=info["title"] + " illustration, " + source_label)
    copied = []
    child_omit_indices = child_omit_indices or {}
    for child in list(source):
        if include_ids is not None and child.get("id") not in include_ids:
            continue
        if include_ids is None and child.get("id") in omit_ids:
            continue
        clone = deepcopy(child)
        omitted = set(child_omit_indices.get(child.get("id"), ()))
        if omitted:
            clone[:] = [element for index, element in enumerate(list(clone)) if index not in omitted]
        if child.get("id") in (drawable_omit_indices or {}):
            omit_source_drawables(clone, drawable_omit_indices[child.get("id")])
        if student_typing_expression and child.get("id") == "Character":
            source_parts = list(clone)
            if len(source_parts) > 23:
                source_parts[19].set("d", "M303,231 Q307,230 311,232 L311,234 Q307,233 303,232 Z")
                source_parts[20].set("d", "M286,230 Q290,230 293,232 L293,234 Q289,232 286,232 Z")
                source_parts[21].set("d", "M305,239 Q308,236 311,239 Q309,242 305,240 Z")
                source_parts[21].set("style", "fill:#FFF6EC;stroke:#2A1F2D;stroke-width:0.8px")
                source_parts[22].set("d", "M286,241 Q289,238 292,241 Q289,244 286,242 Z")
                source_parts[22].set("style", "fill:#FFF6EC;stroke:#2A1F2D;stroke-width:0.8px")
                source_parts[23].set("d", "M298,255 L310,255")
                source_parts[23].set("style", "fill:none;stroke:#263238;stroke-width:2;stroke-linecap:round")
        copied.append(clone)
    mapping = {}
    prefix = "s" + re.sub(r"[^a-z0-9]", "", asset_key.lower()) + "-"
    for child in copied:
        for elem in child.iter():
            old = elem.get("id")
            if old:
                mapping[old] = prefix + old
    for child in copied:
        source_group_id = child.get("id")
        group_color_index = 0
        selected_color_map = dict(color_map or {})
        selected_color_map.update((color_map_by_group or {}).get(source_group_id, {}))
        for elem in child.iter():
            if sad_expression:
                d = elem.get("d", "")
                if d.startswith("M255.56,297.55"):
                    elem.set("d", "M253.8,299.2 Q257.1,294.5 260.5,299.2 L260.5,301.1 Q257.1,297.0 253.8,301.1 Z")
                elif d.startswith("M256.38,297.31") or d.startswith("M257.57,300.16"):
                    elem.set("style", "fill:none")
                elif d.startswith("M257.36,282.6"):
                    elem.set("d", "M257.36,282.9 Q259.2,280.8 261.1,282.3")
                elif d.startswith("M251.2,283.29"):
                    elem.set("d", "M247.9,284.8 Q249.8,282.1 252.2,282.7")
            if raka_expression and asset_key == "coworkers" and elem.get("d", "").startswith("M333.67,194.93"):
                elem.set("d", "M333.5,195.6 Q335.8,192.8 338.2,195.5")
                elem.set("style", "fill:none;stroke:#263238;stroke-width:1.5;stroke-linecap:round")
            if color_map_policy is None:
                apply_color_map = True
            else:
                start, count = color_map_policy.get(source_group_id, (0, 0))
                has_mapped_color = any(
                    re.search(re.escape(old), value, flags=re.IGNORECASE)
                    for old in selected_color_map
                    for value in elem.attrib.values()
                )
                apply_color_map = start <= group_color_index < start + count
                if has_mapped_color:
                    group_color_index += 1
            for key, value in list(elem.attrib.items()):
                for old, new in mapping.items():
                    value = value.replace("url(#" + old + ")", "url(#" + new + ")")
                    value = value.replace("#" + old, "#" + new)
                if apply_color_map:
                    for old, new in selected_color_map.items():
                        value = re.sub(re.escape(old), new, value, flags=re.IGNORECASE)
                elem.set(key, value)
            if elem.text:
                for old, new in mapping.items():
                    elem.text = elem.text.replace("#" + old, "#" + new)
            if elem.get("id") in mapping:
                elem.set("id", mapping[elem.get("id")])
        nested.append(child)
    return nested


def start_scene(scene, bg=CREAM, accent=PLUM, dark=False):
    root = clean_frame(
        f"Scene {scene['no']}: {scene['title']}",
        scene["visual"],
        bg,
    )
    root.find(tag("metadata")).text = (
        f"Scene {scene['no']} · {scene['title']} · Standalone editable animation frame. "
        "Only the scene background, source illustration, animated props, and audience-facing text are included; production-board labels remain outside the artwork."
    )
    return root


def draw_public_comment(root, comment, x, y, width=432, height=174):
    group = add(root, "g", id="public-comment-" + comment["id"], role="group",
                data_sequence=comment["sequence"], data_source_type=comment["kind"], opacity="0")
    rect(group, x + 7, y + 9, width, height, ESPRESSO, rx=20, opacity=0.32)
    rect(group, x, y, width, height, CREAM, rx=20, stroke=PEACH, sw=3)
    circle(group, x + 31, y + 32, 18, PLUM)
    circle(group, x + 31, y + 25, 6, PEACH)
    path(group, f"M{x+19},{y+41} Q{x+31},{y+27} {x+43},{y+41} Z", fill=PEACH)
    txt(group, x + 63, y + 38, comment["alias"].upper(), 15, PLUM, 700, letter=0.9)
    line(group, x + 22, y + 60, x + width - 22, y + 60, PEACH, 2)
    lines = comment["lines"]
    first_y = y + (110 if len(lines) == 2 else 125)
    for index, line_text in enumerate(lines):
        txt(group, x + width / 2, first_y + index * 36, line_text, 25, ESPRESSO, 600, "middle")
    delay = [3.0, 8.0, 13.0, 18.0][comment["sequence"] - 1]
    add(group, "animate", attributeName="opacity", values="0;1", begin=f"{delay}s", dur="0.25s", fill="freeze")
    return group


def build_scene_1(scene):
    root = clean_frame(
        "Scene 1 · The Hook",
        "One continuous plum-background scene. The woman from the Character object in People Creating Robot · Rafiki is isolated in a centered half-body crop; its robot, props, panels, and background are omitted. Four public-comment cards with pseudonyms appear at the sides.",
        PLUM,
    )
    root.set("data-frame", "01")
    root.find(tag("metadata")).text = (
        "Scene 1 is one continuous editable composition. Presenter: the Character object from People Creating Robot · Rafiki (Storyset, free with attribution). "
        "Within that 41-drawable object, six lower-body paths are cropped and six outfit paths recolored; 29/41 (70.7%) remain unchanged. "
        "All other source objects are omitted. There is no presenter frame or pointer; comment cards use pseudonyms and generic avatars. "
        "Public quote sources, access dates, and Storyset attribution are recorded outside this artwork in the storyboard source audit."
    )
    embed_asset(
        root, "presenter", 550, -80, 820, 1080,
        include_ids=("Character",),
        view_box="285 130 190 250",
        color_map={"#407BFF": CREAM},
        drawable_omit_indices={"Character": (4, 5, 6, 7, 8, 9)},
    )

    positions = [(36, 228), (1452, 228), (36, 622), (1452, 622)]
    for comment, (x, y) in zip(SCENE1_COMMENT_SOURCES, positions):
        draw_public_comment(root, comment, x, y)

    question = add(root, "g", id="closing-onscreen-question", data_motion="appears-before-phone-zoom")
    question.set("opacity", "0")
    rect(question, 566, 944, 788, 72, CREAM, rx=22)
    txt(question, 960, 991, "Are we using psychology terms correctly?", 27, PLUM, 700, "middle")
    add(question, "animate", attributeName="opacity", values="0;1", begin="22.5s", dur="0.6s", fill="freeze")
    return root


def build_scene_2(scene):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "scene-title scene-desc"})
    title_node = add(root, "title", id="scene-title")
    title_node.text = "Scene 2A: room, phone, and scrolling comments"
    desc_node = add(root, "desc", id="scene-desc")
    desc_node.text = "A cream living room with a TV and framed art at left, an espresso phone with a scrolling comment feed in the center, and a reclining man on a sage sofa beneath colored shelves at right."
    defs = add(root, "defs")
    tv_gradient = add(defs, "linearGradient", id="tv-screen", x1="0%", y1="0%", x2="100%", y2="100%")
    add(tv_gradient, "stop", offset="0%", stop_color=PLUM)
    add(tv_gradient, "stop", offset="100%", stop_color=ESPRESSO)
    feed_clip = add(defs, "clipPath", id="phone-feed-window", clipPathUnits="userSpaceOnUse")
    rect(feed_clip, 678, 203, 560, 667, CREAM, rx=12)

    backdrop = add(root, "g", id="room-background")
    rect(backdrop, 0, 0, W, H, CREAM)
    floor = add(root, "g", id="room-floor")
    rect(floor, 0, 862, W, 218, PEACH, opacity=0.38)
    line(floor, 0, 862, W, 862, PEACH, 3, opacity=0.8)
    ambience = add(root, "g", id="room-ambient")
    circle(ambience, 175, 170, 185, MUSTARD, opacity=0.10)
    circle(ambience, 1040, 845, 280, WHITE, opacity=0.34)
    circle(ambience, 1860, 540, 205, SAGE, opacity=0.10)

    # Left wall art and television form a simple, colored living-room setting.
    wall_art = add(root, "g", id="wall-art")
    rect(wall_art, 115, 130, 358, 226, ESPRESSO, rx=17)
    rect(wall_art, 132, 147, 324, 192, PEACH, rx=8)
    circle(wall_art, 370, 196, 32, MUSTARD)
    path(wall_art, "M145,307 L242,218 L310,282 L365,238 L447,314 L447,339 L145,339 Z", fill=SAGE)
    path(wall_art, "M145,319 L225,270 L281,315 L350,266 L447,326 L447,339 L145,339 Z", fill=PLUM)
    path(wall_art, "M145,322 L216,293 L276,324 L345,291 L447,329 L447,339 L145,339 Z", fill=TERRACOTTA)

    tv = add(root, "g", id="television")
    rect(tv, 67, 425, 460, 307, ESPRESSO, rx=24)
    rect(tv, 83, 441, 428, 270, "url(#tv-screen)", rx=13)
    circle(tv, 494, 454, 3, MUSTARD, opacity=0.8)
    path(tv, "M99,675 C215,591 284,614 371,527 C427,471 485,493 510,462 L510,711 L99,711 Z", fill=PLUM, opacity=0.55)
    path(tv, "M250,732 L385,732 L405,757 L230,757 Z", fill=ESPRESSO)
    console = add(root, "g", id="media-console")
    rect(console, 29, 758, 582, 91, PLUM, rx=13)
    rect(console, 53, 778, 534, 9, PEACH, rx=4)
    for index, (x, color) in enumerate([(64, MUSTARD), (239, SAGE), (414, TERRACOTTA)], 1):
        book = add(console, "g", id=f"console-book-{index}")
        rect(book, x, 796, 154, 39, color, rx=7)
    rect(console, 72, 849, 18, 26, ESPRESSO, rx=4)
    rect(console, 548, 849, 18, 26, ESPRESSO, rx=4)
    # A small plant and a few colored books keep the console from reading as a UI panel.
    console_plant = add(console, "g", id="console-plant")
    rect(console_plant, 540, 708, 60, 48, TERRACOTTA, rx=8)
    path(console_plant, "M570,709 C557,687 555,672 563,656 C577,674 579,690 570,709 Z", fill=SAGE)
    path(console_plant, "M568,708 C582,687 595,679 607,682 C600,700 587,710 568,708 Z", fill=SAGE)
    path(console_plant, "M570,707 C560,689 543,683 533,689 C541,704 553,711 570,707 Z", fill=SAGE)
    console_books = add(console, "g", id="console-books")
    for index, (x, y, width, height, color) in enumerate([
        (229, 726, 20, 31, TERRACOTTA), (253, 717, 23, 40, MUSTARD), (280, 723, 22, 34, PLUM)
    ], 1):
        book = add(console_books, "g", id=f"console-book-small-{index}")
        rect(book, x, y, width, height, color, rx=3)

    # Right-side shelves: colored books, a vase, and flowers sit above the original character.
    shelves = add(root, "g", id="shelf-set")
    upper_shelf = add(shelves, "g", id="shelf-upper")
    upper_rail = add(upper_shelf, "g", id="shelf-upper-rail")
    rect(upper_rail, 1287, 250, 538, 20, PLUM, rx=7)
    upper_books = add(upper_shelf, "g", id="shelf-upper-books")
    for index, (x, y, width, height, color) in enumerate([
        (1326, 213, 10, 38, ESPRESSO), (1345, 168, 26, 81, TERRACOTTA),
        (1376, 181, 23, 68, MUSTARD), (1405, 157, 31, 92, SAGE),
        (1442, 176, 22, 73, PEACH), (1470, 189, 27, 60, ESPRESSO),
        (1777, 213, 10, 38, ESPRESSO)
    ], 1):
        book = add(upper_books, "g", id=f"shelf-upper-item-{index}")
        rect(book, x, y, width, height, color, rx=5 if width > 15 else 4)
    upper_flower = add(upper_shelf, "g", id="shelf-upper-flower")
    rect(upper_flower, 1714, 207, 63, 42, MUSTARD, rx=8)
    path(upper_flower, "M1745,207 C1736,183 1739,162 1751,144 C1765,166 1764,187 1745,207 Z", fill=SAGE)
    path(upper_flower, "M1749,207 C1762,183 1780,174 1793,181 C1781,199 1767,207 1749,207 Z", fill=SAGE)
    petals = add(upper_flower, "g", id="shelf-upper-petals")
    circle(petals, 1753, 145, 9, TERRACOTTA)
    circle(petals, 1745, 151, 8, MUSTARD)
    circle(petals, 1761, 154, 8, PEACH)
    circle(petals, 1754, 158, 7, CREAM)

    lower_shelf = add(shelves, "g", id="shelf-lower")
    lower_rail = add(lower_shelf, "g", id="shelf-lower-rail")
    rect(lower_rail, 1314, 382, 500, 19, PLUM, rx=7)
    lower_books = add(lower_shelf, "g", id="shelf-lower-books")
    for index, (x, y, width, height, color) in enumerate([
        (1342, 342, 20, 38, SAGE), (1367, 324, 24, 56, TERRACOTTA),
        (1396, 335, 28, 45, PEACH), (1430, 314, 22, 66, MUSTARD),
        (1460, 330, 24, 50, SAGE)
    ], 1):
        book = add(lower_books, "g", id=f"shelf-lower-item-{index}")
        rect(book, x, y, width, height, color, rx=4)
    lower_flower = add(lower_shelf, "g", id="shelf-lower-flower")
    rect(lower_flower, 1705, 336, 69, 43, PLUM, rx=9)
    path(lower_flower, "M1739,336 C1729,316 1730,299 1740,282 C1756,304 1754,322 1739,336 Z", fill=SAGE)
    path(lower_flower, "M1741,335 C1753,316 1770,310 1781,316 C1771,332 1759,338 1741,335 Z", fill=SAGE)

    # Keep the Storyset man and sofa at the same display size; widen the room around them.
    sofa_shadow = add(root, "g", id="sofa-shadow")
    add(sofa_shadow, "ellipse", cx=1630, cy=866, rx=346, ry=47, fill=ESPRESSO, opacity=0.15)
    shared_map = {"#263238": ESPRESSO, "#B97964": PEACH, "#A24E3F": TERRACOTTA,
                  "#E0E0E0": CREAM, "#E6E6E6": PEACH, "#EBEBEB": CREAM,
                  "#F0F0F0": CREAM, "#F5F5F5": CREAM}
    sofa_scene = add(root, "g", id="reclining-sofa")
    embed_asset(sofa_scene, "feed", 1080, 150, 850, 850,
                include_ids={"freepik--Sofa--inject-2"}, color_map={**shared_map, "#407BFF": SAGE})
    character_scene = add(root, "g", id="phone-scroller")
    character_svg = embed_asset(character_scene, "feed", 1080, 150, 850, 850,
                                include_ids={"freepik--Character--inject-2"}, color_map={**shared_map, "#407BFF": PLUM})
    character = next((node for node in character_svg.iter() if node.get("id") == "sfeed-freepik--Character--inject-2"), None)
    source_pose = next(iter(character), None) if character is not None else None
    if source_pose is not None:
        source_parts = list(source_pose)
        source_pose[:] = []
        hand_group = None
        for index, source_part in enumerate(source_parts):
            if 4 <= index <= 8:
                if hand_group is None:
                    hand_group = add(source_pose, "g", id="phone-hand-motion", data_rig_part="scrolling-hand-and-phone")
                hand_group.append(source_part)
            else:
                source_pose.append(source_part)

    # The centered phone is the only place where the scrollable social feed appears.
    phone = add(root, "g", id="phone-shell")
    rect(phone, 664, 79, 622, 934, ESPRESSO, rx=58, opacity=0.22)
    rect(phone, 646, 61, 622, 934, ESPRESSO, rx=58)
    rect(phone, 666, 81, 582, 894, CREAM, rx=42)
    phone_header = add(phone, "g", id="phone-header")
    rect(phone_header, 889, 94, 136, 11, PLUM, rx=6)
    circle(phone_header, 1172, 129, 5, PLUM)
    circle(phone_header, 1187, 129, 5, PLUM)
    rect(phone_header, 1201, 121, 23, 15, "none", rx=3, stroke=PLUM, sw=2)
    rect(phone_header, 1206, 125, 13, 7, SAGE, rx=2)
    txt(phone_header, 704, 179, "Beranda", 23, ESPRESSO, 700)
    circle(phone_header, 1200, 167, 9, "none", stroke=PLUM, sw=2)
    path(phone_header, "M1207,174 L1214,181", stroke=PLUM, sw=2)
    line(phone_header, 686, 193, 1228, 193, PEACH, 2)
    comment_viewport = add(phone, "g", id="comment-viewport", clip_path="url(#phone-feed-window)")
    comment_feed = add(comment_viewport, "g", id="comment-feed")
    for i, comment in enumerate(SCENE2_COMMENTS[:4]):
        draw_social_comment(comment_feed, 686, 209 + i * 166, 542, 153, comment["id"], comment["handle"], comment["lines"], comment["accent"], size=18)
    # A narrow track cues the scroll animation without adding a storyboard caption.
    scroll_track = add(phone, "g", id="phone-scroll-track")
    rect(scroll_track, 1236, 218, 4, 624, PEACH, rx=2)
    scroll_thumb = add(scroll_track, "g", id="phone-scroll-thumb")
    rect(scroll_thumb, 1232, 310, 12, 154, PLUM, rx=6)
    phone_footer = add(phone, "g", id="phone-footer")
    line(phone_footer, 686, 889, 1228, 889, PEACH, 2)
    circle(phone_footer, 760, 926, 9, PLUM)
    circle(phone_footer, 905, 926, 9, MUSTARD)
    circle(phone_footer, 1050, 926, 9, MUSTARD)
    circle(phone_footer, 1195, 926, 9, MUSTARD)
    return root


def build_scene_2_components():
    root = ET.Element(tag("svg"), {"width": "560", "height": "1080", "viewBox": "0 0 560 1080", "role": "img", "aria-labelledby": "components-title components-desc"})
    title_node = add(root, "title", id="components-title")
    title_node.text = "Scene 2 scroll components: six additional social comments"
    desc_node = add(root, "desc", id="components-desc")
    desc_node.text = "Six separate editable comment groups, comment-05 through comment-10, ready to move through the phone during the scroll animation."
    for i, comment in enumerate(SCENE2_COMMENTS[4:]):
        draw_social_comment(root, 20, 14 + i * 176, 520, 153, comment["id"], comment["handle"], comment["lines"], comment["accent"], size=17)
    return root


def build_scene_2b(scene):
    import base64
    import random

    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}",
                                    "role": "img", "aria-labelledby": "scene-title scene-desc",
                                    "data-frame": "02B"})
    title_node = add(root, "title", id="scene-title")
    title_node.text = "Scene 2B · chaotic hashtag word cloud"
    desc_node = add(root, "desc", id="scene-desc")
    desc_node.text = "A dense, overlapping field of handwritten hashtags in varied sizes and rotations fills a deep-plum frame."
    metadata = add(root, "metadata")
    metadata.text = "56 editable Caveat hashtag texts, including six script terms and fifty social-media trigger terms. The bundled Caveat variable font is licensed under SIL OFL 1.1. Tags appear one by one with shrinking intervals and pile into this final full-screen state."
    rect(root, 0, 0, W, H, PLUM)
    defs = add(root, "defs")
    font_data = base64.b64encode((ASSETS_DIR / "Caveat[wght].ttf").read_bytes()).decode("ascii")
    style = add(defs, "style", type="text/css")
    style.text = "@font-face{font-family:Caveat;src:url(data:font/ttf;base64," + font_data + ") format('truetype');font-weight:400 700}"
    # Seeded coordinates keep this editable cloud stable from build to build.
    rng = random.Random(2402)
    everyday_tags = [
        "#redflag", "#healing", "#baper", "#overthinking", "#triggered", "#insecure",
        "#boundaries", "#manipulatif", "#selflove", "#cutoff", "#mentalhealth", "#ghosting",
        "#toxicrelationship", "#attachmentstyle", "#situationship", "#lovebombing", "#selfdiagnose",
        "#traumaresponse", "#innerchild", "#relationshipgoals", "#psychologytok", "#healingjourney",
        "#validasi", "#capekmental", "#sadboy", "#sadgirl", "#peoplepleaser", "#narsis",
        "#emosional", "#coping", "#anxiety", "#burnout", "#selfcare", "#healingera",
        "#redflagcheck", "#staystrong", "#relatable", "#fyp", "#curhat", "#overreacting",
        "#silenttreatment", "#selfdiagnosis", "#innerwork", "#emositidakstabil", "#cintasehat",
        "#toxicbehavior", "#emotionaldamage", "#marup", "#detachment", "#unhealed",
    ]
    tags = []
    # Peripheral tags arrive first and remain behind the script terms and three central heroes.
    for word in everyday_tags:
        # Most tags carry enough visual weight to close gaps. A few stay small,
        # so the largest script terms remain about 6–8x their size.
        size = rng.choices([26, 34, 44, 54, 64, 74], weights=[1, 2, 2, 3, 3, 2])[0]
        x = rng.randint(10, W - 10)
        y = rng.randint(58, H - 24)
        angle = rng.randint(-30, 30)
        color = rng.choices([CREAM, MUSTARD, PEACH], weights=[8, 1, 1])[0]
        tags.append((word, x, y, size, angle, color, False))
    script_terms = [
        ("#narcissist", 330, 760, 132, -18),
        ("#trauma", 1520, 392, 126, 17),
        ("#depressed", 1510, 872, 122, -11),
    ]
    for word, x, y, size, angle in script_terms:
        tags.append((word, x, y, size, angle, TERRACOTTA, True))
    # Three oversized center words form the cloud's 6–8x visual hierarchy.
    tags.extend([
        ("#toxic", 640, 493, 208, -12, TERRACOTTA, True),
        ("#gaslighting", 1195, 552, 176, 9, TERRACOTTA, True),
        ("#avoidant", 885, 735, 166, -19, TERRACOTTA, True),
    ])
    for index, (word, x, y, size, angle, color, is_script) in enumerate(tags):
        group = add(root, "g", id=f"word-cloud-tag-{index + 1:02d}",
                    transform=f"rotate({angle} {x} {y})", data_word=word,
                    data_order=index + 1, data_script_term=str(is_script).lower(), opacity="0")
        node = add(group, "text", x=x, y=y, fill=color, font_family="Caveat, cursive",
                   font_size=size, font_weight=700, text_anchor="middle",
                   letter_spacing="-0.035em")
        node.text = word
        # The shrinking gaps make the initial trickle accelerate into a dense pile.
        delay = 0.38 * (1 - 0.92 ** index) / (1 - 0.92)
        add(group, "animate", attributeName="opacity", values="0;1",
            begin=f"{delay:.2f}s", dur="0.10s", fill="freeze")
    return root

def build_scene_2c(scene):
    root = clean_frame(
        "Scene 2C: understanding psychology and misusing terms",
        "Centered sage pill: Understanding psychology = Good! Centered terracotta pill: Misusing psychological terms = Problematic.",
        CREAM,
    )
    understanding = add(root, "g", id="understanding-statement")
    rect(understanding, 385, 304, 1150, 148, SAGE, rx=30)
    txt(understanding, 960, 398, "Understanding psychology = Good!", 48, ESPRESSO, 700, "middle")
    misuse = add(root, "g", id="misuse-statement")
    rect(misuse, 270, 552, 1380, 176, TERRACOTTA, rx=30)
    txt(misuse, 960, 660, "Misusing psychological terms = Problematic.", 46, WHITE, 700, "middle")
    root.find(tag("metadata")).text = "Editable animation frame. Existing Storyset vector groups are recolored in SVG; all added elements are hand-edited vector shapes. No generated imagery."
    return root


def clean_frame(title, description, background):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "frame-title frame-desc"})
    title_node = add(root, "title", id="frame-title")
    title_node.text = title
    desc_node = add(root, "desc", id="frame-desc")
    desc_node.text = description
    metadata = add(root, "metadata")
    metadata.text = "Editable animation frame. Characters come from credited, downloadable source SVGs; backgrounds and overlays are separate editable vector layers. No generated imagery."
    rect(root, 0, 0, W, H, background)
    return root


SCENE3_CASES = [
    {
        "key": "texting", "background": CASE_TINTS[0], "quote": "They're avoidant!",
        "correction": ["Mungkin sibuk, lelah, atau", "belum sempat."],
        "colors": {"#2E353A": ESPRESSO, "#407BFF": SAGE, "#DE5753": TERRACOTTA, "#E4897B": PEACH,
                   "#E0E0E0": CREAM, "#E9E9E9": CREAM, "#EBEBEB": PEACH, "#F4F4F4": CREAM, "#F5F5F5": CASE_TINTS[0]},
    },
    {
        "key": "student_stress", "background": CASE_TINTS[1], "quote": "I'm depressed!",
        "correction": ["Sedih setelah gagal itu", "emosi yang normal."],
        "colors": {"#407BFF": PLUM, "#E6F0FC": CASE_TINTS[1], "#DCE3FB": CASE_TINTS[1]},
    },
    {
        "key": "couple", "background": CASE_TINTS[2], "quote": "You're gaslighting me!",
        "correction": ["Beda pendapat belum tentu", "gaslighting."],
        "colors": {"#263238": ESPRESSO, "#407BFF": SAGE, "#D58745": MUSTARD, "#D9877C": PEACH,
                   "#DE5753": TERRACOTTA, "#EBB376": PEACH, "#E0E0E0": CREAM, "#EBEBEB": CREAM,
                   "#F0F0F0": CASE_TINTS[2], "#F5F5F5": CREAM, "#FAFAFA": CASE_TINTS[2]},
    },
]


def draw_case_frame(case_index):
    case = SCENE3_CASES[case_index]
    number = case_index + 1
    description = (
        "Raka sits with his head bowed and both hands holding his head after missing a work target. A separate target progress marker cues the failure; the terracotta label says I'm depressed! and a sage correction says sadness after failure is a normal emotion."
        if case["key"] == "student_stress"
        else f"{case['quote']} A terracotta assumption appears beside the illustration, then a sage correction reads: {' '.join(case['correction'])}"
    )
    root = clean_frame(
        f"Scene 3 case {number}: {case['quote']}",
        description,
        case["background"],
    )
    if case["key"] == "student_stress":
        root.find(tag("metadata")).text = (
            "Student Stress · Rafiki supplies the study setting; Co-workers · Rafiki supplies the same Raka as Scene 4B and 4A. "
            "The target progress marker, mistaken label, and sage correction are independent editable overlays; no generated imagery."
        )
    color_map = case.get("colors", {})
    omit_bg = {f"freepik--background-complete--inject-2"}
    if case["key"] == "texting":
        embed_asset(root, "texting", 58, 78, 900, 900,
                    omit_ids=omit_bg | {"freepik--Chat--inject-2", "freepik--character-1--inject-2"},
                    color_map=color_map, sad_expression=True)
        # The recipient avatar belongs in the chat header; both outgoing messages share one side.
        rect(root, 245, 142, 570, 82, CREAM, rx=24, stroke=PEACH, sw=3)
        embed_asset(root, "texting", 260, 153, 58, 58,
                    include_ids={"freepik--character-1--inject-2"}, color_map=color_map,
                    view_box="75 111 58 64")
        txt(root, 338, 193, "Teman", 23, ESPRESSO, 700)
        rect(root, 335, 244, 480, 82, PLUM, rx=24)
        txt(root, 575, 296, "Hei, ada waktu?", 26, WHITE, 700, "middle")
        rect(root, 255, 344, 560, 88, PLUM, rx=24)
        txt(root, 535, 400, "Balas dong :( ", 26, WHITE, 700, "middle")
        txt(root, 800, 468, "Dibaca · 3 jam", 18, ESPRESSO, 700, "end")
    elif case["key"] == "student_stress":
        # Reuse the cleaner Raka source pose from 4A, keeping the target card and
        # case text as separate composition layers for animation.
        embed_raka_puppet(root, 0, 0, 1080, 1080, pose="hands-on-head")
        target = add(root, "g", id="missed-target-prop")
        rect(target, 116, 106, 390, 110, CREAM, rx=22, stroke=PEACH, sw=3)
        txt(target, 148, 151, "TARGET", 20, ESPRESSO, 700, letter=1.2)
        rect(target, 148, 172, 266, 15, CASE_TINTS[1], rx=8)
        rect(target, 148, 172, 154, 15, PLUM, rx=8)
        circle(target, 454, 161, 25, TERRACOTTA)
        path(target, "M444,151 L464,171 M464,151 L444,171", stroke=CREAM, sw=5)
    else:
        embed_asset(root, "couple", 70, 76, 910, 910, omit_ids=omit_bg, color_map=color_map)

    # Separate groups let the animator reveal the correction when the VO reaches “Padahal...”.
    mistaken_label = add(root, "g", id="wrong-label-bubble")
    rect(mistaken_label, 1054, 298, 770, 146, TERRACOTTA, rx=28)
    txt(mistaken_label, 1439, 390, case["quote"], 50 if number < 3 else 43, WHITE, 700, "middle")
    correction = add(root, "g", id="sage-correction-card")
    rect(correction, 1054, 574, 770, 178, SAGE, rx=28)
    txt(correction, 1112, 635, "REALITANYA", 20, ESPRESSO, 700, letter=1.4)
    txt(correction, 1439, 690, case["correction"][0], 30, ESPRESSO, 600, "middle")
    txt(correction, 1439, 731, case["correction"][1], 30, ESPRESSO, 600, "middle")
    return root


def build_scene_3_case_1(scene):
    return draw_case_frame(0)


def build_scene_3_case_2(scene):
    return draw_case_frame(1)


def build_scene_3_case_3(scene):
    return draw_case_frame(2)


def build_scene_3_combined(scene):
    root = clean_frame(
        "Scene 3 combined examples: one behavior does not define a diagnosis",
        "The unanswered message, sadness, and disagreement examples meet behind the centered on-screen statement: One behavior does not equal a psychological diagnosis.",
        CREAM,
    )
    panel_w = 640
    case_positions = [(0, CASE_TINTS[0]), (640, CASE_TINTS[1]), (1280, CASE_TINTS[2])]
    for i, (x, tint) in enumerate(case_positions):
        rect(root, x, 0, panel_w, H, tint)
        case = SCENE3_CASES[i]
        if case["key"] == "student_stress":
            embed_asset(root, "student_stress", x + 115, 12, 410, 410,
                        omit_ids={"speech-bubble"}, color_map=case["colors"])
        else:
            omit_bg = {"freepik--background-complete--inject-2"}
            embed_asset(root, case["key"], x + 70, 24, 500, 410, omit_ids=omit_bg, color_map=case["colors"])
    rect(root, 95, 433, 1730, 214, PLUM, rx=34)
    txt(root, 548, 564, "One behavior", 58, MUSTARD, 700, "middle")
    line(root, 925, 520, 987, 520, TERRACOTTA, 9)
    line(root, 925, 552, 987, 552, TERRACOTTA, 9)
    path(root, "M940,573 L972,501", stroke=TERRACOTTA, sw=9)
    txt(root, 1346, 562, "A psychological diagnosis.", 48, WHITE, 700, "middle")
    return root


def build_scene_3(scene):
    return build_scene_3_combined(scene)


def build_scene_4a(scene):
    root = clean_frame(
        "Scene 4A: normal emotions are misunderstood",
        "A sad person is surrounded by the ordinary feelings sadness, fear, and disappointment. Those feelings collapse into the self-label: Is something wrong with me?",
        CASE_TINTS[0],
    )
    root.find(tag("metadata")).text = "Storyset Feeling Blue Rafiki SVG; 11 of 282 original drawable elements changed (3.9%). Emotion chips and the self-label are separate editable animation groups. No generated imagery."
    root[3].set("id", "scene-background")
    art = add(root, "g", id="feeling-blue-illustration")
    embed_asset(art, "feeling_blue", 20, 12, 1060, 1056,
                color_map={"#407BFF": PLUM})
    emotions = add(root, "g", id="normal-emotions")
    pill(emotions, 1090, 302, 202, 58, "SEDIH", MUSTARD, ESPRESSO, 23)
    pill(emotions, 1310, 302, 202, 58, "TAKUT", PEACH, ESPRESSO, 23)
    pill(emotions, 1530, 302, 270, 58, "KECEWA", SAGE, ESPRESSO, 23)
    label = add(root, "g", id="self-misunderstanding-label")
    rect(label, 1082, 438, 735, 314, CREAM, rx=30, stroke=PEACH, sw=4)
    rect(label, 1082, 438, 18, 314, TERRACOTTA, rx=9)
    txt(label, 1144, 520, "PERASAAN NORMAL", 21, PLUM, 700, letter=1.2)
    txt(label, 1144, 610, "ADA YANG SALAH", 44, ESPRESSO, 700)
    txt(label, 1144, 674, "DENGAN AKU?", 44, TERRACOTTA, 700)
    return root


def build_scene_4b(scene):
    root = clean_frame(
        "Scene 4B: a label turns disagreement into conflict",
        "A couple disagree. A terracotta narcissist label lands between them and breaks the connection, showing the unnecessary conflict caused by labeling someone for disagreeing.",
        CASE_TINTS[2],
    )
    root.find(tag("metadata")).text = "Storyset Couple Stress Rafiki SVG; 43 of 242 original drawable elements recolored (17.8%). The label and broken connector are separate editable animation groups. No generated imagery."
    root[3].set("id", "scene-background")
    art = add(root, "g", id="couple-stress-illustration")
    couple_colors = {
        "#407BFF": PLUM,
        "#D9877C": PEACH,
        "#DE5753": TERRACOTTA,
        "#EBB376": MUSTARD,
        "#D58745": MUSTARD,
    }
    embed_asset(art, "couple", 0, 8, 1100, 1064, color_map=couple_colors)
    conflict = add(root, "g", id="disagreement-and-label")
    pill(conflict, 1162, 252, 592, 66, "BEDA PENDAPAT", SAGE, ESPRESSO, 25)
    path(conflict, "M1000,524 L1040,488 L1080,560 L1120,488 L1160,524", fill="none", stroke=TERRACOTTA, sw=9)
    rect(conflict, 1134, 378, 668, 208, TERRACOTTA, rx=30)
    txt(conflict, 1468, 468, "LANGSUNG DIBILANG", 20, WHITE, 700, "middle", letter=1.2)
    txt(conflict, 1468, 536, "NARCISSIST?", 47, WHITE, 700, "middle")
    return root


def build_scene_4c(scene):
    root = clean_frame(
        "Scene 4C: one label hides a complex experience",
        "One person is singled out by a group; a single-behavior label turns into stigma, while duration, context, impact, and support remain visible as parts of a more complex experience.",
        CASE_TINTS[1],
    )
    root.find(tag("metadata")).text = "Storyset Bullying Rafiki SVG; 48 of 165 original drawable elements changed or omitted (29.1%). Its profanity speech-bubble group is omitted; palette recoloring and context labels are separate composition layers. No generated imagery."
    root[3].set("id", "scene-background")
    art = add(root, "g", id="bullying-illustration")
    bullying_colors = {
        "#407BFF": PLUM,
        "#FF5652": TERRACOTTA,
        "#FF8B7B": PEACH,
        "#B55B52": PEACH,
        "#A02724": TERRACOTTA,
    }
    embed_asset(art, "bullying", 26, 48, 1200, 984,
                omit_ids={"chat-bubble"}, color_map=bullying_colors)
    label_flow = add(root, "g", id="single-behavior-to-stigma")
    pill(label_flow, 1194, 202, 560, 62, "SATU PERILAKU", TERRACOTTA, WHITE, 23)
    path(label_flow, "M1474,270 L1474,316", stroke=TERRACOTTA, sw=7)
    path(label_flow, "M1463,305 L1474,321 L1485,305 Z", fill=TERRACOTTA)
    rect(label_flow, 1194, 336, 560, 145, TERRACOTTA, rx=26)
    txt(label_flow, 1474, 428, "STIGMA", 43, CREAM, 700, "middle", letter=1.5)
    contexts = add(root, "g", id="context-pieces")
    pill(contexts, 1190, 570, 265, 68, "DURASI", CREAM, ESPRESSO, 20, stroke=PEACH)
    pill(contexts, 1480, 570, 285, 68, "KONTEKS", MUSTARD, ESPRESSO, 20)
    pill(contexts, 1190, 664, 265, 68, "DAMPAK", MUSTARD, ESPRESSO, 20)
    pill(contexts, 1480, 664, 285, 68, "DUKUNGAN", SAGE, ESPRESSO, 20)
    return root


def build_scene_4d(scene):
    root = clean_frame(
        "Scene 4D: context and evidence stop the label chain",
        "A researcher reviews charts, notes, and evidence. The final on-screen text reads: Psychology is more than viral terms; it needs context, evidence, and careful understanding.",
        CREAM,
    )
    root.find(tag("metadata")).text = "Storyset Research Paper Rafiki SVG; 75 of 238 original drawable elements recolored (31.5%). The final message is an independent editable overlay. No generated imagery."
    root[3].set("id", "scene-background")
    art = add(root, "g", id="research-paper-illustration")
    research_colors = {
        "#407BFF": PLUM,
        "#577AFA": PLUM,
        "#263238": ESPRESSO,
        "#2E353A": ESPRESSO,
    }
    embed_asset(art, "research", 0, 18, 900, 1044, color_map=research_colors)
    evidence = add(root, "g", id="evidence-terms")
    pill(evidence, 980, 240, 270, 72, "KONTEKS", SAGE, ESPRESSO, 22)
    pill(evidence, 1270, 240, 230, 72, "BUKTI", MUSTARD, ESPRESSO, 22)
    pill(evidence, 1520, 240, 280, 72, "PEMAHAMAN", PEACH, ESPRESSO, 21)
    closing = add(root, "g", id="final-onscreen-text")
    rect(closing, 940, 390, 850, 380, PLUM, rx=34)
    rect(closing, 940, 390, 18, 380, TERRACOTTA, rx=9)
    txt(closing, 1030, 498, "PSIKOLOGI LEBIH DARI", 25, MUSTARD, 700, letter=1)
    txt(closing, 1030, 557, "ISTILAH VIRAL", 42, CREAM, 700)
    line(closing, 1030, 596, 1690, 596, SAGE, 4)
    txt(closing, 1030, 650, "KONTEKS · BUKTI ·", 31, CREAM, 700)
    txt(closing, 1030, 704, "PEMAHAMAN YANG CERMAT", 28, CREAM, 700)
    return root


def build_scene_4(scene):
    return build_scene_4d(scene)


def continuous_office_base(title, description, frame_id, isolate_dina=False, character_color=PLUM, omit_dina=False):
    root = clean_frame(title, description, CREAM)
    if isolate_dina:
        root.find(tag("metadata")).text = (
            "Storyset Co-workers Rafiki. Conservative audit: 63 of 158 original drawable elements changed or omitted (39.9%): "
            "Raka's 54-element body group is counted as omitted from the office composition, eight blue elements in the "
            "background, chat, desk, and Dina's clothing are recolored plum, and Raka's mouth is adjusted in the separate face inset. "
            "At least 95 elements (60.1%) remain unchanged. Dina stands alone in the office; the Raka face inset and UI overlays are separate SVG layers."
        )
    elif omit_dina:
        root.find(tag("metadata")).text = (
            "Storyset Co-workers · Rafiki. The shared office base and Raka character match 4B; Dina's separate source character object is omitted in this frame. "
            "The target card and thought-label cloud are independent overlays. No character paths were redrawn."
        )
    else:
        root.find(tag("metadata")).text = (
            "Storyset Co-workers Rafiki. 36 of 158 original drawable elements changed or omitted (22.8%): "
            "28 blue elements recolored plum, 7 placeholder-chat elements omitted, and Raka's mouth adjusted. "
            "122 of 158 source elements (77.2%) remain unchanged. All dialogue and interface overlays are separate editable SVG groups."
        )
    root[3].set("id", "scene-background")
    if isolate_dina:
        embed_asset(
            root, "coworkers", 0, 74, 1000, 1000,
            omit_ids={"character-2"}, color_map={"#407BFF": PLUM},
            color_map_policy={"background-simple": (0, 1), "Chat": (0, 1), "Desk": (1, 3), "character-1": (6, 3)},
        )
    elif omit_dina:
        embed_asset(
            root, "coworkers", 0, 74, 1000, 1000,
            omit_ids={"Chat", "character-1"}, color_map={"#407BFF": character_color}, raka_expression=True,
        )
    else:
        embed_asset(
            root, "coworkers", 0, 74, 1000, 1000,
            omit_ids={"Chat"}, color_map={"#407BFF": character_color}, raka_expression=True,
        )
    root.set("data-frame", frame_id)
    return root


def embed_raka_puppet(parent, x, y, w, h, pose="hands-on-head", include_ids=None, character_offset=(0, 0)):
    """Reuse the Co-workers Raka from 4B and keep the scene furniture separate."""
    palette = {
        "#407BFF": PLUM, "#263238": ESPRESSO,
        "#EBC0BB": PEACH, "#EEC1BB": PEACH,
        "#D4827D": TERRACOTTA,
    }
    if include_ids is None:
        source_ids = {
            "background-complete", "background-simple", "Shadow", "Lamp", "Mugs",
            "Device", "Table", "books-2", "book-1", "paper-balls",
        }
    else:
        source_ids = set(include_ids) - {"Character", "speech-bubble"}
    background_ids = source_ids & {"background-complete", "background-simple", "Shadow"}
    foreground_ids = source_ids - background_ids
    if background_ids:
        embed_asset(parent, "student_stress", x, y, w, h,
                    include_ids=background_ids, color_map=palette)
    character_art = embed_asset(
        parent, "coworkers", x + character_offset[0], y + character_offset[1], w, h,
        include_ids={"character-2"}, color_map={"#407BFF": PLUM}, raka_expression=True,
    )
    character = next((node for node in character_art.iter() if node.get("id") == "scoworkers-character-2"), None)
    typing_overlay_parts = []
    if character is not None:
        parts = list(character)
        # Keep the source Raka from 4B recognizable while changing only the arm
        # paths needed for the adjacent scene pose. The handheld phone is not
        # carried into either the self-label or laptop-post scene.
        if pose == "hands-on-head":
            left_arm = set(range(19, 24))
            right_arm = {33, 34, 51, 53}
            for index in left_arm:
                parts[index].set("transform", "rotate(90 331.16 226)")
            for index in right_arm:
                parts[index].set("transform", "rotate(120 368.61 223.59)")
        elif pose == "typing":
            left_arm = set(range(19, 24))
            right_arm = {33, 34, 51, 53}
            for index in left_arm:
                parts[index].set(
                    "transform",
                    "translate(331.16 226) rotate(-42) scale(1.45) translate(-331.16 -226)",
                )
            for index in right_arm:
                parts[index].set(
                    "transform",
                    "translate(368.61 223.59) scale(2.6) translate(-368.61 -223.59)",
                )
            typing_overlay_parts = [parts[index] for index in sorted(left_arm | right_arm)]
        if pose in {"hands-on-head", "typing"}:
            character.remove(parts[52])
    if foreground_ids:
        foreground_art = embed_asset(parent, "student_stress", x, y, w, h,
                                     include_ids=foreground_ids, color_map=palette)
        if pose == "typing":
            # Lift the source desk and laptop to meet the reused Raka's hand pose.
            foreground_art.set("transform", "translate(0 -130)")
            # Keep Raka's hands above the lifted desk. They need their own layer
            # so they are not translated with the furniture.
            arms_art = add(
                parent, "svg", id="raka-typing-arms",
                x=x + character_offset[0], y=y + character_offset[1],
                width=w, height=h, viewBox="0 0 500 500",
                preserveAspectRatio="xMidYMid meet", overflow="hidden",
            )
            for part in typing_overlay_parts:
                character.remove(part)
                arms_art.append(part)
    return character_art


def build_story_03a(scene):
    root = continuous_office_base(
        "4A · Raka mislabels his sadness",
        "The same Co-workers office illustration as 4B shows Raka alone after Dina is removed from the composition. Separate target and thought-label overlays carry the 4A beat.",
        "4A",
        omit_dina=True,
    )

    target = add(root, "g", id="missed-target-card")
    rect(target, 1178, 114, 560, 142, PLUM, rx=22)
    txt(target, 1210, 155, "TARGET KERJA", 17, CREAM, 700, letter=1.0)
    txt(target, 1230, 222, "10", 41, CREAM, 700)
    txt(target, 1294, 217, "target", 15, CREAM, 500)
    txt(target, 1452, 222, "6", 41, TERRACOTTA, 700)
    txt(target, 1507, 217, "tercapai", 15, CREAM, 500)
    line(target, 1392, 172, 1392, 224, CREAM, 2, opacity=0.45)

    cloud = add(root, "g", id="raka-small-thought-cloud", data_role="animatable-thought-label")
    rect(cloud, 850, 324, 536, 154, PLUM, rx=56)
    for cx, cy, r in [(904, 330, 40), (973, 319, 48), (1052, 318, 49), (1134, 322, 45), (1213, 328, 37), (1280, 337, 26)]:
        circle(cloud, cx, cy, r, PLUM)
    for cx, cy, r in [(824, 485, 15), (792, 500, 10), (760, 515, 6)]:
        circle(cloud, cx, cy, r, PLUM)
    txt(cloud, 1118, 391, "Aku depresi. Ada yang salah", 24, CREAM, 700, "middle")
    txt(cloud, 1118, 426, "banget sama aku.", 24, CREAM, 700, "middle")
    txt(cloud, 1118, 458, "SEDIH BEBERAPA HARI", 13, PEACH, 700, "middle", letter=0.8)
    return root


def build_story_03b(scene):
    root = continuous_office_base(
        "04B · Raka and Dina argue in a work meeting",
        "At a work meeting, Raka hears Dina's feedback as a personal attack. Their two editable dialogue bubbles show both lines and keep the disagreement ambiguous.",
        "04B",
    )
    raka = add(root, "g", id="raka-quote")
    rect(raka, 1080, 70, 790, 232, PLUM, rx=28)
    txt(raka, 1124, 116, "RAKA", 17, CREAM, 700, letter=1.2)
    txt(raka, 1124, 185, "Kamu selalu bikin aku", 29, CREAM, 600)
    txt(raka, 1124, 229, "merasa salah!", 29, CREAM, 600)

    dina = add(root, "g", id="dina-quote")
    rect(dina, 1080, 394, 790, 242, CREAM, rx=28, stroke=PLUM, sw=4)
    txt(dina, 1124, 442, "DINA", 17, PLUM, 700, letter=1.2)
    txt(dina, 1124, 512, "Aku cuma kasih masukan.", 27, PLUM, 600)
    txt(dina, 1124, 558, "Jangan baper terus, deh.", 27, PLUM, 600)
    return root


def build_story_03c(scene):
    root = continuous_office_base(
        "4B2 · Raka labels Dina narcissistic",
        "Raka turns Dina's one defensive reply into the stamp NARSISTIK. The frame marks this as Raka's conclusion, based on one reply rather than evidence of a pattern.",
        "4B2",
    )
    thought = add(root, "g", id="raka-conclusion")
    pill(thought, 1100, 186, 700, 54, "KESIMPULAN RAKA · DARI SATU BALASAN", CREAM, PLUM, 18, stroke=PLUM)
    stamp = add(thought, "g", id="narcissistic-stamp", transform="rotate(-3 1450 440)")
    rect(stamp, 1094, 320, 716, 220, TERRACOTTA, rx=24)
    rect(stamp, 1113, 339, 678, 182, "none", rx=16, stroke=CREAM, sw=3)
    txt(stamp, 1452, 403, "DIA", 23, CREAM, 700, "middle", letter=1.4)
    txt(stamp, 1452, 474, "NARSISTIK", 52, CREAM, 700, "middle", letter=1.3)
    txt(thought, 1452, 622, "Satu balasan belum membuktikan pola.", 25, PLUM, 600, "middle")
    root.set("data-frame", "4B2")
    return root


def build_story_04a(scene):
    root = continuous_office_base(
        "4C · Stigma spreads through a post",
        "The same Co-workers office illustration as 4B shows Raka and Dina together. A warning post and group notifications appear as separate layers and begin to spread.",
        "4C",
    )
    root.find(tag("metadata")).text = (
        "The full Co-workers · Rafiki office illustration from 4B is reused unchanged, including Raka and Dina together. "
        "The warning post and notification chain are separate editable SVG overlays; no source character pose was changed."
    )
    zoom = add(root, "g", id="stigma-warning-post", data_motion="post-expands-into-notifications")
    rect(zoom, 1082, 132, 690, 356, PLUM, rx=28)
    rect(zoom, 1104, 154, 646, 312, ESPRESSO, rx=19)
    rect(zoom, 1104, 154, 646, 54, ESPRESSO, rx=19)
    txt(zoom, 1140, 190, "POSTINGAN", 16, PEACH, 700, letter=1.2)
    circle(zoom, 1708, 181, 7, TERRACOTTA)
    txt(zoom, 1140, 292, "Hati-hati sama Dina,", 34, CREAM, 700)
    txt(zoom, 1140, 341, "dia narsistik", 34, CREAM, 700)
    txt(zoom, 1140, 407, "Biar yang lain aman.", 19, PEACH, 600)

    notifications = add(root, "g", id="notification-chain", data_motion="notifications-spread-outward")
    notification_connector = path(notifications, "M1160,624 C1240,575 1305,575 1360,624 S1510,680 1590,624", fill="none", stroke=PEACH, sw=5)
    notification_connector.set("stroke-dasharray", "7 12")
    for index, (x, y, number, heading, message) in enumerate([
        (1045, 555, "1", "GRUP KERJA", "Postingan dibagikan"),
        (1300, 555, "2", "GRUP KERJA", "Dina dikeluarkan\ndari grup"),
        (1555, 555, "3", "GRUP KERJA", "Post dilihat 4 orang"),
    ]):
        card_group = add(notifications, "g", id=f"notification-{index + 1}")
        rect(card_group, x, y, 230, 142, CREAM, rx=18, stroke=PEACH, sw=3)
        circle(card_group, x + 28, y + 31, 13, PLUM)
        txt(card_group, x + 53, y + 37, heading, 14, PLUM, 700, letter=0.6)
        txt(card_group, x + 19, y + 83, message, 18, ESPRESSO, 700)
        circle(card_group, x + 207, y + 22, 11, TERRACOTTA)
        txt(card_group, x + 207, y + 27, number, 13, CREAM, 700, "middle")
    return root


def build_story_04b(scene):
    root = clean_frame(
        "4D · Dina alone · closing text",
        "Dina stands alone in a wide, faded field with a small empty speech bubble and unanswered group notifications. A fading NARSISTIK shadow sits behind her; the final on-screen text occupies the open space to her right.",
        CREAM,
    )
    root.set("data-frame", "4D")
    root.find(tag("metadata")).text = (
        "Storyset Co-workers · Rafiki Dina source. Raka's 54 source paths are omitted and three of 158 source elements are recolored (57/158 = 36.1% changed or omitted; 63.9% remain unchanged). "
        "The source room is washed with a cream overlay and the desk area is masked as separate editable composition layers. Dina is the only person; there is no table or Raka."
    )
    background = embed_asset(
        root, "coworkers", 0, 74, 1000, 1000,
        include_ids={"background-complete", "background-simple", "Shadow", "Chat", "Desk"},
    )
    background.set("opacity", "0.24")
    rect(root, 0, 74, 1000, 1000, CREAM, opacity=0.88)
    # Mask the inherited table zone, then place the source character back over it.
    rect(root, 200, 548, 790, 440, CREAM, rx=12)

    shadow = add(root, "g", id="dina-isolation-shadow")
    add(shadow, "ellipse", cx=394, cy=824, rx=200, ry=28, fill=ESPRESSO, opacity=0.12)
    fading_label = add(root, "g", id="fading-narcissistic-stamp", transform="rotate(-7 1000 390)")
    rect(fading_label, 810, 331, 470, 124, TERRACOTTA, rx=17, opacity=0.09)
    txt(fading_label, 1045, 411, "NARSISTIK", 48, PLUM, 700, "middle", letter=1.4, opacity=0.12)

    bubble = add(root, "g", id="dina-empty-speech-bubble")
    rect(bubble, 360, 144, 286, 134, CREAM, rx=36, stroke=PLUM, sw=3)
    path(bubble, "M407,276 L326,351 L501,278 Z", fill=CREAM, stroke=PLUM, sw=3)
    txt(bubble, 503, 229, "...", 58, PLUM, 700, "middle")

    notifications = add(root, "g", id="unanswered-group-notifications")
    for index, (x, y) in enumerate([(974, 208), (1210, 376), (970, 552)]):
        card_group = add(notifications, "g", id=f"unanswered-notification-{index + 1}", opacity=0.56 - index * 0.1)
        rect(card_group, x, y, 264, 92, CREAM, rx=18, stroke=PEACH, sw=2)
        circle(card_group, x + 30, y + 31, 11, SAGE)
        line(card_group, x + 54, y + 29, x + 205, y + 29, PLUM, 3, opacity=0.58)
        line(card_group, x + 24, y + 57, x + 201, y + 57, PLUM, 3, opacity=0.30)
        line(card_group, x + 24, y + 69, x + 132, y + 69, PLUM, 3, opacity=0.23)
        circle(card_group, x + 238, y + 21, 6, TERRACOTTA)

    embed_asset(
        root, "coworkers", 0, 74, 1000, 1000,
        include_ids={"character-1"}, color_map={"#407BFF": PLUM},
        color_map_policy={"character-1": (6, 3)},
    )
    ending = add(root, "g", id="final-onscreen-text")
    rect(ending, 710, 848, 988, 96, PLUM, rx=24)
    txt(ending, 1204, 909, "Don't reduce people to labels.", 34, CREAM, 700, "middle")
    return root

def build_scene_5(scene):
    root = start_scene(scene, CREAM, PLUM)
    card(root, 72, 106, 680, 854, CREAM, PEACH, 26)
    doc_colors = {"#263238": ESPRESSO, "#407BFF": PLUM, "#D58745": MUSTARD, "#EBB376": PEACH,
                  "#E0E0E0": PEACH, "#E6E6E6": CREAM, "#EBEBEB": CREAM, "#F0F0F0": CREAM,
                  "#F5F5F5": CREAM, "#FAFAFA": CREAM}
    research_colors = {**doc_colors, "#2E353A": ESPRESSO, "#577AFA": PLUM, "#D4827D": TERRACOTTA,
                       "#EBC0BB": PEACH, "#EEC1BB": PEACH}
    embed_asset(root, "document", 92, 122, 595, 650, color_map=doc_colors)
    embed_asset(root, "research", 390, 430, 326, 365, color_map=research_colors)
    pill(root, 125, 866, 555, 46, "KODE ETIK PSIKOLOGI INDONESIA", PLUM, MUSTARD, 15, stroke=PEACH)
    articles = [
        ("9", "Dasar ilmiah & profesional", "Keputusan psikologi berlandaskan pengetahuan yang teruji."),
        ("15", "Menghindari dampak buruk", "Sampaikan informasi dengan mempertimbangkan dampaknya."),
        ("28", "Pernyataan publik yang cermat", "Bijaksana, jujur, teliti, hati-hati, dan berbasis ilmiah."),
        ("31", "Komunikasi media bertanggung jawab", "Selaras dengan kode etik dan praktik psikologi yang tepat."),
    ]
    for i, (num, title, desc) in enumerate(articles):
        y = 112 + i * 174
        rect(root, 799, y + 8, 1032, 153, ESPRESSO, rx=18, opacity=0.10)
        rect(root, 792, y, 1032, 153, PLUM, rx=18, stroke=PEACH, sw=2)
        circle(root, 858, y + 76, 39, MUSTARD)
        txt(root, 858, y + 86, num, 27, ESPRESSO, 700, "middle")
        txt(root, 924, y + 58, "PASAL " + num, 14, MUSTARD, 700, letter=1)
        txt(root, 924, y + 92, title, 24, CREAM, 700)
        txt(root, 924, y + 124, desc, 17, CREAM, 400)
    pill(root, 792, 846, 1032, 62, "SCIENTIFIC BASIS  +  PROFESSIONAL CARE", PLUM, CREAM, 19, stroke=PEACH)
    return root


def build_scene_6(scene):
    root = start_scene(scene, CREAM, SAGE)
    card(root, 72, 130, 660, 760, CREAM, PEACH, 24)
    todo_colors = {"#263238": ESPRESSO, "#407BFF": SAGE, "#E6E6E6": CREAM, "#EBEBEB": CREAM,
                   "#F0F0F0": CREAM, "#F5F5F5": CREAM, "#FAFAFA": CREAM,
                   "#ED847E": PEACH, "#FFC3BD": PEACH}
    embed_asset(root, "todo", 90, 145, 620, 720, color_map=todo_colors)
    steps = [
        ("1", "Understand before you label", "Cari tahu makna istilah sebelum memakainya."),
        ("2", "Describe, don’t diagnose", "Jelaskan perilaku spesifik yang membuat tidak nyaman."),
        ("3", "Learn, don’t self-diagnose", "Konten edukasi bukan pengganti asesmen profesional."),
    ]
    for i, (num, title, desc) in enumerate(steps):
        y = 128 + i * 242
        card(root, 770, y, 1070, 206, CREAM, PEACH, 20)
        circle(root, 844, y + 81, 37, MUSTARD)
        txt(root, 844, y + 91, num, 26, ESPRESSO, 700, "middle")
        title_node = add(root, "text", x=910, y=y + 73, fill=PLUM, font_family="Arial, Helvetica, sans-serif", font_size=27, font_weight=700)
        split = [("Understand before you ", "label"), ("Describe, don’t ", "diagnose"), ("Learn, don’t ", "self-diagnose")][i]
        for value, color in [(split[0], PLUM), (split[1], TERRACOTTA)]:
            span = add(title_node, "tspan", fill=color)
            span.text = value
        txt(root, 910, y + 121, desc, 20, ESPRESSO, 400)
        # Small, editable line icons: learn, describe, and know the limit.
        if i == 0:
            circle(root, 1740, y + 79, 25, "none", stroke=SAGE, sw=5)
            line(root, 1758, y + 97, 1779, y + 118, SAGE, 6)
        elif i == 1:
            rect(root, 1700, y + 51, 93, 58, PEACH, rx=12, stroke=SAGE, sw=3)
            path(root, f"M1720,{y+78} L1734,{y+92} L1769,{y+64}", stroke=SAGE, sw=5)
        else:
            circle(root, 1744, y + 81, 29, CREAM, stroke=SAGE, sw=3)
            path(root, f"M1728,{y+82} L1740,{y+94} L1762,{y+67}", stroke=SAGE, sw=5)
    pill(root, 770, 895, 1070, 66, "Allow normal human emotions.", SAGE, ESPRESSO, 25)
    return root


def face(parent, cx, cy, emotion, color):
    circle(parent, cx, cy, 38, color, stroke=CREAM, sw=4)
    circle(parent, cx - 12, cy - 8, 3.5, ESPRESSO)
    circle(parent, cx + 12, cy - 8, 3.5, ESPRESSO)
    if emotion == "joy":
        path(parent, f"M{cx-15},{cy+7} Q{cx},{cy+25} {cx+15},{cy+7}", stroke=ESPRESSO, sw=3)
    elif emotion == "sadness":
        path(parent, f"M{cx-14},{cy+19} Q{cx},{cy+4} {cx+14},{cy+19}", stroke=ESPRESSO, sw=3)
        path(parent, f"M{cx+22},{cy-1} Q{cx+29},{cy+8} {cx+22},{cy+14} Q{cx+16},{cy+8} {cx+22},{cy-1}", fill=PEACH)
    elif emotion == "fear":
        circle(parent, cx, cy + 13, 5, "none", stroke=ESPRESSO, sw=3)
    else:
        line(parent, cx - 15, cy + 12, cx + 15, cy + 12, ESPRESSO, 3)


def build_scene_7(scene):
    root = start_scene(scene, PLUM, TERRACOTTA, dark=True)
    social_colors = {"#263238": ESPRESSO, "#407BFF": SAGE, "#630F0F": ESPRESSO,
                     "#7F3E3B": PEACH, "#A02724": PEACH, "#B55B52": PEACH,
                     "#E0E0E0": PLUM, "#E6E6E6": PLUM, "#EBEBEB": PLUM,
                     "#F0F0F0": PLUM, "#F5F5F5": PLUM, "#FAFAFA": PLUM,
                     "#ED847E": PEACH, "#FFC3BD": PEACH}
    embed_asset(root, "social", 52, 148, 870, 760,
                omit_ids={"freepik--background-complete--inject-2"}, color_map=social_colors)
    # Viral labels fade from terracotta toward the plum closing field.
    pill(root, 82, 118, 180, 40, "toxic", TERRACOTTA, WHITE, 15, opacity=0.58)
    pill(root, 268, 118, 204, 40, "trauma", TERRACOTTA, WHITE, 15, opacity=0.42)
    pill(root, 478, 118, 190, 40, "avoidant", TERRACOTTA, WHITE, 15, opacity=0.28)
    txt(root, 1370, 136, "NORMAL HUMAN EMOTIONS", 17, CREAM, 700, "middle", letter=1.2)
    emotions = [("joy", "Gembira", PEACH), ("sadness", "Sedih", SAGE), ("fear", "Takut", MUSTARD), ("calm", "Tenang", CREAM)]
    for i, (kind, label, color) in enumerate(emotions):
        x = 960 + i * 210
        face(root, x + 95, 230, kind, color)
        txt(root, x + 95, 304, label, 17, CREAM, 700, "middle")
    line(root, 980, 354, 1764, 354, ESPRESSO, 2)
    txt(root, 1250, 514, "THINK BEFORE YOU", 43, CREAM, 700, "middle", letter=0.6)
    txt(root, 1594, 514, "LABEL.", 43, MUSTARD, 700, "middle", letter=0.6)
    txt(root, 1373, 582, "Understand psychology. Use it responsibly.", 20, CREAM, 400, "middle")
    rect(root, 1020, 682, 710, 82, MUSTARD, rx=22)
    txt(root, 1375, 735, "Stay curious. Stay kind.", 30, ESPRESSO, 700, "middle")
    return root


DRAW = [build_scene_1, build_scene_2, build_scene_3, build_scene_4, build_scene_5, build_scene_6, build_scene_7]


def create_previews():
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    SCENES_DIR.mkdir(parents=True, exist_ok=True)
    for stale in [
        SCENES_DIR / "scene-01-the-hook.svg",
        PREVIEW_DIR / "scene-01-the-hook.png",
        SCENES_DIR / "scene-02-what-is-therapy-speak.svg",
        PREVIEW_DIR / "scene-02-what-is-therapy-speak.png",
        SCENES_DIR / "scene-03-normal-emotions-become-labels.svg",
        PREVIEW_DIR / "scene-03-normal-emotions-become-labels.png",
        SCENES_DIR / "scene-04-why-it-matters.svg",
        PREVIEW_DIR / "scene-04-why-it-matters.png",
        SCENES_DIR / "scene-03a-case-01-avoidant.svg",
        PREVIEW_DIR / "scene-03a-case-01-avoidant.png",
        SCENES_DIR / "scene-03a-raka-self-label.svg",
        PREVIEW_DIR / "scene-03a-raka-self-label.png",
        SCENES_DIR / "scene-03b-case-02-depressed.svg",
        PREVIEW_DIR / "scene-03b-case-02-depressed.png",
        SCENES_DIR / "scene-03b-raka-dina-conflict.svg",
        PREVIEW_DIR / "scene-03b-raka-dina-conflict.png",
        SCENES_DIR / "scene-03c-case-03-gaslighting.svg",
        PREVIEW_DIR / "scene-03c-case-03-gaslighting.png",
        SCENES_DIR / "scene-03c-raka-labels-dina.svg",
        PREVIEW_DIR / "scene-03c-raka-labels-dina.png",
        SCENES_DIR / "scene-03d-combined-on-screen-text.svg",
        PREVIEW_DIR / "scene-03d-combined-on-screen-text.png",
        SCENES_DIR / "scene-04a-emotions-misunderstood.svg",
        PREVIEW_DIR / "scene-04a-emotions-misunderstood.png",
        SCENES_DIR / "scene-04b-label-causes-conflict.svg",
        PREVIEW_DIR / "scene-04b-label-causes-conflict.png",
        SCENES_DIR / "scene-04c-one-label-becomes-stigma.svg",
        PREVIEW_DIR / "scene-04c-one-label-becomes-stigma.png",
        SCENES_DIR / "scene-04d-context-evidence-onscreen-text.svg",
        PREVIEW_DIR / "scene-04d-context-evidence-onscreen-text.png",
        SCENES_DIR / "scene-04a-label-becomes-stigma.svg",
        PREVIEW_DIR / "scene-04a-label-becomes-stigma.png",
        SCENES_DIR / "scene-04b-warning-and-removal.svg",
        PREVIEW_DIR / "scene-04b-warning-and-removal.png",
        SCENES_DIR / "scene-01a-presenter-placeholder.svg",
        PREVIEW_DIR / "scene-01a-presenter-placeholder.png",
        SCENES_DIR / "scene-01b-chat-unanswered.svg",
        PREVIEW_DIR / "scene-01b-chat-unanswered.png",
        SCENES_DIR / "scene-01c-chat-cancelled-plan.svg",
        PREVIEW_DIR / "scene-01c-chat-cancelled-plan.png",
        SCENES_DIR / "scene-01d-chat-sad-days.svg",
        PREVIEW_DIR / "scene-01d-chat-sad-days.png",
        SCENES_DIR / "scene-01e-chat-expectation-mismatch.svg",
        PREVIEW_DIR / "scene-01e-chat-expectation-mismatch.png",
        SCENES_DIR / "scene-01f-smartphone-transition.svg",
        PREVIEW_DIR / "scene-01f-smartphone-transition.png",
        SCENES_DIR / "scene-04a-self-misunderstanding.svg",
        PREVIEW_DIR / "scene-04a-self-misunderstanding.png",
        SCENES_DIR / "scene-04b-conflict-in-the-meeting.svg",
        PREVIEW_DIR / "scene-04b-conflict-in-the-meeting.png",
        SCENES_DIR / "scene-04c-narcissistic-label.svg",
        PREVIEW_DIR / "scene-04c-narcissistic-label.png",
        SCENES_DIR / "scene-04d-stigma-and-removal.svg",
        PREVIEW_DIR / "scene-04d-stigma-and-removal.png",
        SCENES_DIR / "scene-04e-dina-alone-closing-text.svg",
        PREVIEW_DIR / "scene-04e-dina-alone-closing-text.png",
    ]:
        if stale.exists():
            stale.unlink()
    thumbs = []
    for scene, draw_scene in zip(SCENES, DRAW):
        if scene["no"] == 1:
            frames = [
                ("01", "01-the-hook", "The Hook · One Continuous Scene", "00:00–00:30", build_scene_1),
            ]
        elif scene["no"] == 2:
            frames = [
                ("02A", "02a-comment-scroll", "Phone Zoom + Comment Scroll", "00:30–00:45", build_scene_2),
                ("02B", "02b-hashtag-cloud", "Hashtag Cloud Transition", "00:45–00:55", build_scene_2b),
                ("02C", "02c-on-screen-text", "Understanding Psychology · On-Screen Text", "00:55–01:05", build_scene_2c),
            ]
        elif scene["no"] == 3:
            frames = [
                ("03A", "03a-case-01-avoidant", "Case 1 · Unanswered Message", "01:05–01:20", build_scene_3_case_1),
                ("03B", "03b-case-02-depressed", "Case 2 · Sadness After Failure", "01:20–01:35", build_scene_3_case_2),
                ("03C", "03c-case-03-gaslighting", "Case 3 · Disagreement", "01:35–01:50", build_scene_3_case_3),
                ("03D", "03d-combined-on-screen-text", "One Behavior · Closing On-Screen Text", "01:50–01:55", build_scene_3_combined),
            ]
        elif scene["no"] == 4:
            frames = [
                ("4A", "4a-self-misunderstanding", "Raka Mislabels His Own Sadness", "01:55–02:07", build_story_03a),
                ("4B", "4b-conflict-in-the-meeting", "Misjudgment Becomes Conflict", "02:07–02:12", build_story_03b),
                ("4B2", "4b2-narcissistic-label", "One Disagreement Becomes a Label", "02:12–02:17", build_story_03c),
                ("4C", "4c-stigma-and-removal", "Stigma Spreads Through the Group", "02:17–02:27", build_story_04a),
                ("4D", "4d-dina-alone-closing-text", "Dina Alone · Closing On-Screen Text", "02:27–02:35", build_story_04b),
            ]
        else:
            frames = [(f"{scene['no']:02d}", f"{scene['no']:02d}-{scene['slug']}", scene["title"], scene["time"], draw_scene)]
        for code, slug, title, time, renderer in frames:
            frame_scene = dict(scene, title=title, time=time)
            root = renderer(frame_scene)
            ET.indent(root, space="  ")
            svg_path = SCENES_DIR / f"scene-{slug}.svg"
            svg_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
            png_path = PREVIEW_DIR / f"scene-{slug}.png"
            preview_root = deepcopy(root)
            if scene["no"] == 1 or (scene["no"] == 2 and code == "02B"):
                for node in preview_root.iter():
                    if node.get("id", "").startswith(("public-comment-", "word-cloud-tag-")) or node.get("id") == "closing-onscreen-question":
                        node.attrib.pop("opacity", None)
                    for child in list(node):
                        if child.tag == tag("animate"):
                            node.remove(child)
                cairosvg.svg2png(bytestring=ET.tostring(preview_root), write_to=str(png_path), output_width=1440, output_height=810)
            else:
                cairosvg.svg2png(url=str(svg_path), write_to=str(png_path), output_width=1440, output_height=810)
            thumbs.append((code, title, png_path))
    components = build_scene_2_components()
    ET.indent(components, space="  ")
    component_svg = SCENES_DIR / "scene-02-komponen.svg"
    component_svg.write_bytes(ET.tostring(components, encoding="utf-8", xml_declaration=True))
    component_png = PREVIEW_DIR / "scene-02-komponen.png"
    cairosvg.svg2png(url=str(component_svg), write_to=str(component_png), output_width=560, output_height=1080)
    transparent_preview = Image.open(component_png).convert("RGBA")
    preview_canvas = Image.new("RGB", transparent_preview.size, CREAM)
    preview_canvas.paste(transparent_preview, (0, 0), transparent_preview.getchannel("A"))
    preview_canvas.save(component_png)
    # Three columns make a compact visual audit sheet while retaining the 16:9 frames.
    cols = 5 if len(thumbs) > 12 else 4
    card_w, card_h, gap, top = (330, 186, 18, 50) if cols == 5 else (420, 236, 26, 56)
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * card_w + (cols + 1) * gap, rows * (card_h + 48) + (rows + 1) * gap + top), "#F1F4F8")
    draw = ImageDraw.Draw(sheet)
    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
    label_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    draw.text((gap, 15), f"STORYBOARD C · VISUAL AUDIT · 7 SCENES / {len(thumbs)} FRAMES", font=title_font, fill=ESPRESSO)
    for i, (code, title, png_path) in enumerate(thumbs):
        col, row = i % cols, i // cols
        x = gap + col * (card_w + gap)
        y = top + gap + row * (card_h + 48 + gap)
        im = Image.open(png_path).convert("RGB").resize((card_w, card_h), Image.Resampling.LANCZOS)
        sheet.paste(im, (x, y))
        draw.text((x, y + card_h + 9), f"{code}  ·  {title}", font=label_font, fill=ESPRESSO)
    contact = PREVIEW_DIR / "contact-sheet.png"
    sheet.save(contact, quality=94)
    return contact


def shade(cell, fill):
    props = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), fill.lstrip("#"))
    props.append(node)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement("w:" + name)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tc_pr.append(margins)


def add_hyperlink(paragraph, label, url):
    part = paragraph.part
    rid = part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), PLUM.lstrip("#"))
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    paragraph._p.append(link)
    return link


def set_run(run, size=10, color=INK, bold=False):
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))


def add_field(cell, label, content, size=10):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(label.upper() + "\n")
    set_run(r, 8, BLUE_DARK, True)
    r = p.add_run(content)
    set_run(r, size, INK)
    return p


def make_docx(contact):
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11.7), Inches(8.3)
    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.45)
    sec.left_margin = Inches(0.5)
    sec.right_margin = Inches(0.5)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(ESPRESSO.lstrip("#"))
    normal.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("THERAPY SPEAK GONE WRONG")
    set_run(r, 28, INK, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run("Storyboard produksi · Layout C · 7 scene · 16 keyframe · 00:00–04:00")
    set_run(r, 15, PLUM, True)
    p = doc.add_paragraph("Setiap frame tersedia sebagai SVG mandiri yang editable. Scene 1 adalah satu scene presenter dengan empat komentar publik anonim. Scene 3 tetap tiga contoh terpisah sesuai naskah. Scene 4 adalah satu cerita sebab-akibat lima beat sepanjang 40 detik dengan VO asli: salah memahami emosi → konflik → label narsistik → stigma → Dina sendirian. Semua frame adalah komposisi animasi; tidak ada judul produksi di dalam artwork dan tidak ada gambar generatif.")
    p.paragraph_format.space_after = Pt(10)
    for r in p.runs:
        set_run(r, 10, MUTED)

    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(4)
    r = heading.add_run("AUDIT VISUAL")
    set_run(r, 14, INK, True)
    audit = [
        "Scene 1 memakai objek Character dari People Creating Robot · Rafiki dalam crop setengah badan di tengah. Robot, dekorasi, panel, dan latar dihilangkan; tablet kecil yang menjadi bagian pose perempuan dipertahankan. Tidak ada pointer atau bingkai presenter. Audit per objek: dari 41 drawable Character, enam terpotong crop dan enam path pakaian biru diwarnai Cream; 29/41 (70,7%) tetap utuh. Kartu komentar memakai nama samaran Naya, Bagas, Rani, dan Dito; tiga kutipan bersumber publik dan satu bertanda tiruan.",
        "Scene 2B kini berupa word cloud Caveat 56 hashtag: enam istilah naskah dan lima puluh pemicu, dengan ukuran 26–208 px, rotasi acak, tumpang tindih, dan tiga hero term di tengah.",
        "Scene 3 memisahkan case avoidant, depresi, dan gaslighting. Label terracotta muncul lebih dulu; koreksi sage dikelompokkan terpisah untuk animasi saat VO mengatakan ‘Padahal...’. Latar case memakai peach, sage, dan mustard muda sesuai naskah.",
        "Scene 4 berlangsung 40 detik dalam urutan baru 4A, 4B, 4B2, 4C, 4D; VO tetap naskah. 4A memakai basis kantor Co-workers yang sama dengan karakter Dina dihilangkan; 4B, 4B2, dan 4C memakai ilustrasi utuh bersama Raka dan Dina. Panah/garis penghubung di 4B dihapus. Di 4C post peringatan menyebar sebagai notifikasi grup; 4D hanya menampilkan Dina dan teks ‘Don't reduce people to labels.’ Tidak ada objek domino di frame Scene 4.",
        "Semua frame mandiri memuat hanya komposisi animasi, latar, source illustration, overlay, properti, dan teks layar. Judul board, waktu, footer, legenda, kredit aset, dan catatan produksi ditempatkan di luar artwork.",
        "Raka Rig tersedia sebagai SVG terpisah dan halaman preview dengan kontrol joint: torso, kepala, bahu, siku, pergelangan, pinggul, lutut, dan pergelangan kaki. Rig hanya memakai Raka dari Co-workers · Rafiki; kursi dan gelas tidak disertakan. 14/42 drawable tubuh diadaptasi atau diwarnai ulang (33,3%); 28/42 tetap utuh.",
        "4A memakai basis Co-workers · Rafiki yang sama dengan objek karakter Dina dihilangkan. 4B, 4B2, dan 4C memakai ilustrasi dasar utuh dengan Raka dan Dina bersama. Semua panah/garis penghubung dialog pada 4B dihapus. Tidak ada objek domino di Scene 4. Dina di 4D tetap sendiri; 57/158 elemen drawable berubah atau dihilangkan (36,1%), sehingga 101/158 tetap utuh.",
        "Student Stress · Rafiki di Scene 3B mempertahankan 220/254 drawable elements (86,6%); 34 elemen (13,4%) berubah atau dihilangkan. Air mata, penanda target, label, dan koreksi adalah overlay terpisah.",
        "Semua karakter berasal dari SVG sumber yang dapat diunduh; chat, label, background, post, dan teks adalah layer vektor editable. Tidak ada gambar yang dihasilkan atau karakter yang digambar dari nol.",
    ]
    for item in audit:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(item)
        set_run(r, 9, INK)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    r = p.add_run("Warna sistem:   ")
    set_run(r, 9, INK, True)
    for name, color in zip(PALETTE_NAMES, PALETTE):
        r = p.add_run(f"   {name} {color}   ")
        set_run(r, 8, ESPRESSO, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("Objek Character dari People Creating Robot · Rafiki dipakai dengan atribusi Storyset. Pada objek Character, enam path bawah terpotong oleh framing setengah badan dan enam path pakaian diwarnai ulang; 29/41 elemen tetap utuh. Tablet yang terhubung dengan pose tangan dipertahankan. Caveat menggunakan SIL Open Font License 1.1; salinan lisensi ikut dalam source bundle. Aset Storyset gratis dengan atribusi; tautan sumber per ilustrasi dicantumkan pada scene terkait. Periksa ketentuan Storyset sebelum penggunaan final.")
    set_run(r, 9, MUTED)
    add_hyperlink(p, "People Creating Robot · Rafiki", "https://storyset.com/illustration/people-creating-robot/rafiki")
    p.add_run("  ·  ")
    add_hyperlink(p, "Caveat · SIL OFL", "https://github.com/google/fonts/blob/main/ofl/caveat/OFL.txt")
    p.add_run("  ·  ")
    add_hyperlink(p, "Storyset terms", "https://storyset.com/terms")

    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("CONTACT SHEET · AUDIT 7 SCENE / 16 FRAME")
    set_run(r, 20, INK, True)
    doc.add_picture(str(contact), width=Inches(10.65))
    p = doc.add_paragraph("Ritme visual bergerak dari presenter dan komentar publik → feed → tiga case terpisah pada Scene 3 → cerita sebab-akibat Raka dan Dina dalam lima beat → etika → checklist → closing. Scene 4 berakhir dengan Dina sendiri dan teks layar; seluruh frame mengikuti palet Opsi B.")
    for r in p.runs:
        set_run(r, 9, MUTED)

    for scene in SCENES:
        doc.add_page_break()
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(f"SCENE {scene['no']:02d}  ·  {scene['time']}")
        set_run(r, 11, PLUM, True)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(scene["title"])
        set_run(r, 21, INK, True)
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Inches(6.5)
        table.columns[1].width = Inches(4.05)
        left, right = table.rows[0].cells
        left.width, right.width = Inches(6.5), Inches(4.05)
        left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        shade(left, CREAM)
        shade(right, CREAM)
        set_cell_margins(left, 100, 100, 100, 100)
        set_cell_margins(right, 100, 140, 100, 140)
        if scene["no"] == 2:
            left.paragraphs[0].text = ""
            inner = left.add_table(rows=1, cols=3)
            inner.autofit = False
            for column in inner.columns:
                column.width = Inches(2.05)
            frame_data = [
                ("02A · 00:30–00:45", "scene-02a-comment-scroll.png"),
                ("02B · 00:45–00:55", "scene-02b-hashtag-cloud.png"),
                ("02C · 00:55–01:05", "scene-02c-on-screen-text.png"),
            ]
            for cell, (label, filename) in zip(inner.rows[0].cells, frame_data):
                cell.width = Inches(2.05)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                caption = cell.paragraphs[0]
                caption.alignment = 1
                caption.paragraph_format.space_after = Pt(5)
                run = caption.add_run(label)
                set_run(run, 7, PLUM, True)
                image_para = cell.add_paragraph()
                image_para.alignment = 1
                image_para.add_run().add_picture(str(PREVIEW_DIR / filename), width=Inches(1.96))
        elif scene["no"] in (1, 3, 4):
            left.paragraphs[0].text = ""
            frame_data = [
                (f"{frame['code']} · {frame['time']}", f"scene-{frame['slug']}.png")
                for frame in scene["frames"]
            ]
            grid_columns = 1 if scene["no"] == 1 else len(frame_data)
            grid_rows = (len(frame_data) + grid_columns - 1) // grid_columns
            inner = left.add_table(rows=grid_rows, cols=grid_columns)
            inner.autofit = False
            column_width = 6.28 / grid_columns
            for column in inner.columns:
                column.width = Inches(column_width)
            cells = [cell for row in inner.rows for cell in row.cells]
            for cell, (label, filename) in zip(cells, frame_data):
                cell.width = Inches(column_width)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                caption = cell.paragraphs[0]
                caption.alignment = 1
                caption.paragraph_format.space_after = Pt(3)
                run = caption.add_run(label)
                set_run(run, 7, PLUM, True)
                image_para = cell.add_paragraph()
                image_para.alignment = 1
                image_para.add_run().add_picture(str(PREVIEW_DIR / filename), width=Inches(column_width - 0.12))
        else:
            image = PREVIEW_DIR / f"scene-{scene['no']:02d}-{scene['slug']}.png"
            p = left.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(str(image), width=Inches(6.2))
        p = right.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("VISUAL")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["visual"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ON-SCREEN TEXT")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["screen"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK, True)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("GERAK / KAMERA")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["motion"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("SCRIPT / VOICE-OVER" + (" PER BEAT" if scene["no"] in (3, 4) else ""))
        set_run(r, 8, PLUM, True)
        if scene["no"] in (3, 4):
            for frame in scene["frames"]:
                if not frame["voice"]:
                    continue
                duration = frame_duration(frame)
                p = right.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(1)
                r = p.add_run(f"{frame['code']} · {frame['time']} ({duration} detik)")
                set_run(r, 8, PLUM, True)
                p = right.add_paragraph(frame["voice"])
                p.paragraph_format.space_after = Pt(3)
                for r in p.runs: set_run(r, 8, INK)
        else:
            p = right.add_paragraph(scene["voice"])
            p.paragraph_format.space_after = Pt(5)
            for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ASET & PERAN")
        set_run(r, 8, PLUM, True)
        for role in scene["roles"]:
            p = right.add_paragraph(role, style="List Bullet")
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs: set_run(r, 8, INK)
        for asset_key in scene["credits"]:
            info = ASSETS[asset_key]
            p = right.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            asset_credit = " · Storyset · free with attribution"
            add_hyperlink(p, "Sumber " + info["title"] + asset_credit, info["url"])
            for r in p.runs: set_run(r, 8, PLUM)
        p = right.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run("BACKGROUND: ")
        set_run(r, 8, PLUM, True)
        r = p.add_run(scene["background"])
        set_run(r, 8, MUTED)

    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("PUBLIC COMMENT SOURCE AUDIT · SCENE 1")
    set_run(r, 19, INK, True)
    p = doc.add_paragraph("Tanggal akses semua tautan: 26 November 2026. Nama di kartu adalah nama samaran, bukan nama akun sumber; foto profil dan merek platform tidak ditampilkan di artwork. Kutipan dibatasi pada potongan pendek. Komentar bertema narcissistic adalah tiruan karena tidak ditemukan kutipan publik yang aman dan sesuai.")
    for r in p.runs:
        set_run(r, 9, MUTED)
    source_table = doc.add_table(rows=1, cols=3)
    source_table.autofit = False
    for column, width in zip(source_table.columns, (Inches(2.15), Inches(3.7), Inches(4.6))):
        column.width = width
    for cell, heading_text in zip(source_table.rows[0].cells, ("KATA / PLATFORM", "POTONGAN KOMENTAR", "URL / STATUS / AKSES")):
        shade(cell, PLUM)
        set_cell_margins(cell, 100, 110, 100, 110)
        cell.text = heading_text
        for run in cell.paragraphs[0].runs:
            set_run(run, 8, CREAM, True)
    for comment in SCENE1_COMMENT_SOURCES:
        cells = source_table.add_row().cells
        for cell in cells:
            set_cell_margins(cell, 90, 110, 90, 110)
            shade(cell, CREAM)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cells[0].text = "Nama samaran: " + comment["alias"] + "\n" + comment["platform"]
        cells[1].text = "“" + comment["quote"] + "”"
        for run in cells[0].paragraphs[0].runs:
            set_run(run, 8, INK, True)
        for run in cells[1].paragraphs[0].runs:
            set_run(run, 9, ESPRESSO)
        if comment["url"]:
            add_hyperlink(cells[2].paragraphs[0], comment["url"], comment["url"])
            status_line = comment["kind"] + " · Tanggal akses: " + comment["accessed"]
        else:
            cells[2].paragraphs[0].add_run("Tidak ada URL publik · ")
            status_line = "TIRUAN (bukan komentar asli) · Tanggal akses: " + comment["accessed"]
        status_paragraph = cells[2].add_paragraph(status_line)
        for run in status_paragraph.runs:
            set_run(run, 8, TERRACOTTA if "TIRUAN" in status_line else MUTED, "TIRUAN" in status_line)

    out = ROOT / "Storyboard_Therapy_Speak_Gone_Wrong.docx"
    doc.save(out)
    return out


def make_website():
    rows = []
    for scene in SCENES:
        asset_links = []
        for key in scene["credits"]:
            info = ASSETS[key]
            asset_links.append(f'<a href="{info["url"]}" target="_blank" rel="noreferrer">{info["title"]} ↗</a>')
        role_text = "<br>".join(scene["roles"])
        comment_source_markup = ""
        if scene["no"] == 1:
            audit_items = []
            for item in SCENE1_COMMENT_SOURCES:
                url_markup = (f'<a href="{escape(item["url"])}" target="_blank" rel="noreferrer">{escape(item["url"])}</a>'
                              if item["url"] else "Tidak ada URL publik")
                audit_items.append(
                    f'<li><b>{escape(item["id"].upper())} · nama samaran: {escape(item["alias"])} · {escape(item["platform"])}</b> · '
                    f'“{escape(item["quote"])}”<br>{url_markup}<br>{escape(item["kind"])} · Tanggal akses: {escape(item["accessed"])}</li>'
                )
            comment_source_markup = '<div><dt>Public comment source audit</dt><dd><ul class="comment-source-list">' + "".join(audit_items) + '</ul></dd></div>'
        if scene["no"] == 2:
            frame_markup = '''<div class="frame-stack">
              <section class="frame-card"><div class="frame-label"><b>02A · PHONE + COMMENT SCROLL</b><span>00:30–00:45</span></div><a href="scenes/scene-02a-comment-scroll.svg" target="_blank" aria-label="Open editable SVG for Scene 2A"><img src="scenes/scene-02a-comment-scroll.svg" alt="Cream living room with TV and art at left, an espresso phone in the center, and the reclining man on a sage sofa at right"></a><div class="file-links"><a href="scenes/scene-02a-comment-scroll.svg" download>Download editable SVG</a><a href="previews/scene-02a-comment-scroll.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02B · HASHTAG CLOUD</b><span>00:45–00:55</span></div><a href="scenes/scene-02b-hashtag-cloud.svg" target="_blank" aria-label="Open editable SVG for Scene 2B"><img src="scenes/scene-02b-hashtag-cloud.svg" alt="Dense cream handwritten hashtags overlap across a deep plum frame"></a><div class="file-links"><a href="scenes/scene-02b-hashtag-cloud.svg" download>Download editable SVG</a><a href="previews/scene-02b-hashtag-cloud.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02C · ON-SCREEN TEXT</b><span>00:55–01:05</span></div><a href="scenes/scene-02c-on-screen-text.svg" target="_blank" aria-label="Open editable SVG for Scene 2C"><img src="scenes/scene-02c-on-screen-text.svg" alt="Centered sage and terracotta statements on a cream background"></a><div class="file-links"><a href="scenes/scene-02c-on-screen-text.svg" download>Download editable SVG</a><a href="previews/scene-02c-on-screen-text.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02 · SCROLL COMPONENTS</b><span>6 extra comments</span></div><p class="component-copy">The six separate SVG groups use IDs comment-05 through comment-10, ready to animate inside the phone feed.</p><div class="file-links"><a href="scenes/scene-02-komponen.svg" target="_blank">Open SVG components</a><a href="scenes/scene-02-komponen.svg" download>Download SVG</a><a href="previews/scene-02-komponen.png" target="_blank">Preview sheet</a></div></section>
            </div>'''
        elif scene["no"] in (1, 3, 4):
            frame_cards = []
            for frame in scene["frames"]:
                svg_name = f"scene-{frame['slug']}.svg"
                png_name = f"scene-{frame['slug']}.png"
                duration = frame_duration(frame)
                frame_cards.append(f'''<section class="frame-card"><div class="frame-label"><b>{frame['code']} · {frame['title']}</b><span>{frame['time']} · {duration}s</span></div><a href="scenes/{svg_name}" target="_blank" aria-label="Open editable SVG for {frame['code']}"><img src="scenes/{svg_name}" alt="{frame['title']} animation frame"></a><div class="file-links"><a href="scenes/{svg_name}" download>Download editable SVG</a><a href="previews/{png_name}" download>Download PNG preview</a></div></section>''')
            frame_markup = '<div class="case-frame-grid">' + "".join(frame_cards) + '</div>'
        else:
            svg_name = f"scene-{scene['no']:02d}-{scene['slug']}.svg"
            png_name = f"scene-{scene['no']:02d}-{scene['slug']}.png"
            frame_markup = f'<a href="scenes/{svg_name}" target="_blank" aria-label="Open editable SVG for scene {scene["no"]}"><img src="scenes/{svg_name}" alt="Editable scene illustration, scene {scene["no"]}"></a><div class="file-links"><a href="scenes/{svg_name}" download>Download editable SVG</a><a href="previews/{png_name}" download>Download PNG preview</a></div>'
        script_markup = escape(scene["voice"])
        if scene["no"] in (3, 4):
            script_markup = "<br>".join(
                f"<strong>{frame['code']} · {frame['time']} · {frame_duration(frame)} detik</strong><br>{escape(frame['voice'])}"
                for frame in scene["frames"] if frame["voice"]
            )
        rows.append(f'''<article class="board-row">
          <header class="row-meta"><span class="scene-number">{scene['no']:02d}</span><span class="time">{scene['time']}</span><h2>{scene['title']}</h2><p>{scene['visual']}</p></header>
          <div class="row-frame">{frame_markup}</div>
          <dl class="row-notes"><div><dt>On-screen text</dt><dd>{scene['screen'].replace(chr(10), '<br>')}</dd></div><div><dt>Motion / camera</dt><dd>{scene['motion']}</dd></div><div><dt>Script / voice-over</dt><dd>{script_markup}</dd></div><div><dt>Assets & roles</dt><dd>{role_text}</dd><p class="asset-sources">{' · '.join(asset_links) if asset_links else 'Custom editable SVG vectors; no external illustration.'}</p></div>{comment_source_markup}</dl>
        </article>''')
    html = '''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Editable production board with sixteen SVG animation frames for Therapy Speak Gone Wrong.">
<title>Therapy Speak Gone Wrong · Production Board</title>
<style>
:root{--ink:#2a1f2d;--plum:#3d2645;--sage:#7fb89a;--terracotta:#e4572e;--mustard:#f2b134;--peach:#ffc4a3;--cream:#fff6ec;--blue:#3d2645;--blue-dark:#3d2645;--paper:#fff6ec;--line:#e8dacf;--muted:#665967;--coral:#e4572e;--white:#fff;--shadow:0 12px 30px rgba(42,31,45,.09)}
*{box-sizing:border-box}html{background:var(--paper);color:var(--ink);font-family:Arial,Helvetica,sans-serif}body{margin:0;padding:42px 28px 92px}.shell{max-width:1560px;margin:0 auto}.hero{display:grid;grid-template-columns:1fr auto;gap:32px;align-items:end;margin:8px 0 32px}.eyebrow{color:var(--blue-dark);font-size:12px;font-weight:700;letter-spacing:.14em;text-transform:uppercase}.hero h1{max-width:920px;margin:12px 0 12px;font-size:clamp(38px,5.5vw,76px);line-height:.98;letter-spacing:-.055em}.hero p{max-width:760px;margin:0;color:var(--muted);font-size:16px;line-height:1.55}.hero-actions{display:flex;flex-direction:column;gap:9px;align-items:stretch}.action{display:inline-block;padding:12px 16px;border:1px solid var(--line);border-radius:999px;background:var(--white);color:var(--ink);font-size:13px;font-weight:700;text-align:center;text-decoration:none}.action.primary{border-color:var(--blue);background:var(--blue);color:var(--white)}.style-strip{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:28px 0 18px;padding:15px 18px;border:1px solid var(--line);border-radius:14px;background:var(--white);color:var(--muted);font-size:12px}.swatch{display:inline-flex;align-items:center;gap:6px}.swatch i{width:15px;height:15px;border:1px solid #d5dce2;border-radius:50%}.board{display:grid;gap:13px}.board-row{display:grid;grid-template-columns:minmax(0,1fr);gap:18px;align-items:stretch;padding:15px;border:1px solid var(--line);border-radius:17px;background:var(--white);box-shadow:var(--shadow)}.row-meta{display:flex;flex-direction:column;align-items:flex-start;padding:5px 0}.scene-number{display:grid;place-items:center;width:34px;height:34px;border-radius:50%;background:var(--ink);color:white;font-size:12px;font-weight:700}.time{margin:12px 0 9px;color:var(--blue-dark);font-size:11px;font-weight:700;letter-spacing:.08em}.row-meta h2{margin:0;font-size:19px;line-height:1.12;letter-spacing:-.02em}.row-meta p{margin:12px 0 0;color:var(--muted);font-size:12px;line-height:1.45}.row-frame{min-width:0;display:flex;flex-direction:column;justify-content:center}.row-frame img{display:block;width:100%;aspect-ratio:16/9;height:auto;object-fit:contain;object-position:center;border-radius:10px;background:#fff6ec}.frame-stack,.case-frame-grid{display:grid;grid-template-columns:minmax(0,1fr);gap:22px}.frame-card{padding:9px;border:1px solid var(--line);border-radius:12px;background:#fffaf5}.frame-label{display:flex;justify-content:space-between;gap:8px;margin-bottom:7px;color:var(--blue-dark);font-size:10px;letter-spacing:.04em}.frame-label span{color:var(--muted);font-weight:400}.component-copy{margin:0;color:var(--muted);font-size:11px;line-height:1.45}.file-links{display:flex;flex-wrap:wrap;justify-content:space-between;gap:10px;padding:9px 2px 0}.file-links a,.asset-sources a{color:var(--blue-dark);font-size:11px;font-weight:700;text-decoration:none}.file-links a:hover,.asset-sources a:hover{text-decoration:underline}.row-notes{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px;align-content:center;margin:0;padding:2px 4px}.row-notes div{min-width:0}.row-notes dt{margin-bottom:4px;color:var(--blue-dark);font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase}.row-notes dd{margin:0;color:#50616d;font-size:11px;line-height:1.4}.asset-sources{margin:7px 0 0}.asset-sources a{margin-right:6px}.comment-source-list{margin:4px 0 0;padding-left:17px}.comment-source-list li{margin:0 0 8px;color:#50616d;font-size:10px;line-height:1.45;overflow-wrap:anywhere}.comment-source-list a{color:var(--blue-dark);font-weight:700}.audit{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:28px 0;padding:20px;border:1px solid var(--line);border-radius:16px;background:#fff0e6}.audit h2{margin:0 0 8px;font-size:16px}.audit p{margin:0;color:#566974;font-size:12px;line-height:1.5}.footer{margin-top:24px;padding-top:16px;border-top:1px solid var(--line);color:var(--muted);font-size:11px;line-height:1.5}.footer a{color:var(--blue-dark)}
@media(max-width:1100px){.board-row{grid-template-columns:minmax(0,1fr)}.row-notes{grid-column:auto;grid-template-columns:1fr 1fr}.hero{grid-template-columns:1fr}.hero-actions{flex-direction:row;justify-content:flex-start}.action{min-width:180px}}
@media(max-width:700px){body{padding:24px 12px 60px}.hero h1{font-size:42px}.board-row{grid-template-columns:1fr;gap:12px;padding:13px}.row-meta{display:grid;grid-template-columns:40px 1fr;column-gap:8px;align-items:center}.scene-number{grid-row:span 2}.time{margin:0 0 4px}.row-meta h2{grid-column:2}.row-meta p{grid-column:1/-1;margin-top:10px}.row-notes{grid-column:auto;grid-template-columns:1fr}.hero-actions{flex-wrap:wrap}.action{min-width:0}.audit{grid-template-columns:1fr}}
</style>
</head>
<body><main class="shell">
<header class="hero"><div><div class="eyebrow">Editable production board · Layout C</div><h1>Therapy Speak Gone Wrong</h1><p>Seven scenes in sixteen animation keyframes. Scene 1 is one continuous presenter scene with four sourced, anonymized public comments. Scene 3 keeps three independent examples. Scene 4 follows one 40-second cause-and-effect story through five beats and ends with Dina alone. The narration stays verbatim from the script. Every frame is an editable SVG.</p></div><div class="hero-actions"><a class="action primary" href="Storyboard_Therapy_Speak_Gone_Wrong.docx" download>Download Word storyboard</a><a class="action" href="motion.html">Open Scene 2 motion study</a><a class="action" href="rigs/index.html">Open Raka character rig</a><a class="action" href="Storyboard_Editable_Source.zip" download>Download editable source bundle</a></div></header>
<div class="style-strip"><strong>Opsi B · Warm Peach &amp; Mustard</strong><span class="swatch"><i style="background:#3d2645"></i>deep plum</span><span class="swatch"><i style="background:#7fb89a"></i>sage</span><span class="swatch"><i style="background:#e4572e"></i>terracotta</span><span class="swatch"><i style="background:#f2b134"></i>mustard</span><span class="swatch"><i style="background:#ffc4a3"></i>peach</span><span class="swatch"><i style="background:#fff6ec"></i>cream</span><span class="swatch"><i style="background:#2a1f2d"></i>espresso</span><span>· terracotta for mistaken labels; sage for corrections</span></div>
<section class="board" aria-label="Seven scene production board">
''' + "\n".join(rows) + '''
</section>
<section class="audit"><div><h2>Visual audit</h2><p>Scene 3 remains three independent case frames and a centered closing frame. Scene 4A uses the Co-workers · Rafiki office base with Dina's character omitted; 4B, 4B2, and 4C use the full illustration with Raka and Dina together. Dialog connector arrows are removed from 4B. Beat-specific labels, the warning post, and notifications are separate overlays. Domino tiles have been removed from every Scene 4 frame. Frame 4D contains Dina alone. Dina's Co-workers source in 4D changes 57/158 drawable elements (36.1%), leaving 101/158 unchanged. Scene 1 uses only the Character object from People Creating Robot · Rafiki. Of its 41 drawable elements, six are cropped and six clothing paths recolored; 29/41 (70.7%) remain unchanged. The robot, external props, panels, background, presenter frame, and pointer are removed; the small tablet integrated with the woman’s pose remains. Four public-comment groups use pseudonyms. Scene 2B contains 56 independently editable handwritten hashtags (six script terms plus fifty related terms) with accelerating appearance timings. All retained frames were rendered and inspected.</p></div><div><h2>Editing &amp; attribution</h2><p>Open or download each SVG to edit its vector layers and text. The source bundle contains the HTML board, all SVG frames, source illustration assets, and the Caveat font/license. The <a href="https://storyset.com/illustration/people-creating-robot/rafiki" target="_blank" rel="noreferrer">People Creating Robot · Rafiki</a> presenter and <a href="https://storyset.com/illustration/co-workers/rafiki" target="_blank" rel="noreferrer">Co-workers · Rafiki</a> and <a href="https://storyset.com/illustration/student-stress/rafiki" target="_blank" rel="noreferrer">Student Stress · Rafiki</a> are free with attribution. <a href="https://github.com/google/fonts/tree/main/ofl/caveat" target="_blank" rel="noreferrer">Caveat</a> is distributed under SIL OFL 1.1; cite Storyset when publishing and review the <a href="https://storyset.com/terms" target="_blank" rel="noreferrer">Storyset terms</a>.</p></div></section>
<footer class="footer">Scene 1 isolates the woman Character object from the credited Storyset illustration in a centered half-body crop, without the robot, external props, pointer, or presenter frame. Her small held tablet stays as part of the source pose. Scene backgrounds, public-comment cards, cloud words, labels, cards, connectors, notifications, and text are editable SVG layers. No images were generated.</footer>
</main></body></html>'''
    output = ROOT / "index.html"
    output.write_text(html, encoding="utf-8")
    return output


def make_zip():
    import zipfile
    out = ROOT / "Storyboard_Editable_Source.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(ROOT / "index.html", "index.html")
        motion_page = ROOT / "motion.html"
        if motion_page.exists():
            archive.write(motion_page, "motion.html")
        archive.write(ROOT / "build_storyboard.py", "build_storyboard.py")
        archive.write(ROOT / "Storyboard_Therapy_Speak_Gone_Wrong.docx", "Storyboard_Therapy_Speak_Gone_Wrong.docx")
        readme = ROOT / "README.md"
        if readme.exists():
            archive.write(readme, "README.md")
        for path in sorted(SCENES_DIR.glob("*.svg")):
            archive.write(path, "scenes/" + path.name)
        for pattern in ("*.svg", "*.ttf", "*.txt"):
            for path in sorted(ASSETS_DIR.glob(pattern)):
                archive.write(path, "assets/" + path.name)
        for info in ASSETS.values():
            source_asset = info["file"]
            if source_asset.exists() and source_asset.parent != ASSETS_DIR:
                archive.write(source_asset, source_asset.name)
        for path in sorted(PREVIEW_DIR.glob("*.png")):
            archive.write(path, "previews/" + path.name)
        hyperframes = ROOT / "hyperframes"
        if hyperframes.exists():
            for path in sorted(hyperframes.rglob("*")):
                if not path.is_file() or ".hyperframes" in path.parts or "node_modules" in path.parts:
                    continue
                archive.write(path, "hyperframes/" + path.relative_to(hyperframes).as_posix())
        rigs = ROOT / "rigs"
        if rigs.exists():
            for path in sorted(rigs.rglob("*")):
                if not path.is_file() or "__pycache__" in path.parts:
                    continue
                archive.write(path, "rigs/" + path.relative_to(rigs).as_posix())
    return out


def main():
    SCENES_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    contact = create_previews()
    docx = make_docx(contact)
    website = make_website()
    source = make_zip()
    print("Generated:")
    for path in [contact, docx, website, source]:
        print(path.relative_to(ROOT), path.stat().st_size)
    print("Scene SVGs:", len(list(SCENES_DIR.glob("*.svg"))))


if __name__ == "__main__":
    main()
