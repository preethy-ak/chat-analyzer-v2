"""
Chat Analyzer Dashboard v2 — Shopee & Lazada
=============================================
Graas.ai-themed Streamlit app.
v2 adds: Sales Intelligence, Upsell Opportunities, OOS Tracker,
         Conversion Funnel, Merch/AM Performance, Lost Sales, Key Improvements.

Run:  streamlit run chat_analyzer_v2.py
Deps: pip install streamlit pandas openpyxl xlsxwriter
"""

import streamlit as st
import pandas as pd
import numpy as np
import re, io, warnings, gc
from datetime import datetime, timedelta
from collections import defaultdict

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Chat Analyzer v2 | Graas.ai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
.main { background: #F4F6FB; }
.block-container { padding: 1.5rem 2rem; }
.graas-header {
    background: linear-gradient(135deg, #1B2A4A 0%, #243554 100%);
    border-radius: 12px; padding: 1.2rem 1.8rem; margin-bottom: 1.5rem;
    display: flex; align-items: center; gap: 1rem;
}
.graas-header h1 { color: #fff; margin: 0; font-size: 1.5rem; font-weight: 700; }
.graas-header p  { color: #A8C0D6; margin: 0; font-size: 0.85rem; }
.graas-logo { color: #00C4B4; font-size: 2rem; }
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: #fff; border-radius: 10px; padding: 1rem 1.3rem;
    flex: 1; min-width: 150px; border-left: 4px solid #00C4B4;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-card.orange { border-left-color: #FF6B35; }
.metric-card.red    { border-left-color: #E74C3C; }
.metric-card.navy   { border-left-color: #1B2A4A; }
.metric-card.green  { border-left-color: #27AE60; }
.metric-card.purple { border-left-color: #8B5CF6; }
.metric-card.gold   { border-left-color: #F59E0B; }
.metric-val   { font-size: 1.9rem; font-weight: 800; color: #1B2A4A; }
.metric-label { font-size: 0.78rem; color: #7A8EA8; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-sub   { font-size: 0.75rem; color: #A0AEC0; margin-top: 2px; }
.section-title {
    font-size: 1rem; font-weight: 700; color: #1B2A4A;
    border-bottom: 2px solid #00C4B4; padding-bottom: 0.4rem; margin: 1.5rem 0 1rem;
}
.section-title-sales {
    font-size: 1rem; font-weight: 700; color: #1B2A4A;
    border-bottom: 2px solid #FF6B35; padding-bottom: 0.4rem; margin: 1.5rem 0 1rem;
}
.section-title-merch {
    font-size: 1rem; font-weight: 700; color: #1B2A4A;
    border-bottom: 2px solid #8B5CF6; padding-bottom: 0.4rem; margin: 1.5rem 0 1rem;
}
.badge-high   { background:#FDECEA; color:#C0392B; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-medium { background:#FEF9E7; color:#D68910; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-low    { background:#EAF4FB; color:#2980B9; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.sent-pos { color:#27AE60; font-weight:600; }
.sent-neu { color:#7F8C8D; font-weight:600; }
.sent-neg { color:#C0392B; font-weight:600; }
section[data-testid="stSidebar"] { background: #1B2A4A !important; }
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 { color: #00C4B4 !important; font-size: 1rem !important; font-weight: 700 !important; }
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown span { color: #FFFFFF !important; }
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox > label,
section[data-testid="stSidebar"] .stMultiSelect > label,
section[data-testid="stSidebar"] .stDateInput > label,
section[data-testid="stSidebar"] .stTextInput > label {
    color: #FFFFFF !important; font-size: 0.85rem !important; font-weight: 600 !important;
}
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div,
section[data-testid="stSidebar"] .stDateInput > div > div > input,
section[data-testid="stSidebar"] .stTextInput > div > div > input {
    background: #FFFFFF !important; color: #1B2A4A !important;
    border-radius: 6px !important; border: 1.5px solid #00C4B4 !important;
}
section[data-testid="stSidebar"] hr { border-color: #2E4A6A !important; }
section[data-testid="stSidebar"] strong { color: #00C4B4 !important; }
.stTabs [data-baseweb="tab-list"] { background: #fff; border-radius:8px; padding:4px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:6px; padding:6px 18px; font-weight:600; color:#7A8EA8; }
.stTabs [aria-selected="true"] { background:#00C4B4 !important; color:#fff !important; }
.reply-box {
    background: #F0FBF9; border: 1px solid #00C4B4; border-radius: 8px;
    padding: 0.9rem 1rem; font-size: 0.85rem; color: #1B2A4A; line-height: 1.6; margin-top: 0.5rem;
}
.reply-label { font-size:0.75rem; color:#00C4B4; font-weight:700; text-transform:uppercase; margin-bottom:4px; }
.sales-box {
    background: #FFF8F0; border: 1px solid #FF6B35; border-radius: 8px;
    padding: 0.9rem 1rem; font-size: 0.85rem; color: #1B2A4A; line-height: 1.6; margin-top: 0.5rem;
}
.oos-box {
    background: #FEF2F2; border: 1px solid #E74C3C; border-radius: 8px;
    padding: 0.9rem 1rem; font-size: 0.85rem; color: #C0392B; line-height: 1.6;
}
.upsell-box {
    background: #F0FDF4; border: 1px solid #27AE60; border-radius: 8px;
    padding: 0.9rem 1rem; font-size: 0.85rem; color: #1B2A4A; line-height: 1.6;
}
.upload-area {
    background: #fff; border: 2px dashed #00C4B4; border-radius: 12px;
    padding: 2rem; text-align: center; margin-bottom: 1rem;
}
.improvement-box {
    background: #F5F3FF; border-left: 4px solid #8B5CF6; border-radius: 6px;
    padding: 0.8rem 1rem; margin-bottom: 0.6rem; font-size: 0.85rem; color: #1B2A4A;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS — OPERATIONS
# ─────────────────────────────────────────────────────────────────────────────

ISSUE_KEYWORDS = {
    "Refund": ["refund","คืนเงิน","pengembalian dana","dana kembali","ibalik","irefund","bayar balik","money back","reimburse"],
    "Return": ["return","คืนสินค้า","retur","rma","send back","ส่งคืน","kembalikan","return item","product return"],
    "Cancellation": ["cancel","cancelled","ยกเลิก","batalkan","batal","cancellation","cancel order"],
    "Delay": ["delay","late","slow","ช้า","lambat","belum sampai","haven't received","not arrived","waiting","รอนาน","still waiting","lama","terlambat","not delivered yet","ยังไม่ได้รับ"],
    "Damaged/Wrong Item": ["wrong item","wrong product","damaged","broken","defective","สินค้าผิด","ของเสีย","ของแตก","rusak","cacat","salah barang","not as described","wrong size","wrong colour","wrong color"],
    "Missing Item": ["missing","not received","didn't receive","never received","ไม่ได้รับ","ของหาย","hilang","tidak diterima","kurang","incomplete","item missing"],
    "Payment Issue": ["payment","ชำระเงิน","bayar","pembayaran","charge","double charge","overcharged","billing","invoice","โอนเงิน","จ่ายเงิน","pay","transfer"],
    "Product Inquiry": ["how to","how do","วิธีใช้","ราคา","price","size","ขนาด","สี","colour","color","spec","specification","ingredient","cara pakai","ukuran","warna","harga","stok","stock","available","variant","model","version"],
    "Promotion Issue": ["voucher","promo","discount","coupon","code","sale","offer","โปรโมชั่น","ส่วนลด","โค้ด","diskon","kode promo","cashback","flash sale","deal","bundle"],
    "Technical Issue": ["error","bug","cannot","can't","unable","failed","not working","app issue","website","login","checkout problem","system","ไม่สามารถ","เกิดข้อผิดพลาด","tidak bisa","gagal"],
    "Complaint": ["complain","complaint","terrible","horrible","awful","worst","ร้องเรียน","ไม่พอใจ","รำคาญ","โกรธ","disappointed","frustrated","unacceptable","poor service","bad service","kecewa","tidak puas","buruk"],
}

PRIORITY_MAP = {
    "High":   ["Refund","Complaint","Damaged/Wrong Item"],
    "Medium": ["Delay","Missing Item","Return","Cancellation"],
    "Low":    ["Product Inquiry","Promotion Issue","Payment Issue","Technical Issue"],
}

TEAM_ASSIGNMENTS = {
    "Yeria":      ["AACMH","FFH","IKU","GED MY","GEDMY","GED_MY"],
    "Syahira":    ["EWG","HFC","AAISS","GED SG","GEDSG","GED_SG"],
    "Keerthana":  ["AABIY","AABIW","AAFTP","GED PH","GEDPH","GED_PH"],
    "Alfian":     ["AADMJ","AAEDD","AADWP","IGZ ID","IGZID"],
    "Jaye":       ["GSK","DBC","IEI","FYW","ILL"],
    "Ratchakorn": ["AABWU","AAFHU","AAFHB"],
}

STORE_TO_AGENT = {
    store.upper(): agent
    for agent, stores in TEAM_ASSIGNMENTS.items()
    for store in stores
}

AGENT_SHIFT = {
    "Yeria":      "GED MY · AACMH / FFH / IKU",
    "Syahira":    "GED SG · EWG / HFC / AAISS",
    "Keerthana":  "GED PH · AABIY / AABIW / AAFTP",
    "Alfian":     "IGZ ID · AADMJ / AAEDD / AADWP",
    "Jaye":       "GSK / DBC / IEI / FYW / ILL",
    "Ratchakorn": "Full-time · AABWU / AAFHU / AAFHB",
}

STALLING_PATTERNS = [
    r"will (check|look|get back|follow up|investigate|verify|review|update)",
    r"let me (check|look into|verify|confirm|see)",
    r"(checking|looking into|investigating|following up|reviewing)",
    r"please (wait|hold on|allow us|bear with)",
    r"i will (check|get back|follow up|update)",
    r"we (are|will) (checking|looking|investigating|getting back|following up)",
    r"get back to you", r"bear with us", r"kindly (wait|allow|hold)",
    r"we'?ll? (check|look|get back|follow up)",
    r"akan (kami|segera) (cek|periksa|tindak lanjut|proses|hubungi)",
    r"mohon (tunggu|ditunggu|bersabar)",
    r"kami (sedang|akan) (cek|periksa|proses|tindak lanjut)",
    r"จะตรวจสอบ",r"กำลังตรวจสอบ",r"จะแจ้งกลับ",r"จะดำเนินการ",r"ขอตรวจสอบ",r"ขอเวลา",
    r"จะติดต่อกลับ",r"ติดตามให้",r"กำลังประสานงาน",r"escalat",
]

RESOLUTION_PATTERNS = [
    r"refund (has been|was|is) (processed|completed|done|issued|approved)",
    r"(your|the) (order|item|package) (has been|was|is) (shipped|dispatched|replaced|delivered)",
    r"(issue|problem|case) (has been|was|is) (resolved|fixed|closed|sorted|handled)",
    r"(cancellation|cancel) (has been|was|is) (processed|done|completed|approved)",
    r"(we have|we've) (processed|completed|resolved|fixed|issued|sent)",
    r"please (expect|allow) (\d|few|some|a couple)",
    r"track.*link.*sent",r"tracking (number|id|code) (is|was|has been)",
    r"you (should|will) (receive|get) (it|your order|the item)",
    r"ดำเนินการเรียบร้อย",r"จัดการเรียบร้อย",r"แก้ไขเรียบร้อย",r"คืนเงินเรียบร้อย",r"ยกเลิกเรียบร้อย",
    r"sudah (diproses|selesai|dikirim|dikembalikan|dibatalkan)",
    r"telah (diproses|selesai|diselesaikan|dikirimkan)",
]

AUTO_REPLY_PATTERNS = [
    r"(thank you for contacting|thanks for reaching out).*auto",
    r"auto.?reply",r"automated (response|message|reply)",
    r"we'?ll? (get back|respond) (to you )?(within|in|shortly|soon)",
    r"our (team|agent).*(will|shall) (respond|reply|contact)",
    r"welcome to .*(official store|store).*\nhow (can|may) (we|i) help",
    r"สวัสดีค่ะ.*แอดมิน.*ยินดีให้บริการ",r"ยินดีต้อนรับ.*ร้าน",r"hi.{0,30}welcome to.{0,40}store",
]

POSITIVE_KWS = [
    "thank","thanks","great","excellent","awesome","perfect","love","good","nice","happy",
    "satisfied","wonderful","amazing","fantastic","superb","appreciate","helpful","fast","quick",
    "well done","recommend","ขอบคุณ","ดีมาก","ประทับใจ","พอใจ","ยอดเยี่ยม","ดีเลย","ดีค่ะ","ดีครับ",
    "terima kasih","bagus","mantap","keren","memuaskan","puas","salamat","maganda","ayos","galing",
]

NEGATIVE_KWS = [
    "terrible","worst","angry","disappointed","frustrated","cheated","scam","fraud","fake",
    "broken","damaged","wrong item","missing","never received","unacceptable","horrible","awful",
    "complain","complaint","refund","ผิดหวัง","โกรธ","ไม่พอใจ","แย่มาก","แย่","หลอกลวง","ของเสีย",
    "ของปลอม","ช้ามาก","รอนาน","สินค้าไม่ตรง","ไม่ได้รับ","ชำรุด","tipu","rusak","cacat",
    "mengecewakan","marah","kecewa","buruk","parah","salah","tidak diterima","hilang",
]

# ── SALES INTELLIGENCE CONSTANTS ─────────────────────────────────────────────

CONVERSION_KEYWORDS = [
    "i want to buy","i'd like to buy","i would like to buy","how to buy","how to order",
    "how do i order","place an order","can i order","add to cart","how to purchase",
    "i want to purchase","proceed to checkout","ready to buy","i'll take it","i want this",
    "i'll buy","i want to get","interested to buy","interested in buying","want to order",
    "อยากสั่ง","สั่งซื้อ","จะซื้อ","ซื้อ","สนใจซื้อ","จะสั่ง",
    "mau beli","mau order","mau pesan","ingin beli","ingin order","cara beli",
    "mag-order","gusto kong bilhin","bibilhin ko","paano mag-order",
]

UPSELL_KEYWORDS = [
    "similar","other option","alternative","recommend","suggestion","bundle","combo",
    "what else","anything else","other product","related","go with","pair with",
    "ตัวอื่น","แนะนำ","ตัวไหนดี","อะไรดี","มีอะไรอีก","คล้ายกัน",
    "yang lain","alternatif","rekomendasi","produk lain","pilihan lain",
    "iba pa","ano pa","katulad","mas maganda",
]

OOS_KEYWORDS = [
    "out of stock","no stock","sold out","not available","unavailable","habis","stok habis",
    "tidak tersedia","kehabisan","out of stock","หมดสต็อก","สินค้าหมด","ไม่มีสินค้า","ไม่มีของ",
    "wala na","wala nang stock","ubos na","hindi available",
    "หมดแล้ว","ไม่มีแล้ว","หมด","ไม่มี stock",
]

OOS_SELLER_PATTERNS = [
    r"(out of stock|sold out|no stock|not available|unavailable)",
    r"(habis|stok habis|tidak tersedia|kehabisan)",
    r"(หมดสต็อก|สินค้าหมด|ไม่มีสินค้า|หมดแล้ว|ไม่มีแล้ว)",
    r"(wala na|wala nang stock|ubos na|hindi available)",
    r"(currently (not|out of|no) (stock|available|inventory))",
    r"(restock|back in stock|will be available)",
    r"(ยังไม่มี|ยังไม่ได้|จะมีเร็วๆนี้)",
    r"(belum ada|belum tersedia|akan restock)",
    r"will not.*selling.*now",r"ไม่มีจำหน่าย",
]

LOST_SALE_INDICATORS = [
    r"(never mind|forget it|cancel it|don't want|not interested anymore)",
    r"(found elsewhere|buying from another|going to another store)",
    r"(too expensive|price is high|cheaper elsewhere|other shop cheaper)",
    r"(if no stock|if not available|if out of stock).*(never mind|forget|cancel|ok bye|goodbye)",
    r"(ไม่เป็นไร|ไม่ซื้อแล้ว|ซื้อที่อื่น|แพงไป|ที่อื่นถูกกว่า)",
    r"(tidak jadi|beli di tempat lain|terlalu mahal|lebih murah di tempat lain)",
    r"(hindi na|bibili na lang sa ibang tinda|mahal|mas mura sa ibang shop)",
]

PRODUCT_INQUIRY_PATTERNS = [
    r"item_id:(\d+)",
    r"(product|item|produk|สินค้า|barang).{0,30}(available|ready|ada|มี|tersedia)",
    r"(price|harga|ราคา|presyo).{0,20}(\d+)",
    r"(size|ukuran|ขนาด|sukat).{0,20}([XSML0-9\-]+)",
    r"(colour|color|warna|สี|kulay).{0,20}(\w+)",
    r"(variant|variation|varian|รุ่น|bersedia)",
    r"stock.{0,15}(berapa|เท่าไหร่|ilan|how many)",
]

SIZE_PATTERNS = [
    r"\b(XS|S|M|L|XL|XXL|XXXL)\b",
    r"\b(\d{1,3})(cm|ml|mg|g|kg|oz|fl oz|inch|\")\b",
    r"\b(size|ukuran|ขนาด)\s*:?\s*([A-Z0-9\-\/]+)\b",
]

COLOR_PATTERNS = [
    r"\b(red|blue|green|black|white|pink|purple|yellow|orange|grey|gray|brown|navy|beige|cream)\b",
    r"\b(merah|biru|hijau|hitam|putih|pink|ungu|kuning|oranye|abu|coklat|krem)\b",
    r"\b(แดง|น้ำเงิน|เขียว|ดำ|ขาว|ชมพู|ม่วง|เหลือง|ส้ม|เทา|น้ำตาล|ครีม)\b",
    r"\b(pula|asul|berde|itim|puti|rosas|lila|dilaw|kahel|kulay-abo|kayumanggi)\b",
]

TEAM_START_DATE = pd.Timestamp("2026-03-30")

SUGGESTED_REPLIES = {
    "Refund": "Thank you for reaching out. We have reviewed your request and your refund of [AMOUNT] has been initiated and will be reflected within 3–5 business days. Order reference: [ORDER_ID]. We value your trust and hope to serve you better. 😊",
    "Return": "Thank you for contacting us about your return. We've initiated the return process for order [ORDER_ID]. Please expect a return label within 24 hours. Once received, your replacement or refund will be processed within 3–5 business days. 😊",
    "Cancellation": "We've received your cancellation request for order [ORDER_ID]. It has been successfully cancelled and any payment will be refunded within 3–5 business days. 😊",
    "Delay": "Thank you for your patience. We've checked with our logistics partner — your package [STATUS]. Estimated delivery: [DATE]. Track here: [TRACKING_LINK]. Please reach out if not received by [DATE+1]. 😊",
    "Damaged/Wrong Item": "We're truly sorry about your order [ORDER_ID]. A replacement will be dispatched within 1–2 business days. No need to return the incorrect/damaged item. We sincerely apologise. 😊",
    "Missing Item": "We're sorry your order [ORDER_ID] had a missing item. We've raised an investigation and will arrange a replacement or refund within 24 hours. Thank you for your patience. 😊",
    "Payment Issue": "Thank you for flagging this. Our finance team has been notified and the discrepancy for order [ORDER_ID] will be resolved within 2–3 business days. 😊",
    "Product Inquiry": "Thank you for your interest in [PRODUCT_NAME]! Here are the details: [DETAILS]. Feel free to ask about specs, sizing, or availability — we're happy to help you find the perfect product. 😊",
    "Promotion Issue": "Thank you for reaching out. We've reviewed order [ORDER_ID] and confirmed the discount of [AMOUNT] is applicable. Adjustment will reflect within 24–48 hours. 😊",
    "Technical Issue": "We apologise for the technical difficulty. Our team has been notified and is working on a resolution. Please try [TROUBLESHOOTING STEP] in the meantime. ETA: [TIMEFRAME]. 😊",
    "Complaint": "Thank you for your feedback. We sincerely apologise — your case [CASE_ID] has been escalated to our senior team for immediate review. A dedicated agent will contact you within 4 hours. 😊",
    "Other": "Thank you for reaching out! We've reviewed your message and our team is addressing your concern. We aim to resolve within 24 hours. 😊",
}

ACTION_STEPS = {
    "Refund": "1. Verify order ID & payment method.\n2. Check refund eligibility (within 15 days).\n3. Initiate via platform refund portal.\n4. Confirm amount & notify buyer (3–5 days).\n5. Log in DKSH tracker.",
    "Return": "1. Verify condition & return reason.\n2. Check return window (Lazada 7d, Shopee 15d).\n3. Approve in Seller Centre.\n4. Send return shipping label.\n5. Inspect & process refund/replacement. 6. Update tracker.",
    "Cancellation": "1. Check order status — cancellable only before 'Ready to Ship'.\n2. Approve cancellation.\n3. If shipped, advise buyer to reject delivery.\n4. Refund auto-processes 3–5 days.",
    "Delay": "1. Check logistics tracking.\n2. Contact logistics if stalled >3 days.\n3. Share tracking link with buyer.\n4. If lost, file claim.\n5. Offer replacement/refund if SLA fails.",
    "Damaged/Wrong Item": "1. Request photo evidence.\n2. Log dispute in Seller Centre.\n3. Approve replacement — do NOT ask buyer to return.\n4. Update tracker.\n5. Report to warehouse for QC investigation.",
    "Missing Item": "1. Request unboxing video/photo.\n2. Check packing list vs order.\n3. Dispatch replacement within 24h.\n4. Raise internal warehouse investigation.",
    "Payment Issue": "1. Verify transaction in payment dashboard.\n2. Check for double-charge.\n3. Raise dispute ticket with platform finance.\n4. Provide buyer case/ticket reference.\n5. Follow up within 2 days.",
    "Product Inquiry": "1. Provide accurate product specs from official sheet.\n2. Check live inventory if stock inquiry.\n3. Share size guide if sizing question.\n4. ⭐ UPSELL: Suggest related/complementary products.\n5. ⭐ CONVERT: Guide buyer to purchase if ready.",
    "Promotion Issue": "1. Verify voucher/promo code validity.\n2. Check eligibility criteria.\n3. If code valid but not applied, advise re-checkout.\n4. If expired, offer alternative discount.\n5. Escalate to marketing for setup errors.",
    "Technical Issue": "1. Identify platform & device.\n2. Advise: clear cache, update app, reinstall.\n3. Check platform status page.\n4. Raise support ticket.\n5. Keep buyer updated.",
    "Complaint": "1. Acknowledge & empathise — do NOT be defensive.\n2. Log in DKSH escalation tracker.\n3. Identify root cause.\n4. Offer concrete resolution: refund/replacement/discount.\n5. Escalate to senior manager if buyer threatens churn.\n6. Follow up within 4 hours.",
    "Other": "1. Understand concern fully.\n2. Route to appropriate team.\n3. Resolve within 24 hours.\n4. Log in DKSH tracker.",
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS — OPERATIONS
# ─────────────────────────────────────────────────────────────────────────────

def detect_sentiment(text):
    if not isinstance(text, str) or not text.strip():
        return "Neutral"
    t = text.lower()
    neg = sum(1 for kw in NEGATIVE_KWS if kw in t)
    pos = sum(1 for kw in POSITIVE_KWS if kw in t)
    if neg > pos: return "Negative"
    if pos > neg: return "Positive"
    return "Neutral"

def detect_issue_type(text):
    if not isinstance(text, str) or not text.strip():
        return "Other"
    t = text.lower()
    scores = {}
    for issue, kws in ISSUE_KEYWORDS.items():
        score = sum(1 for kw in kws if kw.lower() in t)
        if score > 0:
            scores[issue] = score
    if not scores: return "Other"
    return max(scores, key=scores.get)

def get_priority(issue_type):
    for priority, issues in PRIORITY_MAP.items():
        if issue_type in issues:
            return priority
    return "Low"

def matches_any(text, patterns):
    if not isinstance(text, str): return False
    t = text.lower()
    return any(re.search(p, t, re.IGNORECASE) for p in patterns)

def is_auto_reply(text):
    return matches_any(text, AUTO_REPLY_PATTERNS)

def conversation_is_unresolved(seller_msgs):
    stall_found = False
    for msg in seller_msgs:
        if matches_any(msg, STALLING_PATTERNS):
            stall_found = True
        if matches_any(msg, RESOLUTION_PATTERNS):
            stall_found = False
    return stall_found

def compute_csat(sentiment, is_resolved):
    matrix = {
        ("Positive", True): 5.0, ("Positive", False): 3.5,
        ("Neutral",  True): 4.0, ("Neutral",  False): 3.0,
        ("Negative", True): 2.5, ("Negative", False): 1.0,
    }
    return matrix.get((sentiment, is_resolved), 3.0)

def generate_summary(buyer_msgs, issue_type):
    if not buyer_msgs: return "No buyer messages."
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])[:400]
    return f"[{issue_type}] Buyer enquiry: {combined[:200]}{'...' if len(combined) > 200 else ''}"

def fmt_mins(mins):
    if pd.isna(mins) or mins < 0: return "—"
    if mins < 60: return f"{int(mins)}m"
    h = int(mins // 60); m = int(mins % 60)
    return f"{h}h {m}m" if m else f"{h}h"

def get_team_member(store_code):
    code = str(store_code).strip().upper()
    if not code: return "Others"
    return STORE_TO_AGENT.get(code, "Others")

def detect_conversion(buyer_msgs):
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in CONVERSION_KEYWORDS)

def get_action_steps(issue_type):
    return ACTION_STEPS.get(issue_type, ACTION_STEPS["Other"])

# ─────────────────────────────────────────────────────────────────────────────
# SALES INTELLIGENCE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def detect_upsell_opportunity(buyer_msgs):
    """Returns True if buyer asked for alternatives/recommendations."""
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in UPSELL_KEYWORDS)

def detect_oos_inquiry(all_msgs):
    """Returns True if OOS was mentioned by either party."""
    combined = " ".join([m for m in all_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in OOS_KEYWORDS)

def detect_oos_seller_confirmed(seller_msgs):
    """Returns True if seller confirmed OOS."""
    for msg in seller_msgs:
        if matches_any(msg, OOS_SELLER_PATTERNS):
            return True
    return False

def detect_lost_sale(buyer_msgs):
    """Returns True if buyer gave up / went elsewhere."""
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return matches_any(combined, LOST_SALE_INDICATORS)

def detect_similar_suggested(seller_msgs):
    """Returns True if seller suggested similar/alternative product."""
    suggestion_kws = [
        "alternatively","similar product","you might like","we also have","how about",
        "may i suggest","recommend","try this","check out","นอกจากนี้","สินค้าอื่น",
        "kami juga ada","produk lain","alternatif","ลองดู","ลองสั่ง",
        "meron din kami","pwede rin","katulad nito",
    ]
    combined = " ".join([m for m in seller_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in suggestion_kws)

def extract_product_references(buyer_msgs):
    """Extract item IDs and product-related keywords from buyer messages."""
    refs = []
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])
    # item_id references
    item_ids = re.findall(r"item_id:(\d+)", combined, re.IGNORECASE)
    refs.extend([f"item_id:{i}" for i in item_ids])
    return refs

def extract_size_mentions(buyer_msgs):
    """Extract size/dimension mentions from buyer messages."""
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])
    sizes = []
    for pat in SIZE_PATTERNS:
        matches = re.findall(pat, combined, re.IGNORECASE)
        for m in matches:
            if isinstance(m, tuple):
                sizes.append(" ".join(m).strip())
            else:
                sizes.append(str(m).strip())
    return list(set(sizes)) if sizes else []

def extract_color_mentions(buyer_msgs):
    """Extract color mentions from buyer messages."""
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])
    colors = []
    for pat in COLOR_PATTERNS:
        matches = re.findall(pat, combined, re.IGNORECASE)
        colors.extend(matches)
    return list(set([c.lower() for c in colors])) if colors else []

def classify_sales_stage(buyer_msgs, seller_msgs, is_conversion, is_oos_confirmed, is_lost_sale):
    """Classify conversation into sales funnel stage."""
    if is_conversion:
        return "🟢 Converted"
    if is_lost_sale:
        return "🔴 Lost Sale"
    if is_oos_confirmed:
        return "🟠 OOS — Demand Captured"
    b_combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    if any(kw in b_combined for kw in ["want to buy","interested","i want","มอยากได้","อยากได้","mau beli","gusto ko"]):
        return "🟡 High Intent"
    if any(kw in b_combined for kw in ["price","ราคา","harga","how much","berapa","presyo","magkano"]):
        return "🔵 Price Check"
    if any(kw in b_combined for kw in ["size","available","stock","variant","color","colour"]):
        return "🔵 Product Research"
    return "⚪ Awareness"

def compute_wow_mom(conv_df):
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"].notna()].copy()
    if df.empty: return pd.DataFrame(), pd.DataFrame()
    df["WEEK"]  = df["LAST_MSG_TIME"].dt.to_period("W").apply(lambda r: r.start_time)
    df["MONTH"] = df["LAST_MSG_TIME"].dt.to_period("M").apply(lambda r: r.start_time)
    def agg_metrics(df_in, period_col):
        agg = (
            df_in.groupby(period_col).agg(
                Conversations=("CONVERSATION_ID","count"),
                Resolved=("IS_RESOLVED","sum"),
                Unresolved=("IS_UNRESOLVED","sum"),
                Avg_CSAT=("CSAT_PROXY","mean"),
                Avg_CRT_mins=("AVG_CRT_MINS","mean"),
                Negative=("SENTIMENT", lambda x: (x=="Negative").sum()),
                Positive=("SENTIMENT", lambda x: (x=="Positive").sum()),
                Conversions=("IS_CONVERSION","sum"),
                Lost_Sales=("IS_LOST_SALE","sum"),
                OOS_Inquiries=("IS_OOS_INQUIRY","sum"),
            ).reset_index().sort_values(period_col)
        )
        agg["CRR_%"] = (agg["Resolved"] / agg["Conversations"] * 100).round(1)
        agg["Avg_CSAT"] = agg["Avg_CSAT"].round(2)
        agg["Avg_CRT_mins"] = agg["Avg_CRT_mins"].round(1)
        agg["Conv_Rate_%"] = (agg["Conversions"] / agg["Conversations"] * 100).round(1)
        for col in ["Conversations","Avg_CSAT","CRR_%","Avg_CRT_mins","Conversions","Lost_Sales"]:
            agg[f"Δ {col}"] = agg[col].diff().round(2)
        return agg
    return agg_metrics(df,"WEEK"), agg_metrics(df,"MONTH")

def compute_team_performance(conv_df):
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"] >= TEAM_START_DATE].copy()
    if df.empty or "TEAM_MEMBER" not in df.columns: return pd.DataFrame()
    perf = (
        df.groupby("TEAM_MEMBER").agg(
            Conversations=("CONVERSATION_ID","count"),
            Resolved=("IS_RESOLVED","sum"),
            Unresolved=("IS_UNRESOLVED","sum"),
            Avg_CSAT=("CSAT_PROXY","mean"),
            Avg_CRT_mins=("AVG_CRT_MINS","mean"),
            Positive_Sent=("SENTIMENT", lambda x: (x=="Positive").sum()),
            Negative_Sent=("SENTIMENT", lambda x: (x=="Negative").sum()),
            Conversions=("IS_CONVERSION","sum"),
            High_Priority=("PRIORITY", lambda x: (x=="High").sum()),
            Upsell_Opportunities=("IS_UPSELL_OPP","sum"),
            Lost_Sales=("IS_LOST_SALE","sum"),
            Similar_Suggested=("SIMILAR_SUGGESTED","sum"),
            OOS_Handled=("IS_OOS_CONFIRMED","sum"),
        ).reset_index()
    )
    perf["CRR_%"] = (perf["Resolved"] / perf["Conversations"] * 100).round(1)
    perf["Avg_CSAT"] = perf["Avg_CSAT"].round(2)
    perf["Avg_CRT_mins"] = perf["Avg_CRT_mins"].round(1)
    perf["Conv_Rate_%"] = (perf["Conversions"] / perf["Conversations"] * 100).round(1)
    perf["Upsell_Act_Rate_%"] = (perf["Similar_Suggested"] / perf["Upsell_Opportunities"].replace(0,np.nan) * 100).round(1)
    perf["Shift"] = perf["TEAM_MEMBER"].map(AGENT_SHIFT).fillna("Day")
    return perf.sort_values("Conversations", ascending=False).reset_index(drop=True)

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_data(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets_found = xl.sheet_names
    dfs = []
    platform_map = {}
    for s in sheets_found:
        n = s.lower()
        if "lazada" in n: platform_map[s] = "Lazada"
        elif "shopee" in n: platform_map[s] = "Shopee"
        else: platform_map[s] = "Unknown"
        df = xl.parse(s, dtype=str)
        df["PLATFORM"] = platform_map[s]
        dfs.append(df)
    combined = pd.concat(dfs, ignore_index=True)
    combined["MESSAGE_TIME"] = pd.to_datetime(combined["MESSAGE_TIME"], errors="coerce")
    for col in ["STORE_CODE","SITE_NICK_NAME_ID","CHANNEL_NAME","COUNTRY_CODE",
                "CONVERSATION_ID","BUYER_NAME","MESSAGE_PARSED","MESSAGE_TYPE","SENDER"]:
        if col in combined.columns:
            combined[col] = combined[col].fillna("").astype(str).str.strip()
    for flag in ["IS_READ","IS_ANSWERED"]:
        if flag in combined.columns:
            combined[flag] = combined[flag].astype(str).str.strip().str.lower().isin(["true","1","yes"])
    return combined

# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False, max_entries=1)
def analyse(df):
    df = df.copy()
    df["_sender_lower"] = df["SENDER"].str.lower().fillna("")
    df_sorted = df.sort_values(["CONVERSATION_ID","MESSAGE_TIME"])
    buyer_mask  = df_sorted["_sender_lower"] == "buyer"
    seller_mask = df_sorted["_sender_lower"] == "seller"

    buyer_text_per_conv = (
        df_sorted[buyer_mask].groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(lambda msgs: " ".join(m for m in msgs if isinstance(m, str)))
    )
    issue_map     = buyer_text_per_conv.apply(detect_issue_type)
    sentiment_map = buyer_text_per_conv.apply(detect_sentiment)

    meta_cols = [c for c in ["PLATFORM","STORE_CODE","SITE_NICK_NAME_ID","CHANNEL_NAME",
                              "COUNTRY_CODE","BUYER_NAME","BUYER_ID","IS_ANSWERED","IS_READ"]
                 if c in df_sorted.columns]
    meta_df = df_sorted.groupby("CONVERSATION_ID")[meta_cols].first()
    time_df = df_sorted.groupby("CONVERSATION_ID")["MESSAGE_TIME"].agg(FIRST_MSG_TIME="min", LAST_MSG_TIME="max")
    total_msgs        = df_sorted.groupby("CONVERSATION_ID").size().rename("MSG_COUNT")
    buyer_msgs_count  = df_sorted[buyer_mask].groupby("CONVERSATION_ID").size().rename("BUYER_MSG_COUNT")
    seller_msgs_count = df_sorted[seller_mask].groupby("CONVERSATION_ID").size().rename("SELLER_MSG_COUNT")

    seller_msgs_per_conv = df_sorted[seller_mask].groupby("CONVERSATION_ID")["MESSAGE_PARSED"].apply(list)
    buyer_msgs_per_conv  = df_sorted[buyer_mask].groupby("CONVERSATION_ID")["MESSAGE_PARSED"].apply(list)
    all_msgs_per_conv    = df_sorted.groupby("CONVERSATION_ID")["MESSAGE_PARSED"].apply(list)

    rows = []
    for conv_id, grp in df_sorted.groupby("CONVERSATION_ID", sort=False):
        issue_type = issue_map.get(conv_id, "Other")
        sentiment  = sentiment_map.get(conv_id, "Neutral")
        b_msgs = buyer_msgs_per_conv.get(conv_id, [])
        s_msgs = seller_msgs_per_conv.get(conv_id, [])
        a_msgs = all_msgs_per_conv.get(conv_id, [])
        meta   = meta_df.loc[conv_id] if conv_id in meta_df.index else {}

        is_unresolved = conversation_is_unresolved(s_msgs)
        is_resolved   = not is_unresolved
        priority      = get_priority(issue_type)
        csat          = compute_csat(sentiment, is_resolved)

        # Sales intelligence
        is_conversion       = detect_conversion(b_msgs)
        is_upsell_opp       = detect_upsell_opportunity(b_msgs)
        is_oos_inquiry      = detect_oos_inquiry(a_msgs)
        is_oos_confirmed    = detect_oos_seller_confirmed(s_msgs)
        is_lost_sale        = detect_lost_sale(b_msgs)
        similar_suggested   = detect_similar_suggested(s_msgs)
        product_refs        = extract_product_references(b_msgs)
        size_mentions       = extract_size_mentions(b_msgs)
        color_mentions      = extract_color_mentions(b_msgs)
        sales_stage         = classify_sales_stage(b_msgs, s_msgs, is_conversion, is_oos_confirmed, is_lost_sale)

        # CRT
        crt_list = []
        last_buyer_time = None
        for sender, msg_time in zip(grp["_sender_lower"].tolist(), grp["MESSAGE_TIME"].tolist()):
            if sender == "buyer":
                last_buyer_time = msg_time
            elif sender == "seller" and last_buyer_time is not None:
                delta = (msg_time - last_buyer_time).total_seconds() / 60
                if 0 <= delta <= 1440:
                    crt_list.append(delta)
                last_buyer_time = None
        avg_crt = float(np.mean(crt_list)) if crt_list else np.nan

        def _get(field, default=""):
            try: return meta[field] if hasattr(meta,"__getitem__") else getattr(meta, field, default)
            except: return default

        rows.append({
            "CONVERSATION_ID":   conv_id,
            "PLATFORM":          _get("PLATFORM"),
            "STORE_CODE":        _get("STORE_CODE"),
            "SITE_NICK_NAME_ID": _get("SITE_NICK_NAME_ID"),
            "CHANNEL_NAME":      _get("CHANNEL_NAME"),
            "COUNTRY_CODE":      _get("COUNTRY_CODE"),
            "BUYER_NAME":        _get("BUYER_NAME"),
            "BUYER_ID":          _get("BUYER_ID"),
            "FIRST_MSG_TIME":    time_df.loc[conv_id,"FIRST_MSG_TIME"] if conv_id in time_df.index else pd.NaT,
            "LAST_MSG_TIME":     time_df.loc[conv_id,"LAST_MSG_TIME"]  if conv_id in time_df.index else pd.NaT,
            "MSG_COUNT":         int(total_msgs.get(conv_id, 0)),
            "BUYER_MSG_COUNT":   int(buyer_msgs_count.get(conv_id, 0)),
            "SELLER_MSG_COUNT":  int(seller_msgs_count.get(conv_id, 0)),
            "ISSUE_TYPE":        issue_type,
            "PRIORITY":          priority,
            "SENTIMENT":         sentiment,
            "IS_UNRESOLVED":     is_unresolved,
            "IS_RESOLVED":       is_resolved,
            "CSAT_PROXY":        round(csat, 1),
            "AVG_CRT_MINS":      round(avg_crt, 1) if not np.isnan(avg_crt) else None,
            "BUYER_SUMMARY":     generate_summary(b_msgs, issue_type),
            "IS_CONVERSION":     is_conversion,
            "IS_UPSELL_OPP":     is_upsell_opp,
            "IS_OOS_INQUIRY":    is_oos_inquiry,
            "IS_OOS_CONFIRMED":  is_oos_confirmed,
            "IS_LOST_SALE":      is_lost_sale,
            "SIMILAR_SUGGESTED": similar_suggested,
            "PRODUCT_REFS":      "|".join(product_refs) if product_refs else "",
            "SIZE_MENTIONS":     "|".join(size_mentions) if size_mentions else "",
            "COLOR_MENTIONS":    "|".join(color_mentions) if color_mentions else "",
            "SALES_STAGE":       sales_stage,
            "TEAM_MEMBER":       get_team_member(_get("STORE_CODE")),
            "IS_ANSWERED":       str(_get("IS_ANSWERED")).lower() == "true",
            "IS_READ":           str(_get("IS_READ")).lower() == "true",
        })

    result = pd.DataFrame(rows)
    for col in ["PLATFORM","ISSUE_TYPE","PRIORITY","SENTIMENT","STORE_CODE",
                "CHANNEL_NAME","COUNTRY_CODE","TEAM_MEMBER","SITE_NICK_NAME_ID","SALES_STAGE"]:
        if col in result.columns:
            result[col] = result[col].astype("category")
    for col in ["BUYER_SUMMARY"]:
        if col in result.columns:
            result[col] = result[col].str[:300]
    gc.collect()
    return result

# ─────────────────────────────────────────────────────────────────────────────
# SALES INTELLIGENCE AGGREGATIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_oos_tracker(raw_df, conv_df):
    """Build OOS product tracker from raw messages."""
    seller_mask = raw_df["SENDER"].str.lower() == "seller"
    oos_convs = conv_df[conv_df["IS_OOS_CONFIRMED"] == True]["CONVERSATION_ID"].tolist()
    if not oos_convs: return pd.DataFrame()
    oos_raw = raw_df[raw_df["CONVERSATION_ID"].isin(oos_convs)].copy()
    # Extract product references
    oos_raw["PRODUCT_REF"] = oos_raw["MESSAGE_PARSED"].apply(
        lambda x: "|".join(re.findall(r"item_id:(\d+)", str(x))) if pd.notna(x) else ""
    )
    # Get store/country info
    oos_summary = []
    for conv_id in oos_convs:
        msgs = oos_raw[oos_raw["CONVERSATION_ID"] == conv_id]
        store = msgs["STORE_CODE"].iloc[0] if not msgs.empty else ""
        country = msgs["COUNTRY_CODE"].iloc[0] if not msgs.empty and "COUNTRY_CODE" in msgs.columns else ""
        buyer_msgs = msgs[msgs["SENDER"].str.lower() == "buyer"]["MESSAGE_PARSED"].tolist()
        seller_msgs_list = msgs[msgs["SENDER"].str.lower() == "seller"]["MESSAGE_PARSED"].tolist()
        combined_buyer = " ".join([m for m in buyer_msgs if isinstance(m,str)])
        item_ids = re.findall(r"item_id:(\d+)", combined_buyer)
        sizes = extract_size_mentions(buyer_msgs)
        colors = extract_color_mentions(buyer_msgs)
        similar_was_suggested = detect_similar_suggested(seller_msgs_list)
        lost = detect_lost_sale(buyer_msgs)
        oos_summary.append({
            "CONVERSATION_ID": conv_id,
            "STORE_CODE": store,
            "COUNTRY_CODE": country,
            "ITEM_IDS_INQUIRED": "|".join(item_ids) if item_ids else "Unknown",
            "SIZE_REQUESTED": "|".join(sizes) if sizes else "—",
            "COLOR_REQUESTED": "|".join(colors) if colors else "—",
            "SIMILAR_SUGGESTED": similar_was_suggested,
            "LOST_SALE": lost,
        })
    if not oos_summary: return pd.DataFrame()
    oos_df = pd.DataFrame(oos_summary)
    return oos_df

def build_product_inquiry_summary(raw_df, conv_df):
    """Summarise most inquired products/items/categories."""
    product_inquiry_convs = conv_df[conv_df["ISSUE_TYPE"] == "Product Inquiry"]["CONVERSATION_ID"].tolist()
    if not product_inquiry_convs: return pd.DataFrame(), pd.DataFrame()
    pi_raw = raw_df[
        (raw_df["CONVERSATION_ID"].isin(product_inquiry_convs)) &
        (raw_df["SENDER"].str.lower() == "buyer")
    ].copy()
    # Extract item IDs
    item_counts = defaultdict(int)
    size_counts = defaultdict(int)
    color_counts = defaultdict(int)
    for _, row in pi_raw.iterrows():
        msg = str(row.get("MESSAGE_PARSED",""))
        for item_id in re.findall(r"item_id:(\d+)", msg):
            item_counts[item_id] += 1
        store = str(row.get("STORE_CODE",""))
        for pat in SIZE_PATTERNS:
            for m in re.findall(pat, msg, re.IGNORECASE):
                sz = " ".join(m).strip() if isinstance(m, tuple) else str(m).strip()
                if sz: size_counts[sz.upper()] += 1
        for pat in COLOR_PATTERNS:
            for c in re.findall(pat, msg, re.IGNORECASE):
                if c: color_counts[c.lower()] += 1

    item_df = pd.DataFrame(
        [{"Item ID": k, "Inquiry Count": v} for k,v in sorted(item_counts.items(), key=lambda x: -x[1])]
    ).head(20) if item_counts else pd.DataFrame()

    var_rows = []
    for sz, cnt in sorted(size_counts.items(), key=lambda x: -x[1])[:15]:
        var_rows.append({"Variation": sz, "Type": "Size", "Count": cnt})
    for col, cnt in sorted(color_counts.items(), key=lambda x: -x[1])[:15]:
        var_rows.append({"Variation": col.title(), "Type": "Color", "Count": cnt})
    var_df = pd.DataFrame(var_rows).sort_values("Count", ascending=False) if var_rows else pd.DataFrame()

    return item_df, var_df

def compute_conversion_funnel(conv_df):
    """Build conversion funnel data."""
    total = len(conv_df)
    if total == 0:
        return {}
    product_inq = len(conv_df[conv_df["ISSUE_TYPE"] == "Product Inquiry"])
    high_intent = len(conv_df[conv_df["SALES_STAGE"].astype(str).str.contains("High Intent|Converted", na=False)])
    converted   = int(conv_df["IS_CONVERSION"].sum())
    oos         = int(conv_df["IS_OOS_INQUIRY"].sum())
    lost        = int(conv_df["IS_LOST_SALE"].sum())
    upsell_opp  = int(conv_df["IS_UPSELL_OPP"].sum())
    upsell_acted = int(conv_df["SIMILAR_SUGGESTED"].sum())
    return {
        "total_conversations": total,
        "product_inquiries": product_inq,
        "high_intent": high_intent,
        "converted": converted,
        "oos_total": oos,
        "lost_sales": lost,
        "upsell_opportunities": upsell_opp,
        "upsell_acted": upsell_acted,
        "conv_rate_pct": round(converted / total * 100, 1) if total else 0,
        "lost_rate_pct": round(lost / total * 100, 1) if total else 0,
        "upsell_act_rate_pct": round(upsell_acted / upsell_opp * 100, 1) if upsell_opp else 0,
    }

def generate_key_improvements(conv_df, funnel):
    """Generate prioritised improvement recommendations."""
    recs = []
    total = funnel.get("total_conversations", 1)

    # OOS handling
    oos_pct = funnel.get("oos_total", 0) / total * 100
    upsell_act = funnel.get("upsell_act_rate_pct", 0)
    lost_pct = funnel.get("lost_rate_pct", 0)
    conv_rate = funnel.get("conv_rate_pct", 0)

    if oos_pct > 5:
        recs.append(("🔴 HIGH", "Stock Management",
            f"{funnel.get('oos_total',0)} OOS inquiries ({oos_pct:.1f}% of chats). Immediate stock review needed. "
            "Work with merchandiser to prioritise restock of most-inquired OOS items."))

    if upsell_act < 30 and funnel.get("upsell_opportunities",0) > 10:
        recs.append(("🔴 HIGH", "Upsell Execution Gap",
            f"Only {upsell_act:.0f}% of upsell opportunities had a similar product suggested. "
            "Train agents to always suggest alternatives when item is OOS or when buyer asks 'any other options?'"))

    if lost_pct > 3:
        recs.append(("🔴 HIGH", "Lost Sales Recovery",
            f"{funnel.get('lost_sales',0)} buyers ({lost_pct:.1f}%) left without purchasing. "
            "Introduce retention offer (voucher/bundle) as a last resort before buyer disengages."))

    if conv_rate < 10:
        recs.append(("🟡 MEDIUM", "Conversion Rate Improvement",
            f"Overall conversion rate is {conv_rate:.1f}%. Agents should be trained to guide buyers through checkout "
            "when intent is detected. Consider adding CTA phrases and product links in responses."))

    unresolved_pct = conv_df["IS_UNRESOLVED"].mean() * 100 if len(conv_df) > 0 else 0
    if unresolved_pct > 20:
        recs.append(("🟡 MEDIUM", "Resolution Rate",
            f"{unresolved_pct:.1f}% of conversations are unresolved. Review stalled cases and implement "
            "follow-up SLA reminders. High unresolved rate damages CSAT and repeat purchase intent."))

    avg_crt = conv_df["AVG_CRT_MINS"].mean()
    if not np.isnan(avg_crt) and avg_crt > 60:
        recs.append(("🟡 MEDIUM", "Response Time",
            f"Average CRT is {fmt_mins(avg_crt)}. Target <30 minutes during business hours. "
            "Consider auto-reply templates for common product inquiries to reduce manual response time."))

    neg_pct = (conv_df["SENTIMENT"] == "Negative").sum() / total * 100 if total > 0 else 0
    if neg_pct > 15:
        recs.append(("🟡 MEDIUM", "Negative Sentiment Alert",
            f"{neg_pct:.1f}% of conversations have negative sentiment. Review top complaint themes and "
            "implement proactive outreach for at-risk buyers."))

    # Product inquiry without conversion
    pi_conv = conv_df[(conv_df["ISSUE_TYPE"] == "Product Inquiry") & (conv_df["IS_CONVERSION"] == False)]
    pi_no_conv_pct = len(pi_conv) / max(len(conv_df[conv_df["ISSUE_TYPE"]=="Product Inquiry"]),1) * 100
    if pi_no_conv_pct > 60:
        recs.append(("🟢 OPPORTUNITY", "Product Inquiry → Purchase Conversion",
            f"{pi_no_conv_pct:.0f}% of product inquiries did NOT convert. These are warm leads. "
            "Train agents to add product links, size guides, and urgency cues (limited stock, promo ends soon)."))

    recs.append(("🟢 OPPORTUNITY", "Variation Data for Merchandising",
        "Frequently requested sizes and colours in chats = real demand signal. "
        "Share top variation requests with merchandising team monthly for listing optimisation and stock planning."))

    recs.append(("🟢 OPPORTUNITY", "OOS Demand List for Buying Team",
        "Every OOS inquiry is a lost sale AND a demand signal. Build a weekly OOS report from chat data "
        "and share with buying/AM team to prioritise restock or find substitute SKUs."))

    return recs

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(conv_df, raw_df, today_str):
    df = conv_df.copy()
    if "SUGGESTED_REPLY" not in df.columns and "ISSUE_TYPE" in df.columns:
        df["SUGGESTED_REPLY"] = df["ISSUE_TYPE"].astype(str).map(
            lambda it: SUGGESTED_REPLIES.get(it, SUGGESTED_REPLIES["Other"]))
    if "ACTION_STEPS" not in df.columns and "ISSUE_TYPE" in df.columns:
        df["ACTION_STEPS"] = df["ISSUE_TYPE"].astype(str).map(get_action_steps)

    funnel = compute_conversion_funnel(df)
    oos_df = build_oos_tracker(raw_df, df)
    item_df, var_df = build_product_inquiry_summary(raw_df, df)
    team_perf = compute_team_performance(df)
    improvements = generate_key_improvements(df, funnel)
    total = len(df)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt  = wb.add_format({"bold":True,"bg_color":"#1B2A4A","font_color":"#FFFFFF","border":1,"font_size":10,"align":"center","valign":"vcenter"})
        sub_fmt  = wb.add_format({"bold":True,"bg_color":"#00C4B4","font_color":"#FFFFFF","border":1,"font_size":10})
        sal_fmt  = wb.add_format({"bold":True,"bg_color":"#FF6B35","font_color":"#FFFFFF","border":1,"font_size":10})
        mer_fmt  = wb.add_format({"bold":True,"bg_color":"#8B5CF6","font_color":"#FFFFFF","border":1,"font_size":10})
        num_fmt  = wb.add_format({"num_format":"#,##0","border":1})
        cell_fmt = wb.add_format({"border":1,"font_size":9,"text_wrap":True,"valign":"top"})
        red_fmt  = wb.add_format({"border":1,"font_size":9,"bg_color":"#FDECEA","font_color":"#C0392B"})
        grn_fmt  = wb.add_format({"border":1,"font_size":9,"bg_color":"#E9F7EF","font_color":"#196F3D"})

        def write_df(ws, df_in, start_row=0, fmt=hdr_fmt):
            for c_idx, col in enumerate(df_in.columns):
                ws.write(start_row, c_idx, col, fmt)
            for r_idx, row in enumerate(df_in.itertuples(index=False), start=start_row+1):
                for c_idx, val in enumerate(row):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        ws.write(r_idx, c_idx, "", cell_fmt)
                    elif isinstance(val, bool):
                        ws.write(r_idx, c_idx, "Yes" if val else "No", grn_fmt if val else cell_fmt)
                    elif isinstance(val, (int, float)):
                        ws.write_number(r_idx, c_idx, val, num_fmt)
                    else:
                        ws.write(r_idx, c_idx, str(val), cell_fmt)

        # ── Sheet 1: Summary Dashboard ────────────────────────────────────────
        ws1 = wb.add_worksheet("Summary Dashboard")
        writer.sheets["Summary Dashboard"] = ws1
        ws1.set_column(0, 0, 32); ws1.set_column(1, 1, 22)
        ws1.write(0, 0, f"Chat Analyzer v2 — {today_str}", wb.add_format({"bold":True,"font_size":14,"font_color":"#1B2A4A"}))
        ws1.write(1, 0, "Graas.ai | Operations + Sales Intelligence", wb.add_format({"italic":True,"font_color":"#7A8EA8"}))
        resolved = df["IS_RESOLVED"].sum(); unresolved = df["IS_UNRESOLVED"].sum()
        crr = round(resolved/total*100,1) if total else 0
        avg_crt = df["AVG_CRT_MINS"].mean(); avg_csat = df["CSAT_PROXY"].mean()
        ops_data = [
            ["── OPERATIONS METRICS ──",""],
            ["Total Conversations", total],
            ["Resolved", int(resolved)],
            ["Unresolved", int(unresolved)],
            ["Chat Resolution Rate (CRR)", f"{crr}%"],
            ["Avg Chat Response Time", fmt_mins(avg_crt)],
            ["Avg CSAT Proxy (1–5)", round(avg_csat,2) if not np.isnan(avg_csat) else "—"],
        ]
        sales_data = [
            ["── SALES METRICS ──",""],
            ["Total Conversions (Intent Detected)", funnel.get("converted",0)],
            ["Conversion Rate", f"{funnel.get('conv_rate_pct',0)}%"],
            ["Upsell Opportunities", funnel.get("upsell_opportunities",0)],
            ["Upsell Action Rate", f"{funnel.get('upsell_act_rate_pct',0)}%"],
            ["OOS Inquiries", funnel.get("oos_total",0)],
            ["Lost Sales Detected", funnel.get("lost_sales",0)],
            ["Lost Sale Rate", f"{funnel.get('lost_rate_pct',0)}%"],
            ["Product Inquiries", funnel.get("product_inquiries",0)],
        ]
        row = 3
        for label, val in ops_data:
            ws1.write(row, 0, label, sub_fmt if "──" in str(label) else cell_fmt)
            ws1.write(row, 1, val, cell_fmt)
            row += 1
        row += 1
        for label, val in sales_data:
            ws1.write(row, 0, label, sal_fmt if "──" in str(label) else cell_fmt)
            ws1.write(row, 1, val, cell_fmt)
            row += 1

        # Issue breakdown
        row += 2
        ws1.write(row, 0, "ISSUE TYPE BREAKDOWN", sub_fmt); ws1.write(row, 1, "COUNT", hdr_fmt); row += 1
        for issue, cnt in df["ISSUE_TYPE"].value_counts().items():
            ws1.write(row, 0, issue, cell_fmt); ws1.write(row, 1, int(cnt), num_fmt); row += 1

        # Sales stage breakdown
        row += 2
        ws1.write(row, 0, "SALES FUNNEL STAGE", sal_fmt); ws1.write(row, 1, "COUNT", hdr_fmt); row += 1
        for stage, cnt in df["SALES_STAGE"].value_counts().items():
            ws1.write(row, 0, str(stage), cell_fmt); ws1.write(row, 1, int(cnt), num_fmt); row += 1

        # ── Sheet 2: Sales Intelligence ───────────────────────────────────────
        ws2 = wb.add_worksheet("Sales Intelligence")
        writer.sheets["Sales Intelligence"] = ws2
        ws2.set_column(0, 0, 40); ws2.set_column(1, 15, 18)
        ws2.write(0, 0, "Sales Intelligence — Conversion & Upsell Analysis", wb.add_format({"bold":True,"font_size":13,"font_color":"#FF6B35"}))
        sales_cols = [c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER",
                                   "SALES_STAGE","IS_CONVERSION","IS_UPSELL_OPP","SIMILAR_SUGGESTED",
                                   "IS_OOS_INQUIRY","IS_OOS_CONFIRMED","IS_LOST_SALE",
                                   "PRODUCT_REFS","SIZE_MENTIONS","COLOR_MENTIONS","SENTIMENT","BUYER_SUMMARY"] if c in df.columns]
        write_df(ws2, df[sales_cols], start_row=2, fmt=sal_fmt)

        # ── Sheet 3: OOS Tracker ──────────────────────────────────────────────
        ws3 = wb.add_worksheet("OOS Tracker")
        writer.sheets["OOS Tracker"] = ws3
        ws3.write(0, 0, "Out-of-Stock Demand Tracker", wb.add_format({"bold":True,"font_size":13,"font_color":"#E74C3C"}))
        ws3.write(1, 0, "Share with Buying/AM team to prioritise restock", wb.add_format({"italic":True,"font_color":"#7A8EA8"}))
        if not oos_df.empty:
            write_df(ws3, oos_df, start_row=3, fmt=hdr_fmt)
        else:
            ws3.write(3, 0, "No OOS inquiries detected in this dataset.", cell_fmt)

        # ── Sheet 4: Product & Variation Demand ───────────────────────────────
        ws4 = wb.add_worksheet("Product Demand")
        writer.sheets["Product Demand"] = ws4
        ws4.write(0, 0, "Most Inquired Products & Variations", wb.add_format({"bold":True,"font_size":13,"font_color":"#8B5CF6"}))
        ws4.write(1, 0, "Use for merchandising decisions, listing optimisation, and stock planning", wb.add_format({"italic":True,"font_color":"#7A8EA8"}))
        if not item_df.empty:
            ws4.write(3, 0, "TOP INQUIRED ITEM IDs", mer_fmt)
            write_df(ws4, item_df, start_row=4, fmt=mer_fmt)
        if not var_df.empty:
            start = 4 + len(item_df) + 3 if not item_df.empty else 4
            ws4.write(start, 0, "TOP REQUESTED VARIATIONS (Size & Color)", mer_fmt)
            write_df(ws4, var_df, start_row=start+1, fmt=mer_fmt)

        # ── Sheet 5: Team Performance (AM/Merch view) ─────────────────────────
        ws5 = wb.add_worksheet("Team Performance")
        writer.sheets["Team Performance"] = ws5
        ws5.write(0, 0, "Team Performance — Operations + Sales", wb.add_format({"bold":True,"font_size":13,"font_color":"#1B2A4A"}))
        ws5.write(1, 0, f"From {TEAM_START_DATE.date()} onwards", wb.add_format({"italic":True,"font_color":"#7A8EA8"}))
        if not team_perf.empty:
            write_df(ws5, team_perf, start_row=3, fmt=hdr_fmt)

        # ── Sheet 6: Today Priority + Unresolved ─────────────────────────────
        priority_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME","COUNTRY_CODE",
                                      "TEAM_MEMBER","BUYER_NAME","ISSUE_TYPE","PRIORITY","SENTIMENT",
                                      "IS_UNRESOLVED","CSAT_PROXY","AVG_CRT_MINS","IS_CONVERSION",
                                      "SALES_STAGE","BUYER_SUMMARY","SUGGESTED_REPLY"] if c in df.columns]
        if "SUGGESTED_REPLY" not in df.columns:
            df["SUGGESTED_REPLY"] = df["ISSUE_TYPE"].astype(str).map(lambda it: SUGGESTED_REPLIES.get(it,""))
        unres = df[df["IS_UNRESOLVED"]][priority_cols].sort_values("PRIORITY",key=lambda s:s.map({"High":0,"Medium":1,"Low":2}).fillna(3))
        unres.to_excel(writer, sheet_name="Unresolved Chats", index=False)
        ws6 = writer.sheets["Unresolved Chats"]
        for c_idx, col in enumerate(unres.columns):
            ws6.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 7: Detailed All Chats ───────────────────────────────────────
        detail_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER",
                                    "BUYER_NAME","FIRST_MSG_TIME","LAST_MSG_TIME","MSG_COUNT","ISSUE_TYPE",
                                    "PRIORITY","SENTIMENT","IS_RESOLVED","IS_UNRESOLVED","CSAT_PROXY",
                                    "AVG_CRT_MINS","IS_CONVERSION","SALES_STAGE","IS_OOS_CONFIRMED",
                                    "IS_LOST_SALE","IS_UPSELL_OPP","SIMILAR_SUGGESTED","BUYER_SUMMARY","SUGGESTED_REPLY"] if c in df.columns]
        detail = df[detail_cols].copy()
        for tc in ["FIRST_MSG_TIME","LAST_MSG_TIME"]:
            if tc in detail.columns:
                detail[tc] = pd.to_datetime(detail[tc]).dt.strftime("%Y-%m-%d %H:%M")
        detail.to_excel(writer, sheet_name="Detailed Chat Analysis", index=False)
        ws7 = writer.sheets["Detailed Chat Analysis"]
        for c_idx, col in enumerate(detail.columns):
            ws7.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 8: Key Improvements ─────────────────────────────────────────
        ws8 = wb.add_worksheet("Key Improvements")
        writer.sheets["Key Improvements"] = ws8
        ws8.set_column(0, 0, 15); ws8.set_column(1, 1, 30); ws8.set_column(2, 2, 80)
        ws8.write(0, 0, "Key Improvement Areas & Sales Opportunities", wb.add_format({"bold":True,"font_size":13,"font_color":"#1B2A4A"}))
        ws8.write(1, 0, "Priority", hdr_fmt); ws8.write(1, 1, "Area", hdr_fmt); ws8.write(1, 2, "Recommendation", hdr_fmt)
        for i, (priority_level, area, rec) in enumerate(improvements, start=2):
            clr = red_fmt if "HIGH" in priority_level else grn_fmt
            ws8.write(i, 0, priority_level, clr)
            ws8.write(i, 1, area, cell_fmt)
            ws8.write(i, 2, rec, cell_fmt)

    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def metric_card(label, value, sub="", color=""):
    return f"""<div class="metric-card {color}">
        <div class="metric-val">{value}</div>
        <div class="metric-label">{label}</div>
        {"<div class='metric-sub'>"+sub+"</div>" if sub else ""}
    </div>"""

def render_header():
    st.markdown("""
    <div class="graas-header">
        <div class="graas-logo">📊</div>
        <div>
            <h1>Chat Analyzer Dashboard v2</h1>
            <p>Graas.ai · Operations + Sales Intelligence · Shopee & Lazada</p>
        </div>
    </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    render_header()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## ⬆️ Upload Data")
        uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx","xls"], label_visibility="collapsed")
        st.markdown("---")

    if not uploaded:
        st.info("👆 Upload your chat export Excel file to begin.")
        st.markdown("""
        **What this dashboard gives you:**
        
        **📊 Operations**
        - Conversation resolution rate & CRT
        - Priority chats & unresolved cases
        - Agent team performance
        - CSAT proxy scores
        
        **💰 Sales Intelligence (NEW)**
        - Conversion funnel analysis
        - Upsell opportunity detection & action rate
        - OOS product demand tracker
        - Lost sales identification
        - Most inquired products, sizes, colours
        - Sales stage classification per conversation
        
        **🛍️ Merchandising / AM (NEW)**
        - Product variation demand (size & colour heatmap)
        - OOS items with restock priority score
        - Lost sales by store/category
        - Agent upsell performance scorecard
        
        **🎯 Key Improvements (NEW)**
        - Auto-generated prioritised recommendations
        - Data-driven coaching insights for team leads
        """)
        return

    with st.spinner("Loading data…"):
        raw_df = load_data(uploaded.read())

    if raw_df.empty:
        st.error("No data found. Check your Excel file."); return

    with st.spinner("Analysing conversations…"):
        conv_df = analyse(raw_df)

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 🔍 Filters")
        platforms = sorted(conv_df["PLATFORM"].astype(str).unique().tolist())
        sel_platforms = st.multiselect("Platform", platforms, default=platforms)
        countries = sorted(conv_df["COUNTRY_CODE"].astype(str).dropna().unique().tolist())
        sel_countries = st.multiselect("Country", countries, default=countries)
        stores = sorted(conv_df["STORE_CODE"].astype(str).dropna().unique().tolist())
        sel_stores = st.multiselect("Store Code", stores, default=stores)
        agents = sorted(conv_df["TEAM_MEMBER"].astype(str).dropna().unique().tolist())
        sel_agents = st.multiselect("Team Member", agents, default=agents)
        date_range = None
        if conv_df["LAST_MSG_TIME"].notna().any():
            min_d = conv_df["LAST_MSG_TIME"].min().date()
            max_d = conv_df["LAST_MSG_TIME"].max().date()
            date_range = st.date_input("Date Range", value=(min_d, max_d), min_value=min_d, max_value=max_d)
        st.markdown("---")
        st.markdown("## 🗓️ Export")
        today_str = datetime.today().strftime("%Y-%m-%d")
        today_ts  = pd.Timestamp(today_str)
        today_date = datetime.today().date()

    # Apply filters
    flt = conv_df.copy()
    if sel_platforms: flt = flt[flt["PLATFORM"].astype(str).isin(sel_platforms)]
    if sel_countries: flt = flt[flt["COUNTRY_CODE"].astype(str).isin(sel_countries)]
    if sel_stores:    flt = flt[flt["STORE_CODE"].astype(str).isin(sel_stores)]
    if sel_agents:    flt = flt[flt["TEAM_MEMBER"].astype(str).isin(sel_agents)]
    if date_range and len(date_range) == 2:
        flt = flt[(flt["LAST_MSG_TIME"].dt.date >= date_range[0]) & (flt["LAST_MSG_TIME"].dt.date <= date_range[1])]

    total = len(flt)
    if total == 0:
        st.warning("No conversations match the current filters."); return

    funnel = compute_conversion_funnel(flt)

    # ── TABS ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Operations Overview",
        "💰 Sales Intelligence",
        "📦 OOS & Product Demand",
        "🛍️ Merch & AM Performance",
        "👥 Team Performance",
        "🎯 Key Improvements",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1: OPERATIONS OVERVIEW
    # ══════════════════════════════════════════════════════════════════════════
    with tab1:
        resolved   = int(flt["IS_RESOLVED"].sum())
        unresolved = int(flt["IS_UNRESOLVED"].sum())
        crr        = round(resolved / total * 100, 1)
        avg_crt    = flt["AVG_CRT_MINS"].mean()
        avg_csat   = flt["CSAT_PROXY"].mean()
        hi_pri     = int((flt["PRIORITY"] == "High").sum())

        st.markdown('<div class="metric-row">' +
            metric_card("Total Conversations", f"{total:,}", color="navy") +
            metric_card("Resolved", f"{resolved:,}", f"CRR: {crr}%", color="green") +
            metric_card("Unresolved", f"{unresolved:,}", "Needs attention", color="red") +
            metric_card("Avg CRT", fmt_mins(avg_crt), "Response time", color="orange") +
            metric_card("Avg CSAT", f"{avg_csat:.1f}/5" if not np.isnan(avg_csat) else "—", "Proxy score") +
            metric_card("High Priority", f"{hi_pri:,}", "Today's urgent", color="red") +
        '</div>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="section-title">📂 Issue Type Breakdown</div>', unsafe_allow_html=True)
            ib = flt.groupby(["ISSUE_TYPE","PRIORITY"]).agg(
                Count=("CONVERSATION_ID","count"),
                Unresolved=("IS_UNRESOLVED","sum"),
                Avg_CSAT=("CSAT_PROXY","mean"),
                Avg_CRT_mins=("AVG_CRT_MINS","mean"),
            ).reset_index().sort_values("Count", ascending=False)
            ib["Avg_CSAT"] = ib["Avg_CSAT"].round(1)
            ib["Avg_CRT_mins"] = ib["Avg_CRT_mins"].round(0).fillna(0).astype(int)
            ib["Unresolved"] = ib["Unresolved"].astype(int)
            st.dataframe(ib, use_container_width=True, height=300, hide_index=True)

        with c2:
            st.markdown('<div class="section-title">🏪 Store Performance</div>', unsafe_allow_html=True)
            sp = flt.groupby(["STORE_CODE","PLATFORM"]).agg(
                Conversations=("CONVERSATION_ID","count"),
                Unresolved=("IS_UNRESOLVED","sum"),
                Avg_CSAT=("CSAT_PROXY","mean"),
                Avg_CRT_mins=("AVG_CRT_MINS","mean"),
            ).reset_index().sort_values("Conversations", ascending=False)
            sp["Avg_CSAT"] = sp["Avg_CSAT"].round(1)
            sp["Avg_CRT_mins"] = sp["Avg_CRT_mins"].round(0).fillna(0).astype(int)
            sp["Unresolved"] = sp["Unresolved"].astype(int)
            sp["CRR%"] = ((sp["Conversations"]-sp["Unresolved"])/sp["Conversations"]*100).round(1)
            st.dataframe(sp, use_container_width=True, height=300, hide_index=True)

        st.markdown('<div class="section-title">📋 Unresolved Conversations</div>', unsafe_allow_html=True)
        unres_show = flt[flt["IS_UNRESOLVED"]][
            [c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER","ISSUE_TYPE",
                          "PRIORITY","SENTIMENT","CSAT_PROXY","AVG_CRT_MINS","SALES_STAGE","BUYER_SUMMARY"] if c in flt.columns]
        ].sort_values("PRIORITY", key=lambda s: s.map({"High":0,"Medium":1,"Low":2}).fillna(3))
        st.dataframe(unres_show, use_container_width=True, height=350, hide_index=True)

        # WoW / MoM
        st.markdown('<div class="section-title">📈 Trend Analysis (WoW / MoM)</div>', unsafe_allow_html=True)
        wow, mom = compute_wow_mom(flt)
        t1, t2 = st.tabs(["Week-on-Week","Month-on-Month"])
        with t1:
            if not wow.empty:
                wow["WEEK"] = wow["WEEK"].astype(str)
                st.dataframe(wow, use_container_width=True, hide_index=True)
        with t2:
            if not mom.empty:
                mom["MONTH"] = mom["MONTH"].astype(str)
                st.dataframe(mom, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2: SALES INTELLIGENCE
    # ══════════════════════════════════════════════════════════════════════════
    with tab2:
        st.markdown('<div class="section-title-sales">💰 Sales Intelligence Dashboard</div>', unsafe_allow_html=True)

        # KPI row
        st.markdown('<div class="metric-row">' +
            metric_card("Conversions Detected", f"{funnel.get('converted',0):,}", f"Conv Rate: {funnel.get('conv_rate_pct',0)}%", color="green") +
            metric_card("Upsell Opportunities", f"{funnel.get('upsell_opportunities',0):,}", f"Action Rate: {funnel.get('upsell_act_rate_pct',0)}%", color="orange") +
            metric_card("OOS Inquiries", f"{funnel.get('oos_total',0):,}", "Demand captured", color="red") +
            metric_card("Lost Sales", f"{funnel.get('lost_sales',0):,}", f"{funnel.get('lost_rate_pct',0)}% of chats", color="red") +
            metric_card("Product Inquiries", f"{funnel.get('product_inquiries',0):,}", "Warm leads", color="purple") +
        '</div>', unsafe_allow_html=True)

        # Funnel visualization
        st.markdown('<div class="section-title-sales">🔽 Sales Funnel</div>', unsafe_allow_html=True)
        funnel_data = {
            "Stage": ["All Conversations","Product Inquiries","High Intent","Converted","Lost Sales"],
            "Count": [
                funnel.get("total_conversations",0),
                funnel.get("product_inquiries",0),
                funnel.get("high_intent",0),
                funnel.get("converted",0),
                funnel.get("lost_sales",0),
            ],
        }
        funnel_df = pd.DataFrame(funnel_data)
        funnel_df["% of Total"] = (funnel_df["Count"] / max(funnel.get("total_conversations",1),1) * 100).round(1)
        st.dataframe(funnel_df, use_container_width=True, hide_index=True,
            column_config={"Count": st.column_config.NumberColumn(format="%d"),
                           "% of Total": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)})

        # Sales stage breakdown
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="section-title-sales">📊 Sales Stage Distribution</div>', unsafe_allow_html=True)
            stage_ct = flt["SALES_STAGE"].value_counts().reset_index()
            stage_ct.columns = ["Stage","Count"]
            stage_ct["% Share"] = (stage_ct["Count"] / total * 100).round(1)
            st.dataframe(stage_ct, use_container_width=True, hide_index=True,
                column_config={"% Share": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)})

        with c2:
            st.markdown('<div class="section-title-sales">🔁 Upsell Action Scorecard</div>', unsafe_allow_html=True)
            total_upsell = funnel.get("upsell_opportunities",0)
            acted_upsell = int(flt["SIMILAR_SUGGESTED"].sum())
            missed_upsell = total_upsell - acted_upsell
            upsell_rate = funnel.get("upsell_act_rate_pct",0)
            st.markdown(f"""
            <div class="upsell-box">
            <b>Upsell Opportunities:</b> {total_upsell}<br>
            <b>Similar Product Suggested:</b> {acted_upsell} ✅<br>
            <b>Missed (No suggestion):</b> {missed_upsell} ❌<br>
            <b>Action Rate:</b> {upsell_rate}%<br><br>
            <i>💡 Each missed upsell is a potential lost basket. 
            Target: suggest alternative on 100% of upsell-signalled chats.</i>
            </div>""", unsafe_allow_html=True)

        # By store
        st.markdown('<div class="section-title-sales">🏪 Sales Intelligence by Store</div>', unsafe_allow_html=True)
        store_sales = flt.groupby(["STORE_CODE","COUNTRY_CODE"]).agg(
            Conversations=("CONVERSATION_ID","count"),
            Conversions=("IS_CONVERSION","sum"),
            Lost_Sales=("IS_LOST_SALE","sum"),
            OOS_Inquiries=("IS_OOS_INQUIRY","sum"),
            Upsell_Opps=("IS_UPSELL_OPP","sum"),
            Similar_Suggested=("SIMILAR_SUGGESTED","sum"),
            Product_Inquiries=("ISSUE_TYPE", lambda x: (x=="Product Inquiry").sum()),
        ).reset_index()
        store_sales["Conv_Rate_%"] = (store_sales["Conversions"]/store_sales["Conversations"]*100).round(1)
        store_sales["Lost_Rate_%"] = (store_sales["Lost_Sales"]/store_sales["Conversations"]*100).round(1)
        store_sales["Upsell_Act_%"] = (store_sales["Similar_Suggested"]/store_sales["Upsell_Opps"].replace(0,np.nan)*100).round(1)
        st.dataframe(store_sales.sort_values("Conversations", ascending=False), use_container_width=True, hide_index=True)

        # Detailed sales conversations
        with st.expander("🔎 View Detailed Sales-Flagged Conversations"):
            sales_detail = flt[
                (flt["IS_CONVERSION"]==True) | (flt["IS_LOST_SALE"]==True) |
                (flt["IS_OOS_CONFIRMED"]==True) | (flt["IS_UPSELL_OPP"]==True)
            ][
                [c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER","SALES_STAGE",
                               "IS_CONVERSION","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_UPSELL_OPP",
                               "SIMILAR_SUGGESTED","SENTIMENT","BUYER_SUMMARY"] if c in flt.columns]
            ].sort_values("SALES_STAGE" if "SALES_STAGE" in flt.columns else "CONVERSATION_ID")
            st.dataframe(sales_detail, use_container_width=True, height=400, hide_index=True,
                column_config={
                    "IS_CONVERSION": st.column_config.CheckboxColumn("Converted?"),
                    "IS_LOST_SALE": st.column_config.CheckboxColumn("Lost?"),
                    "IS_OOS_CONFIRMED": st.column_config.CheckboxColumn("OOS?"),
                    "IS_UPSELL_OPP": st.column_config.CheckboxColumn("Upsell Opp?"),
                    "SIMILAR_SUGGESTED": st.column_config.CheckboxColumn("Alt Suggested?"),
                })

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3: OOS & PRODUCT DEMAND
    # ══════════════════════════════════════════════════════════════════════════
    with tab3:
        st.markdown('<div class="section-title-sales">📦 Out-of-Stock Demand Tracker</div>', unsafe_allow_html=True)
        st.caption("Every OOS inquiry = a demand signal. Share with Buying/AM/Merch team weekly.")

        oos_df = build_oos_tracker(raw_df, flt)
        if not oos_df.empty:
            oos_kpi1, oos_kpi2, oos_kpi3 = st.columns(3)
            with oos_kpi1:
                st.metric("Total OOS Conversations", len(oos_df))
            with oos_kpi2:
                lost_oos = int(oos_df["LOST_SALE"].sum()) if "LOST_SALE" in oos_df.columns else 0
                st.metric("Lost Sales from OOS", lost_oos)
            with oos_kpi3:
                sugg_oos = int(oos_df["SIMILAR_SUGGESTED"].sum()) if "SIMILAR_SUGGESTED" in oos_df.columns else 0
                st.metric("Alternative Suggested", f"{sugg_oos}/{len(oos_df)}")
            st.dataframe(oos_df, use_container_width=True, height=400, hide_index=True,
                column_config={
                    "LOST_SALE": st.column_config.CheckboxColumn("Lost Sale"),
                    "SIMILAR_SUGGESTED": st.column_config.CheckboxColumn("Alt Suggested"),
                })
        else:
            st.info("No confirmed OOS inquiries detected in the filtered data.")

        # Product demand
        st.markdown('<div class="section-title-merch">🔍 Most Inquired Products & Variations</div>', unsafe_allow_html=True)
        item_df, var_df = build_product_inquiry_summary(raw_df, flt)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**📦 Top Inquired Item IDs**")
            st.caption("These item IDs appear most in buyer messages — prioritise stock & listing quality")
            if not item_df.empty:
                st.dataframe(item_df, use_container_width=True, height=300, hide_index=True,
                    column_config={"Inquiry Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=int(item_df["Inquiry Count"].max()))})
            else:
                st.info("No item IDs detected.")

        with c2:
            st.markdown("**📐 Top Requested Variations (Size & Color)**")
            st.caption("Demand signals for merchandising — stock gaps & listing optimisation")
            if not var_df.empty:
                st.dataframe(var_df, use_container_width=True, height=300, hide_index=True,
                    column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=int(var_df["Count"].max()))})
            else:
                st.info("No variation mentions detected.")

        # Issue type → product inquiry breakdown
        st.markdown('<div class="section-title-merch">📊 Product Inquiry Analysis</div>', unsafe_allow_html=True)
        pi_df = flt[flt["ISSUE_TYPE"] == "Product Inquiry"].copy()
        if not pi_df.empty:
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Product Inquiries", len(pi_df))
            with c2: st.metric("Converted", int(pi_df["IS_CONVERSION"].sum()), f"{int(pi_df['IS_CONVERSION'].sum())/len(pi_df)*100:.1f}%")
            with c3: st.metric("OOS Confirmed", int(pi_df["IS_OOS_CONFIRMED"].sum()))

            pi_by_store = pi_df.groupby("STORE_CODE").agg(
                Inquiries=("CONVERSATION_ID","count"),
                Conversions=("IS_CONVERSION","sum"),
                OOS=("IS_OOS_CONFIRMED","sum"),
                Lost=("IS_LOST_SALE","sum"),
            ).reset_index()
            pi_by_store["Conv_%"] = (pi_by_store["Conversions"]/pi_by_store["Inquiries"]*100).round(1)
            st.dataframe(pi_by_store.sort_values("Inquiries", ascending=False), use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 4: MERCH & AM PERFORMANCE
    # ══════════════════════════════════════════════════════════════════════════
    with tab4:
        st.markdown('<div class="section-title-merch">🛍️ Merchandising & Account Management View</div>', unsafe_allow_html=True)
        st.caption("Chat data as a sales signal — for AM/Merch teams to action")

        # AM scorecard
        am_data = flt.groupby(["STORE_CODE","COUNTRY_CODE","PLATFORM"]).agg(
            Total_Chats=("CONVERSATION_ID","count"),
            Product_Inquiries=("ISSUE_TYPE", lambda x: (x=="Product Inquiry").sum()),
            Conversions=("IS_CONVERSION","sum"),
            Lost_Sales=("IS_LOST_SALE","sum"),
            OOS_Hits=("IS_OOS_CONFIRMED","sum"),
            Upsell_Opps=("IS_UPSELL_OPP","sum"),
            Alt_Suggested=("SIMILAR_SUGGESTED","sum"),
            Avg_CSAT=("CSAT_PROXY","mean"),
            Unresolved=("IS_UNRESOLVED","sum"),
        ).reset_index()
        am_data["Conv_Rate_%"] = (am_data["Conversions"]/am_data["Total_Chats"]*100).round(1)
        am_data["Lost_Rate_%"] = (am_data["Lost_Sales"]/am_data["Total_Chats"]*100).round(1)
        am_data["OOS_Rate_%"]  = (am_data["OOS_Hits"]/am_data["Total_Chats"]*100).round(1)
        am_data["Upsell_Act_%"]= (am_data["Alt_Suggested"]/am_data["Upsell_Opps"].replace(0,np.nan)*100).round(1)
        am_data["Avg_CSAT"] = am_data["Avg_CSAT"].round(1)
        st.markdown("**🏪 Per-Store AM Scorecard**")
        st.dataframe(am_data.sort_values("Total_Chats", ascending=False), use_container_width=True, hide_index=True,
            column_config={
                "Conv_Rate_%": st.column_config.ProgressColumn("Conv%", format="%.1f%%", min_value=0, max_value=100),
                "Lost_Rate_%": st.column_config.ProgressColumn("Lost%", format="%.1f%%", min_value=0, max_value=100),
                "OOS_Rate_%": st.column_config.ProgressColumn("OOS%", format="%.1f%%", min_value=0, max_value=100),
                "Upsell_Act_%": st.column_config.ProgressColumn("Upsell Act%", format="%.1f%%", min_value=0, max_value=100),
            })

        # Lost sales deep dive
        st.markdown('<div class="section-title-merch">💸 Lost Sales Analysis</div>', unsafe_allow_html=True)
        lost_df = flt[flt["IS_LOST_SALE"] == True].copy()
        if not lost_df.empty:
            c1, c2 = st.columns(2)
            with c1:
                lost_by_store = lost_df.groupby("STORE_CODE").agg(
                    Lost_Sales=("CONVERSATION_ID","count"),
                    OOS_Related=("IS_OOS_CONFIRMED","sum"),
                ).reset_index().sort_values("Lost_Sales", ascending=False)
                st.markdown("**Lost Sales by Store**")
                st.dataframe(lost_by_store, use_container_width=True, hide_index=True)
            with c2:
                lost_by_issue = lost_df.groupby("ISSUE_TYPE").agg(
                    Lost_Sales=("CONVERSATION_ID","count")
                ).reset_index().sort_values("Lost_Sales", ascending=False)
                st.markdown("**Lost Sales by Root Cause**")
                st.dataframe(lost_by_issue, use_container_width=True, hide_index=True)
            with st.expander("🔎 Lost Sale Conversations Detail"):
                st.dataframe(
                    lost_df[[c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","ISSUE_TYPE",
                                          "IS_OOS_CONFIRMED","SENTIMENT","SALES_STAGE","BUYER_SUMMARY"] if c in lost_df.columns]],
                    use_container_width=True, height=300, hide_index=True)
        else:
            st.success("✅ No lost sales detected in the filtered data.")

        # Restock priority
        st.markdown('<div class="section-title-merch">📦 Restock Priority List</div>', unsafe_allow_html=True)
        oos_df2 = build_oos_tracker(raw_df, flt)
        if not oos_df2.empty:
            st.caption("Items confirmed OOS by seller, sorted by demand frequency — share with buying team")
            restock = oos_df2.groupby(["STORE_CODE","ITEM_IDS_INQUIRED"]).agg(
                Demand_Count=("CONVERSATION_ID","count"),
                Lost_Sales=("LOST_SALE","sum"),
                Color_Requested=("COLOR_REQUESTED", lambda x: " | ".join(set(str(v) for v in x if v and v!="—"))),
                Size_Requested=("SIZE_REQUESTED", lambda x: " | ".join(set(str(v) for v in x if v and v!="—"))),
            ).reset_index().sort_values("Demand_Count", ascending=False)
            restock["Priority Score"] = restock["Demand_Count"] + restock["Lost_Sales"]*2
            st.dataframe(restock.sort_values("Priority Score", ascending=False).head(30),
                         use_container_width=True, hide_index=True,
                         column_config={"Priority Score": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=int(restock["Priority Score"].max()))})
        else:
            st.info("No OOS data available for restock analysis.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 5: TEAM PERFORMANCE
    # ══════════════════════════════════════════════════════════════════════════
    with tab5:
        st.markdown('<div class="section-title">👥 Team Performance</div>', unsafe_allow_html=True)
        st.caption(f"Data from {TEAM_START_DATE.date()} onwards · Sales metrics included")
        team_perf = compute_team_performance(flt)
        if not team_perf.empty:
            st.dataframe(team_perf, use_container_width=True, hide_index=True,
                column_config={
                    "CRR_%":          st.column_config.ProgressColumn("CRR%",    format="%.1f%%", min_value=0, max_value=100),
                    "Conv_Rate_%":    st.column_config.ProgressColumn("Conv%",   format="%.1f%%", min_value=0, max_value=100),
                    "Upsell_Act_Rate_%": st.column_config.ProgressColumn("Upsell Act%", format="%.1f%%", min_value=0, max_value=100),
                    "Avg_CSAT":       st.column_config.NumberColumn("CSAT",      format="%.2f"),
                    "Avg_CRT_mins":   st.column_config.NumberColumn("CRT (min)", format="%.1f"),
                })
            # Highlight best/worst
            if len(team_perf) > 1:
                st.markdown("---")
                c1, c2, c3 = st.columns(3)
                with c1:
                    best_crt = team_perf.dropna(subset=["Avg_CRT_mins"]).nsmallest(1,"Avg_CRT_mins")
                    if not best_crt.empty:
                        st.success(f"⚡ **Fastest Responder:** {best_crt.iloc[0]['TEAM_MEMBER']} ({fmt_mins(best_crt.iloc[0]['Avg_CRT_mins'])})")
                with c2:
                    best_conv = team_perf.nlargest(1,"Conv_Rate_%")
                    if not best_conv.empty:
                        st.success(f"💰 **Top Converter:** {best_conv.iloc[0]['TEAM_MEMBER']} ({best_conv.iloc[0]['Conv_Rate_%']:.1f}%)")
                with c3:
                    best_csat = team_perf.nlargest(1,"Avg_CSAT")
                    if not best_csat.empty:
                        st.success(f"⭐ **Top CSAT:** {best_csat.iloc[0]['TEAM_MEMBER']} ({best_csat.iloc[0]['Avg_CSAT']:.2f}/5)")
        else:
            st.info("No team data available for the selected date range.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 6: KEY IMPROVEMENTS
    # ══════════════════════════════════════════════════════════════════════════
    with tab6:
        st.markdown('<div class="section-title-merch">🎯 Key Improvement Areas</div>', unsafe_allow_html=True)
        st.caption("Auto-generated from your chat data. Prioritised by business impact.")
        improvements = generate_key_improvements(flt, funnel)
        for priority_level, area, rec in improvements:
            icon = "🔴" if "HIGH" in priority_level else ("🟡" if "MEDIUM" in priority_level else "🟢")
            bg = "#FEF2F2" if "HIGH" in priority_level else ("#FFFBEB" if "MEDIUM" in priority_level else "#F0FDF4")
            border = "#E74C3C" if "HIGH" in priority_level else ("#F59E0B" if "MEDIUM" in priority_level else "#22C55E")
            st.markdown(f"""
            <div style="background:{bg};border-left:4px solid {border};border-radius:6px;padding:0.9rem 1.1rem;margin-bottom:0.8rem">
            <b>{icon} {priority_level} · {area}</b><br>
            <span style="font-size:0.88rem;color:#374151">{rec}</span>
            </div>""", unsafe_allow_html=True)

    # ── Export ─────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 📥 Download Reports")
        cutoff_7d   = today_ts - pd.Timedelta(days=6)
        conv_7day   = conv_df[conv_df["LAST_MSG_TIME"] >= cutoff_7d].copy()

        if st.button("📊 Generate Last 7 Days Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_data = build_excel(conv_7day, raw_df, today_str)
            st.download_button(
                label=f"📥 Download 7-Day ({cutoff_7d.date()} → {today_date})",
                data=excel_data,
                file_name=f"Chat_Analysis_v2_7Days_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        if st.button("📊 Generate Filtered Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_data = build_excel(flt, raw_df, today_str)
            st.download_button(
                label="📥 Download Filtered View",
                data=excel_data,
                file_name=f"Chat_Analysis_v2_Filtered_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption("Click Generate first, then Download. v2 includes 8-sheet Excel with full sales intelligence.")


if __name__ == "__main__":
    main()
