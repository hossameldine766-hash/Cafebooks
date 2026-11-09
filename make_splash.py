# ═══════════════════════════════════════════════════════════════════
# إنشاء شاشة Splash لبرنامج CafeBooks
# ═══════════════════════════════════════════════════════════════════
from PIL import Image, ImageDraw, ImageFont
import os
import arabic_reshaper
from bidi.algorithm import get_display

# ═══════════════════════════════════════════════════════════════════
# الإعدادات
# ═══════════════════════════════════════════════════════════════════
W, H = 600, 480

# الألوان (النمط 1 — كلاسيكي دافئ)
BG_COLOR = '#1a5276'         # بنّي داكن (لون القهوة)
GOLD = '#D4AF37'             # ذهبي دافئ
BEIGE = '#F5DEB3'            # بيج فاتح
WHITE = '#FFFFFF'
GRAY = '#BDC3C7'

# الشعار
LOGO_FILE = 'm.png'
LOGO_SIZE = 140

# النصوص
PROGRAM_NAME = "CafeBooks"
PROGRAM_DESC = "نظام حسابات الكافيتيريا المتكامل"
LOADING_TEXT = "جارٍ التحميل..."

COMPANY_NAME = "شركة الزهراء للبرمجيات"
MANAGER_NAME = "المحاسب / حسام الدين سيد ربيع"
PHONE = "01003796984"


# ═══════════════════════════════════════════════════════════════════
# دالة معالجة النص العربي
# ═══════════════════════════════════════════════════════════════════
def ar(text):
    if not text:
        return ''
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


# ═══════════════════════════════════════════════════════════════════
# تحميل الخطوط
# ═══════════════════════════════════════════════════════════════════
def load_font(name, size):
    paths = [
        f"C:/Windows/Fonts/{name}",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except:
            pass
    return ImageFont.load_default()


font_program    = load_font("arialbd.ttf", 38)
font_desc       = load_font("arial.ttf", 16)
font_loading    = load_font("arial.ttf", 15)
font_company    = load_font("arialbd.ttf", 14)
font_manager    = load_font("arial.ttf", 12)
font_phone      = load_font("arialbd.ttf", 13)


# ═══════════════════════════════════════════════════════════════════
# إنشاء الصورة
# ═══════════════════════════════════════════════════════════════════
img = Image.new('RGB', (W, H), BG_COLOR)
draw = ImageDraw.Draw(img)

# ─── الحدود الذهبية ───
draw.rectangle([(0, 0), (W - 1, H - 1)], outline=GOLD, width=3)
draw.rectangle([(8, 8), (W - 9, H - 9)], outline=GOLD, width=1)

# ═══════════════════════════════════════════════════════════════════
# 1) الشعار
# ═══════════════════════════════════════════════════════════════════
if os.path.exists(LOGO_FILE):
    try:
        logo = Image.open(LOGO_FILE).convert('RGBA')
        logo.thumbnail((LOGO_SIZE, LOGO_SIZE), Image.Resampling.LANCZOS)
        logo_x = (W - logo.width) // 2
        logo_y = 35
        img.paste(logo, (logo_x, logo_y), logo)
    except Exception as e:
        print(f"⚠️ فشل تحميل الشعار: {e}")
else:
    # نص بديل
    fallback = "☕"
    bbox = draw.textbbox((0, 0), fallback, font=font_program)
    fw = bbox[2] - bbox[0]
    draw.text(((W - fw) / 2, 60), fallback, fill=GOLD, font=font_program)

# ═══════════════════════════════════════════════════════════════════
# 2) خط ذهبي فاصل أعلى
# ═══════════════════════════════════════════════════════════════════
y_sep1 = 200
draw.line([(80, y_sep1), (W - 80, y_sep1)], fill=GOLD, width=2)

# ═══════════════════════════════════════════════════════════════════
# 3) اسم البرنامج
# ═══════════════════════════════════════════════════════════════════
name_text = f"☕ {PROGRAM_NAME}"
bbox = draw.textbbox((0, 0), name_text, font=font_program)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) / 2, 215), name_text, fill=GOLD, font=font_program)

# ═══════════════════════════════════════════════════════════════════
# 4) الوصف العربي
# ═══════════════════════════════════════════════════════════════════
desc_text = ar(PROGRAM_DESC)
bbox = draw.textbbox((0, 0), desc_text, font=font_desc)
dw = bbox[2] - bbox[0]
draw.text(((W - dw) / 2, 265), desc_text, fill=BEIGE, font=font_desc)

# ═══════════════════════════════════════════════════════════════════
# 5) خط ذهبي فاصل أسفل
# ═══════════════════════════════════════════════════════════════════
y_sep2 = 300
draw.line([(80, y_sep2), (W - 80, y_sep2)], fill=GOLD, width=2)

# ═══════════════════════════════════════════════════════════════════
# 6) حالة التحميل
# ═══════════════════════════════════════════════════════════════════
loading_text = ar(LOADING_TEXT)
bbox = draw.textbbox((0, 0), loading_text, font=font_loading)
lw = bbox[2] - bbox[0]
draw.text(((W - lw) / 2, 318), loading_text, fill=WHITE, font=font_loading)

# ═══════════════════════════════════════════════════════════════════
# 7) بيانات الشركة
# ═══════════════════════════════════════════════════════════════════

company_text = ar(COMPANY_NAME)
bbox = draw.textbbox((0, 0), company_text, font=font_company)
cw = bbox[2] - bbox[0]
draw.text(((W - cw) / 2, 385), company_text, fill=GOLD, font=font_company)

manager_text = ar(MANAGER_NAME)
bbox = draw.textbbox((0, 0), manager_text, font=font_manager)
mw = bbox[2] - bbox[0]
draw.text(((W - mw) / 2, 412), manager_text, fill=GRAY, font=font_manager)

phone_text = ar(f"📞 {PHONE}")
bbox = draw.textbbox((0, 0), phone_text, font=font_phone)
pw = bbox[2] - bbox[0]
draw.text(((W - pw) / 2, 435), phone_text, fill=GRAY, font=font_phone)

# ═══════════════════════════════════════════════════════════════════
# 8) الحفظ
# ═══════════════════════════════════════════════════════════════════
output_file = 'splash.png'
img.save(output_file, 'PNG')
print(f"✅ تم إنشاء شاشة Splash: {output_file}")
print(f"📐 الأبعاد: {W} × {H}")
print(f"📁 المكان: {os.path.abspath(output_file)}")
