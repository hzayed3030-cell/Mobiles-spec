"""
app.py
تطبيق Streamlit المتكامل والمطور لمقارنة مواصفات الهواتف الذكية (Mobile Specification Comparison Dashboard)
الإصدار المطور:
- نظام تحكم فوري وتفاعلي بحجم الخط وتكبير الواجهة (Interactive Zoom & Font Scale) يعمل بدقة وسلاسة 100%.
- إصلاح تداخل النصوص وضبط المسافات وارتفاعات الأسطر (Line-Heights) لمنع أي تداخل نهائياً.
- تصميم عصري فاخر بتقسيم واضح، وبطاقات تفاعلية أنيقة، وجداول فائقة الوضوح.
- عناوين التبويبات (Tabs Headers) باللون الأبيض الناصع (#FFFFFF) مع تدرجات لونية جذابة للتبويب النشط والتبويبات غير النشطة.
- تلوين الخلية الفائزة بالأخضر مع شارة "⭐ الأفضل" وعمود "سبب الأفضلية".
- مركز تصدير متكامل (Excel RTL / PDF RTL / معاينة الطباعة التفاعلية).
"""

import re
import streamlit as st
import pandas as pd
import altair as alt
import streamlit.components.v1 as components

# استيراد مستودع البيانات وكتالوج الماركات وخدمات التصدير ومحرك التحليل
from specs_data import (
    SPEC_DEFINITIONS,
    BRANDS_CATALOG,
    get_mobile_specs,
    get_popular_mobiles_list,
    search_phone_suggestions
)
from export_service import (
    export_comparison_to_excel,
    export_comparison_to_pdf,
    generate_html_print_view,
    export_s26_settings_to_excel,
    export_phone_settings_to_excel,
    get_phone_settings_dataframe
)
from analyzer import (
    evaluate_spec_winner,
    evaluate_overall_winner
)
from git_service import (
    DEFAULT_PROJECT_GITHUB_URL,
    DEFAULT_PROJECT_WEB_URL,
    get_git_info,
    get_git_details,
    get_git_status_summary,
    set_git_remote,
    check_large_build_dirs,
    delete_build_dirs,
    export_project_to_github,
    update_project_on_github
)

# -----------------------------------------------------------------------------
# 1. إعدادات الصفحة واستغلال العرض الكامل للشاشة (Wide Desktop Layout)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="داشبورد مقارنة مواصفات الهواتف الذكية | Mobile Specs Dashboard",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تهيئة session_state
if "searched" not in st.session_state:
    st.session_state.searched = True  # تلقائي لعرض البيانات فوراً
if "comparison_data" not in st.session_state:
    st.session_state.comparison_data = None
if "final_verdict" not in st.session_state:
    st.session_state.final_verdict = None
if "selected_phones" not in st.session_state:
    st.session_state.selected_phones = ["Samsung Galaxy S26 Plus", "iPhone 17 Pro Max", "OPPO Find X8 Pro"]
if "num_mobiles_input" not in st.session_state:
    st.session_state.num_mobiles_input = 3

brand_options = list(BRANDS_CATALOG.keys())

default_brand_models = [
    ("Samsung (سامسونج)", "Galaxy S26 Plus"),
    ("Apple iPhone (آبل)", "iPhone 17 Pro Max"),
    ("OPPO (أوبو)", "Find X8 Pro"),
    ("Vivo & iQOO (فيفو)", "Vivo X200 Pro"),
    ("Xiaomi & Poco & Redmi (شاومي)", "Xiaomi 15 Ultra"),
    ("Google Pixel (جوجل بكسل)", "Google Pixel 9 Pro XL"),
    ("OnePlus (وان بلس)", "OnePlus 13"),
    ("Honor (هونر)", "Honor Magic 7 Pro"),
]

for idx, (def_b, def_m) in enumerate(default_brand_models):
    b_key = f"brand_val_{idx}"
    m_key = f"model_val_{idx}"
    if b_key not in st.session_state:
        st.session_state[b_key] = def_b
    if m_key not in st.session_state:
        st.session_state[m_key] = def_m


# -----------------------------------------------------------------------------
# 2. إعدادات حجم الخط الثابتة والواجهة (Fixed 13.5px Standard Typography)
# -----------------------------------------------------------------------------
cur_zoom = {
    "zoom_factor": 1.15,
    "h1": "24px", "h2": "18.5px", "h3": "16px", "body": "14.5px",
    "label": "13.5px", "input_font": "13.5px", "input_h": "38px",
    "tab_font": "17.5px", "tbl_th": "15.5px", "tbl_td": "14px", "tbl_pad": "9px 12px",
    "card_pad": "12px 16px", "kpi_num": "22px", "kpi_title": "14.5px"
}

# -----------------------------------------------------------------------------
# 3. إعدادات الشريط الجانبي (Sidebar Controls)
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ الإعدادات العامة (Settings)")
    
    show_category_headers = st.checkbox("إظهار فواصل الفئات الملونة (Categories)", value=True, key="cb_cat")

    st.markdown("<hr style='margin: 18px 0; border-color: rgba(255,255,255,0.15);'>", unsafe_allow_html=True)
    
    st.markdown("#### ⚡ مقارنات سريعة جاهزة (Presets)")
    preset_choice = st.selectbox(
        "اختر للملء التلقائي:",
        options=[
            "تخصيص يدوي (Custom)",
            "مقارنة سامسونج وآبل وأوبو (Flagships Battle)",
            "مقارنة M51 و S25+ و iPhone 17 (Popular)",
            "أفضل كاميرات 2025 (S25 Ultra vs Find X8 Pro vs X200 Pro)",
            "مقارنة هواتف آبل (iPhone 16 vs 16 Pro vs 16 Pro Max)",
            "أقوى هواتف الأندرويد (S25 Ultra vs OnePlus 13 vs Pixel 9 Pro XL)"
        ],
        index=0,
        key="preset_select_box"
    )

    preset_map = {
        "مقارنة سامسونج وآبل وأوبو (Flagships Battle)": [
            ("Samsung (سامسونج)", "Galaxy S26 Plus"),
            ("Apple iPhone (آبل)", "iPhone 17 Pro Max"),
            ("OPPO (أوبو)", "Find X8 Pro")
        ],
        "مقارنة M51 و S25+ و iPhone 17 (Popular)": [
            ("Samsung (سامسونج)", "Galaxy M51"),
            ("Samsung (سامسونج)", "Galaxy S25 Plus"),
            ("Apple iPhone (آبل)", "iPhone 17 Pro Max")
        ],
        "أفضل كاميرات 2025 (S25 Ultra vs Find X8 Pro vs X200 Pro)": [
            ("Samsung (سامسونج)", "Galaxy S25 Ultra"),
            ("OPPO (أوبو)", "Find X8 Pro"),
            ("Vivo & iQOO (فيفو)", "Vivo X200 Pro")
        ],
        "مقارنة هواتف آبل (iPhone 16 vs 16 Pro vs 16 Pro Max)": [
            ("Apple iPhone (آبل)", "iPhone 16"),
            ("Apple iPhone (آبل)", "iPhone 16 Pro"),
            ("Apple iPhone (آبل)", "iPhone 16 Pro Max")
        ],
        "أقوى هواتف الأندرويد (S25 Ultra vs OnePlus 13 vs Pixel 9 Pro XL)": [
            ("Samsung (سامسونج)", "Galaxy S25 Ultra"),
            ("OnePlus (وان بلس)", "OnePlus 13"),
            ("Google Pixel (جوجل بكسل)", "Google Pixel 9 Pro XL")
        ]
    }
    
    if preset_choice in preset_map:
        for i, (b_val, m_val) in enumerate(preset_map[preset_choice]):
            st.session_state[f"brand_val_{i}"] = b_val
            st.session_state[f"model_val_{i}"] = m_val

    st.markdown("<hr style='margin: 18px 0; border-color: rgba(255,255,255,0.15);'>", unsafe_allow_html=True)
    st.markdown("<h4 style='color: #FFFFFF; font-weight: 800; margin-bottom: 8px;'>🐙 مستودع المشروع (GitHub Hub)</h4>", unsafe_allow_html=True)
    g_info = get_git_info()
    st.markdown(f"""
    <div style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12); padding: 10px 14px; border-radius: 10px; font-size: 12.5px; color: #E2E8F0;">
        <div>🌿 <strong>الفرع (Branch):</strong> <code style="color: #93C5FD; background: rgba(0,0,0,0.3); padding: 2px 6px; border-radius: 4px;">{g_info.get('branch', 'main')}</code></div>
        <div style="margin-top: 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">🔗 <strong>المستودع:</strong> <small style="color: #CBD5E1;">{g_info.get('remote', '-')}</small></div>
    </div>
    """, unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 4. حقن الـ CSS الاحترافي المطور (Clean, Uncluttered, Perfect RTL & Zoom)
# -----------------------------------------------------------------------------
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&family=Segoe+UI&display=swap');

    :root {{
        --app-zoom: {cur_zoom['zoom_factor']};
        --app-font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif;
    }}

    /* تطبيق الخط واتجاه الصفحة العام بنظافة دون إتلاف العناصر الداخلية */
    html, body, .stApp {{
        font-family: var(--app-font-family) !important;
        direction: rtl;
        text-align: right;
    }}

    /* تفعيل التكبير الحقيقي والمريح لكامل الواجهة */
    .main .block-container {{
        zoom: var(--app-zoom);
        padding-top: 15px !important;
        padding-bottom: 60px !important;
        padding-left: 28px !important;
        padding-right: 28px !important;
        max-width: 100% !important;
    }}

    /* ضبط المسافات وارتفاعات الأسطر لمنع تداخل النصوص */
    p, .stMarkdown p, span, label {{
        line-height: 1.7 !important;
        word-break: normal;
        overflow-wrap: break-word;
    }}

    /* الشريط الجانبي المفتوح (Sidebar Expanded) */
    [data-testid="stSidebar"][aria-expanded="true"] {{
        background-color: #0F172A !important;
        min-width: 320px !important;
        max-width: 420px !important;
        box-shadow: -4px 0 20px rgba(0, 0, 0, 0.3) !important;
    }}

    [data-testid="stSidebar"][aria-expanded="true"] > div:first-child {{
        background-color: #0F172A !important;
        padding: 24px 18px !important;
    }}

    /* الشريط الجانبي عند الإغلاق التام (Sidebar Collapsed) - منع أي تداخل نهائياً */
    [data-testid="stSidebar"][aria-expanded="false"] {{
        margin-left: 0 !important;
        margin-right: 0 !important;
        padding: 0 !important;
        width: 0 !important;
        min-width: 0 !important;
        max-width: 0 !important;
        visibility: hidden !important;
        pointer-events: none !important;
        border: none !important;
        box-shadow: none !important;
    }}

    [data-testid="stSidebar"][aria-expanded="false"] > div {{
        display: none !important;
        padding: 0 !important;
        width: 0 !important;
    }}

    /* زر إظهار وإخفاء الشريط الجانبي (Sidebar Toggle Arrow) */
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {{
        z-index: 999999 !important;
        top: 14px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    }}

    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="collapsedControl"] button {{
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        border: 2px solid #3B82F6 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4) !important;
        transition: all 0.2s ease-in-out !important;
    }}

    [data-testid="stSidebarCollapsedControl"] button:hover,
    [data-testid="collapsedControl"] button:hover {{
        background-color: #2563EB !important;
        transform: scale(1.05) !important;
    }}

    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="collapsedControl"] svg {{
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }}

    [data-testid="stSidebar"] h3 {{
        font-size: 20px !important;
        font-weight: 800 !important;
        color: #60A5FA !important;
        margin-bottom: 14px !important;
    }}

    [data-testid="stSidebar"] h4 {{
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #93C5FD !important;
        margin-top: 14px !important;
        margin-bottom: 8px !important;
    }}

    [data-testid="stSidebar"] label {{
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-size: 17px !important;
        font-weight: 800 !important;
    }}

    /* البانر الرئيسي الفخم (Hero Banner) */
    .hero-banner-new {{
        background: linear-gradient(135deg, #0B132B 0%, #1C2541 45%, #1E3A8A 100%);
        color: #FFFFFF;
        padding: 26px 32px;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.25);
        margin-bottom: 22px;
        border: 1.5px solid rgba(255, 255, 255, 0.12);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 20px;
    }}

    .hero-title-group h1 {{
        color: #FFFFFF !important;
        font-size: {cur_zoom['h1']} !important;
        font-weight: 900 !important;
        margin: 0 0 8px 0 !important;
        line-height: 1.3 !important;
        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }}

    .hero-title-group p {{
        color: #CBD5E1 !important;
        font-size: {cur_zoom['body']} !important;
        font-weight: 600 !important;
        margin: 0 !important;
        line-height: 1.5 !important;
    }}

    /* شريط التحكم السريع في الخط والتكبير أعلى الصفحة */
    .top-control-card {{
        background: #FFFFFF;
        border: 2px solid #E2E8F0;
        border-radius: 16px;
        padding: 14px 22px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.04);
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 14px;
    }}

    .top-control-label {{
        font-size: {cur_zoom['label']} !important;
        font-weight: 800 !important;
        color: #1E3A8A !important;
        display: flex;
        align-items: center;
        gap: 8px;
    }}

    /* شارات العناوين للأقسام */
    .section-badge-title {{
        display: inline-flex;
        align-items: center;
        gap: 10px;
        background: #EFF6FF;
        color: #1E40AF;
        border: 2px solid #BFDBFE;
        padding: 10px 22px;
        border-radius: 14px;
        font-size: {cur_zoom['label']};
        font-weight: 900;
        margin-bottom: 18px;
        margin-top: 10px;
    }}

    /* بطاقات اختيار الهواتف (Phone Selector Cards) - خلفية كحلية داكنة فخمة بحجم مصغر وأنيق */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        background: #0F172A !important;
        border: 1.5px solid #3B82F6 !important;
        border-radius: 12px !important;
        padding: {cur_zoom['card_pad']} !important;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.25) !important;
        margin-bottom: 12px !important;
        transition: all 0.2s ease-in-out !important;
    }}

    div[data-testid="stVerticalBlockBorderWrapper"]:hover {{
        border-color: #60A5FA !important;
        box-shadow: 0 8px 20px rgba(59, 130, 246, 0.2) !important;
    }}

    .phone-select-header {{
        font-size: {cur_zoom['kpi_title']} !important;
        font-weight: 800 !important;
        color: #FFFFFF !important;
        border-bottom: 1.5px solid rgba(255, 255, 255, 0.15) !important;
        padding-bottom: 6px !important;
        margin-bottom: 8px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
    }}

    .phone-select-header span {{
        color: #FFFFFF !important;
    }}

    /* عناوين القوائم المنسدلة (الماركة والموديل) باللون الأبيض وبحجم مصغر */
    div[data-testid="stSelectbox"] {{
        margin-bottom: 8px !important;
    }}

    div[data-testid="stSelectbox"] label,
    div[data-testid="stSelectbox"] label p,
    div[data-testid="stSelectbox"] label span,
    div[data-testid="stSelectbox"] label div,
    div[data-testid="stMultiSelect"] label,
    div[data-testid="stMultiSelect"] label p,
    div[data-testid="stMultiSelect"] label span {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['label']} !important;
        font-weight: 800 !important;
        color: #FFFFFF !important; /* اللون الأبيض لعناوين القوائم */
        -webkit-text-fill-color: #FFFFFF !important;
        text-shadow: 0 1px 3px rgba(0, 0, 0, 0.8) !important;
        margin-bottom: 3px !important;
        line-height: 1.3 !important;
        display: block !important;
    }}

    /* خلايا وخانات الإدخال والقوائم المنسدلة من الداخل (المكتوب فيها الماركة والموديل) */
    div[data-baseweb="select"] > div {{
        min-height: {cur_zoom['input_h']} !important;
        height: {cur_zoom['input_h']} !important;
        border-radius: 8px !important;
        border: 1.5px solid #60A5FA !important;
        background-color: #FFFFFF !important;
        padding: 2px 10px !important;
        display: flex !important;
        align-items: center !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
    }}

    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div[role="combobox"],
    div[data-baseweb="select"] div[aria-selected],
    div[data-baseweb="select"] input {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['input_font']} !important;
        font-weight: 700 !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        line-height: 1.2 !important;
    }}

    div[data-baseweb="select"] svg {{
        width: 16px !important;
        height: 16px !important;
        fill: #1E3A8A !important;
    }}

    /* خيارات القائمة المنسدلة عند فتحها (Dropdown Popover Options) */
    div[data-baseweb="popover"] {{
        border-radius: 10px !important;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2) !important;
        border: 1.5px solid #3B82F6 !important;
        background-color: #FFFFFF !important;
    }}

    div[data-baseweb="popover"] ul[role="listbox"] {{
        padding: 4px !important;
        max-height: 280px !important;
    }}

    div[data-baseweb="popover"] li[role="option"] {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['input_font']} !important;
        font-weight: 700 !important;
        padding: 6px 12px !important;
        min-height: 32px !important;
        border-radius: 6px !important;
        margin-bottom: 2px !important;
        display: flex !important;
        align-items: center !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
    }}

    div[data-baseweb="popover"] li[role="option"]:hover,
    div[data-baseweb="popover"] li[role="option"][aria-selected="true"] {{
        background-color: #EFF6FF !important;
        color: #1E40AF !important;
        -webkit-text-fill-color: #1E40AF !important;
        font-weight: 800 !important;
    }}

    div[data-baseweb="popover"] li[role="option"] * {{
        font-size: {cur_zoom['input_font']} !important;
        font-weight: 700 !important;
    }}

    /* ==========================================================================
       عناوين التبويبات (Tabs Headers) - لون أبيض ناصع 100% وخلفيات متناسقة
       ========================================================================== */
    div[data-testid="stTabs"] div[data-baseweb="tab-list"],
    div[data-baseweb="tab-list"] {{
        gap: 14px !important;
        border-bottom: 4px solid #2563EB !important;
        padding-bottom: 8px !important;
        background: transparent !important;
        display: flex !important;
        flex-wrap: wrap !important;
        margin-top: 20px !important;
        margin-bottom: 24px !important;
    }}

    /* أزرار التبويبات */
    div[data-testid="stTabs"] button[role="tab"],
    div[data-baseweb="tab-list"] button {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['tab_font']} !important;
        font-weight: 900 !important;
        padding: 14px 26px !important;
        border-radius: 14px 14px 0 0 !important;
        min-height: 60px !important;
        transition: all 0.2s ease-in-out !important;
        border: 2.5px solid #2563EB !important;
        border-bottom: none !important;
        margin-bottom: -4px !important;
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }}

    /* نصوص عناوين التبويبات بالكامل باللون الأبيض */
    div[data-testid="stTabs"] button[role="tab"] *,
    div[data-testid="stTabs"] button[role="tab"] p,
    div[data-testid="stTabs"] button[role="tab"] span,
    div[data-testid="stTabs"] button[role="tab"] div,
    div[data-baseweb="tab-list"] button *,
    div[data-baseweb="tab-list"] button p,
    div[data-baseweb="tab-list"] button span,
    div[data-baseweb="tab-list"] button div {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['tab_font']} !important;
        font-weight: 900 !important;
        color: #FFFFFF !important; /* أبيض ناصع */
        -webkit-text-fill-color: #FFFFFF !important;
        text-shadow: 0 2px 4px rgba(0,0,0,0.8) !important;
        line-height: 1.4 !important;
        margin: 0 !important;
    }}

    /* التبويب النشط المختار (Active Tab) */
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"],
    div[data-baseweb="tab-list"] button[aria-selected="true"] {{
        background: linear-gradient(135deg, #1D4ED8 0%, #2563EB 50%, #3B82F6 100%) !important;
        border: 3px solid #93C5FD !important;
        border-bottom: none !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.6) !important;
        transform: translateY(-3px) !important;
    }}

    /* التبويب غير النشط (Inactive Tab) */
    div[data-testid="stTabs"] button[role="tab"][aria-selected="false"],
    div[data-baseweb="tab-list"] button[aria-selected="false"] {{
        background-color: #0F172A !important;
        border: 2.5px solid #3B82F6 !important;
        border-bottom: none !important;
    }}

    /* تأثير الـ Hover على التبويب */
    div[data-testid="stTabs"] button[role="tab"][aria-selected="false"]:hover,
    div[data-baseweb="tab-list"] button[aria-selected="false"]:hover {{
        background-color: #1E3A8A !important;
        border-color: #93C5FD !important;
    }}

    /* حماية محتوى ومخرجات التبويبات من التأثر بلون الأبيض */
    div[data-testid="stTabContent"],
    div[data-baseweb="tab-panel"],
    div[role="tabpanel"] {{
        color: #0F172A !important;
    }}

    /* ==========================================================================
       تنسيقات تبويب التحليل البياني (Visual Charts Tab Styling) - نصوص بيضاء بالكامل
       ========================================================================== */
    .analytics-tab-container {{
        background: linear-gradient(135deg, #0B132B 0%, #0F172A 100%) !important;
        border: 2px solid #1E3A8A !important;
        border-radius: 18px !important;
        padding: 22px !important;
        margin-top: 10px !important;
        margin-bottom: 24px !important;
        box-shadow: 0 10px 30px rgba(11, 19, 43, 0.45) !important;
        direction: rtl !important;
        text-align: right !important;
    }}

    .analytics-header-banner {{
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%) !important;
        border: 1.5px solid #3B82F6 !important;
        border-radius: 14px !important;
        padding: 16px 20px !important;
        margin-bottom: 20px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
    }}

    .analytics-header-title {{
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-size: 19px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        line-height: 1.3 !important;
    }}

    .analytics-header-desc {{
        color: #CBD5E1 !important;
        -webkit-text-fill-color: #CBD5E1 !important;
        font-size: 13.5px !important;
        font-weight: 600 !important;
        margin: 4px 0 0 0 !important;
        line-height: 1.5 !important;
    }}

    .analytics-header-badge {{
        background: rgba(59, 130, 246, 0.25) !important;
        color: #93C5FD !important;
        -webkit-text-fill-color: #93C5FD !important;
        border: 1.5px solid #3B82F6 !important;
        padding: 6px 14px !important;
        border-radius: 10px !important;
        font-size: 13px !important;
        font-weight: 800 !important;
    }}

    .analytics-kpi-card {{
        background: #0F172A !important;
        border: 1.5px solid #334155 !important;
        border-radius: 14px !important;
        padding: 14px 12px !important;
        text-align: center !important;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.25) !important;
        transition: all 0.2s ease-in-out !important;
        margin-bottom: 16px !important;
    }}

    .analytics-kpi-card:hover {{
        border-color: #60A5FA !important;
        transform: translateY(-2px) !important;
    }}

    .analytics-kpi-title {{
        color: #93C5FD !important;
        -webkit-text-fill-color: #93C5FD !important;
        font-size: 12.5px !important;
        font-weight: 800 !important;
        margin-bottom: 4px !important;
    }}

    .analytics-kpi-val {{
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-size: 17px !important;
        font-weight: 900 !important;
        margin-bottom: 2px !important;
    }}

    .analytics-kpi-phone {{
        color: #E2E8F0 !important;
        -webkit-text-fill-color: #E2E8F0 !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }}

    .analytics-chart-box {{
        background: #0F172A !important;
        border: 1.5px solid #1E293B !important;
        border-radius: 16px !important;
        padding: 18px 20px !important;
        margin-bottom: 20px !important;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.25) !important;
    }}

    .analytics-chart-box:hover {{
        border-color: #3B82F6 !important;
    }}

    .analytics-chart-header {{
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        border-bottom: 1.5px solid rgba(255, 255, 255, 0.1) !important;
        padding-bottom: 10px !important;
        margin-bottom: 14px !important;
        flex-wrap: wrap !important;
        gap: 8px !important;
    }}

    .analytics-chart-title {{
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-size: 16px !important;
        font-weight: 900 !important;
        margin: 0 !important;
        display: flex !important;
        align-items: center !important;
        gap: 8px !important;
    }}

    .analytics-chart-badge {{
        background: rgba(59, 130, 246, 0.2) !important;
        color: #93C5FD !important;
        -webkit-text-fill-color: #93C5FD !important;
        border: 1px solid #3B82F6 !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        padding: 2px 10px !important;
        border-radius: 10px !important;
    }}

    .analytics-table-matrix {{
        width: 100% !important;
        border-collapse: collapse !important;
        direction: rtl !important;
        text-align: right !important;
        margin-top: 10px !important;
        border-radius: 12px !important;
        overflow: hidden !important;
    }}

    .analytics-table-matrix th {{
        background: #1E293B !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 900 !important;
        font-size: 13.5px !important;
        padding: 10px 14px !important;
        border: 1px solid #334155 !important;
        text-align: center !important;
    }}

    .analytics-table-matrix td {{
        background: #0F172A !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        padding: 10px 14px !important;
        border: 1px solid #334155 !important;
        text-align: center !important;
    }}

    .analytics-table-matrix tr:nth-child(even) td {{
        background: #162032 !important;
    }}

    /* فرض اللون الأبيض الصريح على جميع النصوص داخل تبويب التحليل البياني */
    .analytics-tab-container,
    .analytics-tab-container p,
    .analytics-tab-container span,
    .analytics-tab-container div,
    .analytics-tab-container h1,
    .analytics-tab-container h2,
    .analytics-tab-container h3,
    .analytics-tab-container h4,
    .analytics-tab-container h5,
    .analytics-tab-container h6,
    .analytics-tab-container label,
    .analytics-chart-box,
    .analytics-chart-box p,
    .analytics-chart-box span,
    .analytics-chart-box div,
    .analytics-kpi-card,
    .analytics-kpi-card p,
    .analytics-kpi-card span,
    .analytics-kpi-card div {{
        font-family: var(--app-font-family) !important;
    }}

    /* ضبط اتجاه الرسوم البيانية Vega-Lite إلى LTR لمنع تداخل النصوص مع الأشرطة بسبب اتجاه الصفحة RTL */
    div[data-testid="stVegaLiteChart"],
    div[data-testid="stArrowVegaLiteChart"],
    .vega-embed,
    .vega-embed svg,
    .vega-embed canvas,
    .vega-bind,
    svg.marks {{
        direction: ltr !important;
        text-align: left !important;
        unicode-bidi: isolate !important;
    }}

    /* زر التحليل الأساسي الكبير */
    button[kind="primary"] {{
        font-family: var(--app-font-family) !important;
        font-size: {cur_zoom['label']} !important;
        font-weight: 900 !important;
        padding: 16px 32px !important;
        min-height: 56px !important;
        background: linear-gradient(135deg, #1E40AF 0%, #3B82F6 100%) !important;
        border: none !important;
        border-radius: 14px !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }}

    button[kind="primary"]:hover {{
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 25px rgba(37, 99, 235, 0.4) !important;
    }}

    /* بطاقة بطل المقارنة الإجمالي (Overall Champion Banner) */
    .champion-card {{
        background: linear-gradient(135deg, #064E3B 0%, #047857 50%, #059669 100%);
        color: #FFFFFF;
        padding: 22px 28px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(4, 120, 87, 0.25);
        margin-bottom: 24px;
        border: 2px solid #34D399;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 16px;
    }}

    .champion-title {{
        font-size: {cur_zoom['h2']};
        font-weight: 900;
        color: #FFFFFF;
        margin-bottom: 6px;
    }}

    .champion-desc {{
        font-size: {cur_zoom['body']};
        font-weight: 600;
        color: #ECFDF5;
        margin: 0;
    }}

    /* كروت المؤشرات السريعة (KPI Metric Cards) */
    .kpi-metric-box {{
        background: #FFFFFF;
        border: 2px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 16px;
        text-align: center;
        box-shadow: 0 4px 14px rgba(0,0,0,0.03);
        margin-bottom: 16px;
    }}

    .kpi-metric-num {{
        font-size: {cur_zoom['kpi_num']} !important;
        font-weight: 900 !important;
        color: #1E3A8A;
        line-height: 1.2 !important;
        margin-bottom: 4px;
    }}

    .kpi-metric-label {{
        font-size: calc({cur_zoom['body']} * 0.9) !important;
        color: #64748B;
        font-weight: 700;
    }}

    /* بطاقات استعراض الهواتف (Phone Quick Showcase Cards) */
    .phone-showcase-box {{
        background: #FFFFFF;
        border: 2px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.04);
        margin-bottom: 18px;
        text-align: center;
    }}

    .phone-showcase-title {{
        font-size: {cur_zoom['kpi_title']} !important;
        font-weight: 900 !important;
        color: #0F172A;
        margin-bottom: 6px;
    }}

    .phone-showcase-price {{
        font-size: calc({cur_zoom['body']} * 1.05) !important;
        font-weight: 900 !important;
        color: #047857;
        background: #ECFDF5;
        border: 1.5px solid #A7F3D0;
        padding: 4px 14px;
        border-radius: 8px;
        display: inline-block;
        margin-bottom: 10px;
    }}

    /* جدول المقارنة الفائق الوضوح (Desktop Comparison Table) */
    .table-container-modern {{
        background-color: #FFFFFF !important;
        border: 2.5px solid #94A3B8 !important;
        border-radius: 18px !important;
        box-shadow: 0 8px 26px rgba(0,0,0,0.06) !important;
        overflow-x: auto;
        margin: 20px 0;
    }}

    #comparison-table-main {{
        width: 100%;
        border-collapse: collapse;
        direction: rtl;
        text-align: right;
    }}

    #comparison-table-main th {{
        font-size: {cur_zoom['tbl_th']} !important;
        padding: {cur_zoom['tbl_pad']} !important;
        line-height: 1.5 !important;
        font-weight: 900 !important;
    }}

    #comparison-table-main td {{
        font-size: {cur_zoom['tbl_td']} !important;
        padding: {cur_zoom['tbl_pad']} !important;
        line-height: 1.7 !important;
    }}

    /* شارة "⭐ الأفضل" داخل الخلية الفائزة */
    .winner-badge {{
        background: #15803D !important;
        color: #FFFFFF !important;
        font-size: calc({cur_zoom['tbl_td']} * 0.75) !important;
        font-weight: 900 !important;
        padding: 4px 12px !important;
        border-radius: 6px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 4px !important;
        margin-top: 6px !important;
        box-shadow: 0 2px 6px rgba(21, 128, 61, 0.3) !important;
    }}

    /* إخفاء زر Download as CSV الافتراضي من شريط أدوات الجداول */
    button[aria-label="Download as CSV"],
    button[title="Download as CSV"],
    button[aria-label*="Download"],
    button[title*="Download"],
    div[data-testid="stElementToolbar"] button[aria-label*="Download"],
    div[data-testid="stElementToolbar"] button[title*="Download"],
    div[data-testid="stElementToolbar"] [data-testid="stIconDownload"],
    div[data-testid="stElementToolbarButton"]:has(svg[data-testid="stIconDownload"]) {{
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        width: 0 !important;
        height: 0 !important;
    }}
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 5. لوحات الألوان وتدرجات الهواتف
# -----------------------------------------------------------------------------
PHONE_COLOR_GRADIENTS = [
    {"gradient": "linear-gradient(135deg, #1E40AF 0%, #3B82F6 100%)", "border": "#1E3A8A", "light_bg": "#EFF6FF", "text_c": "#1E40AF"},
    {"gradient": "linear-gradient(135deg, #6B21A8 0%, #A855F7 100%)", "border": "#581C87", "light_bg": "#FAF5FF", "text_c": "#6B21A8"},
    {"gradient": "linear-gradient(135deg, #C2410C 0%, #FB923C 100%)", "border": "#9A3412", "light_bg": "#FFF7ED", "text_c": "#C2410C"},
    {"gradient": "linear-gradient(135deg, #0F766E 0%, #14B8A6 100%)", "border": "#115E59", "light_bg": "#F0FDFA", "text_c": "#0F766E"},
    {"gradient": "linear-gradient(135deg, #BE123C 0%, #FB7185 100%)", "border": "#9F1239", "light_bg": "#FFF1F2", "text_c": "#BE123C"},
    {"gradient": "linear-gradient(135deg, #4338CA 0%, #818CF8 100%)", "border": "#3730A3", "light_bg": "#EEF2FF", "text_c": "#4338CA"},
    {"gradient": "linear-gradient(135deg, #0369A1 0%, #38BDF8 100%)", "border": "#075985", "light_bg": "#F0F9FF", "text_c": "#0369A1"},
    {"gradient": "linear-gradient(135deg, #3F6212 0%, #84CC16 100%)", "border": "#365314", "light_bg": "#F7FEE7", "text_c": "#3F6212"},
]

CATEGORY_STYLES = {
    "General": {"name": "معلومات عامة والتعريف", "bg": "#F8FAFC", "color": "#0F172A", "icon": "ℹ️"},
    "Screen": {"name": "الشاشة وتكنولوجيا العرض", "bg": "#EFF6FF", "color": "#1E40AF", "icon": "📺"},
    "Performance": {"name": "الأداء والعتاد والمعالج", "bg": "#FAF5FF", "color": "#6B21A8", "icon": "⚡"},
    "Camera": {"name": "منظومة الكاميرات والتصوير", "bg": "#FFF1F2", "color": "#BE123C", "icon": "📸"},
    "Battery": {"name": "البطارية وسرعة الشحن", "bg": "#F0FDF4", "color": "#166534", "icon": "🔋"},
    "Software": {"name": "نظام التشغيل والتحديثات", "bg": "#F0F9FF", "color": "#0369A1", "icon": "💻"},
    "Design": {"name": "التصميم والوزن والأبعاد", "bg": "#FFFBEB", "color": "#B45309", "icon": "⚖️"},
    "Build": {"name": "المتانة ومقاومة الماء", "bg": "#F8FAFC", "color": "#334155", "icon": "🛡️"},
    "Connectivity": {"name": "الاتصال وشبكات 5G", "bg": "#F5F3FF", "color": "#5B21B6", "icon": "📶"},
    "Price": {"name": "الأسعار بالجنيه المصري (1$ = 50 ج)", "bg": "#ECFDF5", "color": "#047857", "icon": "💰"},
    "Extra": {"name": "الميزات البارزة الإضافية", "bg": "#FEFCE8", "color": "#A16207", "icon": "✨"}
}

# -----------------------------------------------------------------------------
# 5.5 شريط أدوات ومستودع GitHub أعلى الشاشة (Top GitHub Toolbar)
# -----------------------------------------------------------------------------
top_git_col1, top_git_col2, top_git_col3, top_git_col4 = st.columns([0.34, 0.22, 0.22, 0.22], gap="small")

git_info_quick = get_git_info()
git_details = get_git_details()
git_summary = get_git_status_summary()
target_github_url = git_details.get("web_url") or DEFAULT_PROJECT_WEB_URL
target_remote_url = git_details.get("remote") or DEFAULT_PROJECT_GITHUB_URL
current_branch = git_details.get("branch") or "main"

with top_git_col1:
    git_status_icon = "🟢" if git_info_quick.get("is_repo") else "⚪"
    changes_badge = f"<span style='font-size: 11px; color: #FBBF24; background: rgba(245, 158, 11, 0.2); border: 1px solid #F59E0B; padding: 2px 7px; border-radius: 6px; font-weight: 700;'>📝 {git_summary['total_changes']} تعديل</span>" if git_summary.get("total_changes", 0) > 0 else "<span style='font-size: 11px; color: #34D399; background: rgba(16, 185, 129, 0.2); border: 1px solid #10B981; padding: 2px 7px; border-radius: 6px; font-weight: 700;'>✨ متزامن</span>"
    
    st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; padding: 7px 12px; background: #0F172A; border: 1.5px solid #3B82F6; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);">
        <div style="display: flex; align-items: center; gap: 8px;">
            <span style="font-size: 20px;">🐙</span>
            <span style="font-weight: 800; font-size: 13px; color: #FFFFFF;">Mobiles-spec</span>
            <span style="font-size: 11px; color: #93C5FD; background: rgba(59, 130, 246, 0.25); border: 1px solid #3B82F6; padding: 2px 7px; border-radius: 6px; font-weight: 700;">
                {git_status_icon} {current_branch}
            </span>
        </div>
        <div>
            {changes_badge}
        </div>
    </div>
    """, unsafe_allow_html=True)

with top_git_col2:
    if target_github_url and target_github_url != "-":
        st.link_button("👁️ View GitHub", url=target_github_url, type="secondary", use_container_width=True)
    else:
        with st.popover("👁️ View GitHub", use_container_width=True):
            st.markdown("### 👁️ ربط واستعراض مستودع GitHub")
            st.info(f"🔗 المستودع الموصى به: `{DEFAULT_PROJECT_GITHUB_URL}`")
            quick_repo_url = st.text_input(
                "🔗 رابط المستودع الخاص بك على GitHub:",
                value=DEFAULT_PROJECT_GITHUB_URL,
                key="quick_set_repo_url"
            )
            if st.button("💾 حفظ وربط المستودع الآن", key="btn_save_quick_repo", type="primary", use_container_width=True):
                clean_in = quick_repo_url.strip() if quick_repo_url.strip() else DEFAULT_PROJECT_GITHUB_URL
                s_set, msg_set = set_git_remote(clean_in)
                if s_set:
                    st.success("✅ تم حفظ رابط المستودع بنجاح! سيفتح الزر مستودعك مباشرة.")
                    st.rerun()
                else:
                    st.error(f"خطأ: {msg_set}")

with top_git_col3:
    with st.popover("🚀 Export to GitHub", use_container_width=True):
        st.markdown("### 🚀 تصدير وتهيئة المشروع على GitHub")
        st.markdown(f"""
        <div style="background: rgba(59, 130, 246, 0.1); border: 1px solid #3B82F6; padding: 10px 14px; border-radius: 10px; margin-bottom: 12px; font-size: 13px;">
            <b>🔗 المستودع المستهدف:</b> <a href="{DEFAULT_PROJECT_WEB_URL}" target="_blank" style="color: #60A5FA; text-decoration: none; font-weight: 700;">hzayed3030-cell/Mobiles-spec</a><br>
            <b>🌿 الفرع المستهدف:</b> <code style="color: #93C5FD;">main</code>
        </div>
        """, unsafe_allow_html=True)
        
        # فحص فهارس build / dist
        detected_build_dirs = check_large_build_dirs()
        if detected_build_dirs and not st.session_state.get("dismiss_build_warn_export", False):
            tot_mb = sum(d["size_mb"] for d in detected_build_dirs)
            dir_names = " و ".join([f"[{d['name']}]" for d in detected_build_dirs])
            
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #7F1D1D 0%, #991B1B 100%); color: #FFFFFF; padding: 14px 18px; border-radius: 12px; border: 2px solid #EF4444; margin-bottom: 14px;">
                <div style="display: flex; align-items: center; gap: 8px; font-size: 15px; font-weight: 800; color: #FEE2E2;">
                    <span>🚨</span>
                    <span>تنبيه أمان: تم اكتشاف مجلدات ملفات تنفيذية ضخمة ({tot_mb} MB)</span>
                </div>
                <div style="margin-top: 6px; font-size: 12.5px; color: #FECACA;">
                    المجلدات: <b>{dir_names}</b>. يفضل حذفها قبل الرفع لمنع تجاوز سعة مستودع GitHub.
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            col_b_act1, col_b_act2 = st.columns([0.62, 0.38], gap="small")
            with col_b_act1:
                if st.button("🗑️ حذف الفهارس للمتابعة", key="btn_del_build_dirs", type="primary", use_container_width=True):
                    s_del, msg_del = delete_build_dirs()
                    if s_del:
                        st.success("✅ تم حذف فهارس Build و Dist بنجاح!")
                        st.session_state["dismiss_build_warn_export"] = False
                        st.rerun()
                    else:
                        st.error(f"خطأ: {msg_del}")
            with col_b_act2:
                if st.button("🔙 تجاهل والمتابعة", key="btn_cancel_del_build_dirs", type="secondary", use_container_width=True):
                    st.session_state["dismiss_build_warn_export"] = True
                    st.rerun()
            st.divider()
        
        export_url = st.text_input(
            "🔗 رابط المستودع على GitHub (Remote URL):",
            value=target_remote_url if target_remote_url != "-" else DEFAULT_PROJECT_GITHUB_URL,
            key="export_github_url_key"
        )
        export_branch = st.text_input(
            "🌿 اسم الفرع (Branch):",
            value=current_branch if current_branch != "-" else "main",
            key="export_github_branch_key"
        )
        export_commit = st.text_input(
            "📝 رسالة الـ Commit (Commit Message):",
            value="feat: initial export of Mobiles-spec codebase",
            key="export_github_commit_key"
        )
        
        if st.button("🚀 بدء التصدير والرفع الآن (Push to GitHub)", key="btn_exec_export_github", type="primary", use_container_width=True):
            clean_exp = export_url.strip() if export_url.strip() else DEFAULT_PROJECT_GITHUB_URL
            with st.spinner("جاري تهيئة ورفع المشروع إلى GitHub..."):
                success, log_msg = export_project_to_github(
                    repo_url=clean_exp,
                    commit_message=export_commit,
                    branch_name=export_branch
                )
                if success:
                    st.success("✅ تم تصدير ورفع المشروع إلى GitHub بنجاح!")
                    clean_web = clean_exp.replace(".git", "").replace("git@github.com:", "https://github.com/").strip()
                    st.link_button("🌐 فتح المستودع الآن على GitHub", url=clean_web, type="primary", use_container_width=True)
                else:
                    st.warning("⚠️ اكتملت العملية مع التقرير التالي:")
                st.code(log_msg, language="bash")

with top_git_col4:
    with st.popover("🔄 Update GitHub", use_container_width=True):
        st.markdown("### 🔄 تحديث ومزامنة مستودع GitHub")
        st.markdown(f"""
        <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid #10B981; padding: 10px 14px; border-radius: 10px; margin-bottom: 12px; font-size: 13px;">
            <b>🔗 المستودع المرتبط:</b> <a href="{DEFAULT_PROJECT_WEB_URL}" target="_blank" style="color: #34D399; text-decoration: none; font-weight: 700;">hzayed3030-cell/Mobiles-spec</a><br>
            <b>🌿 الفرع الحالي:</b> <code style="color: #A7F3D0;">{current_branch}</code> | <b>آخر Commit:</b> <small>{git_info_quick.get('last_commit', '-')}</small>
        </div>
        """, unsafe_allow_html=True)
        
        # عرض حالة التعديلات الحالية
        if git_summary.get("total_changes", 0) > 0:
            st.info(f"📝 يوجد **{git_summary['total_changes']}** ملفات بها تعديلات جاهزة للمزامنة:")
            with st.expander("📂 استعراض قائمة الملفات المعدلة", expanded=False):
                for f_item in git_summary.get("files", [])[:15]:
                    st.caption(f"{f_item['icon']} `{f_item['name']}` ({f_item['type']})")
        else:
            st.success("🟢 المستودع نظيف ومحدث محلياً، يمكنك عمل مزامنة ورفع فوري.")
            
        # فحص وجود فهارس build / dist
        detected_build_update = check_large_build_dirs()
        if detected_build_update and not st.session_state.get("dismiss_build_warn_update", False):
            st.error("🚨 تنبيه: توجد فهارس Build / Dist خاصة بـ EXE في المشروع. يرجى حذفها أولاً لتتمكن من التحديث.")
            col_u_act1, col_u_act2 = st.columns([0.62, 0.38], gap="small")
            with col_u_act1:
                if st.button("🗑️ حذف فهارس Build و Dist الآن", key="btn_del_build_update", type="primary", use_container_width=True):
                    s_del, msg_del = delete_build_dirs()
                    if s_del:
                        st.success("✅ تم الحذف بنجاح!")
                        st.session_state["dismiss_build_warn_update"] = False
                        st.rerun()
                    else:
                        st.error(f"خطأ: {msg_del}")
            with col_u_act2:
                if st.button("🔙 إلغاء والعودة", key="btn_cancel_del_update", type="secondary", use_container_width=True):
                    st.session_state["dismiss_build_warn_update"] = True
                    st.rerun()
            st.divider()
        elif detected_build_update and st.session_state.get("dismiss_build_warn_update", False):
            st.info("ℹ️ تم إلغاء التحديث السابق. فهارس Build / Dist موجودة كما هي دون تعديل.")
            if st.button("🔄 إعادة فحص الفهارس للتحديث", key="btn_reset_build_warn_update", use_container_width=True):
                st.session_state["dismiss_build_warn_update"] = False
                st.rerun()
            st.divider()
        
        update_branch = st.text_input(
            "🌿 اسم الفرع (Branch):",
            value=current_branch if current_branch != "-" else "main",
            key="update_github_branch_key"
        )
        update_commit = st.text_input(
            "📝 رسالة الـ Commit (Commit Message):",
            value="feat: update Mobiles-spec data & specs",
            key="update_github_commit_key"
        )
        
        if st.button("🔄 حفظ ورفع التحديثات (Update & Refresh)", key="btn_exec_update_github", type="primary", use_container_width=True):
            clean_commit_msg = update_commit.strip() if update_commit.strip() else "feat: update Mobiles-spec data & specs"
            with st.spinner("جاري حفظ التعديلات والمزامنة والرفع إلى GitHub..."):
                success, log_msg = update_project_on_github(
                    commit_message=clean_commit_msg,
                    branch_name=update_branch,
                    repo_url=target_remote_url if target_remote_url != "-" else DEFAULT_PROJECT_GITHUB_URL
                )
                if success:
                    st.success("✅ تم تحديث ومزامنة المستودع بنجاح!")
                    commit_history_url = f"{DEFAULT_PROJECT_WEB_URL}/commits/{update_branch}"
                    st.link_button("📜 عرض سجل التحديثات على GitHub", url=commit_history_url, type="primary", use_container_width=True)
                else:
                    st.warning("⚠️ اكتملت العملية مع التقرير التالي:")
                st.code(log_msg, language="bash")

st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 6. البانر التعريفي
# -----------------------------------------------------------------------------
curr_num_mob = st.session_state.get("num_mobiles_input", 3)
st.html(f"""
<div class="hero-banner-new">
    <div class="hero-title-group">
        <h1>📱 لوحة مقارنة مواصفات الهواتف الذكية (Mobile Specs Dashboard)</h1>
        <p style="color: #FEF08A; font-weight: 700; font-size: 14.5px;">مقارنة ذكية متطورة: اختيار متدرج للماركة والموديل، تمييز الخلية الفائزة بالأخضر مع شارة [⭐ الأفضل]، وعمود أسباب الأفضلية.</p>
    </div>
    <div>
        <span style="background: rgba(255,255,255,0.18); border: 1.5px solid rgba(255,255,255,0.3); padding: 8px 18px; border-radius: 12px; font-weight: 800; font-size: 16px;">
            📊 مقارنة {curr_num_mob} هواتف
        </span>
    </div>
</div>
""")


# -----------------------------------------------------------------------------
# 6.5 نافذة العرض المباشر على كامل الشاشة لدليل الإعدادات (Full Screen Dialog)
# -----------------------------------------------------------------------------
@st.dialog("📱 استعراض دليل الإعدادات والمسارات الشامل", width="large")
def show_settings_guide_dialog(full_phone_name: str, selected_brand: str, selected_model: str):
    # حقن كود CSS لزيادة عرض النافذة وتوحيد مظهر وخلفية جميع الأزرار مع زر Excel
    st.markdown("""
        <style>
        div[data-testid="stDialog"] > div {
            width: 96vw !important;
            max-width: 98vw !important;
            min-height: 85vh !important;
            border-radius: 16px !important;
            padding: 20px 28px !important;
            background: #FFFFFF !important;
        }
        div[data-testid="stDialog"] [data-testid="stVerticalBlock"] {
            gap: 10px !important;
        }
        /* إخفاء زر Download as CSV الافتراضي من شريط أدوات الجدول */
        button[aria-label="Download as CSV"],
        button[title="Download as CSV"],
        div[data-testid="stElementToolbar"] button[aria-label*="Download"],
        div[data-testid="stElementToolbar"] button[title*="Download"],
        div[data-testid="stElementToolbar"] [data-testid="stIconDownload"],
        div[data-testid="stElementToolbarButton"]:has(svg[data-testid="stIconDownload"]) {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
            width: 0 !important;
            height: 0 !important;
        }
        /* توحيد ألوان وخلفيات جميع أزرار النافذة مع زر تحميل Excel */
        div[data-testid="stDialog"] div[data-testid="stPopover"] > button,
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button,
        div[data-testid="stDialog"] button {
            background: linear-gradient(135deg, #1E40AF 0%, #3B82F6 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1.5px solid #60A5FA !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
            font-size: 13.5px !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
            transition: all 0.2s ease-in-out !important;
            min-height: 42px !important;
            height: 42px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            padding: 0 14px !important;
        }
        div[data-testid="stDialog"] div[data-testid="stPopover"] > button:hover,
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button:hover,
        div[data-testid="stDialog"] button:hover {
            background: linear-gradient(135deg, #1D4ED8 0%, #2563EB 100%) !important;
            border-color: #93C5FD !important;
            box-shadow: 0 6px 18px rgba(37, 99, 235, 0.4) !important;
            transform: translateY(-1.5px) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        div[data-testid="stDialog"] div[data-testid="stPopover"] > button p,
        div[data-testid="stDialog"] div[data-testid="stPopover"] > button span,
        div[data-testid="stDialog"] div[data-testid="stPopover"] > button div,
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button p,
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button span,
        div[data-testid="stDialog"] button p,
        div[data-testid="stDialog"] button span {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 800 !important;
            font-size: 13.5px !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # تجهيز ملف الإكسيل المنسق بالكامل مع الأبعاد والتنسيقات الملكية
    guide_excel_bytes = export_phone_settings_to_excel(
        phone_name=full_phone_name,
        brand=selected_brand,
        model=selected_model
    )
    safe_model_filename = f"{full_phone_name.replace(' ', '_').replace('/', '_')}_Settings_Guide.xlsx"

    # استخراج DataFrame الكامل
    full_settings_df = get_phone_settings_dataframe(
        phone_name=full_phone_name,
        brand=selected_brand,
        model=selected_model
    )
    all_columns = list(full_settings_df.columns)

    # شريط رأس النافذة مع أزرار التحكم: البحث، إظهار/إخفاء الأعمدة، وتحميل Excel، والإغلاق
    head_c1, head_c2, head_c3, head_c4, head_c5 = st.columns([0.36, 0.16, 0.20, 0.18, 0.10], gap="small")
    with head_c1:
        st.markdown(f"""
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="font-size: 26px;">📱</span>
            <div>
                <h2 style="margin: 0; font-size: 20px; font-weight: 800; color: #1E293B;">دليل إعدادات: {full_phone_name}</h2>
                <div style="font-size: 12px; color: #64748B;">مسارات الوصول والوظائف باللغتين العربية والإنجليزية</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with head_c2:
        with st.popover("🔍 بحث (Search)", use_container_width=True):
            st.markdown("**🔍 البحث الفوري في القوائم والمسارات:**")
            search_term = st.text_input(
                "اكتب كلمة البحث:",
                placeholder="مثال: Wi-Fi, battery, كاميرا, شحن, 5G...",
                key=f"search_input_modal_{full_phone_name}",
                label_visibility="collapsed"
            )

    with head_c3:
        with st.popover("👁️ إظهار/إخفاء (Show/Hide)", use_container_width=True):
            st.markdown("**👁️ تخصيص الأعمدة المعروضة:**")
            selected_columns = st.multiselect(
                "اختر الأعمدة:",
                options=all_columns,
                default=all_columns,
                key=f"cols_filter_modal_{full_phone_name}",
                label_visibility="collapsed"
            )

    with head_c4:
        st.download_button(
            label="📗 تحميل Excel",
            data=guide_excel_bytes,
            file_name=safe_model_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"modal_dl_top_{full_phone_name}",
            type="primary",
            use_container_width=True
        )

    with head_c5:
        if st.button("✖ إغلاق", key=f"top_close_dlg_{full_phone_name}", type="primary", use_container_width=True):
            st.rerun()

    st.divider()

    # تطبيق فلترة الأعمدة وفلترة البحث
    if not selected_columns:
        selected_columns = all_columns
    filtered_df = full_settings_df[selected_columns]

    if 'search_term' in locals() and search_term and search_term.strip():
        q = search_term.strip().lower()
        mask = filtered_df.astype(str).apply(lambda col: col.str.lower().str.contains(q, na=False)).any(axis=1)
        filtered_df = filtered_df[mask]

    col_config = {
        "Menu #": st.column_config.NumberColumn(label="Menu #", width="small"),
        "Main Menu (القائمة الرئيسية)": st.column_config.TextColumn(label="Main Menu (القائمة الرئيسية)", width="medium"),
        "Full Settings Path (مسار الوصول)": st.column_config.TextColumn(label="Full Settings Path (مسار الوصول الكامل)", width="large"),
        "Description in English (الوصف بالإنجليزية)": st.column_config.TextColumn(label="Description in English (الوصف بالإنجليزية)", width="large"),
        "Description in Arabic (شرح الوظيفة بالعربية)": st.column_config.TextColumn(label="Description in Arabic (شرح الوظيفة بالعربية)", width="large"),
    }

    # جدول بيانات تفاعلي بكامل العرض والارتفاع
    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        height=560,
        column_config={k: v for k, v in col_config.items() if k in selected_columns}
    )

    foot_c1, foot_c2, foot_c3 = st.columns([0.48, 0.32, 0.20], gap="small")
    with foot_c2:
        st.download_button(
            label="📗 تحميل شيت الإكسيل الملكي المنسق (Excel)",
            data=guide_excel_bytes,
            file_name=safe_model_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"modal_dl_bot_{full_phone_name}",
            type="primary",
            use_container_width=True
        )
    with foot_c3:
        if st.button("✖ إغلاق الشاشة (Close)", key=f"bot_close_dlg_{full_phone_name}", type="primary", use_container_width=True):
            st.rerun()


# -----------------------------------------------------------------------------
# 6.6 نافذة عرض وطباعة نتائج المقارنة الشاملة (Comparison Print Dialog)
# -----------------------------------------------------------------------------
@st.dialog("📑 عرض وطباعة نتائج المقارنة الشاملة", width="large")
def show_comparison_print_dialog(df, active_phones, all_specs, winners_list, reasons_list, final_verdict):
    st.markdown("""
        <style>
        div[data-testid="stDialog"] > div {
            width: 96vw !important;
            max-width: 98vw !important;
            min-height: 85vh !important;
            border-radius: 16px !important;
            padding: 20px 28px !important;
            background: #FFFFFF !important;
        }
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button,
        div[data-testid="stDialog"] button {
            background: linear-gradient(135deg, #1E40AF 0%, #3B82F6 100%) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1.5px solid #60A5FA !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
            font-size: 13.5px !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
            min-height: 42px !important;
            height: 42px !important;
        }
        div[data-testid="stDialog"] div[data-testid="stDownloadButton"] > button:hover,
        div[data-testid="stDialog"] button:hover {
            background: linear-gradient(135deg, #1D4ED8 0%, #2563EB 100%) !important;
            border-color: #93C5FD !important;
            transform: translateY(-1.5px) !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # تجهيز ملفات التصدير
    excel_data = export_comparison_to_excel(
        df,
        title="مقارنة مواصفات الهواتف الذكية",
        winners_per_row=winners_list,
        final_verdict=final_verdict
    )
    pdf_data = export_comparison_to_pdf(
        df,
        title="مقارنة مواصفات الهواتف الذكية",
        winners_per_row=winners_list,
        final_verdict=final_verdict
    )
    html_print = generate_html_print_view(
        df,
        title="تقرير مقارنة مواصفات الهواتف الذكية",
        winners_per_row=winners_list,
        final_verdict=final_verdict
    )

    # شريط رأس النافذة مع أزرار التحميل والإغلاق
    h_c1, h_c2, h_c3, h_c4 = st.columns([0.46, 0.20, 0.20, 0.14], gap="small")
    with h_c1:
        st.markdown("""
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="font-size: 26px;">📑</span>
            <div>
                <h2 style="margin: 0; font-size: 20px; font-weight: 800; color: #1E293B;">عرض وطباعة نتائج المقارنة</h2>
                <div style="font-size: 12px; color: #64748B;">معاينة كاملة جاهزة للطباعة المباشرة والتصدير بجودة فائقة</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with h_c2:
        st.download_button(
            label="📗 تحميل Excel (.xlsx)",
            data=excel_data,
            file_name="Mobile_Specs_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="modal_comp_dl_excel",
            type="primary",
            use_container_width=True
        )
    with h_c3:
        st.download_button(
            label="📕 تحميل PDF (.pdf)",
            data=pdf_data,
            file_name="Mobile_Specs_Comparison.pdf",
            mime="application/pdf",
            key="modal_comp_dl_pdf",
            type="primary",
            use_container_width=True
        )
    with h_c4:
        if st.button("✖ إغلاق", key="modal_comp_close_top", type="primary", use_container_width=True):
            st.rerun()

    st.divider()

    # عرض شاشة المعاينة للطباعة التفاعلية
    components.html(html_print, height=620, scrolling=True)

    f_c1, f_c2 = st.columns([0.8, 0.2])
    with f_c2:
        if st.button("✖ إغلاق الشاشة (Close)", key="modal_comp_close_bot", type="primary", use_container_width=True):
            st.rerun()


# -----------------------------------------------------------------------------
# 7. لوحة الإدخال المتدرجة (Dynamic Cascading Inputs)
# -----------------------------------------------------------------------------
st.html("""
<div class="section-badge-title">🎯 خطوة 1: اختيار الماركة والموديل لكل جهاز (Brand & Model Selection)</div>
""")

# مدخل تحديد عدد الموبايلات للمقارنة في مقدمة شاشة الإدخال
num_mob_col1, num_mob_col2 = st.columns([1, 2], gap="medium")
with num_mob_col1:
    num_mobiles = st.number_input(
        "🔢 عدد الموبايلات للمقارنة (Number of Mobiles):",
        min_value=1,
        max_value=8,
        value=st.session_state.get("num_mobiles_input", 3),
        step=1,
        key="num_mobiles_input"
    )
with num_mob_col2:
    st.markdown(f"""
    <div style="padding: 9px 16px; background: #0F172A; border: 1.5px solid #3B82F6; border-radius: 12px; margin-top: 4px; display: flex; align-items: center; justify-content: space-between;">
        <span style="font-weight: 800; color: #FFFFFF; font-size: {cur_zoom['label']};">📱 عدد الأجهزة الحالية في المقارنة: <strong style="color: #60A5FA;">{num_mobiles}</strong></span>
        <span style="font-size: calc({cur_zoom['body']} * 0.85); color: #93C5FD; font-weight: 700;">(يمكنك المقارنة حتى 8 أجهزة)</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

input_mobiles = []
num_mob = int(num_mobiles)
cards_per_row = 3 if num_mob <= 3 else (4 if num_mob == 4 else 3)

for chunk_start in range(0, num_mob, cards_per_row):
    chunk_end = min(chunk_start + cards_per_row, num_mob)
    chunk_size = chunk_end - chunk_start
    row_cols = st.columns(chunk_size, gap="large")
    
    for i in range(chunk_start, chunk_end):
        col = row_cols[i - chunk_start]
        b_key = f"brand_val_{i}"
        m_key = f"model_val_{i}"

        saved_brand = st.session_state.get(b_key, brand_options[0])
        if saved_brand not in brand_options:
            saved_brand = brand_options[0]

        with col:
            with st.container(border=True):
                st.markdown(f"""
                <div class="phone-select-header">
                    <span>📱 الجهاز {i + 1}</span>
                    <span style="font-size: 11.5px; color: #93C5FD; font-weight: 700;">(Device {i + 1})</span>
                </div>
                """, unsafe_allow_html=True)
                
                # اختيار الماركة
                selected_brand = st.selectbox(
                    f"🏷️ الماركة (Brand) - جهاز {i + 1}:",
                    options=brand_options,
                    index=brand_options.index(saved_brand),
                    key=f"select_brand_{i}"
                )
                st.session_state[b_key] = selected_brand

                # اختيار الموديل
                available_models = BRANDS_CATALOG.get(selected_brand, [])
                saved_model = st.session_state.get(m_key, available_models[0] if available_models else "")
                
                model_idx = 0
                if saved_model in available_models:
                    model_idx = available_models.index(saved_model)

                selected_model = st.selectbox(
                    f"📱 الموديل (Model) - جهاز {i + 1}:",
                    options=available_models,
                    index=model_idx,
                    key=f"select_model_{i}"
                )
                st.session_state[m_key] = selected_model

                # تكوين اسم الهاتف
                brand_short = selected_brand.split(" (")[0].replace("Apple iPhone", "Apple").replace("Xiaomi & Poco & Redmi", "Xiaomi").replace("Vivo & iQOO", "Vivo").replace("Nothing & Nokia", "")
                if brand_short.lower() in selected_model.lower():
                    full_phone_name = selected_model
                else:
                    full_phone_name = f"{brand_short} {selected_model}".strip()

                input_mobiles.append(full_phone_name)

                # زر عرض وتحميل دليل إعدادات ومسارات الموديل (شاشة كاملة وتفاعلية)
                if st.button(
                    f"👁️ عرض وتحميل دليل {selected_model}",
                    key=f"btn_view_settings_{i}_{selected_brand}_{selected_model}",
                    type="primary",
                    use_container_width=True
                ):
                    show_settings_guide_dialog(
                        full_phone_name=full_phone_name,
                        selected_brand=selected_brand,
                        selected_model=selected_model
                    )

st.markdown("<div style='margin-top: 18px; margin-bottom: 24px;'>", unsafe_allow_html=True)
btn_search = st.button("🚀 تحديث وبدء المقارنة الشاملة (Compare Phones)", type="primary", use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 8. محرك التحليل ومعالجة البيانات
# -----------------------------------------------------------------------------
active_phones = [m for m in input_mobiles if m]
if active_phones:
    spec_col_name = "المواصفة (Specification)"
    df_dict = {
        spec_col_name: [f"{s['label_ar']} ({s['label_en']})" for s in SPEC_DEFINITIONS]
    }

    all_specs = {}
    for phone_name in active_phones:
        specs = get_mobile_specs(phone_name)
        all_specs[phone_name] = specs
        col_values = []
        for s in SPEC_DEFINITIONS:
            val = specs.get(s["key"], "غير متوفر")
            col_values.append(val)
        df_dict[f"{phone_name}"] = col_values

    winners_list = []
    reasons_list = []
    for s in SPEC_DEFINITIONS:
        winner_phone, reason = evaluate_spec_winner(s["key"], all_specs)
        winners_list.append(winner_phone)
        reasons_list.append(reason)

    df_dict["سبب الأفضلية"] = reasons_list
    final_verdict = evaluate_overall_winner(all_specs, SPEC_DEFINITIONS)
    comparison_df = pd.DataFrame(df_dict)

    st.session_state.comparison_data = comparison_df
    st.session_state.selected_phones = active_phones
    st.session_state.all_specs = all_specs
    st.session_state.winners_list = winners_list
    st.session_state.reasons_list = reasons_list
    st.session_state.final_verdict = final_verdict
    st.session_state.searched = True


# -----------------------------------------------------------------------------
# 9. دالة بناء جدول HTML الفائق الوضوح
# -----------------------------------------------------------------------------
def render_modern_table(
    df: pd.DataFrame,
    active_phones: list,
    all_specs: dict,
    winners_list: list,
    reasons_list: list,
    final_verdict: dict,
    zoom_token: dict,
    show_categories: bool = True
) -> str:
    """توليد جدول HTML احترافي واضح الخط ومنسق بالكامل"""
    num_cols = len(df.columns)
    
    headers_parts = []
    headers_parts.append(
        f'<th style="background: #0F172A; color: #FFFFFF; font-weight: 900; border: 2px solid #1E293B; min-width: 240px;">📋 المواصفة (Specification)</th>'
    )

    for i, phone_name in enumerate(active_phones):
        info = PHONE_COLOR_GRADIENTS[i % len(PHONE_COLOR_GRADIENTS)]
        headers_parts.append(
            f'<th style="background: {info["gradient"]}; color: #FFFFFF; border-bottom: 5px solid {info["border"]}; font-weight: 900; border: 2px solid #CBD5E1; min-width: 220px; text-align: center;">📱 {phone_name}</th>'
        )

    headers_parts.append(
        f'<th style="background: #0F172A; color: #60A5FA; font-weight: 900; border: 2px solid #1E293B; min-width: 240px;">💡 سبب الأفضلية</th>'
    )

    thead_html = f"<tr>{''.join(headers_parts)}</tr>"

    rows_parts = []
    current_category = None

    for row_idx, spec_def in enumerate(SPEC_DEFINITIONS):
        cat = spec_def["category"]
        
        if show_categories and cat != current_category:
            current_category = cat
            c_info = CATEGORY_STYLES.get(cat, {"name": cat, "bg": "#F8FAFC", "color": "#0F172A", "icon": "📌"})
            rows_parts.append(
                f'<tr style="background-color: {c_info["bg"]};"><td colspan="{num_cols}" style="font-weight: 900; color: {c_info["color"]}; border-top: 2.5px solid #94A3B8; border-bottom: 2.5px solid #94A3B8; padding: 12px 18px; font-size: {zoom_token["tbl_th"]};">{c_info["icon"]} {c_info["name"]}</td></tr>'
            )

        row_data = df.iloc[row_idx]
        winner_phone = winners_list[row_idx] if row_idx < len(winners_list) else None
        reason_text = reasons_list[row_idx] if row_idx < len(reasons_list) else "-"

        is_even = (row_idx % 2 == 0)
        row_bg = "#FFFFFF" if is_even else "#F8FAFC"
        cells_parts = []

        spec_label = row_data["المواصفة (Specification)"]
        cells_parts.append(
            f'<td style="background-color: #F1F5F9; font-weight: 900; color: #0F172A; text-align: right; border: 1.5px solid #CBD5E1;">{spec_label}</td>'
        )

        for phone_idx, phone_name in enumerate(active_phones):
            val = row_data[phone_name]
            is_winner = (phone_name == winner_phone and winner_phone != "-")
            
            if is_winner:
                cell_bg = "#DCFCE7"  # أخضر ناعم واضح
                cell_border = "border: 2.5px solid #10B981;"
                text_color = "#14532D"
                font_weight = "900"
                badge_html = '<div style="margin-top: 4px;"><span class="winner-badge">⭐ الأفضل</span></div>'
                cell_content = f'<div>{val}</div>{badge_html}'
            else:
                cell_bg = row_bg
                cell_border = "border: 1.5px solid #CBD5E1;"
                text_color = "#0F172A"
                font_weight = "700"
                cell_content = f'<div>{val}</div>'

            cells_parts.append(
                f'<td style="background-color: {cell_bg}; {cell_border} color: {text_color}; font-weight: {font_weight}; text-align: center;">{cell_content}</td>'
            )

        # خلية سبب الأفضلية
        cells_parts.append(
            f'<td style="background-color: #EFF6FF; border: 1.5px solid #CBD5E1; color: #1E40AF; font-weight: 800; text-align: right;">{reason_text}</td>'
        )

        rows_parts.append(f'<tr style="background-color: {row_bg};">{"".join(cells_parts)}</tr>')

    # صف التقييم النهائي
    if final_verdict and final_verdict.get("winner_name"):
        w_name = final_verdict["winner_name"]
        verdict_cells = []
        verdict_cells.append(
            f'<td style="background: #FEF3C7; color: #92400E; font-weight: 900; text-align: right; border: 2.5px solid #F59E0B;">🏆 التقييم النهائي</td>'
        )

        for phone_name in active_phones:
            if phone_name == w_name:
                v_content = f'<span style="background: #15803D; color: #FFFFFF; font-weight: 900; padding: 6px 16px; border-radius: 8px; font-size: calc({zoom_token["tbl_td"]} * 1.05); display: inline-block;">👑 الفائز الأفضل إجمالاً</span>'
                cell_bg = "#DCFCE7"
            else:
                score = final_verdict.get("all_scores", {}).get(phone_name, 0)
                v_content = f'<span style="color: #78350F; font-weight: 800;">تفوق في ({score}) مواصفات</span>'
                cell_bg = "#FEF3C7"

            verdict_cells.append(
                f'<td style="background: {cell_bg}; text-align: center; border: 2.5px solid #F59E0B;">{v_content}</td>'
            )

        verdict_cells.append(
            f'<td style="background: #FEF3C7; color: #1E40AF; font-weight: 900; text-align: right; border: 2.5px solid #F59E0B;">الأكثر حصداً للمواصفات الفائزة</td>'
        )

        rows_parts.append(f'<tr style="background: #FEF3C7;">{"".join(verdict_cells)}</tr>')

    tbody_html = "".join(rows_parts)

    table_full_html = (
        f'<div class="table-container-modern">'
        f'<table id="comparison-table-main">'
        f'<thead>{thead_html}</thead>'
        f'<tbody>{tbody_html}</tbody>'
        f'</table>'
        f'</div>'
    )
    return table_full_html


# -----------------------------------------------------------------------------
# 9.4 دالة بناء جدول تصفية الأقسام (Filtered Category Table)
# -----------------------------------------------------------------------------
def render_filtered_table(
    df: pd.DataFrame,
    active_phones: list,
    filtered_indices: list,
    winners_list: list,
    reasons_list: list,
    zoom_token: dict,
    show_categories: bool = True
) -> str:
    """
    توليد جدول HTML مخصص لتبويب تصفية الأقسام:
    - عناوين الأعمدة بلون أصفر مميز وفاخر (#FACC15 / #FEF08A).
    - عمود المواصفة بلون لبني فاتح (#E0F2FE).
    - إبراز الخلية الفائزة وشارة الأفضل وتوضيح سبب الأفضلية.
    """
    if not filtered_indices:
        return """
        <div style="background: #FEF2F2; border: 1.5px solid #F87171; border-radius: 12px; padding: 16px 20px; color: #991B1B; font-weight: 800; text-align: center; margin-top: 15px;">
            ⚠️ لم يتم اختيار أي قسم للعرض. يرجى اختيار قسم واحد على الأقل من القائمة أعلاه.
        </div>
        """

    num_cols = len(active_phones) + 2
    
    # 1. عناوين الأعمدة باللون الأصفر المميز والفاخر مع تمييز عمود المواصفات
    headers_parts = []
    headers_parts.append(
        f'<th style="background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%); color: #FFFFFF; font-weight: 900; border: 2.5px solid #38BDF8; min-width: 250px; text-align: right; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_th"]}; box-shadow: inset 0 -3px 0 #38BDF8;">📋 المواصفة (Specification)</th>'
    )

    for i, phone_name in enumerate(active_phones):
        headers_parts.append(
            f'<th style="background: linear-gradient(135deg, #FEF08A 0%, #FDE047 100%); color: #713F12; font-weight: 900; border: 2px solid #CA8A04; min-width: 220px; text-align: center; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_th"]}; box-shadow: inset 0 -2px 0 #CA8A04;">📱 {phone_name}</th>'
        )

    headers_parts.append(
        f'<th style="background: linear-gradient(135deg, #FDE047 0%, #FACC15 100%); color: #713F12; font-weight: 900; border: 2px solid #CA8A04; min-width: 240px; text-align: right; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_th"]}; box-shadow: inset 0 -2px 0 #CA8A04;">💡 سبب الأفضلية</th>'
    )

    thead_html = f"<tr>{''.join(headers_parts)}</tr>"

    rows_parts = []
    current_category = None

    for row_pos, global_idx in enumerate(filtered_indices):
        spec_def = SPEC_DEFINITIONS[global_idx]
        cat = spec_def["category"]
        
        # فاصل الفئة عند تغيرها
        if show_categories and cat != current_category:
            current_category = cat
            c_info = CATEGORY_STYLES.get(cat, {"name": cat, "bg": "#F8FAFC", "color": "#0F172A", "icon": "📌"})
            rows_parts.append(
                f'<tr style="background-color: {c_info["bg"]};"><td colspan="{num_cols}" style="font-weight: 900; color: {c_info["color"]}; border-top: 2.5px solid #94A3B8; border-bottom: 2.5px solid #94A3B8; padding: 10px 18px; font-size: {zoom_token["tbl_th"]};">{c_info["icon"]} {c_info["name"]}</td></tr>'
            )

        row_data = df.iloc[global_idx]
        winner_phone = winners_list[global_idx] if global_idx < len(winners_list) else None
        reason_text = reasons_list[global_idx] if global_idx < len(reasons_list) else "-"

        is_even = (row_pos % 2 == 0)
        row_bg = "#FFFFFF" if is_even else "#F8FAFC"
        cells_parts = []

        # 2. عمود المواصفة بلون لبني فاتح مميز (#E0F2FE مع حافة كحلية داكنة)
        spec_label = row_data["المواصفة (Specification)"]
        cells_parts.append(
            f'<td style="background: linear-gradient(135deg, #E0F2FE 0%, #BAE6FD 100%); font-weight: 900; color: #0C4A6E; text-align: right; border: 2px solid #38BDF8; border-right: 6px solid #0284C7; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_td"]}; box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);">📋 {spec_label}</td>'
        )

        # خلايا الهواتف
        for phone_idx, phone_name in enumerate(active_phones):
            val = row_data[phone_name]
            is_winner = (phone_name == winner_phone and winner_phone != "-")
            
            if is_winner:
                cell_bg = "#DCFCE7"  # أخضر ناعم للفائز
                cell_border = "border: 2.5px solid #10B981;"
                text_color = "#14532D"
                font_weight = "900"
                badge_html = '<div style="margin-top: 4px;"><span class="winner-badge">⭐ الأفضل</span></div>'
                cell_content = f'<div>{val}</div>{badge_html}'
            else:
                cell_bg = row_bg
                cell_border = "border: 1.5px solid #CBD5E1;"
                text_color = "#0F172A"
                font_weight = "700"
                cell_content = f'<div>{val}</div>'

            cells_parts.append(
                f'<td style="background-color: {cell_bg}; {cell_border} color: {text_color}; font-weight: {font_weight}; text-align: center; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_td"]};">{cell_content}</td>'
            )

        # خلية سبب الأفضلية
        cells_parts.append(
            f'<td style="background-color: #EFF6FF; border: 1.5px solid #CBD5E1; color: #1E40AF; font-weight: 800; text-align: right; padding: {zoom_token["tbl_pad"]}; font-size: {zoom_token["tbl_td"]};">{reason_text}</td>'
        )

        rows_parts.append(f'<tr style="background-color: {row_bg};">{"".join(cells_parts)}</tr>')

    tbody_html = "".join(rows_parts)

    table_full_html = f"""
    <div class="table-container-modern" style="border: 2.5px solid #EAB308; box-shadow: 0 8px 24px rgba(234, 179, 8, 0.12);">
        <table id="filter-table-main" style="width: 100%; border-collapse: collapse; direction: rtl; text-align: right;">
            <thead>{thead_html}</thead>
            <tbody>{tbody_html}</tbody>
        </table>
    </div>
    """
    return table_full_html


# -----------------------------------------------------------------------------
# 9.5 دالة بناء تبويب التحليل البياني المطور (Visual Analytics Tab)
# -----------------------------------------------------------------------------
def render_visual_charts_tab(
    active_phones: list,
    all_specs: dict,
    winners_list: list,
    final_verdict: dict,
    zoom_token: dict
):
    """عرض لوحة تحليلات ورسوم بيانية تفاعلية فائقة التنسيق بنصوص بيضاء بالكامل وتصميم منسجم 100%"""
    if not active_phones:
        st.info("يرجى اختيار هاتفين على الأقل لعرض التحليل البياني.")
        return

    # استخراج الأرقام والمواصفات الرقمية لكل هاتف بدقة
    chart_data = []
    for idx, p in enumerate(active_phones):
        spec = all_specs.get(p, {})
        
        # 1. سعة البطارية (mAh)
        bat_str = str(spec.get("battery", ""))
        bat_match = re.search(r'(\d{4})', bat_str)
        battery_mah = int(bat_match.group(1)) if bat_match else 0
        
        # 2. سرعة الشحن (Watt)
        ch_str = str(spec.get("charging", ""))
        ch_match = re.search(r'(\d+)\s*W', ch_str, re.IGNORECASE)
        charging_w = int(ch_match.group(1)) if ch_match else 25

        # 3. الذاكرة العشوائية RAM (GB)
        ram_str = str(spec.get("ram", ""))
        ram_nums = re.findall(r'(\d+)\s*GB', ram_str, re.IGNORECASE)
        ram_gb = max([int(r) for r in ram_nums]) if ram_nums else 8

        # 4. الوزن (Grams)
        w_str = str(spec.get("weight", ""))
        w_match = re.search(r'(\d{2,3})', w_str)
        weight_g = int(w_match.group(1)) if w_match else 0

        # 5. نقاط الفوز
        wins_count = sum(1 for w in winners_list if w == p)
        
        # 6. السعر بالجنيه المصري
        price_str = spec.get("price", "غير محدد")

        chart_data.append({
            "phone": p,
            "wins": wins_count,
            "wins_label": f"{wins_count} مواصفة",
            "battery": battery_mah,
            "battery_label": f"{battery_mah:,} mAh" if battery_mah else "-",
            "charging": charging_w,
            "charging_label": f"{charging_w} W",
            "ram": ram_gb,
            "ram_label": f"{ram_gb} GB",
            "weight": weight_g,
            "weight_label": f"{weight_g} g" if weight_g else "-",
            "price": price_str
        })

    df_charts = pd.DataFrame(chart_data)

    # ألوان ساطعة متناسقة لكل هاتف
    chart_palette = [
        "#3B82F6", "#A855F7", "#FB923C", "#14B8A6",
        "#FB7185", "#818CF8", "#38BDF8", "#84CC16"
    ]

    # تحديد قادة المؤشرات القياسية (Leaders)
    best_wins = df_charts.loc[df_charts["wins"].idxmax()] if not df_charts.empty else None
    best_bat = df_charts.loc[df_charts["battery"].idxmax()] if not df_charts.empty else None
    best_ch = df_charts.loc[df_charts["charging"].idxmax()] if not df_charts.empty else None
    best_ram = df_charts.loc[df_charts["ram"].idxmax()] if not df_charts.empty else None
    
    valid_weight_df = df_charts[df_charts["weight"] > 0]
    best_weight = valid_weight_df.loc[valid_weight_df["weight"].idxmin()] if not valid_weight_df.empty else None

    # دالة مساعدة لإنشاء مخطط Altair أفقي فاخر بنصوص بيضاء ومسافات واسعة تمنع التداخل
    def build_metric_chart(val_col, label_col, x_title, domain_max=None):
        raw_max = df_charts[val_col].max() if not df_charts.empty else 0
        max_v = domain_max or (raw_max * 1.28 if raw_max > 0 else 100)
        
        bars = alt.Chart(df_charts).mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            height=24
        ).encode(
            y=alt.Y("phone:N", title=None, sort=None, axis=alt.Axis(
                orient="left",
                labelColor="#FFFFFF",
                labelFont="Cairo",
                labelFontSize=13,
                labelFontWeight="bold",
                labelLimit=320,
                labelPadding=16,
                domainColor="rgba(255,255,255,0.25)",
                tickColor="rgba(255,255,255,0.25)"
            )),
            x=alt.X(f"{val_col}:Q", title=x_title, scale=alt.Scale(domain=[0, max_v]), axis=alt.Axis(
                labelColor="#93C5FD",
                titleColor="#93C5FD",
                labelFont="Cairo",
                titleFont="Cairo",
                titleFontSize=12,
                gridColor="rgba(255,255,255,0.08)",
                domainColor="rgba(255,255,255,0.25)"
            )),
            color=alt.Color("phone:N", scale=alt.Scale(domain=active_phones, range=chart_palette[:len(active_phones)]), legend=None),
            tooltip=[
                alt.Tooltip("phone:N", title="الهاتف"),
                alt.Tooltip(f"{val_col}:Q", title=x_title)
            ]
        )

        labels = alt.Chart(df_charts).mark_text(
            align="left",
            baseline="middle",
            dx=8,
            color="#FFFFFF",
            font="Cairo",
            fontSize=12.5,
            fontWeight="bold"
        ).encode(
            y=alt.Y("phone:N", sort=None),
            x=alt.X(f"{val_col}:Q"),
            text=alt.Text(f"{label_col}:N")
        )

        return (bars + labels).properties(
            height=max(120, len(active_phones) * 40 + 20),
            autosize=alt.AutoSizeParams(type="pad", contains="padding")
        ).configure_view(
            strokeWidth=0,
            fill="transparent"
        )

    # 1. بانر رأس التبويب
    st.html("""
    <div class="analytics-tab-container">
        <div class="analytics-header-banner">
            <div>
                <div class="analytics-header-title">📈 لوحة التحليل البياني والمؤشرات المقارنة (Visual Charts Dashboard)</div>
                <div class="analytics-header-desc">تحليل بصري متقدم ومقارنة فورية لأهم الأرقام القياسية: البطارية، الشحن، الذاكرة، الوزن، ونقاط التفوق.</div>
            </div>
            <div class="analytics-header-badge">⚡ بيانات دقيقة ومباشرة</div>
        </div>
    """)

    # 2. بطاقات المتصدرين السريعة (Leaderboard KPI Cards)
    k_col1, k_col2, k_col3, k_col4, k_col5 = st.columns(5, gap="small")
    
    with k_col1:
        w_p = best_wins["phone"].split(" (")[0] if best_wins is not None else "-"
        w_v = f"{best_wins['wins']} مواصفة" if best_wins is not None else "-"
        st.html(f"""
        <div class="analytics-kpi-card" style="border-top: 4px solid #10B981;">
            <div class="analytics-kpi-title">🏆 الأكثر فوزاً إجمالاً</div>
            <div class="analytics-kpi-val" style="color: #34D399;">{w_v}</div>
            <div class="analytics-kpi-phone" title="{w_p}">📱 {w_p}</div>
        </div>
        """)

    with k_col2:
        b_p = best_bat["phone"].split(" (")[0] if best_bat is not None else "-"
        b_v = f"{best_bat['battery']:,} mAh" if best_bat is not None and best_bat['battery'] > 0 else "-"
        st.html(f"""
        <div class="analytics-kpi-card" style="border-top: 4px solid #3B82F6;">
            <div class="analytics-kpi-title">🔋 الأكبر بطارية</div>
            <div class="analytics-kpi-val" style="color: #60A5FA;">{b_v}</div>
            <div class="analytics-kpi-phone" title="{b_p}">📱 {b_p}</div>
        </div>
        """)

    with k_col3:
        c_p = best_ch["phone"].split(" (")[0] if best_ch is not None else "-"
        c_v = f"{best_ch['charging']} Watt" if best_ch is not None else "-"
        st.html(f"""
        <div class="analytics-kpi-card" style="border-top: 4px solid #F59E0B;">
            <div class="analytics-kpi-title">⚡ الأسرع شحناً</div>
            <div class="analytics-kpi-val" style="color: #FBBF24;">{c_v}</div>
            <div class="analytics-kpi-phone" title="{c_p}">📱 {c_p}</div>
        </div>
        """)

    with k_col4:
        r_p = best_ram["phone"].split(" (")[0] if best_ram is not None else "-"
        r_v = f"{best_ram['ram']} GB" if best_ram is not None else "-"
        st.html(f"""
        <div class="analytics-kpi-card" style="border-top: 4px solid #A855F7;">
            <div class="analytics-kpi-title">💾 الأعلى رام RAM</div>
            <div class="analytics-kpi-val" style="color: #C084FC;">{r_v}</div>
            <div class="analytics-kpi-phone" title="{r_p}">📱 {r_p}</div>
        </div>
        """)

    with k_col5:
        w_p = best_weight["phone"].split(" (")[0] if best_weight is not None else "-"
        w_v = f"{best_weight['weight']} g" if best_weight is not None else "-"
        st.html(f"""
        <div class="analytics-kpi-card" style="border-top: 4px solid #06B6D4;">
            <div class="analytics-kpi-title">⚖️ الأخف وزناً</div>
            <div class="analytics-kpi-val" style="color: #22D3EE;">{w_v}</div>
            <div class="analytics-kpi-phone" title="{w_p}">📱 {w_p}</div>
        </div>
        """)

    # 3. المخطط الرئيسي لإجمالي نقاط الفوز
    st.html("""
    <div class="analytics-chart-box">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">🏆 إجمالي عدد المواصفات المتفوقة لكل هاتف (Total Specs Won)</div>
            <div class="analytics-chart-badge">مؤشر التفوق الشامل</div>
        </div>
    """)
    chart_wins = build_metric_chart("wins", "wins_label", "عدد المواصفات الفائزة")
    st.altair_chart(chart_wins, use_container_width=True)
    st.html("</div>")

    # 4. الرسوم البيانية التخصصية (بنود المقارنة بالتتابع الرأسي أسفل بعضها البعض)
    # أ. سعة البطارية
    st.html("""
    <div class="analytics-chart-box">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">🔋 سعة البطارية (Battery Capacity - mAh)</div>
            <div class="analytics-chart-badge">الأعلى أفضل</div>
        </div>
    """)
    chart_bat = build_metric_chart("battery", "battery_label", "سعة البطارية (mAh)")
    st.altair_chart(chart_bat, use_container_width=True)
    st.html("</div>")

    # ب. سرعة الشحن السلكي
    st.html("""
    <div class="analytics-chart-box">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">⚡ سرعة الشحن السلكي (Charging Speed - Watt)</div>
            <div class="analytics-chart-badge">سرعة ملء الطاقة</div>
        </div>
    """)
    chart_ch = build_metric_chart("charging", "charging_label", "قدرة الشحن (Watt)")
    st.altair_chart(chart_ch, use_container_width=True)
    st.html("</div>")

    # ج. سعة الذاكرة العشوائية RAM
    st.html("""
    <div class="analytics-chart-box">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">💾 سعة الذاكرة العشوائية (RAM Memory - GB)</div>
            <div class="analytics-chart-badge">أداء المهام المتعددة</div>
        </div>
    """)
    chart_ram = build_metric_chart("ram", "ram_label", "الذاكرة العشوائية RAM (GB)")
    st.altair_chart(chart_ram, use_container_width=True)
    st.html("</div>")

    # د. الوزن الإجمالي
    st.html("""
    <div class="analytics-chart-box">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">⚖️ الوزن الإجمالي (Device Weight - Grams)</div>
            <div class="analytics-chart-badge">الأخف أسهل في الحمل</div>
        </div>
    """)
    chart_weight = build_metric_chart("weight", "weight_label", "الوزن (جرام)")
    st.altair_chart(chart_weight, use_container_width=True)
    st.html("</div>")

    # 5. جدول ملخص الأرقام المباشرة (Direct Matrix Table)
    st.html("""
    <div class="analytics-chart-box" style="margin-top: 10px;">
        <div class="analytics-chart-header">
            <div class="analytics-chart-title">📊 جدول ملخص الأرقام والمؤشرات السريعة (Quick Numerical Matrix)</div>
            <div class="analytics-chart-badge">مقارنة رقمية فورية</div>
        </div>
    """)
    
    # بناء صفوف الجدول
    matrix_rows = []
    for idx, row in df_charts.iterrows():
        p_name = row["phone"]
        info = PHONE_COLOR_GRADIENTS[idx % len(PHONE_COLOR_GRADIENTS)]
        
        # علامات التميز
        is_top_win = (best_wins is not None and row["wins"] == best_wins["wins"] and row["wins"] > 0)
        is_top_bat = (best_bat is not None and row["battery"] == best_bat["battery"] and row["battery"] > 0)
        is_top_ch = (best_ch is not None and row["charging"] == best_ch["charging"] and row["charging"] > 0)
        is_top_ram = (best_ram is not None and row["ram"] == best_ram["ram"] and row["ram"] > 0)
        is_top_w = (best_weight is not None and row["weight"] == best_weight["weight"] and row["weight"] > 0)

        win_pill = f'<span style="color: #34D399; font-weight: 900;">⭐ {row["wins"]} فوز</span>' if is_top_win else f'{row["wins"]}'
        bat_pill = f'<span style="color: #60A5FA; font-weight: 900;">⭐ {row["battery"]:,} mAh</span>' if is_top_bat else f'{row["battery"]:,} mAh'
        ch_pill = f'<span style="color: #FBBF24; font-weight: 900;">⭐ {row["charging"]} W</span>' if is_top_ch else f'{row["charging"]} W'
        ram_pill = f'<span style="color: #C084FC; font-weight: 900;">⭐ {row["ram"]} GB</span>' if is_top_ram else f'{row["ram"]} GB'
        w_pill = f'<span style="color: #22D3EE; font-weight: 900;">⭐ {row["weight"]} g</span>' if is_top_w else f'{row["weight"]} g'

        matrix_rows.append(f"""
        <tr>
            <td style="text-align: right; font-weight: 900; color: #FFFFFF;">
                <span style="display: inline-block; width: 10px; height: 10px; border-radius: 50%; background: {info['border']}; margin-left: 8px;"></span>
                📱 {p_name}
            </td>
            <td>{win_pill}</td>
            <td>{bat_pill}</td>
            <td>{ch_pill}</td>
            <td>{ram_pill}</td>
            <td>{w_pill}</td>
            <td style="color: #34D399; font-weight: 900;">{row['price']}</td>
        </tr>
        """)

    matrix_rows_html = "".join(matrix_rows)
    st.html(f"""
        <div style="overflow-x: auto;">
            <table class="analytics-table-matrix">
                <thead>
                    <tr>
                        <th style="text-align: right;">📱 الهاتف (Phone)</th>
                        <th>🏆 عدد بنود الفوز</th>
                        <th>🔋 البطارية (mAh)</th>
                        <th>⚡ الشحن (Watt)</th>
                        <th>💾 الرام (RAM)</th>
                        <th>⚖️ الوزن (Weight)</th>
                        <th>💰 السعر التقريبي</th>
                    </tr>
                </thead>
                <tbody>
                    {matrix_rows_html}
                </tbody>
            </table>
        </div>
    </div>
    </div>
    """)


# -----------------------------------------------------------------------------
# 10. قسم النتائج والداشبورد (Results Dashboard)
# -----------------------------------------------------------------------------
if st.session_state.searched and st.session_state.comparison_data is not None:
    comparison_df = st.session_state.comparison_data
    all_specs = st.session_state.all_specs
    active_phones = st.session_state.selected_phones
    winners_list = st.session_state.winners_list
    reasons_list = st.session_state.reasons_list
    final_verdict = st.session_state.final_verdict

    st.markdown("""
    <div style="margin: 40px 0 30px 0; border-top: 3.5px solid #2563EB; border-radius: 4px;"></div>
    """, unsafe_allow_html=True)

    # بطاقة بطل المقارنة
    if final_verdict and final_verdict.get("winner_name"):
        w_name = final_verdict["winner_name"]
        win_count = final_verdict.get("win_count", 0)
        total_specs = final_verdict.get("total_specs", len(SPEC_DEFINITIONS))
        st.html(f"""
        <div class="champion-card">
            <div>
                <div class="champion-title">👑 بطل المقارنة: {w_name}</div>
                <div class="champion-desc">حقق المركز الأول بالتفوق في <strong>{win_count} من أصل {total_specs} مواصفة</strong> تم فحصها بدقة.</div>
            </div>
            <div>
                <span style="background: #FFFFFF; color: #065F46; font-weight: 900; padding: 10px 22px; border-radius: 12px; font-size: {cur_zoom['label']}; display: inline-block;">
                    نسبة التفوق: {int(round((win_count / total_specs) * 100))}% ⭐
                </span>
            </div>
        </div>
        """)

    # 1. المؤشرات السريعة (KPIs)
    k_c1, k_c2, k_c3, k_c4 = st.columns(4, gap="medium")
    with k_c1:
        st.html(f"""
        <div class="kpi-metric-box">
            <div class="kpi-metric-num">📱 {len(active_phones)}</div>
            <div class="kpi-metric-label">الهواتف المقارنة</div>
        </div>
        """)

    with k_c2:
        st.html(f"""
        <div class="kpi-metric-box">
            <div class="kpi-metric-num">📋 {len(SPEC_DEFINITIONS)}</div>
            <div class="kpi-metric-label">بنود الفحص والمقارنة</div>
        </div>
        """)

    with k_c3:
        w_short = final_verdict.get("winner_name", "-").split(" (")[0]
        st.html(f"""
        <div class="kpi-metric-box" style="border-color: #10B981; background: #ECFDF5;">
            <div class="kpi-metric-num" style="color: #047857;">⭐ {w_short}</div>
            <div class="kpi-metric-label" style="color: #065F46;">الخيار الأول</div>
        </div>
        """)

    with k_c4:
        st.html(f"""
        <div class="kpi-metric-box" style="border-color: #F59E0B; background: #FFFBEB;">
            <div class="kpi-metric-num" style="color: #B45309;">{final_verdict.get('win_count', 0)} فوز</div>
            <div class="kpi-metric-label" style="color: #92400E;">أعلى عدد مواصفات خضراء</div>
        </div>
        """)

    # 2. بطاقات استعراض الهواتف السريعة (شبكة منظمة تمنع التداخل)
    show_per_row = 3 if len(active_phones) <= 3 else (4 if len(active_phones) == 4 else 3)
    for c_start in range(0, len(active_phones), show_per_row):
        c_end = min(c_start + show_per_row, len(active_phones))
        c_len = c_end - c_start
        s_cols = st.columns(c_len, gap="medium")
        for idx in range(c_start, c_end):
            phone_name = active_phones[idx]
            p_spec = all_specs.get(phone_name, {})
            wins_for_this = sum(1 for w in winners_list if w == phone_name)
            info = PHONE_COLOR_GRADIENTS[idx % len(PHONE_COLOR_GRADIENTS)]
            with s_cols[idx - c_start]:
                st.html(f"""
                <div class="phone-showcase-box" style="border-top: 5px solid {info['border']};">
                    <div class="phone-showcase-title" style="color: {info['border']};">📱 {p_spec.get('brand', '')} {p_spec.get('model', phone_name)}</div>
                    <div class="phone-showcase-price">{p_spec.get('price', 'السعر غير محدد')}</div>
                    <div>
                        <span style="font-size: calc({cur_zoom['body']} * 0.9); color: #166534; font-weight: 900; background: #DCFCE7; padding: 4px 12px; border-radius: 8px; display: inline-block; margin-bottom: 8px;">
                            🟩 فاز في {wins_for_this} مواصفة
                        </span>
                    </div>
                    <div style="font-size: calc({cur_zoom['body']} * 0.88); color: #334155; line-height: 1.6; text-align: right;">
                        • <strong>المعالج:</strong> {p_spec.get('processor', '-').split('(')[0]}<br>
                        • <strong>البطارية:</strong> {p_spec.get('battery', '-')}<br>
                        • <strong>الشاشة:</strong> {p_spec.get('display', '-').split(',')[0]}
                    </div>
                </div>
                """)

    # 3. التبويبات التفاعلية (Tabs) بعناوين بيضاء واضحة
    st.html("""
    <div class="section-badge-title">📑 خطوة 2: استعراض جدول المقارنة والتحليلات والتقارير</div>
    """)

    tab_rich_table, tab_charts, tab_filter = st.tabs([
        "📋 جدول المقارنة (Comparison Table)",
        "📈 التحليل البياني (Visual Charts)",
        "🔍 تصفية الأقسام (Category Filter)"
    ])

    # التبويب 1: جدول المقارنة الملون
    with tab_rich_table:
        st.html("""
        <div style="background: #FEF9C3; border: 1.5px solid #FACC15; border-radius: 12px; padding: 12px 20px; margin-bottom: 16px; display: flex; align-items: center; gap: 10px; box-shadow: 0 2px 8px rgba(250, 204, 21, 0.18);">
            <span style="font-size: 22px;">💡</span>
            <span style="color: #713F12; font-weight: 800; font-size: 14.5px;">
                <strong>ملاحظة:</strong> تم تظليل الخلية الفائزة باللون الأخضر مع شارة <strong style="color: #15803D; background: #DCFCE7; padding: 2px 8px; border-radius: 6px; border: 1px solid #86EFAC;">⭐ الأفضل</strong>، وتوضيح سر التفوق في عمود <strong style="color: #1E40AF; background: #EFF6FF; padding: 2px 8px; border-radius: 6px; border: 1px solid #BFDBFE;">سبب الأفضلية</strong>.
            </span>
        </div>
        """)
        html_table_code = render_modern_table(
            df=comparison_df,
            active_phones=active_phones,
            all_specs=all_specs,
            winners_list=winners_list,
            reasons_list=reasons_list,
            final_verdict=final_verdict,
            zoom_token=cur_zoom,
            show_categories=show_category_headers
        )
        st.html(html_table_code)

        # خطوة 3: زر عرض وطباعة نتائج المقارنة أسفل جدول المقارنة
        st.html("""
        <div style="margin-top: 24px; margin-bottom: 8px;">
            <div class="section-badge-title">🎯 خطوة 3: عرض وطباعة ومشاركة نتائج المقارنة (Print & Export Results)</div>
        </div>
        """)

        col_p1, col_p2, col_p3 = st.columns([0.25, 0.50, 0.25], gap="medium")
        with col_p2:
            if st.button(
                "📑 عرض وطباعة نتائج المقارنة (View & Print Report)",
                key="btn_view_print_comparison_step3",
                type="primary",
                use_container_width=True
            ):
                show_comparison_print_dialog(
                    df=comparison_df,
                    active_phones=active_phones,
                    all_specs=all_specs,
                    winners_list=winners_list,
                    reasons_list=reasons_list,
                    final_verdict=final_verdict
                )

    # التبويب 2: الرسوم البيانية المتطورة والمؤشرات التحليلية
    with tab_charts:
        render_visual_charts_tab(
            active_phones=active_phones,
            all_specs=all_specs,
            winners_list=winners_list,
            final_verdict=final_verdict,
            zoom_token=cur_zoom
        )

    # التبويب 3: تصفية الأقسام
    with tab_filter:
        st.markdown("### 🔍 تخصيص وعرض أقسام محددة (Category Filter):")
        categories = list(set(s["category"] for s in SPEC_DEFINITIONS))
        selected_cat = st.multiselect(
            "اختر الأقسام المراد التركيز عليها:",
            options=categories,
            default=["Screen", "Performance", "Camera", "Battery", "Price"],
            format_func=lambda x: {
                "General": "معلومات عامة (General)",
                "Screen": "الشاشة والعرض (Screen)",
                "Performance": "الأداء والعتاد (Performance)",
                "Camera": "الكاميرات والتصوير (Camera)",
                "Battery": "البطارية والشحن (Battery)",
                "Software": "النظام والبرمجيات (Software)",
                "Design": "التصميم والوزن (Design)",
                "Build": "المتانة والمقاومة (Build)",
                "Connectivity": "الاتصال والشبكات (Connectivity)",
                "Price": "الأسعار (Price)",
                "Extra": "ميزات إضافية (Extra)"
            }.get(x, x)
        )

        filtered_keys = [s["key"] for s in SPEC_DEFINITIONS if s["category"] in selected_cat]
        filtered_indices = [idx for idx, s in enumerate(SPEC_DEFINITIONS) if s["key"] in filtered_keys]
        
        st.html(f"""
        <div style="background: #EFF6FF; border: 1.5px solid #BFDBFE; border-radius: 12px; padding: 12px 18px; margin-bottom: 14px; margin-top: 8px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
            <div style="display: flex; align-items: center; gap: 8px;">
                <span style="font-size: 20px;">🔍</span>
                <span style="color: #1E40AF; font-weight: 800; font-size: 14px;">
                    تم تصفية وعرض <strong>{len(filtered_indices)}</strong> مواصفة من أصل <strong>{len(SPEC_DEFINITIONS)}</strong> مواصفة إجمالية.
                </span>
            </div>
            <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                <span style="background: #FEF08A; color: #713F12; border: 1px solid #EAB308; padding: 3px 10px; border-radius: 8px; font-size: 12px; font-weight: 800;">
                    🟡 عناوين الجدول باللون الأصفر
                </span>
                <span style="background: #E0F2FE; color: #0369A1; border: 1px solid #BAE6FD; padding: 3px 10px; border-radius: 8px; font-size: 12px; font-weight: 800;">
                    🔹 عمود المواصفة باللون اللبني الفاتح
                </span>
            </div>
        </div>
        """)

        filtered_table_html = render_filtered_table(
            df=comparison_df,
            active_phones=active_phones,
            filtered_indices=filtered_indices,
            winners_list=winners_list,
            reasons_list=reasons_list,
            zoom_token=cur_zoom,
            show_categories=show_category_headers
        )
        st.html(filtered_table_html)


# -----------------------------------------------------------------------------
# 11. التذييل (Footer)
# -----------------------------------------------------------------------------
st.html(f"""
<div style="text-align: center; color: #64748B; font-size: calc({cur_zoom['body']} * 0.85); font-weight: 700; margin-top: 40px;">
    🚀 أداة مقارنة مواصفات الهواتف الذكية | تصميم عصري متكامل | تكبير ديناميكي ذكي | عناوين تبويبات باللون الأبيض الناصع | تسعير بالجنيه المصري (1$ = 50 ج).
</div>
""")
