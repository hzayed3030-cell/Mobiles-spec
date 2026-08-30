"""
specs_data.py
مستودع بيانات مواصفات الهواتف الذكية مع كتالوج الماركات والموديلات المتدرج (Cascading Brand & Model Catalog)
وتسعير دقيق بالجنيه المصري (باعتبار 1 دولار = 50 جنيه).
"""

import re
from typing import Dict, Any, List, Optional

USD_TO_EGP_RATE = 50.0  # سعر صرف الدولار مقابل الجنيه المصري

SPEC_DEFINITIONS = [
    {"key": "brand", "label_ar": "الشركة المصنعة", "label_en": "Brand", "category": "General"},
    {"key": "model", "label_ar": "اسم الموديل", "label_en": "Model Name", "category": "General"},
    {"key": "release_date", "label_ar": "تاريخ الإطلاق", "label_en": "Release Date", "category": "General"},
    {"key": "display", "label_ar": "الشاشة", "label_en": "Display", "category": "Screen"},
    {"key": "refresh_rate", "label_ar": "معدل التحديث", "label_en": "Refresh Rate", "category": "Screen"},
    {"key": "processor", "label_ar": "المعالج", "label_en": "Processor / Chipset", "category": "Performance"},
    {"key": "gpu", "label_ar": "معالج الرسوميات", "label_en": "GPU", "category": "Performance"},
    {"key": "ram", "label_ar": "الذاكرة العشوائية", "label_en": "RAM", "category": "Performance"},
    {"key": "storage", "label_ar": "سعة التخزين", "label_en": "Internal Storage", "category": "Performance"},
    {"key": "main_camera", "label_ar": "الكاميرا الخلفية", "label_en": "Main Camera", "category": "Camera"},
    {"key": "selfie_camera", "label_ar": "الكاميرا الأمامية", "label_en": "Selfie Camera", "category": "Camera"},
    {"key": "battery", "label_ar": "سعة البطارية", "label_en": "Battery Capacity", "category": "Battery"},
    {"key": "charging", "label_ar": "سرعة الشحن", "label_en": "Charging Speed", "category": "Battery"},
    {"key": "os", "label_ar": "نظام التشغيل", "label_en": "Operating System", "category": "Software"},
    {"key": "weight", "label_ar": "الوزن", "label_en": "Weight", "category": "Design"},
    {"key": "dimensions", "label_ar": "الأبعاد", "label_en": "Dimensions", "category": "Design"},
    {"key": "water_resistance", "label_ar": "مقاومة الماء والغبار", "label_en": "Water/Dust Resistance", "category": "Build"},
    {"key": "network", "label_ar": "الشبكات والاتصال", "label_en": "Network & 5G", "category": "Connectivity"},
    {"key": "price", "label_ar": "السعر بالجنيه المصري (1$ = 50 ج)", "label_en": "Price (EGP)", "category": "Price"},
    {"key": "features", "label_ar": "ميزات بارزة", "label_en": "Key Features", "category": "Extra"},
]

def format_egp(usd: float, rate: float = 50.0) -> str:
    """تحويل السعر بالدولار إلى الجنيه المصري بتنسيق منظم"""
    egp_val = int(round(usd * rate))
    return f"{egp_val:,} جنيه (EGP)"


# ==============================================================================
# كتالوج الماركات والموديلات المتدرج (Cascading Brand-to-Model Catalog)
# ==============================================================================

BRANDS_CATALOG: Dict[str, List[str]] = {
    "Samsung (سامسونج)": [
        "Galaxy S26 Ultra",
        "Galaxy S26 Plus",
        "Galaxy S26",
        "Galaxy S26 Edge",
        "Galaxy S25 Ultra",
        "Galaxy S25 Plus",
        "Galaxy S25",
        "Galaxy S25 Slim",
        "Galaxy S24 Ultra",
        "Galaxy S24 Plus",
        "Galaxy S24",
        "Galaxy S24 FE",
        "Galaxy S23 Ultra",
        "Galaxy S23 Plus",
        "Galaxy S23",
        "Galaxy S23 FE",
        "Galaxy Z Fold 6",
        "Galaxy Z Flip 6",
        "Galaxy Z Fold 5",
        "Galaxy Z Flip 5",
        "Galaxy A56 5G",
        "Galaxy A55 5G",
        "Galaxy A35 5G",
        "Galaxy A25 5G",
        "Galaxy A15 5G",
        "Galaxy A15 4G",
        "Galaxy A05s",
        "Galaxy M51",
        "Galaxy M54 5G",
        "Galaxy M53 5G",
        "Galaxy M52 5G",
        "Galaxy M34 5G",
        "Galaxy M14 5G"
    ],
    "Apple iPhone (آبل)": [
        "iPhone 17 Pro Max",
        "iPhone 17 Pro",
        "iPhone 17 Air",
        "iPhone 17",
        "iPhone 16 Pro Max",
        "iPhone 16 Pro",
        "iPhone 16 Plus",
        "iPhone 16",
        "iPhone 16e",
        "iPhone 15 Pro Max",
        "iPhone 15 Pro",
        "iPhone 15 Plus",
        "iPhone 15",
        "iPhone 14 Pro Max",
        "iPhone 14 Pro",
        "iPhone 14 Plus",
        "iPhone 14",
        "iPhone 13 Pro Max",
        "iPhone 13 Pro",
        "iPhone 13",
        "iPhone 12 Pro Max",
        "iPhone 12",
        "iPhone 11 Pro Max",
        "iPhone 11"
    ],
    "Xiaomi & Poco & Redmi (شاومي)": [
        "Xiaomi 15 Ultra",
        "Xiaomi 15 Pro",
        "Xiaomi 15",
        "Xiaomi 14 Ultra",
        "Xiaomi 14 Pro",
        "Xiaomi 14",
        "Xiaomi 14T Pro",
        "Xiaomi 14T",
        "Poco X7 Pro 5G",
        "Poco X7 5G",
        "Poco F6 Pro",
        "Poco F6",
        "Poco X6 Pro 5G",
        "Poco M6 Pro",
        "Redmi Note 14 Pro+ 5G",
        "Redmi Note 14 Pro 5G",
        "Redmi Note 13 Pro+ 5G",
        "Redmi Note 13 Pro 4G",
        "Redmi Note 13 4G",
        "Redmi 13",
        "Redmi K80 Pro"
    ],
    "OPPO (أوبو)": [
        "Find X9 Pro",
        "Find X8 Pro",
        "Find X8",
        "Find X7 Ultra",
        "Find N3 Fold",
        "Reno 13 Pro",
        "Reno 12 Pro 5G",
        "Reno 12 5G",
        "Reno 12F 5G",
        "Reno 11 Pro 5G",
        "Reno 11 5G",
        "Reno 11F 5G",
        "Reno 10 Pro+ 5G",
        "A98 5G",
        "A79 5G",
        "A78 4G",
        "A58 4G",
        "A38",
        "A18"
    ],
    "Vivo & iQOO (فيفو)": [
        "Vivo X200 Pro",
        "Vivo X200",
        "Vivo X100 Ultra",
        "Vivo X100 Pro",
        "Vivo X Fold 3 Pro",
        "Vivo V40 Pro",
        "Vivo V40 5G",
        "Vivo V40 Lite 5G",
        "Vivo V30 Pro",
        "Vivo V30 5G",
        "Vivo V30e",
        "Vivo V29 5G",
        "Vivo Y28",
        "Vivo Y18",
        "Vivo Y03",
        "iQOO 13",
        "iQOO 12",
        "iQOO Neo 9 Pro"
    ],
    "Realme (ريلمي)": [
        "Realme GT 7 Pro",
        "Realme GT 6",
        "Realme GT 6T",
        "Realme 13 Pro+ 5G",
        "Realme 13+ 5G",
        "Realme 12 Pro+ 5G",
        "Realme 12 Pro 5G",
        "Realme 12+ 5G",
        "Realme 12 5G",
        "Realme 11 Pro+ 5G",
        "Realme C67",
        "Realme C55",
        "Realme C53",
        "Realme Note 50"
    ],
    "Honor (هونر)": [
        "Honor Magic 7 Pro",
        "Honor Magic 7",
        "Honor Magic 6 Pro",
        "Honor Magic V3",
        "Honor 200 Pro",
        "Honor 200",
        "Honor 200 Lite",
        "Honor 90",
        "Honor X9b 5G",
        "Honor X8b",
        "Honor X7b",
        "Honor X6b"
    ],
    "Google Pixel (جوجل بكسل)": [
        "Google Pixel 9 Pro XL",
        "Google Pixel 9 Pro",
        "Google Pixel 9 Pro Fold",
        "Google Pixel 9",
        "Google Pixel 8a",
        "Google Pixel 8 Pro",
        "Google Pixel 8",
        "Google Pixel 7 Pro",
        "Google Pixel 7a"
    ],
    "OnePlus (وان بلس)": [
        "OnePlus 13",
        "OnePlus 13R",
        "OnePlus 12",
        "OnePlus 12R",
        "OnePlus Open",
        "OnePlus 11",
        "OnePlus Nord 4",
        "OnePlus Nord CE 4 5G",
        "OnePlus Nord CE 3 Lite"
    ],
    "Huawei (هواوي)": [
        "Huawei Pura 70 Ultra",
        "Huawei Pura 70 Pro",
        "Huawei Pura 70",
        "Huawei Mate 60 Pro",
        "Huawei Mate X5",
        "Huawei Nova 12 Pro",
        "Huawei Nova 12 SE",
        "Huawei Nova 11"
    ],
    "Infinix (إنفينيكس)": [
        "Infinix Zero 40 5G",
        "Infinix Zero 30 5G",
        "Infinix Note 40 Pro+ 5G",
        "Infinix Note 40 Pro",
        "Infinix Note 40",
        "Infinix Hot 50 Pro+",
        "Infinix Hot 50 4G",
        "Infinix Smart 8"
    ],
    "Tecno (تكنو)": [
        "Tecno Camon 30 Premier 5G",
        "Tecno Camon 30 Pro 5G",
        "Tecno Camon 30 5G",
        "Tecno Pova 6 Pro 5G",
        "Tecno Spark 30 Pro",
        "Tecno Spark 20 Pro+",
        "Tecno Pop 8"
    ],
    "Motorola (موتورولا)": [
        "Motorola Edge 50 Ultra",
        "Motorola Edge 50 Pro",
        "Motorola Edge 50 Fusion",
        "Motorola Razr 50 Ultra",
        "Moto G84 5G",
        "Moto G54 5G"
    ],
    "Nothing & Nokia (أخرى)": [
        "Nothing Phone (3)",
        "Nothing Phone (2)",
        "Nothing Phone (2a) Plus",
        "Nothing Phone (2a)",
        "Nokia XR21",
        "Nokia G42 5G",
        "Nokia C32"
    ]
}


# ==============================================================================
# قاعدة بيانات مفصلة ومسعرة بالجنيه المصري (1 دولار = 50 جنيه)
# ==============================================================================

MOBILE_DATABASE: Dict[str, Dict[str, Any]] = {
    # --- Samsung Series ---
    "Samsung Galaxy M51": {
        "brand": "Samsung",
        "model": "Galaxy M51",
        "release_date": "September 2020",
        "display": "6.7\" Super AMOLED Plus, 1080 x 2400 pixels, Gorilla Glass 3+",
        "refresh_rate": "60Hz",
        "processor": "Snapdragon 730G (8 nm)",
        "gpu": "Adreno 618",
        "ram": "6GB / 8GB LPDDR4X",
        "storage": "128GB (Dedicated microSDXC slot)",
        "main_camera": "Quad: 64 MP (wide) + 12 MP (ultrawide) + 5 MP (macro) + 5 MP (depth)",
        "selfie_camera": "32 MP, f/2.0, 4K@30fps",
        "battery": "7000 mAh (Monster Battery)",
        "charging": "25W Fast Charging (100% in 115 min) + Reverse wired charging",
        "os": "Android 10 (Upgradable to Android 12, One UI 4.1)",
        "weight": "213 g",
        "dimensions": "163.9 x 76.3 x 9.5 mm",
        "water_resistance": "Plastic Frame and Back (No official IP rating)",
        "network": "4G LTE, Wi-Fi 5, Bluetooth 5.0, NFC, 3.5mm Headphone Jack",
        "price": format_egp(280),  # 14,000 EGP
        "features": "Giant 7000mAh Battery, Dedicated SD Card Slot, 3.5mm Headphone Jack"
    },
    "Samsung Galaxy M52 5G": {
        "brand": "Samsung",
        "model": "Galaxy M52 5G",
        "release_date": "October 2021",
        "display": "6.7\" Super AMOLED Plus, 1080 x 2400 pixels, 120Hz",
        "refresh_rate": "120Hz",
        "processor": "Snapdragon 778G 5G (6 nm)",
        "gpu": "Adreno 642L",
        "ram": "6GB / 8GB",
        "storage": "128GB (microSDXC slot)",
        "main_camera": "Triple: 64 MP (wide) + 12 MP (ultrawide) + 5 MP (macro)",
        "selfie_camera": "32 MP, 4K@30fps",
        "battery": "5000 mAh",
        "charging": "25W Wired",
        "os": "Android 11 (Upgradable to Android 13, One UI 5)",
        "weight": "173 g",
        "dimensions": "164.2 x 76.4 x 7.4 mm",
        "water_resistance": "Plastic Back (Slim profile)",
        "network": "5G, Wi-Fi 6, Bluetooth 5.0, NFC",
        "price": format_egp(310),  # 15,500 EGP
        "features": "120Hz AMOLED Screen, Ultra-Slim (7.4mm), Powerful Snapdragon 778G"
    },
    "Samsung Galaxy M54 5G": {
        "brand": "Samsung",
        "model": "Galaxy M54 5G",
        "release_date": "April 2023",
        "display": "6.7\" Super AMOLED Plus, 1080 x 2400 pixels, 120Hz, Gorilla Glass 5",
        "refresh_rate": "120Hz",
        "processor": "Exynos 1380 (5 nm)",
        "gpu": "Mali-G68 MP5",
        "ram": "8GB LPDDR4X",
        "storage": "128GB / 256GB (microSDXC)",
        "main_camera": "Triple: 108 MP (wide, OIS) + 8 MP (ultrawide) + 2 MP (macro)",
        "selfie_camera": "32 MP, 4K@30fps",
        "battery": "6000 mAh",
        "charging": "25W Wired",
        "os": "Android 13 (Upgradable to Android 15, One UI 7)",
        "weight": "199 g",
        "dimensions": "164.9 x 77.3 x 8.4 mm",
        "water_resistance": "Standard Splash Resistance",
        "network": "5G, Wi-Fi 6, Bluetooth 5.3, NFC",
        "price": format_egp(350),  # 17,500 EGP
        "features": "108MP Camera with OIS, Huge 6000mAh Battery, 120Hz AMOLED"
    },
    "Samsung Galaxy S25 Ultra": {
        "brand": "Samsung",
        "model": "Galaxy S25 Ultra",
        "release_date": "January 2025",
        "display": "6.86\" Dynamic LTPO AMOLED 2X, 1440 x 3120 pixels, 2600 nits, Gorilla Armor 2",
        "refresh_rate": "120Hz (LTPO 1-120Hz)",
        "processor": "Snapdragon 8 Elite for Galaxy (3 nm)",
        "gpu": "Adreno 830",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Quad: 200 MP (wide, OIS) + 50 MP (periscope 5x, OIS) + 10 MP (telephoto 3x) + 50 MP (ultrawide)",
        "selfie_camera": "12 MP, f/2.2, 4K@60fps",
        "battery": "5000 mAh",
        "charging": "45W Wired (65% in 30 min) + 15W Wireless + 4.5W Reverse Wireless",
        "os": "Android 15, One UI 7 (7 Major OS upgrades)",
        "weight": "218 g",
        "dimensions": "162.8 x 77.6 x 8.2 mm",
        "water_resistance": "IP68 (up to 1.5m for 30 min), Titanium Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, UWB, NFC",
        "price": format_egp(1299),  # 64,950 EGP
        "features": "Built-in S-Pen, Galaxy AI v2, Satellite connectivity, Anti-reflective screen"
    },
    "Samsung Galaxy S25 Plus": {
        "brand": "Samsung",
        "model": "Galaxy S25 Plus",
        "release_date": "January 2025",
        "display": "6.7\" Dynamic LTPO AMOLED 2X, 1440 x 3120 pixels, 2600 nits, Gorilla Glass Victus 2",
        "refresh_rate": "120Hz (LTPO 1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm) / Exynos 2500",
        "gpu": "Adreno 830 / Xclipse 950",
        "ram": "12GB LPDDR5X",
        "storage": "256GB / 512GB UFS 4.0",
        "main_camera": "Triple: 50 MP (wide, OIS) + 10 MP (telephoto 3x, OIS) + 12 MP (ultrawide)",
        "selfie_camera": "12 MP, f/2.2, 4K@60fps",
        "battery": "4900 mAh",
        "charging": "45W Wired + 15W Wireless + 4.5W Reverse Wireless",
        "os": "Android 15, One UI 7 (7 Major OS upgrades)",
        "weight": "190 g",
        "dimensions": "158.4 x 75.7 x 7.3 mm",
        "water_resistance": "IP68 (up to 1.5m for 30 min), Armor Aluminum Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, UWB, NFC",
        "price": format_egp(999),  # 49,950 EGP
        "features": "Galaxy AI, Slim profile (7.3mm), UWB support, 7 years updates"
    },
    "Samsung Galaxy S25": {
        "brand": "Samsung",
        "model": "Galaxy S25",
        "release_date": "January 2025",
        "display": "6.2\" Dynamic LTPO AMOLED 2X, 1080 x 2340 pixels, 2600 nits",
        "refresh_rate": "120Hz (LTPO 1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm) / Exynos 2500",
        "gpu": "Adreno 830 / Xclipse 950",
        "ram": "12GB LPDDR5X",
        "storage": "128GB / 256GB / 512GB UFS 4.0",
        "main_camera": "Triple: 50 MP (wide, OIS) + 10 MP (telephoto 3x) + 12 MP (ultrawide)",
        "selfie_camera": "12 MP, f/2.2, 4K@60fps",
        "battery": "4000 mAh",
        "charging": "25W Wired + 15W Wireless",
        "os": "Android 15, One UI 7",
        "weight": "162 g",
        "dimensions": "146.9 x 70.4 x 7.2 mm",
        "water_resistance": "IP68, Armor Aluminum",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC",
        "price": format_egp(799),  # 39,950 EGP
        "features": "Compact Flagship, Lightweight (162g), Galaxy AI suite"
    },
    "Samsung Galaxy S26 Plus": {
        "brand": "Samsung",
        "model": "Galaxy S26 Plus (Expected/Rumored)",
        "release_date": "Early 2026 (Expected)",
        "display": "6.7\" Dynamic AMOLED 2X Pro, 1440 x 3120 pixels, 3000 nits, Gorilla Armor",
        "refresh_rate": "144Hz / 120Hz LTPO",
        "processor": "Snapdragon 8 Elite Gen 2 / Exynos 2600 (2 nm)",
        "gpu": "Next-Gen Adreno / Xclipse",
        "ram": "12GB / 16GB LPDDR6",
        "storage": "256GB / 512GB UFS 4.1",
        "main_camera": "Triple: 50 MP (1\" sensor, OIS) + 50 MP (telephoto 3.5x) + 50 MP (ultrawide)",
        "selfie_camera": "16 MP Under-display / punch hole, 4K@60fps",
        "battery": "5100 mAh (Silicon-Carbon battery)",
        "charging": "65W Wired + 25W Qi2 MagSafe-compatible Wireless",
        "os": "Android 16, One UI 8",
        "weight": "188 g",
        "dimensions": "158.0 x 75.5 x 7.1 mm",
        "water_resistance": "IP68 / IP69, Titanium-Aluminum hybrid",
        "network": "5G Advanced, Wi-Fi 7, Bluetooth 5.5, UWB",
        "price": format_egp(1049),  # 52,450 EGP
        "features": "Silicon-Carbon High Density Battery, On-device Next-Gen AI, Qi2 Magnetic Charging"
    },
    "Samsung Galaxy S26 Ultra": {
        "brand": "Samsung",
        "model": "Galaxy S26 Ultra (Expected/Rumored)",
        "release_date": "Early 2026 (Expected)",
        "display": "6.9\" Dynamic LTPO AMOLED 2X, 1440 x 3200 pixels, 3500 nits, Anti-Reflective Armor 3",
        "refresh_rate": "144Hz LTPO (1-144Hz)",
        "processor": "Snapdragon 8 Elite Gen 2 (2 nm)",
        "gpu": "Adreno 840 Next-Gen",
        "ram": "16GB / 24GB LPDDR6",
        "storage": "256GB / 512GB / 1TB / 2TB UFS 4.1",
        "main_camera": "Quad: 200 MP (1/1.1\" OIS) + 50 MP (periscope 10x) + 50 MP (telephoto 3x) + 50 MP (ultrawide)",
        "selfie_camera": "24 MP, 4K@60fps Dolby Vision",
        "battery": "5500 mAh (Silicon-Carbon)",
        "charging": "65W Wired + 25W Wireless Qi2",
        "os": "Android 16, One UI 8",
        "weight": "215 g",
        "dimensions": "162.5 x 77.4 x 8.1 mm",
        "water_resistance": "IP68 & IP69, Grade 5 Titanium",
        "network": "5.5G Advanced, Wi-Fi 7, Bluetooth 5.5, UWB",
        "price": format_egp(1349),  # 67,450 EGP
        "features": "2nm Superchip, 144Hz Display, 200MP Next-Gen Sensor, Qi2 Magnetic Ring"
    },
    "Samsung Galaxy S24 Ultra": {
        "brand": "Samsung",
        "model": "Galaxy S24 Ultra",
        "release_date": "January 2024",
        "display": "6.8\" Dynamic LTPO AMOLED 2X, 1440 x 3120 pixels, 2600 nits, Gorilla Armor",
        "refresh_rate": "120Hz (LTPO 1-120Hz)",
        "processor": "Snapdragon 8 Gen 3 for Galaxy (4 nm)",
        "gpu": "Adreno 750 (1 GHz)",
        "ram": "12GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Quad: 200 MP (wide, OIS) + 50 MP (periscope 5x) + 10 MP (telephoto 3x) + 12 MP (ultrawide)",
        "selfie_camera": "12 MP, f/2.2",
        "battery": "5000 mAh",
        "charging": "45W Wired + 15W Wireless",
        "os": "Android 14, upgradable to Android 15, One UI 7",
        "weight": "232 g",
        "dimensions": "162.3 x 79.0 x 8.6 mm",
        "water_resistance": "IP68, Titanium Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, UWB",
        "price": format_egp(1099),  # 54,950 EGP
        "features": "Flat Screen, S-Pen, Anti-reflective coating, Galaxy AI"
    },
    "Samsung Galaxy A55 5G": {
        "brand": "Samsung",
        "model": "Galaxy A55 5G",
        "release_date": "March 2024",
        "display": "6.6\" Super AMOLED, 1080 x 2340 pixels, 1000 nits (HBM), Gorilla Glass Victus+",
        "refresh_rate": "120Hz",
        "processor": "Exynos 1480 (4 nm)",
        "gpu": "Xclipse 530 (AMD RDNA2)",
        "ram": "8GB / 12GB",
        "storage": "128GB / 256GB (microSDXC slot)",
        "main_camera": "Triple: 50 MP (wide, OIS) + 12 MP (ultrawide) + 5 MP (macro)",
        "selfie_camera": "32 MP, f/2.2, 4K@30fps",
        "battery": "5000 mAh",
        "charging": "25W Wired",
        "os": "Android 14, One UI 6.1 (4 OS updates)",
        "weight": "213 g",
        "dimensions": "161.1 x 77.4 x 8.2 mm",
        "water_resistance": "IP67, Aluminum Frame + Glass Back",
        "network": "5G, Wi-Fi 6, Bluetooth 5.3, NFC",
        "price": format_egp(380),  # 19,000 EGP
        "features": "Premium Aluminum Build, Knox Vault Security, AMD GPU in Mid-range"
    },

    # --- Apple iPhone Series ---
    "iPhone 17 Pro Max": {
        "brand": "Apple",
        "model": "iPhone 17 Pro Max (Expected/Rumored)",
        "release_date": "September 2025 (Expected)",
        "display": "6.9\" LTPO Super Retina XDR OLED, 1320 x 2868 pixels, 3000 nits, Anti-scratch Ceramic Shield",
        "refresh_rate": "120Hz ProMotion (1-120Hz)",
        "processor": "Apple A19 Pro (2 nm TSMC N2)",
        "gpu": "Apple 6-core GPU with Ray Tracing & Neural Accelerators",
        "ram": "12GB Unified Memory",
        "storage": "256GB / 512GB / 1TB / 2TB NVMe",
        "main_camera": "Triple 48 MP: 48 MP (wide, sensor-shift OIS) + 48 MP (periscope 5x/10x optical) + 48 MP (ultrawide)",
        "selfie_camera": "24 MP, Center Stage, 4K@60fps Dolby Vision",
        "battery": "4850 mAh",
        "charging": "40W Wired + 25W MagSafe Wireless + Qi2",
        "os": "iOS 19 (Apple Intelligence Enhanced)",
        "weight": "225 g",
        "dimensions": "163.0 x 77.5 x 8.2 mm",
        "water_resistance": "IP68 (6m for 30 mins), Grade 5 Titanium",
        "network": "5G (Sub-6 & mmWave), Wi-Fi 7, Bluetooth 5.4, UWB 2nd Gen, Satellite SOS",
        "price": format_egp(1299),  # 64,950 EGP
        "features": "Apple Intelligence 2.0, All 48MP Rear Cameras, Smaller Dynamic Island, 2nm Chipset"
    },
    "iPhone 17 Pro": {
        "brand": "Apple",
        "model": "iPhone 17 Pro (Expected/Rumored)",
        "release_date": "September 2025 (Expected)",
        "display": "6.3\" LTPO Super Retina XDR OLED, 1206 x 2622 pixels, 3000 nits",
        "refresh_rate": "120Hz ProMotion",
        "processor": "Apple A19 Pro (2 nm)",
        "gpu": "Apple 6-core GPU",
        "ram": "12GB Unified Memory",
        "storage": "256GB / 512GB / 1TB NVMe",
        "main_camera": "Triple 48 MP: 48 MP (wide) + 48 MP (5x telephoto) + 48 MP (ultrawide)",
        "selfie_camera": "24 MP, 4K@60fps",
        "battery": "3800 mAh",
        "charging": "35W Wired + 25W MagSafe",
        "os": "iOS 19",
        "weight": "197 g",
        "dimensions": "149.6 x 71.5 x 8.1 mm",
        "water_resistance": "IP68, Titanium Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, UWB",
        "price": format_egp(1099),  # 54,950 EGP
        "features": "Compact 2nm Flagship, 48MP across all lenses, 12GB RAM for on-device AI"
    },
    "iPhone 16 Pro Max": {
        "brand": "Apple",
        "model": "iPhone 16 Pro Max",
        "release_date": "September 2024",
        "display": "6.9\" LTPO Super Retina XDR OLED, 1320 x 2868 pixels, 2000 nits, Ceramic Shield 2024",
        "refresh_rate": "120Hz ProMotion (1-120Hz)",
        "processor": "Apple A18 Pro (3 nm N3E)",
        "gpu": "Apple 6-core GPU with Hardware Ray Tracing",
        "ram": "8GB Unified Memory",
        "storage": "256GB / 512GB / 1TB NVMe",
        "main_camera": "Triple: 48 MP (wide, sensor-shift OIS) + 12 MP (periscope 5x, OIS) + 48 MP (ultrawide)",
        "selfie_camera": "12 MP, f/1.9, 4K@60fps Dolby Vision",
        "battery": "4685 mAh",
        "charging": "30W Wired (50% in 30 min) + 25W MagSafe + 15W Qi2",
        "os": "iOS 18, upgradable to latest iOS",
        "weight": "227 g",
        "dimensions": "163.0 x 77.6 x 8.25 mm",
        "water_resistance": "IP68 (6m for 30 mins), Grade 5 Titanium",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, UWB 2nd Gen, Satellite Emergency",
        "price": format_egp(1199),  # 59,950 EGP
        "features": "Camera Control Button, Apple Intelligence, 4K@120fps video recording, Ultra-thin bezels"
    },
    "iPhone 16 Pro": {
        "brand": "Apple",
        "model": "iPhone 16 Pro",
        "release_date": "September 2024",
        "display": "6.3\" LTPO Super Retina XDR OLED, 1206 x 2622 pixels, 2000 nits",
        "refresh_rate": "120Hz ProMotion",
        "processor": "Apple A18 Pro (3 nm)",
        "gpu": "Apple 6-core GPU",
        "ram": "8GB Unified Memory",
        "storage": "128GB / 256GB / 512GB / 1TB NVMe",
        "main_camera": "Triple: 48 MP (wide) + 12 MP (5x telephoto) + 48 MP (ultrawide)",
        "selfie_camera": "12 MP, f/1.9",
        "battery": "3582 mAh",
        "charging": "27W Wired + 25W MagSafe",
        "os": "iOS 18",
        "weight": "199 g",
        "dimensions": "149.6 x 71.5 x 8.25 mm",
        "water_resistance": "IP68, Grade 5 Titanium",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, UWB",
        "price": format_egp(999),  # 49,950 EGP
        "features": "Camera Control, 5x Optical Zoom in smaller size, Apple Intelligence"
    },
    "iPhone 16": {
        "brand": "Apple",
        "model": "iPhone 16",
        "release_date": "September 2024",
        "display": "6.1\" Super Retina XDR OLED, 1179 x 2556 pixels, 2000 nits",
        "refresh_rate": "60Hz",
        "processor": "Apple A18 (3 nm)",
        "gpu": "Apple 5-core GPU",
        "ram": "8GB Unified Memory",
        "storage": "128GB / 256GB / 512GB NVMe",
        "main_camera": "Dual: 48 MP (wide, sensor-shift OIS) + 12 MP (ultrawide with Macro)",
        "selfie_camera": "12 MP, f/1.9",
        "battery": "3561 mAh",
        "charging": "25W Wired + 25W MagSafe",
        "os": "iOS 18",
        "weight": "170 g",
        "dimensions": "147.6 x 71.6 x 7.8 mm",
        "water_resistance": "IP68, Aluminum Frame + Color-infused Glass",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, UWB",
        "price": format_egp(799),  # 39,950 EGP
        "features": "Action Button, Camera Control Button, Spatial Video capture, Apple Intelligence"
    },
    "iPhone 15 Pro Max": {
        "brand": "Apple",
        "model": "iPhone 15 Pro Max",
        "release_date": "September 2023",
        "display": "6.7\" LTPO Super Retina XDR OLED, 1290 x 2796 pixels, 2000 nits",
        "refresh_rate": "120Hz ProMotion",
        "processor": "Apple A17 Pro (3 nm)",
        "gpu": "Apple 6-core GPU",
        "ram": "8GB Unified Memory",
        "storage": "256GB / 512GB / 1TB NVMe",
        "main_camera": "Triple: 48 MP (wide) + 12 MP (periscope 5x) + 12 MP (ultrawide)",
        "selfie_camera": "12 MP, f/1.9",
        "battery": "4441 mAh",
        "charging": "25W Wired + 15W MagSafe",
        "os": "iOS 17, upgradable to iOS 18",
        "weight": "221 g",
        "dimensions": "159.9 x 76.7 x 8.25 mm",
        "water_resistance": "IP68, Titanium Frame, USB Type-C 3.0",
        "network": "5G, Wi-Fi 6E, Bluetooth 5.3, UWB",
        "price": format_egp(999),  # 49,950 EGP
        "features": "First Titanium iPhone, USB-C 10Gbps, Action Button, Apple Intelligence support"
    },

    # --- Xiaomi & Poco ---
    "Xiaomi 15 Ultra": {
        "brand": "Xiaomi",
        "model": "Xiaomi 15 Ultra",
        "release_date": "February 2025",
        "display": "6.73\" LTPO AMOLED, 1440 x 3200 pixels, 3200 nits, Dolby Vision, Shield Glass",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm)",
        "gpu": "Adreno 830",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Leica Quad: 50 MP (1\" LYT-900, OIS) + 200 MP (periscope 4.3x, OIS) + 50 MP (telephoto 3x) + 50 MP (ultrawide)",
        "selfie_camera": "32 MP, f/2.0, 4K@60fps",
        "battery": "6000 mAh (Silicon-Carbon)",
        "charging": "90W Wired + 80W Wireless + 10W Reverse Wireless",
        "os": "Android 15, HyperOS 2",
        "weight": "226 g",
        "dimensions": "161.4 x 75.3 x 9.2 mm",
        "water_resistance": "IP68 / IP69, Ceramic/Eco Leather Back",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC, IR Blaster",
        "price": format_egp(1150),  # 57,500 EGP
        "features": "Leica Optics with 200MP Periscope, Huge 6000mAh Battery, 90W fast charging"
    },
    "Xiaomi 15 Pro": {
        "brand": "Xiaomi",
        "model": "Xiaomi 15 Pro",
        "release_date": "October 2024",
        "display": "6.73\" LTPO AMOLED, 1440 x 3200 pixels, 3200 nits, Dragon Crystal Glass 2",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm)",
        "gpu": "Adreno 830",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Leica Triple: 50 MP (Light Hunter 900, OIS) + 50 MP (periscope 5x, OIS) + 50 MP (ultrawide)",
        "selfie_camera": "32 MP, 4K@60fps",
        "battery": "6100 mAh (High Energy Density)",
        "charging": "90W Wired + 50W Wireless",
        "os": "Android 15, HyperOS 2",
        "weight": "213 g",
        "dimensions": "161.3 x 75.3 x 8.35 mm",
        "water_resistance": "IP68, Aluminum Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC, IR Blaster",
        "price": format_egp(850),  # 42,500 EGP
        "features": "Massive 6100mAh in slim body (8.35mm), Leica 5x Periscope, HyperOS 2 AI"
    },
    "Poco X7 Pro 5G": {
        "brand": "Xiaomi / Poco",
        "model": "Poco X7 Pro 5G",
        "release_date": "January 2025",
        "display": "6.67\" Flow AMOLED, 1220 x 2712 pixels (1.5K), 2400 nits, Gorilla Glass Victus",
        "refresh_rate": "120Hz",
        "processor": "MediaTek Dimensity 8400-Ultra (4 nm)",
        "gpu": "Mali-G720",
        "ram": "8GB / 12GB LPDDR5X",
        "storage": "256GB / 512GB UFS 4.0",
        "main_camera": "Triple: 50 MP (Sony LYT-600, OIS) + 8 MP (ultrawide) + 2 MP (macro)",
        "selfie_camera": "20 MP, 1080p@60fps",
        "battery": "6000 mAh",
        "charging": "90W HyperCharge (100% in 35 min)",
        "os": "Android 15, HyperOS 2",
        "weight": "196 g",
        "dimensions": "160.5 x 74.3 x 8.1 mm",
        "water_resistance": "IP68, Dual stereo speakers",
        "network": "5G, Wi-Fi 6, Bluetooth 5.4, NFC, IR Blaster",
        "price": format_egp(350),  # 17,500 EGP
        "features": "Flagship Killer Value, 90W Charging, 6000mAh Battery, Dimensity 8400-Ultra"
    },
    "Poco F6 Pro": {
        "brand": "Xiaomi / Poco",
        "model": "Poco F6 Pro",
        "release_date": "May 2024",
        "display": "6.67\" AMOLED WQHD+ 1440 x 3200, 4000 nits peak, 120Hz",
        "refresh_rate": "120Hz",
        "processor": "Snapdragon 8 Gen 2 (4 nm)",
        "gpu": "Adreno 740",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Triple: 50 MP (Light Fusion 800, OIS) + 8 MP (ultrawide) + 2 MP (macro)",
        "selfie_camera": "16 MP, 1080p@60fps",
        "battery": "5000 mAh",
        "charging": "120W HyperCharge (100% in 19 min)",
        "os": "Android 14, HyperOS",
        "weight": "209 g",
        "dimensions": "160.9 x 75.0 x 8.2 mm",
        "water_resistance": "IP54, Glass Back + Metal Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, NFC, IR Blaster",
        "price": format_egp(499),  # 24,950 EGP
        "features": "Ultra-fast 120W Charging (0-100% in 19 min), 2K 4000nits Screen"
    },

    # --- OPPO Series ---
    "OPPO Find X8 Pro": {
        "brand": "OPPO",
        "model": "Find X8 Pro",
        "release_date": "October 2024",
        "display": "6.78\" LTPO AMOLED, 1264 x 2780 pixels, 4500 nits peak, Gorilla Glass Victus 2",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "MediaTek Dimensity 9400 (3 nm)",
        "gpu": "Immortalis-G925",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Hasselblad Quad: 50 MP (wide, OIS) + 50 MP (periscope 3x, OIS) + 50 MP (periscope 6x, OIS) + 50 MP (ultrawide)",
        "selfie_camera": "32 MP, f/2.4, 4K@60fps",
        "battery": "5910 mAh (Glacier Silicon-Carbon)",
        "charging": "80W SuperVOOC + 50W AirVOOC Wireless",
        "os": "Android 15, ColorOS 15",
        "weight": "215 g",
        "dimensions": "162.3 x 76.7 x 8.2 mm",
        "water_resistance": "IP68 & IP69 (Highest Water and Dust Rating)",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC, Satellite SOS",
        "price": format_egp(999),  # 49,950 EGP
        "features": "Dual Periscope Cameras (3x & 6x), Quick Button, 5910mAh Glacier Battery, IP69"
    },
    "OPPO Reno 12 Pro 5G": {
        "brand": "OPPO",
        "model": "Reno 12 Pro 5G",
        "release_date": "June 2024",
        "display": "6.7\" Quad-Curved AMOLED, 1080 x 2412 pixels, 120Hz, Gorilla Glass Victus 2",
        "refresh_rate": "120Hz",
        "processor": "MediaTek Dimensity 7300-Energy (4 nm)",
        "gpu": "Mali-G615 MC2",
        "ram": "12GB LPDDR4X",
        "storage": "256GB / 512GB UFS 3.1",
        "main_camera": "Triple: 50 MP (Sony LYT-600, OIS) + 50 MP (telephoto 2x) + 8 MP (ultrawide)",
        "selfie_camera": "50 MP with autofocus, 4K@60fps",
        "battery": "5000 mAh",
        "charging": "80W SuperVOOC (100% in 46 min)",
        "os": "Android 14, ColorOS 14.1 (AI Portrait Expert)",
        "weight": "180 g",
        "dimensions": "161.5 x 74.8 x 7.4 mm",
        "water_resistance": "IP65 (Splashproof and Drop Resistant)",
        "network": "5G, Wi-Fi 6, Bluetooth 5.4, NFC",
        "price": format_egp(450),  # 22,500 EGP
        "features": "50MP AF Selfie, 50MP Telephoto Portrait, Ultra-Slim (7.4mm), GenAI Photo Eraser"
    },

    # --- Google Pixel ---
    "Google Pixel 9 Pro XL": {
        "brand": "Google",
        "model": "Pixel 9 Pro XL",
        "release_date": "August 2024",
        "display": "6.8\" LTPO OLED (Super Actua), 1344 x 2992 pixels, 3000 nits, Gorilla Glass Victus 2",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "Google Tensor G4 (4 nm) + Titan M2 security",
        "gpu": "Mali-G715 MC7",
        "ram": "16GB LPDDR5X",
        "storage": "128GB / 256GB / 512GB / 1TB UFS 3.1",
        "main_camera": "Triple: 50 MP (wide, OIS) + 48 MP (periscope 5x, OIS) + 48 MP (ultrawide)",
        "selfie_camera": "42 MP, f/2.2, ultrawide with autofocus, 4K@60fps",
        "battery": "5060 mAh",
        "charging": "37W Wired (70% in 30 min) + 23W Wireless (Pixel Stand)",
        "os": "Android 14 (Upgradable to Android 15), 7 years of OS & security updates",
        "weight": "221 g",
        "dimensions": "162.8 x 76.6 x 8.5 mm",
        "water_resistance": "IP68, Polished Aluminum Frame",
        "network": "5G, Wi-Fi 7, Bluetooth 5.3, UWB, Satellite SOS",
        "price": format_egp(1099),  # 54,950 EGP
        "features": "Gemini Live & Advanced AI, 42MP Selfie with AF, Temperature sensor, 7-year support"
    },

    # --- OnePlus ---
    "OnePlus 13": {
        "brand": "OnePlus",
        "model": "OnePlus 13",
        "release_date": "October 2024 / January 2025 Global",
        "display": "6.82\" LTPO AMOLED (BOE X2), 1440 x 3168 pixels (2K), 4500 nits peak, Crystal Shield",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm)",
        "gpu": "Adreno 830",
        "ram": "12GB / 16GB / 24GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Hasselblad Triple 50 MP: 50 MP (Sony LYT-808, OIS) + 50 MP (periscope 3x, OIS) + 50 MP (ultrawide)",
        "selfie_camera": "32 MP, 4K@60fps",
        "battery": "6000 mAh (Silicon-Carbon Glacier Battery)",
        "charging": "100W SuperVOOC Wired (100% in 36 min) + 50W AIRVOOC Wireless + Magnetic support",
        "os": "Android 15, OxygenOS 15 / ColorOS 15",
        "weight": "210 g",
        "dimensions": "162.9 x 76.5 x 8.5 mm",
        "water_resistance": "IP68 & IP69 (Hot water & high pressure jet resistant)",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC, IR Blaster",
        "price": format_egp(899),  # 44,950 EGP
        "features": "IP69 rating, 6000mAh battery + 100W, Hasselblad Color Tuning, Alert Slider"
    },

    # --- Honor ---
    "Honor Magic 7 Pro": {
        "brand": "Honor",
        "model": "Magic 7 Pro",
        "release_date": "October 2024 / Early 2025 Global",
        "display": "6.8\" LTPO OLED, 1280 x 2800 pixels, 5000 nits HDR, 4320Hz PWM dimming, Giant Rhinoceros Glass",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "Snapdragon 8 Elite (3 nm)",
        "gpu": "Adreno 830",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "Triple: 50 MP (OmniVision 1/1.3\", variable aperture f/1.4-f/2.0, OIS) + 200 MP (periscope 3x, OIS) + 50 MP (ultrawide)",
        "selfie_camera": "50 MP + 3D ToF (Face Unlock), 4K@60fps",
        "battery": "5850 mAh (3rd Gen Silicon-Carbon)",
        "charging": "100W Wired + 80W Wireless",
        "os": "Android 15, MagicOS 9.0 (AI Agent)",
        "weight": "223 g",
        "dimensions": "162.7 x 77.1 x 8.8 mm",
        "water_resistance": "IP68 & IP69",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, UWB, Satellite messaging",
        "price": format_egp(950),  # 47,500 EGP
        "features": "200MP Periscope Camera, 3D Face Unlock + Ultrasonic Fingerprint, Eye-care 4320Hz PWM"
    },

    # --- Vivo ---
    "Vivo X200 Pro": {
        "brand": "Vivo",
        "model": "Vivo X200 Pro",
        "release_date": "October 2024",
        "display": "6.78\" LTPO AMOLED, 1260 x 2800 pixels, 4500 nits, 2160Hz PWM, Armor Glass",
        "refresh_rate": "120Hz (1-120Hz)",
        "processor": "MediaTek Dimensity 9400 (3 nm)",
        "gpu": "Immortalis-G925",
        "ram": "12GB / 16GB LPDDR5X",
        "storage": "256GB / 512GB / 1TB UFS 4.0",
        "main_camera": "ZEISS: 50 MP (Sony LYT-818 1/1.28\", OIS) + 200 MP (ZEISS APO periscope 3.7x, OIS) + 50 MP (ultrawide)",
        "selfie_camera": "32 MP, f/2.0, 4K@60fps",
        "battery": "6000 mAh (BlueOcean Silicon Battery)",
        "charging": "90W FlashCharge + 30W Wireless",
        "os": "Android 15, OriginOS 5 / Funtouch OS 15",
        "weight": "228 g",
        "dimensions": "162.4 x 76.0 x 8.2 mm",
        "water_resistance": "IP68 & IP69",
        "network": "5G, Wi-Fi 7, Bluetooth 5.4, NFC",
        "price": format_egp(899),  # 44,950 EGP
        "features": "ZEISS 200MP APO Telephoto (Best-in-class zoom), Dimensity 9400 power, 6000mAh"
    }
}

# قائمة موحدة لجميع الأجهزة
ALL_PHONE_CATALOG = []
for brand_name, models_list in BRANDS_CATALOG.items():
    for model in models_list:
        ALL_PHONE_CATALOG.append(model)


def normalize_string(s: str) -> str:
    """تنظيف وتوحيد نصوص البحث لتسهيل المطابقة"""
    s = s.lower().strip()
    s = re.sub(r'[\-\_\+\/\,\.]', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.replace("samsung galaxy", "samsung")
    s = s.replace("galaxy", "samsung")
    s = s.replace("apple iphone", "iphone")
    s = s.replace("promax", "pro max")
    s = s.replace("plus", "+")
    return s.strip()


def search_phone_suggestions(query: str, limit: int = 6) -> List[str]:
    """محرك البحث والاقتراح التلقائي السريع (Autocomplete Suggestions)"""
    if not query or not query.strip():
        return []

    q_clean = query.strip()
    q_lower = q_clean.lower()
    q_norm = normalize_string(q_clean)
    q_tokens = [t for t in q_norm.split() if t]

    matches_with_score = []
    combined_catalog = list(dict.fromkeys(ALL_PHONE_CATALOG + list(MOBILE_DATABASE.keys())))

    for phone_name in combined_catalog:
        name_lower = phone_name.lower()
        name_norm = normalize_string(phone_name)
        score = 0

        if q_lower == name_lower:
            score += 100
        elif name_lower.startswith(q_lower) or name_norm.startswith(q_norm):
            score += 60
        elif q_lower in name_lower or q_norm in name_norm:
            score += 40

        name_tokens = set(name_norm.split())
        matched_tokens = 0
        for token in q_tokens:
            if token in name_tokens:
                score += 15
                matched_tokens += 1
                if any(c.isdigit() for c in token):
                    score += 25
                if token in ["ultra", "plus", "+", "max", "pro", "fe", "fold", "flip", "edge", "air"]:
                    score += 15
            elif any(token in nt for nt in name_tokens):
                score += 8
                matched_tokens += 1

        if score > 0 or (matched_tokens == len(q_tokens) and len(q_tokens) > 0):
            matches_with_score.append((phone_name, score))

    matches_with_score.sort(key=lambda x: x[1], reverse=True)
    return [name for name, sc in matches_with_score[:limit]]


def find_in_database(query: str) -> Optional[Dict[str, Any]]:
    """البحث الذكي داخل قاعدة البيانات المدمجة"""
    norm_query = normalize_string(query)
    if not norm_query:
        return None

    for db_name, specs in MOBILE_DATABASE.items():
        norm_db_name = normalize_string(db_name)
        if norm_query == norm_db_name:
            return specs.copy()

    query_tokens = set(norm_query.split())
    best_match = None
    best_score = 0

    for db_name, specs in MOBILE_DATABASE.items():
        norm_db_name = normalize_string(db_name)
        db_tokens = set(norm_db_name.split())
        common = query_tokens.intersection(db_tokens)
        score = len(common)

        for token in query_tokens:
            if token in norm_db_name:
                score += 1
                if token in ["ultra", "plus", "+", "max", "pro", "fe", "fold", "flip", "m51", "m52", "m54"]:
                    score += 3
                if any(char.isdigit() for char in token):
                    score += 3

        if score > best_score and len(common) >= 1:
            best_score = score
            best_match = specs.copy()

    if best_match and best_score >= 3:
        return best_match

    return None


def generate_estimated_specs(query: str) -> Dict[str, Any]:
    """توليد مواصفات ذكية واستنتاجية مع تحويل السعر بالجنيه المصري (1$ = 50 ج)"""
    clean_name = query.strip()
    norm = normalize_string(query)

    brand = "Unknown"
    if "samsung" in norm or "galaxy" in norm or norm.startswith("s") or norm.startswith("m") or norm.startswith("a"):
        brand = "Samsung"
    elif "iphone" in norm or "apple" in norm:
        brand = "Apple"
    elif "xiaomi" in norm or "redmi" in norm or "poco" in norm:
        brand = "Xiaomi"
    elif "pixel" in norm or "google" in norm:
        brand = "Google"
    elif "oneplus" in norm:
        brand = "OnePlus"
    elif "honor" in norm or "huawei" in norm:
        brand = "Honor / Huawei"
    elif "oppo" in norm or "reno" in norm or "find" in norm:
        brand = "OPPO"
    elif "vivo" in norm or "iqoo" in norm:
        brand = "Vivo"
    elif "realme" in norm:
        brand = "Realme"
    elif "infinix" in norm:
        brand = "Infinix"
    elif "tecno" in norm:
        brand = "Tecno"

    is_flagship = any(w in norm for w in ["ultra", "pro", "max", "plus", "+", "fold", "s24", "s25", "s26", "15", "16", "17", "find", "magic", "x200", "gt 7"])
    is_apple = brand == "Apple"

    processor = "Next-Gen Flagship Processor (3nm / 2nm)" if is_flagship else "High-Efficiency Octa-Core Processor"
    if is_apple:
        processor = "Apple A-Series Bionic / Pro Chip"
    elif "samsung" in norm:
        processor = "Snapdragon / Exynos High-Performance Chipset"
    elif "oppo" in norm or "vivo" in norm:
        processor = "MediaTek Dimensity / Snapdragon Performance Chip"

    price_str = format_egp(1099 if is_flagship else 380)

    return {
        "brand": brand,
        "model": clean_name,
        "release_date": "Latest Generation / Announced",
        "display": "6.7\" Dynamic AMOLED / OLED 120Hz HDR10+" if is_flagship else "6.6\" FHD+ AMOLED 120Hz/90Hz",
        "refresh_rate": "120Hz LTPO (Adaptive)" if is_flagship else "120Hz / 90Hz",
        "processor": processor,
        "gpu": "High-Performance GPU",
        "ram": ("12GB / 16GB LPDDR5X" if not is_apple else "8GB / 12GB Unified") if is_flagship else "8GB LPDDR4X",
        "storage": "256GB / 512GB / 1TB UFS 4.0" if is_flagship else "128GB / 256GB",
        "main_camera": "Triple / Quad Camera (50MP/200MP Main + Telephoto + Ultrawide, OIS)" if is_flagship else "50MP Main + Ultrawide",
        "selfie_camera": "12MP / 32MP, 4K Video support",
        "battery": "5000 mAh - 6000 mAh (Silicon-Carbon)" if not is_apple else "4500 mAh - 4850 mAh",
        "charging": "45W - 100W Fast Charging + Wireless" if not is_apple else "30W Fast Wired + MagSafe",
        "os": ("iOS (Latest Version)" if is_apple else "Android (Latest Version)"),
        "weight": "~ 190 g - 215 g",
        "dimensions": "~ 162 x 76 x 8.0 mm",
        "water_resistance": "IP68 Water and Dust Resistant" if is_flagship else "IP54 / Splash Resistant",
        "network": "5G, Wi-Fi 7 / 6E, Bluetooth 5.4, NFC",
        "price": price_str,
        "features": f"AI Assistant Integration, HDR Display, High-Resolution Camera System ({clean_name})"
    }


def get_mobile_specs(query: str) -> Dict[str, Any]:
    """الدالة الرئيسية لجلب مواصفات أي هاتف ذكي"""
    if not query or not query.strip():
        return {}

    specs = find_in_database(query)
    if specs:
        return specs

    return generate_estimated_specs(query)


def get_popular_mobiles_list() -> List[str]:
    """قائمة بأشهر الهواتف للاقتراح السريع"""
    return ALL_PHONE_CATALOG
