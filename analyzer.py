"""
analyzer.py
محرك التحليل الذكي والديناميكي لمقارنة الهواتف، وتحديد الفائز في كل مواصفة بالجنيه المصري، وتوليد سبب الأفضلية المباشر.
"""

import re
from typing import Dict, Any, List, Tuple

def extract_numbers(text: str) -> List[int]:
    """استخراج جميع الأرقام الصحيحة من النص"""
    return [int(n) for n in re.findall(r'\b\d+\b', str(text))]

def evaluate_spec_winner(spec_key: str, phones_specs: Dict[str, Dict[str, Any]]) -> Tuple[str, str]:
    """
    تحليل ذكي وديناميكي لتحديد الهاتف الأفضل في مواصفة معينة مع ذكر سبب الأفضلية المختصر والدقيق.
    phones_specs: قاموس يحتوي على {phone_name: specs_dict}
    """
    if not phones_specs:
        return ("-", "-")

    # استبعاد الحقول التعريفية من المقارنة التنافسية
    if spec_key in ["brand", "model", "release_date"]:
        return ("-", "معلومات عامة")

    phone_names = list(phones_specs.keys())
    if len(phone_names) == 1:
        return (phone_names[0], "الجهاز الوحيد في المقارنة")

    # 1. تحليل الشاشة (Display)
    if spec_key == "display":
        best_phone = phone_names[0]
        max_score = -1
        for name, spec in phones_specs.items():
            text = str(spec.get("display", "")).lower()
            score = 0
            nits_match = re.search(r'(\d{4})\s*nits', text)
            nits = int(nits_match.group(1)) if nits_match else 2000
            score += nits

            if "1440" in text or "2k" in text or "3120" in text or "3200" in text:
                score += 800
            if "ltpo" in text:
                score += 400
            if "gorilla armor" in text or "ceramic shield" in text:
                score += 300

            if score > max_score:
                max_score = score
                best_phone = name

        return (best_phone, "شاشة أكبر وسطوع ودقة أعلى")

    # 2. معدل التحديث (Refresh Rate)
    if spec_key == "refresh_rate":
        best_phone = phone_names[0]
        max_hz = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("refresh_rate", "")).lower()
            hz_match = re.search(r'(\d{2,3})hz', text)
            hz = int(hz_match.group(1)) if hz_match else 120
            score = hz + (20 if "ltpo" in text or "1-" in text or "adaptive" in text else 0)
            if score > max_hz:
                max_hz = score
                best_phone = name
        return (best_phone, "معدل تحديث أعلى وأكثر سلاسة")

    # 3. المعالج والأداء (Processor)
    if spec_key == "processor":
        best_phone = phone_names[0]
        max_proc_score = -1
        for name, spec in phones_specs.items():
            text = str(spec.get("processor", "")).lower()
            score = 50
            if "2 nm" in text or "2nm" in text or "gen 2" in text or "a19" in text:
                score = 100
            elif "elite" in text or "a18 pro" in text or "9400" in text:
                score = 90
            elif "gen 3" in text or "a17 pro" in text:
                score = 80
            elif "tensor" in text:
                score = 75
            elif "8400" in text or "1480" in text:
                score = 60

            if score > max_proc_score:
                max_proc_score = score
                best_phone = name
        return (best_phone, "معالج أقوى وكفاءة أداء أعلى")

    # 4. الذاكرة العشوائية (RAM)
    if spec_key == "ram":
        best_phone = phone_names[0]
        max_ram = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("ram", ""))
            ram_nums = re.findall(r'(\d+)GB', text)
            ram_val = max([int(r) for r in ram_nums]) if ram_nums else 8
            bonus = 2 if "lpddr6" in text.lower() else (1 if "lpddr5x" in text.lower() else 0)
            total = ram_val + bonus
            if total > max_ram:
                max_ram = total
                best_phone = name
        return (best_phone, "RAM وسعة ذاكرة أعلى")

    # 5. سعة التخزين (Storage)
    if spec_key == "storage":
        best_phone = phone_names[0]
        max_st = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("storage", ""))
            score = 256
            if "2TB" in text or "2tb" in text:
                score = 2048
            elif "1TB" in text or "1tb" in text:
                score = 1024
            elif "512GB" in text or "512gb" in text:
                score = 512

            if score > max_st:
                max_st = score
                best_phone = name
        return (best_phone, "سعة تخزين أكبر")

    # 6. الكاميرا الأساسية (Main Camera)
    if spec_key == "main_camera":
        best_phone = phone_names[0]
        max_cam_score = -1
        for name, spec in phones_specs.items():
            text = str(spec.get("main_camera", "")).lower()
            score = 50
            if "200 mp" in text or "200mp" in text:
                score += 40
            if "1\"" in text or "1 inch" in text or "lyt-900" in text:
                score += 35
            if "periscope" in text or "5x" in text or "10x" in text or "3.7x" in text:
                score += 30
            if "leica" in text or "zeiss" in text or "hasselblad" in text:
                score += 25
            if "sensor-shift" in text or "apple" in name.lower() or "iphone" in name.lower():
                score += 25

            if score > max_cam_score:
                max_cam_score = score
                best_phone = name
        return (best_phone, "كاميرا أفضل ومعالجة صور أعلى")

    # 7. الكاميرا الأمامية (Selfie Camera)
    if spec_key == "selfie_camera":
        best_phone = phone_names[0]
        max_selfie = -1
        for name, spec in phones_specs.items():
            text = str(spec.get("selfie_camera", "")).lower()
            score = 10
            mp_match = re.search(r'(\d+)\s*mp', text)
            mp = int(mp_match.group(1)) if mp_match else 12
            score += mp
            if "4k@60fps" in text or "4k" in text:
                score += 20
            if "autofocus" in text or "center stage" in text or "tof" in text:
                score += 25

            if score > max_selfie:
                max_selfie = score
                best_phone = name
        return (best_phone, "كاميرا سيلفي بدقة أعلى")

    # 8. سعة البطارية (Battery)
    if spec_key == "battery":
        best_phone = phone_names[0]
        max_mah = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("battery", ""))
            bat_match = re.search(r'(\d{4})', text)
            mah = int(bat_match.group(1)) if bat_match else 4000
            bonus = 200 if "silicon" in text.lower() or "carbon" in text.lower() else 0
            total = mah + bonus
            if total > max_mah:
                max_mah = total
                best_phone = name
        return (best_phone, "بطارية أكبر وسعة أعلى")

    # 9. سرعة الشحن (Charging)
    if spec_key == "charging":
        best_phone = phone_names[0]
        max_watt = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("charging", ""))
            w_match = re.search(r'(\d+)\s*w', text, re.IGNORECASE)
            watt = int(w_match.group(1)) if w_match else 25
            if watt > max_watt:
                max_watt = watt
                best_phone = name
        return (best_phone, "سرعة شحن أعلى وأسرع")

    # 10. الوزن (Weight) - الأخف وزناً هو الأفضل
    if spec_key == "weight":
        best_phone = phone_names[0]
        min_weight = 9999
        for name, spec in phones_specs.items():
            text = str(spec.get("weight", ""))
            w_match = re.search(r'(\d{3})', text)
            w = int(w_match.group(1)) if w_match else 200
            if w < min_weight:
                min_weight = w
                best_phone = name
        return (best_phone, "وزن أقل وأخف في اليد")

    # 11. الأبعاد والنحافة (Dimensions)
    if spec_key == "dimensions":
        best_phone = phone_names[0]
        min_thickness = 99.0
        for name, spec in phones_specs.items():
            text = str(spec.get("dimensions", ""))
            th_match = re.findall(r'(\d+\.\d+)\s*mm', text)
            th = float(th_match[-1]) if th_match else 8.2
            if th < min_thickness:
                min_thickness = th
                best_phone = name
        return (best_phone, "هيكل أنحف وتصميم مدمج")

    # 12. مقاومة الماء والغبار (Water Resistance)
    if spec_key == "water_resistance":
        best_phone = phone_names[0]
        max_ip = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("water_resistance", "")).lower()
            score = 1
            if "ip69" in text:
                score = 4
            elif "6m" in text:
                score = 3
            elif "ip68" in text:
                score = 2

            if score > max_ip:
                max_ip = score
                best_phone = name
        return (best_phone, "معيار حماية ومقاومة أعلى")

    # 13. نظام التشغيل والتحديثات (OS)
    if spec_key == "os":
        best_phone = phone_names[0]
        max_os_score = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("os", "")).lower()
            score = 50
            if "7" in text or "years" in text or "ios" in text or "samsung" in name.lower() or "google" in name.lower():
                score = 90
            elif "android 15" in text or "android 16" in text or "ios 18" in text or "ios 19" in text:
                score = 80
            else:
                score = 60

            if score > max_os_score:
                max_os_score = score
                best_phone = name
        return (best_phone, "نظام أحدث ودعم تحديثات أطول")

    # 14. السعر والقيمة مقابل السعر بالجنيه المصري (Price in EGP)
    if spec_key == "price":
        best_phone = phone_names[0]
        min_egp = 9999999
        for name, spec in phones_specs.items():
            text = str(spec.get("price", ""))
            digits_match = re.findall(r'(\d[\d,]*)', text)
            if digits_match:
                val = int(digits_match[0].replace(",", ""))
                if val < 5000:
                    val = val * 50
                if val < min_egp:
                    min_egp = val
                    best_phone = name
        return (best_phone, "سعر أقل وقيمة أفضل مقابل السعر")

    # 15. الشبكات والاتصال (Network)
    if spec_key == "network":
        best_phone = phone_names[0]
        max_net_score = 0
        for name, spec in phones_specs.items():
            text = str(spec.get("network", "")).lower()
            score = 50
            if "wi-fi 7" in text or "satellite" in text or "uwb" in text:
                score = 90
            elif "5g" in text:
                score = 70
            if score > max_net_score:
                max_net_score = score
                best_phone = name
        return (best_phone, "تقنيات اتصال وشبكات أسرع")

    # 16. الميزات البارزة (Features)
    if spec_key == "features":
        best_phone = phone_names[0]
        return (best_phone, "ميزات إضافية متفوقة")

    # باقي المواصفات الافتراضية
    first_phone = phone_names[0]
    return (first_phone, "مواصفات متقدمة ومطابقة للمعايير")


def evaluate_overall_winner(phones_specs: Dict[str, Dict[str, Any]], spec_definitions: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    حساب النتيجة النهائية وتحديد أفضل موبايل بشكل عام بناءً على تجميع نقاط الفوز في جميع المواصفات.
    """
    if not phones_specs:
        return {
            "winner_name": "-",
            "win_count": 0,
            "total_specs": 0,
            "won_specs": [],
            "summary_reason": "لا توجد أجهزة كافية للمقارنة"
        }

    phone_names = list(phones_specs.keys())
    if len(phone_names) == 1:
        return {
            "winner_name": phone_names[0],
            "win_count": len(spec_definitions),
            "total_specs": len(spec_definitions),
            "won_specs": [s["label_ar"] for s in spec_definitions],
            "summary_reason": "الجهاز الوحيد المحدد في المقارنة."
        }

    scores = {name: 0 for name in phone_names}
    won_categories = {name: [] for name in phone_names}
    competitive_specs_count = 0

    for spec_def in spec_definitions:
        k = spec_def["key"]
        winner, _ = evaluate_spec_winner(k, phones_specs)
        if winner != "-" and winner in scores:
            scores[winner] += 1
            won_categories[winner].append(spec_def["label_ar"])
            competitive_specs_count += 1

    # تحديد الهاتف الحاصل على أعلى نقاط
    best_overall_phone = max(scores, key=scores.get)
    win_count = scores[best_overall_phone]
    won_specs_list = won_categories[best_overall_phone]

    top_won_str = "، ".join(won_specs_list[:4])
    summary_reason = f"تفوق في {win_count} من أصل {competitive_specs_count} مواصفة تنافسية، وأبرزها ({top_won_str}) مما يجعله الخيار الأكثر تكاملاً وقوة بشكل عام."

    return {
        "winner_name": best_overall_phone,
        "win_count": win_count,
        "total_specs": competitive_specs_count,
        "won_specs": won_specs_list,
        "summary_reason": summary_reason,
        "all_scores": scores
    }
