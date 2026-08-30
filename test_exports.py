import sys
import pandas as pd
from specs_data import SPEC_DEFINITIONS, get_mobile_specs, search_phone_suggestions
from export_service import (
    export_comparison_to_excel,
    export_comparison_to_pdf,
    generate_html_print_view,
    export_s26_settings_to_excel,
    export_phone_settings_to_excel
)
from analyzer import evaluate_spec_winner, evaluate_overall_winner

def test_all():
    phones = ["Samsung Galaxy M51", "Samsung Galaxy S25 Plus", "iPhone 17 Pro Max"]
    
    # 1. Test Autocomplete suggestions
    suggs = search_phone_suggestions("M51")
    print("M51 suggestions:", suggs)
    assert len(suggs) > 0, "Suggestions should not be empty"

    # 2. Build specs & comparison DataFrame
    all_specs = {p: get_mobile_specs(p) for p in phones}
    
    df_dict = {
        "المواصفة (Specification)": [f"{s['label_ar']} ({s['label_en']})" for s in SPEC_DEFINITIONS]
    }
    for p in phones:
        df_dict[p] = [all_specs[p].get(s["key"], "غير متوفر") for s in SPEC_DEFINITIONS]

    winners = []
    reasons = []
    for s in SPEC_DEFINITIONS:
        w, r = evaluate_spec_winner(s["key"], all_specs)
        winners.append(w)
        reasons.append(r)

    df_dict["سبب الأفضلية"] = reasons

    df = pd.DataFrame(df_dict)
    verdict = evaluate_overall_winner(all_specs, SPEC_DEFINITIONS)

    print("Overall Winner:", verdict["winner_name"])
    print("Scores:", verdict["all_scores"])
    print("DataFrame shape:", df.shape)
    assert df.shape == (20, 5), f"Expected shape (20, 5), got {df.shape}"
    assert "سبب الأفضلية" in df.columns, "Column 'سبب الأفضلية' must be present"

    # 3. Test Comparison Exports
    excel_bytes = export_comparison_to_excel(df, "Test Export", winners, verdict)
    print("Excel size:", len(excel_bytes))
    assert len(excel_bytes) > 1000

    pdf_bytes = export_comparison_to_pdf(df, "Test Export", winners, verdict)
    print("PDF size:", len(pdf_bytes))
    assert len(pdf_bytes) > 5000

    html_str = generate_html_print_view(df, "Test Export", winners, verdict)
    print("HTML length:", len(html_str))
    assert len(html_str) > 1000

    # 4. Test S26 Plus & Universal Settings Excel Export
    s26_bytes = export_s26_settings_to_excel()
    print("S26 Settings Excel size:", len(s26_bytes))
    assert len(s26_bytes) > 5000, "S26 Settings Excel should be generated properly"

    # Apple iPhone Guide
    apple_bytes = export_phone_settings_to_excel("iPhone 17 Pro Max", "Apple iPhone (آبل)", "iPhone 17 Pro Max")
    print("Apple Settings Excel size:", len(apple_bytes))
    assert len(apple_bytes) > 5000, "Apple Settings Excel should be generated properly"

    # Xiaomi Guide
    xiaomi_bytes = export_phone_settings_to_excel("Xiaomi 15 Ultra", "Xiaomi & Poco & Redmi (شاومي)", "Xiaomi 15 Ultra")
    print("Xiaomi Settings Excel size:", len(xiaomi_bytes))
    assert len(xiaomi_bytes) > 5000, "Xiaomi Settings Excel should be generated properly"

    # Oppo Guide
    oppo_bytes = export_phone_settings_to_excel("Find X8 Pro", "OPPO (أوبو)", "Find X8 Pro")
    print("OPPO Settings Excel size:", len(oppo_bytes))
    assert len(oppo_bytes) > 5000, "OPPO Settings Excel should be generated properly"

    print("ALL TESTS PASSED WITH 100% SUCCESS!")

if __name__ == "__main__":
    test_all()
