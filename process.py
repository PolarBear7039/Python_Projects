import os
import re
import pandas as pd


def clean_sub_prefix(text):
    if not text or pd.isna(text):
        return ""
    text = str(text).strip()
    cleaned = re.sub(
        r"^[\u0600-\u06FF0-9a-zA-Z]{1,3}\s*[\-\–\.\:\)]\s*", "", text
    ).strip()
    return cleaned.rstrip(" .")


def format_sub_item_number(item_no, sub_idx):
    item_no = str(item_no).strip()
    if "/" in item_no:
        parts = item_no.split("/", 1)
        return f"{parts[0]}.{sub_idx}/{parts[1]}"
    else:
        return f"{item_no}.{sub_idx}"


def preprocess_merged_rows(raw_rows):
    merged_rows = []

    for row in raw_rows:
        item_no = str(row[0]).strip() if len(row) > 0 else ""
        desc = str(row[1]).strip() if len(row) > 1 else ""
        unit = str(row[2]).strip() if len(row) > 2 else ""
        price = str(row[3]).strip() if len(row) > 3 else ""

        if "رقم البند" in item_no or "بيان الأعمال" in desc:
            continue

        if not item_no and not desc and not unit and not price:
            continue

        if not item_no and merged_rows:
            prev_item_no, prev_desc, prev_unit, prev_price = merged_rows[-1]

            new_desc = f"{prev_desc} {desc}".strip() if desc else prev_desc
            new_unit = unit if not prev_unit else prev_unit
            new_price = price if not prev_price else prev_price

            merged_rows[-1] = (prev_item_no, new_desc, new_unit, new_price)
        else:
            merged_rows.append((item_no, desc, unit, price))

    return merged_rows


def process_boq_data(input_file_path, output_excel_path):
    if input_file_path.endswith(".xlsx") or input_file_path.endswith(".xls"):
        df_raw = pd.read_excel(input_file_path, header=None, dtype=str)
    else:
        df_raw = pd.read_csv(
            input_file_path,
            sep="\t",
            header=None,
            dtype=str,
            on_bad_lines="skip",
            engine="python",
        )

    df_raw = df_raw.fillna("")
    raw_rows = df_raw.values.tolist()

    rows = preprocess_merged_rows(raw_rows)

    processed_data = []

    groups = []
    current_group = []
    last_item_no = None

    for item_no, desc, unit, price in rows:
        if not item_no:
            processed_data.append({
                "رقم البند": "",
                "بيان البند": desc,
                "البند الفرعي": "",
                "الوحدة": unit,
                "سعر الوحدة": price,
                "ملاحظات": "سطر غريب - رقم البند مفقود",
            })
            continue

        if item_no == last_item_no:
            current_group.append((item_no, desc, unit, price))
        else:
            if current_group:
                groups.append(current_group)
            current_group = [(item_no, desc, unit, price)]
            last_item_no = item_no

    if current_group:
        groups.append(current_group)

    for group in groups:
        item_no = group[0][0]

        if len(group) == 1:
            _, desc, unit, price = group[0]
            notes = ""
            if not desc:
                notes = "بيان البند مفقود"

            processed_data.append({
                "رقم البند": item_no,
                "بيان البند": desc.rstrip(" ."),
                "البند الفرعي": "لا يوجد",
                "الوحدة": unit,
                "سعر الوحدة": price,
                "ملاحظات": notes,
            })
        else:
            first_desc = group[0][1]
            first_unit = group[0][2]
            first_price = group[0][3]

            has_sub_prefix = bool(
                re.match(
                    r"^[\u0600-\u06FF0-9a-zA-Z]{1,3}\s*[\-\–\.\:\)]",
                    first_desc,
                )
            )

            if not has_sub_prefix or (not first_unit and not first_price):
                main_desc = first_desc.rstrip(" .")
                sub_items = group[1:]
            else:
                main_desc = ""
                sub_items = group

            sub_idx = 1
            for _, sub_desc, sub_unit, sub_price in sub_items:
                formatted_item_no = format_sub_item_number(item_no, sub_idx)
                clean_sub = clean_sub_prefix(sub_desc)

                notes = ""
                if not sub_desc and not sub_unit and not sub_price:
                    notes = "سطر فرعي فارغ أو به بيانات مفقودة"

                processed_data.append({
                    "رقم البند": formatted_item_no,
                    "بيان البند": main_desc,
                    "البند الفرعي": clean_sub,
                    "الوحدة": sub_unit,
                    "سعر الوحدة": sub_price,
                    "ملاحظات": notes,
                })
                sub_idx += 1

    df_result = pd.DataFrame(processed_data)
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="المقايسة المفرغة")

    print(
        f"✅ تم معالجة {len(processed_data)} سطر بنجاح واستخراج الملف إلى:"
        f" {output_excel_path}"
    )


if __name__ == "__main__":
    output_file = "جدول_المقايسة_النهائي.xlsx"

    if os.path.exists("data.xlsx"):
        process_boq_data("data.xlsx", output_file)
    elif os.path.exists("data.xls"):
        process_boq_data("data.xls", output_file)
    elif os.path.exists("data.txt"):
        process_boq_data("data.txt", output_file)
    else:
        print(
            "⚠️ لم يتم العثور على ملف data.xlsx أو data.txt داخل الفولدر. يرجى"
            " التأكد من وجوده."
        )