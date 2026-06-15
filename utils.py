from operator import inv

import pandas as pd
import streamlit as st
import re


# -----------------------------------
# READ PURCHASE FILE
# -----------------------------------

def read_purchase_file(file):

    df = pd.read_excel(
        file,
        dtype=str
    )

    # CLEAN COLUMN NAMES
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # REMOVE EMPTY ROWS
    df = df.dropna(
        how="all"
    )

    # REMOVE DUPLICATE COLUMNS
    df = df.loc[
        :,
        ~df.columns.duplicated()
    ]
        
    return df
# -----------------------------------
# READ GSTR2B FILE
# -----------------------------------

def read_gstr2b_file(file):

    xls = pd.ExcelFile(file)

    valid_sheets = [
        "B2B",
        "B2BA",
        "B2B-CDNR",
        "B2B-CDNRA",
        "IMPG"
    ]

    all_data = []

    for sheet in valid_sheets:

        if sheet in xls.sheet_names:

            try:

                # READ SHEET
                temp_df = pd.read_excel(
                    xls,
                    sheet_name=sheet,
                    header=0,
                    dtype=str
                )

                # CLEAN COLUMN NAMES
                temp_df.columns = (
                    temp_df.columns
                    .astype(str)
                    .str.strip()
                    .str.replace("\n", " ", regex=False)
                    .str.upper()
                )

                # REMOVE EMPTY ROWS
                temp_df = temp_df.dropna(
                    how="all"
                )

                # REMOVE DUPLICATE COLUMNS
                temp_df = temp_df.loc[
                    :,
                    ~temp_df.columns.duplicated()
                ]

                # -----------------------------------
                # DOCUMENT TYPE
                # -----------------------------------

                # NORMAL INVOICE
                if sheet in ["B2B", "B2BA"]:

                    temp_df["DOC_TYPE"] = "INV"

                # CREDIT / DEBIT_NOTE
                elif sheet in ["B2B-CDNR", "B2B-CDNRA"]:

                    temp_df["DOC_TYPE"] = "CN"

                    # IF_NOTE TYPE COLUMN AVAILABLE
                    possible_note_cols = [
                        "NOTE TYPE",
                        "NOTETYPE",
                        "DOCUMENT TYPE"
                    ]

                    note_col = None

                    for col in possible_note_cols:

                        if col in temp_df.columns:

                            note_col = col
                            break

                    # DETECT CN / DN
                    if note_col:

                        temp_df.loc[
                            temp_df[note_col]
                            .astype(str)
                            .str.upper()
                            .str.contains("D", na=False),

                            "DOC_TYPE"
                        ] = "DN"

                        temp_df.loc[
                            temp_df[note_col]
                            .astype(str)
                            .str.upper()
                            .str.contains("C", na=False),

                            "DOC_TYPE"
                        ] = "CN"

                else:

                    temp_df["DOC_TYPE"] = "INV"

                # ADD SOURCE SHEET
                temp_df["SOURCE_SHEET"] = sheet
                    
                all_data.append(temp_df)

            except Exception as e:

                st.warning(f"{sheet} Error : {e}")

    # NO DATA
    if len(all_data) == 0:

        return pd.DataFrame()

    # CONCAT
    final_df = pd.concat(
        all_data,
        ignore_index=True,
        sort=False
    )
    # REMOVE DUPLICATE COLUMNS
    final_df = final_df.loc[
        :,
        ~final_df.columns.duplicated()
    ]
    return final_df

# -----------------------------------
# STANDARDIZE COLUMNS
# -----------------------------------

def standardize_columns(df, file_type):

    if df is None or df.empty:

        return pd.DataFrame()
    
    # REMOVE DUPLICATE COLUMNS
    df = df.loc[
        :,
        ~df.columns.duplicated()
    ]

    column_mapping = {}

    # -----------------------------------
    # PURCHASE
    # -----------------------------------

    if file_type == "purchase":

        exact_mapping = {

            "GSTIN": "GSTIN",

            "PARTY NAME": "PARTY_NAME",

            "BILL NO.": "INVOICE_NO",

            "BILL NO": "INVOICE_NO",

            "INVOICE NO": "INVOICE_NO",

            "INVOICE_NO": "INVOICE_NO",

            "BILL DATE": "INVOICE_DATE",

            "INVOICE DATE": "INVOICE_DATE",

            "INVOICE_DATE": "INVOICE_DATE",

            "TAXABLE VALUE OF REC": "TAXABLE_VALUE_REC",

            "TAXABLE VALUE OF PAY": "TAXABLE_VALUE_PAY",

            "IGST RECEIVABLE": "IGST",

            "CGST RECEIVABLE": "CGST",

            "SGST RECEIVABLE": "SGST",

            "BILL AMOUNT": "BILL_AMOUNT",

            "BILL_AMOUNT": "BILL_AMOUNT"
        }

        for col in df.columns:

            clean_col = str(col).strip().upper()

            if clean_col in exact_mapping:

                column_mapping[col] = exact_mapping[clean_col]

        df = df.rename(columns=column_mapping)

        # df = df.loc[
        #     :,
        #     ~df.columns.duplicated()
        # ]
        
        # =========================================
        # SAFE DATE CLEANING
        # =========================================

        if "INVOICE_DATE" in df.columns:

            # CONVERT TO STRING
            df["INVOICE_DATE"] = (
                df["INVOICE_DATE"]
                .astype(str)
                .str.strip()
            )

            # REMOVE TIME
            df["INVOICE_DATE"] = (
                df["INVOICE_DATE"]
                .str.replace(
                    "00:00:00",
                    "",
                    regex=False
                )
                .str.strip()
            )

            # REMOVE NaT
            df["INVOICE_DATE"] = (
                df["INVOICE_DATE"]
                .replace(
                    ["NaT", "nan", "None"],
                    ""
                )
            )

        # CLEAN INVOICE
        if "INVOICE_NO" in df.columns:

            df["INV_NORM"] = (
                df["INVOICE_NO"]
            .astype(str)
            .apply(normalize_invoice)
            )
        # -----------------------------------
        # PURCHASE DOC TYPE
        # -----------------------------------

        def detect_purchase_doc_type(row):

            invoice = str(
                row.get("INVOICE_NO", "")
            ).upper()

            taxable = pd.to_numeric(
                row.get("TAXABLE_VALUE_REC", 0),
                errors="coerce"
            )

            taxable = 0 if pd.isna(taxable) else taxable

            # CREDIT NOTE
            if taxable < 0:

                return "CN"

            # DEBIT NOTE
            elif (
                "DN" in invoice or
                "DR" in invoice or
                "DEBIT" in invoice
            ):

                return "DN"

            # CREDIT NOTE
            elif (
                "CN" in invoice or
                "CR" in invoice or
                "CREDIT" in invoice
            ):

                return "CN"

            # NORMAL INVOICE
            return "INV"

        # APPLY DOC TYPE
        df["DOC_TYPE"] = (
            df.apply(
                detect_purchase_doc_type,
                axis=1
            )
        )

        return df

    # -----------------------------------
    # GSTR2B
    # -----------------------------------
    elif file_type == "gstr2b":
        
        exact_mapping = {

            "GSTIN OF SUPPLIER": "GSTIN",

            "TRADE/LEGAL NAME": "PARTY_NAME",

            "INVOICE NUMBER": "INVOICE_NO",

            "NOTE NUMBER": "INVOICE_NO",

            "INVOICE DATE": "INVOICE_DATE",

            "TAXABLE VALUE (₹)": "TAXABLE_VALUE",

            "CENTRAL TAX(₹)": "CGST",

            "STATE/UT TAX(₹)": "SGST",

            "INTEGRATED TAX(₹)": "IGST",

            "CESS(₹)": "CESS",

            "INVOICE VALUE(₹)": "BILL_AMOUNT",
            
        }
        
        for col in df.columns:

            clean_col = str(col).strip().upper()

            # EXACT MATCH
            if clean_col in exact_mapping:

                column_mapping[col] = exact_mapping[clean_col]

            # FLEXIBLE MATCH
            elif "GSTIN" in clean_col:

                column_mapping[col] = "GSTIN"

            elif "TRADE" in clean_col or "LEGAL NAME" in clean_col:

                column_mapping[col] = "PARTY_NAME"

            elif "INVOICE" in clean_col and "NUMBER" in clean_col:

                column_mapping[col] = "INVOICE_NO"

            elif "NOTE" in clean_col and "NUMBER" in clean_col:

                column_mapping[col] = "INVOICE_NO"
                
            elif "INVOICE" in clean_col and "DATE" in clean_col:

                column_mapping[col] = "INVOICE_DATE"
                
            elif "NOTE" in clean_col and "DATE" in clean_col:

                column_mapping[col] = "INVOICE_DATE"

            elif "TAXABLE VALUE" in clean_col:

                column_mapping[col] = "TAXABLE_VALUE"

            elif "CENTRAL TAX" in clean_col:

                column_mapping[col] = "CGST"

            elif "STATE/UT TAX" in clean_col:

                column_mapping[col] = "SGST"

            elif "INTEGRATED TAX" in clean_col:

                column_mapping[col] = "IGST"

            elif "INVOICE VALUE" in clean_col:

                column_mapping[col] = "BILL_AMOUNT"
            
        # CDNR NOTE NUMBER -> INVOICE NUMBER

        if "NOTE NUMBER" in df.columns:
            df["INVOICE NUMBER"] = df["INVOICE NUMBER"].fillna(
                df["NOTE NUMBER"]
        )

        if "NOTE DATE" in df.columns:
            df["INVOICE DATE"] = df["INVOICE DATE"].fillna(
                df["NOTE DATE"]
        )

        df = df.rename(columns=column_mapping)
        
        # REMOVE DUPLICATE COLUMNS
        df = df.loc[:, ~df.columns.duplicated()]
        
        
        # FORMAT DATE   

        if "INVOICE_DATE" in df.columns:

            df["INVOICE_DATE"] = (
                df["INVOICE_DATE"]
                .astype(str)
                .str[:10]
            )
        # CLEAN INVOICE

        if "INVOICE_NO" in df.columns:

            df["INV_NORM"] = (

                df["INVOICE_NO"]
                .astype(str)
                .apply(normalize_invoice)
            )
            
        
        # -----------------------------------
        # DOCUMENT TYPE
        # -----------------------------------

        df["DOC_TYPE"] = "INV"

        df.loc[
            df["SOURCE_SHEET"]
            .isin([
            "B2B-CDNR",
            "B2B-CDNRA"
        ]),
        "DOC_TYPE"
        ] = "CN"
        numeric_cols = [
            "TAXABLE_VALUE_REC",
            "TOTAL_VALUE",
            "TOTAL_TAX",
            "CGST",
            "SGST",
            "IGST",
            "CESS",
            "BILL_AMOUNT"
        ]

        for col in numeric_cols:

            if col in df.columns:

                df[col] = pd.to_numeric(
                    df[col],
                    errors="coerce"
                ).fillna(0)

        return df
    # -----------------------------------
    # DOCUMENT TYPE
    # -----------------------------------

    df["DOC_TYPE"] = "INV"

    if "SOURCE_SHEET" in df.columns:

        df.loc[
            df["SOURCE_SHEET"]
            .isin(["CDNR", "CDNRA"]),

            "DOC_TYPE"
        ] = "CN"    


# -----------------------------------
# CREATE TOTAL TAX
# -----------------------------------
def create_total_tax(df):

    # COPY DATAFRAME
    temp_df = df.copy()

    tax_columns = [
        "IGST",
        "CGST",
        "SGST",
        "CESS"
    ]

    # CREATE MISSING COLUMNS
    for col in tax_columns:

        if col not in temp_df.columns:

            temp_df[col] = 0

        temp_df[col] = pd.to_numeric(
            temp_df[col],
            errors="coerce"
        ).fillna(0)

    # TOTAL TAX
    temp_df["TOTAL_TAX"] = (

        temp_df[tax_columns]
        .sum(axis=1)
        .round(0)

    )

    return temp_df
    
# =========================================
# FINAL GST INVOICE CLEANER
# =========================================
import re
import pandas as pd

def normalize_invoice(inv):

    if pd.isna(inv):
        return ""

    inv = str(inv).upper().strip()
    
    inv = re.sub(r'([-/])25-26$', '', inv, flags=re.IGNORECASE)
    inv = re.sub(r'([-/])25$', '', inv, flags=re.IGNORECASE)
    
    m = re.match(r'^([A-Z]+)/(\d+)/25-26$', inv)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    
    # 033/2025-26/C -> 33
    m = re.match(r'^0*(\d+)/20\d{2}-\d{2}/[A-Z]$', inv)

    if m:
        return str(int(m.group(1)))
    # GST remove
    inv = re.sub(r'^GST[-/ ]*', '', inv)
    
    # remove all non-alphanumeric
    inv = re.sub(r'[^A-Z0-9]', '', inv)
    
    # MG252611 -> MG11
    inv = re.sub(r'^([A-Z]+)2526(\d+)$', r'\1\2', inv)
    
    # G/101 -> 101
    inv = re.sub(r'^G/', '', inv)

    # 00036 -> 36
    if inv.isdigit():
        inv = str(int(inv))
        
    # remove leading G
    if inv.startswith("G") and len(inv) > 1:
        inv = inv[1:]
        
    # MG/25-26/11 -> MG/11
    inv = re.sub(r'([A-Z]+)/\d{2}\d{2}/(\d+)$', r'\1/\2', inv)
    
    # 15A -> 15
    if re.match(r'^\d+A$', inv):
        inv = inv[:-1]
    return inv
# -----------------------------------
# RECONCILIATION
# -----------------------------------

def reconcile_data(purchase_df, gstr2b_df):
    print(gstr2b_df["SOURCE_SHEET"].value_counts())
    
    purchase_df = purchase_df.copy()
    gstr2b_df = gstr2b_df.copy()

    # ==========================
    # REMOVE CN / DN FROM RECON
    # ==========================

    cdnr_df = gstr2b_df[
        gstr2b_df["SOURCE_SHEET"].isin(
            ["B2B-CDNR", "B2B-CDNRA"]
        )
    ].copy()

    gstr2b_df = gstr2b_df[
        gstr2b_df["SOURCE_SHEET"].isin(
            ["B2B", "B2BA", "IMPG"]
        )
    ].copy()
    # =========================================
    # CLEAN COLUMN NAMES
    # =========================================

    purchase_df.columns = [
        col.strip()
        .upper()
        .replace(" ", "_")
        .replace(".", "")
        .replace("(", "")
        .replace(")", "")
        .replace("%", "")
        .replace("₹", "")
        .replace("/", "_")
        for col in purchase_df.columns
    ]

    gstr2b_df.columns = [
        col.strip()
        .upper()
        .replace(" ", "_")
        .replace(".", "")
        .replace("(", "")
        .replace(")", "")
        .replace("%", "")
        .replace("₹", "")
        .replace("/", "_")

        for col in gstr2b_df.columns
    ]
    
    # =========================================
    # TOTAL TAX - PURCHASE
    # =========================================

    purchase_df["TOTAL_TAX"] = (

        pd.to_numeric(
            purchase_df.get(
                "IGST",
                pd.Series(0, index=purchase_df.index)
            ),
            errors="coerce"
        ).fillna(0)

        +

        pd.to_numeric(
            purchase_df.get(
                "CGST",
                pd.Series(0, index=purchase_df.index)
            ),
            errors="coerce"
        ).fillna(0)

        +

        pd.to_numeric(
            purchase_df.get(
                "SGST",
                pd.Series(0, index=purchase_df.index)
            ),
            errors="coerce"
        ).fillna(0)
    )

    # =========================================
    # TOTAL TAX - 2B
    # =========================================

    gstr2b_df["TOTAL_TAX"] = (

        pd.to_numeric(
            gstr2b_df.get("IGST",0),
            errors="coerce"
        ).fillna(0)

        +

        pd.to_numeric(
            gstr2b_df.get("CGST",0),
            errors="coerce"
        ).fillna(0)

        +

        pd.to_numeric(
            gstr2b_df.get("SGST",0),
            errors="coerce"
        ).fillna(0)
    )
    # =========================================
    # NUMERIC COLUMNS
    # =========================================

    numeric_cols = [
        "TAXABLE_VALUE",
        "TOTAL_TAX"
    ]

    for col in numeric_cols:

        if col in purchase_df.columns:

            purchase_df[col] = pd.to_numeric(
                purchase_df[col],
                errors="coerce"
            ).fillna(0)

        if col in gstr2b_df.columns:

            gstr2b_df[col] = pd.to_numeric(
                gstr2b_df[col],
                errors="coerce"
            ).fillna(0)
    # =========================================
    # DOC TYPE CREATE - PURCHASE
    # =========================================

    purchase_df["DOC_TYPE"] = "INV"

    # CHECK COLUMN EXISTS
    if "TRAN_TYPE" in purchase_df.columns:

        purchase_df.loc[
            purchase_df["TRAN_TYPE"]
            .astype(str)
            .str.upper()
            .str.contains(
                "CREDIT NOTE|PURCHASE RETURN",
                na=False
            ),
            "DOC_TYPE"
        ] = "CN"

        purchase_df.loc[
            purchase_df["TRAN_TYPE"]
            .astype(str)
            .str.upper()
            .str.contains(
                "DEBIT NOTE",
                na=False
            ),

            "DOC_TYPE"
        ] = "DN"

    # =========================================
    # NUMERIC CONVERSION BEFORE MATCH KEY
    # =========================================

    purchase_df["TAXABLE_VALUE_REC"] = pd.to_numeric(
        purchase_df["TAXABLE_VALUE_REC"],
        errors="coerce"
    ).fillna(0)

    purchase_df["TOTAL_TAX"] = pd.to_numeric(
        purchase_df["TOTAL_TAX"],
        errors="coerce"
    ).fillna(0)

    gstr2b_df["TAXABLE_VALUE"] = pd.to_numeric(
        gstr2b_df["TAXABLE_VALUE"],
        errors="coerce"
    ).fillna(0)

    gstr2b_df["TOTAL_TAX"] = pd.to_numeric(
        gstr2b_df["TOTAL_TAX"],
        errors="coerce" 
    ).fillna(0)

    # =========================================
    # SORT BEFORE MATCHING
    # =========================================

    purchase_df = purchase_df.sort_values(
        by=[
            "GSTIN",
            "INV_NORM",
            "TAXABLE_VALUE_REC"
        ]
    )

    gstr2b_df = gstr2b_df.sort_values(
        by=[
            "GSTIN",
            "INV_NORM",
            "TAXABLE_VALUE"
        ]
    )
    # ==========================
    # PURCHASE AGGREGATION
    # ==========================

    purchase_df["TOTAL_TAX"] = (
        purchase_df["CGST"].fillna(0)
        + purchase_df["SGST"].fillna(0)
        + purchase_df["IGST"].fillna(0)
    )
    purchase_df = purchase_df.groupby(
        ["GSTIN", "INV_NORM", "DOC_TYPE"],
        as_index=False
    ).agg({
        "TRAN_TYPE": "first",
        "PARTY_NAME": "first",
        "INVOICE_NO": "first",
        "INVOICE_DATE": "first",
        "BILL_AMOUNT": "sum",
        "INVOICE_TYPE": "first",
        "STATE_CODE": "first",
        "DIVISION": "first",
        "IGST": "sum",
        "CGST": "sum",
        "SGST": "sum",
        "TAXABLE_VALUE_REC": "sum",
        "TOTAL_TAX": "sum"
    })
    # ==========================
    # 2B AGGREGATION
    # ==========================

    gstr2b_df["TOTAL_TAX"] = (
        gstr2b_df["CGST"].fillna(0)
        + gstr2b_df["SGST"].fillna(0)
        + gstr2b_df["IGST"].fillna(0)
    )

    gstr2b_df = gstr2b_df.groupby(
        ["GSTIN", "INV_NORM", "DOC_TYPE"],
        as_index=False
    ).agg({
        "PARTY_NAME": "first",
        "INVOICE_NO": "first",
        "INVOICE_DATE": "first",
        "TAXABLE_VALUE": "sum",
        "TOTAL_TAX": "sum"
    })
    
    purchase_df["MATCH_KEY"] = (
        purchase_df["GSTIN"].astype(str)
        + "_"
        + purchase_df["DOC_TYPE"].astype(str)
        + "_"
        + purchase_df["INV_NORM"].astype(str)
    )

    gstr2b_df["MATCH_KEY"] = (
        gstr2b_df["GSTIN"].astype(str)
        + "_"
        + gstr2b_df["DOC_TYPE"].astype(str)
        + "_"
        + gstr2b_df["INV_NORM"].astype(str)
    )
    purchase_df = purchase_df.drop_duplicates(
        subset=["MATCH_KEY"]
    )

    gstr2b_df = gstr2b_df.drop_duplicates(
        subset=["MATCH_KEY"]
    )
    # =========================================
    # SIMPLE MERGE
    # =========================================
    merged_df = purchase_df.merge(

        gstr2b_df[
            [
                "MATCH_KEY",
                "TAXABLE_VALUE",
                "TOTAL_TAX"
            ]
        ].rename(
            columns={
                "TAXABLE_VALUE": "TAXABLE_VALUE_2B",
                "TOTAL_TAX": "TOTAL_TAX_2B"
            }
        ),

        on="MATCH_KEY",

        how="left"
    )    
    
    merged_df["RECON_STATUS"] = "MISSING IN 2B"
    merged_df["REMARK"] = (
        "INVOICE NOT AVAILABLE IN GSTR2B"
    )
    
    # =========================================
    # RECON STATUS
    # =========================================

    matched_mask = (
        merged_df["TAXABLE_VALUE_2B"].notna()
    )

    merged_df.loc[
        matched_mask,
        "RECON_STATUS"
    ] = "MATCHED"

    merged_df.loc[
        matched_mask,
        "REMARK"
    ] = "INVOICE FOUND IN GSTR2B"

    taxable_diff = (
        merged_df["TAXABLE_VALUE_REC"]
        -
        merged_df["TAXABLE_VALUE_2B"]
    ).abs()

    tax_diff = (
        merged_df["TOTAL_TAX"]
        -
        merged_df["TOTAL_TAX_2B"]
    ).abs()

    merged_df.loc[
        matched_mask &
        (taxable_diff > 5),
        "RECON_STATUS"
    ] = "TAXABLE VALUE MISMATCH"

    merged_df.loc[
        matched_mask &
        (taxable_diff > 5),
        "REMARK"
    ] = "TAXABLE VALUE DIFFERENCE FOUND"

    merged_df.loc[
        matched_mask &
        (taxable_diff <= 5) &
        (tax_diff > 5),
        "RECON_STATUS"
    ] = "TAX AMOUNT DIFFERENCE"

    merged_df.loc[
        matched_mask &
        (taxable_diff <= 5) &
        (tax_diff > 5),
        "REMARK"
    ] = "TAX DIFFERENCE FOUND"

    merged_df.loc[
        matched_mask &
        (taxable_diff <= 5) &
        (tax_diff <= 5),
        "RECON_STATUS"
    ] = "MATCHED"

    merged_df.loc[
        matched_mask &
        (taxable_diff <= 5) &
        (tax_diff <= 5),
        "REMARK"
    ] = "EXACT MATCH FOUND"
    
    # =========================================
    # EXTRA RECORDS IN 2B
    # =========================================

    extra_2b = gstr2b_df[
        ~gstr2b_df["MATCH_KEY"].isin(
            purchase_df["MATCH_KEY"]
        )
    ].copy()

    extra_2b["TRAN_TYPE"] = ""
    extra_2b["STATE_CODE"] = ""
    extra_2b["DIVISION"] = ""

    extra_2b["TAXABLE_VALUE_2B"] = extra_2b["TAXABLE_VALUE"]
    extra_2b["TOTAL_TAX_2B"] = extra_2b["TOTAL_TAX"]

    extra_2b["TAXABLE_VALUE_REC"] = 0
    extra_2b["TOTAL_TAX"] = 0


    extra_2b["RECON_STATUS"] = "MISSING IN PURCHASE"

    extra_2b["REMARK"] = (
        "AVAILABLE IN GSTR2B BUT NOT IN PURCHASE"
    )

    # SAME COLUMNS AS MERGED REPORT
    extra_2b = extra_2b.reindex(
        columns=merged_df.columns
    )

    # APPEND
    merged_df = pd.concat(
        [merged_df, extra_2b],
        ignore_index=True
    )
    print("Extra 2B Records:", len(extra_2b))
    print(merged_df["RECON_STATUS"].value_counts())

    # ========================================
    # FALLBACK MATCH DATASET
    # =========================================

    fallback_2b = gstr2b_df.copy()

    fallback_2b["DATE_VALUE_KEY"] = (
        fallback_2b["GSTIN"].astype(str)
        + "_"
        + fallback_2b["INVOICE_DATE"].astype(str)
    )

    merged_df["DATE_VALUE_KEY"] = (
        merged_df["GSTIN"].astype(str)
        + "_"
        + merged_df["INVOICE_DATE"].astype(str)
    )

    missing_df = merged_df[
        merged_df["RECON_STATUS"] == "MISSING IN 2B"
    ].copy()

    missing_df["ORIGINAL_INDEX"] = missing_df.index

    fallback_match = missing_df.merge(

        fallback_2b[
            [
                "DATE_VALUE_KEY",
                "TAXABLE_VALUE",
                "TOTAL_TAX"
            ]
        ],
        on="DATE_VALUE_KEY",

        how="left",

        suffixes=("", "_FB")
    )

    taxable_diff_fb = (
        fallback_match["TAXABLE_VALUE_REC"]
        -
        fallback_match["TAXABLE_VALUE"]
    ).abs()

    tax_diff_fb = (
        fallback_match["TOTAL_TAX"]
        -
        fallback_match["TOTAL_TAX_FB"]
    ).abs()

    fallback_ok = (
        (taxable_diff_fb <= 5)
        &
        (tax_diff_fb <= 5)
    )

    matched_index = fallback_match.loc[
        fallback_ok,
        "ORIGINAL_INDEX"
    ]

    merged_df.loc[
        matched_index,
        "RECON_STATUS"
    ] = "MATCHED BY VALUE"

    merged_df.loc[
        matched_index,
        "REMARK"
    ] = "MATCHED USING GSTIN + DATE + VALUE"


    # =========================================
    # TOLERANCE MATCH
    # =========================================

    taxable_diff = abs(
        merged_df["TAXABLE_VALUE_REC"].fillna(0)
        -
        merged_df["TAXABLE_VALUE_2B"].fillna(0)
    )

    mask = (
        (merged_df["RECON_STATUS"] == "TAXABLE VALUE MISMATCH")
        &
        (taxable_diff <= 5)
    )

    merged_df.loc[
        mask,
        "RECON_STATUS"
    ] = "MATCHED"
    print(merged_df["RECON_STATUS"].value_counts())
    return merged_df, cdnr_df, extra_2b