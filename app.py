import streamlit as st
import pandas as pd
from lxml import etree
import time

# ✅ Fungsi cache untuk baca Excel
@st.cache_data
def load_excel(file):
    raw = pd.read_excel(file, sheet_name="Sheet1", header=None)
    df = pd.read_excel(file, sheet_name="Sheet1", skiprows=1)
    return raw, df

st.set_page_config(page_title="Excel to Accurate XML", layout="centered")
st.title("📄 Excel to Accurate XML Converter")
st.markdown("#### 💵 Fitur Pembayaran & Penerimaan Modul Kas dan Bank")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
if uploaded_file:
    # Ambil TRANSACTIONID awal dari B1
    raw_excel, df = load_excel(uploaded_file)
    try:
        start_id = int(raw_excel.iloc[0, 1])
    except Exception:
        st.error("❌ Sel B1 harus berisi angka sebagai TRANSACTIONID awal.")
        st.stop()

    # Pilih jenis transaksi
    trans_type = st.selectbox("Pilih jenis transaksi:", ["Pembayaran", "Penerimaan"])

    # Baca data mulai dari baris ke-2 (karena baris 1 = start_id) ambil header
    df.columns = df.columns.str.strip().str.upper()

    # Validasi kolom wajib
    if trans_type == "Pembayaran":
        required_cols = ["NO INVOICE", "TANGGAL","NO AKUN", "NAMA AKUN","TOTAL BAYAR", "DESCRIPTION", "MEMO", "CHEQUE NO", "PAYEE", "AKUN BANK","NAMA BANK"]
    else:
        required_cols = ["NO INVOICE", "TANGGAL", "NO AKUN", "NAMA AKUN","TOTAL TERIMA", "DESCRIPTION", "MEMO", "AKUN BANK", "NAMA BANK","RATE"]

    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        st.error(f"❌ Kolom berikut wajib ada di Excel: {', '.join(missing)}")
        st.stop()

    st.success("✅ File berhasil dibaca. Preview data:")
    # st.dataframe(df) non pagination

    df_display = df.copy()
    df_display.insert(0, "No", range(1, len(df_display) + 1))

    # pagination version
    page_size = 10
    total_rows = len(df_display)  # ✅ pakai df_display, bukan df
    total_pages = (total_rows + page_size - 1) // page_size
    page = st.number_input("Halaman:", min_value=1, max_value=total_pages, value=1)
    start_idx = (page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)

    st.write(f"Menampilkan baris {start_idx + 1} sampai {end_idx} dari total {total_rows}")
    st.data_editor(df_display.iloc[start_idx:end_idx], hide_index=True, disabled=True, key="preview_data")  # ✅ tambahkan key

    # validasi AKUN BANK
    invalid_rows = df[df["AKUN BANK"].astype(str).str.strip().isin(["", "PILIH AKUN BANK"])]

    if not invalid_rows.empty:
        st.error(f"❌ Ada {len(invalid_rows)} baris dengan AKUN BANK kosong atau belum dipilih.")
        # tampilkan data bank bermasalah
        invalid_rows_display = invalid_rows.copy()
        invalid_rows_display.insert(0, "No", range(1, len(invalid_rows_display) + 1))
        st.data_editor(invalid_rows_display, hide_index=True, disabled=True, key="invalid_data")  # ✅ tambahkan key
        st.info("Perbaiki data dulu sebelum lanjut generate XML.")
    else:
        st.success("✅ Data valid. Siap diproses.")
        if st.button("🔄 Generate XML"):
            progress = st.progress(0, text="⏳ Sedang memproses data...")
            root = etree.Element("NMEXML", EximID="12", BranchCode="936526925", ACCOUNTANTCOPYID="")
            transactions = etree.SubElement(root, "TRANSACTIONS", OnError="CONTINUE")

            for idx, row in df.iterrows():
                transaction_id = start_id + idx
                tag_name = "OTHERPAYMENT" if trans_type == "Pembayaran" else "OTHERDEPOSIT"
                entry = etree.SubElement(transactions, tag_name, operation="Add", REQUESTID="1")
                etree.SubElement(entry, "TRANSACTIONID").text = str(transaction_id)

                # ACCOUNTLINE
                accountline = etree.SubElement(entry, "ACCOUNTLINE", operation="Add")
                etree.SubElement(accountline, "KeyID").text = "1"
                etree.SubElement(accountline, "GLACCOUNT").text = str(row["NO AKUN"])
                amount_field = "TOTAL BAYAR" if trans_type == "Pembayaran" else "TOTAL TERIMA"
                etree.SubElement(accountline, "GLAMOUNT").text = str(row[amount_field])
                etree.SubElement(accountline, "DESCRIPTION").text = "" if pd.isna(row["DESCRIPTION"]) else str(row["DESCRIPTION"])
                rate_value = row.get("RATE", "")
                rate_clean = str(rate_value).strip()
                etree.SubElement(accountline, "RATE").text = "1" if rate_clean == "" or pd.isna(rate_value) else rate_clean
                for tag in ["TXDATE", "POSTED", "CURRENCYNAME"]:
                    etree.SubElement(accountline, tag)

                # Common fields
                etree.SubElement(entry, "JVNUMBER").text = str(row["NO INVOICE"])
                etree.SubElement(entry, "TRANSDATE").text = "2025-10-15"
                etree.SubElement(entry, "SOURCE").text = "GL"
                etree.SubElement(entry, "TRANSTYPE").text = "other payment" if trans_type == "Pembayaran" else "other receipt"
                etree.SubElement(entry, "TRANSDESCRIPTION").text = "" if pd.isna(row["MEMO"]) else str(row["MEMO"])
                etree.SubElement(entry, "JVAMOUNT").text = str(row[amount_field])
                etree.SubElement(entry, "GLACCOUNT").text = str(row["AKUN BANK"])
                etree.SubElement(entry, "RATE").text = "1"

                # Field khusus
                # if trans_type == "Pembayaran":
                #     etree.SubElement(entry, "CHEQUENO").text = "" if pd.isna(row["CHEQUE NO"]) else str(row["CHEQUE NO"])
                #     etree.SubElement(entry, "PAYEE").text = "" if pd.isna(row["PAYEE"]) else str(row["PAYEE"])
                #     etree.SubElement(entry, "VOIDCHEQUE").text = "0"
                # else:
                #     etree.SubElement(entry, "RECEIPTNO").text = "" if pd.isna(row["RECEIPT NO"]) else str(row["RECEIPT NO"])
                #     etree.SubElement(entry, "CUSTOMER").text = "" if pd.isna(row["CUSTOMER"]) else str(row["CUSTOMER"])
                #     etree.SubElement(entry, "VOIDRECEIPT").text = "0"

                # Update progress
                progress.progress((idx + 1) / len(df), text=f"⏳ Memproses baris {idx + 1} dari {len(df)}")
                time.sleep(0.05)
                progress.empty()  # Hilangkan progress bar


            xml_bytes = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            nama_file = "pembayaran_accurate.xml" if trans_type == "Pembayaran" else "penerimaan_accurate.xml"
            st.download_button("⬇️ Download XML", data=xml_bytes, file_name=nama_file, mime="application/xml")
