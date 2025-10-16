import pandas as pd
from lxml import etree
import openpyxl

# Ambil nilai B1 sebagai starting TRANSACTIONID
wb = openpyxl.load_workbook("Book1.xlsx", data_only=True)
ws = wb["Sheet1"]
start_id = ws["B1"].value  # Asumsikan B1 berisi angka awal, misalnya 900

# Load Excel, skip baris pertama karena bukan header
df = pd.read_excel("Book1.xlsx", sheet_name="Sheet1", skiprows=1)
df.columns = df.columns.str.strip().str.upper()
print("Kolom yang terbaca:", df.columns.tolist())


# Buat root XML
root = etree.Element("NMEXML", EximID="12", BranchCode="936526925", ACCOUNTANTCOPYID="")
transactions = etree.SubElement(root, "TRANSACTIONS", OnError="CONTINUE")

# Loop tiap baris Excel
for idx, row in df.iterrows():
    transaction_id = start_id + idx  # Increment dari B1
    other = etree.SubElement(transactions, "OTHERPAYMENT", operation="Add", REQUESTID="1")
    etree.SubElement(other, "TRANSACTIONID").text = str(transaction_id)

    accountline = etree.SubElement(other, "ACCOUNTLINE", operation="Add")
    etree.SubElement(accountline, "KeyID").text = "1"
    etree.SubElement(accountline, "GLACCOUNT").text = str(row["NO AKUN"])
    etree.SubElement(accountline, "GLAMOUNT").text = str(row["TOTAL BAYAR"])
    etree.SubElement(accountline, "DESCRIPTION").text = "" if pd.isna(row["DESCRIPTION"]) else str(row["DESCRIPTION"])
    etree.SubElement(accountline, "RATE").text = "1"
    etree.SubElement(accountline, "TXDATE")
    etree.SubElement(accountline, "POSTED")
    etree.SubElement(accountline, "CURRENCYNAME")

    # Tambahan elemen setelah ACCOUNTLINE
    etree.SubElement(other, "JVNUMBER").text = str(row["NO INVOICE"])
    etree.SubElement(other, "TRANSDATE").text = "2025-10-15"
    etree.SubElement(other, "SOURCE").text = "GL"
    etree.SubElement(other, "TRANSTYPE").text = "other payment"
    etree.SubElement(other, "TRANSDESCRIPTION").text = "" if pd.isna(row["MEMO"]) else str(row["MEMO"])
    etree.SubElement(other, "JVAMOUNT").text = str(row["TOTAL BAYAR"])
    etree.SubElement(other, "CHEQUENO").text = "" if pd.isna(row["CHEQUE NO"]) else str(row["CHEQUE NO"])
    etree.SubElement(other, "PAYEE").text = "" if pd.isna(row["PAYEE"]) else str(row["PAYEE"])
    etree.SubElement(other, "VOIDCHEQUE").text = "0"
    etree.SubElement(other, "GLACCOUNT").text = str(row["AKUN BANK"])
    etree.SubElement(other, "RATE").text = "1"

# Simpan ke file XML
tree = etree.ElementTree(root)
tree.write("output.xml", pretty_print=True, xml_declaration=True, encoding="UTF-8")
print("✅ XML berhasil dibuat: output.xml")