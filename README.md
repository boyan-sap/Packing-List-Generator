# 📦 MPO Refill Packing List Generator

Automatically generate branded **packing list PDFs** for MPO (store) refills using Excel data. This tool creates individual PDFs per customer (MPO), as well as a merged PDF with all deliveries.

---

## ✅ Features

- 📊 Reads structured Excel input (`packing_list.xlsx`)
- 🧾 Generates a PDF packing list per MPO (store)
- 📎 Merges all PDFs into a single file (`All_Refills_Combined.pdf`)
- 🎨 Includes company logos, styled tables, and delivery info
- ✍️ Auto-adds signature placeholders (Issued, Taken, Received)
- 📅 Asks user for delivery date via dialog
- 📁 Opens the export folder automatically

---

## 📂 Input File: `packing_list.xlsx`

Your Excel file must include these columns:

| Customer Name (MPO) | Delivery | Material | Article | EAN | Material Name | Qty | Collab Order | Delivery Address |
|---------------------|----------|----------|---------|-----|----------------|-----|----------------|------------------|

Example:

| MPO         | Delivery | Material | Article | EAN         | Material Name | Qty | Collab Order | Delivery Address        |
|-------------|----------|----------|---------|-------------|----------------|-----|----------------|--------------------------|
| Store A     | 123456   | 111111   | 222222  | 1234567890123 | Power Cable   | 10  | ORD001         | Some Street 123, City   |

---

## 🛠 Dependencies

Install required Python libraries:

```bash
pip install pandas reportlab PyPDF2
