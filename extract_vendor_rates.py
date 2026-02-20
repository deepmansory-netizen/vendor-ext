import pdfplumber
import pandas as pd
import re
import os

pdf_folder = r"C:\Users\dpatel\Desktop\Sample"
output_file = r"C:\Users\dpatel\Desktop\all_vendor_rates.xlsx"

print("Starting Vendor Rate Extractor...")

if not os.path.exists(pdf_folder):
    print("ERROR: Folder not found at " + pdf_folder)
    input("Press Enter to exit...")
    quit()

pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith(".pdf")]
total = len(pdf_files)
print("Found " + str(total) + " PDFs")

data = []

for i, pdf_file in enumerate(pdf_files, 1):
    print("[" + str(i) + "/" + str(total) + "] " + pdf_file)
    try:
        with pdfplumber.open(os.path.join(pdf_folder, pdf_file)) as pdf:
            text = ""
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text = text + extracted + "\n"

        op = re.search(r"Peak[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)
        onp = re.search(r"Non-Peak[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)
        or_ = re.search(r"Origin.*?reduced by \$([\d.]+)", text, re.I | re.DOTALL)
        dp = re.search(r"Destination.*?Peak[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I | re.DOTALL)
        dnp = re.search(r"Non-Peak All[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)
        dr = re.search(r"Destination.*?reduced by \$([\d.]+)", text, re.I | re.DOTALL)
        mil = re.search(r"All Weights:[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)
        mld = re.search(r"Destination Services.*?All Weights:[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I | re.DOTALL)
        mg = re.findall(r"Mileage[^\$]*\$([\d.]+)[^\$]*\$([\d.]+)[^\$]*\$([\d.]+)", text, re.I)
        ob = re.search(r"On Base[^\$]*\$([\d.]+)", text, re.I)
        db = re.search(r"Density Bonus[^\$]*\$([\d.]+)", text, re.I)
        ul = re.search(r"Used[^/]*/[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)
        ov = re.search(r"Overflow[^\$]*\$(\d+)", text, re.I)
        sl = re.search(r"Security Seals[^\$]*\$(\d+)", text, re.I)
        rc = re.findall(r"Re-Coop[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)[^\$]*\$(\d+)", text, re.I)

        row = {
            "Filename": pdf_file,
            "Origin_Peak_MA": op.group(1) if op else "N/A",
            "Origin_Peak_VA": op.group(2) if op else "N/A",
            "Origin_Peak_FL": op.group(3) if op else "N/A",
            "Origin_NonPeak_MA": onp.group(1) if onp else "N/A",
            "Origin_NonPeak_VA": onp.group(2) if onp else "N/A",
            "Origin_NonPeak_FL": onp.group(3) if onp else "N/A",
            "Origin_Reduction": or_.group(1) if or_ else "N/A",
            "Dest_Peak_MA": dp.group(1) if dp else "N/A",
            "Dest_Peak_VA": dp.group(2) if dp else "N/A",
            "Dest_Peak_FL": dp.group(3) if dp else "N/A",
            "Dest_NonPeak_MA": dnp.group(1) if dnp else "N/A",
            "Dest_NonPeak_VA": dnp.group(2) if dnp else "N/A",
            "Dest_NonPeak_FL": dnp.group(3) if dnp else "N/A",
            "Dest_Reduction": dr.group(1) if dr else "N/A",
            "Mileage_MA": mg[0][0] if mg else "N/A",
            "Mileage_VA": mg[0][1] if mg else "N/A",
            "Mileage_FL": mg[0][2] if mg else "N/A",
            "OnBase_Fee": ob.group(1) if ob else "N/A",
            "Density_Bonus": db.group(1) if db else "N/A",
            "Used_Liftvans": ul.group(1) if ul else "N/A",
            "New_Liftvans": ul.group(2) if ul else "N/A",
            "Overflow_Boxes": ov.group(1) if ov else "N/A",
            "Seals": sl.group(1) if sl else "N/A",
            "ReCoop_MA": rc[0][0] if rc else "N/A",
            "Assembly_MA": rc[0][1] if rc else "N/A",
            "ReCoop_VA": rc[0][2] if rc else "N/A",
            "Assembly_VA": rc[0][3] if rc else "N/A",
            "ReCoop_FL": rc[0][4] if rc else "N/A",
            "Assembly_FL": rc[0][5] if rc else "N/A",
            "MIL_Origin_MA": mil.group(1) if mil else "N/A",
            "MIL_Origin_VA": mil.group(2) if mil else "N/A",
            "MIL_Origin_FL": mil.group(3) if mil else "N/A",
            "MIL_Dest_MA": mld.group(1) if mld else "N/A",
            "MIL_Dest_VA": mld.group(2) if mld else "N/A",
            "MIL_Dest_FL": mld.group(3) if mld else "N/A"
        }
        data.append(row)

    except Exception as e:
        print("  ERROR: " + str(e))
        data.append({"Filename": pdf_file})
        continue

df = pd.DataFrame(data)
df.to_excel(output_file, index=False)
print("SUCCESS! " + str(len(data)) + " vendors saved to " + output_file)
input("Press Enter to close...")

