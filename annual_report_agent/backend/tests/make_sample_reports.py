"""Generate two small fictional annual reports (FY2022, FY2023) for end-to-end tests.

Each report has prior-year comparative columns, N'000 scale, a directors list, a
scanned (image-only) page and filler pages so the page locator has to work.
"""
import io
import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw, ImageFont

DATA = {
    2022: dict(prev=2021, revenue=("1,845,220", "1,602,114"), pbt=("412,908", "350,771"),
               pat=("298,450", "251,019"), assets=("5,220,931", "4,870,002"),
               liab=("3,101,447", "2,990,385"), equity=("2,119,484", "1,879,617"),
               employees=("2,410", "2,305"),
               directors=[("Adaeze Okafor", "Chairman", "Non-Executive", "Female"),
                          ("Tunde Bakare", "Chief Executive Officer", "Executive", "Male"),
                          ("Ngozi Eze", "Director", "Independent Non-Executive", "Female")],
               scope1=("18,450", "19,920"), scope2=("42,300", "44,010")),
    2023: dict(prev=2022, revenue=("2,104,876", "1,845,220"), pbt=("(35,214)", "412,908"),
               pat=("(61,330)", "298,450"), assets=("5,902,117", "5,220,931"),
               liab=("3,866,941", "3,101,447"), equity=("2,035,176", "2,119,484"),
               employees=("2,512", "2,410"),
               directors=[("Adaeze Okafor", "Chairman", "Non-Executive", "Female"),
                          ("Tunde Bakare", "Chief Executive Officer", "Executive", "Male"),
                          ("Ngozi Eze", "Director", "Independent Non-Executive", "Female"),
                          ("Ibrahim Musa", "Chief Financial Officer", "Executive", "Male")],
               scope1=("17,880", "18,450"), scope2=("40,120", "42,300")),
    # FY2024: long report (60+ pages) with amounts in MILLIONS -> tests page locating and scale
    2024: dict(prev=2023, revenue=("2,488.6", "2,104.9"), pbt=("301.2", "(35.2)"),
               pat=("208.7", "(61.3)"), assets=("6,410.3", "5,902.1"),
               liab=("4,120.8", "3,866.9"), equity=("2,289.5", "2,035.2"),
               employees=("2,630", "2,512"),
               directors=[("Adaeze Okafor", "Chairman", "Non-Executive", "Female"),
                          ("Ibrahim Musa", "Chief Executive Officer", "Executive", "Male"),
                          ("Ngozi Eze", "Director", "Independent Non-Executive", "Female"),
                          ("Kemi Adeyemi", "Director", "Independent Non-Executive", "Female")],
               scope1=("16,940", "17,880"), scope2=("38,870", "40,120"),
               unit="Naira millions (N'm)", padding=50),
}

FILLER = ("Our strategy focuses on customer growth, digital channels and operational efficiency. "
          "We continued to invest in people and technology across our markets. ") * 12


def page(doc, text):
    p = doc.new_page(width=595, height=842)
    p.insert_textbox(pymupdf.Rect(50, 50, 545, 800), text, fontsize=10, fontname="helv")
    return p


def scanned_page(doc, lines):
    img = Image.new("RGB", (1240, 1754), "white")
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 34)
    except OSError:
        font = ImageFont.load_default(size=34)
    y = 120
    for ln in lines:
        d.text((100, y), ln, fill="black", font=font)
        y += 60
    buf = io.BytesIO()
    img.save(buf, "PNG")
    p = doc.new_page(width=595, height=842)
    p.insert_image(p.rect, stream=buf.getvalue())


def build(year: int, out: Path) -> None:
    d = DATA[year]
    py = d["prev"]
    doc = pymupdf.open()
    page(doc, f"ACME TELECOMS PLC\n\nANNUAL REPORT {year}\n\nFor the year ended 31 December {year}\n\n"
              "Connecting lives across Nigeria")
    page(doc, "CHAIRMAN'S STATEMENT\n\n" + FILLER)
    page(doc, "BOARD OF DIRECTORS\n\n" + "\n\n".join(
        f"{n}\n{role} ({kind})\nGender: {g}\nNationality: Nigerian" for n, role, kind, g in d["directors"]))
    page(doc, "OUR PEOPLE\n\n" + FILLER[:600] +
         f"\n\nTotal number of employees at 31 December {year}: {d['employees'][0]} "
         f"({py}: {d['employees'][1]})")
    page(doc, "RISK MANAGEMENT\n\n" + FILLER)
    for i in range(d.get("padding", 0)):   # long narrative section, like a real report
        page(doc, f"OPERATING REVIEW - SECTION {i + 1}\n\n" + FILLER)
    unit = d.get("unit", "Naira thousands (N'000)")
    page(doc, "CONSOLIDATED STATEMENT OF PROFIT OR LOSS\nFor the year ended 31 December "
              f"{year}\nAll amounts in {unit}\n\n"
              f"                                   {year}            {py}\n"
              f"Revenue                         {d['revenue'][0]}      {d['revenue'][1]}\n"
              f"Profit/(loss) before tax        {d['pbt'][0]}        {d['pbt'][1]}\n"
              f"Profit/(loss) for the year      {d['pat'][0]}        {d['pat'][1]}\n")
    page(doc, "CONSOLIDATED STATEMENT OF FINANCIAL POSITION\nAs at 31 December "
              f"{year}\nAll amounts in {unit}\n\n"
              f"                                   {year}            {py}\n"
              f"Total assets                    {d['assets'][0]}      {d['assets'][1]}\n"
              f"Total liabilities               {d['liab'][0]}      {d['liab'][1]}\n"
              f"Total equity                    {d['equity'][0]}      {d['equity'][1]}\n")
    page(doc, "NOTES TO THE FINANCIAL STATEMENTS\n\n" + FILLER)
    scanned_page(doc, ["SUSTAINABILITY REPORT - GHG EMISSIONS",
                       f"Greenhouse gas emissions (tCO2e)       {year}      {py}",
                       f"Scope 1 (direct)                      {d['scope1'][0]}    {d['scope1'][1]}",
                       f"Scope 2 (purchased electricity)       {d['scope2'][0]}    {d['scope2'][1]}",
                       "Boundary: Acme Telecoms Plc and subsidiaries",
                       "Methodology: GHG Protocol Corporate Standard"])
    page(doc, "SHAREHOLDER INFORMATION\n\n" + FILLER[:900])
    doc.save(out)


if __name__ == "__main__":
    out_dir = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent / "fixtures")
    out_dir.mkdir(parents=True, exist_ok=True)
    for y in DATA:
        build(y, out_dir / f"Acme_Telecoms_Annual_Report_{y}.pdf")
    print("written to", out_dir)
