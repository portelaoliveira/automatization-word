import datetime as dt
import random
from docx2pdf import convert
import matplotlib.pyplot as plt
from docxtpl import DocxTemplate, InlineImage

# create a document object
doc = DocxTemplate("template/reportTmpl.docx")

# create data for reports
salesTblRows = []
for k in range(10):
    costPu = random.randint(1, 15)
    nUnits = random.randint(100, 500)
    salesTblRows.append(
        {
            "sNo": k + 1,
            "name": "Item " + str(k + 1),
            "cPu": costPu,
            "nUnits": nUnits,
            "revenue": costPu * nUnits,
        }
    )

topItems = [
    x["name"]
    for x in sorted(salesTblRows, key=lambda x: x["revenue"], reverse=True)
][0:3]

todayStr = dt.datetime.now().strftime("%d-%b-%Y")

# create context to pass data to template
context = {
    "reportDtStr": todayStr,
    "salesTblRows": salesTblRows,
    "topItemsRows": topItems,
}

# inject image into the context
fig, ax = plt.subplots()
ax.bar([x["name"] for x in salesTblRows], [x["revenue"] for x in salesTblRows])
fig.tight_layout()
fig.savefig("imgs/trendImg.png")
context["trendImg"] = InlineImage(doc, "imgs/trendImg.png")

# render context into the document object
doc.render(context)

# save the document object as a word file
reportWordPath = f"test/report_{todayStr}.docx"
doc.save(reportWordPath)

# convert the word file as pdf file
convert(reportWordPath, reportWordPath.replace(".docx", ".pdf"))
