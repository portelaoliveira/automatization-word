from docxtpl import DocxTemplate, InlineImage
import datetime as dt
from docx2pdf import convert

# template word file path
tmplPath = "template/inviteTmpl.docx"

personNames = [
    "Aakav",
    "Aakesh",
    "Aarav",
    "Advik",
    "Chaitanya",
    "Chandran",
    "Darsh",
]

# run for each person in a for loop
for pItr, p in enumerate(personNames):
    # create a document object
    doc = DocxTemplate(tmplPath)

    # create context dictionary
    context = {
        "todayStr": dt.datetime.now().strftime("%d-%b-%Y"),
        "recipientName": p,
        "evntDtStr": "21-Oct-2021",
        "venueStr": "the beach",
        "senderName": "Sanket",
    }

    # inject image into the context
    bannerImgPath = f"imgs/party_banner_{pItr % 3}.png"
    imgObj = InlineImage(doc, bannerImgPath)
    context["bannerImg"] = imgObj

    # render context into the document object
    doc.render(context)

    # save the document object as a word file
    resultFilePath = f"test/invitation_{pItr}.docx"
    doc.save(resultFilePath)

    # convert the word file into pdf
    convert(resultFilePath, resultFilePath.replace(".docx", ".pdf"))

print("execution complete...")
