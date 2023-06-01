from docxtpl import DocxTemplate, InlineImage
import datetime as dt
from docx2pdf import convert

# create a document object
doc = DocxTemplate("template/inviteTmpl.docx")

# The current time
currentDateAndTime = dt.datetime.now()
currentTime = currentDateAndTime.strftime("%H:%M:%S")

# create context dictionary
context = {
    "todayStr": dt.datetime.now().strftime("%d-%b-%Y"),
    "recipientName": "Chaitanya",
    "currentTime": currentTime,
    "venueStr": "the beach",
    "senderName": "Sanket",
}

# inject image into the context
context["bannerImg"] = InlineImage(doc, "imgs/party_banner_0.png")

# render context into the document object
doc.render(context)

# save the document object as a word file
doc.save("test/invitation.docx")

# convert word file to a pdf file - Have word installed
convert("test/invitation.docx", "test/invitation.pdf")
