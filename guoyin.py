import sys, string
from dragonmapper.hanzi import to_zhuyin
from docx import Document

argv = sys.argv
text = ""
document_path = argv[-1]
document = Document(document_path)

text = "".join([para.text+'p' for para in document.paragraphs]) # p = paragraph end

ignore_chars = string.printable + " ，、…。！？：；‘’“”（）《》【】“”！"
escape_sequences = ["\n", "\t", "\r", "\b", "\f", "\v", "\0"]
zhuyin = ""
for char in text:
    if char not in ignore_chars and char not in escape_sequences:
        zhuyin += char + " -> " + to_zhuyin(char) + "\n"
        print(to_zhuyin(char))
    else:
        print("x")
        zhuyin += "x\n"

with open(document_path+"-guoyin.txt", "w") as output_file:
    output_file.write(zhuyin)

sys.exit()