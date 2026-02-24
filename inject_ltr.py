"""
inject_ltr.py  —  Post-processes a pandoc-generated DOCX to force LTR direction.
Usage: python inject_ltr.py <file.docx>
Adds <w:bidi w:val="0"/> to every paragraph in word/document.xml so that
Arabic/Hebrew installations of Word display the text left-to-right.
"""
import zipfile, os, re, sys

def inject(path):
    tmp = path + '.ltr_tmp'
    with zipfile.ZipFile(path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    xml = data.decode('utf-8')
                    # Add bidi=0 to every <w:pPr>
                    xml = xml.replace('<w:pPr>', '<w:pPr><w:bidi w:val="0"/>')
                    # Remove duplicates if it was already there
                    xml = re.sub(r'(<w:bidi w:val="0"/>){2,}', '<w:bidi w:val="0"/>', xml)
                    data = xml.encode('utf-8')
                zout.writestr(item, data)
    os.replace(tmp, path)

if __name__ == '__main__':
    inject(sys.argv[1])
