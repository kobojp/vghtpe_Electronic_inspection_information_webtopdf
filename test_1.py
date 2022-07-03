
import pdfkit
import os

path_wkhtmltopdf = os.path.join('wkhtmltopdf.exe')
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

pdfkit.from_url('https://www.google.com/', '123.pdf', configuration=config)