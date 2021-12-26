import sys, getopt
import openpyxl
import datetime
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML


def main():
    argv=sys.argv[1:]
    invoicenum = ''
    try:
        opts, args = getopt.getopt(argv,"i:",["invoicenum="])
    except getopt.GetoptError:
        print('test.py -i <invoicenum>')
        sys.exit(2)
    for opt, arg in opts:
        if opt in ("-i", "--invoicenum"):
            invoicenum = arg

    xlsx = openpyxl.load_workbook(filename='/Users/zshorvat/OneDrive/Atos/Consulting/kiadas-bevetel.xlsm', data_only=True)

    timetracker = xlsx['Time tracker']
    for row in timetracker.rows:
        if row[0].value == invoicenum:
            break

    invoicedate_excel = datetime.datetime.strptime(str(row[1].value),"%Y-%m-%d %H:%M:%S")
    invoicedate_pdf = str(invoicedate_excel.day)+'/'+str(invoicedate_excel.month)+'/'+str(invoicedate_excel.year)
    env = Environment(loader=FileSystemLoader('/Users/zshorvat/OneDrive/Atos/Consulting/'))
    template = env.get_template("invoicetemplate.html")

    template_vars = {"invoicenum" : row[0].value,
                     "ordernum": row[5].value,
                     "invoicedate": invoicedate_pdf,
                     "hours": str(row[6].value).replace('.',','),
                     "description": row[4].value,
                     "unitprice": str("%.2f" % row[7].value).replace('.',','),
                     "subtotal": str("%.2f" % row[8].value).replace('.',','),
                     "vat": str("%.2f" % row[9].value).replace('.',','),
                     "total": str("%.2f" % row[10].value).replace('.',',')}

    html_out = template.render(template_vars)

    HTML(string=html_out).write_pdf(f"/Users/zshorvat/OneDrive/Atos/Consulting/Invoices/{row[0].value.replace('/','-')}-Zsolt Horvath-{row[5].value}.pdf")

if __name__ == "__main__":
    main()
