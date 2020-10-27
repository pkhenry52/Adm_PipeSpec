from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, Paragraph, SimpleDocTemplate,\
     Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import PyPDF4
import re
import os


class LineDrw(Flowable):

    def __init__(self, width, height=0):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def __repr__(self):
        return "Line(w=%s)" % self.width

    def draw(self):
        self.canv.line(0, self.height, self.width, self.height)


class Report:

    def __init__(self, datatable, rptdata, Colnames, Colwdths, filename, ttl):
        # (table name, table data, table column names, table column
        # widths, name of PDF file)
        self.datatable = datatable
        self.rptdata = rptdata
        self.columnames = Colnames
        self.colms = Colwdths
        self.filename = filename
        self.ttl = ttl

        self.width, self.height = letter

        if datatable.find("_") != -1:
            self.tblname = (datatable.replace("_", " "))
        else:
            self.tblname = (' '.join(re.findall('([A-Z][a-z]*)', datatable)))

    def create_pdf(self):
        textAdjust = 6.5
        body = []

        styles = getSampleStyleSheet()
        spacer1 = Spacer(0, 0.25*inch)
        spacer2 = Spacer(0, 0.5*inch)

        tblstyle = TableStyle([
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])

        # provide the table name for the page header
        if self.ttl is not None:
            txt = self.tblname + ' requirements for ' + self.ttl
        else:
            txt = self.tblname
        ptext = '<font size=14>%s</font>' % txt
        body.append(Paragraph(ptext, styles["Heading2"]))

        line = LineDrw(500)

        if self.rptdata != []:
            body.append(line)
            body.append(spacer1)

            colwdth1 = []
            colwdth2 = []
            for i in self.colms[0]:
                colwdth1.append(i * textAdjust)
            for i in self.colms[1]:
                colwdth2.append(i * textAdjust)

            NumTbls = len(self.rptdata)

            pgbrk = 0
            for pg in range(0, NumTbls, 2):
                n = 0
                for i in self.columnames[0]:
                    ptxt = ("<b>{0}:</b>{1}<br/>".
                            format(i, self.rptdata[pg][n]))
                    body.append(Paragraph(ptxt, styles["Normal"]))
                    n += 1

            #    body.append(spacer2)

                rptdata2 = []
                rptdata2.append(list(self.columnames[1]))

                m = 0
                # remove the individual rows of data from data string
                for seg in range(len(self.rptdata[pg+1])//len(colwdth2)):
                    temp_data = list(self.rptdata[pg+1]
                                     [m * len(colwdth2):
                                     (m * len(colwdth2) + len(colwdth2))])
                    p = 0
                    for item in temp_data:
                        if type(item) == str:
                            if len(item) >= 25:
                                item = Paragraph(item, styles['Normal'])
                                temp_data[p] = item
                        p += 1

                    rptdata2.append(temp_data)

                    m += 1

                tbl2 = Table(rptdata2, colWidths=colwdth2)

                body.append(spacer2)

                tbl2.setStyle(tblstyle)
                body.append(tbl2)
                body.append(spacer2)

                pgbrk += 1

                if pgbrk % 2 == 0:
                    body.append(PageBreak())

        else:
            txt = ' have not been setup.'
            ptext = '<font size=14>%s</font>' % txt
            body.append(Paragraph(ptext, styles["Heading2"]))

        doc = SimpleDocTemplate(
            'tmp_rot_file.pdf', pagesize=landscape(letter),
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)

        doc.build(body)

        pdf_old = open('tmp_rot_file.pdf', 'rb')
        pdf_reader = PyPDF4.PdfFileReader(pdf_old)
        pdf_writer = PyPDF4.PdfFileWriter()

        for pagenum in range(pdf_reader.numPages):
            page = pdf_reader.getPage(pagenum)
            page.rotateCounterClockwise(90)
            pdf_writer.addPage(page)

        pdf_out = open(self.filename, 'wb')
        pdf_writer.write(pdf_out)
        pdf_out.close()
        pdf_old.close()
        os.remove('tmp_rot_file.pdf')
