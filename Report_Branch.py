from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table,\
    TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch


class Report:
    def __init__(self, rows, colsz, ends, filename, ttl):
        # Developed report to print out a pdf of the allowed branch connections
        styles = getSampleStyleSheet()
        doc = []
        fontsz = colsz//3

        text = ('''<para align=center>Branch Connection Chart
                for {code}<br/><br/></para>'''.format(code=ttl))
        pt = Paragraph(text, style=styles["Heading2"])
        doc.append(pt)

        tblstyle = TableStyle([('INNERGRID', (0, 0), (-1, -1), 0.25,
                                colors.red),
                              ('BOX', (0, 0), (-1, -1), 0.4, colors.black),
                              ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                              ('FONT', (0, 0), (-1, -1), 'Helvetica', fontsz),
                               ])
        if rows != []:
            tbl = Table(rows, colWidths=colsz)
            tbl.setStyle(tblstyle)
            doc.append(tbl)

            spacer1 = Spacer(0, 0.1*inch)
            doc.append(spacer1)

            text = ('''<para align=center>Branch Connection Pipe OD
                    <br/><br/></para>''')
            px = Paragraph(text, style=styles["Normal"])
            doc.append(px)

            txte = ('''Allowed end connections for {code} are {ends}
                    <br/><br/>'''.format(code=ttl, ends=ends))
            pe = Paragraph(txte, style=styles["Normal"])
            doc.append(pe)

            txt = '''<b><u>LEGEND</u></b> <br/>
            ET - equal tee<br/>
            RT - reducing tee<br/>
            OL - O-Let (end connection and weight to be specified by \
                commodity property)<br/>
            BO - butt-on or set-on type fabrication, requires enginnering \
                stamped approval<br/>
            SW - sweep outlet requiring engineering design and approval<br/>
            Eng - special engineering designed connection excluding \
                Hot Tapes'''
            pl = Paragraph(txt, style=styles['Normal'])
            doc.append(pl)
        else:
            ptext = ('''<para align=center>has not been set up
                    <br/><br/></para>''')
            pt = Paragraph(ptext, style=styles["Heading2"])
            doc.append(pt)
        '''
        chrt = SimpleDocTemplate(
            filename, pagesize=landscape(letter),
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)'''

        chrt = SimpleDocTemplate(
            filename, pagesize=letter,
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)

        chrt.build(doc)
