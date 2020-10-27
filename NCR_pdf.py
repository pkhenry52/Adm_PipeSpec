
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Flowable, Paragraph, Frame, Spacer
from reportlab.lib.enums import TA_LEFT, TA_RIGHT


class LineDrw(Flowable):

    def __init__(self, width, height=0):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def __repr__(self):
        return "Line(w=%s)" % self.width

    def draw(self):
        self.canv.line(0, self.height, self.width, self.height)


class NCRForm:
    def __init__(self, pdf_file, t_data, ta_data, chk_data=None):
        self.t = t_data
        self.ta = ta_data
        self.chk = chk_data
        self.pdf_file = pdf_file

    def create(self):
        styles = getSampleStyleSheet()

        aW, aH = letter

        nrm = styles['Normal']
        hd1 = styles['Heading1']
        hd2 = styles['Heading2']

        styles.add(ParagraphStyle(name='Right', fontSize=12, spaceAfter=15,
                                  alignment=TA_RIGHT))
        styles.add(ParagraphStyle(name='Left', fontSize=12, spaceAfter=15,
                                  borderColor='black', borderRadius=3,
                                  borderPadding=5, borderWidth=1,
                                  alignment=TA_LEFT))

        basepg = Canvas(self.pdf_file)

        # Form title
        ttl = []
        ttl.append(Paragraph('Nonconformance Report', hd1))
        ttl.append(LineDrw(500))
        frm = Frame(15*mm, 260*mm, width=150*mm, height=25*mm, showBoundary=0)
        frm.addFromList(ttl, basepg)

        # First text area box
        dscrp = []
        dscrp.append(Paragraph('Project Description', hd2))
        dscrp.append(Paragraph(self.ta['ta_1'], nrm))
        frm1 = Frame(15*mm, 240*mm, width=180*mm, height=25*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # Section of labels for input text
        info = []
        info.append(Paragraph('<b><i>Job Number:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Requestor:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Contact Information:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Fabricator:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Date Of Issue:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Pipe Specification Code:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>P&ID Line Number:</i></b>',
                              style=styles['Right']))
        frmLft = Frame(15*mm, 150*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmLft.addFromList(info, basepg)

        # list of first section of input text boxes
        inpt = []
        for n in range(1, 8):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
        frmRgt = Frame(80*mm, 150*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmRgt.addFromList(inpt, basepg)

        # second textarea box
        dscrp = []
        dscrp.append(Paragraph('Related Drawings or Documentation:', hd2))
        dscrp.append(Paragraph(self.ta['ta_2'], nrm))
        frm1 = Frame(15*mm, 120*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # third textarea box
        dscrp = []
        dscrp.append(Paragraph('Nonconformance Description:', hd2))
        dscrp.append(Paragraph(self.ta['ta_3'], nrm))
        frm1 = Frame(15*mm, 70*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # fourth textarea box
        dscrp = []
        dscrp.append(Paragraph('Proposed Repair', hd2))
        dscrp.append(Paragraph(self.ta['ta_4'], nrm))
        frm1 = Frame(15*mm, 30*mm, width=180*mm, height=35*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        basepg.showPage()

        # check box labels
        info = []
        info.append(Paragraph('<b><i>Has QC Manager been notified?</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Has Operations been notified?</i></b>',
                              style=styles['Right']))
        frmLft1 = Frame(15*mm, 200*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmLft1.addFromList(info, basepg)

        # check box checked values
        inpt = []
        txt = 'No'
        if 'chk_1' in self.chk:
            txt = 'Yes'
        inpt.append(Paragraph(txt, style=styles['Left']))
        inpt.append(Spacer(1, 13))
        txt = 'No'
        if 'chk_2' in self.chk:
            txt = 'Yes'
        inpt.append(Paragraph(txt, style=styles['Left']))
        frmLft2 = Frame(80*mm, 197*mm, width=15*mm, height=65*mm,
                        showBoundary=0)
        frmLft2.addFromList(inpt, basepg)

        # fifth textarea box
        dscrp = []
        dscrp.append(Paragraph('Final Repair or Disposition:', hd2))
        dscrp.append(Paragraph('(indicate schedule for repair)', nrm))
        dscrp.append(Paragraph(self.ta['ta_5'], nrm))
        frm1 = Frame(15*mm, 190*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # text box labels
        info = []
        info.append(Paragraph('<b><i>QC Manager Approval</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>Operations Approval</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>Requestor Approval</i></b>',
                    style=styles['Right']))
        frmLft1 = Frame(10*mm, 100*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmLft1.addFromList(info, basepg)

        # input text data
        inpt = []
        for n in range(8, 11):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmLft2 = Frame(80*mm, 100*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmLft2.addFromList(inpt, basepg)

        # text box labels
        info = []
        for n in range(3):
            info.append(Paragraph('<b><i>Date</i></b>', style=styles['Right']))
            info.append(Spacer(1, 13))
        frmRgt1 = Frame(95*mm, 100*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmRgt1.addFromList(info, basepg)

        # text box input data
        inpt = []
        for n in range(11, 14):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmRgt2 = Frame(160*mm, 100*mm, width=35*mm, height=65*mm,
                        showBoundary=0)
        frmRgt2.addFromList(inpt, basepg)

        basepg.save()

'''
if __name__ == '__main__':
    NCRForm().create()
'''
