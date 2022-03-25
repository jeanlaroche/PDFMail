#!/usr/bin/env python
"""
A very simple script that takes two png images that form a recto/version PDF document, and a list of addresses.
The script creates a PDF file that can be used for a mailing.
"""

import csv
try:
    from fpdf import FPDF
    import sys, fitz
except ImportError:
    print("ERROR: You must install fpdf with 'pip install fpdf' and pymupdf with 'pip install pymupdf' ")
    exit(0)

__author__ = "Jean Laroche"
__version__ = 1.0
__license__ = "GPL"

class PDF(FPDF):
    def __init__(self,recto,verso,numPerPage=1,margin=0,xAdjust=0,yAdjust=0,fontSizeAdjust=0):
        """
        :param recto:       path to the image file for the recto
        :param verso:       path to the image file for the verso
        :param numPerPage:  number of prints per page (1 or 2)
        :param margin:      margin, in inches
        """
        if numPerPage not in [1,2]:
            raise Exception("numPerPage must be 1 or 2")
        super().__init__(orientation = 'P' if numPerPage==2 else 'L', unit = 'in', format='letter')
        self.addresses = []
        self.sortByZip = 0
        self.margin = margin
        self.xAdjust = xAdjust
        self.yAdjust = yAdjust
        self.fontSizeAdjust = fontSizeAdjust
        self.recto = recto
        self.verso = verso
        self.numPerPage = numPerPage
        self.defFontSize = 15

    def setAddressList(self,csvFile,headerLines=1,sortByZip=0):
        """
        Sets the csv files for addresses
        :param csvFile:         csv file with addresses "Name","Street","City","State/Province","ZIP/Postal Code"
        :param headerLines:     number of header lines to skip in csv file.
        """
        self.sortByZip = sortByZip
        with open(csvFile, newline='') as csvfile:
            addresses = csv.reader(csvfile, delimiter=',', quotechar='|')
            allRows = []
            for ii,row in enumerate(addresses):
                if ii < headerLines: continue
                # Remove quotes and remove = sign.
                row = [r.replace('"','').replace("=","") for r in row]
                allRows.append(row)

            # Deal with instances where the name includes a \n!
            aR = []
            for ii,row in enumerate(allRows):
                if len(row) < 5:
                    # print(f"Correcting {row[0]}")
                    allRows[ii+1] = ["\n".join(row) + '\n'] + allRows[ii+1]
                else:
                    aR.append(row)
            allRows = aR

            # Now, sort the rows by zip code
            if self.sortByZip:
                allRows = sorted(allRows,key=lambda x: x[-1])

            # Create the address strings
            self.addresses = []
            for row in allRows:
                city = " ".join(row[-3:])
                self.addresses.append("\n".join(row[0:-3]+[city]))

    def addTwoPages(self,onlyVerso=0):
        """
        Adds two pages to the pdf file, one for the recto, one for the verso
        :param onlyVerso:   if 1, only add the verso page.
        """
        if not onlyVerso:
            self.add_page()
            # Sizes for the images.
            w = self.w-2*self.margin
            h = (self.h-2*self.margin)
            self.image(self.recto, x=self.margin, y=self.margin, w=w, h=h)
        self.add_page()
        # Sizes for the images.
        w = self.w-2*self.margin
        h = (self.h-2*self.margin)
        self.image(self.verso, x=self.margin, y=self.margin, w=w, h=h)

    def newPageOne(self,address,onlyVerso=0):
        """
        Adds two pages to the PDF using the address, for 1 print per page.
        :param address: string to be used for the address
        :param onlyVerso: if 1 only output verso
        """
        self.addTwoPages(onlyVerso)
        # Position of the address
        xAddress = .6*self.w + self.xAdjust
        yAddress =  self.h * .55 + self.yAdjust
        # Width of the address box.
        w = self.w-self.margin - .5 - xAddress

        self.set_xy(x = xAddress, y = yAddress)
        self.multi_cell(w = w, h = .25, txt= address, border = 0, align= 'L', fill= False)

    def newPageTwo(self,address1,address2,onlyVerso=0):
        """
        Adds two pages to the PDF, for 2 prints per page.
        :param address1: string to be used for the address on the top half
        :param address2: string to be used for the address on the bottom half
        :param onlyVerso: if 1 only output verso
        """
        self.addTwoPages(onlyVerso)
        # Position of the first address
        xAddress = .6*self.w + self.xAdjust
        yAddress =  self.h * .3 +  self.yAdjust
        # Width of the address box.
        w = self.w-self.margin - .5 - xAddress

        self.set_xy(x = xAddress, y = yAddress)
        self.multi_cell(w = w, h = .2, txt= address1, border = 0, align= 'L', fill= False)

        # Position of the second address
        yAddress += self.h/2
        self.set_xy(x = xAddress, y = yAddress)
        self.multi_cell(w = w, h = .2, txt= address2, border = 0, align= 'L', fill= False)

    def createPDF(self,outFile,numPages=1000,testMode=0):
        """
        Create the PDF document
        :param outFile:     Name of output pdf file
        :param numPages:    Maximum number of pages to create
        :param testMode:    If 1, sort pages by length of output address
        """
        self.set_font('Times', '', self.defFontSize+self.fontSizeAdjust if self.numPerPage == 1 else self.defFontSize-2+self.fontSizeAdjust)
        self.set_margins(left=self.margin, top=self.margin, right=self.margin)

        addresses = self.addresses
        if testMode:
            # Sort addresses by decreasing address length
            addresses = sorted(addresses, key=lambda x: -max([len(y) for y in x.split('\n')]))
        if self.numPerPage == 2:
            # Add an extra empty address if we don't have an even number
            if len(addresses) % 2: addresses.append("")
            L = len(addresses)//2
            if self.sortByZip and not testMode:
                # This is so when the pages are physically cut in half, the top halves are in consecutive zip order, followed by
                # the bottom halves
                addresses = list(zip(addresses[0:L], addresses[L:]))
            else:
                addresses = list(zip(addresses[0:2 * L:2], addresses[1:2 * L:2]))

        _numPages = min(numPages, len(addresses))
        if self.numPerPage == 1:
            for address in addresses[0:numPages]:
                self.newPageOne(address,testMode)
        else:
            for address1,address2 in addresses[0:numPages]:
                self.newPageTwo(address1,address2,testMode)
        self.output(outFile, 'F')
        print(f"Output pdf written to {outFile}")

def extractImgsFromPDF(file):
    """
    Extract Recto and Verso from pdf file as images.
    :param file: input pdf file.
    """
    print(f"Extracting recto and verso from {file}")
    doc = fitz.open(file)
    if doc.page_count < 2:
        raise Exception("Input PDF document must have at least two pages, one for recto, one for verso")
    page = doc[0]
    pix = page.get_pixmap(dpi=600)
    pix.save("recto_.png")
    page = doc[1]
    pix = page.get_pixmap(dpi=600)
    pix.save("verso_.png")
    return "recto_.png","verso_.png"

def doit(args):
    """
    Create the PDF.
    :param args: obtained by argparse.
    """
    if not args.inPDFFile:
        recto,verso,csvFile=args.args
    else:
        csvFile = args.args[0]
        recto,verso = extractImgsFromPDF(args.inPDFFile)

    pdf = PDF(recto,verso,args.npp,margin=args.margin,xAdjust=args.x,yAdjust=args.y,fontSizeAdjust=args.f)
    pdf.setAddressList(csvFile,headerLines=args.skip,sortByZip=args.sort)
    pdf.createPDF(outFile=args.outFile,numPages=args.np,testMode=args.test)


if __name__ == '__main__':
    import argparse
    desc = '''
Creates a mailing PDF by merging a pair of recto/verso images with an address list
See https://pyfpdf.readthedocs.io/en/latest/reference/image/index.html for supported image formats
The csv file is expected to have the following fields: "Name","Street","City","State","ZIP code" separated by commas
If npp is 1, the page is in landscape format and there is page for the recto and one page for the verso.
If npp is 2, the page is in portrait format and there are two rectos on one page and two versos on the next.
The -x and -y flags can be used to fine tune the position of the address.
The -f flag can be used to fine tune font size.
The -test flag is useful for outputting the verso pages only with the longest addresses first
The -i flag can be used to first extract the recto and verso pages from an existing PDF. In that case, you can ommit
the recto and verso inputs.
Example:
PDFMail.py -skip 1 -npp 2 -o mail.pdf Recto.png Verso.png addresses.csv
To extract the recto/verso images from an existing pdf:
PDFMail.py -skip 1 -npp 2 -o mail.pdf -i input.pdf addresses.csv
To decrease the font size by 1
PDFMail.py -skip 1 -f -1 -npp 2 -o mail.pdf Recto.png Verso.png addresses.csv
'''
    parser = argparse.ArgumentParser('PDFMail', description=desc, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-npp', type=int, default=1, help='Number of prints per page: 1 or 2 (default 1)')
    parser.add_argument('-sort', action='store_true', help='Sort zip codes')
    parser.add_argument('-test', action='store_true', help='Output a test pdf with only verso pages and the longest addresses first')
    parser.add_argument('-np', type=int, default=10000, help='Number of pages to print')
    parser.add_argument('-skip', type=int, default=1, help='Number of lines to skip at start of csv file (default 1)')
    parser.add_argument('-margin', type=float, default=0, help='Margins in inch (default 0)')
    parser.add_argument('-x', type=float, default=0, help='Adjustment for horizontal position of address, in +/-inch (default 0)')
    parser.add_argument('-y', type=float, default=0, help='Adjustment for vertical position of address, in +/-inch (default 0)')
    parser.add_argument('-f', type=int, default=0, help='Adjustment for font size (default 0')
    parser.add_argument('-o', dest='outFile', default="output.pdf", help='Name of output file (default output.pdf)')
    parser.add_argument('-i', dest='inPDFFile', help='Optional input pdf file which will be used to extract the recto and verso)')
    parser.add_argument('args', metavar='recto verso csvFile', nargs=argparse.REMAINDER, help='recto verso csvFile or csvFile if the -i option is used')
    args = parser.parse_args()
    doit(args)

