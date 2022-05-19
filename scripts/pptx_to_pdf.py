import win32com.client
import glob


def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    print("Converting PPT to PDF...")
    print("Input file: " + inputFileName)
    print("Output file: " + outputFileName)

    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
        
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)
    deck.Close()
    powerpoint.Quit()

def convert(path):
    pptxs = glob.glob(f"{path}\*.pptx")

    for pptx in pptxs:
        PPTtoPDF(inputFileName = pptx, outputFileName = pptx[:pptx.rfind('.')] + '.pdf')
