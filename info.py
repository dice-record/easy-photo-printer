from PyQt6.QtPrintSupport import QPrinter
for size in QPrinter.PaperSize:
    print(size, QPrinter.pageRect(QPrinter.PageRectOption.Device).size())