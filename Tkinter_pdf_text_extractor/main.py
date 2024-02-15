import tkinter, PyPDF2
from tkinter import filedialog


def openFile():
    filename = filedialog.askopenfilename(title="Open Pdf File",
                                          initialdir=r'C:\Users\DELL\Desktop\Lotus pics\logo\lotus',
                                          filetypes=[('PDF files', '*.pdf')])
    
    filename_label.configure(text= filename)
    outputfile_text.delete("1.0", tkinter.END)
    reader = PyPDF2.PdfReader(filename)
    for page_number in range(len(reader.pages)):
        current_text = reader.pages[page_number].extract_text()
        outputfile_text.insert(tkinter.END,current_text )
    

root = tkinter.Tk()
root.title("PDF Text Extractor")


filename_label = tkinter.Label(root, text="No File Selected")
outputfile_text = tkinter.Text(root)
openfile_button = tkinter.Button(root, text="Open Pdf File", command=openFile)



filename_label.pack()
outputfile_text.pack()
openfile_button.pack()


root.mainloop()