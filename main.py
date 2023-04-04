import os
import docx
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

class PDFConverter:
    def __init__(self, master):
        self.pdf_file = None
        
        # Create the GUI widgets
        self.master = master
        self.master.title('PDF to Word Converter')
        self.master.geometry('1366x768+0+0')
        self.master.configure(bg='#23272a')
        
        # Create the label
        self.label = tk.Label(master, text='PDF to Word Converter', font=('Arial', 24), fg='#ffffff', bg='#23272a')
        self.label.pack(pady=(30,10))
        
        # Create the "Choose PDF File" button
        self.choose_button = tk.Button(master, text='Choose PDF File', font=('Arial', 16), bg='#6ab04c', fg='#ffffff', activebackground='#578e3e', activeforeground='#ffffff', padx=15, pady=10, borderwidth=0, command=self.choose_file)
        self.choose_button.pack(pady=20)
        
        # Create the "Convert to Word" button
        self.convert_button = tk.Button(master, text='Convert to Word', font=('Arial', 16), bg='#1e88e5', fg='#ffffff', activebackground='#1565c0', activeforeground='#ffffff', padx=15, pady=10, borderwidth=0, command=self.convert_to_word, state=tk.DISABLED)
        self.convert_button.pack(pady=20)
        
        # Create the status label
        self.status_label = tk.Label(master, text='', font=('Arial', 14), fg='#ffffff', bg='#23272a')
        self.status_label.pack(pady=10)

        # Create the view menu
        self.view_menu = tk.Menu(self.master)
        self.master.config(menu=self.view_menu)
        
        # Add the "Fullscreen" option to the view menu
        self.view_menu.add_command(label="Fullscreen", command=self.toggle_fullscreen)
        
        # Add the "Windowed" option to the view menu
        self.view_menu.add_command(label="Windowed", command=self.toggle_windowed)
        
        # Set the windowed flag to True
        self.windowed = True
    
    def toggle_fullscreen(self):
        # Toggle fullscreen mode
        self.master.attributes('-fullscreen', True)
        self.windowed = False
    
    def toggle_windowed(self):
        # Toggle windowed mode
        self.master.attributes('-fullscreen', False)
        self.windowed = True


    def choose_file(self):
        # Prompt the user to choose a PDF file
        self.pdf_file = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
        if self.pdf_file is not None:
            self.status_label.config(text=f'Selected file: {os.path.basename(self.pdf_file)}')
            self.convert_button.config(state=tk.NORMAL)
    
    def convert_to_word(self):
        if self.pdf_file is not None:
            # Convert the PDF file to Word format
            try:
                # Prompt the user to choose a destination location
                destination = filedialog.asksaveasfilename(defaultextension='.docx',
                                                           filetypes=[('Microsoft Word Document', '*.docx')])
                if destination is not None:
                    # Create a Word document
                    doc = docx.Document()
                    # Open the PDF file
                    with open(self.pdf_file, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        # Loop through each page of the PDF and add it to the Word document
                        for page_num in range(len(pdf_reader.pages)):
                            # Get the page
                            page = pdf_reader.pages[page_num]
                            # Extract text from the page
                            text = page.extract_text()
                            # Add the text to the Word document
                           

                            doc.add_paragraph(text)
                    # Save the Word document
                    doc.save(destination)
                    self.status_label.config(text=f'Conversion successful. Word file saved as {destination}.')
            except Exception as e:
                messagebox.showerror(title='Conversion error', message=str(e))
        else:
            messagebox.showwarning(title='No PDF file selected', message='Please select a PDF file to convert.')

if __name__ == '__main__':
    root = tk.Tk()
    converter = PDFConverter(root)
    root.mainloop()
