from docx2pdf import convert
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os
from random import randint
import webbrowser
import PyPDF2
import platform


# Define class for creating os_window.
class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Word_to_PDF Converter.")
        self.config(bg='#FFBD73')  # Set the background of the window
        self.resizable(width=False, height=False)  # Disable resizing

        # Cross-platform fullscreen logic
        if platform.system() == "Windows":
            self.state('zoomed')  # Fullscreen for Windows
        else:
            self.attributes('-fullscreen', True)  # Fullscreen for Linux/macOS

        # Styling
        self.style = ttk.Style(self)
        self.style.configure('TButton',
                             font=('Arial', 12, 'bold'),
                             padding=15,
                             relief='flat',
                             background='black',
                             foreground='#3B1E54')
        self.style.configure('TLabel',
                             font=('Arial', 14, 'bold'),
                             foreground='black')
        self.style.configure('TLabelFrame',
                             font=('Arial', 14, 'bold'),
                             foreground='red',
                             background='#f7f7f7',
                             padding=20)
        self.style.configure('TEntry',
                             font=('Arial', 12),
                             padding=10)

        # Centered frame for file selection and conversion options
        self.main_frame = ttk.Frame(self, padding="20")
        self.main_frame.place(relx=0.5, rely=0.5, anchor=CENTER)  # Center the frame in the window

        # Label for file selection
        self.lableFrame = ttk.LabelFrame(self.main_frame, text="Select your Word File", relief="groove")
        self.lableFrame.grid(column=1, row=1, padx=20, pady=20, sticky=N + S + E + W)

        # Centered label at the bottom for credits
        self.lable = ttk.Label(self.main_frame, text="Developed by Sushant", relief="flat", style="TLabel")
        self.lable.grid(row=8, column=1, pady=10, sticky=S)

        # Buttons and inputs
        self.button()
        self.make_dir()
        self.button1()
        self.button2()

        # Password Entry Field
        self.password_label = ttk.Label(self.lableFrame, text="Enter Password for PDF:", style="TLabel")
        self.password_label.grid(column=1, row=3, padx=20, pady=5)

        self.password_entry = ttk.Entry(self.lableFrame, show="*")
        self.password_entry.grid(column=1, row=4, padx=20, pady=5)

        # Show/Hide Password Toggle Button
        self.show_password = False
        self.toggle_button = ttk.Button(self.lableFrame, text="Show", command=self.toggle_password)
        self.toggle_button.grid(column=2, row=4, padx=10, pady=5)

        # Metadata Display Field
        self.metadata_label = ttk.Label(self.lableFrame, text="PDF Metadata:", style="TLabel")
        self.metadata_label.grid(column=1, row=5, padx=20, pady=5)

        self.metadata_text = Text(self.lableFrame, height=8, width=60, font=('Arial', 12), wrap=WORD)
        self.metadata_text.grid(column=1, row=6, padx=20, pady=5)

    def toggle_password(self):
        if self.show_password:
            self.password_entry.config(show="*")
            self.toggle_button.config(text="Show")
        else:
            self.password_entry.config(show="")
            self.toggle_button.config(text="Hide")
        self.show_password = not self.show_password

    def button(self):
        self.button = ttk.Button(self.lableFrame, text="Browse a File", command=self.fileDialog)
        self.button.grid(column=1, row=1)

    def button1(self):
        self.button1 = ttk.Button(self.lableFrame, text="Convert File", command=self.convert)
        self.button1.grid(column=1, row=2, padx=20, pady=50)

    def button2(self):
        self.button2 = ttk.Button(self.lableFrame, text="Open Output Folder", command=self.open_folder)
        self.button2.grid(column=1, row=7, padx=20, pady=20)

    def fileDialog(self):
        self.filename = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="Select a File",
            filetype=(("docx files", "*.docx"), ("All Files", "*.*"))
        )
        if self.filename:
            print(f"Selected file: {self.filename}")

    def make_dir(self):
        self.output_path = os.path.join(os.getcwd(), "Doc_2_PDF_Output")
        os.makedirs(self.output_path, exist_ok=True)

    def convert(self):
        if not hasattr(self, 'filename') or not self.filename:
            print("No file selected. Please select a file.")
            return

        i = str(randint(1, 1000))
        self.output_file = os.path.join(self.output_path, f'output_{i}.pdf')

        convert(self.filename, self.output_file)
        password = self.password_entry.get()

        if password:
            self.apply_password(self.output_file, password)
        else:
            print("No password provided. File is not encrypted.")

        self.show_metadata(self.output_file, password)
        webbrowser.open(self.output_file)

    def apply_password(self, pdf_file, password):
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_writer = PyPDF2.PdfWriter()

            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

            pdf_writer.encrypt(password)

            with open(pdf_file, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)
        except Exception as e:
            print(f"Error applying password: {e}")

    def show_metadata(self, pdf_file, password):
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            if pdf_reader.is_encrypted:
                pdf_reader.decrypt(password)

            metadata = pdf_reader.metadata
            self.metadata_text.delete(1.0, END)

            if metadata:
                for key, value in metadata.items():
                    self.metadata_text.insert(END, f"{key}: {value}\n")
            else:
                self.metadata_text.insert(END, "No metadata available.")
        except Exception as e:
            print(f"Error displaying metadata: {e}")

    def open_folder(self):
        os.startfile(self.output_path)


if __name__ == '__main__':
    root = Root()
    root.mainloop()
