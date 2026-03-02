import tkinter as tk
from tkinter import filedialog, messagebox, font, colorchooser, simpledialog, Menu, Text, Frame, Scrollbar

class MettarolWord:
    def __init__(self, root):
        root.geometry('800x600')
        self.root = root
        self.root.title("Mettarol Word")
        self.text_area = Text(self.root, wrap='word', font=("Arial", 12))
        self.text_area.pack(expand=True, fill='both')
        
        self.is_modified = False  # Variable to track changes
        self.current_file = None  # Name of the opened file

        self.create_menu()

        self.scrollbar = Scrollbar(self.text_area)
        self.scrollbar.pack(side='right', fill='y')
        self.text_area.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.text_area.yview)

        # Event handler for text modification
        self.text_area.bind('<Key>', self.mark_as_modified)

        # Link window close event to check function
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def mark_as_modified(self, event):
        """Marks document as modified."""
        self.is_modified = True

    def on_close(self):
        """Checks file state before closing program."""
        if self.is_modified:
            result = messagebox.askyesno(
                title="Warning!",
                message="The file has not been saved. Do you want to save it?"
            )
            if result:
                self.save_file()
        self.root.destroy()

    def create_menu(self):
        menu = Menu(self.root)
        self.root.config(menu=menu)

        file_menu = Menu(menu)
        menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open", command=self.open_file)
        file_menu.add_command(label="Save", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)

        edit_menu = Menu(menu)
        menu.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Make Bold", command=self.bold_text)
        edit_menu.add_command(label="Underline", command=self.underline_text)
        edit_menu.add_command(label="Italicize", command=self.italic_text)
        edit_menu.add_command(label="Change Text Color", command=self.color_text)
        edit_menu.add_command(label="Change Font Size", command=self.change_font_size)
        edit_menu.add_command(label="Insert Table", command=self.insert_table)

    def open_file(self):
        file_path = filedialog.askopenfilename(defaultextension=".txt",
                                              filetypes=[
                                                  ("Mettarol Word Files", "*.metwc"),
                                                  ("Batch Files", "*.bat"),
                                                  ("VBScript Files", "*.vbs"),
                                                  ("Document Files", "*.docx"),
                                                  ("Markdown Files", "*.md"),
                                                  ("Comma Separated Values", "*.csv"),
                                                  ("Open Document Format", "*.odt"),
                                                  ("Python Scripts", "*.py"),
                                                  ("Rich Text Format", "*.rft"),
                                                  ("HyperText Markup Language", "*.html"),
                                                  ("Cascading Style Sheets", "*.css"),
                                                  ("Game Creation System", "*.gcs"),
                                                  ("Java Trace File", "*.jtf"),
                                                  ("Terraform Configuration", "*.tf"),
                                                  ("Bibliography Files", "*.bib"),
                                                  ("C Source Code", "*.c"),
                                                  ("JavaScript Object Notation", "*.json"),
                                                  ("Initialization Files", "*.ini"),
                                                  ("C++ Source Code", "*.cpp"),
                                                  ("Extensible Markup Language", "*.xml"),
                                                  ("Microsoft Word Documents", "*.doc"),
                                                  ("Java Source Code", "*.java"),
                                                  ("JavaScript", "*.js"),
                                                  ("High Level Shader Language", "*.hlsl"),
                                                  ("Go Programming Language", "*.go"),
                                                  ("FictionBook eBooks", "*.fb2"),
                                                  ("Electronic Publication Format", "*.epub"),
                                                  ("Mobipocket eBooks", "*.mobi"),
                                                  ("Handlebars Templates", "*.handlebars"),
                                                  ("Handlebars Templates", "*.hbs"),
                                                  ("Handlebars Templates", "*.hjs"),
                                                  ("Help Files", "*.help"),
                                                  ("Security Related Files", "*.security"),
                                                  ("Kum Scripts", "*.kum"),
                                                  ("Event Response Model", "*.erm"),
                                                  ("Run Time Environment", "*.rt"),
                                                  ("Hierarchical JavaScript", "*.hj"),
                                                  ("Project Definition", "*.pj"),
                                                  ("Workflow Definition", "*.wj"),
                                                  ("Visual Studio Code Settings", "*.vscode"),
                                                  ("Bourne Shell Scripts", "*.sh"),
                                                  ("PowerShell Scripts", "*.ps1"),
                                                  ("Ruby Programs", "*.rb"),
                                                  ("Perl Scripts", "*.pl"),
                                                  ("Command Line Interface", "*.cmd"),
                                                  ("Plain Text Document", "*.txt"),
                                                  ("All Files", "*.*")
                                              ])
        if not file_path:
            return
        with open(file_path, 'r') as file:
            content = file.read()
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, content)
            self.current_file = file_path
            self.is_modified = False

    def save_file(self):
        if self.current_file is None:
            file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                    filetypes=[
                                                        ("Mettarol Word Files", "*.metwc"),
                                                        ("Batch Files", "*.bat"),
                                                        ("VBScript Files", "*.vbs"),
                                                        ("Document Files", "*.docx"),
                                                        ("Markdown Files", "*.md"),
                                                        ("Comma Separated Values", "*.csv"),
                                                        ("Open Document Format", "*.odt"),
                                                        ("Python Scripts", "*.py"),
                                                        ("Rich Text Format", "*.rft"),
                                                        ("HyperText Markup Language", "*.html"),
                                                        ("Cascading Style Sheets", "*.css"),
                                                        ("Game Creation System", "*.gcs"),
                                                        ("Java Trace File", "*.jtf"),
                                                        ("Terraform Configuration", "*.tf"),
                                                        ("Bibliography Files", "*.bib"),
                                                        ("C Source Code", "*.c"),
                                                        ("JavaScript Object Notation", "*.json"),
                                                        ("Initialization Files", "*.ini"),
                                                        ("C++ Source Code", "*.cpp"),
                                                        ("Extensible Markup Language", "*.xml"),
                                                        ("Microsoft Word Documents", "*.doc"),
                                                        ("Java Source Code", "*.java"),
                                                        ("JavaScript", "*.js"),
                                                        ("High Level Shader Language", "*.hlsl"),
                                                        ("Go Programming Language", "*.go"),
                                                        ("FictionBook eBooks", "*.fb2"),
                                                        ("Electronic Publication Format", "*.epub"),
                                                        ("Mobipocket eBooks", "*.mobi"),
                                                        ("Handlebars Templates", "*.handlebars"),
                                                        ("Handlebars Templates", "*.hbs"),
                                                        ("Handlebars Templates", "*.hjs"),
                                                        ("Help Files", "*.help"),
                                                        ("Security Related Files", "*.security"),
                                                        ("Kum Scripts", "*.kum"),
                                                        ("Event Response Model", "*.erm"),
                                                        ("Run Time Environment", "*.rt"),
                                                        ("Hierarchical JavaScript", "*.hj"),
                                                        ("Project Definition", "*.pj"),
                                                        ("Workflow Definition", "*.wj"),
                                                        ("Visual Studio Code Settings", "*.vscode"),
                                                        ("Bourne Shell Scripts", "*.sh"),
                                                        ("PowerShell Scripts", "*.ps1"),
                                                        ("Ruby Programs", "*.rb"),
                                                        ("Perl Scripts", "*.pl"),
                                                        ("Command Line Interface", "*.cmd"),
                                                        ("Plain Text Document", "*.txt"),
                                                        ("All Files", "*.*")
                                                    ])
            if not file_path:
                return
            self.current_file = file_path
            
        with open(self.current_file, 'w') as file:
            content = self.text_area.get(1.0, tk.END)
            file.write(content)
            self.is_modified = False

    def bold_text(self):
        current_tags = self.text_area.tag_names("sel.first")
        if "bold" in current_tags:
            self.text_area.tag_remove("bold", "sel.first", "sel.last")
        else:
            self.text_area.tag_add("bold", "sel.first", "sel.last")
            self.text_area.tag_config("bold", font=("Arial", 12, "bold"))

    def underline_text(self):
        current_tags = self.text_area.tag_names("sel.first")
        if "underline" in current_tags:
            self.text_area.tag_remove("underline", "sel.first", "sel.last")
        else:
            self.text_area.tag_add("underline", "sel.first", "sel.last")
            self.text_area.tag_config("underline", underline=True)

    def italic_text(self):
        current_tags = self.text_area.tag_names("sel.first")
        if "italic" in current_tags:
            self.text_area.tag_remove("italic", "sel.first", "sel.last")
        else:
            self.text_area.tag_add("italic", "sel.first", "sel.last")
            self.text_area.tag_config("italic", font=("Arial", 12, "italic"))

    def color_text(self):
        color = colorchooser.askcolor()[1]
        if color:
            self.text_area.tag_add("color", "sel.first", "sel.last")
            self.text_area.tag_config("color", foreground=color)

    def change_font_size(self):
        size = simpledialog.askinteger("Font Size", "Choose a font size:")
        if size:
            self.text_area.config(font=("Arial", size))

    def insert_table(self):
        rows = simpledialog.askinteger("Table", "Number of Rows:")
        cols = simpledialog.askinteger("Table", "Number of Columns:")
        if rows and cols:
            table_data = ""
            for _ in range(rows):
                row_data = ["|"] + ["Cell"] * cols + ["|"]
                table_data += " ".join(row_data) + "\n"
            self.text_area.insert(tk.END, table_data)

if __name__ == "__main__":
    root = tk.Tk()
    app = MettarolWord(root)
    root.mainloop()
