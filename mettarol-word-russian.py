import tkinter as tk
from tkinter import filedialog, messagebox, font, colorchooser, simpledialog, Menu, Text, Frame, Scrollbar

class MettarolWord:
    def __init__(self, root):
        root.geometry('800x600')
        self.root = root
        self.root.title("Mettarol Word")
        self.text_area = Text(self.root, wrap='word', font=("Arial", 12))
        self.text_area.pack(expand=True, fill='both')
        
        self.is_modified = False  # Переменная отслеживания изменений
        self.current_file = None  # Имя открытого файла

        self.create_menu()

        self.scrollbar = Scrollbar(self.text_area)
        self.scrollbar.pack(side='right', fill='y')
        self.text_area.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.text_area.yview)

        # Обработчик события изменения текста
        self.text_area.bind('<Key>', self.mark_as_modified)

        # Связываем событие закрытия окна с обработчиком проверки
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def mark_as_modified(self, event):
        """Отмечает документ как измененный."""
        self.is_modified = True

    def on_close(self):
        """Проверяет состояние файла перед закрытием программы."""
        if self.is_modified:
            result = messagebox.askyesno(
                title="Предупреждение!",
                message="Файл не сохранён. Хотите сохранить его?"
            )
            if result:
                self.save_file()
        self.root.destroy()

    def create_menu(self):
        menu = Menu(self.root)
        self.root.config(menu=menu)

        file_menu = Menu(menu)
        menu.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Открыть", command=self.open_file)
        file_menu.add_command(label="Сохранить", command=self.save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Выйти", command=self.on_close)

        edit_menu = Menu(menu)
        menu.add_cascade(label="Редактировать", menu=edit_menu)
        edit_menu.add_command(label="Сделать жирным", command=self.bold_text)
        edit_menu.add_command(label="Подчеркнуть", command=self.underline_text)
        edit_menu.add_command(label="Наклонить", command=self.italic_text)
        edit_menu.add_command(label="Изменить цвет текста", command=self.color_text)
        edit_menu.add_command(label="Изменить размер текста", command=self.change_font_size)
        edit_menu.add_command(label="Создать таблицу", command=self.insert_table)

    def open_file(self):
        file_path = filedialog.askopenfilename(defaultextension=".txt",
                                              filetypes=[
                                                  ("Mettarol Word Файлы", "*.metwc"),
                                                  ("Bat Файлы", "*.bat"),
                                                  ("VBScript Файлы", "*.vbs"),
                                                  ("Docx Файлы", "*.docx"),
                                                  ("MD Файлы", "*.md"),
                                                  ("CSV Файлы", "*.csv"),
                                                  ("ODT Файлы", "*.odt"),
                                                  ("Python файлы", "*.py"),
                                                  ("RFT Файлы", "*.rft"),
                                                  ("HTML Файлы", "*.html"),
                                                  ("CSS Файлы", "*.css"),
                                                  ("GCS Файлы", "*.gcs"),
                                                  ("JTF Файлы", "*.jtf"),
                                                  ("TF Файлы", "*.tf"),
                                                  ("BiB Файлы", "*.bib"),
                                                  ("C Файлы", "*.c"),
                                                  ("JSON Файлы", "*.json"),
                                                  ("Ini Файлы", "*.ini"),
                                                  ("C++ Файлы", "*.c++"),
                                                  ("XML файлы", "*.xml"),
                                                  ("Doc Файлы", "*.doc"),
                                                  ("JAVA Файлы", "*.java"),
                                                  ("JAVA SCRIPT Файлы", "*.js"),
                                                  ("HLSL Файлы", "*.hlsl"),
                                                  ("GO Файлы", "*.go"),
                                                  ("FB2 файлы", "*.fb2"),
                                                  ("EPUB Файлы", "*.epub"),
                                                  ("MOBI Файлы", "*.mobi"),
                                                  ("Handlebars Файлы", "*.handlebars"),
                                                  ("Handlebars Файлы", "*.hbs"),
                                                  ("Handlebars Файлы", "*.hjs"),
                                                  ("Help Файлы", "*.help"),
                                                  ("Security Файлы", "*.security"),
                                                  ("KUM Файлы", "*.kum"),
                                                  ("ERM Файлы", "*.erm"),
                                                  ("RT Файлы", "*.rt"),
                                                  ("HJ Файлы", "*.hj"),
                                                  ("PJ Файлы", "*.pj"),
                                                  ("WJ Файлы", "*.wj"),
                                                  ("Visual Studio Code Файлы", "*.vscode"),
                                                  ("Bash Файлы", "*.sh"),
                                                  ("PowerShell Файлы", "*.ps1"),
                                                  ("Ruby Файлы", "*.rb"),
                                                  ("Perl Файлы", "*.pl"),
                                                  ("CMD Файлы", "*.cmd"),
                                                  ("Текстовый Документ", "*.txt"),
                                                  ("Все файлы", "*.*")
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
                                                        ("Mettarol Word Файлы", "*.metwc"),
                                                        ("Bat Файлы", "*.bat"),
                                                        ("VBScript Файлы", "*.vbs"),
                                                        ("Docx Файлы", "*.docx"),
                                                        ("MD Файлы", "*.md"),
                                                        ("CSV Файлы", "*.csv"),
                                                        ("ODT Файлы", "*.odt"),
                                                        ("Python файлы", "*.py"),
                                                        ("RFT Файлы", "*.rft"),
                                                        ("HTML Файлы", "*.html"),
                                                        ("CSS Файлы", "*.css"),
                                                        ("GCS Файлы", "*.gcs"),
                                                        ("JTF Файлы", "*.jtf"),
                                                        ("TF Файлы", "*.tf"),
                                                        ("BiB Файлы", "*.bib"),
                                                        ("C Файлы", "*.c"),
                                                        ("JSON Файлы", "*.json"),
                                                        ("Ini Файлы", "*.ini"),
                                                        ("C++ Файлы", "*.c++"),
                                                        ("XML файлы", "*.xml"),
                                                        ("Doc Файлы", "*.doc"),
                                                        ("JAVA Файлы", "*.java"),
                                                        ("JAVA SCRIPT Файлы", "*.js"),
                                                        ("HLSL Файлы", "*.hlsl"),
                                                        ("GO Файлы", "*.go"),
                                                        ("FB2 файлы", "*.fb2"),
                                                        ("EPUB Файлы", "*.epub"),
                                                        ("MOBI Файлы", "*.mobi"),
                                                        ("Handlebars Файлы", "*.handlebars"),
                                                        ("Handlebars Файлы", "*.hbs"),
                                                        ("Handlebars Файлы", "*.hjs"),
                                                        ("Help Файлы", "*.help"),
                                                        ("Security Файлы", "*.security"),
                                                        ("KUM Файлы", "*.kum"),
                                                        ("ERM Файлы", "*.erm"),
                                                        ("RT Файлы", "*.rt"),
                                                        ("HJ Файлы", "*.hj"),
                                                        ("PJ Файлы", "*.pj"),
                                                        ("WJ Файлы", "*.wj"),
                                                        ("Visual Studio Code Файлы", "*.vscode"),
                                                        ("Bash Файлы", "*.sh"),
                                                        ("PowerShell Файлы", "*.ps1"),
                                                        ("Ruby Файлы", "*.rb"),
                                                        ("Perl Файлы", "*.pl"),
                                                        ("CMD Файлы", "*.cmd"),
                                                        ("Текстовый Документ", "*.txt"),
                                                        ("Все файлы", "*.*")
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
        size = simpledialog.askinteger("Размер шрифта", "Выберите размер шрифта:")
        if size:
            self.text_area.config(font=("Arial", size))


    def insert_table(self):
        rows = simpledialog.askinteger("Таблица", "Количество строк:")
        cols = simpledialog.askinteger("Таблица", "Количество столбцов:")
        if rows and cols:
            table_data = ""
            for _ in range(rows):
                row_data = ["|"] + ["Cell"]*cols + ["|"]
                table_data += " ".join(row_data) + "\n"
            self.text_area.insert(tk.END, table_data)

if __name__ == "__main__":
    root = tk.Tk()
    app = MettarolWord(root)
    root.mainloop()
