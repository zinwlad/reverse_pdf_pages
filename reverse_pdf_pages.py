import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PyPDF2 import PdfReader, PdfWriter
import os
import logging
from typing import Optional, Tuple

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='pdf_reverser.log'
)

class PDFReverserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Reverser")
        self.root.geometry("500x300")
        
        # Создаем главный фрейм
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(
            main_frame, 
            text="Перетащите PDF-файл сюда или нажмите 'Выбрать PDF'",
            wraplength=400
        )
        self.label.pack(pady=10)

        # Фрейм для кнопок
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=5)

        self.open_button = ttk.Button(
            button_frame, 
            text="Выбрать PDF", 
            command=self.open_file
        )
        self.open_button.pack(side=tk.LEFT, padx=5)

        self.save_button = ttk.Button(
            button_frame, 
            text="Сохранить изменённый PDF", 
            command=self.save_file, 
            state=tk.DISABLED
        )
        self.save_button.pack(side=tk.LEFT, padx=5)

        # Фрейм для выбора страниц
        page_frame = ttk.LabelFrame(main_frame, text="Диапазон страниц", padding="5")
        page_frame.pack(pady=10, fill=tk.X)

        self.page_var = tk.StringVar(value="all")
        ttk.Radiobutton(
            page_frame, 
            text="Все страницы", 
            variable=self.page_var, 
            value="all"
        ).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(
            page_frame, 
            text="Выбранные страницы", 
            variable=self.page_var, 
            value="range"
        ).pack(side=tk.LEFT, padx=5)

        self.range_frame = ttk.Frame(page_frame)
        self.range_frame.pack(side=tk.LEFT, padx=5)
        
        self.start_page = ttk.Entry(self.range_frame, width=5)
        self.start_page.pack(side=tk.LEFT, padx=2)
        ttk.Label(self.range_frame, text="-").pack(side=tk.LEFT)
        self.end_page = ttk.Entry(self.range_frame, width=5)
        self.end_page.pack(side=tk.LEFT, padx=2)

        # Прогресс бар
        self.progress = ttk.Progressbar(
            main_frame, 
            orient=tk.HORIZONTAL, 
            length=300, 
            mode='determinate'
        )
        self.progress.pack(pady=10, fill=tk.X)

        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=10)

        self.file_path = None
        self.total_pages = 0

        # Настройка Drag-and-Drop
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

    def get_page_range(self) -> Optional[Tuple[int, int]]:
        """Получение выбранного диапазона страниц"""
        if self.page_var.get() == "all":
            return None
        
        try:
            start = int(self.start_page.get()) if self.start_page.get() else 1
            end = int(self.end_page.get()) if self.end_page.get() else self.total_pages
            
            if start < 1:
                start = 1
            if end > self.total_pages:
                end = self.total_pages
            if start > end:
                start, end = end, start
                
            return start - 1, end  # PyPDF2 использует 0-based индексацию
        except ValueError:
            messagebox.showwarning(
                "Предупреждение",
                "Неверный формат диапазона страниц. Будут обработаны все страницы."
            )
            return None

    def open_file(self):
        """Открытие файла через диалоговое окно"""
        try:
            self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if self.file_path:
                self.update_status(f"Выбран файл: {self.file_path}")
                self.save_button.config(state=tk.NORMAL)
            else:
                self.update_status("Файл не выбран.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")

    def save_file(self):
        """Сохранение изменённого PDF файла"""
        if not self.file_path:
            self.update_status("Ошибка: Файл не выбран")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            try:
                self.reverse_pdf_pages(self.file_path, output_path)
                self.update_status(f"Изменённый файл сохранён: {output_path}")
                messagebox.showinfo("Успех", f"Файл успешно сохранён: {output_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла:\n{e}")

    def reverse_pdf_pages(self, input_pdf_path, output_pdf_path):
        """Реверсирование страниц PDF и сохранение в новый файл"""
        try:
            reader = PdfReader(input_pdf_path)
            writer = PdfWriter()
            
            self.total_pages = len(reader.pages)
            page_range = self.get_page_range()
            
            if page_range:
                start, end = page_range
                pages_to_process = range(start, end)
            else:
                pages_to_process = range(len(reader.pages))

            # Копируем страницы до начала выбранного диапазона
            if page_range and start > 0:
                for i in range(start):
                    writer.add_page(reader.pages[i])

            # Реверсируем выбранные страницы
            total_pages = len(pages_to_process)
            for idx, page_num in enumerate(reversed(pages_to_process)):
                writer.add_page(reader.pages[page_num])
                progress = (idx + 1) / total_pages * 100
                self.progress['value'] = progress
                self.root.update_idletasks()

            # Копируем страницы после выбранного диапазона
            if page_range and end < len(reader.pages):
                for i in range(end, len(reader.pages)):
                    writer.add_page(reader.pages[i])

            with open(output_pdf_path, "wb") as output_pdf:
                writer.write(output_pdf)
                
            self.progress['value'] = 100
            
        except Exception as e:
            logging.error(f"Error processing PDF: {str(e)}")
            raise Exception(f"Ошибка при реверсировании PDF: {str(e)}")

    def on_drop(self, event):
        """Обработка перетаскивания файла в окно"""
        file_path = self.extract_file_path(event.data)
        if file_path and os.path.isfile(file_path):
            self.file_path = file_path
            self.update_status(f"Выбран файл: {self.file_path}")
            self.save_button.config(state=tk.NORMAL)
        else:
            self.update_status("Ошибка: Не удалось загрузить файл")

    def extract_file_path(self, data):
        """Извлечение пути файла из данных DND"""
        if data.startswith('{'):
            data = data.split('}')[0]  # Удаление лишних символов, характерных для некоторых ОС
        return data.replace('{', '').replace('}', '').replace('/', os.sep)

    def update_status(self, message):
        """Обновление статуса на экране"""
        self.status_label.config(text=message)


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFReverserApp(root)
    root.mainloop()
