import os
import subprocess
from tkinter import Tk, Frame, Label, Button, Entry, filedialog, messagebox, Canvas, END
from tkinter import ttk
from PIL import Image, ImageTk
import openpyxl
from openpyxl import Workbook

class ExcelImageAnnotator:
    def __init__(self, root):
        self.root = root
        self.root.title("برنامه ثبت داده برای عکس‌ها")
        self.root.geometry("1200x700")
        
        # تقسیم فرم به دو بخش:
        # image_frame: 70% عرض فرم (نمایش تصویر)
        # control_frame: 30% عرض فرم (کنترل‌ها)
        self.image_frame = Frame(root, bg="black")
        self.image_frame.pack(side="left", fill="both", expand=True)
        
        self.control_frame = Frame(root, bg="#f0f0f0", width=360)
        self.control_frame.pack(side="right", fill="y")
        self.control_frame.pack_propagate(False)
        
        # Canvas نمایش تصویر در image_frame
        self.canvas = Canvas(self.image_frame, bg="black")
        self.canvas.pack(fill="both", expand=True)
        self.current_image = None  # نگهداری تصویر نمایش داده شده

        # لیست تصاویر و اندیس تصویر جاری
        self.image_list = []
        self.current_index = 0

        # --- کنترل‌های بخش کنترلی در control_frame ---
        Button(self.control_frame, text="انتخاب تصاویر", command=self.select_images, font=("Arial", 15)).pack(pady=10)

        # Treeview نمایش لیست تصاویر (با دو ستون: پوشه و مسیر)
        self.tree = ttk.Treeview(self.control_frame, columns=("Folder", "Path"), show="headings", height=8)
        self.tree.heading("Folder", text="نام پوشه")
        self.tree.heading("Path", text="مسیر عکس")
        self.tree.column("Folder", width=100)
        self.tree.column("Path", width=240)
        self.tree.pack(padx=10, pady=10)

        # دکمه حذف تصویر انتخاب شده
        Button(self.control_frame, text="حذف تصویر انتخاب شده", command=self.remove_selected_image, font=("Arial", 15)).pack(pady=5)

        # انتخاب مسیر فایل اکسل (ورودی با عرض کمتر)
        frame_excel = Frame(self.control_frame, bg="#f0f0f0")
        frame_excel.pack(padx=10, pady=5, fill="x")
        Label(frame_excel, text="مسیر فایل اکسل:", bg="#f0f0f0", font=("Arial", 15)).pack(side="left")
        self.excel_path_entry = Entry(frame_excel, width=15, font=("Arial", 15), justify="right")
        self.excel_path_entry.pack(side="left", padx=5)
        Button(frame_excel, text="انتخاب", command=self.select_excel_file, font=("Arial", 15)).pack(side="left", padx=5)

        Button(self.control_frame, text="استارت", command=self.start_process, font=("Arial", 15)).pack(pady=10)

        # بخش ورود داده (باکس‌های ورودی) به صورت منظم و راست‌چین
        self.data_frame = Frame(self.control_frame, bg="#f0f0f0")
        self.data_frame.pack(padx=10, pady=10, fill="x")
        
        self.fields = [
            "لوله پلیکا 75", "زانو پلیکا", "شیلنگ",
            "بست شیلنگ", "پلمپ", "سرشیلنگی", "لوله 28", "لوله 70"
        ]
        self.entries = {}
        for idx, field in enumerate(self.fields):
            row_frame = Frame(self.data_frame, bg="#f0f0f0")
            row_frame.grid(row=idx, column=0, sticky="e", pady=3)
            Label(row_frame, text=field + ":", bg="#f0f0f0", font=("Arial", 15)).grid(row=0, column=0, sticky="e")
            entry = Entry(row_frame, width=4, font=("Arial", 15), justify="right")
            entry.grid(row=0, column=1, padx=5)
            if field in ["بست شیلنگ", "پلمپ"]:
                entry.insert(0, "2")
            entry.bind("<Return>", lambda event, i=idx: self.focus_next(i))
            self.entries[field] = entry

        Button(self.control_frame, text="ذخیره داده", command=self.save_current_data, font=("Arial", 15)).pack(pady=10)
        Button(self.control_frame, text="خروج", command=root.quit, font=("Arial", 15)).pack(pady=10)

        self.wb = None
        self.ws = None

    def focus_next(self, current_index):
        if current_index < len(self.fields) - 1:
            next_field = self.fields[current_index + 1]
            self.entries[next_field].focus_set()

    def select_images(self):
        files = filedialog.askopenfilenames(
            title="انتخاب تصاویر",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                if f not in self.image_list:
                    self.image_list.append(f)
                    folder = os.path.basename(os.path.dirname(f))
                    self.tree.insert("", "end", values=(folder, f))
            if self.current_index >= len(self.image_list):
                self.current_index = 0

    def remove_selected_image(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("هشدار", "لطفاً یک تصویر برای حذف انتخاب کنید.")
            return
        for item in selected:
            values = self.tree.item(item, "values")
            image_path = values[1]
            if image_path in self.image_list:
                idx = self.image_list.index(image_path)
                del self.image_list[idx]
            self.tree.delete(item)
        if self.current_index >= len(self.image_list):
            self.current_index = 0
        if self.image_list:
            self.show_image(self.image_list[self.current_index])
        else:
            self.canvas.delete("all")
            messagebox.showinfo("اطلاع", "لیست تصاویر خالی است.")

    def select_excel_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.excel_path_entry.delete(0, END)
            self.excel_path_entry.insert(0, file_path)

    def start_process(self):
        if not self.image_list:
            messagebox.showerror("خطا", "لطفاً ابتدا تصاویر را انتخاب نمایید.")
            return
        if not self.excel_path_entry.get():
            messagebox.showerror("خطا", "لطفاً مسیر فایل اکسل را انتخاب نمایید.")
            return
        
        excel_path = self.excel_path_entry.get()
        if os.path.exists(excel_path):
            self.wb = openpyxl.load_workbook(excel_path)
            self.ws = self.wb.active
        else:
            self.wb = Workbook()
            self.ws = self.wb.active
            headers = ["پوشه", "مسیر عکس"] + self.fields
            self.ws.append(headers)
            self.wb.save(excel_path)
        self.show_image(self.image_list[self.current_index])
        first_field = self.fields[0]
        self.entries[first_field].focus_set()

    def show_image(self, image_path):
        """نمایش تصویر کامل با کاهش اندازه برای تطبیق با صفحه نمایش بدون نیاز به بزرگنمایی."""
        try:
            if not os.path.exists(image_path):
                raise FileNotFoundError("فایل تصویر یافت نشد.")
            img = Image.open(image_path)
            canvas_width = self.canvas.winfo_width() or 800
            canvas_height = self.canvas.winfo_height() or 700
            # کاهش اندازه تصویر به طوری که کل تصویر در canvas جای بگیرد
            img.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
            self.current_image = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor="nw", image=self.current_image)
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
        except Exception as e:
            messagebox.showerror("خطا", f"مشکل در نمایش تصویر: {e}")

    def save_current_data(self):
        if not self.wb or not self.ws:
            messagebox.showerror("خطا", "فرآیند استارت اجرا نشده است.")
            return

        for field in ["لوله 70", "لوله 28", "سرشیلنگی"]:
            try:
                int(self.entries[field].get())
            except ValueError:
                messagebox.showerror("خطا", f"مقدار '{field}' باید یک عدد صحیح باشد.")
                return

        try:
            shilung_value = float(self.entries["شیلنگ"].get())
            if shilung_value > 2:
                messagebox.showerror("خطا", "مقدار شیلنگ نباید بیش از 2 باشد.")
                return
        except ValueError:
            messagebox.showerror("خطا", "مقدار شیلنگ باید عددی باشد.")
            return

        try:
            looleh_value = float(self.entries["لوله پلیکا 75"].get())
            if looleh_value > 4:
                messagebox.showerror("خطا", "مقدار لوله پلیکا 75 نباید بیش از 4 متر باشد.")
                return
        except ValueError:
            messagebox.showerror("خطا", "مقدار لوله پلیکا 75 باید عددی باشد.")
            return

        data_values = []
        for field in self.fields:
            value = self.entries[field].get()
            if value == "":
                messagebox.showerror("خطا", f"لطفاً مقدار '{field}' را وارد نمایید.")
                return
            data_values.append(value)
        
        current_image_path = self.image_list[self.current_index]
        folder = os.path.basename(os.path.dirname(current_image_path))
        row_data = [folder, current_image_path] + data_values
        
        try:
            self.ws.append(row_data)
            self.wb.save(self.excel_path_entry.get())
            messagebox.showinfo("موفقیت", "داده‌ها با موفقیت ذخیره شدند.")
        except Exception as e:
            messagebox.showerror("خطا", f"مشکل در ذخیره به اکسل: {e}")
            return
        
        for field in self.fields:
            self.entries[field].delete(0, END)
            if field in ["بست شیلنگ", "پلمپ"]:
                self.entries[field].insert(0, "2")
        
        self.canvas.delete("all")
        self.current_index += 1
        if self.current_index < len(self.image_list):
            self.show_image(self.image_list[self.current_index])
            first_field = self.fields[0]
            self.entries[first_field].focus_set()
        else:
            messagebox.showinfo("پایان", "تمام تصاویر پردازش شدند.")

if __name__ == "__main__":
    root = Tk()
    app = ExcelImageAnnotator(root)
    root.mainloop()
