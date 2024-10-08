import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from timeparser import ScheduleParser

class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Менеджер расписания(Создано Чуевым В.О)")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        self.load_button = tk.Button(root, text="Загрузите файл с расписанием", command=self.load_file)
        self.load_button.pack(pady=20)

        self.teachers_listbox = tk.Listbox(root, width=50, height=10)
        self.teachers_listbox.pack(pady=20, fill=tk.BOTH, expand=True)
        self.teachers_listbox.bind('<<ListboxSelect>>', self.show_teacher_info)

        self.tree = ttk.Treeview(root, columns=("dt", "pair", "subj", "groups"), show='headings', height=10)

        self.tree.heading("dt", text="Дата")
        self.tree.heading("pair", text="Номер пары")
        self.tree.heading("subj", text="Предмет")
        self.tree.heading("groups", text="Группы")

        self.tree.column("dt", width=100, anchor='center')
        self.tree.column("pair", width=60, anchor='center')

        self.tree.column("subj", width=250, anchor='w')
        self.tree.column("groups", width=250, anchor='w')  

        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите эксель файл",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            try:
                self.parser = ScheduleParser(file_path)
                self.parser.parse_schedule()
                self.teachers = self.parser.get_teachers_schedule()
                self.populate_teachers_list()
            except Exception as e:
                messagebox.showerror("Ошибка", f"не удалось обработать: {e}")

    def populate_teachers_list(self):
        self.teachers_listbox.delete(0, tk.END)

        for teacher in self.teachers:
            self.teachers_listbox.insert(tk.END, teacher)

    def show_teacher_info(self, event):
        selection = self.teachers_listbox.curselection()
        if selection:
            teacher_name = self.teachers_listbox.get(selection[0])

            for item in self.tree.get_children():
                self.tree.delete(item)

            schedule = self.teachers[teacher_name]

            schedule = sorted(schedule, key=lambda x: x['dt'])

            grouped_schedule = {}
            items = dict()
            for entry in schedule:
                dt = entry['dt'].strftime('%d.%m.%Y')
                num = entry['lesson_num']
                mem_str = dt + str(num)
                if dt not in grouped_schedule:
                    grouped_schedule[dt] = []
                items[mem_str] = items.get(mem_str, 0) + 1
                entry['mem'] = mem_str
                grouped_schedule[dt].append(entry)

            for date, entries in grouped_schedule.items():
                for entry in entries:
                    item = items.get(entry['mem'], 1)
                    row_id = self.tree.insert("", tk.END, values=(date, entry['lesson_num'], entry['subj'], ", ".join(entry['groups'])))
                    if item > 1:
                        self.tree.tag_configure('red', background='red')
                        self.tree.item(row_id, tags=('red',))
if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()
