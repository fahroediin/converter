import tkinter as tk
from tkinter import messagebox
import csv
import io

class CsvToExcelTool:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Converter (No Gap)")
        self.root.geometry("800x600")

        # --- Frame Input ---
        frame_top = tk.Frame(root)
        frame_top.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        lbl_input = tk.Label(frame_top, text="1. Paste Data Mentah di sini:", font=("Arial", 10, "bold"))
        lbl_input.pack(anchor="w")

        self.txt_input = tk.Text(frame_top, height=10)
        self.txt_input.pack(fill=tk.BOTH, expand=True)

        # --- Frame Controls (Delimiter & Button) ---
        frame_mid = tk.Frame(root, bg="#f0f0f0", bd=1, relief=tk.RAISED)
        frame_mid.pack(fill=tk.X, padx=10, pady=10)

        lbl_delim = tk.Label(frame_mid, text="Pilih Delimiter (Pemisah):", bg="#f0f0f0")
        lbl_delim.pack(side=tk.LEFT, padx=5, pady=10)

        self.entry_delim = tk.Entry(frame_mid, width=5, font=("Arial", 12, "bold"), justify='center')
        self.entry_delim.insert(0, "|") # Default delimiter
        self.entry_delim.pack(side=tk.LEFT, padx=5)

        btn_convert = tk.Button(frame_mid, text="KONVERSI KE EXCEL FORMAT", command=self.convert_data, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        btn_convert.pack(side=tk.LEFT, padx=20)

        btn_clear = tk.Button(frame_mid, text="Hapus Semua", command=self.clear_all, bg="#FF5722", fg="white")
        btn_clear.pack(side=tk.RIGHT, padx=10)

        # --- Frame Output ---
        frame_bottom = tk.Frame(root)
        frame_bottom.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        lbl_output = tk.Label(frame_bottom, text="2. Hasil (Siap Copy ke Excel):", font=("Arial", 10, "bold"))
        lbl_output.pack(anchor="w")

        self.txt_output = tk.Text(frame_bottom, height=10, bg="#e8f5e9")
        self.txt_output.pack(fill=tk.BOTH, expand=True)

        btn_copy = tk.Button(root, text="COPY HASIL KE CLIPBOARD", command=self.copy_to_clipboard, bg="#2196F3", fg="white", font=("Arial", 12, "bold"), pady=10)
        btn_copy.pack(fill=tk.X, padx=10, pady=10)

    def convert_data(self):
        # Ambil data dan bersihkan spasi di awal/akhir string utama
        raw_data = self.txt_input.get("1.0", tk.END).strip()
        delimiter = self.entry_delim.get()

        if not raw_data:
            messagebox.showwarning("Peringatan", "Data input kosong!")
            return
        
        if not delimiter:
            messagebox.showwarning("Peringatan", "Delimiter tidak boleh kosong!")
            return

        try:
            f_input = io.StringIO(raw_data)
            reader = csv.reader(f_input, delimiter=delimiter)
            
            f_output = io.StringIO()
            
            # PERBAIKAN DI SINI: lineterminator='\n' mencegah baris ganda
            writer = csv.writer(f_output, delimiter='\t', lineterminator='\n') 

            row_count = 0
            for row in reader:
                # Cek jika row tidak kosong (list kosong)
                if row and any(field.strip() for field in row):
                    # Membersihkan spasi berlebih di setiap sel
                    clean_row = [cell.strip() for cell in row]
                    writer.writerow(clean_row)
                    row_count += 1

            result = f_output.getvalue()
            
            self.txt_output.delete("1.0", tk.END)
            self.txt_output.insert("1.0", result)
            
            messagebox.showinfo("Sukses", f"Berhasil mengonversi {row_count} baris data.")

        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan format:\n{str(e)}")

    def copy_to_clipboard(self):
        result = self.txt_output.get("1.0", tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(result)
        self.root.update() 
        messagebox.showinfo("Copied", "Data berhasil disalin! Silakan Paste (Ctrl+V) di Excel.")

    def clear_all(self):
        self.txt_input.delete("1.0", tk.END)
        self.txt_output.delete("1.0", tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = CsvToExcelTool(root)
    root.mainloop()