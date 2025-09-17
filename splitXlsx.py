import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 文件按年级班级拆分")
        self.root.geometry("600x400")  # Increased size for better visibility

        # Variables
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        # Set default output directory
        if getattr(sys, 'frozen', False):
            # Running as EXE
            script_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            script_dir = os.path.dirname(os.path.abspath(__file__))
        default_output = os.path.join(script_dir, "OutFiles")
        self.output_dir.set(default_output)

        # GUI Elements
        # Input file selection
        tk.Label(root, text="输入 Excel 文件:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(root, textvariable=self.input_file, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(root, text="浏览", command=self.browse_input_file, width=10).grid(row=0, column=2, padx=10, pady=10)

        # Output folder selection
        tk.Label(root, text="输出文件夹:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(root, textvariable=self.output_dir, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(root, text="浏览", command=self.browse_output_folder, width=10).grid(row=1, column=2, padx=10, pady=10)
        
        # Process button
        tk.Button(root, text="开始处理", command=self.process_excel, width=15, font=("Arial", 12)).grid(row=2, column=1, pady=20)
        
        # Open output folder button
        tk.Button(root, text="打开输出文件夹", command=self.open_output_folder, width=15, font=("Arial", 12)).grid(row=3, column=1, pady=10)
        
        # Status label
        self.status_label = tk.Label(root, text="请先选择输入文件和输出文件夹", font=("Arial", 10), wraplength=500)
        self.status_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择要处理的 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_file.set(file_path)
            self.status_label.config(text=f"已选择输入文件: {file_path}")
            print(f"Selected input file: {file_path}")

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self.output_dir.get()  # Start at current output_dir
        )
        if folder_path:
            self.output_dir.set(folder_path)
            self.status_label.config(text=f"已选择输出文件夹: {folder_path}")
            print(f"Selected output folder: {folder_path}")

    def process_excel(self):
        input_file = self.input_file.get()
        output_dir = self.output_dir.get()
        
        if not input_file:
            messagebox.showerror("错误", "请先选择一个 Excel 文件！")
            self.status_label.config(text="错误：未选择输入文件")
            return
        
        if not output_dir:
            messagebox.showerror("错误", "请先选择输出文件夹！")
            self.status_label.config(text="错误：未选择输出文件夹")
            return
        
        try:
            # Read Excel file
            df = pd.read_excel(input_file, engine='openpyxl')
            
            # Clean data
            df['年级'] = df['年级'].astype(str).str.strip()
            df['班级'] = df['班级'].astype(str).str.strip()
            
            # Print debugging info
            print("读取的数据行数:", len(df))
            print("所有年级:", df['年级'].unique())
            print("所有班级:", df['班级'].unique())
            
            # Check for required columns
            if '年级' not in df.columns or '班级' not in df.columns:
                raise ValueError("Excel 文件必须包含 '年级' 和 '班级' 列")
            
            # Create output directory
            os.makedirs(output_dir, exist_ok=True)
            
            # Process each grade
            grades = df['年级'].unique()
            for grade in grades:
                print(f"处理年级: {grade}")
                grade_folder = os.path.join(output_dir, str(grade))
                os.makedirs(grade_folder, exist_ok=True)
                
                # Process each class in the grade
                grade_df = df[df['年级'] == grade]
                classes = grade_df['班级'].unique()
                print(f"该年级班级: {classes}")
                
                for class_name in classes:
                    print(f"生成文件: {grade}{class_name}.xlsx")
                    class_df = grade_df[grade_df['班级'] == class_name]
                    output_file = os.path.join(grade_folder, f"{grade}{class_name}.xlsx")
                    class_df.to_excel(output_file, index=False, engine='openpyxl')
            
            self.status_label.config(text=f"处理完成！文件已保存至 {output_dir}")
            messagebox.showinfo("成功", f"处理完成！文件已保存至 {output_dir}")
            
        except Exception as e:
            error_msg = f"处理过程中发生错误: {str(e)}"
            print(error_msg)
            self.status_label.config(text=error_msg)
            messagebox.showerror("错误", error_msg)

    def open_output_folder(self):
        output_dir = self.output_dir.get()
        if output_dir and os.path.exists(output_dir):
            # Normalize path for Windows
            output_dir = os.path.normpath(output_dir)
            print(f"Opening folder: {output_dir}")
            subprocess.run(['explorer', output_dir])
            self.status_label.config(text=f"已打开输出文件夹: {output_dir}")
        else:
            messagebox.showerror("错误", "输出文件夹不存在或未选择！")
            self.status_label.config(text="错误：输出文件夹不存在或未选择")

def main():
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()