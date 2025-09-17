import pandas as pd
import os
from tkinter import Tk, filedialog
import shutil

def select_input_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择要处理的 Excel 文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path

def select_output_directory():
    root = Tk()
    root.withdraw()
    directory = filedialog.askdirectory(
       title="选择要保存文件的文件夹" 
    )
    return directory

def create_grade_class_folders(input_file, output_dir):
    try:
        # 读取 Excel 文件
        df = pd.read_excel(input_file)
        
        # 确保存在年级和班级列
        if '年级' not in df.columns or '班级' not in df.columns:
            raise ValueError("Excel 文件必须包含 '年级' 和 '班级' 列")
        
        # 获取所有年级
        grades = df['年级'].unique()
        
        # 遍历每个年级
        for grade in grades:
            # 创建年级文件夹
            grade_folder = os.path.join(output_dir, str(grade))
            os.makedirs(grade_folder, exist_ok=True)
            
            # 获取该年级的所有班级
            grade_df = df[df['年级'] == grade]
            classes = grade_df['班级'].unique()
            
            # 遍历每个班级
            for class_name in classes:
                # 获取该班级的数据
                class_df = grade_df[grade_df['班级'] == class_name]
                
                # 创建输出文件路径
                output_file = os.path.join(grade_folder, f"{grade}{class_name}.xlsx")
                
                # 保存到新的 Excel 文件
                class_df.to_excel(output_file, index=False)
                
        print(f"处理完成！文件已保存至 {output_dir}")
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")

def main():
    # 选择输入文件
    print("请选择要处理的 Excel 文件...")
    input_file = select_input_file()
    if not input_file:
        print("未选择文件，程序退出")
        return
    
    # 选择输出目录
    print("请选择输出文件夹...")    
    output_dir = os.path.join(select_output_directory(), "OutFiles")
    if not output_dir:
        print("未选择输出文件夹，程序退出")
        return
    
    # 处理文件
    create_grade_class_folders(input_file, output_dir)

if __name__ == "__main__":
    main()