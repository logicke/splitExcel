import pandas as pd
import os
from tkinter import Tk, filedialog

def select_input_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def select_output_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def create_grade_class_folders(input_file, output_dir):
    try:
        # 读取 Excel 文件，明确指定 engine='openpyxl'
        df = pd.read_excel(input_file, engine='openpyxl')
        
        # 清理数据
        df['年级'] = df['年级'].astype(str).str.strip()
        df['班级'] = df['班级'].astype(str).str.strip()
        
        # 打印数据概况
        print("读取的数据行数:", len(df))
        print("所有年级:", df['年级'].unique())
        
        # 确保存在年级和班级列
        if '年级' not in df.columns or '班级' not in df.columns:
            raise ValueError("Excel 文件必须包含 '年级' 和 '班级' 列")
        
        # 创建默认输出目录 OutFiles
        os.makedirs(output_dir, exist_ok=True)
        
        # 获取所有年级
        grades = df['年级'].unique()
        
        # 遍历每个年级
        for grade in grades:
            print(f"处理年级: {grade}")
            # 创建年级文件夹
            grade_folder = os.path.join(output_dir, str(grade))
            os.makedirs(grade_folder, exist_ok=True)
            
            # 获取该年级的所有班级
            grade_df = df[df['年级'] == grade]
            classes = grade_df['班级'].unique()
            print(f"该年级班级: {classes}")
            
            # 遍历每个班级
            for class_name in classes:
                print(f"生成文件: {grade}{class_name}.xlsx")
                # 获取该班级的数据
                class_df = grade_df[grade_df['班级'] == class_name]
                
                # 创建输出文件路径
                output_file = os.path.join(grade_folder, f"{grade}{class_name}.xlsx")
                
                # 保存到新的 Excel 文件
                class_df.to_excel(output_file, index=False, engine='openpyxl')
                
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
    
    # 设置默认输出目录为脚本所在目录下的 OutFiles
    script_dir = select_output_file()
    output_dir = os.path.join(script_dir, "OutFiles")
    
    # 处理文件
    create_grade_class_folders(input_file, output_dir)

if __name__ == "__main__":
    main()