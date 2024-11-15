from PyQt6 import uic

def convert_ui_to_py(input_ui_file, output_py_file):
    with open(output_py_file, 'w', encoding='utf-8') as py_file:
        uic.compileUi(input_ui_file, py_file)
        print(f"Chuyển đổi {input_ui_file} thành {output_py_file} thành công!")

# Đường dẫn đến file .ui và file .py
input_ui_file = "E:\\MyProject\\MayChamCong\\GUI_MAY_CHAM_CONG\\gui_main.ui"
output_py_file = "E:\\MyProject\\MayChamCong\\GUI_MAY_CHAM_CONG\\gui_main.py"


convert_ui_to_py(input_ui_file, output_py_file)

