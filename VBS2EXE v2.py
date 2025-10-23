import os
import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QProgressBar, QLineEdit
)
from PyQt6.QtCore import Qt

def generate_cpp_wrapper(vbs_path):
    with open(vbs_path, 'r', encoding='utf-8') as f:
        vbs_content = f.read()
    escaped_vbs = vbs_content.replace('\\', '\\\\').replace('"', '\\"')
    cpp_code = f'''
#include <windows.h>
#include <fstream>
#include <string>
#include <shlobj.h>

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int) {{
    char tempPath[MAX_PATH];
    GetTempPathA(MAX_PATH, tempPath);
    std::string tempFile = std::string(tempPath) + "embedded_script.vbs";

    std::ofstream out(tempFile);
    out << "{escaped_vbs}";
    out.close();

    ShellExecuteA(NULL, "open", "wscript.exe", tempFile.c_str(), NULL, SW_HIDE);
    return 0;
}}
'''
    cpp_output_path = os.path.join(os.path.dirname(vbs_path), 'vbs_embedded_wrapper.cpp')
    with open(cpp_output_path, 'w', encoding='utf-8') as f:
        f.write(cpp_code)
    return cpp_output_path

def compile_cpp_to_exe(cpp_path, output_path):
    compile_command = f'g++ "{cpp_path}" -o "{output_path}" -mwindows'
    result = os.system(compile_command)
    return result == 0

class VBSCompiler(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VBS to EXE Compiler")
        self.setGeometry(100, 100, 500, 250)

        layout = QVBoxLayout()

        self.label = QLabel("Select a .vbs file to compile into .exe")
        layout.addWidget(self.label)

        self.select_button = QPushButton("Choose VBS File")
        self.select_button.clicked.connect(self.select_vbs_file)
        layout.addWidget(self.select_button)

        self.output_name_input = QLineEdit()
        self.output_name_input.setPlaceholderText("Enter EXE name (without .exe)")
        layout.addWidget(self.output_name_input)

        self.output_folder_button = QPushButton("Choose Output Folder")
        self.output_folder_button.clicked.connect(self.select_output_folder)
        layout.addWidget(self.output_folder_button)

        self.output_folder_label = QLabel("No folder selected")
        layout.addWidget(self.output_folder_label)

        self.compile_button = QPushButton("Compile to EXE")
        self.compile_button.clicked.connect(self.compile)
        layout.addWidget(self.compile_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)
        self.vbs_path = ""
        self.output_folder = ""

    def select_vbs_file(self):
        vbs_path, _ = QFileDialog.getOpenFileName(self, "Select VBS File", "", "VBScript Files (*.vbs)")
        if vbs_path:
            self.vbs_path = vbs_path
            self.label.setText(f"Selected: {os.path.basename(vbs_path)}")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_folder_label.setText(f"Output Folder: {folder}")

    def compile(self):
        if not self.vbs_path:
            QMessageBox.warning(self, "Warning", "Please select a VBS file.")
            return

        exe_name = self.output_name_input.text().strip()
        if not exe_name:
            QMessageBox.warning(self, "Warning", "Please enter a name for the output EXE.")
            return

        if not self.output_folder:
            QMessageBox.warning(self, "Warning", "Please select an output folder.")
            return

        self.progress_bar.setValue(10)
        try:
            cpp_path = generate_cpp_wrapper(self.vbs_path)
            self.progress_bar.setValue(40)

            output_exe_path = os.path.join(self.output_folder, exe_name + ".exe")
            success = compile_cpp_to_exe(cpp_path, output_exe_path)
            self.progress_bar.setValue(100 if success else 0)

            if success:
                QMessageBox.information(self, "Success", f"Executable created at:\n{output_exe_path}")
            else:
                QMessageBox.critical(self, "Compilation Failed", "Failed to compile the generated C++ file.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self.progress_bar.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VBSCompiler()
    window.show()
    sys.exit(app.exec())