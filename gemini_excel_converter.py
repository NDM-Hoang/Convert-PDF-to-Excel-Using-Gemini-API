import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import base64
import json
import os
import requests
import subprocess
import sys
import tempfile
import webbrowser
import re
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

class GeminiExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Excel Converter")
        self.root.geometry("900x900")
        
        # Thiết lập style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#4a86e8")
        self.style.configure("Run.TButton", background="#4CAF50")
        self.style.configure("Accent.TButton", background="#4CAF50")
        
        # Biến lưu trữ
        self.api_key = tk.StringVar()
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_folder.set(os.path.expanduser("~/Documents"))  # Mặc định thư mục Documents
        self.generated_code = ""
        self.excel_file_path = ""
        
        # Thiết lập GUI
        self._setup_ui()
        self._load_api_key()
        
    def _setup_ui(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Phần API Key
        api_frame = ttk.LabelFrame(main_frame, text="Cài đặt API")
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="Gemini API Key:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key, width=50, show="*")
        api_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Thêm button hiện/ẩn API key
        self.show_key = tk.BooleanVar(value=False)
        ttk.Checkbutton(api_frame, text="Hiện key", variable=self.show_key, 
                       command=lambda: api_entry.config(show="" if self.show_key.get() else "*")).grid(
                       row=0, column=2, padx=5, pady=5)
        
        # Phần chọn file
        file_frame = ttk.LabelFrame(main_frame, text="Chọn file đầu vào")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(file_frame, text="Chọn PDF/Ảnh", command=self._select_file).grid(
            row=0, column=0, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.input_file).grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Phần thư mục đầu ra
        output_frame = ttk.LabelFrame(main_frame, text="Thư mục lưu file Excel")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(output_frame, text="Chọn thư mục", command=self._select_folder).grid(
            row=0, column=0, padx=5, pady=5)
        ttk.Label(output_frame, textvariable=self.output_folder).grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Phần nhập prompt
        prompt_frame = ttk.LabelFrame(main_frame, text="Yêu cầu xử lý")
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.prompt_text = scrolledtext.ScrolledText(prompt_frame, height=5)
        self.prompt_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.prompt_text.insert(tk.END, "Read file then create code to create Excel file with full data from image without editing or deleting anything, full text.")
        
        # Phần code sinh ra
        code_frame = ttk.LabelFrame(main_frame, text="Code sinh ra")
        code_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.code_view = scrolledtext.ScrolledText(code_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.code_view.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Phần tiến trình
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress.pack(fill=tk.X, pady=5)
        
        # Nút điều khiển
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        # Tạo khung con cho các nút bên trái
        left_btn_frame = ttk.Frame(btn_frame)
        left_btn_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Tạo khung con cho các nút bên phải
        right_btn_frame = ttk.Frame(btn_frame)
        right_btn_frame.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Các nút chức năng chính
        self.prompt_btn = ttk.Button(
            left_btn_frame,
            text="Chạy Prompt", 
            command=self.run_prompt,
            width=15
        )
        self.prompt_btn.pack(side=tk.LEFT, padx=5, pady=2)

        self.run_btn = ttk.Button(
            left_btn_frame,
            text="Chạy Code", 
            command=self.run_code,
            state=tk.DISABLED,
            width=15
        )
        self.run_btn.pack(side=tk.LEFT, padx=5, pady=2)

        self.retry_btn = ttk.Button(
            left_btn_frame,
            text="Chạy lại Prompt", 
            command=self.retry_prompt,
            width=15
        )
        self.retry_btn.pack(side=tk.LEFT, padx=5, pady=2)

        # Nút mở Excel và Reset
        self.open_excel_btn = ttk.Button(
            right_btn_frame,
            text="Mở File Excel", 
            command=self._open_excel_file,
            state=tk.DISABLED,
            width=15
        )
        self.open_excel_btn.pack(side=tk.RIGHT, padx=5, pady=2)

        ttk.Button(
            right_btn_frame,
            text="Reset", 
            command=self._reset,
            width=10
        ).pack(side=tk.RIGHT, padx=5, pady=2)

        # Thanh trạng thái
        self.status_var = tk.StringVar(value="Sẵn sàng")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=5)
        
        # Thiết lập Grid cho các frame
        for frame in [api_frame, file_frame, output_frame]:
            frame.grid_columnconfigure(1, weight=1)
        
    def _select_file(self):
        filetypes = (
            ("PDF files", "*.pdf"),
            ("Image files", "*.png *.jpg *.jpeg"),
            ("All files", "*.*")
        )
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.input_file.set(filename)
            
    def _select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)
    
    def generate_code(self):
        if not self._validate_inputs():
            return
        
        self.progress_var.set(0)
        self.status_var.set("Đang xử lý file...")
        self.root.update()
        
        try:
            # Đọc và mã hóa file
            file_path = self.input_file.get()
            with open(file_path, "rb") as f:
                file_data = base64.b64encode(f.read()).decode("utf-8")
            
            self.progress_var.set(20)
            self.status_var.set("Đang tạo prompt...")
            self.root.update()
            
            # Tạo prompt
            prompt = self._build_prompt(file_path)
            
            self.progress_var.set(30)
            self.status_var.set("Đang gửi yêu cầu đến Gemini API...")
            self.root.update()
            
            # Gọi Gemini API
            response = self._call_gemini_api(prompt, file_data)
            
            self.progress_var.set(80)
            self.status_var.set("Đang trích xuất code...")
            self.root.update()
            
            # Trích xuất code
            self.generated_code = self._extract_code(response)
            
            # Hiển thị code
            self.code_view.delete(1.0, tk.END)
            self.code_view.insert(tk.END, self.generated_code)
            self.run_btn.config(state=tk.NORMAL)
            
            self.progress_var.set(100)
            self.status_var.set("Đã tạo code thành công. Kiểm tra và nhấn 'Chạy Code' để tiếp tục.")
            
            # Lưu API key
            self._save_api_key()
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi tạo code: {str(e)}")
            self.status_var.set("Đã xảy ra lỗi")
    
    def _build_prompt(self, file_path):
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(self.output_folder.get(), f"{base_name}.xlsx")
        self.excel_file_path = output_path
        
        user_prompt = self.prompt_text.get(1.0, tk.END).strip()
        
        return f"""
        I need Python code that extracts all text data from the attached file and creates an Excel file following this EXACT code structure:

        ```
        import openpyxl
        from openpyxl.styles import Font, Alignment, Border, Side
        import os

        def create_excel_report(output_path):
            try:
                # Create output directory
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                # Create workbook
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Extracted Data"

                # Define styles
                bold_font = Font(bold=True)
                center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin'))

                # YOUR CODE HERE: Extract and format all content from the source file
                
                # Save file
                wb.save(output_path)
                print(f"Successfully created Excel file: {{output_path}}")
                return True
            except Exception as e:
                print(f"Error: {{e}}")
                return False

        # Execution
        output_file_path = r"{output_path.replace('\\', '\\\\')}"
        create_excel_report(output_file_path)
        ```

        Requirements:
        1. Extract ALL text/tables from the file
        2. Format with proper headings, alignment, borders
        3. DO NOT change function structure or file path
        4. MUST include wb.save() with the exact output path
        
        User instructions: {user_prompt}
        """
    
    def _call_gemini_api(self, prompt, file_data):
        headers = {
            "Content-Type": "application/json",
        }
        
        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": self._get_mime_type(),
                            "data": file_data
                        }
                    }
                ]
            }],
            "generationConfig": {
                "temperature": 0.3,
                "topP": 0.95,
                "maxOutputTokens": 8192
            }
        }
        
        response = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro-exp-03-25:generateContent?key={self.api_key.get()}",
            headers=headers,
            json=payload
        )
        
        if response.status_code != 200:
            raise Exception(f"API Error: {response.text}")
            
        return response.json()
    
    def _get_mime_type(self):
        ext = os.path.splitext(self.input_file.get())[1].lower()
        mime_types = {
            ".pdf": "application/pdf",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png"
        }
        
        if ext not in mime_types:
            raise ValueError(f"Định dạng file không được hỗ trợ: {ext}")
            
        return mime_types[ext]
    
    def _extract_code(self, response):
        try:
            # Lấy text từ phản hồi API
            content_parts = response["candidates"][0]["content"]["parts"]
            full_text = ""
            for part in content_parts:
                if "text" in part:
                    full_text += part["text"]
            
            # Xóa bỏ cú pháp markdown
            # Trường hợp 1: Có dạng ```python ... ```
            pattern1 = r'```python\s*(.*?)\s*```'
            # Trường hợp 2: Có dạng ``` ... ```
            pattern2 = r'```\s*(.*?)\s*```'
            
            # Tìm kiếm theo pattern 1 trước
            matches = re.findall(pattern1, full_text, re.DOTALL)
            if matches:
                return matches[0].strip()
            
            # Nếu không tìm thấy, tìm theo pattern 2
            matches = re.findall(pattern2, full_text, re.DOTALL)
            if matches:
                return matches[0].strip()
            
            # Làm sạch mã trước khi trả về (xóa các dấu hiệu markdown còn sót)
            clean_text = full_text.strip()
            clean_text = re.sub(r'^```python\s*', '', clean_text)
            clean_text = re.sub(r'^```', '', clean_text)
            clean_text = re.sub(r'\s*```$', '', clean_text)
            
            return clean_text
                
        except (KeyError, IndexError) as e:
            raise Exception(f"Không thể trích xuất code từ phản hồi: {str(e)}")
    
    def run_code(self):
        if not self.generated_code:
            messagebox.showerror("Lỗi", "Không có code để thực thi")
            return
        
        self.status_var.set("Đang thực thi code...")
        self.progress_var.set(0)
        self.root.update()
        
        # Đảm bảo thư mục đầu ra tồn tại
        output_dir = os.path.dirname(self.excel_file_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        try:
            # Cập nhật tiến trình
            for i in range(5):
                self.progress_var.set(i * 20)
                self.root.update()
            
            # Bắt toàn bộ output để debug
            old_stdout = sys.stdout
            from io import StringIO
            captured_output = StringIO()
            sys.stdout = captured_output
            
            # Thực thi code trong cùng tiến trình
            namespace = {
                'os': os,
                'sys': sys,
                'Path': Path,
                'openpyxl': __import__('openpyxl')
            }
            
            try:
                exec(self.generated_code, namespace)
            finally:
                sys.stdout = old_stdout
                
            execution_log = captured_output.getvalue()
            print(execution_log)  # For debugging
            
            self.progress_var.set(100)
            
            # Kiểm tra file đã được tạo hay chưa
            if os.path.exists(self.excel_file_path):
                messagebox.showinfo("Thành công", f"Đã tạo file Excel thành công!\n\nĐường dẫn: {self.excel_file_path}")
                self.status_var.set("Đã tạo file Excel thành công")
                self.open_excel_btn.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("Cảnh báo", f"Code đã chạy nhưng không tìm thấy file Excel.\nKiểm tra output log: {execution_log}")
                self.status_var.set("Không tìm thấy file Excel đầu ra")
            
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("Lỗi", f"Thực thi code thất bại:\n\n{error_msg}")
            self.status_var.set("Thực thi code thất bại")

    
    def _validate_inputs(self):
        if not self.api_key.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập API Key")
            return False
        if not self.input_file.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file đầu vào")
            return False
        if not self.output_folder.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục đầu ra")
            return False
        if not os.path.exists(self.input_file.get()):
            messagebox.showwarning("Cảnh báo", "File đầu vào không tồn tại")
            return False
        
        # Tự động tạo thư mục đầu ra nếu chưa tồn tại
        try:
            os.makedirs(self.output_folder.get(), exist_ok=True)
            return True
        except:
            messagebox.showwarning("Cảnh báo", "Không thể tạo thư mục đầu ra")
            return False
    
    def _open_excel_file(self):
        if os.path.exists(self.excel_file_path):
            webbrowser.open(self.excel_file_path)
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file Excel")
    
    def _save_api_key(self):
        try:
            config_dir = Path.home() / ".excel_converter"
            config_dir.mkdir(exist_ok=True)
            config_file = config_dir / "config.json"
            
            config = {"api_key": self.api_key.get()}
            with open(config_file, "w") as f:
                json.dump(config, f)
        except:
            # Bỏ qua lỗi nếu không thể lưu
            pass
    
    def _load_api_key(self):
        try:
            config_file = Path.home() / ".excel_converter" / "config.json"
            if config_file.exists():
                with open(config_file, "r") as f:
                    config = json.load(f)
                    if "api_key" in config:
                        self.api_key.set(config["api_key"])
        except:
            # Bỏ qua lỗi nếu không thể đọc
            pass
    
    def _reset(self):
        # Xóa các trường nhập liệu
        self.input_file.set("")
        self.code_view.delete(1.0, tk.END)
        self.prompt_text.delete(1.0, tk.END)
        self.prompt_text.insert(tk.END, "Tạo file Excel từ các bảng trong tài liệu, giữ nguyên cấu trúc và định dạng dữ liệu.")
        self.progress_var.set(0)
        self.status_var.set("Đã reset các trường nhập liệu")
        self.run_btn.config(state=tk.DISABLED)
        self.open_excel_btn.config(state=tk.DISABLED)
        
    def run_prompt(self):
        # Thực hiện tạo code từ prompt
        self.generate_code()

    def retry_prompt(self):
        # Xóa code cũ và chạy lại prompt
        self.code_view.delete(1.0, tk.END)
        self.status_var.set("Đang chạy lại prompt...")
        self.progress_var.set(0)
        self.root.update()
        self.generate_code()

if __name__ == "__main__":
    root = tk.Tk()
    app = GeminiExcelConverter(root)
    app.api_key.set("")  # Thiết lập API key
    root.mainloop()
