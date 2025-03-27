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
        
        # Setup style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#4a86e8")
        self.style.configure("Run.TButton", background="#4CAF50")
        self.style.configure("Accent.TButton", background="#4CAF50")
        
        # Variables
        self.api_key = tk.StringVar()
        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_folder.set(os.path.expanduser("~/Documents"))  # Default to Documents folder
        self.generated_code = ""
        self.excel_file_path = ""
        self.model = tk.StringVar(value="gemini-2.5-pro-exp-03-25")  # Default model
        
        # Setup GUI
        self._setup_ui()
        self._load_api_key()
        
    def _setup_ui(self):
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # API Key Section
        api_frame = ttk.LabelFrame(main_frame, text="API Settings")
        api_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(api_frame, text="Gemini API Key:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key, width=50, show="*")
        api_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Add model selection
        ttk.Label(api_frame, text="Model:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        model_combobox = ttk.Combobox(api_frame, textvariable=self.model, state="readonly")
        model_combobox['values'] = ("gemini-2.5-pro-exp-03-25", "gemini-2.0-flash-thinking-exp-01-21", "gemini-2.0-flash")
        model_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        # Add show/hide API key button
        self.show_key = tk.BooleanVar(value=False)
        ttk.Checkbutton(api_frame, text="Show key", variable=self.show_key, 
                       command=lambda: api_entry.config(show="" if self.show_key.get() else "*")).grid(
                       row=0, column=2, padx=5, pady=5)
        
        # Input file selection
        file_frame = ttk.LabelFrame(main_frame, text="Select Input File")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(file_frame, text="Select PDF/Image", command=self._select_file).grid(
            row=0, column=0, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.input_file).grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Output folder selection
        output_frame = ttk.LabelFrame(main_frame, text="Excel Output Folder")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(output_frame, text="Select Folder", command=self._select_folder).grid(
            row=0, column=0, padx=5, pady=5)
        ttk.Label(output_frame, textvariable=self.output_folder).grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")
        
        # Prompt input
        prompt_frame = ttk.LabelFrame(main_frame, text="Processing Request")
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.prompt_text = scrolledtext.ScrolledText(prompt_frame, height=5)
        self.prompt_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.prompt_text.insert(tk.END, "Read file then create code to create Excel file with full data from image without editing or deleting anything, full text.")
        
        # Generated code display
        code_frame = ttk.LabelFrame(main_frame, text="Generated Code")
        code_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.code_view = scrolledtext.ScrolledText(code_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.code_view.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress.pack(fill=tk.X, pady=5)
        
        # Control buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        # Create sub-frame for left buttons
        left_btn_frame = ttk.Frame(btn_frame)
        left_btn_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)

        # Create sub-frame for right buttons
        right_btn_frame = ttk.Frame(btn_frame)
        right_btn_frame.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Main function buttons
        self.prompt_btn = ttk.Button(
            left_btn_frame,
            text="Run Prompt", 
            command=self.run_prompt,
            width=15
        )
        self.prompt_btn.pack(side=tk.LEFT, padx=5, pady=2)

        self.run_btn = ttk.Button(
            left_btn_frame,
            text="Run Code", 
            command=self.run_code,
            state=tk.DISABLED,
            width=15
        )
        self.run_btn.pack(side=tk.LEFT, padx=5, pady=2)

        self.retry_btn = ttk.Button(
            left_btn_frame,
            text="Retry Prompt", 
            command=self.retry_prompt,
            width=15
        )
        self.retry_btn.pack(side=tk.LEFT, padx=5, pady=2)

        # Open Excel and Reset buttons
        self.open_excel_btn = ttk.Button(
            right_btn_frame,
            text="Open Folder", 
            command=self._open_output_folder,
            width=15
        )
        self.open_excel_btn.pack(side=tk.RIGHT, padx=5, pady=2)

        ttk.Button(
            right_btn_frame,
            text="Reset", 
            command=self._reset,
            width=10
        ).pack(side=tk.RIGHT, padx=5, pady=2)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=5)
        
        # Configure grid for frames
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
        self.status_var.set("Processing file...")
        self.root.update()
        
        try:
            # Read and encode file
            file_path = self.input_file.get()
            with open(file_path, "rb") as f:
                file_data = base64.b64encode(f.read()).decode("utf-8")
            
            self.progress_var.set(20)
            self.status_var.set("Creating prompt...")
            self.root.update()
            
            # Create prompt
            prompt = self._build_prompt(file_path)
            
            self.progress_var.set(30)
            self.status_var.set("Sending request to Gemini API...")
            self.root.update()
            
            # Call Gemini API
            response = self._call_gemini_api(prompt, file_data)
            
            self.progress_var.set(80)
            self.status_var.set("Extracting code...")
            self.root.update()
            
            # Extract code
            self.generated_code = self._extract_code(response)
            
            # Display code
            self.code_view.delete(1.0, tk.END)
            self.code_view.insert(tk.END, self.generated_code)
            self.run_btn.config(state=tk.NORMAL)
            
            self.progress_var.set(100)
            
            # Save API key
            self._save_api_key()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating code: {str(e)}")
            self.status_var.set("An error occurred")
    
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
            f"https://generativelanguage.googleapis.com/v1beta/models/{self.model.get()}:generateContent?key={self.api_key.get()}",
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
            raise ValueError(f"Unsupported file format: {ext}")
            
        return mime_types[ext]
    
    def _extract_code(self, response):
        try:
            # Extract text from API response
            content_parts = response["candidates"][0]["content"]["parts"]
            full_text = ""
            for part in content_parts:
                if "text" in part:
                    full_text += part["text"]
            
            # Remove markdown syntax
            # Case 1: Has form ```python ... ```
            pattern1 = r'```python\s*(.*?)\s*```'
            # Case 2: Has form ``` ... ```
            pattern2 = r'```\s*(.*?)\s*```'
            
            # Search with pattern 1 first
            matches = re.findall(pattern1, full_text, re.DOTALL)
            if matches:
                return matches[0].strip()
            
            # If not found, search with pattern 2
            matches = re.findall(pattern2, full_text, re.DOTALL)
            if matches:
                return matches[0].strip()
            
            # Clean code before returning (remove any remaining markdown indicators)
            clean_text = full_text.strip()
            clean_text = re.sub(r'^```python\s*', '', clean_text)
            clean_text = re.sub(r'^```', '', clean_text)
            clean_text = re.sub(r'\s*```$', '', clean_text)
            
            return clean_text
                
        except (KeyError, IndexError) as e:
            raise Exception(f"Cannot extract code from response: {str(e)}")
    
    def run_code(self):
        if not self.generated_code:
            messagebox.showerror("Error", "No code to execute")
            return
        
        self.status_var.set("Executing code...")
        self.progress_var.set(0)
        self.root.update()
        
        # Ensure output directory exists
        output_dir = os.path.dirname(self.excel_file_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        try:
            # Update progress
            for i in range(5):
                self.progress_var.set(i * 20)
                self.root.update()
            
            # Capture all output for debugging
            old_stdout = sys.stdout
            from io import StringIO
            captured_output = StringIO()
            sys.stdout = captured_output
            
            # Execute code in the same process
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
            
            # Check if file was created
            if os.path.exists(self.excel_file_path):
                messagebox.showinfo("Success", f"Excel file created successfully!\n\nPath: {self.excel_file_path}")
                self.status_var.set("Excel file created successfully")
                self.open_excel_btn.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("Warning", f"Code ran but Excel file not found.\nCheck output log: {execution_log}")
                self.status_var.set("Output Excel file not found")
            
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("Error", f"Code execution failed:\n\n{error_msg}")
            self.status_var.set("Code execution failed")

    
    def _validate_inputs(self):
        if not self.api_key.get():
            messagebox.showwarning("Warning", "Please enter API Key")
            return False
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file")
            return False
        if not self.output_folder.get():
            messagebox.showwarning("Warning", "Please select an output folder")
            return False
        if not os.path.exists(self.input_file.get()):
            messagebox.showwarning("Warning", "Input file does not exist")
            return False
        
        # Automatically create output folder if it doesn't exist
        try:
            os.makedirs(self.output_folder.get(), exist_ok=True)
            return True
        except:
            messagebox.showwarning("Warning", "Cannot create output folder")
            return False
    
    def _open_output_folder(self):
        if os.path.exists(self.output_folder.get()):
            webbrowser.open(self.output_folder.get())
        else:
            messagebox.showwarning("Warning", "Output folder not found")
    
    def _save_api_key(self):
        try:
            config_dir = Path.home() / ".excel_converter"
            config_dir.mkdir(exist_ok=True)
            config_file = config_dir / "config.json"
            
            config = {"api_key": self.api_key.get()}
            with open(config_file, "w") as f:
                json.dump(config, f)
        except:
            # Ignore errors if unable to save
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
            # Ignore errors if unable to read
            pass
    
    def _reset(self):
        # Clear input fields
        self.input_file.set("")
        self.code_view.delete(1.0, tk.END)
        self.prompt_text.delete(1.0, tk.END)
        self.prompt_text.insert(tk.END, "Create Excel file from tables in the document, preserving structure and data format.")
        self.progress_var.set(0)
        self.status_var.set("Input fields reset")
        self.run_btn.config(state=tk.DISABLED)
        self.open_excel_btn.config(state=tk.DISABLED)
        
    def run_prompt(self):
        # Generate code from prompt
        self.generate_code()

    def retry_prompt(self):
        # Clear old code and rerun prompt
        self.code_view.delete(1.0, tk.END)
        self.status_var.set("Retrying prompt...")
        self.progress_var.set(0)
        self.root.update()
        self.generate_code()

if __name__ == "__main__":
    root = tk.Tk()
    app = GeminiExcelConverter(root)
    app.api_key.set("")  # Set API key
    root.mainloop()
