import os
import tarfile
import zipfile
import tempfile
import pandas as pd
import numpy as np
import h5py
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path


class FileToLatexConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("File to LaTeX Table Converter")
        self.root.geometry("800x600")
        
        self.setup_ui()
        
        # Temporary directory for extracted files
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def on_closing(self):
        """Clean up temporary directory on application close"""
        self.temp_dir.cleanup()
        self.root.destroy()
        
    def setup_ui(self):
        """Set up the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="Input", padding="10")
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(input_frame, text="Select File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Button(input_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, pady=5)
        
        # Options section
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        # File type detection
        ttk.Label(options_frame, text="File Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_type_var = tk.StringVar(value="auto")
        file_types = ["auto", "csv", "h5", "text", "tar", "zip"]
        ttk.Combobox(options_frame, textvariable=self.file_type_var, values=file_types, width=15).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # LaTeX table options
        ttk.Label(options_frame, text="Table Caption:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.caption_var = tk.StringVar(value="Average Teaching Evaluation Scores for Fall and Spring Semesters (2018-2024) on a 5-point Scale.")
        ttk.Entry(options_frame, textvariable=self.caption_var, width=40).grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Label(options_frame, text="Table Label:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.label_var = tk.StringVar(value="tab:data")
        ttk.Entry(options_frame, textvariable=self.label_var, width=20).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Include caption checkbox
        self.include_caption_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include Caption", variable=self.include_caption_var).grid(row=3, column=0, sticky=tk.W, pady=5)
        
        # Table style selection
        ttk.Label(options_frame, text="Table Style:").grid(row=3, column=1, sticky=tk.W, pady=5)
        self.table_style_var = tk.StringVar(value="booktabs")
        table_styles = ["booktabs", "standard"]
        ttk.Combobox(options_frame, textvariable=self.table_style_var, values=table_styles, width=15).grid(row=3, column=2, sticky=tk.W, pady=5)
        
        # Max rows
        ttk.Label(options_frame, text="Max Rows:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.max_rows_var = tk.StringVar(value="50")
        ttk.Entry(options_frame, textvariable=self.max_rows_var, width=10).grid(row=4, column=1, sticky=tk.W, pady=5)
                
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Convert", command=self.convert_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Preview", command=self.preview_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save LaTeX", command=self.save_latex).pack(side=tk.LEFT, padx=5)
        
        # Output area
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create a notebook with tabs for data preview and LaTeX output
        self.notebook = ttk.Notebook(output_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Data preview tab
        preview_frame = ttk.Frame(self.notebook)
        self.notebook.add(preview_frame, text="Data Preview")
        
        self.preview_text = tk.Text(preview_frame, wrap=tk.NONE, font=("Courier", 10))
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbars for the preview
        preview_y_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        preview_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text['yscrollcommand'] = preview_y_scroll.set
        
        preview_x_scroll = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_text.xview)
        preview_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_text['xscrollcommand'] = preview_x_scroll.set
        
        # LaTeX output tab
        latex_frame = ttk.Frame(self.notebook)
        self.notebook.add(latex_frame, text="LaTeX Output")
        
        self.latex_text = tk.Text(latex_frame, wrap=tk.NONE, font=("Courier", 10))
        self.latex_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbars for the LaTeX output
        latex_y_scroll = ttk.Scrollbar(latex_frame, orient=tk.VERTICAL, command=self.latex_text.yview)
        latex_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.latex_text['yscrollcommand'] = latex_y_scroll.set
        
        latex_x_scroll = ttk.Scrollbar(latex_frame, orient=tk.HORIZONTAL, command=self.latex_text.xview)
        latex_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.latex_text['xscrollcommand'] = latex_x_scroll.set
        
        # Status bar
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, side=tk.BOTTOM)
        
        # Initialize variables
        self.data = None
        self.latex_code = ""
        
    def browse_file(self):
        """Open file browser dialog to select a file"""
        filetypes = [
            ("All Supported Files", "*.csv *.h5 *.txt *.tar *.gz *.zip"),
            ("CSV Files", "*.csv"),
            ("HDF5 Files", "*.h5"),
            ("Text Files", "*.txt"),
            ("TAR Files", "*.tar *.gz"),
            ("ZIP Files", "*.zip"),
            ("All Files", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"File selected: {os.path.basename(file_path)}")
            
            # Auto-detect file type
            extension = os.path.splitext(file_path)[1].lower()
            if extension == ".csv":
                self.file_type_var.set("csv")
            elif extension == ".h5":
                self.file_type_var.set("h5")
            elif extension == ".txt":
                self.file_type_var.set("text")
            elif extension in [".tar", ".gz"]:
                self.file_type_var.set("tar")
            elif extension == ".zip":
                self.file_type_var.set("zip")
    
    def extract_archive(self, file_path):
        """Extract files from a tar or zip archive"""
        file_type = self.file_type_var.get()
        temp_path = self.temp_dir.name
        
        try:
            if file_type == "auto":
                extension = os.path.splitext(file_path)[1].lower()
                if extension in [".tar", ".gz"]:
                    file_type = "tar"
                elif extension == ".zip":
                    file_type = "zip"
            
            if file_type == "tar":
                with tarfile.open(file_path, "r:*") as tar:
                    tar.extractall(path=temp_path)
            elif file_type == "zip":
                with zipfile.ZipFile(file_path, "r") as zip_ref:
                    zip_ref.extractall(temp_path)
            
            # Return a list of extracted files
            extracted_files = []
            for root, _, files in os.walk(temp_path):
                for file in files:
                    extracted_files.append(os.path.join(root, file))
            
            if not extracted_files:
                raise Exception("No files found in archive")
                
            return extracted_files
        
        except Exception as e:
            messagebox.showerror("Extraction Error", f"Failed to extract archive: {str(e)}")
            return []
    
    def load_file(self, file_path):
        """Load data from the selected file"""
        file_type = self.file_type_var.get()
        
        try:
            # Check if we need to extract an archive first
            if file_type in ["tar", "zip"] or (file_type == "auto" and any(file_path.endswith(ext) for ext in [".tar", ".gz", ".zip"])):
                extracted_files = self.extract_archive(file_path)
                if not extracted_files:
                    return None
                
                # For simplicity, use the first file in the archive
                # You might want to add a file selection dialog for multiple files
                file_path = extracted_files[0]
                # Auto-detect the file type from the extracted file
                file_type = "auto"
            
            # Determine file type if set to auto
            if file_type == "auto":
                extension = os.path.splitext(file_path)[1].lower()
                if extension == ".csv":
                    file_type = "csv"
                elif extension == ".h5":
                    file_type = "h5"
                else:
                    file_type = "text"
            
            # Load based on file type
            if file_type == "csv":
                data = pd.read_csv(file_path)
                return data
            
            elif file_type == "h5":
                with h5py.File(file_path, 'r') as f:
                    # HDF5 files can have complex structures
                    # For simplicity, we'll get the first dataset
                    datasets = []
                    
                    def find_datasets(name, obj):
                        if isinstance(obj, h5py.Dataset):
                            datasets.append((name, obj))
                    
                    f.visititems(find_datasets)
                    
                    if not datasets:
                        raise ValueError("No datasets found in HDF5 file")
                    
                    # Use the first dataset
                    name, dataset = datasets[0]
                    # Convert to pandas DataFrame if possible
                    try:
                        data = pd.DataFrame(dataset[()])
                        # If the dataset contains structured arrays, pandas will handle column names
                        return data
                    except:
                        # Fallback for simple arrays
                        data = pd.DataFrame(dataset[()])
                        return data
            
            elif file_type == "text":
                # Try to determine the delimiter in text files
                with open(file_path, 'r') as f:
                    sample = f.read(1024)
                
                # Check if it's tab-delimited
                if '\t' in sample:
                    data = pd.read_csv(file_path, sep='\t')
                # Check if it's comma-delimited
                elif ',' in sample:
                    data = pd.read_csv(file_path, sep=',')
                # Otherwise, treat as space-delimited
                else:
                    data = pd.read_csv(file_path, delim_whitespace=True, header=None)
                
                return data
            
            else:
                raise ValueError(f"Unsupported file type: {file_type}")
        
        except Exception as e:
            messagebox.showerror("File Loading Error", f"Failed to load file: {str(e)}")
            return None
    
    def format_booktabs_style(self, data):
        """
        Format data into a booktabs-style LaTeX table like the example
        This creates the specific visual style shown in the screenshot
        """
        try:
            # Get options from the UI
            caption = self.caption_var.get() if self.include_caption_var.get() else None
            label = self.label_var.get()
            
            # Limit number of rows if specified
            max_rows = int(self.max_rows_var.get()) if self.max_rows_var.get().isdigit() else None
            if max_rows and len(data) > max_rows:
                data = data.head(max_rows)
            
            # Generate the LaTeX code
            latex = []
            
            # Begin table environment
            latex.append("\\begin{table}[htbp]")
            latex.append("  \\begin{center}")
            latex.append("  \\centering")
            
            # Determine the number of columns and create a structure like the example
            num_cols = len(data.columns)
            
            # First column is for row identifiers (like years in the example)
            # Divide remaining columns into two semester groups if possible
            main_cols = num_cols - 1  # Exclude the first column
            
            if main_cols >= 2:
                # Split remaining columns into two groups
                first_group = main_cols // 2
                second_group = main_cols - first_group
                
                # Create the tabular environment with appropriate column formatting
                # First column is 'c', followed by first_group 'c's, then a vertical bar, then second_group 'c's
                alignment = "c " + "c" * first_group + " | " + "c" * second_group
            else:
                # If too few columns, just use all 'c' alignment
                alignment = "c" * num_cols
            
            latex.append(f"  \\begin{{tabular}}{{{alignment}}}")
            latex.append("    \\toprule")
            
            # Create the first header row with semester groups
            if main_cols >= 4:  # Need at least 2 columns per group
                first_header = [""]  # Empty cell in top left
                
                # Use column names if they exist, or create defaults
                first_group_name = "Fall Courses"
                second_group_name = "Spring Courses"
                
                # Add multicolumn headers
                first_header.append(f"\\multicolumn{{{first_group}}}{{c}}{{{first_group_name}}}")
                first_header.append(f"\\multicolumn{{{second_group}}}{{c}}{{{second_group_name}}}")
                
                latex.append(f"    {' & '.join(first_header)} \\\\")
                
                # Add cmidrule separators like in the example
                first_start = 2  # Columns are 1-indexed in LaTeX
                first_end = first_start + first_group - 1
                second_start = first_end + 1
                second_end = second_start + second_group - 1
                
                latex.append(f"    \\cmidrule(lr){{{first_start}-{first_end}}} \\cmidrule(lr){{{second_start}-{second_end}}}")
                
                # Add the subheader row with column names
                column_headers = [""] + list(data.columns)[1:]  # Skip first column for now
                latex.append(f"    {' & '.join(column_headers)} \\\\")
                
                # Add double midrule after headers
                latex.append("    \\midrule    \\midrule")
            else:
                # Simple header for tables with few columns
                headers = [str(col) for col in data.columns]
                latex.append(f"    {' & '.join(headers)} \\\\")
                latex.append("    \\midrule")
            
            # Add data rows
            for i, (_, row) in enumerate(data.iterrows()):
                row_values = []
                for val in row:
                    if pd.isna(val):
                        # Handle missing values with a dash like in the example
                        row_values.append("-")
                    elif isinstance(val, (int, np.integer)):
                        # Format integers without decimal point
                        row_values.append(f"{val:d}")
                    elif isinstance(val, (float, np.floating)):
                        # Format floats with one decimal place like in the example
                        if val.is_integer():
                            row_values.append(f"{int(val)}")
                        else:
                            row_values.append(f"{val:.1f}")
                    else:
                        # Escape special LaTeX characters in text
                        val_str = str(val)
                        for char, replacement in [
                            ("&", "\\&"),
                            ("%", "\\%"),
                            ("$", "\\$"),
                            ("#", "\\#"),
                            ("_", "\\_"),
                            ("{", "\\{"),
                            ("}", "\\}"),
                            ("~", "\\textasciitilde"),
                            ("^", "\\textasciicircum"),
                            ("\\", "\\textbackslash")
                        ]:
                            val_str = val_str.replace(char, replacement)
                        row_values.append(val_str)
                
                latex.append(f"    {' & '.join(row_values)} \\\\")
                
                # Add midrule between rows like in the example
                if i < len(data) - 1:
                    latex.append("    \\midrule")
            
            # Add space before bottomrule and the bottomrule itself
            latex.append("    \\addlinespace")
            latex.append("    \\bottomrule")
            
            # End tabular environment
            latex.append("  \\end{tabular}")
            
            # Add caption
            if caption:
                latex.append(f"  \\caption{{{caption}}}")
            
            # End table environments
            latex.append("  \\end{center}")
            latex.append("\\end{table}")
            
            return "\n".join(latex)
        
        except Exception as e:
            messagebox.showerror("Conversion Error", f"Failed to convert to LaTeX: {str(e)}")
            return ""
    
    def convert_to_latex(self, data):
        """Convert DataFrame to LaTeX table code"""
        style = self.table_style_var.get()
        
        if style == "booktabs":
            return self.format_booktabs_style(data)
        else:
            # Standard LaTeX table style with vertical lines
            try:
                # Get options from the UI
                caption = self.caption_var.get() if self.include_caption_var.get() else None
                label = self.label_var.get()
                
                # Limit number of rows if specified
                max_rows = int(self.max_rows_var.get()) if self.max_rows_var.get().isdigit() else None
                if max_rows and len(data) > max_rows:
                    data = data.head(max_rows)
                
                # Generate the LaTeX code
                latex = []
                
                # Begin table environment
                latex.append("\\begin{table}[htbp]")
                latex.append("  \\centering")
                
                # Add caption if requested
                if caption:
                    latex.append(f"  \\caption{{{caption}}}")
                
                # Add label
                if label:
                    latex.append(f"  \\label{{{label}}}")
                
                # Calculate number of columns
                num_cols = len(data.columns)
                
                # Create the tabular environment
                alignment = "c" * num_cols
                latex.append(f"  \\begin{{tabular}}{{|{alignment}|}}")
                latex.append("    \\hline")
                
                # Add headers
                headers = " & ".join([str(col) for col in data.columns])
                latex.append(f"    {headers} \\\\")
                latex.append("    \\hline")
                
                # Add data rows
                for _, row in data.iterrows():
                    row_values = []
                    for val in row:
                        if pd.isna(val):
                            # Handle missing values
                            row_values.append("--")
                        elif isinstance(val, (int, np.integer)):
                            # Format integers without decimal point
                            row_values.append(f"{val:d}")
                        elif isinstance(val, (float, np.floating)):
                            # Format floats with reasonable precision
                            if val.is_integer():
                                row_values.append(f"{int(val)}")
                            else:
                                row_values.append(f"{val:.2f}")
                        else:
                            # Escape special LaTeX characters in text
                            val_str = str(val)
                            for char, replacement in [
                                ("&", "\\&"),
                                ("%", "\\%"),
                                ("$", "\\$"),
                                ("#", "\\#"),
                                ("_", "\\_"),
                                ("{", "\\{"),
                                ("}", "\\}"),
                                ("~", "\\textasciitilde"),
                                ("^", "\\textasciicircum"),
                                ("\\", "\\textbackslash")
                            ]:
                                val_str = val_str.replace(char, replacement)
                            row_values.append(val_str)
                    
                    latex.append(f"    {' & '.join(row_values)} \\\\")
                    latex.append("    \\hline")
                
                # End table environments
                latex.append("  \\end{tabular}")
                latex.append("\\end{table}")
                
                return "\n".join(latex)
            
            except Exception as e:
                messagebox.showerror("Conversion Error", f"Failed to convert to LaTeX: {str(e)}")
                return ""
    
    def preview_data(self):
        """Preview the data from the selected file"""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File", "Please select a file first.")
            return
        
        self.data = self.load_file(file_path)
        if self.data is not None:
            # Clear previous content
            self.preview_text.delete(1.0, tk.END)
            
            # Show DataFrame preview
            self.preview_text.insert(tk.END, str(self.data.head(10)))
            
            # Show information about the data
            info_text = f"\n\nData Info:\n"
            info_text += f"Rows: {len(self.data)}\n"
            info_text += f"Columns: {len(self.data.columns)}\n"
            info_text += f"Column Names: {list(self.data.columns)}\n"
            
            self.preview_text.insert(tk.END, info_text)
            
            # Switch to the preview tab
            self.notebook.select(0)
            
            self.status_var.set("Data loaded successfully")
    
    def convert_file(self):
        """Convert the file to LaTeX"""
        if self.data is None:
            file_path = self.file_path_var.get()
            if not file_path:
                messagebox.showwarning("No File", "Please select a file first.")
                return
            
            self.data = self.load_file(file_path)
            if self.data is None:
                return
        
        # Convert to LaTeX
        self.latex_code = self.convert_to_latex(self.data)
        
        # Display the LaTeX code
        self.latex_text.delete(1.0, tk.END)
        self.latex_text.insert(tk.END, self.latex_code)
        
        # Switch to the LaTeX output tab
        self.notebook.select(1)
        
        self.status_var.set("Conversion to LaTeX completed")
    
    def save_latex(self):
        """Save the LaTeX code to a file"""
        if not self.latex_code:
            messagebox.showwarning("No LaTeX", "Please convert a file to LaTeX first.")
            return
        
        # Get the input file name to suggest an output name
        input_file = self.file_path_var.get()
        if input_file:
            suggested_name = os.path.splitext(os.path.basename(input_file))[0] + "_table.tex"
        else:
            suggested_name = "table.tex"
        
        # Open save dialog
        file_path = filedialog.asksaveasfilename(
            defaultextension=".tex",
            filetypes=[("LaTeX Files", "*.tex"), ("All Files", "*.*")],
            initialfile=suggested_name
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(self.latex_code)
                
                self.status_var.set(f"LaTeX saved to: {os.path.basename(file_path)}")
                messagebox.showinfo("Save Successful", f"LaTeX code saved to {file_path}")
            
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save file: {str(e)}")


def main():
    root = tk.Tk()
    app = FileToLatexConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()