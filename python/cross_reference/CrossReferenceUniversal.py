import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from fuzzywuzzy import fuzz, process
import os
from openpyxl import load_workbook

class FileCrossReferencer:
    def __init__(self, root):
        self.root = root
        self.root.title("File Cross-Reference Tool")
        self.root.geometry("800x600")
        
        # Data storage
        self.source_df = None
        self.target_df = None
        self.results_df = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # File selection section
        ttk.Label(main_frame, text="File Selection", font=('Arial', 12, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # Source file
        ttk.Label(main_frame, text="Source File:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.source_file_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.source_file_var, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_source_file).grid(row=1, column=2, pady=2)
        
        # Target file
        ttk.Label(main_frame, text="Target File:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.target_file_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.target_file_var, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_target_file).grid(row=2, column=2, pady=2)
        
        # Column selection section
        ttk.Label(main_frame, text="Column Selection", font=('Arial', 12, 'bold')).grid(row=3, column=0, columnspan=3, pady=(20, 10))
        
        # Source column
        ttk.Label(main_frame, text="Source Column:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.source_column_var = tk.StringVar()
        self.source_column_combo = ttk.Combobox(main_frame, textvariable=self.source_column_var, state="readonly")
        self.source_column_combo.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        
        # Target column
        ttk.Label(main_frame, text="Target Column:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.target_column_var = tk.StringVar()
        self.target_column_combo = ttk.Combobox(main_frame, textvariable=self.target_column_var, state="readonly")
        self.target_column_combo.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        
        # Matching options section
        ttk.Label(main_frame, text="Matching Options", font=('Arial', 12, 'bold')).grid(row=6, column=0, columnspan=3, pady=(20, 10))
        
        # Match type
        ttk.Label(main_frame, text="Match Type:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.match_type_var = tk.StringVar(value="exact")
        match_frame = ttk.Frame(main_frame)
        match_frame.grid(row=7, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        ttk.Radiobutton(match_frame, text="Exact Match", variable=self.match_type_var, value="exact").pack(side=tk.LEFT)
        ttk.Radiobutton(match_frame, text="Inference Match", variable=self.match_type_var, value="fuzzy").pack(side=tk.LEFT, padx=(20, 0))
        
        # Fuzzy threshold (only for fuzzy matching)
        ttk.Label(main_frame, text="Inference Threshold:").grid(row=8, column=0, sticky=tk.W, pady=2)
        self.threshold_var = tk.IntVar(value=80)
        threshold_frame = ttk.Frame(main_frame)
        threshold_frame.grid(row=8, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)
        ttk.Scale(threshold_frame, from_=50, to=100, variable=self.threshold_var, orient=tk.HORIZONTAL).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.threshold_label = ttk.Label(threshold_frame, text="80%")
        self.threshold_label.pack(side=tk.RIGHT)
        self.threshold_var.trace('w', self.update_threshold_label)
        
        # Process button
        ttk.Button(main_frame, text="Process Files", command=self.process_files).grid(row=9, column=0, columnspan=3, pady=20)
        
        # Results section
        ttk.Label(main_frame, text="Results", font=('Arial', 12, 'bold')).grid(row=10, column=0, columnspan=3, pady=(20, 10))
        
        # Results treeview
        self.results_frame = ttk.Frame(main_frame)
        self.results_frame.grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        
        # Configure main frame row weight for results
        main_frame.rowconfigure(11, weight=1)
        
        self.results_tree = ttk.Treeview(self.results_frame)
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars for results
        v_scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.HORIZONTAL, command=self.results_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.results_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Export button
        self.export_btn = ttk.Button(main_frame, text="Export Results", command=self.export_results, state="disabled")
        self.export_btn.grid(row=12, column=0, columnspan=3, pady=10)
    
    def update_threshold_label(self, *args):
        self.threshold_label.config(text=f"{self.threshold_var.get()}%")
    
    def browse_source_file(self):
        filename = filedialog.askopenfilename(
            title="Select Source File",
            filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filename:
            self.source_file_var.set(filename)
            self.load_source_file()
    
    def browse_target_file(self):
        filename = filedialog.askopenfilename(
            title="Select Target File",
            filetypes=[("Excel & CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if filename:
            self.target_file_var.set(filename)
            self.load_target_file()
    
    # Figuring out what file format the file is in
    def load_file(self, filepath):
        try:
            if filepath.lower().endswith('.csv'):
                return pd.read_csv(filepath)
            elif filepath.lower().endswith(('.xlsx', '.xls')):
                return pd.read_excel(filepath)
            else:
                raise ValueError("Unsupported file format")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            return None
    
    def load_source_file(self):
        filepath = self.source_file_var.get()
        if filepath:
            self.source_df = self.load_file(filepath)
            if self.source_df is not None:
                self.source_column_combo['values'] = list(self.source_df.columns)
                messagebox.showinfo("Success", f"Source file loaded: {len(self.source_df)} rows")
    
    def load_target_file(self):
        filepath = self.target_file_var.get()
        if filepath:
            self.target_df = self.load_file(filepath)
            if self.target_df is not None:
                self.target_column_combo['values'] = list(self.target_df.columns)
                messagebox.showinfo("Success", f"Target file loaded: {len(self.target_df)} rows")
    
    def exact_match(self, source_values, target_values):
        results = []
        no_match_source = []
        match_source = []

        target_set = set(target_values.dropna().astype(str))
        
        for idx, source_val in source_values.items():
            if pd.isna(source_val):
                results.append({
                    'Source_Index': idx,
                    'Source_Value': source_val,
                    'Target_Value': None,
                    'Match_Type': 'No Match (Source NA)',
                    'Match_Score': 0
                })
                no_match_source.append(self.source_df.loc[idx])
            else:
                source_str = str(source_val)
                if source_str in target_set:
                    results.append({
                        'Source_Index': idx,
                        'Source_Value': source_val,
                        'Target_Value': source_str,
                        'Match_Type': 'Exact Match',
                        'Match_Score': 100
                    })
                    match_source.append(self.source_df.loc[idx])
                else:
                    results.append({
                        'Source_Index': idx,
                        'Source_Value': source_val,
                        'Target_Value': None,
                        'Match_Type': 'No Match',
                        'Match_Score': 0
                    })
                    no_match_source.append(self.source_df.loc[idx])
        
        return pd.DataFrame(results), pd.DataFrame(no_match_source), pd.DataFrame(match_source)
    
    def fuzzy_match(self, source_values, target_values, threshold):
        results = []
        no_match_source = []
        match_source = []
        target_list = target_values.dropna().astype(str).tolist()
        
        for idx, source_val in source_values.items():
            if pd.isna(source_val):
                results.append({
                    'Source_Index': idx,
                    'Source_Value': source_val,
                    'Target_Value': None,
                    'Match_Type': 'No Match (Source NA)',
                    'Match_Score': 0
                })
                no_match_source.append(self.source_df.loc[idx])
            else:
                source_str = str(source_val)
                
                # Find best match
                best_match = process.extractOne(source_str, target_list, scorer=fuzz.ratio)
                
                if best_match and best_match[1] >= threshold:
                    match_type = 'Exact Match' if best_match[1] == 100 else 'Fuzzy Match'
                    results.append({
                        'Source_Index': idx,
                        'Source_Value': source_val,
                        'Target_Value': best_match[0],
                        'Match_Type': match_type,
                        'Match_Score': best_match[1]
                    })
                    match_source.append(self.source_df.loc[idx])
                else:
                    results.append({
                        'Source_Index': idx,
                        'Source_Value': source_val,
                        'Target_Value': best_match[0] if best_match else None,
                        'Match_Type': 'No Match',
                        'Match_Score': best_match[1] if best_match else 0
                    })
                    no_match_source.append(self.source_df.loc[idx])
        
        return pd.DataFrame(results), pd.DataFrame(no_match_source), pd.DataFrame(match_source)
    
    def process_files(self):
        # Validate inputs
        if self.source_df is None or self.target_df is None:
            messagebox.showerror("Error", "Please load both source and target files")
            return
        
        source_col = self.source_column_var.get()
        target_col = self.target_column_var.get()
        
        if not source_col or not target_col:
            messagebox.showerror("Error", "Please select both source and target columns")
            return
        
        try:
            # Get column data
            source_values = self.source_df[source_col]
            target_values = self.target_df[target_col]
            
            # Perform matching
            match_type = self.match_type_var.get()
            
            if match_type == "exact":
                self.results_df, self.no_match_source_df, self.match_source_df = self.exact_match(source_values, target_values)
                threshold = 100
            else:
                threshold = self.threshold_var.get()
                self.results_df, self.no_match_source_df, self.match_source_df= self.fuzzy_match(source_values, target_values, threshold)
            
            # Display results
            self.display_results()
            
            # Show summary
            total_rows = len(self.results_df)
            matches = len(self.results_df[self.results_df['Match_Score'] >= threshold])
            match_rate = (matches / total_rows) * 100 if total_rows > 0 else 0
            
            messagebox.showinfo("Processing Complete", 
                              f"Processing complete!\n"
                              f"Total records: {total_rows}\n"
                              f"Matches found: {matches}\n"
                              f"Match rate: {match_rate:.1f}%")
            
            self.export_btn.config(state="normal")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing files: {str(e)}")
    
    def display_results(self):
        # Clear existing results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if self.results_df is None or self.results_df.empty:
            return
        
        # Configure columns
        columns = list(self.results_df.columns)
        self.results_tree['columns'] = columns
        self.results_tree['show'] = 'headings'
        
        # Configure column headings and widths
        for col in columns:
            self.results_tree.heading(col, text=col)
            self.results_tree.column(col, width=150, minwidth=100)
        
        # Insert data
        for idx, row in self.results_df.iterrows():
            values = [str(val) if not pd.isna(val) else '' for val in row]
            item = self.results_tree.insert('', 'end', values=values)
            
            # Color code based on match type
            match_type = row['Match_Type']
            if match_type == 'Exact Match':
                self.results_tree.set(item, 'Match_Type', match_type)
                self.results_tree.item(item, tags=('exact',))
            elif match_type == 'Fuzzy Match':
                self.results_tree.item(item, tags=('fuzzy',))
            else:
                self.results_tree.item(item, tags=('no_match',))
        
        # Configure tags for coloring
        self.results_tree.tag_configure('exact', background='lightgreen')
        self.results_tree.tag_configure('fuzzy', background='lightyellow')
        self.results_tree.tag_configure('no_match', background='lightcoral')
    
    def export_results(self):
        if self.results_df is None or self.results_df.empty:
            messagebox.showerror("Error", "No results to export")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        print(filename)
        if filename:
            try:
                if filename.lower().endswith('.csv'):
                    self.results_df.to_csv(filename, index=False)
                else:
                    with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
                        self.results_df.to_excel(writer, sheet_name="Results", index=False)
                        self.match_source_df.to_excel(writer, sheet_name="Matched Rows (Source)")
                        self.no_match_source_df.to_excel(writer, sheet_name="No Matched Rows (Source)")
                messagebox.showinfo("Success", f"Results exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Error exporting results: {str(e)}")

def main():
    root = tk.Tk()
    app = FileCrossReferencer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
