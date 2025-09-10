print("Task Monitor Script. Made by Ian Nicholls, 2025.")
import sys, time
if '--startup' in sys.argv:
    print("Run at startup detected; pausing script for 90 seconds to conserve system resources. (Don't close this window!)")
    time.sleep(90)

print("Starting task checker script...")

t0 = time.time()

import openpyxl
import csv
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime, date
from os import path
from dotenv import load_dotenv, set_key
import os

t1 = time.time()
print(f"Python module load complete in {round(t1-t0, 2)}s")

class ColumnSelectionWindow(tk.Toplevel):
    def __init__(self, parent, columns, worksheet_name, saved_config, default_date, default_desc):
        super().__init__(parent)
        self.title(f"Select Columns - {worksheet_name}" if worksheet_name else "Select Columns")
        self.geometry("500x450")
        self.resizable(False, False)
        self.grab_set()

        self.result = {'date_col': None, 'desc_col': None, 'confirmed': False, 'save_config': False}

        self.create_widgets(columns, worksheet_name, saved_config, default_date, default_desc)
        self.center_window()

    def create_widgets(self, columns, worksheet_name, saved_config, default_date, default_desc):
        if worksheet_name:
            tk.Label(self, text=f"Configuring worksheet: {worksheet_name}",
                     font=('Arial', 12, 'bold'), fg='darkblue').pack(pady=5)
        
        if saved_config:
            tk.Label(self, text="Using saved configuration. Verify columns:",
                     font=('Arial', 10, 'italic'), fg='blue').pack(pady=5)

        tk.Label(self, text=f"Available columns: {', '.join(columns)}",
                 font=('Arial', 9), wraplength=450, justify='left').pack(pady=5)

        tk.Label(self, text="Select the DUE BY column:", font=('Arial', 12, 'bold')).pack(pady=10)
        
        date_var = tk.StringVar(value=default_date)
        date_frame = tk.Frame(self)
        date_frame.pack(pady=5)
        
        for col in columns:
            tk.Radiobutton(date_frame, text=col, variable=date_var, value=col).pack(anchor='w')
        
        tk.Label(self, text="Select the TASK DESCRIPTION column:", font=('Arial', 12, 'bold')).pack(pady=(20,10))
        
        desc_var = tk.StringVar(value=default_desc)
        desc_frame = tk.Frame(self)
        desc_frame.pack(pady=5)
        
        for col in columns:
            tk.Radiobutton(desc_frame, text=col, variable=desc_var, value=col).pack(anchor='w')
            
        save_config_var = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="Save these settings for next time",
                    variable=save_config_var, font=('Arial', 10)).pack(pady=10)

        separator = tk.Frame(self, height=2, bg='gray')
        separator.pack(fill='x', padx=20, pady=(15, 10))
        
        button_frame = tk.Frame(self)
        button_frame.pack(pady=15)

        confirm_btn = tk.Button(button_frame, text="✓ Confirm Selection", command=lambda: self.confirm(date_var, desc_var, save_config_var),
                                bg='lightgreen', font=('Arial', 11, 'bold'),
                                padx=20, pady=8)
        confirm_btn.pack(side='left', padx=15)
        
        cancel_btn = tk.Button(button_frame, text="✗ Cancel", command=self.cancel,
                                bg='lightcoral', font=('Arial', 11, 'bold'),
                                padx=20, pady=8)
        cancel_btn.pack(side='left', padx=15)

    def confirm(self, date_var, desc_var, save_config_var):
        self.result['date_col'] = date_var.get()
        self.result['desc_col'] = desc_var.get()
        self.result['confirmed'] = True
        self.result['save_config'] = save_config_var.get()
        self.destroy()

    def cancel(self):
        self.destroy()

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

class TaskChecker:
    def __init__(self):
        # Hide the main tkinter window
        self.root = tk.Tk()
        self.root.withdraw()
        
        # Use appdata directory for storing .env file
        appdata_dir = os.path.join(os.getenv('APPDATA'), 'TaskMonitor')
        os.makedirs(appdata_dir, exist_ok=True)
        self.env_file = os.path.join(appdata_dir, 'config.env')
        
    def load_env_config(self):
        """Load configuration from .env file"""
        if not path.exists(self.env_file):
            # Create the file if it doesn't exist
            with open(self.env_file, 'w') as f:
                pass
        load_dotenv(dotenv_path=self.env_file)
        return os.environ
            
    def save_env_config(self, file_path, worksheet, date_column, description_column, expected_sheets):
        upworksheet = worksheet.upper().replace(" ", "_")
        
        """Save configuration to .env file"""
        set_key(self.env_file, "SPREADSHEET_PATH", file_path)
        set_key(self.env_file, "EXPECTED_SHEETS", str(expected_sheets))
        set_key(self.env_file, f"{upworksheet}_DATE_COLUMN", date_column)
        set_key(self.env_file, f"{upworksheet}_DESCRIPTION_COLUMN", description_column)
        print(f"{worksheet} configuration saved to {self.env_file}")
    
    def get_file_path(self):
        """Get file path from .env or ask user to select"""
        config = self.load_env_config()
        file_path = config.get('SPREADSHEET_PATH', '')
        
        # Check if .env file exists and has a valid file path
        if file_path and path.exists(file_path):
            return file_path, config
        
        # If no .env file, empty file, or user chose to select new file
        print("No file location in settings. Please select a file.")
        file_path = self.select_file()
        return file_path, {}
    
    def select_file(self):
        """Allow user to select the spreadsheet file"""
        file_path = filedialog.askopenfilename(
            title="Select your spreadsheet",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        return file_path
    
    def load_spreadsheet(self, file_path, worksheet_name=None):
        """Load the spreadsheet based on file extension"""
        try:
            if file_path.endswith('.csv'):
                with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    data = list(reader)
                return data, None  # CSV files don't have worksheets
            else:
                # Load Excel file
                workbook = openpyxl.load_workbook(file_path, data_only=True)
                worksheet_names = workbook.sheetnames
                
                if worksheet_name and worksheet_name in worksheet_names:
                    # Load specific worksheet
                    sheet = workbook[worksheet_name]
                    data = list(sheet.values)
                    return data, worksheet_names
                elif len(worksheet_names) == 1:
                    # Only one worksheet, load it automatically
                    sheet = workbook.active
                    data = list(sheet.values)
                    return data, worksheet_names
                else:
                    # Multiple worksheets, need to select
                    return None, worksheet_names
                    
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {str(e)}")
            return None, None
    
    def parse_date(self, date_value):
        """Parse various date formats"""
        if date_value is None:
            return None
            
        # If it's already a datetime object
        if isinstance(date_value, (datetime, date)):
            return date_value.date() if hasattr(date_value, 'date') else date_value
        
        # Try to parse string dates
        try:
            # Common date formats
            date_formats = [
                '%Y-%m-%d %H:%M:%S', # For datetime objects from openpyxl
                '%Y-%m-%d',
                '%m/%d/%Y',
                '%d/%m/%Y',
                '%m-%d-%Y',
                '%d-%m-%Y',
                '%Y/%m/%d',
                '%B %d, %Y',
                '%b %d, %Y',
                '%d %B %Y',
                '%d %b %Y'
            ]
            
            date_str = str(date_value).strip()
            
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt).date()
                    return parsed_date
                except ValueError:
                    continue
            return None
        except:
            return None
    
    def check_due_tasks(self, data, date_column, description_column):
        """Check for tasks that are due today or overdue"""
        today = date.today()
        due_tasks = []
        
        header = data[0]
        try:
            date_col_index = header.index(date_column)
            desc_col_index = header.index(description_column)
        except ValueError as e:
            messagebox.showerror("Column Not Found", f"A specified column was not found in the header: {e}")
            return []

        for i, row in enumerate(data[1:]): # Skip header row
            if not any(row): continue # Skip empty rows

            task_date = self.parse_date(row[date_col_index])
            
            description = str(row[desc_col_index]) if row[desc_col_index] is not None else f"Task in row {i + 2}"

            if task_date and task_date <= today:
                due_tasks.append({
                    'description': description,
                    'due_date': task_date,
                    'days_overdue': (today - task_date).days
                })
            elif not task_date:
                due_tasks.append({
                    'description': description,
                    'due_date': "Invalid Date",
                    'days_overdue': -1
                })
        
        return due_tasks
    
    def show_popup(self, due_tasks):
        """Display popup with due tasks"""
        if not due_tasks:
            messagebox.showinfo("Task Checker", "No tasks are due today or overdue!")
            return
        
        # Create message
        message = "The following tasks are due:\n\n"
        
        for task in due_tasks:
            if task['days_overdue'] == 0:
                status = "DUE TODAY"
            elif task['days_overdue'] == -1:
                status = "Task completion date not entered or invalid."
            else:
                status = f"OVERDUE by {task['days_overdue']} day(s)"
            
            message += f"• {task['description']} ({task['worksheet']})\n" if 'worksheet' in task else f"• {task['description']}\n"
            message += f"  Due: {task['due_date']} ({status})\n\n"
        
        # Show popup
        messagebox.showwarning("Tasks Due!", message)
    
    def get_column_selection(self, data, saved_config=None, worksheet_name=None, sheetnum=1):
        """Handles the logic for column selection, using the UI window if needed."""
        columns = list(data[0]) if data else []
        worksheet_name_env = worksheet_name.replace(" ", "_").upper()

        default_date = saved_config.get(f'{worksheet_name_env}_DATE_COLUMN', columns[2] if len(columns) > 1 else (columns[0] if columns else ''))
        default_desc = saved_config.get(f'{worksheet_name_env}_DESCRIPTION_COLUMN', columns[0] if columns else '')

        if default_date not in columns:
            default_date = columns[2] if len(columns) > 1 else (columns[0] if columns else '')
        if default_desc not in columns:
            default_desc = columns[0] if columns else ''

        expected_sheets = int(saved_config.get("EXPECTED_SHEETS", 0))
        use_saved = (expected_sheets == sheetnum) and all([default_date, default_desc])

        if not use_saved:
            selection_window = ColumnSelectionWindow(self.root, columns, worksheet_name, saved_config, default_date, default_desc)
            self.root.wait_window(selection_window)
            return selection_window.result
        
        return {'date_col': default_date, 'desc_col': default_desc, 'confirmed': True, 'save_config': False}
    
    def run(self):
        """Main execution function"""
        # Get file path (from .env or user selection)
        print("Getting spreadsheet location from settings file.")
        file_path, saved_config = self.get_file_path()
        if not file_path:
            print("Read spreadsheet location from settings file failed. Please try again.")
            return

        # Load spreadsheet and get worksheet info
        print("Reading file content...\n")
        data, worksheet_names = self.load_spreadsheet(file_path)
        
        # Handle different file types
        if data is None and worksheet_names:
            # Multiple worksheets - check all of them
            all_due_tasks = []
            
            for worksheet in worksheet_names:  
                try:
                    # With openpyxl, we need to reload the data for each sheet
                    sheet_data, _ = self.load_spreadsheet(file_path, worksheet_name=worksheet)
                    if not sheet_data: continue

                    print(f"Checking worksheet: {worksheet}\n")
                    
                    # Get column selection for this worksheet
                    selection = self.get_column_selection(sheet_data, saved_config, worksheet, len(worksheet_names))
                    if not selection['confirmed']:
                        continue  # Skip this worksheet if user cancels
                    
                    date_column = selection['date_col']
                    description_column = selection['desc_col']
                    
                    # Save configuration if requested
                    if selection['save_config']:
                        self.save_env_config(file_path, worksheet, date_column, description_column, len(worksheet_names))
                    
                    # Check for due tasks in this worksheet
                    due_tasks = self.check_due_tasks(sheet_data, date_column, description_column)
                    
                    # Add worksheet name to each task for identification
                    for task in due_tasks:
                        task['worksheet'] = worksheet
                    
                    all_due_tasks.extend(due_tasks)
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not load worksheet '{worksheet}': {str(e)}")
                    continue

            t2 = time.time()
            print(f"Completed in {round(t2-t0, 2)}s!\n\n")
            # Show combined results from all worksheets
            self.show_popup(all_due_tasks)
            
        elif data is not None:
            # Single worksheet or CSV file
            selected_worksheet = worksheet_names[0] if worksheet_names else "Main"
            
            # Select columns (using saved config if available)
            selection = self.get_column_selection(data, saved_config, selected_worksheet, len(worksheet_names) if worksheet_names else 1)
            if not selection['confirmed']:
                return
            
            date_column = selection['date_col']
            description_column = selection['desc_col']
            
            # Save configuration if requested
            if selection['save_config']:
                self.save_env_config(file_path, selected_worksheet, date_column, description_column, len(worksheet_names) if worksheet_names else 1)
            
            # Check for due tasks
            due_tasks = self.check_due_tasks(data, date_column, description_column)
            
            # Add worksheet name for consistency
            for task in due_tasks:
                task['worksheet'] = selected_worksheet
            
            t2 = time.time()
            print(f"Completed in {round(t2-t0, 2)}s!\n\n")
            # Show popup
            self.show_popup(due_tasks)
        else:
            return  # Error loading file

def main():
    """Main function with .env file support"""
    checker = TaskChecker()
    checker.run()

if __name__ == "__main__":
    if sys.argv[1] == "--editsettings":
        checker = TaskChecker()
        file_path, _ = checker.get_file_path()
        _, worksheets = checker.load_spreadsheet(file_path)
        for worksheet in worksheets:
            data, _ = checker.load_spreadsheet(file_path, worksheet_name=worksheet)
            checker.get_column_selection(data, {}, worksheet, len(worksheets))
        sys.exit(0)
    else:
        main()
