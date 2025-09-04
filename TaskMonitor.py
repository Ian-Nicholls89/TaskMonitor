import time
print("Starting task checker script...")
t0 = time.time()

from pandas import read_csv, read_excel, ExcelFile, isna, notna, to_datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime, date
from os import path, getenv
t1 = time.time()
print(f"Preload complete in {round(t1-t0, 2)}s")

class TaskChecker:
    def __init__(self):
        # Hide the main tkinter window
        self.root = tk.Tk()
        self.root.withdraw()
        self.env_file = '.env'
        
    def load_env_config(self):
        """Load configuration from .env file"""
        config = {}
        
        if not path.exists(self.env_file):
            return config
            
        try:
            with open(self.env_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        config[key.strip()] = value.strip().strip('"\'')
            return config
        except Exception as e:
            print(f"Error reading .env file: {e}")
            return config
    
    def save_env_config(self, file_path, worksheet, date_column, description_column):
        upworksheet = worksheet.upper()
        upworksheet = upworksheet.replace(" ", "_")
        """Save configuration to .env file"""
        try:
            with open(self.env_file, 'a') as f:
                if getenv("SPREADSHEET_PATH") is None:
                    f.write("# Task Checker Configuration\n")
                    f.write(f"SPREADSHEET_PATH={file_path}\n")
                f.write(f"{upworksheet}_DATE_COLUMN={date_column}\n")
                f.write(f"{upworksheet}_DESCRIPTION_COLUMN={description_column}\n")
            print(f"Configuration saved to {self.env_file}")
        except Exception as e:
            print(f"Error saving .env file: {e}")
    
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
                df = read_csv(file_path)
                return df, None  # CSV files don't have worksheets
            else:
                # Load Excel file
                excel_file = ExcelFile(file_path)
                worksheet_names = excel_file.sheet_names
                
                if worksheet_name and worksheet_name in worksheet_names:
                    # Load specific worksheet
                    df = read_excel(file_path, sheet_name=worksheet_name)
                    return df, worksheet_names
                elif len(worksheet_names) == 1:
                    # Only one worksheet, load it automatically
                    df = read_excel(file_path, sheet_name=worksheet_names[0])
                    return df, worksheet_names
                else:
                    # Multiple worksheets, need to select
                    return None, worksheet_names
                    
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file: {str(e)}")
            return None, None
    
    def parse_date(self, date_value):
        """Parse various date formats"""
        if isna(date_value):
            return None
            
        # If it's already a datetime object
        if isinstance(date_value, (datetime, date)):
            return date_value.date() if hasattr(date_value, 'date') else date_value
        
        # Try to parse string dates
        try:
            # Common date formats
            date_formats = [
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
                    
            # If pandas can parse it
            parsed_date = to_datetime(date_str).date()
            return parsed_date
            
        except:
            return None
    
    def check_due_tasks(self, df, date_column, description_column):
        """Check for tasks that are due today or overdue"""
        today = date.today()
        due_tasks = []
        
        for index, row in df.iterrows():
            task_date = self.parse_date(row[date_column])
            
            if task_date and task_date <= today:
                description = str(row[description_column]) if notna(row[description_column]) else f"Task in row {index + 1}"
                due_tasks.append({
                    'description': description,
                    'due_date': task_date,
                    'days_overdue': (today - task_date).days
                })
            elif not task_date:
                description = str(row[description_column]) if notna(row[description_column]) else f"Task in row {index + 1}"
                due_tasks.append({
                    'description': description,
                    'due_date': task_date,
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
                status = "Task completion date not entered."
            else:
                status = f"OVERDUE by {task['days_overdue']} day(s)"
            
            message += f"• {task['description']} ({task['worksheet']})\n" if 'worksheet' in task else f"• {task['description']}\n"
            message += f"  Due: {task['due_date']} ({status})\n\n"
        
        # Show popup
        messagebox.showwarning("Tasks Due!", message)
    
    def get_column_selection(self, df, saved_config=None, worksheet_name=None):
        """Simple dialog to select columns with saved defaults"""
        columns = list(df.columns)
        worksheet_name = worksheet_name.replace(" ","_").upper()

        # Use saved values as defaults if available
        default_date = saved_config.get(f'{worksheet_name}_DATE_COLUMN', columns[0]) if saved_config else columns[0]
        default_desc = saved_config.get(f'{worksheet_name}_DESCRIPTION_COLUMN', columns[1] if len(columns) > 1 else columns[0]) if saved_config else (columns[1] if len(columns) > 1 else columns[0])
        
        # Validate that saved columns exist in the spreadsheet
        if default_date not in columns:
            default_date = columns[0]
        if default_desc not in columns:
            default_desc = columns[1] if len(columns) > 1 else columns[0]
        
        if saved_config is None:
            # Create a simple selection window
            selection_window = tk.Toplevel()
            title = f"Select Columns - {worksheet_name}" if worksheet_name else "Select Columns"
            selection_window.title(title)
            selection_window.geometry("500x450")
            selection_window.resizable(False, False)
            selection_window.grab_set()  # Make it modal
            
            # Center the window
            selection_window.update_idletasks()
            x = (selection_window.winfo_screenwidth() // 2) - (500 // 2)
            y = (selection_window.winfo_screenheight() // 2) - (450 // 2)
            selection_window.geometry(f"500x450+{x}+{y}")
            
            if worksheet_name:
                tk.Label(selection_window, text=f"Configuring worksheet: {worksheet_name}", 
                        font=('Arial', 12, 'bold'), fg='darkblue').pack(pady=5)
            
            if saved_config:
                tk.Label(selection_window, text="Using saved configuration. Verify columns:", 
                        font=('Arial', 10, 'italic'), fg='blue').pack(pady=5)
            
            # Show available columns at the top
            tk.Label(selection_window, text=f"Available columns: {', '.join(columns)}", 
                    font=('Arial', 9), wraplength=450, justify='left').pack(pady=5)
            
            tk.Label(selection_window, text="Select the DUE BY column:", font=('Arial', 12, 'bold')).pack(pady=10)
            
            date_var = tk.StringVar(value=default_date)
            date_frame = tk.Frame(selection_window)
            date_frame.pack(pady=5)
            
            for col in columns:
                tk.Radiobutton(date_frame, text=col, variable=date_var, value=col).pack(anchor='w')
            
            tk.Label(selection_window, text="Select the TASK DESCRIPTION column:", font=('Arial', 12, 'bold')).pack(pady=(20,10))
            
            desc_var = tk.StringVar(value=default_desc)
            desc_frame = tk.Frame(selection_window)
            desc_frame.pack(pady=5)
            
            for col in columns:
                tk.Radiobutton(desc_frame, text=col, variable=desc_var, value=col).pack(anchor='w')
            
            # Add checkbox to save configuration
            save_config_var = tk.BooleanVar(value=True)
            tk.Checkbutton(selection_window, text="Save these settings for next time", 
                        variable=save_config_var, font=('Arial', 10)).pack(pady=10)
            
            result = {'date_col': None, 'desc_col': None, 'confirmed': False, 'save_config': False}
            
            def confirm():
                result['date_col'] = date_var.get()
                result['desc_col'] = desc_var.get()
                result['confirmed'] = True
                result['save_config'] = save_config_var.get()
                selection_window.destroy()
            
            def cancel():
                selection_window.destroy()
            
            # Add separator line
            separator = tk.Frame(selection_window, height=2, bg='gray')
            separator.pack(fill='x', padx=20, pady=(15, 10))
            
            button_frame = tk.Frame(selection_window)
            button_frame.pack(pady=15)
            
            # Make buttons more prominent
            confirm_btn = tk.Button(button_frame, text="✓ Confirm Selection", command=confirm, 
                                bg='lightgreen', font=('Arial', 11, 'bold'), 
                                padx=20, pady=8)
            confirm_btn.pack(side='left', padx=15)
            
            cancel_btn = tk.Button(button_frame, text="✗ Cancel", command=cancel, 
                                bg='lightcoral', font=('Arial', 11, 'bold'), 
                                padx=20, pady=8)
            cancel_btn.pack(side='left', padx=15)
            
            selection_window.wait_window()
            return result
        
        result = {'date_col': default_date, 'desc_col': default_desc, 'confirmed': True, 'save_config': False} # return saved values, no need to save again
        return result
    
    def run(self):
        """Main execution function"""
        # Get file path (from .env or user selection)
        print("Getting spreadsheet location from settings file.")
        file_path, saved_config = self.get_file_path()
        if not file_path:
            return

        # Load spreadsheet and get worksheet info
        print("Reading...\n")
        df, worksheet_names = self.load_spreadsheet(file_path, saved_config.get('WORKSHEET_NAME'))
        
        # Handle different file types
        if df is None and worksheet_names:
            # Multiple worksheets - check all of them
            all_due_tasks = []
            
            for worksheet in worksheet_names:  
                try:
                    df = read_excel(file_path, sheet_name=worksheet)
                    print(f"Checking worksheet: {worksheet}\n")
                    
                    # Get column selection for this worksheet
                    selection = self.get_column_selection(df, saved_config, worksheet)
                    if not selection['confirmed']:
                        continue  # Skip this worksheet if user cancels
                    
                    date_column = selection['date_col']
                    description_column = selection['desc_col']
                    
                    # Save configuration if requested (only for the first worksheet or if different)
                    if selection['save_config']:
                        self.save_env_config(file_path, worksheet, date_column, description_column)
                    
                    # Check for due tasks in this worksheet
                    due_tasks = self.check_due_tasks(df, date_column, description_column)
                    
                    # Add worksheet name to each task for identification
                    for task in due_tasks:
                        task['worksheet'] = worksheet
                    
                    all_due_tasks.extend(due_tasks)
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Could not load worksheet '{worksheet}': {str(e)}")
                    continue

            t2 = time.time()
            print(f"Completed in {round(t2-t0, 2)}s!\n\n")
            print("To change any settings either edit (in notepad) or delete the .env file in the root directory of this script and run again.")
            # Show combined results from all worksheets
            self.show_popup(all_due_tasks)
            
        elif df is not None:
            # Single worksheet or CSV file
            selected_worksheet = worksheet_names[0] if worksheet_names else "Main"
            
            # Select columns (using saved config if available)
            selection = self.get_column_selection(df, saved_config)
            if not selection['confirmed']:
                return
            
            date_column = selection['date_col']
            description_column = selection['desc_col']
            
            # Save configuration if requested
            if selection['save_config']:
                self.save_env_config(file_path, selected_worksheet, date_column, description_column)
            
            # Check for due tasks
            due_tasks = self.check_due_tasks(df, date_column, description_column)
            
            # Add worksheet name for consistency
            for task in due_tasks:
                task['worksheet'] = selected_worksheet
            
            t2 = time.time()
            print(f"Completed in {round(t2-t0, 2)}s!\n\n")
            print("To change any settings either edit (in notepad) or delete the .env file in the root directory of this script and run again.")
            # Show popup
            self.show_popup(due_tasks)
        else:
            return  # Error loading file

def main():
    """Main function with .env file support"""
    checker = TaskChecker()
    checker.run()

if __name__ == "__main__":
    main()