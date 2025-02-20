import pandas as pd
from sqlalchemy import create_engine
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import Calendar
import openpyxl
import pyodbc

class DatabaseQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("M_POS APPLICATION")
        self.root.geometry("450x400")

        # UI elements
        self.label = tk.Label(root, text="Select Server Details File")
        self.label.pack(pady=10)

        self.select_file_button = tk.Button(root, text="Browse File", command=self.browse_file)
        self.select_file_button.pack(pady=5)

        # Date selection UI elements
        self.date_label = tk.Label(root, text="Select Date Range")
        self.date_label.pack(pady=10)

        self.start_date_label = tk.Label(root, text="Start Date")
        self.start_date_label.pack(pady=5)
        
        self.start_date = Calendar(root, selectmode='day', date_pattern='yyyy-mm-dd')
        self.start_date.pack(pady=5)
        
        self.end_date_label = tk.Label(root, text="End Date")
        self.end_date_label.pack(pady=5)
        
        self.end_date = Calendar(root, selectmode='day', date_pattern='yyyy-mm-dd')
        self.end_date.pack(pady=5)

        self.run_button = tk.Button(root, text="Run Queries", command=self.run_queries)
        self.run_button.pack(pady=20)

        self.status_label = tk.Label(root, text="Status: Waiting", fg="blue")
        self.status_label.pack(pady=20)

        # Path of the selected Excel file
        self.server_details_file = None

    def browse_file(self):
        """Open file dialog to select an Excel file"""
        self.server_details_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.server_details_file:
            self.status_label.config(text=f"File Selected: {self.server_details_file}", fg="green")
        else:
            self.status_label.config(text="No file selected", fg="red")

    def run_queries(self):
        """Run the SQL queries on each server and generate the output file"""
        if not self.server_details_file:
            messagebox.showerror("Error", "Please select the server details file first.")
            return
        
        # Get selected dates from the calendar
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()

        if not start_date or not end_date:
            messagebox.showerror("Error", "Please select both start and end dates.")
            return

        # Ensure start date is before or equal to end date
        if pd.to_datetime(start_date) > pd.to_datetime(end_date):
            messagebox.showerror("Error", "Start date cannot be later than end date.")
            return
        
        # Debug: Print the selected start and end dates
        print(f"Start Date: {start_date}")
        print(f"End Date: {end_date}")

        try:
            # Read the server details from the Excel file
            server_details_df = pd.read_excel(self.server_details_file)
            combined_data = []

            # Iterate through servers and execute query2
            for index, row in server_details_df.iterrows():
                server_name = row['ServerName']
                connection_string = row['ConnectionString']
                try:
                    # Define your database connection using the extracted connection string
                    engine = create_engine(connection_string)

                    # Query to handle date filtering (CAST transdate to date)
                    query2 = f"""
                   SELECT store, COUNT(*) as counts
                  FROM ax.retailtransactiontable
                  WHERE ismposbill = '1'
                AND CAST(transdate AS DATE) >= '{start_date}' 
                 AND CAST(transdate AS DATE) < '{end_date}'
                 GROUP BY store;
                 """

                    # Debug: Print the final query to check the generated SQL
                    print(f"Query 2: {query2}")

                    # Execute the second query
                    df2 = pd.read_sql(query2, engine)

                    # Debug: Print the number of records fetched from query2
                    print(f"Query2 returned {len(df2)} rows.")

                    # Check if the DataFrame is empty
                    if not df2.empty:
                        combined_data.append(df2)

                except Exception as e:
                    messagebox.showwarning("Warning", f"Error for server {server_name}: {str(e)}")
                    continue  # Continue to the next server

            # Concatenate all the data into a single DataFrame, if there's any data
            if combined_data:
                final_df = pd.concat(combined_data, ignore_index=True)

                # Ask user where to save the final file
                output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
                if output_file:
                    # Use context manager to save the final DataFrame to the sheet
                    with pd.ExcelWriter(output_file, engine='openpyxl') as excel_writer:
                        final_df.to_excel(excel_writer, sheet_name='Combined', index=False)

                    self.status_label.config(text=f"Data exported to {output_file}", fg="green")
                    messagebox.showinfo("Success", "Data exported successfully.")
                else:
                    self.status_label.config(text="File not saved", fg="red")
            else:
                self.status_label.config(text="No data to export.", fg="red")
                messagebox.showinfo("No Data", "No data was retrieved to export.")
    
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the Tkinter window
root = tk.Tk()

# Initialize the application
app = DatabaseQueryApp(root)

# Run the Tkinter event loop
root.mainloop()
