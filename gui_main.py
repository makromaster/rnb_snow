import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import subprocess
import os
import sys
import sqlite3
import csv
from create_database import update_database, get_database_stats


class TicketMatchingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("RnB Snow Ticket Matching System")
        self.root.geometry("800x700")

        # Variables
        self.sap_file = tk.StringVar()
        self.snow_file = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")

        self.setup_ui()
        self.refresh_stats()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(main_frame, text="RnB Snow Ticket Matching System",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # SAP file selection
        ttk.Label(file_frame, text="SAP Data File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.sap_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse",
                  command=self.browse_sap_file).grid(row=0, column=2, padx=5, pady=5)

        # Snow file selection
        ttk.Label(file_frame, text="ServiceNow File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.snow_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse",
                  command=self.browse_snow_file).grid(row=1, column=2, padx=5, pady=5)

        # Process button
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=2, column=0, columnspan=3, pady=10)

        self.process_btn = ttk.Button(process_frame, text="Process Files",
                                     command=self.process_files, style="Accent.TButton")
        self.process_btn.pack(side=tk.LEFT, padx=5)

        self.selenium_btn = ttk.Button(process_frame, text="Launch Selenium Session",
                                      command=self.launch_selenium)
        self.selenium_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(process_frame, text="Export to CSV",
                                    command=self.export_to_csv)
        self.export_btn.pack(side=tk.LEFT, padx=5)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Status
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT)
        self.status_label = ttk.Label(status_frame, textvariable=self.status_text)
        self.status_label.pack(side=tk.LEFT, padx=5)

        # Statistics section
        stats_frame = ttk.LabelFrame(main_frame, text="Database Statistics", padding="10")
        stats_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Stats grid
        self.stats_labels = {}
        stats_layout = [
            ("Total Tickets:", "total_tickets", 0, 0),
            ("Matched Tickets:", "matched_tickets", 0, 2),
            ("Match Percentage:", "match_percentage", 1, 0),
            ("SAP Records:", "sap_records", 1, 2),
            ("Extracted Text:", "extracted_count", 2, 0),
            ("Nothing to Extract:", "nothing_to_extract_count", 2, 2),
            ("Pending Extraction:", "pending_extraction_count", 3, 0)
        ]

        for label_text, key, row, col in stats_layout:
            ttk.Label(stats_frame, text=label_text).grid(row=row, column=col, sticky=tk.W, padx=5, pady=2)
            self.stats_labels[key] = ttk.Label(stats_frame, text="0", font=("Arial", 10, "bold"))
            self.stats_labels[key].grid(row=row, column=col+1, sticky=tk.W, padx=5, pady=2)

        # Refresh button
        ttk.Button(stats_frame, text="Refresh Stats",
                  command=self.refresh_stats).grid(row=4, column=0, columnspan=4, pady=10)

        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # Text widget with scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(text_frame, height=15, width=80)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Clear log button
        ttk.Button(log_frame, text="Clear Log",
                  command=self.clear_log).pack(pady=5)

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

    def browse_sap_file(self):
        filename = filedialog.askopenfilename(
            title="Select SAP Data File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.sap_file.set(filename)

    def browse_snow_file(self):
        filename = filedialog.askopenfilename(
            title="Select ServiceNow Data File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.snow_file.set(filename)

    def log_message(self, message):
        """Add message to log with timestamp"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def update_status(self, message):
        """Update status and log message"""
        self.status_text.set(message)
        self.log_message(message)

    def process_files(self):
        """Process the selected files"""
        sap_file = self.sap_file.get().strip()
        snow_file = self.snow_file.get().strip()

        if not sap_file or not snow_file:
            messagebox.showerror("Error", "Please select both SAP and ServiceNow files")
            return

        if not os.path.exists(sap_file):
            messagebox.showerror("Error", f"SAP file not found: {sap_file}")
            return

        if not os.path.exists(snow_file):
            messagebox.showerror("Error", f"ServiceNow file not found: {snow_file}")
            return

        # Disable button and start progress
        self.process_btn.config(state="disabled")
        self.progress.start()

        # Run processing in background thread
        thread = threading.Thread(target=self._process_files_thread, args=(sap_file, snow_file))
        thread.daemon = True
        thread.start()

    def _process_files_thread(self, sap_file, snow_file):
        """Background thread for file processing"""
        try:
            def progress_callback(message):
                self.root.after(0, self.update_status, message)

            self.root.after(0, self.update_status, "Starting file processing...")

            # Process files
            matched = update_database(
                sap_file=sap_file,
                snow_file=snow_file,
                progress_callback=progress_callback
            )

            self.root.after(0, self.update_status, f"Processing complete! {matched} matches found")
            self.root.after(0, self.refresh_stats)

        except Exception as e:
            error_msg = f"Error processing files: {str(e)}"
            self.root.after(0, self.update_status, error_msg)
            self.root.after(0, messagebox.showerror, "Error", error_msg)

        finally:
            # Re-enable button and stop progress
            self.root.after(0, self._processing_complete)

    def _processing_complete(self):
        """Called when processing is complete"""
        self.process_btn.config(state="normal")
        self.progress.stop()

    def refresh_stats(self):
        """Refresh database statistics"""
        try:
            stats = get_database_stats()

            self.stats_labels['total_tickets'].config(text=str(stats['total_tickets']))
            self.stats_labels['matched_tickets'].config(text=str(stats['matched_tickets']))
            self.stats_labels['match_percentage'].config(text=f"{stats['match_percentage']:.1f}%")
            self.stats_labels['sap_records'].config(text=str(stats['sap_records']))
            self.stats_labels['extracted_count'].config(text=str(stats['extracted_count']))
            self.stats_labels['nothing_to_extract_count'].config(text=str(stats['nothing_to_extract_count']))
            self.stats_labels['pending_extraction_count'].config(text=str(stats['pending_extraction_count']))

            self.log_message("Statistics refreshed")

        except Exception as e:
            self.log_message(f"Error refreshing stats: {e}")

    def launch_selenium(self):
        """Launch selenium debug session"""
        try:
            # Check if there are tickets to process
            stats = get_database_stats()
            if stats['pending_extraction_count'] == 0:
                messagebox.showinfo("Info", "No tickets pending extraction. All tickets have been processed or marked as 'nothing to extract'.")
                return

            self.log_message(f"Launching Selenium session for {stats['pending_extraction_count']} pending tickets...")

            # Determine the selenium executable path
            current_dir = os.path.dirname(os.path.abspath(__file__))

            # Check if we're running as executable or from source
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller executable
                selenium_exe = os.path.join(current_dir, 'RnB-Snow-Selenium.exe')
                if os.path.exists(selenium_exe):
                    cmd = [selenium_exe]
                else:
                    # Fallback: try to import and run directly
                    self._run_selenium_direct()
                    return
            else:
                # Running from source - use Python
                cmd = ['python', 'selenium_debug_session.py']

            # Launch selenium process
            process = subprocess.Popen(cmd, cwd=current_dir)

            self.log_message("Selenium session launched in separate process")
            messagebox.showinfo("Selenium Launched",
                               f"Selenium session started for {stats['pending_extraction_count']} tickets.\n"
                               "Check the console window for progress.")

        except Exception as e:
            error_msg = f"Error launching Selenium: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)

    def _run_selenium_direct(self):
        """Run selenium session directly in current process as fallback"""
        try:
            import selenium_debug_session
            self.log_message("Running Selenium session directly...")

            # Run in background thread to avoid blocking GUI
            thread = threading.Thread(target=selenium_debug_session.main)
            thread.daemon = True
            thread.start()

            messagebox.showinfo("Selenium Started", "Selenium session started in background thread.")
        except Exception as e:
            error_msg = f"Error running Selenium directly: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)

    def export_to_csv(self):
        """Export tickets from the current CSV file to a new CSV with all database columns"""
        snow_file = self.snow_file.get().strip()

        if not snow_file:
            messagebox.showerror("Error", "Please select a ServiceNow CSV file first")
            return

        if not os.path.exists(snow_file):
            messagebox.showerror("Error", f"ServiceNow file not found: {snow_file}")
            return

        try:
            # Ask user where to save the CSV file
            export_file = filedialog.asksaveasfilename(
                title="Save CSV Report",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )

            if not export_file:
                return  # User cancelled

            self.log_message("Starting CSV export...")

            # Read the CSV file to get ticket numbers using basic file reading
            ticket_numbers = []
            with open(snow_file, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)

                # Get ticket column name (handle different formats)
                if 'number' in reader.fieldnames:
                    ticket_column = 'number'
                elif 'ticket' in reader.fieldnames:
                    ticket_column = 'ticket'
                else:
                    # Try first column
                    ticket_column = reader.fieldnames[0]

                for row in reader:
                    if ticket_column in row and row[ticket_column]:
                        ticket_numbers.append(row[ticket_column])

            self.log_message(f"Found {len(ticket_numbers)} tickets in CSV file")

            # Connect to database and get matching records
            conn = sqlite3.connect('ticket_matching.db')
            cursor = conn.cursor()

            # Create placeholders for SQL IN clause
            placeholders = ','.join(['?' for _ in ticket_numbers])
            query = f'''
                SELECT ticket, short_description, eml_domain, account_number, account_name, extraction_status
                FROM snow
                WHERE ticket IN ({placeholders})
                ORDER BY ticket
            '''

            # Execute query
            cursor.execute(query, ticket_numbers)
            results = cursor.fetchall()
            conn.close()

            # Write to CSV file
            with open(export_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Write header
                headers = ['ticket', 'short_description', 'eml_domain', 'account_number', 'account_name', 'extraction_status']
                writer.writerow(headers)

                # Write data rows
                for row in results:
                    # Convert None values to empty strings for better CSV display
                    clean_row = ['' if cell is None else str(cell) for cell in row]
                    writer.writerow(clean_row)

            self.log_message(f"CSV export completed: {len(results)} tickets exported to {export_file}")
            messagebox.showinfo("Export Complete",
                               f"Successfully exported {len(results)} tickets to:\n{export_file}")

        except Exception as e:
            error_msg = f"Error exporting to CSV: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Export Error", error_msg)


def main():
    root = tk.Tk()
    app = TicketMatchingGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()