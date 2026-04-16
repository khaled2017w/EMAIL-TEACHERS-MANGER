# ============================================================================
# Grade Email System - Cleaned for GitHub
# ============================================================================
# IMPORTANT: Before using this application, replace the following placeholders:
#   1. YOUR_GMAIL_ADDRESS      -> your Gmail address (e.g., you@gmail.com)
#   2. YOUR_GMAIL_APP_PASSWORD -> your Gmail App Password (NOT your regular password)
#   3. PUT_YOUR_MAIN_FOLDER_ID_HERE -> Google Drive folder ID where student folders are stored
#
# Also, you need to place a 'credentials.json' file (OAuth 2.0 Client ID) in the same directory
# for Google Drive API access. See: https://developers.google.com/drive/api/quickstart/python
# ============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
import threading
import smtplib
import os
import pickle
import io
import json
import schedule  # type: ignore
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
import urllib.parse
from google.auth.transport.requests import Request # type: ignore
from google_auth_oauthlib.flow import InstalledAppFlow # type: ignore
from googleapiclient.discovery import build # type: ignore
from googleapiclient.http import MediaIoBaseDownload # type: ignore

# Changed to full Drive scope to allow folder creation
SCOPES = ['https://www.googleapis.com/auth/drive']
DATABASE_FILE = 'students_database.json'
SCHEDULE_FILE = 'email_schedule.json'
ANALYTICS_FILE = 'email_analytics.json'
SETTINGS_FILE = 'app_settings.json'

class EmailSystemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🎓 Grade Email System")
        self.root.geometry("1000x750")  # Increased size for new features
        self.root.configure(bg='#f0f0f0')
        
        # Initialize theme
        self.current_theme = 'light'
        self.load_settings()
        
        # Initialize output buffer for early logging
        self._output_buffer = []
        
        # Load databases
        self.STUDENT_DATABASE = self.load_student_database()
        self.SCHEDULED_EMAILS = self.load_scheduled_emails()
        self.EMAIL_ANALYTICS = self.load_email_analytics()
        
        # Schedule monitoring thread
        self.schedule_monitor_running = True
        self.schedule_thread = threading.Thread(target=self.monitor_schedules, daemon=True)
        self.schedule_thread.start()
        
        self.setup_gui()
        
        # Flush any buffered output
        self._flush_output_buffer()
        
    def load_settings(self):
        """Load application settings including theme"""
        default_settings = {
            'theme': 'light',
            'auto_save': True,
            'notifications': True
        }
        
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
                    self.current_theme = settings.get('theme', 'light')
            except:
                pass
        
        self.apply_theme()
        
    def save_settings(self):
        """Save application settings"""
        settings = {
            'theme': self.current_theme,
            'auto_save': True,
            'notifications': True
        }
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
            
    def apply_theme(self):
        """Apply current theme to the application"""
        if self.current_theme == 'dark':
            self.bg_color = '#2b2b2b'
            self.fg_color = '#ffffff'
            self.accent_color = '#bb86fc'
            self.secondary_bg = '#3c3c3c'
            self.text_color = '#e0e0e0'
        else:  # light theme
            self.bg_color = '#f0f0f0'
            self.fg_color = '#000000'
            self.accent_color = '#667eea'
            self.secondary_bg = '#ffffff'
            self.text_color = '#333333'
            
        # Update UI if already initialized
        if hasattr(self, 'root'):
            self.update_theme_colors()
        
    def update_theme_colors(self):
        """Update all UI elements with current theme colors"""
        # This would need to be called after GUI setup to update colors
        # For now, theme is applied during initial setup
        pass
        
    def toggle_theme(self):
        """Toggle between dark and light themes"""
        self.current_theme = 'dark' if self.current_theme == 'light' else 'light'
        self.apply_theme()
        self.save_settings()
        self.log_output(f"🌙 Theme changed to {self.current_theme} mode")
        
    def load_student_database(self):
        if os.path.exists(DATABASE_FILE):
            try:
                with open(DATABASE_FILE, 'r') as f:
                    return json.load(f)
            except:
                pass
        
        # If no file exists, create empty database
        empty_database = {}
        self.save_student_database(empty_database)
        return empty_database

    def save_student_database(self, database):
        with open(DATABASE_FILE, 'w') as f:
            json.dump(database, f, indent=4)
            
    def load_scheduled_emails(self):
        """Load scheduled emails from file"""
        if os.path.exists(SCHEDULE_FILE):
            try:
                with open(SCHEDULE_FILE, 'r') as f:
                    return json.load(f)
            except:
                pass
        return []
        
    def save_scheduled_emails(self):
        """Save scheduled emails to file"""
        with open(SCHEDULE_FILE, 'w') as f:
            json.dump(self.SCHEDULED_EMAILS, f, indent=4)
            
    def load_email_analytics(self):
        """Load email analytics from file"""
        if os.path.exists(ANALYTICS_FILE):
            try:
                with open(ANALYTICS_FILE, 'r') as f:
                    return json.load(f)
            except:
                pass
        return {}
        
    def save_email_analytics(self):
        """Save email analytics to file"""
        with open(ANALYTICS_FILE, 'w') as f:
            json.dump(self.EMAIL_ANALYTICS, f, indent=4)
            
    def track_email_sent(self, student_name, email_type, success=True, scheduled=False):
        """Track email sending activity"""
        timestamp = datetime.now().isoformat()
        student_key = student_name.lower()
        
        if student_key not in self.EMAIL_ANALYTICS:
            self.EMAIL_ANALYTICS[student_key] = {
                'total_sent': 0,
                'successful': 0,
                'failed': 0,
                'scheduled': 0,
                'last_sent': None,
                'history': []
            }
            
        analytics = self.EMAIL_ANALYTICS[student_key]
        analytics['total_sent'] += 1
        analytics['last_sent'] = timestamp
        
        if scheduled:
            analytics['scheduled'] += 1
            
        if success:
            analytics['successful'] += 1
            status = '✅ SUCCESS'
        else:
            analytics['failed'] += 1
            status = '❌ FAILED'
            
        analytics['history'].append({
            'timestamp': timestamp,
            'type': email_type,
            'status': status,
            'scheduled': scheduled
        })
        
        # Keep only last 100 history entries
        if len(analytics['history']) > 100:
            analytics['history'] = analytics['history'][-100:]
            
        self.save_email_analytics()

    def setup_gui(self):
        # Header with theme toggle
        header_frame = tk.Frame(self.root, bg=self.accent_color, height=80)
        header_frame.pack(fill='x', padx=10, pady=10)
        header_frame.pack_propagate(False)
        
        header_content = tk.Frame(header_frame, bg=self.accent_color)
        header_content.pack(expand=True, fill='both')
        
        title_label = tk.Label(header_content, text="🎓 Grade Email System", 
                              font=('Arial', 20, 'bold'), bg=self.accent_color, fg='white')
        title_label.pack(side='left', expand=True)
        
        # Theme toggle button
        theme_btn = tk.Button(header_content, text="🌙" if self.current_theme == 'light' else "☀️",
                             command=self.toggle_theme, bg=self.accent_color, fg='white',
                             font=('Arial', 14), relief='flat', bd=0)
        theme_btn.pack(side='right', padx=10)
        
        # Main content frame
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Left frame - Student list and controls
        left_frame = tk.Frame(main_frame, bg=self.bg_color)
        left_frame.pack(side='left', fill='y', padx=(0, 10))
        
        # Student list
        student_list_label = tk.Label(left_frame, text="Students:", font=('Arial', 12, 'bold'), 
                                     bg=self.bg_color, fg=self.fg_color)
        student_list_label.pack(anchor='w', pady=(0, 5))
        
        self.student_listbox = tk.Listbox(left_frame, width=25, height=12, font=('Arial', 10),
                                         bg=self.secondary_bg, fg=self.text_color, selectbackground=self.accent_color)
        self.student_listbox.pack(fill='y', expand=True)
        
        # Student management buttons
        student_buttons_frame = tk.Frame(left_frame, bg=self.bg_color)
        student_buttons_frame.pack(fill='x', pady=5)
        
        tk.Button(student_buttons_frame, text="➕ Add Student", 
                 command=self.add_student_dialog, bg='#4CAF50', fg='white', font=('Arial', 10),
                 width=12).pack(side='left', padx=2)
        
        tk.Button(student_buttons_frame, text="🗑️ Remove", 
                 command=self.remove_student, bg='#f44336', fg='white', font=('Arial', 10),
                 width=10).pack(side='left', padx=2)
        
        # Populate student list
        self.refresh_student_list()
        
        # Right frame - Controls and output
        right_frame = tk.Frame(main_frame, bg=self.bg_color)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # Controls frame with notebook for tabs
        controls_notebook = ttk.Notebook(right_frame)
        controls_notebook.pack(fill='x', pady=(0, 10))
        
        # Email Operations Tab
        email_tab = ttk.Frame(controls_notebook)
        controls_notebook.add(email_tab, text="📧 Email Operations")
        
        # File operations
        file_frame = tk.LabelFrame(email_tab, text="File Operations", font=('Arial', 11, 'bold'), 
                                  bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        file_frame.pack(fill='x', pady=(0, 10))
        
        tk.Button(file_frame, text="Send All Files to Selected Student", 
                 command=self.send_all_to_selected, bg='#4CAF50', fg='white', font=('Arial', 10),
                 width=25).pack(pady=5)
        
        tk.Button(file_frame, text="Send Files to ALL Students", 
                 command=self.send_to_all_students, bg='#2196F3', fg='white', font=('Arial', 10),
                 width=25).pack(pady=5)
        
        # Single file sending
        file_send_frame = tk.Frame(file_frame, bg=self.bg_color)
        file_send_frame.pack(fill='x', pady=5)
        
        tk.Label(file_send_frame, text="Filename:", bg=self.bg_color, fg=self.fg_color).pack(side='left')
        self.filename_entry = tk.Entry(file_send_frame, width=20, font=('Arial', 10),
                                      bg=self.secondary_bg, fg=self.text_color)
        self.filename_entry.pack(side='left', padx=5)
        tk.Button(file_send_frame, text="Send File", 
                 command=self.send_single_file, bg='#FF9800', fg='white', font=('Arial', 9)).pack(side='left', padx=5)
        
        # Message operations
        message_frame = tk.LabelFrame(email_tab, text="Message Operations", font=('Arial', 11, 'bold'), 
                                     bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        message_frame.pack(fill='x', pady=(0, 10))
        
        self.message_text = scrolledtext.ScrolledText(message_frame, height=4, width=50, font=('Arial', 10),
                                                     bg=self.secondary_bg, fg=self.text_color)
        self.message_text.pack(fill='x', pady=5)
        
        msg_buttons_frame = tk.Frame(message_frame, bg=self.bg_color)
        msg_buttons_frame.pack(fill='x')
        
        tk.Button(msg_buttons_frame, text="Send Message to Selected", 
                 command=self.send_message_to_selected, bg='#9C27B0', fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        tk.Button(msg_buttons_frame, text="Send Message to ALL", 
                 command=self.send_message_to_all, bg='#E91E63', fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        # Scheduling Tab
        schedule_tab = ttk.Frame(controls_notebook)
        controls_notebook.add(schedule_tab, text="⏰ Scheduling")
        
        self.setup_scheduling_tab(schedule_tab)
        
        # Analytics Tab
        analytics_tab = ttk.Frame(controls_notebook)
        controls_notebook.add(analytics_tab, text="📊 Analytics")
        
        self.setup_analytics_tab(analytics_tab)
        
        # Utility operations
        util_frame = tk.LabelFrame(email_tab, text="Utilities", font=('Arial', 11, 'bold'), 
                                  bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        util_frame.pack(fill='x')
        
        util_buttons_frame = tk.Frame(util_frame, bg=self.bg_color)
        util_buttons_frame.pack(fill='x')
        
        tk.Button(util_buttons_frame, text="List Student Files", 
                 command=self.list_student_files, bg='#607D8B', fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        tk.Button(util_buttons_frame, text="Refresh Student List", 
                 command=self.refresh_student_list, bg='#795548', fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        # Output console
        output_frame = tk.LabelFrame(right_frame, text="Output Console", font=('Arial', 11, 'bold'), 
                                    bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        output_frame.pack(fill='both', expand=True)
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=12, width=70, font=('Consolas', 9),
                                                    bg=self.secondary_bg, fg=self.text_color)
        self.output_text.pack(fill='both', expand=True)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = tk.Label(self.root, textvariable=self.status_var, relief='sunken', anchor='w', 
                             bg='#e0e0e0', fg='#333333', font=('Arial', 9))
        status_bar.pack(side='bottom', fill='x')

    def setup_scheduling_tab(self, parent):
        """Setup the email scheduling tab"""
        # Schedule email frame
        schedule_frame = tk.LabelFrame(parent, text="Schedule Email", font=('Arial', 11, 'bold'),
                                      bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        schedule_frame.pack(fill='x', pady=5)
        
        # Email type selection
        type_frame = tk.Frame(schedule_frame, bg=self.bg_color)
        type_frame.pack(fill='x', pady=5)
        
        tk.Label(type_frame, text="Email Type:", bg=self.bg_color, fg=self.fg_color).pack(side='left')
        self.schedule_type = tk.StringVar(value="message")
        
        # Radio buttons with command to toggle visibility
        tk.Radiobutton(type_frame, text="Message", variable=self.schedule_type, value="message",
                      bg=self.bg_color, fg=self.fg_color, command=self.toggle_schedule_fields).pack(side='left', padx=10)
        tk.Radiobutton(type_frame, text="All Files", variable=self.schedule_type, value="all_files",
                      bg=self.bg_color, fg=self.fg_color, command=self.toggle_schedule_fields).pack(side='left', padx=10)
        tk.Radiobutton(type_frame, text="Single File", variable=self.schedule_type, value="single_file",
                      bg=self.bg_color, fg=self.fg_color, command=self.toggle_schedule_fields).pack(side='left', padx=10)
        
        # Message input for scheduled messages
        self.schedule_message_frame = tk.Frame(schedule_frame, bg=self.bg_color)
        self.schedule_message_frame.pack(fill='x', pady=5)
        
        tk.Label(self.schedule_message_frame, text="Message:", bg=self.bg_color, fg=self.fg_color).pack(anchor='w')
        self.schedule_message_text = scrolledtext.ScrolledText(self.schedule_message_frame, height=4, width=50, font=('Arial', 10),
                                                          bg=self.secondary_bg, fg=self.text_color)
        self.schedule_message_text.pack(fill='x', pady=5)
        
        # File selection for scheduled single file
        self.schedule_file_frame = tk.Frame(schedule_frame, bg=self.bg_color)
        
        file_selection_frame = tk.Frame(self.schedule_file_frame, bg=self.bg_color)
        file_selection_frame.pack(fill='x', pady=5)
        
        tk.Label(file_selection_frame, text="Select File:", bg=self.bg_color, fg=self.fg_color).pack(side='left')
        self.schedule_filename_var = tk.StringVar(value="No file selected")
        file_label = tk.Label(file_selection_frame, textvariable=self.schedule_filename_var, 
                             bg=self.secondary_bg, fg=self.text_color, relief='sunken', width=30, anchor='w')
        file_label.pack(side='left', padx=5)
        
        tk.Button(file_selection_frame, text="Browse Files", 
                 command=self.browse_schedule_files, bg='#2196F3', fg='white', font=('Arial', 9)).pack(side='left', padx=5)
        
        # Manual filename entry
        manual_file_frame = tk.Frame(self.schedule_file_frame, bg=self.bg_color)
        manual_file_frame.pack(fill='x', pady=5)
        
        tk.Label(manual_file_frame, text="Or enter filename:", bg=self.bg_color, fg=self.fg_color).pack(side='left')
        self.schedule_manual_filename = tk.Entry(manual_file_frame, width=25, font=('Arial', 10),
                                           bg=self.secondary_bg, fg=self.text_color)
        self.schedule_manual_filename.pack(side='left', padx=5)
        
        # Date and time selection
        datetime_frame = tk.Frame(schedule_frame, bg=self.bg_color)
        datetime_frame.pack(fill='x', pady=5)
        
        tk.Label(datetime_frame, text="Date (YYYY-MM-DD):", bg=self.bg_color, fg=self.fg_color).pack(side='left')
        self.schedule_date = tk.Entry(datetime_frame, width=12, font=('Arial', 10),
                                     bg=self.secondary_bg, fg=self.text_color)
        self.schedule_date.pack(side='left', padx=5)
        self.schedule_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        tk.Label(datetime_frame, text="Time (HH:MM):", bg=self.bg_color, fg=self.fg_color).pack(side='left', padx=(10,0))
        self.schedule_time = tk.Entry(datetime_frame, width=8, font=('Arial', 10),
                                     bg=self.secondary_bg, fg=self.text_color)
        self.schedule_time.pack(side='left', padx=5)
        self.schedule_time.insert(0, "09:00")
        
        # Schedule buttons
        schedule_buttons = tk.Frame(schedule_frame, bg=self.bg_color)
        schedule_buttons.pack(fill='x', pady=5)
        
        tk.Button(schedule_buttons, text="Schedule for Selected Student", 
                 command=self.schedule_for_selected, bg='#FF9800', fg='white').pack(side='left', padx=5)
        
        tk.Button(schedule_buttons, text="Schedule for All Students", 
                 command=self.schedule_for_all, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        # Scheduled emails list
        scheduled_list_frame = tk.LabelFrame(parent, text="Scheduled Emails", font=('Arial', 11, 'bold'),
                                           bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        scheduled_list_frame.pack(fill='both', expand=True, pady=5)
        
        # Treeview for scheduled emails
        columns = ('Student', 'Type', 'File/Message', 'Scheduled For', 'Status')
        self.schedule_tree = ttk.Treeview(scheduled_list_frame, columns=columns, show='headings', height=8)
        
        # Configure column widths
        self.schedule_tree.heading('Student', text='Student')
        self.schedule_tree.heading('Type', text='Type')
        self.schedule_tree.heading('File/Message', text='File/Message')
        self.schedule_tree.heading('Scheduled For', text='Scheduled For')
        self.schedule_tree.heading('Status', text='Status')
        
        self.schedule_tree.column('Student', width=120)
        self.schedule_tree.column('Type', width=100)
        self.schedule_tree.column('File/Message', width=150)
        self.schedule_tree.column('Scheduled For', width=120)
        self.schedule_tree.column('Status', width=100)
        
        self.schedule_tree.pack(fill='both', expand=True)
        
        # Schedule management buttons
        schedule_manage_frame = tk.Frame(scheduled_list_frame, bg=self.bg_color)
        schedule_manage_frame.pack(fill='x', pady=5)
        
        tk.Button(schedule_manage_frame, text="🔄 Refresh List", 
                 command=self.refresh_schedule_list, bg='#607D8B', fg='white').pack(side='left', padx=5)
        
        tk.Button(schedule_manage_frame, text="❌ Cancel Selected", 
                 command=self.cancel_scheduled, bg='#f44336', fg='white').pack(side='left', padx=5)
        
        # Initialize field visibility
        self.toggle_schedule_fields()
        self.refresh_schedule_list()

    def toggle_schedule_fields(self):
        """Show/hide message and file inputs based on email type selection"""
        email_type = self.schedule_type.get()
        
        # Hide both frames first
        self.schedule_message_frame.pack_forget()
        self.schedule_file_frame.pack_forget()
        
        # Show appropriate frame
        if email_type == "message":
            self.schedule_message_frame.pack(fill='x', pady=5)
        elif email_type == "single_file":
            self.schedule_file_frame.pack(fill='x', pady=5)

    def browse_schedule_files(self):
        """Browse and select files from the selected student's folder for scheduling"""
        student_name = self.get_selected_student()
        if not student_name:
            messagebox.showwarning("No Selection", "Please select a student first")
            return
        
        try:
            service = self.authenticate_google_drive()
            student_folder_id = self.get_student_folder_id(service, student_name)
            
            if not student_folder_id:
                messagebox.showerror("Error", f"No folder found for {student_name.title()}")
                return
            
            files = self.get_files_from_student_folder(service, student_folder_id)
            
            if not files:
                messagebox.showinfo("No Files", f"No files found for {student_name.title()}")
                return
            
            # Create file selection dialog
            self.show_schedule_file_selection_dialog(files, student_name)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to browse files: {str(e)}")

    def show_schedule_file_selection_dialog(self, files, student_name):
        """Show dialog to select a single file for scheduling"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Select File for Scheduling - {student_name.title()}")
        dialog.geometry("500x300")
        dialog.configure(bg=self.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # File list
        tk.Label(dialog, text=f"Select a file for {student_name.title()}:", 
                 font=('Arial', 11, 'bold'), bg=self.bg_color, fg=self.fg_color).pack(anchor='w', padx=10, pady=10)
        
        file_listbox = tk.Listbox(dialog, selectmode=tk.SINGLE, height=10,
                                 bg=self.secondary_bg, fg=self.text_color, 
                                 selectbackground=self.accent_color)
        file_listbox.pack(fill='both', expand=True, padx=10, pady=5)
        
        for file in files:
            file_listbox.insert(tk.END, file['name'])
        
        # Selection buttons
        button_frame = tk.Frame(dialog, bg=self.bg_color)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        def select_file():
            selection = file_listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a file")
                return
            
            filename = file_listbox.get(selection[0])
            self.schedule_filename_var.set(filename)
            self.schedule_selected_filename = filename
            self.schedule_selected_student = student_name
            dialog.destroy()
        
        tk.Button(button_frame, text="Select File", 
                 command=select_file, bg='#4CAF50', fg='white').pack(side='right', padx=5)
        tk.Button(button_frame, text="Cancel", 
                 command=dialog.destroy, bg='#f44336', fg='white').pack(side='right', padx=5)

    def setup_analytics_tab(self, parent):
        """Setup the email analytics tab"""
        # Analytics overview
        overview_frame = tk.LabelFrame(parent, text="Email Analytics Overview", font=('Arial', 11, 'bold'),
                                     bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        overview_frame.pack(fill='x', pady=5)
        
        # Stats frame
        stats_frame = tk.Frame(overview_frame, bg=self.bg_color)
        stats_frame.pack(fill='x', pady=5)
        
        self.total_emails_var = tk.StringVar(value="Total Emails: 0")
        self.success_rate_var = tk.StringVar(value="Success Rate: 0%")
        self.scheduled_emails_var = tk.StringVar(value="Scheduled: 0")
        
        tk.Label(stats_frame, textvariable=self.total_emails_var, font=('Arial', 10, 'bold'),
                bg=self.bg_color, fg=self.fg_color).pack(side='left', padx=20)
        tk.Label(stats_frame, textvariable=self.success_rate_var, font=('Arial', 10, 'bold'),
                bg=self.bg_color, fg=self.fg_color).pack(side='left', padx=20)
        tk.Label(stats_frame, textvariable=self.scheduled_emails_var, font=('Arial', 10, 'bold'),
                bg=self.bg_color, fg=self.fg_color).pack(side='left', padx=20)
        
        # Student analytics
        student_analytics_frame = tk.LabelFrame(parent, text="Student Email History", font=('Arial', 11, 'bold'),
                                              bg=self.bg_color, fg=self.fg_color, padx=10, pady=10)
        student_analytics_frame.pack(fill='both', expand=True, pady=5)
        
        # Treeview for student analytics
        columns = ('Student', 'Total Sent', 'Successful', 'Failed', 'Scheduled', 'Last Sent')
        self.analytics_tree = ttk.Treeview(student_analytics_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.analytics_tree.heading(col, text=col)
            self.analytics_tree.column(col, width=100)
        
        self.analytics_tree.pack(fill='both', expand=True)
        
        # Analytics buttons
        analytics_buttons = tk.Frame(student_analytics_frame, bg=self.bg_color)
        analytics_buttons.pack(fill='x', pady=5)
        
        tk.Button(analytics_buttons, text="Refresh Analytics", 
                 command=self.refresh_analytics, bg='#4CAF50', fg='white').pack(side='left', padx=5)
        
        tk.Button(analytics_buttons, text="Export Report", 
                 command=self.export_analytics, bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        tk.Button(analytics_buttons, text="Clear History", 
                 command=self.clear_analytics, bg='#f44336', fg='white').pack(side='left', padx=5)
        
        self.refresh_analytics()

    def schedule_for_selected(self):
        """Schedule email for selected student"""
        student_name = self.get_selected_student()
        if not student_name:
            messagebox.showwarning("No Selection", "Please select a student from the list")
            return
            
        self._schedule_email([student_name])
        
    def schedule_for_all(self):
        """Schedule email for all students"""
        student_names = list(self.STUDENT_DATABASE.keys())
        if not student_names:
            messagebox.showwarning("No Students", "No students in database")
            return
            
        self._schedule_email(student_names)
        
    def _schedule_email(self, student_names):
        """Schedule emails for given students"""
        try:
            email_type = self.schedule_type.get()
            date_str = self.schedule_date.get()
            time_str = self.schedule_time.get()
            
            # Validate datetime
            try:
                scheduled_datetime = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
                if scheduled_datetime < datetime.now():
                    messagebox.showerror("Error", "Scheduled time must be in the future")
                    return
            except ValueError:
                messagebox.showerror("Error", "Invalid date or time format. Use YYYY-MM-DD and HH:MM")
                return
                
            # Get content based on email type
            message_text = ""
            filename = ""
            
            if email_type == "message":
                message_text = self.schedule_message_text.get(1.0, tk.END).strip()
                if not message_text:
                    messagebox.showerror("Error", "Please enter a message for scheduled message emails")
                    return
                    
            elif email_type == "single_file":
                # Try manual filename first, then selected file
                filename = self.schedule_manual_filename.get().strip()
                if not filename and hasattr(self, 'schedule_selected_filename'):
                    filename = self.schedule_selected_filename
                
                if not filename:
                    messagebox.showerror("Error", "Please select or enter a filename for scheduled file emails")
                    return
            
            for student_name in student_names:
                schedule_id = f"{student_name}_{scheduled_datetime.isoformat()}"
                
                scheduled_email = {
                    'id': schedule_id,
                    'student_name': student_name,
                    'email_type': email_type,
                    'scheduled_datetime': scheduled_datetime.isoformat(),
                    'message_text': message_text,
                    'filename': filename,
                    'status': 'scheduled',
                    'created_at': datetime.now().isoformat()
                }
                
                self.SCHEDULED_EMAILS.append(scheduled_email)
                
                # Log appropriate message based on type
                if email_type == "message":
                    self.log_output(f"⏰ Scheduled message for {student_name.title()} at {scheduled_datetime}")
                elif email_type == "all_files":
                    self.log_output(f"⏰ Scheduled all files for {student_name.title()} at {scheduled_datetime}")
                else:  # single_file
                    self.log_output(f"⏰ Scheduled file '{filename}' for {student_name.title()} at {scheduled_datetime}")
                    
            self.save_scheduled_emails()
            self.refresh_schedule_list()
            messagebox.showinfo("Success", f"Scheduled {len(student_names)} email(s)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to schedule email: {str(e)}")
            
    def refresh_schedule_list(self):
        """Refresh the scheduled emails list"""
        for item in self.schedule_tree.get_children():
            self.schedule_tree.delete(item)
            
        now = datetime.now()
        
        for scheduled in self.SCHEDULED_EMAILS:
            scheduled_dt = datetime.fromisoformat(scheduled['scheduled_datetime'])
            student_name = scheduled['student_name'].title()
            
            # Determine type display
            if scheduled['email_type'] == "all_files":
                email_type = "All Files"
                content = "All files"
            elif scheduled['email_type'] == "single_file":
                email_type = "Single File"
                content = scheduled.get('filename', 'No file specified')
            else:  # message
                email_type = "Message"
                # Show first 30 chars of message
                message = scheduled.get('message_text', '')
                content = message[:30] + "..." if len(message) > 30 else message
            
            scheduled_for = scheduled_dt.strftime("%Y-%m-%d %H:%M")
            
            # Update status if past due
            if scheduled_dt < now and scheduled['status'] == 'scheduled':
                scheduled['status'] = 'missed'
                
            status = scheduled['status']
            
            self.schedule_tree.insert('', 'end', values=(
                student_name, email_type, content, scheduled_for, status
            ))
            
    def cancel_scheduled(self):
        """Cancel selected scheduled email"""
        selection = self.schedule_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a scheduled email to cancel")
            return
            
        item = selection[0]
        values = self.schedule_tree.item(item, 'values')
        student_name = values[0].lower()
        scheduled_for = values[3]
        
        # Find and remove the scheduled email
        self.SCHEDULED_EMAILS = [s for s in self.SCHEDULED_EMAILS 
                               if not (s['student_name'].lower() == student_name and 
                                      datetime.fromisoformat(s['scheduled_datetime']).strftime("%Y-%m-%d %H:%M") == scheduled_for)]
        
        self.save_scheduled_emails()
        self.refresh_schedule_list()
        self.log_output(f"❌ Cancelled scheduled email for {student_name.title()}")
        
    def monitor_schedules(self):
        """Monitor and execute scheduled emails"""
        while self.schedule_monitor_running:
            try:
                now = datetime.now()
                emails_to_send = []
                
                for scheduled in self.SCHEDULED_EMAILS[:]:  # Copy for iteration
                    if scheduled['status'] != 'scheduled':
                        continue
                        
                    scheduled_dt = datetime.fromisoformat(scheduled['scheduled_datetime'])
                    
                    # If it's time to send (within 1 minute)
                    if abs((scheduled_dt - now).total_seconds()) <= 60:
                        emails_to_send.append(scheduled)
                        scheduled['status'] = 'sending'
                        
                # Send emails
                for scheduled in emails_to_send:
                    try:
                        student_name = scheduled['student_name']
                        email_type = scheduled['email_type']
                        
                        if email_type == 'message':
                            success = self.send_text_to_single_student(
                                student_name, 
                                scheduled['message_text'],
                                scheduled=True
                            )
                        elif email_type == 'all_files':
                            success = self.send_all_files_to_student(
                                student_name,
                                scheduled=True
                            )
                        else:  # single_file
                            filename = scheduled.get('filename', '')
                            if filename:
                                success = self.send_single_file_to_student(
                                    student_name,
                                    filename,
                                    scheduled=True
                                )
                            else:
                                success = False
                                self.log_output(f"❌ No filename specified for scheduled file email to {student_name.title()}")
                                
                        if success:
                            scheduled['status'] = 'sent'
                            scheduled['sent_at'] = datetime.now().isoformat()
                            self.log_output(f"✅ Sent scheduled {email_type} email to {student_name.title()}")
                        else:
                            scheduled['status'] = 'failed'
                            scheduled['failed_at'] = datetime.now().isoformat()
                            self.log_output(f"❌ Failed to send scheduled email to {student_name.title()}")
                            
                    except Exception as e:
                        scheduled['status'] = 'failed'
                        scheduled['failed_at'] = datetime.now().isoformat()
                        self.log_output(f"❌ Error sending scheduled email: {str(e)}")
                        
                self.save_scheduled_emails()
                
            except Exception as e:
                self.log_output(f"❌ Schedule monitor error: {str(e)}")
                
            time.sleep(30)  # Check every 30 seconds
            
    def refresh_analytics(self):
        """Refresh analytics display"""
        # Update overview stats
        total_emails = 0
        successful_emails = 0
        scheduled_emails = 0
        
        for student_analytics in self.EMAIL_ANALYTICS.values():
            total_emails += student_analytics['total_sent']
            successful_emails += student_analytics['successful']
            scheduled_emails += student_analytics['scheduled']
            
        success_rate = (successful_emails / total_emails * 100) if total_emails > 0 else 0
        
        self.total_emails_var.set(f"Total Emails: {total_emails}")
        self.success_rate_var.set(f"Success Rate: {success_rate:.1f}%")
        self.scheduled_emails_var.set(f"Scheduled: {scheduled_emails}")
        
        # Update student analytics tree
        for item in self.analytics_tree.get_children():
            self.analytics_tree.delete(item)
            
        for student_name, analytics in self.EMAIL_ANALYTICS.items():
            last_sent = analytics['last_sent']
            if last_sent:
                last_sent = datetime.fromisoformat(last_sent).strftime("%Y-%m-%d %H:%M")
            else:
                last_sent = "Never"
                
            self.analytics_tree.insert('', 'end', values=(
                student_name.title(),
                analytics['total_sent'],
                analytics['successful'],
                analytics['failed'],
                analytics['scheduled'],
                last_sent
            ))
            
    def export_analytics(self):
        """Export analytics to CSV"""
        try:
            filename = f"email_analytics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w') as f:
                f.write("Student,Total Sent,Successful,Failed,Scheduled,Last Sent\n")
                for student_name, analytics in self.EMAIL_ANALYTICS.items():
                    last_sent = analytics['last_sent'] or "Never"
                    f.write(f"{student_name.title()},{analytics['total_sent']},{analytics['successful']},"
                           f"{analytics['failed']},{analytics['scheduled']},{last_sent}\n")
                    
            self.log_output(f"📊 Analytics exported to {filename}")
            messagebox.showinfo("Success", f"Analytics exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export analytics: {str(e)}")
            
    def clear_analytics(self):
        """Clear all analytics data"""
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all analytics data?"):
            self.EMAIL_ANALYTICS = {}
            self.save_email_analytics()
            self.refresh_analytics()
            self.log_output("🗑️ Cleared all analytics data")

    def add_student_dialog(self):
        """Open dialog to add a new student"""
        dialog = AddStudentDialog(self.root, self)
        self.root.wait_window(dialog.top)
        
    def add_student(self, student_name, student_email, main_folder_id):
        """Add a new student and create their folder in Google Drive"""
        try:
            # Validate inputs
            if not student_name or not student_email:
                messagebox.showerror("Error", "Please provide both name and email")
                return False
                
            # Normalize student name for consistency
            student_name_lower = student_name.lower().strip()
            
            # Check if student already exists
            if student_name_lower in self.STUDENT_DATABASE:
                messagebox.showerror("Error", f"Student '{student_name}' already exists!")
                return False
            
            self.set_status(f"Adding student {student_name}...")
            
            # Create folder in Google Drive
            folder_id = self.create_student_folder(student_name_lower, main_folder_id)
            
            if not folder_id:
                messagebox.showerror("Error", f"Failed to create folder for {student_name}")
                return False
            
            # Add to database
            self.STUDENT_DATABASE[student_name_lower] = {
                "email": student_email.strip(),
                "main_folder_id": main_folder_id,
                "subfolder": student_name_lower
            }
            
            # Save database
            self.save_student_database(self.STUDENT_DATABASE)
            
            # Refresh UI
            self.refresh_student_list()
            
            self.log_output(f"✅ SUCCESS: Added student '{student_name}' with email '{student_email}'")
            self.log_output(f"📁 Created folder: {student_name_lower} in Google Drive")
            self.set_status("Ready")
            
            messagebox.showinfo("Success", f"Student '{student_name}' added successfully!\nFolder created in Google Drive.")
            return True
            
        except Exception as e:
            error_msg = f"Error adding student: {str(e)}"
            self.log_output(f"❌ {error_msg}")
            messagebox.showerror("Error", error_msg)
            self.set_status("Error occurred")
            return False

    def create_student_folder(self, student_name, parent_folder_id):
        """Create a folder for the student in Google Drive"""
        try:
            service = self.authenticate_google_drive()
            
            # Check if folder already exists
            query = f"'{parent_folder_id}' in parents and name='{student_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            results = service.files().list(q=query).execute()
            existing_folders = results.get('files', [])
            
            if existing_folders:
                self.log_output(f"📁 Folder '{student_name}' already exists, using existing folder")
                return existing_folders[0]['id']
            
            # Create new folder
            folder_metadata = {
                'name': student_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }
            
            folder = service.files().create(body=folder_metadata, fields='id').execute()
            self.log_output(f"✅ Created folder '{student_name}' with ID: {folder.get('id')}")
            return folder.get('id')
            
        except Exception as e:
            self.log_output(f"❌ Error creating folder: {str(e)}")
            return None

    def remove_student(self):
        """Remove selected student from database"""
        student_name = self.get_selected_student()
        if not student_name:
            return
            
        if messagebox.askyesno("Confirm Removal", 
                              f"Are you sure you want to remove student '{student_name.title()}'?\n\nThis will remove them from the database but NOT delete their Google Drive folder."):
            try:
                # Remove from database
                if student_name in self.STUDENT_DATABASE:
                    del self.STUDENT_DATABASE[student_name]
                    self.save_student_database(self.STUDENT_DATABASE)
                    self.refresh_student_list()
                    self.log_output(f"✅ Removed student: {student_name.title()}")
                else:
                    messagebox.showerror("Error", f"Student '{student_name}' not found in database")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove student: {str(e)}")

    def log_output(self, message):
        """Safely log output to the console, buffering if GUI not ready"""
        if hasattr(self, 'output_text'):
            self.output_text.insert(tk.END, message + "\n")
            self.output_text.see(tk.END)
            self.root.update_idletasks()
        else:
            # Buffer messages until GUI is ready
            self._output_buffer.append(message)

    def _flush_output_buffer(self):
        """Flush any buffered output messages"""
        if hasattr(self, 'output_text') and self._output_buffer:
            for message in self._output_buffer:
                self.output_text.insert(tk.END, message + "\n")
            self.output_text.see(tk.END)
            self._output_buffer.clear()

    def set_status(self, message):
        if hasattr(self, 'status_var'):
            self.status_var.set(message)
            self.root.update_idletasks()

    def get_selected_student(self):
        selection = self.student_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a student from the list")
            return None
        student_display = self.student_listbox.get(selection[0])
        return student_display.lower()

    def send_all_to_selected(self):
        student_name = self.get_selected_student()
        if student_name:
            self.set_status(f"Sending all files to {student_name}...")
            thread = threading.Thread(target=self._send_all_files, args=(student_name,))
            thread.daemon = True
            thread.start()

    def _send_all_files(self, student_name):
        try:
            self.log_output(f"\n🚀 Sending ALL files to: {student_name.title()}")
            success = self.send_all_files_to_student(student_name)
            if success:
                self.log_output(f"✅ SUCCESS: All files sent to {student_name.title()}")
                self.set_status("Ready")
            else:
                self.log_output(f"❌ FAILED: Could not send files to {student_name.title()}")
                self.set_status("Failed")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def send_to_all_students(self):
        self.set_status("Sending files to ALL students...")
        thread = threading.Thread(target=self._send_to_all)
        thread.daemon = True
        thread.start()

    def _send_to_all(self):
        try:
            self.log_output(f"\n📧 Sending files to ALL students...")
            success = self.send_files_to_all_students()
            if success:
                self.log_output("✅ SUCCESS: Files sent to all students")
                self.set_status("Ready")
            else:
                self.log_output("❌ FAILED: Could not send files to all students")
                self.set_status("Failed")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def send_single_file(self):
        student_name = self.get_selected_student()
        filename = self.filename_entry.get().strip()
        if not filename:
            messagebox.showwarning("No Filename", "Please enter a filename")
            return
        
        if student_name:
            self.set_status(f"Sending {filename} to {student_name}...")
            thread = threading.Thread(target=self._send_single_file, args=(student_name, filename))
            thread.daemon = True
            thread.start()

    def _send_single_file(self, student_name, filename):
        try:
            self.log_output(f"\n🚀 Sending '{filename}' to: {student_name.title()}")
            success = self.send_single_file_to_student(student_name, filename)
            if success:
                self.log_output(f"✅ SUCCESS: File '{filename}' sent to {student_name.title()}")
                self.set_status("Ready")
            else:
                self.log_output(f"❌ FAILED: Could not send file '{filename}' to {student_name.title()}")
                self.set_status("Failed")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def send_message_to_selected(self):
        student_name = self.get_selected_student()
        message = self.message_text.get(1.0, tk.END).strip()
        if not message:
            messagebox.showwarning("No Message", "Please enter a message")
            return
        
        if student_name:
            self.set_status(f"Sending message to {student_name}...")
            thread = threading.Thread(target=self._send_message_to_selected, args=(student_name, message))
            thread.daemon = True
            thread.start()

    def _send_message_to_selected(self, student_name, message):
        try:
            self.log_output(f"\n💬 Sending message to: {student_name.title()}")
            success = self.send_text_to_single_student(student_name, message)
            if success:
                self.log_output(f"✅ SUCCESS: Message sent to {student_name.title()}")
                self.set_status("Ready")
            else:
                self.log_output(f"❌ FAILED: Could not send message to {student_name.title()}")
                self.set_status("Failed")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def send_message_to_all(self):
        message = self.message_text.get(1.0, tk.END).strip()
        if not message:
            messagebox.showwarning("No Message", "Please enter a message")
            return
        
        self.set_status("Sending message to ALL students...")
        thread = threading.Thread(target=self._send_message_to_all, args=(message,))
        thread.daemon = True
        thread.start()

    def _send_message_to_all(self, message):
        try:
            self.log_output(f"\n📢 Sending message to ALL students...")
            success = self.send_text_to_all_students(message)
            if success:
                self.log_output("✅ SUCCESS: Message sent to all students")
                self.set_status("Ready")
            else:
                self.log_output("❌ FAILED: Could not send message to all students")
                self.set_status("Failed")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def list_student_files(self):
        student_name = self.get_selected_student()
        if student_name:
            self.set_status(f"Listing files for {student_name}...")
            thread = threading.Thread(target=self._list_student_files, args=(student_name,))
            thread.daemon = True
            thread.start()

    def _list_student_files(self, student_name):
        try:
            self.log_output(f"\n📁 Listing files for {student_name.title()}...")
            self.list_student_files_func(student_name)
            self.set_status("Ready")
        except Exception as e:
            self.log_output(f"❌ ERROR: {str(e)}")
            self.set_status("Error occurred")

    def refresh_student_list(self):
        self.STUDENT_DATABASE = self.load_student_database()
        self.student_listbox.delete(0, tk.END)
        for student in self.STUDENT_DATABASE.keys():
            self.student_listbox.insert(tk.END, student.title())
        # Don't log here during initialization to avoid the error
        if hasattr(self, 'output_text'):
            self.log_output("✅ Student list refreshed")
        self.set_status("Ready")

    # Core functionality methods
    def authenticate_google_drive(self):
        creds = None
        # Delete the old token to force re-authentication with new scopes
        if os.path.exists('token.pickle'):
            # Check if we need to re-authenticate with new scopes
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
            
            # If the credentials don't have the required scopes, delete the token
            if creds and not creds.has_scopes(SCOPES):
                self.log_output("🔄 Re-authenticating with new permissions...")
                os.remove('token.pickle')
                creds = None
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        return build('drive', 'v3', credentials=creds)

    def get_student_folder_id(self, service, student_name):
        main_folder_id = self.STUDENT_DATABASE[student_name]["main_folder_id"]
        student_subfolder = self.STUDENT_DATABASE[student_name]["subfolder"]
        
        query = f"'{main_folder_id}' in parents and name='{student_subfolder}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        results = service.files().list(q=query).execute()
        folders = results.get('files', [])
        
        return folders[0]['id'] if folders else None

    def get_files_from_student_folder(self, service, folder_id):
        query = f"'{folder_id}' in parents and trashed=false"
        results = service.files().list(q=query).execute()
        return results.get('files', [])

    def download_drive_file(self, service, file_id, filename):
        try:
            request = service.files().get_media(fileId=file_id)
            file_bytes = io.BytesIO()
            downloader = MediaIoBaseDownload(file_bytes, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            return file_bytes.getvalue()
        except Exception as e:
            self.log_output(f"Error downloading {filename}: {e}")
            return None

    def encode_filename(self, filename):
        """Properly encode filename for email attachment, especially for Arabic characters"""
        try:
            # Check if filename contains non-ASCII characters
            if any(ord(char) > 127 for char in filename):
                # Use proper encoding for non-ASCII filenames
                encoded_name = Header(filename, 'utf-8').encode()
                return f"UTF-8''{urllib.parse.quote(filename, safe='')}"
            else:
                return filename
        except Exception as e:
            self.log_output(f"⚠️  Warning: Could not encode filename '{filename}', using original: {e}")
            return filename

    def send_all_files_to_student(self, student_name, recipient_email=None, scheduled=False):
        try:
            if not recipient_email:
                recipient_email = self.STUDENT_DATABASE[student_name]["email"]
            
            self.log_output(f"Connecting to Google Drive...")
            service = self.authenticate_google_drive()
            
            self.log_output(f"Looking for {student_name}'s folder...")
            student_folder_id = self.get_student_folder_id(service, student_name)
            
            if not student_folder_id:
                self.log_output(f"❌ {student_name}'s folder not found")
                return False
            
            self.log_output(f"Getting files from {student_name}'s folder...")
            files = self.get_files_from_student_folder(service, student_folder_id)
            
            if not files:
                self.log_output(f"❌ No files found in {student_name}'s folder")
                return False
            
            self.log_output(f"Found {len(files)} files for {student_name}")
            
            conn = smtplib.SMTP("smtp.gmail.com", 587)
            conn.ehlo()
            conn.starttls()
            conn.login("YOUR_GMAIL_ADDRESS", "YOUR_GMAIL_APP_PASSWORD")
            
            msg = MIMEMultipart()
            msg['From'] = "YOUR_GMAIL_ADDRESS"
            msg['To'] = recipient_email
            msg['Subject'] = f"📚 All Your Grade Files - {student_name.title()}"
            
            html_body = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{
                        font-family: 'Arial', sans-serif;
                        line-height: 1.6;
                        color: #333;
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 30px;
                        text-align: center;
                        border-radius: 10px 10px 0 0;
                    }}
                    .content {{
                        background: #f8f9fa;
                        padding: 30px;
                        border-radius: 0 0 10px 10px;
                    }}
                    .student-name {{
                        color: #667eea;
                        font-size: 24px;
                        font-weight: bold;
                        margin: 10px 0;
                    }}
                    .file-list {{
                        background: white;
                        padding: 20px;
                        border-radius: 8px;
                        margin: 20px 0;
                        border-left: 4px solid #667eea;
                    }}
                    .footer {{
                        text-align: center;
                        margin-top: 30px;
                        color: #666;
                        font-size: 12px;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h1>🎓 Academic Records</h1>
                    <p>All Your Grade Files</p>
                </div>
                
                <div class="content">
                    <div class="student-name">
                        Dear {student_name.title()},
                    </div>
                    
                    <p>All your grade files have been attached to this email.</p>
                    
                    <div class="file-list">
                        <h3>📋 Attached Files ({len(files)} files):</h3>
                        <ul>
            """
            
            for file in files:
                html_body += f"<li>📄 {file['name']}</li>"
            
            html_body += f"""
                        </ul>
                    </div>
                    
                    <p>All files have been retrieved from your personal grade folder.</p>
                    <p>If you have any questions about your grades, please contact your academic advisor.</p>
                </div>
                
                <div class="footer">
                    <p>Best regards,<br>
                    <strong>Academic Records Department</strong></p>
                </div>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(html_body, 'html'))
            
            files_attached = 0
            for file in files:
                try:
                    self.log_output(f"⬇️  Downloading: {file['name']}")
                    file_content = self.download_drive_file(service, file['id'], file['name'])
                    
                    if file_content:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(file_content)
                        encoders.encode_base64(part)
                        
                        # Use proper filename encoding for Arabic/Unicode characters
                        encoded_filename = self.encode_filename(file['name'])
                        part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
                        
                        msg.attach(part)
                        files_attached += 1
                        self.log_output(f"✅ Attached: {file['name']}")
                    else:
                        self.log_output(f"❌ Failed to download: {file['name']}")
                    
                except Exception as file_error:
                    self.log_output(f"❌ Error attaching {file['name']}: {file_error}")
            
            conn.sendmail("YOUR_GMAIL_ADDRESS", recipient_email, msg.as_string())
            conn.quit()
            
            self.log_output(f"🎉 SUCCESS: Sent {files_attached} files to {student_name} at {recipient_email}")
            
            # Track the email
            self.track_email_sent(student_name, "all_files", success=True, scheduled=scheduled)
            return True
            
        except Exception as e:
            self.log_output(f"❌ ERROR: {e}")
            # Track the email
            self.track_email_sent(student_name, "all_files", success=False, scheduled=scheduled)
            return False

    def send_files_to_all_students(self):
        try:
            service = self.authenticate_google_drive()
            total_students = len(self.STUDENT_DATABASE)
            success_count = 0
            
            self.log_output(f"📧 Sending files to all {total_students} students...")
            
            for student_name, student_info in self.STUDENT_DATABASE.items():
                try:
                    self.log_output(f"\n🎯 Processing {student_name.title()}...")
                    
                    student_folder_id = self.get_student_folder_id(service, student_name)
                    
                    if not student_folder_id:
                        self.log_output(f"❌ {student_name}'s folder not found - skipping")
                        continue
                    
                    files = self.get_files_from_student_folder(service, student_folder_id)
                    
                    if not files:
                        self.log_output(f"❌ No files found for {student_name} - skipping")
                        continue
                    
                    self.log_output(f"Found {len(files)} files for {student_name}")
                    
                    conn = smtplib.SMTP("smtp.gmail.com", 587)
                    conn.ehlo()
                    conn.starttls()
                    conn.login("YOUR_GMAIL_ADDRESS", "YOUR_GMAIL_APP_PASSWORD")
                    
                    msg = MIMEMultipart()
                    msg['From'] = "YOUR_GMAIL_ADDRESS"
                    msg['To'] = student_info["email"]
                    msg['Subject'] = f"📚 Your Grade Files - {student_name.title()}"
                    
                    html_body = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                            body {{
                                font-family: 'Arial', sans-serif;
                                line-height: 1.6;
                                color: #333;
                                max-width: 600px;
                                margin: 0 auto;
                                padding: 20px;
                            }}
                            .header {{
                                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                                color: white;
                                padding: 30px;
                                text-align: center;
                                border-radius: 10px 10px 0 0;
                            }}
                            .content {{
                                background: #f8f9fa;
                                padding: 30px;
                                border-radius: 0 0 10px 10px;
                            }}
                            .student-name {{
                                color: #667eea;
                                font-size: 24px;
                                font-weight: bold;
                                margin: 10px 0;
                            }}
                            .file-list {{
                                background: white;
                                padding: 20px;
                                border-radius: 8px;
                                margin: 20px 0;
                                border-left: 4px solid #667eea;
                            }}
                            .footer {{
                                text-align: center;
                                margin-top: 30px;
                                color: #666;
                                font-size: 12px;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <h1>🎓 Academic Records</h1>
                            <p>Your Grade Files</p>
                        </div>
                        
                        <div class="content">
                            <div class="student-name">
                                Dear {student_name.title()},
                            </div>
                            
                            <p>Your grade files have been attached to this email.</p>
                            
                            <div class="file-list">
                                <h3>📋 Attached Files ({len(files)} files):</h3>
                                <ul>
                    """
                    
                    for file in files:
                        html_body += f"<li>📄 {file['name']}</li>"
                    
                    html_body += f"""
                                </ul>
                            </div>
                            
                            <p>These files have been retrieved from your personal grade folder.</p>
                            <p>If you have any questions about your grades, please contact your academic advisor.</p>
                        </div>
                        
                        <div class="footer">
                            <p>Best regards,<br>
                            <strong>Academic Records Department</strong></p>
                        </div>
                    </body>
                    </html>
                    """
                    
                    msg.attach(MIMEText(html_body, 'html'))
                    
                    files_attached = 0
                    for file in files:
                        try:
                            file_content = self.download_drive_file(service, file['id'], file['name'])
                            
                            if file_content:
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(file_content)
                                encoders.encode_base64(part)
                                
                                # Use proper filename encoding for Arabic/Unicode characters
                                encoded_filename = self.encode_filename(file['name'])
                                part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
                                
                                msg.attach(part)
                                files_attached += 1
                                self.log_output(f"✅ Attached: {file['name']}")
                            else:
                                self.log_output(f"❌ Failed to download: {file['name']}")
                            
                        except Exception as file_error:
                            self.log_output(f"❌ Error attaching {file['name']}: {file_error}")
                    
                    conn.sendmail("YOUR_GMAIL_ADDRESS", student_info["email"], msg.as_string())
                    conn.quit()
                    
                    self.log_output(f"🎉 Sent {files_attached} files to {student_name}")
                    success_count += 1
                    
                    # Track the email
                    self.track_email_sent(student_name, "all_files", success=True, scheduled=False)
                    
                except Exception as e:
                    self.log_output(f"❌ Failed to send to {student_name}: {e}")
                    # Track the email
                    self.track_email_sent(student_name, "all_files", success=False, scheduled=False)
            
            self.log_output(f"\n🎊 COMPLETED: Successfully sent files to {success_count}/{total_students} students")
            return True
            
        except Exception as e:
            self.log_output(f"❌ ERROR: {e}")
            return False

    def send_single_file_to_student(self, student_name, filename, recipient_email=None, scheduled=False):
        try:
            if not recipient_email:
                recipient_email = self.STUDENT_DATABASE[student_name]["email"]
            
            self.log_output(f"Connecting to Google Drive...")
            service = self.authenticate_google_drive()
            
            self.log_output(f"Looking for {student_name}'s folder...")
            student_folder_id = self.get_student_folder_id(service, student_name)
            
            if not student_folder_id:
                self.log_output(f"❌ {student_name}'s folder not found")
                return False
            
            self.log_output(f"Getting files from {student_name}'s folder...")
            files = self.get_files_from_student_folder(service, student_folder_id)
            
            if not files:
                self.log_output(f"❌ No files found in {student_name}'s folder")
                return False
            
            target_file = None
            for file in files:
                if file['name'].lower() == filename.lower():
                    target_file = file
                    break
            
            if not target_file:
                self.log_output(f"❌ File '{filename}' not found in {student_name}'s folder")
                self.log_output(f"Available files: {[f['name'] for f in files]}")
                return False
            
            conn = smtplib.SMTP("smtp.gmail.com", 587)
            conn.ehlo()
            conn.starttls()
            conn.login("YOUR_GMAIL_ADDRESS", "YOUR_GMAIL_APP_PASSWORD")
            
            msg = MIMEMultipart()
            msg['From'] = "YOUR_GMAIL_ADDRESS"
            msg['To'] = recipient_email
            msg['Subject'] = f"📄 {filename} - {student_name.title()}"
            
            html_body = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{
                        font-family: 'Arial', sans-serif;
                        line-height: 1.6;
                        color: #333;
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 30px;
                        text-align: center;
                        border-radius: 10px 10px 0 0;
                    }}
                    .content {{
                        background: #f8f9fa;
                        padding: 30px;
                        border-radius: 0 0 10px 10px;
                    }}
                    .student-name {{
                        color: #667eea;
                        font-size: 24px;
                        font-weight: bold;
                        margin: 10px 0;
                    }}
                    .file-info {{
                        background: white;
                        padding: 20px;
                        border-radius: 8px;
                        margin: 20px 0;
                        border-left: 4px solid #667eea;
                        text-align: center;
                    }}
                    .footer {{
                        text-align: center;
                        margin-top: 30px;
                        color: #666;
                        font-size: 12px;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h1>🎓 Academic Records</h1>
                    <p>Your Requested File</p>
                </div>
                
                <div class="content">
                    <div class="student-name">
                        Dear {student_name.title()},
                    </div>
                    
                    <p>Your requested file has been attached to this email.</p>
                    
                    <div class="file-info">
                        <h3>📄 Attached File:</h3>
                        <p><strong>{target_file['name']}</strong></p>
                        <p>This file has been retrieved from your personal grade folder.</p>
                    </div>
                    
                    <p>If you have any questions about this file, please contact your academic advisor.</p>
                </div>
                
                <div class="footer">
                    <p>Best regards,<br>
                    <strong>Academic Records Department</strong></p>
                </div>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(html_body, 'html'))
            
            self.log_output(f"⬇️  Downloading: {target_file['name']}")
            file_content = self.download_drive_file(service, target_file['id'], target_file['name'])
            
            if file_content:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file_content)
                encoders.encode_base64(part)
                
                # Use proper filename encoding for Arabic/Unicode characters
                encoded_filename = self.encode_filename(target_file['name'])
                part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
                
                msg.attach(part)
                self.log_output(f"✅ Attached: {target_file['name']}")
            else:
                self.log_output(f"❌ Failed to download: {target_file['name']}")
                return False
            
            conn.sendmail("YOUR_GMAIL_ADDRESS", recipient_email, msg.as_string())
            conn.quit()
            
            self.log_output(f"🎉 SUCCESS: Sent '{target_file['name']}' to {student_name} at {recipient_email}")
            
            # Track the email
            self.track_email_sent(student_name, "single_file", success=True, scheduled=scheduled)
            return True
            
        except Exception as e:
            self.log_output(f"❌ ERROR: {e}")
            # Track the email
            self.track_email_sent(student_name, "single_file", success=False, scheduled=scheduled)
            return False

    def send_text_to_single_student(self, student_name, message_text, scheduled=False):
        try:
            recipient_email = self.STUDENT_DATABASE[student_name]["email"]
            
            conn = smtplib.SMTP("smtp.gmail.com", 587)
            conn.ehlo()
            conn.starttls()
            conn.login("YOUR_GMAIL_ADDRESS", "YOUR_GMAIL_APP_PASSWORD")
            
            msg = MIMEMultipart()
            msg['From'] = "YOUR_GMAIL_ADDRESS"
            msg['To'] = recipient_email
            msg['Subject'] = "📢 Message from Academic Department"
            
            html_body = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {{
                        font-family: 'Arial', sans-serif;
                        line-height: 1.6;
                        color: #333;
                        max-width: 600px;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .header {{
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 30px;
                        text-align: center;
                        border-radius: 10px 10px 0 0;
                    }}
                    .content {{
                        background: #f8f9fa;
                        padding: 30px;
                        border-radius: 0 0 10px 10px;
                    }}
                    .student-name {{
                        color: #667eea;
                        font-size: 24px;
                        font-weight: bold;
                        margin: 10px 0;
                    }}
                    .message-box {{
                        background: white;
                        padding: 20px;
                        border-radius: 8px;
                        margin: 20px 0;
                        border-left: 4px solid #667eea;
                    }}
                    .footer {{
                        text-align: center;
                        margin-top: 30px;
                        color: #666;
                        font-size: 12px;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h1>🎓 Academic Message</h1>
                    <p>Personal Communication</p>
                </div>
                
                <div class="content">
                    <div class="student-name">
                        Dear {student_name.title()},
                    </div>
                    
                    <div class="message-box">
                        <p>{message_text}</p>
                    </div>
                    
                    <p>If you have any questions, please don't hesitate to contact us.</p>
                </div>
                
                <div class="footer">
                    <p>Best regards,<br>
                    <strong>Academic Records Department</strong></p>
                </div>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(html_body, 'html'))
            conn.sendmail("YOUR_GMAIL_ADDRESS", recipient_email, msg.as_string())
            conn.quit()
            
            self.log_output(f"✅ SUCCESS: Sent message to {student_name} at {recipient_email}")
            
            # Track the email
            self.track_email_sent(student_name, "message", success=True, scheduled=scheduled)
            return True
            
        except Exception as e:
            self.log_output(f"❌ ERROR sending to {student_name}: {e}")
            # Track the email
            self.track_email_sent(student_name, "message", success=False, scheduled=scheduled)
            return False

    def send_text_to_all_students(self, message_text):
        try:
            conn = smtplib.SMTP("smtp.gmail.com", 587)
            conn.ehlo()
            conn.starttls()
            conn.login("YOUR_GMAIL_ADDRESS", "YOUR_GMAIL_APP_PASSWORD")
            
            success_count = 0
            total_students = len(self.STUDENT_DATABASE)
            
            for student_name, student_info in self.STUDENT_DATABASE.items():
                try:
                    msg = MIMEMultipart()
                    msg['From'] = "YOUR_GMAIL_ADDRESS"
                    msg['To'] = student_info["email"]
                    msg['Subject'] = "📢 Important Message from Academic Department"
                    
                    html_body = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                            body {{
                                font-family: 'Arial', sans-serif;
                                line-height: 1.6;
                                color: #333;
                                max-width: 600px;
                                margin: 0 auto;
                                padding: 20px;
                            }}
                            .header {{
                                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                                color: white;
                                padding: 30px;
                                text-align: center;
                                border-radius: 10px 10px 0 0;
                            }}
                            .content {{
                                background: #f8f9fa;
                                padding: 30px;
                                border-radius: 0 0 10px 10px;
                            }}
                            .student-name {{
                                color: #667eea;
                                font-size: 24px;
                                font-weight: bold;
                                margin: 10px 0;
                            }}
                            .message-box {{
                                background: white;
                                padding: 20px;
                                border-radius: 8px;
                                margin: 20px 0;
                                border-left: 4px solid #667eea;
                            }}
                            .footer {{
                                text-align: center;
                                margin-top: 30px;
                                color: #666;
                                font-size: 12px;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <h1>🎓 Academic Announcement</h1>
                            <p>Important Message</p>
                        </div>
                        
                        <div class="content">
                            <div class="student-name">
                                Dear {student_name.title()},
                            </div>
                            
                            <div class="message-box">
                                <p>{message_text}</p>
                            </div>
                            
                            <p>If you have any questions, please don't hesitate to contact us.</p>
                        </div>
                        
                        <div class="footer">
                            <p>Best regards,<br>
                            <strong>Academic Records Department</strong></p>
                        </div>
                    </body>
                    </html>
                    """
                    
                    msg.attach(MIMEText(html_body, 'html'))
                    conn.sendmail("YOUR_GMAIL_ADDRESS", student_info["email"], msg.as_string())
                    success_count += 1
                    self.log_output(f"✅ Sent to {student_name.title()}")
                    
                    # Track the email
                    self.track_email_sent(student_name, "message", success=True, scheduled=False)
                    
                except Exception as e:
                    self.log_output(f"❌ Failed to send to {student_name}: {e}")
                    # Track the email
                    self.track_email_sent(student_name, "message", success=False, scheduled=False)
            
            conn.quit()
            self.log_output(f"🎉 SUCCESS: Sent message to {success_count}/{total_students} students")
            return True
            
        except Exception as e:
            self.log_output(f"❌ ERROR: {e}")
            return False

    def list_student_files_func(self, student_name):
        try:
            service = self.authenticate_google_drive()
            student_folder_id = self.get_student_folder_id(service, student_name)
            
            if not student_folder_id:
                self.log_output(f"❌ {student_name}'s folder not found")
                return
            
            files = self.get_files_from_student_folder(service, student_folder_id)
            
            if not files:
                self.log_output(f"❌ No files found in {student_name}'s folder")
                return
            
            self.log_output(f"\n📁 Files available for {student_name.title()}:")
            for i, file in enumerate(files, 1):
                self.log_output(f"   {i}. {file['name']}")
                
        except Exception as e:
            self.log_output(f"❌ ERROR: {e}")


class AddStudentDialog:
    def __init__(self, parent, main_app):
        self.main_app = main_app
        self.top = tk.Toplevel(parent)
        self.top.title("➕ Add New Student")
        self.top.geometry("500x400")
        self.top.configure(bg='#f0f0f0')
        self.top.transient(parent)
        self.top.grab_set()
        
        self.setup_dialog()
        
    def setup_dialog(self):
        # Simple vertical layout using pack
        tk.Label(self.top, text="➕ Add New Student", 
                font=('Arial', 16, 'bold'), bg='#667eea', fg='white',
                pady=10).pack(fill='x', padx=20, pady=10)
        
        # Student Name
        tk.Label(self.top, text="Student Name:", 
                font=('Arial', 11, 'bold'), bg='#f0f0f0').pack(anchor='w', padx=20, pady=(20, 5))
        self.name_entry = tk.Entry(self.top, width=40, font=('Arial', 11))
        self.name_entry.pack(fill='x', padx=20, pady=(0, 15))
        
        # Email
        tk.Label(self.top, text="Email Address:", 
                font=('Arial', 11, 'bold'), bg='#f0f0f0').pack(anchor='w', padx=20, pady=(0, 5))
        self.email_entry = tk.Entry(self.top, width=40, font=('Arial', 11))
        self.email_entry.pack(fill='x', padx=20, pady=(0, 15))
        
        # Main Folder ID
        tk.Label(self.top, text="Main Folder ID:", 
                font=('Arial', 11, 'bold'), bg='#f0f0f0').pack(anchor='w', padx=20, pady=(0, 5))
        self.folder_entry = tk.Entry(self.top, width=40, font=('Arial', 11))
        self.folder_entry.pack(fill='x', padx=20, pady=(0, 10))
        # Placeholder – user must replace this with their own folder ID
        self.folder_entry.insert(0, "PUT_YOUR_MAIN_FOLDER_ID_HERE")
        
        # Help text
        help_text = "📁 The Main Folder ID is the Google Drive folder where the student's personal folder will be created."
        tk.Label(self.top, text=help_text, font=('Arial', 9), 
                bg='#f0f0f0', fg='#666', justify='left').pack(anchor='w', padx=20, pady=(0, 30))
        
        # Buttons - very simple layout
        button_frame = tk.Frame(self.top, bg='#f0f0f0')
        button_frame.pack(fill='x', padx=20, pady=20)
        
        tk.Button(button_frame, text="Add Student", 
                 command=self.add_student, 
                 bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'),
                 width=15, height=1).pack(side='right', padx=(10, 5))
        
        tk.Button(button_frame, text="Cancel", 
                 command=self.top.destroy,
                 bg='#f44336', fg='white', font=('Arial', 12),
                 width=10, height=1).pack(side='right', padx=5)
        
        self.name_entry.focus_set()
        self.top.bind('<Return>', lambda e: self.add_student())
        
    def add_student(self):
        name = self.name_entry.get().strip()
        email = self.email_entry.get().strip()
        folder_id = self.folder_entry.get().strip()
        
        if not name:
            messagebox.showerror("Error", "Please enter student name")
            return
            
        if not email:
            messagebox.showerror("Error", "Please enter email address")
            return
            
        if not folder_id:
            messagebox.showerror("Error", "Please enter main folder ID")
            return
        
        if '@' not in email or '.' not in email:
            if not messagebox.askyesno("Confirm Email", 
                                      f"The email '{email}' doesn't look valid. Continue anyway?"):
                return
        
        if self.main_app.add_student(name, email, folder_id):
            self.top.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSystemGUI(root)
    root.mainloop()
