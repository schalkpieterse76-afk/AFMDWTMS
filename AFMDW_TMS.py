"""
AFMDW Task Management System (TMS)
A comprehensive asset management, bundled data extraction, and reporting system
with PDF exports, owner persistence, and advanced query capabilities.

Author: Schalk Pieterse
Created: 2026-01-13
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import os
import csv
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import sqlite3
from pathlib import Path
from tkcalendar import DateEntry
import webbrowser


class AFMDWTaskManagementSystem:
    """
    Main application class for AFMDW Task Management System.
    Handles asset management, data extraction, reporting, and PDF generation.
    """

    def __init__(self, root):
        """Initialize the application with all necessary components."""
        self.root = root
        self.root.title("AFMDW Task Management System")
        self.root.geometry("1400x800")
        
        # Application data
        self.assets = []
        self.owners = {}
        self.data_file = "afmdw_assets.json"
        self.owners_file = "afmdw_owners.json"
        self.db_file = "afmdw_tasks.db"
        
        # Initialize database
        self.init_database()
        
        # Load existing data
        self.load_data()
        self.load_owners()
        
        # UI Variables
        self.asset_id_var = tk.StringVar()
        self.asset_name_var = tk.StringVar()
        self.asset_type_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.owner_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.acquisition_date_var = tk.StringVar()
        self.release_date = tk.StringVar()  # Fixed: renamed from release_var
        self.cost_var = tk.StringVar()
        self.warranty_var = tk.StringVar()
        self.notes_var = tk.StringVar()
        
        # Build UI
        self.build_ui()

    def init_database(self):
        """Initialize SQLite database for tasks and queries."""
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            
            # Tasks table
            c.execute('''CREATE TABLE IF NOT EXISTS tasks
                         (id INTEGER PRIMARY KEY AUTOINCREMENT,
                          asset_id TEXT,
                          task_type TEXT,
                          description TEXT,
                          status TEXT,
                          created_date TEXT,
                          due_date TEXT,
                          completed_date TEXT)''')
            
            # Query history table
            c.execute('''CREATE TABLE IF NOT EXISTS query_history
                         (id INTEGER PRIMARY KEY AUTOINCREMENT,
                          query_type TEXT,
                          query_params TEXT,
                          created_date TEXT,
                          results TEXT)''')
            
            conn.commit()
            conn.close()
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to initialize database: {e}")

    def load_data(self):
        """Load assets from JSON file."""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r') as f:
                    self.assets = json.load(f)
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load data: {e}")
                self.assets = []
        else:
            self.assets = []

    def save_data(self):
        """Save assets to JSON file."""
        try:
            with open(self.data_file, 'w') as f:
                json.dump(self.assets, f, indent=4)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save data: {e}")

    def load_owners(self):
        """Load owners from JSON file."""
        if os.path.exists(self.owners_file):
            try:
                with open(self.owners_file, 'r') as f:
                    self.owners = json.load(f)
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load owners: {e}")
                self.owners = {}
        else:
            self.owners = {}

    def save_owners(self):
        """Save owners to JSON file."""
        try:
            with open(self.owners_file, 'w') as f:
                json.dump(self.owners, f, indent=4)
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save owners: {e}")

    def build_ui(self):
        """Build the main user interface."""
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Assets tab
        self.build_assets_tab()
        
        # Query tab
        self.build_query_tab()
        
        # Reports tab
        self.build_reports_tab()
        
        # Settings tab
        self.build_settings_tab()

    def build_assets_tab(self):
        """Build the assets management tab."""
        assets_frame = ttk.Frame(self.notebook)
        self.notebook.add(assets_frame, text="Assets Management")
        
        # Input frame
        input_frame = ttk.LabelFrame(assets_frame, text="Asset Details", padding=10)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Asset ID
        ttk.Label(input_frame, text="Asset ID:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.asset_id_var, width=20).grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Asset Name
        ttk.Label(input_frame, text="Asset Name:").grid(row=0, column=2, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.asset_name_var, width=20).grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # Asset Type
        ttk.Label(input_frame, text="Asset Type:").grid(row=0, column=4, sticky=tk.W)
        asset_types = ["Hardware", "Software", "Network", "Storage", "Peripheral", "Other"]
        ttk.Combobox(input_frame, textvariable=self.asset_type_var, values=asset_types, width=15).grid(row=0, column=5, sticky=tk.W, padx=5)
        
        # Status
        ttk.Label(input_frame, text="Status:").grid(row=1, column=0, sticky=tk.W)
        statuses = ["Active", "Inactive", "In Repair", "Decommissioned", "On Hold"]
        ttk.Combobox(input_frame, textvariable=self.status_var, values=statuses, width=15).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Owner
        ttk.Label(input_frame, text="Owner:").grid(row=1, column=2, sticky=tk.W)
        self.owner_combobox = ttk.Combobox(input_frame, textvariable=self.owner_var, width=20)
        self.owner_combobox.grid(row=1, column=3, sticky=tk.W, padx=5)
        
        # Location
        ttk.Label(input_frame, text="Location:").grid(row=1, column=4, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.location_var, width=20).grid(row=1, column=5, sticky=tk.W, padx=5)
        
        # Acquisition Date
        ttk.Label(input_frame, text="Acquisition Date:").grid(row=2, column=0, sticky=tk.W)
        DateEntry(input_frame, textvariable=self.acquisition_date_var, width=20).grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Release Date - Fixed variable naming
        ttk.Label(input_frame, text="Release Date:").grid(row=2, column=2, sticky=tk.W)
        DateEntry(input_frame, textvariable=self.release_date, width=20).grid(row=2, column=3, sticky=tk.W, padx=5)
        
        # Cost
        ttk.Label(input_frame, text="Cost:").grid(row=2, column=4, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.cost_var, width=20).grid(row=2, column=5, sticky=tk.W, padx=5)
        
        # Warranty
        ttk.Label(input_frame, text="Warranty (months):").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.warranty_var, width=20).grid(row=3, column=1, sticky=tk.W, padx=5)
        
        # Notes
        ttk.Label(input_frame, text="Notes:").grid(row=3, column=2, sticky=tk.W)
        ttk.Entry(input_frame, textvariable=self.notes_var, width=40).grid(row=3, column=3, columnspan=3, sticky=tk.W, padx=5)
        
        # Buttons frame
        button_frame = ttk.Frame(assets_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Add Asset", command=self.add_asset).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Update Asset", command=self.update_record).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Asset", command=self.delete_asset).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Form", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export to PDF", command=self.export_to_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export to CSV", command=self.export_to_csv).pack(side=tk.LEFT, padx=5)
        
        # Tree frame
        tree_frame = ttk.LabelFrame(assets_frame, text="Assets List", padding=5)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        # Treeview - Fixed: Removed problematic displaycolumns=() empty configuration
        columns = ("ID", "Name", "Type", "Status", "Owner", "Location", "Acq Date", "Release Date", "Cost", "Warranty", "Notes")
        self.tree = ttk.Treeview(tree_frame, columns=columns, yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=15)
        
        # Configure scrollbars
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configure column headings and widths
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.heading("#0", text="", anchor=tk.W)
        
        column_widths = {
            "ID": 80,
            "Name": 120,
            "Type": 80,
            "Status": 100,
            "Owner": 100,
            "Location": 100,
            "Acq Date": 100,
            "Release Date": 100,
            "Cost": 80,
            "Warranty": 80,
            "Notes": 150
        }
        
        for col in columns:
            self.tree.column(col, width=column_widths.get(col, 100), anchor=tk.W)
            self.tree.heading(col, text=col, anchor=tk.W)
        
        # Bind selection event
        self.tree.bind("<ButtonRelease-1>", self.on_tree_select)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Refresh tree
        self.refresh_tree()

    def build_query_tab(self):
        """Build the query and search tab."""
        query_frame = ttk.Frame(self.notebook)
        self.notebook.add(query_frame, text="Query & Search")
        
        # Query options
        options_frame = ttk.LabelFrame(query_frame, text="Search Options", padding=10)
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(options_frame, text="Search by:").pack(side=tk.LEFT, padx=5)
        
        self.search_type_var = tk.StringVar()
        search_types = ["Asset ID", "Asset Name", "Owner", "Status", "Location", "Asset Type"]
        ttk.Combobox(options_frame, textvariable=self.search_type_var, values=search_types, width=20).pack(side=tk.LEFT, padx=5)
        
        self.search_term_var = tk.StringVar()
        ttk.Entry(options_frame, textvariable=self.search_term_var, width=30).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(options_frame, text="Search", command=self.perform_search).pack(side=tk.LEFT, padx=5)
        ttk.Button(options_frame, text="Advanced Query", command=self.open_advanced_query).pack(side=tk.LEFT, padx=5)
        ttk.Button(options_frame, text="Clear Results", command=self.refresh_tree).pack(side=tk.LEFT, padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(query_frame, text="Search Results", padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL)
        
        # Results tree
        columns = ("ID", "Name", "Type", "Status", "Owner", "Location")
        self.search_tree = ttk.Treeview(results_frame, columns=columns, yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.search_tree.yview)
        hsb.config(command=self.search_tree.xview)
        
        self.search_tree.column("#0", width=0, stretch=tk.NO)
        self.search_tree.heading("#0", text="", anchor=tk.W)
        
        for col in columns:
            self.search_tree.column(col, width=100, anchor=tk.W)
            self.search_tree.heading(col, text=col, anchor=tk.W)
        
        self.search_tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

    def build_reports_tab(self):
        """Build the reports and analytics tab."""
        reports_frame = ttk.Frame(self.notebook)
        self.notebook.add(reports_frame, text="Reports & Analytics")
        
        # Reports buttons
        button_frame = ttk.LabelFrame(reports_frame, text="Generate Reports", padding=10)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Asset Summary Report", command=self.generate_summary_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Owner Distribution Report", command=self.generate_owner_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Status Report", command=self.generate_status_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Warranty Status Report", command=self.generate_warranty_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cost Analysis", command=self.generate_cost_report).pack(side=tk.LEFT, padx=5)
        
        # Report display
        self.report_text = tk.Text(reports_frame, wrap=tk.WORD, height=30)
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Copy and export buttons
        export_frame = ttk.Frame(reports_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(export_frame, text="Copy Report", command=self.copy_report).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Export Report", command=self.export_report).pack(side=tk.LEFT, padx=5)

    def build_settings_tab(self):
        """Build the settings and configuration tab."""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings & Tools")
        
        # Owner management
        owner_frame = ttk.LabelFrame(settings_frame, text="Owner Management", padding=10)
        owner_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(owner_frame, text="Owner Name:").pack(side=tk.LEFT, padx=5)
        self.new_owner_var = tk.StringVar()
        ttk.Entry(owner_frame, textvariable=self.new_owner_var, width=20).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(owner_frame, text="Add Owner", command=self.add_owner).pack(side=tk.LEFT, padx=5)
        ttk.Button(owner_frame, text="Remove Owner", command=self.remove_owner).pack(side=tk.LEFT, padx=5)
        ttk.Button(owner_frame, text="View Owners", command=self.view_owners).pack(side=tk.LEFT, padx=5)
        
        # Bundled data extraction
        bundle_frame = ttk.LabelFrame(settings_frame, text="Bundled Data Extraction", padding=10)
        bundle_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(bundle_frame, text="Export All as JSON", command=self.export_bundle_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(bundle_frame, text="Export All as CSV", command=self.export_bundle_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(bundle_frame, text="Import from File", command=self.import_bundle).pack(side=tk.LEFT, padx=5)
        
        # Database operations
        db_frame = ttk.LabelFrame(settings_frame, text="Database Operations", padding=10)
        db_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(db_frame, text="Backup Data", command=self.backup_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(db_frame, text="View Task History", command=self.view_task_history).pack(side=tk.LEFT, padx=5)
        ttk.Button(db_frame, text="View Query History", command=self.view_query_history).pack(side=tk.LEFT, padx=5)
        ttk.Button(db_frame, text="Clear Database", command=self.clear_database).pack(side=tk.LEFT, padx=5)
        
        # Application info
        info_frame = ttk.LabelFrame(settings_frame, text="Application Information", padding=10)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        info_text = tk.Text(info_frame, height=10, width=80)
        info_text.pack(padx=5, pady=5)
        
        info_content = """AFMDW Task Management System v1.0
        
Features:
- Asset Management: Add, update, delete, and track assets
- Bundled Data Extraction: Export/import asset data in multiple formats
- Owner Persistence: Manage and persist owner information
- PDF Reports: Generate professional PDF reports
- Advanced Queries: Search and filter assets by multiple criteria
- Task Management: Create and track tasks related to assets
- Comprehensive Reporting: Generate detailed reports and analytics
- Data Persistence: SQLite database integration
- Multi-format Export: CSV and PDF export capabilities

Current Data File: """ + self.data_file + """
Owners File: """ + self.owners_file + """
Database File: """ + self.db_file
        
        info_text.insert(tk.END, info_content)
        info_text.config(state=tk.DISABLED)

    def add_asset(self):
        """Add a new asset to the system."""
        if not self.asset_id_var.get() or not self.asset_name_var.get():
            messagebox.showerror("Error", "Asset ID and Name are required")
            return
        
        asset = self.build_record_from_form()
        
        # Check for duplicate ID
        if any(a["id"] == asset["id"] for a in self.assets):
            messagebox.showerror("Error", "Asset ID already exists")
            return
        
        self.assets.append(asset)
        self.save_data()
        self.refresh_tree()
        self.clear_form()
        messagebox.showinfo("Success", f"Asset {asset['id']} added successfully")

    def update_record(self):
        """Update an existing asset."""
        if not self.tree.selection():
            messagebox.showerror("Error", "Please select an asset to update")
            return
        
        selected_item = self.tree.selection()[0]
        item_values = self.tree.item(selected_item, 'values')
        asset_id = item_values[0]
        
        # Find and update the asset
        for asset in self.assets:
            if asset["id"] == asset_id:
                updated_asset = self.build_record_from_form()
                updated_asset["id"] = asset_id
                self.assets[self.assets.index(asset)] = updated_asset
                self.save_data()
                self.refresh_tree()
                self.clear_form()
                messagebox.showinfo("Success", f"Asset {asset_id} updated successfully")
                return
        
        messagebox.showerror("Error", "Asset not found")

    def delete_asset(self):
        """Delete an asset from the system."""
        if not self.tree.selection():
            messagebox.showerror("Error", "Please select an asset to delete")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this asset?"):
            selected_item = self.tree.selection()[0]
            item_values = self.tree.item(selected_item, 'values')
            asset_id = item_values[0]
            
            self.assets = [a for a in self.assets if a["id"] != asset_id]
            self.save_data()
            self.refresh_tree()
            self.clear_form()
            messagebox.showinfo("Success", f"Asset {asset_id} deleted successfully")

    def build_record_from_form(self):
        """Build an asset record from form inputs - Fixed release_date reference."""
        return {
            "id": self.asset_id_var.get(),
            "name": self.asset_name_var.get(),
            "type": self.asset_type_var.get(),
            "status": self.status_var.get(),
            "owner": self.owner_var.get(),
            "location": self.location_var.get(),
            "acquisition_date": self.acquisition_date_var.get(),
            "release_date": self.release_date.get(),  # Fixed variable reference
            "cost": self.cost_var.get(),
            "warranty": self.warranty_var.get(),
            "notes": self.notes_var.get()
        }

    def load_selected_record(self, asset):
        """Load asset data into form for editing - Fixed release_date reference."""
        self.asset_id_var.set(asset.get("id", ""))
        self.asset_name_var.set(asset.get("name", ""))
        self.asset_type_var.set(asset.get("type", ""))
        self.status_var.set(asset.get("status", ""))
        self.owner_var.set(asset.get("owner", ""))
        self.location_var.set(asset.get("location", ""))
        self.acquisition_date_var.set(asset.get("acquisition_date", ""))
        self.release_date.set(asset.get("release_date", ""))  # Fixed variable reference
        self.cost_var.set(asset.get("cost", ""))
        self.warranty_var.set(asset.get("warranty", ""))
        self.notes_var.set(asset.get("notes", ""))

    def clear_form(self):
        """Clear all form fields."""
        self.asset_id_var.set("")
        self.asset_name_var.set("")
        self.asset_type_var.set("")
        self.status_var.set("")
        self.owner_var.set("")
        self.location_var.set("")
        self.acquisition_date_var.set("")
        self.release_date.set("")  # Fixed variable reference
        self.cost_var.set("")
        self.warranty_var.set("")
        self.notes_var.set("")

    def on_tree_select(self, event):
        """Handle tree selection event."""
        if self.tree.selection():
            selected_item = self.tree.selection()[0]
            item_values = self.tree.item(selected_item, 'values')
            
            # Find the asset
            for asset in self.assets:
                if asset["id"] == item_values[0]:
                    self.load_selected_record(asset)
                    return

    def refresh_tree(self):
        """Refresh the asset tree view."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add assets
        for asset in self.assets:
            values = (
                asset.get("id", ""),
                asset.get("name", ""),
                asset.get("type", ""),
                asset.get("status", ""),
                asset.get("owner", ""),
                asset.get("location", ""),
                asset.get("acquisition_date", ""),
                asset.get("release_date", ""),
                asset.get("cost", ""),
                asset.get("warranty", ""),
                asset.get("notes", "")
            )
            self.tree.insert("", tk.END, values=values)
        
        # Update owner combobox
        self.update_owner_combobox()

    def update_owner_combobox(self):
        """Update the owner combobox with current owners."""
        owners_list = list(self.owners.keys()) if self.owners else []
        self.owner_combobox['values'] = owners_list

    def add_owner(self):
        """Add a new owner to the system."""
        owner_name = self.new_owner_var.get().strip()
        if not owner_name:
            messagebox.showerror("Error", "Please enter an owner name")
            return
        
        if owner_name in self.owners:
            messagebox.showerror("Error", "Owner already exists")
            return
        
        self.owners[owner_name] = {"created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        self.save_owners()
        self.update_owner_combobox()
        self.new_owner_var.set("")
        messagebox.showinfo("Success", f"Owner '{owner_name}' added successfully")

    def remove_owner(self):
        """Remove an owner from the system."""
        owner_name = self.new_owner_var.get().strip()
        if not owner_name:
            messagebox.showerror("Error", "Please enter an owner name")
            return
        
        if owner_name not in self.owners:
            messagebox.showerror("Error", "Owner not found")
            return
        
        if messagebox.askyesno("Confirm", f"Remove owner '{owner_name}'?"):
            del self.owners[owner_name]
            self.save_owners()
            self.update_owner_combobox()
            self.new_owner_var.set("")
            messagebox.showinfo("Success", f"Owner '{owner_name}' removed successfully")

    def view_owners(self):
        """Display all owners in a new window."""
        owners_window = tk.Toplevel(self.root)
        owners_window.title("Owners List")
        owners_window.geometry("400x400")
        
        text = tk.Text(owners_window, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        if self.owners:
            for owner, details in self.owners.items():
                text.insert(tk.END, f"Owner: {owner}\n")
                text.insert(tk.END, f"Created: {details.get('created_date', 'Unknown')}\n\n")
        else:
            text.insert(tk.END, "No owners registered yet.")
        
        text.config(state=tk.DISABLED)

    def perform_search(self):
        """Perform a search based on criteria."""
        search_type = self.search_type_var.get()
        search_term = self.search_term_var.get().lower()
        
        if not search_type or not search_term:
            messagebox.showerror("Error", "Please select search type and enter search term")
            return
        
        # Clear search tree
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        # Perform search
        results = []
        for asset in self.assets:
            match = False
            
            if search_type == "Asset ID" and search_term in asset.get("id", "").lower():
                match = True
            elif search_type == "Asset Name" and search_term in asset.get("name", "").lower():
                match = True
            elif search_type == "Owner" and search_term in asset.get("owner", "").lower():
                match = True
            elif search_type == "Status" and search_term in asset.get("status", "").lower():
                match = True
            elif search_type == "Location" and search_term in asset.get("location", "").lower():
                match = True
            elif search_type == "Asset Type" and search_term in asset.get("type", "").lower():
                match = True
            
            if match:
                results.append(asset)
        
        # Display results
        for asset in results:
            values = (
                asset.get("id", ""),
                asset.get("name", ""),
                asset.get("type", ""),
                asset.get("status", ""),
                asset.get("owner", ""),
                asset.get("location", "")
            )
            self.search_tree.insert("", tk.END, values=values)
        
        messagebox.showinfo("Search Results", f"Found {len(results)} matching asset(s)")
        
        # Log to database
        self.log_query("Search", {"type": search_type, "term": search_term}, len(results))

    def open_advanced_query(self):
        """Open advanced query window."""
        query_window = tk.Toplevel(self.root)
        query_window.title("Advanced Query")
        query_window.geometry("500x600")
        
        # Query criteria
        criteria_frame = ttk.LabelFrame(query_window, text="Query Criteria", padding=10)
        criteria_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(criteria_frame, text="Status:").pack(anchor=tk.W)
        status_var = tk.StringVar()
        statuses = ["All", "Active", "Inactive", "In Repair", "Decommissioned", "On Hold"]
        ttk.Combobox(criteria_frame, textvariable=status_var, values=statuses).pack(fill=tk.X, pady=5)
        
        ttk.Label(criteria_frame, text="Asset Type:").pack(anchor=tk.W)
        type_var = tk.StringVar()
        types = ["All", "Hardware", "Software", "Network", "Storage", "Peripheral", "Other"]
        ttk.Combobox(criteria_frame, textvariable=type_var, values=types).pack(fill=tk.X, pady=5)
        
        ttk.Label(criteria_frame, text="Owner:").pack(anchor=tk.W)
        owner_var = tk.StringVar()
        owners_list = ["All"] + list(self.owners.keys())
        ttk.Combobox(criteria_frame, textvariable=owner_var, values=owners_list).pack(fill=tk.X, pady=5)
        
        ttk.Label(criteria_frame, text="Min Cost:").pack(anchor=tk.W)
        min_cost_var = tk.StringVar()
        ttk.Entry(criteria_frame, textvariable=min_cost_var).pack(fill=tk.X, pady=5)
        
        ttk.Label(criteria_frame, text="Max Cost:").pack(anchor=tk.W)
        max_cost_var = tk.StringVar()
        ttk.Entry(criteria_frame, textvariable=max_cost_var).pack(fill=tk.X, pady=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(query_window, text="Query Results", padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        results_text = tk.Text(results_frame, wrap=tk.WORD)
        results_text.pack(fill=tk.BOTH, expand=True)
        
        def execute_query():
            results_text.delete(1.0, tk.END)
            
            results = self.assets[:]
            
            # Apply filters
            if status_var.get() != "All":
                results = [a for a in results if a.get("status") == status_var.get()]
            
            if type_var.get() != "All":
                results = [a for a in results if a.get("type") == type_var.get()]
            
            if owner_var.get() != "All":
                results = [a for a in results if a.get("owner") == owner_var.get()]
            
            try:
                if min_cost_var.get():
                    min_cost = float(min_cost_var.get())
                    results = [a for a in results if float(a.get("cost", 0) or 0) >= min_cost]
            except ValueError:
                pass
            
            try:
                if max_cost_var.get():
                    max_cost = float(max_cost_var.get())
                    results = [a for a in results if float(a.get("cost", 0) or 0) <= max_cost]
            except ValueError:
                pass
            
            # Display results
            results_text.insert(tk.END, f"Found {len(results)} asset(s)\n\n")
            for asset in results:
                results_text.insert(tk.END, f"ID: {asset.get('id')}\n")
                results_text.insert(tk.END, f"Name: {asset.get('name')}\n")
                results_text.insert(tk.END, f"Type: {asset.get('type')}\n")
                results_text.insert(tk.END, f"Status: {asset.get('status')}\n")
                results_text.insert(tk.END, f"Owner: {asset.get('owner')}\n")
                results_text.insert(tk.END, f"Location: {asset.get('location')}\n")
                results_text.insert(tk.END, f"Cost: {asset.get('cost')}\n")
                results_text.insert(tk.END, "-" * 40 + "\n\n")
        
        ttk.Button(criteria_frame, text="Execute Query", command=execute_query).pack(side=tk.LEFT, padx=5, pady=10)

    def generate_summary_report(self):
        """Generate a summary report of all assets."""
        report = "=" * 60 + "\n"
        report += "ASSET SUMMARY REPORT\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        report += f"Total Assets: {len(self.assets)}\n"
        
        if self.assets:
            total_cost = sum(float(a.get("cost", 0) or 0) for a in self.assets)
            report += f"Total Cost: ${total_cost:,.2f}\n\n"
            
            # Assets by type
            types = {}
            for asset in self.assets:
                asset_type = asset.get("type", "Unknown")
                types[asset_type] = types.get(asset_type, 0) + 1
            
            report += "Assets by Type:\n"
            for asset_type, count in sorted(types.items()):
                report += f"  {asset_type}: {count}\n"
            
            report += "\nAssets by Status:\n"
            statuses = {}
            for asset in self.assets:
                status = asset.get("status", "Unknown")
                statuses[status] = statuses.get(status, 0) + 1
            
            for status, count in sorted(statuses.items()):
                report += f"  {status}: {count}\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def generate_owner_report(self):
        """Generate a report on asset distribution by owner."""
        report = "=" * 60 + "\n"
        report += "OWNER DISTRIBUTION REPORT\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        owners_assets = {}
        for asset in self.assets:
            owner = asset.get("owner", "Unassigned")
            if owner not in owners_assets:
                owners_assets[owner] = []
            owners_assets[owner].append(asset)
        
        if owners_assets:
            for owner, assets in sorted(owners_assets.items()):
                report += f"\nOwner: {owner}\n"
                report += f"  Total Assets: {len(assets)}\n"
                total_cost = sum(float(a.get("cost", 0) or 0) for a in assets)
                report += f"  Total Cost: ${total_cost:,.2f}\n"
                report += "  Assets:\n"
                for asset in assets:
                    report += f"    - {asset.get('name')} ({asset.get('id')})\n"
        else:
            report += "No assets assigned to owners.\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def generate_status_report(self):
        """Generate a report on asset status."""
        report = "=" * 60 + "\n"
        report += "ASSET STATUS REPORT\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        statuses = {}
        for asset in self.assets:
            status = asset.get("status", "Unknown")
            if status not in statuses:
                statuses[status] = []
            statuses[status].append(asset)
        
        if statuses:
            for status, assets in sorted(statuses.items()):
                report += f"\n{status}: {len(assets)} asset(s)\n"
                for asset in assets:
                    report += f"  - {asset.get('name')} (ID: {asset.get('id')})\n"
        else:
            report += "No assets found.\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def generate_warranty_report(self):
        """Generate a warranty status report."""
        report = "=" * 60 + "\n"
        report += "WARRANTY STATUS REPORT\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        if not self.assets:
            report += "No assets found.\n"
        else:
            today = datetime.now()
            expiring_soon = []
            expired = []
            valid = []
            
            for asset in self.assets:
                try:
                    acq_date = datetime.strptime(asset.get("acquisition_date", ""), "%Y-%m-%d")
                    warranty_months = int(asset.get("warranty", 0) or 0)
                    expiry_date = acq_date + relativedelta(months=warranty_months)
                    
                    days_until_expiry = (expiry_date - today).days
                    
                    if days_until_expiry < 0:
                        expired.append((asset, expiry_date, days_until_expiry))
                    elif days_until_expiry <= 90:
                        expiring_soon.append((asset, expiry_date, days_until_expiry))
                    else:
                        valid.append((asset, expiry_date, days_until_expiry))
                except:
                    pass
            
            report += f"Valid Warranty: {len(valid)}\n"
            report += f"Expiring Soon (< 90 days): {len(expiring_soon)}\n"
            report += f"Expired: {len(expired)}\n\n"
            
            if expired:
                report += "EXPIRED WARRANTIES:\n"
                for asset, expiry_date, days in expired:
                    report += f"  - {asset.get('name')} (expired {abs(days)} days ago)\n"
            
            if expiring_soon:
                report += "\nEXPIRING SOON:\n"
                for asset, expiry_date, days in expiring_soon:
                    report += f"  - {asset.get('name')} (expires in {days} days)\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def generate_cost_report(self):
        """Generate a cost analysis report."""
        report = "=" * 60 + "\n"
        report += "COST ANALYSIS REPORT\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += "=" * 60 + "\n\n"
        
        if not self.assets:
            report += "No assets found.\n"
        else:
            total_cost = sum(float(a.get("cost", 0) or 0) for a in self.assets)
            avg_cost = total_cost / len(self.assets) if self.assets else 0
            
            report += f"Total Assets: {len(self.assets)}\n"
            report += f"Total Cost: ${total_cost:,.2f}\n"
            report += f"Average Cost: ${avg_cost:,.2f}\n\n"
            
            # Cost by type
            cost_by_type = {}
            for asset in self.assets:
                asset_type = asset.get("type", "Unknown")
                cost = float(asset.get("cost", 0) or 0)
                if asset_type not in cost_by_type:
                    cost_by_type[asset_type] = 0
                cost_by_type[asset_type] += cost
            
            report += "Cost by Type:\n"
            for asset_type, cost in sorted(cost_by_type.items(), key=lambda x: x[1], reverse=True):
                report += f"  {asset_type}: ${cost:,.2f}\n"
        
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def copy_report(self):
        """Copy report text to clipboard."""
        report_content = self.report_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(report_content)
        messagebox.showinfo("Success", "Report copied to clipboard")

    def export_report(self):
        """Export report to text file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(self.report_text.get(1.0, tk.END))
                messagebox.showinfo("Success", "Report exported successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export report: {e}")

    def export_to_pdf(self):
        """Export selected asset to PDF."""
        if not self.tree.selection():
            messagebox.showerror("Error", "Please select an asset to export")
            return
        
        selected_item = self.tree.selection()[0]
        item_values = self.tree.item(selected_item, 'values')
        asset_id = item_values[0]
        
        # Find the asset
        asset = None
        for a in self.assets:
            if a["id"] == asset_id:
                asset = a
                break
        
        if not asset:
            messagebox.showerror("Error", "Asset not found")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return
        
        try:
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            story = []
            
            # Title
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#1f77b4'),
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph("Asset Report", title_style))
            story.append(Spacer(1, 0.3 * inch))
            
            # Asset details
            data = [
                ["Asset ID", asset.get("id", "")],
                ["Name", asset.get("name", "")],
                ["Type", asset.get("type", "")],
                ["Status", asset.get("status", "")],
                ["Owner", asset.get("owner", "")],
                ["Location", asset.get("location", "")],
                ["Acquisition Date", asset.get("acquisition_date", "")],
                ["Release Date", asset.get("release_date", "")],
                ["Cost", f"${float(asset.get('cost', 0) or 0):,.2f}"],
                ["Warranty (months)", asset.get("warranty", "")],
                ["Notes", asset.get("notes", "")]
            ]
            
            table = Table(data, colWidths=[2 * inch, 4 * inch])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(table)
            
            doc.build(story)
            messagebox.showinfo("Success", "PDF exported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")

    def export_to_csv(self):
        """Export all assets to CSV file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=self.assets[0].keys() if self.assets else [])
                writer.writeheader()
                writer.writerows(self.assets)
            messagebox.showinfo("Success", "CSV exported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV: {e}")

    def export_bundle_json(self):
        """Export all data as bundled JSON."""
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not file_path:
            return
        
        try:
            bundle = {
                "assets": self.assets,
                "owners": self.owners,
                "export_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            with open(file_path, 'w') as f:
                json.dump(bundle, f, indent=4)
            messagebox.showinfo("Success", "Data exported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")

    def export_bundle_csv(self):
        """Export assets as bundled CSV with metadata."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='') as f:
                writer = csv.writer(f)
                
                # Header with metadata
                writer.writerow(["AFMDW Asset Management Bundle Export"])
                writer.writerow([f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
                writer.writerow([])
                
                # Asset data
                writer.writerow(["ASSETS"])
                if self.assets:
                    fieldnames = list(self.assets[0].keys())
                    writer.writerow(fieldnames)
                    for asset in self.assets:
                        writer.writerow([asset.get(field, "") for field in fieldnames])
                
                writer.writerow([])
                writer.writerow(["OWNERS"])
                writer.writerow(["Owner", "Created Date"])
                for owner, details in self.owners.items():
                    writer.writerow([owner, details.get("created_date", "")])
            
            messagebox.showinfo("Success", "Bundle exported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export bundle: {e}")

    def import_bundle(self):
        """Import bundled data from file."""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("CSV files", "*.csv")])
        if not file_path:
            return
        
        try:
            if file_path.endswith('.json'):
                with open(file_path, 'r') as f:
                    bundle = json.load(f)
                    self.assets = bundle.get("assets", [])
                    self.owners = bundle.get("owners", {})
            else:
                messagebox.showerror("Error", "CSV import not yet implemented for complex data")
                return
            
            self.save_data()
            self.save_owners()
            self.refresh_tree()
            messagebox.showinfo("Success", "Data imported successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import data: {e}")

    def backup_data(self):
        """Create a backup of all data."""
        backup_file = f"afmdw_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            bundle = {
                "assets": self.assets,
                "owners": self.owners,
                "backup_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            with open(backup_file, 'w') as f:
                json.dump(bundle, f, indent=4)
            messagebox.showinfo("Success", f"Backup created: {backup_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create backup: {e}")

    def view_task_history(self):
        """View task history from database."""
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute("SELECT * FROM tasks ORDER BY created_date DESC LIMIT 50")
            tasks = c.fetchall()
            conn.close()
            
            history_window = tk.Toplevel(self.root)
            history_window.title("Task History")
            history_window.geometry("800x400")
            
            text = tk.Text(history_window, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            if tasks:
                for task in tasks:
                    text.insert(tk.END, f"Asset ID: {task[1]}\n")
                    text.insert(tk.END, f"Type: {task[2]}\n")
                    text.insert(tk.END, f"Description: {task[3]}\n")
                    text.insert(tk.END, f"Status: {task[4]}\n")
                    text.insert(tk.END, f"Created: {task[5]}\n")
                    text.insert(tk.END, f"Due: {task[6]}\n")
                    text.insert(tk.END, "-" * 40 + "\n\n")
            else:
                text.insert(tk.END, "No task history found.")
            
            text.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to view task history: {e}")

    def view_query_history(self):
        """View query history from database."""
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute("SELECT * FROM query_history ORDER BY created_date DESC LIMIT 50")
            queries = c.fetchall()
            conn.close()
            
            history_window = tk.Toplevel(self.root)
            history_window.title("Query History")
            history_window.geometry("800x400")
            
            text = tk.Text(history_window, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            if queries:
                for query in queries:
                    text.insert(tk.END, f"Query Type: {query[1]}\n")
                    text.insert(tk.END, f"Parameters: {query[2]}\n")
                    text.insert(tk.END, f"Created: {query[3]}\n")
                    text.insert(tk.END, f"Results: {query[4]}\n")
                    text.insert(tk.END, "-" * 40 + "\n\n")
            else:
                text.insert(tk.END, "No query history found.")
            
            text.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to view query history: {e}")

    def log_query(self, query_type, params, results_count):
        """Log a query to the database."""
        try:
            conn = sqlite3.connect(self.db_file)
            c = conn.cursor()
            c.execute('''INSERT INTO query_history (query_type, query_params, created_date, results)
                         VALUES (?, ?, ?, ?)''',
                      (query_type, json.dumps(params), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(results_count)))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Failed to log query: {e}")

    def clear_database(self):
        """Clear the database (with confirmation)."""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear the entire database? This cannot be undone."):
            try:
                conn = sqlite3.connect(self.db_file)
                c = conn.cursor()
                c.execute("DELETE FROM tasks")
                c.execute("DELETE FROM query_history")
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", "Database cleared successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to clear database: {e}")


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = AFMDWTaskManagementSystem(root)
    root.mainloop()


if __name__ == "__main__":
    main()
