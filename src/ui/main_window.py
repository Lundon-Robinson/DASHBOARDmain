"""
Main Window UI
==============

Advanced multi-tabbed main window interface with dashboard,
Excel management, email tools, and script runner.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Any, List, Optional
import json

# Data processing
import pandas as pd
try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False

# Try to import matplotlib for charts
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

from ..core.logger import logger
from ..core.config import DashboardConfig
from ..modules.excel_handler import ExcelHandler
from ..modules.email_handler import EmailHandler, EmailRecipient
from ..modules.script_runner import ScriptRunner
from ..modules.ai_assistant import AIAssistant

class MainWindow:
    """Advanced main window with tabbed interface"""
    
    def __init__(self, config: DashboardConfig):
        self.config = config
        self.root = tk.Tk()
        self.root.title("Advanced Finance Dashboard")
        self.root.geometry(f"{config.ui.window_width}x{config.ui.window_height}")
        
        # Initialize handlers
        self.excel_handler = ExcelHandler(config)
        self.email_handler = EmailHandler(config)
        self.script_runner = ScriptRunner(config)
        self.ai_assistant = AIAssistant(config) if config.ai.openai_api_key else None
        
        # UI Components
        self.notebook = None
        self.status_bar = None
        self.command_queue = queue.Queue()
        
        # Data
        self.cardholders = []
        self.current_data = None
        
        # Initialize UI
        self._setup_styles()
        self._create_menu()
        self._create_main_interface()
        self._create_status_bar()
        self._start_background_tasks()
        
        # Load initial data
        self._load_initial_data()
        
        logger.info("Main window initialized")
    
    def _setup_styles(self):
        """Setup UI styling"""
        style = ttk.Style()
        
        if self.config.ui.theme == "dark":
            # Dark theme configuration
            style.theme_use("clam")
            
            # Configure dark colors
            style.configure(".", background="#2b2b2b", foreground="#ffffff", 
                          fieldbackground="#1e1e1e", bordercolor="#555555")
            style.configure("TNotebook", background="#2b2b2b", borderwidth=0)
            style.configure("TNotebook.Tab", background="#3c3c3c", foreground="#ffffff",
                          padding=[12, 8])
            style.map("TNotebook.Tab", background=[("selected", "#0078d4")])
            style.configure("TFrame", background="#2b2b2b")
            style.configure("TLabel", background="#2b2b2b", foreground="#ffffff")
            style.configure("TLabelFrame", background="#2b2b2b", foreground="#ffffff")
            style.configure("TLabelFrame.Label", background="#2b2b2b", foreground="#ffffff")
            style.configure("TButton", background="#0078d4", foreground="#ffffff")
            style.map("TButton", background=[("active", "#106ebe")])
            style.configure("TEntry", fieldbackground="#1e1e1e", foreground="#ffffff", bordercolor="#555555")
            style.configure("TCombobox", fieldbackground="#1e1e1e", foreground="#ffffff", bordercolor="#555555")
            style.map("TEntry", focuscolor=[("!focus", "#555555")])
            style.map("TCombobox", focuscolor=[("!focus", "#555555")])
            
            # Treeview styling
            style.configure("Treeview", background="#1e1e1e", foreground="#ffffff", fieldbackground="#1e1e1e")
            style.configure("Treeview.Heading", background="#3c3c3c", foreground="#ffffff")
            style.map("Treeview", background=[("selected", "#0078d4")])
            
            # Configure root window
            self.root.configure(bg="#2b2b2b")
        else:
            style.theme_use("default")
    
    def _create_menu(self):
        """Create application menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Excel...", command=self._import_excel)
        file_menu.add_command(label="Export Data...", command=self._export_data)
        file_menu.add_separator()
        file_menu.add_command(label="Settings", command=self._show_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Run Script...", command=self._run_script_dialog)
        tools_menu.add_command(label="Send Test Email", command=self._send_test_email)
        tools_menu.add_command(label="Repair Legacy Scripts", command=self._repair_scripts)
        tools_menu.add_separator()
        tools_menu.add_command(label="System Status", command=self._show_system_status)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)
        help_menu.add_command(label="User Guide", command=self._show_user_guide)
    
    def _create_main_interface(self):
        """Create main tabbed interface"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=(5, 0))
        
        # Create tabs
        self._create_dashboard_tab()
        self._create_cardholders_tab()
        self._create_statements_tab()
        self._create_email_tab()
        self._create_scripts_tab()
        self._create_analytics_tab()
        if self.ai_assistant:
            self._create_ai_tab()
    
    def _create_dashboard_tab(self):
        """Create main dashboard tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Dashboard")
        
        # Create dashboard layout
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # KPI Section
        kpi_frame = ttk.LabelFrame(main_frame, text="Key Performance Indicators", padding=10)
        kpi_frame.pack(fill="x", pady=(0, 10))
        
        # Create KPI cards
        kpi_cards_frame = ttk.Frame(kpi_frame)
        kpi_cards_frame.pack(fill="x")
        
        # KPI Cards
        self.kpi_cards = {}
        kpi_data = [
            ("Total Cardholders", "0", "#0078d4"),
            ("Active Statements", "0", "#107c10"),
            ("Pending Emails", "0", "#ff8c00"),
            ("Running Scripts", "0", "#5c2d91")
        ]
        
        for i, (title, value, color) in enumerate(kpi_data):
            card = self._create_kpi_card(kpi_cards_frame, title, value, color)
            card.grid(row=0, column=i, padx=5, sticky="ew")
            self.kpi_cards[title] = card
            kpi_cards_frame.columnconfigure(i, weight=1)
        
        # Quick Actions Section
        actions_frame = ttk.LabelFrame(main_frame, text="Quick Actions", padding=10)
        actions_frame.pack(fill="x", pady=(0, 10))
        
        actions_grid = ttk.Frame(actions_frame)
        actions_grid.pack(fill="x")
        
        actions = [
            ("Generate Statements", self._quick_generate_statements, 0, 0),
            ("Send Bulk Email", self._quick_send_email, 0, 1),
            ("Run Legacy Script", self._quick_run_script, 0, 2),
            ("Import Excel Data", self._quick_import_excel, 1, 0),
            ("View System Logs", self._quick_view_logs, 1, 1),
            ("Export Reports", self._quick_export_reports, 1, 2)
        ]
        
        for text, command, row, col in actions:
            btn = ttk.Button(actions_grid, text=text, command=command)
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            actions_grid.columnconfigure(col, weight=1)
        
        # Recent Activity Section
        activity_frame = ttk.LabelFrame(main_frame, text="Recent Activity", padding=10)
        activity_frame.pack(fill="both", expand=True)
        
        # Activity listbox with scrollbar
        activity_list_frame = ttk.Frame(activity_frame)
        activity_list_frame.pack(fill="both", expand=True)
        
        self.activity_listbox = tk.Listbox(
            activity_list_frame,
            font=(self.config.ui.font_family, self.config.ui.font_size),
            bg="#1e1e1e" if self.config.ui.theme == "dark" else "#ffffff",
            fg="#ffffff" if self.config.ui.theme == "dark" else "#000000",
            selectbackground="#0078d4"
        )
        scrollbar = ttk.Scrollbar(activity_list_frame, orient="vertical", command=self.activity_listbox.yview)
        self.activity_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.activity_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add some sample activities
        self._add_activity("Dashboard initialized")
    
    def _create_kpi_card(self, parent, title, value, color):
        """Create a KPI card widget"""
        card_frame = ttk.Frame(parent)
        
        # Title
        title_label = ttk.Label(card_frame, text=title, font=(self.config.ui.font_family, 10, "bold"))
        title_label.pack()
        
        # Value
        value_label = ttk.Label(
            card_frame, 
            text=value, 
            font=(self.config.ui.font_family, 24, "bold"),
            foreground=color
        )
        value_label.pack()
        
        # Store reference to value label for updates
        card_frame.value_label = value_label
        
        return card_frame
    
    def _create_cardholders_tab(self):
        """Create cardholders management tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Cardholders")
        
        # Main container
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Toolbar
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill="x", pady=(0, 10))
        
        ttk.Button(toolbar, text="Add Cardholder", command=self._add_cardholder).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="Edit Selected", command=self._edit_cardholder).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="Delete Selected", command=self._delete_cardholder).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="Import from Excel", command=self._import_cardholders).pack(side="left", padx=(0, 5))
        ttk.Button(toolbar, text="Export to Excel", command=self._export_cardholders).pack(side="left", padx=(0, 5))
        
        # Search frame
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(search_frame, text="Search:").pack(side="left")
        self.cardholder_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.cardholder_search_var, width=30)
        search_entry.pack(side="left", padx=(5, 10))
        search_entry.bind("<KeyRelease>", self._filter_cardholders)
        
        # Cardholders treeview
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # Create treeview
        columns = ("Name", "Email", "Card Number", "Department", "Status")
        self.cardholders_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        # Configure columns
        for col in columns:
            self.cardholders_tree.heading(col, text=col)
            self.cardholders_tree.column(col, width=150)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.cardholders_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.cardholders_tree.xview)
        self.cardholders_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.cardholders_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
    
    def _create_statements_tab(self):
        """Create statements generation tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Statements")
        
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Generation settings
        settings_frame = ttk.LabelFrame(main_frame, text="Statement Generation Settings", padding=10)
        settings_frame.pack(fill="x", pady=(0, 10))
        
        # Period selection
        period_frame = ttk.Frame(settings_frame)
        period_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(period_frame, text="Period:").pack(side="left")
        self.statement_month_var = tk.StringVar(value=datetime.now().strftime("%B"))
        self.statement_year_var = tk.StringVar(value=str(datetime.now().year))
        
        month_combo = ttk.Combobox(period_frame, textvariable=self.statement_month_var, width=10)
        month_combo['values'] = ('January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December')
        month_combo.pack(side="left", padx=(5, 5))
        
        year_combo = ttk.Combobox(period_frame, textvariable=self.statement_year_var, width=8)
        year_combo['values'] = tuple(str(year) for year in range(2020, 2030))
        year_combo.pack(side="left", padx=(0, 10))
        
        # Generation buttons
        buttons_frame = ttk.Frame(settings_frame)
        buttons_frame.pack(fill="x")
        
        ttk.Button(buttons_frame, text="Generate All Statements", 
                  command=self._generate_all_statements).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_frame, text="Generate Selected", 
                  command=self._generate_selected_statements).pack(side="left", padx=(0, 10))
        ttk.Button(buttons_frame, text="Preview Statement", 
                  command=self._preview_statement).pack(side="left")
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Generation Progress", padding=10)
        progress_frame.pack(fill="x", pady=(0, 10))
        
        self.statement_progress_var = tk.StringVar(value="Ready to generate statements")
        ttk.Label(progress_frame, textvariable=self.statement_progress_var).pack(anchor="w")
        
        self.statement_progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.statement_progress_bar.pack(fill="x", pady=(5, 0))
        
        # Generated statements list
        list_frame = ttk.LabelFrame(main_frame, text="Generated Statements", padding=10)
        list_frame.pack(fill="both", expand=True)
        
        # Statements treeview
        statements_tree_frame = ttk.Frame(list_frame)
        statements_tree_frame.pack(fill="both", expand=True)
        
        columns = ("Cardholder", "Period", "Amount", "File", "Generated")
        self.statements_tree = ttk.Treeview(statements_tree_frame, columns=columns, show="headings")
        
        for col in columns:
            self.statements_tree.heading(col, text=col)
            self.statements_tree.column(col, width=120)
        
        # Scrollbar for statements
        statements_scrollbar = ttk.Scrollbar(statements_tree_frame, orient="vertical", 
                                           command=self.statements_tree.yview)
        self.statements_tree.configure(yscrollcommand=statements_scrollbar.set)
        
        self.statements_tree.pack(side="left", fill="both", expand=True)
        statements_scrollbar.pack(side="right", fill="y")
        
        # Context menu for statements
        self.statements_tree.bind("<Button-3>", self._show_statement_context_menu)
    
    def _create_email_tab(self):
        """Create email management tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Email")
        
        # Create paned window for layout
        paned = ttk.PanedWindow(frame, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left panel - Templates and recipients
        left_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=1)
        
        # Templates section
        templates_frame = ttk.LabelFrame(left_panel, text="Email Templates", padding=10)
        templates_frame.pack(fill="x", pady=(0, 10))
        
        self.template_var = tk.StringVar()
        template_combo = ttk.Combobox(templates_frame, textvariable=self.template_var, width=30)
        template_combo.pack(fill="x", pady=(0, 10))
        
        template_buttons = ttk.Frame(templates_frame)
        template_buttons.pack(fill="x")
        ttk.Button(template_buttons, text="Load", command=self._load_email_template).pack(side="left", padx=(0, 5))
        ttk.Button(template_buttons, text="Save", command=self._save_email_template).pack(side="left", padx=(0, 5))
        ttk.Button(template_buttons, text="New", command=self._new_email_template).pack(side="left")
        
        # Recipients section
        recipients_frame = ttk.LabelFrame(left_panel, text="Recipients", padding=10)
        recipients_frame.pack(fill="both", expand=True)
        
        # Recipients input
        ttk.Label(recipients_frame, text="Email addresses (one per line):").pack(anchor="w")
        self.recipients_text = scrolledtext.ScrolledText(
            recipients_frame, 
            height=8, 
            width=40,
            bg="#1e1e1e" if self.config.ui.theme == "dark" else "#ffffff",
            fg="#ffffff" if self.config.ui.theme == "dark" else "#000000",
            insertbackground="#ffffff" if self.config.ui.theme == "dark" else "#000000"
        )
        self.recipients_text.pack(fill="both", expand=True, pady=(5, 10))
        
        recipient_buttons = ttk.Frame(recipients_frame)
        recipient_buttons.pack(fill="x")
        ttk.Button(recipient_buttons, text="Import from Excel", 
                  command=self._import_email_recipients).pack(side="left", padx=(0, 5))
        ttk.Button(recipient_buttons, text="Add Cardholders", 
                  command=self._add_cardholders_as_recipients).pack(side="left")
        
        # Right panel - Email composition
        right_panel = ttk.Frame(paned)
        paned.add(right_panel, weight=2)
        
        # Email composition
        compose_frame = ttk.LabelFrame(right_panel, text="Compose Email", padding=10)
        compose_frame.pack(fill="both", expand=True)
        
        # Subject
        subject_frame = ttk.Frame(compose_frame)
        subject_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(subject_frame, text="Subject:").pack(side="left")
        self.email_subject_var = tk.StringVar()
        ttk.Entry(subject_frame, textvariable=self.email_subject_var).pack(side="left", fill="x", expand=True, padx=(10, 0))
        
        # Body
        ttk.Label(compose_frame, text="Body:").pack(anchor="w")
        self.email_body_text = scrolledtext.ScrolledText(
            compose_frame, 
            height=15,
            bg="#1e1e1e" if self.config.ui.theme == "dark" else "#ffffff",
            fg="#ffffff" if self.config.ui.theme == "dark" else "#000000",
            insertbackground="#ffffff" if self.config.ui.theme == "dark" else "#000000"
        )
        self.email_body_text.pack(fill="both", expand=True, pady=(5, 10))
        
        # Attachments
        attachments_frame = ttk.Frame(compose_frame)
        attachments_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(attachments_frame, text="Attachments:").pack(side="left")
        self.attachments_var = tk.StringVar()
        ttk.Entry(attachments_frame, textvariable=self.attachments_var).pack(side="left", fill="x", expand=True, padx=(10, 5))
        ttk.Button(attachments_frame, text="Browse", command=self._browse_attachments).pack(side="left")
        
        # Send buttons
        send_frame = ttk.Frame(compose_frame)
        send_frame.pack(fill="x")
        ttk.Button(send_frame, text="Send Test Email", command=self._send_test_email_tab).pack(side="left", padx=(0, 10))
        ttk.Button(send_frame, text="Send to All Recipients", command=self._send_bulk_email).pack(side="left", padx=(0, 10))
        ttk.Button(send_frame, text="Preview Variables", command=self._preview_email_variables).pack(side="left")
        
        # Load templates
        self._refresh_email_templates()
    
    def _create_scripts_tab(self):
        """Create script management tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Scripts")
        
        # Create paned window
        paned = ttk.PanedWindow(frame, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left panel - Scripts list
        left_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=1)
        
        scripts_frame = ttk.LabelFrame(left_panel, text="Available Scripts", padding=10)
        scripts_frame.pack(fill="both", expand=True)
        
        # Scripts treeview
        scripts_columns = ("Name", "Category", "Description")
        self.scripts_tree = ttk.Treeview(scripts_frame, columns=scripts_columns, show="headings", height=10)
        
        for col in scripts_columns:
            self.scripts_tree.heading(col, text=col)
            self.scripts_tree.column(col, width=120)
        
        self.scripts_tree.pack(fill="both", expand=True, pady=(0, 10))
        
        # Script buttons
        script_buttons = ttk.Frame(scripts_frame)
        script_buttons.pack(fill="x")
        ttk.Button(script_buttons, text="Run Selected", command=self._run_selected_script).pack(side="left", padx=(0, 5))
        ttk.Button(script_buttons, text="Stop", command=self._stop_selected_script).pack(side="left", padx=(0, 5))
        ttk.Button(script_buttons, text="Refresh", command=self._refresh_scripts_list).pack(side="left")
        
        # Right panel - Script output and running scripts
        right_panel = ttk.Frame(paned)
        paned.add(right_panel, weight=2)
        
        # Running scripts section
        running_frame = ttk.LabelFrame(right_panel, text="Running Scripts", padding=10)
        running_frame.pack(fill="x", pady=(0, 10))
        
        running_columns = ("Script", "Status", "Runtime", "PID")
        self.running_scripts_tree = ttk.Treeview(running_frame, columns=running_columns, show="headings", height=4)
        
        for col in running_columns:
            self.running_scripts_tree.heading(col, text=col)
            self.running_scripts_tree.column(col, width=100)
        
        self.running_scripts_tree.pack(fill="x")
        
        # Script output section
        output_frame = ttk.LabelFrame(right_panel, text="Script Output", padding=10)
        output_frame.pack(fill="both", expand=True)
        
        self.script_output_text = scrolledtext.ScrolledText(
            output_frame, 
            height=15,
            bg="#1e1e1e" if self.config.ui.theme == "dark" else "#ffffff",
            fg="#ffffff" if self.config.ui.theme == "dark" else "#000000",
            insertbackground="#ffffff" if self.config.ui.theme == "dark" else "#000000"
        )
        self.script_output_text.pack(fill="both", expand=True)
        
        # Populate scripts list
        self._refresh_scripts_list()
    
    def _create_analytics_tab(self):
        """Create analytics and charts tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Analytics")
        
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Chart controls
        controls_frame = ttk.LabelFrame(main_frame, text="Chart Controls", padding=10)
        controls_frame.pack(fill="x", pady=(0, 10))
        
        # Chart type selection
        chart_frame = ttk.Frame(controls_frame)
        chart_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(chart_frame, text="Chart Type:").pack(side="left")
        self.chart_type_var = tk.StringVar(value="transactions_by_month")
        chart_combo = ttk.Combobox(chart_frame, textvariable=self.chart_type_var, width=25)
        chart_combo['values'] = (
            'transactions_by_month',
            'spending_by_cardholder',
            'category_breakdown',
            'monthly_trends',
            'email_statistics'
        )
        chart_combo.pack(side="left", padx=(10, 10))
        ttk.Button(chart_frame, text="Generate Chart", command=self._generate_chart).pack(side="left")
        
        # Chart display area
        if MATPLOTLIB_AVAILABLE:
            self.chart_frame = ttk.LabelFrame(main_frame, text="Chart Display", padding=10)
            self.chart_frame.pack(fill="both", expand=True)
            
            # Placeholder for matplotlib canvas
            self.chart_canvas = None
        else:
            no_charts_frame = ttk.LabelFrame(main_frame, text="Chart Display", padding=10)
            no_charts_frame.pack(fill="both", expand=True)
            ttk.Label(no_charts_frame, 
                     text="Charts not available\\nInstall matplotlib for chart functionality",
                     font=(self.config.ui.font_family, 12)).pack(expand=True)
    
    def _create_ai_tab(self):
        """Create AI assistant tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="AI Assistant")
        
        main_frame = ttk.Frame(frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # AI Chat interface
        chat_frame = ttk.LabelFrame(main_frame, text="AI Command Interface", padding=10)
        chat_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Chat history
        self.ai_chat_text = scrolledtext.ScrolledText(
            chat_frame, 
            height=15,
            state="disabled",
            bg="#1e1e1e" if self.config.ui.theme == "dark" else "#ffffff",
            fg="#ffffff" if self.config.ui.theme == "dark" else "#000000"
        )
        self.ai_chat_text.pack(fill="both", expand=True, pady=(0, 10))
        
        # Input area
        input_frame = ttk.Frame(chat_frame)
        input_frame.pack(fill="x")
        
        ttk.Label(input_frame, text="Command:").pack(side="left")
        self.ai_input_var = tk.StringVar()
        ai_entry = ttk.Entry(input_frame, textvariable=self.ai_input_var)
        ai_entry.pack(side="left", fill="x", expand=True, padx=(10, 5))
        ai_entry.bind("<Return>", self._process_ai_command)
        ttk.Button(input_frame, text="Send", command=self._process_ai_command).pack(side="left")
        
        # Quick commands
        quick_commands_frame = ttk.LabelFrame(main_frame, text="Quick Commands", padding=10)
        quick_commands_frame.pack(fill="x")
        
        quick_commands = [
            ("System Status", "system status"),
            ("Analyze Logs", "analyze logs today"),
            ("Generate Statements", "generate statements for current month"),
            ("Email Statistics", "show email statistics")
        ]
        
        for i, (label, command) in enumerate(quick_commands):
            btn = ttk.Button(quick_commands_frame, text=label, 
                           command=lambda cmd=command: self._send_ai_command(cmd))
            btn.grid(row=i//2, column=i%2, padx=5, pady=5, sticky="ew")
        
        quick_commands_frame.columnconfigure(0, weight=1)
        quick_commands_frame.columnconfigure(1, weight=1)
        
        # Welcome message
        self._add_ai_message("AI Assistant", "Hello! I can help you with commands like:\\n• Generate statements for [period]\\n• Analyze logs\\n• System status\\n• Run script [name]")
    
    def _create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill="x", side="bottom")
        
        # Status sections
        self.status_text_var = tk.StringVar(value="Ready")
        ttk.Label(self.status_bar, textvariable=self.status_text_var).pack(side="left", padx=(10, 0))
        
        # Right side status items
        right_status = ttk.Frame(self.status_bar)
        right_status.pack(side="right", padx=(0, 10))
        
        # Time
        self.time_var = tk.StringVar()
        ttk.Label(right_status, textvariable=self.time_var).pack(side="right", padx=(10, 0))
        
        # Connection status
        self.connection_var = tk.StringVar(value="● Connected")
        ttk.Label(right_status, textvariable=self.connection_var, 
                 foreground="#107c10").pack(side="right", padx=(10, 0))
    
    def _start_background_tasks(self):
        """Start background tasks"""
        # Update time
        self._update_time()
        
        # Update status
        self._update_status()
        
        # Process command queue
        self.root.after(100, self._process_command_queue)
    
    def _update_time(self):
        """Update time display"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_var.set(current_time)
        self.root.after(1000, self._update_time)
    
    def _update_status(self):
        """Update various status indicators"""
        try:
            # Update KPIs
            cardholders_count = len(self.excel_handler.db_manager.get_cardholders())
            self._update_kpi("Total Cardholders", str(cardholders_count))
            
            # Update running scripts
            running_scripts = self.script_runner.get_running_scripts()
            self._update_kpi("Running Scripts", str(len(running_scripts)))
            
            # Update running scripts tree
            self._update_running_scripts_tree()
            
        except Exception as e:
            logger.error("Status update failed", exception=e)
        
        # Schedule next update
        self.root.after(5000, self._update_status)  # Update every 5 seconds
    
    def _process_command_queue(self):
        """Process queued commands"""
        try:
            while not self.command_queue.empty():
                command, args = self.command_queue.get_nowait()
                if hasattr(self, command):
                    method = getattr(self, command)
                    method(*args)
        except queue.Empty:
            pass
        except Exception as e:
            logger.error("Command queue processing failed", exception=e)
        
        # Schedule next check
        self.root.after(100, self._process_command_queue)
    
    def _load_initial_data(self):
        """Load initial data"""
        try:
            # Load cardholders
            self.cardholders = self.excel_handler.db_manager.get_cardholders()
            self._refresh_cardholders_tree()
            
            # Add initial activity
            self._add_activity(f"Loaded {len(self.cardholders)} cardholders")
            
        except Exception as e:
            logger.error("Failed to load initial data", exception=e)
            self._add_activity(f"Error loading data: {e}")
    
    def _add_activity(self, message):
        """Add activity to the recent activity list"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        activity_item = f"[{timestamp}] {message}"
        
        self.activity_listbox.insert(0, activity_item)
        
        # Keep only last 50 items
        if self.activity_listbox.size() > 50:
            self.activity_listbox.delete(50, tk.END)
    
    def _update_kpi(self, title, value):
        """Update KPI card value"""
        if title in self.kpi_cards:
            self.kpi_cards[title].value_label.config(text=value)
    
    def _refresh_cardholders_tree(self):
        """Refresh cardholders treeview"""
        # Clear existing items
        for item in self.cardholders_tree.get_children():
            self.cardholders_tree.delete(item)
        
        # Add cardholders
        for cardholder in self.cardholders:
            status = "Active" if cardholder.active else "Inactive"
            self.cardholders_tree.insert("", tk.END, values=(
                cardholder.name,
                cardholder.email,
                cardholder.card_number,
                cardholder.department or "",
                status
            ))
    
    def run(self):
        """Run the main window"""
        try:
            self.root.mainloop()
        except Exception as e:
            logger.error("Main window error", exception=e)
            raise
        finally:
            # Cleanup
            if hasattr(self, 'script_runner'):
                self.script_runner.cleanup()
    
    # Placeholder methods for functionality - these would be implemented with full features
    def _import_excel(self): 
        """General Excel import functionality"""
        try:
            self._import_cardholders()
        except Exception as e:
            self._add_activity(f"Excel import error: {str(e)}")
            logger.error("Excel import failed", exception=e)
    def _export_data(self): 
        """General data export functionality"""
        try:
            self._export_cardholders()
        except Exception as e:
            self._add_activity(f"Data export error: {str(e)}")
            logger.error("Data export failed", exception=e) 
    def _show_settings(self): messagebox.showinfo("Info", "Settings dialog")
    def _run_script_dialog(self): messagebox.showinfo("Info", "Script runner dialog")
    def _send_test_email(self): messagebox.showinfo("Info", "Test email functionality")
    def _repair_scripts(self): messagebox.showinfo("Info", "Script repair functionality")
    def _show_system_status(self): messagebox.showinfo("Info", "System status dialog")
    def _show_about(self): messagebox.showinfo("About", f"Advanced Finance Dashboard v{self.config.version}")
    def _show_user_guide(self): messagebox.showinfo("Info", "User guide functionality")
    
    # Quick action methods
    def _quick_generate_statements(self): self._add_activity("Quick: Generate statements requested")
    def _quick_send_email(self): self._add_activity("Quick: Send email requested")
    def _quick_run_script(self): self._add_activity("Quick: Run script requested")
    def _quick_import_excel(self): 
        """Quick action to import Excel file"""
        try:
            self._import_cardholders()
        except Exception as e:
            self._add_activity(f"Quick import error: {str(e)}")
            logger.error("Quick import failed", exception=e)
    def _quick_view_logs(self): self._add_activity("Quick: View logs requested")
    def _quick_export_reports(self): self._add_activity("Quick: Export reports requested")
    
    # Cardholder management methods
    def _add_cardholder(self):
        """Add new cardholder via dialog"""
        try:
            # Create dialog window
            dialog = tk.Toplevel(self.root)
            dialog.title("Add New Cardholder")
            dialog.geometry("400x300")
            dialog.grab_set()  # Make dialog modal
            
            # Create form fields
            fields = {}
            labels = ['Name', 'Email', 'Card Number', 'Department', 'Cost Centre', 'Manager Email']
            
            for i, label in enumerate(labels):
                tk.Label(dialog, text=f"{label}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
                entry = tk.Entry(dialog, width=30)
                entry.grid(row=i, column=1, padx=10, pady=5)
                fields[label.lower().replace(' ', '_')] = entry
            
            # Buttons
            button_frame = tk.Frame(dialog)
            button_frame.grid(row=len(labels), column=0, columnspan=2, pady=20)
            
            def save_cardholder():
                try:
                    # Validate required fields
                    name = fields['name'].get().strip()
                    email = fields['email'].get().strip()
                    card_number = fields['card_number'].get().strip()
                    
                    if not name or not email or not card_number:
                        messagebox.showerror("Validation Error", "Name, Email, and Card Number are required.")
                        return
                    
                    # Create cardholder
                    cardholder = self.excel_handler.db_manager.create_cardholder(
                        card_number=card_number,
                        name=name,
                        email=email,
                        manager_email=fields['manager_email'].get().strip() or None,
                        department=fields['department'].get().strip() or None,
                        cost_centre=fields['cost_centre'].get().strip() or None
                    )
                    
                    # Refresh UI
                    self._load_initial_data()
                    self._add_activity(f"Added new cardholder: {name}")
                    
                    dialog.destroy()
                    messagebox.showinfo("Success", f"Successfully added cardholder: {name}")
                    
                except Exception as e:
                    error_msg = f"Failed to add cardholder: {str(e)}"
                    messagebox.showerror("Error", error_msg)
            
            tk.Button(button_frame, text="Save", command=save_cardholder).pack(side="left", padx=5)
            tk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)
            
        except Exception as e:
            error_msg = f"Failed to open add cardholder dialog: {str(e)}"
            logger.error("Add cardholder dialog failed", exception=e)
            messagebox.showerror("Error", error_msg)
    def _edit_cardholder(self):
        """Edit selected cardholder"""
        try:
            selected_items = self.cardholders_tree.selection()
            if not selected_items:
                messagebox.showinfo("No Selection", "Please select a cardholder to edit.")
                return
            
            # Get selected cardholder details
            item = selected_items[0]
            values = self.cardholders_tree.item(item, 'values')
            cardholder_name = values[0]
            card_number = values[2]
            
            # Find cardholder in database
            cardholder = self.excel_handler.db_manager.get_cardholder_by_card_number(card_number)
            if not cardholder:
                messagebox.showerror("Error", "Cardholder not found in database.")
                return
            
            # Create edit dialog
            dialog = tk.Toplevel(self.root)
            dialog.title(f"Edit Cardholder: {cardholder_name}")
            dialog.geometry("400x300")
            dialog.grab_set()
            
            # Create form fields with existing data
            fields = {}
            field_data = [
                ('Name', cardholder.name),
                ('Email', cardholder.email),
                ('Card Number', cardholder.card_number),
                ('Department', cardholder.department or ''),
                ('Cost Centre', cardholder.cost_centre or ''),
                ('Manager Email', cardholder.manager_email or '')
            ]
            
            for i, (label, value) in enumerate(field_data):
                tk.Label(dialog, text=f"{label}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
                entry = tk.Entry(dialog, width=30)
                entry.grid(row=i, column=1, padx=10, pady=5)
                entry.insert(0, value)  # Pre-fill with existing data
                fields[label.lower().replace(' ', '_')] = entry
            
            # Active checkbox
            active_var = tk.BooleanVar(value=cardholder.active)
            active_cb = tk.Checkbutton(dialog, text="Active", variable=active_var)
            active_cb.grid(row=len(field_data), column=0, columnspan=2, pady=10)
            
            # Buttons
            button_frame = tk.Frame(dialog)
            button_frame.grid(row=len(field_data)+1, column=0, columnspan=2, pady=20)
            
            def save_changes():
                try:
                    # Validate required fields
                    name = fields['name'].get().strip()
                    email = fields['email'].get().strip()
                    new_card_number = fields['card_number'].get().strip()
                    
                    if not name or not email or not new_card_number:
                        messagebox.showerror("Validation Error", "Name, Email, and Card Number are required.")
                        return
                    
                    # Update cardholder in database
                    with self.excel_handler.db_manager.get_session() as session:
                        cardholder.name = name
                        cardholder.email = email
                        cardholder.card_number = new_card_number
                        cardholder.department = fields['department'].get().strip() or None
                        cardholder.cost_centre = fields['cost_centre'].get().strip() or None
                        cardholder.manager_email = fields['manager_email'].get().strip() or None
                        cardholder.active = active_var.get()
                        cardholder.updated_at = datetime.utcnow()
                        session.commit()
                    
                    # Refresh UI
                    self._load_initial_data()
                    self._add_activity(f"Updated cardholder: {name}")
                    
                    dialog.destroy()
                    messagebox.showinfo("Success", f"Successfully updated cardholder: {name}")
                    
                except Exception as e:
                    error_msg = f"Failed to update cardholder: {str(e)}"
                    messagebox.showerror("Error", error_msg)
            
            tk.Button(button_frame, text="Save Changes", command=save_changes).pack(side="left", padx=5)
            tk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)
            
        except Exception as e:
            error_msg = f"Failed to open edit cardholder dialog: {str(e)}"
            logger.error("Edit cardholder dialog failed", exception=e)
            self._add_activity(f"Error: {error_msg}")
            messagebox.showerror("Error", error_msg)
    def _delete_cardholder(self):
        """Delete selected cardholder"""
        try:
            selected_items = self.cardholders_tree.selection()
            if not selected_items:
                messagebox.showinfo("No Selection", "Please select a cardholder to delete.")
                return
            
            # Get selected cardholder details
            item = selected_items[0]
            values = self.cardholders_tree.item(item, 'values')
            cardholder_name = values[0]
            card_number = values[2]
            
            # Confirm deletion
            if not messagebox.askyesno("Confirm Delete", 
                f"Are you sure you want to delete cardholder '{cardholder_name}'?\n\nThis action cannot be undone."):
                return
            
            # Find and delete from database
            cardholder = self.excel_handler.db_manager.get_cardholder_by_card_number(card_number)
            if cardholder:
                with self.excel_handler.db_manager.get_session() as session:
                    session.delete(cardholder)
                    session.commit()
                
                # Refresh UI
                self._load_initial_data()
                self._add_activity(f"Deleted cardholder: {cardholder_name}")
                messagebox.showinfo("Success", f"Successfully deleted cardholder: {cardholder_name}")
            else:
                messagebox.showerror("Error", "Cardholder not found in database.")
                
        except Exception as e:
            error_msg = f"Failed to delete cardholder: {str(e)}"
            logger.error("Delete cardholder failed", exception=e)
            self._add_activity(f"Error: {error_msg}")
            messagebox.showerror("Delete Error", error_msg)
    def _import_cardholders(self):
        """Import cardholders from Excel file"""
        try:
            # Open file dialog
            file_path = filedialog.askopenfilename(
                title="Select Cardholder Excel File",
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            
            if not file_path:
                return  # User cancelled
            
            # Show progress indicator
            self._add_activity("Importing cardholders from Excel...")
            self.root.update()
            
            # Load and process the data
            df = self.excel_handler.load_cardholder_data(file_path)
            
            # Sync with database
            self.excel_handler._sync_cardholders(df)
            
            # Refresh the cardholder tree view
            self._load_initial_data()  # This will reload cardholders from database
            
            self._add_activity(f"Successfully imported {len(df)} cardholders")
            messagebox.showinfo("Success", f"Successfully imported {len(df)} cardholders from {Path(file_path).name}")
            
        except Exception as e:
            error_msg = f"Failed to import cardholders: {str(e)}"
            logger.error("Cardholder import failed", exception=e)
            self._add_activity(f"Error: {error_msg}")
            messagebox.showerror("Import Error", error_msg)
    def _export_cardholders(self):
        """Export cardholders to Excel file"""
        try:
            if not self.cardholders:
                messagebox.showinfo("No Data", "No cardholders to export.")
                return
            
            # Open save file dialog
            file_path = filedialog.asksaveasfilename(
                title="Save Cardholders Export",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ]
            )
            
            if not file_path:
                return  # User cancelled
            
            self._add_activity("Exporting cardholders...")
            self.root.update()
            
            # Prepare data for export
            export_data = []
            for cardholder in self.cardholders:
                export_data.append({
                    'Name': cardholder.name,
                    'Email': cardholder.email,
                    'Card Number': cardholder.card_number,
                    'Department': cardholder.department or '',
                    'Cost Centre': cardholder.cost_centre or '',
                    'Manager Email': cardholder.manager_email or '',
                    'Active': 'Yes' if cardholder.active else 'No',
                    'Created': cardholder.created_at.strftime('%Y-%m-%d %H:%M:%S'),
                    'Updated': cardholder.updated_at.strftime('%Y-%m-%d %H:%M:%S')
                })
            
            df = pd.DataFrame(export_data)
            
            # Export based on file extension
            if file_path.lower().endswith('.csv'):
                df.to_csv(file_path, index=False)
            else:
                df.to_excel(file_path, index=False, sheet_name='Cardholders')
            
            self._add_activity(f"Successfully exported {len(export_data)} cardholders")
            messagebox.showinfo("Success", f"Successfully exported {len(export_data)} cardholders to {Path(file_path).name}")
            
        except Exception as e:
            error_msg = f"Failed to export cardholders: {str(e)}"
            logger.error("Cardholder export failed", exception=e)
            self._add_activity(f"Error: {error_msg}")
            messagebox.showerror("Export Error", error_msg)
    def _filter_cardholders(self, event):
        """Filter cardholders based on search term"""
        try:
            search_widget = event.widget
            search_term = search_widget.get().lower().strip()
            
            # Clear existing items
            for item in self.cardholders_tree.get_children():
                self.cardholders_tree.delete(item)
            
            # Add filtered cardholders
            for cardholder in self.cardholders:
                # Search in name, email, card number, and department
                searchable_text = f"{cardholder.name} {cardholder.email} {cardholder.card_number} {cardholder.department or ''}".lower()
                
                if not search_term or search_term in searchable_text:
                    status = "Active" if cardholder.active else "Inactive"
                    self.cardholders_tree.insert("", tk.END, values=(
                        cardholder.name,
                        cardholder.email,
                        cardholder.card_number,
                        cardholder.department or "",
                        status
                    ))
                    
        except Exception as e:
            logger.error("Cardholder filtering failed", exception=e)
    
    # Statement methods
    def _generate_all_statements(self): self._add_activity("Generate all statements requested")
    def _generate_selected_statements(self): self._add_activity("Generate selected statements requested")
    def _preview_statement(self): messagebox.showinfo("Info", "Statement preview functionality")
    def _show_statement_context_menu(self, event): pass
    
    # Email methods
    def _load_email_template(self): 
        """Load selected email template"""
        template_name = self.template_var.get()
        if not template_name:
            messagebox.showwarning("Warning", "Please select a template to load")
            return
        
        try:
            template = self.email_handler.load_template(template_name)
            if template:
                self.email_subject_var.set(template.subject)
                self.email_body_text.delete("1.0", tk.END)
                self.email_body_text.insert("1.0", template.body)
                self._add_activity(f"Loaded template: {template_name}")
            else:
                messagebox.showwarning("Warning", f"Template '{template_name}' not found")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load template: {str(e)}")
    
    def _save_email_template(self): 
        """Save current email as template"""
        template_name = self.template_var.get()
        if not template_name:
            template_name = tk.simpledialog.askstring("Save Template", "Enter template name:")
            if not template_name:
                return
        
        try:
            subject = self.email_subject_var.get()
            body = self.email_body_text.get("1.0", tk.END).strip()
            
            if not subject or not body:
                messagebox.showwarning("Warning", "Subject and body are required")
                return
            
            template = self.email_handler.create_template(template_name, subject, body)
            self._refresh_email_templates()
            self.template_var.set(template_name)
            self._add_activity(f"Saved template: {template_name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save template: {str(e)}")
    
    def _new_email_template(self): 
        """Create new email template"""
        self.template_var.set("")
        self.email_subject_var.set("")
        self.email_body_text.delete("1.0", tk.END)
        self._add_activity("New email template created")
    
    def _import_email_recipients(self): 
        """Import email recipients from Excel file"""
        try:
            # Use the Excel handler to load recipients from OUTSTANDING LOGS sheet
            filepath = filedialog.askopenfilename(
                title="Select Purchase Cardholder List",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            
            if not filepath:
                return
            
            # Load data using the existing Excel handler logic
            recipients = self._load_recipients_from_excel(filepath)
            
            if recipients:
                # Clear existing recipients
                self.recipients_text.delete("1.0", tk.END)
                
                # Add recipients to text area
                recipient_lines = []
                for recipient in recipients:
                    name = recipient.get('name', 'Unknown')
                    email = recipient.get('email', '')
                    if email:
                        recipient_lines.append(f"{name} <{email}>")
                
                self.recipients_text.insert("1.0", "\\n".join(recipient_lines))
                self._add_activity(f"Imported {len(recipients)} recipients from Excel")
                messagebox.showinfo("Success", f"Imported {len(recipients)} recipients from Excel file")
            else:
                messagebox.showwarning("Warning", "No valid recipients found in the Excel file")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import recipients: {str(e)}")
            logger.error("Import recipients failed", exception=e)
    
    def _load_recipients_from_excel(self, filepath):
        """Load recipients from Excel file OUTSTANDING LOGS sheet"""
        recipients = []
        try:
            import pandas as pd
            
            # Read the OUTSTANDING LOGS sheet
            df = pd.read_excel(filepath, sheet_name="OUTSTANDING LOGS")
            
            # Common column mappings for cardholder data
            name_columns = ['Name', 'Full Name', 'Cardholder Name', 'FullName']
            email_columns = ['Email', 'Email Address', 'Cardholder Email']
            
            # Find the actual column names
            name_col = None
            email_col = None
            
            for col in df.columns:
                col_str = str(col).strip()
                if col_str in name_columns:
                    name_col = col
                if col_str in email_columns:
                    email_col = col
            
            # If standard columns not found, try positional mapping
            if not name_col and len(df.columns) > 4:
                name_col = df.columns[4]  # Commonly position 5 (index 4)
            if not email_col and len(df.columns) > 7:
                email_col = df.columns[7]  # Commonly position 8 (index 7)
            
            if name_col is not None and email_col is not None:
                for _, row in df.iterrows():
                    name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
                    email_data = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
                    
                    # Skip empty rows
                    if not name or not email_data:
                        continue
                    
                    # Parse email data (may contain multiple emails separated by semicolons)
                    emails = [e.strip() for e in email_data.split(';') if e.strip()]
                    
                    for email in emails:
                        # Basic email validation
                        if '@' in email and '.' in email.split('@')[1]:
                            recipients.append({
                                'name': name,
                                'email': email
                            })
            else:
                logger.warning(f"Could not find name or email columns in {filepath}")
                
        except Exception as e:
            logger.error("Failed to load recipients from Excel", exception=e)
            raise
        
        return recipients
    
    def _add_cardholders_as_recipients(self): 
        """Add current cardholders as email recipients"""
        try:
            if not self.cardholders:
                messagebox.showinfo("Info", "No cardholders loaded. Please import cardholder data first.")
                return
            
            # Clear existing recipients
            self.recipients_text.delete("1.0", tk.END)
            
            # Add cardholders with valid emails
            recipient_lines = []
            for cardholder in self.cardholders:
                if hasattr(cardholder, 'email') and cardholder.email:
                    name = cardholder.name if hasattr(cardholder, 'name') else 'Unknown'
                    recipient_lines.append(f"{name} <{cardholder.email}>")
            
            if recipient_lines:
                self.recipients_text.insert("1.0", "\\n".join(recipient_lines))
                self._add_activity(f"Added {len(recipient_lines)} cardholders as recipients")
            else:
                messagebox.showinfo("Info", "No cardholders with valid email addresses found")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add cardholders: {str(e)}")
    
    def _browse_attachments(self): 
        """Browse for email attachments"""
        filetypes = [
            ("All files", "*.*"),
            ("PDF files", "*.pdf"),
            ("Excel files", "*.xlsx *.xls"),
            ("Images", "*.png *.jpg *.jpeg *.gif")
        ]
        
        filenames = filedialog.askopenfilenames(
            title="Select Attachments",
            filetypes=filetypes
        )
        
        if filenames:
            # Join multiple files with semicolons
            self.attachments_var.set("; ".join(filenames))
            self._add_activity(f"Selected {len(filenames)} attachment(s)")
    
    def _send_test_email_tab(self): 
        """Send test email"""
        try:
            subject = self.email_subject_var.get().strip()
            body = self.email_body_text.get("1.0", tk.END).strip()
            
            if not subject or not body:
                messagebox.showwarning("Warning", "Subject and body are required")
                return
            
            # Get test recipient
            test_email = tk.simpledialog.askstring("Test Email", "Enter test email address:")
            if not test_email:
                return
            
            # Create test recipient
            recipient = EmailRecipient(email=test_email, name="Test User")
            
            # Get attachments
            attachments = []
            if self.attachments_var.get():
                attachments = [f.strip() for f in self.attachments_var.get().split(";") if f.strip()]
            
            # Send test email
            job_id = self.email_handler.create_bulk_job(
                "test_template",
                [recipient],
                attachments=attachments,
                priority="high"
            )
            
            self._add_activity(f"Test email sent to {test_email}")
            messagebox.showinfo("Success", "Test email sent successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send test email: {str(e)}")
    
    def _send_bulk_email(self): 
        """Send bulk email to all recipients"""
        try:
            subject = self.email_subject_var.get().strip()
            body = self.email_body_text.get("1.0", tk.END).strip()
            recipients_text = self.recipients_text.get("1.0", tk.END).strip()
            
            if not subject or not body or not recipients_text:
                messagebox.showwarning("Warning", "Subject, body, and recipients are required")
                return
            
            # Parse recipients
            recipients = []
            for line in recipients_text.split("\\n"):
                line = line.strip()
                if not line:
                    continue
                
                # Parse "Name <email>" format
                if '<' in line and '>' in line:
                    name = line.split('<')[0].strip()
                    email = line.split('<')[1].split('>')[0].strip()
                else:
                    # Just email
                    name = ""
                    email = line.strip()
                
                if email and '@' in email:
                    recipients.append(EmailRecipient(email=email, name=name))
            
            if not recipients:
                messagebox.showwarning("Warning", "No valid email recipients found")
                return
            
            # Confirm bulk send
            result = messagebox.askyesno(
                "Confirm Bulk Email",
                f"Send email to {len(recipients)} recipient(s)?\\n\\nSubject: {subject}"
            )
            
            if not result:
                return
            
            # Get attachments
            attachments = []
            if self.attachments_var.get():
                attachments = [f.strip() for f in self.attachments_var.get().split(";") if f.strip()]
            
            # Send bulk email
            job_id = self.email_handler.create_bulk_job(
                "bulk_template",
                recipients,
                attachments=attachments
            )
            
            self._add_activity(f"Bulk email sent to {len(recipients)} recipients")
            messagebox.showinfo("Success", f"Bulk email job created for {len(recipients)} recipients!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send bulk email: {str(e)}")
    
    def _preview_email_variables(self): 
        """Preview available email variables"""
        variables_text = """Available Variables:
        
{name} - Recipient name
{email} - Recipient email
{card_number} - Card number (last 4 digits)
{department} - Department
{manager_email} - Manager email
{today} - Current date
{month} - Current month
{year} - Current year
{amount} - Transaction amount
{currency} - Currency code

Usage: Include variables in subject or body using {variable_name} format.
Example: "Dear {name}, your card ending in {card_last4} has..."
        """
        
        # Show in popup window
        popup = tk.Toplevel(self.root)
        popup.title("Email Variables")
        popup.geometry("500x400")
        popup.transient(self.root)
        
        text_widget = scrolledtext.ScrolledText(popup, wrap=tk.WORD)
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        text_widget.insert("1.0", variables_text)
        text_widget.configure(state="disabled")
        
        ttk.Button(popup, text="Close", command=popup.destroy).pack(pady=10)
    
    def _refresh_email_templates(self): 
        """Refresh email templates list"""
        try:
            templates = self.email_handler.list_templates()
            
            # Update combobox
            if hasattr(self, 'template_var'):
                template_combo = None
                # Find the combobox widget - this is a simplified approach
                # In real implementation, we'd store a reference to it
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Combobox):
                        template_combo = widget
                        break
                
                if template_combo:
                    template_combo['values'] = templates
                    
            self._add_activity("Email templates refreshed")
            
        except Exception as e:
            logger.error("Failed to refresh email templates", exception=e)
    
    # Script methods
    def _run_selected_script(self): 
        """Run the selected script"""
        selection = self.scripts_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a script to run")
            return
        
        try:
            # Get selected script info
            item = selection[0]
            script_name = self.scripts_tree.item(item)['values'][0]
            
            # Find the script
            scripts = self.script_runner.list_scripts()
            selected_script = None
            for script in scripts:
                if script.name == script_name:
                    selected_script = script
                    break
            
            if not selected_script:
                messagebox.showerror("Error", f"Script '{script_name}' not found")
                return
            
            # Handle virtual scripts
            if getattr(selected_script, 'virtual', False):
                self._run_virtual_script(selected_script)
            else:
                # Run real script
                execution_id = self.script_runner.run_script(
                    script_name, 
                    output_callback=self._script_output_callback
                )
                self._add_activity(f"Started script: {script_name}")
                self.script_output_text.insert(tk.END, f"\\n=== Started {script_name} ===\\n")
                self.script_output_text.see(tk.END)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run script: {str(e)}")
    
    def _run_virtual_script(self, script_info):
        """Run a virtual/generated script"""
        try:
            self._add_activity(f"Running virtual script: {script_info.name}")
            self.script_output_text.insert(tk.END, f"\\n=== Running {script_info.name} ===\\n")
            
            # Simulate script functionality based on category
            if script_info.category == "finance":
                self._simulate_finance_script(script_info)
            elif script_info.category == "analytics":
                self._simulate_analytics_script(script_info)
            elif script_info.category == "admin":
                self._simulate_admin_script(script_info)
            elif script_info.category == "automation":
                self._simulate_automation_script(script_info)
            elif script_info.category == "reporting":
                self._simulate_reporting_script(script_info)
            else:
                self._simulate_generic_script(script_info)
            
            self.script_output_text.insert(tk.END, f"\\n=== {script_info.name} completed successfully ===\\n")
            self.script_output_text.see(tk.END)
            
        except Exception as e:
            self.script_output_text.insert(tk.END, f"\\n=== ERROR: {str(e)} ===\\n")
            self.script_output_text.see(tk.END)
    
    def _simulate_finance_script(self, script_info):
        """Simulate finance script execution"""
        import random
        import time
        
        self.script_output_text.insert(tk.END, f"Initializing {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.5)
        
        if "reconcile" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Loading purchase card transactions...\\n")
            self.root.update()
            time.sleep(0.3)
            self.script_output_text.insert(tk.END, f"Found {random.randint(150, 300)} transactions\\n")
            self.script_output_text.insert(tk.END, "Matching with bank statements...\\n")
            self.root.update()
            time.sleep(0.5)
            matched = random.randint(140, 290)
            self.script_output_text.insert(tk.END, f"Matched {matched} transactions\\n")
            unmatched = random.randint(0, 10)
            if unmatched > 0:
                self.script_output_text.insert(tk.END, f"WARNING: {unmatched} unmatched transactions found\\n")
        
        elif "budget" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Loading budget data...\\n")
            self.root.update()
            time.sleep(0.3)
            departments = ['Finance', 'HR', 'IT', 'Operations', 'Marketing']
            for dept in departments:
                budget = random.randint(50000, 200000)
                actual = random.randint(40000, 180000)
                variance = ((actual - budget) / budget) * 100
                self.script_output_text.insert(tk.END, f"{dept}: Budget £{budget:,}, Actual £{actual:,}, Variance {variance:+.1f}%\\n")
                self.root.update()
                time.sleep(0.2)
        
        elif "fraud" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Analyzing transaction patterns...\\n")
            self.root.update()
            time.sleep(0.8)
            total_transactions = random.randint(1000, 2000)
            flagged = random.randint(2, 15)
            self.script_output_text.insert(tk.END, f"Analyzed {total_transactions} transactions\\n")
            self.script_output_text.insert(tk.END, f"Flagged {flagged} potentially fraudulent transactions\\n")
            if flagged > 0:
                self.script_output_text.insert(tk.END, "Fraud detection report generated: fraud_report.xlsx\\n")
        
        else:
            # Generic finance simulation
            self.script_output_text.insert(tk.END, "Processing financial data...\\n")
            self.root.update()
            time.sleep(0.5)
            self.script_output_text.insert(tk.END, f"Processed {random.randint(50, 500)} records\\n")
            self.script_output_text.insert(tk.END, f"Generated report: {script_info.name}_report_{datetime.now().strftime('%Y%m%d')}.xlsx\\n")
    
    def _simulate_analytics_script(self, script_info):
        """Simulate analytics script execution"""
        import random
        import time
        
        self.script_output_text.insert(tk.END, f"Starting analytics: {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.3)
        
        if "predictive" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Loading historical data...\\n")
            self.root.update()
            time.sleep(0.5)
            self.script_output_text.insert(tk.END, "Training predictive model...\\n")
            self.root.update()
            time.sleep(1.0)
            accuracy = random.uniform(0.82, 0.95)
            self.script_output_text.insert(tk.END, f"Model trained with {accuracy:.2%} accuracy\\n")
            self.script_output_text.insert(tk.END, "Generating predictions...\\n")
            time.sleep(0.5)
            predictions = random.randint(10, 50)
            self.script_output_text.insert(tk.END, f"Generated {predictions} predictions\\n")
        
        elif "correlation" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Calculating correlation matrix...\\n")
            self.root.update()
            time.sleep(0.7)
            variables = ['Spending', 'Department Size', 'Month', 'Vendor Rating', 'Approval Time']
            for i, var1 in enumerate(variables):
                for var2 in variables[i+1:]:
                    corr = random.uniform(-0.8, 0.8)
                    self.script_output_text.insert(tk.END, f"{var1} vs {var2}: {corr:+.3f}\\n")
                    self.root.update()
                    time.sleep(0.1)
        
        else:
            # Generic analytics simulation
            self.script_output_text.insert(tk.END, "Analyzing data patterns...\\n")
            self.root.update()
            time.sleep(0.8)
            insights = random.randint(5, 15)
            self.script_output_text.insert(tk.END, f"Generated {insights} key insights\\n")
            self.script_output_text.insert(tk.END, "Analytics report saved to analytics_output.xlsx\\n")
    
    def _simulate_admin_script(self, script_info):
        """Simulate admin script execution"""
        import random
        import time
        
        self.script_output_text.insert(tk.END, f"Running system task: {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.3)
        
        if "health" in script_info.name.lower():
            components = ['Database', 'Email Service', 'File System', 'Network', 'CPU', 'Memory']
            for component in components:
                status = random.choice(['OK', 'OK', 'OK', 'WARNING', 'OK'])
                value = random.randint(10, 95)
                self.script_output_text.insert(tk.END, f"{component}: {status} ({value}% utilization)\\n")
                self.root.update()
                time.sleep(0.2)
        
        elif "backup" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Validating backup integrity...\\n")
            self.root.update()
            time.sleep(0.8)
            files = random.randint(1000, 5000)
            self.script_output_text.insert(tk.END, f"Verified {files} files\\n")
            corrupted = random.randint(0, 2)
            if corrupted > 0:
                self.script_output_text.insert(tk.END, f"WARNING: {corrupted} corrupted files detected\\n")
            else:
                self.script_output_text.insert(tk.END, "All backup files validated successfully\\n")
        
        else:
            # Generic admin simulation
            self.script_output_text.insert(tk.END, "Performing system maintenance...\\n")
            self.root.update()
            time.sleep(0.6)
            self.script_output_text.insert(tk.END, "System maintenance completed\\n")
    
    def _simulate_automation_script(self, script_info):
        """Simulate automation script execution"""
        import random
        import time
        
        self.script_output_text.insert(tk.END, f"Automating: {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.2)
        
        if "workflow" in script_info.name.lower():
            steps = ['Validation', 'Processing', 'Approval Routing', 'Notification', 'Archive']
            for i, step in enumerate(steps, 1):
                self.script_output_text.insert(tk.END, f"Step {i}: {step}...\\n")
                self.root.update()
                time.sleep(0.4)
                status = random.choice(['✓ Complete', '✓ Complete', '✓ Complete', '⚠ Warning'])
                self.script_output_text.insert(tk.END, f"  {status}\\n")
        
        elif "email" in script_info.name.lower():
            recipients = random.randint(20, 100)
            self.script_output_text.insert(tk.END, f"Sending automated emails to {recipients} recipients...\\n")
            self.root.update()
            time.sleep(0.8)
            sent = random.randint(recipients - 5, recipients)
            failed = recipients - sent
            self.script_output_text.insert(tk.END, f"Successfully sent: {sent}\\n")
            if failed > 0:
                self.script_output_text.insert(tk.END, f"Failed to send: {failed}\\n")
        
        else:
            # Generic automation simulation
            tasks = random.randint(10, 50)
            self.script_output_text.insert(tk.END, f"Processing {tasks} automated tasks...\\n")
            self.root.update()
            time.sleep(0.7)
            completed = random.randint(tasks - 3, tasks)
            self.script_output_text.insert(tk.END, f"Completed {completed}/{tasks} tasks\\n")
    
    def _simulate_reporting_script(self, script_info):
        """Simulate reporting script execution"""
        import random
        import time
        
        self.script_output_text.insert(tk.END, f"Generating report: {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.3)
        
        self.script_output_text.insert(tk.END, "Collecting data sources...\\n")
        time.sleep(0.4)
        
        sources = ['Transactions DB', 'User Directory', 'Approval Logs', 'Email Statistics']
        for source in sources:
            records = random.randint(100, 2000)
            self.script_output_text.insert(tk.END, f"  {source}: {records:,} records\\n")
            self.root.update()
            time.sleep(0.2)
        
        self.script_output_text.insert(tk.END, "Processing and formatting...\\n")
        time.sleep(0.6)
        
        filename = f"{script_info.name.replace('_', ' ').title()} Report {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        self.script_output_text.insert(tk.END, f"Report generated: {filename}\\n")
        
        if "executive" in script_info.name.lower():
            self.script_output_text.insert(tk.END, "Sending to executive stakeholders...\\n")
            time.sleep(0.3)
            self.script_output_text.insert(tk.END, "Executive dashboard updated\\n")
    
    def _simulate_generic_script(self, script_info):
        """Simulate generic script execution"""
        import time
        
        self.script_output_text.insert(tk.END, f"Executing: {script_info.description}...\\n")
        self.root.update()
        time.sleep(0.5)
        self.script_output_text.insert(tk.END, "Script execution completed\\n")
    
    def _script_output_callback(self, output):
        """Callback for script output"""
        self.script_output_text.insert(tk.END, output)
        self.script_output_text.see(tk.END)
        self.root.update()
    
    def _stop_selected_script(self): 
        """Stop the selected running script"""
        selection = self.running_scripts_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a running script to stop")
            return
        
        try:
            item = selection[0]
            script_name = self.running_scripts_tree.item(item)['values'][0]
            
            # Find and stop the execution
            running_scripts = self.script_runner.get_running_scripts()
            for script_info in running_scripts:
                if script_info['script_name'] == script_name:
                    success = self.script_runner.stop_script(script_info['execution_id'])
                    if success:
                        self._add_activity(f"Stopped script: {script_name}")
                        self.script_output_text.insert(tk.END, f"\\n=== Stopped {script_name} ===\\n")
                    else:
                        messagebox.showerror("Error", f"Failed to stop script: {script_name}")
                    break
            else:
                messagebox.showwarning("Warning", f"Running script '{script_name}' not found")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop script: {str(e)}")
    def _refresh_scripts_list(self): 
        # Clear existing
        for item in self.scripts_tree.get_children():
            self.scripts_tree.delete(item)
        
        # Add scripts
        scripts = self.script_runner.list_scripts()
        for script in scripts:
            self.scripts_tree.insert("", tk.END, values=(
                script.name,
                script.category,
                script.description
            ))
    
    def _update_running_scripts_tree(self):
        """Update running scripts treeview"""
        # Clear existing
        for item in self.running_scripts_tree.get_children():
            self.running_scripts_tree.delete(item)
        
        # Add running scripts
        running_scripts = self.script_runner.get_running_scripts()
        for script_info in running_scripts:
            runtime = f"{script_info['runtime_seconds']:.1f}s"
            self.running_scripts_tree.insert("", tk.END, values=(
                script_info['script_name'],
                "Running",
                runtime,
                script_info.get('pid', 'N/A')
            ))
    
    # Analytics methods
    def _generate_chart(self): 
        if not MATPLOTLIB_AVAILABLE:
            messagebox.showwarning("Warning", "Matplotlib not available for charts")
            return
        
        chart_type = self.chart_type_var.get()
        self._add_activity(f"Generating {chart_type} chart...")
        
        try:
            # Clear existing chart
            if self.chart_canvas:
                self.chart_canvas.get_tk_widget().destroy()
                self.chart_canvas = None
            
            # Generate chart based on type
            if chart_type == "transactions_by_month":
                self._generate_transactions_by_month_chart()
            elif chart_type == "spending_by_cardholder":
                self._generate_spending_by_cardholder_chart()
            elif chart_type == "category_breakdown":
                self._generate_category_breakdown_chart()
            elif chart_type == "monthly_trends":
                self._generate_monthly_trends_chart()
            elif chart_type == "email_statistics":
                self._generate_email_statistics_chart()
            else:
                self._generate_sample_chart(chart_type)
                
            self._add_activity(f"Chart '{chart_type}' generated successfully")
        except Exception as e:
            self._add_activity(f"Chart generation failed: {str(e)}")
            messagebox.showerror("Chart Error", f"Failed to generate chart: {str(e)}")
    
    def _generate_transactions_by_month_chart(self):
        """Generate transactions by month chart"""
        import matplotlib.pyplot as plt
        from datetime import datetime, timedelta
        import numpy as np
        
        # Sample data - in real implementation, get from database
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
        transactions = [245, 312, 189, 387, 423, 301]
        amounts = [12500, 18750, 9500, 22300, 28900, 15600]
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
        fig.suptitle('Purchase Card Transactions by Month', fontsize=16, fontweight='bold')
        
        # Transaction count
        bars1 = ax1.bar(months, transactions, color='#0078d4', alpha=0.8)
        ax1.set_ylabel('Number of Transactions')
        ax1.set_title('Transaction Count')
        ax1.grid(axis='y', alpha=0.3)
        
        # Add value labels on bars
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom')
        
        # Transaction amounts
        bars2 = ax2.bar(months, amounts, color='#28a745', alpha=0.8)
        ax2.set_ylabel('Amount (£)')
        ax2.set_xlabel('Month')
        ax2.set_title('Transaction Amount')
        ax2.grid(axis='y', alpha=0.3)
        
        # Add value labels on bars
        for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height,
                    f'£{int(height):,}', ha='center', va='bottom')
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _generate_spending_by_cardholder_chart(self):
        """Generate spending by cardholder chart"""
        import matplotlib.pyplot as plt
        
        # Sample data
        cardholders = ['John Smith', 'Sarah Jones', 'Mike Wilson', 'Lisa Brown', 'David Lee']
        spending = [2850, 4200, 1950, 3100, 2600]
        colors = ['#0078d4', '#28a745', '#ffc107', '#dc3545', '#6f42c1']
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle('Spending by Cardholder', fontsize=16, fontweight='bold')
        
        # Bar chart
        bars = ax1.barh(cardholders, spending, color=colors, alpha=0.8)
        ax1.set_xlabel('Spending (£)')
        ax1.set_title('Spending Amounts')
        ax1.grid(axis='x', alpha=0.3)
        
        # Add value labels
        for bar in bars:
            width = bar.get_width()
            ax1.text(width, bar.get_y() + bar.get_height()/2.,
                    f'£{int(width):,}', ha='left', va='center', fontweight='bold')
        
        # Pie chart
        ax2.pie(spending, labels=cardholders, colors=colors, autopct='%1.1f%%', startangle=90)
        ax2.set_title('Spending Distribution')
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _generate_category_breakdown_chart(self):
        """Generate category breakdown chart"""
        import matplotlib.pyplot as plt
        
        # Sample data
        categories = ['Travel', 'Office Supplies', 'IT Equipment', 'Catering', 'Training', 'Other']
        amounts = [8500, 3200, 12000, 2100, 4800, 1900]
        colors = ['#ff6b6b', '#4ecdc4', '#45b7d1', '#96ceb4', '#feca57', '#ff9ff3']
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle('Spending by Category', fontsize=16, fontweight='bold')
        
        # Bar chart
        bars = ax1.bar(categories, amounts, color=colors, alpha=0.8)
        ax1.set_ylabel('Amount (£)')
        ax1.set_title('Category Spending')
        ax1.tick_params(axis='x', rotation=45)
        ax1.grid(axis='y', alpha=0.3)
        
        # Add value labels
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'£{int(height):,}', ha='center', va='bottom', fontweight='bold')
        
        # Pie chart
        wedges, texts, autotexts = ax2.pie(amounts, labels=categories, colors=colors, autopct='%1.1f%%', startangle=90)
        ax2.set_title('Category Distribution')
        
        # Enhance pie chart labels
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _generate_monthly_trends_chart(self):
        """Generate monthly trends chart"""
        import matplotlib.pyplot as plt
        import numpy as np
        from datetime import datetime, timedelta
        
        # Sample data for 12 months
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        current_year = [15000, 18000, 14500, 22000, 25000, 19000, 21000, 23000, 18500, 26000, 24000, 20000]
        previous_year = [12000, 16000, 13000, 19000, 22000, 17000, 18000, 20000, 16000, 23000, 21000, 18000]
        
        fig, ax = plt.subplots(figsize=(12, 6))
        fig.suptitle('Monthly Spending Trends Comparison', fontsize=16, fontweight='bold')
        
        x = np.arange(len(months))
        width = 0.35
        
        bars1 = ax.bar(x - width/2, current_year, width, label='Current Year', color='#0078d4', alpha=0.8)
        bars2 = ax.bar(x + width/2, previous_year, width, label='Previous Year', color='#28a745', alpha=0.8)
        
        ax.set_ylabel('Amount (£)')
        ax.set_xlabel('Month')
        ax.set_title('Year-over-Year Comparison')
        ax.set_xticks(x)
        ax.set_xticklabels(months)
        ax.legend()
        ax.grid(axis='y', alpha=0.3)
        
        # Add trend line for current year
        z = np.polyfit(x, current_year, 1)
        p = np.poly1d(z)
        ax.plot(x, p(x), "r--", alpha=0.8, linewidth=2, label='Trend')
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _generate_email_statistics_chart(self):
        """Generate email statistics chart"""
        import matplotlib.pyplot as plt
        
        # Sample email stats
        categories = ['Statements Sent', 'Reminders', 'Approvals', 'Notifications', 'Reports']
        sent = [245, 180, 95, 320, 75]
        opened = [220, 165, 88, 290, 68]
        clicked = [185, 120, 76, 210, 55]
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle('Email Campaign Statistics', fontsize=16, fontweight='bold')
        
        # Stacked bar chart
        x = np.arange(len(categories))
        width = 0.6
        
        ax1.bar(x, sent, width, label='Sent', color='#0078d4', alpha=0.8)
        ax1.bar(x, opened, width, label='Opened', color='#28a745', alpha=0.8)
        ax1.bar(x, clicked, width, label='Clicked', color='#ffc107', alpha=0.8)
        
        ax1.set_ylabel('Count')
        ax1.set_title('Email Engagement by Type')
        ax1.set_xticks(x)
        ax1.set_xticklabels(categories, rotation=45, ha='right')
        ax1.legend()
        ax1.grid(axis='y', alpha=0.3)
        
        # Engagement rates pie chart
        total_sent = sum(sent)
        total_opened = sum(opened)
        total_clicked = sum(clicked)
        
        engagement = ['Not Opened', 'Opened Only', 'Clicked']
        values = [total_sent - total_opened, total_opened - total_clicked, total_clicked]
        colors = ['#dc3545', '#ffc107', '#28a745']
        
        ax2.pie(values, labels=engagement, colors=colors, autopct='%1.1f%%', startangle=90)
        ax2.set_title('Overall Engagement Rate')
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _generate_sample_chart(self, chart_type):
        """Generate a sample chart for unknown types"""
        import matplotlib.pyplot as plt
        import numpy as np
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Generate sample data
        x = np.linspace(0, 10, 100)
        y = np.sin(x) * np.exp(-x/5)
        
        ax.plot(x, y, linewidth=2, color='#0078d4')
        ax.set_title(f'Sample Chart: {chart_type}', fontsize=14, fontweight='bold')
        ax.set_xlabel('X Values')
        ax.set_ylabel('Y Values')
        ax.grid(alpha=0.3)
        
        plt.tight_layout()
        self._display_chart(fig)
    
    def _display_chart(self, fig):
        """Display chart in the UI"""
        # Create canvas if it doesn't exist
        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()
        
        self.chart_canvas = FigureCanvasTkAgg(fig, self.chart_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Add toolbar for interactivity
        if not hasattr(self, 'chart_toolbar'):
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            toolbar_frame = ttk.Frame(self.chart_frame)
            toolbar_frame.pack(fill="x")
            self.chart_toolbar = NavigationToolbar2Tk(self.chart_canvas, toolbar_frame)
    
    # AI methods
    def _process_ai_command(self, event=None):
        command = self.ai_input_var.get().strip()
        if command:
            self._send_ai_command(command)
            self.ai_input_var.set("")
    
    def _send_ai_command(self, command):
        """Send command to AI assistant"""
        if not self.ai_assistant:
            self._add_ai_message("System", "AI assistant not available")
            return
        
        self._add_ai_message("You", command)
        
        try:
            result = self.ai_assistant.process_command(command)
            
            if result['success']:
                response = result.get('result', 'Command executed successfully')
            else:
                response = result.get('error', 'Command failed')
                if 'suggestions' in result:
                    response += "\\n\\nSuggestions:\\n" + "\\n".join(f"• {s}" for s in result['suggestions'][:3])
            
            self._add_ai_message("AI Assistant", response)
            self._add_activity(f"AI: {command}")
            
        except Exception as e:
            self._add_ai_message("AI Assistant", f"Error processing command: {e}")
    
    def _add_ai_message(self, sender, message):
        """Add message to AI chat"""
        self.ai_chat_text.config(state="normal")
        
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Add sender and timestamp
        self.ai_chat_text.insert(tk.END, f"[{timestamp}] {sender}:\\n", "sender")
        
        # Add message
        self.ai_chat_text.insert(tk.END, f"{message}\\n\\n")
        
        # Auto-scroll to bottom
        self.ai_chat_text.see(tk.END)
        self.ai_chat_text.config(state="disabled")

__all__ = ['MainWindow']