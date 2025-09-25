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
            style.configure("TButton", background="#0078d4", foreground="#ffffff")
            style.map("TButton", background=[("active", "#106ebe")])
            
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
        self.recipients_text = scrolledtext.ScrolledText(recipients_frame, height=8, width=40)
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
        self.email_body_text = scrolledtext.ScrolledText(compose_frame, height=15)
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
    def _import_excel(self): messagebox.showinfo("Info", "Excel import functionality")
    def _export_data(self): messagebox.showinfo("Info", "Data export functionality") 
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
    def _quick_import_excel(self): self._add_activity("Quick: Import Excel requested")
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
    def _edit_cardholder(self): messagebox.showinfo("Info", "Edit cardholder functionality")
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
    def _filter_cardholders(self, event): pass  # Search functionality
    
    # Statement methods
    def _generate_all_statements(self): self._add_activity("Generate all statements requested")
    def _generate_selected_statements(self): self._add_activity("Generate selected statements requested")
    def _preview_statement(self): messagebox.showinfo("Info", "Statement preview functionality")
    def _show_statement_context_menu(self, event): pass
    
    # Email methods
    def _load_email_template(self): pass
    def _save_email_template(self): pass
    def _new_email_template(self): pass
    def _import_email_recipients(self): pass
    def _add_cardholders_as_recipients(self): pass
    def _browse_attachments(self): pass
    def _send_test_email_tab(self): pass
    def _send_bulk_email(self): pass
    def _preview_email_variables(self): pass
    def _refresh_email_templates(self): pass
    
    # Script methods
    def _run_selected_script(self): pass
    def _stop_selected_script(self): pass
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
        self._add_activity("Chart generation requested")
    
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