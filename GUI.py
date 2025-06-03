#!/usr/bin/env python3

import tkinter as tk
from tkinter import ttk, messagebox
import threading

class TelegramExcelGUI:
    def __init__(self, functions_handler):
        """
        Initialize GUI with reference to functions handler
        
        Args:
            functions_handler: Instance of Functions class that handles business logic
        """
        self.functions = functions_handler
        self.root = tk.Tk()
        self.root.geometry("1200x700")
        self.root.title("Telegram Excel Viewer")
        
        # GUI State variables
        self.current_page = 0
        self.page_size = 20
        self.data = []  # Current page data for tree display
        
        # UI Components
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.page_label = ttk.Label(self.root, text="Página 0 de 0")
        self.total_label = ttk.Label(self.root, text="Total: 0 registros")
        self.prev_btn = ttk.Button(self.root, text="⬅️ Anterior", command=self.prev_page, state="disabled")
        self.next_btn = ttk.Button(self.root, text="Siguiente ➡️", command=self.next_page, state="disabled")
        
        self.setup_ui()
        
        # Set functions callback for GUI updates
        if self.functions:
            self.functions.set_gui_callback(self)
        self.setup_bindings()
    
    def setup_ui(self):
        """Configure the main user interface components"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Buttons
        load_btn = ttk.Button(button_frame, text="📁 Cargar archivo Excel", command=self.load_file)
        load_btn.grid(row=0, column=0, padx=(0, 10))
        
        ready_btn = ttk.Button(button_frame, text="✅ Ver links listos", command=self.view_ready_links)
        ready_btn.grid(row=0, column=1, padx=(0, 10))
        
        exit_btn = ttk.Button(button_frame, text="❌ Salir", command=self.exit_app)
        exit_btn.grid(row=0, column=2, padx=(0, 10))
        
        # Progress bar
        self.progress.grid(row=0, column=3, padx=(10, 0), sticky="ew")
        self.progress.grid_remove()
        
        button_frame.columnconfigure(3, weight=1)
        
        # Treeview setup
        self.setup_treeview(main_frame)
        
        # Pagination setup
        self.setup_pagination(main_frame)
        
        # Instructions
        self.setup_instructions(main_frame)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def setup_treeview(self, parent):
        """Setup the main data display treeview"""
        columns = ('Link', 'Formato', 'Duration', 'Size', 'File', 'Text')
        self.tree = ttk.Treeview(parent, columns=columns, show='headings', height=20)
        
        # Configure columns
        for col in columns:
            self.tree.heading(col, text=col)
            if col == 'Link':
                self.tree.column(col, width=300)
            else:
                self.tree.column(col, width=120)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid placement
        self.tree.grid(row=1, column=0, sticky="nsew")
        v_scrollbar.grid(row=1, column=1, sticky="ns")
        h_scrollbar.grid(row=2, column=0, sticky="ew")
    
    def setup_pagination(self, parent):
        """Setup pagination controls"""
        pagination_frame = ttk.Frame(parent)
        pagination_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        
        self.prev_btn = ttk.Button(pagination_frame, text="⬅️ Anterior", command=self.prev_page, state="disabled")
        self.prev_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.page_label = ttk.Label(pagination_frame, text="Página 0 de 0")
        self.page_label.grid(row=0, column=1, padx=(0, 10))
        
        self.next_btn = ttk.Button(pagination_frame, text="Siguiente ➡️", command=self.next_page, state="disabled")
        self.next_btn.grid(row=0, column=2, padx=(0, 10))
        
        self.total_label = ttk.Label(pagination_frame, text="Total: 0 registros")
        self.total_label.grid(row=0, column=3, padx=(10, 0))
    
    def setup_instructions(self, parent):
        """Setup instruction label"""
        instructions = ttk.Label(parent, 
                               text="📖 Instrucciones: Click en link para abrir • Enter para abrir • Cmd+Click para descargar • Cmd+Enter para descargar • Opt+Click/Opt+Enter para reenviar")
        instructions.grid(row=4, column=0, columnspan=2, pady=(10, 0), sticky="w")
    
    def setup_bindings(self):
        """Configure event bindings for user interactions"""
        self.tree.bind('<Button-1>', self.on_single_click)
        self.tree.bind('<Return>', self.on_enter_key)
        self.tree.bind('<Command-Return>', self.on_command_enter)
        self.tree.bind('<Option-Return>', self.on_option_enter)
        self.tree.bind('<FocusIn>', lambda _: self.tree.focus_set())
        self.tree.focus_set()
    
    # Event Handlers
    def on_single_click(self, event):
        """Handle single click events on treeview items"""
        item = self.tree.identify('item', event.x, event.y)
        column = self.tree.identify('column', event.x, event.y)
        
        if item and column == '#1':  # Click on link column
            if event.state & 0x80000:  # Option/Alt key
                self.forward_link(item)
            elif event.state & 0x8:  # Cmd key
                self.download_link(item)
            else:  # Regular click
                self.open_link(item)
    
    def on_enter_key(self, event):
        """Handle Enter key press"""
        selection = self.tree.selection()
        if selection:
            if event.state & 0x8:  # Cmd+Enter
                self.download_link(selection[0])
            else:  # Regular Enter
                self.open_link(selection[0])
    
    def on_command_enter(self, _):
        """Handle Cmd+Enter key combination"""
        selection = self.tree.selection()
        if selection:
            self.download_link(selection[0])

    def on_option_enter(self, _):
        """Handle Option+Enter key combination"""
        selection = self.tree.selection()
        if selection:
            self.forward_link(selection[0])
    
    # Action Methods
    def load_file(self):
        """Initiate file loading process"""
        self.show_progress()
        threading.Thread(target=self._load_file_thread, daemon=True).start()
    
    def _load_file_thread(self):
        """Background thread for file loading"""
        try:
            success, message = self.functions.load_excel_file()
            self.root.after(0, lambda: self._load_complete(success, message))
        except Exception as e:
            self.root.after(0, lambda: self._load_error(str(e)))
    
    def _load_complete(self, success, message):
        """Handle completion of file loading"""
        self.hide_progress()
        if success:
            self.current_page = 0
            self.load_current_page()
            self.update_pagination()
            messagebox.showinfo("✅ Éxito", message)
        else:
            messagebox.showerror("❌ Error", message)
    
    def _load_error(self, error_msg):
        """Handle file loading errors"""
        self.hide_progress()
        messagebox.showerror("❌ Error", f"Error al cargar el archivo: {error_msg}")
    
    def open_link(self, item_id):
        """Open link in Telegram"""
        clicked_data = self.get_data_by_item(item_id)
        if clicked_data:
            try:
                self.functions.open_in_telegram(clicked_data['link'])
                self.functions.mark_as_clicked(clicked_data['excel_row'])
                self.refresh_current_display()
            except Exception as e:
                messagebox.showerror("❌ Error", f"Error al abrir el enlace: {str(e)}")
    
    def download_link(self, item_id):
        """Download link using tlg command"""
        clicked_data = self.get_data_by_item(item_id)
        if clicked_data:
            try:
                self.functions.download_with_tlg(clicked_data['link'])
                self.functions.mark_as_clicked(clicked_data['excel_row'])
                self.refresh_current_display()
            except Exception as e:
                messagebox.showerror("❌ Error", f"Error al descargar: {str(e)}")
    
    def forward_link(self, item_id):
        """Forward link using tdl command"""
        clicked_data = self.get_data_by_item(item_id)
        if clicked_data:
            try:
                self.show_progress()
                threading.Thread(
                    target=self._forward_link_thread, 
                    args=(clicked_data,), 
                    daemon=True
                ).start()
            except Exception as e:
                self.hide_progress()
                messagebox.showerror("❌ Error", f"Error al reenviar: {str(e)}")
    
    def _forward_link_thread(self, clicked_data):
        """Background thread for forwarding links"""
        try:
            success = self.functions.forward_with_tdl(clicked_data['link'])
            self.root.after(0, lambda: self._forward_complete(success, clicked_data))
        except Exception as e:
            self.root.after(0, lambda: self._forward_error(str(e)))
    
    def _forward_complete(self, success, clicked_data):
        """Handle completion of link forwarding"""
        self.hide_progress()
        if success:
            messagebox.showinfo("✅ Éxito", "Video reenviado correctamente al canal")
            self.functions.mark_as_clicked(clicked_data['excel_row'])
            self.refresh_current_display()
        else:
            messagebox.showerror("❌ Error", "No se pudo reenviar el video ni enviar como texto")
    
    def _forward_error(self, error_msg):
        """Handle forwarding errors"""
        self.hide_progress()
        messagebox.showerror("❌ Error", f"Error durante el reenvío: {error_msg}")
    
    # Pagination Methods
    def prev_page(self):
        """Navigate to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.load_current_page()
            self.update_pagination()
    
    def next_page(self):
        """Navigate to next page"""
        total_pages = self.get_total_pages()
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self.load_current_page()
            self.update_pagination()
    
    def get_total_pages(self):
        """Calculate total number of pages"""
        total_records = self.functions.get_total_records()
        return (total_records + self.page_size - 1) // self.page_size if total_records > 0 else 0
    
    def load_current_page(self):
        """Load and display current page data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.data = []
        
        # Get page data from functions
        page_data = self.functions.get_page_data(self.current_page, self.page_size)
        
        for data_item in page_data:
            row_data = data_item['data']
            item_id = self.tree.insert('', tk.END, values=row_data)
            
            # Mark clicked items visually
            if data_item['is_clicked']:
                current_values = list(row_data)
                current_values[0] = f"✅ {data_item['link']}"
                self.tree.item(item_id, values=current_values)
            
            self.data.append({
                'tree_item': item_id,
                'excel_row': data_item['excel_row'],
                'link': data_item['link'],
                'is_clicked': data_item['is_clicked']
            })
    
    def update_pagination(self):
        """Update pagination display and button states"""
        total_pages = self.get_total_pages()
        total_records = self.functions.get_total_records()
        current_page_display = self.current_page + 1 if total_records > 0 else 0
        
        self.page_label.config(text=f"Página {current_page_display} de {total_pages}")
        self.total_label.config(text=f"Total: {total_records} registros")
        
        self.prev_btn.config(state="normal" if self.current_page > 0 else "disabled")
        self.next_btn.config(state="normal" if self.current_page < total_pages - 1 else "disabled")
    
    def refresh_current_display(self):
        """Refresh the current page display"""
        self.load_current_page()
    
    # Utility Methods
    def get_data_by_item(self, item_id):
        """Get data associated with a tree item"""
        for data_item in self.data:
            if data_item['tree_item'] == item_id:
                return data_item
        return None
    
    def show_progress(self):
        """Show progress indicator"""
        self.progress.grid()
        self.progress.start()
    
    def hide_progress(self):
        """Hide progress indicator"""
        self.progress.stop()
        self.progress.grid_remove()
    
    # Ready Links Window
    def view_ready_links(self):
        """Display window with processed/ready links"""
        ready_window = tk.Toplevel(self.root)
        ready_window.title("✅ Links Listos")
        ready_window.geometry("600x400")

        ready_frame = ttk.Frame(ready_window, padding="10")
        ready_frame.grid(row=0, column=0, sticky="nsew")

        # Treeview for ready links
        ready_tree = ttk.Treeview(ready_frame, columns=('Link', 'Status'), show='headings')
        ready_tree.heading('Link', text='Link')
        ready_tree.heading('Status', text='Status')
        ready_tree.column('Link', width=450)
        ready_tree.column('Status', width=100)

        # Scrollbar
        ready_v_scrollbar = ttk.Scrollbar(ready_frame, orient=tk.VERTICAL, command=ready_tree.yview)
        ready_tree.configure(yscrollcommand=ready_v_scrollbar.set)
        
        # Grid layout
        ready_tree.grid(row=0, column=0, sticky="nsew")
        ready_v_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configure expansion
        ready_frame.columnconfigure(0, weight=1)
        ready_frame.rowconfigure(0, weight=1)
        ready_window.columnconfigure(0, weight=1)
        ready_window.rowconfigure(0, weight=1)
        
        # Load ready links
        ready_links = self.functions.get_ready_links()
        for link in ready_links:
            ready_tree.insert('', tk.END, values=(link, '✅ Visto'))
        
        # Add refresh button
        refresh_btn = ttk.Button(ready_frame, text="🔄 Actualizar", 
                               command=lambda: self.refresh_ready_links(ready_tree))
        refresh_btn.grid(row=1, column=0, pady=(10, 0), sticky="w")
        
        # Add count label
        count_label = ttk.Label(ready_frame, text=f"Total links vistos: {len(ready_links)}")
        count_label.grid(row=2, column=0, pady=(5, 0), sticky="w")
    
    def refresh_ready_links(self, ready_tree):
        """Refresh the ready links display"""
        # Clear existing items
        for item in ready_tree.get_children():
            ready_tree.delete(item)
        
        # Reload ready links
        ready_links = self.functions.get_ready_links()
        for link in ready_links:
            ready_tree.insert('', tk.END, values=(link, '✅ Visto'))
        
        messagebox.showinfo("🔄 Actualizado", f"Se encontraron {len(ready_links)} links vistos")
    
    def exit_app(self):
        """Handle application exit"""
        if messagebox.askokcancel("❌ Salir", "¿Estás seguro de que quieres salir?"):
            self.root.quit()
            self.root.destroy()
    
    def run(self):
        """Start the GUI main loop"""
        self.root.mainloop()
