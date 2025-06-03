#!/usr/bin/env python3
"""
TelegramExcelViewer - Main Application Entry Point

This module serves as the coordination layer between GUI and Functions modules.
It handles application initialization, configuration, and the main execution flow.
"""

import os
import sys
from pathlib import Path

# Add the current directory to the Python path for relative imports
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

try:
from GUI import TelegramExcelGUI
from Functions import ExcelHandler, TelegramOperations, DataManager
except ImportError as e:
    print(f"❌ Error importing modules: {e}")
    print("Please ensure GUI.py and Functions.py are in the same directory as Main.py")
    sys.exit(1)

class TelegramExcelApplication:
    """
    Main application coordinator class.
    Manages the lifecycle and coordination between GUI and Functions modules.
    """
    
    def __init__(self):
        """Initialize the application with default configuration"""
        self.config = self._load_default_config()
        self.functions = None
        self.gui = None
        
        # Initialize application components
        self._initialize_components()
    
    def _load_default_config(self):
        """Load default application configuration"""
        return {
            'default_path': os.path.expanduser("~/Downloads/Porno/Descargar/CanalesUnidos"),
            'page_size': 20,
            'window_geometry': "1200x700",
            'app_title': "Telegram Excel Viewer",
            'target_chat': "2532518781",  # Default target chat for forwarding
            'data_number': 1,  # Default tdl data number
            'timeout_seconds': 60,  # Default timeout for operations
        }
    
    def _initialize_components(self):
        """Initialize Functions and GUI components"""
        try:
            # Initialize Functions module with configuration
            self.functions = TelegramExcelFunctions(self.config)
            
            # Initialize GUI module with Functions reference
            self.gui = TelegramExcelGUI(self.functions)
            
            # Configure window title and geometry
            self.gui.root.title(self.config['app_title'])
            self.gui.root.geometry(self.config['window_geometry'])
            
            print("✅ Application components initialized successfully")
            
        except Exception as e:
            print(f"❌ Error initializing application components: {e}")
            raise
    
    def run(self):
        """
        Start the application main loop
        
        This method serves as the primary entry point for running the application.
        It handles any final initialization and starts the GUI main loop.
        """
        try:
            print(f"🚀 Starting {self.config['app_title']}...")
            
            # Ensure default directory exists
            os.makedirs(self.config['default_path'], exist_ok=True)
            
            # Start the GUI main loop
            self.gui.run()
            
        except KeyboardInterrupt:
            print("\n🔴 Application interrupted by user")
            self._cleanup()
        except Exception as e:
            print(f"❌ Critical error during application execution: {e}")
            self._cleanup()
            raise
    
    def _cleanup(self):
        """Perform cleanup operations before application shutdown"""
        try:
            if self.functions:
                self.functions.cleanup()
            print("🧹 Cleanup completed")
        except Exception as e:
            print(f"⚠️ Warning during cleanup: {e}")

def main():
    """
    Main entry point for the application.
    
    This function creates and runs the TelegramExcelApplication instance.
    It also handles any top-level exceptions and provides user-friendly error messages.
    """
    try:
        # Create and run the application
        app = TelegramExcelApplication()
        app.run()
        
    except ImportError as e:
        print(f"❌ Module import error: {e}")
        print("Please ensure all required modules (GUI.py, Functions.py) are available")
        print("and all dependencies are installed:")
        print("  - tkinter (usually included with Python)")
        print("  - openpyxl: pip install openpyxl")
        sys.exit(1)
        
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
        print("Please ensure all application files are in the correct location")
        sys.exit(1)
        
    except PermissionError as e:
        print(f"❌ Permission error: {e}")
        print("Please check file and directory permissions")
        sys.exit(1)
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        print("Please check the application logs and try again")
        sys.exit(1)

if __name__ == "__main__":
    """
    Entry point when script is run directly.
    
    This ensures the main() function is only called when the script
    is executed directly, not when imported as a module.
    """
    main()
