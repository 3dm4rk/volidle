import ctypes
import time
import os
import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
import sys
import json
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

class SystemUtilitiesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("System Utilities")
        self.root.geometry("450x500")
        self.root.resizable(False, False)
        
        # Configuration file
        self.config_file = "config.txt"
        
        # Initialize states
        self.is_running = False
        self.warning_shown = False
        self.last_slider_update = None
        self.ignore_volume_change = False
        
        # Load or initialize configuration
        self.config = self.load_config()
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.create_settings_tab()
        self.create_idle_detector_tab()
        self.create_volume_control_tab()
        
        # Initialize volume control
        self.volume_control = self.init_volume_control()
        
        # Set initial states based on config
        if self.config.get('volume_control_enabled', True) and self.volume_control:
            if self.config.get('saved_volume') is not None:
                self.set_volume(self.config['saved_volume'])
            self.monitor_volume_changes()
        
        if self.config.get('idle_detector_enabled', True):
            self.start_detection()
        
        # Hide window if configured to do so
        if self.config.get('hide_on_startup', False):
            self.root.withdraw()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_config(self):
        """Load configuration from file or create default"""
        default_config = {
            'idle_threshold': 30,
            'shutdown_delay': 30,
            'saved_volume': 50,
            'hide_on_startup': False,
            'idle_detector_enabled': True,
            'volume_control_enabled': True
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_config = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    return {**default_config, **loaded_config}
            return default_config
        except Exception as e:
            print(f"Error loading config: {e}")
            return default_config

    def save_config(self):
        """Save current configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config:\n{str(e)}")

    # ==============================================
    # SETTINGS TAB
    # ==============================================
    def create_settings_tab(self):
        """Create the settings tab for enabling/disabling features"""
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Main frame
        main_frame = ttk.Frame(self.settings_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Feature toggles frame
        toggles_frame = ttk.LabelFrame(main_frame, text="Feature Toggles", padding="10")
        toggles_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Idle detector toggle
        self.idle_detector_enabled = tk.BooleanVar(value=self.config.get('idle_detector_enabled', True))
        idle_toggle = ttk.Checkbutton(
            toggles_frame,
            text="Enable Idle Detector",
            variable=self.idle_detector_enabled,
            command=self.toggle_idle_detector
        )
        idle_toggle.pack(anchor=tk.W, pady=5)
        
        # Volume control toggle
        self.volume_control_enabled = tk.BooleanVar(value=self.config.get('volume_control_enabled', True))
        volume_toggle = ttk.Checkbutton(
            toggles_frame,
            text="Enable Volume Control",
            variable=self.volume_control_enabled,
            command=self.toggle_volume_control
        )
        volume_toggle.pack(anchor=tk.W, pady=5)
        
        # Status info
        ttk.Label(
            main_frame,
            text="Note: Changes take effect immediately",
            font=('Segoe UI', 9)
        ).pack(pady=10)

    def toggle_idle_detector(self):
        """Toggle idle detector on/off"""
        enabled = self.idle_detector_enabled.get()
        self.config['idle_detector_enabled'] = enabled
        self.save_config()
        
        if enabled:
            self.start_detection()
        else:
            self.stop_detection()
        
        # Update tab state
        self.notebook.tab(1, state=tk.NORMAL if enabled else tk.DISABLED)

    def toggle_volume_control(self):
        """Toggle volume control on/off"""
        enabled = self.volume_control_enabled.get()
        self.config['volume_control_enabled'] = enabled
        self.save_config()
        
        if enabled and self.volume_control:
            self.monitor_volume_changes()
            if self.config.get('saved_volume') is not None:
                self.set_volume(self.config['saved_volume'])
        else:
            if hasattr(self, 'last_slider_update') and self.last_slider_update:
                self.root.after_cancel(self.last_slider_update)
        
        # Update tab state and controls
        self.notebook.tab(2, state=tk.NORMAL if enabled else tk.DISABLED)
        self.update_volume_controls_state(enabled)

    def update_volume_controls_state(self, enabled):
        """Enable/disable all volume controls"""
        if not hasattr(self, 'volume_tab'):
            return
            
        state = tk.NORMAL if enabled else tk.DISABLED
        if hasattr(self, 'volume_slider'):
            self.volume_slider.config(state=state)
        if hasattr(self, 'custom_entry'):
            self.custom_entry.config(state=state)
        
        # Update button states in volume tab
        for child in self.volume_tab.winfo_children():
            for widget in child.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=state)
        
        # Update label
        if hasattr(self, 'volume_label'):
            self.volume_label.config(
                text="Current Volume: Checking..." if enabled else "Volume Control Disabled"
            )

    # ==============================================
    # IDLE DETECTOR TAB
    # ==============================================
    def create_idle_detector_tab(self):
        """Create the idle detector tab"""
        self.idle_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.idle_tab, text="Idle Detector")
        
        # Variables for idle detector
        self.idle_threshold = tk.IntVar(value=self.config.get('idle_threshold', 30))
        self.shutdown_delay = tk.IntVar(value=self.config.get('shutdown_delay', 30))
        self.last_active_time = time.time()
        self.countdown_remaining = 0
        self.warning_window = None
        
        # Main frame
        main_frame = ttk.Frame(self.idle_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(settings_frame, text="Idle Threshold (seconds):").grid(row=0, column=0, sticky=tk.W)
        ttk.Spinbox(settings_frame, from_=5, to=300, textvariable=self.idle_threshold, width=5,
                   command=lambda: self.update_config('idle_threshold', self.idle_threshold.get())).grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(settings_frame, text="Shutdown Delay (seconds):").grid(row=1, column=0, sticky=tk.W)
        ttk.Spinbox(settings_frame, from_=5, to=300, textvariable=self.shutdown_delay, width=5,
                   command=lambda: self.update_config('shutdown_delay', self.shutdown_delay.get())).grid(row=1, column=1, sticky=tk.W)
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(status_frame, height=8, state=tk.DISABLED, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_detection)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Set initial tab state based on config
        self.notebook.tab(1, state=tk.NORMAL if self.config.get('idle_detector_enabled', True) else tk.DISABLED)

    def update_config(self, key, value):
        """Update a config value and save to file"""
        self.config[key] = value
        self.save_config()

    def log_status(self, message):
        """Log messages to the status text box"""
        if hasattr(self, 'status_text'):
            self.status_text.config(state=tk.NORMAL)
            self.status_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.status_text.see(tk.END)
            self.status_text.config(state=tk.DISABLED)
        
    def get_idle_time(self):
        """Get the system idle time in seconds"""
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [
                ('cbSize', ctypes.c_uint),
                ('dwTime', ctypes.c_uint),
            ]
        
        lastInputInfo = LASTINPUTINFO()
        lastInputInfo.cbSize = ctypes.sizeof(lastInputInfo)
        ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lastInputInfo))
        
        millis = ctypes.windll.kernel32.GetTickCount() - lastInputInfo.dwTime
        return millis / 1000.0
    
    def update_countdown(self, window, countdown_label):
        """Update the countdown display every second"""
        if hasattr(self, 'countdown_remaining') and hasattr(self, 'warning_shown'):
            if self.countdown_remaining > 0 and self.warning_shown:
                countdown_label.config(text=f"Time remaining: {self.countdown_remaining} seconds")
                self.countdown_remaining -= 1
                window.after(1000, self.update_countdown, window, countdown_label)
    
    def hide_warning(self):
        """Hide the warning window if it exists"""
        if hasattr(self, 'warning_window') and self.warning_window:
            try:
                self.warning_window.destroy()
                self.log_status("Warning dismissed due to user activity")
            except:
                pass
            finally:
                self.warning_window = None
                self.warning_shown = False
    
    def show_warning(self):
        """Show a warning popup with an OK button and countdown"""
        self.warning_shown = True
        self.countdown_remaining = self.shutdown_delay.get()
        self.log_status("Idle detected! Showing warning...")
        
        # Create a top-level window for the warning
        self.warning_window = tk.Toplevel(self.root)
        self.warning_window.title("Idle Warning")
        self.warning_window.geometry("400x220")
        self.warning_window.resizable(False, False)
        
        # Make sure the warning window stays on top
        self.warning_window.attributes('-topmost', True)
        
        # Main message
        message = "Idle detected! This computer will shutdown soon if you don't interact."
        ttk.Label(self.warning_window, text=message, wraplength=380, padding=10).pack(pady=(10, 0))
        
        # Countdown display
        self.countdown_label = ttk.Label(self.warning_window, 
                                       text=f"Time remaining: {self.countdown_remaining} seconds", 
                                       font=('Arial', 10, 'bold'))
        self.countdown_label.pack(pady=5)
        
        # OK button
        ttk.Button(self.warning_window, text="OK", 
                  command=lambda: self.on_warning_response(self.warning_window)).pack(pady=10)
        
        # Start the countdown updates
        self.update_countdown(self.warning_window, self.countdown_label)
        
        # Schedule shutdown if no response
        self.warning_window.after(self.shutdown_delay.get() * 1000, self.shutdown_computer)
        
    def on_warning_response(self, window):
        """Handle user response to the warning"""
        self.hide_warning()
        self.last_active_time = time.time()
        self.log_status("User manually dismissed warning")
        
    def shutdown_computer(self):
        """Shutdown the computer"""
        if hasattr(self, 'warning_shown') and self.warning_shown:  # Only shutdown if warning is still active
            self.log_status("Shutting down computer...")
            os.system("shutdown /s /t 1")
        
    def detection_loop(self):
        """Main detection loop"""
        while hasattr(self, 'is_running') and self.is_running:
            idle_time = self.get_idle_time()
            
            if idle_time >= self.idle_threshold.get() and not self.warning_shown:
                self.root.after(0, self.show_warning)
            elif idle_time < 1:  # User was active
                if hasattr(self, 'warning_shown') and self.warning_shown:
                    self.root.after(0, self.hide_warning)
                self.last_active_time = time.time()
            
            time.sleep(1)
        
    def start_detection(self):
        """Start the idle detection"""
        if hasattr(self, 'is_running') and self.is_running:
            return
            
        self.is_running = True
        self.log_status(f"Idle detection running (Threshold: {self.idle_threshold.get()}s)")
        
        # Start detection in a separate thread
        self.detection_thread = Thread(target=self.detection_loop, daemon=True)
        self.detection_thread.start()
        
    def stop_detection(self):
        """Stop the idle detection"""
        if hasattr(self, 'is_running'):
            self.is_running = False
        if hasattr(self, 'warning_shown') and self.warning_shown:
            self.hide_warning()
        if hasattr(self, 'status_text'):
            self.log_status("Idle detection stopped")

    # ==============================================
    # VOLUME CONTROL TAB
    # ==============================================
    def create_volume_control_tab(self):
        """Create the volume control tab"""
        self.volume_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.volume_tab, text="Volume Control")
        
        # Main frame
        main_frame = ttk.Frame(self.volume_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(
            main_frame,
            text="Windows Volume Control",
            font=('Segoe UI', 12, 'bold')
        ).pack(pady=(0, 10))
        
        # Current volume display
        self.volume_label = ttk.Label(
            main_frame,
            text="Current Volume: Checking..." if self.config.get('volume_control_enabled', True) else "Volume Control Disabled",
            font=('Segoe UI', 10)
        )
        self.volume_label.pack(pady=5)
        
        # Saved volume display
        self.saved_volume_label = ttk.Label(
            main_frame,
            text=f"Saved Volume: {self.config.get('saved_volume', 50)}%",
            font=('Segoe UI', 10)
        )
        self.saved_volume_label.pack(pady=5)
        
        # Volume slider
        self.volume_slider = ttk.Scale(
            main_frame,
            from_=0,
            to=100,
            command=self.on_slider_move,
            state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED
        )
        self.volume_slider.pack(fill=tk.X, pady=10)
        
        # Quick set buttons
        quick_buttons_frame = ttk.Frame(main_frame)
        quick_buttons_frame.pack(pady=10)
        
        for percent in [15, 30, 50, 75]:
            btn = ttk.Button(
                quick_buttons_frame,
                text=f"{percent}%",
                command=lambda p=percent: self.set_volume(p),
                width=5,
                state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED
            )
            btn.pack(side=tk.LEFT, padx=5)
        
        # Custom volume entry
        custom_frame = ttk.Frame(main_frame)
        custom_frame.pack(pady=10)
        
        ttk.Label(custom_frame, text="Custom %:").pack(side=tk.LEFT)
        
        self.custom_entry = ttk.Entry(
            custom_frame,
            width=5,
            validate='key',
            validatecommand=(self.root.register(self.validate_percent), '%P'),
            state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED
        )
        self.custom_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            custom_frame,
            text="Set",
            command=self.set_custom_volume,
            width=5,
            state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED
        ).pack(side=tk.LEFT)
        
        # Hide on startup checkbox
        self.hide_var = tk.BooleanVar(value=self.config.get('hide_on_startup', False))
        hide_checkbox = ttk.Checkbutton(
            main_frame,
            text="Hide program when run",
            variable=self.hide_var,
            command=self.toggle_hide_setting
        )
        hide_checkbox.pack(pady=5)
        
        # Save current volume button
        ttk.Button(
            main_frame,
            text="Save Current Volume",
            command=self.save_current_volume,
            width=20,
            state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED
        ).pack(pady=5)
        
        # Show/Hide window button
        self.toggle_window_button = ttk.Button(
            main_frame,
            text="Show Window" if self.config.get('hide_on_startup', False) else "Hide Window",
            command=self.toggle_window_visibility,
            width=20
        )
        self.toggle_window_button.pack(pady=5)
        
        # Set initial tab state based on config
        self.notebook.tab(2, state=tk.NORMAL if self.config.get('volume_control_enabled', True) else tk.DISABLED)

    def init_volume_control(self):
        """Initialize the volume control interface"""
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize audio control:\n{str(e)}")
            return None

    def toggle_hide_setting(self):
        """Toggle the hide on startup setting"""
        self.config['hide_on_startup'] = self.hide_var.get()
        self.save_config()
        # Update the toggle button text
        if hasattr(self, 'toggle_window_button'):
            self.toggle_window_button.config(text="Show Window" if self.config['hide_on_startup'] else "Hide Window")

    def toggle_window_visibility(self):
        """Toggle window visibility"""
        if self.root.state() == 'withdrawn':
            self.root.deiconify()
            if hasattr(self, 'toggle_window_button'):
                self.toggle_window_button.config(text="Hide Window")
        else:
            self.root.withdraw()
            if hasattr(self, 'toggle_window_button'):
                self.toggle_window_button.config(text="Show Window")

    def validate_percent(self, text):
        """Validate percentage input"""
        if text == "":
            return True
        try:
            return 0 <= int(text) <= 100
        except ValueError:
            return False

    def set_custom_volume(self):
        """Set volume from custom entry"""
        try:
            percent = int(self.custom_entry.get())
            self.set_volume(percent)
        except ValueError:
            messagebox.showerror("Error", "Please enter a number 0-100")

    def save_current_volume(self):
        """Save the current volume as the preferred setting"""
        if not self.volume_control:
            return
            
        try:
            current_vol = self.volume_control.GetMasterVolumeLevelScalar()
            percent = int(current_vol * 100)
            self.config['saved_volume'] = percent
            self.save_config()
            if hasattr(self, 'saved_volume_label'):
                self.saved_volume_label.config(text=f"Saved Volume: {percent}%")
            messagebox.showinfo("Saved", f"Volume setting {percent}% has been saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save volume:\n{str(e)}")

    def on_slider_move(self, value):
        """Handle slider movement with debounce"""
        if not self.config.get('volume_control_enabled', True):
            return
            
        if hasattr(self, 'last_slider_update') and self.last_slider_update:
            self.root.after_cancel(self.last_slider_update)
        
        percent = float(value)
        self.last_slider_update = self.root.after(
            100, 
            lambda: self.set_volume(percent, update_slider=False)
        )

    def set_volume(self, percent, update_slider=True):
        """Set system volume with error handling"""
        if not self.volume_control or not self.config.get('volume_control_enabled', True):
            return
            
        try:
            self.ignore_volume_change = True
            percent = max(0, min(100, float(percent)))
            self.volume_control.SetMasterVolumeLevelScalar(percent/100, None)
            
            if update_slider and hasattr(self, 'volume_slider'):
                self.volume_slider.set(percent)
            self.update_current_volume()
            self.ignore_volume_change = False
        except Exception as e:
            self.ignore_volume_change = False
            messagebox.showerror("Error", f"Volume control failed:\n{str(e)}")

    def update_current_volume(self):
        """Update displayed volume"""
        if not self.volume_control or not self.config.get('volume_control_enabled', True):
            if hasattr(self, 'volume_label'):
                self.volume_label.config(text="Current Volume: Unavailable")
            return
            
        try:
            current_vol = self.volume_control.GetMasterVolumeLevelScalar()
            percent = int(current_vol * 100)
            if hasattr(self, 'volume_label'):
                self.volume_label.config(text=f"Current Volume: {percent}%")
            if hasattr(self, 'volume_slider'):
                self.volume_slider.set(percent)
        except:
            if hasattr(self, 'volume_label'):
                self.volume_label.config(text="Current Volume: Unknown")

    def monitor_volume_changes(self):
        """Check for volume changes and revert to saved setting if changed externally"""
        if not self.volume_control or not self.config.get('volume_control_enabled', True):
            return
            
        try:
            current_vol = self.volume_control.GetMasterVolumeLevelScalar()
            current_percent = int(current_vol * 100)
            
            if (not self.ignore_volume_change and 
                hasattr(self, 'config') and 
                self.config.get('saved_volume') is not None):
                if current_percent != self.config['saved_volume']:
                    self.set_volume(self.config['saved_volume'])
        except:
            pass
            
        # Check again after 1 second
        if hasattr(self, 'root'):
            self.root.after(1000, self.monitor_volume_changes)

    def on_close(self):
        """Clean up on window close"""
        if hasattr(self, 'last_slider_update') and self.last_slider_update:
            self.root.after_cancel(self.last_slider_update)
        if hasattr(self, 'is_running') and self.is_running:
            self.is_running = False
        if hasattr(self, 'warning_shown') and self.warning_shown:
            self.hide_warning()
        if hasattr(self, 'root'):
            self.root.destroy()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = SystemUtilitiesApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Fatal Error", f"Application failed to start:\n{str(e)}")
        sys.exit(1)
