import sys
import mido
import comtypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
from PySide6.QtCore import Qt, QTimer, Signal, Slot, QObject
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QSlider, QPushButton, 
                             QComboBox, QFrame, QDialog, QMessageBox, QSystemTrayIcon, 
                             QMenu, QTabWidget, QTextEdit)

class ChannelStrip(QWidget):
    def __init__(self, channel_idx, parent=None):
        super().__init__(parent)
        self.channel_idx = channel_idx
        self.session_idx = None
        self.is_muted = False
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        
        # App selector
        self.app_selector = QComboBox()
        self.app_selector.setFixedHeight(30)
        self.app_selector.setStyleSheet("""
            QComboBox {
                background-color: #333;
                color: white;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                background-color: #333;
                color: white;
                selection-background-color: #555;
            }
        """)
        
        # Volume label
        self.volume_label = QLabel("0%")
        self.volume_label.setAlignment(Qt.AlignCenter)
        self.volume_label.setStyleSheet("color: white; font-weight: bold;")
        
        # Fader
        self.fader = QSlider(Qt.Vertical)
        self.fader.setMinimum(0)
        self.fader.setMaximum(100)
        self.fader.setValue(0)
        self.fader.setFixedHeight(200)
        self.fader.setStyleSheet("""
            QSlider {
                background: transparent;
            }
            QSlider::groove:vertical {
                background: #444;
                width: 30px;
                border-radius: 4px;
            }
            QSlider::handle:vertical {
                background: #00a8ff;
                height: 20px;
                width: 40px;
                margin: 0 -5px;
                border-radius: 3px;
            }
        """)
        
        # Mute button
        self.mute_btn = QPushButton("M")
        self.mute_btn.setFixedSize(40, 40)
        self.mute_btn.setCheckable(True)
        self.mute_btn.setStyleSheet("""
            QPushButton {
                background-color: #333;
                color: white;
                border: 1px solid #555;
                border-radius: 20px;
                font-weight: bold;
            }
            QPushButton:checked {
                background-color: #bb0000;
                color: white;
            }
            QPushButton:hover {
                background-color: #444;
            }
        """)
        
        # Channel label
        self.channel_label = QLabel(f"CH {self.channel_idx + 1}")
        self.channel_label.setAlignment(Qt.AlignCenter)
        self.channel_label.setStyleSheet("color: #999;")
        
        # Add widgets to layout
        layout.addWidget(self.app_selector)
        layout.addWidget(self.volume_label)
        layout.addWidget(self.fader, 1)
        layout.addWidget(self.mute_btn)
        layout.addWidget(self.channel_label)
        
        # Connect signals
        self.fader.valueChanged.connect(self.on_fader_value_changed)
        self.mute_btn.clicked.connect(self.on_mute_clicked)
        
    def on_fader_value_changed(self, value):
        self.volume_label.setText(f"{value}%")
        
    def on_mute_clicked(self):
        self.is_muted = self.mute_btn.isChecked()
        
    def set_fader_value(self, value):
        """Set fader value from MIDI (0.0-1.0)"""
        value_percent = int(value * 100)
        self.fader.blockSignals(True)
        self.fader.setValue(value_percent)
        self.fader.blockSignals(False)
        self.volume_label.setText(f"{value_percent}%")
        
    def set_app_options(self, apps):
        """Set available apps in the selector"""
        self.app_selector.clear()
        self.app_selector.addItem("Not Assigned", None)
        for idx, name in apps:
            self.app_selector.addItem(name, idx)
            
    def get_selected_session(self):
        """Get the selected session index"""
        return self.app_selector.currentData()
        
    def set_mute_state(self, muted):
        """Set mute button state"""
        self.is_muted = muted
        self.mute_btn.blockSignals(True)
        self.mute_btn.setChecked(muted)
        self.mute_btn.blockSignals(False)

mido.set_backend('mido.backends.rtmidi')

class MIDIHandler(QObject):
    # Signals
    fader_moved = Signal(int, float)  # channel, value
    button_pressed = Signal(int, bool)  # channel, state (True=pressed)
    knob_turned = Signal(int, float)  # channel, value
    raw_message_received = Signal(str)  # raw MIDI message
    
    def __init__(self):
        super().__init__()
        self.midi_in = None
        self.midi_out = None
        self.device_name = ""
        self.fader_values = [0.0] * 8
        self.button_states = [False] * 24
        self.knob_values = [0.0] * 8
        
    def get_available_devices(self):
        """Get lists of available MIDI input and output devices"""
        try:
            mido.set_backend('mido.backends.rtmidi')  # Ensure the correct backend is set
            input_devices = mido.get_input_names()
            output_devices = mido.get_output_names()
            return input_devices, output_devices
        except Exception as e:
            print(f"Error getting MIDI devices: {e}")
            return [], []
    
    def connect_device(self, input_device_name, output_device_name):
        """Connect to specific MIDI devices by name"""
        try:
            # Close existing connections if any
            self.close()
            
            self.midi_in = mido.open_input(input_device_name)
            self.midi_out = mido.open_output(output_device_name)
            self.device_name = input_device_name
            return True
        except Exception as e:
            print(f"Error connecting to MIDI device: {e}")
            return False
    
    def close(self):
        """Close MIDI connections"""
        if self.midi_in:
            self.midi_in.close()
            self.midi_in = None
        if self.midi_out:
            self.midi_out.close()
            self.midi_out = None
    
    def poll_messages(self):
        """Check for new MIDI messages"""
        if not self.midi_in:
            return
            
        for msg in self.midi_in.iter_pending():
            if msg.type == 'control_change':
                self.process_control_change(msg)
            # Emit raw message for debug tab
            self.raw_message_received.emit(str(msg))
    
    def process_control_change(self, msg):
        """Process MIDI control change messages"""
        # Faders are CC 0-7
        if 0 <= msg.control <= 7:
            channel = msg.control
            value = msg.value / 127.0
            self.fader_values[channel] = value
            self.fader_moved.emit(channel, value)
            
        # Knobs are CC 16-23
        elif 16 <= msg.control <= 23:
            channel = msg.control - 16
            value = msg.value / 127.0
            self.knob_values[channel] = value
            self.knob_turned.emit(channel, value)
            
        # S buttons (mute) are CC 32-39
        elif 32 <= msg.control <= 39:
            channel = msg.control - 32
            state = msg.value >= 64  # Treat values >= 64 as pressed
            
            # Only emit if state changed (handles button press and release properly)
            if self.button_states[channel] != state:
                self.button_states[channel] = state
                self.button_pressed.emit(channel, state)
    
    def send_led_feedback(self, button_idx, state):
        """Send LED feedback to the controller for button states"""
        if not self.midi_out:
            return
            
        # Map button index to CC number (S buttons are 32-39)
        cc = 32 + button_idx
        value = 127 if state else 0
        msg = mido.Message('control_change', control=cc, value=value)
        self.midi_out.send(msg)


class WindowsAudioMixer:
    def __init__(self):
        self.sessions = []
        self.update_sessions()
        
        # Get master volume controller
        self.master_devices = AudioUtilities.GetSpeakers()
        self.master_interface = self.master_devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.master_volume = cast(self.master_interface, POINTER(IAudioEndpointVolume))
        
    def update_sessions(self):
        """Update the list of audio sessions"""
        self.sessions = AudioUtilities.GetAllSessions()
        app_list = []
        
        # Add master volume as first option
        app_list.append((-1, "Master Volume"))
        
        # Add application sessions
        for i, session in enumerate(self.sessions):
            if session.Process and session.Process.name():
                app_list.append((i, session.Process.name()))
        
        return app_list
        
    def set_volume(self, session_idx, volume):
        """Set volume for a specific session"""
        try:
            # Handle master volume separately
            if session_idx == -1:
                self.master_volume.SetMasterVolumeLevelScalar(volume, None)
                return
                
            # Handle app volume
            if 0 <= session_idx < len(self.sessions):
                session = self.sessions[session_idx]
                volume_interface = session._ctl.QueryInterface(ISimpleAudioVolume)
                volume_interface.SetMasterVolume(volume, None)
        except Exception as e:
            print(f"Error setting volume: {e}")
    
    def toggle_mute(self, session_idx):
        """Toggle mute state for a specific session"""
        try:
            # Handle master volume separately
            if session_idx == -1:
                current_mute = self.master_volume.GetMute()
                self.master_volume.SetMute(not current_mute, None)
                return not current_mute
                
            # Handle app mute
            if 0 <= session_idx < len(self.sessions):
                session = self.sessions[session_idx]
                volume_interface = session._ctl.QueryInterface(ISimpleAudioVolume)
                current_mute = volume_interface.GetMute()
                volume_interface.SetMute(not current_mute, None)
                return not current_mute
        except Exception as e:
            print(f"Error toggling mute: {e}")
        return False
        
    def is_muted(self, session_idx):
        """Check if session is muted"""
        try:
            # Handle master volume separately
            if session_idx == -1:
                return self.master_volume.GetMute()
                
            # Handle app mute
            if 0 <= session_idx < len(self.sessions):
                session = self.sessions[session_idx]
                volume_interface = session._ctl.QueryInterface(ISimpleAudioVolume)
                return volume_interface.GetMute()
        except Exception as e:
            print(f"Error checking mute status: {e}")
        return False
        
    def get_volume(self, session_idx):
        """Get volume for a specific session"""
        try:
            # Handle master volume separately
            if session_idx == -1:
                return self.master_volume.GetMasterVolumeLevelScalar()
                
            # Handle app volume
            if 0 <= session_idx < len(self.sessions):
                session = self.sessions[session_idx]
                volume_interface = session._ctl.QueryInterface(ISimpleAudioVolume)
                return volume_interface.GetMasterVolume()
        except Exception as e:
            print(f"Error getting volume: {e}")
        return 0.0


class DeviceSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("MIDI Device Selection")
        self.setMinimumWidth(400)
        self.setup_ui()
        
        # Get available devices
        self.midi_handler = MIDIHandler()
        self.refresh_device_lists()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("Select MIDI Device")
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: white; margin-bottom: 15px;")
        
        # Input device selection
        input_label = QLabel("Input Device:")
        input_label.setStyleSheet("color: white;")
        self.input_combo = QComboBox()
        self.input_combo.setStyleSheet("""
            QComboBox {
                background-color: #333;
                color: white;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 6px;
            }
            QComboBox QAbstractItemView {
                background-color: #333;
                color: white;
                selection-background-color: #555;
            }
        """)
        
        # Output device selection
        output_label = QLabel("Output Device:")
        output_label.setStyleSheet("color: white;")
        self.output_combo = QComboBox()
        self.output_combo.setStyleSheet("""
            QComboBox {
                background-color: #333;
                color: white;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 6px;
            }
            QComboBox QAbstractItemView {
                background-color: #333;
                color: white;
                selection-background-color: #555;
            }
        """)
        
        # Auto-select nanokontrol2 checkbox
        self.auto_select_box = QPushButton("Auto-detect nanoKONTROL2")
        self.auto_select_box.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        self.auto_select_box.clicked.connect(self.auto_select_nanokontrol)
        
        # Refresh button
        refresh_btn = QPushButton("Refresh Device List")
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)
        refresh_btn.clicked.connect(self.refresh_device_lists)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        connect_btn = QPushButton("Connect")
        connect_btn.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        connect_btn.clicked.connect(self.accept)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #7f8c8d;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #95a5a6;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(connect_btn)
        button_layout.addWidget(cancel_btn)
        
        # Add widgets to main layout
        layout.addWidget(title_label)
        layout.addWidget(input_label)
        layout.addWidget(self.input_combo)
        layout.addWidget(output_label)
        layout.addWidget(self.output_combo)
        layout.addWidget(self.auto_select_box)
        layout.addWidget(refresh_btn)
        layout.addStretch()
        layout.addLayout(button_layout)
        
        # Set dialog style
        self.setStyleSheet("background-color: #1e1e1e;")
        
    def refresh_device_lists(self):
        """Refresh the lists of available MIDI devices"""
        input_devices, output_devices = self.midi_handler.get_available_devices()
        
        self.input_combo.clear()
        self.output_combo.clear()
        
        for device in input_devices:
            self.input_combo.addItem(device)
            
        for device in output_devices:
            self.output_combo.addItem(device)
            
        # Auto-select nanoKONTROL2 if available
        self.auto_select_nanokontrol()
    
    def auto_select_nanokontrol(self):
        """Automatically select nanoKONTROL2 devices if available"""
        nano_name = "nanoKONTROL2"
        
        # Look for nanoKONTROL2 in input devices
        for i in range(self.input_combo.count()):
            if nano_name in self.input_combo.itemText(i):
                self.input_combo.setCurrentIndex(i)
                break
                
        # Look for nanoKONTROL2 in output devices
        for i in range(self.output_combo.count()):
            if nano_name in self.output_combo.itemText(i):
                self.output_combo.setCurrentIndex(i)
                break
                
    def get_selected_devices(self):
        """Get the selected input and output devices"""
        input_device = self.input_combo.currentText()
        output_device = self.output_combo.currentText()
        return input_device, output_device


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MIDI Mixer Control")
        self.setMinimumSize(800, 500)
        
        # Initialize MIDI and Audio handlers
        self.midi_handler = MIDIHandler()
        self.audio_mixer = WindowsAudioMixer()
        
        # Setup UI after initializing handlers
        self.setup_ui()
        
        # Setup system tray
        self.setup_tray_icon()
        
        # Connect signals
        self.midi_handler.fader_moved.connect(self.on_midi_fader_moved)
        self.midi_handler.button_pressed.connect(self.on_midi_button_pressed)
        self.midi_handler.raw_message_received.connect(self.on_raw_message_received)
        
        # Update timer
        self.update_timer = QTimer(self)
        self.update_timer.setInterval(50)  # 50ms interval
        self.update_timer.timeout.connect(self.update_loop)
        
        # Show device selection dialog
        self.show_device_selection()
    
    def setup_ui(self):
        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        
        # Create tab widget
        self.tabs = QTabWidget()
        
        # Main tab
        self.main_tab = QWidget()
        self.setup_main_tab()
        
        # Debug tab
        self.debug_tab = QWidget()
        self.setup_debug_tab()
        
        self.tabs.addTab(self.main_tab, "Main")
        self.tabs.addTab(self.debug_tab, "Debug")
        
        main_layout.addWidget(self.tabs)
        self.setCentralWidget(main_widget)
        
        # Set style
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1e1e1e;
                color: white;
            }
        """)
    
    def setup_main_tab(self):
        layout = QVBoxLayout(self.main_tab)
        
        # Header area
        header_layout = QHBoxLayout()
        
        # Title
        title_label = QLabel("MIDI Mixer Control")
        title_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setStyleSheet("color: white; margin: 10px;")
        
        # Device info and change button
        self.device_info_label = QLabel("No device connected")
        self.device_info_label.setStyleSheet("color: #999;")
        
        self.change_device_btn = QPushButton("Change Device")
        self.change_device_btn.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        self.change_device_btn.clicked.connect(self.show_device_selection)
        
        header_layout.addWidget(title_label, 1)
        header_layout.addWidget(self.device_info_label)
        header_layout.addWidget(self.change_device_btn)
        
        # Status bar
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.StyledPanel)
        status_frame.setStyleSheet("background-color: #222; border-radius: 4px;")
        status_layout = QHBoxLayout(status_frame)
        
        self.status_label = QLabel("Disconnected")
        self.status_label.setStyleSheet("color: #999;")
        
        self.refresh_btn = QPushButton("Refresh Apps")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: #3498db;
            }
        """)
        self.refresh_btn.clicked.connect(self.update_app_list)
        
        status_layout.addWidget(self.status_label)
        status_layout.addStretch(1)
        status_layout.addWidget(self.refresh_btn)
        
        # Channel strips layout
        strips_layout = QHBoxLayout()
        strips_layout.setSpacing(10)
        
        # Create 8 channel strips
        self.channel_strips = []
        for i in range(8):
            strip = ChannelStrip(i)
            strips_layout.addWidget(strip)
            self.channel_strips.append(strip)
        
        # Add everything to main layout
        layout.addLayout(header_layout)
        layout.addWidget(status_frame)
        layout.addLayout(strips_layout, 1)
    
    def setup_debug_tab(self):
        layout = QVBoxLayout(self.debug_tab)
        
        self.debug_text = QTextEdit()
        self.debug_text.setReadOnly(True)
        self.debug_text.setStyleSheet("""
            QTextEdit {
                background-color: #222;
                color: #0f0;
                font-family: Consolas, monospace;
            }
        """)
        
        layout.addWidget(self.debug_text)
    
    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.png"))  # Replace with your icon
        
        tray_menu = QMenu()
        show_action = tray_menu.addAction("Show")
        show_action.triggered.connect(self.show_normal)
        
        exit_action = tray_menu.addAction("Exit")
        exit_action.triggered.connect(self.close)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        # Connect tray icon click
        self.tray_icon.activated.connect(self.tray_icon_clicked)
    
    def tray_icon_clicked(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_normal()
    
    def show_normal(self):
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()
    
    def closeEvent(self, event):
        """Handle window close event"""
        # Minimize to tray instead of closing
        if self.tray_icon.isVisible():
            self.hide()
            event.ignore()
        else:
            # Clean up resources
            self.update_timer.stop()
            self.midi_handler.close()
            super().closeEvent(event)
    
    def show_device_selection(self):
        """Show device selection dialog"""
        # Stop the update timer if it's running
        if hasattr(self, 'update_timer') and self.update_timer.isActive():
            self.update_timer.stop()
        
        # Close existing MIDI connection
        self.midi_handler.close()
        
        # Show device selection dialog
        dialog = DeviceSelectionDialog(self)
        if dialog.exec():
            input_device, output_device = dialog.get_selected_devices()
            
            # Connect to selected devices
            if self.midi_handler.connect_device(input_device, output_device):
                self.status_label.setText(f"Connected to MIDI device")
                self.status_label.setStyleSheet("color: #00ff00;")
                self.device_info_label.setText(f"Device: {input_device}")
                self.update_timer.start()
                self.update_app_list()
            else:
                self.status_label.setText("Failed to connect to MIDI device")
                self.status_label.setStyleSheet("color: #ff0000;")
                QMessageBox.warning(self, "Connection Error", 
                                   "Failed to connect to the selected MIDI device. Please try again.")
    
    def update_app_list(self):
        """Update the list of audio applications"""
        apps = self.audio_mixer.update_sessions()
        for strip in self.channel_strips:
            current_app = strip.get_selected_session()
            strip.set_app_options(apps)
            
            # Try to restore previous selection if it still exists
            if current_app is not None:
                index = strip.app_selector.findData(current_app)
                if index >= 0:
                    strip.app_selector.setCurrentIndex(index)
    
    def update_loop(self):
        """Main update loop"""
        # Poll MIDI messages
        self.midi_handler.poll_messages()
        
        # Update UI to reflect audio state
        for strip in self.channel_strips:
            session_idx = strip.get_selected_session()
            if session_idx is not None:
                # Update mute state from actual audio state
                muted = self.audio_mixer.is_muted(session_idx)
                if muted != strip.is_muted:
                    strip.set_mute_state(muted)
    
    def on_midi_fader_moved(self, channel, value):
        """Handle MIDI fader movement"""
        if 0 <= channel < len(self.channel_strips):
            strip = self.channel_strips[channel]
            strip.set_fader_value(value)
            
            # Update audio if assigned
            session_idx = strip.get_selected_session()
            if session_idx is not None:
                self.audio_mixer.set_volume(session_idx, value)
    
    def on_midi_button_pressed(self, channel, state):
        """Handle MIDI button press (specifically the mute buttons)"""
        if 0 <= channel < len(self.channel_strips):
            strip = self.channel_strips[channel]
            session_idx = strip.get_selected_session()
            
            if session_idx is not None:
                # Toggle mute state
                new_mute_state = self.audio_mixer.toggle_mute(session_idx)
                strip.set_mute_state(new_mute_state)
                # Send LED feedback to controller
                self.midi_handler.send_led_feedback(channel, new_mute_state)
    
    def on_raw_message_received(self, message):
        """Handle raw MIDI messages for debug tab"""
        self.debug_text.append(message)
        # Keep only the last 100 lines to prevent memory issues
        cursor = self.debug_text.textCursor()
        cursor.movePosition(cursor.Start)
        cursor.movePosition(cursor.Down, cursor.KeepAnchor, 100)
        cursor.removeSelectedText()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Ensure the application doesn't quit when last window is closed
    app.setQuitOnLastWindowClosed(False)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())