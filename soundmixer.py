import sys
import json
import os
import logging
import threading
import time
import keyboard
from PyQt5.QtWidgets import (
    QApplication, QSystemTrayIcon, QMenu, QDialog, 
    QVBoxLayout, QLabel, QPushButton, QHBoxLayout,
    QMessageBox, QComboBox, QWidget, QProgressBar, 
    QAction, QStyle
)
from PyQt5.QtGui import QIcon, QKeySequence, QFont
from PyQt5.QtCore import Qt, QTimer
import win32gui
import win32process
import win32con
import pyautogui
import psutil

# Настройка логирования
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("sound_mixer.log"),
        logging.StreamHandler()
    ]
)

class ConfigManager:
    def __init__(self, config_path='config.json'):
        self.config_path = config_path
        self.default_config = {
            "hotkeys": {
                "volume_up": ["ctrl", "alt", "up"],
                "volume_down": ["ctrl", "alt", "down"],
                "mute": ["ctrl", "alt", "m"],
                "switch_app": ["ctrl", "alt", "tab"]
            },
            "gui": {
                "opacity": 0.85,
                "timeout": 2000
            }
        }
        self.config = self.load_config()
        logging.info("Config loaded")
    
    def load_config(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError as e:
                logging.error(f"Invalid JSON in config: {e}")
                self.create_default_config()
                return self.default_config
            except Exception as e:
                logging.error(f"Error loading config: {e}")
                return self.default_config
        else:
            self.create_default_config()
            return self.default_config
    
    def create_default_config(self):
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.default_config, f, indent=4)
            logging.info("Default config created")
        except Exception as e:
            logging.error(f"Error creating config: {e}")
    
    def save_config(self):
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config, f, indent=4)
            logging.info("Config saved")
        except Exception as e:
            logging.error(f"Error saving config: {e}")
    
    def get_hotkey(self, action):
        keys = self.config["hotkeys"].get(action, [])
        return keys
    
    def set_hotkey(self, action, keys):
        self.config["hotkeys"][action] = keys
        self.save_config()
    
    def get_gui_setting(self, setting):
        return self.config["gui"].get(setting, self.default_config["gui"][setting])

class AudioController:
    def __init__(self):
        self.initialize_audio()
        self.current_pid = None
        self.current_volume_control = None
    
    def initialize_audio(self):
        try:
            from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
            from comtypes import CLSCTX_ALL
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                ISimpleAudioVolume._iid_, 
                CLSCTX_ALL, 
                None
            )
            self.volume_interface = interface.QueryInterface(ISimpleAudioVolume)
            logging.info("Audio controller initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize audio controller: {e}")
            self.volume_interface = None
    
    def get_volume_control_for_app(self, pid):
        """Получаем интерфейс управления громкостью для конкретного приложения"""
        from pycaw.pycaw import AudioUtilities
        
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.ProcessId == pid:
                    return session.SimpleAudioVolume
            return None
        except Exception as e:
            logging.error(f"Error getting volume control for app: {e}")
            return None
    
    def get_active_app_pid(self):
        """Получаем PID активного приложения"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
        except Exception as e:
            logging.error(f"Error getting active app PID: {e}")
            return None
    
    def set_volume(self, level: float):
        """Установка громкости активного приложения"""
        try:
            pid = self.get_active_app_pid()
            if not pid:
                logging.warning("No active app PID found")
                return
                
            # Если приложение изменилось, получаем новый интерфейс
            if pid != self.current_pid:
                self.current_volume_control = self.get_volume_control_for_app(pid)
                self.current_pid = pid
                
            if self.current_volume_control:
                self.current_volume_control.SetMasterVolume(max(0.0, min(1.0, level)), None)
            elif self.volume_interface:
                # Если не нашли интерфейс для приложения, используем общий
                self.volume_interface.SetMasterVolume(max(0.0, min(1.0, level)), None)
        except Exception as e:
            logging.error(f"Error setting volume: {e}")
    
    def get_volume(self) -> float:
        """Получение текущей громкости активного приложения"""
        try:
            pid = self.get_active_app_pid()
            if not pid:
                return 0.5
                
            # Если приложение изменилось, получаем новый интерфейс
            if pid != self.current_pid:
                self.current_volume_control = self.get_volume_control_for_app(pid)
                self.current_pid = pid
                
            if self.current_volume_control:
                return self.current_volume_control.GetMasterVolume()
            elif self.volume_interface:
                return self.volume_interface.GetMasterVolume()
            return 0.5
        except Exception as e:
            logging.error(f"Error getting volume: {e}")
            return 0.5
    
    def toggle_mute(self):
        """Переключение режима без звука для активного приложения"""
        try:
            pid = self.get_active_app_pid()
            if not pid:
                return
                
            # Если приложение изменилось, получаем новый интерфейс
            if pid != self.current_pid:
                self.current_volume_control = self.get_volume_control_for_app(pid)
                self.current_pid = pid
                
            if self.current_volume_control:
                is_muted = self.current_volume_control.GetMute()
                self.current_volume_control.SetMute(not is_muted, None)
            elif self.volume_interface:
                is_muted = self.volume_interface.GetMute()
                self.volume_interface.SetMute(not is_muted, None)
        except Exception as e:
            logging.error(f"Error toggling mute: {e}")

class VolumeOverlay(QWidget):
    def __init__(self, audio_controller, config_manager):
        super().__init__()
        self.audio = audio_controller
        self.config = config_manager
        self.init_ui()
        self.update_info()
        
        self.hide_timer = QTimer(self)
        self.hide_timer.timeout.connect(self.hide)
        self.hide_timer.setSingleShot(True)
        
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_info)
        self.update_timer.start(500)
    
    def init_ui(self):
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool |
            Qt.X11BypassWindowManagerHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background: transparent;")
        
        container = QWidget(self)
        container.setStyleSheet("""
            background-color: rgba(30, 30, 40, 200);
            border-radius: 10px;
            padding: 15px;
        """)
        
        layout = QVBoxLayout(container)
        
        self.app_label = QLabel("Active App")
        self.app_label.setStyleSheet("color: white;")
        self.app_label.setFont(QFont("Arial", 10))
        self.app_label.setAlignment(Qt.AlignCenter)
        
        self.volume_bar = QProgressBar()
        self.volume_bar.setRange(0, 100)
        self.volume_bar.setTextVisible(False)
        self.volume_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #555;
                border-radius: 5px;
                background: #222;
            }
            QProgressBar::chunk {
                background: #4CAF50;
                border-radius: 4px;
            }
        """)
        
        self.volume_label = QLabel("100%")
        self.volume_label.setStyleSheet("color: white;")
        self.volume_label.setFont(QFont("Arial", 12, QFont.Bold))
        self.volume_label.setAlignment(Qt.AlignCenter)
        
        layout.addWidget(self.app_label)
        layout.addWidget(self.volume_bar)
        layout.addWidget(self.volume_label)
        
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(container)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        self.resize(300, 150)
        self.move_to_corner()
        self.hide()
    
    def move_to_corner(self):
        screen = QApplication.primaryScreen().geometry()
        self.move(screen.width() - self.width() - 20, 50)
    
    def update_info(self):
        try:
            # Получаем название активного приложения
            hwnd = win32gui.GetForegroundWindow()
            app_name = win32gui.GetWindowText(hwnd).strip()
            
            # Если название пустое, пытаемся получить имя процесса
            if not app_name:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    app_name = process.name()
                except:
                    app_name = "System"
            
            if not app_name:
                app_name = "System"
            
            if len(app_name) > 40:
                app_name = app_name[:37] + "..."
            
            self.app_label.setText(app_name)
        except Exception as e:
            logging.error(f"Error getting active app: {e}")
            self.app_label.setText("Unknown App")
        
        try:
            volume = int(self.audio.get_volume() * 100)
            self.volume_bar.setValue(volume)
            self.volume_label.setText(f"{volume}%")
        except Exception as e:
            logging.error(f"Error getting volume: {e}")
            self.volume_label.setText("Error")
    
    def show_overlay(self, timeout=None):
        if timeout is None:
            timeout = self.config.get_gui_setting("timeout")
        self.show()
        self.raise_()
        self.activateWindow()
        
        if self.hide_timer.isActive():
            self.hide_timer.stop()
        
        self.hide_timer.start(timeout)

class HotkeyManager:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.hotkeys = {}
        self.active = False
        self.callbacks = {
            "volume_up": lambda: None,
            "volume_down": lambda: None,
            "mute": lambda: None,
            "switch_app": lambda: None
        }
        self.load_hotkeys()
        logging.info("HotkeyManager initialized")

    def load_hotkeys(self):
        for action in self.callbacks.keys():
            keys = self.config_manager.get_hotkey(action)
            if keys:
                combo = '+'.join(keys)
                self.register_hotkey(combo, self.callbacks[action])
                logging.debug(f"Registered hotkey for {action}: {combo}")

    def register_hotkey(self, combo, callback):
        try:
            keyboard.add_hotkey(combo, callback)
            self.hotkeys[combo] = callback
            logging.debug(f"Registered hotkey combination: {combo}")
        except Exception as e:
            logging.error(f"Error registering hotkey {combo}: {e}")

    def update_hotkeys(self):
        keyboard.unhook_all_hotkeys()
        self.hotkeys.clear()
        self.load_hotkeys()
        logging.info("Hotkeys updated")

    def start(self):
        if self.active:
            return
            
        self.active = True
        logging.info("Starting hotkey listener")
        
        def run():
            logging.info("Hotkey listener thread started")
            keyboard.wait()
        
        self.thread = threading.Thread(target=run, daemon=True)
        self.thread.start()

    def stop(self):
        logging.info("Stopping hotkey listener")
        keyboard.unhook_all_hotkeys()
        self.active = False

    def set_callback(self, action, callback):
        if action in self.callbacks:
            self.callbacks[action] = callback
            logging.info(f"Callback set for {action}")
            self.update_hotkeys()

class HotkeySettingsDialog(QDialog):
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config = config_manager
        self.setWindowTitle("Настройка горячих клавиш")
        self.setWindowModality(Qt.ApplicationModal)
        self.setFixedSize(400, 300)
        
        layout = QVBoxLayout()
        
        actions = {
            "volume_up": "Увеличить громкость",
            "volume_down": "Уменьшить громкость",
            "mute": "Без звука",
            "switch_app": "Переключить приложение"
        }
        
        self.comboboxes = {}
        
        for action, description in actions.items():
            hbox = QHBoxLayout()
            hbox.addWidget(QLabel(description))
            
            combo = QComboBox()
            combo.setObjectName(action)
            combo.addItem("+".join(self.config.get_hotkey(action)))
            self.comboboxes[action] = combo
            hbox.addWidget(combo)
            
            btn = QPushButton("Изменить")
            btn.clicked.connect(lambda _, a=action: self.start_key_recording(a))
            hbox.addWidget(btn)
            
            layout.addLayout(hbox)
        
        btn_box = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_settings)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        btn_box.addWidget(save_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box)
        
        self.setLayout(layout)
        self.recording_action = None
        self.new_hotkey = None
    
    def start_key_recording(self, action):
        self.recording_action = action
        self.comboboxes[action].clear()
        self.comboboxes[action].addItem("Нажмите комбинацию...")
        self.comboboxes[action].setFocus()
        self.grabKeyboard()
    
    def keyPressEvent(self, event):
        if not self.recording_action:
            return super().keyPressEvent(event)
        
        if event.isAutoRepeat():
            return
        
        modifiers = []
        key = None
        
        if event.modifiers() & Qt.ControlModifier:
            modifiers.append("ctrl")
        if event.modifiers() & Qt.AltModifier:
            modifiers.append("alt")
        if event.modifiers() & Qt.ShiftModifier:
            modifiers.append("shift")
        
        if event.key() not in [Qt.Key_Control, Qt.Key_Alt, Qt.Key_Shift]:
            key = QKeySequence(event.key()).toString().lower()
        
        keys = modifiers
        if key:
            keys.append(key)
        
        if keys:
            combo_str = "+".join(keys)
            self.comboboxes[self.recording_action].clear()
            self.comboboxes[self.recording_action].addItem(combo_str)
            self.new_hotkey = keys
        
        self.releaseKeyboard()
        self.recording_action = None
    
    def save_settings(self):
        if not self.new_hotkey:
            QMessageBox.warning(self, "Ошибка", "Не выбрана новая комбинация клавиш")
            return
            
        for action, combo in self.comboboxes.items():
            if combo.count() > 0:
                keys = combo.currentText().split('+')
                self.config.set_hotkey(action, keys)
        
        QMessageBox.information(self, "Сохранено", 
                               "Настройки сохранены. Приложение будет перезапущено.")
        self.accept()
        QApplication.exit(199)

class SoundMixerApp:
    def __init__(self):
        self.config = ConfigManager()
        self.audio = AudioController()
        self.hotkey_manager = HotkeyManager(self.config)
        
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        
        self.gui = VolumeOverlay(self.audio, self.config)
        self.gui.setWindowOpacity(self.config.get_gui_setting("opacity"))
        
        self.hotkey_manager.set_callback("volume_up", self.volume_up)
        self.hotkey_manager.set_callback("volume_down", self.volume_down)
        self.hotkey_manager.set_callback("mute", self.toggle_mute)
        self.hotkey_manager.set_callback("switch_app", self.switch_app)
        
        self.tray_icon = QSystemTrayIcon()
        
        try:
            icon = self.app.style().standardIcon(QStyle.SP_MediaVolume)
            self.tray_icon.setIcon(icon)
            logging.info("Standard tray icon set")
        except Exception as e:
            logging.error(f"Error setting tray icon: {e}")
        
        tray_menu = QMenu()
        
        settings_action = QAction("Настройки горячих клавиш", self.app)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)
        
        exit_action = QAction("Выход", self.app)
        exit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        self.tray_icon.activated.connect(self.tray_icon_activated)
        
        self.add_test_hotkey()
        
        logging.info("SoundMixerApp initialized")
    
    def add_test_hotkey(self):
        try:
            keyboard.add_hotkey('ctrl+alt+t', self.test_hotkey)
            logging.info("Test hotkey registered: Ctrl+Alt+T")
        except Exception as e:
            logging.error(f"Error registering test hotkey: {e}")
    
    def test_hotkey(self):
        logging.info("TEST HOTKEY WORKED! Hotkey system is functional")
        self.safe_show_overlay()
        self.tray_icon.showMessage(
            "Тест горячих клавиш",
            "Горячие клавиши работают! Ctrl+Alt+T сработала.",
            QSystemTrayIcon.Information,
            3000
        )
    
    def safe_show_overlay(self):
        """Безопасный вызов show_overlay из главного потока"""
        if not self.gui.isVisible():
            self.gui.show_overlay()
    
    def volume_up(self):
        logging.info("Volume up triggered")
        try:
            current = self.audio.get_volume()
            new_volume = min(1.0, round(current + 0.05, 2))
            self.audio.set_volume(new_volume)
            QTimer.singleShot(0, self.safe_show_overlay)
        except Exception as e:
            logging.error(f"Error in volume_up: {e}")
    
    def volume_down(self):
        logging.info("Volume down triggered")
        try:
            current = self.audio.get_volume()
            new_volume = max(0.0, round(current - 0.05, 2))
            self.audio.set_volume(new_volume)
            QTimer.singleShot(0, self.safe_show_overlay)
        except Exception as e:
            logging.error(f"Error in volume_down: {e}")
    
    def toggle_mute(self):
        logging.info("Mute toggled")
        try:
            self.audio.toggle_mute()
            QTimer.singleShot(0, self.safe_show_overlay)
        except Exception as e:
            logging.error(f"Error in toggle_mute: {e}")
    
    def switch_app(self):
        logging.info("Switching application")
        try:
            pyautogui.hotkey('alt', 'tab')
            QTimer.singleShot(0, self.safe_show_overlay)
        except Exception as e:
            logging.error(f"Error in switch_app: {e}")
    
    def show_settings(self):
        dialog = HotkeySettingsDialog(self.config)
        if dialog.exec_() == QDialog.Accepted:
            os.execl(sys.executable, sys.executable, *sys.argv)
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.safe_show_overlay()
    
    def quit_app(self):
        self.hotkey_manager.stop()
        self.app.quit()
    
    def run(self):
        logging.info("Sound Mixer started")
        
        hotkeys_info = "\n".join([
            f"{action}: {'+'.join(keys)}"
            for action, keys in self.config.config['hotkeys'].items()
        ])
        logging.info(f"Active hotkeys:\n{hotkeys_info}")
        
        self.hotkey_manager.start()
        
        if not QSystemTrayIcon.isSystemTrayAvailable():
            logging.error("System tray is not available!")
            QMessageBox.critical(None, "Ошибка", 
                                "Системный трей недоступен. Приложение не может работать.")
            sys.exit(1)
        
        try:
            self.tray_icon.showMessage(
                "Sound Mixer", 
                "Приложение работает в фоновом режиме. Дважды щелкните по иконке, чтобы показать громкость.",
                QSystemTrayIcon.Information,
                3000
            )
        except Exception as e:
            logging.error(f"Tray showMessage error: {e}")
        
        sys.exit(self.app.exec_())

if __name__ == "__main__":
    restart_code = 199
    while True:
        app = QApplication(sys.argv)
        mixer = SoundMixerApp()
        exit_code = app.exec_()
        
        if exit_code != restart_code:
            break