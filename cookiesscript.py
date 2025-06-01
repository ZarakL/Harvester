import time, subprocess, random, os, json, re, sys
from pathlib import Path
from itertools import cycle, product
from screeninfo import get_monitors
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import fix_paths

try:
    import patch
except ImportError:
    # Patch module is optional
    pass

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# ─── CONSTANT PATHS ─────────────────────────────────────────────
# Browser executables (standard installation paths, can be overridden by config)
BRAVE_EXE   = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
CHROME_EXE  = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
OPERAGX_EXE = r"C:\Users\zarak\AppData\Local\Programs\Opera GX\opera.exe"  # Not used by default
VIVALDI_EXE = r"C:\Users\zarak\AppData\Local\Vivaldi\Application\vivaldi.exe"  # Not used by default

# WebDrivers (paths will be set at runtime)
CHROMEDRIVER = None  # Will be set at runtime
CHROMEDRIVER_133 = None  # Not currently used
CHROMEDRIVER_134 = None  # Not currently used
OPERADRIVER = None  # Not currently used

# Extension path (will be set at runtime)
EXT_PATH = None  # Will be set to local path at runtime
EXT_ID = None    # Will be determined dynamically from the extension folder

# Profile configuration - paths relative to the executable location
script_dir = Path(fix_paths.get_absolute_path("."))
PERSISTENT_PROFILES_BASE = script_dir / "profiles"
PROFILES_CONFIG_FILE = script_dir / "profiles_config.json"

# Browser configuration
BROWSERS = {
    "brave": {
        "exe": BRAVE_EXE,
        "name": "Brave",
        "profile_base": PERSISTENT_PROFILES_BASE / "brave_profiles",
        "kill_cmd": "taskkill /IM brave.exe /F /T",
        "prefs_path": "Default/Preferences",
        "ext_prefs_path": "Default/Secure Preferences",
        "version_cmd": r'powershell -command "&{(Get-Item \'' + BRAVE_EXE + r'\').VersionInfo.FileVersion}"',
        "driver_path": CHROMEDRIVER,
        "driver_type": "chrome",
        "enabled": True
    },
    "chrome": {
        "exe": CHROME_EXE,
        "name": "Chrome", 
        "profile_base": PERSISTENT_PROFILES_BASE / "chrome_profiles",
        "kill_cmd": "taskkill /IM chrome.exe /F /T",
        "prefs_path": "Default/Preferences",
        "ext_prefs_path": "Default/Secure Preferences",
        "version_cmd": r'powershell -command "&{(Get-Item \'' + CHROME_EXE + r'\').VersionInfo.FileVersion}"',
        "driver_path": CHROMEDRIVER,
        "driver_type": "chrome",
        "enabled": True
    },
    "vivaldi": {
        "exe": VIVALDI_EXE,
        "name": "Vivaldi",
        "profile_base": PERSISTENT_PROFILES_BASE / "vivaldi_profiles",
        "kill_cmd": "taskkill /IM vivaldi.exe /F /T",
        "prefs_path": "Default/Preferences",
        "ext_prefs_path": "Default/Secure Preferences",
        "version_cmd": r'powershell -command "&{(Get-Item \'' + VIVALDI_EXE + r'\').VersionInfo.FileVersion}"',
        "driver_path": CHROMEDRIVER_133,  # Using v133 driver for Vivaldi v7.3
        "driver_type": "chrome",
        "enabled": False  # Disabled for now
    },
    "operagx": {
        "exe": OPERAGX_EXE,
        "name": "OperaGX",
        "profile_base": PERSISTENT_PROFILES_BASE / "operagx_profiles",
        "kill_cmd": "taskkill /IM opera.exe /F /T",
        "prefs_path": "Default/Preferences",
        "ext_prefs_path": "Default/Secure Preferences",
        "version_cmd": r'powershell -command "&{(Get-Item \'' + OPERAGX_EXE + r'\').VersionInfo.FileVersion}"',
        "driver_path": CHROMEDRIVER_133,
        "driver_type": "chrome",  # Use chrome driver type since we're using ChromeDriver
        "enabled": False  # Disabled Opera - having issues with persistent profiles
    }
}

PRODUCT_URL = (
    "https://www.target.com/p/pok-233-mon-trading-card-game-zapdos-ex-deluxe-battle-deck/-/A-91351689#lnk=sametab"
)

# ─── CONSTANTS / LABELS ─────────────────────────────────────────
HOTKEY        = "%0"                           # Alt+0
CLEAR_LABEL   = "Clear Harvested Data"
TOGGLE_REGEX  = "^(Start|Stop) Harvesting$"
TARGET_PROXY  = "Local (No Proxy List)"        # Default value, will be overridden by user input
ACTIVE_TIME   = 120                            # 2 minutes (120 seconds)

# ─── EXTENSION SETUP ─────────────────────────────────────────────
BROWSER_LOAD_WAIT = 2                          # Seconds to wait for browser to fully load
MANUAL_CONFIG_WAIT = 300                       # Seconds to wait for manual configuration (5 minutes)

# Monitor configuration - Can be customized
USE_SECOND_MONITOR = True                      # Set to False to use primary monitor
MONITOR_LEFT_OFFSET = 0                        # Left offset for window placement
MONITOR_TOP_OFFSET = 0                         # Top offset for window placement
WINDOW_WIDTH = 600                             # Width of browser windows
WINDOW_HEIGHT = 900                            # Height of browser windows
WINDOW_GAP = 40                                # Gap between windows
# ────────────────────────────────────────────────────────────────


# ─── MONITOR DETECTION ────────────────────────────────────────────
def get_monitor_origin():
    """Determine window placement based on monitor configuration"""
    monitors = get_monitors()
    if not monitors:
        raise RuntimeError("No monitors detected.")
    
    if not USE_SECOND_MONITOR or len(monitors) == 1:
        # Use primary monitor or single monitor
        for mon in monitors:
            if getattr(mon, "is_primary", False) or len(monitors) == 1:
                return mon.x + MONITOR_LEFT_OFFSET, mon.y + MONITOR_TOP_OFFSET
    else:
        # Try to find monitor above primary
        mains = [m for m in monitors if getattr(m, "is_primary", False)]
        if not mains:
            # No primary monitor found, use first non-primary
            for mon in monitors:
                if not getattr(mon, "is_primary", False):
                    return mon.x + MONITOR_LEFT_OFFSET, mon.y + MONITOR_TOP_OFFSET
        
        main = mains[0]
        # Find monitor above primary
        for mon in monitors:
            if mon is main:
                continue
            if mon.y + mon.height <= main.y:  # Monitor is above primary
                return mon.x + MONITOR_LEFT_OFFSET, mon.y + MONITOR_TOP_OFFSET
        
        # No monitor above primary, use first non-primary
        for mon in monitors:
            if not getattr(mon, "is_primary", False):
                return mon.x + MONITOR_LEFT_OFFSET, mon.y + MONITOR_TOP_OFFSET
    
    # Fallback to primary monitor
    for mon in monitors:
        if getattr(mon, "is_primary", False):
            return mon.x + MONITOR_LEFT_OFFSET, mon.y + MONITOR_TOP_OFFSET
    
    # Last resort - use first monitor
    return monitors[0].x + MONITOR_LEFT_OFFSET, monitors[0].y + MONITOR_TOP_OFFSET

# Get monitor coordinates
X0, Y0 = get_monitor_origin()
RECTS = [
    (X0, Y0, WINDOW_WIDTH, WINDOW_HEIGHT),                      # left tile
    (X0 + WINDOW_WIDTH + WINDOW_GAP, Y0, WINDOW_WIDTH, WINDOW_HEIGHT),  # right tile
]
# ────────────────────────────────────────────────────────────────


# ─── PROFILE MANAGEMENT ─────────────────────────────────────────────
def load_profiles_config():
    """Load profiles configuration from JSON file"""
    if PROFILES_CONFIG_FILE.exists():
        try:
            with open(PROFILES_CONFIG_FILE, 'r') as f:
                config = json.load(f)
            print(f"  Loaded profiles configuration from {PROFILES_CONFIG_FILE}")
            return config
        except Exception as e:
            print(f"  Error loading profiles configuration: {e}")
    
    # Default config if file doesn't exist or there's an error
    print(f"  Creating new profiles configuration")
    return {
        "configured": False,
        "browser_profiles": {}
    }


def save_profiles_config(config):
    """Save profiles configuration to JSON file"""
    try:
        # Create parent directory if it doesn't exist
        PROFILES_CONFIG_FILE.parent.mkdir(exist_ok=True, parents=True)
        
        with open(PROFILES_CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
        print(f"  Saved profiles configuration to {PROFILES_CONFIG_FILE}")
        return True
    except Exception as e:
        print(f"  Error saving profiles configuration: {e}")
        return False


def generate_profile_path(browser_name, profile_number=1):
    """Generate a persistent profile path for a browser"""
    browser = BROWSERS[browser_name]
    profile_base = browser["profile_base"]
    
    # Create base directory if it doesn't exist
    Path(profile_base).mkdir(exist_ok=True, parents=True)
    
    # Create a unique profile directory with a simple numbering system
    profile_dir = Path(profile_base) / f"profile_{profile_number}"
    
    return profile_dir


def create_persistent_profile(browser_name, profile_number=1, force_new=False):
    """Create or load a persistent profile directory for the browser"""
    profile_dir = generate_profile_path(browser_name, profile_number)
    
    # Check if profile already exists
    if profile_dir.exists() and not force_new:
        print(f"     Using existing profile at {profile_dir}")
    else:
        # Create new profile
        profile_dir.mkdir(exist_ok=True, parents=True)
        print(f"     Created new persistent profile at {profile_dir}")
        
        # Create Default directory structure
        default_dir = profile_dir / "Default"
        default_dir.mkdir(exist_ok=True)
        
        # Set up extension keyboard shortcut in preferences
        configure_extension_shortcuts(profile_dir, browser_name)
    
    return profile_dir


def get_configured_profiles():
    """Get list of configured profiles from config file"""
    config = load_profiles_config()
    return config.get("browser_profiles", {})


def wait_for_manual_configuration(driver, browser_name, profile_number, profile_dir):
    """Wait for user to manually configure extension shortcuts and settings"""
    global EXT_ID
    import sys
    import select
    import os
    
    print(f"  Waiting for manual configuration of {BROWSERS[browser_name]['name']} profile {profile_number}...")
    print(f"  Please set up the following:")
    print(f"   1. Configure Alt+0 keyboard shortcut for the Refract extension")
    print(f"   2. Ensure extension is pinned to toolbar")
    print(f"   3. Configure any other settings as needed")
    print(f"   You have {MANUAL_CONFIG_WAIT // 60} minutes to complete the configuration.")
    print(f"   Press Enter in the terminal when you're done, or wait for the timeout.")
    
    # Navigate to extensions shortcuts page
    driver.get("chrome://extensions/shortcuts")
    
    # Wait for user input or timeout
    timeout = time.time() + MANUAL_CONFIG_WAIT
    
    # Store profile information in config
    config = load_profiles_config()
    if "browser_profiles" not in config:
        config["browser_profiles"] = {}
    
    if browser_name not in config["browser_profiles"]:
        config["browser_profiles"][browser_name] = {}
    
    # Add profile information
    config["browser_profiles"][browser_name][str(profile_number)] = {
        "path": str(profile_dir),
        "configured": True,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    
    save_profiles_config(config)
    
    # Print important profile information
    print(f"  IMPORTANT: Profile path for {BROWSERS[browser_name]['name']} profile {profile_number}:")
    print(f"   {profile_dir}")
    
    # Wait for user input or timeout with properly detecting Enter key press
    print(f"  Configuration in progress. Press Enter when complete or wait for timeout...")
    
    # Different approach for Windows (WSL) vs native Linux for input detection
    is_windows = os.name == 'nt' or 'microsoft' in sys.platform.lower()
    
    if is_windows or 'microsoft' in sys.platform.lower():  # WSL environment
        # For Windows/WSL we'll use a simple blocking input with timeout check
        import threading
        user_input_received = False
        
        def check_input():
            nonlocal user_input_received
            input()  # This will block until the user presses Enter
            user_input_received = True
        
        # Start a thread to wait for user input
        input_thread = threading.Thread(target=check_input)
        input_thread.daemon = True  # Daemon thread will be killed when main thread exits
        input_thread.start()
        
        # Wait for either user input or timeout
        while not user_input_received and time.time() < timeout:
            # Check if browser is still active
            try:
                if not driver.window_handles:
                    print(f"  Browser window closed, assuming configuration is complete")
                    break
            except Exception:
                print(f"  Browser connection lost, assuming configuration is complete")
                break
                
            time.sleep(1)  # Check every second
            
        if user_input_received:
            print(f"   User confirmed configuration is complete!")
    else:
        # Unix-like systems can use select for non-blocking input
        while time.time() < timeout:
            # Check if there's input ready to be read
            rlist, _, _ = select.select([sys.stdin], [], [], 1.0)
            
            if rlist:
                # User pressed Enter, read the line and continue
                sys.stdin.readline()
                print(f"   User confirmed configuration is complete!")
                break
                
            # Also check if browser is still active
            try:
                if not driver.window_handles:
                    print(f"  Browser window closed, assuming configuration is complete")
                    break
            except Exception:
                print(f"  Browser connection lost, assuming configuration is complete")
                break
    
    print(f" ️ Configuration period ended for {BROWSERS[browser_name]['name']} profile {profile_number}")
    
    # Make sure to properly close the browser after configuration
    try:
        driver.quit()
        print(f"   Browser closed after configuration")
    except Exception:
        pass
        
    return True


# ─── UTILITY FUNCTIONS ─────────────────────────────────────────────
def get_proxy_input():
    """
    Get user input for the proxy to use

    Returns:
        str: The proxy name selected by the user
    """
    print("\n  Proxy Selection")
    print("=" * 50)
    print("Please enter the name of the proxy list you want to use:")
    print("  - Leave empty for default: 'Local (No Proxy List)'")
    print("  - Or enter your proxy list name (e.g. 'really good warmed isp')")
    print("=" * 50)

    proxy_input = input("Proxy List: ").strip()

    # Use default if empty
    if not proxy_input:
        print(f"Using default proxy: 'Local (No Proxy List)'")
        return "Local (No Proxy List)"

    print(f"Using proxy: '{proxy_input}'")
    return proxy_input


def get_active_time_input():
    """
    Get user input for the active time (session duration)

    Returns:
        int: The active time in seconds
    """
    print("\n  Session Duration")
    print("=" * 50)
    print("Please enter the session duration in seconds:")
    print("  - Leave empty for default: 120 seconds (2 minutes)")
    print("  - Or enter a custom duration (e.g. 180 for 3 minutes)")
    print("=" * 50)

    time_input = input("Session Duration (seconds): ").strip()

    # Use default if empty or not a valid number
    if not time_input:
        print(f"Using default duration: 120 seconds (2 minutes)")
        return 120

    try:
        time_value = int(time_input)
        if time_value <= 0:
            print(f"Invalid duration. Using default: 120 seconds (2 minutes)")
            return 120

        minutes = time_value // 60
        seconds = time_value % 60
        time_str = f"{minutes} minutes" if seconds == 0 else f"{minutes} minutes and {seconds} seconds"
        print(f"Using custom duration: {time_value} seconds ({time_str})")
        return time_value
    except ValueError:
        print(f"Invalid duration. Using default: 120 seconds (2 minutes)")
        return 120

def kill_browsers(only_script_windows=True):
    """
    Kill browser processes

    Args:
        only_script_windows: If True, only closes browser windows launched by this script
                            If False, kills all browser processes (original behavior)
    """
    if only_script_windows:
        # Only close Selenium-launched windows by using the driver's quit() method
        # The actual closing will happen in the specific functions that use drivers
        return
    else:
        # Original behavior: Kill all browser processes
        for browser in BROWSERS.values():
            subprocess.run(browser["kill_cmd"], 
                          shell=True, capture_output=True, text=True)
        time.sleep(1)  # Give processes time to terminate


def check_driver_exists(driver_path):
    """Check if the driver exists at the specified path"""
    if not Path(driver_path).exists():
        return False
    return True


def verify_drivers():
    """Check if all required drivers are available and compatible with browser versions"""
    missing_drivers = []
    incompatible_browsers = []
    
    # First check if driver files exist
    driver_paths = set(browser["driver_path"] for browser_name, browser in BROWSERS.items() 
                      if browser["enabled"])
    for driver_path in driver_paths:
        if not check_driver_exists(driver_path):
            missing_drivers.append(driver_path)
    
    # Disable browsers with missing drivers
    if missing_drivers:
        print("  Some WebDrivers are missing:")
        for driver_path in missing_drivers:
            print(f"   - {driver_path}")
            # Disable browsers that use this driver
            for browser_name, browser in BROWSERS.items():
                if browser["driver_path"] == driver_path and browser["enabled"]:
                    print(f"     Disabling {browser['name']} due to missing driver")
                    browser["enabled"] = False
    
    # Check for known version incompatibilities (Vivaldi specifically)
    for browser_name, browser in BROWSERS.items():
        if not browser["enabled"]:
            continue
            
        if browser_name == "vivaldi":
            # Vivaldi 7.x needs ChromeDriver 133, not 134
            print(f"  Note: Vivaldi browser requires specific ChromeDriver versions")
            print(f"   Currently using driver: {browser['driver_path']}")
            print(f"   Use ChromeDriver 133 for Vivaldi 7.3.x versions")
            
    # Display enabled browsers
    enabled_browsers = [name for name, browser in BROWSERS.items() if browser["enabled"]]
    if enabled_browsers:
        print(f"   Enabled browsers: {', '.join([BROWSERS[name]['name'] for name in enabled_browsers])}")
    
    # Ensure we have at least one browser with a working driver
    if not enabled_browsers:
        raise RuntimeError("No browsers with working drivers found")


def get_extension_id():
    """Get the extension ID by looking at the manifest.json file"""
    manifest_path = Path(EXT_PATH) / "manifest.json"
    if not manifest_path.exists():
        raise RuntimeError(f"Extension manifest not found at {manifest_path}")
    
    try:
        with open(manifest_path, 'r') as f:
            manifest = json.load(f)
        # Use key if available, otherwise extract from folder name
        if 'key' in manifest:
            return manifest['key']
        else:
            # Use the folder name as extension ID if nothing else is available
            return EXT_PATH.name
    except Exception as e:
        print(f"Warning: Could not get extension ID from manifest: {e}")
        # Fallback to using the folder name
        return EXT_PATH.name


def configure_extension_shortcuts(profile_dir, browser_name):
    """Configure the extension keyboard shortcut in the preferences file"""
    global EXT_ID
    
    # Get extension ID if not already set
    if EXT_ID is None:
        EXT_ID = get_extension_id()
    
    browser = BROWSERS[browser_name]
    prefs_path = profile_dir / browser["prefs_path"]
    ext_prefs_path = profile_dir / browser["ext_prefs_path"]
    
    # Create directory structure
    default_dir = profile_dir / "Default"
    default_dir.mkdir(exist_ok=True)
    
    # Create a more comprehensive preferences structure with multiple shortcuts configurations
    preferences = {
        "extensions": {
            "commands": {
                f"{EXT_ID}:_execute_browser_action": {
                    "suggested_key": {
                        "default": "Alt+0",
                        "windows": "Alt+0",
                        "mac": "Alt+0",
                        "chromeos": "Alt+0",
                        "linux": "Alt+0"
                    },
                    "description": "Activate the browser action",
                    "global": False,
                    "enabled": True
                }
            },
            "settings": {
                f"{EXT_ID}": {
                    "active_permissions": {
                        "api": ["contextMenus", "cookies", "storage", "tabs", "webRequest", "webRequestBlocking"],
                        "explicit_host": ["*://*/*", "chrome://favicon/*"]
                    },
                    "granted_permissions": {
                        "api": ["contextMenus", "cookies", "storage", "tabs", "webRequest", "webRequestBlocking"],
                        "explicit_host": ["*://*/*", "chrome://favicon/*"]
                    },
                    "incognito": "not_allowed",
                    "location": 1,
                    "manifest": {
                        "name": "Refract",
                        "permissions": ["contextMenus", "cookies", "storage", "tabs", "webRequest", "webRequestBlocking", "*://*/*"],
                        "version": "1.0.0",
                        "manifest_version": 2,
                        "browser_action": {
                            "default_title": "Refract"
                        },
                        "commands": {
                            "_execute_browser_action": {
                                "suggested_key": {
                                    "default": "Alt+0",
                                    "windows": "Alt+0"
                                },
                                "description": "Activate the browser action"
                            }
                        }
                    },
                    "path": str(EXT_PATH),
                    "state": 1,
                    "was_installed_by_default": False,
                    "was_installed_by_oem": False
                }
            },
            "ui": {
                "developer_mode": True
            },
            "toolbar": [f"{EXT_ID}"],  # Add extension to toolbar
            "pinned_extensions": [f"{EXT_ID}"]  # Pin extension for easier access
        },
        "browser": {
            "enabled_labs_experiments": ["extension-apis", "extensions-on-chrome-urls"],
            "pinned_extensions": [f"{EXT_ID}"]  # Pin here too for redundancy
        },
        "extensions_prefs": {
            "enable_all_extensions": True,
            f"{EXT_ID}": {
                "active": True,
                "enabled": True,
                "pinned": True,
            }
        },
        "profile": {
            "exited_cleanly": True,
            "name": f"{browser_name.capitalize()} Profile"
        }
    }
    
    # Write preferences file
    with open(prefs_path, 'w') as f:
        json.dump(preferences, f)
    
    # Create Local State file to configure keyboard shortcuts at browser level
    local_state_path = profile_dir / "Local State"
    
    local_state = {
        "browser": {
            "enabled_labs_experiments": ["extension-apis", "extensions-on-chrome-urls"]
        },
        "extensions": {
            "pinned_extensions": [f"{EXT_ID}"],
            "toolbar": [f"{EXT_ID}"],
            "shortcuts": {
                f"{EXT_ID}:_execute_browser_action": {
                    "key": "Alt+0",
                    "scope": "regular"
                }
            }
        }
    }
    
    with open(local_state_path, 'w') as f:
        json.dump(local_state, f)
    
    # Copy the same preferences to secure preferences
    with open(ext_prefs_path, 'w') as f:
        json.dump(preferences, f)
    
    # Create a Commands file specifically for keyboard shortcuts
    commands_dir = default_dir / "Extensions"
    commands_dir.mkdir(exist_ok=True)
    
    ext_dir = commands_dir / EXT_ID
    ext_dir.mkdir(exist_ok=True)
    
    commands_path = ext_dir / "Commands"
    
    commands = {
        "commands": {
            "_execute_browser_action": {
                "command_key": "Alt+0",
                "command_mac": "Alt+0",
                "enabled": True,
                "global": False
            }
        }
    }
    
    with open(commands_path, 'w') as f:
        json.dump(commands, f)
    
    print(f"     Set up keyboard shortcut Alt+0 for Refract extension in multiple configuration files")
    
    return profile_dir


def launch_driver(browser_name, profile_dir, rect):
    """Launch a browser with specified profile using the appropriate driver"""
    browser = BROWSERS[browser_name]
    
    try:
        print(f"     Creating browser driver for {browser['name']}")
        
        # Create options setup
        options = ChromeOptions()
        options.binary_location = browser["exe"]
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument(f"--disable-extensions-except={EXT_PATH}")
        options.add_argument(f"--load-extension={EXT_PATH}")
        options.add_argument("--no-sandbox")
        
        # Log info for debugging
        print(f"     Using profile: {profile_dir}")
        print(f"     Using extension: {EXT_PATH}")
        print(f"     Using driver: {browser['driver_path']}")
        
        # Create Chrome driver
        service = ChromeService(executable_path=browser["driver_path"])
        drv = webdriver.Chrome(service=service, options=options)
            
        # Set window size and position
        drv.set_window_rect(*rect)
        
        # Navigate directly to the target URL
        print(f"     Navigating to target URL")
        drv.get(PRODUCT_URL)
        time.sleep(BROWSER_LOAD_WAIT)
        
        # Get extension ID if not already set
        global EXT_ID
        if EXT_ID is None:
            print(f"     Getting extension ID")
            EXT_ID = get_extension_id()
            print(f"      Extension ID: {EXT_ID}")
        
        return drv
        
    except Exception as e:
        print(f"     Error creating driver: {e}")
        raise  # Re-raise the exception to allow proper error handling


def open_popup(driver, retries=3):
    """
    Open the extension popup using keyboard shortcut (Alt+0)
    """
    desktop = Desktop(backend="uia")
    print(f"     Attempting to open extension popup with Alt+0 shortcut...")
    
    for attempt in range(retries):
        print(f"     Attempt {attempt+1} of {retries}...")
        before = set(desktop.windows())
        
        # Ensure browser window is in focus
        driver.switch_to.window(driver.current_window_handle)
        
        # Try to activate the extension via keyboard shortcut
        try:
            # First make sure the window is focused
            active_window = desktop.window(active_only=True)
            active_window.set_focus()
            print(f"        Window focused, sending Alt+0...")
            
            # Try opening with keyboard shortcut
            send_keys(HOTKEY)
            time.sleep(0.7)  # Slightly longer wait
            
            # Check for new windows
            new = set(desktop.windows()) - before
            if new:
                popup = desktop.window(handle=new.pop().handle)
                popup.wait("exists enabled visible ready", timeout=5)
                print(f"         Successfully opened extension popup!")
                return popup
        except Exception as e:
            print(f"        Error during focus/keypress: {e}")
        
        print(f"        No popup window detected")
        time.sleep(0.5)
    
    print(f"     Failed to open extension popup with Alt+0 shortcut after {retries} attempts")
    print(f"     Try running reset_profiles.py to fix keyboard shortcuts")
    return None




def set_proxy_local(popup):
    """Switch combo-box to target proxy list"""
    try:
        print(f"    Looking for ComboBox in popup...")
        print(f"    Setting proxy to: '{TARGET_PROXY}'")

        # Wait a bit longer for UI to be fully ready
        time.sleep(1.5)

        # Find and expand the combo box - more flexible search
        combo = None
        try:
            # Try by control type first
            combo = popup.child_window(control_type="ComboBox")
            print(f"      Found ComboBox by control type")
        except Exception:
            # Try other methods
            try:
                combo = popup.child_window(class_name="ComboBox")
                print(f"     Found ComboBox by class name")
            except Exception:
                # Try finding any combo-like control
                all_children = popup.children()
                for child in all_children:
                    if 'combo' in child.friendly_class_name().lower():
                        combo = child
                        print(f"      Found ComboBox by friendly class name")
                        break

        if not combo:
            print(f"     Could not find ComboBox, using keyboard method directly")
            # Direct keyboard method without combo
            popup.set_focus()
            # First press tab to get focus on the combo
            for _ in range(3):  # Try tabbing to get to the combo
                send_keys("{TAB}")
                time.sleep(0.3)
            # Then open the dropdown
            send_keys("%{DOWN}")  # Alt+Down
            time.sleep(0.3)
            # Select by first letter
            first_letter = TARGET_PROXY[0].lower()
            send_keys(first_letter)
            time.sleep(0.3)
            send_keys("{ENTER}")
            print(f"      Set proxy to '{TARGET_PROXY}' using keyboard method")
            return

        # Found combo box, proceed with attempts
        combo.wait("enabled visible ready", timeout=3)
        print(f"     Expanding ComboBox to select '{TARGET_PROXY}'...")
        try:
            combo.expand()
        except Exception as e:
            print(f"     Standard expand failed: {e}, trying alternate methods")
            try:
                combo.click_input()
                time.sleep(0.5)
            except Exception:
                # Try keyboard shortcut to expand
                combo.set_focus()
                send_keys("{F4}")  # Often works to expand combos
                time.sleep(0.5)

        # Try direct selection (fastest method)
        try:
            print(f"     Trying direct selection of '{TARGET_PROXY}'")
            combo.select(TARGET_PROXY)
            print(f"      Successfully set proxy to '{TARGET_PROXY}'")
            return
        except Exception as e:
            print(f"     Direct selection failed: {e}")

        # Try clicking the item directly
        try:
            print(f"     Looking for list item '{TARGET_PROXY}'")
            item = popup.child_window(title=TARGET_PROXY, control_type="ListItem")
            item.wait("exists visible ready", timeout=2)
            print(f"     Clicking list item '{TARGET_PROXY}'")
            item.click_input()
            print(f"      Successfully set proxy to '{TARGET_PROXY}' via item click")
            return
        except Exception as e:
            print(f"     Item click failed: {e}")

        # Last resort: keyboard method
        print(f"     Using keyboard method to select '{TARGET_PROXY}'")
        combo.set_focus()
        send_keys("%{DOWN}")  # Alt+Down
        time.sleep(0.5)
        # Try selecting first item if it's what we want
        first_letter = TARGET_PROXY[0].lower()
        send_keys(first_letter)
        time.sleep(0.5)
        send_keys("{ENTER}")
        print(f"      Successfully set proxy to '{TARGET_PROXY}' using keyboard method")

    except Exception as e:
        print(f"     Failed to set proxy to '{TARGET_PROXY}': {e}")


def clear_and_toggle(popup):
    """Clear and toggle the harvesting"""
    try:
        print(f"     Looking for clear and toggle buttons...")
        # Wait for UI to be fully ready
        time.sleep(1)

        # Try finding the clear button by title
        clear_button = None
        try:
            print(f"     Looking for '{CLEAR_LABEL}' button")
            clear_button = popup.child_window(title=CLEAR_LABEL, control_type="Button")
            clear_button.wait("exists enabled visible ready", timeout=2)
            print(f"     Clicking clear button")
            clear_button.click_input()
            time.sleep(0.5)
        except Exception as e:
            print(f"     Could not find/click clear button by standard method: {e}")

            # Alternative: try finding by partial title match
            try:
                for child in popup.children():
                    if hasattr(child, 'window_text') and 'clear' in child.window_text().lower():
                        clear_button = child
                        print(f"      Found clear button by text: {child.window_text()}")
                        break

                if clear_button:
                    print(f"     Clicking clear button (alt method)")
                    clear_button.click_input()
                    time.sleep(0.5)
                else:
                    # Try keyboard navigation - tab to the clear button
                    print(f"     Using keyboard navigation for clear button")
                    popup.set_focus()
                    # Generally the first button after tabbing
                    send_keys("{TAB}")
                    time.sleep(0.3)
                    send_keys("{ENTER}")
                    time.sleep(0.5)
            except Exception as e2:
                print(f"     Alternative clear button methods failed: {e2}")

        # Try finding the toggle button by title regex
        toggle_button = None
        try:
            print(f"     Looking for toggle button (Start/Stop Harvesting)")
            toggle_button = popup.child_window(title_re=TOGGLE_REGEX, control_type="Button")
            toggle_button.wait("exists enabled visible ready", timeout=2)
            print(f"     Clicking toggle button")
            toggle_button.click_input()
            time.sleep(0.5)
        except Exception as e:
            print(f"     Could not find/click toggle button by standard method: {e}")

            # Alternative: try finding by partial title match
            try:
                for child in popup.children():
                    if (hasattr(child, 'window_text') and
                        ('harvest' in child.window_text().lower() or
                         'start' in child.window_text().lower() or
                         'stop' in child.window_text().lower())):
                        toggle_button = child
                        print(f"      Found toggle button by text: {child.window_text()}")
                        break

                if toggle_button:
                    print(f"     Clicking toggle button (alt method)")
                    toggle_button.click_input()
                    time.sleep(0.5)
                else:
                    # Try keyboard navigation - tab to the toggle button (usually second button)
                    print(f"     Using keyboard navigation for toggle button")
                    popup.set_focus()
                    # Tab twice to get to the second button
                    send_keys("{TAB}")
                    time.sleep(0.3)
                    send_keys("{TAB}")
                    time.sleep(0.3)
                    send_keys("{ENTER}")
                    time.sleep(0.5)
            except Exception as e2:
                print(f"     Alternative toggle button methods failed: {e2}")

        print(f"      Clear and toggle operations completed")
        return True
    except Exception as e:
        print(f"     Error in clear_and_toggle: {e}")
        return False


def start_harvest(driver):
    """Start harvesting process using the keyboard shortcut method"""
    try:
        popup = open_popup(driver)
        
        if popup:
            set_proxy_local(popup)
            clear_and_toggle(popup)
            return popup
        else:
            raise RuntimeError("Failed to open Refract extension with keyboard shortcut")
    except Exception as e:
        print(f"     Failed to start harvesting: {e}")
        return None


def stop_harvest(popup):
    """Stop harvesting"""
    try:
        print(f"     Attempting to stop harvesting...")
        try:
            # Try standard method first
            toggle_button = popup.child_window(title_re=TOGGLE_REGEX, control_type="Button")
            toggle_button.wait("exists enabled visible ready", timeout=2)
            print(f"     Clicking toggle button to stop harvesting")
            toggle_button.click_input()
            time.sleep(0.5)
        except Exception as e:
            print(f"     Could not find/click toggle button by standard method: {e}")

            # Try alternative methods
            try:
                # Look for any button that might be the toggle
                toggle_found = False
                for child in popup.children():
                    if (hasattr(child, 'window_text') and
                        ('harvest' in child.window_text().lower() or
                         'start' in child.window_text().lower() or
                         'stop' in child.window_text().lower())):
                        print(f"      Found toggle button by text: {child.window_text()}")
                        child.click_input()
                        toggle_found = True
                        time.sleep(0.5)
                        break

                if not toggle_found:
                    # Try keyboard navigation - tab to the toggle button (usually second button)
                    print(f"     Using keyboard navigation for toggle button")
                    popup.set_focus()
                    # Tab twice to get to the second button
                    send_keys("{TAB}")
                    time.sleep(0.3)
                    send_keys("{TAB}")
                    time.sleep(0.3)
                    send_keys("{ENTER}")
                    time.sleep(0.5)
            except Exception as e2:
                print(f"     Alternative toggle button methods failed: {e2}")

        # Try closing the popup
        try:
            print(f"     Closing popup window")
            popup.close()
            time.sleep(0.5)
        except Exception as e:
            print(f"     Error closing popup: {e}")
            # Try alt+F4 to close window
            try:
                popup.set_focus()
                send_keys("%{F4}")  # Alt+F4
                time.sleep(0.5)
            except:
                print(f"     Could not close popup with keyboard shortcut")

        print(f"      Stop harvest completed")
        return True
    except Exception as e:
        print(f"     Error stopping harvest: {e}")
        return False


# ─── BROWSER PAIR GENERATOR ───────────────────────────────────────
def generate_browser_pairs():
    """Generate unique browser pairs, mixing different browsers and using all profiles"""
    # Get only enabled browsers
    browser_names = [name for name, browser in BROWSERS.items() if browser["enabled"]]
    
    # We should have Chrome and Brave enabled
    if len(browser_names) < 2:
        if len(browser_names) == 1:
            # Use the same browser for both if only one is available
            print(f"Warning: Only one browser is enabled: {browser_names[0]}")
            print(f"Will use the same browser with different profiles")
            return cycle([(browser_names[0], browser_names[0])])
        else:
            raise RuntimeError("No compatible browsers found.")
    
    # Get available profiles for each browser
    config = load_profiles_config()
    browser_profiles = config.get("browser_profiles", {})
    
    profile_counts = {}
    for browser in browser_names:
        if browser in browser_profiles:
            profile_counts[browser] = len(browser_profiles[browser])
        else:
            profile_counts[browser] = 0
    
    print(f"Available profiles: {profile_counts}")
    
    # Create profile pairs
    pairs = []
    
    # Generate pairs of different browsers with different profiles
    for browser1 in browser_names:
        for browser2 in browser_names:
            if browser1 == browser2:
                continue  # Skip same browser pairs
                
            # Get number of profiles for each browser
            profiles1 = profile_counts.get(browser1, 0)
            profiles2 = profile_counts.get(browser2, 0)
            
            # Skip if either browser has no profiles
            if profiles1 == 0 or profiles2 == 0:
                continue
                
            # Create pairs with all profile combinations
            for profile1 in range(1, profiles1 + 1):
                for profile2 in range(1, profiles2 + 1):
                    pairs.append((browser1, browser2, profile1, profile2))
    
    # If no valid pairs were found, use same browser with different profiles
    if not pairs:
        for browser in browser_names:
            profiles = profile_counts.get(browser, 0)
            if profiles >= 2:
                for i in range(1, profiles):
                    for j in range(i+1, profiles+1):
                        pairs.append((browser, browser, i, j))
    
    # Shuffle the pairs for randomness
    if pairs:
        random.shuffle(pairs)
        print(f"Using browser pairs: {pairs}")
        return cycle(pairs)
    else:
        raise RuntimeError("No valid browser profile pairs found.")


# ─── INITIAL PROFILE SETUP ──────────────────────────────────────────
def setup_initial_profiles():
    """Create and configure initial browser profiles with manual configuration"""
    print("  Initial profile setup for persistent browser usage")
    
    # Load existing configuration
    config = load_profiles_config()
    
    # Check if we've already configured profiles
    if config.get("configured", False):
        print("   Profiles already configured, skipping initial setup")
        return
    
    # Check if all required drivers exist
    verify_drivers()
    
    # Filter out incompatible browsers
    enabled_browsers = [name for name, browser in BROWSERS.items() if browser["enabled"]]
    if not enabled_browsers:
        raise RuntimeError("No compatible browsers found. Please provide working WebDrivers.")
    
    print(f"   Available browsers: {', '.join([BROWSERS[name]['name'] for name in enabled_browsers])}")
    
    # Create base profile directories
    for browser_name in enabled_browsers:
        browser = BROWSERS[browser_name]
        Path(browser["profile_base"]).mkdir(exist_ok=True, parents=True)
    
    # Only close script-launched browsers
    kill_browsers(only_script_windows=True)
    
    try:
        # Create and configure 2 profiles for each browser
        for browser_name in enabled_browsers:
            print(f"\n➡ Setting up {BROWSERS[browser_name]['name']} profiles...")
            
            for profile_number in range(1, 3):  # Create 2 profiles (1 and 2)
                print(f"\n  Creating {BROWSERS[browser_name]['name']} profile {profile_number}...")
                
                # Create and get profile directory
                profile_dir = create_persistent_profile(browser_name, profile_number, force_new=True)
                
                # Launch browser with this profile
                driver = launch_driver(browser_name, profile_dir, RECTS[0])
                
                try:
                    # Wait for manual configuration - browser will be closed by this function
                    wait_for_manual_configuration(driver, browser_name, profile_number, profile_dir)
                    print(f"   Successfully configured {BROWSERS[browser_name]['name']} profile {profile_number}")
                    
                except Exception as e:
                    print(f"  Error during profile configuration: {e}")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    
                # Only close script-launched browser windows
                kill_browsers(only_script_windows=True)
                
                # Allow some time between profile setups
                time.sleep(3)
        
        # Mark configuration as complete
        config["configured"] = True
        save_profiles_config(config)
        
        print("\n   All browser profiles have been set up!")
        
        # Display all configured profiles
        print("\n  Configured Profiles Summary:")
        for browser_name, profiles in config.get("browser_profiles", {}).items():
            print(f"\n{BROWSERS[browser_name]['name']}:")
            for profile_num, profile_data in profiles.items():
                print(f"   Profile {profile_num}: {profile_data['path']}")
        
    except KeyboardInterrupt:
        print("\n  Profile setup interrupted by user — exiting cleanly.")
    except Exception as e:
        print(f"\n  Error during profile setup: {e}")
    finally:
        # Only close script-launched browser windows
        kill_browsers(only_script_windows=True)
        
        print("  Initial profile setup complete!")


# ─── MAIN COOKIE HARVESTING LOOP ──────────────────────────────────
# Global stop flag that can be set by external code
STOP_REQUESTED = False

def run_cookie_harvesting():
    """Run the main cookie harvesting loop using configured profiles"""
    global STOP_REQUESTED
    print("\n  Starting cookie harvesting (Ctrl+C to stop)")
    print(f"  Using proxy: '{TARGET_PROXY}'")

    # Display active time information
    minutes = ACTIVE_TIME // 60
    seconds = ACTIVE_TIME % 60
    time_str = f"{minutes} minutes" if seconds == 0 else f"{minutes} minutes and {seconds} seconds"
    print(f" ️ Session duration: {ACTIVE_TIME} seconds ({time_str})")
    
    # Check if profiles are configured
    config = load_profiles_config()
    
    # If not configured, run setup
    if not config.get("configured", False) or not config.get("browser_profiles", {}):
        print("  Profiles not configured yet. Running initial setup...")
        setup_initial_profiles()
        config = load_profiles_config()
    
    # Check for configured profiles
    browser_profiles = config.get("browser_profiles", {})
    if not browser_profiles:
        print("  No browser profiles found. Please run setup first.")
        return
    
    # Get all enabled browsers with configured profiles
    enabled_browsers = []
    for browser_name, browser in BROWSERS.items():
        if browser["enabled"] and browser_name in browser_profiles:
            enabled_browsers.append(browser_name)
    
    if not enabled_browsers:
        print("  No configured profiles for enabled browsers.")
        return
    
    # Get pairs of different browsers
    browser_pairs = generate_browser_pairs()
    
    # Keep track of which profile to use
    profile_tracker = {}
    for browser in enabled_browsers:
        profile_tracker[browser] = 1  # Start with profile 1
    
    try:
        while not STOP_REQUESTED:
            # Close previous browser windows
            kill_browsers(only_script_windows=True)
            
            # Get next browser pair with specific profiles
            browser1, browser2, profile1_num, profile2_num = next(browser_pairs)
            
            # Get paths
            profile1_dir = Path(browser_profiles.get(browser1, {}).get(str(profile1_num), {}).get("path", ""))
            profile2_dir = Path(browser_profiles.get(browser2, {}).get(str(profile2_num), {}).get("path", ""))
            
            # Skip if any profile is missing
            if not profile1_dir.exists() or not profile2_dir.exists():
                print(f"  Missing profile directory, skipping this pair")
                continue
                
            print(f"Using {browser1} profile {profile1_num} and {browser2} profile {profile2_num}")
            
            # Launch first browser and start harvesting
            driver_a = launch_driver(browser1, profile1_dir, RECTS[0])
            if STOP_REQUESTED: 
                driver_a.quit()
                break
                
            popup_a = start_harvest(driver_a)
            if not popup_a or STOP_REQUESTED:
                driver_a.quit()
                continue
            
            # Launch second browser and start harvesting
            driver_b = launch_driver(browser2, profile2_dir, RECTS[1])
            if STOP_REQUESTED:
                stop_harvest(popup_a)
                driver_a.quit()
                driver_b.quit()
                break
                
            popup_b = start_harvest(driver_b)
            
            # If both browsers are harvesting, wait for the specified time
            if popup_b:
                # Format the time display nicely
                minutes = ACTIVE_TIME // 60
                seconds = ACTIVE_TIME % 60
                if seconds == 0:
                    print(f"     Harvesting for {minutes} minutes...")
                else:
                    print(f"     Harvesting for {minutes} minutes and {seconds} seconds...")
                
                # Wait but check for stop request
                wait_start_time = time.time()
                while time.time() - wait_start_time < ACTIVE_TIME and not STOP_REQUESTED:
                    time.sleep(1)
                
                # Stop harvesting and close browsers
                stop_harvest(popup_a)
                stop_harvest(popup_b)
                
                driver_a.quit()
                driver_b.quit()
            else:
                # Second browser failed, stop first browser
                stop_harvest(popup_a)
                driver_a.quit()
                driver_b.quit()
            
            # Check if stop was requested
            if STOP_REQUESTED:
                break
                
            # Small delay before next cycle
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n  Stopped by user — exiting cleanly.")
    except Exception as e:
        print(f"\n  Error: {e}")
    finally:
        # Only close script-launched browser windows
        kill_browsers(only_script_windows=True)
        print("  Goodbye!")


# ─── MAIN SCRIPT ──────────────────────────────────────────────────
if __name__ == "__main__":
    print("  Persistent Browser Profile Cookie Harvesting Tool")

    # Set extension path to local folder using fix_paths
    EXT_PATH = Path(fix_paths.get_absolute_path("latest/build"))
    if not EXT_PATH.exists():
        print(f"  Extension path {EXT_PATH} not found. Please check that the 'latest/build' folder exists.")
        sys.exit(1)
    else:
        print(f"   Using extension at: {EXT_PATH}")

    # Verify ChromeDriver exists in the local directory using fix_paths
    chromedriver_path = Path(fix_paths.get_absolute_path("chromedriver.exe"))
    if not chromedriver_path.exists():
        print(f"  ChromeDriver not found at {chromedriver_path}. Please make sure it's in the correct location.")
        sys.exit(1)
    else:
        print(f"   Using ChromeDriver at: {chromedriver_path}")
        # Update ChromeDriver path in all browser configurations
        for browser in BROWSERS.values():
            if browser["driver_type"] == "chrome":
                browser["driver_path"] = str(chromedriver_path)

    print("   Enabled browsers: Chrome, Brave")
    print(f"   Chrome: {CHROME_EXE}")
    print(f"   Brave: {BRAVE_EXE}")

    # Create base directories
    PERSISTENT_PROFILES_BASE.mkdir(exist_ok=True, parents=True)

    # First check if we need to set up profiles
    config = load_profiles_config()
    if not config.get("configured", False):
        setup_initial_profiles()

    # Get proxy input from user
    TARGET_PROXY = get_proxy_input()

    # Get active time input from user
    ACTIVE_TIME = get_active_time_input()

    # Run the main cookie harvesting loop
    run_cookie_harvesting()