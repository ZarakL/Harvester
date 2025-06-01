import time
import subprocess
import random
import os
import json
import sys
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
import threading

# ─── CONSTANT PATHS ─────────────────────────────────────────────
# Get script directory
script_dir = Path(os.path.dirname(os.path.abspath(__file__)))

# Browser executables (standard installation paths)
BRAVE_EXE = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# ChromeDriver path
CHROMEDRIVER = script_dir / "chromedriver.exe"

# Extension path
EXT_PATH = script_dir / "latest" / "build"

# Profile configuration 
PERSISTENT_PROFILES_BASE = script_dir / "profiles" 
PROFILES_CONFIG_FILE = script_dir / "profiles_config.json"

# Window positioning and timing
WINDOW_WIDTH = 1024
WINDOW_HEIGHT = 768
PROFILE_SETUP_WAIT = 300  # 5 minutes maximum wait time per profile

# ─── UTILITY FUNCTIONS ─────────────────────────────────────────────
def load_profiles_config():
    """Load profiles configuration from JSON file"""
    if PROFILES_CONFIG_FILE.exists():
        try:
            with open(PROFILES_CONFIG_FILE, 'r') as f:
                config = json.load(f)
            print(f"Loaded profiles configuration from {PROFILES_CONFIG_FILE}")
            return config
        except Exception as e:
            print(f"Error loading profiles configuration: {e}")
    
    # Default config
    print(f"Creating new profiles configuration")
    return {
        "configured": False,
        "browser_profiles": {}
    }

def save_profiles_config(config):
    """Save profiles configuration to JSON file"""
    try:
        PROFILES_CONFIG_FILE.parent.mkdir(exist_ok=True, parents=True)
        with open(PROFILES_CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
        print(f"Saved profiles configuration to {PROFILES_CONFIG_FILE}")
        return True
    except Exception as e:
        print(f"Error saving profiles configuration: {e}")
        return False

def generate_profile_path(browser_name, profile_number=1):
    """Generate a persistent profile path for a browser"""
    profile_base = PERSISTENT_PROFILES_BASE / f"{browser_name}_profiles"
    
    # Create base directory if it doesn't exist
    profile_base.mkdir(exist_ok=True, parents=True)
    
    # Create profile directory
    profile_dir = profile_base / f"profile_{profile_number}"
    
    return profile_dir

def wait_for_user_input(timeout, driver):
    """Wait for user input or timeout"""
    user_input_received = False
    
    def check_input():
        nonlocal user_input_received
        input()  # This will block until the user presses Enter
        user_input_received = True
    
    # Start a thread to wait for user input
    input_thread = threading.Thread(target=check_input)
    input_thread.daemon = True
    input_thread.start()
    
    # Wait for either user input or timeout
    timeout_time = time.time() + timeout
    while not user_input_received and time.time() < timeout_time:
        # Check if browser is still active
        try:
            if not driver.window_handles:
                print(f"Browser window closed, assuming configuration is complete")
                return True
        except Exception:
            print(f"Browser connection lost, assuming configuration is complete")
            return True
            
        time.sleep(1)  # Check every second
        
    if user_input_received:
        print(f"User confirmed configuration is complete!")
        return True
    
    print(f"Setup timed out")
    return False

def configure_profile(browser_name, profile_number):
    """Configure a profile with user interaction"""
    print(f"\nSetting up {browser_name.capitalize()} profile {profile_number}...")
    
    # Create profile directory
    profile_dir = generate_profile_path(browser_name, profile_number)
    profile_dir.mkdir(exist_ok=True, parents=True)
    
    # Create default directory
    default_dir = profile_dir / "Default"
    default_dir.mkdir(exist_ok=True)
    
    # Setup Chrome options
    options = ChromeOptions()
    if browser_name == "brave":
        options.binary_location = BRAVE_EXE
    else:
        options.binary_location = CHROME_EXE
        
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument(f"--load-extension={EXT_PATH}")
    options.add_argument("--no-sandbox")
    
    # Create Chrome service
    service = ChromeService(executable_path=CHROMEDRIVER)
    
    # Launch browser
    print(f"   Launching {browser_name.capitalize()} with profile at {profile_dir}")
    driver = webdriver.Chrome(service=service, options=options)
    
    # Set window size
    driver.set_window_size(WINDOW_WIDTH, WINDOW_HEIGHT)
    
    # Navigate to shortcuts page
    driver.get("chrome://extensions/shortcuts")
    
    # Instructions for the user
    print("\nSETUP INSTRUCTIONS:")
    print("   1. Find the Refract extension in the list")
    print("   2. Click the keyboard shortcut input field")
    print("   3. Press Alt+0 to set the shortcut")
    print("   4. Optional: Pin the extension to the toolbar")
    print(f"   5. After setup is complete, press ENTER in this terminal window")
    print(f"   You have {PROFILE_SETUP_WAIT // 60} minutes to complete this setup")
    
    # Wait for user to complete setup
    setup_complete = wait_for_user_input(PROFILE_SETUP_WAIT, driver)
    
    # Close the browser
    try:
        driver.quit()
    except Exception:
        pass
    
    # Update profiles config
    if setup_complete:
        config = load_profiles_config()
        
        if "browser_profiles" not in config:
            config["browser_profiles"] = {}
            
        if browser_name not in config["browser_profiles"]:
            config["browser_profiles"][browser_name] = {}
            
        config["browser_profiles"][browser_name][str(profile_number)] = {
            "path": str(profile_dir),
            "configured": True,
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        
        save_profiles_config(config)
        print(f"Successfully configured {browser_name.capitalize()} profile {profile_number}")
        return True
    
    print(f"Failed to configure {browser_name.capitalize()} profile {profile_number}")
    return False

def main():
    """Main function to configure browser profiles"""
    print("Browser Profile Setup Tool")
    
    # Check for Chromedriver
    if not CHROMEDRIVER.exists():
        print(f"ChromeDriver not found at {CHROMEDRIVER}")
        print(f"   Please download ChromeDriver and place it in the same directory as this script")
        return
    
    # Check for extension
    if not EXT_PATH.exists():
        print(f"Extension not found at {EXT_PATH}")
        print(f"   Please place the extension in the 'latest/build' directory")
        return
    
    # Create base directories
    PERSISTENT_PROFILES_BASE.mkdir(exist_ok=True, parents=True)
    
    # Configure browser profiles
    browsers = ["chrome", "brave"]
    profiles_per_browser = 4  # Configure 4 profiles per browser
    
    # Count total profiles to configure
    total_profiles = len(browsers) * profiles_per_browser
    profiles_configured = 0
    
    for browser_name in browsers:
        print(f"\nSetting up {browser_name.capitalize()} profiles...")
        
        for profile_number in range(1, profiles_per_browser + 1):
            profiles_configured += 1
            print(f"\nProgress: {profiles_configured}/{total_profiles} profiles")
            
            if configure_profile(browser_name, profile_number):
                print(f"✅ {browser_name.capitalize()} profile {profile_number} configured successfully")
            else:
                print(f"Failed to configure {browser_name.capitalize()} profile {profile_number}")
            
            # Add a small delay between profile setups
            time.sleep(3)
    
    # Mark configuration as complete
    config = load_profiles_config()
    config["configured"] = True
    save_profiles_config(config)
    
    print("\nAll browser profiles have been set up!")
    print("   You can now run cookiesscript.py to start harvesting")

if __name__ == "__main__":
    main()