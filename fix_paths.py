import os
import sys
import json
from pathlib import Path

# Get the current executable directory
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    exe_dir = Path(os.path.dirname(sys.executable))
else:
    # Running as script
    exe_dir = Path(os.path.dirname(os.path.abspath(__file__)))

# Override resource_path function to always use executable directory
def get_absolute_path(relative_path):
    return str(exe_dir / relative_path)

# Update the profiles_config.json file if it exists
config_path = exe_dir / "profiles_config.json"
if config_path.exists():
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)

        # Update paths for all browser profiles
        modified = False
        for browser, profiles in config.get("browser_profiles", {}).items():
            for profile_num, profile_data in profiles.items():
                # Replace absolute path with path relative to exe directory
                new_path = str(exe_dir / "profiles" / f"{browser}_profiles" / f"profile_{profile_num}")
                if profile_data.get("path") != new_path:
                    profile_data["path"] = new_path
                    modified = True

        # Save if modified
        if modified:
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=2)
            print(f"✅ Updated profile paths in {config_path}")
    except Exception as e:
        print(f"⚠️ Error updating profiles_config.json: {e}")