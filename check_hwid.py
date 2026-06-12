import subprocess
import os

def get_hwid():
    try:
        # Get BIOS UUID using wmic
        cmd = "wmic csproduct get uuid"
        uuid = subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip()
        return uuid
    except:
        # Fallback to volume serial number
        try:
            import os
            return os.popen("vol C:").read().split()[-1]
        except:
            return "UNKNOWN_DEVICE"

if __name__ == "__main__":
    print(f"HWID: {get_hwid()}")
