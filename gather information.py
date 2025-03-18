import os
import subprocess
import sys
import importlib.util
import ctypes
import platform
import socket
import shutil
import tempfile
import psutil
import pandas as pd
import win32com.client
from datetime import datetime
from openpyxl import load_workbook
from datetime import datetime
# Required packages and their corresponding import names
REQUIRED_PACKAGES = {
    "psutil": "psutil",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "pywin32": "win32com.client"
}

installed_packages = []

def install_missing_packages(packages):
    """Install missing packages and track what was installed for later removal."""
    global installed_packages
    missing_packages = [pkg for pkg, import_name in packages.items() if importlib.util.find_spec(import_name) is None]

    if missing_packages:
        subprocess.run([sys.executable, "-m", "pip", "install", *missing_packages], check=True)
        installed_packages.extend(missing_packages)  # Keep track of what we installed

# Ensure required packages are installed
install_missing_packages(REQUIRED_PACKAGES)

def get_windows_version():
    """Determine if the system is running Windows 10 or Windows 11"""
    release = platform.release()
    return "Windows 10" if release == "10" else "Windows 11" if release == "11" else "Unknown Windows Version"

def get_system_info():
    """Gather system information"""
    return {
        "PC Name": platform.node(),
        "IP Address": socket.gethostbyname(socket.gethostname()),
        "Windows": platform.release(),
        "Total Memory (GB)": round(psutil.virtual_memory().total / (1024 ** 3), 2),
        "Total Storage (GB)": round(psutil.disk_usage("/").total / (1024 ** 3), 2),
        "Free Disk Space (GB)": round(psutil.disk_usage("/").free / (1024 ** 3), 2),
        "Disk Type": "SSD" if "SSD" in str(subprocess.getoutput("wmic diskdrive get MediaType")) else "HDD",
        "CPU": platform.processor(),
        "GPU": get_gpu(),
        "Mac Address": get_mac_address(),
        "Serial Number": get_serial_number(),
        "Date Collected": datetime.today().strftime('%Y-%m-%d')
    }
def get_gpu():
    """Retrieve the first detected GPU name (handles empty output)."""
    try:
        output = subprocess.check_output("wmic path win32_videocontroller get name", shell=True).decode()
        gpus = [line.strip() for line in output.split("\n") if line.strip() and "Name" not in line]
        return gpus[0] if gpus else "N/A"
    except Exception:
        return "N/A"

def get_serial_number():
    """Retrieve the PC's serial number."""
    try:
        output = subprocess.check_output("wmic bios get serialnumber", shell=True).decode()
        serials = [line.strip() for line in output.split("\n") if line.strip() and "SerialNumber" not in line]
        return serials[0] if serials else "N/A"
    except Exception:
        return "N/A"
def get_mac_address():
    """Retrieve the first available MAC address from active network interfaces."""
    network_interfaces = psutil.net_if_addrs()
    
    for interface_name, addresses in network_interfaces.items():
        for addr in addresses:
            if addr.family == psutil.AF_LINK:  # AF_LINK means it's a MAC address
                return addr.address  # Return the first found MAC address
    
    return "N/A"  # If no MAC address is found
def save_to_excel(info, filename="system_info.xlsx"):
    """Save system information to an Excel file in wide format and auto-adjust column width"""
    file_path = os.path.join(tempfile.gettempdir(), filename)  # Store in temp folder
    df = pd.DataFrame([list(info.values())], columns=list(info.keys()))  # Create a wide table
    df.to_excel(file_path, index=False)  # Save to file

    # Auto-adjust column width
    wb = load_workbook(file_path)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Get Excel column letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2  # Add padding
    wb.save(file_path)

    return file_path

def send_email(recipient_email, attachment_path, pc_name):
    """Send an email using the currently logged-in Outlook account"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = f"System Report - {pc_name} - {datetime.now().strftime('%Y-%m-%d')}"
        mail.Body = f"Attached is the system information report from {pc_name}."
        mail.Attachments.Add(attachment_path)
        mail.Send()
        return True
    except Exception as e:
        return False

def cleanup():
    """Remove installed dependencies and temporary files after script execution"""
    if installed_packages:
        subprocess.run([sys.executable, "-m", "pip", "uninstall", "-y", *installed_packages], check=True)
    
    temp_dir = tempfile.gettempdir()
    for file in os.listdir(temp_dir):
        if file.startswith("system_info") and file.endswith(".xlsx"):
            try:
                os.remove(os.path.join(temp_dir, file))
            except Exception:
                pass
def show_popup():
    """Show a Windows message box when the script finishes, including Windows Update status."""
    message = f"בוצע בהצלחה!"
    ctypes.windll.user32.MessageBoxW(0, message, "System Info Sent", 0x40)
    self_delete()  # Trigger self-deletion


def self_delete():
    """Delete the executable file after user clicks OK."""
    exe_path = sys.argv[0]

    # Only delete if running as an .exe file
    if exe_path.endswith(".exe"):
        delete_script = f"""
        @echo off
        timeout /t 3 >nul
        del "{exe_path}"
        """
        bat_file = os.path.join(tempfile.gettempdir(), "delete_me.bat")

        with open(bat_file, "w") as f:
            f.write(delete_script)

        subprocess.Popen(bat_file, shell=True)  # Run batch file to delete the .exe

def main():
    """Main function to execute the script"""
    info = get_system_info()
    pc_name = info["PC Name"]  # Get PC name for the email subject
    report_file = save_to_excel(info)

    # Hardcoded recipient email (yours)
    recipient_email = "dany@lahav.ac.il"  # Replace with your email

    if send_email(recipient_email, report_file, pc_name):
        cleanup()
        show_popup()  # Show success message and trigger self-deletion

if __name__ == "__main__":
    main()