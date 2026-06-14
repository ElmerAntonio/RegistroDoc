import os
import sys
import ctypes
import time

# List of analyst tools and reverse-engineering software processes
SUSPICIOUS_PROCESSES = [
    "ida.exe", "ida64.exe", "ghidra", "x64dbg.exe", "x32dbg.exe",
    "wireshark.exe", "ollydbg.exe", "processhacker.exe", "procexp.exe",
    "cheatengine.exe", "cheatengine-x86_64.exe", "cheatengine-i386.exe"
]

# Structure for process enumeration using Win32 API
class PROCESSENTRY32(ctypes.Structure):
    _fields_ = [
        ("dwSize", ctypes.c_ulong),
        ("cntUsage", ctypes.c_ulong),
        ("th32ProcessID", ctypes.c_ulong),
        ("th32DefaultHeapID", ctypes.c_void_p),
        ("th32ModuleID", ctypes.c_ulong),
        ("cntThreads", ctypes.c_ulong),
        ("th32ParentProcessID", ctypes.c_ulong),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", ctypes.c_ulong),
        ("szExeFile", ctypes.c_char * 260)
    ]

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("dwTime", ctypes.c_uint)
    ]

def detect_debugger():
    """Detect if a debugger is attached using native OS checks and sys.gettrace."""
    if os.environ.get("REGISTRODOC_DEV_MODE") == "1":
        return False
    # check 1: sys.gettrace
    if sys.gettrace() is not None:
        return True
    
    # check 2: IsDebuggerPresent on Windows
    if sys.platform == "win32":
        try:
            if ctypes.windll.kernel32.IsDebuggerPresent():
                return True
        except Exception:
            pass
    return False

def detect_virtual_machine():
    """Detect common VM signatures in registries or drivers."""
    if sys.platform != "win32":
        return False
        
    vm_drivers = [
        r"C:\windows\System32\Drivers\VBoxMouse.sys",
        r"C:\windows\System32\Drivers\VBoxGuest.sys",
        r"C:\windows\System32\Drivers\vmmouse.sys",
        r"C:\windows\System32\Drivers\vmhgfs.sys"
    ]
    for driver in vm_drivers:
        if os.path.exists(driver):
            return True
            
    # Check common registry indicators of VirtualBox/VMware/QEMU
    # We use ctypes/winreg safely
    import winreg
    vm_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\ACPI\DSDT\VBOX__"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Oracle\VirtualBox Guest Additions"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\VMware, Inc.\VMware Tools")
    ]
    for hkey, path in vm_keys:
        try:
            key = winreg.OpenKey(hkey, path)
            winreg.CloseKey(key)
            return True
        except WindowsError:
            pass
            
    return False

def get_running_processes():
    """List running processes using Toolhelp32Snapshot to avoid subprocess flags."""
    if sys.platform != "win32":
        return []
        
    processes = []
    TH32CS_SNAPPROCESS = 0x00000002
    ctypes.windll.kernel32.CreateToolhelp32Snapshot.argtypes = [ctypes.c_ulong, ctypes.c_ulong]
    ctypes.windll.kernel32.CreateToolhelp32Snapshot.restype = ctypes.c_void_p
    
    hSnapshot = ctypes.windll.kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if hSnapshot == -1 or hSnapshot is None:
        return []
        
    pe = PROCESSENTRY32()
    pe.dwSize = ctypes.sizeof(PROCESSENTRY32)
    
    if ctypes.windll.kernel32.Process32First(hSnapshot, ctypes.byref(pe)):
        while True:
            exe_name = pe.szExeFile.decode('ansi', errors='ignore').lower()
            processes.append(exe_name)
            if not ctypes.windll.kernel32.Process32Next(hSnapshot, ctypes.byref(pe)):
                break
                
    ctypes.windll.kernel32.CloseHandle(hSnapshot)
    return processes

def detect_analysis_tools():
    """Scan active processes for analysis or cracking software."""
    running = get_running_processes()
    for proc in SUSPICIOUS_PROCESSES:
        if proc in running:
            return True
    return False

def run_anti_analysis_checks():
    """Run all checks and terminate immediately if any trigger."""
    if os.environ.get("REGISTRODOC_DEV_MODE") == "1":
        return
    if detect_debugger() or detect_virtual_machine() or detect_analysis_tools():
        # Exit silently to prevent malware analyst from stepping through crash handlers
        os._exit(0)

def get_inactivity_time():
    """Return user inactivity time in seconds using native Windows GetLastInputInfo."""
    if sys.platform != "win32":
        return 0.0
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
        ticks = ctypes.windll.kernel32.GetTickCount()
        # Handle tick count rollover
        elapsed = ticks - lii.dwTime
        return max(0.0, elapsed / 1000.0)
    return 0.0
