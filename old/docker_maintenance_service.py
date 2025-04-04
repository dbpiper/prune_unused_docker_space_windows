import os
import sys
import time
import datetime
import subprocess
import win32serviceutil
import win32service
import win32event
import win32evtlogutil
import win32evtlog  # For error type constants.
import win32con
from concurrent.futures import ThreadPoolExecutor

# Fixed log file and startup file locations.
FIXED_LOG_FILE = r"C:\Temp\maintenance_service.log"
STARTUP_FILE = r"C:\Temp\startup.txt"


def ensure_log_folder() -> None:
    folder = os.path.dirname(FIXED_LOG_FILE)
    if not os.path.exists(folder):
        try:
            os.makedirs(folder)
        except Exception as e:
            sys.stderr.write("Failed to create log folder: " + str(e) + "\n")


def clear_log_file() -> None:
    """Clear the log file contents at service start."""
    try:
        ensure_log_folder()
        with open(FIXED_LOG_FILE, "w") as f:
            f.truncate(0)
    except Exception as e:
        sys.stderr.write("Failed to clear log file: " + str(e) + "\n")


def write_local_log(message: str) -> None:
    try:
        ensure_log_folder()
        with open(FIXED_LOG_FILE, "a") as f:
            f.write(f"{datetime.datetime.now().isoformat()}: {message}\n")
            f.flush()
    except Exception as e:
        sys.stderr.write("File logging failed: " + str(e) + "\n")


def write_startup_file() -> None:
    try:
        ensure_log_folder()
        with open(STARTUP_FILE, "w") as f:
            f.write("Service started at " + str(datetime.datetime.now()))
            f.flush()
    except Exception as e:
        sys.stderr.write("Failed to write startup file: " + str(e) + "\n")


def get_docker_vhd_path() -> list:
    r"""
    Returns the Docker VHD file path for user 'david' located at:
    C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhd
    """
    vhd_paths = []
    target_path = (
        r"C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhd"
    )
    write_local_log(f"DEBUG: Checking for Docker VHD at: {target_path}")
    if os.path.isfile(target_path):
        vhd_paths.append(target_path)
        write_local_log(f"DEBUG: Found Docker VHD file at: {target_path}")
    else:
        write_local_log(f"DEBUG: Docker VHD file not found at: {target_path}")
    return vhd_paths


class DockerMaintenanceService(win32serviceutil.ServiceFramework):
    _svc_name_ = "DockerMaintenanceService"
    _svc_display_name_ = "Docker Maintenance Service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running = False
        win32event.SetEvent(self.stop_event)
        self.log_event("Service stop requested.")

    def SvcDoRun(self):
        try:
            # Clear old logs and write startup file.
            clear_log_file()
            write_startup_file()
            write_local_log("DEBUG: SvcDoRun() entered.")
            write_local_log(
                "DEBUG: Skipped servicemanager.LogMsg in SvcDoRun()."
            )
            self.log_event("Service started.")
            write_local_log("DEBUG: After self.log_event('Service started.')")
            self.log_event(
                "Service environment PATH: " + os.environ.get("PATH", "")
            )
            write_local_log("DEBUG: After logging environment PATH.")
            self.log_event("DEBUG: Starting immediate maintenance tasks.")
            write_local_log("DEBUG: Before calling run_maintenance_tasks().")
            self.run_maintenance_tasks()
            write_local_log("DEBUG: After calling run_maintenance_tasks().")
            self.main()
        except Exception as e:
            write_local_log("Exception in SvcDoRun: " + str(e))
            raise

    def main(self):
        while self.is_running:
            now = datetime.datetime.now()
            scheduled_time = now.replace(
                hour=2, minute=0, second=0, microsecond=0
            )
            if now >= scheduled_time:
                scheduled_time += datetime.timedelta(days=1)
            wait_seconds = (scheduled_time - now).total_seconds()
            self.log_event(
                f"DEBUG: Waiting for {wait_seconds:.0f} seconds until next run at {scheduled_time}"
            )
            ret = win32event.WaitForSingleObject(
                self.stop_event, int(wait_seconds * 1000)
            )
            if ret == win32event.WAIT_OBJECT_0:
                break
            self.log_event(
                "DEBUG: Scheduled time reached. Starting maintenance tasks."
            )
            self.run_maintenance_tasks()
        self.log_event("DEBUG: Service main loop exited; stopping service.")

    def check_docker_running(self) -> bool:
        try:
            self.run_command("docker", "info")
            return True
        except Exception as e:
            self.log_event(
                "Docker daemon is not running: " + str(e),
                eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
            )
            return False

    def run_maintenance_tasks(self):
        try:
            self.log_event(
                "DEBUG: Entered run_maintenance_tasks() at "
                + str(datetime.datetime.now())
            )
            if not self.check_docker_running():
                self.log_event(
                    "Docker is not running; attempting to start Docker Desktop.",
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                )
                docker_desktop_path = (
                    r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
                )
                self.start_docker_desktop(docker_desktop_path)
                self.log_event("Waiting 60 seconds for Docker to start...")
                time.sleep(60)
                if not self.check_docker_running():
                    self.log_event(
                        "Docker still isn't running after attempting to start it.",
                        eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                    )
                    return
                else:
                    self.log_event("Docker is now running.")
            container_count_before = self.get_docker_container_count()
            self.log_event(
                "Docker container count before prune: "
                + str(container_count_before)
            )
            prune_output = self.run_command(
                "powershell", "docker system prune -a --volumes --force"
            )
            self.log_event("Docker prune output: " + prune_output)
            container_count_after = self.get_docker_container_count()
            self.log_event(
                "Docker container count after prune: "
                + str(container_count_after)
            )
            removed_containers = container_count_before - container_count_after
            self.log_event(
                "Removed {} docker containers during prune.".format(
                    removed_containers
                )
            )
            self.kill_docker_processes()
            try:
                wsl_output = self.run_command("wsl", "--shutdown")
                self.log_event("WSL shutdown executed. Output: " + wsl_output)
            except Exception as e:
                self.log_event(
                    "WSL shutdown failed, continuing: " + str(e),
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                )
            vhd_paths = get_docker_vhd_path()
            if not vhd_paths:
                self.log_event(
                    "No Docker VHD files found for user 'david'.",
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                )
            else:
                for path in vhd_paths:
                    optimize_cmd = f'Optimize-VHD -Path "{path}" -Mode Full'
                    self.log_event("Optimizing VHD: " + path)
                    self.run_command("powershell", optimize_cmd)
                    self.log_event(
                        "Docker VHD optimization executed for: " + path
                    )
            docker_desktop_path = (
                r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
            )
            self.start_docker_desktop(docker_desktop_path)
            self.log_event("Docker Desktop restarted minimized.")
            self.log_event("DEBUG: Maintenance tasks completed successfully.")
        except Exception as e:
            self.log_event(
                "Error during maintenance tasks: " + str(e),
                eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
            )

    def kill_docker_processes(self) -> None:
        for proc_name in ["docker.exe", "Docker Desktop.exe"]:
            try:
                self.kill_process(proc_name)
                self.log_event(f"Killed {proc_name} processes if any.")
            except Exception as e:
                self.log_event(
                    f"Failed to kill process {proc_name}: " + str(e),
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                )

    def get_docker_container_count(self) -> int:
        try:
            result = subprocess.run(
                ["docker", "ps", "-a", "-q"],
                capture_output=True,
                text=True,
                timeout=30,
            )
            container_ids = [
                line.strip()
                for line in result.stdout.strip().splitlines()
                if line.strip()
            ]
            return len(container_ids)
        except Exception as e:
            self.log_event(
                "Failed to get docker container count: " + str(e),
                eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
            )
            return 0

    def run_command(self, command: str, args: str) -> str:
        self.log_event("DEBUG: Executing command: {} {}".format(command, args))
        if command.lower() == "powershell":
            cmd = [command, "-Command", args]
        else:
            cmd = [command, args]
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=30
            )
        except subprocess.TimeoutExpired:
            error_msg = f"Command timeout expired: {command} {args}"
            self.log_event(
                error_msg, eventType=win32evtlog.EVENTLOG_ERROR_TYPE
            )
            raise Exception(error_msg)
        if result.returncode != 0:
            error_msg = "Command failed: {} {}. Error: {} (Stdout: {})".format(
                command, args, result.stderr, result.stdout
            )
            self.log_event(
                error_msg, eventType=win32evtlog.EVENTLOG_ERROR_TYPE
            )
            raise Exception(error_msg)
        output = result.stdout.strip()
        self.log_event("DEBUG: Command output: " + output)
        return output

    def kill_process(self, process_name: str) -> None:
        self.log_event("DEBUG: Killing processes named: " + process_name)
        result = subprocess.run(
            ["taskkill", "/F", "/IM", process_name],
            capture_output=True,
            text=True,
            timeout=30,
        )
        self.log_event("Taskkill output: " + result.stdout.strip())

    def start_docker_desktop(self, executable_path: str) -> None:
        """
        Attempts to start Docker Desktop using the Windows 'start' command via cmd
        so that it launches in an interactive session under user 'david'.
        This method logs detailed diagnostic information and uses a retry mechanism.
        Note: To fully run as your user "david", configure the service’s “Log On” account in
        the Windows Services console (services.msc) to use your user account.
        """
        max_retries = 3
        retry_delay = 5  # seconds between attempts
        attempt = 0
        self.log_event(
            "DEBUG: Attempting to start Docker Desktop via cmd: "
            + executable_path
        )
        if not os.path.exists(executable_path):
            self.log_event(
                "ERROR: Docker Desktop executable not found: "
                + executable_path,
                eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
            )
            return
        while attempt < max_retries:
            try:
                # Use the Windows "start" command via cmd with an empty title.
                cmd = f'cmd /c start "" "{executable_path}"'
                self.log_event("DEBUG: Running command: " + cmd)
                result = subprocess.run(
                    cmd, capture_output=True, text=True, shell=True, timeout=30
                )
                self.log_event(
                    "DEBUG: Docker Desktop start stdout: " + result.stdout
                )
                self.log_event(
                    "DEBUG: Docker Desktop start stderr: " + result.stderr
                )
                time.sleep(retry_delay)
                # Use tasklist to check if Docker Desktop is running.
                tasklist = subprocess.run(
                    'tasklist /FI "IMAGENAME eq Docker Desktop.exe"',
                    capture_output=True,
                    text=True,
                    shell=True,
                )
                if "Docker Desktop.exe" in tasklist.stdout:
                    self.log_event(
                        f"DEBUG: Docker Desktop is running (Attempt {attempt+1})."
                    )
                    return
                else:
                    self.log_event(
                        f"ERROR: Docker Desktop does not appear to be running (Attempt {attempt+1}).",
                        eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                    )
            except Exception as e:
                self.log_event(
                    f"ERROR: Failed to start Docker Desktop on attempt {attempt+1}: {str(e)}",
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                )
            attempt += 1
            self.log_event(
                f"DEBUG: Retrying Docker Desktop launch in {retry_delay} seconds..."
            )
            time.sleep(retry_delay)
        self.log_event(
            "ERROR: Failed to start Docker Desktop after maximum retries.",
            eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
        )

    def log_event(
        self, message: str, eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE
    ) -> None:
        write_local_log(message)


if __name__ == "__main__":
    # Note: To fully run as your user "david", configure the service’s “Log On” account in the Windows Services console to use your user account.
    win32serviceutil.HandleCommandLine(DockerMaintenanceService)
