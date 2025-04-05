import os
import sys
import time
import datetime
import subprocess
import win32serviceutil
import win32service
import win32event
import win32api
import win32con

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
    Returns the Docker VHDX file path for user 'david' located at:
    C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhdx
    """
    vhd_paths = []
    target_path = (
        r"C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhdx"
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
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def log_event(self, message: str, eventType=None) -> None:
        write_local_log(message)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running = False
        win32event.SetEvent(self.stop_event)
        write_local_log("Service stop requested.")

    def SvcDoRun(self):
        try:
            clear_log_file()
            write_startup_file()
            # Let SCM know we’re up and running
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)

            write_local_log("DEBUG: SvcDoRun() entered.")
            self.log_event("Service started.")
            self.log_event(
                "Service environment PATH: " + os.environ.get("PATH", "")
            )
            self.log_event("DEBUG: Starting immediate maintenance tasks.")
            self.run_maintenance_tasks()
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
            write_local_log("Docker daemon is not running: " + str(e))
            return False

    def run_maintenance_tasks(self):
        try:
            self.log_event(
                "DEBUG: Entered run_maintenance_tasks() at "
                + str(datetime.datetime.now())
            )
            if not self.check_docker_running():
                write_local_log(
                    "Docker is not running; attempting to start Docker Desktop."
                )
                docker_desktop_path = (
                    r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
                )
                self.start_docker_desktop(docker_desktop_path)
                write_local_log("Waiting 60 seconds for Docker to start...")
                time.sleep(60)
                if not self.check_docker_running():
                    write_local_log(
                        "Docker still isn't running after attempting to start it."
                    )
                    return
                else:
                    write_local_log("Docker is now running.")

            container_count_before = self.get_docker_container_count()
            write_local_log(
                "Docker container count before prune: "
                + str(container_count_before)
            )
            prune_output = self.run_command(
                "powershell", "docker system prune -a --volumes --force"
            )
            write_local_log("Docker prune output: " + prune_output)
            container_count_after = self.get_docker_container_count()
            write_local_log(
                "Docker container count after prune: "
                + str(container_count_after)
            )
            removed_containers = container_count_before - container_count_after
            write_local_log(
                f"Removed {removed_containers} docker containers during prune."
            )
            self.kill_docker_processes()

            # Stop the WSL service so the VHDX isn’t locked
            write_local_log("Stopping WSL service (WslService).")
            self.run_command("powershell", "Stop-Service WslService -Force")
            write_local_log("WslService stopped.")

            vhd_paths = get_docker_vhd_path()
            if not vhd_paths:
                write_local_log("No Docker VHD files found for user 'david'.")
            else:
                for path in vhd_paths:
                    optimize_cmd = f'Optimize-VHD -Path "{path}" -Mode Full'
                    write_local_log("Optimizing VHD: " + path)
                    self.run_command("powershell", optimize_cmd)
                    write_local_log(
                        "Docker VHD optimization executed for: " + path
                    )

            # Restart the WSL service and wait for it
            write_local_log("Starting WSL service (WslService).")
            self.run_command("powershell", "Start-Service WslService")
            write_local_log("WslService start requested.")
            write_local_log(
                "DEBUG: Waiting for WslService to reach 'Running' state with timeout."
            )
            start_time = time.time()
            timeout = 60  # seconds
            while True:
                try:
                    status = self.run_command(
                        "powershell", "(Get-Service WslService).Status"
                    ).strip()
                except Exception as e:
                    write_local_log(
                        "ERROR: Failed to query WslService status: " + str(e)
                    )
                    break
                if status == "Running":
                    write_local_log("WslService is now running.")
                    break
                if time.time() - start_time > timeout:
                    write_local_log(
                        f"ERROR: Timeout waiting for WslService to start after {timeout} seconds."
                    )
                    break
                time.sleep(1)

            # Enhanced WSL status check: verify WSL operational status by executing a simple WSL command.
            try:
                wsl_check_output = self.run_command("wsl.exe", "-l")
                if not wsl_check_output:
                    write_local_log(
                        "ERROR: WSL command returned empty output. WSL might not be operational."
                    )
                else:
                    write_local_log(
                        "WSL operational check succeeded. Output: "
                        + wsl_check_output
                    )
            except Exception as e:
                write_local_log(
                    "ERROR: Exception during WSL operational check: " + str(e)
                )

            # Now start Docker Desktop
            docker_desktop_path = (
                r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
            )
            self.start_docker_desktop(docker_desktop_path)
            write_local_log(
                "Docker Desktop restart attempted with exponential backoff."
            )

            # Additional check: if Docker daemon still isn't responsive, attempt one extra restart
            if not self.check_docker_running():
                write_local_log(
                    "WARNING: Docker daemon still not accessible after startup attempt. Attempting one additional restart."
                )
                self.kill_docker_processes()
                self.start_docker_desktop(docker_desktop_path)
                if self.check_docker_running():
                    write_local_log(
                        "Docker daemon responsive after additional restart."
                    )
                else:
                    write_local_log(
                        "ERROR: Docker daemon still not responsive after additional restart."
                    )

            write_local_log("DEBUG: Maintenance tasks completed successfully.")
        except Exception as e:
            write_local_log("Error during maintenance tasks: " + str(e))

    def kill_docker_processes(self) -> None:
        for proc_name in ["docker.exe", "Docker Desktop.exe"]:
            try:
                self.kill_process(proc_name)
                write_local_log(f"Killed {proc_name} processes if any.")
            except Exception as e:
                write_local_log(
                    f"Failed to kill process {proc_name}: {str(e)}"
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
            write_local_log("Failed to get docker container count: " + str(e))
            return 0

    def run_command(self, command: str, args: str) -> str:
        write_local_log(f"DEBUG: Executing command: {command} {args}")
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
            write_local_log(error_msg)
            raise Exception(error_msg)
        if result.returncode != 0:
            error_msg = f"Command failed: {command} {args}. Error: {result.stderr} (Stdout: {result.stdout})"
            write_local_log(error_msg)
            raise Exception(error_msg)
        output = result.stdout.strip()
        write_local_log("DEBUG: Command output: " + output)
        return output

    def kill_process(self, process_name: str) -> None:
        write_local_log("DEBUG: Killing processes named: " + process_name)
        result = subprocess.run(
            ["taskkill", "/F", "/IM", process_name],
            capture_output=True,
            text=True,
            timeout=30,
        )
        write_local_log("Taskkill output: " + result.stdout.strip())

    def start_docker_desktop(self, executable_path: str) -> None:
        """
        Attempts to start Docker Desktop using ShellExecute via win32api.
        Implements exponential backoff on retries and verifies Docker daemon responsiveness.
        """
        max_retries = 3
        retry_delay = 5  # seconds
        attempt = 0
        write_local_log(
            "DEBUG: Attempting to start Docker Desktop: " + executable_path
        )
        if not os.path.exists(executable_path):
            write_local_log(
                "ERROR: Docker Desktop executable not found: "
                + executable_path
            )
            return

        while attempt < max_retries:
            try:
                win32api.ShellExecute(
                    0,
                    "open",
                    executable_path,
                    None,
                    None,
                    win32con.SW_SHOWNORMAL,
                )
                write_local_log(
                    f"DEBUG: ShellExecute() succeeded (Attempt {attempt+1})."
                )
            except Exception as e:
                write_local_log(
                    f"ERROR: ShellExecute() failed on attempt {attempt+1}: {str(e)}"
                )
            # Wait before checking if the process is running.
            time.sleep(retry_delay)
            tasklist = subprocess.run(
                'tasklist /FI "IMAGENAME eq Docker Desktop.exe"',
                capture_output=True,
                text=True,
                shell=True,
            )
            if "Docker Desktop.exe" in tasklist.stdout:
                write_local_log(
                    f"DEBUG: Docker Desktop is running (Attempt {attempt+1})."
                )
                # Verify Docker daemon responsiveness.
                try:
                    self.run_command("docker", "info")
                    write_local_log(
                        "Docker daemon is responsive after startup."
                    )
                    return
                except Exception as e:
                    write_local_log(
                        "ERROR: Docker daemon not responsive: " + str(e)
                    )
            else:
                write_local_log(
                    f"ERROR: Docker Desktop does not appear to be running (Attempt {attempt+1})."
                )
            attempt += 1
            retry_delay *= 2
            write_local_log(
                f"DEBUG: Retrying Docker Desktop launch in {retry_delay} seconds..."
            )
            time.sleep(retry_delay)
        write_local_log(
            "ERROR: Failed to start Docker Desktop after maximum retries."
        )


if __name__ == "__main__":
    from win32serviceutil import StopService, WaitForServiceStatus
    from win32service import SERVICE_STOPPED

    svc_name = DockerMaintenanceService._svc_name_
    if "update" in sys.argv:
        write_local_log(
            "INFO: update requested, stopping service before update"
        )
        try:
            StopService(svc_name)
            WaitForServiceStatus(
                svc_name, SERVICE_STOPPED, 30_000
            )  # timeout in ms
            write_local_log("INFO: service stopped successfully")
        except Exception as e:
            write_local_log(f"WARN: could not stop service before update: {e}")

    win32serviceutil.HandleCommandLine(DockerMaintenanceService)
