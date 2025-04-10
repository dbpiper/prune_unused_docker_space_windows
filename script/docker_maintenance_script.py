import os
import sys
import time
import datetime
import subprocess
import win32evtlogutil
import win32evtlog  # For error type constants.
import win32con
from concurrent.futures import ThreadPoolExecutor


def log(message: str, eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE) -> None:
    # Print the message with a timestamp to the console.
    timestamp = datetime.datetime.now().isoformat()
    print(f"{timestamp}: {message}")


def get_docker_vhd_path() -> list[str]:
    r"""
    Returns the Docker VHDX file path for user 'david' located at:
    C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhdx
    """
    vhd_paths: list[str] = []
    target_path = (
        r"C:\Users\david\AppData\Local\Docker\wsl\disk\docker_data.vhdx"
    )
    log(f"DEBUG: Checking for Docker VHDX at: {target_path}")
    if os.path.isfile(target_path):
        vhd_paths.append(target_path)
        log(f"DEBUG: Found Docker VHDX file at: {target_path}")
    else:
        log(f"DEBUG: Docker VHDX file not found at: {target_path}")
    return vhd_paths


class DockerMaintenance:
    def log_event(
        self, message: str, eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE
    ) -> None:
        log(message, eventType)

    def run_command(self, command: str, args: str) -> str:
        self.log_event(f"DEBUG: Executing command: {command} {args}")
        # build invocation
        if command.lower() == "powershell":
            cmd = [command, "-Command", args]
        else:
            cmd = [command, args]

        # spawn and stream stdout+stderr
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
            )
        except Exception as e:
            msg = f"Failed to start process: {command} {args}: {e}"
            self.log_event(msg, eventType=win32evtlog.EVENTLOG_ERROR_TYPE)
            raise

        output_lines: list[str] = []
        assert proc.stdout is not None
        for raw in proc.stdout:
            line = raw.rstrip()
            # immediate console feedback
            timestamp = datetime.datetime.now().isoformat()
            print(f"{timestamp}: {line}")
            output_lines.append(line)

        try:
            ret = proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
            msg = f"Command timeout expired: {command} {args}"
            self.log_event(msg, eventType=win32evtlog.EVENTLOG_ERROR_TYPE)
            raise

        if ret != 0:
            msg = f"Command failed: {command} {args}. Exit code: {ret}"
            self.log_event(msg, eventType=win32evtlog.EVENTLOG_ERROR_TYPE)
            raise Exception(msg)

        return "\n".join(output_lines)

    def kill_process(self, process_name: str) -> None:
        # delegate to run_command so output streams live
        self.log_event(f"DEBUG: Killing processes named: {process_name}")
        self.run_command("taskkill", f"/F /IM {process_name}")
        self.log_event(f"Killed processes named: {process_name}")

    def kill_docker_processes(self) -> None:
        for proc_name in ["docker.exe", "Docker Desktop.exe"]:
            try:
                self.kill_process(proc_name)
                self.log_event(f"Killed {proc_name} processes if any.")
            except Exception as e:
                self.log_event(
                    f"Failed to kill process {proc_name}: {str(e)}",
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

    def start_docker_desktop(self, executable_path: str) -> None:
        """
        Attempts to start Docker Desktop using the Windows 'start' command via cmd.
        This method logs detailed diagnostic information and uses a retry mechanism.
        Note: To fully run as your user "david", run this script under your account.
        """
        max_retries = 3
        retry_delay = 5
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
                # start and stream
                start_cmd = f'cmd /c start "" "{executable_path}"'
                self.log_event("DEBUG: Running command: " + start_cmd)
                self.run_command("cmd", start_cmd)

                time.sleep(retry_delay)

                self.log_event("DEBUG: Checking for Docker Desktop process")
                list_output = self.run_command(
                    "cmd", 'tasklist /FI "IMAGENAME eq Docker Desktop.exe"'
                )
                if "Docker Desktop.exe" in list_output:
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

    def run_maintenance_tasks(self) -> None:
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


def main() -> None:
    log("Script started.")
    log("Environment PATH: " + os.environ.get("PATH", ""))
    dm = DockerMaintenance()
    dm.log_event("Starting maintenance tasks.")
    dm.run_maintenance_tasks()
    dm.log_event("Maintenance tasks completed.")


if __name__ == "__main__":
    # Run the maintenance tasks once.
    main()
