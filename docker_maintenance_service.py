import os
import sys
import time
import datetime
import subprocess
import win32serviceutil
import win32service
import win32event
import servicemanager
import win32evtlogutil
import win32con

class DockerMaintenanceService(win32serviceutil.ServiceFramework):
    _svc_name_ = "DockerMaintenanceService"
    _svc_display_name_ = "Docker Maintenance Service"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        # Create an event which we will use to wait on.
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running = False
        win32event.SetEvent(self.stop_event)
        self.log_event("Service stop requested.")

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.log_event("Service started.")
        self.main()

    def main(self):
        # Main loop: wait for scheduled time (2:00 AM by default) and run tasks.
        while self.is_running:
            now = datetime.datetime.now()
            scheduled_time = now.replace(hour=2, minute=0, second=0, microsecond=0)
            # If 2:00 AM today has already passed, schedule for tomorrow.
            if now >= scheduled_time:
                scheduled_time += datetime.timedelta(days=1)
            wait_seconds = (scheduled_time - now).total_seconds()
            self.log_event("Waiting for {:.0f} seconds until next run at {}".format(wait_seconds, scheduled_time))
            
            # Wait until either the time elapses or a stop is requested.
            ret = win32event.WaitForSingleObject(self.stop_event, int(wait_seconds * 1000))
            if ret == win32event.WAIT_OBJECT_0:
                break

            self.log_event("Starting maintenance tasks.")
            self.run_maintenance_tasks()

        self.log_event("Service main loop exited; stopping service.")

    def run_maintenance_tasks(self):
        try:
            # Step 1: Prune Docker resources.
            self.run_command("powershell", "docker system prune -a --volumes")

            # Step 2: Kill Docker processes.
            self.kill_process("docker.exe")  # Adjust process name if needed.

            # Step 3: Shutdown WSL.
            self.run_command("wsl", "--shutdown")

            # Step 4: Optimize the Docker VHD.
            optimize_cmd = "Optimize-VHD -Path \"$env:LOCALAPPDATA\\Docker\\wsl\\disk\\docker_data.vhdx\" -Mode Full"
            self.run_command("powershell", optimize_cmd)

            # Step 5: Restart Docker Desktop minimized.
            docker_desktop_path = r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
            self.start_minimized(docker_desktop_path)

            self.log_event("Maintenance tasks completed successfully.")
        except Exception as e:
            self.log_event("Error during maintenance tasks: " + str(e), eventType=win32evtlogutil.EVENTLOG_ERROR_TYPE)

    def run_command(self, command, args):
        """
        Executes a command and logs its output. Raises an Exception if the command fails.
        """
        self.log_event("Executing command: {} {}".format(command, args))
        # For PowerShell commands, pass the command as an argument to "-Command"
        if command.lower() == "powershell":
            cmd = [command, "-Command", args]
        else:
            cmd = [command, args]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            error_msg = "Command failed: {} {}. Error: {}".format(command, args, result.stderr)
            self.log_event(error_msg, eventType=win32evtlogutil.EVENTLOG_ERROR_TYPE)
            raise Exception(error_msg)
        self.log_event("Command output: " + result.stdout.strip())

    def kill_process(self, process_name):
        """
        Kills all instances of the given process name using taskkill.
        """
        self.log_event("Killing processes named: " + process_name)
        result = subprocess.run(["taskkill", "/F", "/IM", process_name], capture_output=True, text=True)
        self.log_event("Taskkill output: " + result.stdout.strip())

    def start_minimized(self, executable_path):
        """
        Starts an executable minimized.
        """
        self.log_event("Starting executable minimized: " + executable_path)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = win32con.SW_MINIMIZE
        subprocess.Popen([executable_path], startupinfo=startupinfo)

    def log_event(self, message, eventType=servicemanager.EVENTLOG_INFORMATION_TYPE):
        """
        Logs an event message to the Windows Event Log.
        """
        servicemanager.LogInfoMsg(message)
        try:
            win32evtlogutil.ReportEvent(self._svc_name_, eventType, 0, [message])
        except Exception as ex:
            # In case logging fails, fallback to printing to stderr.
            sys.stderr.write("Logging failed: {}\nOriginal message: {}\n".format(str(ex), message))

if __name__ == '__main__':
    # Allow the script to be installed, started, stopped, or debugged from the command line.
    win32serviceutil.HandleCommandLine(DockerMaintenanceService)
