# Docker Maintenance Windows Service

A Python-based Windows Service that performs nightly Docker maintenance tasks. This service automates routine Docker cleanup by pruning unused Docker resources, killing Docker processes, shutting down WSL, optimizing the Docker VHD, and restarting Docker Desktop (minimized) on a daily schedule.

> **Note:** This service requires administrative privileges to run and execute elevated commands.

---

## Table of Contents

- [Docker Maintenance Windows Service](#docker-maintenance-windows-service)
  - [Table of Contents](#table-of-contents)
  - [Features](#features)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Configuration](#configuration)
  - [Usage](#usage)
    - [Installing the Service](#installing-the-service)
    - [Starting the Service](#starting-the-service)
    - [Stopping the Service](#stopping-the-service)
    - [Removing the Service](#removing-the-service)
  - [How It Works](#how-it-works)
  - [Troubleshooting](#troubleshooting)
  - [Contributing](#contributing)
  - [License](#license)
  - [Additional Resources](#additional-resources)

---

## Features

- **Automated Scheduling:** Runs once per day at a scheduled time (default is 2:00 AM).
- **Docker Maintenance:**
  - **Prune Docker Resources:** Runs `docker system prune -a --volumes` via PowerShell.
  - **Kill Docker Processes:** Terminates running Docker processes using `taskkill`.
  - **Shutdown WSL:** Executes `wsl --shutdown`.
  - **Optimize Docker VHD:** Runs `Optimize-VHD` to shrink the Docker VHD.
  - **Restart Docker Desktop:** Restarts Docker Desktop in a minimized state.
- **Robust Logging:** Uses Windows Event Log to record information, errors, and progress.
- **Graceful Shutdown:** Listens for stop requests and cleans up resources accordingly.

---

## Prerequisites

- **Operating System:** Windows 7 or later (with support for Windows Services)
- **Python:** 3.6 or later
- **Administrative Rights:** Required for installation and execution (to run elevated commands)
- **Dependencies:**
  - [pywin32](https://github.com/mhammond/pywin32)
    ```bash
    pip install pywin32
    ```

---

## Installation

1. **Clone the Repository:**

```bash
git clone https://github.com/yourusername/docker-maintenance-service.git
cd docker-maintenance-service
```

2. **Set Up a Virtual Environment (Recommended):**

```bash
python -m venv venv
venv\Scripts\activate
```

3. **Install Python Dependencies:**

```bash
pip install pywin32
```

---

## Configuration

Adjust parameters in `docker_maintenance_service.py` as needed:

- **Scheduled Time:**

```python
scheduled_time = now.replace(hour=2, minute=0, second=0, microsecond=0)
```

- **Docker Commands:** Adjust if necessary.

- **Docker Desktop Path:**

```python
docker_desktop_path = r"C:\Program Files\Docker\Docker\Docker Desktop.exe"
```

- **Optimize VHD Command:**

```python
optimize_cmd = "Optimize-VHD -Path \"$env:LOCALAPPDATA\\Docker\\wsl\\disk\\docker_data.vhdx\" -Mode Full"
```

---

## Usage

### Installing the Service

```bash
python docker_maintenance_service.py install
```

### Starting the Service

```bash
python docker_maintenance_service.py start
```

### Stopping the Service

```bash
python docker_maintenance_service.py stop
```

### Removing the Service

```bash
python docker_maintenance_service.py remove
```

---

## How It Works

- **Initialization:** Service inherits from `win32serviceutil.ServiceFramework`.
- **Scheduled Execution:** Performs tasks at scheduled time, logs each step.
- **Error Handling:** Logs errors, continues operation.
- **Graceful Shutdown:** Cleans up resources on stop request.

---

## Troubleshooting

- **Service Issues:** Check administrative privileges, verify paths.
- **Command Failures:** Run commands manually via PowerShell to identify errors.
- **Logging Issues:** Ensure permissions to write to Windows Event Log.

---

## Contributing

Contributions are welcome!

- Fork the repository.
- Create a branch:
  ```bash
  git checkout -b feature/your-feature-name
  ```
- Commit and push your changes.
- Open a pull request.

---

## License

Licensed under the MIT License.

---

## Additional Resources

- [pywin32 Documentation](https://github.com/mhammond/pywin32)
- [Windows Service Programming in Python](https://github.com/mhammond/pywin32)
- [Docker Documentation](https://docs.docker.com/)
