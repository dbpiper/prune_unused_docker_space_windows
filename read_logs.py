#!/usr/bin/env python

import win32evtlog
import sys


def read_service_logs(
    service_name, server="localhost", log_type="Application"
):
    """
    Reads and prints Windows Event Log entries for the specified service.

    :param service_name: The source name of the service (e.g., "DockerMaintenanceService")
    :param server: The server to query (default 'localhost')
    :param log_type: The event log type (default 'Application')
    """
    try:
        # Open the specified event log.
        log_handle = win32evtlog.OpenEventLog(server, log_type)
    except Exception as e:
        print(f"Error opening event log: {e}")
        sys.exit(1)

    # Set the flags for reading the log in reverse order (most recent first)
    flags = (
        win32evtlog.EVENTLOG_BACKWARDS_READ
        | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    )

    print(f"Reading logs for service: {service_name}\n{'=' * 60}")
    total_records = win32evtlog.GetNumberOfEventLogRecords(log_handle)
    print("Total records in log:", total_records)

    # Loop through the event log records.
    while True:
        events = win32evtlog.ReadEventLog(log_handle, flags, 0)
        if not events:
            break
        for event in events:
            # Filter events by source name.
            if event.SourceName == service_name:
                # Format the event time if possible.
                if hasattr(event.TimeGenerated, "Format"):
                    event_time = event.TimeGenerated.Format()
                else:
                    event_time = str(event.TimeGenerated)
                print("Time:", event_time)
                print("Event ID:", event.EventID)
                # The event message might be stored in StringInserts.
                if event.StringInserts:
                    print("Message:", " ".join(event.StringInserts))
                else:
                    print("Message: (No message)")
                print("-" * 60)

    win32evtlog.CloseEventLog(log_handle)


if __name__ == "__main__":
    # Specify the service name as defined in your Windows service.
    service_name = "DockerMaintenanceService"
    read_service_logs(service_name)
