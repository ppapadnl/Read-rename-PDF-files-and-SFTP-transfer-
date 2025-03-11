import paramiko
import os
import time
import win32serviceutil
import win32service
import win32event
import servicemanager
import shutil
import socket
import sys
  
  
class PullAndArchService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'PullAndArchService'
    _svc_display_name_ = 'Pull And Archive Service'
  
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_alive = True
        self.sftp = None
  
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
  
    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()
  
    def main(self):
        host = 'host'
        password = "pass"
        username = "User"
  
        transfer_freq = 15
  
        while self.is_alive:
            try:
                if self.sftp is None or self.sftp.get_channel().closed:
                    self.connect_to_sftp(host, username, password)
  
                pdf_dir = "/doc"
                local_folder = r"D:\Python_local"
                local_archive_folder = r"D:\Python_archive"
  
                self.transfer_files(pdf_dir, local_folder, local_archive_folder)
            except Exception as e:
                print(f"Error in main loop: {e}")
  
            time.sleep(transfer_freq)
  
    def connect_to_sftp(self, host, username, password):
        transport = paramiko.Transport((host, 22))  # 22 is the port
        transport.connect(username=username, password=password)
        self.sftp = paramiko.SFTPClient.from_transport(transport)
  
    def transfer_files(self, pdf_dir, local_folder, local_archive_folder):
        try:
            files = self.sftp.listdir(pdf_dir)
  
            for pdf_file in files:
                if pdf_file.startswith("WA"):
                    # Build file paths
                    sftp_file_path = f"{pdf_dir}/{pdf_file}"
                    local_file_path = os.path.join(local_folder, pdf_file)
                    local_archive_file_path = os.path.join(local_archive_folder, pdf_file)
  
                    # Check if the file exists in the local folder
                    if not os.path.exists(local_file_path):
                        # Download file to local folder
                        self.sftp.get(sftp_file_path, local_file_path)
  
                        # Copy the file to the local archive
                        shutil.copy(local_file_path, local_archive_file_path)
  
                        # Remove the file from the original directory on the SFTP server
                        self.sftp.remove(sftp_file_path)
  
                        print(f"Archived {pdf_file} to {local_folder} and {local_archive_folder}")
        except Exception as e:
            print(f"Error in transfer_files: {e}")
  
  
if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PullAndArchService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(PullAndArchService)
