import paramiko
import os
import time
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
  
  
class SftpTransferService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'SftpTransferService'
    _svc_display_name_ = 'SFTP Transfer Service'
  
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
        host = 'Host'
        password = "pass"
        username = "User"
  
        transfer_freq = 10
  
        while self.is_alive:
            try:
                if self.sftp is None or self.sftp.get_channel().closed:
                    self.connect_to_sftp(host, username, password)
  
                pdf_dir = "/home/prddba/reports/Print/In"
                local_folder = r"D:\Python_local"
  
                self.transfer_files(pdf_dir, local_folder)
            except Exception as e:
                print(f"Error in main loop: {e}")
  
            time.sleep(transfer_freq)
  
        if self.sftp:
            self.sftp.close()
  
    def connect_to_sftp(self, host, username, password):
        transport = paramiko.Transport((host, 22))
        transport.connect(username=username, password=password)
        self.sftp = paramiko.SFTPClient.from_transport(transport)
  
    def transfer_files(self, pdf_dir, local_folder):
        try:
            # Retrieve a list of files on the SFTP server
            sftp_files = set(self.sftp.listdir(pdf_dir))
  
            # Retrieve a list of files in the local folder
            local_files = set(os.listdir(local_folder))
  
            # Calculate the files that are on the local folder but not on the SFTP server
            files_to_upload = local_files.difference(sftp_files)
  
            for local_file in files_to_upload:
                # Build file paths
                if not local_file.startswith("WA"):
                    local_file_path = os.path.join(local_folder, local_file)
                    sftp_file_path = f"{pdf_dir}/{local_file}"
  
                    # Upload file to SFTP directory
                    self.sftp.put(local_file_path, sftp_file_path)
  
                    print(f"Uploaded {local_file} to {pdf_dir}")
  
                    # Optional: You can remove the local file after uploading
                    os.remove(local_file_path)
                    print(f"Removed {local_file} from {local_folder}")
  
        except Exception as e:
            print(f"Error in transfer_files: {e}")
  
    def sftp_file_exists(self, sftp, file_path):
        try:
            sftp.stat(file_path)
            return True
        except FileNotFoundError:
            return False
  
  
if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SftpTransferService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SftpTransferService)
