import os
import time
import re
import win32serviceutil
import win32service
import win32event
import servicemanager
import sys
from PyPDF2 import PdfReader
  
  
class PdfProcessingService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'PdfProcessingService'
    _svc_display_name_ = 'PDF Processing Service'
  
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
  
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
  
    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_,))
        self.main()
  
    def main(self):
        pdf_dir = r"D:\Python_local"
        invoice_number_check = re.compile(r"Invoice\s*Number\s*[:\s]*([^\n]+)", re.IGNORECASE)
        pulling_freq = 10  # Check for new files every ... seconds
  
        while self.is_alive:
            self.process_pdfs(pdf_dir, invoice_number_check)
            time.sleep(pulling_freq)
  
    def process_pdfs(self, pdf_dir, invoice_number_check):
        for pdf_file in os.listdir(pdf_dir):
            if pdf_file.startswith("WA"):
                file_path = os.path.join(pdf_dir, pdf_file)
  
                try:
                    reader = PdfReader(file_path)
                    page = reader.pages[0]
                    text = page.extract_text()
  
                    match = invoice_number_check.search(text)
                    if match:
                        invoice_number = match.group(1).strip()
                        new_file_name = f"{invoice_number}.pdf"
                        new_file_path = os.path.join(pdf_dir, new_file_name)
                        os.rename(file_path, new_file_path)
  
                        print(f"Renamed {pdf_file} to {new_file_name}")
                    else:
                        print(f"Could not find Invoice Number in {pdf_file}")
                except Exception as e:
                    print(f"Error processing {pdf_file}: {e}")
  
  
if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PdfProcessingService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(PdfProcessingService)
