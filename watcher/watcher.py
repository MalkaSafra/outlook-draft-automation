import time
import os
import json
import sys
import pythoncom  
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    WINDOWS = True
except ImportError:
    WINDOWS = False
    print("⚠️ win32com not available - install with: pip install pywin32")

# Set drafts folder path
SCRIPT_DIR = Path(__file__).parent
DRAFTS_FOLDER = SCRIPT_DIR.parent / "drafts"
ATTACHMENTS_FOLDER = DRAFTS_FOLDER / "attachments"

# Create folders if they don't exist
DRAFTS_FOLDER.mkdir(exist_ok=True)
ATTACHMENTS_FOLDER.mkdir(exist_ok=True)


class OutlookDraftHandler(FileSystemEventHandler):
    """Handles file creation events in the drafts folder"""
    
    def __init__(self):
        self.processed_files = set()
    
    def on_created(self, event):
        """When a new file is created"""
        if event.is_directory:
            return
        
        # Process only JSON files
        if not event.src_path.endswith('.json'):
            return
        
        # Prevent duplicate processing
        if event.src_path in self.processed_files:
            return
        
        print(f"📩 New draft file detected: {Path(event.src_path).name}")
        
        # Brief wait to ensure file is fully written
        time.sleep(0.2)
        
        self.process_draft(event.src_path)
        self.processed_files.add(event.src_path)
    
    def process_draft(self, filepath):
        """Process draft file and open Outlook"""
        try:
            # Read draft data
            with open(filepath, 'r', encoding='utf-8') as f:
                draft_data = json.load(f)
            
            print(f"   📧 To: {draft_data['to']}")
            print(f"   📝 Subject: {draft_data['subject']}")
            
            # Check if Windows and Outlook available
            if not WINDOWS:
                print("   ⚠️ Script running on non-Windows system - cannot open Outlook")
                return
            
            # Open Outlook
            self.open_outlook_draft(draft_data)
            
            # Delete file after successful processing
            time.sleep(0.5)
            os.remove(filepath)
            print(f"   ✅ Draft opened in Outlook and file deleted\n")
            
        except FileNotFoundError:
            print(f"   ❌ File not found: {filepath}")
        except json.JSONDecodeError:
            print(f"   ❌ Error reading JSON from file: {filepath}")
        except Exception as e:
            print(f"   ❌ Error processing draft: {e}\n")
    
    def open_outlook_draft(self, draft_data):
        """Open email draft in Outlook"""
        try:
            # ✅ Initialize COM - this is the solution for the error!
            pythoncom.CoInitialize()
            
            # Create Outlook object
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Create new mail item
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            # Set details
            mail.To = draft_data['to']
            mail.Subject = draft_data['subject']
            mail.Body = draft_data.get('body', '')
            
            # Attach file if exists
            if draft_data.get('attachmentPath') and os.path.exists(draft_data['attachmentPath']):
                mail.Attachments.Add(draft_data['attachmentPath'])
                print(f"   📎 File attached: {draft_data.get('attachmentName', 'file')}")
            
            # Open draft (doesn't send!)
            mail.Display()
            
            # ✅ Clean up COM
            pythoncom.CoUninitialize()
            
        except Exception as e:
            print(f"   ❌ Error opening Outlook: {e}")
            pythoncom.CoUninitialize()
            raise


def main():
    """Main function"""
    print("=" * 60)
    print("🔍 Outlook Draft Watcher")
    print("=" * 60)
    print(f"📁 Monitoring folder: {DRAFTS_FOLDER.absolute()}")
    print(f"📎 Attachments folder: {ATTACHMENTS_FOLDER.absolute()}")
    
    if not WINDOWS:
        print("\n⚠️ Warning: pywin32 not installed!")
        print("   Install with: pip install pywin32")
        print("   Script will continue running but cannot open Outlook\n")
    
    print("\n✅ Watcher is active! Waiting for draft files...")
    print("   (Press Ctrl+C to stop)\n")
    
    # Create Observer
    event_handler = OutlookDraftHandler()
    observer = Observer()
    observer.schedule(event_handler, str(DRAFTS_FOLDER), recursive=False)
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n🛑 Stopping Watcher...")
        observer.stop()
    
    observer.join()
    print("✅ Watcher stopped successfully")


if __name__ == "__main__":
    main()