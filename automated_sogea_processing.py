"""
Automated SoGEA Order Processing from Outlook
1. Monitors Outlook folder for new order emails
2. Extracts Excel attachments
3. Processes orders via bulk ordering script
4. Archives results
"""

import subprocess
import sys
from pathlib import Path
from datetime import datetime
import shutil
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automated_sogea_processing.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).parent.resolve()
INCOMING_DIR = SCRIPT_DIR / "incoming_orders"
PROCESSED_DIR = INCOMING_DIR / "processed"
RESULTS_ARCHIVE = SCRIPT_DIR / "archived_results"

# Ensure directories exist
INCOMING_DIR.mkdir(exist_ok=True)
PROCESSED_DIR.mkdir(exist_ok=True)
RESULTS_ARCHIVE.mkdir(exist_ok=True)


def check_for_new_orders():
    """
    Run the Outlook monitor to check for new orders.
    Returns the number of new orders found.
    """
    logger.info("Checking Outlook for new orders...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "outlook_email_monitor.py"), "--move-to", "Processed"],
            capture_output=True,
            text=True,
            check=True
        )
        
        logger.info(result.stdout)
        
        # Count Excel files in incoming directory
        excel_files = list(INCOMING_DIR.glob("*.xlsx")) + list(INCOMING_DIR.glob("*.xls"))
        return len(excel_files)
        
    except subprocess.CalledProcessError as e:
        logger.error(f"Error checking Outlook: {e}")
        logger.error(f"Output: {e.stdout}")
        logger.error(f"Errors: {e.stderr}")
        return 0


def process_order_file(excel_file: Path):
    """
    Process a single order file using the bulk ordering script.
    """
    logger.info(f"\n{'='*60}")
    logger.info(f"Processing order file: {excel_file.name}")
    logger.info(f"{'='*60}")
    
    # Copy the file to the expected location for the bulk ordering script
    template_path = SCRIPT_DIR / "sogea_bulk_order_template.xlsx"
    backup_path = SCRIPT_DIR / f"sogea_bulk_order_template_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # Backup existing template if it exists
    if template_path.exists():
        shutil.copy(template_path, backup_path)
        logger.info(f"✓ Backed up existing template to: {backup_path.name}")
    
    # Copy the new order file as the template
    shutil.copy(excel_file, template_path)
    logger.info(f"✓ Copied {excel_file.name} to template location")
    
    try:
        # Run the bulk ordering script
        logger.info("Starting bulk order placement...")
        
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "sogea_bulk_ordering.py")],
            capture_output=True,
            text=True,
            cwd=str(SCRIPT_DIR)
        )
        
        logger.info(result.stdout)
        
        if result.returncode != 0:
            logger.error(f"Bulk ordering script failed with code {result.returncode}")
            logger.error(f"Errors: {result.stderr}")
            return False
        
        # Find the results file (most recent bulk_order_results_*.xlsx)
        result_files = sorted(SCRIPT_DIR.glob("bulk_order_results_*.xlsx"), key=lambda p: p.stat().st_mtime)
        
        if result_files:
            latest_result = result_files[-1]
            
            # Archive the results with a reference to the original file
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            archived_name = f"results_{excel_file.stem}_{timestamp}.xlsx"
            archived_path = RESULTS_ARCHIVE / archived_name
            
            shutil.copy(latest_result, archived_path)
            logger.info(f"✓ Archived results to: {archived_name}")
        
        # Move the processed order file
        processed_path = PROCESSED_DIR / excel_file.name
        shutil.move(str(excel_file), str(processed_path))
        logger.info(f"✓ Moved order file to: processed/")
        
        # Restore the backup template
        if backup_path.exists():
            shutil.move(str(backup_path), str(template_path))
            logger.info(f"✓ Restored original template")
        
        return True
        
    except Exception as e:
        logger.error(f"Error processing order: {e}", exc_info=True)
        
        # Restore backup on error
        if backup_path.exists():
            shutil.move(str(backup_path), str(template_path))
            logger.info(f"✓ Restored original template after error")
        
        return False


def main():
    """Main automation workflow"""
    logger.info("="*60)
    logger.info("Automated SoGEA Order Processing")
    logger.info("="*60)
    
    # Step 1: Check Outlook for new orders
    new_orders = check_for_new_orders()
    
    if new_orders == 0:
        logger.info("✓ No new orders found")
        return
    
    logger.info(f"✓ Found {new_orders} new order file(s)")
    
    # Step 2: Process each order file
    excel_files = list(INCOMING_DIR.glob("*.xlsx")) + list(INCOMING_DIR.glob("*.xls"))
    
    success_count = 0
    fail_count = 0
    
    for excel_file in excel_files:
        if process_order_file(excel_file):
            success_count += 1
        else:
            fail_count += 1
    
    # Step 3: Summary
    logger.info(f"\n{'='*60}")
    logger.info("PROCESSING COMPLETE")
    logger.info(f"{'='*60}")
    logger.info(f"✅ Successful: {success_count}")
    logger.info(f"❌ Failed: {fail_count}")
    logger.info(f"📁 Results archived to: {RESULTS_ARCHIVE}")
    logger.info(f"{'='*60}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Process interrupted by user")
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)
