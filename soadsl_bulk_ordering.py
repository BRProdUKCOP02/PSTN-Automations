"""
SOADSL Bulk Ordering Script
Reads orders from the 'SOADSL Order.xlsx' input sheet and places them
through the BMT portal via the SOADSLPortal class.

Order types supported:
- New Orders    (orderType = "New")     — columns A–AD
- Regrades      (orderType = "Regrade") — shared cols A–AD + regrade cols AE–AN
- Migrations    (orderType = "Migration") — TBD

Input sheet column layout
--------------------------
Shared (all order types):
  A  orderType                    B  account
  C  contactName                  D  contactTelephone
  E  contactEmail                 F  emailNotifications
  G  siteCompanyName              H  siteContactEmail
  I  alternaticeSiteContactName   J  alternativeSiteContactTelephone
  K  siteAddress                  L  sitePostcode
  M  nadKey                       N  newLineInstall
  O  approvedEccBanding           P  approvedTrcBanding
  Q  customerReference            R  customerRequiredDate
  S  productName                  T  careLevel
  U  voiceProduct                 V  routedIP
  W  installType                  X  routerProduct
  Y  sameAsSiteAddress            Z  deliveryCompanyName
  AA deliveryContactName          AB deliveryContactEmail
  AC deliveryContactTelephoneNumber  AD deliveryPostcode

Regrade-specific (columns AE–AN):
  AE serviceID
  AF channelPartner
  AG customerName
  AH CLI
  AI broadbandProduct
  AJ newCareLevel
  AK ipAddressOption
  AL technologyType
  AM customerRequiredDate   (regrade date — overrides shared column R for this order type)
  AN installType            (regrade install type — overrides shared column W)
"""
import pandas as pd
from soadsl_portal import SOADSLPortal, SOADSLPortalError
from graph_mailbox_check import send_order_report_email
from datetime import datetime
import os
import glob
import time
import shutil

# =============================================================================
# CONFIGURATION
# =============================================================================

SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER  = r"\\localhost\c$\users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\SOADSL Orders"
OUTPUT_FOLDER    = os.path.join(SCRIPT_DIR, "output")
PROCESSED_FOLDER = os.path.join(INPUT_FOLDER, "processed")

# Seconds to wait between portal orders to avoid overloading BMT
ORDER_DELAY = 2

# Email recipients for the output report (comma/semicolon-separated string or list)
# Set to None or '' to disable email sending
REPORT_RECIPIENTS = 'david.murphy+psmsoutput@gamma.co.uk, psmanaged.delivery@gamma.co.uk'

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# =============================================================================
# COLUMN NAME → header row mapping
# The spreadsheet uses row 0 as headers; these are the exact column names.
# =============================================================================

# ── Shared columns (all order types) ──────────────────────────────────────────
SHARED_COLS = [
    'orderType', 'account', 'contactName', 'contactTelephone', 'contactEmail',
    'emailNotifications', 'siteCompanyName', 'siteContactEmail',
    'alternaticeSiteContactName', 'alternativeSiteContactTelephone',
    'siteAddress', 'sitePostcode', 'nadKey', 'newLineInstall',
    'approvedEccBanding', 'approvedTrcBanding', 'customerReference',
    'customerRequiredDate', 'productName', 'careLevel', 'voiceProduct',
    'routedIP', 'installType', 'routerProduct', 'sameAsSiteAddress',
    'deliveryCompanyName', 'deliveryContactName', 'deliveryContactEmail',
    'deliveryContactTelephoneNumber', 'deliveryPostcode',
]

# ── Regrade-specific columns (AE–AN) ──────────────────────────────────────────
# NOTE: pandas appends .1 to column names that duplicate an earlier header.
# customerRequiredDate also appears at col R (shared), so col AM → customerRequiredDate.1
# installType       also appears at col W (shared), so col AN → installType.1
REGRADE_COLS = [
    'serviceID',              # AE — existing service to regrade
    'channelPartner',         # AF
    'customerName',           # AG
    'CLI',                    # AH — phone number on the line
    'broadbandProduct',       # AI — new product to regrade to
    'newCareLevel',           # AJ
    'ipAddressOption',        # AK
    'technologyType',         # AL
    'customerRequiredDate.1', # AM — regrade date (pandas renames duplicate header)
    'installType.1',          # AN — regrade install type (pandas renames duplicate header)
]

# Required fields for a Regrade row — validation will reject rows missing these
REGRADE_REQUIRED = ['serviceID', 'broadbandProduct', 'customerRequiredDate.1']

# =============================================================================
# HELPERS
# =============================================================================

def _str(value) -> str:
    """Return empty string for NaN/None, otherwise stripped string."""
    if pd.isna(value):
        return ''
    return str(value).strip()


def _opt(value):
    """Return None for blank/NaN, otherwise stripped string."""
    s = _str(value)
    return s if s else None


def convert_to_date_string(value) -> str | None:
    """Convert an Excel date, pandas Timestamp, or string to YYYY-MM-DD."""
    if pd.isna(value) or _str(value) == '':
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime('%Y-%m-%d')
    date_str = str(value).strip().split(' ')[0]   # strip any time component
    return date_str


def format_phone_number(value) -> str:
    """Restore leading zero that Excel may have stripped from UK numbers."""
    if pd.isna(value) or _str(value) == '':
        return ''
    phone_str = str(value).replace('.0', '').strip()
    if phone_str.isdigit() and len(phone_str) == 10 and phone_str[0] in '1237890':
        return '0' + phone_str
    return phone_str


# =============================================================================
# LOADING INPUT FILES
# =============================================================================

def get_input_files():
    """Return a sorted list of (full_path, filename_no_ext) for all xlsx/xls
    files in INPUT_FOLDER (skips Excel temp files starting with ~$)."""
    files = []
    for pattern in ('*.xlsx', '*.xls'):
        for path in glob.glob(os.path.join(INPUT_FOLDER, pattern)):
            name = os.path.basename(path)
            if not name.startswith('~$'):
                files.append((path, os.path.splitext(name)[0]))
    return sorted(files)


def load_sheet(file_path: str) -> pd.DataFrame:
    """Read Sheet1 from the workbook, using row 0 as column headers."""
    return pd.read_excel(file_path, sheet_name='Sheet1', header=0)


# =============================================================================
# VALIDATION
# =============================================================================

def validate_regrade_row(row: pd.Series, excel_row: int) -> tuple[bool, str | None]:
    """
    Validate a Regrade row.

    Args:
        row:       pandas Series for one row (column names as index)
        excel_row: 1-based Excel row number (for readable error messages)

    Returns:
        (True, None) if valid, (False, error_message) if not.
    """
    errors = []
    for field in REGRADE_REQUIRED:
        if _str(row.get(field)) == '':
            errors.append(f"missing required field: '{field}'")

    # Basic date format check
    date_val = row.get('customerRequiredDate.1')
    date_str = convert_to_date_string(date_val)
    if date_str:
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            errors.append(f"invalid customerRequiredDate '{date_val}' — expected YYYY-MM-DD")
    else:
        errors.append("missing customerRequiredDate")

    if errors:
        return False, f"Row {excel_row}: " + "; ".join(errors)
    return True, None


def validate_cancel_regrade_row(row: pd.Series, excel_row: int) -> tuple[bool, str | None]:
    """
    Validate a Regrade Cancellation row.

    Args:
        row:       pandas Series for one row (column names as index)
        excel_row: 1-based Excel row number (for readable error messages)

    Returns:
        (True, None) if valid, (False, error_message) if not.
    """
    errors = []
    
    # Only serviceID is required - date is auto-generated as today
    if _str(row.get('serviceID')) == '':
        errors.append("missing required field: 'serviceID'")

    if errors:
        return False, f"Row {excel_row}: " + "; ".join(errors)
    return True, None


def validate_all_rows(df: pd.DataFrame) -> list[str]:
    """Run validation over every row. Returns a list of error strings (empty = OK)."""
    errors = []
    for idx, row in df.iterrows():
        order_type = _str(row.get('orderType')).lower()
        excel_row = idx + 2   # +1 for 0-index, +1 for header row

        if order_type == 'regrade':
            ok, msg = validate_regrade_row(row, excel_row)
            if not ok:
                errors.append(msg)
        elif order_type == 'regrade - cancel':
            ok, msg = validate_cancel_regrade_row(row, excel_row)
            if not ok:
                errors.append(msg)
        # Future: add validate_new_row(...), validate_migration_row(...) here

    return errors


# =============================================================================
# ROW → dict BUILDERS
# =============================================================================

def build_regrade_row(row: pd.Series) -> dict:
    """
    Extract all fields needed for a Regrade order from the spreadsheet row.

    The regrade-specific columns (AE–AN) are used for portal fields.
    The shared contact/site columns (A–AD) may also be read where needed
    (e.g. for the confirmation email address in contactEmail).

    Returns a flat dict passed directly to SOADSLPortal.place_regrade().
    """
    # ── Regrade-specific portal fields ────────────────────────────────────────
    data = {
        # Identifying the existing service
        'serviceID':      _str(row['serviceID']),
        'channelPartner': _str(row['channelPartner']),
        'customerName':   _str(row['customerName']),
        'CLI':            format_phone_number(row['CLI']),

        # New product details
        'broadbandProduct': _str(row['broadbandProduct']),
        'newCareLevel':     _opt(row.get('newCareLevel')),
        'ipAddressOption':  _opt(row.get('ipAddressOption')),
        'technologyType':   _opt(row.get('technologyType')),

        # Scheduling
        'customerRequiredDate': convert_to_date_string(row['customerRequiredDate.1']),
        'installType':          _opt(row.get('installType.1')),
    }

    # ── Shared contact columns — used for notification / reference ─────────────
    data['account']          = _str(row.get('account'))
    data['contactName']      = _str(row.get('contactName'))
    data['contactTelephone'] = format_phone_number(row.get('contactTelephone'))
    data['contactEmail']     = _str(row.get('contactEmail'))
    data['customerReference']= _opt(row.get('customerReference'))

    return data


def build_cancel_regrade_row(row: pd.Series) -> dict:
    """
    Extract all fields needed for a Regrade Cancellation from the spreadsheet row.

    Only serviceID is required - the cancellation date is auto-generated as today's date
    and the cancellation reason is hardcoded to "Cancellation".

    The shared contact/account columns (A–AD) are included for reporting purposes.

    Returns a flat dict passed directly to SOADSLPortal.cancel_regrade().
    """
    # ── Regrade cancellation portal fields ────────────────────────────────────
    data = {
        # Identifying the existing service to cancel
        'serviceID': _str(row['serviceID']),
    }

    # ── Shared contact columns — used for notification / reference ─────────────
    data['account']          = _str(row.get('account'))
    data['contactName']      = _str(row.get('contactName'))
    data['contactTelephone'] = format_phone_number(row.get('contactTelephone'))
    data['contactEmail']     = _str(row.get('contactEmail'))
    data['customerReference']= _opt(row.get('customerReference'))

    # Also include regrade-specific columns if present (for reporting)
    data['channelPartner'] = _str(row.get('channelPartner'))
    data['customerName']   = _str(row.get('customerName'))
    data['CLI']            = format_phone_number(row.get('CLI'))

    return data


# =============================================================================
# RESULT WRITING
# =============================================================================

# Result/status columns that appear first in the output sheet
_RESULT_COLS = ['status', 'order_ref', 'message']


def write_results(results: list[dict], output_file: str):
    """Write the results list to an Excel output file.

    Result/status columns appear first, followed by all input data columns
    in their original sheet order, so failures are easy to diagnose.
    """
    df_out = pd.DataFrame(results)
    # Build column order: result cols first, then every input col not already listed
    ordered = [c for c in _RESULT_COLS if c in df_out.columns]
    ordered += [c for c in df_out.columns if c not in ordered]
    df_out = df_out[ordered]
    df_out.to_excel(output_file, index=False)
    print(f"\n  [OK] Results written to: {os.path.basename(output_file)}")


# =============================================================================
# MAIN FILE PROCESSING
# =============================================================================

def process_file(file_path: str, filename: str, portal: SOADSLPortal):
    """
    Process one input Excel file end-to-end.

    Args:
        file_path: Full path to the workbook.
        filename:  Filename without extension (used for output naming).
        portal:    An already-started SOADSLPortal instance.
    """
    print("=" * 70)
    print(f"  PROCESSING: {filename}")
    print("=" * 70)

    try:
        df = load_sheet(file_path)
    except Exception as exc:
        print(f"  [ERROR] Could not read '{filename}': {exc}")
        return

    print(f"  [OK] Loaded {len(df)} row(s)")

    # ── Validate all rows before placing any orders ────────────────────────────
    print("\n  Validating rows...")
    validation_errors = validate_all_rows(df)
    if validation_errors:
        print(f"\n  [ERROR] Validation failed - {len(validation_errors)} error(s):\n")
        for err in validation_errors:
            print(f"     - {err}")
        print("\n  Fix the errors above and re-run.")
        return

    print("  [OK] All rows valid\n")

    # ── Place orders ───────────────────────────────────────────────────────────
    timestamp   = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(OUTPUT_FOLDER, f"{filename}_results_{timestamp}.xlsx")
    results     = []

    for idx, row in df.iterrows():
        excel_row  = idx + 2
        order_type = _str(row.get('orderType')).lower()
        print(f"  Row {excel_row:>3} | {order_type.upper():<12}", end=' ', flush=True)

        try:
            if order_type == 'regrade':
                row_data = build_regrade_row(row)
                result   = portal.place_regrade(row_data)

            elif order_type == 'regrade - cancel':
                row_data = build_cancel_regrade_row(row)
                result   = portal.cancel_regrade(row_data)

            # ── Future order types ─────────────────────────────────────────────
            # elif order_type == 'new':
            #     row_data = build_new_row(row)
            #     result   = portal.place_new_order(row_data)
            # elif order_type == 'migration':
            #     row_data = build_migration_row(row)
            #     result   = portal.place_migration(row_data)
            else:
                result = {
                    'order_ref': '',
                    'status':    'SKIPPED',
                    'message':   f"Unknown orderType '{order_type}' - skipped",
                }

            print(f"> {result['status']:<8}  {result.get('order_ref', '')}  {result.get('message', '')}")

        except SOADSLPortalError as exc:
            result = {
                'order_ref': '',
                'status':    'ERROR',
                'message':   str(exc.message),
            }
            print(f"> ERROR  {exc.message}")

        except Exception as exc:
            result = {
                'order_ref': '',
                'status':    'ERROR',
                'message':   str(exc),
            }
            print(f"> ERROR  {exc}")

        # Tag result with status/source columns

        # Merge all input row data so the report is self-contained.
        # Rename .1-suffixed duplicate columns back to readable names.
        row_data = {
            str(k).replace('customerRequiredDate.1', 'customerRequiredDate_regrade')
                  .replace('installType.1', 'installType_regrade'):
            (convert_to_date_string(v) if isinstance(v, (pd.Timestamp, datetime)) else _str(v))
            for k, v in row.items()
        }
        # Input columns must not overwrite result columns
        for k, v in row_data.items():
            if k not in result:
                result[k] = v

        results.append(result)

        time.sleep(ORDER_DELAY)

    # ── Write results & archive input ─────────────────────────────────────────
    write_results(results, output_file)

    # ── Email the report ──────────────────────────────────────────────────────
    if REPORT_RECIPIENTS:
        total      = len(results)
        successful = sum(1 for r in results if r.get('status') == 'SUBMITTED')
        failed     = sum(1 for r in results if r.get('status') in ('FAILED', 'SKIPPED'))
        errors     = sum(1 for r in results if r.get('status') == 'ERROR')
        send_order_report_email(
            report_file_path=output_file,
            recipient_email=REPORT_RECIPIENTS,
            input_filename=filename,
            total_orders=total,
            successful=successful,
            failed=failed,
            errors=errors,
            report_type='Regrade Order',
            subject_prefix='SOADSL Regrade Order Report',
        )

    dest = os.path.join(PROCESSED_FOLDER, os.path.basename(file_path))
    try:
        shutil.move(file_path, dest)
        print(f"  [OK] Input file archived to processed/")
    except Exception as exc:
        print(f"  [WARNING] Could not archive input file: {exc}")


# =============================================================================
# ENTRY POINT
# =============================================================================

def main():
    input_files = get_input_files()

    if not input_files:
        print(f"\n  No input files found in:\n  {INPUT_FOLDER}\n")
        return

    print(f"\n  Found {len(input_files)} file(s) to process\n")

    with SOADSLPortal() as portal:
        for file_path, filename in input_files:
            process_file(file_path, filename, portal)

    print("\n  [OK] All files processed.")


if __name__ == '__main__':
    main()
