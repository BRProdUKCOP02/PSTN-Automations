"""
SOADSL Portal Automation
Handles interactions on BMT for SOADSL ordering 

Supported order types:
- New                : Place a new SOADSL Order
- Regrades           : Change product on an existing circuit to SOADSL
- Migration          : Migrate a non-Gamma Broadband ciruit to SOADSL
- Migration_Gamma    : Migrate a Gamma Broadband circuit to SOADSL
"""
from playwright.sync_api import sync_playwright, Page, Browser, BrowserContext, Dialog
from dotenv import load_dotenv
from datetime import datetime
import os
import time

# Load .env credentials for BMT HTTP Auth (browser popup)
# .env lives in the same folder as this script (PSTN Migration/)
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env'))

# BMT Portal URLs
BMT_UAT_URL = "https://dev-datasvc-03/UAT/broadband_home.php"
BMT_LIVE_URL = "https://bmt/broadband_home.php"

# Set Environment
BMT_USE_PRODUCTION = False  # Set to True for production, False for UAT
BMT_URL = BMT_LIVE_URL if BMT_USE_PRODUCTION else BMT_UAT_URL

# Portal Exception Handling - raised when a portal action failes (e.g. element not found, timeout, etc.)
class SOADSLPortalError(Exception):
    def __init__(self, message: str, order_type: str = '', row_num: int = 0):
        self.message = message
        self.order_type = order_type
        self.row_num = row_num
        super().__init__(message)


class SOADSLPortal:
    """
    Controls the BMT Portal using Playwright library commands.

    Usage:
        portal = SOADSLPortal()
        portal.start()
        result = portal.place_new_order(row)
        portal.close()

    Or use as a context manager:
        with SOADSLPortal() as portal:
            result = portal.place_new_order(row)
    """

    def __init__(self):
        # Prefer dedicated BMT credentials; fall back to GAMMA_ credentials.
        # .env keys (in order of precedence):
        #   BMT_USERNAME / BMT_PASSWORD   ← use these for BMT-specific login
        #   GAMMA_USERNAME / GAMMA_PASSWORD ← fallback (Gamma API creds)
        self.username = os.getenv("BMT_USERNAME") or os.getenv("GAMMA_USERNAME")
        self.password = os.getenv("BMT_PASSWORD") or os.getenv("GAMMA_PASSWORD")

        if not self.username or not self.password:
            raise SOADSLPortalError(
                "BMT credentials not found. Add BMT_USERNAME and BMT_PASSWORD "
                "(or GAMMA_USERNAME / GAMMA_PASSWORD) to the .env file in the "
                "'PSTN Migration' folder."
            )
        # Playwright browser and context will be initialized in start() method
        self._playwright = None
        self._browser: Browser = None
        self._context: BrowserContext = None
        self._page: Page = None

    # ── Lifecycle ─────────────────────────────────────────────────────────────

    def start(self, headless: bool = False):
        """
        Launch the browser and authenticate with BMT.

        headless=False lets you see the browser while testing.
        Set headless=True for unattended / scheduled runs.
        """
        self._playwright = sync_playwright().start()

        # ignore HTTPS errors - BMT UAT uses a self-signed certificate
        self._browser = self._playwright.chromium.launch(headless=headless)

        # http_credentials handles the HTTP Basic Auth popup automatically.
        # Playwright intercepts the 401 challenge and responds before the
        # popup is ever shown on screen.
        self._context = self._browser.new_context(
            ignore_https_errors=True,
            http_credentials={
                "username": self.username,
                "password": self.password
            }
        )
        self._page = self._context.new_page()

        # Verify credentials with a lightweight probe against the BMT root —
        # this catches a 401 early without loading any full portal page.
        masked_pw = ('*' * (len(self.password) - 2) + self.password[-2:]) if self.password else ''
        bmt_root = BMT_URL.split('/')[0] + '//' + BMT_URL.split('/')[2] + '/'
        print(f"  Auth probe : {bmt_root}")
        print(f"  Username   : {self.username}")
        print(f"  Password   : {masked_pw}")

        response = self._page.goto(bmt_root, wait_until="commit", timeout=15000)
        status = response.status if response else 0
        if status == 401:
            raise SOADSLPortalError(
                f"BMT returned 401 Unauthorized.\n"
                f"  URL      : {bmt_root}\n"
                f"  Username : {self.username}\n"
                f"  Check / update BMT_USERNAME and BMT_PASSWORD in:\n"
                f"  {os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')}"
            )
        if "login" in self._page.url.lower() or (status not in (0, 200, 302) and status >= 400):
            raise SOADSLPortalError(
                f"BMT login failed (HTTP {status}). Landed on: {self._page.url} — "
                f"check credentials and that {BMT_URL} is reachable."
            )
        print(f"  [OK] BMT authenticated  (HTTP {status})")

    def close(self):
        """Close the browser and end the Playwright session cleanly."""
        try:
            if self._browser:
                self._browser.close()
            if self._playwright:
                self._playwright.stop()
        except Exception:
            pass  # always close silently

    # Context manager support — allows `with SOADSLPortal() as portal:`
    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False  # do not suppress exceptions

    # ── Navigation helpers ────────────────────────────────────────────────────

    def _go_home(self):
        """Navigate back to the BMT Home Page between orders."""
        self._page.goto(BMT_URL, wait_until="load", timeout=30000)

    def _wait_and_fill(self, selector: str, value: str, timeout: int = 90000):
        """Wait for a text field to appear, then fill it."""
        self._page.wait_for_selector(selector, timeout=timeout)
        self._page.fill(selector, str(value))

    def _wait_and_select(self, selector: str, value: str, timeout: int = 90000):
        """Wait for a <select> dropdown, then choose a value."""
        self._page.wait_for_selector(selector, timeout=timeout)
        self._page.select_option(selector, value=str(value))

    def _wait_and_click(self, selector: str, timeout: int = 90000):
        """Wait for a button or link, then click it."""
        self._page.wait_for_selector(selector, timeout=timeout)
        self._page.click(selector)

    def _set_date_field(self, name: str, value_ddmmyyyy: str):
        """
        Set a datepicker <input> by injecting the value via JavaScript and
        firing a 'change' event so any JS listeners (e.g. validation) react.

        The BMT datepickers use DD/MM/YYYY format and start disabled/readonly;
        after the Edit button is clicked they become enabled but remain
        attached to jQuery UI datepicker, so direct fill() is unreliable.
        """
        self._page.evaluate(
            """
            (args) => {
                const el = document.querySelector(`input[name="${args.name}"]`);
                if (!el) throw new Error('Date field not found: ' + args.name);
                el.removeAttribute('disabled');
                el.removeAttribute('readonly');
                el.value = args.value;
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('input',  { bubbles: true }));
            }
            """,
            {"name": name, "value": value_ddmmyyyy}
        )

    @staticmethod
    def _to_ddmmyyyy(date_str: str) -> str:
        """
        Convert YYYY-MM-DD → DD/MM/YYYY as expected by BMT datepickers.
        Returns the value unchanged if it cannot be parsed.
        """
        try:
            return datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%Y')
        except (ValueError, TypeError):
            return date_str  # pass through and let the portal reject it

    # ── Order placement methods ───────────────────────────────────────────────

    def place_new_order(self, row: dict) -> dict:
        """
        Place a new SOADSL order.

        Args:
            row: dict of values from the spreadsheet row
        Returns:
            dict with keys: order_ref, status, message
        """
        raise NotImplementedError(
            "place_new_order() will be built once we inspect the BMT form together."
        )

    def place_regrade(self, row: dict) -> dict:
        """
        Regrade an existing SOADSL circuit to a different product via BMT portal.

        Portal flow
        -----------
        1. Navigate to service_details.php?serviceid={serviceID}
        2. Click the Edit button to unlock the regrade form fields.
        3. Fill in:
               - New Broadband Product  (REG_BroadbandProduct)
               - New Care Level         (REG_CareLevel)          — optional
               - IP Address Option      (REG_CGNATId)            — optional
               - Technology Type        (REG_TechnologyType)     — optional
               - Customer Required Date (REG_CustomerRequiredDate, DD/MM/YYYY)
        4. Click Submit Regrade — a browser confirm() dialog appears.
        5. Accept the dialog automatically.
        6. Wait for the page to settle and return the result.

        Args:
            row: dict produced by build_regrade_row() with keys:
                 serviceID, broadbandProduct, newCareLevel, ipAddressOption,
                 technologyType, customerRequiredDate, ...
        Returns:
            dict with keys: order_ref, status, message
        Raises:
            SOADSLPortalError on any portal interaction failure.
        """
        service_id = row['serviceID']

        # Build the service details URL by replacing the home-page path segment
        # e.g. https://dev-datasvc-03/UAT/broadband_home.php
        #   -> https://dev-datasvc-03/UAT/service_details.php?serviceid=568439
        service_url = BMT_URL.replace(
            'broadband_home.php',
            f'service_details.php?serviceid={service_id}'
        )

        # ── Persistent dialog handler — active for the entire regrade flow ──────
        # BMT can fire browser dialogs at several points:
        #   • On page load  : "Alert - This service is being regraded"
        #                     (appears if the circuit is already under regrade)
        #   • After submit 1: confirm "Do you want to mark this service under regrade?"
        #   • After reload  : alert  "Alert - This service is being regraded"
        #   • After submit 2: same confirm/alert pair for the install type step
        # Registering ONE persistent handler here (before goto) ensures ALL of
        # them are accepted automatically and the page never stays blocked.
        def _accept_all(dialog):
            try:
                dialog.accept()
            except Exception:
                pass  # ignore if already dismissed

        self._page.on('dialog', _accept_all)

        try:
            # ── 1. Open the service details page ─────────────────────────────
            # Use 'load' not 'networkidle' — the service details page has
            # persistent connections that prevent networkidle from ever firing.
            # UAT can be slow — allow up to 90 seconds for the page to load.
            self._page.goto(service_url, wait_until='load', timeout=90000)
            # Brief pause — let any on-load dialog fire and be dismissed before
            # we start querying the DOM for form elements.
            time.sleep(1)

            # ── 2. Click Edit to enable the regrade form ──────────────────────
            self._wait_and_click('input[name="REG_InitiateEdit"]')

            # Wait until the Submit Regrade button loses its 'hidden' class,
            # which confirms the form is now in edit mode.
            self._page.wait_for_function(
                """() => {
                    const btn = document.querySelector('input[name="REG_InitiateRegrade"]');
                    return btn && !btn.classList.contains('hidden');
                }""",
                timeout=90000
            )

            # ── 3. Fill in the regrade fields ─────────────────────────────────

            # Broadband Product — select by visible option text.
            # Option values are JSON blobs; select_option(label=) matches text.
            self._page.select_option(
                'select[name="REG_BroadbandProduct"]',
                label=row['broadbandProduct']
            )
            # Allow any onchange JS (product refresh, care level check, etc.) to settle
            time.sleep(0.5)

            # Care Level (optional)
            if row.get('newCareLevel'):
                self._wait_and_select('select[name="REG_CareLevel"]', row['newCareLevel'])

            # IP Address Option (optional)
            # Spreadsheet value is the label text; map to the option value (1 or 2)
            if row.get('ipAddressOption'):
                ip_label = str(row['ipAddressOption']).strip().lower()
                ip_value_map = {
                    'static public ip address': '1',
                    'static': '1',
                    'cgnat broadband': '2',
                    'cgnat': '2',
                }
                ip_val = ip_value_map.get(ip_label)
                if ip_val:
                    self._wait_and_select('select[name="REG_CGNATId"]', ip_val)
                else:
                    # Fall back to selecting by label text directly
                    self._page.select_option(
                        'select[name="REG_CGNATId"]',
                        label=row['ipAddressOption']
                    )
                time.sleep(0.3)  # let ToggleRoutedIPsDropdown() run

            # Technology Type (optional)
            # Option values match the visible text (e.g. "ADSL2+ (WBC)")
            if row.get('technologyType'):
                self._wait_and_select('select[name="REG_TechnologyType"]', row['technologyType'])

            # Customer Required Date — BMT expects DD/MM/YYYY
            # The field starts disabled; _set_date_field() removes that via JS.
            if row.get('customerRequiredDate'):
                date_ddmmyyyy = self._to_ddmmyyyy(row['customerRequiredDate'])
                self._set_date_field('REG_CustomerRequiredDate', date_ddmmyyyy)

            # ── 4–7. Submit — dialogs handled by top-level _accept_all handler ──
            self._wait_and_click('input[name="REG_InitiateRegrade"]')
            # Wait for the page reload that follows the submit confirm being accepted.
            self._page.wait_for_load_state('load', timeout=90000)
            # Brief pause so any post-reload alert can fire and be dismissed
            time.sleep(2)

            # ── 8. Set Install Type (if provided) ─────────────────────────────
            install_type = row.get('installType')
            if install_type:
                # After submission the Edit button name changes from
                # REG_InitiateEdit  (pre-regrade)  →  REG_InProgressEdit (post-regrade).
                # Use a CSS selector that matches either name to be safe.
                self._wait_and_click('input[name="REG_InProgressEdit"], input[name="REG_InitiateEdit"]')

                # Wait for the Install Type dropdown to become editable — this
                # confirms the form is in edit mode, without relying on the
                # submit button name which changes between portal states.
                self._page.wait_for_selector(
                    'select[name="REG_PCPInstallTypeId"]',
                    state='visible',
                    timeout=90000
                )
                time.sleep(0.5)  # let any JS initialisation settle

                # REG_PCPInstallTypeId uses numeric option values.
                # Map the spreadsheet label to the correct value.
                install_type_map = {
                    'self install':              '2',
                    'managed install':           '1',
                    'premium managed install':   '3',
                }
                it_val = install_type_map.get(str(install_type).strip().lower())

                # The select has comboboxReadOnlyOnFocus / comboboxReadOnlyOnChange
                # handlers that intercept and reset changes.  Remove them first,
                # then set the value directly, bypassing those guards.
                self._page.evaluate(
                    """
                    (args) => {
                        const el = document.querySelector('select[name="REG_PCPInstallTypeId"]');
                        if (!el) throw new Error('REG_PCPInstallTypeId not found');
                        el.removeAttribute('onchange');
                        el.removeAttribute('onfocus');
                        el.removeAttribute('disabled');
                        el.removeAttribute('readonly');
                        if (args.value) {
                            el.value = args.value;
                        } else {
                            const lower = args.label.toLowerCase();
                            for (const opt of el.options) {
                                if (opt.text.toLowerCase() === lower) {
                                    el.value = opt.value;
                                    break;
                                }
                            }
                        }
                    }
                    """,
                    {"value": it_val or "", "label": str(install_type)}
                )

                # Find the submit button — name varies by portal state:
                #   REG_UpdateStatus      (install type / status update form)
                #   REG_InProgressRegrade (post-regrade edit form)
                #   REG_InitiateRegrade   (fresh regrade form)
                # Use JS to find whichever is present and not hidden.
                submit_name = self._page.evaluate(
                    """() => {
                        for (const name of ['REG_UpdateStatus', 'REG_InProgressRegrade', 'REG_InitiateRegrade']) {
                            const btn = document.querySelector('input[name="' + name + '"]');
                            if (btn && !btn.classList.contains('hidden') &&
                                btn.offsetParent !== null) {
                                return name;
                            }
                        }
                        return null;
                    }"""
                )
                if not submit_name:
                    raise SOADSLPortalError(
                        f'Could not find the regrade submit button for service {service_id}',
                        order_type='Regrade'
                    )
                self._page.click(f'input[name="{submit_name}"]')
                self._page.wait_for_load_state('load', timeout=90000)
                time.sleep(2)

            return {
                'order_ref': '',
                'status':    'SUBMITTED',
                'message':   (
                    f'Regrade submitted for service {service_id}'
                    + (f' | Install Type: {install_type}' if install_type else '')
                ),
            }

        except SOADSLPortalError:
            raise  # re-raise portal errors unchanged
        except Exception as exc:
            raise SOADSLPortalError(
                message=f'Regrade failed for service {service_id}: {exc}',
                order_type='Regrade',
            ) from exc
        finally:
            # Always remove the persistent dialog handler when leaving this method
            try:
                self._page.remove_listener('dialog', _accept_all)
            except Exception:
                pass

    def cancel_regrade(self, row: dict) -> dict:
        """
        Cancel an existing regrade order via BMT portal.

        Portal flow
        -----------
        1. Navigate to service_details.php?serviceid={serviceID}
        2. Click the "Cancel Regrade" button (REG_CancelRegradeExpand)
        3. Wait for cancellation form fields to appear
        4. Fill in:
               - Cancellation Date (auto-populated with today's date in DD/MM/YYYY)
               - Cancellation Reason (REG_CancellationReason, value="Cancellation")
        5. Click "Submit Cancel" button (REG_CancelRegrade)
        6. Accept the browser confirmation dialog automatically

        Args:
            row: dict with keys:
                 serviceID (required) - existing service with active regrade to cancel
                 account, contactName, contactEmail, etc. (optional - for reporting)
        Returns:
            dict with keys: order_ref, status, message
        Raises:
            SOADSLPortalError on any portal interaction failure.
        """
        service_id = row['serviceID']

        # Build the service details URL
        service_url = BMT_URL.replace(
            'broadband_home.php',
            f'service_details.php?serviceid={service_id}'
        )

        # ── Persistent dialog handler for all popups (on-load alerts + confirmation) ──
        # BMT fires an alert "This service is being regraded" when loading
        # a service page that's already under regrade. This handler auto-accepts
        # ALL dialogs (alerts and confirms) throughout the cancel flow.
        def _accept_all(dialog):
            try:
                dialog.accept()
            except Exception:
                pass

        self._page.on('dialog', _accept_all)

        try:
            # ── 1. Open the service details page ─────────────────────────────
            # Use 'load' wait state - the page has persistent connections that
            # prevent 'networkidle' from firing. UAT can be slow (90s timeout).
            print(f"      Navigating to service {service_id}...")
            self._page.goto(service_url, wait_until='load', timeout=90000)
            print("      ✓ Page loaded")
            
            # Critical: wait for page to be fully interactive after any on-load
            # dialogs are dismissed. The "Cancel Regrade" button existence
            # confirms the page is ready and any blocking dialogs are gone.
            time.sleep(2)  # Extended pause for dialog dismissal and page stabilization
            
            # Verify page is ready by waiting for a key element to be present
            # (either the Cancel Regrade button or the edit button)
            print("      Waiting for page elements...")
            try:
                self._page.wait_for_selector(
                    'input[name="REG_CancelRegradeExpand"], input[name="REG_InitiateEdit"], input[name="REG_InProgressEdit"]',
                    timeout=15000
                )
                print("      ✓ Page ready")
            except Exception:
                # If we can't find any expected buttons, the page may still be blocked
                print("      ⚠ Key elements not found, waiting longer...")
                time.sleep(2)  # Give it more time

            # ── 2. Click "Cancel Regrade" button ──────────────────────────────
            print(f"      Clicking Cancel Regrade button...")
            self._wait_and_click('input[name="REG_CancelRegradeExpand"]')
            print("      ✓ Cancel Regrade button clicked")

            # ── 3. Wait for cancellation form to appear ───────────────────────
            # The cancellation reason dropdown appears when Cancel Regrade is clicked
            print("      Waiting for cancellation form to appear...")
            try:
                self._page.wait_for_selector(
                    'select[name="REG_CancellationReason"]',
                    state='visible',
                    timeout=10000
                )
                print("      ✓ Cancellation form visible")
            except Exception as e:
                print(f"      ✗ Cancellation form not found: {e}")
                raise SOADSLPortalError(message=f"Cancellation form did not appear after clicking Cancel Regrade button")
            
            time.sleep(0.5)  # Let form fields fully initialize

            # ── 4. Fill in cancellation date (today) ──────────────────────────
            # The date field should accept today's date in DD/MM/YYYY format
            today_ddmmyyyy = datetime.now().strftime('%d/%m/%Y')
            today_day = datetime.now().day
            print(f"      Filling cancellation date: {today_ddmmyyyy}")
            
            # The field name is REG_RegradeCancellationDate and it's readonly with a datepicker
            try:
                date_input = self._page.query_selector('input[name="REG_RegradeCancellationDate"]')
                if date_input:
                    print(f"      ✓ Found date field: REG_RegradeCancellationDate")
                    
                    # Remove readonly attribute so we can fill it
                    self._page.evaluate('''
                        () => {
                            const input = document.querySelector('input[name="REG_RegradeCancellationDate"]');
                            if (input) {
                                input.removeAttribute('readonly');
                                input.removeAttribute('disabled');
                            }
                        }
                    ''')
                    
                    # Fill the value directly
                    date_input.fill(today_ddmmyyyy)
                    
                    # Trigger events
                    date_input.dispatch_event('input')
                    date_input.dispatch_event('change')
                    date_input.dispatch_event('blur')
                    
                    print(f"      ✓ Date filled: {today_ddmmyyyy}")
                else:
                    print(f"      ✗ Date field not found")
                    # Try clicking today's date in the open calendar as fallback
                    print(f"      Trying to click day {today_day} in calendar...")
                    try:
                        # Click today's date in the calendar
                        self._page.click(f'a.ui-state-default:has-text("{today_day}")', timeout=3000)
                        print(f"      ✓ Clicked day {today_day} in calendar")
                    except Exception as e:
                        print(f"      ✗ Calendar click failed: {e}")
            except Exception as e:
                print(f"      ✗ Date fill failed: {e}")

            # ── 5. Select "Cancellation" reason ───────────────────────────────
            print("      Selecting cancellation reason: 'Cancellation'")
            try:
                # Wait for the dropdown to be visible and enabled
                reason_dropdown = self._page.wait_for_selector(
                    'select[name="REG_CancellationReason"]',
                    state='visible',
                    timeout=5000
                )
                # Select by value or text - try both
                try:
                    reason_dropdown.select_option(label='Cancellation')
                    print("      ✓ Reason selected by label")
                except Exception:
                    try:
                        reason_dropdown.select_option(value='Cancellation')
                        print("      ✓ Reason selected by value")
                    except Exception as e:
                        # Try selecting by index (assuming Cancellation is first real option after placeholder)
                        reason_dropdown.select_option(index=1)
                        print("      ✓ Reason selected by index")
                
                # Trigger change event
                reason_dropdown.dispatch_event('change')
                time.sleep(0.3)
            except Exception as e:
                print(f"      ✗ Failed to select cancellation reason: {e}")
                raise SOADSLPortalError(message=f"Could not select cancellation reason: {e}")

            # ── 6. Click "Submit Cancel" button ───────────────────────────────
            print("      Clicking Submit Cancel button...")
            try:
                submit_button = self._page.wait_for_selector(
                    'input[id="REG_CancelRegrade"]',
                    state='visible',
                    timeout=5000
                )
                submit_button.click()
                print("      ✓ Submit clicked")
            except Exception as e:
                print(f"      ✗ Failed to click Submit Cancel: {e}")
                # Try alternative selector
                try:
                    self._page.click('input[name="REG_CancelRegrade"]')
                    print("      ✓ Submit clicked (alternative selector)")
                except Exception as e2:
                    raise SOADSLPortalError(message=f"Could not click Submit Cancel button: {e2}")
            
            # Wait for page reload after submission
            self._page.wait_for_load_state('load', timeout=90000)
            time.sleep(2)  # Let any post-submit dialog fire and be dismissed
            print("      ✓ Page reloaded after submission")

            return {
                'order_ref': '',
                'status':    'SUBMITTED',
                'message':   f'Regrade cancellation submitted for service {service_id}',
            }

        except SOADSLPortalError:
            raise  # re-raise portal errors unchanged
        except Exception as exc:
            raise SOADSLPortalError(
                message=f'Cancel regrade failed for service {service_id}: {exc}',
                order_type='Regrade - Cancel',
            ) from exc
        finally:
            # Always remove the persistent dialog handler
            try:
                self._page.remove_listener('dialog', _accept_all)
            except Exception:
                pass

    def place_migration(self, row: dict) -> dict:
        """
        Migrate a non-Gamma broadband circuit to Gamma SOADSL.

        Args:
            row: dict of values from the spreadsheet row
        Returns:
            dict with keys: order_ref, status, message
        """

        raise NotImplementedError("place_migration() — to be implemented.")

#    def place_migration_gamma(self, row: dict) -> dict:
        """
        Migrate an existing Gamma circuit to SOADSL.

        Args:
            row: dict of values from the spreadsheet row
        Returns:
            dict with keys: order_ref, status, message
        """
        raise NotImplementedError("place_migration_gamma() — to be implemented.")


# =============================================================================
# STANDALONE CONNECTION TEST
# Run this file directly to verify BMT portal connectivity:
#   python soadsl_portal.py
#   python soadsl_portal.py 568439        ← also open a service page
# =============================================================================

if __name__ == '__main__':
    import sys

    service_id = sys.argv[1] if len(sys.argv) > 1 else None

    print(f"\n  BMT URL  : {BMT_URL}")
    print(f"  Env      : {'PRODUCTION' if BMT_USE_PRODUCTION else 'UAT'}")
    print(f"  Username : {os.getenv('BMT_USERNAME') or os.getenv('GAMMA_USERNAME') or '(not set)'}\n")

    try:
        portal = SOADSLPortal()
        portal.start(headless=False)      # headless=False so you can see the browser
        print("  [OK] Portal connection OK")

        if service_id:
            service_url = BMT_URL.replace(
                'broadband_home.php',
                f'service_details.php?serviceid={service_id}'
            )
        else:
            service_url = BMT_URL.replace('broadband_home.php', '')
            print(f"\n  Tip: pass a service ID to also open its page, e.g.:")
            print(f"       python soadsl_portal.py 568439\n")

        print(f"\n  Opening: {service_url}")
        portal._page.goto(service_url, wait_until='load', timeout=30000)
        print(f"  [OK] Page loaded  (title: {portal._page.title()!r})")

        input("\n  Browser is open — press Enter to close it.\n")
        portal.close()

    except SOADSLPortalError as exc:
        print(f"\n  [ERROR] Portal error: {exc.message}")
    except Exception as exc:
        print(f"\n  [ERROR] Unexpected error: {exc}")