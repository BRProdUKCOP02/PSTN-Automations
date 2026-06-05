# SoGEA Bulk Ordering - Design Document

## Solution Overview
Automated broadband order placement system that reads SoGEA (Single Order Generic Ethernet Access) order data from Excel templates and places bulk orders via the Gamma Broadband Ordering API. Supports NEW orders (fresh installations) and REGRADE orders (service upgrades/migrations). Handles installation coordination, router provisioning, delivery logistics, and number porting.

## How It Works

### Workflow
1. **File Discovery**: Scans input folder (`\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\...`) for Excel files
2. **File Locking**: Immediately moves file to processing folder to prevent duplicate processing if automation triggers multiple times
3. **Authentication**: Obtains JWT bearer token using Gamma API credentials from config.py
4. **Order Type Detection**: Determines NEW vs. REGRADE from `orderType` column (defaults to NEW)
5. **Validation**: Validates required fields based on order type
6. **Order Placement**: 
   - **NEW**: Places complete order with site, installation, equipment, reseller contact
   - **REGRADE**: Upgrades existing order - only requires product, date, router details
7. **Order Tracking**: Captures Gamma order ID for status checking
8. **Copy for Status Check**: Copies order IDs to order_check_input folder for 24h+ status verification
9. **Report Generation**: Outputs Excel report with order IDs, status, errors
10. **Email Notification**: Sends report to psmanaged.delivery@gamma.co.uk and Dave.Siddall@gamma.co.uk
11. **File Archival**: Moves completed file to processed folder with timestamp
12. **Delay Management**: 2-second delay between orders to avoid API rate limiting

### Order Types

#### NEW Orders
- **Purpose**: Place brand new broadband installation
- **Requires**: Full site details, NAD key, installation type, customer required date, care level, reseller contact
- **Optional**: Router, delivery address, number porting, routed IP, ECC/TRC bands

#### REGRADE Orders (Currently Disabled - Server-Side Issue)
- **Purpose**: Upgrade/downgrade existing broadband service
- **Requires**: existingOrderId, broadbandProduct, customerRequiredDate, routerRequired, router details (if needed)
- **Reuses**: Account number, care level, site details, reseller contact from existing order
- **Note**: Regrade functionality temporarily disabled due to Gamma API server-side bug

### Error Handling
- **Validation Errors**: Skips row and logs missing/invalid fields in report
- **API Errors**: Captures HTTP status codes and error messages, continues processing
- **Authentication Failures**: Stops processing and reports authentication issue
- **Duplicate Prevention**: File locking prevents concurrent processing

## API Integration

### Endpoints
- **Auth**: `https://api-test.gamma.co.uk/auth/token` (Test) / `https://api.gamma.co.uk/auth/token` (Production)
- **Orders**: `https://api-test.gamma.co.uk/broadband/v2/orders` - POST to place NEW orders
- **Regrade**: `https://api-test.gamma.co.uk/broadband/v2/orders/{orderId}/regrade` - POST to upgrade existing order

### Authentication
OAuth2 Password Grant - JWT Bearer Token obtained via POST:
```
Content-Type: application/x-www-form-urlencoded
grant_type=password&username=<api_user>&password=<api_password>
```

### Request Payload Structure

#### NEW Order Payload
```json
{
  "accountNumber": 12345,
  "broadbandProduct": "SoGEA_FTTC_40_10",
  "careLevel": "Standard",
  "resellerEmailNotifications": true,
  "ipAddressOption": "Single",
  "voiceProduct": "None" | "SIP2",
  "routedIpOption": "/29" (optional),
  "installation": {
    "lineType": "Existing" | "New",
    "cli": "01234567890" (if Existing),
    "customerRequiredDate": "2026-03-01",
    "type": "engineer" | "self" (optional),
    "customerReference": "string" (optional),
    "eccBand": "string" (optional),
    "trcBand": "string" (optional),
    "site": {
      "companyName": "string",
      "nadKey": "A00010584545",
      "address": {
        "building": "string",
        "subPremises": "string" (optional),
        "street": "string",
        "town": "string",
        "county": "string",
        "postcode": "string"
      },
      "contact": {
        "name": "string",
        "emailAddress": "string",
        "telephoneNumber": "string"
      }
    }
  },
  "resellerContact": {
    "name": "string",
    "emailAddress": "string",
    "telephoneNumber": "string"
  },
  "equipment": {
    "routerRequired": true,
    "router": "Technicolor TG582n v3",
    "deliveryAddress": {
      "building": "string",
      "street": "string",
      "town": "string",
      "county": "string",
      "postcode": "string"
    },
    "deliveryContact": {
      "name": "string",
      "emailAddress": "string",
      "telephoneNumber": "string"
    }
  },
  "numberPort": {
    "voipReference": "string" (optional)
  }
}
```

#### REGRADE Order Payload
```json
{
  "broadbandProduct": "SoGEA_FTTC_80_20",
  "customerRequiredDate": "2026-03-15",
  "routerRequired": false,
  "router": "string" (if routerRequired true),
  "voiceProduct": "string" (optional),
  "routedIpOption": "string" (optional),
  "careLevel": "string" (optional),
  "installType": "engineer" (optional - required when migrating to FTTC),
  "companyName": "string" (optional),
  "deliveryAddress": { /* same structure */ } (optional),
  "deliveryContact": { /* same structure */ } (optional)
}
```

## Key Variables & Configuration

### Script Configuration (in script)
| Variable | Value | Description |
|----------|-------|-------------|
| `ORDER_DELAY` | 2s | Delay between orders to avoid rate limiting |
| `INPUT_FOLDER` | OneDrive network path | Source folder for Excel files |
| `OUTPUT_FOLDER` | ./output | Results report destination |
| `CHECK_INPUT_FOLDER` | ./order_check_input | Order IDs copied here for 24h+ status check |
| `PROCESSED_FOLDER` | OneDrive/processed | Archive for completed files |

### External Configuration (config.py)
- `ENVIRONMENT`: "test" or "production" - determines API endpoint
- Gamma API credentials: username, password

## Mandatory Inputs (Excel Columns)

### Required for NEW Orders
| Column | Type | Description | Validation |
|--------|------|-------------|------------|
| `accountNumber` | Integer | Gamma account number | Numeric |
| `broadbandProduct` | String | Product code (e.g., "SoGEA_FTTC_40_10") | Non-empty |
| `careLevel` | String | "Standard", "Priority", "Enhanced" | Non-empty |
| `ipAddressOption` | String | "Single", "Block5", "Block13" | Non-empty |
| `lineType` | String | "Existing" or "New" | Must be "Existing" or "New" |
| `cli` | String | Existing line CLI (if lineType = Existing) | 11-digit UK number |
| `customerRequiredDate` | Date | Installation date (YYYY-MM-DD) | Valid date format |
| `site_companyName` | String | Site company name | Non-empty |
| `site_nadKey` | String | NAD key (e.g., "A00010584545") | Non-empty |
| `site_building` | String | Site building number/name | Non-empty |
| `site_street` | String | Site street name | Non-empty |
| `site_town` | String | Site town/city | Non-empty |
| `site_postcode` | String | Site postcode | Non-empty |
| `site_contact_name` | String | Site contact name | Non-empty |
| `site_contact_email` | String | Site contact email | Valid email |
| `site_contact_phone` | String | Site contact phone | Non-empty (leading zero preserved) |
| `reseller_contact_name` | String | Reseller contact name | Non-empty |
| `reseller_contact_email` | String | Reseller contact email | Valid email |
| `reseller_contact_phone` | String | Reseller contact phone | Non-empty (leading zero preserved) |

### Required for REGRADE Orders
| Column | Type | Description | Validation |
|--------|------|-------------|------------|
| `orderType` | String | Must be "REGRADE" | "REGRADE" |
| `existingOrderId` | Integer | Order ID to regrade | Numeric |
| `broadbandProduct` | String | New product code | Non-empty |
| `customerRequiredDate` | Date | Regrade date (YYYY-MM-DD) | Valid date format |
| `routerRequired` | Boolean | TRUE/FALSE | Boolean |

### Optional Columns
- `voiceProduct`: "None", "SIP2" (voice over IP)
- `routedIpOption`: "/29", "/30" (skip if not needed or use "none")
- `resellerEmailNotifications`: TRUE/FALSE (defaults to TRUE)
- `installation_type`: "engineer", "self"
- `customerReference`: Customer reference string
- `eccBand`, `trcBand`: Installation time bands
- `site_subPremises`, `site_county`: Additional address fields
- `routerRequired`: TRUE/FALSE
- `router`: Router model (required if routerRequired = TRUE)
- `router_delivery_*`: Delivery address fields (building, street, town, county, postcode)
- `router_contact_*`: Delivery contact fields (name, email, phone)
- `voipReference`: Number porting reference

## Output Data

### Report Excel Structure
**Sheet 1 - Order Results**: Row per order with:
- `row`: Excel row number
- `timestamp`: Processing timestamp
- Input echo fields (accountNumber, broadbandProduct, careLevel, site details, etc.)
- `success`: Boolean order placement status
- `order_id`: Gamma order ID (for tracking)
- `order_status`: Initial order status from API
- `error`: Error message if failed

**Sheet 2 - Summary**: Processing statistics (total, successful, failed, environment, timestamp)

### Order ID File (for Status Checking)
Separate Excel file created in `order_check_input` folder containing:
- Column: `Order_ID` with all successfully placed order IDs
- Filename: `{input_filename}_order_ids_{timestamp}.xlsx`
- **Purpose**: Used by bulk_check_orders.py to verify order status 24+ hours after placement

## Technical Notes
- **Date Handling**: Automatically converts Excel dates to YYYY-MM-DD format, strips time component
- **Phone Number Formatting**: Adds leading zero to UK numbers if Excel stripped it (07xxx, 01/02/03/08)
- **NAD Key Validation**: Disabled by default (uncomment to re-enable) - should start with 'A' and be 11+ chars
- **Boolean Conversion**: Accepts TRUE/FALSE strings or Excel boolean cells
- **Router Requirement**: Equipment object always included in payload (API requires it) even if routerRequired=FALSE
- **Duplicate Prevention**: File locking prevents concurrent processing if Power Automate triggers multiple times
- **Module Dependencies**: `bb_ordering_api`, `config`, `graph_mailbox_check`
- **Regrade Limitation**: Regrade orders currently disabled due to Gamma server-side API issue - will be re-enabled when fixed
- **Status Check Workflow**: Orders should be checked 24+ hours after placement using bulk_check_orders.py with generated order ID file

---
**Version**: 1.1 | **Last Updated**: Feb 2026 | **Environment**: Test/Production | **Note**: Regrade functionality temporarily disabled
