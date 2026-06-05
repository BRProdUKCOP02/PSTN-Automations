# Bulk Check Orders - Design Document

## Solution Overview
Automated broadband order status verification system that reads order IDs from Excel files and retrieves detailed status information via the Gamma Broadband Ordering API. Designed for status checks 24+ hours after order placement. Provides comprehensive order details including installation progress, engineer appointments, equipment delivery, and historical updates.

## How It Works

### Workflow
1. **Manual Trigger**: User places Excel file with order IDs in SharePoint/OneDrive folder (24+ hours after order placement)
2. **File Discovery**: Script scans input folder for Excel files
3. **Authentication**: Obtains JWT bearer token using Gamma API credentials from config.py
4. **Order ID Loading**: Reads Order_ID column from Excel
5. **Status Retrieval**: Fetches detailed order information from API for each order ID
6. **Data Flattening**: Converts nested JSON responses into flat Excel-friendly format
7. **Update Extraction**: Separates order history updates into dedicated sheet
8. **Report Generation**: Creates multi-sheet Excel with:
   - **Order Details**: Flattened order data (one row per order)
   - **Order Updates**: Historical status changes (one row per update)
   - **Summary**: Processing statistics
9. **Email Notification**: Sends report to psmanaged.delivery@gamma.co.uk and Dave.Siddall@gamma.co.uk
10. **File Archival**: Moves input file to processed folder
11. **Delay Management**: 0.5-second delay between API calls

### Use Case Scenario
1. SoGEA bulk ordering script places 50 orders on Monday morning
2. Script generates order ID file: `orders_order_ids_20260224_093000.xlsx`
3. User waits 24-48 hours for orders to progress
4. User moves order ID file to Bulk Status Checker folder
5. This script auto-runs, fetches status for all 50 orders
6. Report shows current status (e.g., "In Progress", "Awaiting Engineer", "Completed")
7. Report includes engineer appointment dates/times if scheduled

### Error Handling
- **Order Not Found**: Reports HTTP 404 error and continues processing remaining orders
- **API Errors**: Captures error response (status code + message), continues processing
- **Authentication Failures**: Stops processing and reports authentication issue
- **Missing Order_ID Column**: Stops and reports validation error

## API Integration

### Endpoints
- **Auth**: `https://api-test.gamma.co.uk/auth/token` (Test) / `https://api.gamma.co.uk/auth/token` (Production)
- **Get Order**: `https://api-test.gamma.co.uk/broadband/v2/orders/{orderId}` - GET to retrieve order details

### Authentication
OAuth2 Password Grant - JWT Bearer Token obtained via POST:
```
Content-Type: application/x-www-form-urlencoded
grant_type=password&username=<api_user>&password=<api_password>
```

### API Response Structure (Flattened in Report)
```json
{
  "id": 123456,
  "status": "In Progress",
  "accountNumber": 12345,
  "broadbandProduct": "SoGEA_FTTC_40_10",
  "voiceProduct": "None",
  "careLevel": "Standard",
  "routedIpOption": "",
  "ipAddressOption": "Single",
  "resellerEmailNotifications": true,
  "underRegrade": false,
  "installation": {
    "lineType": "Existing",
    "cli": "01234567890",
    "customerRequiredDate": "2026-03-01",
    "customerReference": "REF123",
    "type": "engineer",
    "supplierPromisedDate": "2026-03-05",
    "site": {
      "companyName": "ACME Corp",
      "nadKey": "A00010584545",
      "address": { /* ... */ },
      "contact": { /* ... */ }
    },
    "engineerAppointment": {
      "date": "2026-03-05",
      "timeslot": "AM",
      "timeslotStart": "08:00",
      "timeslotEnd": "13:00"
    }
  },
  "resellerContact": { /* ... */ },
  "equipment": {
    "routerRequired": true,
    "router": "Technicolor TG582n v3",
    "companyName": "ACME Corp",
    "deliveryContact": { /* ... */ },
    "deliveryAddress": { /* ... */ },
    "routerConfiguration": {
      "dslUsername": "user@gamma",
      "wanIpAddress": "123.45.67.89"
    }
  },
  "numberPort": {
    "voipReference": "PORT123",
    "portOrderId": 789,
    "portOrderStatus": "Pending"
  },
  "updates": [
    {
      "timestamp": "2026-02-24T10:30:00Z",
      "status": "Order Placed",
      "message": "Order received and validated"
    },
    {
      "timestamp": "2026-02-25T14:15:00Z",
      "status": "In Progress",
      "message": "Engineer appointment scheduled"
    }
  ]
}
```

## Key Variables & Configuration

### Script Configuration (in script)
| Variable | Value | Description |
|----------|-------|-------------|
| `CHECK_DELAY` | 0.5s | Delay between API calls to avoid rate limiting |
| `INPUT_FOLDER` | OneDrive network path | Source folder for order ID files |
| `OUTPUT_FOLDER` | ./order_check_output | Results report destination |
| `PROCESSED_FOLDER` | INPUT_FOLDER/processed | Archive for completed check files |

### External Configuration (config.py)
- `ENVIRONMENT`: "test" or "production" - determines API endpoint
- Gamma API credentials: username, password

## Mandatory Inputs (Excel Columns)

### Required Column
| Column | Type | Description | Validation |
|--------|------|-------------|------------|
| `Order_ID` | Integer | Gamma order ID to check | Numeric, non-empty |

### No Other Columns Required
This checker only needs order IDs - all order data is fetched from the API.

### Input File Source
Typically generated by `sogea_bulk_ordering.py` with filename pattern:
- `{original_input}_order_ids_{timestamp}.xlsx`

## Output Data

### Report Excel Structure

#### Sheet 1 - Order Details
Flattened order data with columns:
- `Check_Status`: "SUCCESS", "FAILED", "ERROR"
- `Check_Error`: Error message if check failed
- **Basic Order Info**:
  - `Order_ID`: Gamma order ID
  - `Status`: Current order status
  - `Account_Number`: Account number
  - `Broadband_Product`: Product code
  - `Voice_Product`: Voice product
  - `Care_Level`: Service care level
  - `Routed_IP_Option`: Routed IP allocation
  - `IP_Address_Option`: IP address type
  - `Reseller_Email_Notifications`: Boolean
  - `Under_Regrade`: Boolean (true if being upgraded)
- **Installation Details**:
  - `Line_Type`: "Existing" or "New"
  - `CLI`: Circuit line identifier
  - `Customer_Required_Date`: Requested installation date
  - `Customer_Reference`: Customer reference string
  - `Installation_Type`: "engineer" or "self"
  - `Supplier_Promised_Date`: Gamma's promised completion date
- **Site Information**:
  - `Site_Company_Name`: Site company name
  - `Site_NAD_Key`: NAD key
  - `Site_Contact_Name/Email/Phone`: Site contact details
  - `Site_Building/SubPremises/Street/Town/County/Postcode`: Full address
- **Engineer Appointment**:
  - `Engineer_Appointment_Date`: Scheduled date
  - `Engineer_Appointment_Timeslot`: "AM", "PM", "Specific"
  - `Engineer_Appointment_Start`: Start time (HH:MM)
  - `Engineer_Appointment_End`: End time (HH:MM)
- **Reseller Contact**:
  - `Reseller_Contact_Name/Email/Phone`: Reseller details
- **Equipment Details**:
  - `Router_Required`: Boolean
  - `Router_Model`: Router model name
  - `Equipment_Company_Name`: Company name for delivery
  - `Delivery_Contact_Name/Email/Phone`: Delivery contact
  - `Delivery_Building/Street/Town/County/Postcode`: Delivery address
  - `Router_DSL_Username`: DSL authentication username
  - `Router_WAN_IP_Address`: WAN IP address
- **Number Port Details**:
  - `Number_Port_VOIP_Reference`: Porting reference
  - `Number_Port_Order_ID`: Port order ID
  - `Number_Port_Status`: Port order status
- **Update Summary** (for last update only):
  - `Last_Update_Timestamp`: Most recent update time
  - `Last_Update_Status`: Most recent status
  - `Last_Update_Message`: Most recent message
  - `Total_Updates`: Count of all updates

#### Sheet 2 - Order Updates
Historical order status changes (one row per update):
- `Order_ID`: Order ID
- `Update_Number`: Sequential update number (1, 2, 3...)
- `Timestamp`: Update timestamp (ISO format)
- `Status`: Status at this update
- `Message`: Status message/description

#### Sheet 3 - Summary
Processing statistics:
- `Total Orders Checked`: Count of order IDs in input
- `Successful Checks`: Count where API returned data
- `Failed Checks`: Count where API returned error (e.g., 404)
- `Errors`: Count of unexpected errors
- `Environment`: "test" or "production"
- `Check Date/Time`: Processing timestamp

## Common Order Statuses

### Status Values
- `Order Placed`: Order submitted to Gamma
- `In Progress`: Order actively being provisioned
- `Awaiting Engineer`: Pending engineer appointment
- `Awaiting Customer`: Waiting on customer action
- `Completed`: Order successfully fulfilled
- `Cancelled`: Order cancelled
- `Failed`: Order failed provisioning
- `On Hold`: Order paused

### Appointment Timeslots
- `AM`: 08:00-13:00
- `PM`: 13:00-18:00
- `Specific`: Custom time range (check timeslotStart/timeslotEnd)

## Technical Notes
- **Flat Structure**: Nested JSON converted to flat column structure for Excel compatibility
- **Update Separation**: Main sheet shows only last update; all updates in dedicated sheet
- **Error Resilience**: Individual order failures don't stop processing of remaining orders
- **Auto-Column Sizing**: Report columns auto-sized for readability (max 50 chars)
- **Empty Values**: Missing fields shown as empty strings in report
- **Timestamp Format**: ISO 8601 format preserved from API
- **Module Dependencies**: `bb_ordering_api`, `config`, `graph_mailbox_check`
- **Intended Use**: Run 24-48 hours after order placement for meaningful status updates
- **Rate Limiting**: 0.5s delay between calls prevents API throttling
- **File Pattern Handling**: Skips temporary Excel files (starting with ~$)
- **Duplicate Safety**: Appends timestamp to moved file if destination already exists

---
**Version**: 1.0 | **Last Updated**: Feb 2026 | **Environment**: Test/Production
