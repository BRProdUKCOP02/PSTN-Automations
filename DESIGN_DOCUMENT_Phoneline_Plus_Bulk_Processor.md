# Phoneline+ Bulk Customer Processor - Design Document

## Solution Overview
Automated bulk customer provisioning system that reads customer data from Excel templates and creates customers and users via the Gamma Phoneline+ Partner API. Supports Hardware Only account types with SIP credentials, multi-user creation per customer, and automatic phone number allocation with in-flight inventory conflict resolution.

## How It Works

### Workflow
1. **File Discovery**: Scans input folder (`\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\...`) for Excel files
2. **File Locking**: Immediately moves file to processing folder to prevent duplicate processing if automation triggers multiple times
3. **Authentication**: Generates JWT token using partner API credentials (keyID/secret from Excel row)
4. **Number Pre-Allocation**: Pre-fetches 20 available phone numbers into pool for user creation
5. **Customer Creation**: Creates Hardware Only customer with retry logic (up to 3 attempts) for in-flight number conflicts
6. **User Provisioning**: Creates additional users for customer using pre-allocated numbers (up to 10 retries per user)
7. **Report Generation**: Outputs Excel report with customer IDs, SIP credentials, success/failure status
8. **Email Notification**: Sends report to psmanaged.delivery@gamma.co.uk and Dave.Siddall@gamma.co.uk
9. **File Archival**: Moves completed file to processed/completed folder with timestamp

### Error Handling
- **In-Flight Conflicts**: If selected number has existing order, automatically fetches replacement and retries
- **Timeout Recovery**: Prompts operator to manually verify/provide customer ID for user creation continuation
- **Validation Flags**: Marks records needing verification when API returns success but no customer/user ID

## API Integration

### Endpoints
- **Auth**: `https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/auth` (UAT) / `gammaapi.net` (Production)
- **Customers**: `.../partner/v1/customers` - POST to create Hardware Only customers
- **Users**: `.../partner/v1/customers/{customerId}/users` - POST to add users to customers
- **Numbers**: `.../partner/v1/numbers/available` - GET to fetch available numbers (standard_geographic, area 44161)

### Authentication
JWT Bearer Token generated via POST request with partner credentials:
```json
{"keyID": "<partner_key>", "secret": "<partner_secret>"}
```

### Request Payload Structure
**Customer Creation**:
```json
{
  "plan": "hardware_only" | "hardware_only_cp",
  "companyName": "string",
  "fullName": "string",
  "contactNumber": "string",
  "email": "string" (optional),
  "number": "string" (required for hardware_only, optional for hardware_only_cp),
  "address": {
    "premises": "string",
    "street": "string", 
    "town": "string",
    "county": "string",
    "postcode": "string"
  }
}
```

**User Creation**:
```json
{
  "fullName": "string",
  "email": "string",
  "phoneNumber": "string" (auto-allocated from available pool),
  "type": "standard" | other,
  "address": { /* same structure */ }
}
```

## Key Variables & Configuration

### Class Constants
| Variable | Value | Description |
|----------|-------|-------------|
| `CYCLE_DELAY` | 3.5s | Delay between processing customers to allow API preparation |
| `USER_NUMBER_TYPE` | "standard_geographic" | Type of numbers to allocate for users |
| `USER_AREA_CODE` | "44161" | Manchester area code for allocated numbers |
| `USER_NUMBER_MAX_RETRIES` | 8 | Max retries when fetching available numbers |
| `USER_NUMBER_RETRY_DELAY_SECONDS` | 2.0 | Delay between number fetch retries |
| `USER_INFLIGHT_RETRY_ATTEMPTS` | 10 | Max retry attempts per user when number in-flight |
| `CUSTOMER_INFLIGHT_RETRY_ATTEMPTS` | 3 | Max retry attempts per customer when number in-flight |
| `NUMBER_POOL_SIZE` | 20 | Pre-fetch pool size for efficient user creation |

### Environment Variables
- `ENVIRONMENT`: "uat" or "production" - determines API endpoint selection
- `INPUT_FOLDER`: Network path for incoming Excel files
- `OUTPUT_FOLDER`: Local path for generated reports
- `PROCESSED_FOLDER`: Archive location for processed input files

## Mandatory Inputs (Excel Columns)

### Required Columns per Customer Row
| Column | Type | Description | Validation |
|--------|------|-------------|------------|
| `keyID` | String | Partner API Key ID | Non-empty |
| `secret` | String | Partner API Secret | Non-empty |
| `companyName` | String | Customer company name | Non-empty |
| `fullName` | String | Primary contact full name | Non-empty |
| `number` | String | Phone number for customer | Required if plan=hardware_only, optional if hardware_only_cp |
| `contactNumber` | String | Contact phone number | Non-empty |
| `plan` | String | Account plan type | Must be "hardware_only" or "hardware_only_cp" |
| `premises` | String | Address line 1 | Non-empty |
| `street` | String | Street name | Non-empty |
| `town` | String | Town/City | Non-empty |
| `county` | String | County | Non-empty |
| `postcode` | String | Postcode (auto-formatted, spaces removed) | Non-empty |

### Optional Columns
- `email`: Customer email address (null if not provided)

### Multi-User Columns (Repeating Pattern)
Supports unlimited users per customer using pattern: `userN_<field>` where N = 1, 2, 3, ...
- `userN_fullName` (required for user)
- `userN_email` (required for user)
- `userN_phoneNumber` (optional - auto-allocated if not provided)
- `userN_type` (optional - defaults to "standard")
- `userN_premises`, `userN_street`, `userN_town`, `userN_county`, `userN_postcode` (all optional)

**Example**: `user1_fullName`, `user2_fullName`, `user3_fullName` creates 3 users

## Output Data

### Report Excel Structure
**Sheet 1 - Customer Results**: Row per customer/user with:
- Input echo fields (companyName, fullName, email, etc.)
- `customer_id`: Gamma-assigned customer ID
- `user_id`: Portal user ID (from customer creation)
- `assigned_number`: Allocated phone number (E.164 format)
- `registration_server`: SIP registration server
- `sip_id`: SIP username
- `sip_password`: SIP password
- `users_created`, `users_failed`: User creation counts
- `success`: Boolean status
- `needs_verification`: Flag for manual verification required
- `error`: Error message if failed

**Sheet 2 - Summary**: Processing statistics (total, successful, failed, needs verification, users created/failed, environment, timestamp)

## Technical Notes
- **Number Formatting**: Automatically adds leading zero to UK numbers (07xxx mobile, 01/02/03/08 landlines) if stripped by Excel
- **Duplicate Prevention**: File moved to processing folder immediately to prevent concurrent execution duplicates
- **Retry Strategy**: Exponential backoff not used; fixed delays with excluded number tracking
- **Excel Text Format**: Numbers formatted as text (`@`) to prevent scientific notation in output
- **Module Dependencies**: `phoneline_plus_jwt_auth`, `phoneline_plus_create_customer`, `phoneline_plus_create_user`, `graph_mailbox_check`
- **Timeout Handling**: 409 customer timeout errors trigger manual intervention prompt with option to continue user creation
- **Logging**: Unique run ID generated per execution for traceability

---
**Version**: 1.0 | **Last Updated**: Feb 2026 | **Environment**: UAT/Production
