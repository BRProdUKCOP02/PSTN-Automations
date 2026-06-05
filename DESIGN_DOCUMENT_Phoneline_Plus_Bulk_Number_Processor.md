# Phoneline+ Bulk Number Processor - Design Document

## Solution Overview
Automated geographical number reallocation system that swaps user-assigned non-geographical numbers with available geographical numbers on customer accounts. Fetches all numbers and users for a customer, identifies assigned numbers, prioritizes non-geographical number users, and performs allocation/deallocation operations via the Gamma Phoneline+ Partner API.

## How It Works

### Workflow
1. **File Discovery**: Scans input folder for Excel files with customer IDs
2. **File Locking**: Immediately moves file to processing folder to prevent duplicate processing
3. **Authentication**: Generates JWT token using partner API credentials (keyID/secret from Excel)
4. **Customer Analysis**: Fetches all numbers and users for customer
5. **Number Classification**: Categorizes numbers as geographical vs. non-geographical
6. **User Mapping**: Identifies which users have which numbers assigned
7. **Prioritization**: Prioritizes non-geographical users first when available numbers limited
8. **Number Swapping**: 
   - Allocates available geographical number to user
   - Deallocates old non-geographical number from user
   - Both operations tracked per user
9. **Report Generation**: Outputs Excel report with swap details, success/failure status
10. **Email Notification**: Sends report to delivery team
11. **File Archival**: Moves completed file to processed folder
12. **Delay Management**: 2-second delay between customers

### Reallocation Logic
- **Fetch Numbers**: Get all numbers on customer account (assigned + available)
- **Identify Available**: Find numbers not assigned to any user
- **Prioritize Geo Numbers**: Use geographical numbers first from available pool
- **Match Users**: Limit swaps to min(users_with_numbers, available_numbers)
- **Prioritize Non-Geo Users**: When numbers scarce, swap non-geographical users first
- **Two-Step Swap**: Allocate new number first, then deallocate old number
- **Track Failures**: Records allocation/deallocation success independently

### Error Handling
- **No Available Numbers**: Logs error and skips customer
- **Partial Failures**: If allocation succeeds but deallocation fails, flags error but counts allocation
- **API Errors**: Captures error response and continues processing remaining customers
- **Authentication Failures**: Stops processing and reports authentication issue

## API Integration

### Endpoints
- **Auth**: `https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/auth` (UAT) / `gammaapi.net` (Production)
- **Numbers**: `.../partner/v1/customers/{customerId}/numbers` - GET to fetch all customer numbers
- **Users**: `.../partner/v1/customers/{customerId}/users` - GET to fetch all customer users
- **Allocate**: `.../partner/v1/customers/{customerId}/numbers/{numberId}` - PUT to assign number to user
- **Deallocate**: `.../partner/v1/customers/{customerId}/numbers/{numberId}` - PUT to unassign number

### Authentication
JWT Bearer Token generated via POST request with partner credentials:
```json
{"keyID": "<partner_key>", "secret": "<partner_secret>"}
```

### Request Payload Structure
**Allocate Number**:
```json
{
  "allocateTo": {
    "userID": "user-uuid"
  }
}
```

**Deallocate Number**:
```json
{
  "allocateTo": null
}
```

## Key Variables & Configuration

### Class Constants
| Variable | Value | Description |
|----------|-------|-------------|
| `CYCLE_DELAY` | 2.0s | Delay between processing customers to allow API preparation |

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
| `customer_id` | String (UUID) | Customer UUID to process | Non-empty |

### No Other Columns Required
This processor only needs the customer ID - all number and user data is fetched automatically from the API.

## Output Data

### Report Excel Structure
**Sheet 1 - Reallocation Results**: Row per swap operation with:
- `timestamp`: Processing timestamp
- `input_customer_id`: Customer UUID from input
- `success`: Boolean overall customer processing status
- `total_numbers`: Total numbers on account
- `geo_numbers`: Count of geographical numbers
- `nongeo_numbers`: Count of non-geographical numbers
- `total_users`: Total users on account
- `users_with_numbers`: Count of users with assigned numbers
- `users_without_numbers`: Count of users without numbers
- `reallocations`: Count of successful allocations
- `new_allocations`: Count of new number assignments (not used in current version)
- `deallocations`: Count of successful deallocations
- `failed_operations`: Count of failed operations
- `customer_error`: Overall customer-level error message
- **Operation Details (per row)**:
  - `operation_type`: "swap"
  - `user_id`: User UUID
  - `old_number`: Previous number (E.164 format)
  - `old_type`: Previous number type display name
  - `new_number`: New allocated number (E.164 format)
  - `new_type`: New number type display name
  - `allocation_success`: Boolean allocation outcome
  - `deallocation_success`: Boolean deallocation outcome
  - `operation_error`: Operation-specific error message

**Sheet 2 - Summary**: Processing statistics (total customers, successful, failed, total swaps, environment, timestamp)

### Number Type Classifications
- **Geographical** (`standard_geographic`): 01/02/03 UK area code numbers
- **Non-Geographical** (`standard_nongeographic`): 033/0800/0845 UK service numbers
- **Display Names**: User-friendly type names from API (e.g., "UK Geographic Number", "UK Non-Geographic Number")

## Swap Operation Priority Matrix

### When Available Numbers < Users with Numbers
1. **Sort Users by Priority**:
   - Priority 0: Users with non-geographical numbers (swapped first)
   - Priority 1: Users with geographical numbers (swapped later)
2. **Limit Swaps**: Only swap up to count of available numbers
3. **Warning Logged**: Reports that not all users were processed due to limited availability

### Example Scenario
- Customer has: 5 users with numbers (3 non-geo, 2 geo)
- Available numbers: 3 geographical
- **Result**: 3 non-geo users swapped to geographical, 2 geo users unchanged

## Technical Notes
- **Two-Phase Swap**: Allocation happens first, deallocation second - reduces risk of user losing number
- **E.164 Format**: All numbers displayed in international format (e.g., +441234567890)
- **Number Pool Ordering**: Available pool ordered with geographical numbers first for preference
- **Independent Tracking**: Allocation and deallocation tracked separately - allocation success + deallocation failure = partial success
- **Module Dependencies**: `phoneline_plus_jwt_auth`, `phoneline_plus_number_management`, `graph_mailbox_check`
- **No Automatic Retry**: Single attempt per operation
- **User ID Mapping**: Uses internal number allocation mapping (`allocatedTo.userID`) to determine current assignments
- **Empty Account Handling**: If no users have numbers, reports success with "No operations needed"

---
**Version**: 1.0 | **Last Updated**: Feb 2026 | **Environment**: UAT/Production
