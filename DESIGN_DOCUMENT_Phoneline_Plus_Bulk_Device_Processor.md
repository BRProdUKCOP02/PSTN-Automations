# Phoneline+ Bulk Device Processor - Design Document

## Solution Overview
Automated hardware device ordering system that reads device orders from Excel templates and places bulk orders via the Gamma Phoneline+ Partner API. Supports multi-product orders, device-to-user assignment, and delivery address specification. Designed for ordering desk phones, headsets, and other hardware for existing Phoneline+ customers.

## How It Works

### Workflow
1. **File Discovery**: Scans input folder (`\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\...`) for Excel files
2. **File Locking**: Immediately moves file to processing folder to prevent duplicate processing
3. **Authentication**: Generates JWT token using partner API credentials from first row (keyID/secret)
4. **Validation**: Validates each row for required fields (customer ID, delivery details, product SKU, etc.)
5. **Order Processing**: Places hardware orders via API with delivery details and SKU-to-user mapping
6. **Report Generation**: Outputs Excel report with order numbers, product details, success/failure status
7. **Email Notification**: Sends report to psmanaged.delivery@gamma.co.uk and Dave.Siddall@gamma.co.uk
8. **File Archival**: Moves completed file to processed/completed folder with timestamp
9. **Delay Management**: 2-second delay between orders to avoid API rate limiting

### Order Structure
- **Single Row = One Order** containing multiple products/devices (same delivery address)
- Each product specified with: Product UUID, SKU code (e.g., "W73P"), quantity
- Optional SKU-to-user mapping assigns specific devices to specific user IDs
- Tracks order numbers returned by API for shipment tracking

### Error Handling
- **Validation Errors**: Skips row and logs missing/invalid fields in report
- **API Errors**: Captures error response and continues processing remaining orders
- **Authentication Failures**: Stops processing and reports authentication issue
- **File Locking**: Prevents concurrent processing if automation triggers multiple times

## API Integration

### Endpoints
- **Auth**: `https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/auth` (UAT) / `gammaapi.net` (Production)
- **Orders**: `.../partner/v1/customers/{customerId}/orders` - POST to place hardware orders

### Authentication
JWT Bearer Token generated via POST request with partner credentials:
```json
{"keyID": "<partner_key>", "secret": "<partner_secret>"}
```

### Request Payload Structure
```json
{
  "deliveryAddress": {
    "line1": "string",
    "line2": "string" (optional),
    "line3": "string" (optional),
    "town": "string",
    "county": "string" (optional),
    "country": "string" (optional, defaults to "United Kingdom"),
    "postcode": "string"
  },
  "products": [
    {
      "ID": "uuid-of-product-sku",
      "quantity": 1
    }
  ],
  "name": "string",
  "email": "string",
  "phoneNumber": "string",
  "trackingEmail": "string" (optional, defaults to email),
  "devices": [ /* same format as products */ ],
  "skuToUserMapping": {
    "SKU_CODE": [
      {"userID": "user-uuid"}
    ]
  } (optional)
}
```

## Key Variables & Configuration

### Class Constants
| Variable | Value | Description |
|----------|-------|-------------|
| `CYCLE_DELAY` | 2.0s | Delay between processing orders to avoid API rate limiting |

### Column Name Mapping (Aliases)
Flexible column naming - the processor normalizes these aliases to canonical names:
- `key id` / `keyid` → `keyID`
- `customer id` / `customerid` → `customer_id`
- `full name` / `contact name` → `name`
- `e-mail` / `email address` / `mail` → `email`
- `tracking email` / `trackingemail` → `tracking_email`
- `devices json` / `devices_json` → `devices_json`
- `sku mapping json` / `sku_mapping_json` → `sku_mapping_json`
- `phone` / `phone number` → `phone_number`
- `delivery line1` / `delivery line 1` → `delivery_line1`
- `delivery postcode` / `postcode` → `delivery_postcode`
- `product sku` → `product_id`
- `sku` / `sku code` / `product model` / `model` → `sku_code`
- `qty` → `quantity`
- `user uuid` / `userid` / `user id` → `user_id`

### Environment Variables
- `ENVIRONMENT`: "uat" or "production" - determines API endpoint selection
- `INPUT_FOLDER`: Network path for incoming Excel files
- `OUTPUT_FOLDER`: Local path for generated reports
- `PROCESSED_FOLDER`: Archive location for processed input files

## Mandatory Inputs (Excel Columns)

### Required Columns per Order Row
| Column | Type | Description | Validation |
|--------|------|-------------|------------|
| `keyID` | String | Partner API Key ID (first row only) | Non-empty |
| `secret` | String | Partner API Secret (first row only) | Non-empty |
| `customer_id` | String (UUID) | Customer UUID to place order for | Non-empty |
| `name` | String | Delivery contact name | Non-empty |
| `email` | String | Delivery contact email | Valid email format |
| `phone_number` | String | Delivery contact phone | Non-empty (leading zero preserved) |
| `delivery_line1` | String | Delivery address line 1 | Non-empty |
| `delivery_town` | String | Town/City | Non-empty |
| `delivery_postcode` | String | Postcode | Non-empty |
| `product_id` | String (UUID) | Product SKU UUID from catalog | Non-empty |
| `sku_code` | String | Product SKU code (e.g., "W73P", "W56H") | Non-empty (needed for device mapping) |

### Optional Columns
- `tracking_email`: Alternative email for shipment tracking (defaults to `email`)
- `delivery_line2`, `delivery_line3`: Additional address lines
- `delivery_county`: County name
- `delivery_country`: Country (defaults to "United Kingdom")
- `quantity`: Number of units (defaults to 1)
- `user_id`: User UUID for device assignment (creates SKU-to-user mapping)
- `devices_json`: JSON array override for devices payload (advanced use)
- `sku_mapping_json`: JSON object override for SKU-to-user mapping (advanced use)

### Advanced JSON Column Format
**devices_json** (optional array for custom device list):
```json
[
  {"ID": "product-uuid", "quantity": 2},
  {"ID": "another-product-uuid", "quantity": 1}
]
```

**sku_mapping_json** (optional object for multi-user device assignment):
```json
{
  "W73P": [{"userID": "user-uuid-1"}, {"userID": "user-uuid-2"}],
  "W56H": [{"userID": "user-uuid-3"}]
}
```

## Output Data

### Report Excel Structure
**Sheet 1 - Order Results**: Row per order with:
- `row`: Excel row number
- `timestamp`: Processing timestamp (ISO format)
- Input echo fields (customer_id, name, email, phone_number, delivery address, product details)
- `success`: Boolean order placement status
- `order_number`: Gamma order number (for tracking)
- `order_created`: Order creation timestamp from API
- `product_category`, `product_make`, `product_name`: Product details from API response
- `product_unit_cost`: Unit cost returned by API
- `error`: Error message if failed

**Sheet 2 - Summary**: Processing statistics (total, successful, failed, environment, timestamp)

## Technical Notes
- **Number Formatting**: Automatically adds leading zero to UK phone numbers (07xxx mobile, 01/02/03/08 landlines) if stripped by Excel
- **Duplicate Prevention**: File moved to processing folder immediately to prevent concurrent execution
- **Column Flexibility**: Accepts various column name formats (case-insensitive, space/underscore interchangeable)
- **Retry Strategy**: No automatic retries - single attempt per order
- **Module Dependencies**: `phoneline_plus_jwt_auth`, `phoneline_plus_place_order`, `graph_mailbox_check`
- **Multi-Product Support**: Single order can contain multiple product types with different quantities
- **Device Assignment**: SKU code (not UUID) used as key in mapping - critical for API to match devices correctly
- **Rate Limiting**: 2-second delay between orders prevents API throttling

---
**Version**: 1.0 | **Last Updated**: Feb 2026 | **Environment**: UAT/Production
