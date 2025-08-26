# Contacts API Documentation

## Overview
The Contacts module provides comprehensive management of Microsoft Outlook contacts through the Microsoft Graph API. It supports full CRUD operations, advanced search capabilities, and contact folder management.

## Available Operations

### List Contacts
Retrieve a list of contacts with pagination support.

```javascript
{
  operation: "list",
  folderId: "optional-folder-id",  // Optional: specific folder
  top: 25,                         // Optional: max results (default: 25)
  skip: 0,                        // Optional: pagination offset
  select: "displayName,emailAddresses",  // Optional: fields to return
  orderBy: "displayName"         // Optional: sort field
}
```

### Search Contacts
Search contacts by various criteria with hybrid approach (server-side for emails, client-side for other fields).

```javascript
{
  operation: "search",
  query: "search term",           // Required: search query
  searchFields: ["displayName", "emailAddresses", "companyName"],  // Optional
  top: 10                         // Optional: max results
}
```

### Get Contact
Retrieve a specific contact by ID.

```javascript
{
  operation: "get",
  contactId: "contact-id-here"    // Required: contact ID
}
```

### Create Contact
Create a new contact with comprehensive field support.

```javascript
{
  operation: "create",
  displayName: "John Smith",       // Required
  givenName: "John",               // Optional
  surname: "Smith",                // Optional
  middleName: "Michael",           // Optional
  nickName: "Johnny",              // Optional
  emailAddresses: "john@example.com" or ["john@example.com", "smith@work.com"],
  businessPhones: ["555-1234"],
  homePhones: ["555-5678"],
  mobilePhone: "555-9999",
  companyName: "Acme Corp",
  department: "Engineering",
  jobTitle: "Senior Developer",
  officeLocation: "Building A",
  birthday: "2000-01-15",          // ISO 8601 format
  personalNotes: "Met at conference",
  businessAddress: {
    street: "123 Main St",
    city: "Seattle",
    state: "WA",
    postalCode: "98101",
    countryOrRegion: "USA"
  },
  homeAddress: {
    street: "456 Oak Ave",
    city: "Bellevue",
    state: "WA",
    postalCode: "98004",
    countryOrRegion: "USA"
  },
  categories: ["Client", "Important"]
}
```

### Update Contact
Update an existing contact (partial updates supported).

```javascript
{
  operation: "update",
  contactId: "contact-id-here",    // Required
  // Include any fields to update (same as create)
  jobTitle: "Director of Engineering",
  mobilePhone: "555-0000"
}
```

### Delete Contact
Delete a contact permanently.

```javascript
{
  operation: "delete",
  contactId: "contact-id-here"     // Required
}
```

### List Contact Folders
Get all contact folders in the mailbox.

```javascript
{
  operation: "list_folders"
}
```

### Create Contact Folder
Create a new contact folder.

```javascript
{
  operation: "create_folder",
  displayName: "Important Clients",  // Required
  parentFolderId: "optional-parent-id"  // Optional: for subfolders
}
```

## Field Types and Formats

### Email Addresses
Can be provided as:
- Single string: `"john@example.com"`
- Array of strings: `["john@example.com", "john.smith@work.com"]`
- Array of objects: `[{address: "john@example.com", name: "John Smith"}]`

### Phone Numbers
Can be provided as:
- Single string: `"555-1234"`
- Array of strings: `["555-1234", "555-5678"]`

### Addresses
Object format:
```javascript
{
  street: "123 Main St, Suite 100",
  city: "Seattle",
  state: "WA",
  postalCode: "98101",
  countryOrRegion: "USA"
}
```

### Categories
Array of strings representing categories/tags:
```javascript
["VIP", "Client", "Friend", "Colleague"]
```

### Birthday
ISO 8601 date format (year is optional):
- With year: `"1990-05-15"`
- Without year: `"--05-15"`

## Search Capabilities

The search operation uses a hybrid approach:
1. **Server-side search** for email addresses (exact matches)
2. **Client-side search** for all other fields (substring matches)

Searchable fields:
- `displayName` - Full name
- `givenName` - First name
- `surname` - Last name
- `emailAddresses` - Email addresses
- `companyName` - Company/organization
- `department` - Department
- `jobTitle` - Job title
- `businessPhones` - Business phone numbers
- `mobilePhone` - Mobile phone
- `categories` - Categories/tags

## Error Handling

Common error responses:
- `404 Not Found` - Contact or folder not found
- `400 Bad Request` - Invalid input format
- `403 Forbidden` - Insufficient permissions
- `429 Too Many Requests` - Rate limit exceeded

## Permissions Required

The following Microsoft Graph permission is required:
- `Contacts.ReadWrite` - Full read and write access to contacts (includes read permissions)

## Rate Limits

Microsoft Graph API rate limits apply:
- Recommended: Max 4 requests per second
- Batch operations available for bulk updates

## Examples

### Create a business contact
```javascript
await contacts({
  operation: "create",
  displayName: "Sarah Johnson",
  givenName: "Sarah",
  surname: "Johnson",
  emailAddresses: ["sarah.johnson@techcorp.com", "sarah@personal.com"],
  businessPhones: ["555-0100", "555-0101"],
  mobilePhone: "555-0102",
  companyName: "TechCorp Industries",
  department: "Sales",
  jobTitle: "Regional Sales Manager",
  officeLocation: "Tower B, Floor 15",
  businessAddress: {
    street: "789 Corporate Blvd",
    city: "San Francisco",
    state: "CA",
    postalCode: "94105",
    countryOrRegion: "USA"
  },
  categories: ["VIP", "Sales", "West Coast"],
  personalNotes: "Prefers email communication. Decision maker for west coast operations."
});
```

### Search for contacts at a company
```javascript
await contacts({
  operation: "search",
  query: "TechCorp",
  searchFields: ["companyName"],
  top: 50
});
```

### Update contact's job information
```javascript
await contacts({
  operation: "update",
  contactId: "AAMkADcyOG...",
  jobTitle: "Vice President of Sales",
  department: "Executive Team",
  officeLocation: "Executive Suite"
});
```