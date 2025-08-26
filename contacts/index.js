/**
 * Microsoft Contacts API module
 * Provides full contact management capabilities through Microsoft Graph API
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Main contacts handler
 */
async function handleContacts(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, search, get, create, update, delete, list_folders, create_folder" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listContacts(accessToken, params);
      case 'search':
        return await searchContacts(accessToken, params);
      case 'get':
        return await getContact(accessToken, params);
      case 'create':
        return await createContact(accessToken, params);
      case 'update':
        return await updateContact(accessToken, params);
      case 'delete':
        return await deleteContact(accessToken, params);
      case 'list_folders':
        return await listContactFolders(accessToken, params);
      case 'create_folder':
        return await createContactFolder(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, search, get, create, update, delete, list_folders, create_folder` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in contacts ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in contacts ${operation}: ${error.message}` }]
    };
  }
}

/**
 * List all contacts
 */
async function listContacts(accessToken, params) {
  const { 
    top = 50,
    skip = 0,
    orderBy = 'displayName',
    select,
    folderId
  } = params;
  
  const queryParams = {
    $top: top,
    $skip: skip,
    $orderby: orderBy
  };
  
  if (select) {
    queryParams.$select = select;
  }
  
  const endpoint = folderId 
    ? `/me/contactFolders/${folderId}/contacts`
    : '/me/contacts';
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No contacts found." }]
    };
  }
  
  const contactsList = response.value.map((contact, index) => {
    const emails = contact.emailAddresses || [];
    const phones = contact.businessPhones || [];
    // Handle both string and array formats for emails
    const emailStr = Array.isArray(emails) 
      ? emails.map(e => typeof e === 'object' ? e.address : e).join(', ')
      : emails;
    const phoneStr = Array.isArray(phones) ? phones.join(', ') : phones || '';
    
    return `${index + 1}. ${contact.displayName || contact.givenName + ' ' + contact.surname}
   Email: ${emailStr || 'N/A'}
   Phone: ${phoneStr || 'N/A'}
   Company: ${contact.companyName || 'N/A'}
   Title: ${contact.jobTitle || 'N/A'}
   ID: ${contact.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} contacts:\n\n${contactsList}` 
    }]
  };
}

/**
 * Search contacts by query
 * Due to Microsoft Graph API limitations, only email address filtering is supported server-side
 * Name searches are performed client-side
 */
async function searchContacts(accessToken, params) {
  const { 
    query,
    searchFields = ['displayName', 'emailAddresses', 'givenName', 'surname', 'companyName'],
    top = 100  // Fetch more for client-side filtering
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  const lowerQuery = query.toLowerCase();
  let response;
  
  // Check if query looks like an email address
  const isEmailSearch = query.includes('@');
  
  if (isEmailSearch) {
    // Use server-side filter for email searches (only supported filter)
    // Microsoft Graph only supports exact email matching with specific syntax
    const queryParams = {
      $filter: `emailAddresses/any(a:a/address eq '${query}')`,
      $top: top,
      $orderby: 'displayName'
    };
    
    try {
      response = await callGraphAPI(
        accessToken,
        'GET',
        '/me/contacts',
        null,
        queryParams
      );
    } catch (error) {
      // If exact match fails, fall back to fetching all and filtering client-side
      console.error('Email filter failed, falling back to client-side search:', error);
      const fallbackParams = {
        $top: top,
        $orderby: 'displayName'
      };
      response = await callGraphAPI(
        accessToken,
        'GET',
        '/me/contacts',
        null,
        fallbackParams
      );
      
      // Filter client-side for partial email matches
      if (response.value) {
        response.value = response.value.filter(contact => {
          const emails = contact.emailAddresses || [];
          return emails.some(e => e.address && e.address.toLowerCase().includes(lowerQuery));
        });
      }
    }
  } else {
    // For name searches, fetch all contacts and filter client-side
    // This is necessary because Microsoft Graph doesn't support filtering by name fields
    const queryParams = {
      $top: top,
      $orderby: 'displayName',
      $select: 'id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,businessPhones,mobilePhone'
    };
    
    response = await callGraphAPI(
      accessToken,
      'GET',
      '/me/contacts',
      null,
      queryParams
    );
    
    // Client-side filtering for name fields
    if (response.value) {
      response.value = response.value.filter(contact => {
        // Check each field for matches
        const fieldsToCheck = [];
        
        if (searchFields.includes('displayName') && contact.displayName) {
          fieldsToCheck.push(contact.displayName);
        }
        if (searchFields.includes('givenName') && contact.givenName) {
          fieldsToCheck.push(contact.givenName);
        }
        if (searchFields.includes('surname') && contact.surname) {
          fieldsToCheck.push(contact.surname);
        }
        if (searchFields.includes('companyName') && contact.companyName) {
          fieldsToCheck.push(contact.companyName);
        }
        if (searchFields.includes('emailAddresses') && contact.emailAddresses) {
          contact.emailAddresses.forEach(email => {
            if (email.address) fieldsToCheck.push(email.address);
            if (email.name) fieldsToCheck.push(email.name);
          });
        }
        
        // Check if any field contains the search query
        return fieldsToCheck.some(field => 
          field.toLowerCase().includes(lowerQuery)
        );
      });
    }
  }
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: `No contacts found matching "${query}".` }]
    };
  }
  
  // Limit results for display
  const displayResults = response.value.slice(0, 25);
  
  const contactsList = displayResults.map((contact, index) => {
    const emails = contact.emailAddresses || [];
    const phones = contact.businessPhones || [];
    const emailStr = emails.map(e => {
      if (e.name && e.address) {
        return `${e.name} <${e.address}>`;
      }
      return e.address || '';
    }).filter(e => e).join(', ');
    const phoneStr = phones.join(', ');
    
    return `${index + 1}. ${contact.displayName || `${contact.givenName || ''} ${contact.surname || ''}`.trim() || 'No Name'}
   Email: ${emailStr || 'N/A'}
   Phone: ${phoneStr || contact.mobilePhone || 'N/A'}
   Company: ${contact.companyName || 'N/A'}
   Title: ${contact.jobTitle || 'N/A'}
   ID: ${contact.id}`;
  }).join('\n\n');
  
  const resultCount = response.value.length;
  const displayCount = displayResults.length;
  const countMessage = resultCount > displayCount 
    ? `Showing ${displayCount} of ${resultCount} contacts matching "${query}":`
    : `Found ${resultCount} contact${resultCount === 1 ? '' : 's'} matching "${query}":`;
  
  return {
    content: [{ 
      type: "text", 
      text: `${countMessage}\n\n${contactsList}` 
    }]
  };
}

/**
 * Get a specific contact by ID
 */
async function getContact(accessToken, params) {
  const { contactId } = params;
  
  if (!contactId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: contactId" 
      }]
    };
  }
  
  const contact = await callGraphAPI(
    accessToken,
    'GET',
    `/me/contacts/${contactId}`
  );
  
  const emails = contact.emailAddresses || [];
  const businessPhones = contact.businessPhones || [];
  const homePhones = contact.homePhones || [];
  const addresses = contact.businessAddress || contact.homeAddress || {};
  
  let details = `Contact Details:
Name: ${contact.displayName || contact.givenName + ' ' + contact.surname}
Given Name: ${contact.givenName || 'N/A'}
Surname: ${contact.surname || 'N/A'}
Middle Name: ${contact.middleName || 'N/A'}
Nickname: ${contact.nickName || 'N/A'}

Contact Information:`;
  
  if (emails.length > 0) {
    details += '\nEmails:';
    emails.forEach(email => {
      details += `\n  - ${email.name || 'Primary'}: ${email.address}`;
    });
  }
  
  if (businessPhones.length > 0) {
    details += '\nBusiness Phones:';
    businessPhones.forEach(phone => {
      details += `\n  - ${phone}`;
    });
  }
  
  if (homePhones.length > 0) {
    details += '\nHome Phones:';
    homePhones.forEach(phone => {
      details += `\n  - ${phone}`;
    });
  }
  
  details += `\nMobile: ${contact.mobilePhone || 'N/A'}

Professional Information:
Company: ${contact.companyName || 'N/A'}
Job Title: ${contact.jobTitle || 'N/A'}
Department: ${contact.department || 'N/A'}
Office Location: ${contact.officeLocation || 'N/A'}

Personal Information:
Birthday: ${contact.birthday || 'N/A'}
Personal Notes: ${contact.personalNotes || 'N/A'}

ID: ${contact.id}`;
  
  return {
    content: [{ type: "text", text: details }]
  };
}

/**
 * Create a new contact
 */
async function createContact(accessToken, params) {
  const {
    givenName,
    surname,
    displayName,
    middleName,
    nickName,
    emailAddresses,
    businessPhones,
    homePhones,
    mobilePhone,
    companyName,
    jobTitle,
    department,
    officeLocation,
    businessAddress,
    homeAddress,
    birthday,
    personalNotes,
    categories,
    folderId
  } = params;
  
  if (!givenName && !surname && !displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "At least one of givenName, surname, or displayName is required" 
      }]
    };
  }
  
  // Build contact object
  const contactData = {};
  
  // Name fields
  if (givenName) contactData.givenName = givenName;
  if (surname) contactData.surname = surname;
  if (displayName) {
    contactData.displayName = displayName;
  } else if (givenName && surname) {
    contactData.displayName = `${givenName} ${surname}`;
  }
  if (middleName) contactData.middleName = middleName;
  if (nickName) contactData.nickName = nickName;
  
  // Email addresses - handle both string and array formats
  if (emailAddresses) {
    if (typeof emailAddresses === 'string') {
      contactData.emailAddresses = [{
        address: emailAddresses,
        name: contactData.displayName || emailAddresses
      }];
    } else if (Array.isArray(emailAddresses)) {
      contactData.emailAddresses = emailAddresses.map(email => {
        if (typeof email === 'string') {
          return {
            address: email,
            name: contactData.displayName || email
          };
        }
        // Ensure it's a proper object with address property
        return {
          address: email.address || email,
          name: email.name || contactData.displayName || email.address || email
        };
      });
    }
  }
  
  // Phone numbers
  if (businessPhones) {
    contactData.businessPhones = Array.isArray(businessPhones) ? businessPhones : [businessPhones];
  }
  if (homePhones) {
    contactData.homePhones = Array.isArray(homePhones) ? homePhones : [homePhones];
  }
  if (mobilePhone) contactData.mobilePhone = mobilePhone;
  
  // Professional info
  if (companyName) contactData.companyName = companyName;
  if (jobTitle) contactData.jobTitle = jobTitle;
  if (department) contactData.department = department;
  if (officeLocation) contactData.officeLocation = officeLocation;
  
  // Addresses
  if (businessAddress) contactData.businessAddress = businessAddress;
  if (homeAddress) contactData.homeAddress = homeAddress;
  
  // Other fields
  if (birthday) contactData.birthday = birthday;
  if (personalNotes) contactData.personalNotes = personalNotes;
  if (categories) contactData.categories = Array.isArray(categories) ? categories : [categories];
  
  const endpoint = folderId 
    ? `/me/contactFolders/${folderId}/contacts`
    : '/me/contacts';
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    endpoint,
    contactData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Contact created successfully!\nName: ${response.displayName}\nID: ${response.id}` 
    }]
  };
}

/**
 * Update an existing contact
 */
async function updateContact(accessToken, params) {
  const { contactId, ...updateData } = params;
  
  if (!contactId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: contactId" 
      }]
    };
  }
  
  // Handle email addresses format conversion if needed
  if (updateData.emailAddresses) {
    if (typeof updateData.emailAddresses === 'string') {
      updateData.emailAddresses = [{
        address: updateData.emailAddresses,
        name: updateData.displayName || updateData.emailAddresses
      }];
    } else if (Array.isArray(updateData.emailAddresses)) {
      updateData.emailAddresses = updateData.emailAddresses.map(email => {
        if (typeof email === 'string') {
          return {
            address: email,
            name: updateData.displayName || email
          };
        }
        return email;
      });
    }
  }
  
  // Handle phone arrays
  if (updateData.businessPhones && !Array.isArray(updateData.businessPhones)) {
    updateData.businessPhones = [updateData.businessPhones];
  }
  if (updateData.homePhones && !Array.isArray(updateData.homePhones)) {
    updateData.homePhones = [updateData.homePhones];
  }
  if (updateData.categories && !Array.isArray(updateData.categories)) {
    updateData.categories = [updateData.categories];
  }
  
  const response = await callGraphAPI(
    accessToken,
    'PATCH',
    `/me/contacts/${contactId}`,
    updateData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Contact updated successfully!\nName: ${response.displayName}\nID: ${response.id}` 
    }]
  };
}

/**
 * Delete a contact
 */
async function deleteContact(accessToken, params) {
  const { contactId } = params;
  
  if (!contactId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: contactId" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `/me/contacts/${contactId}`
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Contact deleted successfully! ID: ${contactId}` 
    }]
  };
}

/**
 * List contact folders
 */
async function listContactFolders(accessToken, params) {
  const { top = 20 } = params;
  
  const queryParams = {
    $top: top,
    $orderby: 'displayName'
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    '/me/contactFolders',
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No contact folders found." }]
    };
  }
  
  const foldersList = response.value.map((folder, index) => {
    return `${index + 1}. ${folder.displayName}
   ID: ${folder.id}
   Parent: ${folder.parentFolderId || 'Root'}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} contact folders:\n\n${foldersList}` 
    }]
  };
}

/**
 * Create a contact folder
 */
async function createContactFolder(accessToken, params) {
  const { displayName, parentFolderId } = params;
  
  if (!displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: displayName" 
      }]
    };
  }
  
  const folderData = { displayName };
  
  const endpoint = parentFolderId 
    ? `/me/contactFolders/${parentFolderId}/childFolders`
    : '/me/contactFolders';
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    endpoint,
    folderData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Contact folder created successfully!\nName: ${response.displayName}\nID: ${response.id}` 
    }]
  };
}

// Export contacts tools
const contactsTools = [
  {
    name: "contacts",
    description: "Manage Outlook contacts: list, search, get, create, update, delete contacts and folders",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "search", "get", "create", "update", "delete", "list_folders", "create_folder"],
          description: "The operation to perform" 
        },
        // Common parameters
        contactId: { type: "string", description: "Contact ID (for get, update, delete)" },
        folderId: { type: "string", description: "Contact folder ID" },
        
        // List parameters
        top: { type: "number", description: "Maximum number of results" },
        skip: { type: "number", description: "Number of results to skip" },
        orderBy: { type: "string", description: "Field to order by" },
        select: { type: "string", description: "Fields to select" },
        
        // Search parameters
        query: { type: "string", description: "Search query" },
        searchFields: { 
          type: "array",
          items: { type: "string" },
          description: "Fields to search in" 
        },
        
        // Contact fields (for create/update)
        givenName: { type: "string", description: "First name" },
        surname: { type: "string", description: "Last name" },
        displayName: { type: "string", description: "Display name" },
        middleName: { type: "string", description: "Middle name" },
        nickName: { type: "string", description: "Nickname" },
        emailAddresses: { 
          description: "Email addresses (string or array)",
          oneOf: [
            { type: "string" },
            { 
              type: "array",
              items: {
                oneOf: [
                  { type: "string" },
                  { 
                    type: "object",
                    properties: {
                      address: { type: "string" },
                      name: { type: "string" }
                    }
                  }
                ]
              }
            }
          ]
        },
        businessPhones: { 
          description: "Business phone numbers",
          oneOf: [
            { type: "string" },
            { type: "array", items: { type: "string" } }
          ]
        },
        homePhones: { 
          description: "Home phone numbers",
          oneOf: [
            { type: "string" },
            { type: "array", items: { type: "string" } }
          ]
        },
        mobilePhone: { type: "string", description: "Mobile phone number" },
        companyName: { type: "string", description: "Company name" },
        jobTitle: { type: "string", description: "Job title" },
        department: { type: "string", description: "Department" },
        officeLocation: { type: "string", description: "Office location" },
        businessAddress: { 
          type: "object",
          description: "Business address",
          properties: {
            street: { type: "string" },
            city: { type: "string" },
            state: { type: "string" },
            countryOrRegion: { type: "string" },
            postalCode: { type: "string" }
          }
        },
        homeAddress: { 
          type: "object",
          description: "Home address",
          properties: {
            street: { type: "string" },
            city: { type: "string" },
            state: { type: "string" },
            countryOrRegion: { type: "string" },
            postalCode: { type: "string" }
          }
        },
        birthday: { type: "string", description: "Birthday (ISO 8601)" },
        personalNotes: { type: "string", description: "Personal notes" },
        categories: { 
          description: "Categories",
          oneOf: [
            { type: "string" },
            { type: "array", items: { type: "string" } }
          ]
        }
      },
      required: ["operation"]
    },
    handler: handleContacts
  }
];

module.exports = { contactsTools };