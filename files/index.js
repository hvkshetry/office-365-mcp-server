/**
 * OneDrive and SharePoint Files module
 * Provides comprehensive file management capabilities
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Unified files handler for all file operations
 */
async function handleFiles(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, get, upload, download, delete, share, search, move, copy, create_folder" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listFiles(accessToken, params);
      case 'get':
        return await getFile(accessToken, params);
      case 'upload':
        return await uploadFile(accessToken, params);
      case 'download':
        return await downloadFile(accessToken, params);
      case 'delete':
        return await deleteFile(accessToken, params);
      case 'share':
        return await shareFile(accessToken, params);
      case 'search':
        return await searchFiles(accessToken, params);
      case 'move':
        return await moveFile(accessToken, params);
      case 'copy':
        return await copyFile(accessToken, params);
      case 'create_folder':
        return await createFolder(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, get, upload, download, delete, share, search, move, copy, create_folder` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in files ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in files ${operation}: ${error.message}` }]
    };
  }
}

/**
 * List files in OneDrive or SharePoint
 */
async function listFiles(accessToken, params) {
  const { 
    path = '/me/drive/root', 
    siteId, 
    driveId,
    folderId,
    includeSubfolders = false,
    maxResults = 50 
  } = params;
  
  let endpoint = path;
  
  // Handle different contexts
  if (siteId && driveId) {
    endpoint = `/sites/${siteId}/drives/${driveId}/root`;
  } else if (driveId) {
    endpoint = `/drives/${driveId}/root`;
  } else if (folderId) {
    endpoint = `/me/drive/items/${folderId}`;
  }
  
  const queryParams = {
    $select: 'id,name,size,file,folder,createdDateTime,lastModifiedDateTime,webUrl',
    $top: maxResults
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${endpoint}/children`,
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No files found in the specified location." }]
    };
  }
  
  const filesList = response.value.map((item, index) => {
    const type = item.folder ? 'Folder' : 'File';
    const size = item.size ? `${(item.size / 1024 / 1024).toFixed(2)} MB` : 'N/A';
    return `${index + 1}. [${type}] ${item.name}
   Size: ${size}
   Modified: ${new Date(item.lastModifiedDateTime).toLocaleString()}
   ID: ${item.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} items:\n\n${filesList}` 
    }]
  };
}

/**
 * Get file metadata
 */
async function getFile(accessToken, params) {
  const { fileId, path } = params;
  
  if (!fileId && !path) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: fileId or path" 
      }]
    };
  }
  
  const endpoint = fileId 
    ? `/me/drive/items/${fileId}`
    : `/me/drive/root:${path}`;
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    {
      $select: 'id,name,size,file,folder,createdDateTime,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl'
    }
  );
  
  let details = `Name: ${response.name}\n`;
  details += `Type: ${response.folder ? 'Folder' : 'File'}\n`;
  details += `Size: ${response.size ? `${(response.size / 1024 / 1024).toFixed(2)} MB` : 'N/A'}\n`;
  details += `Created: ${new Date(response.createdDateTime).toLocaleString()}\n`;
  details += `Modified: ${new Date(response.lastModifiedDateTime).toLocaleString()}\n`;
  details += `Web URL: ${response.webUrl}\n`;
  
  if (response['@microsoft.graph.downloadUrl']) {
    details += `Download URL: ${response['@microsoft.graph.downloadUrl']}\n`;
  }
  
  return {
    content: [{ type: "text", text: details }]
  };
}

/**
 * Upload a file to OneDrive
 */
async function uploadFile(accessToken, params) {
  const { 
    fileName, 
    content, 
    parentPath = '/me/drive/root',
    parentId,
    conflictBehavior = 'rename' // rename, replace, fail
  } = params;
  
  if (!fileName || !content) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: fileName and content" 
      }]
    };
  }
  
  const endpoint = parentId
    ? `/me/drive/items/${parentId}:/${fileName}:/content`
    : `${parentPath}:/${fileName}:/content`;
  
  const queryParams = {
    '@microsoft.graph.conflictBehavior': conflictBehavior
  };
  
  // For small files (< 4MB), use simple upload
  const response = await callGraphAPI(
    accessToken,
    'PUT',
    endpoint,
    content,
    queryParams,
    {
      'Content-Type': 'application/octet-stream'
    }
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `File uploaded successfully!\nFile ID: ${response.id}\nName: ${response.name}\nSize: ${(response.size / 1024).toFixed(2)} KB` 
    }]
  };
}

/**
 * Download file content
 */
async function downloadFile(accessToken, params) {
  const { fileId, path } = params;
  
  if (!fileId && !path) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: fileId or path" 
      }]
    };
  }
  
  const endpoint = fileId 
    ? `/me/drive/items/${fileId}/content`
    : `/me/drive/root:${path}:/content`;
  
  try {
    const content = await callGraphAPI(
      accessToken,
      'GET',
      endpoint
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `File content downloaded successfully.\nContent length: ${content.length} bytes\n\nContent:\n${content}` 
      }]
    };
  } catch (error) {
    // For binary files, we might get a download URL instead
    const metadataEndpoint = fileId 
      ? `/me/drive/items/${fileId}`
      : `/me/drive/root:${path}`;
    
    const metadata = await callGraphAPI(
      accessToken,
      'GET',
      metadataEndpoint,
      null,
      { $select: '@microsoft.graph.downloadUrl' }
    );
    
    if (metadata['@microsoft.graph.downloadUrl']) {
      return {
        content: [{ 
          type: "text", 
          text: `File is binary. Download URL: ${metadata['@microsoft.graph.downloadUrl']}` 
        }]
      };
    }
    
    throw error;
  }
}

/**
 * Delete a file or folder
 */
async function deleteFile(accessToken, params) {
  const { fileId, path } = params;
  
  if (!fileId && !path) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: fileId or path" 
      }]
    };
  }
  
  const endpoint = fileId 
    ? `/me/drive/items/${fileId}`
    : `/me/drive/root:${path}`;
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    endpoint
  );
  
  return {
    content: [{ type: "text", text: "File/folder deleted successfully!" }]
  };
}

/**
 * Share a file
 */
async function shareFile(accessToken, params) {
  const { 
    fileId, 
    type = 'view', // view, edit, embed
    scope = 'anonymous', // anonymous, organization, users
    password,
    expirationDateTime,
    recipients
  } = params;
  
  if (!fileId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: fileId" 
      }]
    };
  }
  
  const body = {
    type: type,
    scope: scope
  };
  
  if (password) body.password = password;
  if (expirationDateTime) body.expirationDateTime = expirationDateTime;
  
  if (recipients && recipients.length > 0) {
    // Send sharing invitation
    const inviteBody = {
      requireSignIn: true,
      sendInvitation: true,
      roles: [type === 'edit' ? 'write' : 'read'],
      recipients: recipients.map(email => ({
        email: email
      })),
      message: "I've shared a file with you"
    };
    
    const response = await callGraphAPI(
      accessToken,
      'POST',
      `/me/drive/items/${fileId}/invite`,
      inviteBody
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `File shared successfully with ${recipients.join(', ')}` 
      }]
    };
  } else {
    // Create sharing link
    const response = await callGraphAPI(
      accessToken,
      'POST',
      `/me/drive/items/${fileId}/createLink`,
      body
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Sharing link created:\n${response.link.webUrl}\nType: ${type}\nScope: ${scope}` 
      }]
    };
  }
}

/**
 * Search for files
 */
async function searchFiles(accessToken, params) {
  const { 
    query, 
    scope = 'me', // me, sites, all
    fileTypes,
    maxResults = 25 
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  let searchQuery = query;
  if (fileTypes && fileTypes.length > 0) {
    const typeFilter = fileTypes.map(type => `filetype:${type}`).join(' OR ');
    searchQuery = `${query} AND (${typeFilter})`;
  }
  
  const endpoint = scope === 'me' 
    ? `/me/drive/search(q='${encodeURIComponent(searchQuery)}')`
    : `/search/query`;
  
  const body = scope !== 'me' ? {
    requests: [{
      entityTypes: ['driveItem'],
      query: {
        queryString: searchQuery
      },
      size: maxResults
    }]
  } : null;
  
  const response = await callGraphAPI(
    accessToken,
    scope === 'me' ? 'GET' : 'POST',
    endpoint,
    body,
    scope === 'me' ? { $top: maxResults } : {}
  );
  
  const items = scope === 'me' ? response.value : response.value[0].hitsContainers[0].hits;
  
  if (!items || items.length === 0) {
    return {
      content: [{ type: "text", text: "No files found matching your search." }]
    };
  }
  
  const results = items.map((item, index) => {
    const file = scope === 'me' ? item : item.resource;
    return `${index + 1}. ${file.name}
   Path: ${file.parentReference?.path || 'N/A'}
   WebURL: ${file.webUrl || 'N/A'}
   Modified: ${new Date(file.lastModifiedDateTime).toLocaleString()}
   ID: ${file.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${items.length} files:\n\n${results}` 
    }]
  };
}

/**
 * Move a file or folder
 */
async function moveFile(accessToken, params) {
  const { fileId, destinationId, newName } = params;
  
  if (!fileId || !destinationId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: fileId and destinationId" 
      }]
    };
  }
  
  const body = {
    parentReference: {
      id: destinationId
    }
  };
  
  if (newName) {
    body.name = newName;
  }
  
  const response = await callGraphAPI(
    accessToken,
    'PATCH',
    `/me/drive/items/${fileId}`,
    body
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `File moved successfully!\nNew location: ${response.parentReference.path}/${response.name}` 
    }]
  };
}

/**
 * Copy a file or folder
 */
async function copyFile(accessToken, params) {
  const { fileId, destinationId, newName } = params;
  
  if (!fileId || !destinationId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: fileId and destinationId" 
      }]
    };
  }
  
  const body = {
    parentReference: {
      id: destinationId
    }
  };
  
  if (newName) {
    body.name = newName;
  }
  
  // Copy operation returns a monitor URL
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `/me/drive/items/${fileId}/copy`,
    body
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Copy operation initiated. The file will be copied in the background.\nYou can monitor the operation if needed.` 
    }]
  };
}

/**
 * Create a folder
 */
async function createFolder(accessToken, params) {
  const { name, parentId, parentPath = '/me/drive/root' } = params;
  
  if (!name) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: name" 
      }]
    };
  }
  
  const endpoint = parentId
    ? `/me/drive/items/${parentId}/children`
    : `${parentPath}/children`;
  
  const body = {
    name: name,
    folder: {},
    '@microsoft.graph.conflictBehavior': 'rename'
  };
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    endpoint,
    body
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Folder created successfully!\nName: ${response.name}\nID: ${response.id}\nPath: ${response.parentReference.path}/${response.name}` 
    }]
  };
}

// Export consolidated tool
const filesTools = [
  {
    name: "files",
    description: "Manage OneDrive and SharePoint files: list, get, upload, download, delete, share, search, move, copy, create_folder",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "get", "upload", "download", "delete", "share", "search", "move", "copy", "create_folder"],
          description: "The operation to perform" 
        },
        // Common parameters
        fileId: { type: "string", description: "File or folder ID" },
        path: { type: "string", description: "File path (alternative to fileId)" },
        // List parameters
        siteId: { type: "string", description: "SharePoint site ID" },
        driveId: { type: "string", description: "Drive ID" },
        folderId: { type: "string", description: "Folder ID to list contents" },
        includeSubfolders: { type: "boolean", description: "Include subfolders in listing" },
        // Upload parameters
        fileName: { type: "string", description: "Name for uploaded file" },
        content: { type: "string", description: "File content to upload" },
        parentPath: { type: "string", description: "Parent folder path" },
        parentId: { type: "string", description: "Parent folder ID" },
        conflictBehavior: { 
          type: "string", 
          enum: ["rename", "replace", "fail"],
          description: "How to handle naming conflicts" 
        },
        // Share parameters
        type: { 
          type: "string", 
          enum: ["view", "edit", "embed"],
          description: "Share permission type" 
        },
        scope: { 
          type: "string", 
          enum: ["anonymous", "organization", "users"],
          description: "Share scope" 
        },
        password: { type: "string", description: "Password for shared link" },
        expirationDateTime: { type: "string", description: "Expiration for shared link" },
        recipients: { 
          type: "array", 
          items: { type: "string" },
          description: "Email addresses to share with" 
        },
        // Search parameters
        query: { type: "string", description: "Search query" },
        fileTypes: { 
          type: "array", 
          items: { type: "string" },
          description: "File types to filter (e.g., ['docx', 'pdf'])" 
        },
        // Move/Copy parameters
        destinationId: { type: "string", description: "Destination folder ID" },
        newName: { type: "string", description: "New name for file/folder" },
        // Create folder parameters
        name: { type: "string", description: "Folder name" },
        // General parameters
        maxResults: { type: "number", description: "Maximum number of results" }
      },
      required: ["operation"]
    },
    handler: handleFiles
  },
  {
    name: "files_map_sharepoint_path",
    description: "Map SharePoint webUrl to local sync path",
    inputSchema: {
      type: "object",
      properties: {
        webUrl: { 
          type: "string", 
          description: "SharePoint web URL to map to local sync path" 
        }
      },
      required: ["webUrl"]
    },
    handler: async (args) => {
      const { webUrl } = args;
      
      if (!webUrl) {
        return {
          content: [{
            type: "text",
            text: "Error: webUrl is required"
          }]
        };
      }
      
      // Base sync path
      const baseSyncPath = "/mnt/c/Users/hvksh/Circle H2O LLC";
      
      try {
        // Parse URL and extract path components
        const url = new URL(webUrl);
        const pathParts = decodeURIComponent(url.pathname).split('/').filter(p => p);
        
        // Find the site name and document library
        let siteName = '';
        let docLibraryIndex = -1;
        let pathAfterDocLib = [];
        
        // Look for pattern: /sites/[site-name]/[document-library]/...
        const sitesIndex = pathParts.findIndex(p => p === 'sites');
        if (sitesIndex !== -1 && sitesIndex + 1 < pathParts.length) {
          siteName = pathParts[sitesIndex + 1];
          
          // Find document library (usually "Shared Documents" or similar)
          for (let i = sitesIndex + 2; i < pathParts.length; i++) {
            if (pathParts[i].includes('Documents') || pathParts[i] === 'Shared') {
              docLibraryIndex = i;
              // Skip the library name itself, get the path after it
              pathAfterDocLib = pathParts.slice(i + 1);
              break;
            }
          }
        }
        
        // If we couldn't find the document library, try a simpler approach
        if (docLibraryIndex === -1) {
          // Look for any "Documents" in the path
          for (let i = 0; i < pathParts.length; i++) {
            if (pathParts[i].includes('Documents')) {
              docLibraryIndex = i;
              pathAfterDocLib = pathParts.slice(i + 1);
              break;
            }
          }
        }
        
        // Transform the site name for local folder
        let localSiteFolderName = '';
        if (siteName) {
          // Transform SharePoint site name to local folder name
          // Examples:
          // "Intema-UASBInternals" -> "Intema - UASB Internals - Documents"
          // "ChemfieldCellulosePvtLtd" -> "Chemfield Cellulose Pvt Ltd - Documents"
          // "CBGMeerut" -> "CBG Meerut - Documents"
          // "Proposals" -> "Proposals - Documents"
          // "ProjectName-SubProject" -> "Project Name - Sub Project - Documents"
          
          // Replace hyphens with spaces and add proper formatting
          localSiteFolderName = siteName
            .replace(/([a-z])([A-Z])/g, '$1 $2') // Add space between camelCase
            .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2') // Add space between consecutive caps and following word
            .replace(/(Pvt|Ltd|LLC|Inc|Corp)([A-Z])/g, '$1 $2') // Add space after common abbreviations
            .replace(/-/g, ' - ') // Replace hyphens with space-hyphen-space
            .replace(/\s+/g, ' ') // Clean up multiple spaces
            .trim();
          
          // Add "- Documents" suffix to ALL project folders
          // SharePoint sites sync locally with "- Documents" suffix
          if (!localSiteFolderName.endsWith('- Documents') && !localSiteFolderName.endsWith('Documents')) {
            localSiteFolderName += ' - Documents';
          }
        }
        
        // Build the complete local path
        const pathComponents = [];
        
        // Add the transformed site folder if we have one
        if (localSiteFolderName) {
          pathComponents.push(localSiteFolderName);
        }
        
        // Add the rest of the path
        pathComponents.push(...pathAfterDocLib);
        
        // Construct the local path
        const localPath = pathComponents.length > 0 
          ? `${baseSyncPath}/${pathComponents.join('/')}`
          : baseSyncPath;
        
        // Also provide symlink path for subagent use
        const symlinkPath = pathComponents.length > 0
          ? `/home/hvksh/admin/temp/sharepoint/${pathComponents.join('/')}`
          : '/home/hvksh/admin/temp/sharepoint';
        
        return {
          content: [{
            type: "text",
            text: `SharePoint URL mapped successfully:\n\nSharePoint: ${webUrl}\n\nLocal path (for main agent):\n${localPath}\n\nSymlink path (for subagent):\n${symlinkPath}`
          }]
        };
      } catch (error) {
        return {
          content: [{
            type: "text",
            text: `Error mapping SharePoint URL: ${error.message}`
          }]
        };
      }
    }
  }
];

module.exports = { filesTools };