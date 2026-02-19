/**
 * Email categories management
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Manage email categories
 */
async function handleEmailCategories(args) {
  const { operation, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations are: list, create, update, delete, apply, remove"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'list':
        return await listCategories(accessToken);
      case 'create':
        return await createCategory(accessToken, params);
      case 'update':
        return await updateCategory(accessToken, params);
      case 'delete':
        return await deleteCategory(accessToken, params);
      case 'apply':
        return await applyCategory(accessToken, params);
      case 'remove':
        return await removeCategory(accessToken, params);
      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in categories ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error: ${error.message}` }]
    };
  }
}

async function listCategories(accessToken) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/outlook/masterCategories'
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No categories found." }]
    };
  }

  const categoriesList = response.value.map((cat, index) => {
    return `${index + 1}. ${cat.displayName} (Color: ${cat.color})`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Categories:\n${categoriesList}`
    }]
  };
}

async function createCategory(accessToken, params) {
  const { displayName, color = 'preset0', mailbox } = params;

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: displayName"
      }]
    };
  }

  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/outlook/masterCategories',
    { displayName, color }
  );

  return {
    content: [{
      type: "text",
      text: `Category '${displayName}' created with color ${color}`
    }]
  };
}

async function updateCategory(accessToken, params) {
  const { categoryId, displayName, color } = params;

  if (!categoryId) {
    return {
      content: [{ type: "text", text: "Missing required parameter: categoryId" }]
    };
  }

  const update = {};
  if (displayName) update.displayName = displayName;
  if (color) update.color = color;

  await callGraphAPI(
    accessToken,
    'PATCH',
    `me/outlook/masterCategories/${categoryId}`,
    update
  );

  return {
    content: [{ type: "text", text: "Category updated successfully!" }]
  };
}

async function deleteCategory(accessToken, params) {
  const { categoryId } = params;

  if (!categoryId) {
    return {
      content: [{ type: "text", text: "Missing required parameter: categoryId" }]
    };
  }

  await callGraphAPI(
    accessToken,
    'DELETE',
    `me/outlook/masterCategories/${categoryId}`
  );

  return {
    content: [{ type: "text", text: "Category deleted successfully!" }]
  };
}

async function applyCategory(accessToken, params) {
  const { emailId, categories, mailbox } = params;

  if (!emailId || !categories) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameters: emailId and categories"
      }]
    };
  }

  await callGraphAPI(
    accessToken,
    'PATCH',
    `${config.getMailboxPrefix(mailbox)}/messages/${emailId}`,
    { categories: Array.isArray(categories) ? categories : [categories] }
  );

  return {
    content: [{ type: "text", text: "Categories applied successfully!" }]
  };
}

async function removeCategory(accessToken, params) {
  const { emailId, categories, mailbox } = params;

  if (!emailId || !categories) {
    return {
      content: [{ type: "text", text: "Missing required parameters: emailId and categories" }]
    };
  }

  // Get current categories, then remove the specified ones
  const email = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/messages/${emailId}`,
    null,
    { $select: 'categories' }
  );

  const categoriesToRemove = Array.isArray(categories) ? categories : [categories];
  const updatedCategories = (email.categories || []).filter(c => !categoriesToRemove.includes(c));

  await callGraphAPI(
    accessToken,
    'PATCH',
    `${config.getMailboxPrefix(mailbox)}/messages/${emailId}`,
    { categories: updatedCategories }
  );

  return {
    content: [{ type: "text", text: "Categories removed successfully!" }]
  };
}

module.exports = { handleEmailCategories };
