/**
 * Email rules management - list and create inbox rules
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Unified email rules handler
 */
async function handleEmailRules(args) {
  const { operation, enhanced = false, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations are: list, create"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'list':
        return enhanced ?
          await listEmailRulesEnhanced(accessToken, params.mailbox) :
          await listEmailRules(accessToken, params.mailbox);
      case 'create':
        return enhanced ?
          await createEmailRuleEnhanced(accessToken, params) :
          await createEmailRule(accessToken, params);
      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid operations are: list, create`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email rules ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in email rules: ${error.message}` }]
    };
  }
}

async function listEmailRules(accessToken, mailbox = null) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/mailFolders/inbox/messageRules`
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email rules found." }]
    };
  }

  const rulesList = response.value.map((rule, index) => {
    return `${index + 1}. ${rule.displayName}\n   Enabled: ${rule.isEnabled}\n   ID: ${rule.id}\n`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} email rules:\n\n${rulesList}`
    }]
  };
}

async function listEmailRulesEnhanced(accessToken, mailbox = null) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/mailFolders/inbox/messageRules`
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email rules found." }]
    };
  }

  const rulesList = response.value.map((rule, index) => {
    let details = `${index + 1}. ${rule.displayName}\n`;
    details += `   Enabled: ${rule.isEnabled}\n`;
    details += `   ID: ${rule.id}\n`;

    if (rule.conditions?.fromAddresses?.length > 0) {
      details += `   From: ${rule.conditions.fromAddresses.map(a => a.emailAddress.address).join(', ')}\n`;
    }

    if (rule.actions?.moveToFolder) {
      details += `   Action: Move to folder ${rule.actions.moveToFolder}\n`;
    }

    return details;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} email rules (enhanced view):\n\n${rulesList}`
    }]
  };
}

async function createEmailRule(accessToken, params) {
  const { displayName, fromAddresses, moveToFolder, forwardTo, mailbox } = params;

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: displayName"
      }]
    };
  }

  const rule = {
    displayName,
    sequence: 1,
    isEnabled: true,
    conditions: {},
    actions: {}
  };

  if (fromAddresses && fromAddresses.length > 0) {
    rule.conditions.fromAddresses = fromAddresses.map(email => ({
      emailAddress: { address: email }
    }));
  }

  if (moveToFolder) {
    rule.actions.moveToFolder = moveToFolder;
  }

  if (forwardTo && forwardTo.length > 0) {
    rule.actions.forwardTo = forwardTo.map(email => ({
      emailAddress: { address: email }
    }));
  }

  const response = await callGraphAPI(
    accessToken,
    'POST',
    `${config.getMailboxPrefix(mailbox)}/mailFolders/inbox/messageRules`,
    rule
  );

  return {
    content: [{
      type: "text",
      text: `Email rule created successfully!\nRule ID: ${response.id}`
    }]
  };
}

async function createEmailRuleEnhanced(accessToken, params) {
  const { displayName, fromAddresses, moveToFolder, forwardTo, subjectContains, importance, mailbox } = params;

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: displayName"
      }]
    };
  }

  const rule = {
    displayName,
    sequence: 1,
    isEnabled: true,
    conditions: {},
    actions: {}
  };

  if (fromAddresses && fromAddresses.length > 0) {
    rule.conditions.fromAddresses = fromAddresses.map(email => ({
      name: email,
      address: email
    }));
  }

  if (subjectContains && subjectContains.length > 0) {
    rule.conditions.subjectContains = subjectContains;
  }

  if (importance) {
    rule.conditions.importance = importance;
  }

  if (moveToFolder) {
    rule.actions.moveToFolder = moveToFolder;
  }

  if (forwardTo && forwardTo.length > 0) {
    rule.actions.forwardTo = forwardTo.map(email => ({
      emailAddress: {
        name: email,
        address: email
      }
    }));
  }

  const response = await callGraphAPI(
    accessToken,
    'POST',
    `${config.getMailboxPrefix(mailbox)}/mailFolders/inbox/messageRules`,
    rule
  );

  return {
    content: [{
      type: "text",
      text: `Enhanced email rule created successfully!\nRule ID: ${response.id}`
    }]
  };
}

module.exports = { handleEmailRules };
