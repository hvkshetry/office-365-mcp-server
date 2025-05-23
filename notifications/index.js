/**
 * Change notifications functionality module
 */
const { subscriptionTools } = require('./subscriptions');

// Export all Notification tools
const notificationTools = [
  ...subscriptionTools
];

module.exports = {
  notificationTools
};
