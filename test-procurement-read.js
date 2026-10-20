const assert = require('node:assert/strict');
const { requestFor } = require('./procurement-read');
const base = { operation: 'mail_delta', mailbox: 'hersh@circleh2o.com', folder_id: 'folder' };
assert.equal(requestFor(base).endpoint, 'users/hersh%40circleh2o.com/mailFolders/folder/messages/delta');
const cursor = 'https://graph.microsoft.com/v1.0/users/hersh%40circleh2o.com/mailFolders/folder/messages/delta?$deltatoken=opaque';
assert.ok(requestFor({ ...base, cursor }).endpoint.endsWith('?$deltatoken=opaque'));
for (const bad of [
  cursor.replace('graph.microsoft.com', 'evil.test'),
  cursor.replace('/folder/', '/other/'),
  cursor.replace('https:', 'http:'),
  cursor.replace('/messages/delta', '/sendMail')
]) assert.throws(() => requestFor({ ...base, cursor: bad }));
assert.throws(() => requestFor({ ...base, folder_id: '../Inbox' }));
assert.throws(() => requestFor({ operation: 'send' }));
assert.ok(!('$filter' in requestFor(base).query), 'Initial mailbox sync must not hit the filtered-delta ceiling');
assert.equal(requestFor({ operation: 'folders', mailbox: 'hersh@circleh2o.com' }).query.includeHiddenFolders, true);
console.log('procurement read endpoint and cursor tests passed');
