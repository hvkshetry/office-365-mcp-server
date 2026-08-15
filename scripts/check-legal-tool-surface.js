#!/usr/bin/env node

/** Fail closed when the Office surface reviewed for the legal persona drifts. */

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const modules = [
  require('../auth').authTools,
  require('../calendar').calendarTools,
  require('../email').emailTools,
  require('../teams'),
  require('../notifications').notificationTools,
  require('../planner').plannerTools,
  require('../files').filesTools,
  require('../search').searchTools,
  require('../contacts').contactsTools,
  require('../todo').todoTools,
  require('../groups').groupsTools,
  require('../directory').directoryTools
];
const { LEGAL_TOOL_MANIFEST } = require('../policy');

function normalize(value) {
  if (Array.isArray(value)) {
    return value.map(normalize);
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, item]) => [key, normalize(item)])
    );
  }
  return value;
}

function digest(value) {
  return crypto
    .createHash('sha256')
    .update(JSON.stringify(normalize(value)))
    .digest('hex');
}

function currentSurface() {
  const tools = modules
    .flat()
    .map(tool => ({ name: tool.name, input_schema: tool.inputSchema || {} }))
    .sort((left, right) => left.name.localeCompare(right.name));
  const reviewed = {};
  for (const [name, operations] of Object.entries(LEGAL_TOOL_MANIFEST)) {
    reviewed[name] = operations === null ? null : Array.from(operations).sort();
  }
  return {
    schema_version: 'puran.office-legal-tool-surface/v1',
    tool_names: tools.map(tool => tool.name),
    tool_schema_sha256: digest(tools),
    legal_persona_allowlist: reviewed
  };
}

const current = currentSurface();
if (process.argv.includes('--print-current')) {
  process.stdout.write(`${JSON.stringify(normalize(current), null, 2)}\n`);
  process.exit(0);
}

const lockPath = path.join(__dirname, '..', 'legal-tool-surface.lock.json');
const expected = JSON.parse(fs.readFileSync(lockPath, 'utf8'));
if (JSON.stringify(normalize(current)) !== JSON.stringify(normalize(expected))) {
  process.stderr.write('Legal-persona Office tool surface drifted; review before updating the lock.\n');
  process.stderr.write(`${JSON.stringify(normalize(current), null, 2)}\n`);
  process.exit(1);
}

process.stdout.write(`legal Office tool surface verified: ${current.tool_names.length} tools\n`);
