import os from 'os';
import { watch } from 'fs';

const name = (await Bun.file("package.json").json()).name as string;

let defaultLogsPath = '';
switch (process.platform) {
  case 'darwin':
    defaultLogsPath = os.homedir() + '/Library/Logs/Claude/mcp-server-' + name + '.log';
    break;
  case 'win32':
    defaultLogsPath = os.homedir() + '/AppData/Local/Claude/mcp-server-' + name + '.log';
    break;
  default:
    break;
}
const logsPath = process.env.CLAUDE_LOGS_PATH || defaultLogsPath;

if (!logsPath)
  throw new Error("CLAUDE_LOGS_PATH is not set");

const args = process.argv.slice(2);

if (args.includes('--clear')) {
  await Bun.write(logsPath, '');
  console.log('Logs cleared');  
  process.exit(0);
}

const initialLogs = await Bun.file(logsPath).text();
console.log(initialLogs);

if (args.includes('--follow')) {
  let lastSize = (await Bun.file(logsPath).stat())?.size || 0;
  
  watch(logsPath, async (eventType) => {
    if (eventType === 'change') {
      const newSize = (await Bun.file(logsPath).stat())?.size || 0;
      if (newSize > lastSize) {
        const newContent = await Bun.file(logsPath).text();
        console.log(newContent.slice(lastSize));
        lastSize = newSize;
      }
    }
  });
}

