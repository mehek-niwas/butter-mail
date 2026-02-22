const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { ImapFlow } = require('imapflow');

const configPath = path.join(app.getPath('userData'), 'imap-config.json');

function loadConfig() {
  try {
    const data = fs.readFileSync(configPath, 'utf8');
    return JSON.parse(data);
  } catch {
    return null;
  }
}

function saveConfig(config) {
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
}

function parseEmlFromSource(buffer) {
  if (!buffer) return { from: '', to: '', subject: '(no subject)', date: '', body: '', fromEmail: '' };
  const text = Buffer.isBuffer(buffer) ? buffer.toString('utf8') : String(buffer);
  const lines = text.split(/\r?\n/);
  const headers = {};
  let bodyStart = -1;
  let i = 0;

  for (; i < lines.length; i++) {
    const line = lines[i];
    if (line === '') {
      bodyStart = i;
      break;
    }
    const match = line.match(/^([\w-]+):\s*(.*)$/i);
    if (match) {
      const key = match[1].toLowerCase();
      let val = match[2];
      while (i + 1 < lines.length && /^\s/.test(lines[i + 1])) {
        i++;
        val += lines[i].trim();
      }
      headers[key] = val;
    }
  }

  const body = bodyStart >= 0 ? lines.slice(bodyStart + 1).join('\n').trim() : '';
  function extractEmail(str) {
    if (!str) return '';
    const m = str.match(/<([^>]+)>/);
    return m ? m[1] : str.trim();
  }

  return {
    from: headers.from || '',
    to: headers.to || '',
    subject: (headers.subject || '(no subject)').replace(/\s+/g, ' ').trim(),
    date: headers.date || '',
    body,
    fromEmail: extractEmail(headers.from)
  };
}

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  app.quit();
});

// IPC handlers
ipcMain.handle('imap:getConfig', () => loadConfig());

ipcMain.handle('imap:saveConfig', (_, config) => {
  saveConfig(config);
  return true;
});

ipcMain.handle('imap:fetch', async (_, limit = 50) => {
  const config = loadConfig();
  if (!config || !config.host || !config.user || !config.pass) {
    return { ok: false, error: 'IMAP not configured' };
  }

  const client = new ImapFlow({
    host: config.host,
    port: config.port || 993,
    secure: config.secure !== false,
    auth: { user: config.user, pass: config.pass },
    logger: false
  });

  try {
    await client.connect();
    await client.mailboxOpen('INBOX');
    const exists = client.mailbox.exists;
    if (exists === 0) {
      await client.logout();
      return { ok: true, emails: [] };
    }

    const count = Math.min(limit, exists);
    const start = Math.max(1, exists - count + 1);
    const range = start === count ? `${start}` : `${start}:${exists}`;

    const emails = [];
    for await (const msg of client.fetch(range, { envelope: true })) {
      const fromAddr = msg.envelope?.from?.[0];
      const fromStr = fromAddr
        ? (fromAddr.name ? `${fromAddr.name} <${fromAddr.address}>` : fromAddr.address)
        : '';
      emails.push({
        id: `imap-${msg.uid}`,
        uid: msg.uid,
        subject: msg.envelope?.subject || '(no subject)',
        from: fromStr,
        date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : '',
        body: '',
        fromEmail: fromAddr?.address || ''
      });
    }

    await client.logout();
    return { ok: true, emails };
  } catch (err) {
    const detail = err.response || err.responseCode || err.code || '';
    return { ok: false, error: err.message + (detail ? ` (${detail})` : '') };
  }
});

ipcMain.handle('imap:fetchOne', async (_, uid) => {
  const config = loadConfig();
  if (!config || !config.host || !config.user || !config.pass) {
    return { ok: false, error: 'IMAP not configured' };
  }

  const client = new ImapFlow({
    host: config.host,
    port: config.port || 993,
    secure: config.secure !== false,
    auth: { user: config.user, pass: config.pass },
    logger: false
  });

  try {
    await client.connect();
    await client.mailboxOpen('INBOX');
    const msg = await client.fetchOne(String(uid), { envelope: true, source: true }, { uid: true });
    let email = null;
    if (msg) {
      const parsed = parseEmlFromSource(msg.source);
      const fromAddr = msg.envelope?.from?.[0];
      email = {
        id: `imap-${msg.uid}`,
        subject: msg.envelope?.subject || parsed.subject,
        from: fromAddr ? (fromAddr.name ? `${fromAddr.name} <${fromAddr.address}>` : fromAddr.address) : parsed.from,
        date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : parsed.date,
        body: parsed.body,
        fromEmail: fromAddr?.address || parsed.fromEmail
      };
    }
    await client.logout();
    return { ok: true, email };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('imap:test', async (_, config) => {
  const client = new ImapFlow({
    host: config.host,
    port: config.port || 993,
    secure: config.secure !== false,
    auth: { user: config.user, pass: config.pass },
    logger: false
  });

  try {
    await client.connect();
    await client.mailboxOpen('INBOX');
    await client.logout();
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});
