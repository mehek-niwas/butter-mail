const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const nodemailer = require('nodemailer');
const embeddingsService = require('./embeddings-service');
const pcaUtils = require('./pca-utils');
const clustering = require('./clustering');
const searchService = require('./search-service');

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

function getSmtpOptions(config) {
  if (!config || !config.user || !config.pass) return null;
  let host = config.smtpHost;
  let port = config.smtpPort;
  if (!host && config.host) {
    if (config.host.includes('gmail.com')) {
      host = 'smtp.gmail.com';
      port = port || 587;
    } else if (config.host.includes('office365.com') || config.host.includes('outlook.')) {
      host = 'smtp.office365.com';
      port = port || 587;
    } else {
      host = config.host.replace(/^imap\./i, 'smtp.');
      port = port || 587;
    }
  }
  if (!port) port = 587;
  return { host, port, user: config.user, pass: config.pass };
}

async function parseEmlFromSource(buffer) {
  if (!buffer) return { from: '', to: '', subject: '(no subject)', date: '', body: '', bodyIsHtml: false, fromEmail: '', toDisplay: '', messageId: '', inReplyTo: '', references: '' };
  try {
    const parsed = await simpleParser(Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer));
    const body = parsed.html || parsed.text || '';
    const bodyIsHtml = !!parsed.html;
    const fromAddr = parsed.from?.value?.[0];
    const messageId = parsed.messageId ? (Array.isArray(parsed.messageId) ? parsed.messageId[0] : String(parsed.messageId)).trim() : '';
    const inReplyTo = parsed.inReplyTo ? (Array.isArray(parsed.inReplyTo) ? parsed.inReplyTo.join(' ') : String(parsed.inReplyTo)).trim() : '';
    const references = parsed.references ? (Array.isArray(parsed.references) ? parsed.references.join(' ') : String(parsed.references)).trim() : '';
    return {
      from: parsed.from?.text || '',
      to: parsed.to?.text || '',
      subject: (parsed.subject || '(no subject)').replace(/\s+/g, ' ').trim(),
      date: parsed.date ? parsed.date.toISOString() : '',
      body,
      bodyIsHtml,
      fromEmail: fromAddr?.address || '',
      toDisplay: parsed.to?.text || '',
      messageId,
      inReplyTo,
      references
    };
  } catch {
    return { from: '', to: '', subject: '(no subject)', date: '', body: '', bodyIsHtml: false, fromEmail: '', toDisplay: '', messageId: '', inReplyTo: '', references: '' };
  }
}

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    icon: path.join(__dirname, 'assets', 'bread-butter.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile('landing.html');
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
      const parsed = await parseEmlFromSource(msg.source);
      const fromAddr = msg.envelope?.from?.[0];
      email = {
        id: `imap-${msg.uid}`,
        subject: msg.envelope?.subject || parsed.subject,
        from: fromAddr ? (fromAddr.name ? `${fromAddr.name} <${fromAddr.address}>` : fromAddr.address) : parsed.from,
        date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : parsed.date,
        body: parsed.body,
        bodyIsHtml: parsed.bodyIsHtml,
        toDisplay: parsed.toDisplay,
        fromEmail: fromAddr?.address || parsed.fromEmail,
        messageId: parsed.messageId || '',
        inReplyTo: parsed.inReplyTo || '',
        references: parsed.references || ''
      };
    }
    await client.logout();
    return { ok: true, email };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('imap:fetchThreadHeaders', async (_, uids) => {
  const config = loadConfig();
  if (!config || !config.host || !config.user || !config.pass || !uids || uids.length === 0) {
    return { ok: true, headers: {} };
  }
  console.log('[butter-mail] timeline: fetchThreadHeaders starting for', uids.length, 'uids');
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
    const headers = {};
    const maxUids = Math.min(uids.length, 300);
    for (let i = 0; i < maxUids; i++) {
      const uid = uids[i];
      try {
        const msg = await client.fetchOne(String(uid), { source: { start: 0, maxLength: 8192 } }, { uid: true });
        if (msg && msg.source) {
          const parsed = await parseEmlFromSource(msg.source);
          headers[uid] = {
            messageId: parsed.messageId || '',
            inReplyTo: parsed.inReplyTo || '',
            references: parsed.references || ''
          };
        }
      } catch (_) {}
    }
    await client.logout();
    console.log('[butter-mail] timeline: fetchThreadHeaders done. headers:', Object.keys(headers).length);
    return { ok: true, headers };
  } catch (err) {
    return { ok: false, error: err.message || String(err), headers: {} };
  }
});

ipcMain.handle('imap:fetchBodies', async (_, uids) => {
  const config = loadConfig();
  if (!config || !config.host || !config.user || !config.pass || !uids || uids.length === 0) {
    return { ok: true, bodies: {} };
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
    const bodies = {};
    for (const uid of uids) {
      const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
      if (msg) {
        const parsed = await parseEmlFromSource(msg.source);
        bodies[uid] = { body: parsed.body, bodyIsHtml: parsed.bodyIsHtml };
      }
    }
    await client.logout();
    return { ok: true, bodies };
  } catch (err) {
    return { ok: false, error: err.message || String(err), bodies: {} };
  }
});

ipcMain.handle('embeddings:compute', async (_, emails) => {
  try {
    console.log('[butter-mail] embeddings:compute starting for', emails ? emails.length : 0, 'emails');
    const onProgress = (p) => {
      console.log(
        '[butter-mail] embeddings:compute progress',
        p && typeof p.current === 'number' ? p.current : '?',
        '/',
        p && typeof p.total === 'number' ? p.total : '?',
        p && p.message ? p.message : ''
      );
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('embeddings:progress', p);
      }
    };
    const results = await embeddingsService.computeEmbeddings(emails, onProgress);
    console.log('[butter-mail] embeddings:compute finished successfully.');
    return { ok: true, embeddings: results };
  } catch (err) {
    console.error('[butter-mail] embeddings:compute failed:', err);
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('embeddings:pca', async (_, embeddings, emailIds) => {
  try {
    const matrix = emailIds.map((id) => embeddings[id]).filter(Boolean);
    const ids = emailIds.filter((id) => embeddings[id]);
    if (matrix.length === 0) return { ok: true, points: [], model: null };
    console.log('[butter-mail] PCA (graph view): fitting and projecting', ids.length, 'points to 3D');
    const { points, model } = pcaUtils.fitAndProject(matrix, 3);
    const pointMap = {};
    ids.forEach((id, i) => { pointMap[id] = points[i]; });
    console.log('[butter-mail] PCA (graph view): done.');
    return { ok: true, points: pointMap, model };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('embeddings:cluster', async (_, embeddings, emailIds) => {
  try {
    const matrix = emailIds.map((id) => embeddings[id]).filter(Boolean);
    const ids = emailIds.filter((id) => embeddings[id]);
    if (matrix.length === 0) return { ok: true, assignments: {}, meta: {} };
    console.log('[butter-mail] clustering: running DBSCAN on', ids.length, 'emails');
    const { assignments, meta } = clustering.cluster(matrix, ids);
    console.log('[butter-mail] clustering: done.');
    return { ok: true, assignments, meta };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('embeddings:query', async (_, query) => {
  try {
    const vec = await embeddingsService.computeQueryEmbedding(query);
    return { ok: true, embedding: vec };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
});

ipcMain.handle('search:hybrid', async (_, query, emails, embeddings) => {
  try {
    const results = await searchService.hybridSearch(query, emails, embeddings);
    return { ok: true, emails: results };
  } catch (err) {
    return { ok: false, error: err.message || String(err), emails: [] };
  }
});

ipcMain.handle('embeddings:promptCluster', async (_, prompt, embeddings, emailIds, threshold = 0.5) => {
  try {
    console.log('[butter-mail] embeddings:promptCluster starting. Prompt:', prompt, 'Emails:', emailIds ? emailIds.length : 0, 'Threshold:', threshold);
    const ids = await searchService.emailsBySimilarityToPrompt(prompt, embeddings, emailIds, threshold);
    console.log('[butter-mail] embeddings:promptCluster finished. Matched emails:', ids ? ids.length : 0);
    return { ok: true, emailIds: ids };
  } catch (err) {
    console.error('[butter-mail] embeddings:promptCluster failed:', err);
    return { ok: false, error: err.message || String(err), emailIds: [] };
  }
});

ipcMain.handle('embeddings:promptClusterScored', async (_, prompt, embeddings, emailIds) => {
  try {
    console.log('[butter-mail] embeddings:promptClusterScored starting. Prompt:', prompt, 'Emails:', emailIds ? emailIds.length : 0);
    const scored = await searchService.emailsBySimilarityToPromptScored(prompt, embeddings, emailIds);
    console.log('[butter-mail] embeddings:promptClusterScored finished. Results:', scored ? scored.length : 0);
    return { ok: true, scored };
  } catch (err) {
    console.error('[butter-mail] embeddings:promptClusterScored failed:', err);
    return { ok: false, error: err.message || String(err), scored: [] };
  }
});

ipcMain.handle('smtp:send', async (_, { to, subject, text, html, replyToMessageId }) => {
  const config = loadConfig();
  const smtpOpts = getSmtpOptions(config);
  if (!smtpOpts) {
    return { ok: false, error: 'SMTP not configured. Set IMAP settings (same account is used for sending).' };
  }
  if (!to || !to.trim()) {
    return { ok: false, error: 'Recipient (To) is required.' };
  }
  try {
    const transporter = nodemailer.createTransport({
      host: smtpOpts.host,
      port: smtpOpts.port,
      secure: smtpOpts.port === 465,
      auth: { user: smtpOpts.user, pass: smtpOpts.pass }
    });
    const mailOptions = {
      from: smtpOpts.user,
      to: to.trim(),
      subject: (subject || '').trim() || '(no subject)',
      text: (text || '').trim() || '',
      html: (html || '').trim() || undefined
    };
    if (replyToMessageId) mailOptions.references = replyToMessageId;
    if (replyToMessageId) mailOptions.inReplyTo = replyToMessageId;
    await transporter.sendMail(mailOptions);
    return { ok: true };
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
