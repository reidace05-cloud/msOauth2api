const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

async function fetchWithRetry(url, options, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            const controller = new AbortController();
            const timeout = setTimeout(() => controller.abort(), 15000);
            const res = await fetch(url, {
                ...options,
                signal: controller.signal,
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    ...(options.headers || {})
                }
            });
            clearTimeout(timeout);
            return res;
        } catch (err) {
            if (i === retries - 1) throw err;
            await new Promise(r => setTimeout(r, 800));
        }
    }
}

async function get_access_token(refresh_token, client_id) {
    const response = await fetchWithRetry('https://microsoftonline.com', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Token error (${response.status}): ${errorText}`);
    }

    const data = await response.json();
    return data.access_token;
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

async function graph_api(refresh_token, client_id) {
    const response = await fetchWithRetry('https://microsoftonline.com', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'scope': 'https://microsoft.com'
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph token error (${response.status}): ${errorText}`);
    }

    const data = await response.json();
    const hasGraphScope = data.scope && data.scope.indexOf('https://microsoft.com') !== -1;

    return {
        access_token: data.access_token,
        status: hasGraphScope
    };
}

async function get_emails(access_token, mailbox) {
    if (!access_token) {
        throw new Error("Failed to obtain access token");
    }

    const response = await fetchWithRetry(`https://microsoft.com{mailbox}/messages?$top=50`, {
        method: 'GET',
        headers: {
            'Accept': 'application/json',
            'Authorization': `Bearer ${access_token}`
        }
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error (${response.status}): ${errorText}`);
    }

    const responseData = await response.json();
    const emails = responseData.value || [];

    return emails.map(item => ({
        send: item.from?.emailAddress?.address || '未知发件人',
        subject: item.subject || '(无主题)',
        text: item.bodyPreview || '',
        html: item.body?.content || item.bodyPreview || '',
        date: item.createdDateTime || '',
    }));
}

module.exports = async (req, res) => {
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({ error: '密码验证失败' });
    }

    let { refresh_token, client_id, email, mailbox } = req.method === 'GET' ? req.query : req.body;

    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters' });
    }

    const ALLOWED_MAILBOXES = ['INBOX', 'Junk'];
    if (!ALLOWED_MAILBOXES.includes(mailbox)) {
        return res.status(400).json({ error: 'Invalid mailbox' });
    }

    try {
        const graph_api_result = await graph_api(refresh_token, client_id);

        if (graph_api_result.status) {
            const targetMailbox = mailbox === 'Junk' ? 'junkemail' : 'inbox';
            try {
                const result = await get_emails(graph_api_result.access_token, targetMailbox);
                return res.status(200).json(result);
            } catch (graphErr) {
                console.log("Graph 切换 IMAP:", graphErr.message);
            }
        }

        const access_token = await get_access_token(refresh_token, client_id);
        const authString = generateAuthString(email, access_token);

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: '://office365.com',
            port: 993,
            tls: true,
            connTimeout: 15000,
            authTimeout: 15000,
            tlsOptions: { rejectUnauthorized: false }
        });

        const emailList = [];
        imap.once("ready", async () => {
            try {
                await new Promise((resolve, reject) => {
                    imap.openBox(mailbox, true, (err, box) => err ? reject(err) : resolve(box));
                });

                const results = await new Promise((resolve, reject) => {
                    imap.search(["ALL"], (err, results) => err ? reject(err) : resolve(results));
                });

                if (!results || results.length === 0) {
                    imap.end();
                    return;
                }

                const recentResults = results.slice(-50);
                const f = imap.fetch(recentResults, { bodies: "" });

                f.on("message", (msg) => {
                    msg.on("body", (stream) => {
                        simpleParser(stream, (err, mail) => {
                            if (err) return;
                            emailList.push({
                                send: mail.from?.text || '未知发件人',
                                subject: mail.subject || '(无主题)',
                                text: mail.text || '',
                                html: mail.html || '',
                                date: mail.date || '',
                            });
                        });
                    });
                });

                f.once("end", () => imap.end());
            } catch (err) {
                imap.end();
                if (!res.headersSent) res.status(500).json({ error: err.message });
            }
        });

        imap.once('error', (err) => {
            if (!res.headersSent) res.status(500).json({ error: `IMAP error: ${err.message}` });
        });

        imap.once('end', () => {
            if (!res.headersSent) res.status(200).json(emailList);
        });

        imap.connect();

    } catch (error) {
        if (!res.headersSent) res.status(500).json({ error: error.message });
    }
};
