const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

async function get_access_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();

    try {
        const data = JSON.parse(responseText);
        return data.access_token;
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

async function graph_api(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
            'scope': 'https://graph.microsoft.com/.default'
        }).toString()
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP error! status: ${response.status}, response: ${errorText}`);
    }

    const responseText = await response.text();

    try {
        const data = JSON.parse(responseText);

        if (data.scope && data.scope.indexOf('https://graph.microsoft.com/Mail.Read') != -1) {
            return {
                access_token: data.access_token,
                status: true
            }
        }

        return {
            access_token: data.access_token,
            status: false
        }
    } catch (parseError) {
        throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
    }
}

async function get_emails(access_token, mailbox) {
    if (!access_token) {
        throw new Error("Failed to obtain access token");
    }

    const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=100`, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            "Authorization": `Bearer ${access_token}`
        },
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error (${response.status}): ${errorText}`);
    }

    const responseData = await response.json();
    const emails = responseData.value || [];

    return emails.map(item => {
        return {
            send: item.from?.emailAddress?.address || '未知发件人',
            subject: item.subject || '(无主题)',
            text: item.bodyPreview || '',
            html: item.body?.content || item.bodyPreview || '',
            date: item.createdDateTime || '',
        };
    });
}

module.exports = async (req, res) => {

    const { password } = req.method === 'GET' ? req.query : req.body;

    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({
            error: '密码验证失败'
        });
    }

    let { refresh_token, client_id, email, mailbox } = req.method === 'GET' ? req.query : req.body;

    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
    }

    const ALLOWED_MAILBOXES = ['INBOX', 'Junk'];
    if (!ALLOWED_MAILBOXES.includes(mailbox)) {
        return res.status(400).json({ error: 'Invalid mailbox. Allowed: INBOX, Junk' });
    }

    try {
        console.log("判断是否graph_api");
        const graph_api_result = await graph_api(refresh_token, client_id);

        if (graph_api_result.status) {
            console.log("是graph_api");

            let targetMailbox = 'inbox';
            if (mailbox === 'Junk') {
                targetMailbox = 'junkemail';
            }

            const result = await get_emails(graph_api_result.access_token, targetMailbox);
            return res.status(200).json(result);
        }

        const access_token = await get_access_token(refresh_token, client_id);
        const authString = generateAuthString(email, access_token);

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: {
                rejectUnauthorized: false
            }
        });

        const emailList = [];
        imap.once("ready", async () => {
            try {
                await new Promise((resolve, reject) => {
                    imap.openBox(mailbox, true, (err, box) => {
                        if (err) return reject(err);
                        resolve(box);
                    });
                });

                const results = await new Promise((resolve, reject) => {
                    imap.search(["ALL"], (err, results) => {
                        if (err) return reject(err);
                        resolve(results);
                    });
                });

                if (!results || results.length === 0) {
                    imap.end();
                    return;
                }

                const f = imap.fetch(results, { bodies: "" });

                f.on("message", (msg, seqno) => {
                    msg.on("body", (stream, info) => {
                        simpleParser(stream, (err, mail) => {
                            if (err) return;
                            const data = {
                                send: mail.from?.text || '未知发件人',
                                subject: mail.subject || '(无主题)',
                                text: mail.text || '',
                                html: mail.html || '',
                                date: mail.date || '',
                            };

                            emailList.push(data);
                        });
                    });
                });

                f.once("end", () => {
                    imap.end();
                });
            } catch (err) {
                imap.end();
                if (!res.headersSent) {
                    res.status(500).json({ error: err.message });
                }
            }
        });

        imap.once('error', (err) => {
            console.error('IMAP error:', err);
            if (!res.headersSent) {
                res.status(500).json({ error: err.message });
            }
        });

        imap.once('end', () => {
            if (!res.headersSent) {
                res.status(200).json(emailList);
            }
            console.log('IMAP connection ended');
        });

        imap.connect();

    } catch (error) {
        console.error('Error:', error);
        if (!res.headersSent) {
            res.status(500).json({ error: error.message });
        }
    }
};
