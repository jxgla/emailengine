'use strict';

const Boom = require('@hapi/boom');
const { fetch: fetchCmd } = require('undici');
const { prepareUrl, retryAgent } = require('./tools');

const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
const OUTLOOK_SCOPE_SETS = {
    global: {
        api: ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send', 'offline_access', 'https://graph.microsoft.com/User.Read'],
        imap: ['https://outlook.office.com/IMAP.AccessAsUser.All', 'offline_access']
    },
    'gcc-high': {
        api: ['https://graph.microsoft.us/Mail.ReadWrite', 'https://graph.microsoft.us/Mail.Send', 'offline_access', 'https://graph.microsoft.us/User.Read'],
        imap: ['https://outlook.office365.us/IMAP.AccessAsUser.All', 'offline_access']
    },
    dod: {
        api: [
            'https://dod-graph.microsoft.us/Mail.ReadWrite',
            'https://dod-graph.microsoft.us/Mail.Send',
            'offline_access',
            'https://dod-graph.microsoft.us/User.Read'
        ],
        imap: ['https://outlook.office365.us/IMAP.AccessAsUser.All', 'offline_access']
    },
    china: {
        api: [
            'https://microsoftgraph.chinacloudapi.cn/Mail.ReadWrite',
            'https://microsoftgraph.chinacloudapi.cn/Mail.Send',
            'offline_access',
            'https://microsoftgraph.chinacloudapi.cn/User.Read'
        ],
        imap: ['https://partner.outlook.cn/IMAP.AccessAsUser.All', 'offline_access']
    }
};
const OUTLOOK_IMAP_SETTINGS = {
    global: {
        host: 'outlook.live.com',
        port: 993,
        secure: true
    },
    'gcc-high': {
        host: 'outlook.office365.us',
        port: 993,
        secure: true
    },
    dod: {
        host: 'outlook.office365.us',
        port: 993,
        secure: true
    },
    china: {
        host: 'partner.outlook.cn',
        port: 993,
        secure: true
    }
};

function sanitizeValue(value, maxLength) {
    let text = (value || '').toString();
    text = text.replace(/\r/g, '').replace(/\n/g, '').replace(/\t/g, '').trim();

    if (typeof maxLength === 'number' && maxLength > 0 && text.length > maxLength) {
        text = text.slice(0, maxLength);
    }

    return text;
}

function normalizeAccountId(email, accountIdPrefix) {
    const normalizedEmail = sanitizeValue(email, 320).toLowerCase();
    return `${sanitizeValue(accountIdPrefix, 128)}${normalizedEmail}`;
}

function parseRemoteOutlookAccountLines(input, options = {}) {
    const accountIdPrefix = options.accountIdPrefix || '';
    const entries = [];
    const errors = [];

    for (let [index, rawLine] of String(input || '')
        .split(/\r?\n/)
        .entries()) {
        const lineNumber = index + 1;
        const line = (rawLine || '').trim();

        if (!line || line.startsWith('#')) {
            continue;
        }

        const parts = line.split('----');
        if (parts.length < 4) {
            errors.push({
                line: lineNumber,
                error: 'Invalid format, expected email----password----client_id----refresh_token'
            });
            continue;
        }

        const email = sanitizeValue(parts[0], 320).toLowerCase();
        const password = sanitizeValue(parts[1], 512);
        const clientId = sanitizeValue(parts[2], 256);
        const refreshToken = sanitizeValue(parts.slice(3).join('----'), 4 * 4096);

        if (!email || !EMAIL_RE.test(email)) {
            errors.push({
                line: lineNumber,
                email,
                error: 'Invalid email address'
            });
            continue;
        }

        if (!clientId || !refreshToken) {
            errors.push({
                line: lineNumber,
                email,
                error: 'Client ID and refresh token are required'
            });
            continue;
        }

        entries.push({
            line: lineNumber,
            account: normalizeAccountId(email, accountIdPrefix),
            email,
            password,
            clientId,
            refreshToken
        });
    }

    return {
        entries,
        errors,
        summary: {
            total: entries.length + errors.length,
            parsed: entries.length,
            failed: errors.length
        }
    };
}

function buildRemoteManagedOutlookAccount(entry, options = {}) {
    const defaults = options.defaults || {};
    const normalizedCloud = sanitizeValue(options.cloud || 'global', 32) || 'global';
    const normalizedBaseScopes = sanitizeValue(options.baseScopes || 'api', 32) || 'api';

    const accountData = {
        account: entry.account,
        name: defaults.name || entry.email,
        email: entry.email,
        logs: !!defaults.logs
    };

    if (normalizedBaseScopes === 'imap') {
        const imapSettings = (OUTLOOK_IMAP_SETTINGS[normalizedCloud] || OUTLOOK_IMAP_SETTINGS.global);
        accountData.imap = Object.assign(
            {
                useAuthServer: true
            },
            imapSettings
        );
    } else {
        accountData.oauth2 = {
            provider: options.app,
            auth: {
                user: entry.email
            },
            useAuthServer: true
        };
    }

    for (let key of ['notifyFrom', 'syncFrom', 'path', 'proxy', 'smtpEhloName', 'webhooks', 'locale', 'tz']) {
        if (typeof defaults[key] !== 'undefined') {
            accountData[key] = defaults[key];
        }
    }

    return accountData;
}

function buildBrokerImportEntry(entry, options = {}) {
    return {
        account: entry.account,
        email: entry.email,
        user: entry.email,
        clientId: entry.clientId,
        refreshToken: entry.refreshToken,
        tenant: sanitizeValue(options.tenant || 'common', 256),
        proxyUrl: sanitizeValue(options.proxyUrl || '', 1024),
        cloud: sanitizeValue(options.cloud || 'global', 32) || 'global',
        baseScopes: sanitizeValue(options.baseScopes || 'api', 32) || 'api'
    };
}

function serializeRemoteOutlookAccountEntries(entries) {
    return []
        .concat(entries || [])
        .map(entry => [entry.email || '', entry.password || '', entry.clientId || '', entry.refreshToken || ''].join('----'))
        .join('\n');
}

function getOutlookBaseScopes(baseScopes, cloud) {
    const normalizedCloud = sanitizeValue(cloud || 'global', 32) || 'global';
    const normalizedBaseScopes = sanitizeValue(baseScopes || 'api', 32) || 'api';
    return ((OUTLOOK_SCOPE_SETS[normalizedCloud] || OUTLOOK_SCOPE_SETS.global)[normalizedBaseScopes] || OUTLOOK_SCOPE_SETS.global.api).slice();
}

async function importRemoteOutlookAccountsToAuthServer(authServer, entries, options = {}) {
    if (!entries.length) {
        return {
            summary: {
                total: 0,
                parsed: 0,
                created: 0,
                updated: 0,
                failed: 0
            },
            errors: []
        };
    }

    const parsedAuthServer = new URL(authServer);
    const headers = {
        'Content-Type': 'application/json'
    };

    if (parsedAuthServer.username || parsedAuthServer.password) {
        headers.Authorization = `Basic ${Buffer.from(
            `${decodeURIComponent(parsedAuthServer.username || '')}:${decodeURIComponent(parsedAuthServer.password || '')}`
        ).toString('base64')}`;
        parsedAuthServer.username = '';
        parsedAuthServer.password = '';
    }

    parsedAuthServer.search = '';
    parsedAuthServer.hash = '';

    const importUrl = prepareUrl('v1/accounts/import', parsedAuthServer.toString());
    const response = await fetchCmd(importUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify({
            accounts: serializeRemoteOutlookAccountEntries(entries),
            accountIdPrefix: options.accountIdPrefix || '',
            tenant: options.tenant || 'common',
            proxyUrl: options.proxyUrl || '',
            cloud: options.cloud || 'global',
            baseScopes: options.baseScopes || 'api'
        }),
        dispatcher: retryAgent
    });

    const responseText = await response.text();

    let payload;
    try {
        payload = responseText ? JSON.parse(responseText) : {};
    } catch (err) {
        payload = {
            text: responseText
        };
    }

    if (!response.ok) {
        const error = Boom.badGateway(`Authentication server import failed with status ${response.status}`);
        error.output.payload.details = payload;
        throw error;
    }

    if (payload?.summary?.failed) {
        const error = Boom.badGateway(`Authentication server rejected ${payload.summary.failed} imported account(s)`);
        error.output.payload.details = payload;
        throw error;
    }

    return payload;
}

module.exports = {
    buildBrokerImportEntry,
    buildRemoteManagedOutlookAccount,
    getOutlookBaseScopes,
    importRemoteOutlookAccountsToAuthServer,
    normalizeAccountId,
    parseRemoteOutlookAccountLines,
    serializeRemoteOutlookAccountEntries
};
