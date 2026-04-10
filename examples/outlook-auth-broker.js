'use strict';

const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const Hapi = require('@hapi/hapi');
const hapiPino = require('hapi-pino');
const Joi = require('joi');
const { fetch, ProxyAgent } = require('undici');
const { buildBrokerImportEntry, getOutlookBaseScopes, parseRemoteOutlookAccountLines } = require('../lib/remote-outlook');

const STORE_PATH = process.env.OUTLOOK_AUTH_BROKER_STORE || path.join(process.cwd(), 'data', 'outlook-auth-broker.json');
const PORT = Number(process.env.OUTLOOK_AUTH_BROKER_PORT || 3081);
const HOST = process.env.OUTLOOK_AUTH_BROKER_HOST || '127.0.0.1';
const BASIC_AUTH_USERNAME = process.env.OUTLOOK_AUTH_BROKER_USERNAME || '';
const BASIC_AUTH_PASSWORD = process.env.OUTLOOK_AUTH_BROKER_PASSWORD || '';
const TOKEN_CACHE_TTL_BUFFER = 5 * 60 * 1000;

const tokenCache = new Map();
const proxyAgents = new Map();
const refreshLocks = new Map();

function safeEqual(left, right) {
    const leftBuffer = Buffer.from(String(left || ''));
    const rightBuffer = Buffer.from(String(right || ''));

    if (leftBuffer.length !== rightBuffer.length) {
        return false;
    }

    return crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function checkBasicAuth(request) {
    if (!BASIC_AUTH_USERNAME && !BASIC_AUTH_PASSWORD) {
        return true;
    }

    const header = request.headers.authorization || '';
    if (!header.startsWith('Basic ')) {
        return false;
    }

    let decoded;
    try {
        decoded = Buffer.from(header.slice(6), 'base64').toString('utf-8');
    } catch (err) {
        return false;
    }

    const splitterPos = decoded.indexOf(':');
    const username = splitterPos >= 0 ? decoded.slice(0, splitterPos) : decoded;
    const password = splitterPos >= 0 ? decoded.slice(splitterPos + 1) : '';

    return safeEqual(username, BASIC_AUTH_USERNAME) && safeEqual(password, BASIC_AUTH_PASSWORD);
}

async function readStore() {
    try {
        const content = await fs.readFile(STORE_PATH, 'utf-8');
        const parsed = JSON.parse(content);
        if (!parsed.accounts || typeof parsed.accounts !== 'object') {
            parsed.accounts = {};
        }
        return parsed;
    } catch (err) {
        if (err.code === 'ENOENT') {
            return {
                accounts: {}
            };
        }
        throw err;
    }
}

async function writeStore(store) {
    await fs.mkdir(path.dirname(STORE_PATH), { recursive: true });

    const tmpPath = `${STORE_PATH}.${process.pid}.tmp`;
    const payload = JSON.stringify(store, null, 2);
    await fs.writeFile(tmpPath, payload);
    await fs.rename(tmpPath, STORE_PATH);
}

function maskValue(value, visible = 6) {
    const text = String(value || '');
    if (text.length <= visible * 2) {
        return text ? `${text.slice(0, visible)}...` : '';
    }
    return `${text.slice(0, visible)}...${text.slice(-visible)}`;
}

function getProxyDispatcher(proxyUrl) {
    if (!proxyUrl) {
        return undefined;
    }

    if (!proxyAgents.has(proxyUrl)) {
        proxyAgents.set(proxyUrl, new ProxyAgent(proxyUrl));
    }

    return proxyAgents.get(proxyUrl);
}

function getCachedToken(cacheKey) {
    const cached = tokenCache.get(cacheKey);
    if (!cached || cached.expires <= Date.now()) {
        tokenCache.delete(cacheKey);
        return null;
    }
    return cached;
}

async function readJsonResponse(response) {
    const text = await response.text();
    if (!text) {
        return {};
    }

    try {
        return JSON.parse(text);
    } catch (err) {
        return {
            text
        };
    }
}

async function withRefreshLock(account, handler) {
    if (refreshLocks.has(account)) {
        return await refreshLocks.get(account);
    }

    const lockPromise = (async () => {
        try {
            return await handler();
        } finally {
            refreshLocks.delete(account);
        }
    })();

    refreshLocks.set(account, lockPromise);
    return await lockPromise;
}

async function refreshGraphAccessToken(entry) {
    const cacheKey = `${entry.account}:${entry.baseScopes || 'api'}`;
    const cached = getCachedToken(cacheKey);
    if (cached) {
        return {
            accessToken: cached.accessToken,
            expires: cached.expires,
            rotated: false
        };
    }

    const tenant = entry.tenant || 'common';
    const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams({
            client_id: entry.clientId,
            grant_type: 'refresh_token',
            refresh_token: entry.refreshToken,
            scope: getOutlookBaseScopes(entry.baseScopes || 'api', entry.cloud || 'global').join(' ')
        }).toString(),
        dispatcher: getProxyDispatcher(entry.proxyUrl)
    });

    const payload = await readJsonResponse(response);
    if (!response.ok || !payload.access_token) {
        const error = new Error(`Token refresh failed: ${response.status}`);
        error.statusCode = response.status;
        error.payload = payload;
        throw error;
    }

    const expires = Date.now() + Math.max(0, Number(payload.expires_in || 0) * 1000 - TOKEN_CACHE_TTL_BUFFER);
    tokenCache.set(cacheKey, {
        accessToken: payload.access_token,
        expires
    });

    const nextRefreshToken = typeof payload.refresh_token === 'string' && payload.refresh_token.trim() ? payload.refresh_token.trim() : null;
    return {
        accessToken: payload.access_token,
        expires,
        nextRefreshToken
    };
}

async function refreshAndPersistAccount(accountId) {
    return await withRefreshLock(accountId, async () => {
        const store = await readStore();
        const entry = store.accounts[accountId];
        if (!entry) {
            const error = new Error('Account was not found');
            error.statusCode = 404;
            throw error;
        }

        const tokenResult = await refreshGraphAccessToken(entry);
        if (tokenResult.nextRefreshToken && tokenResult.nextRefreshToken !== entry.refreshToken) {
            store.accounts[accountId] = Object.assign({}, entry, {
                refreshToken: tokenResult.nextRefreshToken,
                updatedAt: new Date().toISOString(),
                rotatedAt: new Date().toISOString()
            });
            await writeStore(store);
        }

        return {
            entry: store.accounts[accountId] || entry,
            tokenResult
        };
    });
}

async function init() {
    const server = Hapi.server({
        port: PORT,
        host: HOST,
        routes: {
            cors: true
        }
    });

    await server.register({
        plugin: hapiPino,
        options: {
            level: process.env.LOG_LEVEL || 'info'
        }
    });

    server.ext('onRequest', (request, h) => {
        if (!checkBasicAuth(request)) {
            return h.response({ error: 'Unauthorized' }).code(401).header('WWW-Authenticate', 'Basic realm="Outlook Auth Broker"').takeover();
        }
        return h.continue;
    });

    server.route({
        method: 'GET',
        path: '/healthz',
        handler() {
            return {
                status: 'ok'
            };
        }
    });

    server.route({
        method: 'GET',
        path: '/',
        options: {
            validate: {
                query: Joi.object({
                    account: Joi.string().max(320).required(),
                    proto: Joi.string().valid('api').required()
                })
            }
        },
        async handler(request) {
            const { account } = request.query;
            const { entry, tokenResult } = await refreshAndPersistAccount(account);

            request.logger.info({
                msg: 'Resolved auth server credentials',
                account,
                proto: 'api',
                rotated: !!tokenResult.nextRefreshToken
            });

            return {
                user: entry.user || entry.email,
                accessToken: tokenResult.accessToken
            };
        }
    });

    server.route({
        method: 'GET',
        path: '/v1/accounts',
        async handler() {
            const store = await readStore();
            return {
                accounts: Object.values(store.accounts).map(entry => ({
                    account: entry.account,
                    email: entry.email,
                    user: entry.user,
                    tenant: entry.tenant || 'common',
                    cloud: entry.cloud || 'global',
                    baseScopes: entry.baseScopes || 'api',
                    proxyUrl: entry.proxyUrl || '',
                    clientId: maskValue(entry.clientId),
                    hasRefreshToken: !!entry.refreshToken,
                    updatedAt: entry.updatedAt || null,
                    rotatedAt: entry.rotatedAt || null
                }))
            };
        }
    });

    server.route({
        method: 'POST',
        path: '/v1/accounts/import',
        options: {
            validate: {
                payload: Joi.object({
                    accounts: Joi.string()
                        .max(1024 * 1024)
                        .required(),
                    accountIdPrefix: Joi.string().empty('').trim().max(128).default(''),
                    tenant: Joi.string().empty('').trim().max(256).default('common'),
                    proxyUrl: Joi.string()
                        .empty('')
                        .trim()
                        .uri({ scheme: ['http', 'https', 'socks5', 'socks4'] })
                        .allow('')
                        .default(''),
                    cloud: Joi.string().valid('global', 'gcc-high', 'dod', 'china').default('global'),
                    baseScopes: Joi.string().valid('api').default('api')
                })
            }
        },
        async handler(request) {
            const { entries, errors, summary } = parseRemoteOutlookAccountLines(request.payload.accounts, {
                accountIdPrefix: request.payload.accountIdPrefix
            });

            const store = await readStore();
            let created = 0;
            let updated = 0;

            for (let entry of entries) {
                const brokerEntry = buildBrokerImportEntry(entry, {
                    tenant: request.payload.tenant,
                    proxyUrl: request.payload.proxyUrl,
                    cloud: request.payload.cloud,
                    baseScopes: request.payload.baseScopes
                });

                if (store.accounts[brokerEntry.account]) {
                    updated++;
                } else {
                    created++;
                }

                store.accounts[brokerEntry.account] = Object.assign({}, store.accounts[brokerEntry.account] || {}, brokerEntry, {
                    updatedAt: new Date().toISOString()
                });
            }

            await writeStore(store);

            return {
                summary: {
                    total: summary.total,
                    parsed: summary.parsed,
                    created,
                    updated,
                    failed: errors.length
                },
                errors
            };
        }
    });

    server.route({
        method: 'POST',
        path: '/v1/accounts/{account}/refresh',
        options: {
            validate: {
                params: Joi.object({
                    account: Joi.string().max(320).required()
                })
            }
        },
        async handler(request) {
            const { account } = request.params;
            const { entry, tokenResult } = await refreshAndPersistAccount(account);

            return {
                account: entry.account,
                email: entry.email,
                user: entry.user,
                expiresAt: new Date(tokenResult.expires).toISOString(),
                rotated: !!tokenResult.nextRefreshToken
            };
        }
    });

    await server.start();
    console.log(`Outlook auth broker running at ${server.info.uri}`);
    console.log(`Store path: ${STORE_PATH}`);
}

init().catch(err => {
    console.error(err);
    process.exit(1);
});
