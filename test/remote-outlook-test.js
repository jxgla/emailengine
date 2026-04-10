'use strict';

const test = require('node:test');
const assert = require('node:assert').strict;
const Joi = require('joi');

const dbPath = require.resolve('../lib/db');
const fakeRedis = {
    duplicate() {
        return this;
    },
    subscribe() {
        return Promise.resolve();
    },
    on() {
        return this;
    },
    defineCommand(name) {
        this[name] = async () => null;
    }
};

require.cache[dbPath] = {
    id: dbPath,
    filename: dbPath,
    loaded: true,
    exports: {
        redis: fakeRedis
    }
};

const { oauthCreateSchema, oauthUpdateSchema } = require('../lib/schemas');
const {
    buildBrokerImportEntry,
    buildRemoteManagedOutlookAccount,
    getOutlookBaseScopes,
    parseRemoteOutlookAccountLines,
    serializeRemoteOutlookAccountEntries
} = require('../lib/remote-outlook');

test('parse remote Outlook import lines', async () => {
    const { entries, errors, summary } = parseRemoteOutlookAccountLines(
        ['# comment', 'User.One@Outlook.com----pw----client-1----rt-1----tail', 'broken-line'].join('\n'),
        {
            accountIdPrefix: 'remote:'
        }
    );

    assert.strictEqual(summary.total, 2);
    assert.strictEqual(summary.parsed, 1);
    assert.strictEqual(summary.failed, 1);

    assert.strictEqual(entries.length, 1);
    assert.strictEqual(entries[0].account, 'remote:user.one@outlook.com');
    assert.strictEqual(entries[0].email, 'user.one@outlook.com');
    assert.strictEqual(entries[0].clientId, 'client-1');
    assert.strictEqual(entries[0].refreshToken, 'rt-1----tail');

    assert.strictEqual(errors.length, 1);
    assert.match(errors[0].error, /Invalid format/i);
});

test('build remote managed Outlook account payload', async () => {
    const accountData = buildRemoteManagedOutlookAccount(
        {
            account: 'remote:user.two@outlook.com',
            email: 'user.two@outlook.com'
        },
        {
            app: 'outlook-api-app',
            defaults: {
                logs: true,
                proxy: 'http://127.0.0.1:8080',
                path: ['*']
            }
        }
    );

    assert.deepStrictEqual(accountData.oauth2, {
        provider: 'outlook-api-app',
        auth: {
            user: 'user.two@outlook.com'
        },
        useAuthServer: true
    });
    assert.strictEqual(accountData.logs, true);
    assert.strictEqual(accountData.proxy, 'http://127.0.0.1:8080');
    assert.deepStrictEqual(accountData.path, ['*']);
});

test('build remote managed Outlook IMAP account payload', async () => {
    const accountData = buildRemoteManagedOutlookAccount(
        {
            account: 'remote:user.two@outlook.com',
            email: 'user.two@outlook.com'
        },
        {
            app: 'outlook-imap-app',
            baseScopes: 'imap',
            cloud: 'global',
            defaults: {
                logs: true,
                proxy: 'http://127.0.0.1:8080',
                path: ['INBOX']
            }
        }
    );

    assert.deepStrictEqual(accountData.imap, {
        useAuthServer: true,
        host: 'outlook.live.com',
        port: 993,
        secure: true
    });
    assert.ok(!accountData.oauth2);
    assert.strictEqual(accountData.logs, true);
    assert.strictEqual(accountData.proxy, 'http://127.0.0.1:8080');
    assert.deepStrictEqual(accountData.path, ['INBOX']);
});

test('build broker import entry strips account to broker essentials', async () => {
    const brokerEntry = buildBrokerImportEntry(
        {
            account: 'remote:user.three@outlook.com',
            email: 'user.three@outlook.com',
            clientId: 'client-3',
            refreshToken: 'rt-3'
        },
        {
            tenant: 'consumers',
            proxyUrl: 'http://127.0.0.1:18080',
            cloud: 'global',
            baseScopes: 'api'
        }
    );

    assert.deepStrictEqual(brokerEntry, {
        account: 'remote:user.three@outlook.com',
        email: 'user.three@outlook.com',
        user: 'user.three@outlook.com',
        clientId: 'client-3',
        refreshToken: 'rt-3',
        tenant: 'consumers',
        proxyUrl: 'http://127.0.0.1:18080',
        cloud: 'global',
        baseScopes: 'api'
    });
});

test('serialize remote Outlook entries for broker sync', async () => {
    const text = serializeRemoteOutlookAccountEntries([
        {
            email: 'user.one@outlook.com',
            password: 'pw',
            clientId: 'client-1',
            refreshToken: 'rt-1'
        },
        {
            email: 'user.two@outlook.com',
            password: '',
            clientId: 'client-2',
            refreshToken: 'rt-2----tail'
        }
    ]);

    assert.strictEqual(text, 'user.one@outlook.com----pw----client-1----rt-1\nuser.two@outlook.com--------client-2----rt-2----tail');
});

test('baseScopes=api uses delegated Graph scopes instead of .default', async () => {
    assert.deepStrictEqual(getOutlookBaseScopes('api', 'global'), [
        'https://graph.microsoft.com/Mail.ReadWrite',
        'https://graph.microsoft.com/Mail.Send',
        'offline_access',
        'https://graph.microsoft.com/User.Read'
    ]);
});

test('baseScopes=imap uses minimal IMAP delegated scopes', async () => {
    assert.deepStrictEqual(getOutlookBaseScopes('imap', 'global'), [
        'https://outlook.office.com/IMAP.AccessAsUser.All',
        'offline_access'
    ]);
});

test('outlook api OAuth app can omit secret and redirect when broker manages tokens', async () => {
    const schema = Joi.object(oauthCreateSchema);
    const { error, value } = schema.validate({
        name: 'Remote Outlook API App',
        provider: 'outlook',
        enabled: true,
        clientId: 'public-client-id',
        clientSecret: '',
        baseScopes: 'api',
        authority: 'consumers',
        cloud: 'global',
        redirectUrl: ''
    });

    assert.ifError(error);
    assert.strictEqual(value.baseScopes, 'api');
    assert.strictEqual(value.provider, 'outlook');
});

test('outlook IMAP OAuth app still requires secret and redirect URL', async () => {
    const schema = Joi.object(oauthCreateSchema);
    const { error } = schema.validate({
        name: 'Remote Outlook IMAP App',
        provider: 'outlook',
        enabled: true,
        clientId: 'public-client-id',
        clientSecret: '',
        baseScopes: 'imap',
        authority: 'consumers',
        cloud: 'global',
        redirectUrl: ''
    });

    assert.ok(error);
    assert.match(error.message, /clientsecret|client secret|redirect/i);
});

test('outlook api OAuth app update can omit redirect when broker manages tokens', async () => {
    const schema = Joi.object(oauthUpdateSchema);
    const { error, value } = schema.validate({
        app: 'outlook-api-app',
        provider: 'outlook',
        name: 'Remote Outlook API App',
        baseScopes: 'api',
        clientId: 'public-client-id',
        authority: 'consumers',
        cloud: 'global',
        redirectUrl: ''
    });

    assert.ifError(error);
    assert.strictEqual(value.baseScopes, 'api');
});
