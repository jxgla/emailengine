'use strict';

const test = require('node:test');
const assert = require('node:assert').strict;

const {
    DEFAULT_VERIFICATION_CODE_REGEX,
    createVerificationError,
    extractVerificationCodeFromMessage,
    pickNewestMessageCandidate,
    resolveVerificationMailboxes,
    summarizeCandidate
} = require('../lib/verification-code');

test('resolve inbox and junk mailboxes using special use flags', async () => {
    const mailboxes = [
        { path: 'Archive', name: 'Archive' },
        { path: 'Junk', name: 'Junk Email', specialUse: '\\Junk' },
        { path: 'INBOX', name: 'Inbox', specialUse: '\\Inbox' }
    ];

    const resolved = resolveVerificationMailboxes(mailboxes);

    assert.strictEqual(resolved.inbox.path, 'INBOX');
    assert.strictEqual(resolved.junk.path, 'Junk');
});

test('resolve junk mailbox using localized names when special use is missing', async () => {
    const mailboxes = [
        { path: 'INBOX', name: 'Inbox' },
        { path: '垃圾邮件', name: '垃圾邮件' }
    ];

    const resolved = resolveVerificationMailboxes(mailboxes);

    assert.strictEqual(resolved.inbox.path, 'INBOX');
    assert.strictEqual(resolved.junk.path, '垃圾邮件');
});

test('pick newest candidate between inbox and junk', async () => {
    const candidates = [
        {
            slot: 'inbox',
            messageId: 'msg-1',
            messageTimestamp: Date.parse('2026-04-10T09:59:00.000Z')
        },
        {
            slot: 'junk',
            messageId: 'msg-2',
            messageTimestamp: Date.parse('2026-04-10T10:00:00.000Z')
        }
    ];

    const selected = pickNewestMessageCandidate(candidates);

    assert.strictEqual(selected.slot, 'junk');
    assert.strictEqual(selected.messageId, 'msg-2');
});

test('summarize candidate normalizes timestamp and sender address', async () => {
    const candidate = summarizeCandidate(
        'inbox',
        { path: 'INBOX', name: 'Inbox', specialUse: '\\Inbox' },
        {
            id: 'msg-1',
            date: '2026-04-10T10:05:00.000Z',
            subject: 'Your code',
            from: { address: 'sender@example.com' }
        }
    );

    assert.strictEqual(candidate.slot, 'inbox');
    assert.strictEqual(candidate.path, 'INBOX');
    assert.strictEqual(candidate.messageFrom, 'sender@example.com');
    assert.ok(candidate.messageTimestamp > 0);
});

test('extract verification code prefers plain text, then html, then subject', async () => {
    const extracted = extractVerificationCodeFromMessage({
        subject: 'Backup subject 555555',
        text: {
            plain: 'Use verification code 654321 to continue',
            html: '<p>Use verification code <strong>123456</strong> to continue</p>'
        }
    });

    assert.strictEqual(extracted.code, '654321');
    assert.strictEqual(extracted.matchSource, 'plain');
    assert.strictEqual(extracted.regex, DEFAULT_VERIFICATION_CODE_REGEX);
});

test('extract verification code falls back to html text when plain text is empty', async () => {
    const extracted = extractVerificationCodeFromMessage({
        subject: 'No code here',
        text: {
            html: '<div>Your verification code is <b>778899</b></div>'
        }
    });

    assert.strictEqual(extracted.code, '778899');
    assert.strictEqual(extracted.matchSource, 'html');
});

test('extract verification code supports custom capture-group regex', async () => {
    const extracted = extractVerificationCodeFromMessage(
        {
            subject: 'Latest message',
            text: {
                plain: 'Token=ABC-7788-END'
            }
        },
        {
            codeRegex: 'Token=([A-Z]{3}-\\d{4}-END)'
        }
    );

    assert.strictEqual(extracted.code, 'ABC-7788-END');
    assert.strictEqual(extracted.matchSource, 'plain');
    assert.strictEqual(extracted.regex, 'Token=([A-Z]{3}-\\d{4}-END)');
});

test('invalid regex is reported with a stable error code', async () => {
    assert.throws(
        () =>
            extractVerificationCodeFromMessage(
                {
                    text: {
                        plain: '123456'
                    }
                },
                {
                    codeRegex: '['
                }
            ),
        err => {
            assert.strictEqual(err.code, 'VERIFICATION_CODE_REGEX_INVALID');
            assert.strictEqual(err.statusCode, 400);
            return true;
        }
    );
});

test('createVerificationError attaches code and status', async () => {
    const error = createVerificationError('TEST_ERROR', 'broken', 409);

    assert.strictEqual(error.code, 'TEST_ERROR');
    assert.strictEqual(error.statusCode, 409);
    assert.strictEqual(error.message, 'broken');
});
