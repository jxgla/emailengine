'use strict';

const { convert: htmlToText } = require('html-to-text');

const DEFAULT_VERIFICATION_CODE_REGEX = String.raw`\b\d{4,8}\b`;
const DEFAULT_REGEX_FLAGS = 'i';

const INBOX_HINTS = ['inbox', '收件'];
const JUNK_HINTS = ['junk', 'junk email', 'spam', 'bulk', '垃圾'];

function createVerificationError(code, message, statusCode) {
    const error = new Error(message);
    error.code = code;
    error.statusCode = statusCode;
    return error;
}

function getMailboxSpecialUses(mailbox) {
    if (!mailbox) {
        return [];
    }

    if (Array.isArray(mailbox.specialUse)) {
        return mailbox.specialUse.map(entry => String(entry || '').trim()).filter(Boolean);
    }

    if (mailbox.specialUse) {
        return [String(mailbox.specialUse).trim()].filter(Boolean);
    }

    return [];
}

function mailboxSearchValues(mailbox) {
    return [mailbox && mailbox.path, mailbox && mailbox.name]
        .concat(getMailboxSpecialUses(mailbox))
        .map(value => String(value || '').trim().toLowerCase())
        .filter(Boolean);
}

function includesHint(values, hints) {
    return values.some(value => hints.some(hint => value.includes(hint)));
}

function resolveVerificationMailboxes(mailboxes) {
    const list = Array.isArray(mailboxes) ? mailboxes.filter(Boolean) : [];

    const inbox =
        list.find(mailbox => getMailboxSpecialUses(mailbox).includes('\\Inbox')) ||
        list.find(mailbox => includesHint(mailboxSearchValues(mailbox), INBOX_HINTS)) ||
        null;

    const junk =
        list.find(mailbox => getMailboxSpecialUses(mailbox).includes('\\Junk')) ||
        list.find(mailbox => includesHint(mailboxSearchValues(mailbox), JUNK_HINTS)) ||
        null;

    return { inbox, junk };
}

function parseMessageTimestamp(...values) {
    for (const value of values) {
        if (!value) {
            continue;
        }

        const timestamp = Date.parse(value);
        if (Number.isFinite(timestamp)) {
            return timestamp;
        }
    }

    return 0;
}

function formatAddress(value) {
    if (!value) {
        return '';
    }

    if (typeof value === 'string') {
        return value;
    }

    if (Array.isArray(value)) {
        return value.map(formatAddress).filter(Boolean).join(', ');
    }

    if (value.address) {
        return String(value.address || '');
    }

    if (value.emailAddress && value.emailAddress.address) {
        return String(value.emailAddress.address || '');
    }

    return '';
}

function summarizeCandidate(slot, mailbox, message) {
    if (!mailbox || !message) {
        return null;
    }

    const timestamp = parseMessageTimestamp(
        message.date,
        message.created,
        message.createdAt,
        message.receivedAt,
        message.receivedDateTime,
        message.internalDate
    );

    return {
        slot,
        path: String(mailbox.path || ''),
        name: String(mailbox.name || mailbox.path || ''),
        specialUse: getMailboxSpecialUses(mailbox),
        messageId: String(message.id || ''),
        messageDate: String(message.date || message.created || message.createdAt || message.receivedAt || ''),
        messageTimestamp: timestamp,
        messageSubject: String(message.subject || ''),
        messageFrom: formatAddress(message.from)
    };
}

function pickNewestMessageCandidate(candidates) {
    return (Array.isArray(candidates) ? candidates : [])
        .filter(candidate => candidate && candidate.messageId)
        .sort((a, b) => {
            if (b.messageTimestamp !== a.messageTimestamp) {
                return b.messageTimestamp - a.messageTimestamp;
            }

            return String(a.slot || '').localeCompare(String(b.slot || ''));
        })[0] || null;
}

function htmlToPlainText(html) {
    const input = String(html || '').trim();
    if (!input) {
        return '';
    }

    try {
        return htmlToText(input, {
            wordwrap: false,
            selectors: [
                { selector: 'a', options: { ignoreHref: true } },
                { selector: 'img', format: 'skip' }
            ]
        }).trim();
    } catch (err) {
        return input.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    }
}

function firstNonEmpty(...values) {
    for (const value of values) {
        if (typeof value === 'string' && value.trim()) {
            return value.trim();
        }
    }

    return '';
}

function getVerificationSources(message) {
    const subject = String(message && message.subject ? message.subject : '').trim();
    const plain = firstNonEmpty(
        message && message.text && message.text.plain,
        message && message.text && message.text.text,
        message && message.textPlain,
        message && message.plain,
        typeof (message && message.text) === 'string' ? message.text : '',
        message && message.content
    );
    const html = firstNonEmpty(
        message && message.text && message.text.html,
        message && message.html,
        message && message.textHtml
    );
    const htmlText = htmlToPlainText(html);
    const combined = [plain, htmlText, subject].filter(Boolean).join('\n');

    return [
        { name: 'plain', value: plain },
        { name: 'html', value: htmlText },
        { name: 'subject', value: subject },
        { name: 'combined', value: combined }
    ];
}

function compileVerificationRegex(pattern, flags) {
    const normalizedPattern = String(pattern || DEFAULT_VERIFICATION_CODE_REGEX).trim() || DEFAULT_VERIFICATION_CODE_REGEX;
    const normalizedFlags = typeof flags === 'string' ? flags : DEFAULT_REGEX_FLAGS;

    try {
        return {
            pattern: normalizedPattern,
            flags: normalizedFlags,
            regex: new RegExp(normalizedPattern, normalizedFlags)
        };
    } catch (err) {
        throw createVerificationError('VERIFICATION_CODE_REGEX_INVALID', `Invalid verification regex: ${err.message}`, 400);
    }
}

function pickMatchedValue(match) {
    if (!match) {
        return '';
    }

    for (let index = 1; index < match.length; index++) {
        if (typeof match[index] === 'string' && match[index]) {
            return match[index];
        }
    }

    return match[0] || '';
}

function extractVerificationCodeFromMessage(message, options = {}) {
    const compiled = compileVerificationRegex(options.codeRegex, options.codeFlags);

    for (const source of getVerificationSources(message)) {
        if (!source.value) {
            continue;
        }

        compiled.regex.lastIndex = 0;
        const match = compiled.regex.exec(source.value);
        if (match) {
            return {
                code: pickMatchedValue(match).trim(),
                matchSource: source.name,
                regex: compiled.pattern,
                flags: compiled.flags
            };
        }
    }

    return {
        code: '',
        matchSource: '',
        regex: compiled.pattern,
        flags: compiled.flags
    };
}

module.exports = {
    DEFAULT_REGEX_FLAGS,
    DEFAULT_VERIFICATION_CODE_REGEX,
    createVerificationError,
    extractVerificationCodeFromMessage,
    pickNewestMessageCandidate,
    resolveVerificationMailboxes,
    summarizeCandidate
};
