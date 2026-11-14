const { promisify } = require('node:util');
const {
  createHash,
  randomBytes,
  scrypt,
  timingSafeEqual,
} = require('node:crypto');
const { db } = require('./db');

const scryptAsync = promisify(scrypt);
const SESSION_COOKIE = 'household_session';
const sessionDurationSeconds = 8 * 60 * 60;
const failedLogins = new Map();

const sha256 = (value) => createHash('sha256').update(value).digest('hex');

async function hashPassword(password) {
  const salt = randomBytes(16).toString('hex');
  const hash = await scryptAsync(password, salt, 64);
  return `scrypt$${salt}$${hash.toString('hex')}`;
}

async function verifyPassword(password, encoded) {
  const [algorithm, salt, expectedHex] = String(encoded || '').split('$');
  if (algorithm !== 'scrypt' || !salt || !expectedHex) return false;
  const actual = await scryptAsync(password, salt, 64);
  const expected = Buffer.from(expectedHex, 'hex');
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}

function cookieToken(request) {
  const cookies = String(request.headers.cookie || '').split(';');
  const cookie = cookies.map((item) => item.trim()).find((item) => item.startsWith(`${SESSION_COOKIE}=`));
  if (!cookie) return null;
  try {
    return decodeURIComponent(cookie.slice(SESSION_COOKIE.length + 1));
  } catch {
    return null;
  }
}

function setSessionCookie(response, token) {
  const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
  response.setHeader(
    'Set-Cookie',
    `${SESSION_COOKIE}=${encodeURIComponent(token)}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${sessionDurationSeconds}${secure}`,
  );
}

function clearSessionCookie(response) {
  const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
  response.setHeader(
    'Set-Cookie',
    `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0${secure}`,
  );
}

function loginIsLimited(ip) {
  const attempts = failedLogins.get(ip);
  if (!attempts) return false;
  if (attempts.expiresAt <= Date.now()) {
    failedLogins.delete(ip);
    return false;
  }
  return attempts.count >= 8;
}

function noteFailedLogin(ip) {
  const current = failedLogins.get(ip);
  if (!current || current.expiresAt <= Date.now()) {
    failedLogins.set(ip, { count: 1, expiresAt: Date.now() + 15 * 60 * 1000 });
    return;
  }
  current.count += 1;
}

function clearFailedLogins(ip) {
  failedLogins.delete(ip);
}

async function resolveSession(request) {
  const token = cookieToken(request);
  if (!token || token.length > 200) return null;
  return db.prepare(`
    SELECT u.id, u.email, u.role, s.token_hash
    FROM user_sessions s
    JOIN users u ON u.id = s.user_id
    WHERE s.token_hash = ? AND s.expires_at > CURRENT_TIMESTAMP AND u.is_active = TRUE
  `).get(sha256(token));
}

async function createSession(userId) {
  const token = randomBytes(32).toString('base64url');
  await db.prepare(`
    INSERT INTO user_sessions (token_hash, user_id, expires_at)
    VALUES (?, ?, CURRENT_TIMESTAMP + INTERVAL '8 hours')
  `).run(sha256(token), userId);
  return token;
}

async function deleteSession(request) {
  const token = cookieToken(request);
  if (token) {
    await db.prepare('DELETE FROM user_sessions WHERE token_hash = ?').run(sha256(token));
  }
}

function safeRedirect(value) {
  const target = String(value || '/');
  return target.startsWith('/') && !target.startsWith('//') && !target.includes('\\')
    ? target
    : '/';
}

module.exports = {
  clearFailedLogins,
  clearSessionCookie,
  createSession,
  deleteSession,
  hashPassword,
  loginIsLimited,
  noteFailedLogin,
  resolveSession,
  safeRedirect,
  SESSION_COOKIE,
  setSessionCookie,
  verifyPassword,
};
