// Express 4 doesn't catch rejected promises from async route handlers — wrap every one.
module.exports = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);
