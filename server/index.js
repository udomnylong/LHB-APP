require('dotenv').config();
const express = require('express');
const cors = require('./middleware/cors');

const app = express();
app.use(cors);
app.use(express.json({ limit: '2mb' }));

app.get('/health', (req, res) => res.json({ status: 'ok' }));

app.use('/api/auth', require('./routes/auth'));
app.use('/api/users', require('./routes/users'));
app.use('/api/staff', require('./routes/staff'));
app.use('/api', require('./routes/attendance')); // exposes /api/checkins, /api/checkouts, /api/attendance
app.use('/api/telegram', require('./routes/telegram'));
app.use('/api/admin', require('./routes/admin'));

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ status: 'error', msg: err.message || 'Internal error' });
});

const port = process.env.PORT || 8080;
app.listen(port, () => console.log(`lhb-hr-api listening on :${port}`));
