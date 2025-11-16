// server.js
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { MongoClient, ObjectId } from 'mongodb';
import engine from 'ejs-mate';
import entriesRouter from './src/routes/entries.js';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// EJS + ejs-mate(layout)
app.engine('ejs', engine);
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Body parser
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Static
app.use(express.static(path.join(__dirname, 'public')));

// MongoDB 연결
const MONGO_URI =
  process.env.MONGO_URI ||
  'mongodb+srv://krogy123:rlarudfhr1262@cluster0.qnjcx2e.mongodb.net/?retryWrites=true&w=majority&tls=true&tlsAllowInvalidCertificates=true';

const client = new MongoClient(MONGO_URI, { serverSelectionTimeoutMS: 15000 });

async function bootstrap() {
  await client.connect();
  const dbName = process.env.MONGO_DB || 'fund_ledger';
  const db = client.db(dbName);
  app.locals.db = db;
  app.locals.ObjectId = ObjectId;

  console.log('MongoDB connected');

  // 라우터
  app.use('/', entriesRouter);

  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`http://localhost:${PORT}`);
  });
}

bootstrap().catch((err) => {
  console.error('Mongo bootstrap error:', err);
  process.exit(1);
});
