import mongoose from 'mongoose';

export default async function connectDB() {
  const uri = process.env.MONGODB_URI;
  if (!uri) throw new Error('MONGODB_URI missing');

  // Atlas 기본값에 맞는 안전한 옵션들
  await mongoose.connect(uri, {
    autoIndex: true,
    serverSelectionTimeoutMS: 5000,
    family: 4,                 // IPv4 우선. 네트워크 이슈 회피용
    appName: 'fund-ledger'
  });
  console.log('MongoDB connected');
}
