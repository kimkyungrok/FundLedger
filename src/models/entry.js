import mongoose from 'mongoose';

const EntrySchema = new mongoose.Schema({
  date: { type: Date, required: true, index: true },
  desc: { type: String, required: true, trim: true }, // 적요
  income: { type: Number, default: 0, min: 0 },
  expense: { type: Number, default: 0, min: 0 },
  note: { type: String, trim: true },                 // 비고(선택)
  tag: { type: String, trim: true }                   // 분류(선택)
}, { timestamps: true });

EntrySchema.virtual('net').get(function() {
  return (this.income || 0) - (this.expense || 0);
});

export default mongoose.model('Entry', EntrySchema);
