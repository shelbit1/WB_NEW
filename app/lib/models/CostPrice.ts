import mongoose, { Schema, Document } from 'mongoose';

export interface ICostPrice extends Document {
  tokenId: string;
  productKey: string; // Формат: nmID-barcode
  nmID: number;
  barcode: string;
  costPrice: number;
  updatedBy?: string;
  createdAt: Date;
  updatedAt: Date;
}

const CostPriceSchema: Schema = new Schema({
  tokenId: {
    type: String,
    required: true,
    trim: true
  },
  productKey: {
    type: String,
    required: true,
    trim: true
  },
  nmID: {
    type: Number,
    required: true
  },
  barcode: {
    type: String,
    required: true,
    trim: true
  },
  costPrice: {
    type: Number,
    required: true,
    min: 0
  },
  updatedBy: {
    type: String,
    trim: true
  }
}, {
  timestamps: true
});

// Создаем составной индекс для быстрого поиска
CostPriceSchema.index({ tokenId: 1, productKey: 1 }, { unique: true });
CostPriceSchema.index({ tokenId: 1, nmID: 1, barcode: 1 });

export default mongoose.models.CostPrice || mongoose.model<ICostPrice>('CostPrice', CostPriceSchema); 