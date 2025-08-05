import mongoose, { Schema, Document } from 'mongoose';

export interface IToken extends Document {
  name: string;
  apiKey: string;
  paymentStatus: 'free' | 'trial' | 'pending' | 'disabled';
  comment?: string;
  createdAt: Date;
  updatedAt: Date;
}

const TokenSchema: Schema = new Schema({
  name: {
    type: String,
    required: true,
    trim: true
  },
  apiKey: {
    type: String,
    required: true,
    trim: true
  },
  paymentStatus: {
    type: String,
    enum: ['free', 'trial', 'pending', 'disabled'],
    default: 'free'
  },
  comment: {
    type: String,
    trim: true
  }
}, {
  timestamps: true
});

export default mongoose.models.Token || mongoose.model<IToken>('Token', TokenSchema); 