import mongoose, { Schema, Document } from 'mongoose';

export interface IToken extends Document {
  name: string;
  apiKey: string;
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
  }
}, {
  timestamps: true
});

export default mongoose.models.Token || mongoose.model<IToken>('Token', TokenSchema); 