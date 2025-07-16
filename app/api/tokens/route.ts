import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';

export async function GET() {
  try {
    console.log('Attempting to connect to MongoDB...');
    await connectDB();
    console.log('Connected to MongoDB successfully');
    const tokens = await Token.find({}).sort({ createdAt: -1 });
    console.log('Tokens fetched:', tokens.length);
    return NextResponse.json(tokens);
  } catch (error) {
    console.error('Error in GET /api/tokens:', error);
    return NextResponse.json(
      { 
        error: 'Ошибка при получении токенов',
        details: error instanceof Error ? error.message : 'Unknown error',
        mongoUri: process.env.MONGODB_URI ? 'Set' : 'Not set'
      },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    console.log('POST request to /api/tokens');
    await connectDB();
    console.log('Connected to MongoDB for POST');
    
    const body = await request.json();
    console.log('Request body:', body);
    
    const { name, apiKey } = body;
    
    if (!name || !apiKey) {
      return NextResponse.json(
        { error: 'Название и API-ключ обязательны' },
        { status: 400 }
      );
    }

    const token = new Token({
      name,
      apiKey
    });

    await token.save();
    console.log('Token saved successfully:', token._id);
    
    return NextResponse.json(token, { status: 201 });
  } catch (error) {
    console.error('Error in POST /api/tokens:', error);
    return NextResponse.json(
      { 
        error: 'Ошибка при создании токена',
        details: error instanceof Error ? error.message : 'Unknown error',
        mongoUri: process.env.MONGODB_URI ? 'Set' : 'Not set'
      },
      { status: 500 }
    );
  }
} 