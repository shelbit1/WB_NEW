import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';

export async function GET() {
  try {
    await connectDB();
    const tokens = await Token.find({}).sort({ createdAt: -1 });
    return NextResponse.json(tokens);
  } catch (error) {
    return NextResponse.json(
      { error: 'Ошибка при получении токенов' },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    await connectDB();
    const body = await request.json();
    
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
    
    return NextResponse.json(token, { status: 201 });
  } catch (error) {
    return NextResponse.json(
      { error: 'Ошибка при создании токена' },
      { status: 500 }
    );
  }
} 