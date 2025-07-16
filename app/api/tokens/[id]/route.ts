import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';

interface RouteParams {
  params: Promise<{ id: string }>;
}

export async function PUT(request: NextRequest, { params }: RouteParams) {
  try {
    await connectDB();
    const { id } = await params;
    const body = await request.json();
    
    const { name, apiKey } = body;
    
    if (!name || !apiKey) {
      return NextResponse.json(
        { error: 'Название и API-ключ обязательны' },
        { status: 400 }
      );
    }

    const token = await Token.findByIdAndUpdate(
      id,
      { name, apiKey },
      { new: true }
    );

    if (!token) {
      return NextResponse.json(
        { error: 'Токен не найден' },
        { status: 404 }
      );
    }

    return NextResponse.json(token);
  } catch (error) {
    return NextResponse.json(
      { error: 'Ошибка при обновлении токена' },
      { status: 500 }
    );
  }
}

export async function DELETE(request: NextRequest, { params }: RouteParams) {
  try {
    await connectDB();
    const { id } = await params;
    
    const token = await Token.findByIdAndDelete(id);
    
    if (!token) {
      return NextResponse.json(
        { error: 'Токен не найден' },
        { status: 404 }
      );
    }

    return NextResponse.json({ message: 'Токен удален' });
  } catch (error) {
    return NextResponse.json(
      { error: 'Ошибка при удалении токена' },
      { status: 500 }
    );
  }
} 