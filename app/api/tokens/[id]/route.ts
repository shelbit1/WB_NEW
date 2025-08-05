import { NextRequest, NextResponse } from 'next/server';
import connectDB from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';

export async function PUT(request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const body = await request.json();
    
    await connectDB();
    
    // Позволяем обновлять любые поля из модели
    const { name, apiKey, paymentStatus, comment } = body;

    const updateData: {
      name?: string;
      apiKey?: string;
      paymentStatus?: 'free' | 'trial' | 'pending' | 'disabled';
      comment?: string;
    } = {};

    if (name) updateData.name = name;
    if (apiKey) updateData.apiKey = apiKey;
    if (paymentStatus) updateData.paymentStatus = paymentStatus;
    if (comment !== undefined) updateData.comment = comment;
    
    const updatedToken = await Token.findByIdAndUpdate(
      id,
      { $set: updateData },
      { new: true, runValidators: true }
    );

    if (!updatedToken) {
      return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
    }

    return NextResponse.json(updatedToken);
  } catch (error) {
    console.error(`Error in PUT /api/tokens/[id]:`, error);
    return NextResponse.json(
      { 
        error: 'Ошибка при обновлении токена',
        details: error instanceof Error ? error.message : 'Unknown error',
      },
      { status: 500 }
    );
  }
}

export async function DELETE(request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    await connectDB();
    const { id } = await params;
    const deletedToken = await Token.findByIdAndDelete(id);

    if (!deletedToken) {
      return NextResponse.json(
        { error: 'Токен не найден' },
        { status: 404 }
      );
    }

    return NextResponse.json({ message: 'Токен успешно удален' });
  } catch {
    return NextResponse.json(
      { error: 'Ошибка при удалении токена' },
      { status: 500 }
    );
  }
} 