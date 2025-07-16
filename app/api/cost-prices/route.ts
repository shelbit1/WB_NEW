import { NextRequest, NextResponse } from 'next/server';
import connectToDatabase from '@/app/lib/mongodb';
import CostPrice from '@/app/lib/models/CostPrice';

// GET - получение всех себестоимостей для токена
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const tokenId = searchParams.get('tokenId');

    if (!tokenId) {
      return NextResponse.json({ error: 'Требуется указать tokenId' }, { status: 400 });
    }

    await connectToDatabase();

    const costPrices = await CostPrice.find({ tokenId }).lean();
    
    // Преобразуем в формат для getCostPriceData
    const costPricesMap: { [key: string]: string } = {};
    costPrices.forEach((item) => {
      costPricesMap[item.productKey] = item.costPrice.toString();
    });

    return NextResponse.json({ costPrices: costPricesMap });
  } catch (error) {
    console.error('Ошибка при получении себестоимости:', error);
    return NextResponse.json(
      { error: 'Ошибка при получении себестоимости' },
      { status: 500 }
    );
  }
}

// POST - сохранение/обновление себестоимости
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { tokenId, costPrices } = body;

    if (!tokenId || !costPrices || typeof costPrices !== 'object') {
      return NextResponse.json({ error: 'Неверные данные запроса' }, { status: 400 });
    }

    await connectToDatabase();

    const results = [];
    
    for (const [productKey, costPriceValue] of Object.entries(costPrices)) {
      const costPrice = parseFloat(costPriceValue as string);
      
      if (isNaN(costPrice) || costPrice < 0) {
        continue; // Пропускаем некорректные значения
      }

      // Извлекаем nmID и barcode из productKey (формат: nmID-barcode)
      const [nmIDStr, barcode] = productKey.split('-', 2);
      const nmID = parseInt(nmIDStr);

      if (isNaN(nmID)) {
        continue; // Пропускаем некорректные ключи
      }

      try {
        const result = await CostPrice.findOneAndUpdate(
          { tokenId, productKey },
          {
            tokenId,
            productKey,
            nmID,
            barcode: barcode || '',
            costPrice,
            updatedBy: 'system' // Можно заменить на user ID если есть аутентификация
          },
          { 
            upsert: true, 
            new: true,
            runValidators: true
          }
        );
        
        results.push(result);
      } catch (error) {
        console.warn(`Ошибка сохранения себестоимости для ${productKey}:`, error);
      }
    }

    return NextResponse.json({ 
      message: `Сохранено ${results.length} записей себестоимости`,
      saved: results.length 
    });
  } catch (error) {
    console.error('Ошибка при сохранении себестоимости:', error);
    return NextResponse.json(
      { error: 'Ошибка при сохранении себестоимости' },
      { status: 500 }
    );
  }
}

// DELETE - удаление себестоимости для конкретного товара
export async function DELETE(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const tokenId = searchParams.get('tokenId');
    const productKey = searchParams.get('productKey');

    if (!tokenId || !productKey) {
      return NextResponse.json({ error: 'Требуется указать tokenId и productKey' }, { status: 400 });
    }

    await connectToDatabase();

    const result = await CostPrice.findOneAndDelete({ tokenId, productKey });

    if (!result) {
      return NextResponse.json({ error: 'Запись не найдена' }, { status: 404 });
    }

    return NextResponse.json({ message: 'Себестоимость удалена' });
  } catch (error) {
    console.error('Ошибка при удалении себестоимости:', error);
    return NextResponse.json(
      { error: 'Ошибка при удалении себестоимости' },
      { status: 500 }
    );
  }
} 