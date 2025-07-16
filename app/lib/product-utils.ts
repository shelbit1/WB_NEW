import { getStorageData, StorageItem } from './storage-utils';

// Интерфейсы для работы с товарами
export interface ProductCard {
  nmID: number;
  vendorCode: string;
  object?: string;
  subjectName?: string;
  brand?: string;
  createdAt?: string;
  updatedAt?: string;
  sizes?: ProductSize[];
}

export interface ProductSize {
  techSize?: string;
  wbSize?: string;
  price?: number;
  skus?: string[];
}

export interface CostPriceData {
  nmID: number;
  vendorCode: string;
  object: string;
  brand: string;
  sizeName: string;
  barcode: string;
  price: number;
  costPrice: number;
  createdAt: string;
  updatedAt: string;
  source?: string;
}

// Используем интерфейс DetailReportItem из route.ts
export interface RealizationItem {
  nm_id?: number;
  sa_name?: string;
  barcode?: string;
  retail_price_withdisc_rub?: number;
}

/**
 * Получает карточки товаров из Content API WB с поддержкой пагинации
 */
async function getProductCards(token: string, maxPages: number = 10): Promise<ProductCard[]> {
  if (!token || typeof token !== 'string' || token.trim().length === 0) {
    console.error('❌ Некорректный токен для API карточек');
    return [];
  }

  try {
    console.log('💰 Получение карточек товаров из Content API');
    
    const allCards: ProductCard[] = [];
    let currentPage = 0;
    let hasMore = true;
    let lastCursor: string | undefined;

    while (hasMore && currentPage < maxPages) {
      const requestBody = {
        settings: {
          cursor: {
            limit: 100,
            ...(lastCursor && { updatedAt: lastCursor })
          },
          filter: {
            withPhoto: -1
          }
        }
      };

      console.log(`📄 Загрузка страницы ${currentPage + 1}/${maxPages}...`);

      const response = await fetch('https://content-api.wildberries.ru/content/v2/get/cards/list', {
        method: 'POST',
        headers: {
          'Authorization': token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody),
        signal: AbortSignal.timeout(30000) // 30 секунд таймаут
      });

      if (response.ok) {
        const data = await response.json();
        
        if (data.cards && Array.isArray(data.cards)) {
          // Валидация карточек
          const validCards = data.cards.filter((card: unknown) => {
            if (!card || typeof card !== 'object') return false;
            const cardObj = card as Record<string, unknown>;
            return (cardObj.nmID || cardObj.vendorCode) && 
                   !isNaN(parseInt(cardObj.nmID as string));
          });
          
          allCards.push(...validCards);
          
          if (validCards.length !== data.cards.length) {
            console.warn(`⚠️ Отфильтровано ${data.cards.length - validCards.length} некорректных карточек на странице ${currentPage + 1}`);
          }

          // Проверяем, есть ли еще данные
          if (data.cards.length < 100) {
            hasMore = false;
          } else {
            // Используем cursor для следующей страницы
            lastCursor = data.cursor?.updatedAt;
            if (!lastCursor) {
              hasMore = false;
            }
          }
        } else {
          console.warn('⚠️ API карточек не вернул ожидаемую структуру данных');
          hasMore = false;
        }
      } else {
        const errorText = await response.text().catch(() => 'Неизвестная ошибка');
        console.warn(`⚠️ Не удалось получить карточки товаров из API контента: ${response.status} - ${errorText}`);
        
        if (response.status === 401) {
          throw new Error('Неверный или истекший токен авторизации');
        } else if (response.status === 429) {
          console.warn('⚠️ Лимит запросов превышен, ждем 2 секунды...');
          await new Promise(resolve => setTimeout(resolve, 2000));
          continue; // Повторяем попытку для той же страницы
        } else {
          hasMore = false;
        }
      }

      currentPage++;
      
      // Пауза между запросами для соблюдения rate limits
      if (hasMore && currentPage < maxPages) {
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }

    console.log(`📦 Итого получено карточек из API контента: ${allCards.length} (загружено ${currentPage} страниц)`);
    return allCards;

  } catch (error) {
    if (error instanceof Error) {
      console.error(`❌ Ошибка при получении карточек из API контента: ${error.message}`);
      throw error;
    } else {
      console.error('❌ Неизвестная ошибка при получении карточек из API контента:', error);
      return [];
    }
  }
}

/**
 * Основная функция получения данных себестоимости товаров
 */
export async function getCostPriceData(
  token: string, 
  savedCostPrices: {[key: string]: string} = {}, 
  realizationData: RealizationItem[] = []
): Promise<CostPriceData[]> {
  // Валидация входных параметров
  if (!token || typeof token !== 'string' || token.trim().length === 0) {
    throw new Error('Токен не указан или некорректен');
  }

  if (!Array.isArray(realizationData)) {
    console.warn('⚠️ realizationData не является массивом, используем пустой массив');
    realizationData = [];
  }

  if (typeof savedCostPrices !== 'object' || savedCostPrices === null) {
    console.warn('⚠️ savedCostPrices не является объектом, используем пустой объект');
    savedCostPrices = {};
  }

  try {
    console.log('💰 Получение данных себестоимости товаров');
    console.log(`📊 Параметры: сохраненных цен: ${Object.keys(savedCostPrices).length}, записей реализации: ${realizationData.length}`);
    
    const costPriceData: CostPriceData[] = [];
    
    // Шаг 1: Получаем карточки товаров из API контента
    let cards: ProductCard[] = [];
    try {
      cards = await getProductCards(token);
    } catch (error) {
      console.error('❌ Критическая ошибка при получении карточек:', error);
      // Продолжаем работу с пустым массивом карточек
      cards = [];
    }

    // Получаем данные о хранении для получения предметов
    let storageData: StorageItem[] = [];
    try {
      const today = new Date().toISOString().split('T')[0];
      storageData = await getStorageData(token, today, today);
      console.log(`📦 Получено данных хранения: ${storageData.length} записей`);
    } catch (error) {
      console.warn('⚠️ Не удалось получить данные хранения:', error);
    }

    // Создаем карту предметов из данных хранения
    const subjectMap = new Map<string, string>();
    storageData.forEach((item: StorageItem) => {
      const nmId = item.nmId?.toString() || '';
      const vendorCode = item.vendorCode || '';
      const subject = item.subject || '';
      
      if (subject && nmId) {
        subjectMap.set(nmId, subject);
      }
      if (subject && vendorCode) {
        subjectMap.set(vendorCode, subject);
      }
    });

    // Создаем карту существующих карточек для быстрого поиска
    const cardsMap = new Map<string, ProductCard>();
    cards.forEach((card: ProductCard) => {
      cardsMap.set(card.nmID?.toString() || '', card);
      cardsMap.set(card.vendorCode || '', card);
    });

    // Шаг 2: Преобразуем карточки в данные себестоимости (оптимизированно)
    console.log(`📊 Преобразование ${cards.length} карточек в данные себестоимости...`);
    
    // Предварительно выделяем массив с приблизительным размером
    const estimatedSize = cards.reduce((sum, card) => {
      if (!card.sizes) return sum + 1;
      return sum + card.sizes.reduce((sizeSum, size) => 
        sizeSum + (size.skus?.length || 1), 0);
    }, 0);
    
    console.log(`📊 Ожидаемое количество записей: ~${estimatedSize}`);
    
    cards.forEach((card: ProductCard, cardIndex: number) => {
      // Логирование прогресса для больших объемов
      if (cardIndex % 1000 === 0 && cardIndex > 0) {
        console.log(`📊 Обработано ${cardIndex}/${cards.length} карточек`);
      }

      const nmId = card.nmID?.toString() || '';
      const vendorCode = card.vendorCode || '';
      
      // Получаем предмет из разных источников (кешированно)
      const subject = card.object || card.subjectName || 
                     subjectMap.get(nmId) || subjectMap.get(vendorCode) || '';
      
      const baseProduct = {
        nmID: card.nmID,
        vendorCode: card.vendorCode,
        object: subject,
        brand: card.brand || '',
        createdAt: card.createdAt || '',
        updatedAt: card.updatedAt || ''
      };

      if (card.sizes && card.sizes.length > 0) {
        // Для каждого размера создаем отдельную запись
        card.sizes.forEach((size: ProductSize) => {
          if (size.skus && size.skus.length > 0) {
            // Для каждого SKU в размере создаем отдельную строку
            size.skus.forEach((sku: string) => {
              const productKey = `${card.nmID}-${sku}`;
              const savedCostPrice = savedCostPrices[productKey] ? parseFloat(savedCostPrices[productKey]) : 0;
              
              costPriceData.push({
                ...baseProduct,
                sizeName: size.techSize || size.wbSize || 'Без размера',
                barcode: sku,
                price: size.price || 0,
                costPrice: savedCostPrice // Используем сохраненную себестоимость
              });
            });
          } else {
            // Если нет SKU, все равно добавляем размер
            const productKey = `${card.nmID}-`;
            const savedCostPrice = savedCostPrices[productKey] ? parseFloat(savedCostPrices[productKey]) : 0;
            
            costPriceData.push({
              ...baseProduct,
              sizeName: size.techSize || size.wbSize || 'Без размера',
              barcode: '',
              price: size.price || 0,
              costPrice: savedCostPrice
            });
          }
        });
      } else {
        // Если нет размеров, добавляем товар без размера
        const productKey = `${card.nmID}-`;
        const savedCostPrice = savedCostPrices[productKey] ? parseFloat(savedCostPrices[productKey]) : 0;
        
        costPriceData.push({
          ...baseProduct,
          sizeName: 'Без размера',
          barcode: '',
          price: 0,
          costPrice: savedCostPrice
        });
      }
    });

    // Шаг 3: Добавляем товары из реализации, которых нет в карточках (оптимизированно)
    console.log('📦 Добавление товаров из реализации, отсутствующих в карточках...');
    
    // Создаем оптимизированные индексы для быстрого поиска
    const existingProducts = new Set<string>();
    const costPriceIndex = new Map<string, CostPriceData>();
    
    // Индексируем существующие товары
    costPriceData.forEach(item => {
      const key1 = `${item.nmID}-${item.barcode}`;
      const key2 = `${item.vendorCode}-${item.barcode}`;
      existingProducts.add(key1);
      existingProducts.add(key2);
      costPriceIndex.set(key1, item);
      costPriceIndex.set(key2, item);
    });

    const missingProducts = new Map<string, CostPriceData>();
    const batchSize = 1000; // Обрабатываем батчами для экономии памяти
    
    console.log(`📊 Обработка ${realizationData.length} записей реализации батчами по ${batchSize}...`);
    
    for (let i = 0; i < realizationData.length; i += batchSize) {
      const batch = realizationData.slice(i, i + batchSize);
      
      batch.forEach((item: RealizationItem, index: number) => {
        // Логирование прогресса для больших объемов
        if ((i + index) % 5000 === 0 && (i + index) > 0) {
          console.log(`📊 Обработано ${i + index}/${realizationData.length} записей реализации`);
        }

        const nmId = item.nm_id?.toString() || '';
        const vendorCode = item.sa_name || '';
        const barcode = item.barcode || '';
        
        if (!vendorCode) return; // Пропускаем товары без артикула продавца
        
        const productKey1 = `${nmId}-${barcode}`;
        const productKey2 = `${vendorCode}-${barcode}`;
        
        // Если товара нет среди карточек, добавляем его
        if (!existingProducts.has(productKey1) && !existingProducts.has(productKey2)) {
          const uniqueKey = vendorCode; // Используем артикул продавца как уникальный ключ
          
          if (!missingProducts.has(uniqueKey)) {
            const costKey = `${nmId}-${barcode}`;
            const savedCostPrice = savedCostPrices[costKey] ? parseFloat(savedCostPrices[costKey]) : 0;
            
            // Получаем предмет из карты субъектов или используем артикул как fallback
            const subject = subjectMap.get(nmId) || subjectMap.get(vendorCode) || vendorCode;
            
            missingProducts.set(uniqueKey, {
              nmID: parseInt(nmId) || 0,
              vendorCode: vendorCode,
              object: subject,
              brand: 'Неизвестный бренд',
              sizeName: 'Размер не указан',
              barcode: barcode,
              price: item.retail_price_withdisc_rub || 0,
              costPrice: savedCostPrice,
              createdAt: '',
              updatedAt: '',
              source: 'realization' // Помечаем источник данных
            });
          }
        }
      });
      
      // Короткая пауза между батчами для снижения нагрузки
      if (i + batchSize < realizationData.length) {
        await new Promise(resolve => setTimeout(resolve, 10));
      }
    }

    // Добавляем отсутствующие товары оптимизированно
    console.log(`📦 Добавление ${missingProducts.size} отсутствующих товаров...`);
    costPriceData.push(...Array.from(missingProducts.values()));

    console.log(`✅ Обработано карточек из API: ${cards.length}`);
    console.log(`📦 Добавлено товаров из реализации: ${missingProducts.size}`);
    console.log(`📊 Итого позиций в листе "Список товаров": ${costPriceData.length}`);
    
    return costPriceData;
  } catch (error) {
    console.error('❌ Ошибка при получении данных себестоимости:', error);
    return [];
  }
}

/**
 * Преобразует данные себестоимости в формат для Excel
 */
export function transformCostPriceToExcel(costPriceData: CostPriceData[]) {
  return costPriceData.map((item) => ({
    "Артикул ВБ": item.nmID || "",
    "Артикул продавца": item.vendorCode || "",
    "Предмет": item.object || "",
    "Бренд": item.brand || "",
    "Размер": item.sizeName || "",
    "Штрихкод": item.barcode || "",
    "Цена": item.price || 0,
    "Себестоимость": item.costPrice || 0,
    "Маржа": item.costPrice > 0 ? (item.price - item.costPrice) : 0,
    "Рентабельность (%)": item.costPrice > 0 && item.price > 0 ? 
      ((item.price - item.costPrice) / item.price * 100).toFixed(2) : 0,
    "Источник данных": item.source === 'realization' ? 'Из реализации' : 'Из карточек API',
    "Дата создания": item.createdAt || "",
    "Дата обновления": item.updatedAt || ""
  }));
}

/**
 * Загружает сохраненные себестоимости из базы данных
 */
export async function loadSavedCostPrices(tokenId: string): Promise<{ [key: string]: string }> {
  if (!tokenId || typeof tokenId !== 'string' || tokenId.trim().length === 0) {
    console.error('❌ Некорректный tokenId для загрузки себестоимости');
    return {};
  }

  try {
    const CostPrice = (await import('./models/CostPrice')).default;
    const connectToDatabase = (await import('./mongodb')).default;
    
    await connectToDatabase();
    
    const costPricesFromDB = await CostPrice.find({ tokenId }).lean();
    
    const savedCostPrices: { [key: string]: string } = {};
    let validCount = 0;
    
    costPricesFromDB.forEach((item) => {
      if (item && item.productKey && typeof item.costPrice === 'number' && !isNaN(item.costPrice)) {
        savedCostPrices[item.productKey] = item.costPrice.toString();
        validCount++;
      } else {
        console.warn(`⚠️ Пропущена некорректная запись себестоимости:`, item);
      }
    });
    
    console.log(`📊 Загружено ${validCount} сохраненных себестоимостей для токена ${tokenId}`);
    
    if (validCount !== costPricesFromDB.length) {
      console.warn(`⚠️ Отфильтровано ${costPricesFromDB.length - validCount} некорректных записей себестоимости`);
    }
    
    return savedCostPrices;
  } catch (error) {
    console.error('❌ Ошибка при загрузке сохраненных себестоимостей:', error);
    return {};
  }
}

/**
 * Сохраняет себестоимости в базу данных
 */
export async function saveCostPrices(tokenId: string, costPrices: { [key: string]: string }): Promise<number> {
  if (!tokenId || typeof tokenId !== 'string' || tokenId.trim().length === 0) {
    console.error('❌ Некорректный tokenId для сохранения себестоимости');
    return 0;
  }

  if (!costPrices || typeof costPrices !== 'object' || Array.isArray(costPrices)) {
    console.error('❌ Некорректные данные себестоимости для сохранения');
    return 0;
  }

  const entries = Object.entries(costPrices);
  if (entries.length === 0) {
    console.warn('⚠️ Нет данных для сохранения');
    return 0;
  }

  try {
    const CostPrice = (await import('./models/CostPrice')).default;
    const connectToDatabase = (await import('./mongodb')).default;
    
    await connectToDatabase();
    
    let savedCount = 0;
    let skippedCount = 0;
    
    console.log(`💾 Начинаем сохранение ${entries.length} записей себестоимости`);
    
    for (const [productKey, costPriceValue] of entries) {
      // Валидация productKey
      if (!productKey || typeof productKey !== 'string' || !productKey.includes('-')) {
        console.warn(`⚠️ Некорректный productKey: ${productKey}`);
        skippedCount++;
        continue;
      }

      // Валидация цены
      const costPrice = parseFloat(costPriceValue);
      if (isNaN(costPrice) || costPrice < 0) {
        console.warn(`⚠️ Некорректная себестоимость для ${productKey}: ${costPriceValue}`);
        skippedCount++;
        continue;
      }

      // Извлекаем nmID и barcode из productKey (формат: nmID-barcode)
      const parts = productKey.split('-');
      if (parts.length < 2) {
        console.warn(`⚠️ Некорректный формат productKey: ${productKey}`);
        skippedCount++;
        continue;
      }

      const nmIDStr = parts[0];
      const barcode = parts.slice(1).join('-'); // На случай, если в barcode есть дефисы
      const nmID = parseInt(nmIDStr);

      if (isNaN(nmID) || nmID <= 0) {
        console.warn(`⚠️ Некорректный nmID в productKey: ${productKey}`);
        skippedCount++;
        continue;
      }

      try {
        await CostPrice.findOneAndUpdate(
          { tokenId, productKey },
          {
            tokenId,
            productKey,
            nmID,
            barcode: barcode || '',
            costPrice,
            updatedBy: 'system'
          },
          { 
            upsert: true, 
            new: true,
            runValidators: true
          }
        );
        
        savedCount++;
      } catch (error) {
        console.error(`❌ Ошибка сохранения себестоимости для ${productKey}:`, error);
        skippedCount++;
      }
    }

    console.log(`💾 Сохранение завершено: сохранено ${savedCount}, пропущено ${skippedCount} записей`);
    return savedCount;
  } catch (error) {
    console.error('❌ Критическая ошибка при сохранении себестоимости:', error);
    return 0;
  }
} 