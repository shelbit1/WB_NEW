import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import connectToDatabase from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';
import { getStorageData, StorageItem } from '@/app/lib/storage-utils';



interface AcceptanceReportItem {
  giCreateDate: string;      // Дата создания поставки
  incomeId: number;          // Номер поставки
  nmID: number;              // Артикул Wildberries
  shkCreateDate: string;     // Дата приёмки
  subjectName: string;       // Предмет (подкатегория)
  count: number;             // Количество товаров, шт.
  total: number;             // Суммарная стоимость приёмки, ₽
}

async function fetchAcceptanceReport(apiKey: string, dateFrom: string, dateTo: string): Promise<AcceptanceReportItem[]> {
  // Детальное диагностическое логирование
  const isProblematicPeriod = dateFrom.includes('2025-06-16') || dateFrom.includes('16.06.2025') || dateFrom.includes('2025-06-22');
  
  console.log('🔍 ДИАГНОСТИКА ПЛАТНОЙ ПРИЕМКИ:');
  console.log(`   📅 Период: ${dateFrom} → ${dateTo}`);
  console.log(`   🔑 Токен (первые 20 символов): ${apiKey.substring(0, 20)}...`);
  console.log(`   ⚠️  Проблемный период (16-22 июня): ${isProblematicPeriod}`);
  console.log(`   📊 Текущее время: ${new Date().toISOString()}`);
  console.log(`   🌍 Локальное время: ${new Date().toLocaleString('ru-RU', { timeZone: 'Europe/Moscow' })}`);
  
  try {
    // Пробуем новый асинхронный API
    console.log('🔄 ЭТАП 1: Попытка использования нового асинхронного API...');
    
    // 1. Создаем задачу (GET запрос с query параметрами)
    const createUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report?dateFrom=${dateFrom}&dateTo=${dateTo}`;
    console.log(`   🌐 URL создания задачи: ${createUrl}`);
    
    // Повторяем запрос создания задачи, если получаем 429 Too Many Requests
    const maxCreateAttempts = 7; // Ограничиваем количество повторов
    let createAttempt = 0;
    let createResponse: any = null;

    while (createAttempt < maxCreateAttempts) {
      createAttempt++;
      console.log(`   🚀 Попытка ${createAttempt}/${maxCreateAttempts} создать задачу платной приёмки`);

      createResponse = await fetch(createUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'Content-Type': 'application/json'
        }
      });

      console.log(`   📡 Ответ создания задачи: ${createResponse.status} ${createResponse.statusText}`);
      console.log('   📋 Заголовки ответа:', Object.fromEntries(createResponse.headers.entries()));

      // Успешный ответ – выходим из цикла
      if (createResponse.ok) {
        break;
      }

      // Ответ 429 – ждём и повторяем
      if (createResponse.status === 429) {
        const retryAfterHeader = createResponse.headers.get('Retry-After');
        const retryAfterSeconds = retryAfterHeader ? parseInt(retryAfterHeader, 10) : 20; // По умолчанию 20 с
        const waitMs = (retryAfterSeconds + 5) * 1000; // небольшой «джиттер» 5 с
        console.log(`   ⏳ Получен 429. Ждём ${waitMs / 1000}s перед новой попыткой...`);
        await new Promise(r => setTimeout(r, waitMs));
        continue;
      }

      // Для остальных ошибок – завершаем
      const errorText = await createResponse.text();
      console.log(`   ❌ Полный ответ ошибки: ${errorText}`);
      throw new Error(`Асинхронный API недоступен: ${createResponse.status} ${createResponse.statusText}. Ответ: ${errorText}`);
    }

    if (!createResponse || !createResponse.ok) {
      const errorText = createResponse ? await createResponse.text() : 'нет ответа';
      throw new Error(`Асинхронный API недоступен после ${maxCreateAttempts} попыток. Ответ: ${errorText}`);
    }

    const createData = await createResponse.json();
    console.log(`   📋 Полный ответ создания задачи:`, JSON.stringify(createData, null, 2));
    
    const taskId = createData.data?.taskId;
    
    if (!taskId) {
      console.log(`   ❌ TaskId не найден в ответе. Структура ответа:`, createData);
      throw new Error('Не получен taskId от API');
    }
    
    console.log(`📋 ЭТАП 2: Создана задача платной приемки: ${taskId}`);

    // 2. Ждем готовности (с таймаутом)
    let status = 'processing';
    let attempts = 0;
    const maxAttempts = isProblematicPeriod ? 24 : 72; // Сокращаем таймаут для проблемного периода (2 минуты вместо 6)
    
    console.log(`⏰ Максимальное количество попыток: ${maxAttempts}`);

    while (status !== 'done' && attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, 5000)); // ждем 5 сек (согласно лимитам API)
      
      const statusUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/status`;
      console.log(`   🔍 Проверка статуса (попытка ${attempts + 1}): ${statusUrl}`);
      
      const statusResponse = await fetch(statusUrl, { 
        headers: { 'Authorization': `Bearer ${apiKey}` },
        method: 'GET'
      });

      console.log(`   📡 Ответ статуса: ${statusResponse.status} ${statusResponse.statusText}`);

      if (!statusResponse.ok) {
        const statusErrorText = await statusResponse.text();
        console.log(`   ❌ Ошибка проверки статуса: ${statusErrorText}`);
        throw new Error(`Ошибка проверки статуса: ${statusResponse.status}. Ответ: ${statusErrorText}`);
      }

      const statusData = await statusResponse.json();
      console.log(`   📊 Полный ответ статуса:`, JSON.stringify(statusData, null, 2));
      
      status = statusData.data?.status || 'unknown';
      attempts++;

      console.log(`⏳ Попытка ${attempts}: статус = ${status}`);
      
      if (status === 'done') {
        break;
      }
      
      if (status === 'error' || status === 'failed') {
        console.log(`   ❌ Задача завершилась с ошибкой: ${status}`);
        console.log(`   🔍 Детали ошибки в статусе:`, statusData);
        throw new Error(`Задача завершилась с ошибкой: ${status}. Детали: ${JSON.stringify(statusData)}`);
      }
    }

    if (status !== 'done') {
      console.log(`   ⏰ Таймаут ожидания. Финальный статус: ${status}, попыток: ${attempts}`);
      throw new Error(`Таймаут ожидания выполнения задачи. Последний статус: ${status}, попыток: ${attempts}`);
    }

    // 3. Загружаем готовые данные
    console.log('📥 ЭТАП 3: Загрузка готовых данных...');
    const downloadUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/download`;
    console.log(`   🌐 URL загрузки: ${downloadUrl}`);
    
    const downloadResponse = await fetch(downloadUrl, { 
      headers: { 'Authorization': `Bearer ${apiKey}` },
      method: 'GET'
    });

    console.log(`   📡 Ответ загрузки: ${downloadResponse.status} ${downloadResponse.statusText}`);

    if (!downloadResponse.ok) {
      const downloadErrorText = await downloadResponse.text();
      console.log(`   ❌ Ошибка загрузки: ${downloadErrorText}`);
      throw new Error(`Ошибка загрузки данных: ${downloadResponse.status}. Ответ: ${downloadErrorText}`);
    }

    const downloadData = await downloadResponse.json();
    console.log(`   📊 Размер полученных данных: ${JSON.stringify(downloadData).length} символов`);
    console.log(`   📋 Тип данных: ${typeof downloadData}, является массивом: ${Array.isArray(downloadData)}`);
    
    if (downloadData && Array.isArray(downloadData) && downloadData.length > 0) {
      console.log(`   📄 Первая запись:`, JSON.stringify(downloadData[0], null, 2));
    }
    
    if (!downloadData || !Array.isArray(downloadData)) {
      console.log('⚠️ Новый API вернул пустые данные, данных за этот период нет');
      return [];
    }

    console.log(`✅ УСПЕХ: Новый асинхронный API успешно вернул ${downloadData.length} записей`);
    return downloadData;
    
  } catch (asyncError) {
    console.log(`❌ ЭТАП 4: Ошибка нового асинхронного API: ${asyncError}`);
    console.log(`   🔍 Детали ошибки:`, asyncError);
    
    // Fallback: пробуем старый синхронный API
    try {
      console.log('🔄 ЭТАП 5: Переключение на старый синхронный API...');
      
      const fallbackUrl = `https://seller-analytics-api.wildberries.ru/api/v1/analytics/acceptance-report?dateFrom=${dateFrom}&dateTo=${dateTo}`;
      console.log(`   🌐 URL старого API: ${fallbackUrl}`);
      
      const fallbackResponse = await fetch(fallbackUrl, {
        method: 'GET',
        headers: {
          'Authorization': apiKey, // Старый API может использовать токен без Bearer
          'Content-Type': 'application/json'
        }
      });

      console.log(`   📡 Ответ старого API: ${fallbackResponse.status} ${fallbackResponse.statusText}`);
      console.log(`   📋 Заголовки ответа старого API:`, Object.fromEntries(fallbackResponse.headers.entries()));

      if (!fallbackResponse.ok) {
        const fallbackErrorText = await fallbackResponse.text();
        console.log(`   ❌ Ошибка старого API: ${fallbackErrorText}`);
        throw new Error(`Старый API тоже недоступен: ${fallbackResponse.status} ${fallbackResponse.statusText}. Ответ: ${fallbackErrorText}`);
      }

      const fallbackData = await fallbackResponse.json();
      console.log(`   📊 Размер данных старого API: ${JSON.stringify(fallbackData).length} символов`);
      console.log(`   📋 Тип данных: ${typeof fallbackData}, является массивом: ${Array.isArray(fallbackData)}`);
      console.log(`   📄 Полный ответ старого API:`, JSON.stringify(fallbackData, null, 2));
      
      const processedData = fallbackData.report || fallbackData || [];
      
      if (!processedData || !Array.isArray(processedData)) {
        console.log('⚠️ Старый API тоже вернул пустые данные');
        return [];
      }

      console.log(`✅ УСПЕХ: Старый API успешно вернул ${processedData.length} записей`);
      return processedData;

    } catch (oldApiError) {
      console.log(`❌ ФИНАЛ: И старый API тоже недоступен: ${oldApiError}`);
      console.log(`   🔍 Детали ошибки старого API:`, oldApiError);
      throw new Error(`Оба API недоступны. Новый: ${asyncError}. Старый: ${oldApiError}`);
    }
  }
}



export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { reportType, startDate, endDate, tokenId } = body;

    // Создаем новую рабочую книгу
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Отчёт');

    let headers: string[] = [];
    let fileName = '';

    switch (reportType) {
      case 'details':
        // Подключаемся к базе данных и получаем токен
        await connectToDatabase();
        
        let detailsTokenDoc;
        try {
          detailsTokenDoc = await Token.findById(tokenId);
        } catch {
          return NextResponse.json({ error: 'Невалидный ID токена' }, { status: 400 });
        }
        
        if (!detailsTokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        console.log('🚀 Начало создания отчета детализации реализации...');
        const detailsStartTime = Date.now();

        fileName = `Отчет детализации - ${startDate}–${endDate}.xlsx`;

        try {
          // Импортируем функции для работы с реализацией
          const { fetchRealizationData, addDetailedRealizationToWorkbook } = await import('@/app/lib/realization-utils');

          // Получаем данные реализации
          const realizationData = await fetchRealizationData(detailsTokenDoc.apiKey, startDate, endDate);

          console.log(`📊 Проверка данных для отчета детализации. Записей: ${realizationData.length}`);

          // Лист создается ТОЛЬКО при наличии данных реализации
          console.log(`📊 Создание листа "Полный отчет". Количество записей: ${realizationData?.length || 0}`);

          if (realizationData && realizationData.length > 0) {
            console.log("🚀 Создание отчета детализации с данными...");
            
            // Используем функцию для добавления детализированных данных в workbook
            addDetailedRealizationToWorkbook(workbook, realizationData);

            console.log(`✅ Отчет детализации создан за ${Date.now() - detailsStartTime}ms с ${realizationData.length} записями`);
            
          } else {
            // Лист НЕ СОЗДАЕТСЯ, если data пустой!
            console.log("⚠️ Лист 'Полный отчет' не создан - нет данных реализации");
            console.log("ℹ️ Нет данных для создания отчета детализации - создаем пустой шаблон");
            
            // Создаем пустой шаблон если нет данных
            headers = [
              'ID строки отчета',
              'ID отчета реализации', 
              'Дата начала',
              'Дата окончания',
              'Дата создания',
              'Валюта',
              'Код договора поставщика',
              'Дата реализации',
              'ID поставки',
              'Предмет',
              'Артикул WB',
              'Бренд',
              'Артикул поставщика',
              'Размер',
              'Штрихкод',
              'Тип документа',
              'Количество',
              'Цена розничная',
              'Сумма продаж',
              'Скидка продавца %',
              'Комиссия %',
              'Склад',
              'Тип операции',
              'Дата заказа',
              'Дата продажи',
              'ШК',
              'Цена розничная с учетом скидки',
              'Сумма доставки',
              'Сумма возврата',
              'Стоимость логистики',
              'Тип коробки',
              'Скидка товара для отчета',
              'Промо продавца',
              'Номер поставки',
              'СПП %',
              'КВ без НДС %',
              'КВ %',
              'Доплата',
              'КГВПв2',
              'Комиссия продаж',
              'К перечислению продавцу',
              'Вознаграждение с продаж',
              'Эквайринг',
              'Банк эквайринга',
              'Логистика',
              'НДС с логистики',
              'ID офиса',
              'Название офиса',
              'ID поставщика',
              'Название поставщика',
              'ИНН',
              'Номер декларации',
              'Тип бонуса',
              'ID стикера',
              'Страна',
              'Штрафы',
              'Доплаты',
              'SRID'
            ];

            // Добавляем заголовки для пустого шаблона
            worksheet.addRow(headers);
            const emptyHeaderRow = worksheet.getRow(1);
            emptyHeaderRow.font = { bold: true };
            emptyHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length / 2, 8);
            });
          }
        } catch (error) {
          console.error('❌ Ошибка при получении данных реализации:', error);
          
          // Проверяем тип ошибки для лучшей обработки
          if (error instanceof Error && error.message.includes('API_ENDPOINT_NOT_FOUND')) {
            console.log('🔧 Обнаружена проблема с endpoint API - возможно, API Wildberries изменился');
            console.log('📞 Рекомендуется обратиться к документации Wildberries API или в техподдержку');
            console.log('🔄 Создаем пустой шаблон отчета детализации');

            // Создаем пустой шаблон при ошибке API
            headers = [
              'ID строки отчета',
              'ID отчета реализации', 
              'Дата начала',
              'Дата окончания',
              'Дата создания',
              'Валюта',
              'Код договора поставщика',
              'Дата реализации',
              'ID поставки',
              'Предмет',
              'Артикул WB',
              'Бренд',
              'Артикул поставщика',
              'Размер',
              'Штрихкод',
              'Тип документа',
              'Количество',
              'Цена розничная',
              'Сумма продаж',
              'Скидка продавца %',
              'Комиссия %',
              'Склад',
              'Тип операции',
              'Дата заказа',
              'Дата продажи',
              'ШК',
              'Цена розничная с учетом скидки',
              'Сумма доставки',
              'Сумма возврата',
              'Стоимость логистики'
            ];

            // Добавляем заголовки для пустого шаблона
            worksheet.addRow(headers);
            const emptyHeaderRow = worksheet.getRow(1);
            emptyHeaderRow.font = { bold: true };
            emptyHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length / 2, 8);
            });
            
            fileName = `Шаблон отчета детализации - ${startDate}–${endDate}.xlsx`;
          } else {
            console.error('💥 Критическая ошибка при создании отчета детализации');
            return NextResponse.json({ 
              error: `Ошибка при получении данных реализации: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}` 
            }, { status: 500 });
          }
        }
        break;

      case 'storage':
        // Подключаемся к базе данных и получаем токен
        await connectToDatabase();
        
        let storageTokenDoc;
        try {
          storageTokenDoc = await Token.findById(tokenId);
        } catch {
          return NextResponse.json({ error: 'Невалидный ID токена' }, { status: 400 });
        }
        
        if (!storageTokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        console.log('🚀 Начало создания отчета "Платное хранение"...');
        const storageStartTime = Date.now();

        let storageData: any[] = [];
        try {
          // Получаем данные о платном хранении
          storageData = await getStorageData(storageTokenDoc.apiKey, startDate, endDate);
          console.log(`📊 Проверка данных для отчета "Платное хранение". Записей: ${storageData.length}`);
        } catch (error) {
          console.error('❌ Ошибка при получении данных платного хранения:', error);
          // Продолжаем создание пустого шаблона
          storageData = [];
        }

        fileName = `Платное хранение - ${startDate}–${endDate}.xlsx`;

        if (storageData && storageData.length > 0) {
          console.log("🚀 Создание отчета 'Платное хранение' с данными...");
          
          // Подготавливаем данные для Excel
          const storageExcelData = storageData.map((item: StorageItem) => ({
            "Дата": item.date || "",
            "Склад": item.warehouse || "",
            "Артикул Wildberries": item.nmId || "",
            "Размер": item.size || "",
            "Баркод": item.barcode || "",
            "Предмет": item.subject || "",
            "Бренд": item.brand || "",
            "Артикул продавца": item.vendorCode || "",
            "Объем (дм³)": item.volume || 0,
            "Тип расчета": item.calcType || "",
            "Сумма хранения": item.warehousePrice || 0,
            "Количество баркодов": item.barcodesCount || 0,
            "Коэффициент склада": item.warehouseCoef || 0,
            "Скидка лояльности (%)": item.loyaltyDiscount || 0,
            "Дата фиксации тарифа": item.tariffFixDate || "",
            "Дата снижения тарифа": item.tariffLowerDate || "",
          }));

          headers = [
            'Дата',
            'Склад',
            'Артикул Wildberries',
            'Размер',
            'Баркод',
            'Предмет',
            'Бренд',
            'Артикул продавца',
            'Объем (дм³)',
            'Тип расчета',
            'Сумма хранения',
            'Количество баркодов',
            'Коэффициент склада',
            'Скидка лояльности (%)',
            'Дата фиксации тарифа',
            'Дата снижения тарифа'
          ];

          // Добавляем заголовки
          worksheet.addRow(headers);
          const storageHeaderRow = worksheet.getRow(1);
          storageHeaderRow.font = { bold: true };
          storageHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };

          // Настройка ширины колонок для оптимального отображения
          const storageColumnWidths = [
            { wch: 12 }, // Дата
            { wch: 15 }, // Склад
            { wch: 20 }, // Артикул Wildberries
            { wch: 10 }, // Размер
            { wch: 15 }, // Баркод
            { wch: 25 }, // Предмет
            { wch: 15 }, // Бренд
            { wch: 15 }, // Артикул продавца
            { wch: 12 }, // Объем
            { wch: 20 }, // Тип расчета
            { wch: 15 }, // Сумма хранения
            { wch: 12 }, // Количество баркодов
            { wch: 12 }, // Коэффициент склада
            { wch: 15 }, // Скидка лояльности
            { wch: 15 }, // Дата фиксации тарифа
            { wch: 15 }, // Дата снижения тарифа
          ];

          storageColumnWidths.forEach((width, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = width.wch;
          });

          // Добавляем данные
          storageExcelData.forEach(item => {
            worksheet.addRow([
              item["Дата"],
              item["Склад"],
              item["Артикул Wildberries"],
              item["Размер"],
              item["Баркод"],
              item["Предмет"],
              item["Бренд"],
              item["Артикул продавца"],
              item["Объем (дм³)"],
              item["Тип расчета"],
              item["Сумма хранения"],
              item["Количество баркодов"],
              item["Коэффициент склада"],
              item["Скидка лояльности (%)"],
              item["Дата фиксации тарифа"],
              item["Дата снижения тарифа"]
            ]);
          });

          // Форматируем числовые колонки БЕЗ разделителей тысяч (формат 0,00)
          const storageNumericColumns = [9, 11, 12, 13, 14]; // Объем, Сумма хранения, Количество баркодов, Коэффициент склада, Скидка лояльности
          storageNumericColumns.forEach(columnIndex => {
            const column = worksheet.getColumn(columnIndex);
            column.eachCell((cell, rowNumber) => {
              if (rowNumber > 1) { // Пропускаем заголовок
                cell.numFmt = '0,00'; // Формат БЕЗ разделителей тысяч согласно инструкциям
              }
            });
          });

          console.log(`✅ Отчет 'Платное хранение' создан с ${storageData.length} записями за ${Date.now() - storageStartTime}мс`);
        } else {
          console.log("ℹ️ Нет данных для создания отчета 'Платное хранение' - создаем пустой шаблон");
          
          headers = [
            'Дата',
            'Склад',
            'Артикул Wildberries',
            'Размер',
            'Баркод',
            'Предмет',
            'Бренд',
            'Артикул продавца',
            'Объем (дм³)',
            'Тип расчета',
            'Сумма хранения',
            'Количество баркодов',
            'Коэффициент склада',
            'Скидка лояльности (%)',
            'Дата фиксации тарифа',
            'Дата снижения тарифа'
          ];

          // Добавляем заголовки для пустого шаблона
          worksheet.addRow(headers);
          const emptyHeaderRow = worksheet.getRow(1);
          emptyHeaderRow.font = { bold: true };
          emptyHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };
          
          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = Math.max(header.length + 5, 15);
          });
        }
        break;

      case 'acceptance':
        // Подключаемся к базе данных и получаем токен
        await connectToDatabase();
        
        let acceptanceTokenDoc;
        try {
          acceptanceTokenDoc = await Token.findById(tokenId);
        } catch {
          return NextResponse.json({ error: 'Невалидный ID токена' }, { status: 400 });
        }
        
        if (!acceptanceTokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        console.log('🚀 Начало создания отчета "Платная приемка"...');
        const acceptanceStartTime = Date.now();

        try {
          // Получаем данные о платной приемке
          const acceptanceData = await fetchAcceptanceReport(acceptanceTokenDoc.apiKey, startDate, endDate);

          console.log(`📊 Проверка данных для отчета "Платная приемка". Записей: ${acceptanceData.length}`);

          fileName = `Платная приемка - ${startDate}–${endDate}.xlsx`;

          if (acceptanceData && acceptanceData.length > 0) {
            console.log("🚀 Создание отчета 'Платная приемка' с данными...");
            
            // Заголовки согласно требованиям
            headers = [
              'Дата создания поставки',
              'Номер поставки',
              'Артикул Wildberries',
              'Дата приёмки',
              'Предмет (подкатегория)',
              'Количество товаров, шт.',
              'Суммарная стоимость приёмки, ₽'
            ];

            // Добавляем заголовки
            worksheet.addRow(headers);
            const acceptanceHeaderRow = worksheet.getRow(1);
            acceptanceHeaderRow.font = { bold: true };
            acceptanceHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };

            // Настройка ширины колонок
            const acceptanceColumnWidths = [
              { wch: 20 }, // Дата создания поставки
              { wch: 15 }, // Номер поставки
              { wch: 20 }, // Артикул Wildberries
              { wch: 20 }, // Дата приёмки
              { wch: 30 }, // Предмет (подкатегория)
              { wch: 20 }, // Количество товаров, шт.
              { wch: 25 }  // Суммарная стоимость приёмки, ₽
            ];
            
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = acceptanceColumnWidths[index]?.wch || Math.max(header.length + 5, 15);
            });

            // Добавляем данные
            acceptanceData.forEach(item => {
              worksheet.addRow([
                item.giCreateDate ? new Date(item.giCreateDate).toLocaleDateString('ru-RU') : '',
                item.incomeId || '',
                item.nmID || '',
                item.shkCreateDate ? new Date(item.shkCreateDate).toLocaleDateString('ru-RU') : '',
                item.subjectName || '',
                item.count || 0,
                item.total || 0
              ]);
            });

            // Форматируем числовые колонки для отчета платной приемки в российском формате
            const acceptanceNumericColumns = [6, 7]; // Количество товаров, Суммарная стоимость приёмки
            acceptanceNumericColumns.forEach(columnIndex => {
              const column = worksheet.getColumn(columnIndex);
              column.eachCell((cell, rowNumber) => {
                if (rowNumber > 1) { // Пропускаем заголовок
                  cell.numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00'; // Российский формат с локалью
                }
              });
            });

            console.log(`✅ Отчет "Платная приемка" создан за ${Date.now() - acceptanceStartTime}ms с ${acceptanceData.length} записями`);
            
          } else {
            console.log("ℹ️ Нет данных для создания отчета 'Платная приемка' - создаем пустой шаблон");
            
            headers = [
              'Дата создания поставки',
              'Номер поставки',
              'Артикул Wildberries',
              'Дата приёмки',
              'Предмет (подкатегория)',
              'Количество товаров, шт.',
              'Суммарная стоимость приёмки, ₽'
            ];

            // Добавляем заголовки для пустого шаблона
            worksheet.addRow(headers);
            const emptyHeaderRow = worksheet.getRow(1);
            emptyHeaderRow.font = { bold: true };
            emptyHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length + 5, 15);
            });
          }
        } catch (error) {
          console.error('❌ Ошибка при получении данных платной приемки:', error);
          return NextResponse.json({ 
            error: `Ошибка при получении данных платной приемки: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}` 
          }, { status: 500 });
        }
        break;

      case 'products':
        // Подключаемся к базе данных и получаем токен
        await connectToDatabase();
        
        let productsTokenDoc;
        try {
          productsTokenDoc = await Token.findById(tokenId);
        } catch {
          return NextResponse.json({ error: 'Невалидный ID токена' }, { status: 400 });
        }
        
        if (!productsTokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        console.log('🚀 Начало создания списка товаров...');
        const productsStartTime = Date.now();

                 // Данные реализации больше не используются для товаров
         const realizationData: any[] = [];
         console.log('ℹ️ Данные реализации не загружаются для списка товаров');

                 // Импортируем функции для работы с товарами
         const { getCostPriceData, transformCostPriceToExcel, loadSavedCostPrices } = await import('@/app/lib/product-utils');

         // Получаем сохраненные себестоимости из базы данных
         const savedCostPrices = await loadSavedCostPrices(tokenId);

         // Получаем данные себестоимости с сохраненными ценами
         const costPriceData = await getCostPriceData(productsTokenDoc.apiKey, savedCostPrices, realizationData);

        fileName = `Список товаров - ${startDate}–${endDate}.xlsx`;

        if (costPriceData && costPriceData.length > 0) {
          console.log("🚀 Создание списка товаров с данными...");
          
          // Преобразуем данные для Excel
          const productsExcelData = transformCostPriceToExcel(costPriceData);

          headers = [
            'Артикул ВБ',
            'Артикул продавца',
            'Предмет',
            'Бренд',
            'Размер',
            'Штрихкод',
            'Цена',
            'Себестоимость',
            'Маржа',
            'Рентабельность (%)',
            'Источник данных',
            'Дата создания',
            'Дата обновления'
          ];

          // Добавляем заголовки
          worksheet.addRow(headers);
          const productsHeaderRow = worksheet.getRow(1);
          productsHeaderRow.font = { bold: true };
          productsHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };

          // Настройка ширины колонок для списка товаров
          const productsColumnWidths = [
            { wch: 15 }, // Артикул ВБ
            { wch: 20 }, // Артикул продавца
            { wch: 35 }, // Предмет
            { wch: 15 }, // Бренд
            { wch: 15 }, // Размер
            { wch: 15 }, // Штрихкод
            { wch: 15 }, // Цена
            { wch: 15 }, // Себестоимость
            { wch: 15 }, // Маржа
            { wch: 20 }, // Рентабельность
            { wch: 20 }, // Источник данных
            { wch: 20 }, // Дата создания
            { wch: 20 }, // Дата обновления
          ];

          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = productsColumnWidths[index]?.wch || Math.max(header.length + 5, 15);
          });

          // Добавляем данные
          productsExcelData.forEach(item => {
            worksheet.addRow([
              item["Артикул ВБ"],
              item["Артикул продавца"],
              item["Предмет"],
              item["Бренд"],
              item["Размер"],
              item["Штрихкод"],
              item["Цена"],
              item["Себестоимость"],
              item["Маржа"],
              item["Рентабельность (%)"],
              item["Источник данных"],
              item["Дата создания"],
              item["Дата обновления"]
            ]);
          });

          // Форматируем числовые колонки для списка товаров в российском формате
          const productsNumericColumns = [7, 8, 9, 10]; // Цена, Себестоимость, Маржа, Рентабельность
          productsNumericColumns.forEach(columnIndex => {
            const column = worksheet.getColumn(columnIndex);
            column.eachCell((cell, rowNumber) => {
              if (rowNumber > 1) { // Пропускаем заголовок
                cell.numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00'; // Российский формат с локалью
              }
            });
          });

          console.log(`✅ Список товаров создан за ${Date.now() - productsStartTime}ms с ${productsExcelData.length} записями`);
          
        } else {
          console.log("ℹ️ Нет данных для создания списка товаров - создаем пустой шаблон");
          
          headers = [
            'Артикул ВБ',
            'Артикул продавца',
            'Предмет',
            'Бренд',
            'Размер',
            'Штрихкод',
            'Цена',
            'Себестоимость',
            'Маржа',
            'Рентабельность (%)',
            'Источник данных',
            'Дата создания',
            'Дата обновления'
          ];

          // Добавляем заголовки для пустого шаблона
          worksheet.addRow(headers);
          const emptyHeaderRow = worksheet.getRow(1);
          emptyHeaderRow.font = { bold: true };
          emptyHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };
          
          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = Math.max(header.length + 5, 15);
          });
        }
        break;

      case 'finances':
        // Подключаемся к базе данных и получаем токен
        await connectToDatabase();
        
        let financeTokenDoc;
        try {
          financeTokenDoc = await Token.findById(tokenId);
        } catch {
          return NextResponse.json({ error: 'Невалидный ID токена' }, { status: 400 });
        }
        
        if (!financeTokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        // Импортируем функции для работы с финансами РК
        const { fetchCampaigns, fetchSKUData, fetchFinancialData } = await import('@/app/lib/finance-utils');

        console.log('🚀 Начало создания листа "Финансы РК"...');
        const financeStartTime = Date.now();

        // Получаем кампании
        const campaigns = await fetchCampaigns(financeTokenDoc.apiKey);
        
        // Получаем финансовые данные с буферными днями
        const financialData = await fetchFinancialData(financeTokenDoc.apiKey, startDate, endDate);

        console.log(`📊 Проверка данных для листа "Финансы РК". Кампаний: ${campaigns.length}, финансовых записей: ${financialData.length}`);

        if (campaigns.length > 0 && financialData.length > 0) {
          console.log("🚀 Создание листа 'Финансы РК'...");
          
          // Получаем SKU данные
          console.log(`📊 Получение SKU данных для кампаний...`);
          const skuMap = await fetchSKUData(financeTokenDoc.apiKey, campaigns);
          console.log(`📊 Использование готовых SKU данных: ${skuMap.size} кампаний`);
          
          // Создаем карту кампаний для быстрого поиска
          const campaignMap = new Map(campaigns.map(c => [c.advertId, c]));
          
          // Подготавливаем данные для листа "Финансы РК"
          const financeExcelData = financialData.map(record => {
            const campaign = campaignMap.get(record.advertId);
            return {
              "ID кампании": record.advertId,
              "Название кампании": campaign?.name || 'Неизвестная кампания',
              "SKU ID": skuMap.get(record.advertId) || '',
              "Дата": record.date,
              "Сумма": record.sum,
              "Источник списания": record.bill === 1 ? 'Счет' : 'Баланс',
              "Тип операции": record.type,
              "Номер документа": record.docNumber
            };
          });

          headers = [
            'ID кампании',
            'Название кампании', 
            'SKU ID',
            'Дата',
            'Сумма',
            'Источник списания',
            'Тип операции',
            'Номер документа'
          ];

          fileName = `Финансы РК - ${startDate}–${endDate}.xlsx`;

          // Добавляем заголовки
          worksheet.addRow(headers);

          // Стилизуем заголовки
          const financesHeaderRow = worksheet.getRow(1);
          financesHeaderRow.font = { bold: true };
          financesHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };

          // Настройка ширины колонок для листа "Финансы РК"
          const financeColumnWidths = [
            { wch: 15 }, // ID кампании
            { wch: 30 }, // Название кампании
            { wch: 15 }, // SKU ID
            { wch: 15 }, // Дата
            { wch: 15 }, // Сумма
            { wch: 20 }, // Источник списания
            { wch: 15 }, // Тип операции
            { wch: 20 }, // Номер документа
          ];
          
          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = financeColumnWidths[index]?.wch || Math.max(header.length + 5, 15);
          });

          // Добавляем данные
          financeExcelData.forEach(item => {
            worksheet.addRow([
              item["ID кампании"],
              item["Название кампании"],
              item["SKU ID"],
              item["Дата"],
              item["Сумма"],
              item["Источник списания"],
              item["Тип операции"],
              item["Номер документа"]
            ]);
          });

          // Форматируем числовые колонки для отчета финансов РК в российском формате
          const financeNumericColumns = [5]; // Сумма
          financeNumericColumns.forEach(columnIndex => {
            const column = worksheet.getColumn(columnIndex);
            column.eachCell((cell, rowNumber) => {
              if (rowNumber > 1) { // Пропускаем заголовок
                cell.numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00'; // Российский формат с локалью
              }
            });
          });

          console.log(`✅ Лист "Финансы РК" создан за ${Date.now() - financeStartTime}ms с ${financeExcelData.length} записями`);
          
        } else {
          console.log("ℹ️ Нет данных для создания листа 'Финансы РК' - создаем пустой шаблон");
          
          headers = [
            'ID кампании',
            'Название кампании', 
            'SKU ID',
            'Дата',
            'Сумма',
            'Источник списания',
            'Тип операции',
            'Номер документа'
          ];

          fileName = `Финансы РК - ${startDate}–${endDate}.xlsx`;

          // Добавляем заголовки для пустого шаблона
          worksheet.addRow(headers);
          const emptyHeaderRow = worksheet.getRow(1);
          emptyHeaderRow.font = { bold: true };
          emptyHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
          };
          
          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = Math.max(header.length + 5, 15);
          });
        }
        break;

      default:
        return NextResponse.json({ error: 'Неизвестный тип отчёта' }, { status: 400 });
    }

    // Генерируем буфер
    const buffer = await workbook.xlsx.writeBuffer();

    // Возвращаем файл
    return new NextResponse(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(fileName)}"`,
      },
    });

  } catch (error) {
    console.error('Ошибка генерации отчёта:', error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : 'Ошибка при генерации отчёта' },
      { status: 500 }
    );
  }
} 