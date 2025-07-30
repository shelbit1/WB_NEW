import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import connectToDatabase from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';

// Интерфейс для данных детализации реализации
interface RealizationDetailItem {
  realizationreport_id?: string;
  date_from?: string;
  date_to?: string;
  create_dt?: string;
  currency_name?: string;
  suppliercontract_code?: string;
  rrd_id?: string;
  gi_id?: string;
  subject_name?: string;
  nm_id?: number;
  brand_name?: string;
  sa_name?: string;
  ts_name?: string;
  barcode?: string;
  doc_type_name?: string;
  quantity?: number;
  retail_price?: number;
  retail_amount?: number;
  sale_percent?: number;
  commission_percent?: number;
  office_name?: string;
  supplier_oper_name?: string;
  order_dt?: string;
  sale_dt?: string;
  rr_dt?: string;
  shk_id?: string;
  retail_price_withdisc_rub?: number;
  delivery_amount?: number;
  return_amount?: number;
  delivery_rub?: number;
  gi_box_type_name?: string;
  product_discount_for_report?: number;
  supplier_promo?: number;
  rid?: string;
  ppvz_spp_prc?: number;
  ppvz_kvw_prc_base?: number;
  ppvz_kvw_prc?: number;
  sup_rating_prc_up?: number;
  is_kgvp_v2?: number;
  ppvz_sales_commission?: number;
  ppvz_for_pay?: number;
  ppvz_reward?: number;
  acquiring_fee?: number;
  acquiring_percent?: number;
  acquiring_bank?: string;
  ppvz_vw?: number;
  ppvz_vw_nds?: number;
  ppvz_office_id?: string;
  ppvz_office_name?: string;
  ppvz_supplier_id?: string;
  ppvz_supplier_name?: string;
  ppvz_inn?: string;
  declaration_number?: string;
  bonus_type_name?: string;
  sticker_id?: string;
  site_country?: string;
  penalty?: number;
  additional_payment?: number;
  rebill_logistic_cost?: number;
  rebill_logistic_org?: string;
  kiz?: string;
  storage_fee?: number;
  deduction?: number;
  acceptance?: number;
  srid?: string;
}

// Функция получения данных детализации реализации
async function fetchRealizationDetailData(apiKey: string, dateFrom: string, dateTo: string): Promise<RealizationDetailItem[]> {
  console.log(`📊 Начало получения данных детализации: ${dateFrom} - ${dateTo}`);
  
  // Конвертируем даты в формат RFC3339
  const startDate = new Date(dateFrom + 'T00:00:00Z').toISOString();
  const endDate = new Date(dateTo + 'T23:59:59Z').toISOString();
  
  console.log(`📅 Конвертированные даты: ${startDate} - ${endDate}`);
  
  const allData: RealizationDetailItem[] = [];
  let rrd_id = 0;
  let hasMoreData = true;
  let attempts = 0;
  const maxAttempts = 50; // Максимум 50 страниц для безопасности
  
  while (hasMoreData && attempts < maxAttempts) {
    attempts++;
    console.log(`📄 Запрос страницы ${attempts}, rrd_id: ${rrd_id}`);
    
    try {
      const url = `https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=${startDate}&dateTo=${endDate}&limit=100000&rrd_id=${rrd_id}`;
      console.log(`🌐 URL: ${url}`);
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': apiKey,
          'Content-Type': 'application/json'
        }
      });

      console.log(`📡 Ответ API: ${response.status} ${response.statusText}`);

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`❌ Ошибка API детализации: ${response.status}`, errorText);
        
        if (response.status === 401) {
          throw new Error('Неверный или недействительный API токен');
        } else if (response.status === 429) {
          console.log('⏳ Превышен лимит запросов, ждем 65 секунд...');
          await new Promise(resolve => setTimeout(resolve, 65000));
          continue; // Повторяем тот же запрос
        } else if (response.status === 400) {
          throw new Error('Некорректные параметры запроса. Проверьте период дат.');
        }
        throw new Error(`Ошибка API: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      console.log(`📦 Получено записей: ${Array.isArray(data) ? data.length : 0}`);
      
      if (!Array.isArray(data) || data.length === 0) {
        console.log('📄 Больше данных нет, завершаем');
        hasMoreData = false;
        break;
      }

      // Добавляем данные (фильтруем дубликаты если это не первая страница)
      if (attempts === 1) {
        // Первая страница - добавляем все данные
        allData.push(...data);
      } else {
        // Последующие страницы - исключаем первую запись, так как она дублирует последнюю с предыдущей страницы
        const filteredData = data.filter((item, index) => {
          if (index === 0) {
            // Проверяем, не дублируется ли первая запись
            return !allData.some(existingItem => existingItem.rrd_id === item.rrd_id);
          }
          return true;
        });
        allData.push(...filteredData);
        console.log(`📦 После фильтрации дубликатов добавлено: ${filteredData.length} записей`);
      }
      
      // Получаем rrd_id для следующей страницы
      const lastItem = data[data.length - 1];
      const nextRrdId = lastItem?.rrd_id;
      
      if (nextRrdId && nextRrdId !== rrd_id) {
        rrd_id = nextRrdId;
        console.log(`➡️ Следующий rrd_id: ${rrd_id}`);
      } else {
        console.log('📄 Достигнут конец данных');
        hasMoreData = false;
      }

      // Пауза между запросами (1 запрос в минуту)
      if (hasMoreData) {
        console.log('⏳ Ожидание 65 секунд между запросами...');
        await new Promise(resolve => setTimeout(resolve, 65000));
      }

    } catch (error) {
      console.error(`❌ Ошибка на странице ${attempts}:`, error);
      throw error;
    }
  }

  console.log(`📊 Получено всего записей до дедупликации: ${allData.length}`);
  
  // Финальная дедупликация по rrd_id для полной гарантии
  const uniqueData = allData.filter((item, index, array) => {
    return array.findIndex(t => t.rrd_id === item.rrd_id) === index;
  });
  
  const duplicatesCount = allData.length - uniqueData.length;
  if (duplicatesCount > 0) {
    console.log(`🔍 Удалено дубликатов: ${duplicatesCount}`);
  }
  
  console.log(`✅ Завершено получение данных детализации: ${uniqueData.length} уникальных записей за ${attempts} запросов`);
  return uniqueData;
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
          // Получаем данные детализации реализации
          const realizationData = await fetchRealizationDetailData(detailsTokenDoc.apiKey, startDate, endDate);

          console.log(`📊 Получено записей детализации: ${realizationData.length}`);

          if (realizationData && realizationData.length > 0) {
            console.log("🚀 Создание отчета детализации с данными...");
            
                        // Заголовки согласно скриншоту
            headers = [
              'Номер отчёта',
              'Дата начала отчётного периода',
              'Дата конца отчётного периода',
              'Дата формирования отчёта',
              'Валюта отчёта',
              'Договор',
              'Номер строки',
              'Номер поставки',
              'Предмет',
              'Артикул WB',
              'Бренд',
              'Артикул продавца',
              'Размер',
              'Баркод',
              'Тип документа',
              'Количество',
              'Цена розничная',
              'Сумма продаж (возвратов)',
              'Согласованная скидка',
              'Процент комиссии',
              'Склад',
              'Обоснование для оплаты',
              'Дата заказа',
              'Дата продажи',
              'Дата операции',
              'Штрих-код',
              'Цена розничная с учетом согласованной скидки',
              'Количество доставок',
              'Количество возвратов',
              'Стоимость логистики',
              'Тип коробов',
              'Согласованный продуктовый дисконт',
              'Промокод',
              'Уникальный идентификатор',
              'Скидка постоянного покупателя',
              'Размер кВВ без НДС, %',
              'Итоговый кВВ без НДС, %',
              'Размер снижения кВВ из-за рейтинга',
              'Размер снижения кВВ из-за акции',
              'Вознаграждение с продажи до вычета',
              'К перечислению продавцу',
              'Возмещение за выдачу и возврат',
              'Возмещение издержек по эквайрингу',
              'Наименование банка-эквайера',
              'Вознаграждение WB без НДС',
              'НДС с вознаграждения WB',
              'Номер офиса',
              'Наименование офиса доставки',
              'Номер партнера',
              'Партнер',
              'ИНН',
              'Номер таможенной декларации',
              'Обоснование штрафов и доплат',
              'Цифровое значение стикера, который клеится на товар',
              'Страна продажи',
              'Штрафы',
              'Доплаты',
              'Возмещение издержек по',
              'Организатор перевозки',
              'Код маркировки',
              'Стоимость хранения',
              'Прочие удержания/выплаты',
              'Стоимость платной приёмки',
              'Уникальный идентификатор автора'
            ];

            // Добавляем заголовки
            worksheet.addRow(headers);
            const headerRow = worksheet.getRow(1);
            headerRow.font = { bold: true };
            headerRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            
            // Настройка ширины колонок
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length + 2, 12);
          });

          // Добавляем данные
            realizationData.forEach(item => {
            worksheet.addRow([
                item.realizationreport_id || '',
                item.date_from ? new Date(item.date_from).toLocaleDateString('ru-RU') : '',
                item.date_to ? new Date(item.date_to).toLocaleDateString('ru-RU') : '',
                item.create_dt ? new Date(item.create_dt).toLocaleDateString('ru-RU') : '',
                item.currency_name || '',
                item.suppliercontract_code || '',
                item.rrd_id || '',
                item.gi_id || '',
                item.subject_name || '',
                item.nm_id || '',
                item.brand_name || '',
                item.sa_name || '',
                item.ts_name || '',
                item.barcode || '',
                item.doc_type_name || '',
                item.quantity || 0,
                item.retail_price || 0,
                item.retail_amount || 0,
                item.sale_percent || 0,
                item.commission_percent || 0,
                item.office_name || '',
                item.supplier_oper_name || '',
                item.order_dt ? new Date(item.order_dt).toLocaleDateString('ru-RU') : '',
                item.sale_dt ? new Date(item.sale_dt).toLocaleDateString('ru-RU') : '',
                item.rr_dt ? new Date(item.rr_dt).toLocaleDateString('ru-RU') : '',
                item.shk_id || '',
                item.retail_price_withdisc_rub || 0,
                item.delivery_amount || 0,
                item.return_amount || 0,
                item.delivery_rub || 0,
                item.gi_box_type_name || '',
                item.product_discount_for_report || 0,
                item.supplier_promo || 0,
                item.rid || '',
                item.ppvz_spp_prc || 0,
                item.ppvz_kvw_prc_base || 0,
                item.ppvz_kvw_prc || 0,
                item.sup_rating_prc_up || 0,
                item.is_kgvp_v2 || 0,
                item.ppvz_sales_commission || 0,
                item.ppvz_for_pay || 0,
                item.ppvz_reward || 0,
                item.acquiring_fee || 0,
                item.acquiring_percent || 0,
                item.acquiring_bank || '',
                item.ppvz_vw || 0,
                item.ppvz_vw_nds || 0,
                item.ppvz_office_id || '',
                item.ppvz_supplier_id || '',
                item.ppvz_supplier_name || '',
                item.ppvz_inn || '',
                item.declaration_number || '',
                item.bonus_type_name || '',
                item.sticker_id || '',
                item.site_country || '',
                item.penalty || 0,
                item.additional_payment || 0,
                item.rebill_logistic_cost || 0,
                item.rebill_logistic_org || '',
                item.kiz || '',
                item.storage_fee || 0,
                item.deduction || 0,
                item.acceptance || 0,
                item.srid || ''
            ]);
          });

            // Добавляем заголовок и формулы в столбец BN "Даты отчета"
            console.log("📝 Добавляем столбец 'Даты отчета' в BN...");
            
            // Устанавливаем заголовок в BN1
            worksheet.getCell('BN1').value = 'Даты отчета';
            
            // Добавляем формулы для данных
            for (let rowIndex = 2; rowIndex <= realizationData.length + 1; rowIndex++) {
              const cell = worksheet.getCell(`BN${rowIndex}`);
              cell.value = { formula: `B${rowIndex}&" - "&C${rowIndex}` };
            }

            // Форматируем числовые колонки в российском формате (числовой формат с 2 знаками после запятой)
            const numericColumns = [16, 17, 18, 27, 28, 29, 30, 40, 41, 42, 43, 45, 46, 55, 56, 60, 61, 62]; // Количество, Цена розничная, Сумма продаж, Цена с учетом скидки, Количество доставок, Количество возвратов, Стоимость логистики, Комиссия за продажи, К доплате, Вознаграждение, Эквайринг, Логистика ВВ, НДС логистики ВВ, Штрафы, Доплаты, Стоимость хранения, Прочие удержания, Стоимость платной приёмки
            numericColumns.forEach(columnIndex => {
              const column = worksheet.getColumn(columnIndex);
              column.eachCell((cell, rowNumber) => {
                if (rowNumber > 1) { // Пропускаем заголовок
                  // Российский числовой формат с запятой как десятичный разделитель
                  cell.numFmt = '[$-419]# ##0,00';
                  cell.value = Number(cell.value) || 0; // Принудительно делаем значение числом
                }
              });
            });

            // Создаем новый лист "Товары"
            const productsWorksheet = workbook.addWorksheet('Товары');
            const productHeaders = ['Артикул продавца', 'Артикул WB'];
            const productHeaderRow = productsWorksheet.addRow(productHeaders);
            productHeaderRow.font = { bold: true };
            productHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            productsWorksheet.getColumn(1).width = 30;
            productsWorksheet.getColumn(2).width = 20;

            // Собираем уникальные товары
            const uniqueProducts = new Map<string, { vendorCode: string, nmId: number }>();
            realizationData.forEach(item => {
              if (item.nm_id) {
                const key = `${item.nm_id}`;
                if (!uniqueProducts.has(key)) {
                  uniqueProducts.set(key, {
                    vendorCode: item.sa_name || '',
                    nmId: item.nm_id
                  });
                }
              }
            });

            // Добавляем уникальные товары на лист
            uniqueProducts.forEach(prod => {
              productsWorksheet.addRow([prod.vendorCode, prod.nmId]);
            });


            console.log(`✅ Отчет детализации создан за ${Date.now() - detailsStartTime}ms с ${realizationData.length} записями`);
            
          } else {
            console.log("ℹ️ Нет данных для создания отчета детализации - создаем пустой шаблон");
            
                        // Создаем пустой шаблон с правильными заголовками
            headers = [
              'Номер отчёта',
              'Дата начала отчётного периода',
              'Дата конца отчётного периода',
              'Дата формирования отчёта',
              'Валюта отчёта',
              'Договор',
              'Номер строки',
              'Номер поставки',
              'Предмет',
              'Артикул WB',
              'Бренд',
              'Артикул продавца',
              'Размер',
              'Баркод',
              'Тип документа',
              'Количество',
              'Цена розничная',
              'Сумма продаж (возвратов)',
              'Согласованная скидка',
              'Процент комиссии',
              'Склад',
              'Обоснование для оплаты',
              'Дата заказа',
              'Дата продажи',
              'Дата операции',
              'Штрих-код',
              'Цена розничная с учетом согласованной скидки',
              'Количество доставок',
              'Количество возвратов',
              'Стоимость логистики',
              'Тип коробов',
              'Согласованный продуктовый дисконт',
              'Промокод',
              'Уникальный идентификатор',
              'Скидка постоянного покупателя',
              'Размер кВВ без НДС, %',
              'Итоговый кВВ без НДС, %',
              'Размер снижения кВВ из-за рейтинга',
              'Размер снижения кВВ из-за акции',
              'Вознаграждение с продажи до вычета',
              'К перечислению продавцу',
              'Возмещение за выдачу и возврат',
              'Возмещение издержек по эквайрингу',
              'Наименование банка-эквайера',
              'Вознаграждение WB без НДС',
              'НДС с вознаграждения WB',
              'Номер офиса',
              'Наименование офиса доставки',
              'Номер партнера',
              'Партнер',
              'ИНН',
              'Номер таможенной декларации',
              'Обоснование штрафов и доплат',
              'Цифровое значение стикера, который клеится на товар',
              'Страна продажи',
              'Штрафы',
              'Доплаты',
              'Возмещение издержек по',
              'Организатор перевозки',
              'Код маркировки',
              'Стоимость хранения',
              'Прочие удержания/выплаты',
              'Стоимость платной приёмки',
              'Уникальный идентификатор автора'
            ];

            // Добавляем заголовки для пустого шаблона
            worksheet.addRow(headers);
            
            // Добавляем заголовок "Даты отчета" в BN1 для пустого шаблона
            worksheet.getCell('BN1').value = 'Даты отчета';
            const emptyHeaderRow = worksheet.getRow(1);
            emptyHeaderRow.font = { bold: true };
            emptyHeaderRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };
            
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length + 2, 12);
            });
          }
        } catch (error) {
          console.error('❌ Ошибка при получении данных детализации:', error);
          return NextResponse.json({ 
            error: `Ошибка при получении данных детализации: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}` 
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

                 // Импортируем функции для работы с товарами
         const { getCostPriceData, transformCostPriceToExcel, loadSavedCostPrices } = await import('@/app/lib/product-utils');

         // Получаем сохраненные себестоимости из базы данных
         const savedCostPrices = await loadSavedCostPrices(tokenId);

         // Получаем данные себестоимости с сохраненными ценами
        const costPriceData = await getCostPriceData(productsTokenDoc.apiKey, savedCostPrices);

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
        const { fetchCampaignSkus, fetchFinancialData } = await import('@/app/lib/finance-utils');

        console.log('🚀 Начало создания листа "Финансы РК"...');
        const financeStartTime = Date.now();

        // Получаем финансовые данные с буферными днями (содержат названия кампаний)
        const financialData = await fetchFinancialData(financeTokenDoc.apiKey, startDate, endDate);

        console.log(`📊 Проверка данных для листа "Финансы РК". Финансовых записей: ${financialData.length}`);

        if (financialData.length > 0) {
          console.log("🚀 Создание листа 'Финансы РК'...");
          
          // Собираем уникальные кампании из финансовых данных
          const uniqueCampaigns = new Map();
          financialData.forEach(record => {
            if (!uniqueCampaigns.has(record.advertId)) {
              uniqueCampaigns.set(record.advertId, {
                advertId: record.advertId,
                name: record.campName || 'Неизвестная кампания'
              });
            }
          });
          
          console.log(`📊 Найдено уникальных кампаний: ${uniqueCampaigns.size}`);
          
          // Получаем детальные данные артикулов для столбца I
          const campaignsArray = Array.from(uniqueCampaigns.values());
          console.log(`📊 Получение детальных данных артикулов для кампаний...`);
          const skusMap = await fetchCampaignSkus(financeTokenDoc.apiKey, campaignsArray);
          console.log(`📊 Получено детальных данных артикулов: ${skusMap.size} кампаний`);
          
          // Подготавливаем данные для листа "Финансы РК"
          const financeExcelData = financialData.map(record => {
            return {
              "ID кампании": record.advertId,
              "Название кампании": record.campName || 'Неизвестная кампания',
              "Дата": record.date,
              "Сумма": record.sum,
              "Источник списания": record.bill === 1 ? 'Счет' : 'Баланс',
              "Тип операции": record.type,
              "Номер документа": record.docNumber,
              "SKU ID": skusMap.get(record.advertId) || 'Нет данных'
            };
          });

          headers = [
            'ID кампании',
            'Название кампании', 
            'Дата',
            'Сумма',
            'Источник списания',
            'Тип операции',
            'Номер документа',
            'SKU ID',
            'Период отчета'
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
            { wch: 15 }, // Дата
            { wch: 15 }, // Сумма
            { wch: 20 }, // Источник списания
            { wch: 15 }, // Тип операции
            { wch: 20 }, // Номер документа
            { wch: 40 }, // SKU ID (широкий столбец для списка артикулов)
            { wch: 25 }  // Период отчета
          ];
          
          headers.forEach((header, index) => {
            const column = worksheet.getColumn(index + 1);
            column.width = financeColumnWidths[index]?.wch || Math.max(header.length + 5, 15);
          });

          // Добавляем данные
          const reportPeriod = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;
          financeExcelData.forEach(item => {
            worksheet.addRow([
              item["ID кампании"],
              item["Название кампании"],
              item["Дата"],
              item["Сумма"],
              item["Источник списания"],
              item["Тип операции"],
              item["Номер документа"],
              item["SKU ID"],
              reportPeriod
            ]);
          });

          // Форматируем числовые колонки для отчета финансов РК в российском формате
          const financeNumericColumns = [4]; // Сумма
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
            'Дата',
            'Сумма',
            'Источник списания',
            'Тип операции',
            'Номер документа',
            'SKU ID',
            'Период отчета'
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

        console.log('�� Начало создания отчёта платной приёмки...');
        const acceptanceStart = Date.now();
        fileName = `Платная приёмка - ${startDate}–${endDate}.xlsx`;

        try {
          // 1. Создаём задачу на генерацию
          const createUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report?dateFrom=${startDate}&dateTo=${endDate}`;
          const createResp = await fetch(createUrl, {
            method: 'GET',
            headers: { Authorization: acceptanceTokenDoc.apiKey }
          });
          if (!createResp.ok) throw new Error(`Ошибка создания задачи: ${createResp.status}`);
          const { data: { taskId } = { taskId: null } } = await createResp.json();
          if (!taskId) throw new Error('taskId не получен');
          console.log(`📋 Задача создана: ${taskId}`);

          // 2. Ожидаем готовности
          let ready = false; let attempts = 0;
          while (!ready && attempts < 30) {
            attempts++;
            const statusUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/status`;
            const st = await fetch(statusUrl, { headers: { Authorization: acceptanceTokenDoc.apiKey } });
            if (st.ok) {
              const stData = await st.json();
              const status = stData.data?.status;
              console.log(`🔄 Статус: ${status}`);
              if (status === 'done') { ready = true; break; }
              if (status === 'error') throw new Error('Ошибка генерации на стороне WB');
            }
            if (!ready) await new Promise(r => setTimeout(r, 5000));
          }
          if (!ready) throw new Error('Превышено время ожидания отчёта');

          // 3. Скачиваем отчёт
          const dlUrl = `https://seller-analytics-api.wildberries.ru/api/v1/acceptance_report/tasks/${taskId}/download`;
          const dlResp = await fetch(dlUrl, { headers: { Authorization: acceptanceTokenDoc.apiKey } });
          if (!dlResp.ok) throw new Error(`Ошибка скачивания: ${dlResp.status}`);
          const accData = await dlResp.json();
          console.log(`📦 Записей платной приёмки: ${accData.length}`);

          if (Array.isArray(accData) && accData.length) {
            headers = [
              'Кол-во',
              'Дата создания GI',
              'Income ID',
              'Артикул WB',
              'Дата создания ШК',
              'Предмет',
              'Сумма (руб)',
              'Дата отчета'
            ];
            worksheet.addRow(headers);
            const hr = worksheet.getRow(1);
            hr.font = { bold: true };
            hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };

            const periodStr = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;

            accData.forEach((item: { count?: number; giCreateDate?: string; incomeId?: number; nmID?: number; shkCreateDate?: string; subjectName?: string; total?: number; }) => {
              worksheet.addRow([
                item.count || 0,
                item.giCreateDate ? new Date(item.giCreateDate).toLocaleDateString('ru-RU') : '',
                item.incomeId || '',
                item.nmID || '',
                item.shkCreateDate ? new Date(item.shkCreateDate).toLocaleDateString('ru-RU') : '',
                item.subjectName || '',
                item.total || 0,
                periodStr
              ]);
            });

            // формат числовых столбцов
            [1,7].forEach(col => {
              const c = worksheet.getColumn(col);
              c.eachCell((cell, row) => { if (row>1) { cell.numFmt='[$-419]# ##0,00'; cell.value=Number(cell.value)||0;} });
            });

            headers.forEach((h,i)=>{ worksheet.getColumn(i+1).width=Math.max(h.length+5,15); });
          } else {
            headers = ['Кол-во','Дата создания GI','Income ID','Артикул WB','Дата создания ШК','Предмет','Сумма (руб)','Дата отчета'];
            worksheet.addRow(headers);
            worksheet.getCell('A1').font={bold:true};
          }
          console.log(`✅ Отчёт платной приёмки готов за ${Date.now()-acceptanceStart}ms`);
        } catch(err: unknown) {
          console.error('❌ Ошибка платной приёмки:', err);
          const message = err instanceof Error ? err.message : 'неизвестно';
          return NextResponse.json({ error: `Ошибка при получении данных платной приёмки: ${message}` }, { status: 500 });
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

        console.log('🚀 Начало создания отчета платного хранения...');
        const storageStartTime = Date.now();

        fileName = `Платное хранение - ${startDate}–${endDate}.xlsx`;

        try {
          // Шаг 1: Создаем задание на генерацию отчета
          console.log('📋 Создание задания на генерацию отчета платного хранения...');
          const createTaskUrl = `https://seller-analytics-api.wildberries.ru/api/v1/paid_storage?dateFrom=${startDate}&dateTo=${endDate}`;
          
          const createTaskResponse = await fetch(createTaskUrl, {
            method: 'GET',
            headers: {
              'Authorization': storageTokenDoc.apiKey,
              'Content-Type': 'application/json'
            }
          });

          if (!createTaskResponse.ok) {
            if (createTaskResponse.status === 401) {
              throw new Error('Неверный или недействительный API токен');
            }
            throw new Error(`Ошибка создания задания: ${createTaskResponse.status}`);
          }

          const taskResponse = await createTaskResponse.json();
          const taskId = taskResponse.data?.taskId;
          
          if (!taskId) {
            throw new Error('Не получен ID задания');
          }

          console.log(`📋 Задание создано с ID: ${taskId}`);

          // Шаг 2: Ждем готовности отчета (проверяем статус)
          console.log('⏳ Ожидание готовности отчета...');
          let attempts = 0;
          const maxAttempts = 30; // 30 попыток по 5 секунд = 2.5 минуты максимум
          let isReady = false;

          while (!isReady && attempts < maxAttempts) {
            attempts++;
            console.log(`🔄 Проверка статуса, попытка ${attempts}/${maxAttempts}...`);
            
            const statusUrl = `https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/status`;
            const statusResponse = await fetch(statusUrl, {
              method: 'GET',
              headers: {
                'Authorization': storageTokenDoc.apiKey,
                'Content-Type': 'application/json'
              }
            });

            if (statusResponse.ok) {
              const statusData = await statusResponse.json();
              console.log(`📊 Статус задания: ${statusData.data?.status}`);
              
              if (statusData.data?.status === 'done') {
                isReady = true;
                break;
              } else if (statusData.data?.status === 'error') {
                throw new Error('Ошибка при генерации отчета на сервере WB');
              }
            }

            if (!isReady && attempts < maxAttempts) {
              await new Promise(resolve => setTimeout(resolve, 5000)); // Ждем 5 секунд
            }
          }

          if (!isReady) {
            throw new Error('Превышено время ожидания генерации отчета (2.5 минуты)');
          }

          // Шаг 3: Получаем готовый отчет
          console.log('📥 Скачивание готового отчета...');
          const downloadUrl = `https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/download`;
          const downloadResponse = await fetch(downloadUrl, {
            method: 'GET',
            headers: {
              'Authorization': storageTokenDoc.apiKey,
              'Content-Type': 'application/json'
            }
          });

          if (!downloadResponse.ok) {
            throw new Error(`Ошибка скачивания отчета: ${downloadResponse.status}`);
          }

          const storageData = await downloadResponse.json();
          console.log(`📊 Получено записей платного хранения: ${storageData.length || 0}`);

          if (storageData && storageData.length > 0) {
            headers = [
              'Дата',
              'Артикул WB',
              'Артикул поставщика',
              'Предмет',
              'Бренд',
              'Размер',
              'Объем (л)',
              'Склад',
              'Коэффициент склада',
              'Логистический коэффициент',
              'Стоимость хранения (руб)',
              'Количество баркодов',
              'Тип расчета'
            ];

            // Добавляем заголовки
            worksheet.addRow(headers);
            const headerRow = worksheet.getRow(1);
            headerRow.font = { bold: true };
            headerRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFE0E0E0' }
            };

            // Добавляем данные согласно новой структуре API
            storageData.forEach((item: {
              date?: string;
              nmId?: number;
              vendorCode?: string;
              subject?: string;
              brand?: string;
              size?: string;
              volume?: number;
              warehouse?: string;
              warehouseCoef?: number;
              logWarehouseCoef?: number;
              warehousePrice?: number;
              barcodesCount?: number;
              calcType?: string;
            }) => {
              worksheet.addRow([
                item.date ? new Date(item.date).toLocaleDateString('ru-RU') : '',
                item.nmId || '',
                item.vendorCode || '',
                item.subject || '',
                item.brand || '',
                item.size || '',
                item.volume || 0,
                item.warehouse || '',
                item.warehouseCoef || 0,
                item.logWarehouseCoef || 0,
                item.warehousePrice || 0,
                item.barcodesCount || 0,
                item.calcType || ''
              ]);
            });

            // Форматируем числовые колонки в российском формате
            const numericColumns = [7, 9, 10, 11, 12]; // Объем, Коэффициенты, Стоимость, Количество
            numericColumns.forEach(columnIndex => {
              const column = worksheet.getColumn(columnIndex);
              column.eachCell((cell, rowNumber) => {
                if (rowNumber > 1) {
                  cell.numFmt = '[$-419]# ##0,00';
                  cell.value = Number(cell.value) || 0;
                }
              });
            });

            // Настройка ширины колонок
            headers.forEach((header, index) => {
              const column = worksheet.getColumn(index + 1);
              column.width = Math.max(header.length + 5, 15);
            });

            // Добавляем столбец "Дата отчета" в O1
            console.log("📝 Добавляем столбец 'Дата отчета' в O...");
            worksheet.getCell('O1').value = 'Дата отчета';
            
            // Добавляем период отчета для всех строк данных
            for (let rowIndex = 2; rowIndex <= storageData.length + 1; rowIndex++) {
              const cell = worksheet.getCell(`O${rowIndex}`);
              cell.value = `${new Date(startDate).toLocaleDateString('ru-RU')} - ${new Date(endDate).toLocaleDateString('ru-RU')}`;
            }
          } else {
            // Пустой шаблон
            headers = [
              'Дата',
              'Артикул WB',
              'Артикул поставщика',
              'Предмет',
              'Бренд',
              'Размер',
              'Объем (л)',
              'Склад',
              'Коэффициент склада',
              'Логистический коэффициент',
              'Стоимость хранения (руб)',
              'Количество баркодов',
              'Тип расчета'
            ];

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

            // Добавляем столбец "Дата отчета" в O1 для пустого шаблона
            worksheet.getCell('O1').value = 'Дата отчета';
          }

          console.log(`✅ Отчет платного хранения создан за ${Date.now() - storageStartTime}ms`);
        } catch (error) {
          console.error('❌ Ошибка при получении данных платного хранения:', error);
          return NextResponse.json({ 
            error: `Ошибка при получении данных платного хранения: ${error instanceof Error ? error.message : 'Неизвестная ошибка'}` 
          }, { status: 500 });
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