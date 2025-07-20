/**
 * Утилиты для работы с данными реализации Wildberries
 */

// Полный интерфейс данных реализации (детализации)
export interface RealizationDetailItem {
  realizationreport_id?: string;
  date_from?: string;
  date_to?: string;
  create_dt?: string;
  currency_name?: string;
  suppliercontract_code?: string;
  rrd_id?: string;
  gi_id?: string;
  dlv_prc?: number;
  fix_tariff_date_from?: string;
  fix_tariff_date_to?: string;
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
  payment_processing?: string;
  acquiring_bank?: string;
  ppvz_vw?: number;
  ppvz_vw_nds?: number;
  ppvz_office_name?: string;
  ppvz_office_id?: string;
  ppvz_supplier_id?: string;
  ppvz_supplier_name?: string;
  ppvz_inn?: string;
  declaration_number?: string;
  bonus_type_name?: string;
  sticker_id?: string;
  site_country?: string;
  srv_dbs?: boolean;
  penalty?: number;
  additional_payment?: number;
  rebill_logistic_cost?: number;
  rebill_logistic_org?: string;
  storage_fee?: number;
  deduction?: number;
  acceptance?: number;
  assembly_id?: string;
  kiz?: string;
  srid?: string;
  report_type?: number;
  is_legal_entity?: boolean;
  trbx_id?: string;
  installment_cofinancing_amount?: number;
  wibes_wb_discount_percent?: number;
}

/**
 * Функция получения данных реализации (детализации) с API Wildberries
 * @param apiKey - API токен Wildberries
 * @param dateFrom - Дата начала в формате YYYY-MM-DD
 * @param dateTo - Дата окончания в формате YYYY-MM-DD
 * @returns Массив данных детализации реализации
 */
export async function fetchRealizationData(apiKey: string, dateFrom: string, dateTo: string): Promise<RealizationDetailItem[]> {
  // Проверка периода - максимум 31 день
  const startDate = new Date(dateFrom);
  const endDate = new Date(dateTo);
  const diffTime = Math.abs(endDate.getTime() - startDate.getTime());
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays > 31) {
    console.warn(`⚠️ Период ${diffDays} дней превышает лимит 31 день для API реализации`);
    throw new Error(`Период не должен превышать 31 день. Текущий период: ${diffDays} дней`);
  }

  try {
    console.log('🔄 ЭТАП 1: Создание задачи получения данных реализации...');
    
    // 1. Получаем данные детализации реализации напрямую
    // ВАЖНО: Правильный URL API Wildberries для получения детализации реализации
    const cleanToken = apiKey.startsWith('Bearer ') ? apiKey : apiKey;
    const apiUrl = `https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=${dateFrom}&dateTo=${dateTo}`;
    console.log(`   🌐 URL API реализации: ${apiUrl}`);
    
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Authorization': cleanToken,
        'User-Agent': 'Mozilla/5.0 (compatible; WB-API-Client/1.0)',
        'Accept': 'application/json',
        'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
      }
    });

    console.log(`📊 Ответ API реализации: ${response.status} ${response.statusText}`);
    console.log(`📋 Заголовки ответа:`, Object.fromEntries(response.headers.entries()));
    console.log(`📊 Content-Type: ${response.headers.get('content-type')}`);
    console.log(`📊 Response Size: ${response.headers.get('content-length')} bytes`);
    
    if (!response.ok) {
      console.error(`❌ Ошибка API реализации: ${response.status} ${response.statusText}`);
      
      let errorMessage = `Ошибка API Wildberries: ${response.status}`;
      let userHelp = '';
      
      if (response.status === 401) {
        errorMessage = "Неверный или недействительный API токен";
        userHelp = "Проверьте токен в личном кабинете Wildberries в разделе 'Настройки → Доступ к API'";
      } else if (response.status === 403) {
        errorMessage = "Нет доступа к данным";
        userHelp = "Проверьте права токена. Токен должен иметь доступ к статистике.";
      } else if (response.status === 429) {
        errorMessage = "Превышен лимит запросов";
        userHelp = "Попробуйте позже. Wildberries ограничивает количество запросов в минуту.";
      } else if (response.status === 400) {
        errorMessage = "Некорректный запрос";
        userHelp = "Проверьте период дат. Максимальный период для отчета - 30 дней.";
      } else if (response.status === 500) {
        errorMessage = "Внутренняя ошибка сервера Wildberries";
        userHelp = "Попробуйте позже. Проблема на стороне Wildberries.";
      } else if (response.status === 404) {
        errorMessage = "API endpoint не найден";
        userHelp = "Возможно, API Wildberries изменился. Обратитесь в техподдержку.";
      }
      
      const errorText = await response.text();
      console.log(`❌ Детали ошибки: ${errorText}`);
      console.log(`💡 Рекомендация: ${userHelp}`);
      
      throw new Error(`${errorMessage}. ${userHelp}. Ответ сервера: ${errorText}`);
    }

    // 2. Получаем данные напрямую из ответа API
    const data = await response.json();
    console.log(`📊 Размер полученных данных: ${JSON.stringify(data).length} символов`);
    console.log(`📋 Тип данных: ${typeof data}, является массивом: ${Array.isArray(data)}`);
    
    // КРИТИЧЕСКИ ВАЖНАЯ ПРОВЕРКА
    if (!Array.isArray(data) || data.length === 0) {
      console.warn("⚠️ Нет данных реализации за указанный период");
      console.log(`📊 Тип полученных данных: ${typeof data}`);
      console.log(`📊 Является массивом: ${Array.isArray(data)}`);
      console.log(`📊 Данные:`, JSON.stringify(data, null, 2));
      // ЭТО ОСНОВНАЯ ПРИЧИНА ПУСТОГО ЛИСТА!
      return [];
    }
    
    // Диагностика структуры данных
    if (data.length > 0) {
      console.log("🔍 Первая запись данных:", JSON.stringify(data[0], null, 2));
      console.log("🏷️ Поля записи:", Object.keys(data[0] || {}));
      
      // Проверяем ключевые поля
      const firstItem = data[0];
      if (firstItem.realizationreport_id) {
        console.log(`📋 Найдены данные с realizationreport_id: ${firstItem.realizationreport_id}`);
      } else {
        console.warn("⚠️ В данных отсутствует поле realizationreport_id!");
      }
    }
    
    console.log(`✅ Успешно получено ${data.length} записей детализации реализации`);
    return data;
    
  } catch (error) {
    console.error('❌ Ошибка при получении данных реализации:', error);
    throw error;
  }
}

/**
 * Функция для добавления данных детализации реализации в ExcelJS workbook с группировкой по realizationreport_id
 * @param workbook - рабочая книга ExcelJS
 * @param data - данные реализации для добавления
 */
export function addDetailedRealizationToWorkbook(workbook: any, data: RealizationDetailItem[]): void {
  console.log("📊 Создание листа 'Полный отчет' с разделением по realizationreport_id...");
  console.log(`📦 Входящие данные: ${data.length} записей`);
  
  // Диагностика структуры данных
  if (data.length > 0) {
    console.log("🔍 Первая запись данных:", JSON.stringify(data[0], null, 2));
    console.log("🏷️ Поля записи:", Object.keys(data[0] || {}));
  }
  
  // Группируем данные по realizationreport_id
  const reportGroups: { [key: string]: RealizationDetailItem[] } = {};
  
  data.forEach(item => {
    const reportId = item.realizationreport_id || 'Без ID';
    if (!reportGroups[reportId]) {
      reportGroups[reportId] = [];
    }
    reportGroups[reportId].push(item);
  });
  
  console.log(`📋 Найдено ${Object.keys(reportGroups).length} уникальных отчетов реализации`);
  
  // Диагностика групп
  Object.keys(reportGroups).forEach(reportId => {
    console.log(`   - Отчет ${reportId}: ${reportGroups[reportId].length} записей`);
  });
  
  if (Object.keys(reportGroups).length === 0) {
    console.error("❌ ПРОБЛЕМА: Нет групп отчетов для создания листа!");
    return;
  }
  
  // Создаем лист
  const worksheet = workbook.addWorksheet('Отчет детализации');
  
  // Определяем заголовки колонок
  const headers = [
    'Номер отчёта', 'Дата начала отчётного периода', 'Дата конца отчётного периода', 
    'Дата формирования отчёта', 'Валюта', 'Код договора поставщика', 'ID записи', 
    'Номер поставки', 'Процент логистики', 'Дата фиксации тарифа с', 
    'Дата фиксации тарифа по', 'Предмет', 'Артикул WB', 'Бренд', 'Артикул продавца', 
    'Размер', 'Баркод', 'Тип документа', 'Количество', 'Розничная цена', 
    'Сумма продаж', 'Согласованная скидка (%)', 'Процент комиссии', 'Склад', 
    'Обоснование для оплаты', 'Дата заказа', 'Дата продажи', 'Дата отчета', 
    'Штрихкод', 'Цена розничная с учетом согласованной скидки', 'Количество доставок', 
    'Количество возвратов', 'Стоимость логистики', 'Тип коробки', 
    'Скидка товара для отчета', 'Промо от поставщика', 'Rid', 'SPP процент', 
    'КВВ процент базовый', 'КВВ процент', 'Процент повышения рейтинга', 
    'Флаг KGVP v2', 'Комиссия за продажи', 'К доплате', 'Вознаграждение', 
    'Комиссия эквайринга', 'Процент эквайринга', 'Обработка платежей', 
    'Банк эквайринга', 'PPVZ VW', 'PPVZ VW НДС', 'Название офиса PPVZ', 
    'ID офиса PPVZ', 'ID поставщика PPVZ', 'Имя поставщика PPVZ', 'ИНН PPVZ', 
    'Номер декларации', 'Тип бонуса', 'ID стикера', 'Страна сайта', 'Флаг DBS', 
    'Штраф', 'Доплата', 'Перерасчет логистики', 'Организация перерасчета', 
    'Стоимость хранения', 'Удержания', 'Приемка', 'ID сборочного задания', 
    'КИЗ', 'SRID', 'Тип отчета', 'Юридическое лицо', 'ID TRBX', 
    'Сумма рассрочки', 'Процент скидки WB'
  ];
  
  // Добавляем заголовки
  worksheet.addRow(headers);
  
  // Стилизуем заголовки
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };
  
  let rowIndex = 2;
  
  // Добавляем данные по группам
  Object.keys(reportGroups).forEach((reportId, groupIndex) => {
    const group = reportGroups[reportId];
    
    // Добавляем пустую строку для разделения между группами (кроме первой)
    if (groupIndex > 0) {
      worksheet.addRow([]);
      rowIndex++;
    }
    
    // Добавляем заголовок группы
    const groupHeaderRow = worksheet.addRow([
      `=== ОТЧЕТ РЕАЛИЗАЦИИ ID: ${reportId} ===`,
      group[0]?.date_from || '',
      group[0]?.date_to || '',
      group[0]?.create_dt || '',
      group[0]?.currency_name || '',
      '',
      `Количество записей: ${group.length}`,
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]);
    
    // Стилизуем заголовок группы
    groupHeaderRow.font = { bold: true, color: { argb: 'FF0066CC' } };
    groupHeaderRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF0F8FF' }
    };
    rowIndex++;
    
    // Добавляем все записи группы
    group.forEach(item => {
      // Вспомогательная функция получения числового значения
      const num = (val: any) => {
        if (val === undefined || val === null || val === '') return 0;
        return typeof val === 'number' ? val : parseFloat(val.toString().replace(',', '.')) || 0;
      };

      // Функция для удаления точек из строковых значений
      const removeDots = (val: any) => {
        if (val === undefined || val === null || val === '') return "";
        return val.toString().replace(/\./g, '');
      };

      worksheet.addRow([
        removeDots(item.realizationreport_id),  // Столбец A - убираем точки
        item.date_from || "",
        item.date_to || "",
        item.create_dt || "",
        item.currency_name || "",
        item.suppliercontract_code || "",
        removeDots(item.rrd_id),                // Столбец G - убираем точки
        removeDots(item.gi_id),                 // Столбец H - убираем точки
        num(item.dlv_prc),
        item.fix_tariff_date_from || "",
        item.fix_tariff_date_to || "",
        item.subject_name || "",
        num(item.nm_id),
        item.brand_name || "",
        item.sa_name || "",
        item.ts_name || "",
        item.barcode || "",
        item.doc_type_name || "",
        num(item.quantity),
        num(item.retail_price),
        num(item.retail_amount),
        num(item.sale_percent),
        num(item.commission_percent),
        item.office_name || "",
        item.supplier_oper_name || "",
        item.order_dt || "",
        item.sale_dt || "",
        item.rr_dt || "",
        removeDots(item.shk_id),                // Столбец AC - убираем точки
        num(item.retail_price_withdisc_rub),
        num(item.delivery_amount),
        num(item.return_amount),
        num(item.delivery_rub),
        item.gi_box_type_name || "",
        num(item.product_discount_for_report),
        num(item.supplier_promo),
        item.rid || "",
        num(item.ppvz_spp_prc),
        num(item.ppvz_kvw_prc_base),
        num(item.ppvz_kvw_prc),
        num(item.sup_rating_prc_up),
        num(item.is_kgvp_v2),
        num(item.ppvz_sales_commission),
        num(item.ppvz_for_pay),
        num(item.ppvz_reward),
        num(item.acquiring_fee),
        num(item.acquiring_percent),
        item.payment_processing || "",
        item.acquiring_bank || "",
        num(item.ppvz_vw),
        num(item.ppvz_vw_nds),
        item.ppvz_office_name || "",
        item.ppvz_office_id || "",
        item.ppvz_supplier_id || "",
        item.ppvz_supplier_name || "",
        item.ppvz_inn || "",
        item.declaration_number || "",
        item.bonus_type_name || "",
        item.sticker_id || "",
        item.site_country || "",
        item.srv_dbs ? "Да" : "Нет",
        num(item.penalty),
        num(item.additional_payment),
        num(item.rebill_logistic_cost),
        item.rebill_logistic_org || "",
        num(item.storage_fee),
        num(item.deduction),
        num(item.acceptance),
        item.assembly_id || "",
        item.kiz || "",
        item.srid || "",
        num(item.report_type),
        item.is_legal_entity ? "Да" : "Нет",
        item.trbx_id || "",
        num(item.installment_cofinancing_amount),
        num(item.wibes_wb_discount_percent)
      ]);
      rowIndex++;
    });
  });
  
  // Настраиваем ширину колонок - все по 15 символов
  headers.forEach((header, index) => {
    const column = worksheet.getColumn(index + 1);
    column.width = 15;
  });
  
     // Форматируем числовые колонки в российском формате (пробел как разделитель тысяч, запятая как десятичный разделитель)
   const numericColumns = [9, 19, 20, 21, 22, 23, 30, 31, 32, 33, 35, 36, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 49, 50, 51, 58, 60, 61, 62, 64, 65, 66, 67, 70, 72, 74, 75];
   numericColumns.forEach(columnIndex => {
     const column = worksheet.getColumn(columnIndex);
     column.eachCell((cell: any, rowNumber: number) => {
       if (rowNumber > 1) { // Пропускаем заголовок
         // Российский формат с точной спецификацией
         cell.numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00';
       }
     });
   });

  // После того как все данные добавлены — обеспечиваем формат чисел (минимум 2 знака)
  worksheet.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    if (rowNumber === 1) return; // пропускаем заголовок
    row.eachCell((cell: any) => {
      if (typeof cell.value === 'number') {
        cell.numFmt = '[$-419]# ##0,00;[$-419]-# ##0,00';
      }
    });
  });
  
  console.log(`✅ Лист детализации реализации создан с ${rowIndex - 1} строками данных`);
  console.log(`📊 Статистика по отчетам реализации:`);
  Object.keys(reportGroups).forEach(reportId => {
    console.log(`   - Отчет ${reportId}: ${reportGroups[reportId].length} записей`);
  });
} 