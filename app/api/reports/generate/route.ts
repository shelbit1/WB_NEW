import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import connectToDatabase from '@/app/lib/mongodb';
import Token from '@/app/lib/models/Token';
import { getStorageData, StorageItem } from '@/app/lib/storage-utils';

interface DetailReportItem {
  rrd_id: number;
  realizationreport_id: number;
  date_from: string;
  date_to: string;
  create_dt: string;
  currency_name: string;
  suppliercontract_code: string;
  rr_dt: string;
  gi_id: number;
  subject_name: string;
  nm_id: number;
  brand_name: string;
  sa_name: string;
  ts_name: string;
  barcode: string;
  doc_type_name: string;
  quantity: number;
  retail_price: number;
  retail_amount: number;
  sale_percent: number;
  commission_percent: number;
  office_name: string;
  supplier_oper_name: string;
  order_dt: string;
  sale_dt: string;
  shk_id: number;
  retail_price_withdisc_rub: number;
  delivery_amount: number;
  return_amount: number;
  delivery_rub: number;
  gi_box_type_name: string;
  product_discount_for_report: number;
  supplier_promo: number;
  rid: number;
  ppvz_spp_prc: number;
  ppvz_kvw_prc_base: number;
  ppvz_kvw_prc: number;
  sup_rating_prc_up: number;
  is_kgvp_v2: number;
  ppvz_sales_commission: number;
  ppvz_for_pay: number;
  ppvz_reward: number;
  acquiring_fee: number;
  acquiring_bank: string;
  ppvz_vw: number;
  ppvz_vw_nds: number;
  ppvz_office_id: number;
  ppvz_office_name: string;
  ppvz_supplier_id: number;
  ppvz_supplier_name: string;
  ppvz_inn: string;
  declaration_number: string;
  bonus_type_name: string;
  sticker_id: string;
  site_country: string;
  penalty: number;
  additional_payment: number;
  srid: string;
  
  // Дополнительные поля из полного API ответа
  dlv_prc: number;
  fix_tariff_date: string;
  retail_commission: number;
  sale_commission: number;
  gi_box_code: string;
  gi_box_count: number;
  product_discount: number;
  product_discount_rub: number;
  kiz: string;
  spp: number;
  finished_price: number;
  price_with_disc: number;
  nm_id_old: number;
  bonus_percent: number;
  wb_discount: number;
  customer_reward: number;
  marketplace: string;
  inc_id: number;
  order_id: number;
  gi_sku_id: number;
  country_name: string;
  oblast_okrug_name: string;
  region_name: string;
  income_id: number;
  odid: number;
  spp_office: number;
  for_pay_nds: number;
  ppvz_inn_nds: number;
  ppvz_office_address: string;
  ppvz_office_name_full: string;
  sticker_number: string;
  gi_barcode: string;
  gi_id_old: number;
  posting_number: string;
  imei: string;
  box_type_name: string;
  cluster_from: string;
  cluster_to: string;
  nm_id_supplier: string;
  clusterization_charge: number;
  clusterization_bonus: number;
  consolidation_charge: number;
  consolidated_goods_qty: number;
  consolidated_coefficient: number;
  consolidated_handling_fee: number;
  consolidated_storage_fee: number;
  payment_method: string;
  payment_date: string;
  bank_percent_fee: number;
  insurance_premium: number;
  insurance_discount: number;
  return_shipping_charge: number;
  return_not_selling_fee: number;
  acceptance_charge: number;
  placing_charge: number;
  additional_charge: number;
  storage_charge: number;
  deduction_charge: number;
  incorrect_attachments_charge: number;
  redirection_charge: number;
  shipping_charge: number;
  fulfillment_charge: number;
  acquiring_percent: number;
  tax_rate: number;
  vat_rate: number;
  total_payment: number;
  currency_code: string;
  exchange_rate: number;
  created_at: string;
  updated_at: string;
}

async function fetchDetailReport(apiKey: string, dateFrom: string, dateTo: string): Promise<DetailReportItem[]> {
  const url = new URL('https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod');
  url.searchParams.append('dateFrom', dateFrom);
  url.searchParams.append('dateTo', dateTo);
  url.searchParams.append('limit', '100000');

  const response = await fetch(url.toString(), {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Ошибка API Wildberries: ${response.status} ${response.statusText}`);
  }

  return await response.json();
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
        const tokenDoc = await Token.findById(tokenId);
        
        if (!tokenDoc) {
          return NextResponse.json({ error: 'Токен не найден' }, { status: 404 });
        }

        // Получаем данные с API Wildberries
        const detailData = await fetchDetailReport(tokenDoc.apiKey, startDate, endDate);

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
          'SRID',
          
          // Дополнительные поля
          'Процент доставки',
          'Дата фикс тарифа',
          'Комиссия розничная',
          'Комиссия продаж',
          'Код коробки',
          'Количество коробок',
          'Скидка товара',
          'Скидка товара руб',
          'КИЗ',
          'СПП',
          'Финальная цена',
          'Цена со скидкой',
          'Старый артикул WB',
          'Процент бонуса',
          'Скидка WB',
          'Вознаграждение покупателя',
          'Маркетплейс',
          'ID поступления',
          'ID заказа',
          'SKU ID',
          'Страна',
          'Область/округ',
          'Регион',
          'ID поступления',
          'ODID',
          'Офис СПП',
          'К оплате НДС',
          'ИНН НДС',
          'Адрес офиса',
          'Полное название офиса',
          'Номер стикера',
          'Штрихкод поставки',
          'Старый ID поставки',
          'Номер отправления',
          'IMEI',
          'Тип коробки',
          'Кластер откуда',
          'Кластер куда',
          'Артикул поставщика',
          'Плата за кластеризацию',
          'Бонус кластеризации',
          'Плата за консолидацию',
          'Количество консолидированного товара',
          'Коэффициент консолидации',
          'Плата за обработку консолидации',
          'Плата за хранение консолидации',
          'Способ оплаты',
          'Дата оплаты',
          'Процент банка',
          'Страховая премия',
          'Страховая скидка',
          'Плата за возвратную доставку',
          'Плата за не продажу возврата',
          'Плата за приемку',
          'Плата за размещение',
          'Дополнительная плата',
          'Плата за хранение',
          'Плата за удержание',
          'Плата за неправильные вложения',
          'Плата за перенаправление',
          'Плата за доставку',
          'Плата за фулфилмент',
          'Процент эквайринга',
          'Налоговая ставка',
          'Ставка НДС',
          'Общая оплата',
          'Код валюты',
          'Курс обмена',
          'Дата создания',
          'Дата обновления'
        ];

        fileName = `Отчет детализации - ${startDate}–${endDate}.xlsx`;

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

        // Автоматически подгоняем ширину колонок
        headers.forEach((header, index) => {
          const column = worksheet.getColumn(index + 1);
          column.width = Math.max(header.length / 2, 8);
        });

        // Добавляем данные
        detailData.forEach(item => {
          worksheet.addRow([
            item.rrd_id || '',
            item.realizationreport_id || '',
            item.date_from || '',
            item.date_to || '',
            item.create_dt || '',
            item.currency_name || '',
            item.suppliercontract_code || '',
            item.rr_dt ? new Date(item.rr_dt).toLocaleDateString('ru-RU') : '',
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
            item.acquiring_bank || '',
            item.ppvz_vw || 0,
            item.ppvz_vw_nds || 0,
            item.ppvz_office_id || '',
            item.ppvz_office_name || '',
            item.ppvz_supplier_id || '',
            item.ppvz_supplier_name || '',
            item.ppvz_inn || '',
            item.declaration_number || '',
            item.bonus_type_name || '',
            item.sticker_id || '',
            item.site_country || '',
            item.penalty || 0,
            item.additional_payment || 0,
            item.srid || '',
            
            // Дополнительные поля
            item.dlv_prc || 0,
            item.fix_tariff_date || '',
            item.retail_commission || 0,
            item.sale_commission || 0,
            item.gi_box_code || '',
            item.gi_box_count || 0,
            item.product_discount || 0,
            item.product_discount_rub || 0,
            item.kiz || '',
            item.spp || 0,
            item.finished_price || 0,
            item.price_with_disc || 0,
            item.nm_id_old || '',
            item.bonus_percent || 0,
            item.wb_discount || 0,
            item.customer_reward || 0,
            item.marketplace || '',
            item.inc_id || '',
            item.order_id || '',
            item.gi_sku_id || '',
            item.country_name || '',
            item.oblast_okrug_name || '',
            item.region_name || '',
            item.income_id || '',
            item.odid || '',
            item.spp_office || 0,
            item.for_pay_nds || 0,
            item.ppvz_inn_nds || '',
            item.ppvz_office_address || '',
            item.ppvz_office_name_full || '',
            item.sticker_number || '',
            item.gi_barcode || '',
            item.gi_id_old || '',
            item.posting_number || '',
            item.imei || '',
            item.box_type_name || '',
            item.cluster_from || '',
            item.cluster_to || '',
            item.nm_id_supplier || '',
            item.clusterization_charge || 0,
            item.clusterization_bonus || 0,
            item.consolidation_charge || 0,
            item.consolidated_goods_qty || 0,
            item.consolidated_coefficient || 0,
            item.consolidated_handling_fee || 0,
            item.consolidated_storage_fee || 0,
            item.payment_method || '',
            item.payment_date || '',
            item.bank_percent_fee || 0,
            item.insurance_premium || 0,
            item.insurance_discount || 0,
            item.return_shipping_charge || 0,
            item.return_not_selling_fee || 0,
            item.acceptance_charge || 0,
            item.placing_charge || 0,
            item.additional_charge || 0,
            item.storage_charge || 0,
            item.deduction_charge || 0,
            item.incorrect_attachments_charge || 0,
            item.redirection_charge || 0,
            item.shipping_charge || 0,
            item.fulfillment_charge || 0,
            item.acquiring_percent || 0,
            item.tax_rate || 0,
            item.vat_rate || 0,
            item.total_payment || 0,
            item.currency_code || '',
            item.exchange_rate || 0,
            item.created_at || '',
            item.updated_at || ''
          ]);
        });

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

        // Получаем данные о платном хранении
        const storageData = await getStorageData(storageTokenDoc.apiKey, startDate, endDate);

        console.log(`📊 Проверка данных для отчета "Платное хранение". Записей: ${storageData.length}`);

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
        headers = ['Дата', 'Артикул', 'Количество', 'Стоимость приёмки'];
        fileName = `Платная приемка - ${startDate}–${endDate}.xlsx`;
        
        worksheet.addRow(headers);
        const acceptanceHeaderRow = worksheet.getRow(1);
        acceptanceHeaderRow.font = { bold: true };
        acceptanceHeaderRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE0E0E0' }
        };
        
        headers.forEach((header, index) => {
          const column = worksheet.getColumn(index + 1);
          column.width = Math.max(header.length + 5, 15);
        });
        
        worksheet.addRow(['01.01.2025', 'WB123456', '100', '250']);
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

                 // Получаем данные детализации для товаров из реализации
         let realizationData: DetailReportItem[] = [];
         try {
           realizationData = await fetchDetailReport(productsTokenDoc.apiKey, startDate, endDate);
           console.log(`📊 Получено записей реализации: ${realizationData.length}`);
         } catch (error) {
           console.warn('⚠️ Не удалось получить данные реализации:', error);
         }

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