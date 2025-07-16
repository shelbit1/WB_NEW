import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { reportType, clientName, startDate, endDate } = body;

    // Создаем новую рабочую книгу
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Отчёт');

    // Добавляем заголовки в зависимости от типа отчёта
    let headers: string[] = [];
    let fileName = '';

    switch (reportType) {
      case 'details':
        headers = ['Дата', 'Артикул', 'Наименование', 'Количество', 'Сумма'];
        fileName = `Отчет детализации - ${startDate}–${endDate}.xlsx`;
        break;
      case 'storage':
        headers = ['Дата', 'Артикул', 'Склад', 'Дни хранения', 'Стоимость'];
        fileName = `Платное хранение - ${startDate}–${endDate}.xlsx`;
        break;
      case 'acceptance':
        headers = ['Дата', 'Артикул', 'Количество', 'Стоимость приёмки'];
        fileName = `Платная приемка - ${startDate}–${endDate}.xlsx`;
        break;
      case 'products':
        headers = ['Артикул', 'Наименование', 'Бренд', 'Категория', 'Цена'];
        fileName = `Список товаров - ${startDate}–${endDate}.xlsx`;
        break;
      case 'finances':
        headers = ['Дата', 'Тип операции', 'Сумма', 'К доплате', 'К перечислению'];
        fileName = `Финансы РК - ${startDate}–${endDate}.xlsx`;
        break;
      default:
        return NextResponse.json({ error: 'Неизвестный тип отчёта' }, { status: 400 });
    }

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
      column.width = Math.max(header.length + 5, 15);
    });

    // Добавляем пример данных (в будущем здесь будет реальная интеграция с API WB)
    if (reportType === 'details') {
      worksheet.addRow(['01.01.2025', 'WB12345', 'Пример товара', '10', '1500']);
    } else if (reportType === 'finances') {
      worksheet.addRow(['01.01.2025', 'Продажа', '1500', '0', '1275']);
    }
    // Добавить другие примеры данных для других типов отчётов...

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
      { error: 'Ошибка при генерации отчёта' },
      { status: 500 }
    );
  }
} 