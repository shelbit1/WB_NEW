// Интерфейс для данных платного хранения согласно документации WB API
export interface StorageItem {
  date?: string;                    // Дата
  logWarehouseCoef?: number;        // Логистический коэффициент склада
  officeId?: number;                // ID офиса
  warehouse?: string;               // Название склада
  warehouseCoef?: number;           // Коэффициент склада
  giId?: number;                    // ID поставки
  chrtId?: number;                  // ID размера товара в WB
  size?: string;                    // Размер товара
  barcode?: string;                 // Штрихкод
  subject?: string;                 // Предмет
  brand?: string;                   // Бренд
  vendorCode?: string;              // Артикул продавца
  nmId?: number;                    // Артикул WB
  volume?: number;                  // Объем товара
  calcType?: string;                // Тип расчета
  warehousePrice?: number;          // Стоимость хранения
  barcodesCount?: number;           // Количество штрихкодов
  palletPlaceCode?: number;         // Код места на паллете
  palletCount?: number;             // Количество паллет
  originalDate?: string;            // Оригинальная дата
  loyaltyDiscount?: number;         // Скидка лояльности
  tariffFixDate?: string;           // Дата фиксации тарифа
  tariffLowerDate?: string;         // Дата снижения тарифа
}

/**
 * Функция получения данных хранения с API ограничениями
 * @param token - API токен Wildberries
 * @param startDate - Дата начала в формате YYYY-MM-DD
 * @param endDate - Дата окончания в формате YYYY-MM-DD
 * @returns Массив данных о платном хранении
 */
export async function getStorageData(token: string, startDate: string, endDate: string): Promise<StorageItem[]> {
  try {
    console.log(`🔍 Запрос данных хранения: ${startDate} - ${endDate}`);
    
    // 1. Проверка периода (максимум 8 дней)
    const start = new Date(startDate);
    const end = new Date(endDate);
    const diffDays = Math.ceil(Math.abs(end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
    
    if (diffDays > 8) {
      console.warn(`⚠️ Период ${diffDays} дней превышает лимит 8 дней для API хранения`);
      // Ограничиваем период до 8 дней
      const limitedEndDate = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
      const limitedEnd = limitedEndDate.toISOString().split('T')[0];
      console.log(`📅 Ограничиваем период до: ${startDate} - ${limitedEnd}`);
      return await getStorageData(token, startDate, limitedEnd);
    }
    
    // 2. Создание задачи
    console.log("📤 Создание задачи для отчета по хранению...");
    const taskUrl = `https://seller-analytics-api.wildberries.ru/api/v1/paid_storage?dateFrom=${startDate}&dateTo=${endDate}`;
    console.log(`📡 URL создания задачи: ${taskUrl}`);
    
    const taskResponse = await fetch(taskUrl, { 
      method: "GET", 
      headers: { 
        "Authorization": token,
        "Content-Type": "application/json"
      } 
    });

    console.log(`📊 Ответ создания задачи хранения: ${taskResponse.status} ${taskResponse.statusText}`);

    // 3. Обработка ошибок создания задачи
    if (!taskResponse.ok) {
      const errorText = await taskResponse.text();
      console.error(`❌ Ошибка создания taskId для хранения: ${taskResponse.status}`, errorText);
      
      if (taskResponse.status === 429) {
        console.warn("⚠️ Превышен лимит запросов (1 запрос в минуту)");
        console.log("⏳ Ожидание 2 минуты перед повторной попыткой...");
        // Ждем 2 минуты при превышении лимита
        await new Promise(resolve => setTimeout(resolve, 120000));
        // Повторная попытка
        return await getStorageData(token, startDate, endDate);
      } else if (taskResponse.status === 401) {
        throw new Error("Неверный токен авторизации");
      }
      throw new Error(`Ошибка создания задачи: ${taskResponse.status}`);
    }
    
    // 4. Получение taskId
    const taskData = await taskResponse.json();
    console.log("📋 Ответ создания задачи:", JSON.stringify(taskData, null, 2));
    
    const taskId = taskData?.data?.taskId;
    if (!taskId) {
      console.error("❌ Не удалось получить taskId для хранения. Структура ответа:", taskData);
      throw new Error("Не удалось получить taskId");
    }
    
    console.log(`✅ Создан taskId для хранения: ${taskId}`);

    // 5. Ожидание готовности отчета
    console.log("⏳ Ожидание подготовки отчета о хранении...");
    const maxStatusChecks = 12; // 12 проверок по 5 секунд = 1 минута
    
    for (let i = 0; i < maxStatusChecks; i++) {
      await new Promise(resolve => setTimeout(resolve, 5000));
      console.log(`🔍 Проверка статуса ${i + 1}/${maxStatusChecks}...`);
      
      const statusResponse = await fetch(`https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/status`, {
        method: "GET",
        headers: { 
          "Authorization": token,
          "Content-Type": "application/json"
        }
      });
      
      if (statusResponse.ok) {
        const statusData = await statusResponse.json();
        console.log(`📊 Статус задачи: ${statusData?.data?.status}`);
        
        if (statusData?.data?.status === "done") {
          console.log("✅ Отчет готов! Загружаем...");
          break;
        }
      } else {
        console.warn(`⚠️ Ошибка проверки статуса: ${statusResponse.status}`);
      }
    }
    
    // 6. Загрузка данных с повторными попытками
    const maxRetries = 5;
    for (let retryCount = 0; retryCount < maxRetries; retryCount++) {
      console.log(`📥 Попытка загрузки отчета ${retryCount + 1}/${maxRetries}...`);
      
      const downloadResponse = await fetch(`https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/download`, {
        method: "GET",
        headers: { 
          "Authorization": token,
          "Content-Type": "application/json"
        }
      });

      console.log(`📊 Ответ загрузки хранения: ${downloadResponse.status} ${downloadResponse.statusText}`);
      
      if (downloadResponse.ok) {
        const responseText = await downloadResponse.text();
        console.log(`📄 Сырой ответ загрузки хранения (длина: ${responseText.length})`);
        
        if (!responseText || responseText.trim().length === 0) {
          console.warn("⚠️ API хранения вернул пустой ответ");
          return [];
        }
        
        let storageData;
        try {
          storageData = JSON.parse(responseText);
        } catch (parseError) {
          console.error("❌ Ошибка парсинга JSON хранения:", parseError);
          console.log("📄 Полный ответ для отладки:", responseText);
          return [];
        }
        
        const dataArray = Array.isArray(storageData) ? storageData : [];
        console.log(`✅ Получено ${dataArray.length} записей о хранении`);
        
        if (dataArray.length > 0) {
          console.log("🔍 Первая запись хранения:", JSON.stringify(dataArray[0], null, 2));
          console.log("🏷️ Поля записи:", Object.keys(dataArray[0] || {}));
        } else {
          console.warn("⚠️ API вернул пустой массив данных хранения");
          console.log("📊 Это может означать:");
          console.log("   - Нет операций хранения за указанный период");
          console.log("   - Неправильный период дат");
          console.log("   - Проблемы с токеном или правами доступа");
        }
        
        return dataArray;
      } else if (downloadResponse.status === 429) {
        // Экспоненциальная задержка: 2, 4, 8 минут
        const delayMinutes = Math.min(2 ** retryCount, 8);
        console.log(`⚠️ Ошибка 429. Ожидание ${delayMinutes} минут...`);
        await new Promise(resolve => setTimeout(resolve, delayMinutes * 60 * 1000));
      } else {
        const errorText = await downloadResponse.text();
        console.error(`❌ Ошибка загрузки отчета о хранении: ${downloadResponse.status}`, errorText);
        
        if (retryCount < maxRetries - 1) {
          console.log(`🔄 Повторная попытка через 10 секунд...`);
          await new Promise(resolve => setTimeout(resolve, 10000));
        }
      }
    }
    
    throw new Error("Не удалось загрузить данные после всех попыток");
  } catch (error) {
    console.error("💥 Критическая ошибка при получении данных о хранении:", error);
    throw error;
  }
} 