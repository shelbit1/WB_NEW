// Интерфейс для данных платного хранения
export interface StorageItem {
  date?: string;
  warehouse?: string;
  nmId?: number;
  size?: string;
  barcode?: string;
  subject?: string;
  brand?: string;
  vendorCode?: string;
  volume?: number;
  calcType?: string;
  warehousePrice?: number;
  barcodesCount?: number;
  warehouseCoef?: number;
  loyaltyDiscount?: number;
  tariffFixDate?: string;
  tariffLowerDate?: string;
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
    
    // Проверяем период (максимум 8 дней согласно документации)
    const start = new Date(startDate);
    const end = new Date(endDate);
    const diffTime = Math.abs(end.getTime() - start.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    if (diffDays > 8) {
      console.warn(`⚠️ Период ${diffDays} дней превышает лимит 8 дней для API хранения`);
      // Ограничиваем период до 8 дней
      const limitedEndDate = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
      const limitedEnd = limitedEndDate.toISOString().split('T')[0];
      console.log(`📅 Ограничиваем период до: ${startDate} - ${limitedEnd}`);
      return await getStorageData(token, startDate, limitedEnd);
    }
    
    // Создаем taskId для хранения с GET запросом (согласно документации)
    console.log("📤 Создание задачи для отчета по хранению...");
    const taskUrl = `https://seller-analytics-api.wildberries.ru/api/v1/paid_storage?dateFrom=${startDate}&dateTo=${endDate}`;
    console.log(`📡 URL создания задачи: ${taskUrl}`);
    
    const taskResponse = await fetch(taskUrl, {
      method: "GET",
      headers: {
        "Authorization": token,
      },
    });

    console.log(`📊 Ответ создания задачи хранения: ${taskResponse.status} ${taskResponse.statusText}`);

    let taskId: string | undefined;

    if (!taskResponse.ok) {
      const errorText = await taskResponse.text();
      console.error(`❌ Ошибка создания taskId для хранения: ${taskResponse.status}`, errorText);
      
      // Проверяем специфичные ошибки
      if (taskResponse.status === 429) {
        console.warn("⚠️ Превышен лимит запросов (1 запрос в минуту)");
        console.log("⏳ Ожидание 2 минуты перед повторной попыткой...");
        
        // Ждем 2 минуты и повторяем попытку
        await new Promise(resolve => setTimeout(resolve, 120000));
        
        console.log("🔄 Повторная попытка создания taskId...");
        const retryTaskResponse = await fetch(taskUrl, {
          method: "GET",
          headers: {
            "Authorization": token,
          },
        });
        
        if (retryTaskResponse.ok) {
          const retryTaskData = await retryTaskResponse.json();
          taskId = retryTaskData?.data?.taskId;
          
          if (taskId) {
            console.log(`✅ Создан taskId после повторной попытки: ${taskId}`);
          } else {
            console.error("❌ Не удалось получить taskId после повторной попытки");
            return [];
          }
        } else {
          console.error(`❌ Повторная попытка создания taskId не удалась: ${retryTaskResponse.status}`);
          return [];
        }
      } else if (taskResponse.status === 401) {
        console.error("❌ Неверный токен авторизации");
        return [];
      } else {
        return [];
      }
    } else {
      const taskData = await taskResponse.json();
      console.log("📋 Ответ создания задачи:", JSON.stringify(taskData, null, 2));
      
      taskId = taskData?.data?.taskId;

      if (!taskId) {
        console.error("❌ Не удалось получить taskId для хранения. Структура ответа:", taskData);
        return [];
      }
    }

    console.log(`✅ Создан taskId для хранения: ${taskId}`);

    // Ждем подготовки отчета (рекомендуется проверять статус)
    console.log("⏳ Ожидание подготовки отчета о хранении...");
    
    // Проверяем статус задачи
    const maxStatusChecks = 12; // 12 проверок по 5 секунд = 1 минута
    let statusCheckCount = 0;
    
    while (statusCheckCount < maxStatusChecks) {
      await new Promise(resolve => setTimeout(resolve, 5000)); // Ждем 5 секунд
      statusCheckCount++;
      
      console.log(`🔍 Проверка статуса ${statusCheckCount}/${maxStatusChecks}...`);
      
      try {
        const statusResponse = await fetch(`https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/status`, {
          method: "GET",
          headers: {
            "Authorization": token,
          },
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
      } catch (statusError) {
        console.warn("⚠️ Ошибка при проверке статуса:", statusError);
      }
    }

    // Загружаем отчет с повторными попытками (увеличено для 429 ошибок)
    const maxRetries = 5;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      console.log(`📥 Попытка загрузки отчета ${retryCount + 1}/${maxRetries}...`);
      
      const downloadResponse = await fetch(`https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/${taskId}/download`, {
        method: "GET",
        headers: {
          "Authorization": token,
        },
      });

      console.log(`📊 Ответ загрузки хранения: ${downloadResponse.status} ${downloadResponse.statusText}`);

      if (downloadResponse.ok) {
        const responseText = await downloadResponse.text();
        console.log(`📄 Сырой ответ загрузки хранения (длина: ${responseText.length}):`, responseText.substring(0, 500) + "...");
        
        // Дополнительные проверки перед парсингом JSON
        if (!responseText || responseText.trim().length === 0) {
          console.warn("⚠️ API хранения вернул пустой ответ");
          return [];
        }
        
        const trimmedResponse = responseText.trim();
        if (!trimmedResponse.startsWith('{') && !trimmedResponse.startsWith('[')) {
          console.warn("⚠️ API хранения вернул не-JSON ответ:", trimmedResponse.substring(0, 100));
          return [];
        }
        
        let storageData;
        try {
          storageData = JSON.parse(responseText);
        } catch (parseError) {
          console.error("❌ Ошибка парсинга JSON хранения:", parseError);
          console.log("📄 Полный ответ для отладки:", responseText);
          console.log("📏 Длина ответа:", responseText.length);
          return [];
        }
        
        console.log("📦 Структура ответа хранения:", Object.keys(storageData || {}));
        
        // Данные должны быть массивом согласно документации
        const dataArray = Array.isArray(storageData) ? storageData : [];
        
        console.log(`✅ Получено ${dataArray.length} записей о хранении`);
        
        if (dataArray.length > 0) {
          console.log("🔍 Первая запись хранения:", JSON.stringify(dataArray[0], null, 2));
          console.log("🏷️ Поля записи:", Object.keys(dataArray[0] || {}));
        }
        
        return dataArray;
      } else if (downloadResponse.status === 429) {
        // Экспоненциальная задержка для 429 ошибок
        const delayMinutes = Math.min(2 ** retryCount, 8); // 2, 4, 8 минут максимум
        const delayMs = delayMinutes * 60 * 1000;
        
        console.log(`⚠️ Ошибка 429 для taskId ${taskId}. Попытка ${retryCount + 1} из ${maxRetries}`);
        console.log(`⏳ Ожидание ${delayMinutes} минут перед повторной попыткой...`);
        
        await new Promise(resolve => setTimeout(resolve, delayMs));
        retryCount++;
      } else {
        const errorText = await downloadResponse.text();
        console.error(`❌ Ошибка загрузки отчета о хранении: ${downloadResponse.status}`, errorText);
        retryCount++;
        
        if (retryCount < maxRetries) {
          console.log(`🔄 Повторная попытка через 10 секунд...`);
          await new Promise(resolve => setTimeout(resolve, 10000));
        }
      }
    }

    console.error(`❌ Не удалось загрузить отчет о хранении после ${maxRetries} попыток`);
    return [];
  } catch (error) {
    console.error("💥 Критическая ошибка при получении данных о хранении:", error);
    return [];
  }
} 