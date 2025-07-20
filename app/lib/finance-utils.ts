// Интерфейсы для финансовых данных
export interface FinancialData {
  advertId: number;
  date: string;
  sum: number;
  bill: number;
  type: string;
  docNumber: string;
  sku?: string;
  campName?: string;  // Название кампании из API
}

export interface Campaign {
  advertId: number;
  name: string;
  type: string;
  status: string;
}

export interface CampaignInfo {
  advertId: number;
  name: string;
}

// Интерфейсы для API ответов
interface WildberriesAdvertData {
  advertId: number;
  name?: string;
  type?: string;
  status?: string;
}

interface WildberriesCampaignsResponse {
  adverts?: WildberriesAdvertData[];
}

interface WildberriesFinanceRecord {
  advertId: number;
  updTime?: string;
  updSum?: number;
  paymentType?: string;
  type?: string;
  updNum?: string;
  campName?: string;      // Название кампании из API
  advertType?: number;    // Тип рекламы из API  
  advertStatus?: number;  // Статус кампании из API
}





// Функция добавления дней к дате
export function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

// Функция форматирования даты
export function formatDate(date: Date): string {
  return date.toISOString().split('T')[0];
}

// Функция получения кампаний
export async function fetchCampaigns(apiKey: string): Promise<Campaign[]> {
  try {
    console.log('📊 Получение списка кампаний...');
    
    const response = await fetch('https://advert-api.wildberries.ru/adv/v1/promotion/count', {
      method: 'GET',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      console.error(`❌ Ошибка получения кампаний: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json() as WildberriesCampaignsResponse;
    console.log(`📦 Получено ${Array.isArray(data.adverts) ? data.adverts.length : 0} кампаний`);
    
    // Преобразуем данные в нужный формат
    const campaigns: Campaign[] = (data.adverts || []).map((advert: WildberriesAdvertData) => ({
      advertId: advert.advertId,
      name: advert.name || `Кампания ${advert.advertId}`,
      type: advert.type || 'Неизвестно',
      status: advert.status || 'Неизвестно'
    }));

    console.log(`✅ Обработано ${campaigns.length} кампаний`);
    return campaigns;
  } catch (error) {
    console.error('❌ Ошибка при получении кампаний:', error);
    return [];
  }
}

// Интерфейсы для API fullstats
interface FullStatsResponse {
  views: number;
  clicks: number;
  sum: number;
  dates: string[];
  days: FullStatsDay[];
  advertId: number;
}

interface FullStatsDay {
  date: string;
  views: number;
  clicks: number;
  sum: number;
  apps: FullStatsApp[];
}

interface FullStatsApp {
  views: number;
  clicks: number;
  sum: number;
  nm: FullStatsNm[];
  appType: number;
}

interface FullStatsNm {
  views: number;
  clicks: number;
  sum: number;
  name: string;
  nmId: number;
}

// Новая функция получения детальной статистики с артикулами через API v2/fullstats
export async function fetchCampaignFullStats(apiKey: string, campaigns: Campaign[], startDate: string, endDate: string): Promise<Map<number, string>> {
  const articlesMap = new Map<number, string>();
  
  try {
    console.log(`📊 Получение полной статистики с артикулами для ${campaigns.length} кампаний через API v2/fullstats...`);
    
    // Получаем статистику для каждой кампании (пакетами для оптимизации)
    const batchSize = 3; // Уменьшаем размер пакета так как API может быть тяжелым
    for (let i = 0; i < campaigns.length; i += batchSize) {
      const batch = campaigns.slice(i, i + batchSize);
      const promises = batch.map(async (campaign) => {
        try {
          console.log(`📈 Запрос статистики для кампании ${campaign.advertId}...`);
          
          const response = await fetch(`https://advert-api.wildberries.ru/adv/v2/fullstats?id=${campaign.advertId}&from=${startDate}&to=${endDate}`, {
            method: 'GET',
            headers: {
              'Authorization': apiKey,
              'Content-Type': 'application/json'
            }
          });

          if (response.ok) {
            const data: FullStatsResponse[] = await response.json();
            console.log(`✅ Получена статистика для кампании ${campaign.advertId}:`, data.length > 0 ? 'есть данные' : 'нет данных');
            
            if (data && data.length > 0) {
              const campaignStats = data[0]; // Берем первую запись
              const allArticles = new Set<string>(); // Используем Set для уникальных артикулов
              
              // Собираем все nmId из всех дней и приложений
              campaignStats.days?.forEach(day => {
                day.apps?.forEach(app => {
                  app.nm?.forEach(nm => {
                    if (nm.nmId) {
                      const articleInfo = nm.name ? `${nm.nmId}:${nm.name}` : `${nm.nmId}`;
                      allArticles.add(articleInfo);
                    }
                  });
                });
              });
              
              if (allArticles.size > 0) {
                const articlesList = Array.from(allArticles).join(', ');
                articlesMap.set(campaign.advertId, articlesList);
                console.log(`📦 Найдено ${allArticles.size} артикулов для кампании ${campaign.advertId}`);
              } else {
                articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId} (нет артикулов в статистике)`);
              }
            } else {
              articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId} (нет данных в fullstats)`);
            }
          } else {
            console.warn(`⚠️ API fullstats для кампании ${campaign.advertId} вернул ${response.status}`);
            articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId} (API Error:${response.status})`);
          }
        } catch (error) {
          console.error(`❌ Ошибка при получении fullstats для кампании ${campaign.advertId}:`, error);
          articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId} (Ошибка запроса)`);
        }
      });
      
      await Promise.all(promises);
      console.log(`📊 Обработано ${Math.min(i + batchSize, campaigns.length)} из ${campaigns.length} кампаний`);
      
      // Добавляем небольшую задержку между пакетами для соблюдения лимитов API
      if (i + batchSize < campaigns.length) {
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
    }
    
    console.log(`✅ Получена полная статистика для ${articlesMap.size} кампаний`);
    return articlesMap;
  } catch (error) {
    console.error('❌ Ошибка при получении полной статистики:', error);
    
    // В случае полной ошибки создаем базовые записи
    campaigns.forEach(campaign => {
      articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId} (Ошибка API)`);
    });
    
    return articlesMap;
  }
}

// Функция получения баланса счета (новая)
export async function fetchAccountBalance(apiKey: string): Promise<{balance: number, net: number, bonus: number} | null> {
  try {
    console.log('💰 Получение баланса счета...');
    
    const response = await fetch('https://advert-api.wildberries.ru/adv/v1/balance', {
      method: 'GET',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      console.log(`✅ Баланс получен: счет ${data.balance}, баланс ${data.net}, бонусы ${data.bonus}`);
      return data;
    } else {
      console.warn(`⚠️ Не удалось получить баланс: ${response.status}`);
      return null;
    }
  } catch (error) {
    console.error('❌ Ошибка при получении баланса:', error);
    return null;
  }
}

// Функция получения детальных данных по артикулам кампаний (упрощенная версия)
export async function fetchCampaignArticles(apiKey: string, campaigns: Campaign[]): Promise<Map<number, string>> {
  const articlesMap = new Map<number, string>();
  
  try {
    console.log(`📊 Попытка получения данных артикулов для ${campaigns.length} кампаний...`);
    
    // Пробуем получить данные через API бюджета (может содержать дополнительную информацию)
    const batchSize = 5; // Уменьшаем размер пакета для надежности
    for (let i = 0; i < campaigns.length; i += batchSize) {
      const batch = campaigns.slice(i, i + batchSize);
      const promises = batch.map(async (campaign) => {
        try {
          // Пробуем API бюджета кампании (более надежный)
          const response = await fetch(`https://advert-api.wildberries.ru/adv/v1/budget?id=${campaign.advertId}`, {
            method: 'GET',
            headers: {
              'Authorization': apiKey,
              'Content-Type': 'application/json'
            }
          });

          if (response.ok) {
            const budgetData = await response.json();
            // Есть бюджет - значит кампания активна
            const budgetInfo = `Бюджет: ${budgetData.total || 0} руб.`;
            articlesMap.set(campaign.advertId, `${campaign.name || campaign.advertId} (${budgetInfo})`);
          } else {
            // API не работает - используем базовую информацию
            articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId}`);
          }
        } catch {
          // При любой ошибке используем базовую информацию о кампании
          articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId}`);
        }
      });
      
      await Promise.all(promises);
      console.log(`📊 Обработано ${Math.min(i + batchSize, campaigns.length)} из ${campaigns.length} кампаний`);
    }
    
    console.log(`✅ Подготовлено описаний для ${articlesMap.size} кампаний`);
    return articlesMap;
  } catch (error) {
    console.error('❌ Ошибка при получении данных кампаний:', error);
    
    // В случае полной ошибки создаем базовые записи
    campaigns.forEach(campaign => {
      articlesMap.set(campaign.advertId, `${campaign.name || 'Кампания'} ID:${campaign.advertId}`);
    });
    
    return articlesMap;
  }
}



// Функция получения финансовых данных с логикой буферных дней
export async function fetchFinancialData(apiKey: string, startDate: string, endDate: string): Promise<FinancialData[]> {
  try {
    console.log(`📊 Получение финансовых данных: ${startDate} - ${endDate}`);
    
    // Добавляем буферные дни
    const originalStart = new Date(startDate);
    const originalEnd = new Date(endDate);
    const bufferStart = addDays(originalStart, -1);
    const bufferEnd = addDays(originalEnd, 1);
    
    const adjustedStartDate = formatDate(bufferStart);
    const adjustedEndDate = formatDate(bufferEnd);
    
    console.log(`📅 Расширенный период с буферными днями: ${adjustedStartDate} - ${adjustedEndDate}`);
    
    const response = await fetch(`https://advert-api.wildberries.ru/adv/v1/upd?from=${adjustedStartDate}&to=${adjustedEndDate}`, {
      method: 'GET',
      headers: {
        'Authorization': apiKey,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      console.error(`❌ Ошибка получения финансовых данных: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json() as WildberriesFinanceRecord[];
    console.log(`📦 Получено ${Array.isArray(data) ? data.length : 0} финансовых записей`);
    
    // Преобразуем данные в нужный формат
    const financialData: FinancialData[] = data.map((record: WildberriesFinanceRecord) => ({
      advertId: record.advertId,
      date: record.updTime ? new Date(record.updTime).toISOString().split('T')[0] : '',
      sum: record.updSum || 0,
      bill: record.paymentType === 'Счет' ? 1 : 0,
      type: record.type || 'Неизвестно',
      docNumber: record.updNum || '',
      sku: `Type:${record.advertType || 0} Status:${record.advertStatus || 0}`, // Добавляем тип и статус как SKU
      campName: record.campName || 'Неизвестная кампания' // Добавляем название кампании
    }));

    console.log(`✅ Обработано ${financialData.length} финансовых записей`);
    
    // Применяем логику буферных дней
    const filteredData = applyBufferDayLogic(financialData, originalStart, originalEnd);
    console.log(`📊 После применения логики буферных дней: ${filteredData.length} записей`);
    
    return filteredData;
  } catch (error) {
    console.error('❌ Ошибка при получении финансовых данных:', error);
    return [];
  }
}

// Логика фильтрации буферных дней
export function applyBufferDayLogic(data: FinancialData[], originalStart: Date, originalEnd: Date): FinancialData[] {
  const originalStartStr = formatDate(originalStart);
  const originalEndStr = formatDate(originalEnd);
  const bufferStartStr = formatDate(addDays(originalStart, -1));
  const bufferEndStr = formatDate(addDays(originalEnd, 1));
  
  console.log(`📊 Применение логики буферных дней:`);
  console.log(`   Основной период: ${originalStartStr} - ${originalEndStr}`);
  console.log(`   Буферные дни: ${bufferStartStr} (предыдущий), ${bufferEndStr} (следующий)`);
  
  // Разделяем данные по периодам
  const mainPeriodData = data.filter(record => 
    record.date >= originalStartStr && record.date <= originalEndStr
  );
  
  const previousBufferData = data.filter(record => record.date === bufferStartStr);
  const nextBufferData = data.filter(record => record.date === bufferEndStr);
  
  console.log(`   Основной период: ${mainPeriodData.length} записей`);
  console.log(`   Предыдущий буферный день: ${previousBufferData.length} записей`);
  console.log(`   Следующий буферный день: ${nextBufferData.length} записей`);
  
  // Получаем номера документов из основного периода
  const mainDocNumbers = new Set(mainPeriodData.map(record => record.docNumber));
  console.log(`   Номера документов основного периода: ${Array.from(mainDocNumbers).join(', ')}`);
  
  // Исключаем из основного периода записи с номерами документов из следующего буферного дня
  const nextBufferDocNumbers = new Set(nextBufferData.map(record => record.docNumber));
  const filteredMainData = mainPeriodData.filter(record => !nextBufferDocNumbers.has(record.docNumber));
  
  if (nextBufferDocNumbers.size > 0) {
    console.log(`   Исключаем из основного периода записи с номерами документов: ${Array.from(nextBufferDocNumbers).join(', ')}`);
    console.log(`   Исключено записей: ${mainPeriodData.length - filteredMainData.length}`);
  }
  
  // Добавляем записи из предыдущего буферного дня, если есть совпадения по номерам документов
  const previousBufferToAdd = previousBufferData.filter(record => 
    mainDocNumbers.has(record.docNumber) && previousBufferData.filter(r => r.docNumber === record.docNumber).length >= 2
  );
  
  if (previousBufferToAdd.length > 0) {
    console.log(`   Добавляем из предыдущего буферного дня: ${previousBufferToAdd.length} записей`);
  }
  
  const result = [...filteredMainData, ...previousBufferToAdd];
  console.log(`   Итоговое количество записей: ${result.length}`);
  
  return result;
} 