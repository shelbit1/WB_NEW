'use client';

import { useState, useEffect } from 'react';
import Link from 'next/link';
import { ArrowLeft, Download, Calendar } from 'lucide-react';

interface Token {
  _id: string;
  name: string;
}

export default function ReportsPage() {
  const [tokens, setTokens] = useState<Token[]>([]);
  const [selectedToken, setSelectedToken] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [loading, setLoading] = useState(true);
  const [downloading, setDownloading] = useState<string | null>(null);

  useEffect(() => {
    fetchTokens();
    // Устанавливаем даты по умолчанию (последние 30 дней)
    const today = new Date();
    const monthAgo = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate());
    
    setEndDate(today.toISOString().split('T')[0]);
    setStartDate(monthAgo.toISOString().split('T')[0]);
  }, []);

  const fetchTokens = async () => {
    try {
      const response = await fetch('/api/tokens');
      const data = await response.json();
      setTokens(data);
      if (data.length > 0) {
        setSelectedToken(data[0]._id);
      }
    } catch (error) {
      console.error('Ошибка при загрузке токенов:', error);
    } finally {
      setLoading(false);
    }
  };

  const downloadReport = async (reportType: string, reportName: string) => {
    if (!selectedToken || !startDate || !endDate) {
      alert('Пожалуйста, выберите клиента и период');
      return;
    }

    setDownloading(reportType);

    try {
      const selectedTokenData = tokens.find(t => t._id === selectedToken);
      const response = await fetch('/api/reports/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          reportType,
          clientName: selectedTokenData?.name,
          tokenId: selectedToken,
          startDate,
          endDate
        })
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${reportName} - ${startDate}–${endDate}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      } else {
        const errorData = await response.json();
        alert(errorData.error || 'Ошибка при скачивании отчёта');
      }
    } catch (error) {
      console.error('Ошибка:', error);
      alert('Ошибка при скачивании отчёта');
    } finally {
      setDownloading(null);
    }
  };

  const reports = [
    { key: 'details', name: 'Отчет детализации', color: 'blue' },
    { key: 'storage', name: 'Платное хранение', color: 'green' },
    { key: 'acceptance', name: 'Платная приемка', color: 'purple' },
    { key: 'products', name: 'Список товаров', color: 'orange' },
    { key: 'finances', name: 'Финансы РК', color: 'red' }
  ];

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50 dark:bg-gray-900 flex items-center justify-center">
        <div className="text-gray-600 dark:text-gray-400">Загрузка...</div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 dark:bg-gray-900">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Header */}
        <div className="flex items-center space-x-4 mb-8">
          <Link 
            href="/"
            className="p-2 rounded-lg bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors"
          >
            <ArrowLeft className="w-5 h-5 text-gray-600 dark:text-gray-400" />
          </Link>
          <h1 className="text-3xl font-bold text-gray-900 dark:text-white">
            Выгрузка отчётов
          </h1>
        </div>

        {tokens.length === 0 ? (
          <div className="bg-white dark:bg-gray-800 rounded-xl p-8 text-center border border-gray-200 dark:border-gray-700">
            <p className="text-gray-600 dark:text-gray-400 mb-4">
              У вас нет сохранённых токенов. Сначала добавьте API-ключ.
            </p>
            <Link 
              href="/tokens"
              className="inline-flex items-center space-x-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg transition-colors"
            >
              <span>Управление токенами</span>
            </Link>
          </div>
        ) : (
          <div className="space-y-8">
            {/* Filters */}
            <div className="bg-white dark:bg-gray-800 rounded-xl p-6 border border-gray-200 dark:border-gray-700">
              <h2 className="text-xl font-semibold text-gray-900 dark:text-white mb-4">
                Параметры выгрузки
              </h2>
              <div className="grid md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                    Клиент
                  </label>
                  <select
                    value={selectedToken}
                    onChange={(e) => setSelectedToken(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  >
                    {tokens.map((token) => (
                      <option key={token._id} value={token._id}>
                        {token.name}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                    Дата начала
                  </label>
                  <input
                    type="date"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                    Дата окончания
                  </label>
                  <input
                    type="date"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
              </div>
            </div>

            {/* Reports */}
            <div className="bg-white dark:bg-gray-800 rounded-xl p-6 border border-gray-200 dark:border-gray-700">
              <h2 className="text-xl font-semibold text-gray-900 dark:text-white mb-6">
                Доступные отчёты
              </h2>
              <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
                {reports.map((report) => (
                  <button
                    key={report.key}
                    onClick={() => downloadReport(report.key, report.name)}
                    disabled={downloading === report.key}
                    className={`
                      p-4 rounded-lg border-2 border-dashed transition-all duration-200
                      ${downloading === report.key 
                        ? 'border-gray-300 bg-gray-50 cursor-not-allowed' 
                        : `border-${report.color}-300 hover:border-${report.color}-400 hover:bg-${report.color}-50 dark:hover:bg-${report.color}-900/10`
                      }
                      text-left group
                    `}
                  >
                    <div className="flex items-center justify-between mb-2">
                      <Download className={`w-5 h-5 text-${report.color}-600`} />
                      {downloading === report.key && (
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-gray-900"></div>
                      )}
                    </div>
                    <h3 className={`font-medium text-gray-900 dark:text-white group-hover:text-${report.color}-700 dark:group-hover:text-${report.color}-300`}>
                      {report.name}
                    </h3>
                    <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
                      {downloading === report.key ? 'Генерация...' : 'Скачать Excel файл'}
                    </p>
                  </button>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
} 