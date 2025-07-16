import Link from "next/link";
import { Key, Download } from "lucide-react";

export default function Home() {
  return (
    <div className="min-h-screen bg-gray-50 dark:bg-gray-900">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="py-12">
          {/* Header */}
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold text-gray-900 dark:text-white mb-4">
              WB Reports
            </h1>
            <p className="text-lg text-gray-600 dark:text-gray-300 max-w-2xl mx-auto">
              Платформа для управления API-ключами Wildberries и выгрузки отчётов
            </p>
          </div>

          {/* Navigation Cards */}
          <div className="grid md:grid-cols-2 gap-8 max-w-4xl mx-auto">
            <Link 
              href="/tokens"
              className="group bg-white dark:bg-gray-800 rounded-2xl p-8 shadow-sm border border-gray-200 dark:border-gray-700 hover:shadow-md transition-all duration-200 hover:border-blue-300 dark:hover:border-blue-600"
            >
              <div className="flex items-center space-x-4 mb-4">
                <div className="bg-blue-100 dark:bg-blue-900 p-3 rounded-xl">
                  <Key className="w-6 h-6 text-blue-600 dark:text-blue-300" />
                </div>
                <h2 className="text-2xl font-semibold text-gray-900 dark:text-white">
                  Управление токенами
                </h2>
              </div>
              <p className="text-gray-600 dark:text-gray-300 group-hover:text-gray-700 dark:group-hover:text-gray-200 transition-colors">
                Добавление, редактирование и управление API-ключами Wildberries
              </p>
            </Link>

            <Link 
              href="/reports"
              className="group bg-white dark:bg-gray-800 rounded-2xl p-8 shadow-sm border border-gray-200 dark:border-gray-700 hover:shadow-md transition-all duration-200 hover:border-green-300 dark:hover:border-green-600"
            >
              <div className="flex items-center space-x-4 mb-4">
                <div className="bg-green-100 dark:bg-green-900 p-3 rounded-xl">
                  <Download className="w-6 h-6 text-green-600 dark:text-green-300" />
                </div>
                <h2 className="text-2xl font-semibold text-gray-900 dark:text-white">
                  Выгрузка отчётов
                </h2>
              </div>
              <p className="text-gray-600 dark:text-gray-300 group-hover:text-gray-700 dark:group-hover:text-gray-200 transition-colors">
                Выбор периода и скачивание отчётов в формате Excel
              </p>
            </Link>
          </div>
        </div>
      </div>
    </div>
  );
}
