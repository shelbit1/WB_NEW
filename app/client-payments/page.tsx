"use client";

import { useEffect, useState, useCallback, useMemo } from "react";
import Link from "next/link";
import { debounce } from 'lodash';

interface Token {
  _id: string;
  name: string;
  paymentStatus: 'free' | 'trial' | 'pending' | 'disabled';
  comment?: string;
}

export default function ClientPaymentsPage() {
  const [tokens, setTokens] = useState<Token[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchTokens = async () => {
      try {
        setLoading(true);
        setError(null);
        const res = await fetch('/api/tokens');
        if (!res.ok) {
          throw new Error('Не удалось получить токены');
        }
        const data = await res.json();
        setTokens(data);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Произошла ошибка');
      } finally {
        setLoading(false);
      }
    };

    fetchTokens();
  }, []);

  const updateTokenOnServer = useCallback(async (tokenId: string, data: Partial<Pick<Token, 'paymentStatus' | 'comment'>>) => {
    const originalTokens = [...tokens];
    
    // Optimistic update
    setTokens(currentTokens =>
      currentTokens.map(t =>
        t._id === tokenId ? { ...t, ...data } : t
      )
    );

    try {
      const res = await fetch(`/api/tokens/${tokenId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

      if (!res.ok) throw new Error('Не удалось сохранить изменения');
      
      const updatedToken = await res.json();

      // Sync with server state
      setTokens(currentTokens =>
        currentTokens.map(t => (t._id === tokenId ? updatedToken : t))
      );
    } catch (error) {
      setError(error instanceof Error ? error.message : 'Ошибка сохранения');
      // Rollback on error
      setTokens(originalTokens);
    }
  }, [tokens]);
  
  const debouncedCommentUpdate = useMemo(() => debounce(updateTokenOnServer, 500), [updateTokenOnServer]);

  const handleStatusChange = (tokenId: string, newStatus: Token['paymentStatus']) => {
    updateTokenOnServer(tokenId, { paymentStatus: newStatus });
  };

  const handleCommentChange = (tokenId: string, newComment: string) => {
    // We only do optimistic update for comments locally before debounce kicks in
     setTokens(currentTokens =>
      currentTokens.map(t =>
        t._id === tokenId ? { ...t, comment: newComment } : t
      )
    );
    debouncedCommentUpdate(tokenId, { comment: newComment });
  };

  return (
    <div className="container mx-auto px-4 py-8">
      <div className="flex justify-between items-center mb-6">
        <h1 className="text-3xl font-bold">Оплаты клиентов</h1>
        <Link href="/" className="text-blue-500 hover:underline">
          На главную
        </Link>
      </div>

      {loading && <p>Загрузка...</p>}
      {error && <p className="text-red-500 bg-red-100 p-3 rounded-md">{error}</p>}
      
      {!loading && !error && (
        <div className="bg-white rounded-lg shadow-md">
          <ul className="divide-y divide-gray-200">
            {tokens.length > 0 ? (
              tokens.map((token) => (
                <li key={token._id} className="px-6 py-4 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="font-medium text-lg">{token.name}</span>
                    <select
                      value={token.paymentStatus}
                      onChange={(e) => handleStatusChange(token._id, e.target.value as Token['paymentStatus'])}
                      className="ml-4 p-2 border rounded-md"
                    >
                      <option value="free">Бесплатно</option>
                      <option value="trial">Пробный период</option>
                      <option value="pending">Ожидаем оплату</option>
                      <option value="disabled">Отключен</option>
                    </select>
                  </div>
                  <div>
                    <textarea
                      placeholder="Комментарий..."
                      value={token.comment || ''}
                      onChange={(e) => handleCommentChange(token._id, e.target.value)}
                      className="p-2 border rounded-md w-full text-sm"
                      rows={2}
                    />
                  </div>
                </li>
              ))
            ) : (
              <li className="px-6 py-4 text-center text-gray-500">
                Клиенты не найдены. Добавьте их в разделе «Управление токенами».
              </li>
            )}
          </ul>
        </div>
      )}
    </div>
  );
} 