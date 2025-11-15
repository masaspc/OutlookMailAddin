import * as React from 'react';
import { CheckResult } from '../../types';
import './CheckResultList.css';

interface CheckResultListProps {
  results: CheckResult[];
}

export const CheckResultList: React.FC<CheckResultListProps> = ({ results }) => {
  if (results.length === 0) {
    return (
      <div className="no-issues">
        <div className="success-icon">✓</div>
        <h3>問題は見つかりませんでした</h3>
        <p>メールは送信可能です。</p>
      </div>
    );
  }

  // カテゴリごとにグループ化
  const groupedResults = results.reduce((acc, result) => {
    if (!acc[result.category]) {
      acc[result.category] = [];
    }
    acc[result.category].push(result);
    return acc;
  }, {} as Record<string, CheckResult[]>);

  const categoryNames: Record<string, string> = {
    attachment: '📎 添付ファイル',
    recipient: '👥 宛先',
    subject: '📋 件名',
    body: '📝 本文',
    time: '⏰ 送信時刻',
  };

  const severityIcons: Record<string, string> = {
    error: '❌',
    warning: '⚠️',
    info: 'ℹ️',
  };

  const severityColors: Record<string, string> = {
    error: '#d13438',
    warning: '#f7b500',
    info: '#0078d4',
  };

  return (
    <div className="check-results">
      {Object.entries(groupedResults).map(([category, categoryResults]) => (
        <div key={category} className="category-group">
          <h3 className="category-title">{categoryNames[category] || category}</h3>
          {categoryResults.map((result, index) => (
            <div
              key={index}
              className="result-item"
              style={{ borderLeftColor: severityColors[result.severity] }}
            >
              <div className="result-header">
                <span className="severity-icon">{severityIcons[result.severity]}</span>
                <span className="result-message">{result.message}</span>
              </div>
              {result.details && (
                <div className="result-details">{result.details}</div>
              )}
            </div>
          ))}
        </div>
      ))}
    </div>
  );
};
