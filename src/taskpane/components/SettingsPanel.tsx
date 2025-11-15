import * as React from 'react';
import { Settings } from '../../types';
import './SettingsPanel.css';

interface SettingsPanelProps {
  settings: Settings;
  onSave: (settings: Settings) => void;
  onClose: () => void;
}

export const SettingsPanel: React.FC<SettingsPanelProps> = ({ settings, onSave, onClose }) => {
  const [localSettings, setLocalSettings] = React.useState<Settings>(settings);

  const handleToggle = (category: keyof Settings, field: string) => {
    setLocalSettings({
      ...localSettings,
      [category]: {
        ...localSettings[category],
        [field]: !localSettings[category][field as keyof typeof localSettings[typeof category]],
      },
    });
  };

  const handleArrayChange = (category: keyof Settings, field: string, value: string) => {
    const items = value.split(',').map((item) => item.trim()).filter((item) => item !== '');
    setLocalSettings({
      ...localSettings,
      [category]: {
        ...localSettings[category],
        [field]: items,
      },
    });
  };

  const handleNumberChange = (category: keyof Settings, field: string, value: string) => {
    setLocalSettings({
      ...localSettings,
      [category]: {
        ...localSettings[category],
        [field]: parseInt(value, 10) || 0,
      },
    });
  };

  const handleSave = () => {
    onSave(localSettings);
  };

  return (
    <div className="settings-panel">
      <div className="settings-header">
        <h2>設定</h2>
        <button className="close-button" onClick={onClose}>
          ✕
        </button>
      </div>

      <div className="settings-content">
        {/* 添付ファイルチェック */}
        <section className="settings-section">
          <h3>📎 添付ファイルチェック</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.attachmentCheck.enabled}
              onChange={() => handleToggle('attachmentCheck', 'enabled')}
            />
            有効化
          </label>

          <div className="setting-item">
            <label>チェックキーワード（カンマ区切り）:</label>
            <textarea
              value={localSettings.attachmentCheck.keywords.join(', ')}
              onChange={(e) => handleArrayChange('attachmentCheck', 'keywords', e.target.value)}
              rows={3}
            />
          </div>

          <div className="setting-item">
            <label>除外キーワード（カンマ区切り）:</label>
            <textarea
              value={localSettings.attachmentCheck.excludeKeywords.join(', ')}
              onChange={(e) => handleArrayChange('attachmentCheck', 'excludeKeywords', e.target.value)}
              rows={2}
            />
          </div>
        </section>

        {/* 宛先チェック */}
        <section className="settings-section">
          <h3>👥 宛先チェック</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.recipientCheck.enabled}
              onChange={() => handleToggle('recipientCheck', 'enabled')}
            />
            有効化
          </label>

          <div className="setting-item">
            <label>社内ドメイン（カンマ区切り）:</label>
            <textarea
              value={localSettings.recipientCheck.internalDomains.join(', ')}
              onChange={(e) => handleArrayChange('recipientCheck', 'internalDomains', e.target.value)}
              rows={2}
              placeholder="例: example.com, company.co.jp"
            />
          </div>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.recipientCheck.warnExternalRecipients}
              onChange={() => handleToggle('recipientCheck', 'warnExternalRecipients')}
            />
            社外メール送信時に警告
          </label>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.recipientCheck.warnMultipleExternal}
              onChange={() => handleToggle('recipientCheck', 'warnMultipleExternal')}
            />
            複数の社外宛先でBccを推奨
          </label>

          <div className="setting-item">
            <label>宛先件数の上限:</label>
            <input
              type="number"
              value={localSettings.recipientCheck.maxRecipientCount}
              onChange={(e) => handleNumberChange('recipientCheck', 'maxRecipientCount', e.target.value)}
              min={1}
            />
          </div>
        </section>

        {/* 件名チェック */}
        <section className="settings-section">
          <h3>📋 件名チェック</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.subjectCheck.enabled}
              onChange={() => handleToggle('subjectCheck', 'enabled')}
            />
            有効化
          </label>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.subjectCheck.requireSubject}
              onChange={() => handleToggle('subjectCheck', 'requireSubject')}
            />
            件名必須
          </label>
        </section>

        {/* 本文チェック */}
        <section className="settings-section">
          <h3>📝 本文チェック</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.bodyCheck.enabled}
              onChange={() => handleToggle('bodyCheck', 'enabled')}
            />
            有効化
          </label>

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={localSettings.bodyCheck.requireBody}
              onChange={() => handleToggle('bodyCheck', 'requireBody')}
            />
            本文必須
          </label>

          <div className="setting-item">
            <label>TODOキーワード（カンマ区切り）:</label>
            <textarea
              value={localSettings.bodyCheck.todoKeywords.join(', ')}
              onChange={(e) => handleArrayChange('bodyCheck', 'todoKeywords', e.target.value)}
              rows={2}
            />
          </div>
        </section>
      </div>

      <div className="settings-footer">
        <button className="save-button" onClick={handleSave}>
          保存
        </button>
        <button className="cancel-button" onClick={onClose}>
          キャンセル
        </button>
      </div>
    </div>
  );
};
