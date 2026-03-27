import { useState } from 'react';
import TemplatesTab from './components/TemplatesTab.jsx';
import ProcessTab from './components/ProcessTab.jsx';
import LogsTab from './components/LogsTab.jsx';

const TABS = [
  { id: 'templates', label: 'Templates' },
  { id: 'process',   label: 'Process' },
  { id: 'logs',      label: 'Logs' },
];

export default function App() {
  const [tab, setTab] = useState('templates');

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <header style={{
        background: 'var(--ink)',
        borderBottom: '3px solid var(--gold)',
        padding: '0 32px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        height: 64,
        position: 'sticky',
        top: 0,
        zIndex: 100,
      }}>
        <div style={{ display: 'flex', alignItems: 'baseline', gap: 10 }}>
          <span style={{
            fontFamily: 'var(--font-serif)',
            color: '#fff',
            fontSize: 22,
            letterSpacing: '-0.01em',
          }}>
            Financial Disclosure
          </span>
          <span style={{ color: 'var(--gold)', fontSize: 13, fontWeight: 500, letterSpacing: '0.05em' }}>
            GENERATOR
          </span>
        </div>

        {/* Tabs */}
        <nav style={{ display: 'flex', gap: 4 }}>
          {TABS.map(t => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                padding: '6px 18px',
                borderRadius: 6,
                background: tab === t.id ? 'var(--gold)' : 'transparent',
                color: tab === t.id ? 'var(--ink)' : 'rgba(255,255,255,0.65)',
                fontWeight: tab === t.id ? 600 : 400,
                fontSize: 14,
                transition: 'all 0.15s',
                letterSpacing: '0.01em',
              }}
            >
              {t.label}
            </button>
          ))}
        </nav>
      </header>

      {/* Content */}
      <main style={{ flex: 1, padding: '32px', maxWidth: 1100, margin: '0 auto', width: '100%' }}>
        {tab === 'templates' && <TemplatesTab />}
        {tab === 'process'   && <ProcessTab />}
        {tab === 'logs'      && <LogsTab />}
      </main>
    </div>
  );
}
