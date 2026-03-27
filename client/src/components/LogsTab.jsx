import { useState, useEffect } from 'react';
import { RefreshCw, CheckCircle, XCircle, Clock, Download } from 'lucide-react';

const STATUS_STYLE = {
  success: { color: 'var(--green)', bg: 'var(--green-bg)', icon: CheckCircle },
  error:   { color: 'var(--red)',   bg: 'var(--red-bg)',   icon: XCircle },
  pending: { color: '#996600',      bg: '#fff8e6',         icon: Clock },
};

const SERVICE_COLOR = {
  Negotiation: 'var(--green)',
  Assisted:    'var(--blue)',
};

export default function LogsTab() {
  const [logs, setLogs]       = useState([]);
  const [loading, setLoading] = useState(false);
  const [expanded, setExpanded] = useState(null);

  const fetchLogs = async () => {
    setLoading(true);
    try {
      const res  = await fetch('/api/logs');
      const data = await res.json();
      setLogs(data);
    } catch (e) { console.error(e); }
    finally { setLoading(false); }
  };

  useEffect(() => {
    fetchLogs();
    const interval = setInterval(fetchLogs, 15000);
    return () => clearInterval(interval);
  }, []);

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', marginBottom: 28 }}>
        <div>
          <h2 style={{ fontSize: 28, marginBottom: 6 }}>Processing Log</h2>
          <p style={{ color: 'var(--ink-soft)', fontSize: 15 }}>
            Last 50 requests — auto-refreshes every 15 seconds.
          </p>
        </div>
        <button
          onClick={fetchLogs}
          style={{
            display: 'flex', alignItems: 'center', gap: 6,
            padding: '8px 16px', borderRadius: 8,
            background: '#fff', border: '1px solid var(--border)',
            color: 'var(--ink)', fontSize: 14, fontWeight: 500,
            cursor: 'pointer',
          }}
        >
          <RefreshCw size={14} style={{ animation: loading ? 'spin 0.7s linear infinite' : 'none' }} />
          Refresh
        </button>
      </div>

      {logs.length === 0 ? (
        <div style={{
          textAlign: 'center', padding: '60px 0',
          color: 'var(--ink-muted)', fontSize: 15,
          background: '#fff', borderRadius: 12,
          border: '1px solid var(--border)',
        }}>
          No processing records yet
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {logs.map((log, i) => {
            const style    = STATUS_STYLE[log.status] || STATUS_STYLE.pending;
            const Icon     = style.icon;
            const isOpen   = expanded === log.id;

            return (
              <div
                key={log.id}
                style={{
                  background: '#fff',
                  border: '1px solid var(--border)',
                  borderRadius: 10,
                  overflow: 'hidden',
                  boxShadow: 'var(--shadow)',
                  animation: `fadeIn 0.2s ${i * 0.03}s ease both`,
                }}
              >
                <div
                  onClick={() => setExpanded(isOpen ? null : log.id)}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 14,
                    padding: '14px 18px', cursor: 'pointer',
                    transition: 'background 0.1s',
                  }}
                  onMouseEnter={e => e.currentTarget.style.background = 'var(--cream)'}
                  onMouseLeave={e => e.currentTarget.style.background = '#fff'}
                >
                  {/* Status icon */}
                  <Icon size={18} color={style.color} />

                  {/* Service badge */}
                  <span style={{
                    fontSize: 12, fontWeight: 700,
                    color: SERVICE_COLOR[log.service] || 'var(--ink-soft)',
                    letterSpacing: '0.06em',
                    minWidth: 90,
                  }}>
                    {log.service?.toUpperCase()}
                  </span>

                  {/* Object ID */}
                  <span style={{
                    fontFamily: 'monospace', fontSize: 13,
                    color: 'var(--ink)', flex: 1,
                  }}>
                    hs_object_id: <strong>{log.hs_object_id}</strong>
                  </span>

                  {/* Status pill */}
                  <span style={{
                    fontSize: 11, fontWeight: 600,
                    padding: '3px 10px', borderRadius: 20,
                    background: style.bg, color: style.color,
                    letterSpacing: '0.04em',
                  }}>
                    {log.status.toUpperCase()}
                  </span>

                  {/* Date */}
                  <span style={{ fontSize: 12, color: 'var(--ink-muted)', minWidth: 130, textAlign: 'right' }}>
                    {new Date(log.created_at).toLocaleString('en-GB', {
                      day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
                    })}
                  </span>

                  {/* Download button — only shown when a generated file exists */}
                  {log.generated_file_name && (
                    <a
                      href={`/api/logs/${log.id}/download`}
                      download={log.generated_file_name}
                      onClick={e => e.stopPropagation()}
                      title={`Download ${log.generated_file_name}`}
                      style={{
                        display: 'flex', alignItems: 'center',
                        padding: '5px 10px', borderRadius: 6,
                        background: 'var(--cream)', border: '1px solid var(--border)',
                        color: 'var(--blue)', textDecoration: 'none', flexShrink: 0,
                      }}
                    >
                      <Download size={14} />
                    </a>
                  )}
                </div>

                {/* Expanded error details */}
                {isOpen && log.error_message && (
                  <div style={{
                    borderTop: '1px solid var(--border)',
                    padding: '14px 18px',
                    background: 'var(--cream)',
                    animation: 'fadeIn 0.15s ease',
                  }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--red)', marginBottom: 4 }}>Error</div>
                    <pre style={{ fontSize: 12, fontFamily: 'monospace', color: 'var(--red)', whiteSpace: 'pre-wrap' }}>
                      {log.error_message}
                    </pre>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
