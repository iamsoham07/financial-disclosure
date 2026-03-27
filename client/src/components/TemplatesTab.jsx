import { useState, useEffect, useRef } from 'react';
import { Upload, Download, Trash2, CheckCircle, AlertCircle, FileSpreadsheet } from 'lucide-react';

const SERVICE_INFO = {
  Assisted:    { color: '#2255aa', bg: '#eef3fc', desc: 'Consent Order — Summary of Financial Disclosure' },
  Negotiation: { color: '#2d7a4f', bg: '#edf7f1', desc: 'Financial Disclosure & Net Effect Table' },
};

function TemplateCard({ service, template, onUpload, onDelete }) {
  const fileRef  = useRef();
  const [drag, setDrag]         = useState(false);
  const [uploading, setUploading] = useState(false);
  const [msg, setMsg]           = useState(null);
  const info = SERVICE_INFO[service];

  const handleFile = async (file) => {
    if (!file) return;
    setUploading(true);
    setMsg(null);
    try {
      const fd = new FormData();
      fd.append('service', service);
      fd.append('name', file.name);
      fd.append('file', file);
      const res = await fetch('/api/templates', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      setMsg({ type: 'success', text: 'Template uploaded successfully' });
      onUpload();
    } catch (e) {
      setMsg({ type: 'error', text: e.message });
    } finally {
      setUploading(false);
    }
  };

  const handleDelete = async () => {
    if (!confirm(`Remove the "${service}" template?`)) return;
    try {
      await fetch(`/api/templates/${service}`, { method: 'DELETE' });
      setMsg({ type: 'success', text: 'Template removed' });
      onDelete();
    } catch (e) {
      setMsg({ type: 'error', text: e.message });
    }
  };

  return (
    <div style={{
      background: '#fff',
      border: `1px solid var(--border)`,
      borderRadius: 12,
      overflow: 'hidden',
      boxShadow: 'var(--shadow)',
      animation: 'fadeIn 0.3s ease both',
    }}>
      {/* Card header */}
      <div style={{
        background: info.bg,
        borderBottom: `3px solid ${info.color}`,
        padding: '20px 24px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
      }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
            <FileSpreadsheet size={18} color={info.color} />
            <span style={{
              fontFamily: 'var(--font-serif)',
              fontSize: 20,
              color: info.color,
            }}>
              {service}
            </span>
            {template && (
              <span style={{
                background: info.color,
                color: '#fff',
                fontSize: 11,
                fontWeight: 600,
                padding: '2px 8px',
                borderRadius: 20,
                letterSpacing: '0.04em',
              }}>ACTIVE</span>
            )}
          </div>
          <p style={{ fontSize: 13, color: 'var(--ink-soft)' }}>{info.desc}</p>
        </div>
      </div>

      <div style={{ padding: '20px 24px' }}>
        {/* Current template info */}
        {template ? (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            background: 'var(--cream)',
            borderRadius: 8,
            padding: '12px 16px',
            marginBottom: 16,
            border: '1px solid var(--border)',
          }}>
            <div>
              <div style={{ fontWeight: 500, fontSize: 14, marginBottom: 2 }}>{template.file_name}</div>
              <div style={{ fontSize: 12, color: 'var(--ink-muted)' }}>
                Updated {new Date(template.updated_at).toLocaleDateString('en-GB', {
                  day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit'
                })}
              </div>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <a href={`/api/templates/${service}/download`}
                style={{
                  display: 'flex', alignItems: 'center', gap: 6,
                  padding: '7px 14px', borderRadius: 7,
                  background: 'var(--cream-dark)', color: 'var(--ink)',
                  fontSize: 13, fontWeight: 500, textDecoration: 'none',
                  border: '1px solid var(--border)',
                }}>
                <Download size={14} /> Download
              </a>
              <button onClick={handleDelete} style={{
                display: 'flex', alignItems: 'center', gap: 6,
                padding: '7px 14px', borderRadius: 7,
                background: 'var(--red-bg)', color: 'var(--red)',
                fontSize: 13, fontWeight: 500,
                border: '1px solid rgba(192,57,43,0.2)',
              }}>
                <Trash2 size={14} /> Remove
              </button>
            </div>
          </div>
        ) : (
          <div style={{
            padding: '12px 16px',
            background: 'var(--red-bg)',
            borderRadius: 8,
            marginBottom: 16,
            fontSize: 13,
            color: 'var(--red)',
            border: '1px solid rgba(192,57,43,0.15)',
          }}>
            No template uploaded yet — processing will fail without one.
          </div>
        )}

        {/* Drop zone */}
        <div
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
          onClick={() => fileRef.current?.click()}
          style={{
            border: `2px dashed ${drag ? info.color : 'var(--border)'}`,
            borderRadius: 8,
            padding: '24px',
            textAlign: 'center',
            cursor: uploading ? 'not-allowed' : 'pointer',
            background: drag ? info.bg : 'transparent',
            transition: 'all 0.15s',
          }}
        >
          <input ref={fileRef} type="file" accept=".xlsx" hidden
            onChange={e => handleFile(e.target.files[0])} />
          {uploading ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, color: 'var(--ink-soft)' }}>
              <div style={{ width: 18, height: 18, border: '2px solid var(--border)', borderTopColor: info.color, borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} />
              Uploading…
            </div>
          ) : (
            <>
              <Upload size={22} color={info.color} style={{ marginBottom: 8 }} />
              <div style={{ fontSize: 14, fontWeight: 500, color: 'var(--ink)' }}>
                {template ? 'Replace template' : 'Upload template'}
              </div>
              <div style={{ fontSize: 12, color: 'var(--ink-muted)', marginTop: 4 }}>
                Drop an .xlsx file here or click to browse
              </div>
            </>
          )}
        </div>

        {/* Feedback message */}
        {msg && (
          <div style={{
            display: 'flex', alignItems: 'center', gap: 8,
            marginTop: 12, padding: '10px 14px', borderRadius: 7,
            background: msg.type === 'success' ? 'var(--green-bg)' : 'var(--red-bg)',
            color: msg.type === 'success' ? 'var(--green)' : 'var(--red)',
            fontSize: 13, border: `1px solid ${msg.type === 'success' ? 'rgba(45,122,79,0.2)' : 'rgba(192,57,43,0.2)'}`,
          }}>
            {msg.type === 'success' ? <CheckCircle size={15} /> : <AlertCircle size={15} />}
            {msg.text}
          </div>
        )}
      </div>
    </div>
  );
}

export default function TemplatesTab() {
  const [templates, setTemplates] = useState([]);

  const fetchTemplates = async () => {
    try {
      const res = await fetch('/api/templates');
      const data = await res.json();
      setTemplates(data);
    } catch (e) { console.error(e); }
  };

  useEffect(() => { fetchTemplates(); }, []);

  const getTemplate = (service) => templates.find(t => t.service === service) || null;

  return (
    <div>
      <div style={{ marginBottom: 28 }}>
        <h2 style={{ fontSize: 28, marginBottom: 6 }}>Templates</h2>
        <p style={{ color: 'var(--ink-soft)', fontSize: 15 }}>
          Upload and manage the Excel templates for each service type. Templates are stored in the database.
        </p>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
        {['Assisted', 'Negotiation'].map(service => (
          <TemplateCard
            key={service}
            service={service}
            template={getTemplate(service)}
            onUpload={fetchTemplates}
            onDelete={fetchTemplates}
          />
        ))}
      </div>
    </div>
  );
}
