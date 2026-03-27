import { useState } from 'react';
import { Send, CheckCircle, AlertCircle, Download } from 'lucide-react';

export default function ProcessTab() {
  const [form, setForm]         = useState({ Service: 'Negotiation', hs_object_id: '', consent_order_json: '' });
  const [loading, setLoading]   = useState(false);
  const [result, setResult]     = useState(null);
  const [jsonError, setJsonError] = useState(null);

  const validateJson = (val) => {
    if (!val) { setJsonError(null); return; }
    try { JSON.parse(val); setJsonError(null); }
    catch (e) { setJsonError('Invalid JSON'); }
  };

  const handleSubmit = async () => {
    if (!form.hs_object_id || !form.consent_order_json || jsonError) return;
    setLoading(true);
    setResult(null);
    try {
      const res = await fetch('/api/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          Service: form.Service,
          hs_object_id: Number(form.hs_object_id),
          consent_order_json: JSON.parse(form.consent_order_json),
        }),
      });

      if (!res.ok) {
        const err = await res.json();
        setResult({ ok: false, error: err.error });
        return;
      }

      // Response is the xlsx binary — trigger download
      const blob = await res.blob();
      const filename = res.headers.get('Content-Disposition')?.match(/filename="(.+?)"/)?.[1]
        || `Financial_Disclosure_${form.Service}_${form.hs_object_id}.xlsx`;
      const hsObjectId = res.headers.get('X-HS-Object-ID');
      const service    = res.headers.get('X-Service');

      const url = URL.createObjectURL(blob);
      const a   = document.createElement('a');
      a.href = url; a.download = filename; a.click();
      URL.revokeObjectURL(url);

      setResult({ ok: true, filename, hs_object_id: hsObjectId, service });
    } catch (e) {
      setResult({ ok: false, error: e.message });
    } finally {
      setLoading(false);
    }
  };

  const canSubmit = !loading && form.hs_object_id && form.consent_order_json && !jsonError;

  return (
    <div style={{ maxWidth: 620 }}>
      <div style={{ marginBottom: 28 }}>
        <h2 style={{ fontSize: 28, marginBottom: 6 }}>Process</h2>
        <p style={{ color: 'var(--ink-soft)', fontSize: 15 }}>
          Submit a consent order to generate and download the prefilled financial disclosure spreadsheet.
        </p>
      </div>

      <div style={{
        background: '#fff',
        border: '1px solid var(--border)',
        borderRadius: 12,
        padding: 28,
        boxShadow: 'var(--shadow)',
      }}>
        {/* Endpoint info */}
        <div style={{
          background: 'var(--cream)',
          borderRadius: 8,
          padding: '10px 14px',
          marginBottom: 24,
          fontFamily: 'monospace',
          fontSize: 13,
          color: 'var(--ink-soft)',
          border: '1px solid var(--border)',
        }}>
          <span style={{ color: 'var(--green)', fontWeight: 600 }}>POST</span>
          {' '}/api/process
        </div>

        {/* Service */}
        <div style={{ marginBottom: 18 }}>
          <label style={{ display: 'block', fontSize: 13, fontWeight: 600, marginBottom: 8, color: 'var(--ink)' }}>
            Service
          </label>
          <div style={{ display: 'flex', gap: 10 }}>
            {['Negotiation', 'Assisted'].map(s => (
              <button
                key={s}
                onClick={() => setForm(f => ({ ...f, Service: s }))}
                style={{
                  flex: 1,
                  padding: '10px 0',
                  borderRadius: 8,
                  border: `2px solid ${form.Service === s
                    ? s === 'Negotiation' ? 'var(--green)' : 'var(--blue)'
                    : 'var(--border)'}`,
                  background: form.Service === s
                    ? s === 'Negotiation' ? 'var(--green-bg)' : 'var(--blue-bg)'
                    : '#fff',
                  color: form.Service === s
                    ? s === 'Negotiation' ? 'var(--green)' : 'var(--blue)'
                    : 'var(--ink-soft)',
                  fontWeight: form.Service === s ? 600 : 400,
                  fontSize: 14,
                  transition: 'all 0.15s',
                  cursor: 'pointer',
                }}
              >
                {s}
              </button>
            ))}
          </div>
        </div>

        {/* HubSpot Object ID */}
        <div style={{ marginBottom: 18 }}>
          <label style={{ display: 'block', fontSize: 13, fontWeight: 600, marginBottom: 8, color: 'var(--ink)' }}>
            HS Object ID
          </label>
          <input
            type="number"
            value={form.hs_object_id}
            onChange={e => setForm(f => ({ ...f, hs_object_id: e.target.value }))}
            placeholder="e.g. 358206255303"
            style={{
              width: '100%',
              padding: '10px 14px',
              border: '1px solid var(--border)',
              borderRadius: 8,
              fontSize: 15,
              color: 'var(--ink)',
              background: '#fff',
              outline: 'none',
              boxSizing: 'border-box',
              transition: 'border-color 0.15s',
            }}
            onFocus={e => e.target.style.borderColor = 'var(--gold)'}
            onBlur={e => e.target.style.borderColor = 'var(--border)'}
          />
        </div>

        {/* Consent Order JSON */}
        <div style={{ marginBottom: 24 }}>
          <label style={{ display: 'block', fontSize: 13, fontWeight: 600, marginBottom: 8, color: 'var(--ink)' }}>
            Consent Order JSON
          </label>
          <textarea
            value={form.consent_order_json}
            onChange={e => { setForm(f => ({ ...f, consent_order_json: e.target.value })); validateJson(e.target.value); }}
            placeholder='Paste the consent_order_json object here…'
            rows={8}
            style={{
              width: '100%',
              padding: '10px 14px',
              border: `1px solid ${jsonError ? 'var(--red)' : 'var(--border)'}`,
              borderRadius: 8,
              fontSize: 13,
              fontFamily: 'monospace',
              color: 'var(--ink)',
              background: '#fff',
              outline: 'none',
              resize: 'vertical',
              boxSizing: 'border-box',
              transition: 'border-color 0.15s',
            }}
            onFocus={e => e.target.style.borderColor = jsonError ? 'var(--red)' : 'var(--gold)'}
            onBlur={e => e.target.style.borderColor = jsonError ? 'var(--red)' : 'var(--border)'}
          />
          {jsonError && (
            <div style={{ fontSize: 12, color: 'var(--red)', marginTop: 4 }}>{jsonError}</div>
          )}
        </div>

        {/* Submit */}
        <button
          onClick={handleSubmit}
          disabled={!canSubmit}
          style={{
            width: '100%',
            padding: '12px',
            borderRadius: 8,
            background: canSubmit ? 'var(--ink)' : 'var(--cream-dark)',
            color: canSubmit ? '#fff' : 'var(--ink-muted)',
            fontWeight: 600,
            fontSize: 15,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: 8,
            border: 'none',
            transition: 'all 0.15s',
            cursor: canSubmit ? 'pointer' : 'not-allowed',
          }}
        >
          {loading
            ? <><div style={{ width: 18, height: 18, border: '2px solid rgba(255,255,255,0.3)', borderTopColor: '#fff', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }} /> Generating…</>
            : <><Download size={16} /> Generate &amp; Download</>
          }
        </button>

        {/* Result */}
        {result && (
          <div style={{
            marginTop: 16,
            padding: '14px 16px',
            borderRadius: 8,
            background: result.ok ? 'var(--green-bg)' : 'var(--red-bg)',
            border: `1px solid ${result.ok ? 'rgba(45,122,79,0.2)' : 'rgba(192,57,43,0.2)'}`,
            animation: 'fadeIn 0.2s ease',
          }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              color: result.ok ? 'var(--green)' : 'var(--red)',
              fontWeight: 600, fontSize: 14, marginBottom: result.ok ? 6 : 0,
            }}>
              {result.ok ? <CheckCircle size={16} /> : <AlertCircle size={16} />}
              {result.ok ? 'File downloaded' : result.error}
            </div>
            {result.ok && (
              <div style={{ fontSize: 12, fontFamily: 'monospace', color: 'var(--green)' }}>
                <div>service: {result.service}</div>
                <div>hs_object_id: {result.hs_object_id}</div>
                <div>filename: {result.filename}</div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
