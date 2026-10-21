import React, { useState, useEffect } from 'react';
import api from './api';

export function broadcastSitesUpdated() {
  window.dispatchEvent(new CustomEvent('mb-sites-updated'));
  try {
    if (typeof BroadcastChannel !== 'undefined') {
      const ch = new BroadcastChannel('mb_sites_sync');
      ch.postMessage({ type: 'SITES_UPDATED', timestamp: Date.now() });
      ch.close();
    }
  } catch (e) {}
  api.post('/public/sites/notify', { source: 'dashboard' }).catch(() => {});
}

export default function WebIntegrationSettingsPanel({ children }) {
  const [apiKey, setApiKey] = useState('');
  const [loadingKey, setLoadingKey] = useState(true);
  const [keyMasked, setKeyMasked] = useState(true);
  const [regeneratingKey, setRegeneratingKey] = useState(false);

  const [copiedCode, setCopiedCode] = useState(null);
  const [activeCodeTab, setActiveCodeTab] = useState('api');
  const [testBroadcasting, setTestBroadcasting] = useState(false);
  const [broadcastResult, setBroadcastResult] = useState(null);
  const [showEmbedKit, setShowEmbedKit] = useState(false);

  const origin = window.location.origin;

  // Fetch integration API key
  const fetchApiKey = async () => {
    setLoadingKey(true);
    try {
      const res = await api.get('/integration/key');
      if (res.data?.apiKey) {
        setApiKey(res.data.apiKey);
      }
    } catch (err) {
      console.warn('Failed to load integration key:', err);
    } finally {
      setLoadingKey(false);
    }
  };

  useEffect(() => {
    fetchApiKey();
  }, []);

  const handleRegenerateKey = async () => {
    const confirmed = window.confirm(
      '⚠️ Are you sure you want to REGENERATE your Integration API Key?\n\n' +
      'Once regenerated, the old key will be permanently invalidated. Any external websites, WordPress plugins, or apps using the old key will stop loading sites until they are updated with the new key.'
    );
    if (!confirmed) return;

    setRegeneratingKey(true);
    try {
      const res = await api.post('/integration/key/regenerate');
      if (res.data?.apiKey) {
        setApiKey(res.data.apiKey);
        setKeyMasked(false);
        setBroadcastResult({ success: true, count: 'New API key generated & applied to all embeds' });
      }
    } catch (err) {
      alert('Failed to regenerate key: ' + (err.response?.data?.message || err.message));
    } finally {
      setRegeneratingKey(false);
      setTimeout(() => setBroadcastResult(null), 5000);
    }
  };

  // Generate secured URLs with the API key attached
  const publicWebUrl = apiKey ? `${origin}/embed/sites?key=${apiKey}` : `${origin}/embed/sites`;
  const iframeCode = `<iframe\n  src="${publicWebUrl}"\n  width="100%"\n  height="800"\n  style="border:0;border-radius:14px;box-shadow:0 8px 30px rgba(0,0,0,0.15);width:100%;min-height:750px;"\n  allow="geolocation"\n  loading="lazy"\n  title="Media Buzz Outdoor Sites"\n></iframe>`;
  const scriptCode = `<!-- Media Buzz OOH Sites Auto-Sync Widget -->\n<div id="mb-sites-widget"></div>\n<script src="${origin}/embed.js"${apiKey ? ` data-api-key="${apiKey}"` : ''}></script>`;
  const apiUrl = apiKey ? `${origin}/api/public/sites?api_key=${apiKey}` : `${origin}/api/public/sites`;
  const sseUrl = apiKey ? `${origin}/api/public/sites/stream?api_key=${apiKey}` : `${origin}/api/public/sites/stream`;

  const copyToClipboard = (text, type) => {
    navigator.clipboard.writeText(text);
    setCopiedCode(type);
    setTimeout(() => setCopiedCode(null), 2500);
  };

  const handleTestBroadcast = async () => {
    setTestBroadcasting(true);
    setBroadcastResult(null);
    try {
      broadcastSitesUpdated();
      const res = await api.post('/public/sites/notify', { source: 'test_sync_button' });
      setBroadcastResult({ success: true, count: res.data?.clientCount ?? 0 });
    } catch (e) {
      setBroadcastResult({ success: false, message: e.message });
    } finally {
      setTestBroadcasting(false);
      setTimeout(() => setBroadcastResult(null), 4000);
    }
  };

  const maskedKeyDisplay = apiKey
    ? (keyMasked ? `${apiKey.slice(0, 10)}${'•'.repeat(24)}` : apiKey)
    : 'Loading API Key...';

  return (
    <div className="web-integration-settings-panel">
      {/* Top Banner / Integration Center */}
      <div className="web-integration-banner">
        <div className="web-integration-banner-left">
          <div className="web-integration-banner-title">
            <span>🌐 Web Integration & Real-time Site Directory</span>
            <span style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              fontSize: '11px',
              fontWeight: 800,
              padding: '4px 10px',
              borderRadius: '999px',
              background: 'rgba(16, 185, 129, 0.18)',
              color: '#34d399',
              border: '1px solid rgba(16, 185, 129, 0.4)'
            }}>
              <span className="mb-status-dot" style={{ width: '6px', height: '6px' }}></span>
              Live Auto-Sync Active
            </span>
            <span style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              fontSize: '11px',
              fontWeight: 800,
              padding: '4px 10px',
              borderRadius: '999px',
              background: 'rgba(56, 189, 248, 0.15)',
              color: '#38bdf8',
              border: '1px solid rgba(56, 189, 248, 0.4)'
            }}>
              🔒 Protected by API Key
            </span>
          </div>
          <p className="web-integration-banner-desc">
            Your outdoor advertising sites are integrated live into your website. The inventory API is protected with a secure API key so only authorized websites and applications can access your sites and rates.
          </p>
        </div>

        <div className="web-integration-banner-actions">
          <button
            type="button"
            className="scooh-btn purple-btn"
            onClick={() => window.open(publicWebUrl, '_blank')}
            title="Open live public integrated site in a new tab"
            style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            <span>🚀</span> Open Live Web View
          </button>

          <button
            type="button"
            className="scooh-btn"
            onClick={() => copyToClipboard(apiUrl, 'banner_json')}
            title="Copy full JSON API endpoint URL with your API key for external website integration"
            style={{ display: 'flex', alignItems: 'center', gap: '6px', color: '#38bdf8', borderColor: 'rgba(56,189,248,0.4)' }}
          >
            <span>{copiedCode === 'banner_json' ? '✓' : '🔗'}</span>
            {copiedCode === 'banner_json' ? 'Full JSON URL Copied!' : 'Copy Full JSON URL'}
          </button>

          <button
            type="button"
            className="scooh-btn"
            onClick={() => setShowEmbedKit(!showEmbedKit)}
            style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            <span>📋</span> {showEmbedKit ? 'Hide Integration Kit' : 'Get Embed & API URLs'}
          </button>

          <button
            type="button"
            className="scooh-btn"
            disabled={testBroadcasting}
            onClick={handleTestBroadcast}
            title="Broadcast test update to all connected web pages"
            style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            <span>{testBroadcasting ? '⏳' : '⚡'}</span> {testBroadcasting ? 'Broadcasting...' : 'Test Live Sync'}
          </button>
        </div>
      </div>

      {/* API Key Security Card */}
      <div style={{
        marginTop: '14px',
        padding: '16px 20px',
        background: 'rgba(15, 23, 42, 0.85)',
        border: '1px solid rgba(139, 92, 246, 0.35)',
        borderRadius: '12px',
        backdropFilter: 'blur(10px)',
        boxShadow: '0 4px 20px rgba(0, 0, 0, 0.25)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexWrap: 'wrap',
        gap: '12px'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
          <div style={{
            width: '38px',
            height: '38px',
            borderRadius: '10px',
            background: 'linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '18px',
            boxShadow: '0 2px 10px rgba(139, 92, 246, 0.4)'
          }}>
            🔑
          </div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontSize: '13px', fontWeight: 700, color: '#f8fafc' }}>
                Inventory Integration API Key
              </span>
              <span style={{
                fontSize: '10px',
                padding: '2px 7px',
                borderRadius: '6px',
                background: 'rgba(16, 185, 129, 0.15)',
                color: '#34d399',
                fontWeight: 700,
                border: '1px solid rgba(16, 185, 129, 0.3)'
              }}>
                Enforced & Private
              </span>
            </div>
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              marginTop: '4px'
            }}>
              <code style={{
                fontSize: '12.5px',
                fontFamily: 'monospace',
                background: 'rgba(0, 0, 0, 0.45)',
                padding: '3px 10px',
                borderRadius: '6px',
                border: '1px solid rgba(255, 255, 255, 0.1)',
                color: '#a78bfa',
                letterSpacing: keyMasked ? '1.5px' : '0.5px'
              }}>
                {loadingKey ? 'Fetching API Key...' : maskedKeyDisplay}
              </code>
              {apiKey && (
                <button
                  type="button"
                  onClick={() => setKeyMasked(!keyMasked)}
                  style={{
                    background: 'none',
                    border: 'none',
                    color: '#94a3b8',
                    cursor: 'pointer',
                    fontSize: '12px',
                    padding: '2px 6px',
                    textDecoration: 'underline'
                  }}
                  title={keyMasked ? 'Reveal full key' : 'Mask key'}
                >
                  {keyMasked ? 'Show' : 'Hide'}
                </button>
              )}
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
          <button
            type="button"
            className="scooh-btn purple-btn"
            disabled={!apiKey || loadingKey}
            onClick={() => copyToClipboard(apiKey, 'direct_api_key')}
            style={{ fontSize: '12px', padding: '6px 14px' }}
          >
            {copiedCode === 'direct_api_key' ? '✓ API Key Copied!' : '📋 Copy API Key'}
          </button>

          <button
            type="button"
            className="scooh-btn"
            disabled={regeneratingKey || loadingKey}
            onClick={handleRegenerateKey}
            style={{ fontSize: '12px', padding: '6px 14px', color: '#f87171', borderColor: 'rgba(239, 68, 68, 0.35)' }}
            title="Generate a new API key and invalidate the old one"
          >
            {regeneratingKey ? '⏳ Regenerating...' : '🔄 Regenerate Key'}
          </button>
        </div>
      </div>

      {broadcastResult && (
        <div className="scooh-banner" style={{
          display: 'block',
          marginTop: '12px',
          background: broadcastResult.success ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)',
          border: `1px solid ${broadcastResult.success ? '#10b981' : '#ef4444'}`,
          color: broadcastResult.success ? '#34d399' : '#f87171'
        }}>
          {broadcastResult.success
            ? (typeof broadcastResult.count === 'string'
                ? `⚡ ${broadcastResult.count}`
                : `⚡ Auto-sync broadcast signal sent successfully! Active listening clients: ${broadcastResult.count}`)
            : `Failed to broadcast sync: ${broadcastResult.message}`}
        </div>
      )}

      {/* Embed Code & API Drawer */}
      {showEmbedKit && (
        <div className="web-integration-kit-drawer" style={{ marginTop: '14px' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
            <h4 className="web-integration-kit-title">Website Integration & Full JSON API</h4>
            <span style={{ fontSize: '11.5px', color: '#94a3b8' }}>
              Secured with API Key • Compatible with WordPress, Webflow, Shopify, React & Mobile Apps
            </span>
          </div>

          <div className="web-integration-tabs-row">
            <button
              type="button"
              className={`scooh-btn ${activeCodeTab === 'api' ? 'purple-btn' : ''}`}
              onClick={() => setActiveCodeTab('api')}
              style={{ padding: '6px 12px', fontSize: '12px' }}
            >
              ⚡ Full JSON API (REST)
            </button>
            <button
              type="button"
              className={`scooh-btn ${activeCodeTab === 'iframe' ? 'purple-btn' : ''}`}
              onClick={() => setActiveCodeTab('iframe')}
              style={{ padding: '6px 12px', fontSize: '12px' }}
            >
              Responsive iFrame
            </button>
            <button
              type="button"
              className={`scooh-btn ${activeCodeTab === 'script' ? 'purple-btn' : ''}`}
              onClick={() => setActiveCodeTab('script')}
              style={{ padding: '6px 12px', fontSize: '12px' }}
            >
              Drop-in Script Tag
            </button>
            <button
              type="button"
              className={`scooh-btn ${activeCodeTab === 'link' ? 'purple-btn' : ''}`}
              onClick={() => setActiveCodeTab('link')}
              style={{ padding: '6px 12px', fontSize: '12px' }}
            >
              Direct Public URL
            </button>
          </div>

          {activeCodeTab === 'api' && (
            <div>
              <p style={{ margin: '0 0 10px', fontSize: '12.5px', color: '#cbd5e1' }}>
                Use this <strong>Full JSON API URL</strong> to pull live sites inventory and availability directly into WordPress, Webflow, mobile apps, or custom frontends. All photos are returned with full absolute URLs. <strong>Your API key is automatically embedded in the query string and headers below.</strong>
              </p>

              {/* Primary Full JSON API URL */}
              <div style={{ marginBottom: '14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
                  <strong style={{ fontSize: '12px', color: '#a78bfa' }}>Full JSON API Endpoint URL (with API key):</strong>
                  <span style={{ fontSize: '11px', color: '#34d399', fontWeight: 700 }}>✓ Key Authentication Active</span>
                </div>
                <div className="web-integration-code-box" style={{ fontSize: '13px', color: '#f8fafc', fontWeight: 600 }}>
                  {apiUrl}
                </div>
                <div style={{ display: 'flex', gap: '10px', marginTop: '8px', flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="scooh-btn purple-btn"
                    onClick={() => copyToClipboard(apiUrl, 'api_main')}
                  >
                    {copiedCode === 'api_main' ? '✓ Copied Full JSON URL!' : '📋 Copy Full JSON URL'}
                  </button>
                  <button
                    type="button"
                    className="scooh-btn"
                    onClick={() => window.open(apiUrl, '_blank')}
                  >
                    🚀 Open JSON Endpoint in New Tab ↗
                  </button>
                </div>
              </div>

              {/* Date Filtered URL */}
              <div style={{ marginBottom: '14px' }}>
                <strong style={{ fontSize: '11.5px', color: '#cbd5e1', display: 'block', marginBottom: '4px' }}>
                  Full JSON URL with Date-Availability Filter:
                </strong>
                <div className="web-integration-code-box" style={{ fontSize: '12px' }}>
                  {`${apiUrl}&date=${new Date().toISOString().slice(0, 10)}`}
                </div>
                <button
                  type="button"
                  className="scooh-btn"
                  onClick={() => copyToClipboard(`${apiUrl}&date=${new Date().toISOString().slice(0, 10)}`, 'api_date')}
                  style={{ marginTop: '6px', fontSize: '11.5px', padding: '5px 12px' }}
                >
                  {copiedCode === 'api_date' ? '✓ Copied Date URL!' : '📋 Copy Date Filtered URL'}
                </button>
              </div>

              {/* Real-time SSE Stream URL */}
              <div style={{ marginBottom: '14px' }}>
                <strong style={{ fontSize: '11.5px', color: '#cbd5e1', display: 'block', marginBottom: '4px' }}>
                  Full Real-time Server-Sent Events (SSE) Stream URL:
                </strong>
                <div className="web-integration-code-box" style={{ fontSize: '12px' }}>
                  {sseUrl}
                </div>
                <button
                  type="button"
                  className="scooh-btn"
                  onClick={() => copyToClipboard(sseUrl, 'api_stream')}
                  style={{ marginTop: '6px', fontSize: '11.5px', padding: '5px 12px' }}
                >
                  {copiedCode === 'api_stream' ? '✓ Copied Stream URL!' : '📋 Copy Stream URL'}
                </button>
              </div>

              {/* HTTP Header Authentication Option */}
              <div style={{ marginBottom: '14px' }}>
                <strong style={{ fontSize: '11.5px', color: '#cbd5e1', display: 'block', marginBottom: '4px' }}>
                  cURL / HTTP Header Authentication (Alternative to query string):
                </strong>
                <div className="web-integration-code-box" style={{ fontSize: '11.5px', color: '#38bdf8' }}>
                  {`curl -H "x-api-key: ${apiKey || 'YOUR_API_KEY'}" "${origin}/api/public/sites"`}
                </div>
              </div>

              {/* WordPress Integration Example */}
              <div style={{ marginTop: '14px', borderTop: '1px solid rgba(255,255,255,0.08)', paddingTop: '12px' }}>
                <strong style={{ fontSize: '12px', color: '#cbd5e1', display: 'block', marginBottom: '6px' }}>
                  WordPress PHP Integration (Copy into functions.php):
                </strong>
                <pre className="web-integration-code-box" style={{ whiteSpace: 'pre-wrap', fontSize: '11.5px' }}>
{`// Fetch live Media Buzz sites in WordPress (Secured with API Key):
$response = wp_remote_get('${origin}/api/public/sites', [
    'headers' => [
        'x-api-key' => '${apiKey || 'YOUR_API_KEY'}'
    ],
    'timeout' => 15
]);

if (!is_wp_error($response)) {
    $body = wp_remote_retrieve_body($response);
    $data = json_decode($body, true);
    $sites = $data['sites']; // Array of sites with live availability and full photo URLs
}`}
                </pre>
              </div>
            </div>
          )}

          {activeCodeTab === 'iframe' && (
            <div>
              <p style={{ margin: '0 0 8px', fontSize: '12px', color: '#94a3b8' }}>
                Paste this snippet into any page, blog post, or custom HTML block on your website. The iframe automatically includes your API key for seamless access:
              </p>
              <div className="web-integration-code-box">{iframeCode}</div>
              <button
                type="button"
                className="scooh-btn purple-btn"
                onClick={() => copyToClipboard(iframeCode, 'iframe')}
                style={{ marginTop: '10px' }}
              >
                {copiedCode === 'iframe' ? '✓ Copied to Clipboard!' : 'Copy iFrame Code'}
              </button>
            </div>
          )}

          {activeCodeTab === 'script' && (
            <div>
              <p style={{ margin: '0 0 8px', fontSize: '12px', color: '#94a3b8' }}>
                Add this script tag wherever you want the inventory widget to appear on your site. The <code>data-api-key</code> attribute authorizes the embed:
              </p>
              <div className="web-integration-code-box">{scriptCode}</div>
              <button
                type="button"
                className="scooh-btn purple-btn"
                onClick={() => copyToClipboard(scriptCode, 'script')}
                style={{ marginTop: '10px' }}
              >
                {copiedCode === 'script' ? '✓ Copied to Clipboard!' : 'Copy Script Code'}
              </button>
            </div>
          )}

          {activeCodeTab === 'link' && (
            <div>
              <p style={{ margin: '0 0 8px', fontSize: '12px', color: '#94a3b8' }}>
                Share this direct link with authorized clients or link to it from your main website navigation:
              </p>
              <div className="web-integration-code-box">{publicWebUrl}</div>
              <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                <button
                  type="button"
                  className="scooh-btn purple-btn"
                  onClick={() => copyToClipboard(publicWebUrl, 'link')}
                >
                  {copiedCode === 'link' ? '✓ Copied Link!' : 'Copy Public Link'}
                </button>
                <button
                  type="button"
                  className="scooh-btn"
                  onClick={() => window.open(publicWebUrl, '_blank')}
                >
                  Open in New Tab ↗
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Sites Directory */}
      <div className="web-integration-sites-wrapper">
        {children}
      </div>
    </div>
  );
}
