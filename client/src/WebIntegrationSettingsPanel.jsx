import React, { useState } from 'react';
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
  const [copiedCode, setCopiedCode] = useState(null);
  const [activeCodeTab, setActiveCodeTab] = useState('iframe');
  const [testBroadcasting, setTestBroadcasting] = useState(false);
  const [broadcastResult, setBroadcastResult] = useState(null);
  const [showEmbedKit, setShowEmbedKit] = useState(false);

  const origin = window.location.origin;
  const publicWebUrl = `${origin}/embed/sites`;
  const iframeCode = `<iframe\n  src="${publicWebUrl}"\n  width="100%"\n  height="800"\n  style="border:0;border-radius:14px;box-shadow:0 8px 30px rgba(0,0,0,0.15);width:100%;min-height:750px;"\n  allow="geolocation"\n  loading="lazy"\n  title="Media Buzz Outdoor Sites"\n></iframe>`;
  const scriptCode = `<!-- Media Buzz OOH Sites Auto-Sync Widget -->\n<div id="mb-sites-widget"></div>\n<script src="${origin}/embed.js"></script>`;
  const apiUrl = `${origin}/api/public/sites`;
  const sseUrl = `${origin}/api/public/sites/stream`;

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
          </div>
          <p className="web-integration-banner-desc">
            Your outdoor advertising sites are integrated live into your website. When any site is created, modified, photos updated, or availability dates change below, the integrated site automatically updates in real-time.
          </p>
        </div>

        <div className="web-integration-banner-actions">
          <button
            type="button"
            className="scooh-btn purple-btn"
            onClick={() => window.open('/embed/sites', '_blank')}
            title="Open live public integrated site in a new tab"
            style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            <span>🚀</span> Open Live Web View
          </button>

          <button
            type="button"
            className="scooh-btn"
            onClick={() => copyToClipboard(apiUrl, 'banner_json')}
            title="Copy full JSON API endpoint URL for external website integration"
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

      {broadcastResult && (
        <div className="scooh-banner" style={{
          display: 'block',
          background: broadcastResult.success ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)',
          border: `1px solid ${broadcastResult.success ? '#10b981' : '#ef4444'}`,
          color: broadcastResult.success ? '#34d399' : '#f87171'
        }}>
          {broadcastResult.success
            ? `⚡ Auto-sync broadcast signal sent successfully! Active listening clients: ${broadcastResult.count}`
            : `Failed to broadcast sync: ${broadcastResult.message}`}
        </div>
      )}

      {/* Embed Code & API Drawer */}
      {showEmbedKit && (
        <div className="web-integration-kit-drawer">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
            <h4 className="web-integration-kit-title">Website Integration & Full JSON API</h4>
            <span style={{ fontSize: '11.5px', color: '#94a3b8' }}>Compatible with WordPress, Webflow, Shopify, React, PHP & Mobile Apps</span>
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
                Use this <strong>Full JSON API URL</strong> to pull live sites inventory and availability directly into WordPress, Webflow, mobile apps, or custom frontends. All photos are returned with full absolute URLs.
              </p>

              {/* Primary Full JSON API URL */}
              <div style={{ marginBottom: '14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
                  <strong style={{ fontSize: '12px', color: '#a78bfa' }}>Full JSON API Endpoint URL:</strong>
                  <span style={{ fontSize: '11px', color: '#34d399', fontWeight: 700 }}>✓ CORS Enabled for All Domains</span>
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
                  {`${apiUrl}?date=${new Date().toISOString().slice(0, 10)}`}
                </div>
                <button
                  type="button"
                  className="scooh-btn"
                  onClick={() => copyToClipboard(`${apiUrl}?date=${new Date().toISOString().slice(0, 10)}`, 'api_date')}
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

              {/* WordPress Integration Example */}
              <div style={{ marginTop: '14px', borderTop: '1px solid rgba(255,255,255,0.08)', paddingTop: '12px' }}>
                <strong style={{ fontSize: '12px', color: '#cbd5e1', display: 'block', marginBottom: '6px' }}>
                  WordPress PHP Integration (Copy into functions.php):
                </strong>
                <pre className="web-integration-code-box" style={{ whiteSpace: 'pre-wrap', fontSize: '11.5px' }}>
{`// Fetch live Media Buzz sites in WordPress:
$response = wp_remote_get('${apiUrl}');
if (!is_wp_error($response)) {
    $body = wp_remote_retrieve_body($response);
    $data = json_decode($body, true);
    $sites = $data['sites']; // Array of sites with full photo URLs
}`}
                </pre>
              </div>
            </div>
          )}

          {activeCodeTab === 'iframe' && (
            <div>
              <p style={{ margin: '0 0 8px', fontSize: '12px', color: '#94a3b8' }}>
                Paste this snippet into any page, blog post, or custom HTML block on your website:
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
                Add this script tag wherever you want the inventory widget to appear on your site:
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
                Share this direct public link with clients or link to it from your main website navigation:
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
