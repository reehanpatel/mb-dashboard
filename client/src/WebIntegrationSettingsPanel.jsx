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
            onClick={() => setShowEmbedKit(!showEmbedKit)}
            style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            <span>📋</span> {showEmbedKit ? 'Hide Embed Kit' : 'Get Embed Code'}
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

      {/* Embed Code Drawer */}
      {showEmbedKit && (
        <div className="web-integration-kit-drawer">
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
            <h4 className="web-integration-kit-title">Website Integration Code & Embed Options</h4>
            <span style={{ fontSize: '11.5px', color: '#94a3b8' }}>Compatible with WordPress, Webflow, Shopify, React, HTML</span>
          </div>

          <div className="web-integration-tabs-row">
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
            <button
              type="button"
              className={`scooh-btn ${activeCodeTab === 'api' ? 'purple-btn' : ''}`}
              onClick={() => setActiveCodeTab('api')}
              style={{ padding: '6px 12px', fontSize: '12px' }}
            >
              REST API & SSE Stream
            </button>
          </div>

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

          {activeCodeTab === 'api' && (
            <div>
              <p style={{ margin: '0 0 8px', fontSize: '12px', color: '#94a3b8' }}>
                For custom web developers building headless/custom frontends:
              </p>
              <div style={{ marginBottom: '8px' }}>
                <strong style={{ fontSize: '11px', color: '#cbd5e1' }}>JSON API:</strong>
                <div className="web-integration-code-box">GET {apiUrl}</div>
              </div>
              <div>
                <strong style={{ fontSize: '11px', color: '#cbd5e1' }}>Real-time SSE Stream:</strong>
                <div className="web-integration-code-box">GET {sseUrl}</div>
              </div>
              <button
                type="button"
                className="scooh-btn purple-btn"
                onClick={() => copyToClipboard(apiUrl, 'api')}
                style={{ marginTop: '10px' }}
              >
                {copiedCode === 'api' ? '✓ Copied API URL!' : 'Copy JSON API URL'}
              </button>
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
