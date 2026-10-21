/**
 * Media Buzz OOH - Web Integration Embed Script
 * Embed your live outdoor advertising sites catalog onto any website.
 *
 * Usage:
 * <div id="mb-sites-widget"></div>
 * <script src="https://YOUR-DOMAIN/embed.js"></script>
 */
(function() {
  'use strict';
  
  // Find current script tag to determine dashboard host
  var scripts = document.getElementsByTagName('script');
  var currentScript = document.currentScript || scripts[scripts.length - 1];
  var scriptSrc = currentScript ? currentScript.src : window.location.origin;
  var baseUrl = scriptSrc ? new URL(scriptSrc).origin : window.location.origin;

  // Find target container or auto-create one
  var container = document.getElementById('mb-sites-widget');
  if (!container) {
    container = document.createElement('div');
    container.id = 'mb-sites-widget';
    if (currentScript && currentScript.parentNode) {
      currentScript.parentNode.insertBefore(container, currentScript);
    } else {
      document.body.appendChild(container);
    }
  }

  // Create responsive iframe
  var iframe = document.createElement('iframe');
  iframe.src = baseUrl + '/embed/sites';
  iframe.style.width = '100%';
  iframe.style.minHeight = '750px';
  iframe.style.border = 'none';
  iframe.style.borderRadius = '14px';
  iframe.style.boxShadow = '0 10px 30px rgba(0, 0, 0, 0.15)';
  iframe.style.display = 'block';
  iframe.setAttribute('loading', 'lazy');
  iframe.setAttribute('allow', 'geolocation');
  iframe.title = 'Media Buzz OOH Sites Inventory';

  container.innerHTML = '';
  container.appendChild(iframe);

  console.log('[Media Buzz] Web Integration Widget initialized successfully from:', baseUrl);
})();
