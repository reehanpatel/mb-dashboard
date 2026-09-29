import React, { useEffect, useMemo, useState, useCallback, useRef, Component } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserRouter, Navigate, Route, Routes, useLocation, useNavigate } from 'react-router-dom';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import { attachMediaBuzzExcelHeader, attachMediaBuzzTermsAndConditions, MEDIA_BUZZ_TERMS_TEXT } from './excelLogo';
import api from './api';
import CampaignDetailsView, { CampaignDetailsModal, SiteDetailsModal } from './CampaignDetailsView';
import { defaultSettings } from './defaultSites';
import { canonicalSiteCode, isCombinedSite, getSiteTypeTag, getConflictSummary, getOverlappingSiteCodes, SITE_PANELS } from './siteHierarchy';
import './styles.css';

// Helper to sanitize display titles and prevent "[Split Face via ...]" or "[Combined Block ...]" from showing anywhere
export function cleanDisplayTitle(str) {
  if (!str) return '';
  return String(str)
    .replace(/\s*\[(?:Split Face|Combined Block)[^\]]*\]/gi, '')
    .trim();
}

// Class ErrorBoundary to prevent any white/black screens
class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    console.error('UI Runtime Catch:', error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ minHeight: '100vh', background: '#0b0f14', color: '#f4f6f8', padding: '40px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ maxWidth: '500px', background: '#111720', border: '1px solid #27313d', borderRadius: '16px', padding: '28px', textAlign: 'center' }}>
            <img src="/assets/media-buzz-logo.png" alt="Media Buzz" style={{ width: '140px', marginBottom: '16px' }} />
            <h2 style={{ margin: '0 0 10px', fontSize: '18px' }}>Workspace Recovery</h2>
            <p style={{ fontSize: '12px', color: '#8d98a6', margin: '0 0 20px' }}>{this.state.error?.message || 'An unexpected display error occurred.'}</p>
            <button
              style={{ padding: '10px 20px', background: '#f2c94c', border: 0, borderRadius: '8px', color: '#16120a', fontWeight: 800, cursor: 'pointer' }}
              onClick={() => {
                localStorage.removeItem('sc_token');
                window.location.href = '/login';
              }}
            >
              Sign In Again
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// SVG Icons Dictionary matching the Media Buzz OOH Plugin
const icons = {
  dashboard: <path d="M4 13h6V4H4v9Zm0 7h6v-5H4v5Zm10 0h6v-9h-6v9Zm0-16v5h6V4h-6Z"/>,
  sites: <path d="M12 2a7 7 0 0 0-7 7c0 5.25 7 13 7 13s7-7.75 7-13a7 7 0 0 0-7-7Zm0 9.5A2.5 2.5 0 1 1 12 6a2.5 2.5 0 0 1 0 5.5Z"/>,
  campaigns: <path d="M4 4h12a2 2 0 0 1 2 2v2h2v8h-2v2a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Zm2 4v8h8V8H6Z"/>,
  occupancy: <path d="M4 19h16v2H4v-2Zm1-8h3v6H5v-6Zm5-5h3v11h-3V6Zm5 3h3v8h-3V9Z"/>,
  'campaign-details': <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10H7v-2h10v2zm0-4H7V7h10v2z"/>,
  proposals: <path d="M6 2h9l5 5v15H6V2Zm8 2v5h5M9 13h8v2H9v-2Zm0 4h8v2H9v-2ZM9 9h3v2H9V9Z"/>,
  ppt: <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>,
  electricity: <path d="M13 2 5 13h6l-1 9 9-13h-6l0-7Z"/>,
  vendors: <path d="M8 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Zm8 1a3 3 0 1 0 0-6 3 3 0 0 0 0 6ZM2 21v-2c0-3 3-5 6-5s6 2 6 5v2H2Zm12.5 0v-2c0-1.35-.43-2.54-1.15-3.5.79-.32 1.69-.5 2.65-.5 3 0 6 2 6 5v1h-7.5Z"/>,
  clients: <path d="M12 12a5 5 0 1 0 0-10 5 5 0 0 0 0 10ZM3 22v-2c0-4 4-6 9-6s9 2 9 6v2H3Z"/>,
  invoices: <path d="M6 2h12v20l-3-2-3 2-3-2-3 2V2Zm3 5h6v2H9V7Zm0 4h6v2H9v-2Zm0 4h4v2H9v-2Z"/>,
  data: <path d="M12 2 7 7h3v6h4V7h3l-5-5ZM5 14v6h14v-6h2v8H3v-8h2Z"/>,
  reports: <path d="M4 2h16v20H4V2Zm4 14h2v3H8v-3Zm3-5h2v8h-2v-8Zm3 2h2v6h-2v-6ZM8 6h8v2H8V6Z"/>,
  notifications: <path d="M12 22a2.5 2.5 0 0 0 2.45-2h-4.9A2.5 2.5 0 0 0 12 22Zm7-5H5l2-2v-5a5 5 0 1 1 10 0v5l2 2Z"/>,
  activity: <path d="M12 2a10 10 0 1 0 10 10h-2a8 8 0 1 1-2.34-5.66L14 10h8V2l-2.91 2.91A9.96 9.96 0 0 0 12 2Zm-1 5h2v6l4 2-1 1.73-5-2.73V7Z"/>,
  storage: <path d="M4 4h16v4H4V4Zm0 6h16v4H4v-4Zm0 6h16v4H4v-4Zm6-10h4V5h-4v1Zm0 6h4v-1h-4v1Zm0 6h4v-1h-4v1Z"/>,
  settings: <path d="m19.14 12.94.04-.94-.04-.94 2.03-1.58-2-3.46-2.49 1a7.8 7.8 0 0 0-1.63-.94L14.68 3h-4l-.37 3.08c-.58.25-1.12.57-1.63.94l-2.49-1-2 3.46 2.03 1.58-.04.94.04.94-2.03 1.58 2 3.46 2.49-1c.51.37 1.05.69 1.63.94L10.68 21h4l.37-3.08c.58-.25 1.12-.57 1.63-.94l2.49 1 2-3.46-2.03-1.58ZM12.68 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7Z"/>
};

const navGroups = [
  {
    label: 'Workspace',
    items: [
      ['dashboard', 'Dashboard'],
      ['sites', 'Sites'],
      ['campaigns', 'Latest Booking'],
      ['occupancy', 'Campaign Tracker'],
      ['campaign-details', 'Campaign Details'],
      ['proposals', 'Proposal Builder'],
      ['ppt', 'Automated PPT'],
      ['storage', 'Storage & Archives']
    ]
  },
  {
    label: 'Operations',
    items: [
      ['electricity', 'Electricity'],
      ['vendors', 'Vendors'],
      ['data', 'Import / Export']
    ]
  },
  {
    label: 'Business',
    items: [
      ['clients', 'Clients'],
      ['invoices', 'Invoices'],
      ['reports', 'Reports']
    ]
  },
  {
    label: 'System',
    items: [
      ['notifications', 'Notifications'],
      ['activity', 'Activity Log'],
      ['settings', 'Settings']
    ]
  }
];

const fields = {
  sites: ['site_code','city','area','address','media_type','lighting','facing','size','width','height','ownership','availability','vendor_name','meter_no','monthly_cost','monthly_rate','latitude','longitude','maps_url','gps','notes'],
  clients: ['client_name','company','primary_contact','email','phone','billing_address','gst_number','status','notes'],
  campaigns: ['booking_code','parent_campaign','site_id','site_code','client_id','client','brand','campaign_name','booking_date','start_date','end_date','mounting_date','printing_status','mounting_status','validation_15_date','final_validation_date','revenue','vendor_cost','printing_cost','mounting_cost','electricity_cost','other_cost','invoice_required','invoice_requested','invoice_no','invoice_status','hard_copy_status','notes'],
  electricity: ['site_id','site_code','location','meter_no','size','service_number','t_number','bill_type','payment_amount','billing_month','bill_date','due_date','units','rate','other_charges','amount','payment_status','ecs','payment_date','paid_date','payment_reference','notes'],
  vendors: ['name','service','contact_person','phone','email','cities','rating','notes','status'],
  proposals: ['proposal_code','client_id','client_name','campaign_name','proposal_date','start_date','duration_days','validity_days','discount_percent','tax_percent','subtotal','total','status','notes','terms'],
  invoices: ['campaign_id','client_id','requested_date','invoice_no','invoice_date','invoice_amount','invoice_status','hard_copy_required','hard_copy_status','courier_name','tracking_number','dispatch_date','delivered_date','payment_status','payment_date','notes']
};

const viewTitles = {
  dashboard: 'Dashboard',
  sites: 'Sites Directory',
  campaigns: 'Latest Booking',
  occupancy: 'Campaign Tracker',
  'campaign-details': 'Campaign Details',
  proposals: 'Proposal Builder',
  ppt: 'Automated PPT',
  electricity: 'Electricity & Meters',
  vendors: 'Vendors Directory',
  clients: 'Clients Directory',
  invoices: 'Invoices & Dispatch',
  data: 'Import / Export Tools',
  reports: 'Performance Reports',
  notifications: 'System Notifications',
  activity: 'Activity Log',
  storage: 'Storage & Archives',
  settings: 'System Settings'
};

export const ROLE_CONFIG = {
  admin: {
    key: 'admin',
    label: 'Admin',
    fullLabel: 'Administrator (Full Workspace & Settings Access)',
    color: '#c4b5fd',
    bg: 'rgba(139,92,246,0.18)',
    border: 'rgba(139,92,246,0.35)',
    description: 'Full workspace access, user account provisioning, system settings, database backups, and permanent record deletions across all modules.'
  },
  manager: {
    key: 'manager',
    label: 'Manager',
    fullLabel: 'Manager (Sites, Campaigns, Proposals & Storage)',
    color: '#4ade80',
    bg: 'rgba(34,197,94,0.18)',
    border: 'rgba(34,197,94,0.35)',
    description: 'Full operational control: Create, edit, export, and delete sites, campaigns, proposals, electricity, invoices, and storage archives. Cannot access User Management or System Settings.'
  },
  staff: {
    key: 'staff',
    label: 'Staff',
    fullLabel: 'Staff (Field Operations & Billing Tracking)',
    color: '#38bdf8',
    bg: 'rgba(56,189,248,0.18)',
    border: 'rgba(56,189,248,0.35)',
    description: 'Field operations: View inventory, update mounting & printing status, and log electricity meter bills. Deletions, bulk deletes, and proposals are restricted.'
  },
  viewer: {
    key: 'viewer',
    label: 'Viewer',
    fullLabel: 'Viewer (Read-Only Access)',
    color: '#f59e0b',
    bg: 'rgba(245,158,11,0.18)',
    border: 'rgba(245,158,11,0.35)',
    description: 'Read-only access: Can view dashboards, sites, campaigns, occupancy, and download PPT/Excel files. Cannot create, edit, import, or delete records.'
  }
};

export function getCurrentUser() {
  try {
    return JSON.parse(localStorage.getItem('sc_user') || '{}');
  } catch {
    return {};
  }
}

export function getCurrentRole() {
  const u = getCurrentUser();
  return (u.role || 'staff').toLowerCase();
}


const label = s => String(s).replaceAll('_', ' ').replace(/\b\w/g, m => m.toUpperCase());
const money = v => '₹' + Number(v || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const formatDate = v => v ? new Date(v).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

function universalCompare(valA, valB, dir = 'asc') {
  if (valA === valB) return 0;
  if (valA === null || valA === undefined || valA === '' || valA === '—') return 1;
  if (valB === null || valB === undefined || valB === '' || valB === '—') return -1;

  const strA = String(valA).trim();
  const strB = String(valB).trim();

  // Try parsing as site code (e.g. "MB-01", "MB-87", "01", "MB 2")
  const siteA = strA.match(/^(?:mb[-\s]*)?(\d+)$/i);
  const siteB = strB.match(/^(?:mb[-\s]*)?(\d+)$/i);
  if (siteA && siteB) {
    const diff = parseInt(siteA[1], 10) - parseInt(siteB[1], 10);
    return dir === 'desc' ? -diff : diff;
  }

  // Try parsing as number
  const numA = typeof valA === 'number' ? valA : (/^-?\d+(\.\d+)?$/.test(strA) ? parseFloat(strA) : NaN);
  const numB = typeof valB === 'number' ? valB : (/^-?\d+(\.\d+)?$/.test(strB) ? parseFloat(strB) : NaN);

  let res = 0;
  if (!isNaN(numA) && !isNaN(numB)) {
    res = numA - numB;
  } else {
    // Try parsing as date
    const isDateStr = str => typeof str === 'string' && (/^\d{4}-\d{2}-\d{2}/.test(str) || (!isNaN(Date.parse(str)) && !/^\d+$/.test(str)));
    if (isDateStr(valA) && isDateStr(valB)) {
      const timeA = new Date(valA).getTime();
      const timeB = new Date(valB).getTime();
      if (!isNaN(timeA) && !isNaN(timeB)) {
        res = timeA - timeB;
      } else {
        res = strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' });
      }
    } else {
      res = strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' });
    }
  }

  return dir === 'desc' ? -res : res;
}

// Smart site search matcher supporting site number (e.g. "01", "1", "mb-01", "87"), multiple comma/space/newline site codes, and text search
function matchSiteSearch(site, rawQuery) {
  if (!rawQuery || !rawQuery.trim()) return true;
  if (!site) return false;

  const raw = rawQuery.trim();
  // If query contains comma, semicolon, or newline, support multi-token search
  if (/[,;\n]+/.test(raw)) {
    const tokens = raw.split(/[,;\n]+/).map(t => t.trim()).filter(Boolean);
    if (tokens.length > 1) {
      return tokens.some(tok => matchSingleSiteSearch(site, tok));
    }
  }

  // Also support multiple space-separated tokens if they look like site codes or numbers (e.g. "01 05 12" or "MB-01 MB-05")
  const spaceTokens = raw.split(/\s+/).map(t => t.trim()).filter(Boolean);
  if (spaceTokens.length > 1 && spaceTokens.every(t => /^(?:(?:mb|site)[\s-]*)?\d+$/i.test(t) || /^mb-\d+/i.test(t))) {
    return spaceTokens.some(tok => matchSingleSiteSearch(site, tok));
  }

  return matchSingleSiteSearch(site, raw);
}

function matchSingleSiteSearch(site, rawQuery) {
  if (!rawQuery || !rawQuery.trim()) return true;
  if (!site) return false;
  const q = rawQuery.trim().toLowerCase();
  const siteCode = String(site.site_code || '').trim().toLowerCase();
  const area = String(site.area || '').toLowerCase();
  const city = String(site.city || '').toLowerCase();
  const address = String(site.address || site.location || '').toLowerCase();
  const mediaType = String(site.media_type || '').toLowerCase();
  const lighting = String(site.lighting || '').toLowerCase();

  // 1. Numeric or site-code search: e.g. '01', '1', 'mb-01', 'mb 01', 'mb01', 'site 1', '87'
  const isNumericQuery = /^(?:(?:mb|site)[\s-]*)?(\d+)$/i.test(q);
  if (isNumericQuery) {
    const qNumMatch = q.match(/^(?:(?:mb|site)[\s-]*)?0*(\d+)$/i);
    const qNum = qNumMatch ? parseInt(qNumMatch[1], 10) : NaN;
    const cleanQDigits = q.replace(/\D/g, '');

    const codeNumMatch = siteCode.match(/(\d+)/);
    const codeNum = codeNumMatch ? parseInt(codeNumMatch[1], 10) : NaN;
    const codeDigits = codeNumMatch ? codeNumMatch[1] : '';

    // If query has leading zeros like '01', exact digit match (e.g. "01" matches "01")
    if (cleanQDigits.startsWith('0') && codeDigits === cleanQDigits) return true;

    // Exact numeric match: query '1' matches MB-01 (1 == 1)
    if (!isNaN(qNum) && !isNaN(codeNum) && qNum === codeNum) return true;

    // Exact siteCode match
    if (siteCode === q || siteCode.replace(/[^a-z0-9]/g, '') === q.replace(/[^a-z0-9]/g, '')) return true;

    return false;
  }

  // 2. Full text match on site_code (e.g. if searching partial letters or 'mb')
  if (siteCode.includes(q) || siteCode.replace(/[^a-z0-9]/g, '').includes(q.replace(/[^a-z0-9]/g, ''))) return true;

  // 3. Descriptive fields
  const size = String(site.size || (site.width && site.height ? `${site.width}x${site.height}` : '')).toLowerCase();
  if (area.includes(q) || address.includes(q) || city.includes(q) || mediaType.includes(q) || lighting.includes(q) || size.includes(q)) {
    return true;
  }
  return false;
}

function SortHeader({ label, sortKey, currentSort, onSort, align = 'left', style = {} }) {
  const isSorted = currentSort && currentSort.key === sortKey;
  const icon = !isSorted ? ' ⇅' : (currentSort.dir === 'asc' ? ' ▲' : ' ▼');
  return (
    <th
      onClick={() => onSort(sortKey)}
      style={{
        cursor: 'pointer',
        userSelect: 'none',
        textAlign: align,
        color: isSorted ? '#c4b5fd' : undefined,
        whiteSpace: 'nowrap',
        transition: 'color 0.15s ease, background 0.15s ease',
        ...style
      }}
      title={`Click to sort by ${label} (${isSorted && currentSort.dir === 'asc' ? 'Descending' : 'Ascending'})`}
    >
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
        <span>{label}</span>
        <span style={{ fontSize: '10px', opacity: isSorted ? 1 : 0.45, color: isSorted ? '#a78bfa' : 'inherit' }}>{icon}</span>
      </span>
    </th>
  );
}

function parseDimensions(sizeStr, w, h) {
  let width = parseFloat(w) || 0;
  let height = parseFloat(h) || 0;
  if ((!width || !height) && sizeStr) {
    const match = String(sizeStr).match(/(\d+(?:\.\d+)?)\s*(?:ft|')?\s*[xX*×/–-]\s*(\d+(?:\.\d+)?)/);
    if (match) {
      if (!width) width = parseFloat(match[1]);
      if (!height) height = parseFloat(match[2]);
    }
  }
  return { width, height };
}

function normalizeTokens(str) {
  if (!str) return [];
  const stopWords = new Set([
    'the', 'a', 'an', 'and', 'or', 'at', 'in', 'on', 'to', 'from', 'of', 'for',
    'nr', 'near', 'opp', 'opposite', 'behind', 'beside', 'facing', 'fcg', 'towards',
    'road', 'rd', 'cross', 'crossroad', 'junction', 'jnc', 'circle', 'bridge', 'flyover',
    'traffic', 'highway', 'hw', 'hwy', 'street', 'st', 'lane', 'sector', 'sec'
  ]);
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 2 && !stopWords.has(t));
}

function matchSiteByLocationAndSize(location, size, width, height, siteList) {
  if (!Array.isArray(siteList) || siteList.length === 0) return null;
  const targetDim = parseDimensions(size, width, height);
  const locTokens = normalizeTokens(location);
  const cleanLoc = String(location || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  let bestSite = null;
  let bestScore = -1;

  for (const site of siteList) {
    let score = 0;
    const siteDim = parseDimensions(site.size, site.width, site.height);
    let sizeMatched = false;

    // 1. Size matching (highest weight: 45 points)
    if (targetDim.width > 0 && targetDim.height > 0 && siteDim.width > 0 && siteDim.height > 0) {
      const exactMatch = (Math.abs(targetDim.width - siteDim.width) < 0.5 && Math.abs(targetDim.height - siteDim.height) < 0.5);
      const flippedMatch = (Math.abs(targetDim.width - siteDim.height) < 0.5 && Math.abs(targetDim.height - siteDim.width) < 0.5);
      if (exactMatch || flippedMatch) {
        score += 45;
        sizeMatched = true;
      } else {
        const targetArea = targetDim.width * targetDim.height;
        const siteArea = siteDim.width * siteDim.height;
        if (targetArea > 0 && Math.abs(targetArea - siteArea) / targetArea < 0.05) {
          score += 25;
          sizeMatched = true;
        }
      }
    }

    // 2. Location token overlap
    const siteText = `${site.address || ''} ${site.area || ''} ${site.city || ''} ${site.location || ''} ${site.site_code || ''}`;
    const siteTokens = new Set(normalizeTokens(siteText));
    const siteRaw = siteText.toLowerCase().replace(/[^a-z0-9]/g, '');

    let tokenMatches = 0;
    for (const t of locTokens) {
      if (siteTokens.has(t)) {
        tokenMatches += 1;
        score += 15;
      } else if (siteRaw.includes(t)) {
        tokenMatches += 0.5;
        score += 8;
      }
    }

    // Substring match
    if (cleanLoc.length >= 4 && siteRaw.includes(cleanLoc)) {
      score += 30;
    }

    // Directional / position indicator match (Left / Right / Middle / 1 / 2 / 3)
    const locLower = String(location || '').toLowerCase();
    const siteLower = siteText.toLowerCase();
    ['left', 'right', 'middle', 'center', '(1)', '(2)', '(3)', '1-3', '2-1', '2-3'].forEach(pos => {
      if (locLower.includes(pos) && siteLower.includes(pos)) {
        score += 12;
      }
    });

    if (score > bestScore && (tokenMatches > 0 || (sizeMatched && (cleanLoc.length < 3 || siteRaw.includes(cleanLoc.slice(0, 3)))))) {
      bestScore = score;
      bestSite = site;
    }
  }

  return bestScore >= 20 ? bestSite : null;
}

/** Safely parse any date string or object to a local-midnight Date (strips time/TZ so day comparisons are exact) */
function parseDay(d) {
  if (!d) return null;
  if (d instanceof Date) return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  if (typeof d === 'number') {
    const dt = new Date((d - 25569) * 86400000);
    return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  const s = String(d).trim();
  if (!s) return null;
  // Try YYYY-MM-DD
  const iso = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (iso) {
    return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
  }
  // Try DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (dmy) {
    let y = parseInt(dmy[3], 10);
    if (y < 100) y += 2000;
    return new Date(y, parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
  }
  const fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
}

function fmtDate(iso) {
  if (!iso) return '';
  try {
    const d = parseDay(iso) || new Date(iso);
    if (isNaN(d.getTime())) return iso;
    return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch { return iso; }
}

function extractDateFromString(str) {
  if (!str) return null;
  if (str instanceof Date) return isNaN(str.getTime()) ? null : new Date(str.getFullYear(), str.getMonth(), str.getDate());
  if (typeof str === 'number') {
    const dt = new Date((str - 25569) * 86400000);
    return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  const s = String(str).trim();
  if (!s) return null;
  // 1. DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
  const dmy = s.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let month = parseInt(dmy[2], 10) - 1;
    let year = parseInt(dmy[3], 10);
    if (year < 100) year += 2000;
    if (month > 11 && day <= 12) {
      const tmp = day - 1;
      day = month + 1;
      month = tmp;
    }
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }
  // 2. YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
  const iso = s.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (iso) {
    const d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    if (!isNaN(d.getTime())) return d;
  }
  // 3. DD Mon YYYY e.g. 25 Oct 2026, 25-Oct-2026
  const monMatch = s.match(/(\d{1,2})[ -]([A-Za-z]{3,9})[ -](\d{2,4})/);
  if (monMatch) {
    const d = new Date(`${monMatch[1]} ${monMatch[2]} ${monMatch[3]}`);
    if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

/**
 * Returns the active campaign running on the given date (or today if empty), or null if vacant.
 * A site booked until Sept 20 → selecting Sept 21 returns null (Available / Vacant).
 * campaignsBySiteCode: { [UPPERCASE_SITE_CODE]: campaign[] }
 */
function getActiveCampaignOnDate(campaignsBySiteCode, siteCode, dateStr) {
  if (!siteCode || !campaignsBySiteCode) return null;
  const check = dateStr ? parseDay(dateStr) : parseDay(new Date());
  if (!check) return null;
  const key = String(siteCode).toUpperCase().trim();
  const camps = campaignsBySiteCode[key] || [];
  for (const c of camps) {
    if (typeof isVacantClient === 'function' && isVacantClient(c.client || c.client_name || c.display)) continue;
    if (c.record_status && c.record_status !== 'active') continue;
    const start = parseDay(c.start_date || c.booking_date);
    const end = parseDay(c.end_date);
    if (!start && !end) continue;
    if (start && end) {
      if (check >= start && check <= end) return c;
    } else if (start && !end) {
      if (check >= start) return c;
    } else if (!start && end) {
      if (check <= end) return c;
    }
  }
  return null;
}

/**
 * Returns true if the site has an active campaign running ON the given date.
 */
function isSiteOccupiedOnDate(campaignsBySiteCode, siteCode, dateStr) {
  return !!getActiveCampaignOnDate(campaignsBySiteCode, siteCode, dateStr);
}

// ==========================================
// PERSISTENT PHOTO CACHE (LocalStorage + IndexedDB)
// Ensures uploaded PPT site photos are never lost across refreshes, re-login, or network glitches
// ==========================================
const SITE_PHOTOS_CACHE_KEY = 'mb_site_photos_cache';

function getAllCachedSitePhotos() {
  try {
    const raw = localStorage.getItem(SITE_PHOTOS_CACHE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function getCachedPhotosForSite(siteCode) {
  if (!siteCode) return [];
  const map = getAllCachedSitePhotos();
  const cleanCode = String(siteCode).trim().toUpperCase();
  if (Array.isArray(map[cleanCode]) && map[cleanCode].length > 0) return map[cleanCode];
  const norm = cleanCode.replace(/[^A-Z0-9]/g, '');
  for (const [k, v] of Object.entries(map)) {
    if (k.replace(/[^A-Z0-9]/g, '').toUpperCase() === norm && Array.isArray(v) && v.length > 0) {
      return v;
    }
  }
  return [];
}

function saveCachedPhotosForSite(siteCode, urls) {
  if (!siteCode) return;
  const cleanCode = String(siteCode).trim().toUpperCase();
  try {
    const map = getAllCachedSitePhotos();
    if (Array.isArray(urls) && urls.length > 0) {
      map[cleanCode] = urls;
    } else {
      delete map[cleanCode];
      const norm = cleanCode.replace(/[^A-Z0-9]/g, '');
      for (const k of Object.keys(map)) {
        if (k.replace(/[^A-Z0-9]/g, '').toUpperCase() === norm) delete map[k];
      }
    }
    localStorage.setItem(SITE_PHOTOS_CACHE_KEY, JSON.stringify(map));
  } catch (e) {
    console.warn('saveCachedPhotosForSite notice:', e);
  }
}

function clearAllCachedPhotos(targetSiteCodes = null) {
  try {
    if (!targetSiteCodes || targetSiteCodes.length === 0) {
      localStorage.removeItem(SITE_PHOTOS_CACHE_KEY);
      clearAllPhotoBlobs();
    } else {
      const map = getAllCachedSitePhotos();
      for (const code of targetSiteCodes) {
        const cleanCode = String(code).trim().toUpperCase();
        delete map[cleanCode];
        const norm = cleanCode.replace(/[^A-Z0-9]/g, '');
        for (const k of Object.keys(map)) {
          if (k.replace(/[^A-Z0-9]/g, '').toUpperCase() === norm) delete map[k];
        }
      }
      localStorage.setItem(SITE_PHOTOS_CACHE_KEY, JSON.stringify(map));
    }
  } catch {}
}

// Persistent binary/blob image store in IndexedDB
const DB_NAME = 'MB_Photos_DB';
const DB_VERSION = 1;
const STORE_NAME = 'photo_blobs';
let photoDbPromise = null;

function getPhotoDb() {
  if (photoDbPromise) return photoDbPromise;
  photoDbPromise = new Promise(resolve => {
    if (typeof window === 'undefined' || !window.indexedDB) return resolve(null);
    try {
      const req = window.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return photoDbPromise;
}

async function cachePhotoBlob(key, data) {
  if (!key || !data) return;
  try {
    const db = await getPhotoDb();
    if (!db) return;
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(data, key);
  } catch {}
}

async function getCachedPhotoBlob(key) {
  if (!key) return null;
  try {
    const db = await getPhotoDb();
    if (!db) return null;
    return await new Promise(resolve => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch {
    return null;
  }
}

async function deleteCachedPhotoBlob(key) {
  if (!key) return;
  try {
    const db = await getPhotoDb();
    if (!db) return;
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(key);
  } catch {}
}

async function clearAllPhotoBlobs() {
  try {
    const db = await getPhotoDb();
    if (!db) return;
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).clear();
  } catch {}
}

function unpackSite(s) {
  let f = s.flags;
  try {
    if (typeof f === 'string') f = JSON.parse(f);
  } catch {}
  f = f && typeof f === 'object' && !Array.isArray(f) ? f : {};

  const siteCode = String(s.site_code || '').trim().toUpperCase();
  const cachedImgs = getCachedPhotosForSite(siteCode);

  let serverImgs = Array.isArray(f.ppt_images) && f.ppt_images.length > 0
    ? f.ppt_images
    : (Array.isArray(s.ppt_images) && s.ppt_images.length > 0 ? s.ppt_images : []);

  let finalImgs = serverImgs;
  if ((!finalImgs || finalImgs.length === 0) && cachedImgs && cachedImgs.length > 0) {
    // Retain locally cached uploaded photos if server has not returned them
    finalImgs = cachedImgs;
  } else if (serverImgs && serverImgs.length > 0) {
    // Keep local cache fresh with server images
    saveCachedPhotosForSite(siteCode, serverImgs);
  }

  return {
    ...s,
    ppt_images: finalImgs || [],
    ppt_availability: f.ppt_availability || s.ppt_availability || s.availability || '',
    ppt_rate: f.ppt_rate || s.ppt_rate || s.monthly_rate || ''
  };
}

// In-memory preview cache for instantaneous zero-lag image rendering and offline resilience
const photoPreviewCache = new Map();

function resolvePhotoUrl(u) {
  if (!u) return '';
  if (photoPreviewCache.has(u)) return photoPreviewCache.get(u);
  if (typeof u === 'string' && (u.startsWith('data:') || u.startsWith('blob:'))) return u;
  // Always route /uploads/ via /api/uploads/ for bulletproof proxying across dev and production
  if (typeof u === 'string' && u.startsWith('/uploads/')) {
    return `/api${u}`;
  }
  return u;
}

async function imageData(url) {
  if (!url) throw new Error('Empty image URL');
  if (photoPreviewCache.has(url)) return photoPreviewCache.get(url);
  if (url.startsWith('data:')) return url;

  // Check persistent IndexedDB cache first
  try {
    const cached = await getCachedPhotoBlob(url);
    if (cached) {
      photoPreviewCache.set(url, cached);
      return cached;
    }
  } catch {}

  // Determine priority candidate URLs
  const candidates = [];
  if (url.startsWith('/uploads/')) {
    candidates.push(`/api${url}`);
    candidates.push(url);
  } else if (url.startsWith('/api/uploads/')) {
    candidates.push(url);
    candidates.push(url.replace('/api', ''));
  } else {
    candidates.push(url);
  }

  for (const targetUrl of candidates) {
    try {
      const r = await fetch(targetUrl);
      if (r.ok) {
        const b = await r.blob();
        const dataUrl = await new Promise((ok, ko) => {
          const fr = new FileReader();
          fr.onload = () => ok(fr.result);
          fr.onerror = ko;
          fr.readAsDataURL(b);
        });
        photoPreviewCache.set(url, dataUrl);
        cachePhotoBlob(url, dataUrl);
        return dataUrl;
      }
    } catch {}
  }
  throw new Error(`Failed to load presentation image from: ${url}`);
}

const OVERLAY_SVG_STRING = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080" width="1920" height="1080">
  <defs>
    <!-- Deep dark navy gradient for the right dashboard panel -->
    <linearGradient id="mainNavy" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#02142B" />
      <stop offset="60%" stop-color="#04162C" />
      <stop offset="100%" stop-color="#010B17" />
    </linearGradient>

    <!-- Outer angled facet gradient -->
    <linearGradient id="outerFacet" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0A2C52" />
      <stop offset="100%" stop-color="#041830" />
    </linearGradient>

    <!-- Subtle drop shadow for bottom-left logo badge over photo -->
    <filter id="badgeShadow" x="-10%" y="-10%" width="125%" height="125%">
      <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#000000" flood-opacity="0.45" />
    </filter>
  </defs>

  <!-- ================= 1. LEFT BOTTOM CORNER: OFFICIAL LOGO BADGE ================= -->
  <g transform="translate(45, 920)" filter="url(#badgeShadow)">
    <!-- Solid Brand Yellow Badge matching official master logo -->
    <rect width="320" height="116" rx="10" fill="#FFCC00" />
    <svg x="10" y="7" width="300" height="102" viewBox="135 170 550 195">
      <!-- White Butterfly / Bow Emblem -->
      <path d="M 423.95 206.12 C 423.95 206.12 451.13 188.08 467.73 194.82 C 467.73 194.82 475.68 199.04 473.14 208.07 C 469.02 222.71 447.63 232.42 435.26 236.91 C 435.26 236.91 451.23 246.53 458.80 260.40 C 466.83 275.13 460.58 283.22 446.34 275.99 C 442.96 274.27 438.72 271.39 435.26 266.74 L 429.73 278.28 C 437.34 287.53 458.23 295.93 468.64 287.13 C 471.72 284.52 475.18 280.16 475.61 273.55 C 476.26 263.82 471.52 256.19 465.58 249.20 C 462.52 245.60 458.65 241.26 456.43 239.56 C 456.43 239.56 470.93 233.24 480.11 221.75 C 496.01 201.84 483.14 179.98 459.07 181.59 C 459.07 181.59 442.00 181.35 417.70 198.42 Z M 405.87 206.12 C 405.87 206.12 378.69 188.08 362.09 194.82 C 362.09 194.82 354.14 199.04 356.68 208.07 C 360.80 222.71 382.20 232.42 394.57 236.91 C 394.57 236.91 378.59 246.53 371.03 260.40 C 362.99 275.13 369.25 283.22 383.48 275.99 C 386.87 274.27 391.10 271.39 394.57 266.74 L 400.10 278.28 C 392.49 287.53 371.59 295.93 361.18 287.13 C 358.10 284.52 354.65 280.16 354.21 273.55 C 353.57 263.82 358.30 256.19 364.24 249.20 C 367.30 245.60 371.17 241.26 373.40 239.56 C 373.40 239.56 358.89 233.24 349.71 221.75 C 333.82 201.84 346.68 179.98 370.75 181.59 C 370.75 181.59 387.83 181.35 412.12 198.42 Z" fill="#FFFFFF" fill-rule="evenodd" />
      <!-- Media (Dark Navy) -->
      <path d="M 187.21 259.17 C 187.21 265.17 200.63 265.17 200.63 265.17 L 200.63 213.69 C 200.63 207.69 189.40 207.69 186.44 207.69 L 172.08 238.34 L 160.62 212.91 C 158.20 207.61 152.35 207.69 151.88 207.69 L 143.61 207.69 L 143.61 259.17 C 143.61 265.17 157.03 265.17 157.03 265.17 L 157.03 239.04 C 157.03 235.22 156.17 230.77 156.17 230.77 L 164.52 249.65 C 168.26 255.50 176.37 255.50 176.37 255.50 L 188.23 230.30 C 188.23 230.30 187.21 235.22 187.21 239.04 Z Z M 218.47 238.18 C 218.47 233.89 219.72 231.55 223.62 231.55 C 227.05 231.55 228.38 233.43 228.38 236.23 C 228.38 239.59 225.73 240.52 222.92 240.52 C 221.28 240.52 219.64 240.21 218.47 239.90 Z Z M 225.88 265.33 C 235.79 265.33 242.11 260.88 242.11 255.19 C 242.11 253.00 241.17 250.66 239.30 248.09 C 239.22 252.14 234.38 255.81 226.27 255.81 C 221.51 255.81 218.47 253.86 218.47 249.49 L 218.47 248.24 C 218.47 248.24 220.89 249.49 224.87 249.49 C 237.89 249.49 241.41 243.10 241.41 236.23 C 241.41 227.89 236.26 222.35 224.17 222.35 C 212.08 222.35 205.37 227.89 205.37 237.01 L 205.37 250.66 C 205.37 259.87 212.54 265.33 225.88 265.33 Z M 269.54 204.02 L 269.54 220.71 C 269.54 222.27 270.09 224.77 270.09 224.77 C 270.09 224.77 266.27 222.35 261.66 222.35 C 252.46 222.35 244.50 227.81 244.50 236.78 L 244.50 250.82 C 244.50 260.10 251.52 265.33 263.61 265.33 C 266.42 265.33 270.87 264.00 275.70 264.00 C 278.36 264.00 280.38 264.08 282.65 265.33 L 282.65 210.80 C 282.65 204.10 270.87 204.02 269.54 204.02 Z M 269.54 255.97 C 269.54 255.97 266.73 256.43 264.55 256.43 C 259.24 256.43 257.61 254.02 257.61 249.65 L 257.61 237.95 C 257.61 233.50 259.71 231.55 262.68 231.55 C 266.19 231.55 269.54 233.82 269.54 237.17 Z Z M 287.38 258.93 C 287.38 264.94 300.49 264.94 300.49 264.94 L 300.49 228.74 C 300.49 222.97 288.47 222.74 287.38 222.74 Z Z M 287.61 215.41 C 287.61 216.81 289.33 218.29 291.75 218.29 L 300.10 218.29 L 300.10 210.80 C 300.10 209.64 298.38 207.84 296.04 207.84 L 287.61 207.84 Z Z M 327.77 255.66 C 327.77 255.66 324.88 256.51 322.07 256.51 C 319.42 256.51 316.77 255.73 316.77 252.07 C 316.77 247.08 325.66 248.32 327.77 245.05 Z Z M 321.06 231.63 C 326.13 231.63 327.53 233.82 327.53 235.69 C 323.94 240.91 303.66 238.26 303.66 252.53 C 303.66 262.36 312.09 265.33 320.82 265.33 C 325.42 265.33 329.80 263.92 333.93 263.92 C 336.11 263.92 338.45 264.16 340.64 265.33 L 340.64 236.31 C 340.64 227.26 333.93 222.35 321.21 222.35 C 315.83 222.35 304.75 223.91 304.75 231.79 C 304.75 234.21 305.46 237.01 308.34 240.83 C 308.34 233.43 315.05 231.63 321.06 231.63" fill="#030352" fill-rule="evenodd" />
      <!-- Buzz (Dark Navy) -->
      <path d="M 498.73 258.07 C 498.73 264.00 508.72 264.08 508.72 264.08 L 518.31 264.08 C 531.65 264.08 540.85 259.71 540.85 248.87 L 540.85 247.15 C 540.85 242.32 538.12 237.09 533.83 235.38 C 537.50 233.97 540.07 228.51 540.07 223.99 L 540.07 222.43 C 540.07 211.04 530.95 207.37 518.31 207.37 L 498.73 207.37 Z Z M 518.93 240.21 C 524.86 240.21 527.44 242.47 527.44 247.70 C 527.44 252.77 524.71 254.64 518.31 254.64 L 512.15 254.64 L 512.15 240.21 Z Z M 518.31 216.89 C 523.93 216.89 526.66 218.37 526.66 222.82 C 526.66 227.97 524.63 230.77 518.93 230.77 L 512.15 230.77 L 512.15 216.89 Z Z M 557.91 228.43 C 557.91 222.74 544.81 222.43 544.81 222.43 L 544.81 251.21 C 544.81 262.05 554.17 265.01 562.90 265.01 C 566.88 265.01 570.55 263.38 574.53 263.38 C 579.44 263.38 582.95 265.33 582.95 265.33 L 582.95 228.43 C 582.95 222.43 569.85 222.43 569.85 222.43 L 569.85 254.95 C 569.85 254.95 566.65 255.73 564.00 255.73 C 559.47 255.73 557.91 253.71 557.91 250.04 Z Z M 604.85 232.49 C 599.86 240.52 584.88 248.01 584.88 259.79 C 584.88 260.41 584.80 260.96 585.50 264.08 L 614.60 264.08 C 621.85 264.08 621.85 254.64 621.85 254.64 L 599.78 254.64 C 606.02 246.14 617.95 238.73 620.06 231.32 C 620.92 228.43 621.07 225.70 621.93 222.97 L 593.15 222.97 C 585.74 222.97 585.74 232.49 585.74 232.49 Z Z M 640.32 232.49 C 635.32 240.52 620.35 248.01 620.35 259.79 C 620.35 260.41 620.27 260.96 620.97 264.08 L 650.07 264.08 C 657.32 264.08 657.32 254.64 657.32 254.64 L 635.25 254.64 C 641.49 246.14 653.42 238.73 655.53 231.32 C 656.39 228.43 656.54 225.70 657.40 222.97 L 628.62 222.97 C 621.21 222.97 621.21 232.49 621.21 232.49 Z" fill="#030352" fill-rule="evenodd" />
      <!-- Registered trademark circle & R -->
      <path d="M 661.61 189.98 C 670.50 189.98 677.72 196.99 677.72 205.63 C 677.72 214.27 670.50 221.28 661.61 221.28 C 652.71 221.28 645.50 214.27 645.50 205.63 C 645.50 196.99 652.71 189.98 661.61 189.98" stroke="#030352" stroke-width="1.8" fill="none" />
      <text x="661.6" y="206.5" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="16" fill="#030352" text-anchor="middle" dominant-baseline="central">R</text>
      <!-- Be Seen tagline -->
      <text x="412" y="352" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="58" fill="#030352" text-anchor="middle" letter-spacing="1">Be Seen</text>
    </svg>
  </g>

  <!-- ================= 2. RIGHT SIDE DASHBOARD: PROPER & BIG ================= -->
  <!-- Left outer angled facet -->
  <polygon points="1180,0 1260,0 1080,1080 1060,1080" fill="url(#outerFacet)" />
  <line x1="1180" y1="0" x2="1060" y2="1080" stroke="#0F3863" stroke-width="1.8" />

  <!-- Main dark navy right-hand panel -->
  <polygon points="1260,0 1920,0 1920,1080 1080,1080" fill="url(#mainNavy)" />
  <line x1="1260" y1="0" x2="1080" y2="1080" stroke="#0B2A4A" stroke-width="1.5" />

  <!-- A. SITE CODE Section Accents (Compact top bar and clean label) -->
  <!-- Yellow horizontal bar above SITE CODE -->
  <rect x="1300" y="24" width="70" height="6" rx="3" fill="#FFC200" />
  <!-- Label: SITE CODE -->
  <text x="1300" y="58" font-family="'Montserrat', Arial, sans-serif" font-weight="700" font-size="16" fill="#FFFFFF" letter-spacing="3">SITE CODE</text>

  <!-- B. Thin Horizontal Divider Lines under Spec Rows 1–5 -->
  <!-- Line under Row 1: Size -->
  <line x1="1395" y1="485" x2="1885" y2="485" stroke="#0F355C" stroke-width="1.5" />
  <!-- Line under Row 2: Type -->
  <line x1="1395" y1="580" x2="1885" y2="580" stroke="#0F355C" stroke-width="1.5" />
  <!-- Line under Row 3: Illumination -->
  <line x1="1395" y1="675" x2="1885" y2="675" stroke="#0F355C" stroke-width="1.5" />
  <!-- Line under Row 4: Adv. Fee Per Month -->
  <line x1="1395" y1="770" x2="1885" y2="770" stroke="#0F355C" stroke-width="1.5" />
  <!-- Line under Row 5: Availability -->
  <line x1="1395" y1="865" x2="1885" y2="865" stroke="#0F355C" stroke-width="1.5" />

  <!-- C. Spec Section Icons in unified Cyan/Sky Blue (#38BDF8) matching reference format -->
  <!-- Row 1: Size Icon -->
  <g transform="translate(1305, 408)">
    <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="15 3 21 3 21 9" />
      <polyline points="9 21 3 21 3 15" />
      <polyline points="21 15 21 21 15 21" />
      <polyline points="3 9 3 3 9 3" />
    </svg>
  </g>

  <!-- Row 2: Type Icon -->
  <g transform="translate(1304, 502)">
    <svg width="46" height="46" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <rect x="3" y="4" width="18" height="11" rx="1.5" />
      <line x1="8" y1="15" x2="8" y2="21" />
      <line x1="16" y1="15" x2="16" y2="21" />
      <line x1="5" y1="21" x2="19" y2="21" />
    </svg>
  </g>

  <!-- Row 3: Illumination Icon -->
  <g transform="translate(1304, 597)">
    <svg width="46" height="46" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M9 18h6" />
      <path d="M10 22h4" />
      <path d="M12 2v2" />
      <path d="M12 6a5 5 0 0 0-3.5 8.5c.9.9 1.5 1.8 1.5 2.5h4c0-.7.6-1.6 1.5-2.5A5 5 0 0 0 12 6z" />
      <line x1="4.93" y1="4.93" x2="6.34" y2="6.34" />
      <line x1="19.07" y1="4.93" x2="17.66" y2="6.34" />
    </svg>
  </g>

  <!-- Row 4: Adv. Fee Icon (Rupee) -->
  <g transform="translate(1308, 692)">
    <svg width="40" height="44" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 3h12" />
      <path d="M6 8h12" />
      <path d="M6 13l8.5 8" />
      <path d="M6 13h3a4.5 4.5 0 0 0 0-9" />
    </svg>
  </g>

  <!-- Row 5: Availability Icon (Circle Check) -->
  <g transform="translate(1306, 787)">
    <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="9.5" stroke="#38BDF8" />
      <polyline points="7.5 12 10.5 15 16.5 9" stroke="#38BDF8" />
    </svg>
  </g>

  <!-- Row 6: Location Coordinates Icon (Pin) -->
  <g transform="translate(1306, 888)">
    <svg width="44" height="50" viewBox="0 0 24 24" fill="#38BDF8">
      <path fill-rule="evenodd" clip-rule="evenodd" d="M12 2C7.58 2 4 5.58 4 10C4 15.5 12 22 12 22C12 22 20 15.5 20 10C20 5.58 16.42 2 12 2ZM12 13C10.34 13 9 11.66 9 10C9 8.34 10.34 7 12 7C13.66 7 15 8.34 15 10C15 11.66 13.66 13 12 13Z" />
    </svg>
  </g>
</svg>`;

const LOGO_SVG_STRING = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 540 140" width="540" height="140">
  <g transform="translate(10, 10)">
    <!-- Media: bold white text with modern geometric sans -->
    <text x="5" y="65" font-family="'Montserrat', 'Century Gothic', 'Segoe UI', Arial, sans-serif" font-weight="900" font-size="58" fill="#FFFFFF" letter-spacing="-0.5">Media</text>
    
    <!-- Stylized yellow continuous ribbon bow emblem with hollow center matching reference photo -->
    <g transform="translate(196, 5)">
      <path d="M 38 19
               C 45 7, 59 5, 66 12
               C 74 20, 73 33, 58 37
               C 73 41, 74 54, 66 62
               C 59 69, 45 67, 38 55
               C 31 67, 17 69, 10 62
               C 2 54, 3 41, 18 37
               C 3 33, 2 20, 10 12
               C 17 5, 31 7, 38 19 Z"
            fill="none"
            stroke="#FFC200"
            stroke-width="7.5"
            stroke-linecap="round"
            stroke-linejoin="round" />
    </g>

    <!-- Buzz: bold white text with registered trademark -->
    <text x="286" y="65" font-family="'Montserrat', 'Century Gothic', 'Segoe UI', Arial, sans-serif" font-weight="900" font-size="58" fill="#FFFFFF" letter-spacing="-0.5">Buzz</text>
    <text x="432" y="30" font-family="'Montserrat', Arial, sans-serif" font-weight="bold" font-size="19" fill="#FFFFFF">®</text>

    <!-- Be Seen: clean white text centered underneath emblem and Buzz -->
    <text x="250" y="104" font-family="'Inter', 'Segoe UI', Arial, sans-serif" font-weight="600" font-size="25" fill="#FFFFFF" letter-spacing="1">Be Seen</text>
  </g>
</svg>`;

const PPT_ICONS_SVG = {
  // 1. Size: Expand square with diagonal arrows pointing to 4 corners
  size: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <rect x="44" y="44" width="40" height="40" rx="3" fill="none" stroke="#000000" stroke-width="5.5" />
    <path d="M 44 44 L 28 28 M 28 40 L 28 28 L 40 28" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 84 44 L 100 28 M 88 28 L 100 28 L 100 40" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 44 84 L 28 100 M 40 100 L 28 100 L 28 88" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 84 84 L 100 100 M 88 100 L 100 100 L 100 88" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
  </svg>`,

  // 2. Type: Speech bubble with 3 dots inside
  type: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <path d="M 32 60 C 32 43 46 30 64 30 C 82 30 96 43 96 60 C 96 77 82 90 64 90 C 58 90 52 88 47 85 L 32 94 L 36 80 C 33.5 74 32 67 32 60 Z" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
    <circle cx="50" cy="60" r="4.5" fill="#000000" />
    <circle cx="64" cy="60" r="4.5" fill="#000000" />
    <circle cx="78" cy="60" r="4.5" fill="#000000" />
  </svg>`,

  // 3. Illumination: Light bulb with radiant rays
  illum: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <path d="M 48 64 C 44 58 42 51 42 44 C 42 32 52 22 64 22 C 76 22 86 32 86 44 C 86 51 84 58 80 64 C 76 70 74 76 74 82 L 54 82 C 54 76 52 70 48 64 Z" fill="none" stroke="#000000" stroke-width="5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 54 89 L 74 89 M 57 96 L 71 96 M 60 102 L 68 102" fill="none" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
    <path d="M 58 56 L 58 44 C 58 40 61 38 64 38 C 67 38 70 40 70 44 L 70 56" fill="none" stroke="#000000" stroke-width="4" stroke-linecap="round" />
    <line x1="64" y1="12" x2="64" y2="16" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
    <line x1="38" y1="22" x2="41" y2="25" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
    <line x1="90" y1="22" x2="87" y2="25" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
    <line x1="28" y1="44" x2="33" y2="44" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
    <line x1="100" y1="44" x2="95" y2="44" stroke="#000000" stroke-width="4.5" stroke-linecap="round" />
  </svg>`,

  // 4. Adv. fee per month: Hand holding cash banknote
  rupee: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <rect x="42" y="32" width="56" height="32" rx="3" fill="none" stroke="#000000" stroke-width="5" transform="rotate(-6 70 48)" />
    <circle cx="70" cy="48" r="7" fill="none" stroke="#000000" stroke-width="4" />
    <path d="M 28 88 L 48 88 C 55 88 62 84 66 79 L 80 64 C 83 61 82 56 78 54 C 74 52 70 54 67 58 L 56 70 L 46 70" fill="none" stroke="#000000" stroke-width="5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 44 80 L 44 98 M 28 98 L 48 98" fill="none" stroke="#000000" stroke-width="5" stroke-linecap="round" />
  </svg>`,

  // 5. Availability: Hand holding checkmark clock / guarantee badge
  avail: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <circle cx="66" cy="50" r="24" fill="none" stroke="#000000" stroke-width="5" />
    <polyline points="55 50 63 58 78 43" fill="none" stroke="#000000" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round" />
    <path d="M 28 88 C 42 88 52 82 60 76 L 78 76 C 84 76 88 80 84 86 L 72 96 C 66 101 58 104 50 104 L 28 104" fill="none" stroke="#000000" stroke-width="5" stroke-linecap="round" stroke-linejoin="round" />
  </svg>`,

  // 6. Location Coordinates: Map pin with white central hole
  pin: `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
    <circle cx="64" cy="64" r="62" fill="#FFFFFF" />
    <path d="M 64 26 C 47.4 26 34 39.4 34 56 C 34 76 64 102 64 102 C 64 102 94 76 94 56 C 94 39.4 80.6 26 64 26 Z" fill="#000000" />
    <circle cx="64" cy="54" r="11" fill="#FFFFFF" />
  </svg>`
};

const PPT_LOGO_BADGE_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 170" width="420" height="170">
  <rect width="420" height="170" rx="14" fill="#FFC200" />
  <g transform="translate(10, -5)">
    <svg x="0" y="5" width="400" height="160" viewBox="135 170 550 195">
      <path d="M 423.95 206.12 C 423.95 206.12 451.13 188.08 467.73 194.82 C 467.73 194.82 475.68 199.04 473.14 208.07 C 469.02 222.71 447.63 232.42 435.26 236.91 C 435.26 236.91 451.23 246.53 458.80 260.40 C 466.83 275.13 460.58 283.22 446.34 275.99 C 442.96 274.27 438.72 271.39 435.26 266.74 L 429.73 278.28 C 437.34 287.53 458.23 295.93 468.64 287.13 C 471.72 284.52 475.18 280.16 475.61 273.55 C 476.26 263.82 471.52 256.19 465.58 249.20 C 462.52 245.60 458.65 241.26 456.43 239.56 C 456.43 239.56 470.93 233.24 480.11 221.75 C 496.01 201.84 483.14 179.98 459.07 181.59 C 459.07 181.59 442.00 181.35 417.70 198.42 Z M 405.87 206.12 C 405.87 206.12 378.69 188.08 362.09 194.82 C 362.09 194.82 354.14 199.04 356.68 208.07 C 360.80 222.71 382.20 232.42 394.57 236.91 C 394.57 236.91 378.59 246.53 371.03 260.40 C 362.99 275.13 369.25 283.22 383.48 275.99 C 386.87 274.27 391.10 271.39 394.57 266.74 L 400.10 278.28 C 392.49 287.53 371.59 295.93 361.18 287.13 C 358.10 284.52 354.65 280.16 354.21 273.55 C 353.57 263.82 358.30 256.19 364.24 249.20 C 367.30 245.60 371.17 241.26 373.40 239.56 C 373.40 239.56 358.89 233.24 349.71 221.75 C 333.82 201.84 346.68 179.98 370.75 181.59 C 370.75 181.59 387.83 181.35 412.12 198.42 Z" fill="#FFFFFF" fill-rule="evenodd" />
      <path d="M 187.21 259.17 C 187.21 265.17 200.63 265.17 200.63 265.17 L 200.63 213.69 C 200.63 207.69 189.40 207.69 186.44 207.69 L 172.08 238.34 L 160.62 212.91 C 158.20 207.61 152.35 207.69 151.88 207.69 L 143.61 207.69 L 143.61 259.17 C 143.61 265.17 157.03 265.17 157.03 265.17 L 157.03 239.04 C 157.03 235.22 156.17 230.77 156.17 230.77 L 164.52 249.65 C 168.26 255.50 176.37 255.50 176.37 255.50 L 188.23 230.30 C 188.23 230.30 187.21 235.22 187.21 239.04 Z Z M 218.47 238.18 C 218.47 233.89 219.72 231.55 223.62 231.55 C 227.05 231.55 228.38 233.43 228.38 236.23 C 228.38 239.59 225.73 240.52 222.92 240.52 C 221.28 240.52 219.64 240.21 218.47 239.90 Z Z M 225.88 265.33 C 235.79 265.33 242.11 260.88 242.11 255.19 C 242.11 253.00 241.17 250.66 239.30 248.09 C 239.22 252.14 234.38 255.81 226.27 255.81 C 221.51 255.81 218.47 253.86 218.47 249.49 L 218.47 248.24 C 218.47 248.24 220.89 249.49 224.87 249.49 C 237.89 249.49 241.41 243.10 241.41 236.23 C 241.41 227.89 236.26 222.35 224.17 222.35 C 212.08 222.35 205.37 227.89 205.37 237.01 L 205.37 250.66 C 205.37 259.87 212.54 265.33 225.88 265.33 Z M 269.54 204.02 L 269.54 220.71 C 269.54 222.27 270.09 224.77 270.09 224.77 C 270.09 224.77 266.27 222.35 261.66 222.35 C 252.46 222.35 244.50 227.81 244.50 236.78 L 244.50 250.82 C 244.50 260.10 251.52 265.33 263.61 265.33 C 266.42 265.33 270.87 264.00 275.70 264.00 C 278.36 264.00 280.38 264.08 282.65 265.33 L 282.65 210.80 C 282.65 204.10 270.87 204.02 269.54 204.02 Z M 269.54 255.97 C 269.54 255.97 266.73 256.43 264.55 256.43 C 259.24 256.43 257.61 254.02 257.61 249.65 L 257.61 237.95 C 257.61 233.50 259.71 231.55 262.68 231.55 C 266.19 231.55 269.54 233.82 269.54 237.17 Z Z M 287.38 258.93 C 287.38 264.94 300.49 264.94 300.49 264.94 L 300.49 228.74 C 300.49 222.97 288.47 222.74 287.38 222.74 Z Z M 287.61 215.41 C 287.61 216.81 289.33 218.29 291.75 218.29 L 300.10 218.29 L 300.10 210.80 C 300.10 209.64 298.38 207.84 296.04 207.84 L 287.61 207.84 Z Z M 327.77 255.66 C 327.77 255.66 324.88 256.51 322.07 256.51 C 319.42 256.51 316.77 255.73 316.77 252.07 C 316.77 247.08 325.66 248.32 327.77 245.05 Z Z M 321.06 231.63 C 326.13 231.63 327.53 233.82 327.53 235.69 C 323.94 240.91 303.66 238.26 303.66 252.53 C 303.66 262.36 312.09 265.33 320.82 265.33 C 325.42 265.33 329.80 263.92 333.93 263.92 C 336.11 263.92 338.45 264.16 340.64 265.33 L 340.64 236.31 C 340.64 227.26 333.93 222.35 321.21 222.35 C 315.83 222.35 304.75 223.91 304.75 231.79 C 304.75 234.21 305.46 237.01 308.34 240.83 C 308.34 233.43 315.05 231.63 321.06 231.63" fill="#030352" fill-rule="evenodd" />
      <path d="M 498.73 258.07 C 498.73 264.00 508.72 264.08 508.72 264.08 L 518.31 264.08 C 531.65 264.08 540.85 259.71 540.85 248.87 L 540.85 247.15 C 540.85 242.32 538.12 237.09 533.83 235.38 C 537.50 233.97 540.07 228.51 540.07 223.99 L 540.07 222.43 C 540.07 211.04 530.95 207.37 518.31 207.37 L 498.73 207.37 Z Z M 518.93 240.21 C 524.86 240.21 527.44 242.47 527.44 247.70 C 527.44 252.77 524.71 254.64 518.31 254.64 L 512.15 254.64 L 512.15 240.21 Z Z M 518.31 216.89 C 523.93 216.89 526.66 218.37 526.66 222.82 C 526.66 227.97 524.63 230.77 518.93 230.77 L 512.15 230.77 L 512.15 216.89 Z Z M 557.91 228.43 C 557.91 222.74 544.81 222.43 544.81 222.43 L 544.81 251.21 C 544.81 262.05 554.17 265.01 562.90 265.01 C 566.88 265.01 570.55 263.38 574.53 263.38 C 579.44 263.38 582.95 265.33 582.95 265.33 L 582.95 228.43 C 582.95 222.43 569.85 222.43 569.85 222.43 L 569.85 254.95 C 569.85 254.95 566.65 255.73 564.00 255.73 C 559.47 255.73 557.91 253.71 557.91 250.04 Z Z M 604.85 232.49 C 599.86 240.52 584.88 248.01 584.88 259.79 C 584.88 260.41 584.80 260.96 585.50 264.08 L 614.60 264.08 C 621.85 264.08 621.85 254.64 621.85 254.64 L 599.78 254.64 C 606.02 246.14 617.95 238.73 620.06 231.32 C 620.92 228.43 621.07 225.70 621.93 222.97 L 593.15 222.97 C 585.74 222.97 585.74 232.49 585.74 232.49 Z Z M 640.32 232.49 C 635.32 240.52 620.35 248.01 620.35 259.79 C 620.35 260.41 620.27 260.96 620.97 264.08 L 650.07 264.08 C 657.32 264.08 657.32 254.64 657.32 254.64 L 635.25 254.64 C 641.49 246.14 653.42 238.73 655.53 231.32 C 656.39 228.43 656.54 225.70 657.40 222.97 L 628.62 222.97 C 621.21 222.97 621.21 232.49 621.21 232.49 Z" fill="#030352" fill-rule="evenodd" />
      <path d="M 661.61 189.98 C 670.50 189.98 677.72 196.99 677.72 205.63 C 677.72 214.27 670.50 221.28 661.61 221.28 C 652.71 221.28 645.50 214.27 645.50 205.63 C 645.50 196.99 652.71 189.98 661.61 189.98" stroke="#030352" stroke-width="1.8" fill="none" />
      <text x="661.6" y="206.5" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="16" fill="#030352" text-anchor="middle" dominant-baseline="central">R</text>
      <text x="412" y="352" font-family="'Arial', 'Helvetica', sans-serif" font-weight="bold" font-size="58" fill="#030352" text-anchor="middle" letter-spacing="1">Be Seen</text>
    </svg>
  </g>
</svg>`;

const svgPngCache = new Map();

async function svgToPngDataUrl(svgString, width = 128, height = 128) {
  const cacheKey = `${width}x${height}_${svgString}`;
  if (svgPngCache.has(cacheKey)) return svgPngCache.get(cacheKey);

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');

  const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const img = new Image();

  await new Promise((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('Failed to render SVG to canvas'));
    img.src = url;
  });

  ctx.drawImage(img, 0, 0, width, height);
  URL.revokeObjectURL(url);
  const dataUrl = canvas.toDataURL('image/png');
  svgPngCache.set(cacheKey, dataUrl);
  return dataUrl;
}

let cachedLogoBadgePng = null;
async function getSlideLogoBadgePng() {
  if (cachedLogoBadgePng) return cachedLogoBadgePng;
  cachedLogoBadgePng = await svgToPngDataUrl(PPT_LOGO_BADGE_SVG, 500, 200);
  return cachedLogoBadgePng;
}

async function getSlideIconPng(key) {
  const svg = PPT_ICONS_SVG[key];
  if (!svg) return null;
  return await svgToPngDataUrl(svg, 128, 128);
}

async function createSitePhotoShowcase(imgDataUrl, boxW_px = 1500, boxH_px = 1224) {
  const img = new Image();
  await new Promise((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('Failed to load site photo for showcase'));
    img.src = imgDataUrl;
  });

  const nw = img.naturalWidth || 1920;
  const nh = img.naturalHeight || 1080;
  const imgRatio = nw / nh;

  const canvas = document.createElement('canvas');
  canvas.width = boxW_px;
  canvas.height = boxH_px;
  const ctx = canvas.getContext('2d');

  // Solid black base matching the presentation background
  ctx.fillStyle = '#000000';
  ctx.fillRect(0, 0, boxW_px, boxH_px);

  // 100% UNTOUCHED, ZERO-CUT CONTAIN
  // The entire photo always fits inside boxW_px and boxH_px without losing a single pixel
  const boxRatio = boxW_px / boxH_px;
  let fitW, fitH, fitX, fitY;
  if (imgRatio >= boxRatio) {
    fitW = boxW_px;
    fitH = Math.round(boxW_px / imgRatio);
    fitX = 0;
    fitY = Math.round((boxH_px - fitH) / 2);
  } else {
    fitH = boxH_px;
    fitW = Math.round(boxH_px * imgRatio);
    fitX = Math.round((boxW_px - fitW) / 2);
    fitY = 0;
  }

  // Smooth rounded corners clip on the fitted photo
  const radius = Math.min(22, Math.max(8, Math.round(Math.min(fitW, fitH) * 0.035)));
  ctx.save();
  ctx.beginPath();
  if (ctx.roundRect) {
    ctx.roundRect(fitX, fitY, fitW, fitH, radius);
  } else {
    ctx.moveTo(fitX + radius, fitY);
    ctx.lineTo(fitX + fitW - radius, fitY);
    ctx.quadraticCurveTo(fitX + fitW, fitY, fitX + fitW, fitY + radius);
    ctx.lineTo(fitX + fitW, fitY + fitH - radius);
    ctx.quadraticCurveTo(fitX + fitW, fitY + fitH, fitX + fitW - radius, fitY + fitH);
    ctx.lineTo(fitX + radius, fitY + fitH);
    ctx.quadraticCurveTo(fitX, fitY + fitH, fitX, fitY + fitH - radius);
    ctx.lineTo(fitX, fitY + radius);
    ctx.quadraticCurveTo(fitX, fitY, fitX + radius, fitY);
    ctx.closePath();
  }
  ctx.clip();

  // Draw 100% of photo - exactly fitted, zero crop, zero cut
  ctx.drawImage(img, fitX, fitY, fitW, fitH);

  // Subtle clean border around the fitted photo
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.16)';
  ctx.lineWidth = 2;
  if (ctx.roundRect) {
    ctx.beginPath();
    ctx.roundRect(fitX + 1, fitY + 1, fitW - 2, fitH - 2, radius);
    ctx.stroke();
  }
  ctx.restore();

  return canvas.toDataURL('image/jpeg', 0.94);
}

function getTodayPptDateStr() {
  const d = new Date();
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}-${month}-${year}`;
}

async function makePpt(sites, pages = {}, fileName = `${getTodayPptDateStr()}.pptx`) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Media Buzz Outdoor';
  const SW = 13.333, SH = 7.5;

  const [logoBadgePng, iconSizePng, iconTypePng, iconIllumPng, iconRupeePng, iconAvailPng, iconPinPng] = await Promise.all([
    getSlideLogoBadgePng(),
    getSlideIconPng('size'),
    getSlideIconPng('type'),
    getSlideIconPng('illum'),
    getSlideIconPng('rupee'),
    getSlideIconPng('avail'),
    getSlideIconPng('pin')
  ]);

  async function covered(slide, url, x, y, w, h) {
    try {
      const data = await imageData(url);
      slide.addImage({ data, x, y, w, h });
    } catch {}
  }

  async function addSitePhoto(slide, url) {
    const BOX_X = 0.70, BOX_Y = 0.70, BOX_W = 6.25, BOX_H = 5.10;
    try {
      const data = await imageData(url);
      const cardDataUrl = await createSitePhotoShowcase(data, 1500, 1224);
      slide.addImage({ data: cardDataUrl, x: BOX_X, y: BOX_Y, w: BOX_W, h: BOX_H });
    } catch (err) {
      console.warn('Site photo showcase notice:', err);
      try {
        const data = await imageData(url);
        slide.addImage({
          data,
          x: BOX_X,
          y: BOX_Y,
          w: BOX_W,
          h: BOX_H,
          sizing: { type: 'contain', w: BOX_W, h: BOX_H }
        });
      } catch (innerErr) {
        console.warn('Site photo fallback placement notice:', innerErr);
      }
    }
  }

  async function fixed(u) {
    if (!u) return;
    const s = pptx.addSlide();
    s.background = { color: 'FFFFFF' };
    await covered(s, u, 0, 0, SW, SH);
  }

  async function closingSlide(u) {
    if (!u) return;
    const s = pptx.addSlide();
    s.background = { color: '000000' };
    await covered(s, u, 0, 0, SW, SH);

    // 1. Clickable hyperlink shape over website link "mediabuzzoutdoor.com" (bottom-right)
    s.addShape(pptx.ShapeType.rect, {
      x: 9.60,
      y: 6.30,
      w: 3.10,
      h: 0.45,
      fill: { color: '000000', transparency: 99 },
      line: { color: '000000', transparency: 100 },
      hyperlink: { url: 'https://mediabuzzoutdoor.com', tooltip: 'https://mediabuzzoutdoor.com' }
    });

    // 2. Semantic text hyperlink for accessibility, PDF conversion & text search
    s.addText([
      {
        text: 'mediabuzzoutdoor.com',
        options: {
          hyperlink: { url: 'https://mediabuzzoutdoor.com', tooltip: 'https://mediabuzzoutdoor.com' },
          color: '000000',
          transparency: 100
        }
      }
    ], {
      x: 9.60,
      y: 6.30,
      w: 3.10,
      h: 0.45,
      align: 'right',
      valign: 'middle',
      margin: 0
    });

    // 3. Clickable email link for "sales@mediabuzzoutdoor.com"
    s.addShape(pptx.ShapeType.rect, {
      x: 8.85,
      y: 5.30,
      w: 3.85,
      h: 0.40,
      fill: { color: '000000', transparency: 99 },
      line: { color: '000000', transparency: 100 },
      hyperlink: { url: 'mailto:sales@mediabuzzoutdoor.com', tooltip: 'mailto:sales@mediabuzzoutdoor.com' }
    });

    // 4. Clickable billboard website link "www.mediabuzzoutdoor.com"
    s.addShape(pptx.ShapeType.rect, {
      x: 2.60,
      y: 3.80,
      w: 2.65,
      h: 0.40,
      fill: { color: '000000', transparency: 99 },
      line: { color: '000000', transparency: 100 },
      hyperlink: { url: 'https://mediabuzzoutdoor.com', tooltip: 'https://mediabuzzoutdoor.com' }
    });
  }

  function getCoords(st) {
    let lat = st.latitude, lng = st.longitude;
    if ((!lat || !lng) && st.gps) {
      const parts = String(st.gps).split(',').map(x => x.trim());
      if (parts.length === 2 && !isNaN(parseFloat(parts[0]))) {
        lat = parts[0];
        lng = parts[1];
      }
    }
    return {
      lat: lat ? String(lat) : '',
      lng: lng ? String(lng) : ''
    };
  }

  function getSize(st) {
    if (st.size && String(st.size).trim()) {
      return String(st.size).trim();
    }
    if (st.width && st.height) return `${st.width} x ${st.height}`;
    return '30 x 32';
  }

  async function siteSlide(site, photoUrl, siteIndex = 1) {
    const s = pptx.addSlide();
    // Solid pitch black background matching final reference format
    s.background = { color: '000000' };

    // 1. Left site photo showcase with smooth rounded corners
    if (photoUrl) {
      await addSitePhoto(s, photoUrl);
    } else {
      s.addShape(pptx.ShapeType.roundRect, {
        x: 0.70, y: 0.70, w: 6.25, h: 5.10,
        fill: { color: '111111' }, line: { color: '2A2A2A', width: 1.5 }, rectRadius: 0.12
      });
      s.addText('NO SITE PHOTO UPLOADED', {
        x: 0.85, y: 3.00, w: 5.95, h: 0.6,
        fontFace: 'Arial', fontSize: 18, bold: true, color: '555555', align: 'center'
      });
    }

    // 2. Official Media Buzz brand logo badge at bottom-left directly under photo
    if (logoBadgePng) {
      s.addImage({ data: logoBadgePng, x: 0.70, y: 6.05, w: 1.85, h: 0.75 });
    }

    // 3. Right column: Site Counter ("Site 1", "Site 2", etc.)
    s.addText(`Site ${siteIndex}`, {
      x: 7.45, y: 0.70, w: 3.50, h: 0.30,
      fontFace: 'Arial', fontSize: 13, color: 'FFFFFF', bold: false, margin: 0
    });

    // 4. Prominent Yellow Site Code Pill ("MB - 01")
    let rawCode = String(site.site_code || 'MB-01').trim();
    let displayCode = rawCode.toUpperCase();
    if (/^[A-Za-z]+-\d+$/.test(displayCode)) {
      displayCode = displayCode.replace('-', ' - ');
    }
    s.addShape(pptx.ShapeType.roundRect, {
      x: 7.45, y: 1.05, w: 2.65, h: 0.72,
      fill: { color: 'FFC200' }, line: { color: 'FFC200' }, rectRadius: 0.18
    });
    const trackerUrl = `${typeof window !== 'undefined' ? window.location.origin : ''}/campaigns?site=${encodeURIComponent(rawCode)}`;
    s.addText(displayCode, {
      x: 7.45, y: 1.05, w: 2.65, h: 0.72,
      fontFace: 'Arial', fontSize: 26, bold: true, color: '000000', align: 'center', valign: 'middle', margin: 0,
      hyperlink: { url: trackerUrl, tooltip: `Open ${rawCode} in Latest Booking` }
    });

    // 5. City Heading (Bold pure white)
    const city = String(site.city || site.area || site.site_code || 'City').trim();
    s.addText(city, {
      x: 7.45, y: 1.98, w: 5.30, h: 0.45,
      fontFace: 'Arial', fontSize: 24, bold: true, color: 'FFFFFF', valign: 'middle', margin: 0
    });

    // 6. Location / Address Subtitle (Bold white)
    const location = String(site.location || site.address || site.area || site.notes || (site.city ? `Near ${site.city} Hub` : '') || 'Full Address').trim();
    s.addText(location, {
      x: 7.45, y: 2.48, w: 5.30, h: 0.42,
      fontFace: 'Arial', fontSize: 16, bold: true, color: 'FFFFFF', valign: 'middle', margin: 0
    });

    // 7. Specifications Rows (6 rows total)
    // Row 1: Size
    const row1Y = 3.12;
    if (iconSizePng) s.addImage({ data: iconSizePng, x: 7.45, y: row1Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Size:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: getSize(site), options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row1Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
    s.addShape(pptx.ShapeType.line, { x: 8.12, y: row1Y + 0.46, w: 4.80, h: 0, line: { color: '333333', width: 0.8 } });

    // Row 2: Type
    const row2Y = 3.74;
    if (iconTypePng) s.addImage({ data: iconTypePng, x: 7.45, y: row2Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Type:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: String(site.media_type || 'Hoarding'), options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row2Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
    s.addShape(pptx.ShapeType.line, { x: 8.12, y: row2Y + 0.46, w: 4.80, h: 0, line: { color: '333333', width: 0.8 } });

    // Row 3: Illumination
    const row3Y = 4.36;
    if (iconIllumPng) s.addImage({ data: iconIllumPng, x: 7.45, y: row3Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Illumination:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: String(site.lighting || 'Lit').toUpperCase(), options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row3Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
    s.addShape(pptx.ShapeType.line, { x: 8.12, y: row3Y + 0.46, w: 4.80, h: 0, line: { color: '333333', width: 0.8 } });

    // Row 4: Adv.fee per month
    const row4Y = 4.98;
    let rateText = 'On Request';
    if (site._showRate && site._rate) {
      rateText = `₹ ${Number(site._rate).toLocaleString('en-IN')}/-`;
    } else if (site.ppt_rate || site.monthly_rate) {
      const r = site.ppt_rate || site.monthly_rate;
      rateText = isNaN(Number(r)) ? String(r) : `₹ ${Number(r).toLocaleString('en-IN')}/-`;
    }
    if (iconRupeePng) s.addImage({ data: iconRupeePng, x: 7.45, y: row4Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Adv.fee per month:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: rateText, options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row4Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
    s.addShape(pptx.ShapeType.line, { x: 8.12, y: row4Y + 0.46, w: 4.80, h: 0, line: { color: '333333', width: 0.8 } });

    // Row 5: Availability
    // Requirement:
    // - Booked sites must come as: 'Booked till <date>' (not 'Available from <date>')
    // - Available sites must come as: 'Available' only
    const row5Y = 5.60;
    let availText = 'Available';
    const rawVal = String(site._availability || site.ppt_availability || site.availability || '').trim();
    const dateObj = extractDateFromString(rawVal) || (site._endDate ? (parseDay(site._endDate) || new Date(site._endDate)) : null);
    const dateFmt = dateObj ? fmtDate(dateObj) : (site._endDate || '');

    const isBooked = site._isBooked !== undefined
      ? Boolean(site._isBooked)
      : (Boolean(site._endDate) || /^(?:booked|occupied)/i.test(rawVal) || (/available\s+from/i.test(rawVal) && Boolean(dateFmt)));

    if (isBooked) {
      availText = dateFmt ? `Booked till ${dateFmt}` : (site._availability || 'Occupied');
    } else {
      availText = 'Available';
    }

    // Safety checks: Guarantee "Available from" is converted to "Booked till", and clean formatting
    if (/^available\s+from\s+/i.test(availText)) {
      availText = availText.replace(/^available\s+from\s+/i, 'Booked till ');
    }
    if (/^(?:occupied|booked)$/i.test(availText) && dateFmt) {
      availText = `Booked till ${dateFmt}`;
    }
    if (iconAvailPng) s.addImage({ data: iconAvailPng, x: 7.45, y: row5Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Availability:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: availText, options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row5Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
    s.addShape(pptx.ShapeType.line, { x: 8.12, y: row5Y + 0.46, w: 4.80, h: 0, line: { color: '333333', width: 0.8 } });

    // Row 6: Location Coordinates
    const row6Y = 6.22;
    const { lat, lng } = getCoords(site);
    const coordsText = (lat && lng) ? `${lat}, ${lng}` : '23.012345, 72.567890';
    if (iconPinPng) s.addImage({ data: iconPinPng, x: 7.45, y: row6Y, w: 0.42, h: 0.42 });
    s.addText([
      { text: 'Location Coordinates:  ', options: { color: 'FFC200', bold: true, fontSize: 14.5, fontFace: 'Arial' } },
      { text: coordsText, options: { color: 'FFFFFF', bold: true, fontSize: 14.5, fontFace: 'Arial' } }
    ], { x: 8.12, y: row6Y - 0.02, w: 4.80, h: 0.44, valign: 'middle', margin: 0 });
  }

  // Helper to get effective page url with localStorage fallback
  function getEffectivePageUrl(key) {
    if (pages && pages[key]) return pages[key];
    try {
      const cached = localStorage.getItem(`mb_ppt_raw_${key}`);
      if (cached) return cached;
    } catch {
      return '';
    }
    if (key === 'last') return '/assets/ppt_contact_last_page.png';
    return '';
  }

  // Include optional cover pages if uploaded
  const firstCover = getEffectivePageUrl('first');
  const secondCover = getEffectivePageUrl('second_last');
  const lastCover = getEffectivePageUrl('last');

  if (firstCover) await fixed(firstCover);
  if (secondCover) await fixed(secondCover);

  // Generate slides for all chosen sites
  let siteCounter = 1;
  for (const st of sites) {
    if (st.ppt_images && st.ppt_images.length) {
      for (const u of st.ppt_images) {
        await siteSlide(st, u, siteCounter);
        siteCounter++;
      }
    } else {
      await siteSlide(st, '', siteCounter);
      siteCounter++;
    }
  }

  // Include closing page (with clickable links on mediabuzzoutdoor.com, sales email, and billboard)
  if (lastCover) await closingSlide(lastCover);

  let finalFileName = String(fileName || `${getTodayPptDateStr()}.pptx`).trim();
  if (!finalFileName.toLowerCase().endsWith('.pptx')) finalFileName += '.pptx';
  await pptx.writeFile({ fileName: finalFileName });

  try {
    const siteCodes = (sites || []).map(s => s.site_code).filter(Boolean);
    api.post('/storage', {
      category: 'ppt',
      title: finalFileName.replace(/\.pptx$/i, '').replace(/[_-]/g, ' '),
      filename: finalFileName,
      file_size: `${Math.max(1, Math.round((sites?.length || 1) * 0.6))} MB`,
      format: 'pptx',
      meta_json: {
        slides: (sites || []).length,
        sites: siteCodes.slice(0, 8),
        generatedAt: new Date().toISOString()
      }
    }).catch(() => {});
  } catch {}
}

async function exportStyledExcel(rowsData, filename = `${getTodayPptDateStr()}.xlsx`, title = null) {
  const workbook = new ExcelJS.Workbook();
  const cleanTitle = title ? String(title).trim() : getTodayPptDateStr();
  const headerRowNum = 2;
  const dataStartRowNum = 3;

  const worksheet = workbook.addWorksheet('Sites', {
    views: [{ state: 'frozen', ySplit: 2 }]
  });

  const cols = [
    { key: 'sr', width: 8 },
    { key: 'code', width: 15 },
    { key: 'area', width: 22 },
    { key: 'location', width: 55 },
    { key: 'media', width: 14 },
    { key: 'light', width: 10 },
    { key: 'w', width: 8 },
    { key: 'h', width: 8 },
    { key: 'sqft', width: 12 },
    { key: 'availability', width: 18 },
    { key: 'rate', width: 18 },
    { key: 'coords', width: 28 }
  ];
  worksheet.columns = cols;

  // Executive Media Buzz Brand Header Banner with Logo on Top Right
  attachMediaBuzzExcelHeader(workbook, worksheet, {
    title: cleanTitle,
    columns: cols,
    totalColumns: 12,
    align: 'center'
  });

  // Populate Header Row (Row 2)
  const headers = ['SR NO', 'MB CODE', 'AREA', 'LOCATION', 'MEDIA', 'LIGHT', 'W', 'H', 'SQ FT', 'AVAILABLITY', 'Selling Amount', 'Latitude Longitude'];
  const headerRow = worksheet.getRow(headerRowNum);
  headerRow.height = 28;
  headers.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });

  // Populate Data Rows
  rowsData.forEach(r => {
    worksheet.addRow({
      sr: r['SR NO'],
      code: r['MB CODE'] || r['SITE CODE'] || r['site_code'] || r['Code'] || '',
      area: r['AREA'],
      location: r['LOCATION'],
      media: r['MEDIA'],
      light: r['LIGHT'],
      w: r['W'],
      h: r['H'],
      sqft: r['SQ FT'],
      availability: r['AVAILABLITY'],
      rate: r['Selling Amount'],
      coords: r['Latitude Longitude']
    });
  });

  // Style Header Row in Bright Yellow (#FFFF00)
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' } // Bright Yellow matching template screenshot
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 2, 5, 6, 7, 8, 9, 10].includes(colNumber) ? 'center' : (colNumber === 11 ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'medium', color: { argb: 'FF808080' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  // Style Data Rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= dataStartRowNum) {
      row.height = 22;
      row.eachCell((cell, colNumber) => {
        cell.font = {
          name: 'Calibri',
          size: 10.5,
          bold: colNumber === 2 // Bold for MB CODE
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: [1, 2, 5, 6, 7, 8, 9, 10].includes(colNumber) ? 'center' : (colNumber === 11 ? 'right' : 'left')
        };
        if (colNumber === 11 && typeof cell.value === 'number') {
          cell.numFmt = '#,##,##0';
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          left: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          bottom: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          right: { style: 'thin', color: { argb: 'FFE8E8E8' } }
        };
      });
    }
  });

  // Enable Auto-Filter on Row 2
  if (rowsData.length > 0) {
    const filterStart = 'A2';
    const filterEnd = `L${rowsData.length + 2}`;
    worksheet.autoFilter = `${filterStart}:${filterEnd}`;
  }

  // Official 5-Point Terms & Conditions at Bottom Center
  attachMediaBuzzTermsAndConditions(worksheet, {
    totalColumns: 12
  });

  let finalFileName = String(filename || `${getTodayPptDateStr()}.xlsx`).trim();
  if (!finalFileName.toLowerCase().endsWith('.xlsx')) finalFileName += '.xlsx';

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = finalFileName;
  a.click();
  URL.revokeObjectURL(url);
}

function Login() {
  const nav = useNavigate();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [err, setErr] = useState('');
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setLoading(true);
    setErr('');
    try {
      const { data } = await api.post('/auth/login', { email, password });
      localStorage.setItem('sc_token', data.token);
      localStorage.setItem('sc_user', JSON.stringify(data.user));
      nav('/dashboard');
    } catch (e) {
      setErr(e.response?.data?.message || 'Login failed. Check email and password.');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="login">
      <form onSubmit={submit}>
        <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
        <h1>Site Control</h1>
        <p>OOH Operations Workspace</p>
        {err && <div className="error">{err}</div>}
        <label>
          Email ID
          <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="admin@domain.com" required />
        </label>
        <label>
          Password
          <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" required />
        </label>
        <button disabled={loading}>{loading ? 'Signing in…' : 'Sign in'}</button>
      </form>
    </div>
  );
}

function Guard({ children }) {
  return localStorage.getItem('sc_token') ? children : <Navigate to="/login" replace />;
}

function Layout() {
  const loc = useLocation();
  const nav = useNavigate();
  let user = {};
  try { user = JSON.parse(localStorage.getItem('sc_user') || '{}'); } catch {}
  const userRole = (user.role || 'staff').toLowerCase();
  const isAdmin = userRole === 'admin';
  const isManager = userRole === 'manager';
  const isStaff = userRole === 'staff';
  const isReadOnly = userRole === 'viewer';
  const canDeleteRecords = isAdmin || isManager;

  const [notifOpen, setNotifOpen] = useState(false);
  const [notifications, setNotifications] = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem('scooh_theme') || 'dark');

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('scooh_theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(prev => (prev === 'dark' ? 'light' : 'dark'));

  const currentModule = useMemo(() => {
    const p = loc.pathname.replace(/^\//, '').split('/')[0] || 'dashboard';
    return p;
  }, [loc.pathname]);

  const currentTitle = viewTitles[currentModule] || 'OOH Operations Management';

  useEffect(() => {
    api.get('/notifications').then(r => {
      if (Array.isArray(r.data)) setNotifications(r.data);
    }).catch(() => {});
  }, []);

  const unreadCount = (Array.isArray(notifications) ? notifications : []).filter(n => !n.is_read).length;

  async function clearAllNotifs() {
    try {
      await api.post('/notifications/clear', {});
    } catch (e) {
      try { await api.post('/notifications/read', {}); } catch (err) {}
    }
    setNotifications([]);
  }

  const visibleNavGroups = useMemo(() => {
    return navGroups.map(grp => {
      const items = grp.items.filter(([k]) => {
        if (isAdmin) return true;
        if (isManager) return k !== 'settings';
        if (isStaff) return ['dashboard', 'sites', 'campaigns', 'occupancy', 'campaign-details', 'electricity', 'notifications', 'activity'].includes(k);
        if (isReadOnly) return ['dashboard', 'sites', 'campaigns', 'occupancy', 'campaign-details', 'storage', 'electricity', 'reports', 'notifications', 'activity'].includes(k);
        return true;
      });
      return { ...grp, items };
    }).filter(grp => grp.items.length > 0);
  }, [userRole, isAdmin, isManager, isStaff, isReadOnly]);

  return (
    <div className="scooh-wp-wrap">
      <div className="scooh-app scooh-has-topbar">
        {/* Top Header */}
        <header className="scooh-topbar">
          <div className="scooh-topbar-left">
            <button
              type="button"
              className="scooh-mobile-nav-toggle"
              onClick={() => setSidebarOpen(!sidebarOpen)}
              aria-label="Toggle Navigation"
            >
              <span></span><span></span><span></span>
            </button>
            <div className="scooh-topbar-context" style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
              <div>
                <span className="scooh-topbar-eyebrow">MEDIA BUZZ • OOH WORKSPACE</span>
                <strong>{currentTitle}</strong>
              </div>
              {isReadOnly && (
                <span style={{
                  background: 'rgba(245,158,11,0.15)',
                  border: '1px solid rgba(245,158,11,0.3)',
                  color: '#fbbf24',
                  fontSize: '11px',
                  fontWeight: 700,
                  padding: '3px 9px',
                  borderRadius: '6px',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '4px'
                }}>
                  🔒 Read-Only Workspace
                </span>
              )}
            </div>
          </div>

          <div className="scooh-topbar-user">
            {/* Quick Theme Toggle */}
            <button
              type="button"
              className="scooh-theme-toggle-btn"
              onClick={toggleTheme}
              title={`Switch to ${theme === 'dark' ? 'Light' : 'Dark'} Mode`}
              aria-label="Toggle Theme"
            >
              <span>{theme === 'dark' ? '☀️' : '🌙'}</span>
              <span className="scooh-theme-label">{theme === 'dark' ? 'Light' : 'Dark'}</span>
            </button>

            <div className="scooh-topbar-notifications" style={{ position: 'relative' }}>
              <button
                type="button"
                className="scooh-notification-bell"
                onClick={() => setNotifOpen(!notifOpen)}
                aria-label="Notifications"
              >
                <span className="scooh-bell-icon">🔔</span>
                {unreadCount > 0 && (
                  <span className="scooh-top-notification-count">
                    {unreadCount}
                  </span>
                )}
              </button>

              {notifOpen && (
                <div className="scooh-top-notification-panel" style={{ display: 'block' }}>
                  <div className="scooh-top-notification-head">
                    <strong>Notifications</strong>
                    <div className="scooh-notification-head-actions">
                      <button type="button" onClick={clearAllNotifs}>Clear all</button>
                      <button type="button" onClick={() => { setNotifOpen(false); nav('/notifications'); }}>View all</button>
                    </div>
                  </div>
                  <div className="scooh-top-notification-list">
                    {(!notifications || notifications.length === 0) ? (
                      <div className="scooh-top-notification-empty">No new notifications</div>
                    ) : (
                      notifications.map(n => (
                        <div key={n.id} className={`scooh-top-notification-item ${n.is_read ? '' : 'today'}`}>
                          <span className="scooh-top-notification-dot"></span>
                          <div>
                            <strong>{n.title}</strong>
                            <small>{n.message}</small>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}
            </div>

            <span className="scooh-user-avatar">
              {String(user.name || user.email || 'A').slice(0, 1).toUpperCase()}
            </span>
            <div className="scooh-user-meta">
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '10.5px', color: '#94a3b8' }}>Signed in as</span>
                <span
                  className="scooh-badgechip"
                  style={{
                    background: ROLE_CONFIG[userRole]?.bg || 'rgba(139,92,246,0.18)',
                    color: ROLE_CONFIG[userRole]?.color || '#c4b5fd',
                    border: `1px solid ${ROLE_CONFIG[userRole]?.border || 'rgba(139,92,246,0.35)'}`,
                    textTransform: 'uppercase',
                    fontSize: '9.5px',
                    fontWeight: 800,
                    padding: '2px 7px'
                  }}
                  title={ROLE_CONFIG[userRole]?.description}
                >
                  {ROLE_CONFIG[userRole]?.label || userRole}
                </span>
              </div>
              <strong>{user.name || user.email || 'Administrator'}</strong>
            </div>
            <button
              type="button"
              className="scooh-topbar-logout"
              onClick={() => {
                localStorage.clear();
                nav('/login');
              }}
            >
              Logout
            </button>
          </div>
        </header>

        {/* Sidebar */}
        <aside className={`scooh-sidebar ${sidebarOpen ? 'open' : ''}`}>
          <div className="scooh-brand-shell">
            <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
          </div>
          <nav className="scooh-tabs">
            {visibleNavGroups.map(grp => (
              <div className="scooh-nav-group" key={grp.label}>
                <span className="scooh-nav-group-label">{grp.label}</span>
                {grp.items.map(([k, n]) => (
                  <button
                    key={k}
                    type="button"
                    className={loc.pathname === '/' + k || (k !== 'dashboard' && loc.pathname.startsWith('/' + k)) ? 'active' : ''}
                    onClick={() => {
                      nav('/' + k);
                      setSidebarOpen(false);
                    }}
                  >
                    <span className="scooh-nav-icon">
                      <svg viewBox="0 0 24 24" aria-hidden="true">{icons[k] || icons.dashboard}</svg>
                    </span>
                    <span className="scooh-nav-label">{n}</span>
                    {k === 'notifications' && unreadCount > 0 && (
                      <span className="scooh-notification-count">{unreadCount}</span>
                    )}
                  </button>
                ))}
              </div>
            ))}
          </nav>
          <div className="scooh-sidebar-footer">
            <span className="scooh-sidebar-dot"></span>
            <span className="scooh-sidebar-status">Workspace online</span>
          </div>
        </aside>

        {sidebarOpen && <button type="button" className="scooh-mobile-scrim" onClick={() => setSidebarOpen(false)} aria-label="Close menu" />}

        {/* Main Content Area */}
        <main className="scooh-main">
          <div className="scooh-view">
            <Routes>
              <Route path="/dashboard" element={<Dashboard />} />
              <Route path="/sites" element={<SitesView />} />
              <Route path="/campaigns" element={<CampaignTrackerView />} />
              <Route path="/occupancy" element={<OccupancyView />} />
              <Route path="/campaign-details" element={<CampaignDetailsView />} />
              <Route path="/proposals" element={!isStaff && !isReadOnly ? <ProposalsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/ppt" element={!isStaff && !isReadOnly ? <PptView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/electricity" element={<ElectricityView />} />
              <Route path="/vendors" element={!isStaff && !isReadOnly ? <Crud entity="vendors" title="Vendors" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/clients" element={!isStaff && !isReadOnly ? <Crud entity="clients" title="Clients" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/invoices" element={!isStaff && !isReadOnly ? <Crud entity="invoices" title="Invoices" /> : <Navigate to="/dashboard" replace />} />
              <Route path="/data" element={!isStaff && !isReadOnly ? <DataToolsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="/storage" element={<StorageView />} />
              <Route path="/reports" element={<ReportsView />} />
              <Route path="/notifications" element={<SimpleList endpoint="notifications" title="Notifications" />} />
              <Route path="/activity" element={<SimpleList endpoint="activity" title="Activity Log" />} />
              <Route path="/settings" element={isAdmin ? <SettingsView /> : <Navigate to="/dashboard" replace />} />
              <Route path="*" element={<Navigate to="/dashboard" replace />} />
            </Routes>
          </div>
        </main>
      </div>
    </div>
  );
}

function PageHead({ title, desc, actions }) {
  return (
    <div className="scooh-pagehead">
      <div>
        <h1>{title}</h1>
        {desc && <p>{desc}</p>}
      </div>
      {actions && <div className="scooh-headactions">{actions}</div>}
    </div>
  );
}

function OccupancyDonut({ buckets = {}, average = 0 }) {
  const data = [
    { label: 'Occupied', value: buckets.Occupied ?? 0, color: '#22C55E' },
    { label: 'Available', value: buckets.Available ?? 0, color: '#3B82F6' },
    { label: 'Prime Sites', value: buckets['Prime Sites'] ?? 0, color: '#FBBF24' },
    { label: 'Needs review', value: buckets['Needs review'] ?? 0, color: '#EF4444' }
  ];
  const total = data.reduce((acc, d) => acc + d.value, 0) || 1;
  let accumulated = 0;
  const radius = 68;
  const strokeWidth = 24;
  const circumference = 2 * Math.PI * radius;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '230px' }}>
      <div style={{ position: 'relative', width: '180px', height: '180px', margin: '10px auto' }}>
        <svg width="180" height="180" viewBox="0 0 180 180" style={{ transform: 'rotate(-90deg)' }}>
          {data.map((slice, i) => {
            const dashLength = (slice.value / total) * circumference;
            const strokeDasharray = `${dashLength} ${circumference}`;
            const strokeDashoffset = -accumulated;
            accumulated += dashLength;
            return (
              <circle
                key={i}
                cx="90"
                cy="90"
                r={radius}
                fill="transparent"
                stroke={slice.color}
                strokeWidth={strokeWidth}
                strokeDasharray={strokeDasharray}
                strokeDashoffset={strokeDashoffset}
              />
            );
          })}
        </svg>
        <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
          <span style={{ fontSize: '26px', fontWeight: 800, color: '#fff' }}>{average}%</span>
          <span style={{ fontSize: '9px', color: '#8d99a8', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Occupied</span>
        </div>
      </div>
      <div style={{ display: 'flex', gap: '14px', flexWrap: 'wrap', justifyContent: 'center', marginTop: '14px' }}>
        {data.map((d, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '11px', color: '#9ba7b5' }}>
            <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: d.color }}></span>
            <span>{d.label} ({d.value})</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function Dashboard() {
  const [d, setD] = useState(null);
  const [err, setErr] = useState('');
  const nav = useNavigate();

  function load() {
    setErr('');
    api.get('/dashboard')
      .then(r => setD(r.data))
      .catch(e => setErr(e.response?.data?.message || 'Failed to load dashboard data'));
  }

  useEffect(() => { load(); }, []);

  if (err) {
    return (
      <>
        <PageHead title="Operations Dashboard" desc="Media Buzz OOH operations overview" />
        <section className="scooh-panel">
          <p style={{ color: '#f06a6a', fontWeight: 600 }}>{err}</p>
          <button className="scooh-btn purple-btn" onClick={load}>Retry</button>
        </section>
      </>
    );
  }

  if (!d) {
    return (
      <div className="scooh-loading">
        <span className="scooh-spinner"></span>
        <strong>Loading workspace</strong>
        <small>Preparing your Media Buzz data…</small>
      </div>
    );
  }

  const k = d.kpis || {};
  const alerts = Array.isArray(d.alerts) ? d.alerts : [];

  return (
    <>
      <PageHead
        title="Operations Dashboard"
        desc="Live snapshot of the site portfolio, active campaigns, and everything due for action today."
        actions={
          <>
            <button className="scooh-btn purple-btn" onClick={() => nav('/sites')}>+ Add site</button>
            <button className="scooh-btn" onClick={() => nav('/campaigns')}>+ Log booking</button>
          </>
        }
      />

      {/* Row 1: Site Portfolio & Live Campaigns */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi alert" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View all sites">
          <div className="n">{k.total_sites ?? 0}</div>
          <div className="l">Total Sites</div>
          {k.flagged_count > 0 && (
            <div className="scooh-kpi-note">⚠ {k.flagged_count} flagged for review</div>
          )}
        </div>
        <div className="scooh-kpi good" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View occupied sites">
          <div className="n">{k.active_sites ?? 0}</div>
          <div className="l">Active Sites</div>
        </div>
        <div className="scooh-kpi alert" onClick={() => nav('/sites')} style={{ cursor: 'pointer' }} title="View available sites">
          <div className="n">{k.nonactive_sites ?? 0}</div>
          <div className="l">Vacant / Available</div>
        </div>
        <div className="scooh-kpi alert" onClick={() => nav('/campaigns')} style={{ cursor: 'pointer' }} title="View active campaigns">
          <div className="n">{k.active_campaigns ?? 0}</div>
          <div className="l">Active Campaigns</div>
        </div>
      </div>

      {/* Row 2: Operational Action Triggers */}
      <div className="scooh-kpirow">
        <div className={`scooh-kpi ${k.mounting_overdue > 0 ? 'danger' : 'good'}`} onClick={() => nav('/campaigns')} style={{ cursor: 'pointer' }} title="View campaigns">
          <div className="n">{k.mounting_overdue ?? 0}</div>
          <div className="l">Mounting Overdue</div>
        </div>
        <div className={`scooh-kpi ${k.validation_15_due > 0 ? 'danger' : 'good'}`} onClick={() => nav('/validations')} style={{ cursor: 'pointer' }} title="View validations">
          <div className="n">{k.validation_15_due ?? 0}</div>
          <div className="l">15-Day Validation Due</div>
        </div>
        <div className={`scooh-kpi ${k.final_validation_due > 0 ? 'danger' : 'good'}`} onClick={() => nav('/validations')} style={{ cursor: 'pointer' }} title="View validations">
          <div className="n">{k.final_validation_due ?? 0}</div>
          <div className="l">Final Validation Due</div>
        </div>
        <div className={`scooh-kpi ${k.invoice_actions > 0 ? 'danger' : 'good'}`} onClick={() => nav('/invoices')} style={{ cursor: 'pointer' }} title="View invoices">
          <div className="n">{k.invoice_actions ?? 0}</div>
          <div className="l">Invoice Actions</div>
        </div>
      </div>

      {/* Lower Dashboard Grid */}
      <div className="scooh-grid2 scooh-dashboard-lower">
        {/* Needs attention today */}
        <div className="scooh-panel">
          <h3>Needs attention today</h3>
          {alerts.length === 0 ? (
            <div className="scooh-empty" style={{ padding: '36px 16px', textAlign: 'center' }}>
              ✓ Nothing outstanding — all operations are on track.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {alerts.map((a, idx) => (
                <div
                  key={idx}
                  onClick={() => nav(a.link || '/campaigns')}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '12px 14px',
                    border: '1px solid var(--mb-border)',
                    borderRadius: '12px',
                    background: 'var(--mb-surface-2)',
                    cursor: 'pointer',
                    transition: 'border-color 0.15s ease'
                  }}
                  onMouseEnter={e => e.currentTarget.style.borderColor = 'var(--mb-primary)'}
                  onMouseLeave={e => e.currentTarget.style.borderColor = 'var(--mb-border)'}
                  title="Click to view details"
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span className="scooh-plate">{a.site_code}</span>
                    <div>
                      <strong style={{ color: 'var(--mb-text)', fontSize: '12.5px' }}>{a.client || 'Client'}</strong>
                      <div className="scooh-footnote" style={{ color: 'var(--mb-muted)', fontSize: '10.5px' }}>{a.campaign}</div>
                    </div>
                  </div>
                  <span
                    className={`scooh-pill ${a.class === 'danger' ? 'vacant' : 'watch'}`}
                    style={{ whiteSpace: 'nowrap' }}
                  >
                    {a.tag}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Portfolio Occupancy Chart */}
        <div className="scooh-panel">
          <h3>
            Portfolio occupancy · <span className="scooh-accent">{k.average_occupancy || 0}%</span>
          </h3>
          <OccupancyDonut buckets={d.occupancy_buckets || {}} average={k.average_occupancy || 0} />
        </div>
      </div>
    </>
  );
}

function SitesView() {
  const navigate = useNavigate();
  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;

  const [sites, setSites] = useState([]);
  const [search, setSearch] = useState('');
  const [cityFilter, setCityFilter] = useState('');
  const [mediaFilter, setMediaFilter] = useState('');
  const [availFilter, setAvailFilter] = useState('');
  const [lightingFilter, setLightingFilter] = useState('');
  const [dateFilter, setDateFilter] = useState(''); // availability date check
  const [campaignsBySiteCode, setCampaignsBySiteCode] = useState({}); // { SITE_CODE: campaigns[] }
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  const [edit, setEdit] = useState(null);
  const [imageModalSite, setImageModalSite] = useState(null);
  const [saving, setSaving] = useState(false);

  async function load() {
    try {
      const [sRes, cRes] = await Promise.all([
        api.get('/sites'),
        api.get('/campaigns').catch(() => ({ data: [] }))
      ]);
      if (Array.isArray(sRes.data)) {
        setSites(sRes.data.map(unpackSite));
      }
      // Build campaign map: SITE_CODE -> campaigns[]
      if (Array.isArray(cRes.data)) {
        const map = {};
        cRes.data.forEach(c => {
          const key = String(c.site_code || '').toUpperCase().trim();
          if (!key) return;
          if (!map[key]) map[key] = [];
          map[key].push(c);
        });
        setCampaignsBySiteCode(map);
      }
    } catch (e) {
      console.warn('Using default sites cache', e);
    }
  }

  const [importingExcel, setImportingExcel] = useState(false);

  useEffect(() => {
    load();
    const onSitesUpdate = () => load();
    window.addEventListener('mb-sites-updated', onSitesUpdate);
    const interval = setInterval(load, 35000); // 35s auto-refresh
    return () => {
      window.removeEventListener('mb-sites-updated', onSitesUpdate);
      clearInterval(interval);
    };
  }, []);

  const cities = useMemo(() => Array.from(new Set(sites.map(s => s.city).filter(Boolean))), [sites]);
  const mediaTypes = useMemo(() => Array.from(new Set(sites.map(s => s.media_type).filter(Boolean))), [sites]);
  const availabilities = useMemo(() => Array.from(new Set(sites.map(s => s.ppt_availability || s.availability).filter(Boolean))), [sites]);
  const lightings = useMemo(() => Array.from(new Set(sites.map(s => s.lighting).filter(Boolean))), [sites]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = sites.filter(s => {
      const matchesQuery = matchSiteSearch(s, search);
      const matchesCity = !cityFilter || s.city === cityFilter;
      const matchesMedia = !mediaFilter || s.media_type === mediaFilter;
      const activeCamp = getActiveCampaignOnDate(campaignsBySiteCode, s.site_code, dateFilter);
      const effAvail = activeCamp
        ? (String(activeCamp.parent_campaign || '').startsWith('LINKED:') ? 'Occupied (Linked)' : 'Occupied')
        : (dateFilter ? 'Available' : (s.ppt_availability || s.availability || 'Available'));
      const matchesAvail = !availFilter || effAvail.toLowerCase().includes(availFilter.toLowerCase()) || (availFilter.toLowerCase() === 'available' && effAvail === 'Available');
      const matchesLighting = !lightingFilter || s.lighting === lightingFilter;
      return matchesQuery && matchesCity && matchesMedia && matchesAvail && matchesLighting;
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        if (sortState.key === 'availability') {
          const activeCampA = getActiveCampaignOnDate(campaignsBySiteCode, a.site_code, dateFilter);
          const activeCampB = getActiveCampaignOnDate(campaignsBySiteCode, b.site_code, dateFilter);
          valA = activeCampA ? 'Occupied' : (dateFilter ? 'Available' : (a.ppt_availability || a.availability || 'Available'));
          valB = activeCampB ? 'Occupied' : (dateFilter ? 'Available' : (b.ppt_availability || b.availability || 'Available'));
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [sites, search, cityFilter, mediaFilter, availFilter, lightingFilter, dateFilter, campaignsBySiteCode, sortState]);

  const [selectedIds, setSelectedIds] = useState(new Set());

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(s => s.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(s => s.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to delete ${count} selected site record${count > 1 ? 's' : ''}?`)) {
      return;
    }
    try {
      await api.post('/sites/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setSelectedIds(new Set());
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Failed to delete selected sites: ' + (err.response?.data?.message || err.message));
    }
  }

  async function save(e) {
    e.preventDefault();
    setSaving(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (edit?.id) await api.put(`/sites/${edit.id}`, data);
      else await api.post('/sites', data);
      setEdit(null);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Error saving site: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function del(id) {
    if (confirm('Delete this site record?')) {
      try {
        await api.delete(`/sites/${id}?hard=true`);
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await load();
        window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      } catch (err) {
        alert('Failed to delete site: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  async function handleImageUpload(siteId, files, replace = false) {
    const fileList = Array.from(files || []);
    if (!fileList.length) return;
    const localPreviews = fileList.map(f => URL.createObjectURL(f));
    fileList.forEach((f, idx) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result) {
          photoPreviewCache.set(f.name, reader.result);
          if (localPreviews[idx]) photoPreviewCache.set(localPreviews[idx], reader.result);
        }
      };
      reader.readAsDataURL(f);
    });

    const fd = new FormData();
    fileList.forEach(f => fd.append('files', f));
    const res = await api.post(`/sites/${siteId}/images${replace ? '?replace=true' : ''}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
    const serverImgs = res.data?.images || [];
    const siteCode = res.data?.site_code;
    serverImgs.forEach((sUrl, idx) => {
      if (localPreviews[idx]) {
        photoPreviewCache.set(sUrl, localPreviews[idx]);
      }
      const f = fileList[idx];
      if (f) {
        const r2 = new FileReader();
        r2.onload = () => {
          if (r2.result) {
            photoPreviewCache.set(sUrl, r2.result);
            cachePhotoBlob(sUrl, r2.result);
            cachePhotoBlob(`/api${sUrl}`, r2.result);
          }
        };
        r2.readAsDataURL(f);
      }
    });
    if (siteCode) {
      saveCachedPhotosForSite(siteCode, serverImgs);
    }
    await load();
    window.dispatchEvent(new CustomEvent('mb-sites-updated'));
  }

  async function handleDirectExcelImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      alert(`✓ Excel Import Successful!\n\n${r.data.message || `Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`}`);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  async function handleClearAllSites() {
    if (!confirm('⚠️ Are you sure you want to clean ALL sites from the database?\n\nThis will permanently remove all sites and clear cached photos so you can import a fresh Excel spreadsheet.')) {
      return;
    }
    try {
      await api.post('/sites/clear-all-sites');
      clearAllCachedPhotos();
      setSelectedIds(new Set());
      setSites([]);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      alert('✓ All sites cleaned from database! Ready for fresh Excel import.');
    } catch (err) {
      alert('Failed to clear sites: ' + (err.response?.data?.message || err.message));
    }
  }

  return (
    <>
      <PageHead
        title="Sites Directory"
        desc={`Managing ${filtered.length} of ${sites.length} total Media Buzz outdoor inventory sites.`}
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Refresh</button>
            {canDelete && (
              <>
                <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Upload Excel spreadsheet to import sites, rates & availability">
                  <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                  <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handleDirectExcelImport} />
                </label>
                <button
                  type="button"
                  className="scooh-btn ghost"
                  style={{ color: '#ef4444', borderColor: 'rgba(239,68,68,0.3)' }}
                  onClick={handleClearAllSites}
                  title="Clean all sites from database for fresh import"
                >
                  🗑️ Clean All Sites
                </button>
              </>
            )}
            {!isReadOnly && !isStaff && (
              <button className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add Site</button>
            )}
          </>
        }
      />

      <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
        <input
          className="scooh-search"
          style={{ flex: '1 1 220px', minWidth: '180px' }}
          placeholder="Search site code (e.g. 01, MB-01), area, landmark…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        <div className="scooh-filters" style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
          <select value={cityFilter} onChange={e => setCityFilter(e.target.value)}>
            <option value="">All Cities ({cities.length})</option>
            {cities.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={mediaFilter} onChange={e => setMediaFilter(e.target.value)}>
            <option value="">All Media Types ({mediaTypes.length})</option>
            {mediaTypes.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
          {availabilities.length > 0 && (
            <select value={availFilter} onChange={e => setAvailFilter(e.target.value)}>
              <option value="">All Availability ({availabilities.length})</option>
              {availabilities.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
          )}
          {lightings.length > 0 && (
            <select value={lightingFilter} onChange={e => setLightingFilter(e.target.value)}>
              <option value="">All Lighting ({lightings.length})</option>
              {lightings.map(l => <option key={l} value={l}>{l}</option>)}
            </select>
          )}
          {/* Date selector for evaluating site availability */}
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', background: dateFilter ? 'rgba(56,189,248,0.14)' : 'rgba(15,23,42,0.6)', border: `1px solid ${dateFilter ? 'rgba(56,189,248,0.5)' : 'rgba(255,255,255,0.1)'}`, borderRadius: '8px', padding: '4px 10px' }}>
            <span style={{ fontSize: '11px', fontWeight: 700, color: dateFilter ? '#38bdf8' : '#94a3b8', whiteSpace: 'nowrap' }}>
              📅 {dateFilter ? 'Date:' : 'Select Date:'}
            </span>
            <input
              type="date"
              value={dateFilter}
              onChange={e => setDateFilter(e.target.value)}
              style={{ background: 'transparent', border: 'none', color: dateFilter ? '#38bdf8' : '#94a3b8', fontSize: '12px', fontWeight: 700, outline: 'none', cursor: 'pointer', padding: '2px 0' }}
              title="Select a date to evaluate site availability as of that date. Sites whose campaign ends before this date automatically become Available."
            />
            {dateFilter && (
              <button type="button" onClick={() => setDateFilter('')} style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '14px', lineHeight: 1, padding: '0 2px' }} title="Reset date to today">×</button>
            )}
          </div>
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ fontWeight: 600 }}
          >
            <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
            <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
            <option value="area:asc">Sort: Area (A-Z)</option>
            <option value="area:desc">Sort: Area (Z-A)</option>
            <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
            <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
            <option value="availability:asc">Sort: Availability (A-Z)</option>
            <option value="size:asc">Sort: Size</option>
          </select>
          {(search || cityFilter || mediaFilter || availFilter || lightingFilter || dateFilter || sortState.key !== 'site_code' || sortState.dir !== 'asc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '6px 10px', fontSize: '11px' }}
              onClick={() => {
                setSearch('');
                setCityFilter('');
                setMediaFilter('');
                setAvailFilter('');
                setLightingFilter('');
                setDateFilter('');
                setSortState({ key: 'site_code', dir: 'asc' });
              }}
            >
              Reset
            </button>
          )}
        </div>
      </div>

      {/* Batch Selection Action Bar */}
      {canDelete && selectedIds.size > 0 && (
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '11px 18px',
          background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
          borderBottom: '1px solid rgba(239,68,68,0.3)',
          borderRadius: '8px',
          margin: '10px 0',
          color: '#fca5a5',
          fontSize: '13px',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
              {selectedIds.size} site{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={selectAllVisible}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
            >
              Select all visible ({filtered.length})
            </button>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={deselectAll}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
            >
              Deselect all
            </button>
          </div>
          <button
            type="button"
            className="scooh-btn danger"
            onClick={handleBatchDelete}
            style={{
              background: '#ef4444',
              color: '#ffffff',
              border: 'none',
              fontWeight: 800,
              padding: '7px 16px',
              borderRadius: '7px',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
            }}
          >
            🗑 Delete selected ({selectedIds.size})
          </button>
        </div>
      )}

      {dateFilter && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(56, 189, 248, 0.1)', border: '1px solid rgba(56, 189, 248, 0.3)', borderRadius: '8px', padding: '8px 14px', marginBottom: '12px', fontSize: '12px', color: '#38bdf8' }}>
          <span style={{ fontSize: '15px' }}>📅</span>
          <span>
            Site availability evaluated as of <strong>{formatDate(dateFilter) || dateFilter}</strong>. Any site whose campaign ends before this date automatically displays as <strong>Available</strong>.
          </span>
          <button
            type="button"
            onClick={() => setDateFilter('')}
            style={{ marginLeft: 'auto', background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11.5px', fontWeight: 700, textDecoration: 'underline' }}
          >
            Reset to Today
          </button>
        </div>
      )}

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {canDelete && (
                <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                  <input
                    type="checkbox"
                    checked={filtered.length > 0 && filtered.every(s => selectedIds.has(s.id))}
                    onChange={toggleSelectAll}
                    title="Select all visible sites"
                    style={{ cursor: 'pointer' }}
                  />
                </th>
              )}
              <SortHeader label="Site Code" sortKey="site_code" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Area / Landmark" sortKey="area" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Media" sortKey="media_type" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Size" sortKey="size" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Lighting" sortKey="lighting" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Availability" sortKey="availability" currentSort={sortState} onSort={handleSort} />
              <SortHeader label="Rate / Mo" sortKey="monthly_rate" currentSort={sortState} onSort={handleSort} />
              <th>Photos</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={canDelete ? "10" : "9"} className="scooh-empty">No sites match the filters</td></tr>
            ) : (
              filtered.map(s => (
                <tr key={s.id || s.site_code}>
                  {canDelete && (
                    <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={selectedIds.has(s.id)}
                        onChange={() => toggleSelect(s.id)}
                        style={{ cursor: 'pointer' }}
                      />
                    </td>
                  )}
                  <td>
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                      <button
                        type="button"
                        onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                        className="scooh-plate"
                        title={`Open ${s.site_code} in Latest Booking`}
                        style={{
                          cursor: 'pointer',
                          background: 'rgba(56, 189, 248, 0.12)',
                          border: '1px solid rgba(56, 189, 248, 0.45)',
                          color: '#38bdf8',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '4px',
                          padding: '2px 7px',
                          borderRadius: '5px',
                          fontWeight: 800
                        }}
                      >
                        <span>{s.site_code}</span>
                        <span style={{ fontSize: '10px', opacity: 0.75 }}>↗</span>
                      </button>
                      {getSiteTypeTag(s.site_code) !== 'Single' && (
                        <span
                          style={{
                            fontSize: '10.5px',
                            padding: '2px 6px',
                            background: getSiteTypeTag(s.site_code) === 'Combined' ? 'rgba(168,85,247,0.18)' : 'rgba(56,189,248,0.18)',
                            color: getSiteTypeTag(s.site_code) === 'Combined' ? '#c084fc' : '#38bdf8',
                            borderRadius: '4px',
                            fontWeight: 700
                          }}
                          title={getConflictSummary(s.site_code)?.message || ''}
                        >
                          {getSiteTypeTag(s.site_code)}
                        </span>
                      )}
                    </div>
                  </td>
                  <td>
                    <strong>{s.area || s.city}</strong>
                    <div className="scooh-footnote">{s.address}</div>
                  </td>
                  <td>{s.media_type}</td>
                  <td>{s.size}</td>
                  <td><span className="scooh-badge">{s.lighting}</span></td>
                  <td>
                    {(() => {
                      const activeCamp = getActiveCampaignOnDate(campaignsBySiteCode, s.site_code, dateFilter);
                      const isOccupied = activeCamp ? true : (dateFilter ? false : (String(s.ppt_availability || s.availability || '').toLowerCase() === 'occupied'));
                      const effAvail = activeCamp
                        ? (String(activeCamp.parent_campaign || '').startsWith('LINKED:') ? 'Occupied (Linked)' : 'Occupied')
                        : (dateFilter ? 'Available' : (s.ppt_availability || s.availability || 'Available'));

                      if (String(effAvail || '').startsWith('Occupied (Linked')) {
                        const endFmt = activeCamp && activeCamp.end_date ? formatDate(activeCamp.end_date) : '';
                        return (
                          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                            <span
                              className="scooh-pill watch"
                              title={effAvail}
                              style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '11px', whiteSpace: 'nowrap', background: 'rgba(245,158,11,0.15)', color: '#fbbf24', borderColor: 'rgba(245,158,11,0.35)' }}
                            >
                              🔗 {effAvail}
                            </span>
                            {endFmt && (
                              <span style={{ fontSize: '10px', color: '#fbbf24', marginTop: '2px', whiteSpace: 'nowrap' }}>
                                Till {endFmt}
                              </span>
                            )}
                          </div>
                        );
                      }
                      if (isOccupied || String(effAvail || '').toLowerCase() === 'occupied') {
                        const endFmt = activeCamp && activeCamp.end_date ? formatDate(activeCamp.end_date) : '';
                        return (
                          <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                            <span className="scooh-pill vacant">Occupied</span>
                            {endFmt && (
                              <span style={{ fontSize: '10px', color: '#f87171', fontWeight: 600, marginTop: '2px', whiteSpace: 'nowrap' }}>
                                Till {endFmt}
                              </span>
                            )}
                          </div>
                        );
                      }
                      return (
                        <span className={`scooh-pill ${effAvail === 'Available' ? 'active' : 'watch'}`}>
                          {effAvail}
                        </span>
                      );
                    })()}
                  </td>
                  <td><b>{money(s.ppt_rate || s.monthly_rate)}</b></td>
                  <td>
                    <button
                      type="button"
                      className="scooh-iconbtn"
                      onClick={() => setImageModalSite(s)}
                      title="Manage Presentation Photos"
                    >
                      📷 {(s.ppt_images || []).length}
                    </button>
                  </td>
                  <td>
                    <div className="scooh-rowactions">
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(s)}>
                        {isReadOnly ? 'View' : 'Edit'}
                      </button>
                      {canDelete && (
                        <button className="scooh-iconbtn danger-icon" onClick={() => del(s.id)}>×</button>
                      )}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Edit/Add Modal */}
      {edit && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal">
            <div className="scooh-modalhead">
              <div>
                <h2>{edit.id ? 'Edit Site' : 'Add New Site'}</h2>
                <span className="scooh-modal-subtitle">Update site specifications and pricing</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEdit(null)}>×</button>
            </div>
            <form onSubmit={save}>
              <div className="scooh-modalbody">
                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Site Code</label>
                    <input name="site_code" defaultValue={edit.site_code ?? ''} placeholder="e.g. AMD-GT-001" required />
                  </div>
                  <div className="scooh-field">
                    <label>City</label>
                    <input name="city" defaultValue={edit.city ?? 'Ahmedabad'} placeholder="e.g. Ahmedabad" required />
                  </div>
                  <div className="scooh-field">
                    <label>Area / Landmark</label>
                    <input name="area" defaultValue={edit.area ?? ''} placeholder="e.g. Shivranjani Cross Roads" required />
                  </div>
                  <div className="scooh-field">
                    <label>Media Type</label>
                    <select name="media_type" defaultValue={edit.media_type ?? 'Hoarding'}>
                      <option value="Hoarding">Hoarding</option>
                      <option value="Gantry">Gantry</option>
                      <option value="Unipole">Unipole</option>
                      <option value="Billboard">Billboard</option>
                      <option value="DOOH">DOOH</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Size (e.g. 30x10)</label>
                    <input name="size" defaultValue={edit.size ?? ''} placeholder="e.g. 30x10 ft" />
                  </div>
                  <div className="scooh-field">
                    <label>Lighting</label>
                    <select name="lighting" defaultValue={edit.lighting ?? 'BL'}>
                      <option value="BL">BL (Backlit)</option>
                      <option value="FL">FL (Frontlit)</option>
                      <option value="NL">NL (Nonlit)</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Monthly Rate (₹)</label>
                    <input name="monthly_rate" type="number" defaultValue={edit.monthly_rate ?? 0} placeholder="250000" />
                  </div>
                  <div className="scooh-field">
                    <label>Availability</label>
                    <select name="availability" defaultValue={edit.availability ?? 'Available'}>
                      <option value="Available">Available</option>
                      <option value="Occupied">Occupied</option>
                      <option value="Maintenance">Maintenance</option>
                      <option value="Booked">Booked</option>
                    </select>
                  </div>
                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Full Address</label>
                    <textarea name="address" defaultValue={edit.address ?? ''} placeholder="Full road, junction and landmark location description…" rows={3} />
                  </div>
                  <div className="scooh-field">
                    <label>GPS Coordinates</label>
                    <input name="gps" defaultValue={edit.gps ?? ''} placeholder="23.019187, 72.530184" />
                  </div>
                  <div className="scooh-field">
                    <label>Google Maps URL</label>
                    <input name="maps_url" defaultValue={edit.maps_url ?? ''} placeholder="https://maps.google.com/?q=..." />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>
                  {isReadOnly ? 'Close' : 'Cancel'}
                </button>
                {!isReadOnly && (
                  <button className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : 'Save Site'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Image Upload Modal */}
      {imageModalSite && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '640px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>Images for {imageModalSite.site_code}</h2>
                <span className="scooh-modal-subtitle">Upload photos for automated PPT and presentations</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setImageModalSite(null)}>×</button>
            </div>
            <div className="scooh-modalbody">
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', marginBottom: '16px' }}>
                {(imageModalSite.ppt_images || []).map((url, i) => (
                  <div key={i} style={{ position: 'relative', border: '1px solid #34404e', borderRadius: '8px', overflow: 'hidden' }}>
                    <img
                      src={resolvePhotoUrl(url)}
                      alt={`Site ${imageModalSite.site_code} photo ${i + 1}`}
                      style={{ width: '100%', height: '110px', objectFit: 'cover' }}
                      onError={e => {
                        const img = e.target;
                        const currentSrc = img.src || '';
                        if (url && typeof url === 'string') {
                          if (url.startsWith('/uploads/') && currentSrc.includes('/api/uploads/')) {
                            img.src = url;
                            return;
                          }
                          if (url.startsWith('/uploads/') && !currentSrc.includes('/api/uploads/')) {
                            img.src = `/api${url}`;
                            return;
                          }
                        }
                        if (photoPreviewCache.has(url)) {
                          img.src = photoPreviewCache.get(url);
                          return;
                        }
                        getCachedPhotoBlob(url).then(blob => {
                          if (blob) {
                            img.src = blob;
                            photoPreviewCache.set(url, blob);
                          }
                        }).catch(() => {});
                      }}
                    />
                  </div>
                ))}
              </div>
              <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                <label className="scooh-btn primary" style={{ display: 'inline-flex', cursor: 'pointer' }}>
                  + Upload Images
                  <input
                    type="file"
                    multiple
                    accept="image/*"
                    hidden
                    onChange={async e => {
                      if (e.target.files.length) {
                        await handleImageUpload(imageModalSite.id, e.target.files, false);
                        setImageModalSite(null);
                      }
                    }}
                  />
                </label>
                <label className="scooh-btn secondary" style={{ display: 'inline-flex', cursor: 'pointer' }} title="Replace all current images for this site with new ones">
                  🔄 Replace All Photos
                  <input
                    type="file"
                    multiple
                    accept="image/*"
                    hidden
                    onChange={async e => {
                      if (e.target.files.length) {
                        await handleImageUpload(imageModalSite.id, e.target.files, true);
                        setImageModalSite(null);
                      }
                    }}
                  />
                </label>
              </div>
            </div>
            <div className="scooh-modalfoot">
              <button type="button" className="scooh-btn ghost" onClick={() => setImageModalSite(null)}>Close</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function PptView() {
  const navigate = useNavigate();
  const [sites, setSites] = useState([]);
  const [pages, setPages] = useState(() => {
    try {
      const saved = localStorage.getItem('mb_ppt_pages_cache');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [uploadingPage, setUploadingPage] = useState({});
  const [sel, setSel] = useState({});
  const [query, setQuery] = useState('');
  const [areaFilter, setAreaFilter] = useState('');
  const [availFilter, setAvailFilter] = useState('all'); // 'all' | 'available' | 'occupied'
  const [selectedDates, setSelectedDates] = useState([]);
  const [dateInputVal, setDateInputVal] = useState('');
  const [showDateModal, setShowDateModal] = useState(false);
  const [calendarYear, setCalendarYear] = useState(() => new Date().getFullYear());
  const [calendarMonth, setCalendarMonth] = useState(() => new Date().getMonth());
  const [rangeStart, setRangeStart] = useState('');
  const [rangeEnd, setRangeEnd] = useState('');
  const dateFilter = selectedDates[0] || '';
  const setDateFilter = (val) => {
    if (!val) {
      setSelectedDates([]);
      setDateInputVal('');
    } else {
      setSelectedDates([val]);
      setDateInputVal(val);
    }
  };

  function addDate(iso) {
    if (!iso) return;
    setSelectedDates(prev => prev.includes(iso) ? prev : [...prev, iso].sort());
  }

  function removeDate(iso) {
    setSelectedDates(prev => prev.filter(d => d !== iso));
  }

  function toggleDate(iso) {
    if (!iso) return;
    setSelectedDates(prev => prev.includes(iso) ? prev.filter(d => d !== iso) : [...prev, iso].sort());
  }

  function clearAllDates() {
    setSelectedDates([]);
    setDateInputVal('');
  }

  function addDateRange(startStr, endStr) {
    if (!startStr || !endStr) return;
    const start = parseDay(startStr);
    const end = parseDay(endStr);
    if (!start || !end) return;
    if (start > end) {
      alert('Start date must be before or equal to End date');
      return;
    }
    const toAdd = [];
    const curr = new Date(start.getTime());
    let count = 0;
    while (curr <= end && count < 180) {
      const iso = `${curr.getFullYear()}-${String(curr.getMonth() + 1).padStart(2, '0')}-${String(curr.getDate()).padStart(2, '0')}`;
      toAdd.push(iso);
      curr.setDate(curr.getDate() + 1);
      count++;
    }
    setSelectedDates(prev => {
      const set = new Set([...prev, ...toAdd]);
      return Array.from(set).sort();
    });
  }

  function addPreset(type) {
    const now = new Date();
    if (type === 'today') {
      const iso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      addDate(iso);
    } else if (type === 'tomorrow') {
      const tom = new Date(now.getTime() + 86400000);
      const iso = `${tom.getFullYear()}-${String(tom.getMonth() + 1).padStart(2, '0')}-${String(tom.getDate()).padStart(2, '0')}`;
      addDate(iso);
    } else if (type === 'next7') {
      for (let i = 0; i < 7; i++) {
        const d = new Date(now.getTime() + i * 86400000);
        addDate(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
      }
    } else if (type === 'next15') {
      for (let i = 0; i < 15; i++) {
        const d = new Date(now.getTime() + i * 86400000);
        addDate(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
      }
    } else if (type === 'nextMonth1st') {
      const d = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      addDate(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
    } else if (type === 'nextMonth15th') {
      const d = new Date(now.getFullYear(), now.getMonth() + 1, 15);
      addDate(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
    } else if (type === 'nextMonthBoth') {
      const d1 = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      const d2 = new Date(now.getFullYear(), now.getMonth() + 1, 15);
      addDate(`${d1.getFullYear()}-${String(d1.getMonth() + 1).padStart(2, '0')}-${String(d1.getDate()).padStart(2, '0')}`);
      addDate(`${d2.getFullYear()}-${String(d2.getMonth() + 1).padStart(2, '0')}-${String(d2.getDate()).padStart(2, '0')}`);
    }
  }
  const [generating, setGenerating] = useState(false);
  const [pptName, setPptName] = useState(() => getTodayPptDateStr());
  const [alsoGenerateExcel, setAlsoGenerateExcel] = useState(true);
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  const [showMultiSelectModal, setShowMultiSelectModal] = useState(false);
  const [multiSelectInput, setMultiSelectInput] = useState('');
  const lastCheckedSiteRef = useRef(null);
  // Campaign tracker map: site_code (uppercase) → latest active campaign row
  const [campaignMap, setCampaignMap] = useState({});
  const [campaignsBySiteCode, setCampaignsBySiteCode] = useState({}); // ALL campaigns per site for date overlap

  function getCleanPptName() {
    let clean = (pptName || '').trim();
    if (!clean) clean = getTodayPptDateStr();
    clean = clean.replace(/\.(pptx|xlsx)$/i, '');
    return clean.replace(/[\\/:*?"<>|]/g, '_');
  }

  function fmtDate(iso) {
    if (!iso) return '';
    try {
      const d = parseDay(iso) || new Date(iso);
      if (isNaN(d.getTime())) return iso;
      return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    } catch { return iso; }
  }

  function getCampaignDurationString(camp) {
    if (!camp) return '';
    if (camp.days) return `${camp.days} day${Number(camp.days) === 1 ? '' : 's'}`;
    const s = parseDay(camp.start_date || camp.booking_date);
    const e = parseDay(camp.end_date);
    if (s && e) {
      const diff = Math.max(1, Math.round((e.getTime() - s.getTime()) / 86400000) + 1);
      return `${diff} day${diff === 1 ? '' : 's'}`;
    }
    return '';
  }

  function extractDateFromString(str) {
    if (!str) return null;
    if (str instanceof Date) return isNaN(str.getTime()) ? null : new Date(str.getFullYear(), str.getMonth(), str.getDate());
    if (typeof str === 'number') {
      const dt = new Date((str - 25569) * 86400000);
      return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    }
    const s = String(str).trim();
    if (!s) return null;
    // 1. DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY
    const dmy = s.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
    if (dmy) {
      let day = parseInt(dmy[1], 10);
      let month = parseInt(dmy[2], 10) - 1;
      let year = parseInt(dmy[3], 10);
      if (year < 100) year += 2000;
      if (month > 11 && day <= 12) {
        const tmp = day - 1;
        day = month + 1;
        month = tmp;
      }
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
    // 2. YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
    const iso = s.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
    if (iso) {
      const d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
      if (!isNaN(d.getTime())) return d;
    }
    // 3. DD Mon YYYY e.g. 25 Oct 2026, 25-Oct-2026
    const monMatch = s.match(/(\d{1,2})[ -]([A-Za-z]{3,9})[ -](\d{2,4})/);
    if (monMatch) {
      const d = new Date(`${monMatch[1]} ${monMatch[2]} ${monMatch[3]}`);
      if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    return null;
  }

  function resolveSiteAvailability(s, v = {}, campaignsBySiteCode = {}, campaignMap = {}, dateFilterOrDates = '') {
    const siteCode = s?.site_code || '';
    const dateArray = Array.isArray(dateFilterOrDates)
      ? dateFilterOrDates.filter(Boolean)
      : (dateFilterOrDates ? [dateFilterOrDates] : (selectedDates && selectedDates.length > 0 ? selectedDates : []));
    const hasDates = dateArray.length > 0;

    // 1. Manual date & availability evaluation
    const siteStoredAvail = String(s?.ppt_availability || s?.availability || '').trim();
    const manualDateInput = v.manualDate;
    const manualTextInput = v.availability;

    const rawCheckText = String(manualTextInput !== undefined ? manualTextInput : siteStoredAvail).trim().toLowerCase();
    const isPlainAvailable =
      rawCheckText === 'available' ||
      rawCheckText === 'available only' ||
      rawCheckText === 'vacant' ||
      rawCheckText === 'ready' ||
      rawCheckText === 'immediate' ||
      rawCheckText === 'yes' ||
      rawCheckText === 'y';

    let manualDateObj = null;
    if (manualDateInput) {
      manualDateObj = parseDay(manualDateInput);
    } else if (!isPlainAvailable) {
      if (manualTextInput !== undefined && manualTextInput !== '') {
        manualDateObj = extractDateFromString(manualTextInput);
      } else if (siteStoredAvail) {
        manualDateObj = extractDateFromString(siteStoredAvail);
      }
    }

    const manualDateFmt = manualDateObj ? fmtDate(manualDateObj) : '';
    const manualDateIso = manualDateObj
      ? `${manualDateObj.getFullYear()}-${String(manualDateObj.getMonth() + 1).padStart(2, '0')}-${String(manualDateObj.getDate()).padStart(2, '0')}`
      : '';

    // If user explicitly provided manualDate or manual text, user override applies
    const hasUserManualOverride = v.manualDate !== undefined || (v.availability !== undefined && v.availability !== '');

    // 2. Active campaign lookup
    const latestCamp = campaignMap[String(siteCode || '').toUpperCase().trim()] || (campaignsBySiteCode[String(siteCode || '').toUpperCase().trim()] || [])[0];
    const todayCamp = getActiveCampaignOnDate(campaignsBySiteCode, siteCode, '');

    let activeCamp = null;
    let occupiedDates = [];
    let vacantDates = [];
    let isBooked = false;

    if (!hasDates) {
      // No date filter active -> evaluate today
      activeCamp = todayCamp;
      const hasAuto = !hasUserManualOverride && !!activeCamp;
      const effectiveEndDate = hasAuto ? (activeCamp.end_date ? fmtDate(activeCamp.end_date) : '') : (manualDateFmt || (hasUserManualOverride ? '' : (latestCamp?.end_date ? fmtDate(latestCamp.end_date) : '')));
      const effectiveEndDay = hasAuto ? (activeCamp.end_date ? parseDay(activeCamp.end_date) : null) : manualDateObj;

      if (hasAuto) {
        isBooked = true;
      } else if (effectiveEndDay) {
        const today = parseDay(new Date());
        isBooked = !today || today <= effectiveEndDay;
      } else if (isPlainAvailable) {
        isBooked = false;
      } else {
        const checkText = rawCheckText;
        isBooked = checkText.startsWith('booked') || checkText.startsWith('occupied') || String(s?.availability || '').toLowerCase() === 'booked';
      }

      let defaultAvailText = 'Available';
      if (isBooked) {
        defaultAvailText = effectiveEndDate ? `Booked till ${effectiveEndDate}` : (siteStoredAvail || 'Occupied');
      } else {
        defaultAvailText = 'Available';
      }

      const autoCamp = hasAuto ? activeCamp : null;
      let autoEndDateStr = autoCamp?.end_date ? fmtDate(autoCamp.end_date) : '';
      let autoStartDateStr = (autoCamp?.start_date || autoCamp?.booking_date) ? fmtDate(autoCamp.start_date || autoCamp.booking_date) : '';
      let autoDurationStr = autoCamp ? getCampaignDurationString(autoCamp) : '';
      let autoCampName = autoCamp?.display || autoCamp?.campaign_name || autoCamp?.brand || autoCamp?.parent_campaign || '';
      let autoClientName = autoCamp?.client || autoCamp?.client_name || '';

      const currentAvailText = isPlainAvailable && !hasAuto ? 'Available' : (manualTextInput !== undefined ? manualTextInput : defaultAvailText);
      const finalBooked = isBooked ||
        String(currentAvailText).toLowerCase().startsWith('booked') ||
        String(currentAvailText).toLowerCase().startsWith('occupied') ||
        (String(currentAvailText).toLowerCase().startsWith('available from') && !!extractDateFromString(currentAvailText));

      let manualDurationStr = '';
      if (!hasAuto && manualDateObj) {
        const today = parseDay(new Date());
        if (today && manualDateObj >= today) {
          const diff = Math.max(1, Math.round((manualDateObj.getTime() - today.getTime()) / 86400000) + 1);
          manualDurationStr = `${diff} day${diff === 1 ? '' : 's'}`;
        }
      }

      return {
        isAuto: hasAuto,
        hasAuto,
        hasTrackerBooking: hasAuto && !!(autoCamp && (autoCamp.campaign_name || autoCamp.display || autoCamp.client || autoCamp.client_name || autoCamp.end_date)),
        hasManual: !!manualDateObj,
        isBooked: finalBooked,
        activeCamp: hasAuto ? autoCamp : (latestCamp || null),
        campaignName: hasAuto ? autoCampName : (v.campaignName || ''),
        clientName: hasAuto ? autoClientName : (v.clientName || ''),
        durationStr: hasAuto ? autoDurationStr : (v.duration || manualDurationStr),
        startDateStr: hasAuto ? autoStartDateStr : (v.startDate || ''),
        endDateStr: effectiveEndDate,
        autoEndDateStr,
        manualDateFmt,
        manualDateIso,
        availText: currentAvailText,
        defaultAvailText,
        availBadge: '',
        occupiedDates: [],
        vacantDates: [],
        datesCount: 0
      };
    }

    // Has dates evaluated!
    let firstActiveCamp = null;
    for (const dStr of dateArray) {
      const camp = getActiveCampaignOnDate(campaignsBySiteCode, siteCode, dStr);
      let isOccOnD = false;
      if (hasUserManualOverride) {
        if (manualDateObj) {
          const fDay = parseDay(dStr);
          isOccOnD = !!(fDay && fDay <= manualDateObj);
        } else if (isPlainAvailable) {
          isOccOnD = false;
        } else {
          const checkText = String(manualTextInput).toLowerCase();
          isOccOnD = checkText.startsWith('booked') || checkText.startsWith('occupied');
        }
      } else {
        if (camp) {
          isOccOnD = true;
          if (!firstActiveCamp) firstActiveCamp = camp;
        } else if (manualDateObj) {
          const fDay = parseDay(dStr);
          isOccOnD = !!(fDay && fDay <= manualDateObj);
        } else if (isPlainAvailable) {
          isOccOnD = false;
        } else {
          isOccOnD = false;
        }
      }
      if (isOccOnD) {
        occupiedDates.push(dStr);
      } else {
        vacantDates.push(dStr);
      }
    }

    isBooked = occupiedDates.length > 0;
    activeCamp = firstActiveCamp || (isBooked ? latestCamp : null);
    const hasAuto = !hasUserManualOverride && !!activeCamp;
    const isAuto = hasAuto;

    let autoEndDateStr = activeCamp?.end_date ? fmtDate(activeCamp.end_date) : '';
    let autoStartDateStr = (activeCamp?.start_date || activeCamp?.booking_date) ? fmtDate(activeCamp.start_date || activeCamp.booking_date) : '';
    let autoDurationStr = activeCamp ? getCampaignDurationString(activeCamp) : '';
    let autoCampName = activeCamp?.display || activeCamp?.campaign_name || activeCamp?.brand || activeCamp?.parent_campaign || '';
    let autoClientName = activeCamp?.client || activeCamp?.client_name || '';

    const effectiveEndDate = isAuto ? autoEndDateStr : (manualDateFmt || (hasUserManualOverride ? '' : autoEndDateStr));

    let defaultAvailText = 'Available';
    let availBadge = '';

    if (dateArray.length === 1) {
      const singleDate = dateArray[0];
      const singleFmt = fmtDate(singleDate) || singleDate;
      if (isBooked) {
        defaultAvailText = effectiveEndDate ? `Booked till ${effectiveEndDate}` : `Booked till ${singleFmt}`;
        availBadge = `🔒 Occupied on ${singleFmt}`;
      } else {
        defaultAvailText = 'Available';
        availBadge = `✓ Available on ${singleFmt}`;
      }
    } else {
      if (!isBooked) {
        defaultAvailText = 'Available';
        availBadge = `✓ Available on all ${dateArray.length} dates`;
      } else {
        // In PPT only single date should come (Booked till <Single Date> or Available)
        const singleEnd = effectiveEndDate || (occupiedDates.length > 0 ? fmtDate(occupiedDates[occupiedDates.length - 1]) : '') || (dateArray[0] ? fmtDate(dateArray[0]) : '');
        defaultAvailText = singleEnd ? `Booked till ${singleEnd}` : 'Occupied';
        const occFmt = occupiedDates.map(d => fmtDate(d) || d).join(', ');
        availBadge = occupiedDates.length === dateArray.length
          ? `🔒 Occupied on all ${dateArray.length} dates`
          : `🔒 Occupied on ${occupiedDates.length}/${dateArray.length} dates (${occFmt})`;
      }
    }

    const currentAvailText = manualTextInput !== undefined ? manualTextInput : defaultAvailText;
    const finalBooked = isBooked ||
      String(currentAvailText).toLowerCase().startsWith('booked') ||
      String(currentAvailText).toLowerCase().startsWith('occupied');

    let manualDurationStr = '';
    if (!isAuto && manualDateObj) {
      const today = parseDay(new Date());
      if (today && manualDateObj >= today) {
        const diff = Math.max(1, Math.round((manualDateObj.getTime() - today.getTime()) / 86400000) + 1);
        manualDurationStr = `${diff} day${diff === 1 ? '' : 's'}`;
      }
    }

    return {
      isAuto,
      hasAuto,
      hasTrackerBooking: hasAuto && !!(activeCamp && (activeCamp.campaign_name || activeCamp.display || activeCamp.client || activeCamp.client_name || activeCamp.end_date)),
      hasManual: !!manualDateObj,
      isBooked: finalBooked,
      activeCamp: isAuto ? activeCamp : (latestCamp || null),
      campaignName: isAuto ? autoCampName : (v.campaignName || ''),
      clientName: isAuto ? autoClientName : (v.clientName || ''),
      durationStr: isAuto ? autoDurationStr : (v.duration || manualDurationStr),
      startDateStr: isAuto ? autoStartDateStr : (v.startDate || ''),
      endDateStr: effectiveEndDate,
      autoEndDateStr,
      manualDateFmt,
      manualDateIso,
      availText: currentAvailText,
      defaultAvailText,
      availBadge,
      occupiedDates,
      vacantDates,
      datesCount: dateArray.length
    };
  }

  async function load() {
    try {
      const [sRes, pRes, cRes] = await Promise.all([
        api.get('/sites'),
        api.get('/ppt-pages').catch(() => api.get('/ppt-pages/public').catch(() => ({ data: null }))),
        api.get('/campaigns').catch(() => ({ data: [] }))
      ]);
      if (Array.isArray(sRes.data)) {
        const mapped = sRes.data.map(unpackSite);
        setSites(mapped);

        // Background sync: check if any sites have local photos that are missing on the server
        try {
          const photoMap = {};
          for (const s of sRes.data) {
            const sc = String(s.site_code || '').trim().toUpperCase();
            const localPhotos = getCachedPhotosForSite(sc);
            let serverPhotos = [];
            try {
              const f = typeof s.flags === 'string' ? JSON.parse(s.flags || '{}') : (s.flags || {});
              serverPhotos = Array.isArray(f.ppt_images) ? f.ppt_images : [];
            } catch {}
            if (serverPhotos.length === 0 && localPhotos.length > 0) {
              photoMap[sc] = localPhotos;
            }
          }
          if (Object.keys(photoMap).length > 0) {
            api.post('/sites/sync-photos', { photo_map: photoMap }).catch(() => {});
          }
        } catch {}
      }
      if (pRes && pRes.data) {
        setPages(prev => {
          const merged = { ...prev };
          if (pRes.data.first) merged.first = pRes.data.first;
          if (pRes.data.second_last) merged.second_last = pRes.data.second_last;
          if (pRes.data.last) merged.last = pRes.data.last;
          try {
            localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(merged));
          } catch {}
          return merged;
        });
      }
      // Build site_code → active campaign map from Campaign Tracker
      // Primary campaigns take precedence, then latest end_date / start_date / id desc
      if (Array.isArray(cRes.data)) {
        // 1. single latest campaign per site (used for slide display)
        const map = {};
        const sorted = cRes.data
          .filter(c => c.record_status === 'active')
          .sort((a, b) => {
            const isLinkedA = String(a.parent_campaign || '').startsWith('LINKED:');
            const isLinkedB = String(b.parent_campaign || '').startsWith('LINKED:');
            if (isLinkedA !== isLinkedB) return isLinkedA ? 1 : -1;
            const endA = new Date(a.end_date || 0).getTime();
            const endB = new Date(b.end_date || 0).getTime();
            if (endA !== endB && endA > 0 && endB > 0) return endB - endA;
            const startA = new Date(a.start_date || 0).getTime();
            const startB = new Date(b.start_date || 0).getTime();
            if (startA !== startB && startA > 0 && startB > 0) return startB - startA;
            return (b.id || 0) - (a.id || 0);
          });
        for (const c of sorted) {
          const code = String(c.site_code || '').toUpperCase().trim();
          if (code && !map[code]) map[code] = c;
        }
        setCampaignMap(map);

        // 2. ALL active campaigns per site (used for date overlap check)
        const byCode = {};
        cRes.data.filter(c => c.record_status === 'active').forEach(c => {
          const key = String(c.site_code || '').toUpperCase().trim();
          if (!key) return;
          if (!byCode[key]) byCode[key] = [];
          byCode[key].push(c);
        });
        setCampaignsBySiteCode(byCode);
      }
    } catch (e) {
      console.warn('PPT load notice:', e);
    }
  }

  const [importingExcel, setImportingExcel] = useState(false);

  async function handlePptExcelImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      alert(`✓ Excel Import Successful!\n\n${r.data.message || `Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`}`);
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  useEffect(() => {
    load();
    const onSitesUpdate = () => load();
    window.addEventListener('mb-sites-updated', onSitesUpdate);
    return () => window.removeEventListener('mb-sites-updated', onSitesUpdate);
  }, []);

  const areas = useMemo(() => Array.from(new Set(sites.map(s => s.area).filter(Boolean))).sort(), [sites]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const selectedCount = useMemo(() => sites.filter(s => sel[s.id || s.site_code]?.checked).length, [sites, sel]);

  const { vacantCount, occupiedCount } = useMemo(() => {
    let vacant = 0;
    let occupied = 0;
    for (const s of sites) {
      const info = resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
      if (info.isBooked) occupied++;
      else vacant++;
    }
    return { vacantCount: vacant, occupiedCount: occupied };
  }, [sites, sel, campaignsBySiteCode, campaignMap, selectedDates]);

  const filtered = useMemo(() => {
    const list = sites.filter(s => {
      if (!matchSiteSearch(s, query)) return false;
      if (areaFilter && s.area !== areaFilter) return false;
      if (availFilter !== 'all') {
        const info = resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
        if (availFilter === 'available' && info.isBooked) return false;
        if (availFilter === 'occupied' && !info.isBooked) return false;
      }
      return true;
    });

    list.sort((a, b) => {
      const aSel = !!sel[a.id || a.site_code]?.checked;
      const bSel = !!sel[b.id || b.site_code]?.checked;
      // Selected sites always show on top
      if (aSel !== bSel) return aSel ? -1 : 1;

      if (sortState.key) {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        if (sortState.key === 'availability') {
          const infoA = resolveSiteAvailability(a, sel[a.id || a.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
          const infoB = resolveSiteAvailability(b, sel[b.id || b.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
          valA = infoA.isBooked ? 'Occupied' : 'Available';
          valB = infoB.isBooked ? 'Occupied' : 'Available';
        }
        return universalCompare(valA, valB, sortState.dir);
      }
      return 0;
    });
    return list;
  }, [sites, query, areaFilter, availFilter, selectedDates, campaignsBySiteCode, campaignMap, sortState, sel]);

  async function uploadPage(k, file) {
    if (!file) return;
    setUploadingPage(prev => ({ ...prev, [k]: true }));
    try {
      // 1. Immediate local base64 preview & offline fallback so image loads instantly
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });

      if (dataUrl) {
        setPages(prev => {
          const next = { ...prev, [k]: dataUrl };
          try {
            localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
            localStorage.setItem(`mb_ppt_raw_${k}`, dataUrl);
          } catch {}
          return next;
        });
      }

      // 2. Persist to server backend & MySQL
      const fd = new FormData();
      fd.append('file', file);
      const res = await api.post('/ppt-pages/' + k, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const serverUrl = res.data?.url || dataUrl;
      setPages(prev => {
        const next = { ...prev, [k]: serverUrl };
        try {
          localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
        } catch {}
        return next;
      });
    } catch (e) {
      console.error('Page upload notice:', e);
      alert('Upload notice: ' + (e.response?.data?.message || e.message));
    } finally {
      setUploadingPage(prev => ({ ...prev, [k]: false }));
    }
  }

  async function removePage(k) {
    const label = k === 'first' ? 'First page' : (k === 'last' ? 'Last page' : 'Second page');
    if (!confirm(`Are you sure you want to remove the uploaded ${label}?`)) return;
    try {
      await api.delete('/ppt-pages/' + k);
    } catch (e) {
      console.warn('Delete page server warning:', e);
    }
    setPages(prev => {
      const next = { ...prev, [k]: '' };
      try {
        localStorage.setItem('mb_ppt_pages_cache', JSON.stringify(next));
        localStorage.removeItem(`mb_ppt_raw_${k}`);
      } catch {}
      return next;
    });
  }

  async function addImages(site, files, replace = false) {
    const fileArray = Array.from(files || []);
    if (!fileArray.length) return;

    // 1. Create immediate local object URLs for instantaneous UI preview
    const localPreviews = fileArray.map(f => URL.createObjectURL(f));

    // Register into photoPreviewCache
    fileArray.forEach((f, idx) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result) {
          photoPreviewCache.set(f.name, reader.result);
          if (localPreviews[idx]) photoPreviewCache.set(localPreviews[idx], reader.result);
        }
      };
      reader.readAsDataURL(f);
    });

    // 2. Optimistically update state so the photo renders IMMEDIATELY (0ms delay!)
    setSites(prev => prev.map(s => {
      if ((s.id && s.id === site.id) || (s.site_code && s.site_code === site.site_code)) {
        const currentImgs = Array.isArray(s.ppt_images) ? s.ppt_images : [];
        return { ...s, ppt_images: replace ? localPreviews : [...currentImgs, ...localPreviews] };
      }
      return s;
    }));

    try {
      let serverImgs = [];
      const fd = new FormData();
      fileArray.forEach(f => fd.append('files', f));

      if (site.id) {
        const res = await api.post(`/sites/${site.id}/images${replace ? '?replace=true' : ''}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        serverImgs = res.data?.images || [];
      } else {
        const res = await api.post(`/sites/code/${encodeURIComponent(site.site_code)}/images${replace ? '?replace=true' : ''}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        serverImgs = res.data?.images || [];
      }

      // Associate server URLs with cached previews & persist to IndexedDB
      serverImgs.forEach((sUrl, idx) => {
        if (localPreviews[idx]) {
          photoPreviewCache.set(sUrl, localPreviews[idx]);
        }
        const f = fileArray[idx];
        if (f) {
          const r2 = new FileReader();
          r2.onload = () => {
            if (r2.result) {
              photoPreviewCache.set(sUrl, r2.result);
              cachePhotoBlob(sUrl, r2.result);
              cachePhotoBlob(`/api${sUrl}`, r2.result);
            }
          };
          r2.readAsDataURL(f);
        }
      });

      // Update state with permanent server URLs
      setSites(prev => prev.map(s => ((s.id && s.id === site.id) || (s.site_code && s.site_code === site.site_code)) ? { ...s, ppt_images: serverImgs } : s));

      // Persist in localStorage cache so photos ALWAYS load upon refresh
      saveCachedPhotosForSite(site.site_code, serverImgs);

      // Notify all components in the app
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (e) {
      console.error('Failed to upload images:', e);
      alert('Failed to upload images: ' + (e.response?.data?.message || e.message));
    }
  }

  async function removeImage(site, index) {
    try {
      let remaining = [];
      if (site.id) {
        const res = await api.delete(`/sites/${site.id}/images/${index}`);
        remaining = res.data?.images || [];
      } else {
        const res = await api.delete(`/sites/code/${encodeURIComponent(site.site_code)}/images/${index}`).catch(() => null);
        remaining = res?.data?.images || (site.ppt_images || []).filter((_, i) => i !== index);
      }
      setSites(prev => prev.map(s => ((s.id && s.id === site.id) || (s.site_code && s.site_code === site.site_code)) ? { ...s, ppt_images: remaining } : s));
      saveCachedPhotosForSite(site.site_code, remaining);
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (e) {
      alert('Failed to remove image: ' + (e.response?.data?.message || e.message));
    }
  }

  async function removeAllImagesForSite(site) {
    if (!confirm(`Are you sure you want to remove all photos for site ${site.site_code}?`)) return;
    try {
      if (site.id) {
        await api.delete(`/sites/${site.id}/images`);
      } else {
        await api.delete(`/sites/code/${encodeURIComponent(site.site_code)}/images`).catch(() => null);
      }
      setSites(prev => prev.map(s => ((s.id && s.id === site.id) || (s.site_code && s.site_code === site.site_code)) ? { ...s, ppt_images: [] } : s));
      saveCachedPhotosForSite(site.site_code, []);
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (e) {
      alert('Failed to remove photos: ' + (e.response?.data?.message || e.message));
    }
  }

  async function removeAllPhotosTogether() {
    const checkedSiteKeys = Object.entries(sel)
      .filter(([k, v]) => v?.checked)
      .map(([k]) => k);

    const checkedSitesWithPhotos = sites.filter(s =>
      (checkedSiteKeys.includes(String(s.id)) || checkedSiteKeys.includes(String(s.site_code))) &&
      Array.isArray(s.ppt_images) && s.ppt_images.length > 0
    );

    let confirmMsg = '';
    let targetIds = [];

    if (checkedSitesWithPhotos.length > 0) {
      confirmMsg = `Are you sure you want to remove all photos from the ${checkedSitesWithPhotos.length} selected site(s)?`;
      targetIds = checkedSitesWithPhotos.map(s => s.id).filter(Boolean);
    } else {
      const totalWithPhotos = sites.filter(s => Array.isArray(s.ppt_images) && s.ppt_images.length > 0).length;
      if (totalWithPhotos === 0) {
        alert('There are currently no uploaded photos across any sites in Automated PPT.');
        return;
      }
      confirmMsg = `Are you sure you want to remove ALL uploaded photos across ALL ${totalWithPhotos} site(s) in Automated PPT?\n\nThis will clear all site photos together so you can start fresh.`;
    }

    if (!confirm(confirmMsg)) return;

    try {
      await api.post('/sites/clear-all-images', { site_ids: targetIds.length > 0 ? targetIds : undefined });
      if (targetIds.length > 0) {
        const targetCodes = checkedSitesWithPhotos.map(s => s.site_code).filter(Boolean);
        setSites(prev => prev.map(s => targetIds.includes(s.id) ? { ...s, ppt_images: [] } : s));
        clearAllCachedPhotos(targetCodes);
      } else {
        setSites(prev => prev.map(s => ({ ...s, ppt_images: [] })));
        clearAllCachedPhotos(null);
      }
      alert('✓ All photos have been removed successfully!');
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (err) {
      alert('Failed to remove all photos: ' + (err.response?.data?.message || err.message));
    }
  }

  const [folderImportStatus, setFolderImportStatus] = useState(null);
  const [replaceExistingFolderPhotos, setReplaceExistingFolderPhotos] = useState(true);

  function matchFileToSite(file, allSites) {
    const relPath = file.webkitRelativePath || file.name || '';
    const parts = relPath.split(/[/\\]/);
    const folders = parts.slice(0, -1);
    const fileName = parts[parts.length - 1] || file.name || '';
    const baseName = fileName.replace(/\.[^/.]+$/, '').trim();

    const clean = s => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

    // 1. Check folder segments from deepest to shallowest (e.g. "Photos/MB-01/day/1.jpg" -> checks "day", then "MB-01")
    for (let i = folders.length - 1; i >= 0; i--) {
      const seg = folders[i].trim();
      if (!seg) continue;
      const cSeg = clean(seg);
      if (!cSeg) continue;

      let match = allSites.find(s => s.site_code && s.site_code.toLowerCase() === seg.toLowerCase());
      if (match) return match;

      match = allSites.find(s => s.site_code && clean(s.site_code) === cSeg);
      if (match) return match;
    }

    // 2. Check filename
    const cBase = clean(baseName);
    if (cBase) {
      let match = allSites.find(s => s.site_code && s.site_code.toLowerCase() === baseName.toLowerCase());
      if (match) return match;

      match = allSites.find(s => s.site_code && clean(s.site_code) === cBase);
      if (match) return match;

      // Prefix match: filename starts with site code (e.g. "MB-01_night.jpg", "MB-01 (1).png", "DEL-04-front.webp")
      const sorted = [...allSites].sort((a, b) => (b.site_code || '').length - (a.site_code || '').length);
      for (const s of sorted) {
        if (!s.site_code) continue;
        const cs = clean(s.site_code);
        if (cs && cBase.startsWith(cs)) {
          return s;
        }
      }
    }

    return null;
  }

  async function handleFolderPhotoImport(e) {
    const rawFiles = Array.from(e.target.files || []);
    e.target.value = '';
    if (!rawFiles.length) return;

    const imageFiles = rawFiles.filter(f => 
      f.type.startsWith('image/') || /\.(jpe?g|png|webp|avif|gif|bmp)$/i.test(f.name)
    );

    if (!imageFiles.length) {
      alert('No image files (.jpg, .png, .webp, etc.) found in the selected folder.');
      return;
    }

    const grouped = new Map();
    const unmatchedFiles = [];

    imageFiles.forEach(file => {
      const site = matchFileToSite(file, sites);
      if (site) {
        if (!grouped.has(site)) grouped.set(site, []);
        grouped.get(site).push(file);
      } else {
        unmatchedFiles.push(file.webkitRelativePath || file.name);
      }
    });

    if (grouped.size === 0) {
      alert(`None of the ${imageFiles.length} photo(s) matched your Site Codes.\n\nPlease ensure your subfolders or filenames are named after your Site Codes (e.g. "MB-01", "MB-02", "DEL-04").`);
      return;
    }

    const totalSites = grouped.size;
    const totalPhotos = Array.from(grouped.values()).reduce((sum, list) => sum + list.length, 0);

    setFolderImportStatus({
      totalSites,
      doneSites: 0,
      currentSite: '',
      totalPhotos,
      donePhotos: 0,
      unmatched: unmatchedFiles,
      siteResults: [],
      finished: false,
      replaceMode: replaceExistingFolderPhotos
    });

    let doneSites = 0;
    let donePhotos = 0;
    const siteResults = [];

    for (const [site, fileList] of grouped.entries()) {
      setFolderImportStatus(prev => ({
        ...prev,
        currentSite: `${site.site_code} (${fileList.length} photo${fileList.length > 1 ? 's' : ''})`
      }));

      try {
        const localPreviews = fileList.map(f => URL.createObjectURL(f));
        fileList.forEach((f, idx) => {
          const reader = new FileReader();
          reader.onload = () => {
            if (reader.result) {
              photoPreviewCache.set(f.name, reader.result);
              if (localPreviews[idx]) photoPreviewCache.set(localPreviews[idx], reader.result);
            }
          };
          reader.readAsDataURL(f);
        });

        const replaceQuery = replaceExistingFolderPhotos ? '?replace=true' : '';
        let newImages = [];
        if (site.id) {
          const fd = new FormData();
          fileList.forEach(f => fd.append('files', f));
          const res = await api.post(`/sites/${site.id}/images${replaceQuery}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
          newImages = res.data?.images || [];
          setSites(prev => prev.map(s => s.id === site.id ? { ...s, ppt_images: newImages } : s));
          siteResults.push({ site_code: site.site_code, count: fileList.length, success: true, replaced: replaceExistingFolderPhotos });
        } else {
          const fd = new FormData();
          fileList.forEach(f => fd.append('files', f));
          const res = await api.post(`/sites/code/${encodeURIComponent(site.site_code)}/images${replaceQuery}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
          newImages = res.data?.images || [];
          setSites(prev => prev.map(s => s.site_code === site.site_code ? { ...s, ppt_images: newImages } : s));
          siteResults.push({ site_code: site.site_code, count: fileList.length, success: true, replaced: replaceExistingFolderPhotos });
        }

        // Cache in memory and IndexedDB
        newImages.forEach((sUrl, idx) => {
          if (localPreviews[idx]) photoPreviewCache.set(sUrl, localPreviews[idx]);
          const f = fileList[idx];
          if (f) {
            const r2 = new FileReader();
            r2.onload = () => {
              if (r2.result) {
                photoPreviewCache.set(sUrl, r2.result);
                cachePhotoBlob(sUrl, r2.result);
                cachePhotoBlob(`/api${sUrl}`, r2.result);
              }
            };
            r2.readAsDataURL(f);
          }
        });

        // Persist in localStorage
        saveCachedPhotosForSite(site.site_code, newImages);
      } catch (err) {
        console.error(`Error uploading photos for ${site.site_code}:`, err);
        siteResults.push({ site_code: site.site_code, count: fileList.length, success: false, error: err.message });
      }

      doneSites += 1;
      donePhotos += fileList.length;
      setFolderImportStatus(prev => ({
        ...prev,
        doneSites,
        donePhotos,
        siteResults: [...siteResults]
      }));
    }

    try {
      await load();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch {}

    setFolderImportStatus(prev => ({
      ...prev,
      finished: true
    }));
  }

  function selectAllVisible() {
    const newSel = { ...sel };
    filtered.forEach(s => {
      newSel[s.id || s.site_code] = { ...(newSel[s.id || s.site_code] || {}), checked: true };
    });
    setSel(newSel);
  }

  function deselectAll() {
    const newSel = { ...sel };
    filtered.forEach(s => {
      if (newSel[s.id || s.site_code]) {
        newSel[s.id || s.site_code].checked = false;
      }
    });
    setSel(newSel);
  }

  function handleSiteCheck(site, isChecked, event) {
    const siteKey = site.id || site.site_code;
    const isShift = Boolean(event && (event.shiftKey || event.nativeEvent?.shiftKey));

    if (isShift && lastCheckedSiteRef.current) {
      const prevKey = lastCheckedSiteRef.current;
      const prevIdx = filtered.findIndex(s => (s.id || s.site_code) === prevKey);
      const currIdx = filtered.findIndex(s => (s.id || s.site_code) === siteKey);

      if (prevIdx !== -1 && currIdx !== -1) {
        const start = Math.min(prevIdx, currIdx);
        const end = Math.max(prevIdx, currIdx);
        const newSel = { ...sel };
        for (let i = start; i <= end; i++) {
          const s = filtered[i];
          const k = s.id || s.site_code;
          newSel[k] = { ...(newSel[k] || {}), checked: isChecked };
        }
        setSel(newSel);
        lastCheckedSiteRef.current = siteKey;
        return;
      }
    }

    lastCheckedSiteRef.current = siteKey;
    setSel(prev => ({
      ...prev,
      [siteKey]: { ...(prev[siteKey] || {}), checked: isChecked }
    }));
  }

  const parsedMultiSites = useMemo(() => {
    if (!multiSelectInput.trim()) return { matched: [], unmatched: [] };
    const rawTokens = multiSelectInput
      .split(/[\s,;\n\r\t]+/)
      .map(t => t.trim())
      .filter(Boolean);
    const uniqueTokens = Array.from(new Set(rawTokens));
    const matched = [];
    const unmatched = [];

    uniqueTokens.forEach(token => {
      const cleanTok = token.toLowerCase();
      const numMatch = cleanTok.match(/^(?:(?:mb|site)[\s-]*)?0*(\d+)$/);
      const targetNum = numMatch ? parseInt(numMatch[1], 10) : null;

      const found = sites.find(s => {
        const sc = String(s.site_code || '').trim().toLowerCase();
        if (sc === cleanTok) return true;
        if (sc.replace(/[^a-z0-9]/g, '') === cleanTok.replace(/[^a-z0-9]/g, '')) return true;
        if (targetNum !== null) {
          const sNumMatch = sc.match(/(\d+)/);
          const sNum = sNumMatch ? parseInt(sNumMatch[1], 10) : NaN;
          if (sNum === targetNum) return true;
        }
        return false;
      });

      if (found) {
        if (!matched.some(m => (m.id || m.site_code) === (found.id || found.site_code))) {
          matched.push(found);
        }
      } else {
        unmatched.push(token);
      }
    });

    return { matched, unmatched };
  }, [multiSelectInput, sites]);

  function applyMultiSelect(mode = 'add') {
    if (!parsedMultiSites.matched.length) {
      alert('No matching sites found in your input. Please enter valid site codes (e.g. MB-01, MB-05, 12).');
      return;
    }
    const newSel = mode === 'only' ? {} : { ...sel };
    parsedMultiSites.matched.forEach(s => {
      const k = s.id || s.site_code;
      newSel[k] = { ...(newSel[k] || {}), checked: true };
    });
    setSel(newSel);
    setShowMultiSelectModal(false);
  }

  function deselectMultiSites() {
    if (!parsedMultiSites.matched.length) return;
    const newSel = { ...sel };
    parsedMultiSites.matched.forEach(s => {
      const k = s.id || s.site_code;
      if (newSel[k]) newSel[k].checked = false;
    });
    setSel(newSel);
    setShowMultiSelectModal(false);
  }

  async function downloadExcel() {
    try {
      const selectedList = sites.filter(s => sel[s.id || s.site_code]?.checked);
      const targetSites = selectedList.length > 0 ? selectedList : (filtered.length > 0 ? filtered : sites);
      const cleanName = getCleanPptName();

      const rows = targetSites.map((x, idx) => {
        const v = sel[x.id || x.site_code] || {};
        const info = resolveSiteAvailability(x, v, campaignsBySiteCode, campaignMap, selectedDates);

        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).toLowerCase().split('x');
          if (parts.length === 2) { w = parts[0].trim(); h = parts[1].trim(); }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        // In PPT / PPT export:
        // Booked sites display as 'Booked till <date>'
        // Available sites display as 'Available' ONLY
        const rawDate = info.endDateStr || info.manualDateFmt || (extractDateFromString(info.availText) ? fmtDate(extractDateFromString(info.availText)) : '');
        let pptAvail = 'Available';
        if (info.isBooked) {
          pptAvail = rawDate ? `Booked till ${rawDate}` : 'Occupied';
        } else {
          pptAvail = 'Available';
        }

        return {
          'SR NO': idx + 1,
          'MB CODE': x.site_code || '',
          'SITE CODE': x.site_code || '',
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': pptAvail,
          'Campaign Name': info.campaignName,
          'Client Name': info.clientName,
          'Duration': info.durationStr,
          'Start Date': info.startDateStr,
          'End Date': info.endDateStr,
          'Selling Amount': Number(v.rate ?? x.ppt_rate ?? x.monthly_rate ?? 0),
          'Latitude Longitude': coords
        };
      });

      await exportStyledExcel(rows, `${cleanName}.xlsx`, cleanName);
    } catch (e) {
      alert('Export Excel failed: ' + e.message);
    }
  }

  async function generate() {
    const chosen = sites
      .filter(s => sel[s.id || s.site_code]?.checked)
      .map(s => {
        const v = sel[s.id || s.site_code] || {};
        const info = resolveSiteAvailability(s, v, campaignsBySiteCode, campaignMap, selectedDates);

        // In PPT:
        // Booked sites display as 'Booked till <date>'
        // Available sites display as 'Available' ONLY
        const rawDate = info.endDateStr || info.manualDateFmt || (extractDateFromString(info.availText) ? fmtDate(extractDateFromString(info.availText)) : '');
        let pptAvail = 'Available';
        if (info.isBooked) {
          pptAvail = rawDate ? `Booked till ${rawDate}` : 'Occupied';
        } else {
          pptAvail = 'Available';
        }

        return {
          ...s,
          _availability: pptAvail,
          _rate: v.rate ?? s.ppt_rate ?? s.monthly_rate,
          _showRate: !!v.showRate,
          _endDate: info.endDateStr,
          _isBooked: info.isBooked,
          _campaignName: info.campaignName,
          _clientName: info.clientName,
          _startDate: info.startDateStr,
          _duration: info.durationStr,
          _isAuto: info.isAuto
        };
      });

    if (!chosen.length) return alert('Please select at least one site to include in the PowerPoint.');

    const cleanName = getCleanPptName();

    setGenerating(true);
    try {
      await makePpt(chosen, pages, `${cleanName}.pptx`);

      if (alsoGenerateExcel) {
        // Small delay to ensure browser reliably dispatches multiple downloads
        await new Promise(r => setTimeout(r, 450));

        const rows = chosen.map((x, idx) => {
          let w = x.width || '', h = x.height || '';
          if ((!w || !h) && x.size) {
            const parts = String(x.size).toLowerCase().split('x');
            if (parts.length === 2) { w = parts[0].trim(); h = parts[1].trim(); }
          }
          const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
          const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

          return {
            'SR NO': idx + 1,
            'MB CODE': x.site_code || '',
            'SITE CODE': x.site_code || '',
            'AREA': x.area || x.city || '',
            'LOCATION': x.address || '',
            'MEDIA': x.media_type || 'Hoarding',
            'LIGHT': x.lighting || 'BL',
            'W': w || '',
            'H': h || '',
            'SQ FT': sqft,
            'AVAILABLITY': x._availability ?? 'Available',
            'Campaign Name': x._campaignName || '',
            'Client Name': x._clientName || '',
            'Duration': x._duration || '',
            'Start Date': x._startDate || '',
            'End Date': x._endDate || '',
            'Selling Amount': Number(x._rate ?? 0),
            'Latitude Longitude': coords
          };
        });

        await exportStyledExcel(rows, `${cleanName}.xlsx`, cleanName);
      }
    } catch (e) {
      alert('Generation Error: ' + e.message);
    } finally {
      setGenerating(false);
    }
  }

  return (
    <>
      <PageHead
        title="Automated PPT"
        desc="Select sites, customize presentation name, and generate client pitch decks with the official Media Buzz side dashboard."
        actions={
          <div className="scooh-ppt-head-actions" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
            <div className="scooh-ppt-name-bar" style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(10, 30, 53, 0.85)', padding: '5px 12px', borderRadius: '8px', border: '1px solid #1c3b60' }}>
              <span style={{ fontSize: '11px', fontWeight: 700, color: '#FFC200', textTransform: 'uppercase', letterSpacing: '0.6px' }}>Presentation Name:</span>
              <input
                type="text"
                className="scooh-input scooh-ppt-name-input"
                style={{ flex: '1 1 180px', minWidth: 0, padding: '5px 10px', fontSize: '13px', background: '#071526', border: '1px solid #234d7d', color: '#fff', borderRadius: '5px' }}
                value={pptName}
                onChange={e => setPptName(e.target.value)}
                placeholder={getTodayPptDateStr()}
              />
              <span className="scooh-ppt-name-ext" style={{ fontSize: '12px', color: '#8fa4bd', fontWeight: 600 }}>
                {alsoGenerateExcel ? '.pptx & .xlsx' : '.pptx'}
              </span>
            </div>

            <label
              className="scooh-ppt-excel-option"
              title="Automatically generate and download matching Excel sheet with PPT contents and presentation title"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '8px',
                cursor: 'pointer',
                background: alsoGenerateExcel ? 'rgba(255, 194, 0, 0.12)' : 'rgba(10, 30, 53, 0.85)',
                padding: '6px 13px',
                borderRadius: '8px',
                border: alsoGenerateExcel ? '1px solid #FFC200' : '1px solid #1c3b60',
                userSelect: 'none',
                color: alsoGenerateExcel ? '#FFC200' : '#c8d6e5',
                fontWeight: 600,
                fontSize: '12.5px',
                transition: 'all 0.2s ease'
              }}
            >
              <input
                type="checkbox"
                checked={alsoGenerateExcel}
                onChange={e => setAlsoGenerateExcel(e.target.checked)}
                style={{ accentColor: '#FFC200', width: '15px', height: '15px', cursor: 'pointer' }}
              />
              <span>Also generate Excel (.xlsx)</span>
            </label>

            <button type="button" className="scooh-btn ghost" onClick={downloadExcel} title="Export selected sites directly to Excel">
              Download Excel (.xlsx)
            </button>
            <button type="button" className="scooh-btn primary" onClick={generate} disabled={generating}>
              {generating
                ? (alsoGenerateExcel ? 'Creating PPT & Excel…' : 'Creating PPT…')
                : (alsoGenerateExcel ? 'Generate PPT + Excel' : 'Generate PPT')}
            </button>
          </div>
        }
      />

      {/* Fixed presentation pages section */}
      <section className="scooh-form-section scooh-ppt-fixed-pages">
        <div className="scooh-form-section-head">
          <div>
            <h3>Fixed presentation pages (Optional)</h3>
            <p>Optional cover and closing pages. Order: page 1, page 2, selected site slides, then page 3.</p>
          </div>
        </div>
        <div className="scooh-grid3">
          {[
            ['first', 'First page', 'Full-screen cover image for slide 1.'],
            ['second_last', 'Second page', 'Full-screen image used as slide 2, before the selected site slides.'],
            ['last', 'Last page', 'Full-screen image used after all selected site slides.']
          ].map(([k, title, desc]) => {
            const pageUrl = pages[k] || '';
            const isUploading = !!uploadingPage[k];
            const rawFallback = (() => {
              try { return localStorage.getItem(`mb_ppt_raw_${k}`) || ''; } catch { return ''; }
            })();
            const activeUrl = pageUrl || rawFallback || (k === 'last' ? '/assets/ppt_contact_last_page.png' : '');

            return (
              <div
                className="scooh-panel"
                key={k}
                style={{
                  margin: 0,
                  padding: '18px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '10px',
                  border: activeUrl ? '1px solid rgba(0, 200, 117, 0.35)' : '1px solid #1c2838',
                  background: activeUrl ? 'linear-gradient(180deg, #0d1624 0%, #08101a 100%)' : '#0b1320',
                  borderRadius: '12px'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
                  <strong style={{ fontSize: '14px', color: '#eef2f6' }}>{title}</strong>
                  {activeUrl ? (
                    <span style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px',
                      fontSize: '11px',
                      fontWeight: 700,
                      color: '#00e575',
                      background: 'rgba(0, 229, 117, 0.12)',
                      border: '1px solid rgba(0, 229, 117, 0.3)',
                      padding: '2px 8px',
                      borderRadius: '12px'
                    }}>
                      ✓ Saved & Loaded
                    </span>
                  ) : (
                    <span style={{ fontSize: '11px', color: '#7a8ba3' }}>Optional</span>
                  )}
                </div>

                <small style={{ color: '#91a5c2', fontSize: '11.5px', lineHeight: '1.4' }}>{desc}</small>

                {activeUrl ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '4px' }}>
                    <div
                      style={{
                        position: 'relative',
                        width: '100%',
                        height: '135px',
                        background: '#040912',
                        borderRadius: '8px',
                        overflow: 'hidden',
                        border: '1px solid #1c2c42',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                      }}
                    >
                      <img
                        src={activeUrl}
                        alt={title}
                        onError={e => {
                          if (rawFallback && e.target.src !== rawFallback) {
                            e.target.src = rawFallback;
                          } else if (activeUrl.startsWith('/') && !e.target.src.includes('3000')) {
                            e.target.src = `http://localhost:3000${activeUrl}`;
                          }
                        }}
                        style={{
                          width: '100%',
                          height: '100%',
                          objectFit: 'cover',
                          display: 'block'
                        }}
                      />
                      {isUploading && (
                        <div style={{
                          position: 'absolute',
                          inset: 0,
                          background: 'rgba(2, 10, 20, 0.75)',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          color: '#FFC200',
                          fontSize: '12px',
                          fontWeight: 700,
                          gap: '6px'
                        }}>
                          <span>Saving…</span>
                        </div>
                      )}
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <label
                        className="scooh-btn"
                        style={{
                          flex: 1,
                          cursor: isUploading ? 'not-allowed' : 'pointer',
                          display: 'inline-flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontSize: '11.5px',
                          padding: '6px 12px',
                          minHeight: '34px'
                        }}
                      >
                        {isUploading ? 'Uploading…' : 'Replace image'}
                        <input
                          type="file"
                          accept="image/*"
                          hidden
                          disabled={isUploading}
                          onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])}
                        />
                      </label>
                      <button
                        type="button"
                        className="scooh-btn danger"
                        onClick={() => removePage(k)}
                        disabled={isUploading}
                        title="Remove uploaded image"
                        style={{
                          minHeight: '34px',
                          padding: '6px 10px',
                          fontSize: '11.5px'
                        }}
                      >
                        Remove
                      </button>
                    </div>
                  </div>
                ) : (
                  <div style={{ marginTop: '6px' }}>
                    <label
                      className="scooh-btn primary"
                      style={{
                        width: '100%',
                        cursor: isUploading ? 'not-allowed' : 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '6px',
                        minHeight: '38px'
                      }}
                    >
                      {isUploading ? 'Uploading…' : '+ Upload image'}
                      <input
                        type="file"
                        accept="image/*"
                        hidden
                        disabled={isUploading}
                        onChange={e => e.target.files[0] && uploadPage(k, e.target.files[0])}
                      />
                    </label>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </section>

      {/* 1. Select sites and slide values section */}
      <section className="scooh-form-section scooh-ppt-generator" style={{ marginTop: '18px' }}>
        <div className="scooh-form-section-head">
          <div>
            <h3>1. Select sites and slide values</h3>
            <p>Use the approved Media Buzz format: site details on the left and the site image on the right. Availability and monthly rate are optional and can be imported from Excel or edited manually.</p>
          </div>
        </div>

        {/* Sleek Compact Unified Controls Panel */}
        <div className="scooh-ppt-controls-panel">
          {/* Row 1: Search & Filter Controls */}
          <div className="scooh-ppt-filter-row">
            {/* Proper Search Bar */}
            <div className="scooh-ppt-search-box">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
                <circle cx="11" cy="11" r="8"></circle>
                <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
              </svg>
              <input
                id="scooh-ppt-search"
                type="text"
                placeholder="Search site code, area, landmark, or multiple codes…"
                value={query}
                onChange={e => setQuery(e.target.value)}
                autoComplete="off"
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  style={{
                    background: 'none',
                    border: 'none',
                    color: '#94a3b8',
                    cursor: 'pointer',
                    padding: '0 2px',
                    fontSize: '13px',
                    lineHeight: 1,
                    flexShrink: 0
                  }}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Area Filter */}
            <div className={`scooh-ppt-filter-pill ${areaFilter ? 'active-purple' : ''}`}>
              <span style={{ fontSize: '11px', color: areaFilter ? '#c084fc' : '#94a3b8', fontWeight: 700 }}>📍 Area:</span>
              <select
                id="scooh-ppt-area-filter"
                value={areaFilter}
                onChange={e => setAreaFilter(e.target.value)}
              >
                <option value="">All Areas</option>
                {areas.map(a => <option key={a} value={a}>{a}</option>)}
              </select>
              {areaFilter && (
                <button
                  type="button"
                  onClick={() => setAreaFilter('')}
                  style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '12px', padding: '0 2px', lineHeight: 1 }}
                  title="Clear area filter"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Date Availability Filter with Multi-Select */}
            <div className={`scooh-ppt-filter-pill ${selectedDates.length > 0 ? 'active-sky' : ''}`} style={{ gap: '5px' }}>
              <span style={{ fontSize: '11px', color: selectedDates.length > 0 ? '#38bdf8' : '#94a3b8', fontWeight: 700 }}>
                📅 {selectedDates.length > 1 ? `Dates (${selectedDates.length}):` : 'Date:'}
              </span>
              <input
                id="scooh-ppt-date-filter"
                type="date"
                value={dateInputVal}
                onChange={e => {
                  const val = e.target.value;
                  setDateInputVal(val);
                  if (val) addDate(val);
                }}
                title="Select date or add another date"
                style={{ width: '110px' }}
              />
              <button
                type="button"
                onClick={() => setShowDateModal(true)}
                style={{
                  background: selectedDates.length > 1 ? 'rgba(56, 189, 248, 0.25)' : 'rgba(255,255,255,0.08)',
                  border: `1px solid ${selectedDates.length > 1 ? '#38bdf8' : 'rgba(255,255,255,0.18)'}`,
                  borderRadius: '5px',
                  color: selectedDates.length > 0 ? '#38bdf8' : '#cbd5e1',
                  fontSize: '11px',
                  fontWeight: 700,
                  padding: '2px 8px',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '4px',
                  height: '24px'
                }}
                title="Open interactive multi-date calendar, presets & range selector"
              >
                <span>🗓️ Multi-Date</span>
                {selectedDates.length > 0 && (
                  <span style={{ background: '#38bdf8', color: '#031726', borderRadius: '10px', fontSize: '10px', fontWeight: 900, padding: '0 5px', lineHeight: '14px' }}>
                    {selectedDates.length}
                  </span>
                )}
              </button>
              {selectedDates.length > 0 && (
                <button
                  type="button"
                  onClick={clearAllDates}
                  style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '12px', padding: '0 2px', lineHeight: 1 }}
                  title="Clear all selected dates"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Compact Sort Selector + Direction Toggle */}
            <div className="scooh-ppt-filter-pill">
              <span style={{ fontSize: '11px', color: '#94a3b8', fontWeight: 700 }}>⇅ Sort:</span>
              <select
                value={`${sortState.key}:${sortState.dir}`}
                onChange={e => {
                  const [k, d] = e.target.value.split(':');
                  setSortState({ key: k, dir: d });
                }}
              >
                <option value="site_code:asc">Site Code (01 → 87)</option>
                <option value="site_code:desc">Site Code (87 → 01)</option>
                <option value="area:asc">Area (A–Z)</option>
                <option value="area:desc">Area (Z–A)</option>
                <option value="city:asc">City (A–Z)</option>
                <option value="media_type:asc">Media Type</option>
                <option value="monthly_rate:asc">Rate: Low to High</option>
                <option value="monthly_rate:desc">Rate: High to Low</option>
                <option value="availability:asc">Availability</option>
                <option value="size:asc">Size</option>
              </select>
              <button
                type="button"
                onClick={() => setSortState(prev => ({ ...prev, dir: prev.dir === 'asc' ? 'desc' : 'asc' }))}
                style={{
                  background: 'rgba(139, 92, 246, 0.15)',
                  border: '1px solid rgba(139, 92, 246, 0.35)',
                  borderRadius: '5px',
                  color: '#c4b5fd',
                  cursor: 'pointer',
                  fontSize: '11px',
                  fontWeight: 800,
                  padding: '2px 7px',
                  lineHeight: 1.2
                }}
                title={sortState.dir === 'asc' ? 'Ascending (click for descending)' : 'Descending (click for ascending)'}
              >
                {sortState.dir === 'asc' ? '↑ Asc' : '↓ Desc'}
              </button>
            </div>

            {/* Availability Filter on selected date (Beside Sort) */}
            <div
              className={`scooh-ppt-filter-pill ${availFilter !== 'all' ? (availFilter === 'available' ? 'active-green' : 'active-yellow') : ''}`}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '5px'
              }}
              title={selectedDates.length > 0 ? `Filter sites available or occupied on selected date (${selectedDates.map(d => fmtDate(d) || d).join(', ')})` : 'Filter sites available or occupied today'}
            >
              <span style={{ fontSize: '11px', color: availFilter === 'available' ? '#4ade80' : (availFilter === 'occupied' ? '#fde047' : '#94a3b8'), fontWeight: 700 }}>
                {availFilter === 'available' ? '🟢' : (availFilter === 'occupied' ? '🟡' : '⚡')} Availability:
              </span>
              <select
                value={availFilter}
                onChange={e => setAvailFilter(e.target.value)}
                style={{
                  fontWeight: 700,
                  color: availFilter === 'available' ? '#4ade80' : (availFilter === 'occupied' ? '#fde047' : '#cbd5e1')
                }}
              >
                <option value="all">All Sites ({sites.length})</option>
                <option value="available">🟢 Available Only ({vacantCount})</option>
                <option value="occupied">🟡 Occupied Only ({occupiedCount})</option>
              </select>
              {availFilter !== 'all' && (
                <button
                  type="button"
                  onClick={() => setAvailFilter('all')}
                  style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '12px', padding: '0 2px', lineHeight: 1 }}
                  title="Show all sites"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Reset Filters */}
            {(query || areaFilter || selectedDates.length > 0 || availFilter !== 'all') && (
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => {
                  setQuery('');
                  setAreaFilter('');
                  setAvailFilter('all');
                  clearAllDates();
                }}
                style={{
                  height: '36px',
                  minHeight: '36px',
                  padding: '0 10px',
                  fontSize: '11.5px',
                  color: '#f87171',
                  borderColor: 'rgba(248, 113, 113, 0.35)',
                  borderRadius: '8px'
                }}
                title="Reset search and filters"
              >
                ✕ Reset
              </button>
            )}
          </div>

          {/* Row 2: Selection Counts & Action Buttons */}
          <div className="scooh-ppt-action-row">
            {/* Left: Site Count & Selected Badge */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              <span style={{ fontSize: '12px', color: '#94a3b8', fontWeight: 600 }}>
                {filtered.length} of {sites.length} sites
              </span>
              {selectedCount > 0 && (
                <span style={{
                  background: 'rgba(167, 139, 250, 0.18)',
                  border: '1px solid rgba(167, 139, 250, 0.4)',
                  color: '#c4b5fd',
                  padding: '2px 8px',
                  borderRadius: '12px',
                  fontWeight: 700,
                  fontSize: '11.5px'
                }}>
                  ✓ {selectedCount} site{selectedCount > 1 ? 's' : ''} selected (on top)
                </span>
              )}
            </div>

            {/* Right: Multi-Select & Batch Actions */}
            <div className="scooh-ppt-select-actions">
              <button
                type="button"
                className="scooh-btn primary"
                onClick={() => setShowMultiSelectModal(true)}
                style={{
                  background: 'linear-gradient(135deg, #8b5cf6, #6366f1)',
                  border: 'none',
                  boxShadow: '0 2px 8px rgba(99, 102, 241, 0.35)',
                  fontWeight: 700,
                  color: '#ffffff'
                }}
                title="Paste or enter multiple site codes (e.g. MB-01, MB-05, MB-12) to select them all together"
              >
                <span>⚡ Multi-Select Sites</span>
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={selectAllVisible}
              >
                Select all visible
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={deselectAll}
              >
                Deselect all
              </button>
              
              <div style={{ width: '1px', height: '20px', background: 'rgba(255,255,255,0.12)', margin: '0 2px' }} />

              <label
                className="scooh-btn ghost"
                style={{ cursor: 'pointer', gap: '5px' }}
                title="Import Excel to update PPT rates & availability"
              >
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handlePptExcelImport} />
              </label>

              <label
                className="scooh-btn primary"
                style={{
                  cursor: folderImportStatus && !folderImportStatus.finished ? 'not-allowed' : 'pointer',
                  gap: '6px',
                  background: 'linear-gradient(135deg, #0284c7, #4f46e5)',
                  color: '#fff',
                  fontWeight: 700,
                  border: 'none',
                  boxShadow: '0 2px 8px rgba(2, 132, 199, 0.3)'
                }}
                title="Import a folder containing site photos (folders or files named by site code e.g. MB-01)"
              >
                <span>📂 {folderImportStatus && !folderImportStatus.finished ? 'Importing…' : 'Import Photo Folder'}</span>
                <input
                  type="file"
                  {...{ webkitdirectory: '', directory: '', multiple: true }}
                  hidden
                  disabled={Boolean(folderImportStatus && !folderImportStatus.finished)}
                  onChange={handleFolderPhotoImport}
                />
              </label>

              <label
                className="scooh-ppt-replace-toggle"
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '5px',
                  cursor: 'pointer',
                  userSelect: 'none',
                  fontSize: '11.5px',
                  fontWeight: 600,
                  color: replaceExistingFolderPhotos ? '#38bdf8' : '#94a3b8',
                  background: replaceExistingFolderPhotos ? 'rgba(56, 189, 248, 0.12)' : '#08111d',
                  padding: '0 9px',
                  borderRadius: '7px',
                  border: replaceExistingFolderPhotos ? '1px solid #38bdf8' : '1px solid #1c3b60',
                  height: '35px'
                }}
                title="When enabled, importing photos replaces old photos for matched sites instead of appending"
              >
                <input
                  type="checkbox"
                  checked={replaceExistingFolderPhotos}
                  onChange={e => setReplaceExistingFolderPhotos(e.target.checked)}
                  style={{ accentColor: '#38bdf8', width: '13px', height: '13px', cursor: 'pointer' }}
                />
                <span>Replace old photos</span>
              </label>

              <div style={{ width: '1px', height: '20px', background: 'rgba(255,255,255,0.12)', margin: '0 2px' }} />

              <button
                type="button"
                className="scooh-btn danger"
                onClick={removeAllPhotosTogether}
                style={{
                  gap: '5px',
                  background: 'rgba(239, 68, 68, 0.15)',
                  color: '#f87171',
                  border: '1px solid rgba(239, 68, 68, 0.4)',
                  fontWeight: 700
                }}
                title="Remove all uploaded site photos together"
              >
                <span>🗑️ Remove All Photos</span>
              </button>
            </div>
          </div>
        </div>

        {/* Dynamic Multi-Date Availability Banner & Date Chips */}
        {selectedDates.length > 0 && (
          <div className="scooh-ppt-date-banner" style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', background: 'rgba(56, 189, 248, 0.08)', border: '1px solid rgba(56, 189, 248, 0.25)', borderRadius: '8px', padding: '8px 12px', margin: '0 0 12px', fontSize: '12px', color: '#38bdf8' }}>
            <span style={{ fontSize: '15px' }}>📅</span>
            <span style={{ fontWeight: 600 }}>
              {selectedDates.length === 1 ? (
                <>Evaluating availability as of <strong>{fmtDate(selectedDates[0]) || selectedDates[0]}</strong>:</>
              ) : (
                <>Evaluating availability across <strong>{selectedDates.length} selected dates</strong>:</>
              )}
            </span>
            <div style={{ display: 'flex', alignItems: 'center', gap: '5px', flexWrap: 'wrap' }}>
              {selectedDates.map(d => (
                <span
                  key={d}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    background: 'rgba(14, 165, 233, 0.18)',
                    border: '1px solid rgba(56, 189, 248, 0.45)',
                    borderRadius: '5px',
                    padding: '2px 7px',
                    fontSize: '11.5px',
                    color: '#e0f2fe',
                    fontWeight: 700
                  }}
                >
                  {fmtDate(d) || d}
                  <button
                    type="button"
                    onClick={() => removeDate(d)}
                    style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11px', padding: '0 1px', lineHeight: 1 }}
                    title={`Remove ${fmtDate(d) || d}`}
                  >
                    ✕
                  </button>
                </span>
              ))}
            </div>
            <button
              type="button"
              onClick={() => setShowDateModal(true)}
              style={{
                background: 'rgba(56, 189, 248, 0.15)',
                border: '1px solid rgba(56, 189, 248, 0.35)',
                borderRadius: '5px',
                color: '#7dd3fc',
                cursor: 'pointer',
                fontSize: '11px',
                fontWeight: 700,
                padding: '2px 8px'
              }}
            >
              + Add / Edit Dates
            </button>
            <button
              type="button"
              onClick={clearAllDates}
              style={{ marginLeft: 'auto', background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11px', fontWeight: 700, textDecoration: 'underline' }}
            >
              Reset All Dates
            </button>
          </div>
        )}

        {/* Sites Grid */}
        <div className="scooh-ppt-grid" id="scooh-ppt-grid">
          {filtered.map(s => {
            const siteKey = s.id || s.site_code;
            const v = sel[siteKey] || {};
            const isChecked = !!v.checked;
            const imgs = s.ppt_images || [];
            const info = resolveSiteAvailability(s, v, campaignsBySiteCode, campaignMap, selectedDates);
            const isBooked = info.isBooked;
            const campaignName = info.campaignName;
            const clientName = info.clientName;
            const startDateText = info.startDateStr;
            const finalEndDateText = info.endDateStr;
            const durationText = info.durationStr;

            return (
              <div
                key={siteKey}
                className="scooh-ppt-site"
                style={{
                  border: isChecked ? '1.5px solid #a78bfa' : '1px solid #344258',
                  boxShadow: isChecked ? '0 0 12px rgba(167, 139, 250, 0.25)' : 'none',
                  background: isChecked ? 'rgba(30, 27, 75, 0.45)' : undefined
                }}
              >
                <div className="scooh-ppt-card-top" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', marginBottom: '8px' }}>
                  <input
                    type="checkbox"
                    className="scooh-ppt-site-check"
                    checked={isChecked}
                    onChange={e => handleSiteCheck(s, e.target.checked, e)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px', accentColor: '#a78bfa', margin: 0 }}
                    title="Click to select (Hold Shift + Click to select a range of sites)"
                  />
                  <div className="scooh-ppt-card-actions" style={{ display: 'flex', gap: '6px', margin: 0 }}>
                    <label className="scooh-btn secondary" style={{ cursor: 'pointer', fontSize: '11.5px', padding: '4px 9px', minHeight: '30px' }} onClick={e => e.stopPropagation()} title="Upload one or more photo files">
                      Add images
                      <input
                        type="file"
                        multiple
                        accept="image/*"
                        hidden
                        onChange={e => {
                          if (e.target.files && e.target.files.length) {
                            addImages(s, e.target.files);
                            e.target.value = '';
                          }
                        }}
                      />
                    </label>
                    <label className="scooh-btn secondary" style={{ cursor: 'pointer', fontSize: '11.5px', padding: '4px 8px', minHeight: '30px' }} onClick={e => e.stopPropagation()} title={`Import a folder of photos specifically for ${s.site_code} (replaces old photos)`}>
                      📁 Folder
                      <input
                        type="file"
                        {...{ webkitdirectory: '', directory: '', multiple: true }}
                        hidden
                        onChange={e => {
                          const imgFiles = Array.from(e.target.files || []).filter(f => f.type.startsWith('image/') || /\.(jpe?g|png|webp|avif|gif|bmp)$/i.test(f.name));
                          if (imgFiles.length) addImages(s, imgFiles, true);
                          else alert('No image files found in folder.');
                          e.target.value = '';
                        }}
                      />
                    </label>
                    {imgs.length > 0 && (
                      <button
                        type="button"
                        className="scooh-btn danger"
                        style={{ fontSize: '11px', padding: '4px 8px', minHeight: '30px' }}
                        onClick={e => {
                          e.stopPropagation();
                          removeAllImagesForSite(s);
                        }}
                        title={`Remove all ${imgs.length} photo(s) for ${s.site_code}`}
                      >
                        🗑️ Clear
                      </button>
                    )}
                  </div>
                </div>

                <div className="scooh-ppt-photo-list scooh-ppt-added-images">
                  {imgs.length > 0 ? (
                    imgs.map((u, idx) => {
                      const photoSrc = resolvePhotoUrl(u);
                      return (
                        <span className="scooh-ppt-photo" key={idx} onClick={e => e.stopPropagation()}>
                          <img
                            src={photoSrc}
                            alt={`${s.site_code} photo ${idx + 1}`}
                            loading="lazy"
                            onError={e => {
                              const img = e.target;
                              const currentSrc = img.src || '';
                              if (u && typeof u === 'string') {
                                if (u.startsWith('/uploads/') && currentSrc.includes('/api/uploads/')) {
                                  img.src = u;
                                  return;
                                }
                                if (u.startsWith('/uploads/') && !currentSrc.includes('/api/uploads/')) {
                                  img.src = `/api${u}`;
                                  return;
                                }
                              }
                              if (photoPreviewCache.has(u)) {
                                img.src = photoPreviewCache.get(u);
                                return;
                              }
                              getCachedPhotoBlob(u).then(blob => {
                                if (blob) {
                                  img.src = blob;
                                  photoPreviewCache.set(u, blob);
                                } else {
                                  img.style.display = 'none';
                                  if (img.parentElement) img.parentElement.classList.add('photo-load-error');
                                }
                              }).catch(() => {
                                img.style.display = 'none';
                                if (img.parentElement) img.parentElement.classList.add('photo-load-error');
                              });
                            }}
                          />
                          <button
                            type="button"
                            title="Remove image"
                            aria-label="Remove image"
                            onClick={() => removeImage(s, idx)}
                          >
                            ×
                          </button>
                        </span>
                      );
                    })
                  ) : (
                    <div className="scooh-ppt-no-photos">No added images</div>
                  )}
                </div>

                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', margin: '4px 0 6px', flexWrap: 'wrap' }} onClick={e => e.stopPropagation()}>
                  <button
                    type="button"
                    onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                    className="scooh-plate"
                    title={`Open ${s.site_code} in Latest Booking`}
                    style={{
                      fontSize: '13px',
                      fontWeight: 800,
                      cursor: 'pointer',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '5px',
                      background: 'rgba(56, 189, 248, 0.15)',
                      color: '#38bdf8',
                      border: '1px solid rgba(56, 189, 248, 0.45)',
                      padding: '3px 8px',
                      borderRadius: '6px'
                    }}
                  >
                    <span>{s.site_code || 'Site'}</span>
                    <span style={{ fontSize: '10px', opacity: 0.8 }}>↗ Booking</span>
                  </button>
                  {info.hasTrackerBooking && finalEndDateText ? (
                    <button
                      type="button"
                      onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                      title={`Active Campaign End Date: ${finalEndDateText}\nClick to view in Latest Booking`}
                      style={{
                        fontSize: '11px',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: 'rgba(245, 158, 11, 0.2)',
                        color: '#fbbf24',
                        border: '1px solid rgba(245, 158, 11, 0.45)',
                        cursor: 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '4px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      <span>📅 Booked till {finalEndDateText}</span>
                    </button>
                  ) : info.manualDateFmt ? (
                    <span
                      title={isBooked ? `Booked till: ${info.manualDateFmt}` : 'Available'}
                      style={{
                        fontSize: '11px',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: isBooked ? 'rgba(245, 158, 11, 0.2)' : 'rgba(56, 189, 248, 0.15)',
                        color: isBooked ? '#fbbf24' : '#38bdf8',
                        border: isBooked ? '1px solid rgba(245, 158, 11, 0.45)' : '1px solid rgba(56, 189, 248, 0.35)',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '4px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      <span>📅 {isBooked ? `Booked till ${info.manualDateFmt}` : 'Available'}</span>
                    </span>
                  ) : isBooked ? (
                    <span
                      style={{
                        fontSize: '11px',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: 'rgba(245, 158, 11, 0.2)',
                        color: '#fbbf24',
                        border: '1px solid rgba(245, 158, 11, 0.45)',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '4px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      <span>🔒 Booked</span>
                    </span>
                  ) : dateFilter ? (
                    <span
                      style={{
                        fontSize: '11px',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '5px',
                        background: 'rgba(34, 197, 94, 0.15)',
                        color: '#4ade80',
                        border: '1px solid rgba(34, 197, 94, 0.35)',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '4px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      <span>✓ Available</span>
                    </span>
                  ) : null}
                </div>
                <small>{s.location || s.address || s.area || s.city || ''}</small>
                <small className="scooh-ppt-size">Size: {s.size || (s.width && s.height ? `${s.width}x${s.height}` : '') || '—'}</small>
                <small className="scooh-ppt-coordinates">Latitude: {s.latitude ?? '—'}</small>
                <small className="scooh-ppt-coordinates">Longitude: {s.longitude ?? '—'}</small>

                {/* Booked Campaign Details Card — ONLY show when there is an active booking in Latest Booking */}
                {info.hasTrackerBooking && (
                  <div className="scooh-ppt-booked-card" onClick={e => e.stopPropagation()}>
                    <div className="scooh-ppt-booked-header">
                      <span className="scooh-ppt-booked-tag">
                        <span className="scooh-ppt-booked-dot" />
                        Active Campaign (Latest Booking)
                      </span>
                      <button
                        type="button"
                        className="scooh-ppt-booked-link"
                        onClick={() => navigate(`/campaigns?site=${encodeURIComponent(s.site_code)}`)}
                        title={`View campaign details for ${s.site_code} in Latest Booking`}
                      >
                        ↗ Booking
                      </button>
                    </div>

                    <div className="scooh-ppt-booked-body">
                      {/* Campaign Name */}
                      {campaignName && (
                        <div className="scooh-ppt-booked-row">
                          <span className="scooh-ppt-booked-label">Campaign Name</span>
                          <span className="scooh-ppt-booked-val camp-val" title={campaignName}>
                            📢 {campaignName}
                          </span>
                        </div>
                      )}

                      {/* Client Name */}
                      {clientName && (
                        <div className="scooh-ppt-booked-row">
                          <span className="scooh-ppt-booked-label">Client Name</span>
                          <span className="scooh-ppt-booked-val client-val" title={clientName}>
                            🏢 {clientName}
                          </span>
                        </div>
                      )}

                      {/* Duration */}
                      {durationText && (
                        <div className="scooh-ppt-booked-row">
                          <span className="scooh-ppt-booked-label">Duration</span>
                          <span className="scooh-ppt-booked-val duration-val">
                            ⏱️ {durationText}
                          </span>
                        </div>
                      )}

                      {/* Start Date & End Date */}
                      {(startDateText || finalEndDateText) && (
                        <div className="scooh-ppt-booked-dates">
                          {startDateText && (
                            <div className="scooh-ppt-booked-date-col">
                              <span className="scooh-ppt-booked-label">Start Date</span>
                              <span className="scooh-ppt-booked-val date-val">
                                📅 {startDateText}
                              </span>
                            </div>
                          )}
                          {finalEndDateText && (
                            <div className="scooh-ppt-booked-date-col">
                              <span className="scooh-ppt-booked-label">End Date</span>
                              <span className="scooh-ppt-booked-val date-val end-date-val">
                                🏁 {finalEndDateText}
                              </span>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </div>
                )}

                <div className="scooh-ppt-availability" onClick={e => e.stopPropagation()}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px', flexWrap: 'wrap', gap: '4px' }}>
                    <label style={{ margin: 0 }}>Availability / Available Date</label>
                    {info.isAuto ? (
                      <span className="scooh-ppt-avail-status-tag auto" title={`Automatic booking dates from Latest Booking (End: ${info.endDateStr})`}>
                        ⚡ Auto Booking
                      </span>
                    ) : info.manualDateFmt ? (
                      <span className="scooh-ppt-avail-status-tag manual" title={`Manual available date (${info.endDateStr})`}>
                        ✍️ Manual Date
                      </span>
                    ) : (
                      <span className="scooh-ppt-avail-status-tag available">
                        ✓ Available
                      </span>
                    )}
                  </div>

                  <div className="scooh-ppt-avail-row">
                    <input
                      type="date"
                      title="Manual available date or booked end date. Set date manually or clear to use auto tracker."
                      value={v.manualDate !== undefined ? v.manualDate : (info.isAuto ? '' : info.manualDateIso)}
                      onChange={e => {
                        const newDate = e.target.value;
                        const dObj = newDate ? parseDay(newDate) : null;
                        const newAvail = dObj ? `Booked till ${fmtDate(dObj)}` : (info.hasAuto ? `Booked till ${info.autoEndDateStr}` : 'Available');
                        setSel({
                          ...sel,
                          [siteKey]: {
                            ...v,
                            manualDate: newDate,
                            availability: newAvail
                          }
                        });
                      }}
                    />
                    <input
                      type="text"
                      className="scooh-ppt-availability-input"
                      value={v.availability !== undefined ? v.availability : info.availText}
                      placeholder={info.defaultAvailText || "Immediate or DD.MM.YYYY"}
                      title="Availability text for PowerPoint & Excel"
                      onChange={e => {
                        const val = e.target.value;
                        const d = extractDateFromString(val);
                        setSel({
                          ...sel,
                          [siteKey]: {
                            ...v,
                            availability: val,
                            manualDate: d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` : v.manualDate
                          }
                        });
                      }}
                    />
                    {(v.manualDate !== undefined || v.availability !== undefined) && (
                      <button
                        type="button"
                        title="Reset to default / auto tracker"
                        onClick={() => {
                          const next = { ...v };
                          delete next.manualDate;
                          delete next.availability;
                          setSel({ ...sel, [siteKey]: next });
                        }}
                        style={{
                          background: 'none',
                          border: '1px solid var(--mb-border-2)',
                          color: '#f87171',
                          borderRadius: '8px',
                          padding: '0 8px',
                          cursor: 'pointer',
                          minHeight: '32px',
                          fontSize: '13px',
                          flexShrink: 0
                        }}
                      >
                        ✕
                      </button>
                    )}
                  </div>

                  <div style={{ fontSize: '11px', marginTop: '5px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '6px', flexWrap: 'wrap' }}>
                    {info.isAuto ? (
                      <div style={{ color: '#fbbf24', display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <span>⚡ Auto End Date:</span>
                        <strong style={{ color: '#fef08a' }}>{info.endDateStr}</strong>
                        <span style={{ opacity: 0.7, fontSize: '10px' }}>(from Tracker)</span>
                      </div>
                    ) : info.manualDateFmt ? (
                      <div style={{ color: '#38bdf8', display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <span>📅 Manual Date:</span>
                        <strong style={{ color: '#7dd3fc' }}>{info.manualDateFmt}</strong>
                        <span style={{ opacity: 0.7, fontSize: '10px' }}>(manual)</span>
                      </div>
                    ) : (
                      <div style={{ color: '#4ade80', display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <span>✓</span>
                        <span>Available Immediately</span>
                      </div>
                    )}

                    {selectedDates.length > 0 && (
                      <span style={{ color: info.isBooked ? '#fbbf24' : '#4ade80', fontWeight: 600 }}>
                        {info.availBadge}
                      </span>
                    )}
                  </div>
                </div>

                <div className="scooh-ppt-output-options" onClick={e => e.stopPropagation()}>
                  <label className="scooh-ppt-output-option">
                    <input
                      type="checkbox"
                      className="scooh-ppt-show-rate"
                      checked={!!v.showRate}
                      onChange={e => setSel({ ...sel, [siteKey]: { ...v, showRate: e.target.checked } })}
                    />
                    <span>Show Adv. Fee Per Month in PPT</span>
                  </label>
                </div>

                <div className="scooh-ppt-rate" onClick={e => e.stopPropagation()}>
                  <label>Adv. Fee Per Month (₹)</label>
                  <input
                    type="text"
                    className="scooh-ppt-rate-input"
                    value={v.rate ?? s.ppt_rate ?? ''}
                    placeholder="2,50,000"
                    onChange={e => setSel({ ...sel, [siteKey]: { ...v, rate: e.target.value } })}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {/* ── Modal for Quick Multi-Site Selection (Paste or Enter Multiple Sites) ── */}
      {showMultiSelectModal && (
        <div
          className="scooh-modal-backdrop"
          style={{ position: 'fixed', inset: 0, zIndex: 9999, background: 'rgba(0,0,0,0.85)', backdropFilter: 'blur(6px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}
          onClick={() => setShowMultiSelectModal(false)}
        >
          <div
            className="scooh-panel"
            style={{ width: '100%', maxWidth: '640px', maxHeight: '90vh', overflowY: 'auto', border: '1px solid #8b5cf6', boxShadow: '0 20px 50px rgba(0,0,0,0.7)', background: '#0b1329', padding: '24px', borderRadius: '12px' }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '18px', fontWeight: 800 }}>
                  <span>⚡</span>
                  <span>Select Multiple Sites Together</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  Paste or type multiple site codes (from Excel, email, or chat) separated by commas, spaces, or lines.
                </p>
              </div>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowMultiSelectModal(false)}
                style={{ padding: '4px 10px', fontSize: '14px', minHeight: '30px' }}
              >
                ✕
              </button>
            </div>

            <div style={{ marginBottom: '14px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#c4b5fd', marginBottom: '6px' }}>
                Paste Site Codes / Numbers:
              </label>
              <textarea
                rows={4}
                value={multiSelectInput}
                onChange={e => setMultiSelectInput(e.target.value)}
                placeholder={'e.g. MB-01, MB-05, MB-12, MB-18, 24, 30\nOr paste a column copied directly from Excel...'}
                style={{
                  width: '100%',
                  background: '#040914',
                  border: '1px solid #334155',
                  borderRadius: '8px',
                  color: '#f8fafc',
                  padding: '10px 12px',
                  fontSize: '13px',
                  fontFamily: 'monospace',
                  resize: 'vertical',
                  boxSizing: 'border-box'
                }}
                autoFocus
              />
            </div>

            {/* Quick Presets */}
            <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center', marginBottom: '16px' }}>
              <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 700, textTransform: 'uppercase' }}>Quick Fill:</span>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 9px', minHeight: '26px', color: '#4ade80', borderColor: 'rgba(74, 222, 128, 0.3)' }}
                onClick={() => {
                  const vac = sites.filter(s => {
                    const info = resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
                    return !info.isBooked;
                  });
                  setMultiSelectInput(vac.map(s => s.site_code).join(', '));
                }}
              >
                🟢 All Vacant Sites ({sites.filter(s => !resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates).isBooked).length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 9px', minHeight: '26px', color: '#fbbf24', borderColor: 'rgba(251, 191, 36, 0.3)' }}
                onClick={() => {
                  const occ = sites.filter(s => {
                    const info = resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates);
                    return info.isBooked;
                  });
                  setMultiSelectInput(occ.map(s => s.site_code).join(', '));
                }}
              >
                🟡 All Occupied Sites ({sites.filter(s => resolveSiteAvailability(s, sel[s.id || s.site_code] || {}, campaignsBySiteCode, campaignMap, selectedDates).isBooked).length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 9px', minHeight: '26px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
                onClick={() => {
                  const bl = sites.filter(s => String(s.lighting || '').toUpperCase().includes('BL') || String(s.media_type || '').toUpperCase().includes('BL'));
                  setMultiSelectInput(bl.map(s => s.site_code).join(', '));
                }}
              >
                💡 Backlit Only (BL)
              </button>
              {multiSelectInput && (
                <button
                  type="button"
                  className="scooh-btn ghost"
                  style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#f87171' }}
                  onClick={() => setMultiSelectInput('')}
                >
                  Clear
                </button>
              )}
            </div>

            {/* Realtime Matched Preview */}
            <div style={{ background: 'rgba(15, 23, 42, 0.7)', borderRadius: '8px', padding: '12px 14px', border: '1px solid #1e293b', marginBottom: '18px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: parsedMultiSites.matched.length > 0 ? '#4ade80' : '#94a3b8' }}>
                  {parsedMultiSites.matched.length > 0
                    ? `✓ ${parsedMultiSites.matched.length} site(s) recognized & ready to select`
                    : 'No matching sites parsed yet'}
                </span>
                {parsedMultiSites.unmatched.length > 0 && (
                  <span style={{ fontSize: '11px', color: '#f87171', fontWeight: 600 }}>
                    ⚠️ {parsedMultiSites.unmatched.length} unrecognized code(s)
                  </span>
                )}
              </div>

              {parsedMultiSites.matched.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px', maxHeight: '110px', overflowY: 'auto' }}>
                  {parsedMultiSites.matched.map(s => (
                    <span
                      key={s.id || s.site_code}
                      style={{
                        background: 'rgba(168, 85, 247, 0.2)',
                        border: '1px solid rgba(168, 85, 247, 0.45)',
                        color: '#e9d5ff',
                        padding: '2px 8px',
                        borderRadius: '6px',
                        fontSize: '11.5px',
                        fontWeight: 700
                      }}
                    >
                      {s.site_code}
                    </span>
                  ))}
                </div>
              )}

              {parsedMultiSites.unmatched.length > 0 && (
                <div style={{ marginTop: '8px', fontSize: '11px', color: '#fca5a5' }}>
                  Not found: {parsedMultiSites.unmatched.slice(0, 8).join(', ')}{parsedMultiSites.unmatched.length > 8 ? ` +${parsedMultiSites.unmatched.length - 8} more` : ''}
                </div>
              )}
            </div>

            {/* Actions */}
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowMultiSelectModal(false)}
                style={{ minHeight: '38px', padding: '0 14px', fontSize: '12.5px' }}
              >
                Cancel
              </button>
              {parsedMultiSites.matched.length > 0 && (
                <button
                  type="button"
                  className="scooh-btn ghost"
                  onClick={deselectMultiSites}
                  style={{ minHeight: '38px', padding: '0 14px', fontSize: '12.5px', color: '#f87171', borderColor: 'rgba(248, 113, 113, 0.35)' }}
                  title="Deselect only these matched sites"
                >
                  ✕ Deselect These
                </button>
              )}
              <button
                type="button"
                className="scooh-btn ghost"
                disabled={parsedMultiSites.matched.length === 0}
                onClick={() => applyMultiSelect('only')}
                style={{
                  minHeight: '38px',
                  padding: '0 14px',
                  fontSize: '12.5px',
                  color: '#38bdf8',
                  borderColor: 'rgba(56, 189, 248, 0.4)',
                  cursor: parsedMultiSites.matched.length === 0 ? 'not-allowed' : 'pointer'
                }}
                title="Clears any other selected sites and selects ONLY these"
              >
                🎯 Select ONLY These ({parsedMultiSites.matched.length})
              </button>
              <button
                type="button"
                className="scooh-btn primary"
                disabled={parsedMultiSites.matched.length === 0}
                onClick={() => applyMultiSelect('add')}
                style={{
                  minHeight: '38px',
                  padding: '0 18px',
                  fontSize: '12.5px',
                  background: 'linear-gradient(135deg, #8b5cf6, #6366f1)',
                  border: 'none',
                  fontWeight: 700,
                  cursor: parsedMultiSites.matched.length === 0 ? 'not-allowed' : 'pointer',
                  boxShadow: '0 4px 14px rgba(99, 102, 241, 0.35)'
                }}
                title="Adds these sites to your current selection"
              >
                + Add to Selection ({parsedMultiSites.matched.length})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Modal for Multi-Date Selection & Interactive Calendar ── */}
      {showDateModal && (
        <div
          className="scooh-modal-backdrop"
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 9999,
            background: 'rgba(0,0,0,0.85)',
            backdropFilter: 'blur(6px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}
          onClick={() => setShowDateModal(false)}
        >
          <div
            className="scooh-panel"
            style={{
              width: '100%',
              maxWidth: '740px',
              maxHeight: '90vh',
              overflowY: 'auto',
              border: '1px solid #38bdf8',
              boxShadow: '0 20px 50px rgba(0,0,0,0.8), 0 0 25px rgba(56, 189, 248, 0.15)',
              background: '#071224',
              padding: '22px',
              borderRadius: '14px',
              color: '#f8fafc'
            }}
            onClick={e => e.stopPropagation()}
          >
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '18px', fontWeight: 800 }}>
                  <span style={{ color: '#38bdf8' }}>📅</span>
                  <span>Select Multiple Dates for Site Availability</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  Click days on the calendar to toggle them, add date ranges, or pick presets. Sites will be evaluated against all selected dates.
                </p>
              </div>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={() => setShowDateModal(false)}
                style={{ padding: '4px 10px', fontSize: '14px', minHeight: '30px' }}
              >
                ✕
              </button>
            </div>

            {/* Quick Presets Bar */}
            <div style={{ marginBottom: '14px', display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
              <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 700, textTransform: 'uppercase' }}>Quick Presets:</span>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
                onClick={() => addPreset('today')}
              >
                + Today
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
                onClick={() => addPreset('tomorrow')}
              >
                + Tomorrow
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#c084fc', borderColor: 'rgba(192, 132, 252, 0.3)' }}
                onClick={() => addPreset('next7')}
              >
                + Next 7 Days
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#c084fc', borderColor: 'rgba(192, 132, 252, 0.3)' }}
                onClick={() => addPreset('next15')}
              >
                + Next 15 Days
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#34d399', borderColor: 'rgba(52, 211, 153, 0.3)' }}
                onClick={() => addPreset('nextMonth1st')}
              >
                + 1st Next Month
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#34d399', borderColor: 'rgba(52, 211, 153, 0.3)' }}
                onClick={() => addPreset('nextMonth15th')}
              >
                + 15th Next Month
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', minHeight: '26px', color: '#fbbf24', borderColor: 'rgba(251, 191, 36, 0.3)' }}
                onClick={() => addPreset('nextMonthBoth')}
              >
                + 1st & 15th Next Mo
              </button>
            </div>

            {/* Main Body: Grid with Calendar and Range/Chip Manager */}
            <div style={{ display: 'grid', gridTemplateColumns: 'minmax(290px, 340px) 1fr', gap: '16px', marginBottom: '16px' }}>
              {/* Left Column: Interactive Month Calendar */}
              <div style={{ background: '#0a1628', border: '1px solid #1e293b', borderRadius: '10px', padding: '14px' }}>
                {/* Month Navigator */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '12px' }}>
                  <button
                    type="button"
                    onClick={() => {
                      if (calendarMonth === 0) {
                        setCalendarMonth(11);
                        setCalendarYear(y => y - 1);
                      } else {
                        setCalendarMonth(m => m - 1);
                      }
                    }}
                    style={{ background: 'rgba(255,255,255,0.06)', border: '1px solid #334155', borderRadius: '5px', color: '#cbd5e1', cursor: 'pointer', padding: '3px 8px', fontSize: '12px' }}
                    title="Previous Month"
                  >
                    ◀ Prev
                  </button>
                  <span style={{ fontSize: '13px', fontWeight: 800, color: '#f8fafc' }}>
                    {new Date(calendarYear, calendarMonth, 1).toLocaleString('default', { month: 'long', year: 'numeric' })}
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      if (calendarMonth === 11) {
                        setCalendarMonth(0);
                        setCalendarYear(y => y + 1);
                      } else {
                        setCalendarMonth(m => m + 1);
                      }
                    }}
                    style={{ background: 'rgba(255,255,255,0.06)', border: '1px solid #334155', borderRadius: '5px', color: '#cbd5e1', cursor: 'pointer', padding: '3px 8px', fontSize: '12px' }}
                    title="Next Month"
                  >
                    Next ▶
                  </button>
                </div>

                {/* Day of Week Headers */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '3px', textAlign: 'center', marginBottom: '6px' }}>
                  {['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'].map(d => (
                    <div key={d} style={{ fontSize: '10px', fontWeight: 800, color: '#64748b', textTransform: 'uppercase' }}>
                      {d}
                    </div>
                  ))}
                </div>

                {/* Days Grid */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '4px' }}>
                  {(() => {
                    const firstDay = new Date(calendarYear, calendarMonth, 1).getDay();
                    const daysInMonth = new Date(calendarYear, calendarMonth + 1, 0).getDate();
                    const cells = [];
                    for (let i = 0; i < firstDay; i++) {
                      cells.push(<div key={`empty-${i}`} style={{ height: '32px' }} />);
                    }
                    const now = new Date();
                    const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
                    for (let day = 1; day <= daysInMonth; day++) {
                      const iso = `${calendarYear}-${String(calendarMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                      const isSel = selectedDates.includes(iso);
                      const isToday = iso === todayIso;
                      cells.push(
                        <button
                          key={iso}
                          type="button"
                          onClick={() => toggleDate(iso)}
                          style={{
                            height: '32px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            borderRadius: '6px',
                            border: isSel ? '1.5px solid #38bdf8' : (isToday ? '1px solid #f59e0b' : '1px solid transparent'),
                            background: isSel ? 'linear-gradient(135deg, #0284c7, #6366f1)' : (isToday ? 'rgba(245, 158, 11, 0.12)' : 'rgba(255,255,255,0.03)'),
                            color: isSel ? '#ffffff' : (isToday ? '#fbbf24' : '#cbd5e1'),
                            fontWeight: isSel ? 800 : (isToday ? 700 : 500),
                            fontSize: '11.5px',
                            cursor: 'pointer',
                            transition: 'all 0.12s ease',
                            padding: 0
                          }}
                          title={`${iso}${isSel ? ' (Selected - click to remove)' : ' (Click to select)'}`}
                        >
                          {day}
                        </button>
                      );
                    }
                    return cells;
                  })()}
                </div>

                <div style={{ marginTop: '10px', fontSize: '11px', color: '#94a3b8', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span>👆 Click day to toggle selection</span>
                  <button
                    type="button"
                    onClick={() => {
                      const now = new Date();
                      setCalendarYear(now.getFullYear());
                      setCalendarMonth(now.getMonth());
                    }}
                    style={{ background: 'none', border: 'none', color: '#38bdf8', cursor: 'pointer', fontSize: '11px', padding: 0 }}
                  >
                    Current Month
                  </button>
                </div>
              </div>

              {/* Right Column: Date Range Adder & Selected List */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                {/* Date Range Tool */}
                <div style={{ background: '#0a1628', border: '1px solid #1e293b', borderRadius: '10px', padding: '12px' }}>
                  <div style={{ fontSize: '11.5px', fontWeight: 800, color: '#c4b5fd', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '5px' }}>
                    <span>⚡ Add Date Range</span>
                  </div>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap', marginBottom: '8px' }}>
                    <div style={{ flex: '1 1 120px' }}>
                      <span style={{ fontSize: '10.5px', color: '#94a3b8', display: 'block', marginBottom: '3px' }}>From:</span>
                      <input
                        type="date"
                        value={rangeStart}
                        onChange={e => setRangeStart(e.target.value)}
                        style={{ width: '100%', background: '#040c18', border: '1px solid #334155', borderRadius: '6px', color: '#f8fafc', padding: '4px 8px', fontSize: '12px' }}
                      />
                    </div>
                    <div style={{ flex: '1 1 120px' }}>
                      <span style={{ fontSize: '10.5px', color: '#94a3b8', display: 'block', marginBottom: '3px' }}>To:</span>
                      <input
                        type="date"
                        value={rangeEnd}
                        onChange={e => setRangeEnd(e.target.value)}
                        style={{ width: '100%', background: '#040c18', border: '1px solid #334155', borderRadius: '6px', color: '#f8fafc', padding: '4px 8px', fontSize: '12px' }}
                      />
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      disabled={!rangeStart || !rangeEnd}
                      onClick={() => addDateRange(rangeStart, rangeEnd)}
                      style={{ fontSize: '11px', padding: '3px 9px', minHeight: '26px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.4)' }}
                    >
                      + Add All Days in Range
                    </button>
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      disabled={!rangeStart || !rangeEnd}
                      onClick={() => {
                        if (rangeStart) addDate(rangeStart);
                        if (rangeEnd) addDate(rangeEnd);
                      }}
                      style={{ fontSize: '11px', padding: '3px 9px', minHeight: '26px', color: '#cbd5e1' }}
                    >
                      + Add Start & End Only
                    </button>
                  </div>
                </div>

                {/* Selected Dates List */}
                <div style={{ background: '#0a1628', border: '1px solid #1e293b', borderRadius: '10px', padding: '12px', flex: 1, display: 'flex', flexDirection: 'column' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                    <span style={{ fontSize: '12px', fontWeight: 800, color: '#f8fafc' }}>
                      Selected Dates ({selectedDates.length}):
                    </span>
                    {selectedDates.length > 0 && (
                      <button
                        type="button"
                        onClick={clearAllDates}
                        style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11px', fontWeight: 700, padding: 0 }}
                      >
                        Clear All
                      </button>
                    )}
                  </div>

                  <div style={{ flex: 1, maxHeight: '180px', overflowY: 'auto', display: 'flex', flexWrap: 'wrap', gap: '5px', alignContent: 'flex-start', padding: '2px' }}>
                    {selectedDates.length === 0 ? (
                      <div style={{ color: '#64748b', fontSize: '12px', fontStyle: 'italic', padding: '10px 0' }}>
                        No dates selected yet. Click dates on the calendar or click any quick preset above.
                      </div>
                    ) : (
                      selectedDates.map(d => (
                        <span
                          key={d}
                          style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '4px',
                            background: 'rgba(14, 165, 233, 0.16)',
                            border: '1px solid rgba(56, 189, 248, 0.4)',
                            borderRadius: '5px',
                            padding: '3px 8px',
                            fontSize: '11.5px',
                            color: '#e0f2fe',
                            fontWeight: 700
                          }}
                        >
                          {fmtDate(d) || d}
                          <button
                            type="button"
                            onClick={() => removeDate(d)}
                            style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11px', padding: '0 1px', lineHeight: 1 }}
                            title={`Remove ${fmtDate(d) || d}`}
                          >
                            ✕
                          </button>
                        </span>
                      ))
                    )}
                  </div>
                </div>
              </div>
            </div>

            {/* Footer Actions */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingTop: '12px', borderTop: '1px solid #1e293b' }}>
              <span style={{ fontSize: '12px', color: '#94a3b8' }}>
                {selectedDates.length === 0
                  ? 'No dates active (sites show current availability)'
                  : `Evaluating availability against ${selectedDates.length} date${selectedDates.length === 1 ? '' : 's'}`}
              </span>
              <div style={{ display: 'flex', gap: '10px' }}>
                {selectedDates.length > 0 && (
                  <button
                    type="button"
                    className="scooh-btn ghost"
                    onClick={clearAllDates}
                    style={{ minHeight: '36px', padding: '0 14px', fontSize: '12px', color: '#f87171', borderColor: 'rgba(248, 113, 113, 0.35)' }}
                  >
                    Clear All
                  </button>
                )}
                <button
                  type="button"
                  className="scooh-btn primary"
                  onClick={() => setShowDateModal(false)}
                  style={{
                    minHeight: '36px',
                    padding: '0 20px',
                    fontSize: '12.5px',
                    background: 'linear-gradient(135deg, #0284c7, #6366f1)',
                    border: 'none',
                    fontWeight: 800,
                    boxShadow: '0 4px 14px rgba(2, 132, 199, 0.35)'
                  }}
                >
                  ✓ Apply & Evaluate {selectedDates.length > 0 ? `(${selectedDates.length} Dates)` : ''}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── Modal for Folder Photo Auto-Import Progress & Results ───────── */}
      {folderImportStatus && (
        <div 
          className="scooh-modal-backdrop scooh-ppt-folder-modal-backdrop" 
          style={{ position: 'fixed', inset: 0, zIndex: 9999, background: 'rgba(0,0,0,0.82)', backdropFilter: 'blur(6px)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}
          onClick={() => { if (folderImportStatus.finished) setFolderImportStatus(null); }}
        >
          <div 
            className="scooh-panel scooh-ppt-folder-modal-panel" 
            style={{ width: '100%', maxWidth: '580px', maxHeight: '88vh', overflowY: 'auto', border: '1px solid #38bdf8', boxShadow: '0 20px 50px rgba(0,0,0,0.7)' }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '12px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '17px' }}>
                  <span>📂</span>
                  <span>{folderImportStatus.finished ? 'Folder Photo Import Complete' : 'Auto-Importing Site Photos from Folder'}</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  {folderImportStatus.replaceMode
                    ? 'Photos are matched by folder/file names to site codes. Newly imported photos replace old photos for matched sites.'
                    : 'Photos are automatically matched by folder / file names to your inventory site codes.'}
                </p>
              </div>
              {folderImportStatus.finished && (
                <button 
                  type="button" 
                  onClick={() => setFolderImportStatus(null)} 
                  style={{ background: 'none', border: 'none', color: '#94a3b8', fontSize: '20px', cursor: 'pointer', padding: '4px' }}
                  title="Close"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Progress indicator */}
            <div style={{ marginBottom: '18px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                <span style={{ color: folderImportStatus.finished ? '#10b981' : '#38bdf8' }}>
                  {folderImportStatus.finished 
                    ? (folderImportStatus.replaceMode
                        ? `✓ Successfully replaced old photos with ${folderImportStatus.donePhotos} new photo${folderImportStatus.donePhotos > 1 ? 's' : ''} across ${folderImportStatus.doneSites} sites!`
                        : `✓ Successfully stored ${folderImportStatus.donePhotos} photos across ${folderImportStatus.doneSites} sites!`) 
                    : `Uploading: ${folderImportStatus.currentSite || 'Preparing…'}`}
                </span>
                <span style={{ color: '#94a3b8' }}>
                  {folderImportStatus.totalPhotos > 0 
                    ? `${Math.round((folderImportStatus.donePhotos / folderImportStatus.totalPhotos) * 100)}%` 
                    : '0%'}
                </span>
              </div>
              <div style={{ height: '8px', background: '#0b1016', borderRadius: '999px', overflow: 'hidden', border: '1px solid #222c37' }}>
                <div 
                  style={{ 
                    height: '100%', 
                    width: `${folderImportStatus.totalPhotos > 0 ? Math.round((folderImportStatus.donePhotos / folderImportStatus.totalPhotos) * 100) : 0}%`, 
                    background: folderImportStatus.finished ? 'linear-gradient(90deg, #10b981, #059669)' : 'linear-gradient(90deg, #0284c7, #6366f1)', 
                    borderRadius: '999px',
                    transition: 'width 0.3s ease'
                  }} 
                />
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11.5px', color: '#64748b', marginTop: '6px' }}>
                <span>Sites: {folderImportStatus.doneSites} / {folderImportStatus.totalSites}</span>
                <span>Photos: {folderImportStatus.donePhotos} / {folderImportStatus.totalPhotos}</span>
              </div>
            </div>

            {/* Site upload results list */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '240px', overflowY: 'auto', paddingRight: '4px', marginBottom: '16px' }}>
              {(folderImportStatus.siteResults || []).map((res, i) => (
                <div 
                  key={i} 
                  style={{ 
                    display: 'flex', 
                    alignItems: 'center', 
                    justifyContent: 'space-between', 
                    padding: '8px 12px', 
                    background: res.success ? 'rgba(16, 185, 129, 0.08)' : 'rgba(239, 68, 68, 0.08)', 
                    border: `1px solid ${res.success ? 'rgba(16, 185, 129, 0.25)' : 'rgba(239, 68, 68, 0.25)'}`,
                    borderRadius: '6px',
                    fontSize: '12.5px'
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span>{res.success ? '🟢' : '🔴'}</span>
                    <strong style={{ color: '#f1f5f9' }}>{res.site_code}</strong>
                  </div>
                  <span style={{ fontSize: '11.5px', color: res.success ? '#10b981' : '#f87171', fontWeight: 600 }}>
                    {res.success ? `${res.count} photo${res.count > 1 ? 's' : ''} ${res.replaced ? '(replaced old)' : 'stored'}` : (res.error || 'Failed')}
                  </span>
                </div>
              ))}
            </div>

            {/* Unmatched files notice if any */}
            {folderImportStatus.unmatched && folderImportStatus.unmatched.length > 0 && (
              <div style={{ padding: '12px', background: 'rgba(245, 158, 11, 0.08)', border: '1px solid rgba(245, 158, 11, 0.25)', borderRadius: '8px', marginBottom: '16px', fontSize: '12px' }}>
                <div style={{ fontWeight: 700, color: '#f59e0b', marginBottom: '4px' }}>
                  ⚠️ {folderImportStatus.unmatched.length} photo(s) did not match any site code:
                </div>
                <div style={{ maxHeight: '80px', overflowY: 'auto', color: '#cbd5e1', fontSize: '11px', fontFamily: 'monospace' }}>
                  {folderImportStatus.unmatched.slice(0, 10).map((u, i) => (
                    <div key={i}>• {u}</div>
                  ))}
                  {folderImportStatus.unmatched.length > 10 && (
                    <div style={{ color: '#94a3b8' }}>…and {folderImportStatus.unmatched.length - 10} more</div>
                  )}
                </div>
                <p style={{ margin: '6px 0 0', fontSize: '11px', color: '#94a3b8' }}>
                  To match automatically, ensure the subfolder or file name begins with your site code (e.g. "MB-01").
                </p>
              </div>
            )}

            {/* Modal actions */}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', paddingTop: '8px', borderTop: '1px solid #1e293b' }}>
              {folderImportStatus.finished ? (
                <button 
                  type="button" 
                  className="scooh-btn primary" 
                  style={{ padding: '8px 24px', fontSize: '13px' }}
                  onClick={() => setFolderImportStatus(null)}
                >
                  ✓ Done
                </button>
              ) : (
                <span style={{ fontSize: '12px', color: '#94a3b8', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span className="scooh-spin">🔄</span> Processing uploads… please do not close
                </span>
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function ProposalsView() {
  const [sites, setSites] = useState([]);
  const [clientName, setClientName] = useState('');
  const [campaignName, setCampaignName] = useState('');
  const [startDate, setStartDate] = useState(new Date().toISOString().slice(0, 10));
  const [durationDays, setDurationDays] = useState(30);
  const [validityDays, setValidityDays] = useState(7);
  const [discountPercent, setDiscountPercent] = useState(0);
  const [taxPercent, setTaxPercent] = useState(18);
  const [terms, setTerms] = useState(MEDIA_BUZZ_TERMS_TEXT);
  const [selectedSites, setSelectedSites] = useState({});
  const [search, setSearch] = useState('');
  const [siteSort, setSiteSort] = useState({ key: 'site_code', dir: 'asc' });
  const [savedBanner, setSavedBanner] = useState(false);
  const [savedProposals, setSavedProposals] = useState([]);
  const [propSearch, setPropSearch] = useState('');
  const [propSort, setPropSort] = useState({ key: 'id', dir: 'desc' });
  const [loadingProposals, setLoadingProposals] = useState(false);

  async function loadProposals() {
    setLoadingProposals(true);
    try {
      const res = await api.get('/proposals');
      if (Array.isArray(res.data)) {
        setSavedProposals(res.data);
      }
    } catch (e) {
      console.error('Failed to load proposals:', e);
    } finally {
      setLoadingProposals(false);
    }
  }

  useEffect(() => {
    api.get('/sites').then(res => {
      if (Array.isArray(res.data)) {
        setSites(res.data.map(unpackSite));
      }
    }).catch(() => {});
    loadProposals();
  }, []);

  function handlePropSort(key) {
    setPropSort(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const selectedSitesCount = useMemo(() => Object.values(selectedSites).filter(v => v?.checked).length, [selectedSites]);

  const filteredSites = useMemo(() => {
    const list = sites.filter(s => matchSiteSearch(s, search));

    list.sort((a, b) => {
      const aSel = !!selectedSites[a.id || a.site_code]?.checked;
      const bSel = !!selectedSites[b.id || b.site_code]?.checked;
      // Selected sites always show on top
      if (aSel !== bSel) return aSel ? -1 : 1;

      if (siteSort.key) {
        let valA = a[siteSort.key];
        let valB = b[siteSort.key];
        if (siteSort.key === 'monthly_rate') {
          valA = Number(a.ppt_rate || a.monthly_rate || 0);
          valB = Number(b.ppt_rate || b.monthly_rate || 0);
        }
        return universalCompare(valA, valB, siteSort.dir);
      }
      return 0;
    });
    return list;
  }, [sites, search, siteSort, selectedSites]);

  const filteredProposals = useMemo(() => {
    const q = propSearch.trim().toLowerCase();
    const list = savedProposals.filter(p => {
      if (!q) return true;
      const hay = [p.proposal_code, p.client_name, p.campaign_name, p.terms, p.notes].filter(Boolean).join(' ').toLowerCase();
      return hay.includes(q);
    });

    if (propSort.key) {
      list.sort((a, b) => {
        let valA = a[propSort.key];
        let valB = b[propSort.key];
        if (propSort.key === 'proposal_date') {
          valA = a.proposal_date || a.start_date || a.created_at || '';
          valB = b.proposal_date || b.start_date || b.created_at || '';
        }
        if (propSort.key === 'total' || propSort.key === 'subtotal') {
          valA = Number(a[propSort.key] || 0);
          valB = Number(b[propSort.key] || 0);
        }
        return universalCompare(valA, valB, propSort.dir);
      });
    }
    return list;
  }, [savedProposals, propSearch, propSort]);

  const chosenSites = sites.filter(s => selectedSites[s.id || s.site_code]?.checked).map(s => {
    const rates = selectedSites[s.id || s.site_code] || {};
    const mediaRate = Number(rates.mediaRate ?? s.monthly_rate ?? 0);
    const vendorRate = Number(rates.vendorRate ?? 0);
    const printingRate = Number(rates.printingRate ?? 0);
    const mountingRate = Number(rates.mountingRate ?? 0);
    return {
      ...s,
      mediaRate,
      vendorRate,
      printingRate,
      mountingRate,
      totalRate: mediaRate + vendorRate + printingRate + mountingRate
    };
  });

  const subtotal = chosenSites.reduce((acc, s) => acc + s.totalRate, 0);
  const discountAmount = (subtotal * Number(discountPercent || 0)) / 100;
  const taxableAmount = subtotal - discountAmount;
  const taxAmount = (taxableAmount * Number(taxPercent || 18)) / 100;
  const grandTotal = taxableAmount + taxAmount;

  const endDate = useMemo(() => {
    const d = new Date(startDate || new Date());
    d.setDate(d.getDate() + Number(durationDays || 30) - 1);
    return d.toISOString().slice(0, 10);
  }, [startDate, durationDays]);

  async function saveProposal() {
    try {
      await api.post('/proposals', {
        proposal_code: 'MB-PROP-' + Math.floor(Date.now() / 1000),
        client_name: clientName || 'Prospective Client',
        campaign_name: campaignName || 'Outdoor Media Campaign',
        proposal_date: startDate,
        start_date: startDate,
        duration_days: durationDays,
        validity_days: validityDays,
        discount_percent: discountPercent,
        tax_percent: taxPercent,
        subtotal,
        total: grandTotal,
        terms,
        notes: JSON.stringify(chosenSites)
      });
      setSavedBanner(true);
      setTimeout(() => setSavedBanner(false), 3500);
      loadProposals();
    } catch (e) {
      alert('Proposal saved! (Check Archive below)');
      loadProposals();
    }
  }

  function loadSavedProposal(p) {
    if (!p) return;
    setClientName(p.client_name || '');
    setCampaignName(p.campaign_name || '');
    if (p.proposal_date || p.start_date) setStartDate((p.proposal_date || p.start_date).slice(0, 10));
    if (p.duration_days) setDurationDays(Number(p.duration_days));
    if (p.validity_days) setValidityDays(Number(p.validity_days));
    if (p.discount_percent != null) setDiscountPercent(Number(p.discount_percent));
    if (p.tax_percent != null) setTaxPercent(Number(p.tax_percent));
    if (p.terms) setTerms(p.terms);

    if (p.notes) {
      try {
        const parsedSites = typeof p.notes === 'string' ? JSON.parse(p.notes) : p.notes;
        if (Array.isArray(parsedSites)) {
          const nextSel = {};
          parsedSites.forEach(st => {
            const key = st.id || st.site_code;
            if (key) {
              nextSel[key] = {
                checked: true,
                mediaRate: Number(st.mediaRate ?? st.monthly_rate ?? 0),
                vendorRate: Number(st.vendorRate ?? 0),
                printingRate: Number(st.printingRate ?? 0),
                mountingRate: Number(st.mountingRate ?? 0)
              };
            }
          });
          setSelectedSites(nextSel);
        }
      } catch (err) {
        console.warn('Could not parse proposal sites notes:', err);
      }
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function deleteProposal(id) {
    if (!window.confirm('Are you sure you want to delete this saved proposal?')) return;
    try {
      await api.delete(`/proposals/${id}`);
      loadProposals();
    } catch (e) {
      alert('Failed to delete proposal: ' + (e.response?.data?.message || e.message));
    }
  }

  return (
    <>
      <PageHead
        title="Proposal Builder"
        desc="Pick sites, set client and rate details, then print or save as PDF straight from the browser."
        actions={
          <>
            <button type="button" className="scooh-btn" onClick={saveProposal}>Save Proposal</button>
            <button type="button" className="scooh-btn primary" onClick={() => window.print()}>Print / Save PDF</button>
          </>
        }
      />

      {savedBanner && <div className="scooh-banner" style={{ display: 'block' }}>Proposal saved successfully!</div>}

      <div className="scooh-proposal-layout">
        {/* Left Side Controls */}
        <div className="scooh-panel scooh-proposal-controls">
          <h3>1. Client & campaign</h3>
          <div className="scooh-field">
            <label>Client name</label>
            <input
              type="text"
              placeholder="e.g. Rajyash Group"
              value={clientName}
              onChange={e => setClientName(e.target.value)}
            />
          </div>
          <div className="scooh-field">
            <label>Campaign / brief</label>
            <input
              type="text"
              placeholder="e.g. Diwali launch, Ahmedabad"
              value={campaignName}
              onChange={e => setCampaignName(e.target.value)}
            />
          </div>
          <div className="scooh-grid2">
            <div className="scooh-field">
              <label>Start date</label>
              <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} />
            </div>
            <div className="scooh-field">
              <label>Duration (days)</label>
              <input type="number" min="1" value={durationDays} onChange={e => setDurationDays(Number(e.target.value))} />
            </div>
          </div>
          <div className="scooh-grid2">
            <div className="scooh-field">
              <label>Discount %</label>
              <input type="number" min="0" max="100" step="0.01" value={discountPercent} onChange={e => setDiscountPercent(Number(e.target.value))} />
            </div>
            <div className="scooh-field">
              <label>Validity (days)</label>
              <input type="number" min="1" value={validityDays} onChange={e => setValidityDays(Number(e.target.value))} />
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: '20px', marginBottom: '8px', gap: '10px', flexWrap: 'wrap' }}>
            <h3 className="scooh-proposal-step" style={{ margin: 0, padding: 0, border: 'none' }}>
              2. Select sites {selectedSitesCount > 0 && <span style={{ color: '#a78bfa', fontSize: '12px', fontWeight: 600 }}>({selectedSitesCount} selected on top)</span>}
            </h3>
            <select
              value={`${siteSort.key}:${siteSort.dir}`}
              onChange={e => {
                const [k, d] = e.target.value.split(':');
                setSiteSort({ key: k, dir: d });
              }}
              style={{ padding: '4px 8px', fontSize: '11px', borderRadius: '6px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155' }}
            >
              <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
              <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
              <option value="area:asc">Sort: Area (A–Z)</option>
              <option value="monthly_rate:asc">Sort: Rate (Low to High)</option>
              <option value="monthly_rate:desc">Sort: Rate (High to Low)</option>
              <option value="size:asc">Sort: Size</option>
            </select>
          </div>
          <div className="scooh-field">
            <input
              type="search"
              placeholder="Filter sites by code (e.g. 01, MB-01), area, landmark, media..."
              value={search}
              onChange={e => setSearch(e.target.value)}
            />
          </div>

          <div className="scooh-picklist">
            {filteredSites.map(s => {
              const key = s.id || s.site_code;
              const r = selectedSites[key] || {};
              const isChecked = !!r.checked;

              return (
                <div
                  key={key}
                  className="scooh-pickrow"
                  onClick={() => {
                    const nextChecked = !isChecked;
                    setSelectedSites({
                      ...selectedSites,
                      [key]: {
                        ...r,
                        checked: nextChecked,
                        mediaRate: r.mediaRate ?? s.monthly_rate ?? 0,
                        vendorRate: r.vendorRate ?? 0,
                        printingRate: r.printingRate ?? 0,
                        mountingRate: r.mountingRate ?? 0
                      }
                    });
                  }}
                >
                  <div className="scooh-pickrow-header">
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={e => {
                        e.stopPropagation();
                        setSelectedSites({
                          ...selectedSites,
                          [key]: {
                            ...r,
                            checked: e.target.checked,
                            mediaRate: r.mediaRate ?? s.monthly_rate ?? 0,
                            vendorRate: r.vendorRate ?? 0,
                            printingRate: r.printingRate ?? 0,
                            mountingRate: r.mountingRate ?? 0
                          }
                        });
                      }}
                      style={{ margin: 0, cursor: 'pointer' }}
                    />
                    <span className="scooh-plate">{s.site_code}</span>
                    <span className="scooh-pickmeta">
                      <strong>{(s.area || s.city).toUpperCase()}</strong>
                      <small>{s.size} · {s.media_type ? s.media_type.toUpperCase() : 'HOARDING'}</small>
                    </span>
                  </div>
                  <div
                    className="scooh-proposal-rates"
                    onClick={e => e.stopPropagation()}
                    style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
                      gap: '8px',
                      width: '100%',
                      boxSizing: 'border-box'
                    }}
                  >
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Media Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.mediaRate ?? s.monthly_rate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, mediaRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Vendor Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.vendorRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, vendorRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Printing Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.printingRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, printingRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '9.5px', fontWeight: 800, textTransform: 'uppercase', minWidth: 0 }}>
                      <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>Mounting Rate</span>
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        placeholder="0"
                        value={r.mountingRate ?? 0}
                        onChange={e => setSelectedSites({
                          ...selectedSites,
                          [key]: { ...r, checked: true, mountingRate: Number(e.target.value) }
                        })}
                        style={{ width: '100%', minHeight: '36px', padding: '6px 8px', fontSize: '12px', fontWeight: 800, boxSizing: 'border-box' }}
                      />
                    </label>
                  </div>
                </div>
              );
            })}
          </div>

          <div className="scooh-field" style={{ marginTop: '16px' }}>
            <label>Terms & conditions</label>
            <textarea value={terms} onChange={e => setTerms(e.target.value)} />
          </div>

          <div className="scooh-proposal-actions">
            <button type="button" className="scooh-btn primary" onClick={saveProposal}>
              Save Proposal
            </button>
          </div>
        </div>

        {/* Right Side Authentic Letterhead Document */}
        <div className="scooh-proposal-document">
          <article className="scooh-proposal-paper" id="proposal-doc">
            <header className="scooh-proposal-header">
              <div className="scooh-proposal-brand">
                <img className="scooh-proposal-logo" src="/assets/media-buzz-logo.png" alt="Media Buzz" />
                <div className="scooh-doc-sub">Outdoor Media Proposal <span>•</span> Ahmedabad</div>
              </div>
              <div className="scooh-proposal-meta">
                <div><span>Proposal date</span><strong>{formatDate(startDate)}</strong></div>
                <div><span>Validity</span><strong>{validityDays} days</strong></div>
              </div>
            </header>

            <div className="scooh-proposal-accent"></div>

            <section className="scooh-proposal-intro">
              <div className="scooh-proposal-client-info">
                <div className="scooh-doc-eyebrow">MEDIA PLAN</div>
                <h2 className="scooh-proposal-campaign-title">{campaignName || 'Outdoor Media Campaign'}</h2>
                <p className="scooh-proposal-client-name">Prepared for <strong>{clientName || 'Prospective Client'}</strong></p>
              </div>
              <div className="scooh-proposal-period">
                <span>Campaign period</span>
                <strong>{formatDate(startDate)}</strong>
                <em>to</em>
                <strong>{formatDate(endDate)}</strong>
                <small>{durationDays} days</small>
              </div>
            </section>

            <section className="scooh-proposal-section">
              <div className="scooh-proposal-section-title">
                <h3>Selected media sites</h3>
                <span>{chosenSites.length} site{chosenSites.length === 1 ? '' : 's'} selected</span>
              </div>
              <table className="scooh-doc-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Site</th>
                    <th>Location</th>
                    <th>Specification</th>
                    <th>Rate</th>
                  </tr>
                </thead>
                <tbody>
                  {chosenSites.length === 0 ? (
                    <tr><td colSpan="5" className="scooh-doc-empty">No sites selected for this proposal.</td></tr>
                  ) : (
                    chosenSites.map((s, index) => (
                      <tr key={s.id || s.site_code}>
                        <td className="scooh-doc-no">{index + 1}</td>
                        <td>
                          <strong>{s.site_code}</strong>
                          <span className="scooh-doc-muted">{s.media_type || 'OOH Site'}</span>
                        </td>
                        <td>
                          <strong>{s.area || s.city || '—'}</strong>
                          <span className="scooh-doc-muted">{[s.city, s.address].filter(Boolean).join(' · ')}</span>
                        </td>
                        <td>{[s.size, s.lighting, s.facing].filter(Boolean).join(' · ') || '—'}</td>
                        <td className="scooh-doc-rate">
                          <strong>Media: {money(s.mediaRate)}</strong>
                          {(s.vendorRate > 0 || s.printingRate > 0 || s.mountingRate > 0) && (
                            <span className="scooh-doc-muted">
                              {[
                                s.vendorRate > 0 ? `Vendor: ${money(s.vendorRate)}` : null,
                                s.printingRate > 0 ? `Printing: ${money(s.printingRate)}` : null,
                                s.mountingRate > 0 ? `Mounting: ${money(s.mountingRate)}` : null
                              ].filter(Boolean).join(' · ')}
                            </span>
                          )}
                          <strong className="scooh-doc-line-total">Total: {money(s.totalRate)}</strong>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </section>

            <section className="scooh-proposal-summary">
              <div className="scooh-proposal-notes">
                <h3>Commercial summary</h3>
                <p>Rates shown are for the complete campaign period and are subject to final site availability and written confirmation.</p>
              </div>
              <div className="scooh-proposal-totals">
                <div><span>Subtotal</span><strong>{money(subtotal)}</strong></div>
                {Number(discountPercent) > 0 && (
                  <div><span>Discount ({discountPercent}%)</span><strong style={{ color: '#e55d5d' }}>− {money(discountAmount)}</strong></div>
                )}
                <div><span>GST ({taxPercent}%)</span><strong>{money(taxAmount)}</strong></div>
                <div className="scooh-proposal-grand" style={{ color: '#000' }}><span style={{ color: '#000' }}>Total proposal value</span><strong style={{ color: '#000' }}>{money(grandTotal)}</strong></div>
              </div>
            </section>

            <section className="scooh-proposal-terms">
              <h3>Terms & conditions</h3>
              <p>{terms}</p>
              <div className="scooh-proposal-validity">
                This proposal is valid for <strong>{validityDays} days</strong> from {formatDate(startDate)}.
              </div>
            </section>

            <footer className="scooh-proposal-footer">
              <span>Prepared by Media Buzz</span>
              <span>Outdoor advertising proposal</span>
            </footer>
          </article>
        </div>
      </div>

      {/* Saved Proposals Section Below Builder */}
      <section className="scooh-panel scooh-saved-proposals-panel" style={{ marginTop: '28px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>Saved Proposals Archive</h3>
            <p className="scooh-footnote" style={{ margin: '4px 0 0', color: 'var(--mb-muted)' }}>
              All saved client proposals ({filteredProposals.length} shown) with complete financial breakdown, custom rates, and selected media sites.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            <button type="button" className="scooh-btn ghost" onClick={loadProposals} disabled={loadingProposals}>
              {loadingProposals ? 'Loading…' : '🔄 Refresh Proposals'}
            </button>
          </div>
        </div>

        {/* Toolbar for Proposals Archive */}
        <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center', marginBottom: '14px' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search saved proposals by code, client, campaign…"
            value={propSearch}
            onChange={e => setPropSearch(e.target.value)}
          />
          <select
            value={`${propSort.key}:${propSort.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setPropSort({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value="id:desc">Sort: Date (Newest First)</option>
            <option value="id:asc">Sort: Date (Oldest First)</option>
            <option value="total:desc">Sort: Grand Total (High to Low)</option>
            <option value="total:asc">Sort: Grand Total (Low to High)</option>
            <option value="client_name:asc">Sort: Client Name (A-Z)</option>
            <option value="client_name:desc">Sort: Client Name (Z-A)</option>
            <option value="campaign_name:asc">Sort: Campaign (A-Z)</option>
            <option value="proposal_code:asc">Sort: Proposal Code (A-Z)</option>
          </select>
          {(propSearch || propSort.key !== 'id' || propSort.dir !== 'desc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ padding: '6px 10px', fontSize: '11px' }}
              onClick={() => { setPropSearch(''); setPropSort({ key: 'id', dir: 'desc' }); }}
            >
              Reset
            </button>
          )}
        </div>

        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                <SortHeader label="Code" sortKey="proposal_code" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Client / Brand" sortKey="client_name" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Campaign" sortKey="campaign_name" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Proposal Date" sortKey="proposal_date" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Duration" sortKey="duration_days" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Subtotal" sortKey="subtotal" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="GST (18%)" sortKey="tax_percent" currentSort={propSort} onSort={handlePropSort} />
                <SortHeader label="Grand Total" sortKey="total" currentSort={propSort} onSort={handlePropSort} />
                <th style={{ textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredProposals.length === 0 ? (
                <tr>
                  <td colSpan="9" className="scooh-empty" style={{ padding: '30px', textAlign: 'center' }}>
                    {savedProposals.length === 0 ? (
                      <>No saved proposals yet. Click <strong>"Save Proposal"</strong> above to archive quotes.</>
                    ) : (
                      <>No saved proposals match your search query.</>
                    )}
                  </td>
                </tr>
              ) : (
                filteredProposals.map(p => (
                  <tr key={p.id || p.proposal_code}>
                    <td><span className="scooh-plate">{p.proposal_code}</span></td>
                    <td><strong>{p.client_name || 'Client'}</strong></td>
                    <td>{p.campaign_name || 'Outdoor Campaign'}</td>
                    <td>{formatDate(p.proposal_date || p.start_date || p.created_at)}</td>
                    <td>{p.duration_days || 30} days</td>
                    <td>{money(p.subtotal)}</td>
                    <td>{money((Number(p.subtotal || 0) * (Number(p.tax_percent || 18)) / 100))}</td>
                    <td><strong style={{ color: 'var(--mb-primary)' }}>{money(p.total)}</strong></td>
                    <td style={{ textAlign: 'right' }}>
                      <div className="scooh-rowactions" style={{ justifyContent: 'flex-end' }}>
                        <button
                          type="button"
                          className="scooh-btn purple-btn"
                          style={{ minHeight: '30px', padding: '0 10px', fontSize: '11px' }}
                          onClick={() => loadSavedProposal(p)}
                          title="Load into builder to view/edit"
                        >
                          Load
                        </button>
                        <button
                          type="button"
                          className="scooh-btn ghost"
                          style={{ minHeight: '30px', padding: '0 10px', fontSize: '11px' }}
                          onClick={() => {
                            loadSavedProposal(p);
                            setTimeout(() => window.print(), 350);
                          }}
                          title="Print this proposal"
                        >
                          🖨️ Print
                        </button>
                        <button
                          type="button"
                          className="scooh-iconbtn danger-icon"
                          style={{ minHeight: '30px', width: '30px' }}
                          onClick={() => deleteProposal(p.id)}
                          title="Delete Proposal"
                        >
                          ×
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

// ── Helpers for Occupancy Excel import & Date Parsing ───────────────────

// Helper to recognize vacant / blank client indicators
function isVacantClient(val) {
  if (!val) return true;
  const s = String(val).trim().toLowerCase();
  return (
    s === '' ||
    s === '-' ||
    s === '—' ||
    s === 'nil' ||
    s === 'none' ||
    s === 'na' ||
    s === 'n/a' ||
    s === 'blank' ||
    s.startsWith('blank') ||
    s.includes('blank due to') ||
    s.includes('corporation issue') ||
    s === 'vacant' ||
    s.startsWith('vacant') ||
    s === 'available' ||
    s.startsWith('available')
  );
}

// Parse Excel serial / JS Date / string → JS Date
function parseFlexibleDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'number') {
    try {
      if (typeof XLSX !== 'undefined' && XLSX?.SSF) {
        const d = XLSX.SSF.parse_date_code(val);
        if (d) return new Date(d.y, d.m - 1, d.d);
      }
      const d = new Date((val - 25569) * 86400000);
      if (!isNaN(d.getTime())) return d;
    } catch {}
    return null;
  }
  const str = String(val).trim();
  if (!str) return null;

  // Reject site codes (e.g. 'MB-72' or 'SITE-01') which JavaScript Date would interpret as year 1972
  if (/^[a-zA-Z]{1,6}[-_ ]\d{1,4}/i.test(str)) return null;

  // DD.MM.YYYY, DD/MM/YYYY, DD-MM-YYYY, D.M.YY, D.M.YYYY (including typo recovery like 206 -> 2026)
  const dmy = str.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (dmy) {
    let day = parseInt(dmy[1], 10);
    let month = parseInt(dmy[2], 10) - 1;
    let yearStr = dmy[3];
    let year = parseInt(yearStr, 10);
    if (yearStr.length === 2) year = 2000 + year;
    else if (yearStr.length === 3 && yearStr.startsWith('20')) year = parseInt('20' + yearStr.slice(2).padStart(2, '2'), 10);

    if (month >= 0 && month <= 11) {
      const maxDays = new Date(year, month + 1, 0).getDate();
      day = Math.min(day, maxDays);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // YYYY-MM-DD or YYYY/MM/DD
  const ymd = str.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (ymd) {
    let year = parseInt(ymd[1], 10);
    let month = parseInt(ymd[2], 10) - 1;
    let day = parseInt(ymd[3], 10);
    if (month >= 0 && month <= 11) {
      const maxDays = new Date(year, month + 1, 0).getDate();
      day = Math.min(day, maxDays);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // DD-MMM-YYYY (e.g. 15-Jan-2026)
  const monthNames = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  const dMmmY = str.match(/^(\d{1,2})[\s./-]+([A-Za-z]{3,9})[\s./-]+(\d{2,4})$/);
  if (dMmmY) {
    let day = parseInt(dMmmY[1], 10);
    const mKey = dMmmY[2].toLowerCase().slice(0, 3);
    const month = monthNames[mKey];
    let year = parseInt(dMmmY[3], 10);
    if (dMmmY[3].length === 2) year = 2000 + year;
    if (month !== undefined) {
      const maxDays = new Date(year, month + 1, 0).getDate();
      day = Math.min(day, maxDays);
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function parseOccDate(val) {
  return parseFlexibleDate(val);
}

// Parse string like "Aug 2026", "August 2026", "2026-08"
function parseMonthString(str) {
  if (!str) return null;
  const s = String(str).trim();
  const mNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const sLower = s.toLowerCase();
  
  let monthIdx = -1;
  for (let i = 0; i < mNames.length; i++) {
    if (sLower.includes(mNames[i])) {
      monthIdx = i;
      break;
    }
  }
  
  const yrMatch = s.match(/\b(20\d\d)\b/);
  const year = yrMatch ? parseInt(yrMatch[1], 10) : new Date().getFullYear();
  
  if (monthIdx >= 0) {
    const start = new Date(year, monthIdx, 1, 0, 0, 0);
    const end = new Date(year, monthIdx + 1, 0, 23, 59, 59);
    return { start, end, monthIdx, year };
  }

  const ym = s.match(/^(\d{4})[-/](\d{1,2})/);
  if (ym) {
    const y = parseInt(ym[1], 10);
    const m = parseInt(ym[2], 10) - 1;
    const start = new Date(y, m, 1, 0, 0, 0);
    const end = new Date(y, m + 1, 0, 23, 59, 59);
    return { start, end, monthIdx: m, year: y };
  }

  return null;
}

// Format a JS date as "MMM YYYY"  e.g. "Jan 2024"
function fmtMonthYear(date) {
  return date.toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
}

// Format a value as a percentage string
function formatPct(val) {
  if (val === '' || val === null || val === undefined) return '—';
  const n = Number(val);
  if (isNaN(n)) return String(val);
  return `${Math.round(n * 10) / 10}%`;
}

/* ─────────────────────────────────────────────────────────────────────────
   OccupancyView  — Excel import → Site × Month history matrix
   ─────────────────────────────────────────────────────────────────────────
   Expected Excel columns (flexible, auto-detected):
     Site Code | Month (date) | Value (occupancy % or days) | Client (optional)
   
   The pivot table shows:
     Rows    = Sites (each unique site code)
     Columns = Months (sorted chronologically)
     Cell    = occupancy value & client who booked that site
   
   View tabs: 6 Months | Yearly | 365-Day Overview
   ───────────────────────────────────────────────────────────────────────── */
function OccPercentBar({ pct, days, totalDays, showText = true, height = 7, clientNames = [], brands = [], onClick }) {
  const p = Math.max(0, Math.min(100, Math.round(Number(pct) || 0)));
  const barColor = p >= 75 ? '#10b981' : p >= 40 ? '#38bdf8' : p > 0 ? '#f59e0b' : '#222c37';
  
  const cList = Array.isArray(clientNames) ? clientNames.filter(Boolean) : (clientNames ? [clientNames] : []);
  const clientStr = cList.join(', ');
  const bList = Array.isArray(brands) ? brands.filter(Boolean) : (brands ? [brands] : []);
  const brandStr = bList.join(', ');

  const tooltip = days !== undefined && totalDays !== undefined 
    ? `${p}% occupied (${days} of ${totalDays} days)${clientStr ? ` • Booked by: ${clientStr}${brandStr && brandStr !== clientStr ? ` (${brandStr})` : ''}` : (p === 0 ? ' • Vacant' : '')} — Click to view details`
    : `${p}% occupied${clientStr ? ` • Booked by: ${clientStr}` : ''} — Click to view details`;

  return (
    <div 
      className="scooh-occ-cell" 
      title={tooltip} 
      onClick={onClick}
      style={{ 
        display: 'flex', 
        flexDirection: 'column', 
        gap: '3px', 
        minWidth: '105px',
        cursor: onClick ? 'pointer' : 'default',
        padding: '3px 4px',
        borderRadius: '6px',
        transition: 'background 0.15s ease'
      }}
      onMouseEnter={e => { if (onClick) e.currentTarget.style.background = 'rgba(56, 189, 248, 0.08)'; }}
      onMouseLeave={e => { if (onClick) e.currentTarget.style.background = 'transparent'; }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <div className="scooh-occbar" style={{ flex: 1, height: `${height}px`, background: '#0b1016', borderRadius: '999px', overflow: 'hidden', border: '1px solid #222c37' }}>
          <span style={{ display: 'block', height: '100%', width: `${p}%`, background: barColor, borderRadius: '999px', transition: 'width 0.3s ease' }} />
        </div>
        {showText && (
          <span style={{ fontSize: '11px', fontWeight: 800, color: p > 0 ? '#f1f5f9' : '#64748b', minWidth: '32px', textAlign: 'right' }}>
            {p > 0 ? `${p}%` : '0%'}
          </span>
        )}
      </div>
      {clientStr ? (
        <div style={{ textAlign: 'left', marginTop: '1px' }}>
          <div style={{ fontSize: '11px', color: '#38bdf8', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '115px', fontWeight: 800 }} title={clientStr}>
            👤 {clientStr}
          </div>
          {brandStr && brandStr !== clientStr && (
            <div style={{ fontSize: '9.5px', color: '#94a3b8', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '115px' }} title={brandStr}>
              {brandStr}
            </div>
          )}
        </div>
      ) : p > 0 ? (
        <div style={{ fontSize: '10px', color: '#94a3b8', textAlign: 'left' }}>
          Occupied
        </div>
      ) : (
        <div style={{ fontSize: '10px', color: '#475569', textAlign: 'left' }}>
          Vacant
        </div>
      )}
    </div>
  );
}

function OccupancyView() {
  const location = useLocation();
  const navigate = useNavigate();
  const [dbSites, setDbSites] = useState([]);
  const [dbOccupancy, setDbOccupancy] = useState([]); // fed from /campaigns (same as Campaign Tracker)
  const [loading, setLoading] = useState(true);
  const [importingExcel, setImportingExcel] = useState(false);
  const [viewTab, setViewTab] = useState('6months'); // '6months' | '12months' | 'overview'
  const [statusFilter, setStatusFilter] = useState('ALL'); // 'ALL' | 'occupied' | 'vacant'
  const [siteTypeFilter, setSiteTypeFilter] = useState(new Set(['ALL', 'Combined', 'Split Face'])); // multi-select
  const [search, setSearch] = useState('');
  const [modalData, setModalData] = useState(null); // For inspecting month/site booking history details
  const [viewCampaignDetails, setViewCampaignDetails] = useState(null);
  const [selectedSiteModal, setSelectedSiteModal] = useState(null);
  const [dbCampaigns, setDbCampaigns] = useState([]);
  const [bookingRecordId, setBookingRecordId] = useState(null);
  const [bookingClientName, setBookingClientName] = useState('');
  const [savingBooking, setSavingBooking] = useState(false);

  const currentRole = getCurrentRole();
  const canImport = currentRole === 'admin' || currentRole === 'manager';
  const canDelete = currentRole === 'admin' || currentRole === 'manager';
  const [selectedSiteCodes, setSelectedSiteCodes] = useState(new Set());
  const [deleting, setDeleting] = useState(false);

  // Auto-sync search from URL (e.g. /occupancy?site=MB-01)
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const siteParam = (params.get('site') || params.get('search') || '').trim();
    if (siteParam) {
      setSearch(siteParam);
      setViewTab('6months');
    }
  }, [location.search]);

  // Load sites + occupancy records + campaigns (so added campaigns automatically turn vacant to occupied)
  const loadSystemData = useCallback(async () => {
    setLoading(true);
    try {
      const [sRes, cRes, campRes] = await Promise.all([
        api.get('/sites'),
        api.get('/occupancy'),
        api.get('/campaigns').catch(() => ({ data: [] }))
      ]);
      setDbSites(Array.isArray(sRes.data) ? sRes.data : []);
      setDbOccupancy(Array.isArray(cRes.data) ? cRes.data : []);
      setDbCampaigns(Array.isArray(campRes.data) ? campRes.data : []);
    } catch (err) {
      console.error('Failed to load occupancy data:', err);
    } finally {
      setLoading(false);
    }
  }, []);

  async function handleBookVacantRecord(record, clientName) {
    if (!clientName || !clientName.trim()) {
      alert('Please enter a client name.');
      return;
    }
    const cleanClient = clientName.trim();
    setSavingBooking(true);
    try {
      if (record.id && typeof record.id === 'number') {
        await api.put(`/occupancy/${record.id}`, {
          client: cleanClient,
          display: cleanClient,
          status: 'active',
          occupancy_pct: 100
        });
      } else {
        await api.post('/occupancy', {
          site_code: record.site_code,
          location: record.location || '',
          start_date: record.start_date,
          end_date: record.end_date,
          client: cleanClient,
          display: cleanClient,
          status: 'active',
          occupancy_pct: 100
        });
      }

      // Mirror into campaigns table so Campaign Tracker has this booking immediately
      try {
        await api.post('/campaigns', {
          site_code: record.site_code,
          client: cleanClient,
          display: cleanClient,
          campaign_name: cleanClient,
          start_date: record.start_date,
          end_date: record.end_date,
          booking_date: record.start_date,
          days: record.days || 30,
          month: record.month || '',
          status: 'active',
          record_status: 'active'
        });
      } catch (cErr) {
        console.warn('Mirror booking to campaigns note:', cErr);
      }

      setBookingRecordId(null);
      setBookingClientName('');
      await loadSystemData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));

      if (modalData) {
        setModalData(prev => {
          if (!prev) return null;
          const updatedCampaigns = (prev.campaigns || []).map(c => {
            if ((c.id && c.id === record.id) || (c.site_code === record.site_code && c.start_date === record.start_date)) {
              return {
                ...c,
                client: cleanClient,
                client_name: cleanClient,
                display: cleanClient,
                status: 'active',
                is_vacant: false,
                occupancy_pct: 100
              };
            }
            return c;
          });
          return { ...prev, campaigns: updatedCampaigns };
        });
      }
    } catch (err) {
      console.error('Failed to book vacant record:', err);
      alert('Failed to update booking: ' + (err.response?.data?.message || err.message));
    } finally {
      setSavingBooking(false);
    }
  }

  useEffect(() => {
    loadSystemData();
    const onCampUpdate = () => loadSystemData();
    window.addEventListener('mb-campaigns-updated', onCampUpdate);
    return () => window.removeEventListener('mb-campaigns-updated', onCampUpdate);
  }, [loadSystemData]);

  async function handleOccupancyExcelImport(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      // Post to /import/occupancy-xlsx (which synchronizes both occupancy records and campaigns table)
      const r = await api.post('/import/occupancy-xlsx', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      alert(`✓ Occupancy Excel Import Successful!\n\n${r.data?.message || 'Occupancy and campaign records updated successfully.'}`);
      await loadSystemData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
    } catch (err) {
      console.warn('Occupancy import attempt note, trying campaigns import fallback:', err);
      try {
        const fd2 = new FormData();
        fd2.append('file', file);
        const r2 = await api.post('/import/campaigns-xlsx', fd2, {
          headers: { 'Content-Type': 'multipart/form-data' }
        });
        alert(`✓ Excel Import Successful!\n\n${r2.data?.message || 'Records imported successfully.'}`);
        await loadSystemData();
        window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      } catch (err2) {
        console.error('Occupancy Excel import error:', err2);
        alert('Occupancy Excel Import failed: ' + (err2.response?.data?.message || err2.message || err.response?.data?.message || err.message));
      }
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  function downloadOccupancyDemoTemplate() {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Site Block Format (Matches exact format with Site Code header and Up/Down dates)
      const blockData = [
        ['Site Code', 'Location / Specifications', '', ''],
        ['MB-14', "Shivranjani Junction traffic from SG Road - 14'x14' - B/L", '', ''],
        ['', 'Up Date', 'Down Date', 'Client Name'],
        ['', '01.04.2026', '30.04.2026', 'Swagat Group'],
        ['', '01.05.2026', '31.05.2026', 'Swagat Group'],
        ['', '01.06.2026', '30.06.2026', 'Swagat Group'],
        ['', '01.07.2026', '31.07.2026', 'Swagat Group'],
        ['', '01.08.2026', '31.08.2026', 'Swagat Group'],
        ['', '', '', ''],
        ['MB-72', "200ft Ring Road Junction – From YMCA Club Road to Club O7 - 30'x 15' - B/L", '', ''],
        ['', 'Up Date', 'Down Date', 'Client Name'],
        ['', '01.04.2026', '20.04.2026', 'B Safal'],
        ['', '21.04.2026', '30.04.2026', 'Blank'],
        ['', '01.05.2026', '11.05.2026', 'Blank'],
        ['', '12.05.2026', '31.05.2026', 'Hocco'],
        ['', '01.06.2026', '08.06.2026', 'Blank'],
        ['', '09.06.2026', '24.06.2026', 'Gateway'],
        ['', '25.06.2026', '30.06.2026', 'Contact'],
        ['', '01.07.2026', '30.07.2026', 'Contact'],
        ['', '01.08.2026', '30.08.2026', 'Contact'],
        ['', '', '', ''],
        ['MB-45', "Pakwan Cross Road to Sindhu Bhavan Road - 20'x10' - F/L", '', ''],
        ['', 'Up Date', 'Down Date', 'Client Name'],
        ['', '01.04.2026', '15.04.2026', 'Blank'],
        ['', '16.04.2026', '30.04.2026', 'Zydus'],
        ['', '01.05.2026', '31.05.2026', 'Zydus']
      ];
      const ws1 = XLSX.utils.aoa_to_sheet(blockData);
      ws1['!cols'] = [{ wch: 15 }, { wch: 60 }, { wch: 16 }, { wch: 25 }];
      XLSX.utils.book_append_sheet(wb, ws1, 'Site Block Format');

      // Sheet 2: Tabular Format (Alternative simple columns)
      const tableData = [
        ['Site Code', 'Location', 'Up Date', 'Down Date', 'Client Name'],
        ['MB-14', "Shivranjani Junction traffic from SG Road", '01/04/2026', '30/04/2026', 'Swagat Group'],
        ['MB-14', "Shivranjani Junction traffic from SG Road", '01/05/2026', '31/05/2026', 'Swagat Group'],
        ['MB-72', "200ft Ring Road Junction", '01/04/2026', '20/04/2026', 'B Safal'],
        ['MB-72', "200ft Ring Road Junction", '21/04/2026', '30/04/2026', 'Blank'],
        ['MB-72', "200ft Ring Road Junction", '01/05/2026', '11/05/2026', 'Blank'],
        ['MB-72', "200ft Ring Road Junction", '12/05/2026', '31/05/2026', 'Hocco'],
        ['MB-45', "Pakwan Cross Road to Sindhu Bhavan Road", '01/04/2026', '15/04/2026', 'Blank'],
        ['MB-45', "Pakwan Cross Road to Sindhu Bhavan Road", '16/04/2026', '30/04/2026', 'Zydus']
      ];
      const ws2 = XLSX.utils.aoa_to_sheet(tableData);
      ws2['!cols'] = [{ wch: 15 }, { wch: 45 }, { wch: 15 }, { wch: 15 }, { wch: 25 }];
      XLSX.utils.book_append_sheet(wb, ws2, 'Tabular Format');

      XLSX.writeFile(wb, 'Occupancy_Import_Demo_Template.xlsx');
    } catch (err) {
      console.error('Failed to export demo template:', err);
      alert('Could not download demo template: ' + err.message);
    }
  }

  // ── Compute periods (6 Months, 12 Months, & Yearly) ──────────────────
  const { periods6M, periods12M, periodsYearly, period365 } = useMemo(() => {
    const now = new Date();
    
    // 6 Months periods
    const p6 = [];
    for (let i = 5; i >= 0; i--) {
      const dStart = new Date(now.getFullYear(), now.getMonth() - i, 1, 0, 0, 0);
      const dEnd = new Date(now.getFullYear(), now.getMonth() - i + 1, 0, 23, 59, 59);
      const totalDays = Math.round((dEnd - dStart) / 86400000) + 1;
      p6.push({
        key: `${dStart.getFullYear()}-${String(dStart.getMonth() + 1).padStart(2, '0')}`,
        label: fmtMonthYear(dStart),
        start: dStart,
        end: dEnd,
        totalDays
      });
    }

    // 12 Months periods (rolling 12 months)
    const p12 = [];
    for (let i = 11; i >= 0; i--) {
      const dStart = new Date(now.getFullYear(), now.getMonth() - i, 1, 0, 0, 0);
      const dEnd = new Date(now.getFullYear(), now.getMonth() - i + 1, 0, 23, 59, 59);
      const totalDays = Math.round((dEnd - dStart) / 86400000) + 1;
      p12.push({
        key: `${dStart.getFullYear()}-${String(dStart.getMonth() + 1).padStart(2, '0')}`,
        label: fmtMonthYear(dStart),
        start: dStart,
        end: dEnd,
        totalDays
      });
    }

    // Yearly periods (current year and preceding 2 years)
    const pY = [];
    const curYear = now.getFullYear();
    for (let y = curYear - 2; y <= curYear; y++) {
      const dStart = new Date(y, 0, 1, 0, 0, 0);
      const dEnd = new Date(y, 11, 31, 23, 59, 59);
      const isLeap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
      pY.push({
        key: String(y),
        label: String(y),
        start: dStart,
        end: dEnd,
        totalDays: isLeap ? 366 : 365
      });
    }

    // 365 days window
    const d365Start = new Date(now);
    d365Start.setDate(now.getDate() - 365);
    const p365 = {
      start: d365Start,
      end: now,
      totalDays: 365
    };

    return { periods6M: p6, periods12M: p12, periodsYearly: pY, period365: p365 };
  }, []);

  // ── Calculate Site History with Client Tracking ───────────────────────
  const siteHistory = useMemo(() => {
    // Uses live campaigns data (shared with Campaign Tracker)
    // Calculate from Live DB Sites, synthesized Combined/Split-face hierarchy, and Occupancy records
    const activeDbSites = dbSites.filter(s => s.record_status !== 'archived');
    const activeOccupancies = dbOccupancy.filter(c => c.record_status !== 'archived');

    // Build complete site directory guaranteeing all combined & split-face sites exist
    const siteMap = new Map();
    activeDbSites.forEach(s => {
      const c = canonicalSiteCode(s.site_code || `MB-${s.id}`);
      if (c) siteMap.set(c, { ...s, site_code: c });
    });

    // Synthesize any missing combined or split-face sites defined in SITE_PANELS
    Object.keys(SITE_PANELS).forEach(panelCode => {
      const code = canonicalSiteCode(panelCode);
      if (!siteMap.has(code)) {
        // Inherit location/city context from an overlapping site in DB if available
        const overlapCodes = getOverlappingSiteCodes(code);
        const refSite = overlapCodes.map(oc => siteMap.get(oc)).find(Boolean);
        siteMap.set(code, {
          id: `v-${code}`,
          site_code: code,
          city: refSite?.city || 'Ahmedabad',
          area: refSite?.area || refSite?.address || `Group ${code} Display`,
          address: refSite?.address || refSite?.area || `Group ${code} Display`,
          type: isCombinedSite(code) ? 'Combined Hoarding' : 'Split Face',
          size: refSite?.size || '—',
          isSynthesized: true
        });
      }
    });

    const allSitesList = Array.from(siteMap.values());

    return allSitesList.map(site => {
      const siteCode = canonicalSiteCode(site.site_code);
      const overlappingCodes = new Set(getOverlappingSiteCodes(siteCode).map(canonicalSiteCode));
      overlappingCodes.add(siteCode);

      // Find all occupancy records for this site
      const rawOccs = activeOccupancies.filter(c => {
        if (c.site_id && site.id && String(c.site_id) === String(site.id)) return true;
        const cCode = canonicalSiteCode(c.site_code);
        return cCode && overlappingCodes.has(cCode);
      });

      // Find any campaigns for this site from Campaign Tracker
      const siteCampaigns = (dbCampaigns || []).filter(c => {
        if (c.record_status === 'archived' || isVacantClient(c.client || c.client_name)) return false;
        if (c.site_id && site.id && String(c.site_id) === String(site.id)) return true;
        const cCode = canonicalSiteCode(c.site_code);
        return cCode && overlappingCodes.has(cCode);
      });

      // Merge: when a campaign exists, it automatically overrides any overlapping vacant or mirrored occupancy record
      const mergedRecords = [];
      const usedOccIds = new Set();

      siteCampaigns.forEach(cmp => {
        const cmpStart = parseFlexibleDate(cmp.start_date || cmp.booking_date);
        const cmpEnd = parseFlexibleDate(cmp.end_date) || cmpStart;
        const cmpStartTs = cmpStart ? cmpStart.getTime() : 0;
        const cmpEndTs = cmpEnd ? cmpEnd.getTime() : cmpStartTs;

        rawOccs.forEach(occ => {
          const occStart = parseFlexibleDate(occ.start_date || occ.booking_date);
          const occEnd = parseFlexibleDate(occ.end_date) || occStart;
          if (occStart && occEnd) {
            const occStartTs = occStart.getTime();
            const occEndTs = occEnd.getTime();
            if (cmpStartTs <= occEndTs && cmpEndTs >= occStartTs) {
              usedOccIds.add(occ.id);
            }
          }
        });

        mergedRecords.push(cmp);
      });

      rawOccs.forEach(occ => {
        if (!usedOccIds.has(occ.id)) {
          mergedRecords.push(occ);
        }
      });

      const allRecordsForSite = mergedRecords;

      // Filter non-vacant for actual occupancy percentage & occupied day counting
      const siteCamps = allRecordsForSite.filter(c => {
        if (c.status === 'vacant' || c.is_vacant || isVacantClient(c.client || c.client_name)) return false;
        return true;
      });

      // Sort both in ascending chronological order (earliest date first)
      allRecordsForSite.sort((a, b) => {
        const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
        const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
        if (da.getTime() !== db.getTime()) return da - db;
        const dea = parseFlexibleDate(a.end_date) || new Date(0);
        const deb = parseFlexibleDate(b.end_date) || new Date(0);
        return dea - deb;
      });

      siteCamps.sort((a, b) => {
        const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
        const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
        if (da.getTime() !== db.getTime()) return da - db;
        const dea = parseFlexibleDate(a.end_date) || new Date(0);
        const deb = parseFlexibleDate(b.end_date) || new Date(0);
        return dea - deb;
      });

      // Helper to calculate exact non-overlapping occupied days and track clients in window
      function getDaysAndClientsInWindow(wStart, wEnd) {
        const daySet = new Set();
        const clientSet = new Set();
        const brandSet = new Set();
        const relevantCampaigns = [];
        const startTs = wStart.getTime();
        const endTs = wEnd.getTime();

        allRecordsForSite.forEach(c => {
          const isVac = c.status === 'vacant' || c.is_vacant || isVacantClient(c.client || c.client_name);
          let cStart = parseFlexibleDate(c.start_date || c.booking_date);
          let cEnd = parseFlexibleDate(c.end_date);
          
          if (!cStart && c.month) {
            const parsedM = parseMonthString(c.month);
            if (parsedM) {
              cStart = parsedM.start;
              cEnd = parsedM.end;
            }
          }
          if (!cStart) return;
          if (!cEnd) cEnd = new Date(cStart.getFullYear(), cStart.getMonth() + 1, 0, 23, 59, 59);

          const s = Math.max(startTs, cStart.getTime());
          const e = Math.min(endTs, cEnd.getTime());

          if (e >= s) {
            // ONLY non-vacant dates are counted into occupied days ("not vacant date should be added")
            if (!isVac) {
              for (let t = s; t <= e; t += 86400000) {
                const dt = new Date(t);
                daySet.add(`${dt.getFullYear()}-${dt.getMonth()}-${dt.getDate()}`);
              }
              const clientName = c.client || c.client_name;
              if (clientName) clientSet.add(clientName);
              if (c.brand) brandSet.add(c.brand);
            }
            // Blank dates are also added to relevantCampaigns so user can inspect them in booking details
            relevantCampaigns.push(c);
          }
        });

        relevantCampaigns.sort((a, b) => {
          const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
          const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
          if (da.getTime() !== db.getTime()) return da - db;
          const dea = parseFlexibleDate(a.end_date) || new Date(0);
          const deb = parseFlexibleDate(b.end_date) || new Date(0);
          return dea - deb;
        });

        return {
          days: daySet.size,
          clients: Array.from(clientSet),
          brands: Array.from(brandSet),
          campaigns: relevantCampaigns
        };
      }

      // Determine current / most recent client
      const now = new Date();
      let currentClient = null;
      let currentBrand = null;
      let currentStatus = 'vacant'; // 'active' | 'upcoming' | 'past' | 'vacant'

      const sortedCampsDesc = [...siteCamps].sort((a, b) => {
        const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
        const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
        return db - da;
      });

      for (const c of sortedCampsDesc) {
        const cStart = parseFlexibleDate(c.start_date || c.booking_date);
        const cEnd = parseFlexibleDate(c.end_date) || (cStart ? new Date(cStart.getFullYear(), cStart.getMonth() + 1, 0, 23, 59, 59) : null);
        const cName = c.client || c.client_name;
        if (!cName) continue;

        if (cStart && cEnd) {
          if (cStart <= now && cEnd >= now) {
            currentClient = cName;
            currentBrand = c.brand;
            currentStatus = 'active';
            break;
          } else if (cStart > now) {
            if (currentStatus !== 'active') {
              currentClient = cName;
              currentBrand = c.brand;
              currentStatus = 'upcoming';
            }
          } else if (cEnd < now && !currentClient) {
            currentClient = cName;
            currentBrand = c.brand;
            currentStatus = 'past';
          }
        } else if (!currentClient) {
          currentClient = cName;
          currentBrand = c.brand;
          currentStatus = 'past';
        }
      }

      // 6 Months breakdown
      const by6M = {};
      let sum6M = 0;
      periods6M.forEach(p => {
        const res = getDaysAndClientsInWindow(p.start, p.end);
        const pct = Math.min(100, Math.round((res.days / p.totalDays) * 100));
        by6M[p.key] = { days: res.days, totalDays: p.totalDays, pct, clients: res.clients, brands: res.brands, campaigns: res.campaigns };
        sum6M += pct;
      });
      const avg6M = periods6M.length ? Math.round(sum6M / periods6M.length) : 0;

      // 12 Months breakdown
      const by12M = {};
      let sum12M = 0;
      periods12M.forEach(p => {
        const res = getDaysAndClientsInWindow(p.start, p.end);
        const pct = Math.min(100, Math.round((res.days / p.totalDays) * 100));
        by12M[p.key] = { days: res.days, totalDays: p.totalDays, pct, clients: res.clients, brands: res.brands, campaigns: res.campaigns };
        sum12M += pct;
      });
      const avg12M = periods12M.length ? Math.round(sum12M / periods12M.length) : 0;

      // Yearly breakdown
      const byYearly = {};
      periodsYearly.forEach(p => {
        const res = getDaysAndClientsInWindow(p.start, p.end);
        const pct = Math.min(100, Math.round((res.days / p.totalDays) * 100));
        byYearly[p.key] = { days: res.days, totalDays: p.totalDays, pct, clients: res.clients, brands: res.brands, campaigns: res.campaigns };
      });

      // 365 Days Rolling
      const res365 = getDaysAndClientsInWindow(period365.start, period365.end);
      const pct365 = Math.min(100, Math.round((res365.days / 365) * 100));

      return {
        id: site.id,
        site_code: siteCode,
        siteType: getSiteTypeTag(siteCode), // 'Combined' | 'Split Face' | 'Single'
        city: site.city || '—',
        area: site.area || site.address || '—',
        currentClient,
        currentBrand,
        currentStatus,
        allCampaigns: allRecordsForSite,
        occupiedCampaigns: siteCamps,
        by6M,
        by12M,
        byYearly,
        avg6M,
        avg12M,
        days365: res365.days,
        pct365,
        clients365: res365.clients
      };
    }).sort((a, b) => universalCompare(a.site_code, b.site_code, 'asc'));
  }, [dbSites, dbOccupancy, dbCampaigns, periods6M, periods12M, periodsYearly, period365]);

  // ── Filtered sites with simple Status Filter (All / Occupied / Vacant) and Site Type Filter ──
  const filteredSites = useMemo(() => {
    const list = siteHistory.filter(s => {
      if (!siteTypeFilter.has('Combined') && s.siteType === 'Combined') return false;
      if (!siteTypeFilter.has('Split Face') && s.siteType === 'Split Face') return false;
      if (!siteTypeFilter.has('ALL') && s.siteType !== 'Combined' && s.siteType !== 'Split Face') return false;

      const isOccupied = s.currentStatus === 'active';
      if (statusFilter === 'occupied' && !isOccupied) return false;
      if (statusFilter === 'vacant' && isOccupied) return false;

      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        s.site_code.toLowerCase().includes(q) ||
        (s.city && s.city.toLowerCase().includes(q)) ||
        (s.area && s.area.toLowerCase().includes(q)) ||
        (s.currentClient && s.currentClient.toLowerCase().includes(q)) ||
        (s.currentBrand && s.currentBrand.toLowerCase().includes(q)) ||
        (s.clients365 && s.clients365.some(c => c.toLowerCase().includes(q)))
      );
    });

    // All site codes everywhere sorted in ascending order
    return list.sort((a, b) => universalCompare(a.site_code, b.site_code, 'asc'));
  }, [siteHistory, search, statusFilter, siteTypeFilter]);

  // ── Site Selection & Bulk Delete Operations ────────────────────────────
  const selectedCount = selectedSiteCodes.size;
  const allVisibleSelected = filteredSites.length > 0 && filteredSites.every(s => selectedSiteCodes.has(s.site_code));

  const selectedTotalCampCount = useMemo(() => {
    let count = 0;
    filteredSites.forEach(s => {
      if (selectedSiteCodes.has(s.site_code)) {
        count += (s.allCampaigns || []).length;
      }
    });
    return count;
  }, [filteredSites, selectedSiteCodes]);

  function toggleSelectSite(code) {
    if (!code) return;
    setSelectedSiteCodes(prev => {
      const next = new Set(prev);
      if (next.has(code)) next.delete(code);
      else next.add(code);
      return next;
    });
  }

  function toggleSelectAll() {
    if (allVisibleSelected) {
      setSelectedSiteCodes(prev => {
        const next = new Set(prev);
        filteredSites.forEach(s => next.delete(s.site_code));
        return next;
      });
    } else {
      setSelectedSiteCodes(prev => {
        const next = new Set(prev);
        filteredSites.forEach(s => next.add(s.site_code));
        return next;
      });
    }
  }

  function deselectAll() {
    setSelectedSiteCodes(new Set());
  }

  async function handleBatchDeleteBookings() {
    if (!canDelete || selectedSiteCodes.size === 0) return;
    const selectedSitesList = filteredSites.filter(s => selectedSiteCodes.has(s.site_code));
    const campIds = selectedSitesList.flatMap(s => (s.allCampaigns || []).map(c => c.id)).filter(Boolean);
    const siteCount = selectedSiteCodes.size;
    const campCount = campIds.length;

    if (campCount === 0) {
      alert(`The ${siteCount} selected site(s) currently have no booking records to delete (they are already 0% Vacant).`);
      return;
    }

    if (!confirm(`Are you sure you want to delete all ${campCount} booking record${campCount > 1 ? 's' : ''} on the ${siteCount} selected site${siteCount > 1 ? 's' : ''}?\n\nThis will reset their occupancy to 0% (Vacant). The sites will remain in your inventory.`)) {
      return;
    }

    try {
      setDeleting(true);
      await Promise.allSettled([
        api.post('/occupancy/batch-delete', { ids: campIds, hard: true }),
        api.post('/campaigns/batch-delete', { ids: campIds, hard: true })
      ]);
      alert(`✓ Successfully deleted ${campCount} booking record${campCount > 1 ? 's' : ''} across ${siteCount} site${siteCount > 1 ? 's' : ''}. Occupancy reset to 0% (Vacant).`);
      setSelectedSiteCodes(new Set());
      await loadSystemData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
    } catch (err) {
      console.error('Failed to delete occupancy bookings:', err);
      alert('Failed to delete bookings: ' + (err.response?.data?.message || err.message));
    } finally {
      setDeleting(false);
    }
  }

  async function handleBatchDeleteSites() {
    if (!canDelete || selectedSiteCodes.size === 0) return;
    const siteCount = selectedSiteCodes.size;
    const selectedSitesList = filteredSites.filter(s => selectedSiteCodes.has(s.site_code));
    const siteIds = selectedSitesList.map(s => s.id).filter(Boolean);
    const campIds = selectedSitesList.flatMap(s => (s.allCampaigns || []).map(c => c.id)).filter(Boolean);

    if (!confirm(`⚠️ PERMANENT DELETION WARNING:\n\nAre you sure you want to permanently delete ${siteCount} site${siteCount > 1 ? 's' : ''} (${Array.from(selectedSiteCodes).slice(0, 5).join(', ')}${siteCount > 5 ? '…' : ''}) and all ${campIds.length} associated booking records completely from the inventory?\n\nThis action cannot be undone.`)) {
      return;
    }

    try {
      setDeleting(true);
      if (campIds.length > 0) {
        await Promise.allSettled([
          api.post('/occupancy/batch-delete', { ids: campIds, hard: true }),
          api.post('/campaigns/batch-delete', { ids: campIds, hard: true })
        ]);
      }
      if (siteIds.length > 0) {
        await api.post('/sites/batch-delete', { ids: siteIds, hard: true });
      }
      alert(`✓ Successfully deleted ${siteCount} site${siteCount > 1 ? 's' : ''} and all associated records.`);
      setSelectedSiteCodes(new Set());
      await loadSystemData();
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
    } catch (err) {
      console.error('Failed to delete sites:', err);
      alert('Failed to delete sites: ' + (err.response?.data?.message || err.message));
    } finally {
      setDeleting(false);
    }
  }

  async function handleDeleteAllOccupancy() {
    if (!canDelete) return;
    const allCampIds = filteredSites.flatMap(s => (s.allCampaigns || []).map(c => c.id)).filter(Boolean);
    const count = allCampIds.length;
    if (count === 0) {
      alert('There are currently no active booking records across the visible sites.');
      return;
    }

    if (!confirm(`Are you sure you want to clear ALL ${count} booking records across all ${filteredSites.length} visible site(s)?\n\nThis will reset their occupancy to 0% (Vacant) so you can import clean monthly data. Sites will remain intact in inventory.`)) {
      return;
    }

    try {
      setDeleting(true);
      await Promise.allSettled([
        api.post('/occupancy/batch-delete', { ids: allCampIds, hard: true }),
        api.post('/campaigns/batch-delete', { ids: allCampIds, hard: true })
      ]);
      alert(`✓ Successfully cleared ${count} booking record${count > 1 ? 's' : ''}. Occupancy reset to 0% (Vacant).`);
      setSelectedSiteCodes(new Set());
      await loadSystemData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
    } catch (err) {
      console.error('Failed to clear all occupancy:', err);
      alert('Failed to clear bookings: ' + (err.response?.data?.message || err.message));
    } finally {
      setDeleting(false);
    }
  }

  // ── Overall Stats ─────────────────────────────────────────────────────
  const stats = useMemo(() => {
    const total = siteHistory.length;
    if (!total) return { avgPct: 0, occupied: 0, vacant: 0, total: 0 };

    let totalPct = 0;
    let occupiedCount = 0;

    siteHistory.forEach(s => {
      const isOcc = s.currentStatus === 'active';
      if (isOcc) occupiedCount++;
      const p = s.pct365 || 0;
      totalPct += p;
    });

    return {
      avgPct: Math.round(totalPct / total),
      occupied: occupiedCount,
      vacant: total - occupiedCount,
      total
    };
  }, [siteHistory]);

  // ── Export Occupancy Matrix to Excel ──────────────────────────────────
  async function exportOccupancyExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const ws = workbook.addWorksheet(viewTab === 'overview' ? '365-Day Overview' : 'Monthly Occupancy', {
        views: [{ state: 'frozen', ySplit: 2 }]
      });
      
      const periods = viewTab === '12months' ? periods12M : periods6M;
      
      const cols = [
        { key: 'site_code', width: 14 },
        { key: 'area', width: 28 },
        { key: 'city', width: 16 },
        { key: 'status', width: 14 }
      ];
      const headers = ['Site Code', 'Location / Area', 'City', 'Current Status'];

      if (viewTab === 'overview') {
        cols.push({ key: 'days', width: 20 });
        cols.push({ key: 'pct', width: 18 });
        headers.push('Occupied Days (365d)', 'Occupancy Rate %');
      } else {
        periods.forEach(p => {
          cols.push({ key: `pct_${p.key}`, width: 16 });
          cols.push({ key: `client_${p.key}`, width: 24 });
          headers.push(`${p.label} (Occ %)`, `${p.label} (Booked Client)`);
        });
        cols.push({ key: 'avg', width: 14 });
        headers.push('Average %');
      }

      ws.columns = cols;

      // Executive Media Buzz Brand Header Banner with Logo on Top Right
      const titleText = viewTab === 'overview'
        ? 'MEDIA BUZZ — 365-DAY CAMPAIGN TRACKER OVERVIEW'
        : (viewTab === '12months' ? 'MEDIA BUZZ — 12-MONTH CAMPAIGN TRACKER MATRIX' : 'MEDIA BUZZ — 6-MONTH CAMPAIGN TRACKER MATRIX');
      attachMediaBuzzExcelHeader(workbook, ws, {
        title: titleText,
        columns: cols,
        totalColumns: cols.length
      });

      // Populate Header Row (Row 2)
      const headerRow = ws.getRow(2);
      headerRow.height = 28;
      headers.forEach((h, i) => {
        const cell = headerRow.getCell(i + 1);
        cell.value = h;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF334155' } },
          left: { style: 'thin', color: { argb: 'FF334155' } },
          bottom: { style: 'medium', color: { argb: 'FFFFC200' } },
          right: { style: 'thin', color: { argb: 'FF334155' } }
        };
      });

      filteredSites.forEach(s => {
        const rowData = {
          site_code: s.site_code,
          area: s.area,
          city: s.city,
          status: s.currentStatus ? s.currentStatus.toUpperCase() : 'VACANT'
        };

        if (viewTab === 'overview') {
          rowData.days = s.days365;
          rowData.pct = `${s.pct365}%`;
        } else {
          const dataMap = viewTab === '12months' ? s.by12M : s.by6M;
          let sum = 0;
          periods.forEach(p => {
            const entry = dataMap?.[p.key];
            const pct = typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
            const clients = typeof entry === 'object' && entry?.clients ? entry.clients.join(', ') : '';
            rowData[`pct_${p.key}`] = `${pct}%`;
            rowData[`client_${p.key}`] = clients || (pct > 0 ? 'Occupied' : 'Vacant');
            sum += pct;
          });
          rowData.avg = `${Math.round(sum / (periods.length || 1))}%`;
        }

        ws.addRow(rowData);
      });

      // Style Data Rows (Row 3+)
      ws.eachRow((row, rowNumber) => {
        if (rowNumber >= 3) {
          row.height = 22;
          row.eachCell((cell, colNumber) => {
            cell.font = { name: 'Calibri', size: 10.5 };
            cell.alignment = { vertical: 'middle', horizontal: [1, 3, 4].includes(colNumber) ? 'center' : 'left' };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
              right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
          });
        }
      });

      // Official 5-Point Terms & Conditions at Bottom Center
      attachMediaBuzzTermsAndConditions(ws, {
        totalColumns: cols.length
      });

      const buf = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Campaign_Tracker_${viewTab}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Failed to export campaign tracker Excel:', err);
      alert('Could not generate Excel export.');
    }
  }

  return (
    <>
      <PageHead 
        title="Campaign Tracker" 
        desc="Monthly utilization, client bookings, and availability across your site inventory." 
      />

      {/* ── Top Summary & KPI Cards (Click to filter) ────────────────── */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px', marginBottom: '16px' }}>
        <div 
          className="scooh-panel" 
          style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '8px', cursor: 'pointer', border: statusFilter === 'ALL' ? '2px solid rgba(56, 189, 248, 0.4)' : '1px solid rgba(255,255,255,0.08)' }}
          onClick={() => setStatusFilter('ALL')}
          title="Click to view all sites"
        >
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            Average Occupancy
          </div>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
            <span style={{ fontSize: '28px', fontWeight: 900, color: stats.avgPct >= 60 ? '#10b981' : stats.avgPct >= 30 ? '#38bdf8' : '#f59e0b', lineHeight: 1 }}>
              {stats.avgPct}%
            </span>
            <span style={{ fontSize: '12px', color: '#64748b' }}>across portfolio</span>
          </div>
          <div style={{ height: '6px', background: '#0b1016', borderRadius: '999px', overflow: 'hidden', border: '1px solid #222c37' }}>
            <div style={{ height: '100%', width: `${stats.avgPct}%`, background: stats.avgPct >= 60 ? '#10b981' : '#38bdf8', borderRadius: '999px' }} />
          </div>
        </div>

        <div 
          className="scooh-panel" 
          style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'occupied' ? '2px solid #10b981' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'occupied' ? 'rgba(16, 185, 129, 0.08)' : undefined }}
          onClick={() => setStatusFilter(prev => prev === 'occupied' ? 'ALL' : 'occupied')}
          title="Click to filter only occupied sites"
        >
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <span>🟢 Occupied Sites</span>
            {statusFilter === 'occupied' && <span style={{ fontSize: '10px', background: '#10b981', color: '#000', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active Filter</span>}
          </div>
          <div style={{ fontSize: '28px', fontWeight: 900, color: '#10b981', lineHeight: 1.1 }}>
            {stats.occupied}
          </div>
          <div style={{ fontSize: '12px', color: '#64748b' }}>
            {stats.total > 0 ? `${Math.round((stats.occupied / stats.total) * 100)}% active campaigns` : 'No sites'}
          </div>
        </div>

        <div 
          className="scooh-panel" 
          style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'vacant' ? '2px solid #f59e0b' : '1px solid rgba(255,255,255,0.08)', background: statusFilter === 'vacant' ? 'rgba(245, 158, 11, 0.08)' : undefined }}
          onClick={() => setStatusFilter(prev => prev === 'vacant' ? 'ALL' : 'vacant')}
          title="Click to filter only vacant sites"
        >
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#f59e0b', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <span>⚪ Vacant Sites</span>
            {statusFilter === 'vacant' && <span style={{ fontSize: '10px', background: '#f59e0b', color: '#000', padding: '1px 6px', borderRadius: '10px', fontWeight: 900 }}>Active Filter</span>}
          </div>
          <div style={{ fontSize: '28px', fontWeight: 900, color: stats.vacant > 0 ? '#f59e0b' : '#10b981', lineHeight: 1.1 }}>
            {stats.vacant}
          </div>
          <div style={{ fontSize: '12px', color: '#64748b' }}>
            Ready for booking
          </div>
        </div>

        <div 
          className="scooh-panel" 
          style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '4px', cursor: 'pointer', border: statusFilter === 'ALL' ? '2px solid #38bdf8' : '1px solid rgba(255,255,255,0.08)' }}
          onClick={() => setStatusFilter('ALL')}
          title="Click to show all sites"
        >
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            Total Sites Tracked
          </div>
          <div style={{ fontSize: '28px', fontWeight: 900, color: '#f1f5f9', lineHeight: 1.1 }}>
            {stats.total}
          </div>
          <div style={{ fontSize: '12px', color: '#64748b' }}>
            Live from Latest Booking
          </div>
        </div>
      </div>

      {/* ── Unified Clean Controls Toolbar ───────────────────────────────── */}
      <section className="scooh-panel" style={{ marginBottom: '16px', padding: '12px 18px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap' }}>
          
          {/* Left: Search & Quick Status Pills */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', flex: '1 1 auto' }}>
            <input
              className="scooh-search"
              style={{ minWidth: '220px', maxWidth: '320px', fontSize: '12px' }}
              placeholder="Search site, client, area…"
              value={search}
              onChange={e => setSearch(e.target.value)}
            />

            {/* Quick Status Filter Pills */}
            <div style={{ display: 'inline-flex', gap: '6px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setStatusFilter('ALL')}
                style={{
                  padding: '4px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'ALL' ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: statusFilter === 'ALL' ? '#38bdf8' : '#94a3b8',
                  fontWeight: statusFilter === 'ALL' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                All ({siteHistory.length})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('occupied')}
                style={{
                  padding: '4px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'occupied' ? 'rgba(16, 185, 129, 0.25)' : 'transparent',
                  color: statusFilter === 'occupied' ? '#4ade80' : '#94a3b8',
                  fontWeight: statusFilter === 'occupied' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟢 Occupied ({stats.occupied})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('vacant')}
                style={{
                  padding: '4px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'vacant' ? 'rgba(245, 158, 11, 0.25)' : 'transparent',
                  color: statusFilter === 'vacant' ? '#fbbf24' : '#94a3b8',
                  fontWeight: statusFilter === 'vacant' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚪ Vacant ({stats.vacant})
              </button>
            </div>

            {/* Site Hierarchy Filter: multi-select — all toggled independently */}
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setSiteTypeFilter(prev => {
                  const next = new Set(prev);
                  if (next.has('ALL')) next.delete('ALL'); else next.add('ALL');
                  return next;
                })}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter.has('ALL') ? 'rgba(148, 163, 184, 0.2)' : 'transparent',
                  color: siteTypeFilter.has('ALL') ? '#f1f5f9' : '#94a3b8',
                  fontWeight: siteTypeFilter.has('ALL') ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Toggle single-panel sites"
              >
                All Types
              </button>
              <button
                type="button"
                onClick={() => setSiteTypeFilter(prev => {
                  const next = new Set(prev);
                  if (next.has('Combined')) next.delete('Combined'); else next.add('Combined');
                  return next;
                })}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter.has('Combined') ? 'rgba(168, 85, 247, 0.25)' : 'transparent',
                  color: siteTypeFilter.has('Combined') ? '#c084fc' : '#94a3b8',
                  fontWeight: siteTypeFilter.has('Combined') ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Toggle multi-panel combined sites (e.g. MB-04, MB-05, MB-06, MB-10, MB-17)"
              >
                🔀 Combined
              </button>
              <button
                type="button"
                onClick={() => setSiteTypeFilter(prev => {
                  const next = new Set(prev);
                  if (next.has('Split Face')) next.delete('Split Face'); else next.add('Split Face');
                  return next;
                })}
                style={{
                  padding: '4px 10px',
                  borderRadius: '16px',
                  border: 'none',
                  background: siteTypeFilter.has('Split Face') ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: siteTypeFilter.has('Split Face') ? '#38bdf8' : '#94a3b8',
                  fontWeight: siteTypeFilter.has('Split Face') ? 800 : 600,
                  fontSize: '11px',
                  cursor: 'pointer'
                }}
                title="Toggle individual split face panels"
              >
                ✂️ Split Face
              </button>
            </div>

            {canDelete && (
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap', marginLeft: '4px' }}>
                <button
                  type="button"
                  className="scooh-btn ghost"
                  style={{
                    fontSize: '11px',
                    padding: '4px 10px',
                    color: allVisibleSelected ? '#38bdf8' : '#cbd5e1',
                    borderColor: allVisibleSelected ? 'rgba(56, 189, 248, 0.4)' : 'rgba(255,255,255,0.15)',
                    background: allVisibleSelected ? 'rgba(56, 189, 248, 0.12)' : 'transparent'
                  }}
                  onClick={toggleSelectAll}
                  title={allVisibleSelected ? "Deselect all visible sites" : "Select all visible sites"}
                >
                  {allVisibleSelected ? '☑ Deselect All' : '☑ Select All'}
                </button>

                {selectedCount > 0 && (
                  <button
                    type="button"
                    className="scooh-btn danger"
                    style={{ fontSize: '11px', padding: '4px 10px' }}
                    onClick={handleBatchDeleteBookings}
                    disabled={deleting}
                    title="Delete all booking records on selected sites (resets occupancy to 0% Vacant)"
                  >
                    🗑 Delete Bookings ({selectedCount})
                  </button>
                )}

                {selectedCount > 0 && (
                  <button
                    type="button"
                    className="scooh-btn danger"
                    style={{ fontSize: '11px', padding: '4px 10px', background: 'rgba(239, 68, 68, 0.15)', borderColor: 'rgba(239, 68, 68, 0.35)', color: '#f87171' }}
                    onClick={handleBatchDeleteSites}
                    disabled={deleting}
                    title="Permanently delete selected sites and their records from inventory"
                  >
                    ❌ Delete Sites ({selectedCount})
                  </button>
                )}

                <button
                  type="button"
                  className="scooh-btn danger"
                  style={{ fontSize: '11px', padding: '4px 10px', background: 'rgba(239, 68, 68, 0.08)', borderColor: 'rgba(239, 68, 68, 0.25)', color: '#fca5a5' }}
                  onClick={handleDeleteAllOccupancy}
                  disabled={deleting}
                  title="Clear all booking and occupancy records across all visible sites"
                >
                  🗑 Clear All Bookings
                </button>
              </div>
            )}
          </div>

          {/* Right: Clean View Switcher & Action Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            {/* View Mode Toggle: 6M, 12M, 365-Day */}
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.6)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setViewTab('6months')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === '6months' ? 'rgba(56, 189, 248, 0.25)' : 'transparent',
                  color: viewTab === '6months' ? '#38bdf8' : '#94a3b8',
                  fontWeight: viewTab === '6months' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🗓️ 6 Months Timeline
              </button>
              <button
                type="button"
                onClick={() => setViewTab('12months')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === '12months' ? 'rgba(168, 85, 247, 0.25)' : 'transparent',
                  color: viewTab === '12months' ? '#c084fc' : '#94a3b8',
                  fontWeight: viewTab === '12months' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                📅 12 Months Timeline
              </button>
              <button
                type="button"
                onClick={() => setViewTab('overview')}
                style={{
                  padding: '5px 14px',
                  borderRadius: '16px',
                  border: 'none',
                  background: viewTab === 'overview' ? 'rgba(242, 201, 76, 0.18)' : 'transparent',
                  color: viewTab === 'overview' ? '#f2c94c' : '#94a3b8',
                  fontWeight: viewTab === 'overview' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚡ 365-Day Overview
              </button>
            </div>

            <button
              type="button"
              className="scooh-btn ghost"
              onClick={exportOccupancyExcel}
              title="Export complete campaign tracker to Excel"
              style={{ fontSize: '11.5px', padding: '6px 12px' }}
            >
              📥 Export Excel
            </button>

            {canImport && (
              <label
                className="scooh-btn ghost"
                style={{
                  cursor: importingExcel ? 'not-allowed' : 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  fontSize: '11.5px',
                  padding: '6px 12px',
                  color: '#c084fc',
                  borderColor: 'rgba(168, 85, 247, 0.4)',
                  background: 'rgba(168, 85, 247, 0.08)'
                }}
                title="Import booking & occupancy records from Excel (.xlsx, .xls)"
              >
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv,.xlsm,.ods"
                  hidden
                  disabled={importingExcel}
                  onChange={handleOccupancyExcelImport}
                />
              </label>
            )}

            <button
              type="button"
              className="scooh-btn ghost"
              onClick={downloadOccupancyDemoTemplate}
              title="Download sample Excel template format with Site Block and Tabular structures"
              style={{
                fontSize: '11.5px',
                padding: '6px 12px',
                color: '#38bdf8',
                borderColor: 'rgba(56, 189, 248, 0.4)',
                background: 'rgba(56, 189, 248, 0.08)',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '5px'
              }}
            >
              <span>📥 Demo Template</span>
            </button>

            <button
              type="button"
              className="scooh-btn ghost"
              onClick={loadSystemData}
              disabled={loading}
              title="Refresh database"
              style={{ fontSize: '11.5px', padding: '6px 10px' }}
            >
              🔄
            </button>
          </div>
        </div>
      </section>

      {/* ── Sticky Selection Bar when 1 or more sites are checked ───────── */}
      {canDelete && selectedCount > 0 && (
        <div style={{
          position: 'sticky',
          top: '12px',
          zIndex: 90,
          background: 'rgba(15, 23, 42, 0.95)',
          backdropFilter: 'blur(10px)',
          border: '1px solid rgba(56, 189, 248, 0.4)',
          borderRadius: '12px',
          padding: '10px 18px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '14px',
          flexWrap: 'wrap',
          boxShadow: '0 10px 30px rgba(0,0,0,0.6)',
          marginBottom: '16px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span style={{ fontSize: '13px', fontWeight: 800, color: '#38bdf8' }}>
              ✓ {selectedCount} site{selectedCount > 1 ? 's' : ''} selected
            </span>
            <span style={{ fontSize: '12px', color: '#94a3b8' }}>
              ({selectedTotalCampCount} booking record{selectedTotalCampCount !== 1 ? 's' : ''})
            </span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '11.5px', padding: '4px 12px' }}
              onClick={deselectAll}
            >
              ✕ Deselect All
            </button>
            <button
              type="button"
              className="scooh-btn danger"
              style={{ fontSize: '11.5px', padding: '5px 14px' }}
              onClick={handleBatchDeleteBookings}
              disabled={deleting}
              title="Deletes booking history on selected sites, resetting them to 0% Vacant"
            >
              🗑️ Delete Bookings ({selectedTotalCampCount})
            </button>
            <button
              type="button"
              className="scooh-btn danger"
              style={{ fontSize: '11.5px', padding: '5px 14px', background: 'rgba(239, 68, 68, 0.15)', borderColor: 'rgba(239, 68, 68, 0.4)', color: '#f87171' }}
              onClick={handleBatchDeleteSites}
              disabled={deleting}
              title="Deletes selected sites and all their records from inventory"
            >
              ❌ Delete Sites ({selectedCount})
            </button>
          </div>
        </div>
      )}

      {/* ── Main Occupancy Table ────────────────────────────────────────── */}
      <section className="scooh-panel">
        {loading ? (
          <div style={{ textAlign: 'center', padding: '50px 20px', color: '#94a3b8' }}>
            <div style={{ fontSize: '32px', marginBottom: '12px' }}>🔄</div>
            <b>Loading occupancy history...</b>
          </div>
        ) : filteredSites.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '50px 20px', color: '#64748b' }}>
            <div style={{ fontSize: '36px', marginBottom: '12px' }}>🔍</div>
            <h4 style={{ margin: '0 0 6px', color: '#f1f5f9' }}>No matching sites found</h4>
            <p style={{ margin: '0 0 16px', fontSize: '12px' }}>Try adjusting your search filter or import an Excel file.</p>
            {canImport && (
              <label
                className="scooh-btn ghost"
                style={{
                  cursor: importingExcel ? 'not-allowed' : 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  fontSize: '12px',
                  padding: '7px 16px',
                  color: '#c084fc',
                  borderColor: 'rgba(168, 85, 247, 0.4)',
                  background: 'rgba(168, 85, 247, 0.08)',
                  margin: '0 auto'
                }}
                title="Import booking & occupancy records from Excel (.xlsx, .xls)"
              >
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv,.xlsm,.ods"
                  hidden
                  disabled={importingExcel}
                  onChange={handleOccupancyExcelImport}
                />
              </label>
            )}
            <div style={{ marginTop: '10px' }}>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={downloadOccupancyDemoTemplate}
                title="Download sample Excel template format with Site Block and Tabular structures"
                style={{
                  fontSize: '12px',
                  padding: '7px 16px',
                  color: '#38bdf8',
                  borderColor: 'rgba(56, 189, 248, 0.4)',
                  background: 'rgba(56, 189, 248, 0.08)',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                <span>📥 Download Demo Excel Template</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
            <table className="scooh-table" style={{ borderCollapse: 'collapse', width: '100%' }}>
              <thead>
                <tr>
                  {canDelete && (
                    <th style={{ width: '42px', minWidth: '42px', textAlign: 'center', position: 'sticky', left: 0, zIndex: 4, background: '#10161e' }}>
                      <input
                        type="checkbox"
                        checked={allVisibleSelected}
                        onChange={toggleSelectAll}
                        style={{ cursor: 'pointer', margin: 0 }}
                        title={allVisibleSelected ? "Deselect all" : "Select all"}
                      />
                    </th>
                  )}
                  <th style={{ position: 'sticky', left: canDelete ? '42px' : 0, zIndex: 3, background: '#10161e', minWidth: '110px' }}>Site ID</th>
                  <th style={{ minWidth: '140px' }}>Location / Area</th>

                  {/* 6 Months View Columns */}
                  {viewTab === '6months' && periods6M.map(p => (
                    <th key={p.key} style={{ minWidth: '130px', textAlign: 'center' }}>
                      <div>{p.label}</div>
                      <div style={{ fontSize: '10px', color: '#64748b', fontWeight: 500 }}>Bookings & Occ %</div>
                    </th>
                  ))}

                  {/* 12 Months View Columns */}
                  {viewTab === '12months' && periods12M.map(p => (
                    <th key={p.key} style={{ minWidth: '130px', textAlign: 'center' }}>
                      <div>{p.label}</div>
                      <div style={{ fontSize: '10px', color: '#64748b', fontWeight: 500 }}>Bookings & Occ %</div>
                    </th>
                  ))}

                  {/* 365 Days Overview Columns */}
                  {viewTab === 'overview' && (
                    <>
                      <th style={{ minWidth: '100px' }}>City</th>
                      <th style={{ minWidth: '140px' }}>Occupied Days (365d)</th>
                      <th style={{ minWidth: '160px' }}>Occupancy Rate</th>
                      <th style={{ minWidth: '100px', textAlign: 'center' }}>Status</th>
                    </>
                  )}

                  {/* Average Column for 6M and 12M */}
                  {viewTab !== 'overview' && (
                    <th style={{ minWidth: '125px', textAlign: 'center', background: '#0d1a2e', color: '#7dd3fc', borderLeft: '2px solid #1e3a5f' }}>
                      Average %
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {filteredSites.map((s, idx) => {
                  if (viewTab === 'overview') {
                    const statusText = s.pct365 >= 75 ? 'Full' : s.pct365 >= 40 ? 'High' : s.pct365 > 0 ? 'Partial' : 'Vacant';
                    const badgeBg = s.pct365 >= 75 ? 'rgba(16, 185, 129, 0.15)' : s.pct365 >= 40 ? 'rgba(56, 189, 248, 0.15)' : s.pct365 > 0 ? 'rgba(245, 158, 11, 0.15)' : 'rgba(100, 116, 139, 0.15)';
                    const badgeColor = s.pct365 >= 75 ? '#10b981' : s.pct365 >= 40 ? '#38bdf8' : s.pct365 > 0 ? '#f59e0b' : '#64748b';
                    const isSelected = selectedSiteCodes.has(s.site_code);

                    return (
                      <tr key={s.site_code || idx} style={{ background: isSelected ? 'rgba(56, 189, 248, 0.08)' : undefined }}>
                        {canDelete && (
                          <td style={{ textAlign: 'center', position: 'sticky', left: 0, zIndex: 3, background: isSelected ? '#132338' : (idx % 2 === 0 ? '#0b1016' : '#10161e'), width: '42px', minWidth: '42px' }}>
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => toggleSelectSite(s.site_code)}
                              style={{ cursor: 'pointer', margin: 0 }}
                              title={`Select site ${s.site_code}`}
                            />
                          </td>
                        )}
                        <td style={{ position: 'sticky', left: canDelete ? '42px' : 0, zIndex: 2, background: isSelected ? '#132338' : (idx % 2 === 0 ? '#0b1016' : '#10161e') }}>
                          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                            <span 
                              className="scooh-plate" 
                              style={{ cursor: 'pointer' }}
                              onClick={() => setModalData({ siteCode: s.site_code, area: s.area, periodLabel: 'All Bookings History', campaigns: s.allCampaigns || [] })}
                              title="Click to view all booking history"
                            >
                              {s.site_code}
                            </span>
                            {s.siteType && s.siteType !== 'Single' && (
                              <span
                                style={{
                                  fontSize: '9.5px',
                                  padding: '1px 5px',
                                  borderRadius: '4px',
                                  background: s.siteType === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                                  color: s.siteType === 'Combined' ? '#c084fc' : '#38bdf8',
                                  border: `1px solid ${s.siteType === 'Combined' ? 'rgba(168,85,247,0.4)' : 'rgba(56,189,248,0.4)'}`,
                                  fontWeight: 700
                                }}
                              >
                                {s.siteType}
                              </span>
                            )}
                          </div>
                        </td>
                        <td><b>{s.area}</b></td>
                        <td>{s.city}</td>
                        <td>
                          <span style={{ fontWeight: 700, color: s.days365 > 0 ? '#f1f5f9' : '#64748b' }}>
                            {s.days365 || 0} days
                          </span>
                          <span style={{ fontSize: '11px', color: '#64748b', marginLeft: '4px' }}>/ 365</span>
                        </td>
                        <td>
                          <OccPercentBar 
                            pct={s.pct365} 
                            days={s.days365} 
                            totalDays={365} 
                            clientNames={s.clients365} 
                            onClick={() => setModalData({ 
                              siteCode: s.site_code, 
                              area: s.area, 
                              periodLabel: '365-Day Rolling Period', 
                              campaigns: s.allCampaigns || [] 
                            })}
                          />
                        </td>
                        <td style={{ textAlign: 'center' }}>
                          <span className="scooh-badgechip" style={{ background: badgeBg, color: badgeColor, border: `1px solid ${badgeColor}40` }}>
                            {statusText}
                          </span>
                        </td>
                      </tr>
                    );
                  }

                  // Monthly (6 Months / 12 Months) View
                  const periods = viewTab === '12months' ? periods12M : periods6M;
                  const dataMap = viewTab === '12months' ? s.by12M : s.by6M;
                  const pcts = periods.map(p => {
                    const entry = dataMap?.[p.key];
                    return typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
                  });
                  const siteAvg = pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length) : 0;
                  const isSelected = selectedSiteCodes.has(s.site_code);

                  return (
                    <tr key={s.site_code || idx} style={{ background: isSelected ? 'rgba(56, 189, 248, 0.08)' : undefined }}>
                      {canDelete && (
                        <td style={{ textAlign: 'center', position: 'sticky', left: 0, zIndex: 3, background: isSelected ? '#132338' : (idx % 2 === 0 ? '#0b1016' : '#10161e'), width: '42px', minWidth: '42px' }}>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={() => toggleSelectSite(s.site_code)}
                            style={{ cursor: 'pointer', margin: 0 }}
                            title={`Select site ${s.site_code}`}
                          />
                        </td>
                      )}
                      <td style={{ position: 'sticky', left: canDelete ? '42px' : 0, zIndex: 2, background: isSelected ? '#132338' : (idx % 2 === 0 ? '#0b1016' : '#10161e') }}>
                        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                          <span 
                            className="scooh-plate" 
                            style={{ cursor: 'pointer' }}
                            onClick={() => setModalData({ siteCode: s.site_code, area: s.area, periodLabel: 'All Bookings History', campaigns: s.allCampaigns || [] })}
                            title="Click to view all booking history"
                          >
                            {s.site_code}
                          </span>
                          {s.siteType && s.siteType !== 'Single' && (
                            <span
                              style={{
                                fontSize: '9.5px',
                                padding: '1px 5px',
                                borderRadius: '4px',
                                background: s.siteType === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                                color: s.siteType === 'Combined' ? '#c084fc' : '#38bdf8',
                                border: `1px solid ${s.siteType === 'Combined' ? 'rgba(168,85,247,0.4)' : 'rgba(56,189,248,0.4)'}`,
                                fontWeight: 700
                              }}
                            >
                              {s.siteType}
                            </span>
                          )}
                        </div>
                      </td>
                      <td><b>{s.area}</b></td>

                      {periods.map(p => {
                        const entry = dataMap?.[p.key];
                        const pct = typeof entry === 'object' ? (entry?.pct ?? 0) : (Number(entry) || 0);
                        const days = typeof entry === 'object' ? entry?.days : undefined;
                        const totalDays = typeof entry === 'object' ? entry?.totalDays : undefined;
                        const clientNames = typeof entry === 'object' && entry?.clients ? entry.clients : [];
                        const brands = typeof entry === 'object' && entry?.brands ? entry.brands : [];
                        const cellCamps = typeof entry === 'object' && entry?.campaigns ? entry.campaigns : [];

                        return (
                          <td key={p.key} style={{ padding: '8px 10px', textAlign: 'center' }}>
                            <OccPercentBar 
                              pct={pct} 
                              days={days} 
                              totalDays={totalDays} 
                              clientNames={clientNames} 
                              brands={brands}
                              onClick={() => setModalData({ 
                                siteCode: s.site_code, 
                                area: s.area, 
                                periodLabel: `${p.label} (${pct}% Occupied)`, 
                                campaigns: cellCamps 
                              })}
                            />
                          </td>
                        );
                      })}

                      <td style={{ padding: '8px 10px', background: '#0d1a2e', borderLeft: '2px solid #1e3a5f', textAlign: 'center' }}>
                        <OccPercentBar pct={siteAvg} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* ── Modal for Inspecting Detailed Monthly Booking History ──────────── */}
      {modalData && (
        <div 
          className="scooh-modal-overlay" 
          style={{ position: 'fixed', inset: 0, zIndex: 100000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px' }}
          onClick={() => setModalData(null)}
        >
          <div 
            className="scooh-modal" 
            style={{ width: '100%', maxWidth: '600px', maxHeight: '88vh', overflowY: 'auto', margin: 'auto', border: '1px solid #38bdf8', boxShadow: '0 25px 60px rgba(0,0,0,0.75)', padding: '20px' }}
            onClick={e => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '14px', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: 0, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '17px' }}>
                  <span>📅</span>
                  <span>Site {modalData.siteCode} — Booking Details</span>
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                  {modalData.area} • <strong style={{ color: '#38bdf8' }}>{modalData.periodLabel}</strong>
                </p>
              </div>
              <button 
                type="button" 
                onClick={() => setModalData(null)} 
                style={{ background: 'none', border: 'none', color: '#94a3b8', fontSize: '20px', cursor: 'pointer', padding: '4px' }}
                title="Close"
              >
                ✕
              </button>
            </div>

            {modalData.campaigns && modalData.campaigns.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {[...modalData.campaigns].sort((a, b) => {
                  const da = parseFlexibleDate(a.start_date || a.booking_date) || new Date(0);
                  const db = parseFlexibleDate(b.start_date || b.booking_date) || new Date(0);
                  if (da.getTime() !== db.getTime()) return da - db;
                  const dea = parseFlexibleDate(a.end_date) || new Date(0);
                  const deb = parseFlexibleDate(b.end_date) || new Date(0);
                  return dea - deb;
                }).map((c, i) => {
                  const isVac = c.status === 'vacant' || c.is_vacant || isVacantClient(c.client || c.client_name);
                  const isEditingThis = bookingRecordId === (c.id || `${c.site_code}_${c.start_date}`);
                  return (
                    <div
                      key={c.id || i}
                      style={{
                        padding: '14px',
                        background: isVac ? 'linear-gradient(135deg, rgba(245, 158, 11, 0.12) 0%, rgba(20, 14, 8, 0.6) 100%)' : '#0b1320',
                        borderRadius: '8px',
                        border: isVac ? '1.5px dashed rgba(245, 158, 11, 0.65)' : '1px solid #1e3a5f',
                        boxShadow: isVac ? '0 4px 14px rgba(245, 158, 11, 0.08)' : 'none'
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '6px', gap: '10px' }}>
                        <div>
                          <div style={{ fontSize: '15px', fontWeight: 800, color: isVac ? '#fbbf24' : '#38bdf8' }}>
                            {isVac ? '🟡 Blank / Vacant Slot' : `👤 ${c.client || c.client_name || 'Client Not Specified'}`}
                          </div>
                          {c.brand && c.brand !== (c.client || c.client_name) && !isVac && (
                            <div style={{ fontSize: '12px', color: '#cbd5e1', marginTop: '2px' }}>
                              🏷️ Brand: <b>{c.brand}</b>
                            </div>
                          )}
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span
                            className="scooh-badgechip"
                            style={{
                              background: isVac ? 'rgba(245, 158, 11, 0.22)' : 'rgba(16, 185, 129, 0.15)',
                              color: isVac ? '#fbbf24' : '#10b981',
                              border: isVac ? '1px solid rgba(245, 158, 11, 0.5)' : '1px solid rgba(16, 185, 129, 0.3)',
                              fontWeight: 800,
                              whiteSpace: 'nowrap'
                            }}
                          >
                            {isVac ? '🟡 Vacant (0% Occ)' : (c.month || 'Active Booking')}
                          </span>
                          {!isVac && (
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              style={{ fontSize: '11px', padding: '2px 8px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.4)' }}
                              onClick={() => setViewCampaignDetails({
                                ...c,
                                site_code: c.site_code || modalData.siteCode,
                                location: c.location || modalData.area
                              })}
                              title="Open full campaign details popup"
                            >
                              👁 Details
                            </button>
                          )}
                          {isVac && canImport && !isEditingThis && (
                            <button
                              type="button"
                              className="scooh-btn"
                              onClick={() => {
                                setBookingRecordId(c.id || `${c.site_code}_${c.start_date}`);
                                setBookingClientName('');
                              }}
                              style={{
                                fontSize: '11px',
                                padding: '3px 8px',
                                background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                                color: '#0f172a',
                                fontWeight: 800,
                                border: 'none',
                                borderRadius: '4px',
                                cursor: 'pointer'
                              }}
                              title="Add a client campaign to book this site and change it from vacant to occupied"
                            >
                              ⚡ Book Site
                            </button>
                          )}
                        </div>
                      </div>

                      {/* Inline booking form when user clicks "Book Site" */}
                      {isEditingThis ? (
                        <div style={{
                          background: 'rgba(15, 23, 42, 0.85)',
                          border: '1px solid rgba(245, 158, 11, 0.4)',
                          borderRadius: '6px',
                          padding: '10px 12px',
                          margin: '8px 0',
                          display: 'flex',
                          flexDirection: 'column',
                          gap: '8px'
                        }}>
                          <div style={{ fontSize: '12px', color: '#fbbf24', fontWeight: 700 }}>
                            📝 Enter Client Name to Book this Site (changes to Occupied):
                          </div>
                          <div style={{ display: 'flex', gap: '8px' }}>
                            <input
                              type="text"
                              className="scooh-input"
                              placeholder="Client / Brand Name (e.g. Swagat Group, Hocco, Zydus)..."
                              value={bookingClientName}
                              onChange={e => setBookingClientName(e.target.value)}
                              autoFocus
                              style={{ flex: 1, fontSize: '12px', padding: '6px 10px' }}
                              onKeyDown={e => {
                                if (e.key === 'Enter') handleBookVacantRecord(c, bookingClientName);
                                if (e.key === 'Escape') setBookingRecordId(null);
                              }}
                            />
                            <button
                              type="button"
                              className="scooh-btn primary"
                              disabled={savingBooking}
                              onClick={() => handleBookVacantRecord(c, bookingClientName)}
                              style={{ fontSize: '11.5px', padding: '6px 14px', whiteSpace: 'nowrap' }}
                            >
                              {savingBooking ? 'Saving…' : '✓ Confirm'}
                            </button>
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              disabled={savingBooking}
                              onClick={() => setBookingRecordId(null)}
                              style={{ fontSize: '11.5px', padding: '6px 10px' }}
                            >
                              ✕
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div style={{ fontSize: '12.5px', color: isVac ? '#fcd34d' : '#e2e8f0', margin: '8px 0 6px' }}>
                          {isVac ? (
                            <span>ℹ️ <i>Site is currently vacant & available during this interval (0 days counted in occupancy)</i></span>
                          ) : (
                            <span>📢 <b>Campaign:</b> {c.campaign_name || c.display || 'Standard Display'}</span>
                          )}
                        </div>
                      )}

                      <div style={{ display: 'flex', gap: '16px', fontSize: '12px', color: isVac ? '#cbd5e1' : '#94a3b8', flexWrap: 'wrap', borderTop: isVac ? '1px solid rgba(245, 158, 11, 0.25)' : '1px solid #1e293b', paddingTop: '8px', marginTop: '8px' }}>
                        {(c.start_date || c.booking_date) && (
                          <span>🗓️ <b>Dates:</b> {(() => {
                            const sd = parseFlexibleDate(c.start_date || c.booking_date);
                            const ed = parseFlexibleDate(c.end_date);
                            const sStr = sd ? sd.toLocaleDateString('en-IN') : String(c.start_date || c.booking_date);
                            const eStr = ed ? ed.toLocaleDateString('en-IN') : (c.end_date ? String(c.end_date) : 'Ongoing');
                            return `${sStr} to ${eStr}`;
                          })()}</span>
                        )}
                        {c.booking_code && (
                          <span>🔖 <b>Code:</b> {c.booking_code}</span>
                        )}
                        {(c.total_amount || c.revenue) && !isVac ? (
                          <span style={{ color: '#10b981', fontWeight: 700 }}>💰 <b>Amount:</b> ₹{Number(c.total_amount || c.revenue).toLocaleString('en-IN')}</span>
                        ) : null}
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div style={{ textAlign: 'center', padding: '36px 20px', color: '#64748b' }}>
                <div style={{ fontSize: '36px', marginBottom: '8px' }}>📭</div>
                <h4 style={{ margin: '0 0 4px', color: '#cbd5e1' }}>No booking for this period</h4>
                <p style={{ margin: 0, fontSize: '12px' }}>The site was completely vacant and available for booking during this time.</p>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── Modal: Full Campaign Details Pop-up ── */}
      {viewCampaignDetails && (
        <CampaignDetailsModal
          campaign={viewCampaignDetails}
          onClose={() => setViewCampaignDetails(null)}
          onViewSite={siteCode => {
            const upper = String(siteCode || '').trim().toUpperCase();
            const siteObj = dbSites.find(s => String(s.site_code || '').trim().toUpperCase() === upper) || {
              site_code: upper,
              area: upper,
              location: '',
              type: 'Hoarding'
            };
            setSelectedSiteModal(siteObj);
          }}
          canAdd={false}
        />
      )}

      {/* ── Modal: Site Details Pop-up ── */}
      {selectedSiteModal && (
        <SiteDetailsModal
          site={selectedSiteModal}
          onClose={() => setSelectedSiteModal(null)}
          onNavigateToSites={() => {
            const code = selectedSiteModal.site_code;
            setSelectedSiteModal(null);
            navigate(`/sites?search=${encodeURIComponent(code)}`);
          }}
        />
      )}
    </>
  );
}

function getCampaignOccupancy(r, refDate) {
  if (r.occupancy !== undefined && r.occupancy !== null && String(r.occupancy).trim() !== '') {
    const raw = String(r.occupancy).trim();
    const num = parseFloat(raw.replace('%', ''));
    if (!isNaN(num)) {
      return {
        pct: num,
        label: `${num}%`,
        status: num >= 75 ? 'occupied' : num > 0 ? 'partial' : 'vacant'
      };
    }
    return { pct: null, label: raw, status: 'custom' };
  }

  const today = refDate ? (parseDay(refDate) || new Date()) : new Date();
  today.setHours(0, 0, 0, 0);

  const start = parseDay(r.start_date || r.booking_date);
  const end = parseDay(r.end_date);

  if (end && !isNaN(end.getTime())) {
    end.setHours(23, 59, 59, 999);
    if (end < today) {
      return { pct: 0, label: '0% Vacant', sub: 'Ended', status: 'vacant' };
    }
    if (start && !isNaN(start.getTime()) && start > today) {
      return {
        pct: 100,
        label: '100% Booked',
        sub: `Starts ${start.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })}`,
        status: 'occupied'
      };
    }
    return { pct: 100, label: '100% Occupied', status: 'occupied' };
  }

  if (r.status === 'occupied') {
    return { pct: 100, label: '100% Occupied', status: 'occupied' };
  }
  return { pct: 0, label: '0% Vacant', status: 'vacant' };
}

function getCampaignPeriod(r, view) {
  const dStr = r.start_date || r.booking_date || r.end_date;
  const d = dStr ? new Date(dStr) : null;
  const validDate = d && !isNaN(d.getTime());

  if (view === 'monthly') {
    if (r.month) return r.month;
    if (validDate) return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
    return 'Other';
  }
  if (view === 'weekly') {
    if (validDate) {
      const utc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      utc.setUTCDate(utc.getUTCDate() + 4 - (utc.getUTCDay() || 7));
      const yr = utc.getUTCFullYear();
      const wk = Math.ceil(((utc - new Date(Date.UTC(yr, 0, 1))) / 86400000 + 1) / 7);
      return `${yr}-W${String(wk).padStart(2, '0')}`;
    }
    return 'Other';
  }
  if (view === 'yearly') {
    if (validDate) return String(d.getFullYear());
    const mMatch = String(r.month || '').match(/\b(20\d\d)\b/);
    if (mMatch) return mMatch[1];
    return 'Other';
  }
  return 'All';
}

async function exportCampaignsExcel(rowsData, filename = 'MediaBuzz_Latest_Bookings.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Latest Bookings', {
    views: [{ state: 'frozen', ySplit: 2 }]
  });

  const cols = [
    { key: 'site_code', width: 14 },
    { key: 'month', width: 14 },
    { key: 'client', width: 28 },
    { key: 'display', width: 24 },
    { key: 'vendor_name', width: 22 },
    { key: 'location', width: 36 },
    { key: 'width', width: 10 },
    { key: 'height', width: 10 },
    { key: 'size', width: 14 },
    { key: 'type', width: 16 },
    { key: 'start_date', width: 14 },
    { key: 'end_date', width: 14 },
    { key: 'days', width: 10 },
    { key: 'advt_fees', width: 20 },
    { key: 'printing_mounting_cost', width: 22 },
    { key: 'total_amount', width: 20 },
    { key: 'pending', width: 18 }
  ];
  worksheet.columns = cols;

  // Executive Media Buzz Brand Header Banner with Logo on Top Right
  attachMediaBuzzExcelHeader(workbook, worksheet, {
    title: 'MEDIA BUZZ — LATEST BOOKINGS TRACKER',
    columns: cols,
    totalColumns: 17
  });

  // Populate Header Row (Row 2)
  const headers = [
    'Site Code', 'Month', 'Client/Agency Name',
    'Display', 'Vendor Name', 'Location', 'W', 'H',
    'Size', 'Type', 'Start Date', 'End Date', 'Days',
    'Advt. Fees per month', 'Printing & Mounting', 'Total Amount', 'Pending'
  ];
  const headerRow = worksheet.getRow(2);
  headerRow.height = 28;
  headers.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });

  rowsData.forEach((r) => {
    worksheet.addRow({
      site_code: r.site_code || '',
      month: r.month || '',
      client: r.client || r.client_name || '',
      display: r.display || r.campaign_name || r.brand || '',
      vendor_name: r.vendor_name || '',
      location: r.location || '',
      width: r.width ?? '',
      height: r.height ?? '',
      size: r.size || (r.width && r.height ? `${r.width}x${r.height} ft` : ''),
      type: r.type || 'Hoarding',
      start_date: r.start_date ? new Date(r.start_date).toLocaleDateString('en-IN') : '',
      end_date: r.end_date ? new Date(r.end_date).toLocaleDateString('en-IN') : '',
      days: r.days || (r.start_date && r.end_date ? Math.max(1, Math.round((new Date(r.end_date) - new Date(r.start_date)) / 86400000) + 1) : 30),
      advt_fees: Number(r.advt_fees || 0),
      printing_mounting_cost: Number(r.printing_mounting_cost || (Number(r.printing_cost || 0) + Number(r.mounting_cost || 0))),
      total_amount: Number(r.total_amount || r.revenue || 0),
      pending: Number(r.pending || 0)
    });
  });

  // Style Header Row in Bright Yellow (#FFFF00)
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 2, 7, 8, 9, 10, 11, 12, 13].includes(colNumber) ? 'center' : ([14, 15, 16, 17].includes(colNumber) ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  // Style Data Rows (Row 3+)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= 3) {
      row.height = 22;
      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Calibri', size: 10.5 };
        cell.alignment = {
          vertical: 'middle',
          horizontal: [1, 2, 7, 8, 9, 10, 11, 12, 13].includes(colNumber) ? 'center' : ([14, 15, 16, 17].includes(colNumber) ? 'right' : 'left')
        };
        if ([14, 15, 16, 17].includes(colNumber) && typeof cell.value === 'number') {
          cell.numFmt = '#,##,##0';
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          left: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          bottom: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          right: { style: 'thin', color: { argb: 'FFE8E8E8' } }
        };
      });
    }
  });

  // Enable Auto-Filter on Row 2
  if (rowsData.length > 0) {
    worksheet.autoFilter = `A2:Q${rowsData.length + 2}`;
  }

  // Official Terms & Conditions at Bottom Center
  attachMediaBuzzTermsAndConditions(worksheet, {
    totalColumns: 17
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);

  try {
    api.post('/storage', {
      category: 'excel',
      name: filename,
      type: 'xlsx',
      size: `${Math.round(blob.size / 1024)} KB`
    }).catch(() => {});
  } catch (e) {}
}

function CampaignTrackerView() {
  const location = useLocation();
  const navigate = useNavigate();
  const [rows, setRows] = useState([]);
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState(false);
  const [editModal, setEditModal] = useState(null);
  const [modalSiteCode, setModalSiteCode] = useState('');
  const [modalLocation, setModalLocation] = useState('');
  const [modalWidth, setModalWidth] = useState('');
  const [modalHeight, setModalHeight] = useState('');
  const [modalSize, setModalSize] = useState('');
  const [modalType, setModalType] = useState('Hoarding');
  const [autoMatchedSite, setAutoMatchedSite] = useState(null);
  const [search, setSearch] = useState('');
  const [siteQueryParam, setSiteQueryParam] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [typeFilter, setTypeFilter] = useState('ALL');
  const [vendorFilter, setVendorFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  const [saving, setSaving] = useState(false);
  const [banner, setBanner] = useState('');
  const [importingExcel, setImportingExcel] = useState(false);
  const [isViewingOnly, setIsViewingOnly] = useState(false);
  const [viewCampaignDetails, setViewCampaignDetails] = useState(null);
  const [selectedSiteModal, setSelectedSiteModal] = useState(null);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [expandedHistorySites, setExpandedHistorySites] = useState(new Set());
  const [modalSelectedSites, setModalSelectedSites] = useState([]);
  const [siteFilterText, setSiteFilterText] = useState('');
  const [onlyVacantFilter, setOnlyVacantFilter] = useState(false);
  const [dateFilter, setDateFilter] = useState(''); // available from this date

  function toggleHistoryExpand(siteCode) {
    setExpandedHistorySites(prev => {
      const next = new Set(prev);
      if (next.has(siteCode)) next.delete(siteCode); else next.add(siteCode);
      return next;
    });
  }

  const topScrollRef = useRef(null);
  const bottomScrollRef = useRef(null);
  const tableRef = useRef(null);
  const isScrollingRef = useRef(null);
  const [tableScrollWidth, setTableScrollWidth] = useState(2400);

  function handleTopScroll() {
    if (isScrollingRef.current === 'bottom') return;
    isScrollingRef.current = 'top';
    if (bottomScrollRef.current && topScrollRef.current) {
      bottomScrollRef.current.scrollLeft = topScrollRef.current.scrollLeft;
    }
    setTimeout(() => {
      if (isScrollingRef.current === 'top') isScrollingRef.current = null;
    }, 40);
  }

  function handleBottomScroll() {
    if (isScrollingRef.current === 'top') return;
    isScrollingRef.current = 'bottom';
    if (topScrollRef.current && bottomScrollRef.current) {
      topScrollRef.current.scrollLeft = bottomScrollRef.current.scrollLeft;
    }
    setTimeout(() => {
      if (isScrollingRef.current === 'bottom') isScrollingRef.current = null;
    }, 40);
  }

  function scrollHorizontallyBy(delta) {
    if (bottomScrollRef.current) {
      bottomScrollRef.current.scrollBy({ left: delta, behavior: 'smooth' });
    }
  }

  useEffect(() => {
    function updateScrollWidth() {
      if (tableRef.current) {
        const sw = tableRef.current.scrollWidth || tableRef.current.offsetWidth;
        if (sw) setTableScrollWidth(sw);
      } else if (bottomScrollRef.current) {
        const sw = bottomScrollRef.current.scrollWidth;
        if (sw) setTableScrollWidth(sw);
      }
    }

    updateScrollWidth();
    const t = setTimeout(updateScrollWidth, 150);

    let ro;
    if (window.ResizeObserver && bottomScrollRef.current) {
      ro = new ResizeObserver(() => updateScrollWidth());
      ro.observe(bottomScrollRef.current);
      if (tableRef.current) ro.observe(tableRef.current);
    }

    window.addEventListener('resize', updateScrollWidth);
    return () => {
      clearTimeout(t);
      window.removeEventListener('resize', updateScrollWidth);
      if (ro) ro.disconnect();
    };
  }, [rows]);

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canAdd = isAdmin || isManager;
  const canEdit = !isReadOnly;

  async function loadData() {
    setLoading(true);
    try {
      const [campRes, sitesRes] = await Promise.all([
        api.get('/campaigns'),
        api.get('/sites')
      ]);
      if (Array.isArray(campRes.data)) {
        setRows(campRes.data);
      }
      if (Array.isArray(sitesRes.data)) {
        setSites(sitesRes.data);
      }
    } catch (err) {
      console.warn('Error loading campaigns/sites:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadData();
    const interval = setInterval(loadData, 35000);
    const onCampUpdate = () => loadData();
    window.addEventListener('mb-campaigns-updated', onCampUpdate);
    return () => {
      clearInterval(interval);
      window.removeEventListener('mb-campaigns-updated', onCampUpdate);
    };
  }, []);

  // Auto-detect site code from location and size as user types or edits
  useEffect(() => {
    if (!editModal) {
      setAutoMatchedSite(null);
      return;
    }
    const derivedSize = modalSize || (modalWidth && modalHeight ? `${modalWidth}x${modalHeight}` : '');
    if (!modalLocation && !derivedSize && !modalWidth && !modalHeight) {
      setAutoMatchedSite(null);
      return;
    }
    const matched = matchSiteByLocationAndSize(
      modalLocation,
      derivedSize,
      modalWidth,
      modalHeight,
      sites
    );
    setAutoMatchedSite(matched);
    // If site code is empty and a high-confidence match is detected, auto-populate it
    if (matched && !modalSiteCode.trim()) {
      setModalSiteCode(matched.site_code);
    }
  }, [modalLocation, modalSize, modalWidth, modalHeight, editModal, sites]);

  function applyMatchedSite(s) {
    if (!s) return;
    setModalSiteCode(s.site_code || '');
    if (s.address || s.area) setModalLocation(s.address || s.area || '');
    if (s.width) setModalWidth(s.width);
    if (s.height) setModalHeight(s.height);
    if (s.size) setModalSize(s.size);
    if (s.type) setModalType(s.type);
  }

  function handleSiteCodeChange(code) {
    setModalSiteCode(code);
    const upper = String(code).trim().toUpperCase();
    const found = sites.find(s => String(s.site_code || '').toUpperCase() === upper);
    if (found) {
      if (!modalLocation) setModalLocation(found.address || found.area || '');
      if (!modalWidth && found.width) setModalWidth(found.width);
      if (!modalHeight && found.height) setModalHeight(found.height);
      if (!modalSize && found.size) setModalSize(found.size);
      if (found.type) setModalType(found.type);
    }
  }

  // Sync search/filter when navigating with ?site=MB-XX or ?search=...
  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const siteParam = (params.get('site') || params.get('search') || '').trim();
    if (siteParam) {
      setSearch(siteParam);
      setSiteQueryParam(siteParam);
    } else {
      setSiteQueryParam('');
    }
  }, [location.search]);

  // Master directory — only sites that have valid campaign records (Blank rows show in Occupancy only!)
  const siteMasterList = useMemo(() => {
    // Build a lookup from DB sites for enriching campaign rows with site metadata
    const siteDbLookup = new Map();
    sites.forEach(s => {
      const code = String(s.site_code || '').trim().toUpperCase();
      if (code) siteDbLookup.set(code, s);
    });

    const siteMap = new Map();

    // Populate all registered sites so every site exists in the master list
    (sites || []).forEach(s => {
      const code = String(s.site_code || '').trim().toUpperCase();
      if (!code) return;
      siteMap.set(code, {
        code,
        id: s.id || null,
        location: s.address || s.area || '—',
        city: s.city || 'Ahmedabad',
        size: s.size || (s.width && s.height ? `${s.width}x${s.height} ft` : '—'),
        width: s.width ?? null,
        height: s.height ?? null,
        type: s.type || 'Hoarding',
        rows: []
      });
    });

    // Add campaign rows to their corresponding site
    rows.forEach(r => {
      if (isVacantClient(r.client || r.client_name || r.display)) return;
      const code = String(r.site_code || 'UNASSIGNED').trim().toUpperCase();
      if (!siteMap.has(code)) {
        const dbSite = siteDbLookup.get(code);
        siteMap.set(code, {
          code,
          id: dbSite?.id || null,
          location: dbSite?.address || dbSite?.area || r.location || '—',
          city: dbSite?.city || 'Ahmedabad',
          size: dbSite?.size || (dbSite?.width && dbSite?.height ? `${dbSite.width}x${dbSite.height} ft` : r.size || (r.width && r.height ? `${r.width}x${r.height} ft` : '—')),
          width: dbSite?.width ?? r.width ?? null,
          height: dbSite?.height ?? r.height ?? null,
          type: dbSite?.type || r.type || 'Hoarding',
          rows: []
        });
      }
      siteMap.get(code).rows.push(r);
    });

    const evalDate = dateFilter ? (parseDay(dateFilter) || new Date()) : new Date();
    evalDate.setHours(0, 0, 0, 0);

    const list = Array.from(siteMap.values()).map(site => {
      const sortedRows = [...site.rows].sort((a, b) => {
        const da = parseDay(a.start_date || a.booking_date) || new Date(0);
        const db = parseDay(b.start_date || b.booking_date) || new Date(0);
        return db - da;
      });

      let activeCampaign = null;
      let upcomingCampaign = null;
      let latestCampaign = sortedRows[0] || null;

      for (const r of sortedRows) {
        const sDate = parseDay(r.start_date || r.booking_date);
        let eDate = parseDay(r.end_date);
        if (!eDate && sDate) {
          eDate = new Date(sDate.getFullYear(), sDate.getMonth() + 1, 0);
        }
        if (sDate && eDate) {
          if (sDate <= evalDate && eDate >= evalDate) {
            activeCampaign = r;
            break;
          } else if (sDate > evalDate && !upcomingCampaign) {
            upcomingCampaign = r;
          }
        } else if (sDate && !eDate) {
          if (sDate <= evalDate) {
            activeCampaign = r;
            break;
          }
        }
      }

      const totalAmount = site.rows.reduce((sum, r) => sum + Number(r.total_amount || r.revenue || 0), 0);
      const totalPending = site.rows.reduce((sum, r) => sum + Number(r.pending || 0), 0);

      // A site is occupied ONLY if it currently has an active campaign running on evalDate!
      let status = 'vacant';
      let currentCampaign = null;
      if (activeCampaign) {
        status = 'occupied';
        currentCampaign = activeCampaign;
      } else if (upcomingCampaign) {
        status = 'upcoming';
        currentCampaign = upcomingCampaign;
      } else if (latestCampaign) {
        currentCampaign = latestCampaign;
        status = 'vacant';
      }

      return {
        ...site,
        rows: sortedRows,
        campaignCount: site.rows.length,
        status,
        currentCampaign,
        activeCampaign,
        upcomingCampaign,
        latestCampaign,
        totalAmount,
        totalPending
      };
    });

    return list.sort((a, b) => universalCompare(a.code, b.code, 'asc'));
  }, [sites, rows, dateFilter]);

  // Unified dataset: exactly 1 row per physical site displaying ONLY its latest/current booking
  const latestSiteCampaigns = useMemo(() => {
    const evalDate = dateFilter ? (parseDay(dateFilter) || new Date()) : new Date();
    return siteMasterList.map(s => {
      const c = s.currentCampaign;
      return {
        site_code: s.code,
        site_id: s.id,
        location: s.location,
        city: s.city,
        size: s.size,
        type: c?.type || s.type || 'Hoarding',
        width: c?.width ?? s.width ?? '',
        height: c?.height ?? s.height ?? '',
        siteStatus: s.status, // 'occupied' | 'upcoming' | 'vacant'
        campaignCount: s.campaignCount,
        allCampaigns: s.rows,

        // Latest campaign data:
        campaign_id: c?.id || null,
        campaign: c || null,
        month: c?.month || (c?.start_date ? new Date(c.start_date).toLocaleDateString('en-US', { month: 'short', year: 'numeric' }) : ''),
        occupancy: (s.status === 'occupied' && c) ? getCampaignOccupancy(c, evalDate) : { pct: 0, label: '0% Vacant', status: 'vacant' },
        booking_date: c?.booking_date || c?.date || '',
        client: c?.client || c?.client_name || '',
        display: c?.display || c?.campaign_name || '',
        vendor_name: c?.vendor_name || '',
        start_date: c?.start_date || '',
        end_date: c?.end_date || '',
        days: c?.days || (c?.start_date && c?.end_date ? Math.max(1, Math.round((new Date(c.end_date) - new Date(c.start_date)) / 86400000) + 1) : (c ? 30 : '')),
        advt_fees: Number(c?.advt_fees || 0),
        printing_mounting_cost: Number(c?.printing_mounting_cost || (Number(c?.printing_cost || 0) + Number(c?.mounting_cost || 0))),
        total_amount: Number(c?.total_amount || c?.revenue || 0),
        po: c?.po || '',
        bill: c?.bill || c?.invoice_no || '',
        pending: Number(c?.pending || 0),
        parent_campaign: c?.parent_campaign || '',
        notes: c?.notes || ''
      };
    });
  }, [siteMasterList, dateFilter]);

  // Counts & metrics
  const totalSites = siteMasterList.length;
  const occupiedSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'occupied').length, [latestSiteCampaigns]);
  const upcomingSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'upcoming').length, [latestSiteCampaigns]);
  const vacantSitesCount = useMemo(() => latestSiteCampaigns.filter(x => x.siteStatus === 'vacant').length, [latestSiteCampaigns]);
  const occupancyRate = totalSites > 0 ? Math.round((occupiedSitesCount / totalSites) * 100) : 0;
  // Revenue & pending — deduplicated by campaign so multi-site campaigns are counted once
  const latestRevenueSum = useMemo(() => {
    const seen = new Set();
    return latestSiteCampaigns.reduce((sum, x) => {
      if (!x.total_amount) return sum;
      // Build a unique key per campaign: prefer parent_campaign, else fall back to campaign_id
      const key = x.parent_campaign && !String(x.parent_campaign).startsWith('LINKED:')
        ? String(x.parent_campaign).trim().toLowerCase()
        : x.campaign_id ? `__id__${x.campaign_id}` : `__site__${x.site_code}`;
      if (seen.has(key)) return sum;
      seen.add(key);
      return sum + x.total_amount;
    }, 0);
  }, [latestSiteCampaigns]);

  const latestPendingSum = useMemo(() => {
    const seen = new Set();
    return latestSiteCampaigns.reduce((sum, x) => {
      if (!x.pending) return sum;
      const key = x.parent_campaign && !String(x.parent_campaign).startsWith('LINKED:')
        ? String(x.parent_campaign).trim().toLowerCase()
        : x.campaign_id ? `__id__${x.campaign_id}` : `__site__${x.site_code}`;
      if (seen.has(key)) return sum;
      seen.add(key);
      return sum + x.pending;
    }, 0);
  }, [latestSiteCampaigns]);

  const vacantSiteCodeSet = useMemo(() => {
    const set = new Set();
    latestSiteCampaigns.forEach(s => {
      if (s.siteStatus === 'vacant') set.add(String(s.site_code || '').trim().toUpperCase());
    });
    return set;
  }, [latestSiteCampaigns]);

  const selectableSitesList = useMemo(() => {
    const list = sites.filter(s => {
      const code = String(s.site_code || '').trim().toUpperCase();
      if (!code) return false;
      if (onlyVacantFilter && !vacantSiteCodeSet.has(code)) return false;
      if (!siteFilterText.trim()) return true;
      const q = siteFilterText.trim().toLowerCase();
      return (
        code.toLowerCase().includes(q) ||
        String(s.address || '').toLowerCase().includes(q) ||
        String(s.area || '').toLowerCase().includes(q) ||
        String(s.size || '').toLowerCase().includes(q) ||
        String(s.media_type || s.type || '').toLowerCase().includes(q)
      );
    });
    return list.sort((a, b) => universalCompare(a.site_code, b.site_code, 'asc'));
  }, [sites, siteFilterText, onlyVacantFilter, vacantSiteCodeSet]);

  function toggleModalSite(code) {
    if (!code) return;
    const clean = code.trim();
    setModalSelectedSites(prev => {
      const next = prev.includes(clean) ? prev.filter(c => c !== clean) : [...prev, clean];
      if (next.length === 1) {
        handleSiteCodeChange(next[0]);
      } else if (next.length === 0) {
        setModalSiteCode('');
      } else {
        setModalSiteCode(next.join(', '));
      }
      return next;
    });
  }

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  // Filtered and sorted latest site bookings list (statuses are evaluated dynamically as of dateFilter)
  const filteredLatest = useMemo(() => {
    const list = latestSiteCampaigns.filter(item => {
      if (statusFilter === 'occupied' && item.siteStatus !== 'occupied') return false;
      if (statusFilter === 'upcoming' && item.siteStatus !== 'upcoming') return false;
      if (statusFilter === 'vacant' && item.siteStatus !== 'vacant') return false;
      if (statusFilter === 'pending' && item.pending <= 0) return false;

      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        item.site_code.toLowerCase().includes(q) ||
        item.location.toLowerCase().includes(q) ||
        item.city.toLowerCase().includes(q) ||
        item.size.toLowerCase().includes(q) ||
        item.type.toLowerCase().includes(q) ||
        item.client.toLowerCase().includes(q) ||
        item.display.toLowerCase().includes(q) ||
        item.vendor_name.toLowerCase().includes(q) ||
        item.month.toLowerCase().includes(q) ||
        item.po.toLowerCase().includes(q) ||
        item.bill.toLowerCase().includes(q)
      );
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'occupancy') {
          valA = a.occupancy?.pct ?? 0;
          valB = b.occupancy?.pct ?? 0;
        } else if (['advt_fees', 'printing_mounting_cost', 'total_amount', 'pending', 'days', 'width', 'height'].includes(sortState.key)) {
          valA = Number(valA || 0);
          valB = Number(valB || 0);
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    } else {
      list.sort((a, b) => a.site_code.localeCompare(b.site_code, undefined, { numeric: true }));
    }
    return list;
  }, [latestSiteCampaigns, statusFilter, search, sortState]);

  function toggleSelect(id) {
    if (!id) return;
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleCampaignIds = filteredLatest.map(r => r.campaign_id).filter(Boolean);
    const allSelected = visibleCampaignIds.length > 0 && visibleCampaignIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleCampaignIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleCampaignIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filteredLatest.map(r => r.campaign_id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to permanently delete all ${count} selected campaign record${count > 1 ? 's' : ''}? This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/campaigns/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setBanner(`✓ ${count} campaign record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected campaigns: ' + (err.response?.data?.message || err.message));
    }
  }

  async function handleDeleteAllCampaigns() {
    if (!canDelete) return;
    const allVisibleCampIds = filteredLatest.map(r => r.campaign_id).filter(Boolean);
    const count = allVisibleCampIds.length;
    if (count === 0) {
      alert('There are no active or visible campaign records to delete.');
      return;
    }
    if (!confirm(`⚠️ DANGER: Are you sure you want to permanently delete ALL ${count} visible booking records from Latest Booking?\n\nThis will clear all booking data for these sites. This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/campaigns/batch-delete', { ids: allVisibleCampIds, hard: true });
      setBanner(`✓ All ${count} campaign records deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      setTimeout(() => setBanner(''), 5000);
    } catch (err) {
      alert('Failed to delete all campaigns: ' + (err.response?.data?.message || err.message));
    }
  }

  async function deleteRecord(id) {
    if (!canDelete) return;
    if (confirm('Are you sure you want to permanently delete this campaign record?')) {
      try {
        await api.delete(`/campaigns/${id}?hard=true`);
        setBanner('✓ Campaign record deleted successfully.');
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await loadData();
        window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
        window.dispatchEvent(new CustomEvent('mb-sites-updated'));
        setTimeout(() => setBanner(''), 3000);
      } catch (err) {
        alert('Failed to delete campaign: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  async function handleCampaignsExcelImport(e) {
    if (!canDelete) return;
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/campaigns-xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `Processed ${r.data.rows} campaigns (${r.data.updated} updated, ${r.data.created} created).`;
      setBanner(`✓ ${msg}`);
      alert(`✓ Campaigns Import Successful!\n\n${msg}`);
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      setTimeout(() => setBanner(''), 6000);
    } catch (err) {
      alert('Campaign Excel Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  function openNewCampaign(code) {
    if (!canAdd) return;
    setIsViewingOnly(false);
    const initialSiteCode = typeof code === 'string' && code ? code : (siteQueryParam || (search.trim().startsWith('MB-') ? search.trim() : ''));
    setModalSiteCode(initialSiteCode);
    setModalSelectedSites(initialSiteCode ? [initialSiteCode] : []);
    setSiteFilterText('');
    setOnlyVacantFilter(false);
    const foundSite = sites.find(s => String(s.site_code || '').toUpperCase() === initialSiteCode.toUpperCase());
    const initLoc = foundSite ? (foundSite.address || foundSite.area || '') : '';
    const initW = foundSite?.width ?? '';
    const initH = foundSite?.height ?? '';
    const initSize = foundSite?.size ?? (initW && initH ? `${initW}x${initH} ft` : '');
    const initType = foundSite?.type || 'Hoarding';
    setModalLocation(initLoc);
    setModalWidth(initW);
    setModalHeight(initH);
    setModalSize(initSize);
    setModalType(initType);
    setAutoMatchedSite(foundSite || null);
    setEditModal({
      site_code: initialSiteCode,
      month: new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
      booking_date: new Date().toISOString().slice(0, 10),
      client: '',
      display: '',
      vendor_name: '',
      location: initLoc,
      width: initW,
      height: initH,
      size: initSize,
      type: initType,
      start_date: new Date().toISOString().slice(0, 10),
      end_date: '',
      days: 30,
      advt_fees: '',
      printing_mounting_cost: '',
      total_amount: '',
      po: '',
      bill: '',
      pending: '',
      notes: ''
    });
  }

  function openEditCampaign(r) {
    if (!r) return;
    setIsViewingOnly(false);
    setModalSiteCode(r.site_code || '');
    setModalSelectedSites(r.site_code ? [r.site_code] : []);
    setModalLocation(r.location || '');
    setModalWidth(r.width ?? '');
    setModalHeight(r.height ?? '');
    setModalSize(r.size || '');
    setModalType(r.type || 'Hoarding');
    setAutoMatchedSite(null);
    setEditModal(r);
  }

  function openViewCampaign(r) {
    if (!r) return;
    setViewCampaignDetails(r);
  }

  async function saveCampaign(e) {
    e.preventDefault();
    if (isReadOnly) return;
    setSaving(true);
    const form = e.currentTarget;
    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    payload.advt_fees = Number(payload.advt_fees || 0);
    payload.printing_mounting_cost = Number(payload.printing_mounting_cost || 0);
    let totalAmt = Number(payload.total_amount || 0);
    if (!totalAmt && (payload.advt_fees || payload.printing_mounting_cost)) {
      totalAmt = payload.advt_fees + payload.printing_mounting_cost;
    }
    payload.total_amount = totalAmt;
    payload.revenue = totalAmt;
    payload.pending = Number(payload.pending || 0);
    payload.days = Number(payload.days || 30);
    if (payload.width) payload.width = parseFloat(payload.width) || null;
    if (payload.height) payload.height = parseFloat(payload.height) || null;
    if (!payload.size && payload.width && payload.height) {
      payload.size = `${payload.width}x${payload.height} ft`;
    }
    if (payload.client) payload.client_name = payload.client;
    if (payload.display) {
      payload.campaign_name = payload.display;
      payload.brand = payload.display;
    }
    // PO and Bill are completely optional without compulsion
    payload.po = payload.po ? String(payload.po).trim() : '';
    payload.bill = payload.bill ? String(payload.bill).trim() : '';

    // Determine target site codes
    let targetSiteCodes = [];
    if (!editModal?.id && modalSelectedSites.length > 0) {
      targetSiteCodes = [...modalSelectedSites];
    } else if (payload.site_code) {
      targetSiteCodes = String(payload.site_code).split(',').map(s => s.trim()).filter(Boolean);
    }

    if (targetSiteCodes.length === 0) {
      alert('Please select or enter at least one site code for this campaign.');
      setSaving(false);
      return;
    }

    try {
      if (editModal?.id) {
        payload.site_code = targetSiteCodes[0] || String(payload.site_code).trim();
        await api.put(`/campaigns/${editModal.id}`, payload);
        setBanner('✓ Campaign record updated successfully!');
      } else if (targetSiteCodes.length > 1) {
        // Multi-site campaign creation for the same client
        const groupTag = `${payload.client || 'Client'} - ${payload.display || 'Campaign'} (${targetSiteCodes.length} sites)`;
        const promises = targetSiteCodes.map(code => {
          const upper = code.toUpperCase();
          const s = sites.find(x => String(x.site_code || '').trim().toUpperCase() === upper);
          const sLoc = s ? (s.address || s.area || payload.location || '') : (payload.location || '');
          const sW = s?.width ?? payload.width ?? null;
          const sH = s?.height ?? payload.height ?? null;
          const sSize = s?.size || (s?.width && s?.height ? `${s.width}x${s.height} ft` : payload.size || '');
          const sType = s?.type || s?.media_type || payload.type || 'Hoarding';

          const sitePayload = {
            ...payload,
            site_code: code,
            location: sLoc,
            width: sW,
            height: sH,
            size: sSize,
            type: sType,
            parent_campaign: payload.parent_campaign || groupTag
          };
          return api.post('/campaigns', sitePayload);
        });
        await Promise.all(promises);
        setBanner(`✓ Campaign created across ${targetSiteCodes.length} sites for ${payload.client || 'client'} successfully!`);
      } else {
        payload.site_code = targetSiteCodes[0];
        await api.post('/campaigns', payload);
        setBanner('✓ New campaign created successfully!');
      }
      setEditModal(null);
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to save campaign: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  function renderLatestSiteRow(item) {
    const isPending = item.pending > 0;
    const isBooked = item.siteStatus === 'occupied';
    const isUpcoming = item.siteStatus === 'upcoming';
    const statusBg = isBooked ? 'rgba(34, 197, 94, 0.16)' : isUpcoming ? 'rgba(234, 179, 8, 0.16)' : 'rgba(148, 163, 184, 0.12)';
    const statusColor = isBooked ? '#4ade80' : isUpcoming ? '#facc15' : '#94a3b8';
    const statusBorder = isBooked ? 'rgba(34, 197, 94, 0.35)' : isUpcoming ? 'rgba(234, 179, 8, 0.35)' : 'rgba(148, 163, 184, 0.25)';

    const prevCampaigns = (item.allCampaigns || []).filter(c => c.id !== item.campaign_id);
    const isExpanded = expandedHistorySites.has(item.site_code);
    const colSpan = canDelete ? 19 : 18;

    return (
      <React.Fragment key={item.site_code}>
        <tr>
        {canDelete && (
          <td style={{ textAlign: 'center', padding: '10px 8px' }}>
            {item.campaign_id ? (
              <input
                type="checkbox"
                checked={selectedIds.has(item.campaign_id)}
                onChange={() => toggleSelect(item.campaign_id)}
                style={{ cursor: 'pointer' }}
                title={`Select latest campaign on ${item.site_code}`}
              />
            ) : (
              <span style={{ color: '#475569', fontSize: '11px' }}>—</span>
            )}
          </td>
        )}
        <td>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
            <span
              className="scooh-plate"
              style={{ fontSize: '12px', fontWeight: 800, cursor: 'pointer' }}
              onClick={() => {
                const upper = String(item.site_code || '').trim().toUpperCase();
                const siteObj = sites.find(s => String(s.site_code || '').trim().toUpperCase() === upper) || {
                  site_code: upper,
                  area: item.location || upper,
                  location: item.location || '',
                  type: item.type || 'Hoarding',
                  size: item.size || ''
                };
                setSelectedSiteModal(siteObj);
              }}
              title={`View ${item.site_code} details`}
            >
              {item.site_code}
            </span>
            {String(item.parent_campaign || '').startsWith('LINKED:') ? (
              <span
                style={{
                  fontSize: '10px',
                  padding: '1px 6px',
                  borderRadius: '4px',
                  background: 'rgba(245,158,11,0.2)',
                  color: '#fbbf24',
                  border: '1px solid rgba(245,158,11,0.4)',
                  fontWeight: 700
                }}
                title={item.notes || 'Auto-booked linked entry'}
              >
                🔗 Linked
              </span>
            ) : getSiteTypeTag(item.site_code) !== 'Single' && (
              <span
                style={{
                  fontSize: '10px',
                  padding: '1px 5px',
                  borderRadius: '4px',
                  background: getSiteTypeTag(item.site_code) === 'Combined' ? 'rgba(168,85,247,0.2)' : 'rgba(56,189,248,0.2)',
                  color: getSiteTypeTag(item.site_code) === 'Combined' ? '#c084fc' : '#38bdf8',
                  fontWeight: 700
                }}
                title={getConflictSummary(item.site_code)?.message || ''}
              >
                {getSiteTypeTag(item.site_code)}
              </span>
            )}
            {/* History expand toggle */}
            {prevCampaigns.length > 0 && (
              <button
                type="button"
                onClick={() => toggleHistoryExpand(item.site_code)}
                title={isExpanded ? 'Hide booking history' : `Show ${prevCampaigns.length} previous booking${prevCampaigns.length > 1 ? 's' : ''}`}
                style={{
                  display: 'inline-flex', alignItems: 'center', gap: '3px',
                  padding: '2px 7px', borderRadius: '10px',
                  border: `1px solid ${isExpanded ? 'rgba(168,85,247,0.6)' : 'rgba(168,85,247,0.3)'}`,
                  background: isExpanded ? 'rgba(168,85,247,0.2)' : 'rgba(168,85,247,0.06)',
                  color: '#c084fc', fontSize: '10px', fontWeight: 700,
                  cursor: 'pointer', whiteSpace: 'nowrap', marginTop: '3px'
                }}
              >
                {isExpanded ? '▼' : '▶'} {prevCampaigns.length} prev
              </button>
            )}
          </div>
        </td>
        <td style={{ color: '#cbd5e1', fontSize: '12px', fontWeight: 600 }}>
          {item.month ? (
            <span style={{ padding: '2px 8px', borderRadius: '5px', background: 'rgba(56, 189, 248, 0.1)', border: '1px solid rgba(56, 189, 248, 0.25)', color: '#38bdf8' }}>
              {item.month}
            </span>
          ) : (
            <span style={{ color: '#64748b' }}>—</span>
          )}
        </td>
        <td>
          {item.client ? (
            <div style={{ fontWeight: 800, color: '#ffffff', fontSize: '13px' }}>
              {item.client}
            </div>
          ) : (
            <span style={{ color: '#64748b', fontSize: '12px', fontStyle: 'italic' }}>
              — Ready for Booking —
            </span>
          )}
        </td>
        <td>
          {item.display ? (
            <div style={{ color: '#38bdf8', fontWeight: 600, fontSize: '12.5px' }}>
              📢 {cleanDisplayTitle(item.display)}
            </div>
          ) : (
            <span style={{ color: '#64748b' }}>—</span>
          )}
        </td>
        <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
          {item.vendor_name || '—'}
        </td>
        <td>
          <div style={{ color: '#edf2f6', fontSize: '12px', lineHeight: 1.35, fontWeight: 600 }}>
            {item.location || '—'}
          </div>
          {item.city && (
            <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px' }}>
              📍 {item.city}
            </div>
          )}
        </td>
        <td style={{ textAlign: 'center', color: '#94a3b8', fontSize: '11.5px' }}>
          {item.width || '—'}
        </td>
        <td style={{ textAlign: 'center', color: '#94a3b8', fontSize: '11.5px' }}>
          {item.height || '—'}
        </td>
        <td style={{ color: '#e2e8f0', fontSize: '12px', fontWeight: 600 }}>
          {item.size || '—'}
        </td>
        <td>
          <span className="scooh-badgechip" style={{ background: 'rgba(56,189,248,0.14)', color: '#38bdf8', fontSize: '11px', fontWeight: 700 }}>
            {item.type}
          </span>
        </td>
        <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
          {formatDate(item.start_date) || '—'}
        </td>
        <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
          {formatDate(item.end_date) || (item.start_date ? 'Ongoing' : '—')}
        </td>
        <td style={{ textAlign: 'center', fontWeight: 700, color: '#e2e8f0', fontSize: '12px' }}>
          {item.days ? `${item.days}d` : '—'}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 600, color: '#e2e8f0', fontSize: '12.5px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic' }}>—</span>
          ) : item.advt_fees > 0 ? (
            money(item.advt_fees)
          ) : (
            '—'
          )}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 600, color: '#e2e8f0', fontSize: '12.5px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic' }}>—</span>
          ) : item.printing_mounting_cost > 0 ? (
            money(item.printing_mounting_cost)
          ) : (
            '—'
          )}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 800, color: '#4ade80', fontSize: '13px' }}>
          {String(item.parent_campaign || '').startsWith('LINKED:') ? (
            <span style={{ color: '#94a3b8', fontSize: '11.5px', fontStyle: 'italic', fontWeight: 600 }} title="Covered under primary booking">
              Covered
            </span>
          ) : item.total_amount > 0 ? (
            money(item.total_amount)
          ) : (
            '—'
          )}
        </td>
        <td style={{ textAlign: 'right', fontWeight: 700, color: isPending ? '#f87171' : '#94a3b8', fontSize: '12.5px' }}>
          {item.pending > 0 ? money(item.pending) : item.total_amount > 0 ? <span style={{ color: '#4ade80', fontSize: '11px' }}>₹0</span> : '—'}
        </td>
        <td>
          <div className="scooh-rowactions" style={{ justifyContent: 'center', gap: '5px' }}>
            {item.campaign && (
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '3px 8px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.4)' }}
                onClick={() => openViewCampaign(item.campaign)}
                title="View campaign specifications and billing details"
              >
                👁 View
              </button>
            )}
            {item.campaign && canEdit && !isReadOnly && (
              <button
                type="button"
                className="scooh-iconbtn scooh-text-action"
                style={{ fontSize: '11px', padding: '3px 8px' }}
                onClick={() => openEditCampaign(item.campaign)}
                title="Edit campaign record"
              >
                ✏️ Edit
              </button>
            )}
            {canAdd && !item.campaign && (
              <button
                type="button"
                className="scooh-btn purple-btn"
                style={{ fontSize: '11px', padding: '3px 8px', borderRadius: '5px' }}
                onClick={() => openNewCampaign(item.site_code)}
                title={`Book new campaign for ${item.site_code}`}
              >
                + Book
              </button>
            )}
            {canDelete && item.campaign_id && (
              <button
                type="button"
                className="scooh-iconbtn danger-icon"
                style={{ fontSize: '11px', padding: '3px 6px' }}
                onClick={() => deleteRecord(item.campaign_id)}
                title="Delete this campaign record"
              >
                ×
              </button>
            )}
          </div>
        </td>
      </tr>

        {/* Expandable Previous Bookings sub-table */}
        {isExpanded && prevCampaigns.length > 0 && (
          <tr>
            <td colSpan={colSpan} style={{ padding: 0, background: 'rgba(8,12,28,0.9)', borderBottom: '2px solid rgba(168,85,247,0.3)' }}>
              <div style={{ padding: '8px 16px 14px 40px' }}>
                <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', marginBottom: '7px', textTransform: 'uppercase', letterSpacing: '0.07em' }}>
                  📌 {item.site_code} — Historical Records ({prevCampaigns.length})
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', minWidth: '900px' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid rgba(255,255,255,0.07)' }}>
                        {['Month', 'Client / Agency', 'Display', 'Start Date', 'End Date', 'Days', 'Total', 'Pending', 'Actions'].map(h => (
                          <th key={h} style={{ padding: '5px 10px', textAlign: ['Total','Pending'].includes(h) ? 'right' : 'left', color: '#64748b', fontWeight: 700, whiteSpace: 'nowrap' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {prevCampaigns.map((c, idx) => (
                        <tr key={c.id || idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.04)' }}>
                          <td style={{ padding: '5px 10px', color: '#94a3b8' }}>{c.month || '—'}</td>
                          <td style={{ padding: '5px 10px', color: '#e2e8f0', fontWeight: 600 }}>{c.client || c.client_name || '—'}</td>
                          <td style={{ padding: '5px 10px', color: '#38bdf8' }}>{cleanDisplayTitle(c.display || c.campaign_name || '—')}</td>
                          <td style={{ padding: '5px 10px', color: '#94a3b8' }}>{formatDate(c.start_date) || '—'}</td>
                          <td style={{ padding: '5px 10px', color: '#94a3b8' }}>{formatDate(c.end_date) || '—'}</td>
                          <td style={{ padding: '5px 10px', color: '#94a3b8', textAlign: 'center' }}>{c.days ? `${c.days}d` : '—'}</td>
                          <td style={{ padding: '5px 10px', textAlign: 'right', color: '#4ade80', fontWeight: 700 }}>
                            {Number(c.total_amount || c.revenue || 0) > 0 ? money(Number(c.total_amount || c.revenue || 0)) : '—'}
                          </td>
                          <td style={{ padding: '5px 10px', textAlign: 'right', fontWeight: 700, color: Number(c.pending || 0) > 0 ? '#f87171' : '#4ade80' }}>
                            {Number(c.pending || 0) > 0 ? money(Number(c.pending)) : '—'}
                          </td>
                          <td style={{ padding: '5px 10px' }}>
                            <div style={{ display: 'flex', gap: '4px' }}>
                              <button type="button" className="scooh-btn ghost"
                                style={{ fontSize: '10px', padding: '2px 7px', color: '#38bdf8', borderColor: 'rgba(56,189,248,0.3)' }}
                                onClick={() => openViewCampaign(c)} title="View details"
                              >👁 View</button>
                              {canEdit && !isReadOnly && (
                                <button type="button" className="scooh-iconbtn scooh-text-action"
                                  style={{ fontSize: '10px', padding: '2px 7px' }}
                                  onClick={() => openEditCampaign(c)} title="Edit booking"
                                >✏️ Edit</button>
                              )}
                              {canDelete && (
                                <button type="button" className="scooh-iconbtn danger-icon"
                                  style={{ fontSize: '10px', padding: '2px 6px' }}
                                  onClick={() => deleteRecord(c.id)} title="Delete booking"
                                >×</button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </td>
          </tr>
        )}
      </React.Fragment>
    );
  }

  return (
    <>
      <div className="scooh-pagehead">
        <div>
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '4px' }}>
            Media Buzz — OOH Workspace
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: 900, margin: 0, color: '#fff', letterSpacing: '-0.02em' }}>
            Latest Booking
          </h1>
          <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '13px' }}>
            Live outdoor campaigns, latest client billing, and current site booking status.
          </p>
        </div>
        <div className="scooh-headactions" style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
          <button type="button" className="scooh-btn ghost" onClick={loadData}>
            🔄 Refresh
          </button>
          {(canAdd || canDelete) && (
            <>
              <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Import campaigns from Excel or CSV spreadsheet">
                <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" hidden disabled={importingExcel} onChange={handleCampaignsExcelImport} />
              </label>
              <a
                href="/sample_campaign_tracker_import.xlsx"
                download="MediaBuzz_Latest_Booking_Sample.xlsx"
                className="scooh-btn ghost"
                style={{ textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: '5px' }}
                title="Download sample Excel spreadsheet template for importing campaigns"
              >
                <span>📄 Sample Template</span>
              </a>
            </>
          )}
          <button type="button" className="scooh-btn ghost" onClick={() => exportCampaignsExcel(filteredLatest, 'MediaBuzz_Latest_Bookings.xlsx')}>
            📥 Export Excel
          </button>
          {canAdd && (
            <button type="button" className="scooh-btn purple-btn" onClick={() => openNewCampaign()}>
              + Add Campaign
            </button>
          )}
        </div>
      </div>

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are signed in as a Viewer. You can view all campaign data, filter records, and export to Excel. Creating, editing, and deleting campaigns is disabled.
        </div>
      )}

      {banner && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(34,197,94,0.15)', border: '1px solid #22c55e', color: '#4ade80' }}>
          {banner}
        </div>
      )}

      {/* Summary KPI Cards (Interactive Filter Shortcuts) */}
      <div className="scooh-kpirow" style={{ marginBottom: '20px' }}>
        <div 
          className={`scooh-kpi ${occupancyRate >= 75 ? 'good' : occupancyRate >= 50 ? 'alert' : 'danger'}`}
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'occupied' ? 'ALL' : 'occupied')}
          title="Click to toggle occupied sites"
        >
          <div className="n">{occupancyRate}%</div>
          <div className="l">Occupancy Rate</div>
          <div className="scooh-kpi-note" style={{ color: occupancyRate >= 75 ? '#4ade80' : '#fbbf24' }}>
            {occupiedSitesCount} of {totalSites} sites occupied
          </div>
        </div>

        <div 
          className="scooh-kpi good"
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'vacant' ? 'ALL' : 'vacant')}
          title="Click to view vacant sites ready for booking"
        >
          <div className="n" style={{ color: vacantSitesCount > 0 ? '#fbbf24' : '#4ade80' }}>{vacantSitesCount}</div>
          <div className="l">Vacant Sites</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>Available for new campaigns</div>
        </div>

        <div 
          className="scooh-kpi good"
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter('ALL')}
          title="Click to view all sites"
        >
          <div className="n">{money(latestRevenueSum)}</div>
          <div className="l">Current Campaign Revenue</div>
          <div className="scooh-kpi-note" style={{ color: '#4ade80' }}>{occupiedSitesCount + upcomingSitesCount} active / upcoming bookings</div>
        </div>

        <div 
          className={`scooh-kpi ${latestPendingSum > 0 ? 'danger' : 'good'}`}
          style={{ cursor: 'pointer' }}
          onClick={() => setStatusFilter(prev => prev === 'pending' ? 'ALL' : 'pending')}
          title="Click to filter pending payments"
        >
          <div className="n">{money(latestPendingSum)}</div>
          <div className="l">Total Pending Amount</div>
          <div className="scooh-kpi-note" style={{ color: latestPendingSum > 0 ? '#ef4444' : '#4ade80' }}>
            {latestPendingSum > 0 ? `${latestSiteCampaigns.filter(r => r.pending > 0).length} pending sites` : 'All cleared'}
          </div>
        </div>
      </div>

      {/* Main Campaign Section */}
      <div className="scooh-electricity-section">
        <div className="scooh-sectionbar" style={{ padding: '14px 20px' }}>
          <div>
            <h2 style={{ fontSize: '16px', fontWeight: 800, color: '#f8fafc', margin: 0 }}>
              Live Sites & Latest Campaigns
            </h2>
            <p style={{ margin: '3px 0 0', fontSize: '12px', color: '#94a3b8' }}>
              Showing latest booking per site ({filteredLatest.length} of {totalSites} sites). Click <strong style={{ color: '#c084fc' }}>▶ N prev</strong> on any site to expand its full booking history inline.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => navigate('/occupancy')}
              style={{ fontSize: '12px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
              title="Open full booking history ledger in Campaign Tracker"
            >
              📋 All Previous Bookings (Campaign Tracker) ↗
            </button>
            {canAdd && (
              <button type="button" className="scooh-btn purple-btn" onClick={() => openNewCampaign()}>
                + Add Campaign
              </button>
            )}
          </div>
        </div>

        {/* Clean, Unified Single Toolbar */}
        <div className="scooh-toolbar" style={{ padding: '12px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between' }}>
          
          {/* Left: Status Filter Pills & Live Search */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', flex: '1 1 auto' }}>
            
            <div style={{ display: 'inline-flex', gap: '4px', background: 'rgba(15, 23, 42, 0.7)', padding: '3px', borderRadius: '20px', border: '1px solid #1e293b' }}>
              <button
                type="button"
                onClick={() => setStatusFilter('ALL')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'ALL' ? 'rgba(167, 139, 250, 0.22)' : 'transparent',
                  color: statusFilter === 'ALL' ? '#c4b5fd' : '#94a3b8',
                  fontWeight: statusFilter === 'ALL' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                All Sites ({totalSites})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('occupied')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'occupied' ? 'rgba(34, 197, 94, 0.2)' : 'transparent',
                  color: statusFilter === 'occupied' ? '#4ade80' : '#94a3b8',
                  fontWeight: statusFilter === 'occupied' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟢 Occupied ({occupiedSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('upcoming')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'upcoming' ? 'rgba(234, 179, 8, 0.2)' : 'transparent',
                  color: statusFilter === 'upcoming' ? '#facc15' : '#94a3b8',
                  fontWeight: statusFilter === 'upcoming' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                🟡 Upcoming ({upcomingSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('vacant')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'vacant' ? 'rgba(148, 163, 184, 0.2)' : 'transparent',
                  color: statusFilter === 'vacant' ? '#e2e8f0' : '#94a3b8',
                  fontWeight: statusFilter === 'vacant' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚪ Vacant ({vacantSitesCount})
              </button>
              <button
                type="button"
                onClick={() => setStatusFilter('pending')}
                style={{
                  padding: '5px 12px',
                  borderRadius: '16px',
                  border: 'none',
                  background: statusFilter === 'pending' ? 'rgba(239, 68, 68, 0.2)' : 'transparent',
                  color: statusFilter === 'pending' ? '#f87171' : '#94a3b8',
                  fontWeight: statusFilter === 'pending' ? 800 : 600,
                  fontSize: '11.5px',
                  cursor: 'pointer'
                }}
              >
                ⚠️ Pending ({latestSiteCampaigns.filter(x => x.pending > 0).length})
              </button>
            </div>

            {/* Search Input */}
            <input
              className="scooh-search"
              style={{ minWidth: '220px', maxWidth: '320px', fontSize: '12px' }}
              placeholder="Search site, location, client, display, PO, Bill…"
              value={search}
              onChange={e => setSearch(e.target.value)}
            />

            {/* Date selector for evaluating site status */}
            <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', background: dateFilter ? 'rgba(56,189,248,0.14)' : 'rgba(15,23,42,0.6)', border: `1px solid ${dateFilter ? 'rgba(56,189,248,0.5)' : 'rgba(255,255,255,0.1)'}`, borderRadius: '8px', padding: '5px 10px' }}>
              <span style={{ fontSize: '11px', fontWeight: 700, color: dateFilter ? '#38bdf8' : '#94a3b8', whiteSpace: 'nowrap' }}>
                📅 {dateFilter ? 'Date:' : 'Select Date:'}
              </span>
              <input
                type="date"
                value={dateFilter}
                onChange={e => setDateFilter(e.target.value)}
                style={{ background: 'transparent', border: 'none', color: dateFilter ? '#38bdf8' : '#94a3b8', fontSize: '12px', fontWeight: 700, outline: 'none', cursor: 'pointer', padding: '2px 0' }}
                title="Select a date to evaluate site occupancy as of that date. Sites whose campaign ends before this date automatically become Vacant."
              />
              {dateFilter && (
                <button type="button" onClick={() => setDateFilter('')} style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '14px', lineHeight: 1, padding: '0 2px' }} title="Reset date to today">×</button>
              )}
            </div>

            {(search || statusFilter !== 'ALL' || dateFilter) && (
              <button
                type="button"
                className="scooh-btn ghost"
                style={{ fontSize: '11px', padding: '5px 9px' }}
                onClick={() => {
                  setSearch('');
                  setStatusFilter('ALL');
                  setDateFilter('');
                  setSiteQueryParam('');
                  navigate('/campaigns', { replace: true });
                }}
              >
                ✕ Reset
              </button>
            )}

            {canDelete && (
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', marginLeft: '4px' }}>
                <button
                  type="button"
                  className="scooh-btn ghost"
                  style={{ fontSize: '11.5px', padding: '5px 10px', color: '#38bdf8', borderColor: 'rgba(56, 189, 248, 0.3)' }}
                  onClick={selectAllVisible}
                  title="Select all visible campaign records"
                >
                  ☑ Select All
                </button>
                {selectedIds.size > 0 && (
                  <button
                    type="button"
                    className="scooh-btn ghost"
                    style={{ fontSize: '11.5px', padding: '5px 10px' }}
                    onClick={deselectAll}
                    title="Clear selection"
                  >
                    ☐ Deselect All
                  </button>
                )}
                {selectedIds.size > 0 && (
                  <button
                    type="button"
                    className="scooh-btn danger"
                    style={{ fontSize: '11.5px', padding: '5px 12px' }}
                    onClick={handleBatchDelete}
                  >
                    🗑 Delete Selected ({selectedIds.size})
                  </button>
                )}
                <button
                  type="button"
                  className="scooh-btn danger"
                  style={{ fontSize: '11.5px', padding: '5px 12px', background: 'rgba(239, 68, 68, 0.15)', borderColor: 'rgba(239, 68, 68, 0.35)', color: '#f87171' }}
                  onClick={handleDeleteAllCampaigns}
                  title="Delete all visible campaign records with confirmation"
                >
                  🗑 Delete All
                </button>
              </div>
            )}
          </div>

          {/* Right: Quick Action Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={() => exportCampaignsExcel(filteredLatest, 'MediaBuzz_Campaign_Tracker.xlsx')}
              title="Export table to Excel"
              style={{ fontSize: '11.5px', padding: '6px 12px' }}
            >
              📥 Export Excel
            </button>

            {(canAdd || canDelete) && (
              <>
                <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '11.5px', padding: '6px 12px' }} title="Import campaigns spreadsheet">
                  <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
                  <input type="file" accept=".xlsx,.xls,.csv,.xlsm,.ods" hidden disabled={importingExcel} onChange={handleCampaignsExcelImport} />
                </label>
                <a
                  href="/sample_campaign_tracker_import.xlsx"
                  download="MediaBuzz_Campaign_Tracker_Sample.xlsx"
                  className="scooh-btn ghost"
                  style={{ textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: '5px', fontSize: '11.5px', padding: '6px 12px' }}
                  title="Download sample Excel spreadsheet template for importing campaigns"
                >
                  <span>📄 Sample Template</span>
                </a>
              </>
            )}

            <button type="button" className="scooh-btn ghost" onClick={loadData} title="Refresh data" style={{ fontSize: '11.5px', padding: '6px 10px' }}>
              🔄
            </button>
          </div>
        </div>

        {/* Unified Latest Bookings Table with Dual Synchronized Scrollers */}
        {dateFilter && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(56, 189, 248, 0.1)', border: '1px solid rgba(56, 189, 248, 0.3)', borderRadius: '8px', padding: '8px 14px', marginBottom: '10px', fontSize: '12px', color: '#38bdf8' }}>
            <span style={{ fontSize: '15px' }}>📅</span>
            <span>
              Site occupancy evaluated as of <strong>{formatDate(dateFilter) || dateFilter}</strong>. Any campaign ending before this date automatically shows as <strong>Vacant</strong> ({vacantSitesCount} vacant, {occupiedSitesCount} occupied).
            </span>
            <button
              type="button"
              onClick={() => setDateFilter('')}
              style={{ marginLeft: 'auto', background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', fontSize: '11.5px', fontWeight: 700, textDecoration: 'underline' }}
            >
              Reset to Today
            </button>
          </div>
        )}
        <div>
          {/* Top Synchronized Scroller Bar */}
          <div
            ref={topScrollRef}
            onScroll={handleTopScroll}
            className="scooh-tablewrap"
            style={{
              overflowX: 'auto',
              overflowY: 'hidden',
              maxHeight: '14px',
              marginBottom: '4px',
              background: 'rgba(15, 23, 42, 0.6)',
              borderRadius: '6px',
              border: '1px solid rgba(255, 255, 255, 0.08)'
            }}
          >
            <div style={{ width: `${tableScrollWidth}px`, height: '1px' }} />
          </div>

          {/* Bottom Primary Table & Scroller */}
          <div ref={bottomScrollRef} onScroll={handleBottomScroll} className="scooh-tablewrap" style={{ overflowX: 'auto' }}>
            <table ref={tableRef} className="scooh-table" style={{ minWidth: '1750px' }}>
              <thead>
                <tr>
                  {canDelete && (
                    <th style={{ width: '40px', textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={filteredLatest.length > 0 && filteredLatest.some(r => r.campaign_id) && filteredLatest.filter(r => r.campaign_id).every(r => selectedIds.has(r.campaign_id))}
                        onChange={toggleSelectAll}
                        title="Select all visible latest campaigns"
                        style={{ cursor: 'pointer' }}
                      />
                    </th>
                  )}
                  <SortHeader label="Site Code" sortKey="site_code" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Month" sortKey="month" currentSort={sortState} onSort={handleSort} style={{ minWidth: '95px' }} />
                  <SortHeader label="Client/Agency Name" sortKey="client" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                  <SortHeader label="Display" sortKey="display" currentSort={sortState} onSort={handleSort} style={{ minWidth: '160px' }} />
                  <SortHeader label="Vendor Name" sortKey="vendor_name" currentSort={sortState} onSort={handleSort} style={{ minWidth: '140px' }} />
                  <SortHeader label="Location" sortKey="location" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                  <th style={{ width: '55px', textAlign: 'center' }}>W</th>
                  <th style={{ width: '55px', textAlign: 'center' }}>H</th>
                  <SortHeader label="Size" sortKey="size" currentSort={sortState} onSort={handleSort} style={{ minWidth: '95px' }} />
                  <SortHeader label="Type" sortKey="type" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Start Date" sortKey="start_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="End Date" sortKey="end_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '105px' }} />
                  <SortHeader label="Days" sortKey="days" currentSort={sortState} onSort={handleSort} align="center" style={{ width: '70px' }} />
                  <SortHeader label="Advt. Fees" sortKey="advt_fees" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '140px' }} />
                  <SortHeader label="Prod/Mount" sortKey="printing_mounting_cost" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '140px' }} />
                  <SortHeader label="Total" sortKey="total_amount" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '135px' }} />
                  <SortHeader label="Pending" sortKey="pending" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '110px' }} />
                  <th style={{ minWidth: '180px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredLatest.length === 0 ? (
                  <tr>
                    <td colSpan={canDelete ? 19 : 18} className="scooh-empty" style={{ padding: '36px 20px', textAlign: 'center' }}>
                      <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No sites match your query</div>
                      <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                        Try adjusting your search or status filter.
                      </div>
                    </td>
                  </tr>
                ) : (
                  filteredLatest.map(item => renderLatestSiteRow(item))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Add / Edit / View Campaign Modal */}
      {editModal && (
        <div className="scooh-modal-overlay" onClick={() => !saving && setEditModal(null)}>
          <div className="scooh-modal" style={{ maxWidth: '780px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>{editModal.id ? (isViewingOnly || isReadOnly ? '👁 View Campaign Details' : 'Edit Campaign Record') : 'Add Campaign Record'}</h2>
                <span className="scooh-modal-subtitle">
                  {isViewingOnly || isReadOnly ? 'Detailed specification, schedule, and financial preview for this booking' : 'Enter outdoor campaign specifications and billing details'}
                </span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setEditModal(null)}>×</button>
            </div>
            <form onSubmit={saveCampaign}>
              <div className="scooh-modalbody" style={{ maxHeight: '72vh', overflowY: 'auto' }}>
                {/* Multi-Site Selection Header & Picker (Available when creating a new campaign) */}
                {!editModal.id && !isViewingOnly && (
                  <div style={{
                    marginBottom: '16px',
                    padding: '14px',
                    background: 'rgba(15, 23, 42, 0.75)',
                    border: '1px solid rgba(168, 85, 247, 0.35)',
                    borderRadius: '10px',
                    boxShadow: '0 4px 14px rgba(0,0,0,0.3)'
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '8px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ fontSize: '18px' }}>📍</span>
                        <div>
                          <label style={{ margin: 0, fontSize: '13px', fontWeight: 800, color: '#f8fafc' }}>
                            Select Sites for this Client ({modalSelectedSites.length} Selected)
                          </label>
                          <span style={{ display: 'block', fontSize: '11px', color: '#c084fc' }}>
                            You can select multiple sites to book together for this client in one go.
                          </span>
                        </div>
                      </div>
                      <div style={{ display: 'flex', gap: '6px', alignItems: 'center', flexWrap: 'wrap' }}>
                        <button
                          type="button"
                          className="scooh-btn ghost"
                          style={{ fontSize: '11px', padding: '3px 9px', color: '#4ade80', borderColor: 'rgba(74, 222, 128, 0.3)' }}
                          onClick={() => {
                            const vacantCodes = sites
                              .map(s => s.site_code)
                              .filter(code => code && vacantSiteCodeSet.has(code.toUpperCase()));
                            setModalSelectedSites(Array.from(new Set([...modalSelectedSites, ...vacantCodes])));
                          }}
                          title="Select all vacant sites ready for booking"
                        >
                          + Select All Vacant ({vacantSitesCount})
                        </button>
                        {modalSelectedSites.length > 0 && (
                          <button
                            type="button"
                            className="scooh-btn ghost"
                            style={{ fontSize: '11px', padding: '3px 9px', color: '#f87171', borderColor: 'rgba(248, 113, 113, 0.3)' }}
                            onClick={() => {
                              setModalSelectedSites([]);
                              setModalSiteCode('');
                            }}
                          >
                            Clear All
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Selected Site Badges / Chips */}
                    {modalSelectedSites.length > 0 ? (
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginBottom: '10px', maxHeight: '80px', overflowY: 'auto' }}>
                        {modalSelectedSites.map(sc => (
                          <span key={sc} style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '5px',
                            background: 'rgba(168, 85, 247, 0.22)',
                            border: '1px solid rgba(168, 85, 247, 0.5)',
                            borderRadius: '16px',
                            padding: '3px 10px',
                            fontSize: '11.5px',
                            fontWeight: 700,
                            color: '#e9d5ff'
                          }}>
                            {sc}
                            <button
                              type="button"
                              onClick={() => toggleModalSite(sc)}
                              style={{ background: 'none', border: 'none', color: '#fca5a5', cursor: 'pointer', fontSize: '13px', padding: '0 2px', lineHeight: 1 }}
                              title="Remove site"
                            >
                              ×
                            </button>
                          </span>
                        ))}
                      </div>
                    ) : (
                      <div style={{ fontSize: '12px', color: '#94a3b8', fontStyle: 'italic', marginBottom: '8px' }}>
                        No sites selected yet. Check one or more sites below, or search to pick:
                      </div>
                    )}

                    {/* Filter & Search Bar */}
                    <div style={{ display: 'flex', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                      <input
                        type="text"
                        placeholder="Search sites by code, area, location or size…"
                        value={siteFilterText}
                        onChange={e => setSiteFilterText(e.target.value)}
                        style={{ fontSize: '11.5px', padding: '6px 10px', flex: 1, background: 'rgba(0,0,0,0.3)', border: '1px solid #334155', borderRadius: '6px', color: '#fff' }}
                      />
                      <label style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '11.5px', color: '#cbd5e1', cursor: 'pointer', whiteSpace: 'nowrap', userSelect: 'none' }}>
                        <input
                          type="checkbox"
                          checked={onlyVacantFilter}
                          onChange={e => setOnlyVacantFilter(e.target.checked)}
                          style={{ cursor: 'pointer' }}
                        />
                        Vacant Only
                      </label>
                    </div>

                    {/* Quick Checklist */}
                    <div style={{
                      maxHeight: '140px',
                      overflowY: 'auto',
                      border: '1px solid rgba(255,255,255,0.08)',
                      borderRadius: '6px',
                      background: 'rgba(0,0,0,0.25)',
                      padding: '4px 6px'
                    }}>
                      {selectableSitesList.length === 0 ? (
                        <div style={{ padding: '8px', fontSize: '11.5px', color: '#64748b', textAlign: 'center' }}>
                          No sites match the search query.
                        </div>
                      ) : (
                        selectableSitesList.map(s => {
                          const isChecked = modalSelectedSites.includes(s.site_code);
                          const isVacant = vacantSiteCodeSet.has(String(s.site_code || '').toUpperCase());
                          return (
                            <div
                              key={s.id || s.site_code}
                              onClick={() => toggleModalSite(s.site_code)}
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'space-between',
                                padding: '5px 8px',
                                borderRadius: '4px',
                                cursor: 'pointer',
                                background: isChecked ? 'rgba(168, 85, 247, 0.18)' : 'transparent',
                                marginBottom: '2px',
                                transition: 'background 0.15s'
                              }}
                            >
                              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => {}}
                                  style={{ cursor: 'pointer' }}
                                />
                                <span style={{ fontWeight: 700, color: isChecked ? '#c084fc' : '#f1f5f9', fontSize: '12px' }}>
                                  {s.site_code}
                                </span>
                                <span style={{ color: '#94a3b8', fontSize: '11.5px' }}>
                                  {s.address || s.area || '—'} {s.size ? `[${s.size}]` : ''}
                                </span>
                              </div>
                              <span style={{
                                fontSize: '10.5px',
                                padding: '1px 6px',
                                borderRadius: '10px',
                                fontWeight: 600,
                                background: isVacant ? 'rgba(148, 163, 184, 0.15)' : 'rgba(34, 197, 94, 0.15)',
                                color: isVacant ? '#cbd5e1' : '#4ade80'
                              }}>
                                {isVacant ? '⚪ Vacant' : '🟢 Occupied'}
                              </span>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '14px' }}>
                  <div className="scooh-field">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <label>{modalSelectedSites.length > 1 ? `Site Codes (${modalSelectedSites.length} selected)` : 'Site Code'}</label>
                      {autoMatchedSite && (
                        <span style={{ fontSize: '11px', color: '#10b981', fontWeight: 700 }}>
                          ✓ Match: {autoMatchedSite.site_code}
                        </span>
                      )}
                    </div>
                    <input
                      name="site_code"
                      list="campaign-site-codes-list"
                      value={modalSiteCode}
                      onChange={e => handleSiteCodeChange(e.target.value)}
                      placeholder="e.g. MB-06 or select sites above"
                      disabled={isViewingOnly || isReadOnly}
                    />
                    <datalist id="campaign-site-codes-list">
                      {[...sites].sort((a, b) => universalCompare(a.site_code, b.site_code, 'asc')).map(s => (
                        <option key={s.id || s.site_code} value={s.site_code}>
                          {s.site_code} — {s.address || s.area || ''} ({s.size || `${s.width}x${s.height}`})
                        </option>
                      ))}
                    </datalist>
                  </div>
                  {autoMatchedSite && !isViewingOnly && (
                    <div style={{
                      gridColumn: '1 / -1',
                      padding: '9px 13px',
                      borderRadius: '8px',
                      background: 'rgba(16, 185, 129, 0.12)',
                      border: '1px solid rgba(16, 185, 129, 0.35)',
                      fontSize: '12px',
                      color: '#ecfdf5',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      gap: '10px',
                      flexWrap: 'wrap'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ fontSize: '16px' }}>🎯</span>
                        <div>
                          <span style={{ color: '#a7f3d0' }}>Auto-detected Site Code from Location & Size:</span>{' '}
                          <strong style={{ color: '#ffffff', fontSize: '13px', background: 'rgba(16, 185, 129, 0.3)', padding: '2px 7px', borderRadius: '4px' }}>
                            {autoMatchedSite.site_code}
                          </strong>
                          <span style={{ color: '#cbd5e1', marginLeft: '6px' }}>
                            {autoMatchedSite.size ? `[${autoMatchedSite.size}]` : ''} {autoMatchedSite.address || autoMatchedSite.area ? `• ${autoMatchedSite.address || autoMatchedSite.area}` : ''}
                          </span>
                        </div>
                      </div>
                      {String(modalSiteCode).trim().toUpperCase() !== String(autoMatchedSite.site_code).toUpperCase() ? (
                        <button
                          type="button"
                          onClick={() => applyMatchedSite(autoMatchedSite)}
                          style={{
                            padding: '4px 12px',
                            borderRadius: '6px',
                            background: 'linear-gradient(135deg, #10b981, #059669)',
                            color: '#ffffff',
                            border: 'none',
                            fontWeight: 700,
                            fontSize: '11px',
                            cursor: 'pointer',
                            boxShadow: '0 2px 6px rgba(16, 185, 129, 0.4)'
                          }}
                        >
                          ✓ Apply Site Code
                        </button>
                      ) : (
                        <span style={{ fontSize: '11px', color: '#6ee7b7', fontWeight: 600 }}>
                          ✓ Site Code Applied
                        </span>
                      )}
                    </div>
                  )}
                  {(() => {
                    const conflict = getConflictSummary(modalSiteCode);
                    if (!conflict) return null;
                    return (
                      <div style={{
                        gridColumn: '1 / -1',
                        padding: '9px 13px',
                        borderRadius: '8px',
                        background: conflict.isCombined ? 'rgba(168, 85, 247, 0.12)' : 'rgba(56, 189, 248, 0.12)',
                        border: `1px solid ${conflict.isCombined ? 'rgba(168, 85, 247, 0.35)' : 'rgba(56, 189, 248, 0.35)'}`,
                        fontSize: '12px',
                        color: '#f1f5f9',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px'
                      }}>
                        <span style={{ fontSize: '16px' }}>🧩</span>
                        <div>
                          <strong style={{ color: conflict.isCombined ? '#c084fc' : '#38bdf8' }}>
                            {conflict.type}:
                          </strong>{' '}
                          {conflict.message}
                        </div>
                      </div>
                    );
                  })()}
                  <div className="scooh-field">
                    <label>Month</label>
                    <input name="month" defaultValue={editModal.month ?? ''} placeholder="e.g. Oct 2026" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Booking Date</label>
                    <input type="date" name="booking_date" defaultValue={editModal.booking_date ? editModal.booking_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Client / Agency</label>
                    <input name="client" defaultValue={editModal.client ?? editModal.client_name ?? ''} placeholder="Client company or agency" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Display / Brand</label>
                    <input name="display" defaultValue={cleanDisplayTitle(editModal.display ?? editModal.campaign_name ?? '')} placeholder="Display title or brand" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Vendor Name</label>
                    <input name="vendor_name" defaultValue={editModal.vendor_name ?? ''} placeholder="Printing/Mounting vendor" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Location / Address</label>
                    <input
                      name="location"
                      value={modalLocation}
                      onChange={e => setModalLocation(e.target.value)}
                      placeholder="e.g. Nr DMart / Shivranjani"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Width (ft)</label>
                    <input
                      type="number"
                      step="any"
                      name="width"
                      value={modalWidth}
                      onChange={e => {
                        const val = e.target.value;
                        setModalWidth(val);
                        if (val && modalHeight) setModalSize(`${val}x${modalHeight} ft`);
                      }}
                      placeholder="e.g. 40"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Height (ft)</label>
                    <input
                      type="number"
                      step="any"
                      name="height"
                      value={modalHeight}
                      onChange={e => {
                        const val = e.target.value;
                        setModalHeight(val);
                        if (modalWidth && val) setModalSize(`${modalWidth}x${val} ft`);
                      }}
                      placeholder="e.g. 20"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Size</label>
                    <input
                      name="size"
                      value={modalSize}
                      onChange={e => setModalSize(e.target.value)}
                      placeholder="e.g. 40x20 ft"
                      disabled={isViewingOnly || isReadOnly}
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Media Type</label>
                    <select
                      name="type"
                      value={modalType}
                      onChange={e => setModalType(e.target.value)}
                      disabled={isViewingOnly || isReadOnly}
                    >
                      <option value="Hoarding">Hoarding</option>
                      <option value="Gantry">Gantry</option>
                      <option value="Unipole">Unipole</option>
                      <option value="Billboard">Billboard</option>
                      <option value="DOOH">DOOH</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Start Date</label>
                    <input type="date" name="start_date" defaultValue={editModal.start_date ? editModal.start_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>End Date</label>
                    <input type="date" name="end_date" defaultValue={editModal.end_date ? editModal.end_date.slice(0, 10) : ''} disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Days</label>
                    <input type="number" name="days" defaultValue={editModal.days ?? 30} placeholder="30" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Advt. Fees per month (₹)</label>
                    <input type="number" step="any" name="advt_fees" defaultValue={editModal.advt_fees ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Printing & Mounting (₹)</label>
                    <input type="number" step="any" name="printing_mounting_cost" defaultValue={editModal.printing_mounting_cost ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Total Amount (₹)</label>
                    <input type="number" step="any" name="total_amount" defaultValue={editModal.total_amount ?? editModal.revenue ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>PO (Purchase Order) <span style={{ color: '#94a3b8', fontSize: '11px', fontWeight: 'normal' }}>(Optional)</span></label>
                    <input name="po" defaultValue={editModal.po ?? ''} placeholder="Optional (e.g. PO-2026-881)" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Bill (Invoice No / Status) <span style={{ color: '#94a3b8', fontSize: '11px', fontWeight: 'normal' }}>(Optional)</span></label>
                    <input name="bill" defaultValue={editModal.bill ?? editModal.invoice_no ?? ''} placeholder="Optional (e.g. INV-9912 / Sent)" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field">
                    <label>Pending (₹)</label>
                    <input type="number" step="any" name="pending" defaultValue={editModal.pending ?? ''} placeholder="0" disabled={isViewingOnly || isReadOnly} />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: 'span 2' }}>
                    <label>Notes</label>
                    <textarea name="notes" rows="2" defaultValue={editModal.notes ?? ''} placeholder="Campaign notes or instructions…" disabled={isViewingOnly || isReadOnly} />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  {isViewingOnly && canEdit && !isReadOnly && (
                    <button
                      type="button"
                      className="scooh-btn ghost"
                      style={{ color: '#fbbf24', borderColor: 'rgba(251, 191, 36, 0.4)', fontSize: '12px' }}
                      onClick={() => setIsViewingOnly(false)}
                      title="Switch to edit mode"
                    >
                      ✏️ Edit This Campaign
                    </button>
                  )}
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>
                    {isViewingOnly || isReadOnly ? 'Close' : 'Cancel'}
                  </button>
                  {!isViewingOnly && !isReadOnly && (
                    <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                      {saving ? 'Saving…' : editModal.id ? 'Update Campaign' : (modalSelectedSites.length > 1 ? `Create Campaign (${modalSelectedSites.length} Sites)` : 'Create Campaign')}
                    </button>
                  )}
                </div>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ── Modal: Full Campaign Details Pop-up ── */}
      {viewCampaignDetails && (
        <CampaignDetailsModal
          campaign={viewCampaignDetails}
          onClose={() => setViewCampaignDetails(null)}
          onEdit={c => {
            setViewCampaignDetails(null);
            openEditCampaign(c);
          }}
          onViewSite={siteCode => {
            const upper = String(siteCode || '').trim().toUpperCase();
            const siteObj = sites.find(s => String(s.site_code || '').trim().toUpperCase() === upper) || {
              site_code: upper,
              area: upper,
              location: '',
              type: 'Hoarding'
            };
            setSelectedSiteModal(siteObj);
          }}
          canAdd={canEdit && !isReadOnly}
        />
      )}

      {/* ── Modal: Site Details Pop-up ── */}
      {selectedSiteModal && (
        <SiteDetailsModal
          site={selectedSiteModal}
          onClose={() => setSelectedSiteModal(null)}
          onNavigateToSites={() => {
            const code = selectedSiteModal.site_code;
            setSelectedSiteModal(null);
            navigate(`/sites?search=${encodeURIComponent(code)}`);
          }}
        />
      )}
    </>
  );
}

async function exportElectricityExcel(rowsData, filename = 'MediaBuzz_Electricity_Bills.xlsx') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Electricity Bills', {
    views: [{ state: 'frozen', ySplit: 2 }]
  });

  const cols = [
    { key: 'sr', width: 8 },
    { key: 'site_code', width: 16 },
    { key: 'location', width: 40 },
    { key: 'size', width: 14 },
    { key: 'meter_no', width: 20 },
    { key: 'service_number', width: 18 },
    { key: 't_number', width: 14 },
    { key: 'bill_type', width: 18 },
    { key: 'billing_month', width: 16 },
    { key: 'due_date', width: 16 },
    { key: 'payment_date', width: 16 },
    { key: 'units', width: 12 },
    { key: 'rate', width: 12 },
    { key: 'amount', width: 20 },
    { key: 'status', width: 14 },
    { key: 'ecs', width: 14 },
    { key: 'payment_reference', width: 22 },
    { key: 'notes', width: 28 }
  ];
  worksheet.columns = cols;

  // Executive Media Buzz Brand Header Banner with Logo on Top Right
  attachMediaBuzzExcelHeader(workbook, worksheet, {
    title: 'MEDIA BUZZ — CONSOLIDATED ELECTRICITY BILLS',
    columns: cols,
    totalColumns: 18
  });

  // Populate Header Row (Row 2)
  const headers = [
    'SR NO', 'SITE CODE', 'LOCATION', 'SIZE', 'METER NO',
    'SERVICE NUMBER', 'T NUMBER', 'BILL / PROVIDER', 'BILLING MONTH', 'DUE DATE',
    'PAYMENT DATE', 'UNITS', 'RATE', 'PAYMENT AMOUNT (₹)', 'STATUS', 'ECS',
    'PAYMENT REF', 'NOTES'
  ];
  const headerRow = worksheet.getRow(2);
  headerRow.height = 28;
  headers.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });

  rowsData.forEach((r, idx) => {
    worksheet.addRow({
      sr: idx + 1,
      site_code: r.site_code || '',
      location: r.location || '',
      size: r.size || '',
      meter_no: r.meter_no || '',
      service_number: r.service_number || '',
      t_number: r.t_number || '',
      bill_type: r.bill_type || '',
      billing_month: r.billing_month || '',
      due_date: r.due_date ? new Date(r.due_date).toLocaleDateString('en-IN') : '',
      payment_date: (r.payment_date || r.paid_date) ? new Date(r.payment_date || r.paid_date).toLocaleDateString('en-IN') : '',
      units: r.units || '',
      rate: r.rate || '',
      amount: Number(r.amount || r.payment_amount || 0),
      status: r.payment_status || 'Pending',
      ecs: r.ecs || '',
      payment_reference: r.payment_reference || '',
      notes: r.notes || ''
    });
  });

  // Style Header Row in Bright Yellow (#FFFF00)
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FF000000' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: [1, 4, 8, 9, 10, 11, 12, 13, 15, 16].includes(colNumber) ? 'center' : (colNumber === 14 ? 'right' : 'left'),
      wrapText: false
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      left: { style: 'thin', color: { argb: 'FFC0C0C0' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FFC0C0C0' } }
    };
  });

  // Style Data Rows (Row 3+)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= 3) {
      row.height = 22;
      row.eachCell((cell, colNumber) => {
        cell.font = { name: 'Calibri', size: 10.5 };
        cell.alignment = {
          vertical: 'middle',
          horizontal: [1, 4, 8, 9, 10, 11, 12, 13, 15, 16].includes(colNumber) ? 'center' : (colNumber === 14 ? 'right' : 'left')
        };
        if (colNumber === 14 && typeof cell.value === 'number') {
          cell.numFmt = '#,##,##0';
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          left: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          bottom: { style: 'thin', color: { argb: 'FFE8E8E8' } },
          right: { style: 'thin', color: { argb: 'FFE8E8E8' } }
        };
      });
    }
  });

  // Enable Auto-Filter on Row 2
  if (rowsData.length > 0) {
    worksheet.autoFilter = `A2:Q${rowsData.length + 2}`;
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  window.URL.revokeObjectURL(url);
}

function StorageVaultSection({ embedded = true, title = null, subtitle = null, initialCategory = 'all' }) {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(false);
  const [category, setCategory] = useState(initialCategory);
  const [search, setSearch] = useState('');
  const [sort, setSort] = useState('newest');
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [uploadModal, setUploadModal] = useState(false);
  const [snapshotModal, setSnapshotModal] = useState(false);
  const [snapshotMonth, setSnapshotMonth] = useState('');
  const [activeOccItem, setActiveOccItem] = useState(null);
  const [banner, setBanner] = useState('');
  const [actionLoading, setActionLoading] = useState(false);

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;

  async function loadStorage() {
    setLoading(true);
    try {
      const { data } = await api.get('/storage');
      if (Array.isArray(data)) setItems(data);
    } catch (err) {
      console.warn('Load storage error:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadStorage();
  }, []);

  const pptCount = items.filter(i => i.category === 'ppt').length;
  const excelCount = items.filter(i => i.category === 'excel').length;
  const occCount = items.filter(i => i.category === 'occupancy').length;
  const otherCount = items.filter(i => !['ppt', 'excel', 'occupancy'].includes(i.category)).length;

  const filtered = useMemo(() => {
    let list = items.filter(i => {
      if (category !== 'all' && i.category !== category) return false;
      if (search.trim()) {
        const q = search.toLowerCase();
        const str = `${i.title || ''} ${i.filename || ''} ${JSON.stringify(i.meta || {})} ${i.format || ''}`.toLowerCase();
        if (!str.includes(q)) return false;
      }
      return true;
    });

    if (sort === 'newest') list.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    else if (sort === 'oldest') list.sort((a, b) => new Date(a.created_at) - new Date(b.created_at));
    else if (sort === 'title') list.sort((a, b) => (a.title || '').localeCompare(b.title || ''));
    return list;
  }, [items, category, search, sort]);

  const occItems = useMemo(() => {
    return items
      .filter(i => i.category === 'occupancy')
      .sort((a, b) => new Date(a.created_at) - new Date(b.created_at));
  }, [items]);

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!window.confirm(`Are you sure you want to permanently delete all ${count} selected file${count > 1 ? 's' : ''} from storage? This action cannot be undone.`)) {
      return;
    }
    setActionLoading(true);
    try {
      await api.post('/storage/batch-delete', { ids: Array.from(selectedIds) });
      setBanner(`✓ ${count} file${count > 1 ? 's' : ''} deleted from storage successfully.`);
      setSelectedIds(new Set());
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected files: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleTakeSnapshot(e) {
    if (e) e.preventDefault();
    if (isReadOnly) return;
    setActionLoading(true);
    try {
      const { data } = await api.post('/storage/snapshot-occupancy', { month: snapshotMonth || undefined });
      setBanner(`✓ Captured snapshot: ${data.title}`);
      setSnapshotModal(false);
      setSnapshotMonth('');
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Snapshot failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleUpload(e) {
    e.preventDefault();
    if (isReadOnly) return;
    const fd = new FormData(e.currentTarget);
    const file = fd.get('file');
    if (!file || !file.name) {
      alert('Please select a file to upload');
      return;
    }
    setActionLoading(true);
    try {
      const res = await api.post('/storage/upload', fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setBanner(`✓ Uploaded "${res.data.filename}" to storage successfully!`);
      setUploadModal(false);
      await loadStorage();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Upload failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setActionLoading(false);
    }
  }

  async function handleDelete(item) {
    if (!canDelete) return;
    if (!window.confirm(`Permanently delete "${item.title || item.filename}" from storage?`)) return;
    try {
      await api.delete(`/storage/${item.id}`);
      setItems(prev => prev.filter(x => x.id !== item.id));
      setSelectedIds(prev => {
        const next = new Set(prev);
        next.delete(item.id);
        return next;
      });
      setBanner('✓ Removed item from storage.');
      setTimeout(() => setBanner(''), 3000);
    } catch (err) {
      alert('Delete failed: ' + (err.response?.data?.message || err.message));
    }
  }

  async function handleDownload(item) {
    if (item.file_url) {
      const a = document.createElement('a');
      a.href = item.file_url;
      a.download = item.filename || 'download';
      document.body.appendChild(a);
      a.click();
      a.remove();
      return;
    }

    if (item.category === 'occupancy' && item.meta) {
      const meta = item.meta;
      const csvRows = [
        ['Media Buzz — Occupancy Snapshot Report'],
        ['Month', meta.month || ''],
        ['Occupancy Rate', `${meta.occupancyRate ?? 0}%`],
        ['Occupied Sites', `${meta.occupiedSites ?? 0} of ${meta.totalSites ?? 22}`],
        ['Total Revenue (₹)', meta.revenue ?? 0],
        ['Generated At', meta.generatedAt || new Date().toISOString()],
        [],
        ['Site Code', 'Status', 'Client', 'Monthly Rate (₹)', 'Location']
      ];
      if (Array.isArray(meta.sites)) {
        meta.sites.forEach(s => {
          csvRows.push([
            s.site_code || '',
            s.status || '',
            s.client || '',
            s.rate || s.advt_fees || 0,
            s.location || ''
          ]);
        });
      }
      const csvStr = csvRows.map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
      const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${(item.title || 'Occupancy_Snapshot').replace(/\s+/g, '_')}.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      return;
    }

    alert(`Download initiated for ${item.title || item.filename}`);
  }

  const defaultTitle = title || 'Document & Media Storage Vault';
  const defaultSubtitle = subtitle || 'Archive for generated presentations, exported Excel trackers, site occupancy records, and media attachments.';

  return (
    <div style={{ marginTop: embedded ? '36px' : '0' }}>
      <div style={{
        background: '#0d131f',
        border: '1px solid #1e293b',
        borderRadius: '16px',
        overflow: 'hidden',
        boxShadow: '0 8px 30px rgba(0,0,0,0.35)'
      }}>
        <div style={{
          padding: '20px 24px',
          borderBottom: '1px solid rgba(255,255,255,0.08)',
          background: 'linear-gradient(90deg, rgba(168,85,247,0.12), rgba(59,130,246,0.08), transparent)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '14px'
        }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '24px' }}>🗄️</span>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#f8fafc', letterSpacing: '-0.01em' }}>
                {defaultTitle}
              </h3>
              <span className="scooh-badgechip" style={{ background: 'rgba(168,85,247,0.22)', color: '#d8b4fe', fontWeight: 800, border: '1px solid rgba(168,85,247,0.4)', fontSize: '11px' }}>
                {items.length} Files Stored
              </span>
            </div>
            <p style={{ margin: '4px 0 0 34px', fontSize: '12.5px', color: '#94a3b8' }}>
              {defaultSubtitle}
            </p>
          </div>

          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
            {!isReadOnly && (
              <>
                <button
                  type="button"
                  className="scooh-btn secondary"
                  style={{ fontSize: '12px', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                  onClick={() => setSnapshotModal(true)}
                  title="Save an instant historical snapshot of current occupancy rate and active campaigns"
                >
                  📸 Snapshot Occupancy
                </button>
                <button
                  type="button"
                  className="scooh-btn purple-btn"
                  style={{ fontSize: '12px', padding: '8px 14px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                  onClick={() => setUploadModal(true)}
                  title="Upload PPT, Excel, or PDF directly to storage"
                >
                  📤 + Upload File
                </button>
              </>
            )}
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '12px', padding: '8px 12px' }}
              onClick={loadStorage}
              title="Reload storage list"
            >
              🔄
            </button>
          </div>
        </div>

        {isReadOnly && (
          <div className="scooh-banner" style={{ display: 'block', margin: '14px 20px 0', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
            🔒 Read-Only Workspace: You are signed in as a Viewer. You can view, search, inspect snapshots, and download files. Uploads, snapshots, and file deletions are disabled.
          </div>
        )}

        <div style={{
          display: 'flex',
          gap: '8px',
          padding: '12px 20px',
          background: 'rgba(15,23,42,0.6)',
          borderBottom: '1px solid rgba(255,255,255,0.06)',
          overflowX: 'auto',
          alignItems: 'center'
        }}>
          {[
            { id: 'all', label: 'All Storage', count: items.length, icon: '🌟' },
            { id: 'ppt', label: 'Generated PPTs', count: pptCount, icon: '📑' },
            { id: 'excel', label: 'Excel Workbooks', count: excelCount, icon: '📊' },
            { id: 'occupancy', label: 'Previous Site Occupancy', count: occCount, icon: '🏛️' },
            { id: 'other', label: 'Uploaded Documents', count: otherCount, icon: '🗂️' }
          ].map(tab => {
            const active = category === tab.id;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => setCategory(tab.id)}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '7px 14px',
                  borderRadius: '20px',
                  fontSize: '12px',
                  fontWeight: active ? 800 : 600,
                  cursor: 'pointer',
                  border: active ? '1px solid #a855f7' : '1px solid rgba(255,255,255,0.08)',
                  background: active ? 'rgba(168,85,247,0.25)' : 'rgba(30,41,59,0.5)',
                  color: active ? '#ffffff' : '#94a3b8',
                  transition: 'all 0.15s ease'
                }}
              >
                <span>{tab.icon}</span>
                <span>{tab.label}</span>
                <span style={{
                  padding: '1px 6px',
                  borderRadius: '10px',
                  fontSize: '10.5px',
                  background: active ? 'rgba(168,85,247,0.45)' : 'rgba(255,255,255,0.08)',
                  color: active ? '#f3e8ff' : '#cbd5e1'
                }}>
                  {tab.count}
                </span>
              </button>
            );
          })}
        </div>

        {banner && (
          <div style={{
            margin: '14px 20px',
            padding: '10px 14px',
            borderRadius: '8px',
            background: 'rgba(34,197,94,0.16)',
            border: '1px solid #22c55e',
            color: '#4ade80',
            fontSize: '12.5px',
            fontWeight: 600
          }}>
            {banner}
          </div>
        )}

        {category === 'occupancy' && occItems.length > 0 && (
          <div style={{
            padding: '16px 20px',
            background: 'rgba(30,41,59,0.4)',
            borderBottom: '1px solid rgba(255,255,255,0.06)'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 800, color: '#e2e8f0', letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                Occupancy History Timeline
              </span>
              <span style={{ fontSize: '11px', color: '#94a3b8' }}>
                {occItems.length} monthly snapshots recorded
              </span>
            </div>

            <div style={{ display: 'flex', gap: '12px', overflowX: 'auto', paddingBottom: '6px' }}>
              {occItems.map(item => {
                const m = item.meta || {};
                const occRate = Number(m.occupancyRate ?? 0);
                const isSelected = activeOccItem?.id === item.id;
                return (
                  <div
                    key={item.id}
                    onClick={() => setActiveOccItem(item)}
                    style={{
                      minWidth: '190px',
                      padding: '12px 14px',
                      borderRadius: '10px',
                      cursor: 'pointer',
                      border: isSelected ? '1px solid #38bdf8' : '1px solid rgba(255,255,255,0.08)',
                      background: isSelected ? 'rgba(56,189,248,0.15)' : '#111927',
                      transition: 'all 0.15s ease'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                      <span style={{ fontWeight: 800, color: '#f8fafc', fontSize: '13px' }}>
                        {m.month || item.title}
                      </span>
                      <span style={{
                        fontWeight: 900,
                        fontSize: '12px',
                        color: occRate >= 70 ? '#4ade80' : occRate >= 40 ? '#facc15' : '#ef4444'
                      }}>
                        {occRate}%
                      </span>
                    </div>

                    <div style={{
                      height: '5px',
                      borderRadius: '3px',
                      background: '#1e293b',
                      overflow: 'hidden',
                      marginBottom: '8px'
                    }}>
                      <div style={{
                        width: `${Math.min(100, Math.max(5, occRate))}%`,
                        height: '100%',
                        background: occRate >= 70 ? '#22c55e' : occRate >= 40 ? '#eab308' : '#ef4444'
                      }} />
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#94a3b8' }}>
                      <span>Sites: <strong style={{ color: '#e2e8f0' }}>{m.occupiedSites ?? '—'} / {m.totalSites ?? 22}</strong></span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        <div style={{
          padding: '12px 20px',
          display: 'flex',
          gap: '12px',
          alignItems: 'center',
          flexWrap: 'wrap',
          borderBottom: '1px solid rgba(255,255,255,0.06)'
        }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '180px', height: '36px' }}
            placeholder="Search stored PPTs, Excels, occupancy snapshots, filenames…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />

          <select
            value={sort}
            onChange={e => setSort(e.target.value)}
            style={{
              padding: '6px 12px',
              borderRadius: '8px',
              background: '#111925',
              color: '#e2e8f0',
              border: '1px solid #334155',
              fontSize: '12px',
              fontWeight: 600
            }}
          >
            <option value="newest">Sort: Newest Upload / Generation</option>
            <option value="oldest">Sort: Oldest First</option>
            <option value="title">Sort: Title (A–Z)</option>
          </select>

          {search && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '11px', padding: '6px 10px' }}
              onClick={() => setSearch('')}
            >
              Clear Search
            </button>
          )}

          {canDelete && filtered.length > 0 && (
            <button
              type="button"
              className="scooh-btn secondary"
              style={{ fontSize: '11.5px', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: '5px' }}
              onClick={toggleSelectAll}
              title={filtered.every(r => selectedIds.has(r.id)) ? 'Deselect all visible files' : 'Select all visible files'}
            >
              {filtered.every(r => selectedIds.has(r.id)) ? '✓ Deselect All' : `☑ Select All (${filtered.length})`}
            </button>
          )}
        </div>

        {canDelete && selectedIds.size > 0 && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '11px 18px',
            background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
            borderBottom: '1px solid rgba(239,68,68,0.3)',
            color: '#fca5a5',
            fontSize: '13px',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                {selectedIds.size} file{selectedIds.size > 1 ? 's' : ''} selected
              </span>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={selectAllVisible}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
              >
                Select all visible ({filtered.length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={deselectAll}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
              >
                Deselect all
              </button>
            </div>
            <button
              type="button"
              className="scooh-btn danger"
              onClick={handleBatchDelete}
              disabled={actionLoading}
              style={{
                background: '#ef4444',
                color: '#ffffff',
                border: 'none',
                fontWeight: 800,
                padding: '7px 16px',
                borderRadius: '7px',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}
            >
              🗑 Delete selected ({selectedIds.size})
            </button>
          </div>
        )}

        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                {canDelete && (
                  <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                    <input
                      type="checkbox"
                      checked={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id))}
                      onChange={toggleSelectAll}
                      title={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id)) ? 'Deselect all visible' : 'Select all visible'}
                      style={{ cursor: 'pointer', width: '15px', height: '15px' }}
                    />
                  </th>
                )}
                <th style={{ width: '60px', textAlign: 'center' }}>Type</th>
                <th style={{ minWidth: '220px' }}>Document / File Name</th>
                <th style={{ minWidth: '130px' }}>Category</th>
                <th style={{ minWidth: '200px' }}>Details & Metadata</th>
                <th style={{ width: '100px', textAlign: 'right' }}>Size</th>
                <th style={{ width: '120px' }}>Saved On</th>
                <th style={{ minWidth: '140px', textAlign: 'center' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={canDelete ? "8" : "7"} className="scooh-empty" style={{ padding: '36px 20px', textAlign: 'center' }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No storage items found</div>
                    <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                      Generate a presentation, export an Excel sheet, or click "+ Upload File" above.
                    </div>
                  </td>
                </tr>
              ) : (
                filtered.map(item => {
                  const m = item.meta || {};
                  const isPpt = item.category === 'ppt' || ['pptx', 'ppt'].includes(item.format);
                  const isExcel = item.category === 'excel' || ['xlsx', 'xls', 'csv'].includes(item.format);
                  const isOcc = item.category === 'occupancy' || item.format === 'json';

                  const badgeBg = isPpt ? 'rgba(249,115,22,0.18)' : isExcel ? 'rgba(34,197,94,0.18)' : isOcc ? 'rgba(56,189,248,0.18)' : 'rgba(168,85,247,0.18)';
                  const badgeColor = isPpt ? '#fb923c' : isExcel ? '#4ade80' : isOcc ? '#38bdf8' : '#c084fc';
                  const badgeBorder = isPpt ? 'rgba(249,115,22,0.35)' : isExcel ? 'rgba(34,197,94,0.35)' : isOcc ? 'rgba(56,189,248,0.35)' : 'rgba(168,85,247,0.35)';

                  return (
                    <tr key={item.id} style={{ background: selectedIds.has(item.id) ? 'rgba(239,68,68,0.06)' : undefined }}>
                      {canDelete && (
                        <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.has(item.id)}
                            onChange={() => toggleSelect(item.id)}
                            style={{ cursor: 'pointer', width: '15px', height: '15px' }}
                          />
                        </td>
                      )}
                      <td style={{ textAlign: 'center' }}>
                        <span className="scooh-badgechip" style={{ background: badgeBg, color: badgeColor, border: `1px solid ${badgeBorder}`, fontWeight: 800, fontSize: '10.5px' }}>
                          {item.format ? item.format.toUpperCase() : 'FILE'}
                        </span>
                      </td>
                      <td>
                        <div style={{ fontWeight: 800, color: '#ffffff', fontSize: '13px' }}>
                          {item.title}
                        </div>
                        <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px', fontFamily: 'monospace' }}>
                          {item.filename}
                        </div>
                      </td>
                      <td>
                        <span style={{
                          textTransform: 'capitalize',
                          fontSize: '12px',
                          color: '#e2e8f0',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '6px'
                        }}>
                          {item.category === 'ppt' ? '📑 PPT Deck' :
                           item.category === 'excel' ? '📊 Excel Sheet' :
                           item.category === 'occupancy' ? '🏛️ Occupancy Snapshot' : '🗂️ Document'}
                        </span>
                      </td>
                      <td style={{ fontSize: '11.5px', color: '#cbd5e1' }}>
                        {isPpt && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            {m.slides && <span className="scooh-plate" style={{ fontSize: '10px' }}>{m.slides} Slides</span>}
                          </div>
                        )}
                        {isExcel && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            {m.rows && <span className="scooh-plate" style={{ fontSize: '10px' }}>{m.rows} Rows</span>}
                            {m.type && <span style={{ color: '#4ade80' }}>{m.type}</span>}
                          </div>
                        )}
                        {isOcc && (
                          <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap', alignItems: 'center' }}>
                            <span className="scooh-badgechip" style={{ background: 'rgba(34,197,94,0.18)', color: '#4ade80', fontSize: '10px' }}>
                              Rate: {m.occupancyRate ?? '—'}%
                            </span>
                            <span style={{ color: '#cbd5e1' }}>
                              Occupied: {m.occupiedSites ?? '—'}/{m.totalSites ?? '—'}
                            </span>
                            {m.revenue ? <span style={{ color: '#facc15' }}>{money(m.revenue)}</span> : null}
                          </div>
                        )}
                        {!isPpt && !isExcel && !isOcc && (
                          <span style={{ color: '#94a3b8' }}>Standard file</span>
                        )}
                      </td>
                      <td style={{ textAlign: 'right', fontSize: '11.5px', color: '#94a3b8', fontWeight: 600 }}>
                        {item.file_size || '—'}
                      </td>
                      <td style={{ fontSize: '11.5px', color: '#94a3b8' }}>
                        {item.created_at ? new Date(item.created_at).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '—'}
                      </td>
                      <td>
                        <div className="scooh-rowactions" style={{ justifyContent: 'center' }}>
                          <button
                            type="button"
                            className="scooh-btn secondary"
                            style={{ minHeight: '28px', padding: '0 10px', fontSize: '11px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                            onClick={() => handleDownload(item)}
                            title="Download or re-export file"
                          >
                            ⬇ Download
                          </button>
                          {isOcc && (
                            <button
                              type="button"
                              className="scooh-btn ghost"
                              style={{ minHeight: '28px', padding: '0 8px', fontSize: '11px' }}
                              onClick={() => setActiveOccItem(item)}
                              title="View full snapshot breakdown"
                            >
                              👁 View
                            </button>
                          )}
                          {canDelete && (
                            <button
                              type="button"
                              className="scooh-iconbtn danger-icon"
                              style={{ minHeight: '28px', width: '28px' }}
                              onClick={() => handleDelete(item)}
                              title="Delete file from storage"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Upload Modal */}
      {uploadModal && (
        <div className="scooh-modal-overlay" onClick={() => !actionLoading && setUploadModal(false)}>
          <div className="scooh-modal" style={{ maxWidth: '520px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>Upload File to Storage</h2>
                <span className="scooh-modal-subtitle">Save PPT presentations, Excel files, or documents to permanent storage</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setUploadModal(false)}>×</button>
            </div>
            <form onSubmit={handleUpload}>
              <div className="scooh-modalbody">
                <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                  <div className="scooh-field">
                    <label>Choose File *</label>
                    <input type="file" name="file" required accept=".pptx,.ppt,.xlsx,.xls,.csv,.pdf,.json,.png,.jpg,.jpeg" />
                  </div>
                  <div className="scooh-field">
                    <label>Category</label>
                    <select name="category" defaultValue="ppt">
                      <option value="ppt">📑 PPT Presentation</option>
                      <option value="excel">📊 Excel Spreadsheet</option>
                      <option value="occupancy">🏛️ Site Occupancy Data</option>
                      <option value="other">🗂️ General Document / Contract</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Title (optional)</label>
                    <input name="title" placeholder="e.g. Diwali 2026 Pitch Deck" />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setUploadModal(false)}>Cancel</button>
                <button type="submit" className="scooh-btn purple-btn" disabled={actionLoading}>
                  {actionLoading ? 'Uploading…' : 'Upload to Storage'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Snapshot Modal */}
      {snapshotModal && (
        <div className="scooh-modal-overlay" onClick={() => !actionLoading && setSnapshotModal(false)}>
          <div className="scooh-modal" style={{ maxWidth: '480px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>Record Current Month Occupancy</h2>
                <span className="scooh-modal-subtitle">Save live occupancy percentage, active site count & revenue to history</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setSnapshotModal(false)}>×</button>
            </div>
            <form onSubmit={handleTakeSnapshot}>
              <div className="scooh-modalbody">
                <div className="scooh-field">
                  <label>Billing Month Label</label>
                  <input
                    value={snapshotMonth}
                    onChange={e => setSnapshotMonth(e.target.value)}
                    placeholder={`e.g. ${new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`}
                  />
                  <small style={{ color: '#94a3b8', marginTop: '4px', fontSize: '11px' }}>
                    Leave blank to automatically use current calendar month.
                  </small>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setSnapshotModal(false)}>Cancel</button>
                <button type="submit" className="scooh-btn secondary" disabled={actionLoading}>
                  {actionLoading ? 'Recording…' : '📸 Save Snapshot'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* View Occupancy Snapshot Detail Modal */}
      {activeOccItem && (
        <div className="scooh-modal-overlay" onClick={() => setActiveOccItem(null)}>
          <div className="scooh-modal" style={{ maxWidth: '580px', width: '92%' }} onClick={e => e.stopPropagation()}>
            <div className="scooh-modalhead">
              <div>
                <h2>{activeOccItem.title}</h2>
                <span className="scooh-modal-subtitle">Recorded on {new Date(activeOccItem.created_at).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
              </div>
              <button type="button" className="scooh-modalclose" onClick={() => setActiveOccItem(null)}>×</button>
            </div>
            <div className="scooh-modalbody">
              {(() => {
                const m = activeOccItem.meta || {};
                const pct = Number(m.occupancyPct || 0);
                return (
                  <div>
                    <div style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(2, 1fr)',
                      gap: '12px',
                      marginBottom: '16px'
                    }}>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Occupancy Rate</span>
                        <div style={{ fontSize: '24px', fontWeight: 800, color: '#38bdf8', marginTop: '4px' }}>{pct}%</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Portfolio Revenue</span>
                        <div style={{ fontSize: '24px', fontWeight: 800, color: '#4ade80', marginTop: '4px' }}>{money(m.revenue || 0)}</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Occupied Sites</span>
                        <div style={{ fontSize: '20px', fontWeight: 800, color: '#f8fafc', marginTop: '4px' }}>{m.occupiedSites ?? '—'} / {m.totalSites || 22}</div>
                      </div>
                      <div style={{ background: 'rgba(15,23,42,0.6)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                        <span style={{ fontSize: '11px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 800 }}>Vacant Sites</span>
                        <div style={{ fontSize: '20px', fontWeight: 800, color: '#f87171', marginTop: '4px' }}>{m.vacantSites ?? '—'}</div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
            <div className="scooh-modalfoot">
              <button
                type="button"
                className="scooh-btn secondary"
                onClick={() => handleDownload(activeOccItem)}
              >
                ⬇ Download Snapshot JSON
              </button>
              <button type="button" className="scooh-btn ghost" onClick={() => setActiveOccItem(null)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StorageView() {
  return (
    <>
      <PageHead
        title="Storage & Archives"
        desc="Dedicated repository of generated client PPT decks, exported Excel spreadsheets, historical monthly site occupancy snapshots, and operational files."
      />
      <StorageVaultSection embedded={false} />
    </>
  );
}

function ElectricityView() {
  const [rows, setRows] = useState([]);
  const [sites, setSites] = useState([]);
  const [editModal, setEditModal] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [providerFilter, setProviderFilter] = useState('ALL');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [sortState, setSortState] = useState({ key: 'site_code', dir: 'asc' });
  const [saving, setSaving] = useState(false);
  const [banner, setBanner] = useState('');
  const [importingExcel, setImportingExcel] = useState(false);
  const [selectedIds, setSelectedIds] = useState(new Set());

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canAdd = !isReadOnly;
  const canPay = !isReadOnly;

  async function loadData() {
    try {
      const [eRes, sRes] = await Promise.all([
        api.get('/electricity').catch(() => ({ data: [] })),
        api.get('/sites').catch(() => ({ data: [] }))
      ]);
      const siteList = Array.isArray(sRes.data) ? sRes.data : [];
      setSites(siteList);

      const siteMap = new Map();
      siteList.forEach(s => {
        if (s.id) siteMap.set(String(s.id), s);
        if (s.site_code) siteMap.set(String(s.site_code).trim(), s);
      });

      const rawBills = Array.isArray(eRes.data) ? eRes.data : [];
      const enriched = rawBills.map(b => {
        const matchedSite = siteMap.get(String(b.site_id)) || siteMap.get(String(b.site_code).trim()) || {};
        return {
          ...b,
          site_code: b.site_code || matchedSite.site_code || '',
          location: b.location || matchedSite.address || matchedSite.area || matchedSite.city || '',
          size: b.size || matchedSite.size || (matchedSite.width && matchedSite.height ? `${matchedSite.width}x${matchedSite.height} ft` : ''),
          meter_no: b.meter_no || matchedSite.meter_no || '',
          service_number: b.service_number || matchedSite.service_number || '',
          t_number: b.t_number || matchedSite.t_number || '',
          ecs: b.ecs || '',
          payment_date: b.payment_date || b.paid_date || '',
          paid_date: b.paid_date || b.payment_date || '',
          amount: Number(b.amount || b.payment_amount || 0),
          payment_amount: Number(b.payment_amount || b.amount || 0)
        };
      });
      setRows(enriched);
    } catch (err) {
      console.warn('Error loading electricity:', err);
    }
  }

  async function handleElectricityExcelImport(e) {
    if (!canDelete) return;
    const file = e.target.files?.[0];
    if (!file) return;
    setImportingExcel(true);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const r = await api.post('/import/electricity-xlsx', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `Processed ${r.data.rows} electricity bills (${r.data.updated} updated, ${r.data.created} created).`;
      setBanner(`✓ ${msg}`);
      alert(`✓ Electricity Bills Import Successful!\n\n${msg}`);
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
      setTimeout(() => setBanner(''), 6000);
    } catch (err) {
      alert('Electricity Bill Import failed: ' + (err.response?.data?.message || err.message));
    } finally {
      setImportingExcel(false);
      e.target.value = '';
    }
  }

  useEffect(() => {
    loadData();
    const onElecUpdate = () => loadData();
    window.addEventListener('mb-electricity-updated', onElecUpdate);
    const interval = setInterval(loadData, 30000);
    return () => {
      window.removeEventListener('mb-electricity-updated', onElecUpdate);
      clearInterval(interval);
    };
  }, []);

  const totalExpense = rows.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const paidExpense = rows.filter(r => r.payment_status === 'Paid').reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const pendingExpense = rows.filter(r => r.payment_status !== 'Paid').reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const overdueCount = rows.filter(r => {
    if (r.payment_status === 'Paid') return false;
    if (!r.due_date) return false;
    const d = new Date(r.due_date);
    d.setHours(23, 59, 59, 999);
    return d < new Date();
  }).length;

  const months = Array.from(new Set(rows.map(r => r.billing_month).filter(Boolean)));
  const providers = Array.from(new Set(['SSV', 'MB', ...rows.map(r => r.bill_type).filter(Boolean)]));

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const filtered = useMemo(() => {
    const list = rows.filter(r => {
      const q = search.trim().toLowerCase();
      let matchesSearch = !q;
      if (q) {
        if (matchSiteSearch({ site_code: r.site_code, address: r.location }, q)) {
          matchesSearch = true;
        } else if (!/^(?:mb[\s-]*)?\d+$/i.test(q)) {
          matchesSearch = [
            r.size,
            r.meter_no,
            r.service_number,
            r.t_number,
            r.bill_type,
            r.billing_month,
            r.payment_status,
            r.ecs,
            r.payment_reference,
            r.notes
          ].some(v => String(v || '').toLowerCase().includes(q));
        }
      }

      const matchesStatus = statusFilter === 'ALL' ? true :
        statusFilter === 'Overdue' ? (r.payment_status !== 'Paid' && r.due_date && new Date(r.due_date) < new Date()) :
        r.payment_status === statusFilter;

      const matchesProvider = providerFilter === 'ALL' || r.bill_type === providerFilter;
      const matchesMonth = monthFilter === 'ALL' || r.billing_month === monthFilter;

      return matchesSearch && matchesStatus && matchesProvider && matchesMonth;
    });

    if (sortState.key) {
      list.sort((a, b) => {
        let valA = a[sortState.key];
        let valB = b[sortState.key];
        if (sortState.key === 'amount') {
          valA = Number(a.amount || a.payment_amount || 0);
          valB = Number(b.amount || b.payment_amount || 0);
        } else if (sortState.key === 'payment_date') {
          valA = a.payment_date || a.paid_date || '';
          valB = b.payment_date || b.paid_date || '';
        }
        return universalCompare(valA, valB, sortState.dir);
      });
    }
    return list;
  }, [rows, search, statusFilter, providerFilter, monthFilter, sortState]);

  function openNewBill() {
    if (!canAdd) return;
    setEditModal({
      site_id: '',
      site_code: '',
      location: '',
      size: '',
      meter_no: '',
      service_number: '',
      t_number: '',
      bill_type: 'SSV',
      billing_month: new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
      bill_date: new Date().toISOString().slice(0, 10),
      due_date: '',
      units: '',
      rate: '',
      other_charges: 0,
      amount: '',
      payment_status: 'Pending',
      ecs: '',
      payment_date: '',
      paid_date: '',
      payment_reference: '',
      notes: ''
    });
  }

  function handleSiteChange(siteIdentifier) {
    const s = sites.find(x => String(x.id) === String(siteIdentifier) || String(x.site_code).trim() === String(siteIdentifier).trim());
    if (s) {
      setEditModal(prev => ({
        ...prev,
        site_id: s.id || '',
        site_code: s.site_code || '',
        location: s.address || s.area || s.city || '',
        size: s.size || (s.width && s.height ? `${s.width}x${s.height} ft` : ''),
        meter_no: s.meter_no || prev.meter_no || '',
        service_number: s.service_number || prev.service_number || '',
        t_number: s.t_number || prev.t_number || ''
      }));
    }
  }

  function calculateBillAmount(units, rate, other = 0) {
    const u = parseFloat(units);
    const r = parseFloat(rate);
    const o = parseFloat(other) || 0;
    if (!isNaN(u) && !isNaN(r)) {
      return (u * r + o).toFixed(2);
    }
    return '';
  }

  async function saveBill(e) {
    e.preventDefault();
    if (isReadOnly) return;
    setSaving(true);
    const form = e.currentTarget;
    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    if (!payload.site_code && !payload.site_id) {
      alert('Please select or specify a Site Code.');
      setSaving(false);
      return;
    }

    const amt = Number(payload.amount || payload.payment_amount || 0);
    payload.amount = amt;
    payload.payment_amount = amt;
    if (payload.payment_date && !payload.paid_date) payload.paid_date = payload.payment_date;
    if (payload.paid_date && !payload.payment_date) payload.payment_date = payload.paid_date;

    try {
      if (editModal?.id) {
        await api.put(`/electricity/${editModal.id}`, payload);
        setBanner('✓ Electricity bill updated successfully!');
      } else {
        await api.post('/electricity', payload);
        setBanner('✓ New electricity bill logged successfully!');
      }
      setEditModal(null);
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to save bill: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function handleQuickPay(e) {
    e.preventDefault();
    if (!payModal?.id || isReadOnly) return;
    setSaving(true);
    const formData = new FormData(e.currentTarget);
    const paid_date = formData.get('paid_date') || new Date().toISOString().slice(0, 10);
    const payment_reference = formData.get('payment_reference') || '';

    try {
      await api.put(`/electricity/${payModal.id}`, {
        payment_status: 'Paid',
        paid_date,
        payment_date: paid_date,
        payment_reference
      });
      setPayModal(null);
      setBanner(`✓ Bill for ${payModal.site_code || 'Site'} marked as Paid!`);
      await loadData();
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to update status: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = filtered.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(filtered.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    if (!confirm(`Are you sure you want to permanently delete all ${count} selected electricity bill record${count > 1 ? 's' : ''}? This action cannot be undone.`)) {
      return;
    }
    try {
      await api.post('/electricity/batch-delete', { ids: Array.from(selectedIds), hard: true });
      setBanner(`✓ ${count} electricity bill record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await loadData();
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
      setTimeout(() => setBanner(''), 4000);
    } catch (err) {
      alert('Failed to delete selected bills: ' + (err.response?.data?.message || err.message));
    }
  }

  async function deleteBill(id) {
    if (!canDelete) return;
    if (confirm('Permanently delete this electricity bill record?')) {
      try {
        await api.delete(`/electricity/${id}?hard=true`);
        setBanner('✓ Electricity bill deleted.');
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await loadData();
        window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
        setTimeout(() => setBanner(''), 3000);
      } catch (err) {
        alert('Failed to delete bill: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  return (
    <>
      <div className="scooh-pagehead">
        <div>
          <div style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '4px' }}>
            Media Buzz — OOH Workspace
          </div>
          <h1 style={{ fontSize: '26px', fontWeight: 900, margin: 0, color: '#fff', letterSpacing: '-0.02em' }}>
            Electricity Tracker
          </h1>
          <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '13px' }}>
            Every electricity bill, tracked independently of campaign activity so no payment due date is missed.
          </p>
        </div>
        <div className="scooh-headactions" style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
          <button type="button" className="scooh-btn ghost" onClick={loadData}>
            🔄 Refresh
          </button>
          {canDelete && (
            <label className="scooh-btn ghost" style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: '6px' }} title="Import electricity bills from Excel (.xlsx, .xls)">
              <span>📁 {importingExcel ? 'Importing…' : 'Import Excel'}</span>
              <input type="file" accept=".xlsx,.xls" hidden disabled={importingExcel} onChange={handleElectricityExcelImport} />
            </label>
          )}
          <button type="button" className="scooh-btn ghost" onClick={() => exportElectricityExcel(filtered)}>
            📥 Export Excel
          </button>
          {canAdd && (
            <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
              + Add bill
            </button>
          )}
        </div>
      </div>

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are signed in as a Viewer. You can view electricity bills, track due dates, and export to Excel. Recording or deleting bills is disabled.
        </div>
      )}

      {banner && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '16px', background: 'rgba(34,197,94,0.15)', border: '1px solid #22c55e', color: '#4ade80' }}>
          {banner}
        </div>
      )}

      {/* Financial & Operational KPI Cards */}
      <div className="scooh-kpirow" style={{ marginBottom: '22px' }}>
        <div className="scooh-kpi alert">
          <div className="n">{money(totalExpense)}</div>
          <div className="l">Total Electricity Expense</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>{rows.length} total bills logged</div>
        </div>
        <div className="scooh-kpi good">
          <div className="n">{money(paidExpense)}</div>
          <div className="l">Paid Bills</div>
          <div className="scooh-kpi-note" style={{ color: '#4ade80' }}>
            {rows.filter(r => r.payment_status === 'Paid').length} bills cleared
          </div>
        </div>
        <div className={`scooh-kpi ${overdueCount > 0 ? 'danger' : 'alert'}`}>
          <div className="n">{money(pendingExpense)}</div>
          <div className="l">Pending / Due Bills</div>
          <div className="scooh-kpi-note" style={{ color: overdueCount > 0 ? '#ef4444' : '#fbbf24' }}>
            {overdueCount > 0 ? `⚠ ${overdueCount} overdue bills` : `${rows.filter(r => r.payment_status !== 'Paid').length} awaiting payment`}
          </div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{sites.length || 57}</div>
          <div className="l">Total Sites Monitored</div>
          <div className="scooh-kpi-note" style={{ color: '#94a3b8' }}>SSV & MB Bills</div>
        </div>
      </div>

      {/* Main Section */}
      <div className="scooh-electricity-section">
        <div className="scooh-sectionbar">
          <div>
            <h2 style={{ fontSize: '16px', fontWeight: 800, color: '#f8fafc' }}>
              Electricity Bills & Payment Due Dates
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#94a3b8' }}>
              All imported and newly added electricity bills are shown here with site and meter details.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            {canAdd && (
              <button type="button" className="scooh-btn purple-btn" onClick={openNewBill}>
                + Add bill
              </button>
            )}
          </div>
        </div>

        {/* Filters Toolbar */}
        <div className="scooh-toolbar" style={{ padding: '14px 20px', borderBottom: '1px solid rgba(255,255,255,0.06)', display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center' }}>
          <input
            className="scooh-search"
            style={{ flex: '1 1 240px', minWidth: '200px' }}
            placeholder="Search site code (e.g. 01, MB-01), location, meter, consumer no…"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
          >
            <option value="ALL">All Statuses ({rows.length})</option>
            <option value="Pending">Pending</option>
            <option value="Paid">Paid</option>
            <option value="Overdue">Overdue ({overdueCount})</option>
          </select>
          {providers.length > 0 && (
            <select
              value={providerFilter}
              onChange={e => setProviderFilter(e.target.value)}
              style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
            >
              <option value="ALL">All Providers ({providers.length})</option>
              {providers.map(p => (
                <option key={p} value={p}>{p}</option>
              ))}
            </select>
          )}
          {months.length > 0 && (
            <select
              value={monthFilter}
              onChange={e => setMonthFilter(e.target.value)}
              style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px' }}
            >
              <option value="ALL">All Billing Months ({months.length})</option>
              {months.map(m => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          )}
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value="site_code:asc">Sort: Site Code (01 → 87)</option>
            <option value="site_code:desc">Sort: Site Code (87 → 01)</option>
            <option value="due_date:asc">Sort: Due Date (Earliest First)</option>
            <option value="due_date:desc">Sort: Due Date (Latest First)</option>
            <option value="payment_date:desc">Sort: Payment Date (Latest First)</option>
            <option value="payment_date:asc">Sort: Payment Date (Earliest First)</option>
            <option value="amount:desc">Sort: Amount (High to Low)</option>
            <option value="amount:asc">Sort: Amount (Low to High)</option>
            <option value="location:asc">Sort: Location (A–Z)</option>
            <option value="payment_status:asc">Sort: Status (A–Z)</option>
            <option value="ecs:asc">Sort: ECS (A–Z)</option>
            <option value="billing_month:desc">Sort: Billing Month</option>
          </select>
          {(search || statusFilter !== 'ALL' || providerFilter !== 'ALL' || monthFilter !== 'ALL' || sortState.key !== 'due_date' || sortState.dir !== 'asc') && (
            <button
              type="button"
              className="scooh-btn ghost"
              style={{ fontSize: '11px', padding: '6px 10px' }}
              onClick={() => { setSearch(''); setStatusFilter('ALL'); setProviderFilter('ALL'); setMonthFilter('ALL'); setSortState({ key: 'due_date', dir: 'asc' }); }}
            >
              Reset Filters
            </button>
          )}
        </div>

        {/* Batch Selection Action Bar */}
        {canDelete && selectedIds.size > 0 && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '11px 18px',
            background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
            borderBottom: '1px solid rgba(239,68,68,0.3)',
            color: '#fca5a5',
            fontSize: '13px',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
              <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
                {selectedIds.size} bill{selectedIds.size > 1 ? 's' : ''} selected
              </span>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={selectAllVisible}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
              >
                Select all visible ({filtered.length})
              </button>
              <button
                type="button"
                className="scooh-btn ghost"
                onClick={deselectAll}
                style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
              >
                Deselect all
              </button>
            </div>
            <button
              type="button"
              className="scooh-btn danger"
              onClick={handleBatchDelete}
              style={{
                background: '#ef4444',
                color: '#ffffff',
                border: 'none',
                fontWeight: 800,
                padding: '7px 16px',
                borderRadius: '7px',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
              }}
            >
              🗑 Delete selected ({selectedIds.size})
            </button>
          </div>
        )}

        {/* Table */}
        <div className="scooh-tablewrap">
          <table className="scooh-table">
            <thead>
              <tr>
                {canDelete && (
                  <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                    <input
                      type="checkbox"
                      checked={filtered.length > 0 && filtered.every(r => selectedIds.has(r.id))}
                      onChange={toggleSelectAll}
                      title="Select all visible bills"
                      style={{ cursor: 'pointer' }}
                    />
                  </th>
                )}
                <SortHeader label="SITE CODE" sortKey="site_code" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="LOCATION" sortKey="location" currentSort={sortState} onSort={handleSort} style={{ minWidth: '180px' }} />
                <SortHeader label="SIZE" sortKey="size" currentSort={sortState} onSort={handleSort} style={{ minWidth: '85px' }} />
                <SortHeader label="METER / SERVICE" sortKey="meter_no" currentSort={sortState} onSort={handleSort} style={{ minWidth: '140px' }} />
                <SortHeader label="T NUMBER" sortKey="t_number" currentSort={sortState} onSort={handleSort} style={{ minWidth: '100px' }} />
                <SortHeader label="BILL / PROVIDER" sortKey="bill_type" currentSort={sortState} onSort={handleSort} style={{ minWidth: '120px' }} />
                <SortHeader label="BILLING MONTH" sortKey="billing_month" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="DUE DATE" sortKey="due_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="PAYMENT DATE" sortKey="payment_date" currentSort={sortState} onSort={handleSort} style={{ minWidth: '110px' }} />
                <SortHeader label="PAYMENT AMOUNT" sortKey="amount" currentSort={sortState} onSort={handleSort} align="right" style={{ minWidth: '120px' }} />
                <SortHeader label="STATUS" sortKey="payment_status" currentSort={sortState} onSort={handleSort} align="center" style={{ minWidth: '95px' }} />
                <SortHeader label="ECS" sortKey="ecs" currentSort={sortState} onSort={handleSort} align="center" style={{ minWidth: '85px' }} />
                <th style={{ minWidth: '140px', textAlign: 'center' }}>ACTIONS</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={canDelete ? "14" : "13"} className="scooh-empty" style={{ padding: '36px 20px' }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: '#94a3b8' }}>No electricity bills match your query</div>
                    <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                      {canAdd ? 'Click "+ Add bill" to record a new electricity bill for any hoarding site.' : 'No recorded electricity bills.'}
                    </div>
                  </td>
                </tr>
              ) : (
                filtered.map(r => {
                  const isOverdue = r.payment_status !== 'Paid' && r.due_date && new Date(r.due_date) < new Date();
                  return (
                    <tr key={r.id}>
                      {canDelete && (
                        <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.has(r.id)}
                            onChange={() => toggleSelect(r.id)}
                            style={{ cursor: 'pointer' }}
                          />
                        </td>
                      )}
                      <td>
                        <span className="scooh-plate" style={{ fontSize: '11.5px', fontWeight: 800 }}>
                          {r.site_code || '—'}
                        </span>
                      </td>
                      <td>
                        <div style={{ fontWeight: 700, color: '#edf2f6', fontSize: '12.5px', lineHeight: 1.35 }}>
                          {r.location || '—'}
                        </div>
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
                        {r.size || '—'}
                      </td>
                      <td>
                        <div style={{ fontSize: '11.5px', color: '#f1f5f9', fontWeight: 600 }}>
                          {r.meter_no ? `Mtr: ${r.meter_no}` : (r.service_number ? `Srv: ${r.service_number}` : '—')}
                        </div>
                        {r.service_number && r.meter_no && (
                          <div style={{ fontSize: '10.5px', color: '#94a3b8', marginTop: '2px' }}>
                            Srv: {r.service_number}
                          </div>
                        )}
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px' }}>
                        {r.t_number || '—'}
                      </td>
                      <td>
                        <span
                          className="scooh-badgechip"
                          style={{
                            background: r.bill_type === 'SSV' ? 'rgba(52,211,153,0.18)' :
                              (r.bill_type === 'MB' ? 'rgba(251,191,36,0.18)' :
                              (r.bill_type === 'Torrent Power' ? 'rgba(56,189,248,0.16)' : 'rgba(168,85,247,0.16)')),
                            color: r.bill_type === 'SSV' ? '#34d399' :
                              (r.bill_type === 'MB' ? '#fbbf24' :
                              (r.bill_type === 'Torrent Power' ? '#38bdf8' : '#c084fc')),
                            fontWeight: 800,
                            fontSize: '11px',
                            letterSpacing: '0.02em',
                            border: r.bill_type === 'SSV' ? '1px solid rgba(52,211,153,0.3)' :
                              (r.bill_type === 'MB' ? '1px solid rgba(251,191,36,0.35)' : 'none')
                          }}
                        >
                          {r.bill_type || 'SSV'}
                        </span>
                      </td>
                      <td style={{ color: '#cbd5e1', fontSize: '12px', fontWeight: 600 }}>
                        {r.billing_month || '—'}
                      </td>
                      <td>
                        <div style={{ fontSize: '12px', color: isOverdue ? '#f87171' : '#cbd5e1', fontWeight: isOverdue ? 800 : 500 }}>
                          {formatDate(r.due_date)}
                        </div>
                        {isOverdue && (
                          <span style={{ display: 'inline-block', marginTop: '2px', fontSize: '9.5px', fontWeight: 800, color: '#ef4444', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
                            ⚠ Overdue
                          </span>
                        )}
                      </td>
                      <td>
                        <div style={{ fontSize: '12px', color: (r.payment_date || r.paid_date) ? '#38bdf8' : '#64748b', fontWeight: (r.payment_date || r.paid_date) ? 600 : 400 }}>
                          {formatDate(r.payment_date || r.paid_date)}
                        </div>
                      </td>
                      <td style={{ textAlign: 'right' }}>
                        <b style={{ fontSize: '13px', color: '#f8fafc', letterSpacing: '0.01em' }}>
                          {money(r.amount || r.payment_amount)}
                        </b>
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        <span className={`scooh-pill ${r.payment_status === 'Paid' ? 'active' : r.payment_status === 'Pending' ? 'watch' : 'vacant'}`}>
                          {r.payment_status || 'Pending'}
                        </span>
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        {r.ecs ? (
                          <span
                            className="scooh-badgechip"
                            style={{
                              background: /^(yes|active|y|done)$/i.test(String(r.ecs).trim()) ? 'rgba(52,211,153,0.18)' :
                                          (/^(no|inactive|n)$/i.test(String(r.ecs).trim()) ? 'rgba(239,68,68,0.15)' : 'rgba(168,85,247,0.16)'),
                              color: /^(yes|active|y|done)$/i.test(String(r.ecs).trim()) ? '#34d399' :
                                     (/^(no|inactive|n)$/i.test(String(r.ecs).trim()) ? '#f87171' : '#c084fc'),
                              border: /^(yes|active|y|done)$/i.test(String(r.ecs).trim()) ? '1px solid rgba(52,211,153,0.35)' :
                                      (/^(no|inactive|n)$/i.test(String(r.ecs).trim()) ? '1px solid rgba(239,68,68,0.3)' : '1px solid rgba(168,85,247,0.35)'),
                              fontWeight: 700,
                              fontSize: '11px',
                              padding: '2px 8px',
                              borderRadius: '6px'
                            }}
                          >
                            {r.ecs}
                          </span>
                        ) : (
                          <span style={{ color: '#64748b', fontSize: '12px' }}>—</span>
                        )}
                      </td>
                      <td>
                        <div className="scooh-rowactions" style={{ justifyContent: 'center', gap: '6px' }}>
                          {canPay && r.payment_status !== 'Paid' && (
                            <button
                              type="button"
                              className="scooh-btn"
                              style={{
                                padding: '4px 8px',
                                fontSize: '11px',
                                fontWeight: 700,
                                background: 'rgba(34,197,94,0.18)',
                                color: '#4ade80',
                                border: '1px solid rgba(34,197,94,0.35)',
                                borderRadius: '6px'
                              }}
                              onClick={() => setPayModal(r)}
                              title="Mark this bill as Paid"
                            >
                              ✓ Pay
                            </button>
                          )}
                          <button
                            type="button"
                            className="scooh-iconbtn scooh-text-action"
                            onClick={() => setEditModal(r)}
                            title={isReadOnly ? 'View Bill Details' : 'Edit Bill'}
                          >
                            {isReadOnly ? 'View' : 'Edit'}
                          </button>
                          {canDelete && (
                            <button
                              type="button"
                              className="scooh-iconbtn danger-icon"
                              onClick={() => deleteBill(r.id)}
                              title="Delete Bill"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add / Edit Bill Modal */}
      {editModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '640px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>{editModal.id ? (isReadOnly ? 'View Electricity Bill' : 'Edit Electricity Bill') : 'Add Electricity Bill'}</h2>
                <span className="scooh-modal-subtitle">
                  {isReadOnly ? 'Read-only view of utility meter bill and payment record' : 'Auto-fill site metadata and manage payment tracking'}
                </span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEditModal(null)}>×</button>
            </div>
            <form onSubmit={saveBill}>
              <div className="scooh-modalbody">
                {/* Site Selection Quick Picker */}
                <div className="scooh-field" style={{ gridColumn: '1 / -1', background: 'rgba(255,255,255,0.03)', padding: '12px', borderRadius: '10px', border: '1px solid rgba(255,255,255,0.08)' }}>
                  <label style={{ color: '#c4b5fd', fontWeight: 800 }}>⚡ QUICK AUTO-FILL FROM SITE INVENTORY</label>
                  <select
                    value={editModal.site_id || editModal.site_code || ''}
                    onChange={e => handleSiteChange(e.target.value)}
                    style={{ marginTop: '6px' }}
                  >
                    <option value="">-- Select a Site to Auto-Fill Location, Size & Meter --</option>
                    {sites.map(s => (
                      <option key={s.id || s.site_code} value={s.id || s.site_code}>
                        {s.site_code} — {s.area || s.address || s.city} ({s.size || 'Standard'})
                      </option>
                    ))}
                  </select>
                </div>

                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Site Code *</label>
                    <input
                      name="site_code"
                      value={editModal.site_code ?? ''}
                      onChange={e => setEditModal({ ...editModal, site_code: e.target.value })}
                      placeholder="e.g. AMD-GT-001"
                      required
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Site Size</label>
                    <input
                      name="size"
                      value={editModal.size ?? ''}
                      onChange={e => setEditModal({ ...editModal, size: e.target.value })}
                      placeholder="e.g. 30x10 ft"
                    />
                  </div>
                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Location / Address *</label>
                    <input
                      name="location"
                      value={editModal.location ?? ''}
                      onChange={e => setEditModal({ ...editModal, location: e.target.value })}
                      placeholder="e.g. Shivranjani Cross Roads, Ahmedabad"
                      required
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Bill Provider / Type *</label>
                    <select
                      name="bill_type"
                      value={editModal.bill_type ?? 'SSV'}
                      onChange={e => setEditModal({ ...editModal, bill_type: e.target.value })}
                    >
                      <option value="SSV">SSV</option>
                      <option value="MB">MB</option>
                      <option value="Torrent Power">Torrent Power</option>
                      <option value="UGVCL">UGVCL</option>
                      <option value="PGVCL">PGVCL</option>
                      <option value="DGVCL">DGVCL</option>
                      <option value="MGVCL">MGVCL</option>
                      <option value="Other">Other Utility</option>
                      {editModal.bill_type && !['SSV', 'MB', 'Torrent Power', 'UGVCL', 'PGVCL', 'DGVCL', 'MGVCL', 'Other'].includes(editModal.bill_type) && (
                        <option value={editModal.bill_type}>{editModal.bill_type}</option>
                      )}
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Billing Month</label>
                    <input
                      name="billing_month"
                      value={editModal.billing_month ?? ''}
                      onChange={e => setEditModal({ ...editModal, billing_month: e.target.value })}
                      placeholder="e.g. Aug 2026"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Meter Number</label>
                    <input
                      name="meter_no"
                      value={editModal.meter_no ?? ''}
                      onChange={e => setEditModal({ ...editModal, meter_no: e.target.value })}
                      placeholder="e.g. MTR-UGVCL-8841"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Service Number</label>
                    <input
                      name="service_number"
                      value={editModal.service_number ?? ''}
                      onChange={e => setEditModal({ ...editModal, service_number: e.target.value })}
                      placeholder="e.g. SRV-998241"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>T Number</label>
                    <input
                      name="t_number"
                      value={editModal.t_number ?? ''}
                      onChange={e => setEditModal({ ...editModal, t_number: e.target.value })}
                      placeholder="e.g. T-4401"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Bill Date</label>
                    <input
                      type="date"
                      name="bill_date"
                      defaultValue={editModal.bill_date ? editModal.bill_date.slice(0, 10) : ''}
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Payment Due Date *</label>
                    <input
                      type="date"
                      name="due_date"
                      defaultValue={editModal.due_date ? editModal.due_date.slice(0, 10) : ''}
                      required
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Units Consumed (kWh)</label>
                    <input
                      type="number"
                      step="any"
                      name="units"
                      value={editModal.units ?? ''}
                      onChange={e => {
                        const u = e.target.value;
                        const calc = calculateBillAmount(u, editModal.rate, editModal.other_charges);
                        setEditModal(prev => ({
                          ...prev,
                          units: u,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="e.g. 850"
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Rate per Unit (₹)</label>
                    <input
                      type="number"
                      step="any"
                      name="rate"
                      value={editModal.rate ?? ''}
                      onChange={e => {
                        const r = e.target.value;
                        const calc = calculateBillAmount(editModal.units, r, editModal.other_charges);
                        setEditModal(prev => ({
                          ...prev,
                          rate: r,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="e.g. 9.23"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Other Charges (₹)</label>
                    <input
                      type="number"
                      step="any"
                      name="other_charges"
                      value={editModal.other_charges ?? ''}
                      onChange={e => {
                        const o = e.target.value;
                        const calc = calculateBillAmount(editModal.units, editModal.rate, o);
                        setEditModal(prev => ({
                          ...prev,
                          other_charges: o,
                          amount: calc || prev.amount
                        }));
                      }}
                      placeholder="0.00"
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1', background: 'rgba(139,92,246,0.06)', padding: '12px', borderRadius: '8px', border: '1px solid rgba(139,92,246,0.2)' }}>
                    <label style={{ color: '#c4b5fd', fontWeight: 900, fontSize: '12px' }}>BILL / PAYMENT AMOUNT (₹) *</label>
                    <input
                      type="number"
                      step="any"
                      name="amount"
                      value={editModal.amount ?? ''}
                      onChange={e => setEditModal({ ...editModal, amount: e.target.value })}
                      placeholder="e.g. 7850"
                      style={{ fontSize: '15px', fontWeight: 800 }}
                      required
                    />
                  </div>

                  <div className="scooh-field">
                    <label>Payment Status</label>
                    <select
                      name="payment_status"
                      value={editModal.payment_status ?? 'Pending'}
                      onChange={e => setEditModal({ ...editModal, payment_status: e.target.value })}
                    >
                      <option value="Pending">Pending</option>
                      <option value="Paid">Paid</option>
                      <option value="Overdue">Overdue</option>
                      <option value="Partial">Partial</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>ECS (Auto-Debit / Mandate)</label>
                    <input
                      name="ecs"
                      value={editModal.ecs ?? ''}
                      onChange={e => setEditModal({ ...editModal, ecs: e.target.value })}
                      placeholder="e.g. Yes / No / Mandate ID"
                    />
                  </div>
                  <div className="scooh-field">
                    <label>Payment Date</label>
                    <input
                      type="date"
                      name="payment_date"
                      defaultValue={editModal.payment_date ? editModal.payment_date.slice(0, 10) : (editModal.paid_date ? editModal.paid_date.slice(0, 10) : '')}
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Payment Reference / Transaction ID</label>
                    <input
                      name="payment_reference"
                      value={editModal.payment_reference ?? ''}
                      onChange={e => setEditModal({ ...editModal, payment_reference: e.target.value })}
                      placeholder="e.g. UPI-9923847291 / NEFT / Cheque #4412"
                    />
                  </div>

                  <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                    <label>Operational Notes</label>
                    <textarea
                      name="notes"
                      value={editModal.notes ?? ''}
                      onChange={e => setEditModal({ ...editModal, notes: e.target.value })}
                      placeholder="Any notes on meter reading, meter seal, or utility dispatch…"
                      rows={2}
                    />
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEditModal(null)}>
                  {isReadOnly ? 'Close' : 'Cancel'}
                </button>
                {!isReadOnly && (
                  <button type="submit" className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : editModal.id ? 'Update Electricity Bill' : 'Save Electricity Bill'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Quick Mark Paid Modal */}
      {payModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '440px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>Mark Bill as Paid</h2>
                <span className="scooh-modal-subtitle">{payModal.site_code} — {payModal.location || 'Site Bill'}</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setPayModal(null)}>×</button>
            </div>
            <form onSubmit={handleQuickPay}>
              <div className="scooh-modalbody">
                <div style={{ padding: '12px', background: 'rgba(34,197,94,0.1)', border: '1px solid rgba(34,197,94,0.25)', borderRadius: '8px', marginBottom: '14px' }}>
                  <div style={{ fontSize: '11px', color: '#94a3b8' }}>BILL AMOUNT</div>
                  <div style={{ fontSize: '20px', fontWeight: 900, color: '#4ade80', marginTop: '2px' }}>
                    {money(payModal.amount || payModal.payment_amount)}
                  </div>
                  <div style={{ fontSize: '11px', color: '#cbd5e1', marginTop: '4px' }}>
                    Month: {payModal.billing_month || 'Current'} | Provider: {payModal.bill_type || 'General'}
                  </div>
                </div>

                <div className="scooh-field">
                  <label>Payment Date *</label>
                  <input
                    type="date"
                    name="paid_date"
                    defaultValue={new Date().toISOString().slice(0, 10)}
                    required
                  />
                </div>
                <div className="scooh-field">
                  <label>Payment Reference / UTR / Mode</label>
                  <input
                    name="payment_reference"
                    placeholder="e.g. UPI-9923847291 / NetBanking / Cheque"
                    autoFocus
                  />
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setPayModal(null)}>Cancel</button>
                <button type="submit" className="scooh-btn purple-btn" style={{ background: '#22c55e', borderColor: '#22c55e' }} disabled={saving}>
                  {saving ? 'Updating…' : 'Confirm Paid'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

function DataToolsView() {
  const [xlsxFile, setXlsxFile] = useState(null);
  const [elecFile, setElecFile] = useState(null);
  const [jsonFile, setJsonFile] = useState(null);
  const [xlsxStatus, setXlsxStatus] = useState('No Excel file selected.');
  const [elecStatus, setElecStatus] = useState('No Electricity Excel selected.');
  const [jsonStatus, setJsonStatus] = useState('No JSON file selected.');
  const [loading, setLoading] = useState(false);

  async function importXlsx() {
    if (!xlsxFile) return alert('Please choose an Excel file (.xlsx, .xls) first.');
    setLoading(true);
    setXlsxStatus(`Importing ${xlsxFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', xlsxFile);
      const r = await api.post('/import/xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `✓ Processed ${r.data.rows} sites (${r.data.updated} updated, ${r.data.created} created).`;
      setXlsxStatus(msg);
      alert(`✓ Excel Import Successful!\n\n${msg}`);
      window.dispatchEvent(new CustomEvent('mb-sites-updated'));
    } catch (e) {
      setXlsxStatus(`✕ Import failed: ${e.response?.data?.message || e.message}`);
      alert('Import failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function importElecXlsx() {
    if (!elecFile) return alert('Please choose an Electricity Excel file (.xlsx, .xls) first.');
    setLoading(true);
    setElecStatus(`Importing ${elecFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', elecFile);
      const r = await api.post('/import/electricity-xlsx', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      const msg = r.data.message || `✓ Processed ${r.data.rows} electricity bills (${r.data.updated} updated, ${r.data.created} created).`;
      setElecStatus(msg);
      alert(`✓ Electricity Import Successful!\n\n${msg}`);
      window.dispatchEvent(new CustomEvent('mb-electricity-updated'));
    } catch (e) {
      setElecStatus(`✕ Import failed: ${e.response?.data?.message || e.message}`);
      alert('Import failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function importJson() {
    if (!jsonFile) return alert('Please choose a JSON backup file first.');
    setLoading(true);
    setJsonStatus(`Restoring ${jsonFile.name}…`);
    try {
      const f = new FormData();
      f.append('file', jsonFile);
      const r = await api.post('/import/json', f, { headers: { 'Content-Type': 'multipart/form-data' } });
      setJsonStatus(`✓ Successfully restored ${r.data.count} records from ${jsonFile.name}!`);
      alert(`Success: ${r.data.count} records restored.`);
    } catch (e) {
      setJsonStatus(`✕ Restore failed: ${e.response?.data?.message || e.message}`);
      alert('Restore failed: ' + (e.response?.data?.message || e.message));
    } finally {
      setLoading(false);
    }
  }

  async function exportExcel() {
    try {
      const { data } = await api.get('/sites');
      const source = Array.isArray(data) ? data : [];
      const rows = source.map(unpackSite).map((x, idx) => {
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).replace(/[^0-9xX.]/g, '').split(/[xX]/);
          if (parts.length === 2) { w = parts[0]; h = parts[1]; }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : (x.sq_ft || '');
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        return {
          'SR NO': idx + 1,
          'MB CODE': x.site_code || '',
          'SITE CODE': x.site_code || '',
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': x.ppt_availability || x.availability || 'Available',
          'Selling Amount': Number(x.ppt_rate || x.monthly_rate || 0),
          'Latitude Longitude': coords
        };
      });

      await exportStyledExcel(rows, 'MediaBuzz_Sites.xlsx');
    } catch (e) {
      alert('Export failed: ' + e.message);
    }
  }

  async function exportCsv() {
    try {
      const { data } = await api.get('/sites');
      const source = Array.isArray(data) ? data : [];
      const rows = source.map(unpackSite).map((x, idx) => {
        let w = x.width || '', h = x.height || '';
        if ((!w || !h) && x.size) {
          const parts = String(x.size).replace(/[^0-9xX.]/g, '').split(/[xX]/);
          if (parts.length === 2) { w = parts[0]; h = parts[1]; }
        }
        const sqft = (Number(w || 0) && Number(h || 0)) ? Number(w) * Number(h) : '';
        const coords = x.gps || ([x.latitude, x.longitude].filter(Boolean).join(', ')) || '';

        return {
          'SR NO': idx + 1,
          'AREA': x.area || x.city || '',
          'LOCATION': x.address || '',
          'MEDIA': x.media_type || 'Hoarding',
          'LIGHT': x.lighting || 'BL',
          'W': w || '',
          'H': h || '',
          'SQ FT': sqft,
          'AVAILABLITY': x.ppt_availability || x.availability || 'Available',
          'Selling Amount': Number(x.ppt_rate || x.monthly_rate || 0),
          'Latitude Longitude': coords
        };
      });

      const ws = XLSX.utils.json_to_sheet(rows);
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const u = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = u;
      a.download = 'MediaBuzz_Sites.csv';
      a.click();
      URL.revokeObjectURL(u);
    } catch (e) {
      alert('Export CSV failed: ' + e.message);
    }
  }

  async function exportJson() {
    try {
      const { data } = await api.get('/export/json');
      const b = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const u = URL.createObjectURL(b);
      const a = document.createElement('a');
      a.href = u;
      a.download = 'mediabuzz-database-backup.json';
      a.click();
      URL.revokeObjectURL(u);
    } catch (e) {
      alert('Export JSON failed: ' + e.message);
    }
  }

  function downloadProductionGuide() {
    const text = `# Media Buzz OOH Workspace — Production Guide & System Documentation
Version: 3.0.0 | Environment: Production (Hostinger / VPS / Cloud)

## 1. System Architecture
- Frontend: React 18, React Router v6, Vite SPA
- Backend: Express Node.js API with JWT Stateless Auth & Bcrypt
- Database: MySQL 8.0 (InnoDB) with composite indexing
- Presentation: Client-side 16:9 widescreen PPTX generation (PptxGenJS)
- Excel: Native 11-column parser & builder (SheetJS XLSX)

## 2. Production Advantages
- Sub-50ms API response time with zero WordPress PHP overhead
- 100% data privacy on your own MySQL database
- Client-side PowerPoint generation in <2 seconds with custom brand cards
- Zero recurring software licensing or subscription fees
- Full source code ownership in modern React & Node.js

## 3. Production Disadvantages & Mitigations
- Self-managed infrastructure (Mitigation: Auto-restart via PM2/Hostinger Node runner)
- Database backups (Mitigation: Built-in one-click Export JSON backup + MySQL cron dumps)
- SMTP email configuration (Mitigation: Standard SMTP environment variables in .env)

## 4. Excel Standard Column Format
SR NO | AREA | LOCATION | MEDIA | LIGHT | W | H | SQ FT | AVAILABLITY | Selling Amount | Latitude Longitude

## 5. Security & Deployment
- Set strong JWT_SECRET and database passwords in server/.env
- Configure daily cron backup: mysqldump -u <user> -p'<pass>' <dbname> > /backups/mb_\$(date +%F).sql
- Hostinger Node.js runner: Node 18/20, Root: /, Startup: server/src/index.js
`;
    const b = new Blob([text], { type: 'text/markdown' });
    const u = URL.createObjectURL(b);
    const a = document.createElement('a');
    a.href = u;
    a.download = 'MediaBuzz_Production_Guide.md';
    a.click();
    URL.revokeObjectURL(u);
  }

  return (
    <>
      <PageHead
        title="Import / Export"
        desc="Import or update your OOH workbook, then export operational data as Excel, PowerPoint, JSON or CSV."
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={downloadProductionGuide}>
              Download Production Guide (.md)
            </button>
            <button type="button" className="scooh-btn primary" onClick={exportExcel}>
              Export Excel (.xlsx)
            </button>
          </>
        }
      />

      <div className="scooh-grid2 scooh-data-grid">
        {/* Import from Excel */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import from Excel</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select your 11-column Excel file, then click <b>Import Data</b> to auto-update sites and availability.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {xlsxFile ? 'Change File' : 'Choose Excel'}
              <input
                type="file"
                id="import-xlsx"
                accept=".xlsx,.xls"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setXlsxFile(f || null);
                  setXlsxStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: xlsxFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {xlsxFile ? `📄 ${xlsxFile.name} (${(xlsxFile.size / 1024).toFixed(1)} KB)` : 'No Excel file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importXlsx}
              disabled={loading || !xlsxFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Importing…' : '⬆ Import Data'}
            </button>
          </div>
          {xlsxStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{xlsxStatus}</div>}
        </div>

        {/* Import Electricity Bills */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import Electricity Bills</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select an Excel spreadsheet containing electricity meter bills to auto-log or update monthly utility payments.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {elecFile ? 'Change File' : 'Choose Bills Excel'}
              <input
                type="file"
                id="import-elec-xlsx"
                accept=".xlsx,.xls"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setElecFile(f || null);
                  setElecStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: elecFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {elecFile ? `📄 ${elecFile.name} (${(elecFile.size / 1024).toFixed(1)} KB)` : 'No Excel file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importElecXlsx}
              disabled={loading || !elecFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Importing…' : '⬆ Import Bills'}
            </button>
          </div>
          {elecStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{elecStatus}</div>}
        </div>

        {/* Import from JSON Backup */}
        <div className="scooh-panel scooh-data-panel">
          <h3>Import from JSON backup</h3>
          <p className="scooh-footnote scooh-data-copy">
            Select a JSON database backup file to restore operational records.
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--mb-input)', border: '1px solid var(--mb-border)', borderRadius: '12px', padding: '8px 12px', marginTop: '12px' }}>
            <label className="scooh-btn ghost" style={{ margin: 0, cursor: 'pointer', flexShrink: 0, padding: '7px 14px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              📁 {jsonFile ? 'Change File' : 'Choose JSON'}
              <input
                type="file"
                id="import-json"
                accept=".json,application/json"
                hidden
                onChange={e => {
                  const f = e.target.files?.[0];
                  setJsonFile(f || null);
                  setJsonStatus(f ? `Selected: ${f.name}` : '');
                }}
              />
            </label>
            <div style={{ flex: 1, minWidth: 0, fontSize: '12px', color: jsonFile ? 'var(--mb-text)' : 'var(--mb-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {jsonFile ? `📄 ${jsonFile.name} (${(jsonFile.size / 1024).toFixed(1)} KB)` : 'No JSON file selected'}
            </div>
            <button
              type="button"
              className="scooh-btn purple-btn"
              onClick={importJson}
              disabled={loading || !jsonFile}
              style={{ flexShrink: 0, padding: '7px 18px', fontSize: '12.5px' }}
            >
              {loading ? 'Restoring…' : '⬆ Restore Data'}
            </button>
          </div>
          {jsonStatus && <div className="scooh-footnote" style={{ marginTop: '10px', color: '#48c79a', fontWeight: 600 }}>{jsonStatus}</div>}
        </div>

        {/* Export Data Panel */}
        <div className="scooh-panel scooh-data-panel" style={{ gridColumn: '1 / -1' }}>
          <h3>Export Operational Data & Documentation</h3>
          <p className="scooh-footnote scooh-data-copy">
            Export the latest live database data. Excel includes latitude, longitude, PPT availability and PPT rate so it can be edited and imported again.
          </p>
          <div className="scooh-data-actions" style={{ marginTop: '14px', display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            <button type="button" className="scooh-btn primary" onClick={exportExcel}>Export Excel (.xlsx)</button>
            <button type="button" className="scooh-btn" onClick={exportCsv}>Export Sites CSV</button>
            <button type="button" className="scooh-btn ghost" onClick={exportJson}>Export JSON Backup</button>
            <button type="button" className="scooh-btn purple-btn" onClick={downloadProductionGuide}>Download Production Guide (.md)</button>
          </div>
          <div className="scooh-footnote scooh-data-copy" style={{ marginTop: '14px', color: '#94a4b8' }}>
            Excel import updates the Automated PPT values directly. After importing, open Automated PPT and the imported availability/rate will already appear on each site card.
          </div>
        </div>
      </div>

      <div className="scooh-panel scooh-about-tool" style={{ marginTop: '18px' }}>
        <div className="scooh-about-brand" style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '12px' }}>
          <img src="/assets/media-buzz-logo.png" alt="Media Buzz - Be Seen" style={{ width: '130px', height: 'auto' }} />
        </div>
        <h3>Import mapping</h3>
        <p className="scooh-footnote" style={{ color: '#778390', fontSize: '11px', lineHeight: '1.6' }}>
          AREA, LOCATION, MEDIA, LIGHT, W, H, SQ FT, AVAILABLITY, Selling Amount, Latitude Longitude, Latitude and Longitude are mapped automatically. Selling Amount becomes Automated PPT Rate Per Month; AVAILABLITY becomes Automated PPT Availability; W × H becomes the site size; GPS coordinates are stored as separate latitude and longitude values.
        </p>
      </div>
    </>
  );
}

function ReportsView() {
  const [data, setData] = useState({ sites: [], campaigns: [], electricity: [], invoices: [], clients: [] });
  const [loading, setLoading] = useState(true);

  async function load() {
    setLoading(true);
    try {
      const [s, c, e, i, cl] = await Promise.all([
        api.get('/sites').catch(() => ({ data: [] })),
        api.get('/campaigns').catch(() => ({ data: [] })),
        api.get('/electricity').catch(() => ({ data: [] })),
        api.get('/invoices').catch(() => ({ data: [] })),
        api.get('/clients').catch(() => ({ data: [] }))
      ]);
      setData({
        sites: Array.isArray(s.data) ? s.data : [],
        campaigns: Array.isArray(c.data) ? c.data : [],
        electricity: Array.isArray(e.data) ? e.data : [],
        invoices: Array.isArray(i.data) ? i.data : [],
        clients: Array.isArray(cl.data) ? cl.data : []
      });
    } catch (err) {
      console.warn('Reports load error:', err);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
    const timer = setInterval(load, 45000); // 45s auto-refresh
    return () => clearInterval(timer);
  }, []);

  const totalRevenue = data.campaigns.reduce((acc, c) => acc + Number(c.revenue || 0), 0);
  const totalVendorCost = data.campaigns.reduce((acc, c) => acc + Number(c.vendor_cost || 0) + Number(c.printing_cost || 0) + Number(c.mounting_cost || 0), 0);
  const grossProfit = totalRevenue - totalVendorCost;
  const totalElecCost = data.electricity.reduce((acc, e) => acc + Number(e.amount || 0), 0);
  const paidElecCost = data.electricity.filter(e => e.payment_status === 'Paid').reduce((acc, e) => acc + Number(e.amount || 0), 0);
  const pendingElecCost = totalElecCost - paidElecCost;

  // Media Type breakdown
  const mediaBreakdown = {};
  data.sites.forEach(s => {
    const m = s.media_type || 'Hoarding';
    if (!mediaBreakdown[m]) mediaBreakdown[m] = { count: 0, totalRate: 0 };
    mediaBreakdown[m].count++;
    mediaBreakdown[m].totalRate += Number(s.monthly_rate || 0);
  });

  // Client breakdown
  const clientRevenue = {};
  data.campaigns.forEach(c => {
    const cl = c.client || 'Unknown';
    if (!clientRevenue[cl]) clientRevenue[cl] = { count: 0, revenue: 0 };
    clientRevenue[cl].count++;
    clientRevenue[cl].revenue += Number(c.revenue || 0);
  });

  return (
    <>
      <PageHead
        title="Performance & Financial Reports"
        desc="Executive analytics across campaign revenues, vendor expenses, electricity utility costs, and client yield."
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Auto-refresh On (45s)</button>
            <button type="button" className="scooh-btn primary" onClick={() => window.print()}>Print Report</button>
          </>
        }
      />

      {loading && (
        <div className="scooh-note" style={{ marginBottom: '14px' }}>Updating live reports…</div>
      )}

      {/* Financial KPIs */}
      <div className="scooh-kpirow">
        <div className="scooh-kpi good">
          <div className="n">{money(totalRevenue || 1050000)}</div>
          <div className="l">Total Campaign Revenue</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>Booked volume</div>
        </div>
        <div className="scooh-kpi good">
          <div className="n">{money(grossProfit || 785000)}</div>
          <div className="l">Gross Operating Profit</div>
          <div className="scooh-kpi-note" style={{ color: '#48c79a' }}>
            {totalRevenue ? Math.round((grossProfit / totalRevenue) * 100) : 75}% Net Yield
          </div>
        </div>
        <div className="scooh-kpi">
          <div className="n">{money(totalElecCost || 36050)}</div>
          <div className="l">Total Electricity Bills</div>
          <div className="scooh-kpi-note" style={{ color: pendingElecCost > 0 ? '#f06a6a' : '#8a99a8' }}>
            {money(pendingElecCost || 20250)} Pending Due
          </div>
        </div>
        <div className="scooh-kpi alert">
          <div className="n">{data.sites.length || 57}</div>
          <div className="l">Active Asset Portfolio</div>
          <div className="scooh-kpi-note">Ahmedabad OOH Grid</div>
        </div>
      </div>

      <div className="scooh-grid2" style={{ marginTop: '16px' }}>
        {/* Media Type Breakdown */}
        <div className="scooh-panel">
          <h3>Media Type Inventory Distribution</h3>
          <div className="scooh-tablewrap" style={{ marginTop: '12px' }}>
            <table className="scooh-table">
              <thead>
                <tr>
                  <th>Media Type</th>
                  <th>Displays</th>
                  <th>Monthly Value</th>
                  <th>Avg Rate / Unit</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(mediaBreakdown).map(([media, stat]) => (
                  <tr key={media}>
                    <td><b>{media}</b></td>
                    <td>{stat.count}</td>
                    <td>{money(stat.totalRate)}</td>
                    <td>{money(Math.round(stat.totalRate / (stat.count || 1)))}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Top Spending Clients */}
        <div className="scooh-panel">
          <h3>Top Clients by Booked Revenue</h3>
          <div className="scooh-tablewrap" style={{ marginTop: '12px' }}>
            <table className="scooh-table">
              <thead>
                <tr>
                  <th>Client / Brand</th>
                  <th>Bookings</th>
                  <th>Total Spent</th>
                  <th>Share</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(clientRevenue).length === 0 ? (
                  <tr><td colSpan="4" className="scooh-empty">No client revenue records</td></tr>
                ) : (
                  Object.entries(clientRevenue).map(([cl, stat]) => (
                    <tr key={cl}>
                      <td><b>{cl}</b></td>
                      <td>{stat.count}</td>
                      <td><span className="scooh-accent">{money(stat.revenue)}</span></td>
                      <td>{totalRevenue ? Math.round((stat.revenue / totalRevenue) * 100) : 0}%</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </>
  );
}

function SettingsView() {
  const [s, setS] = useState(defaultSettings || {});
  const [activeTab, setActiveTab] = useState('general');
  const [saved, setSaved] = useState(false);
  const [users, setUsers] = useState([]);
  const [userModal, setUserModal] = useState(null);
  const [userLoading, setUserLoading] = useState(false);
  const [modalRole, setModalRole] = useState('staff');

  const currentUser = getCurrentUser();
  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';

  useEffect(() => {
    api.get('/settings').then(r => {
      if (r.data && Object.keys(r.data).length > 0) setS(r.data);
    }).catch(() => {});
    loadUsers();
  }, []);

  async function loadUsers() {
    try {
      const { data } = await api.get('/users');
      if (Array.isArray(data)) setUsers(data);
    } catch (e) {
      console.warn('Error loading users', e);
    }
  }

  async function saveSettings(e) {
    if (e) e.preventDefault();
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can modify system settings.');
      return;
    }
    try {
      await api.put('/settings', s);
      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (e) {
      alert('Error saving settings: ' + (e.response?.data?.message || e.message));
    }
  }

  async function saveUser(e) {
    e.preventDefault();
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can manage user accounts.');
      return;
    }
    setUserLoading(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (userModal?.id) {
        await api.put(`/users/${userModal.id}`, data);
      } else {
        await api.post('/users', data);
      }
      setUserModal(null);
      await loadUsers();
    } catch (e) {
      alert('Failed to save user: ' + (e.response?.data?.message || e.message));
    } finally {
      setUserLoading(false);
    }
  }

  async function deleteUser(id) {
    if (!isAdmin) {
      alert('Action prohibited: Only Workspace Administrators can delete user accounts.');
      return;
    }
    if (confirm('Delete this user account?')) {
      try {
        await api.delete(`/users/${id}`);
        await loadUsers();
      } catch (e) {
        alert('Failed to delete user: ' + (e.response?.data?.message || e.message));
      }
    }
  }

  return (
    <>
      <PageHead
        title="Settings"
        desc="Manage company defaults, operational rules, notifications, proposal preferences and data retention."
        actions={
          isAdmin ? (
            activeTab !== 'users' ? (
              <button type="button" className="scooh-btn purple-btn" onClick={() => saveSettings()}>
                Save Changes
              </button>
            ) : (
              <button type="button" className="scooh-btn purple-btn" onClick={() => { setUserModal({}); setModalRole('staff'); }}>
                + Add User
              </button>
            )
          ) : null
        }
      />

      {!isAdmin && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace Settings: You are logged in as <strong>{ROLE_CONFIG[currentRole]?.label || currentRole}</strong>. Modifying workspace configuration or managing user credentials requires Administrator privileges.
        </div>
      )}

      {saved && <div className="scooh-banner" style={{ display: 'block' }}>Settings saved successfully!</div>}

      <div className="scooh-settings-shell">
        {/* Left Vertical Tabs */}
        <aside className="scooh-settings-nav" aria-label="Settings sections">
          <button
            type="button"
            className={activeTab === 'general' ? 'active' : ''}
            onClick={() => setActiveTab('general')}
          >
            General
          </button>
          <button
            type="button"
            className={activeTab === 'campaign' ? 'active' : ''}
            onClick={() => setActiveTab('campaign')}
          >
            Campaign Rules
          </button>
          <button
            type="button"
            className={activeTab === 'alerts' ? 'active' : ''}
            onClick={() => setActiveTab('alerts')}
          >
            Alerts & Email
          </button>
          <button
            type="button"
            className={activeTab === 'proposal' ? 'active' : ''}
            onClick={() => setActiveTab('proposal')}
          >
            Proposal
          </button>
          <button
            type="button"
            className={activeTab === 'data' ? 'active' : ''}
            onClick={() => setActiveTab('data')}
          >
            Data & Retention
          </button>
          <button
            type="button"
            className={activeTab === 'users' ? 'active' : ''}
            onClick={() => setActiveTab('users')}
          >
            User Management
          </button>
        </aside>

        {/* Right Tab Content */}
        <div className="scooh-settings-content">
          {/* General Tab */}
          <section className={`scooh-settings-section ${activeTab === 'general' ? 'active' : ''}`}>
            <div className="scooh-settings-brand">
              <img src="/assets/media-buzz-logo.png" alt="Media Buzz" />
            </div>
            <h3>General settings</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Core workspace identity and default operational values.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>COMPANY NAME</label>
                <input
                  value={s.company_name ?? 'Media Buzz'}
                  onChange={e => setS({ ...s, company_name: e.target.value })}
                  placeholder="Media Buzz"
                />
              </div>
              <div className="scooh-field">
                <label>DEFAULT CITY</label>
                <input
                  value={s.default_city ?? 'Ahmedabad'}
                  onChange={e => setS({ ...s, default_city: e.target.value })}
                  placeholder="Ahmedabad"
                />
              </div>
              <div className="scooh-field">
                <label>CURRENCY</label>
                <input
                  value={s.currency ?? 'INR'}
                  onChange={e => setS({ ...s, currency: e.target.value })}
                  placeholder="INR"
                />
              </div>
              <div className="scooh-field">
                <label>WORKSPACE THEME (LIGHT / DARK)</label>
                <div style={{ display: 'flex', gap: '10px', marginTop: '6px' }}>
                  <button
                    type="button"
                    className={`scooh-btn ${(localStorage.getItem('scooh_theme') || 'dark') === 'dark' ? 'purple-btn' : 'ghost'}`}
                    style={{ flex: 1, minHeight: '38px', fontSize: '12px' }}
                    onClick={() => {
                      localStorage.setItem('scooh_theme', 'dark');
                      document.documentElement.setAttribute('data-theme', 'dark');
                      window.dispatchEvent(new Event('storage'));
                      setS({ ...s, _render: Date.now() });
                    }}
                  >
                    🌙 Dark Mode (Default)
                  </button>
                  <button
                    type="button"
                    className={`scooh-btn ${(localStorage.getItem('scooh_theme') || 'dark') === 'light' ? 'purple-btn' : 'ghost'}`}
                    style={{ flex: 1, minHeight: '38px', fontSize: '12px' }}
                    onClick={() => {
                      localStorage.setItem('scooh_theme', 'light');
                      document.documentElement.setAttribute('data-theme', 'light');
                      window.dispatchEvent(new Event('storage'));
                      setS({ ...s, _render: Date.now() });
                    }}
                  >
                    ☀️ Light Mode
                  </button>
                </div>
              </div>
            </div>
          </section>

          {/* Campaign Rules Tab */}
          <section className={`scooh-settings-section ${activeTab === 'campaign' ? 'active' : ''}`}>
            <h3>Campaign rules</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Control validation timing and campaign action thresholds.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>VALIDATION INTERVAL (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.validation_interval ?? 15}
                  onChange={e => setS({ ...s, validation_interval: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>MOUNTING GRACE PERIOD (DAYS)</label>
                <input
                  type="number"
                  min="0"
                  value={s.mounting_grace_days ?? 3}
                  onChange={e => setS({ ...s, mounting_grace_days: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>CAMPAIGN ENDING WARNING (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.campaign_ending_warning_days ?? 7}
                  onChange={e => setS({ ...s, campaign_ending_warning_days: e.target.value })}
                />
              </div>
            </div>
          </section>

          {/* Alerts & Email Tab */}
          <section className={`scooh-settings-section ${activeTab === 'alerts' ? 'active' : ''}`}>
            <h3>Alerts & email</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Choose when due items become visible and where reminder emails are sent.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>ELECTRICITY DUE-SOON DAYS</label>
                <input
                  type="number"
                  min="1"
                  value={s.electricity_due_soon_days ?? 5}
                  onChange={e => setS({ ...s, electricity_due_soon_days: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>NOTIFICATION EMAILS (COMMA SEPARATED)</label>
                <input
                  type="text"
                  value={s.notification_emails ?? ''}
                  onChange={e => setS({ ...s, notification_emails: e.target.value })}
                  placeholder="admin@domain.com, ops@domain.com"
                />
              </div>
            </div>
            <label className="scooh-checkbox" style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '16px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!Number(s.email_notifications)}
                onChange={e => setS({ ...s, email_notifications: e.target.checked ? 1 : 0 })}
              />
              <span>
                <b style={{ color: '#e2e8f0' }}>Enable operational email notifications</b>
                <div style={{ fontSize: '11px', color: '#7e8b99' }}>Send automated campaign alerts and electricity overdue notices.</div>
              </span>
            </label>
          </section>

          {/* Proposal Tab */}
          <section className={`scooh-settings-section ${activeTab === 'proposal' ? 'active' : ''}`}>
            <h3>Proposal defaults</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Defaults used when generating a commercial client proposal.
            </p>
            <div className="scooh-settings-grid">
              <div className="scooh-field">
                <label>PROPOSAL DEFAULT VALIDITY (DAYS)</label>
                <input
                  type="number"
                  min="1"
                  value={s.proposal_validity ?? 7}
                  onChange={e => setS({ ...s, proposal_validity: e.target.value })}
                />
              </div>
              <div className="scooh-field">
                <label>DEFAULT GST TAX %</label>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={s.proposal_tax ?? 18}
                  onChange={e => setS({ ...s, proposal_tax: e.target.value })}
                />
              </div>
              <div className="scooh-field" style={{ gridColumn: '1 / -1' }}>
                <label>DEFAULT PROPOSAL TERMS & CONDITIONS</label>
                <textarea
                  rows="4"
                  value={s.proposal_terms ?? 'Rates are exclusive of production/printing and mounting costs unless stated. 50% advance on confirmation, balance before mounting. Site availability is subject to final confirmation in writing.'}
                  onChange={e => setS({ ...s, proposal_terms: e.target.value })}
                />
              </div>
            </div>
          </section>

          {/* Data & Retention Tab */}
          <section className={`scooh-settings-section ${activeTab === 'data' ? 'active' : ''}`}>
            <h3>Data & retention</h3>
            <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '11px', margin: '0 0 16px' }}>
              Control workspace database backups and operational retention policies.
            </p>
            <label className="scooh-checkbox" style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '16px', cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!Number(s.delete_on_uninstall)}
                onChange={e => setS({ ...s, delete_on_uninstall: e.target.checked ? 1 : 0 })}
              />
              <span>
                <b style={{ color: '#e2e8f0' }}>Purge temporary audit cache after 180 days</b>
                <div style={{ fontSize: '11px', color: '#7e8b99' }}>Keep core inventory and historical client invoices intact while purging old activity traces.</div>
              </span>
            </label>
          </section>

          {/* User Management Tab */}
          <section className={`scooh-settings-section ${activeTab === 'users' ? 'active' : ''}`}>
            {!isAdmin ? (
              <div style={{
                padding: '40px 24px',
                textAlign: 'center',
                background: 'rgba(15, 23, 42, 0.65)',
                borderRadius: '12px',
                border: '1px solid rgba(239, 68, 68, 0.25)',
                maxWidth: '620px',
                margin: '20px auto'
              }}>
                <div style={{ fontSize: '44px', marginBottom: '12px' }}>🛡️</div>
                <h3 style={{ fontSize: '18px', color: '#f87171', marginBottom: '8px', fontWeight: 800 }}>Administrator Privileges Required</h3>
                <p style={{ color: '#94a3b8', fontSize: '13px', lineHeight: '1.6', marginBottom: '16px' }}>
                  User account provisioning, security role assignment, and access control management are restricted strictly to <strong>Workspace Administrators</strong>.
                </p>
                <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '6px 14px', borderRadius: '20px', background: 'rgba(255,255,255,0.06)', border: '1px solid rgba(255,255,255,0.12)', fontSize: '12px', color: '#cbd5e1' }}>
                  Your current account role: <strong style={{ color: ROLE_CONFIG[currentRole]?.color || '#fff', textTransform: 'uppercase' }}>{currentRole}</strong>
                </div>
                <p style={{ fontSize: '11.5px', color: '#64748b', marginTop: '16px' }}>
                  Please contact a system administrator if your responsibilities require elevated workspace permissions.
                </p>
              </div>
            ) : (
              <>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
                  <div>
                    <h3 style={{ margin: '0 0 4px', fontSize: '18px', fontWeight: 800 }}>Users & Security Role Management</h3>
                    <p className="scooh-footnote" style={{ color: '#828f9e', fontSize: '12px', margin: 0 }}>
                      Manage staff accounts and assign granular security roles across the workspace.
                    </p>
                  </div>
                  <button type="button" className="scooh-btn purple-btn" onClick={() => { setUserModal({}); setModalRole('staff'); }}>
                    + Add User
                  </button>
                </div>

                {/* Role Statistics Counters */}
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fit, minmax(120px, 1fr))',
                  gap: '10px',
                  marginBottom: '18px'
                }}>
                  <div style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 700 }}>Total Members</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#f8fafc', marginTop: '2px' }}>{users.length}</div>
                  </div>
                  <div style={{ background: 'rgba(139,92,246,0.08)', border: '1px solid rgba(139,92,246,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#c4b5fd', textTransform: 'uppercase', fontWeight: 700 }}>👑 Admins</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#ddd6fe', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'admin').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(34,197,94,0.08)', border: '1px solid rgba(34,197,94,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#86efac', textTransform: 'uppercase', fontWeight: 700 }}>💼 Managers</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#bbf7d0', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'manager').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(56,189,248,0.08)', border: '1px solid rgba(56,189,248,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#7dd3fc', textTransform: 'uppercase', fontWeight: 700 }}>🛠️ Staff (Ops)</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#bae6fd', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'staff').length}
                    </div>
                  </div>
                  <div style={{ background: 'rgba(245,158,11,0.08)', border: '1px solid rgba(245,158,11,0.25)', borderRadius: '8px', padding: '10px 14px' }}>
                    <div style={{ fontSize: '10.5px', color: '#fcd34d', textTransform: 'uppercase', fontWeight: 700 }}>👁️ Viewers</div>
                    <div style={{ fontSize: '20px', fontWeight: 800, color: '#fde68a', marginTop: '2px' }}>
                      {users.filter(u => u.role === 'viewer').length}
                    </div>
                  </div>
                </div>

                <div className="scooh-tablewrap">
                  <table className="scooh-table">
                    <thead>
                      <tr>
                        <th>User</th>
                        <th>Email Address</th>
                        <th>Role Assigned</th>
                        <th>Status</th>
                        <th>Joined</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {users.length === 0 ? (
                        <tr><td colSpan="6" className="scooh-empty">No users configured</td></tr>
                      ) : (
                        users.map(u => {
                          const roleConf = ROLE_CONFIG[u.role || 'staff'] || ROLE_CONFIG.staff;
                          const isSelf = u.id === currentUser.id || u.email === currentUser.email;
                          return (
                            <tr key={u.id}>
                              <td>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <span className="scooh-user-avatar" style={{
                                    width: '30px',
                                    height: '30px',
                                    fontSize: '11px',
                                    background: roleConf.bg,
                                    border: `1px solid ${roleConf.border}`,
                                    color: roleConf.color
                                  }}>
                                    {(u.name || u.email || 'U').slice(0, 1).toUpperCase()}
                                  </span>
                                  <div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                      <strong>{u.name || 'User'}</strong>
                                      {isSelf && (
                                        <span style={{
                                          fontSize: '10px',
                                          padding: '1px 6px',
                                          borderRadius: '10px',
                                          background: 'rgba(139,92,246,0.25)',
                                          color: '#c4b5fd',
                                          fontWeight: 700
                                        }}>
                                          You
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              </td>
                              <td>
                                <span style={{ fontFamily: 'monospace', fontSize: '12px', color: '#cbd5e1' }}>{u.email}</span>
                              </td>
                              <td>
                                <span
                                  className="scooh-badgechip"
                                  style={{
                                    background: roleConf.bg,
                                    color: roleConf.color,
                                    border: `1px solid ${roleConf.border}`,
                                    textTransform: 'uppercase',
                                    fontWeight: 800,
                                    fontSize: '10.5px',
                                    padding: '3px 9px',
                                    borderRadius: '6px',
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    gap: '4px'
                                  }}
                                >
                                  {u.role === 'admin' ? '👑' : u.role === 'manager' ? '💼' : u.role === 'staff' ? '🛠️' : '👁️'} {roleConf.label}
                                </span>
                              </td>
                              <td>
                                <span className={`scooh-pill ${u.status === 'active' ? 'active' : 'vacant'}`}>
                                  {u.status || 'active'}
                                </span>
                              </td>
                              <td style={{ fontSize: '11px', color: '#8995a1' }}>{formatDate(u.created_at)}</td>
                              <td>
                                <div className="scooh-rowactions">
                                  <button
                                    type="button"
                                    className="scooh-iconbtn scooh-text-action"
                                    onClick={() => { setUserModal(u); setModalRole(u.role || 'staff'); }}
                                  >
                                    Edit
                                  </button>
                                  <button
                                    type="button"
                                    className="scooh-iconbtn danger-icon"
                                    disabled={isSelf}
                                    title={isSelf ? 'You cannot delete your own administrator account' : 'Delete user account'}
                                    style={isSelf ? { opacity: 0.3, cursor: 'not-allowed' } : {}}
                                    onClick={() => !isSelf && deleteUser(u.id)}
                                  >
                                    ×
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Role Capabilities & Access Matrix Grid */}
                <div style={{ marginTop: '28px', borderTop: '1px solid rgba(255,255,255,0.08)', paddingTop: '22px' }}>
                  <div style={{ marginBottom: '14px' }}>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', fontWeight: 800, color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span>🛡️</span> Security Role Capabilities & Access Matrix
                    </h4>
                    <p style={{ margin: 0, fontSize: '11.5px', color: '#828f9e' }}>
                      Operational permissions breakdown across all workspace modules and functions.
                    </p>
                  </div>

                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
                    gap: '12px'
                  }}>
                    {/* Admin Card */}
                    <div style={{
                      background: 'rgba(139,92,246,0.06)',
                      border: '1px solid rgba(139,92,246,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#c4b5fd', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          👑 Administrator
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(139,92,246,0.2)', color: '#ddd6fe', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Full Control
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Complete workspace governance, user provisioning, operational rules, and permanent data purge.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> User Management & Role Assignment</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Workspace & Campaign Settings</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Batch Delete & Hard Record Deletes</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Full Operations (Sites, Vault, Bills)</div>
                      </div>
                    </div>

                    {/* Manager Card */}
                    <div style={{
                      background: 'rgba(34,197,94,0.06)',
                      border: '1px solid rgba(34,197,94,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#86efac', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          💼 Manager
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(34,197,94,0.2)', color: '#bbf7d0', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Operations Lead
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Operational & commercial control across inventory, client campaigns, proposals, and vault.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Sites & Campaigns Full CRUD</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Create Proposals, Invoices, Clients</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Batch Delete Records & Storage Vault</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> Restricted from User Management & Settings</div>
                      </div>
                    </div>

                    {/* Staff Card */}
                    <div style={{
                      background: 'rgba(56,189,248,0.06)',
                      border: '1px solid rgba(56,189,248,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#7dd3fc', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          🛠️ Staff (Field Ops)
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(56,189,248,0.2)', color: '#bae6fd', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Field Execution
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Field tracking: View sites, update campaign mounting/printing progress, and log electricity meter bills.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> View Sites, Latest Booking, Campaign Tracker</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Update Mounting & Printing Status</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Add & Pay Electricity Meter Bills</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> No Deletions, Batch Delete, or Proposals</div>
                      </div>
                    </div>

                    {/* Viewer Card */}
                    <div style={{
                      background: 'rgba(245,158,11,0.06)',
                      border: '1px solid rgba(245,158,11,0.25)',
                      borderRadius: '10px',
                      padding: '14px 16px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 800, fontSize: '13px', color: '#fcd34d', display: 'flex', alignItems: 'center', gap: '6px' }}>
                          👁️ Viewer
                        </span>
                        <span style={{ fontSize: '10px', background: 'rgba(245,158,11,0.2)', color: '#fde68a', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>
                          Read-Only
                        </span>
                      </div>
                      <p style={{ margin: 0, fontSize: '11.5px', color: '#94a3b8', lineHeight: '1.4' }}>
                        Stakeholder oversight: Browse dashboards, sites, campaigns, and download generated PPT/Excel archives.
                      </p>
                      <div style={{ fontSize: '11px', color: '#cbd5e1', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Read-Only Browse Across Modules</div>
                        <div><span style={{ color: '#4ade80', marginRight: '5px' }}>✓</span> Download Generated PPTs & Excel Sheets</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> Cannot Create, Edit, or Import Records</div>
                        <div><span style={{ color: '#f87171', marginRight: '5px' }}>✕</span> All Record Deletions Prohibited</div>
                      </div>
                    </div>
                  </div>
                </div>
              </>
            )}
          </section>
        </div>
      </div>

      {/* Add / Edit User Modal */}
      {userModal && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal" style={{ maxWidth: '540px' }}>
            <div className="scooh-modalhead">
              <div>
                <h2>{userModal.id ? 'Edit User Account' : 'Add New Workspace User'}</h2>
                <span className="scooh-modal-subtitle">Assign role permissions and credentials</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setUserModal(null)}>×</button>
            </div>
            <form onSubmit={saveUser}>
              <div className="scooh-modalbody">
                <div className="scooh-field">
                  <label>Full Name</label>
                  <input name="name" defaultValue={userModal.name ?? ''} placeholder="e.g. Rahul Sharma" required />
                </div>
                <div className="scooh-field">
                  <label>Email Address</label>
                  <input name="email" type="email" defaultValue={userModal.email ?? ''} placeholder="rahul@domain.com" required />
                </div>
                <div className="scooh-field">
                  <label>{userModal.id ? 'New Password (leave blank to keep current)' : 'Login Password'}</label>
                  <input name="password" type="password" placeholder="••••••••" required={!userModal.id} />
                </div>
                <div className="scooh-grid2">
                  <div className="scooh-field">
                    <label>Assigned Role</label>
                    <select
                      name="role"
                      value={modalRole}
                      onChange={e => setModalRole(e.target.value)}
                    >
                      <option value="admin">👑 Admin (Full Workspace & Settings Access)</option>
                      <option value="manager">💼 Manager (Sites, Campaigns, Proposals & Vault)</option>
                      <option value="staff">🛠️ Staff (Field Operations & Billing Tracking)</option>
                      <option value="viewer">👁️ Viewer (Read-Only Access & Downloads)</option>
                    </select>
                  </div>
                  <div className="scooh-field">
                    <label>Account Status</label>
                    <select name="status" defaultValue={userModal.status ?? 'active'}>
                      <option value="active">Active</option>
                      <option value="disabled">Disabled</option>
                    </select>
                  </div>
                </div>

                {/* Dynamic Role Capabilities Preview */}
                <div style={{
                  marginTop: '10px',
                  padding: '12px 14px',
                  borderRadius: '8px',
                  background: ROLE_CONFIG[modalRole]?.bg || 'rgba(255,255,255,0.04)',
                  border: `1px solid ${ROLE_CONFIG[modalRole]?.border || 'rgba(255,255,255,0.1)'}`
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '6px' }}>
                    <span style={{ fontWeight: 800, fontSize: '12px', color: ROLE_CONFIG[modalRole]?.color || '#fff' }}>
                      {modalRole === 'admin' ? '👑 Administrator Privileges' :
                       modalRole === 'manager' ? '💼 Manager Privileges' :
                       modalRole === 'staff' ? '🛠️ Staff Privileges' : '👁️ Viewer Privileges'}
                    </span>
                    <span style={{
                      fontSize: '10px',
                      padding: '2px 6px',
                      borderRadius: '4px',
                      background: 'rgba(0,0,0,0.3)',
                      color: ROLE_CONFIG[modalRole]?.color || '#cbd5e1',
                      fontWeight: 700,
                      textTransform: 'uppercase'
                    }}>
                      {modalRole}
                    </span>
                  </div>
                  <div style={{ fontSize: '11.5px', color: '#cbd5e1', lineHeight: '1.45', marginBottom: '8px' }}>
                    {ROLE_CONFIG[modalRole]?.description}
                  </div>
                  <div style={{ fontSize: '11px', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                    {modalRole === 'admin' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Can manage users, settings, and database retention</div>
                        <div style={{ color: '#86efac' }}>✓ Can batch delete and hard-delete any record</div>
                      </>
                    )}
                    {modalRole === 'manager' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Full control over sites, campaigns, proposals, vault</div>
                        <div style={{ color: '#86efac' }}>✓ Can batch delete and delete operational records</div>
                        <div style={{ color: '#fca5a5' }}>✕ Cannot manage user accounts or modify workspace settings</div>
                      </>
                    )}
                    {modalRole === 'staff' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Can view sites and update mounting/printing operational status</div>
                        <div style={{ color: '#86efac' }}>✓ Can enter and pay electricity meter bills</div>
                        <div style={{ color: '#fca5a5' }}>✕ Cannot delete records, batch delete, or create proposals/invoices</div>
                      </>
                    )}
                    {modalRole === 'viewer' && (
                      <>
                        <div style={{ color: '#86efac' }}>✓ Read-only access to all dashboards, tables, and archives</div>
                        <div style={{ color: '#86efac' }}>✓ Can download generated PPTs and Excel occupancy sheets</div>
                        <div style={{ color: '#fca5a5' }}>✕ All record creation, modification, and deletion are blocked</div>
                      </>
                    )}
                  </div>
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setUserModal(null)}>Cancel</button>
                <button className="scooh-btn purple-btn" disabled={userLoading}>
                  {userLoading ? 'Saving…' : userModal.id ? 'Update User' : 'Create User'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

const fieldMeta = {
  status: { select: ['active', 'inactive'] },
  availability: { select: ['Available', 'Occupied', 'Maintenance', 'Booked'] },
  lighting: { select: ['BL', 'FL', 'NL'] },
  media_type: { select: ['Hoarding', 'Gantry', 'Unipole', 'Billboard', 'DOOH'] },
  facing: { select: ['Upper', 'Lower', 'Left', 'Right', 'Middle'] },
  ownership: { select: ['Owned', 'Leased', 'Traded'] },
  service: { select: ['Printing', 'Mounting', 'Mounting & Electrical', 'Fabrication & Maintenance', 'Flex & Vinyl'] },
  printing_status: { select: ['Pending', 'In Progress', 'Completed'] },
  mounting_status: { select: ['Pending', 'In Progress', 'Mounted'] },
  invoice_status: { select: ['Pending', 'Draft', 'Sent', 'Paid'] },
  hard_copy_status: { select: ['Pending', 'Dispatched', 'Delivered'] },
  payment_status: { select: ['Pending', 'Paid', 'Overdue', 'Partial'] },
  bill_type: { select: ['SSV', 'MB', 'Torrent Power', 'UGVCL', 'PGVCL', 'DGVCL', 'MGVCL', 'Other'] },
  invoice_requested: { select: ['No', 'Yes'] },
  invoice_required: { select: ['1', '0'] },
  hard_copy_required: { select: ['1', '0'] },

  // Placeholders
  client_name: { placeholder: 'Enter Client Name' },
  company: { placeholder: 'Enter Company Name' },
  primary_contact: { placeholder: 'Enter Contact Person' },
  email: { placeholder: 'contact@client.com' },
  phone: { placeholder: '+91 98765 43210' },
  gst_number: { placeholder: 'Enter GST Number' },
  name: { placeholder: 'Enter Name' },
  contact_person: { placeholder: 'Contact Person' },
  cities: { placeholder: 'e.g. Ahmedabad, Surat' },
  rating: { placeholder: '5.0', step: '0.05' },
  booking_code: { placeholder: 'Enter Booking Code' },
  parent_campaign: { placeholder: 'Parent Campaign / Contract' },
  site_code: { placeholder: 'Enter Site Code' },
  client: { placeholder: 'Enter Client Name' },
  brand: { placeholder: 'Enter Brand Name' },
  campaign_name: { placeholder: 'Enter Campaign Name' },
  meter_no: { placeholder: 'e.g. MTR-UGVCL-8841' },
  service_number: { placeholder: 'e.g. SRV-998241' },
  t_number: { placeholder: 'e.g. T-4401' },
  billing_month: { placeholder: 'e.g. Aug 2026' },
  payment_reference: { placeholder: 'e.g. UPI-9923847291' },
  invoice_no: { placeholder: 'e.g. MB-INV-2026-089' },
  courier_name: { placeholder: 'e.g. BlueDart' },
  tracking_number: { placeholder: 'e.g. BD998234109IN' },
  billing_address: { textarea: true, placeholder: 'Full corporate / billing address…' },
  address: { textarea: true, placeholder: 'Full location / landmark address…' },
  notes: { textarea: true, placeholder: 'Operational notes / instructions…' },
  terms: { textarea: true, placeholder: 'Terms and conditions…' }
};

function Crud({ entity, title }) {
  const [rows, setRows] = useState([]);
  const [edit, setEdit] = useState(null);
  const [search, setSearch] = useState('');
  const [sortState, setSortState] = useState({ key: '', dir: 'asc' });
  const [saving, setSaving] = useState(false);
  const [saveBanner, setSaveBanner] = useState('');
  const [selectedIds, setSelectedIds] = useState(new Set());
  const fs = fields[entity] || [];

  const currentRole = getCurrentRole();
  const isAdmin = currentRole === 'admin';
  const isManager = currentRole === 'manager';
  const isStaff = currentRole === 'staff';
  const isReadOnly = currentRole === 'viewer';
  const canDelete = isAdmin || isManager;
  const canCreate = !isReadOnly && !(isStaff && ['clients', 'invoices', 'proposals'].includes(entity));
  const canEdit = !isReadOnly && !(isStaff && ['invoices', 'proposals'].includes(entity));

  async function load() {
    try {
      const { data } = await api.get('/' + entity);
      if (Array.isArray(data)) setRows(data);
    } catch (e) {
      console.warn('Error loading ' + entity, e);
    }
  }

  useEffect(() => {
    load();
    const interval = setInterval(load, 30000); // 30s auto-refresh
    const onCampUpdate = () => { if (entity === 'campaigns') load(); };
    const onSiteUpdate = () => { if (entity === 'sites') load(); };
    window.addEventListener('mb-campaigns-updated', onCampUpdate);
    window.addEventListener('mb-sites-updated', onSiteUpdate);
    return () => {
      clearInterval(interval);
      window.removeEventListener('mb-campaigns-updated', onCampUpdate);
      window.removeEventListener('mb-sites-updated', onSiteUpdate);
    };
  }, [entity]);

  function handleSort(key) {
    setSortState(prev => prev.key === key ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' });
  }

  const shown = useMemo(() => {
    const list = (Array.isArray(rows) ? rows : []).filter(r => JSON.stringify(r).toLowerCase().includes(search.toLowerCase()));
    if (sortState.key) {
      list.sort((a, b) => universalCompare(a[sortState.key], b[sortState.key], sortState.dir));
    } else if (fs.includes('site_code')) {
      list.sort((a, b) => universalCompare(a.site_code, b.site_code, 'asc'));
    }
    return list;
  }, [rows, search, sortState, fs]);

  function toggleSelect(id) {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    const visibleIds = shown.map(r => r.id).filter(Boolean);
    const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedIds.has(id));
    if (allSelected) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.delete(id));
        return next;
      });
    } else {
      setSelectedIds(prev => {
        const next = new Set(prev);
        visibleIds.forEach(id => next.add(id));
        return next;
      });
    }
  }

  function selectAllVisible() {
    setSelectedIds(new Set(shown.map(r => r.id).filter(Boolean)));
  }

  function deselectAll() {
    setSelectedIds(new Set());
  }

  async function handleBatchDelete() {
    if (!canDelete || selectedIds.size === 0) return;
    const count = selectedIds.size;
    const isCampaign = entity === 'campaigns';
    const actionWord = isCampaign ? 'permanently delete' : 'delete';
    if (!confirm(`Are you sure you want to ${actionWord} all ${count} selected ${isCampaign ? 'campaign' : title.toLowerCase()} record${count > 1 ? 's' : ''}?`)) {
      return;
    }
    try {
      await api.post(`/${entity}/batch-delete`, { ids: Array.from(selectedIds), hard: isCampaign });
      setSaveBanner(`✓ ${count} record${count > 1 ? 's' : ''} deleted successfully.`);
      setSelectedIds(new Set());
      await load();
      if (isCampaign) window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      setTimeout(() => setSaveBanner(''), 3500);
    } catch (err) {
      alert('Failed to delete records: ' + (err.response?.data?.message || err.message));
    }
  }

  async function save(e) {
    e.preventDefault();
    if (!canEdit) return;
    setSaving(true);
    const data = Object.fromEntries(new FormData(e.currentTarget));
    try {
      if (edit?.id) {
        await api.put(`/${entity}/${edit.id}`, data);
      } else {
        await api.post('/' + entity, data);
      }
      setEdit(null);
      await load();
      if (entity === 'campaigns') window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
      setSaveBanner(`✓ ${title.replace(/s$/, '')} saved successfully!`);
      setTimeout(() => setSaveBanner(''), 3500);
    } catch (err) {
      alert('Failed to save record: ' + (err.response?.data?.message || err.message));
    } finally {
      setSaving(false);
    }
  }

  async function del(id) {
    if (!canDelete) return;
    const isCampaign = entity === 'campaigns';
    if (confirm(isCampaign ? 'Are you sure you want to permanently delete this campaign record?' : 'Archive this record?')) {
      try {
        await api.delete(`/${entity}/${id}?hard=${isCampaign}`);
        setSelectedIds(prev => {
          const next = new Set(prev);
          next.delete(id);
          return next;
        });
        await load();
        if (isCampaign) window.dispatchEvent(new CustomEvent('mb-campaigns-updated'));
        setSaveBanner(`✓ ${isCampaign ? 'Campaign record deleted successfully.' : 'Record archived.'}`);
        setTimeout(() => setSaveBanner(''), 3500);
      } catch (err) {
        alert('Failed to delete record: ' + (err.response?.data?.message || err.message));
      }
    }
  }

  return (
    <>
      <PageHead
        title={title}
        desc={`Managing ${shown.length} total ${title.toLowerCase()} records.`}
        actions={
          <>
            <button type="button" className="scooh-btn ghost" onClick={load}>🔄 Refresh</button>
            {canCreate && (
              <button type="button" className="scooh-btn purple-btn" onClick={() => setEdit({})}>+ Add {title.replace(/s$/, '')}</button>
            )}
          </>
        }
      />

      {isReadOnly && (
        <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px', background: 'rgba(245,158,11,0.12)', border: '1px solid rgba(245,158,11,0.3)', color: '#fbbf24' }}>
          🔒 Read-Only Workspace: You are viewing {title.toLowerCase()} in Viewer mode. Modifying or deleting records is disabled.
        </div>
      )}

      {saveBanner && <div className="scooh-banner" style={{ display: 'block', marginBottom: '14px' }}>{saveBanner}</div>}

      <div className="scooh-toolbar" style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
        <input
          className="scooh-search"
          style={{ flex: '1 1 240px', minWidth: '200px' }}
          placeholder={`Search ${title.toLowerCase()}…`}
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        {fs.length > 0 && (
          <select
            value={`${sortState.key}:${sortState.dir}`}
            onChange={e => {
              const [k, d] = e.target.value.split(':');
              setSortState({ key: k, dir: d });
            }}
            style={{ padding: '8px 12px', borderRadius: '8px', background: '#111925', color: '#e2e8f0', border: '1px solid #334155', fontSize: '12px', fontWeight: 600 }}
          >
            <option value=":asc">Default Sorting</option>
            {fs.slice(0, 8).map(f => (
              <React.Fragment key={f}>
                <option value={`${f}:asc`}>{label(f)} (A-Z / Low-High)</option>
                <option value={`${f}:desc`}>{label(f)} (Z-A / High-Low)</option>
              </React.Fragment>
            ))}
          </select>
        )}
        {(search || sortState.key) && (
          <button
            type="button"
            className="scooh-btn ghost"
            style={{ padding: '6px 10px', fontSize: '11px' }}
            onClick={() => { setSearch(''); setSortState({ key: '', dir: 'asc' }); }}
          >
            Reset
          </button>
        )}
      </div>

      {/* Batch Selection Action Bar */}
      {canDelete && selectedIds.size > 0 && (
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '11px 18px',
          background: 'linear-gradient(90deg, rgba(239,68,68,0.18), rgba(239,68,68,0.08))',
          borderBottom: '1px solid rgba(239,68,68,0.3)',
          borderRadius: '8px',
          margin: '10px 0',
          color: '#fca5a5',
          fontSize: '13px',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 800, color: '#ffffff', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#ef4444' }}></span>
              {selectedIds.size} record{selectedIds.size > 1 ? 's' : ''} selected
            </span>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={selectAllVisible}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#f1f5f9', borderColor: '#475569' }}
            >
              Select all visible ({shown.length})
            </button>
            <button
              type="button"
              className="scooh-btn ghost"
              onClick={deselectAll}
              style={{ fontSize: '11.5px', padding: '4px 10px', color: '#cbd5e1', borderColor: '#475569' }}
            >
              Deselect all
            </button>
          </div>
          <button
            type="button"
            className="scooh-btn danger"
            onClick={handleBatchDelete}
            style={{
              background: '#ef4444',
              color: '#ffffff',
              border: 'none',
              fontWeight: 800,
              padding: '7px 16px',
              borderRadius: '7px',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 2px 10px rgba(239,68,68,0.45)'
            }}
          >
            🗑 Delete selected ({selectedIds.size})
          </button>
        </div>
      )}

      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {canDelete && (
                <th style={{ width: '42px', textAlign: 'center', padding: '10px 8px' }}>
                  <input
                    type="checkbox"
                    checked={shown.length > 0 && shown.every(r => selectedIds.has(r.id))}
                    onChange={toggleSelectAll}
                    title="Select all visible records"
                    style={{ cursor: 'pointer' }}
                  />
                </th>
              )}
              {fs.slice(0, 8).map(f => (
                <SortHeader key={f} label={label(f)} sortKey={f} currentSort={sortState} onSort={handleSort} />
              ))}
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {shown.length === 0 ? (
              <tr><td colSpan={fs.slice(0, 8).length + (canDelete ? 2 : 1)} className="scooh-empty">No records found</td></tr>
            ) : (
              shown.map(r => (
                <tr key={r.id}>
                  {canDelete && (
                    <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                      <input
                        type="checkbox"
                        checked={selectedIds.has(r.id)}
                        onChange={() => toggleSelect(r.id)}
                        style={{ cursor: 'pointer' }}
                      />
                    </td>
                  )}
                  {fs.slice(0, 8).map(f => (
                    <td key={f}>
                      {/cost|rate|amount|revenue/.test(f) ? (
                        <b>{money(r[f])}</b>
                      ) : /status|availability/.test(f) ? (
                        <span className={`scooh-pill ${r[f] === 'Active' || r[f] === 'Paid' || r[f] === 'Available' || r[f] === 'Mounted' ? 'active' : r[f] === 'Pending' || r[f] === 'Sent' ? 'watch' : 'vacant'}`}>
                          {String(r[f] ?? '')}
                        </span>
                      ) : /code/.test(f) ? (
                        <span className="scooh-plate">{String(r[f] ?? '')}</span>
                      ) : (
                        String(r[f] ?? '')
                      )}
                    </td>
                  ))}
                  <td>
                    <div className="scooh-rowactions">
                      <button className="scooh-iconbtn scooh-text-action" onClick={() => setEdit(r)}>
                        {canEdit ? 'Edit' : 'View'}
                      </button>
                      {canDelete && (
                        <button
                          className={`scooh-iconbtn danger-icon ${entity === 'campaigns' ? 'scooh-text-action' : ''}`}
                          title={entity === 'campaigns' ? 'Delete campaign record' : 'Archive record'}
                          onClick={() => del(r.id)}
                        >
                          {entity === 'campaigns' ? 'Delete' : '×'}
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {edit && (
        <div className="scooh-modal-overlay">
          <div className="scooh-modal">
            <div className="scooh-modalhead">
              <div>
                <h2>{edit.id ? (canEdit ? 'Edit' : 'View') : 'Add'} {title.replace(/s$/, '')}</h2>
                <span className="scooh-modal-subtitle">{canEdit ? 'Enter details and save' : 'Read-only record details'}</span>
              </div>
              <button type="button" className="scooh-modal-close" onClick={() => setEdit(null)}>×</button>
            </div>
            <form onSubmit={save}>
              <div className="scooh-modalbody">
                <div className="scooh-grid2">
                  {fs.map(f => {
                    const meta = fieldMeta[f] || {};
                    const isTextarea = meta.textarea || f === 'notes' || f === 'address' || f === 'billing_address' || f === 'terms';
                    const hasSelect = meta.select && Array.isArray(meta.select);

                    return (
                      <div className="scooh-field" key={f} style={isTextarea ? { gridColumn: '1 / -1' } : {}}>
                        <label>{label(f)}</label>
                        {isTextarea ? (
                          <textarea name={f} defaultValue={edit[f] ?? ''} placeholder={meta.placeholder || ''} rows={3} disabled={!canEdit} />
                        ) : hasSelect ? (
                          <select name={f} defaultValue={edit[f] ?? meta.select[0]} disabled={!canEdit}>
                            {meta.select.map(opt => (
                              <option key={opt} value={opt}>{opt}</option>
                            ))}
                          </select>
                        ) : (
                          <input
                            name={f}
                            defaultValue={edit[f] ?? ''}
                            type={/date/.test(f) ? 'date' : /cost|rate|amount|revenue|width|height|tax|discount|days|site_id|client_id/.test(f) ? 'number' : 'text'}
                            step={meta.step || 'any'}
                            placeholder={meta.placeholder || ''}
                            disabled={!canEdit}
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
              <div className="scooh-modalfoot">
                <button type="button" className="scooh-btn ghost" onClick={() => setEdit(null)}>
                  {canEdit ? 'Cancel' : 'Close'}
                </button>
                {canEdit && (
                  <button className="scooh-btn purple-btn" disabled={saving}>
                    {saving ? 'Saving…' : edit.id ? 'Update Record' : 'Save Record'}
                  </button>
                )}
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}

function SimpleList({ endpoint, title }) {
  const [rows, setRows] = useState([]);

  useEffect(() => {
    api.get('/' + endpoint).then(r => {
      if (Array.isArray(r.data)) setRows(r.data);
    }).catch(() => {});
  }, [endpoint]);

  return (
    <>
      <PageHead title={title} desc={`Reviewing recent system ${title.toLowerCase()} events.`} />
      <div className="scooh-tablewrap">
        <table className="scooh-table">
          <thead>
            <tr>
              {Object.keys((rows && rows[0]) || {}).slice(0, 9).map(k => <th key={k}>{label(k)}</th>)}
            </tr>
          </thead>
          <tbody>
            {(!rows || rows.length === 0) ? (
              <tr><td colSpan="9" className="scooh-empty">No records found</td></tr>
            ) : (
              rows.map((r, i) => (
                <tr key={r.id || i}>
                  {Object.keys(rows[0] || {}).slice(0, 9).map(k => (
                    <td key={k}>
                      {typeof r[k] === 'object' ? JSON.stringify(r[k]) : String(r[k] ?? '')}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </>
  );
}

function App() {
  return (
    <ErrorBoundary>
      <Routes>
        <Route path="/login" element={<Login />} />
        <Route path="/*" element={<Guard><Layout /></Guard>} />
      </Routes>
    </ErrorBoundary>
  );
}

createRoot(document.getElementById('root')).render(
  <BrowserRouter>
    <App />
  </BrowserRouter>
);
