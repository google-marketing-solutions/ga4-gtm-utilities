/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Readable names of GTM entity types (as shown in the GTM UI).
 */

/** Tag types. */
const TAG_TYPE_NAMES = {
  // Google tags (web).
  googtag: 'Google Tag',
  dest_ga: 'Google Analytics Destination',
  dest_aw: 'Google Ads Destination',
  dest_dc: 'Floodlight Destination',
  dest_mc: 'Merchant Center Destination',
  dest_other: 'Other Google Destination',
  gaawe: 'GA4 Event',
  gaawc: 'GA4 Configuration (legacy)',
  awct: 'Google Ads Conversion Tracking',
  sp: 'Google Ads Remarketing',
  awud: 'Google Ads User-Provided Data Event',
  awcc: 'Google Ads Calls from Website Conversion',
  gclidw: 'Conversion Linker',
  flc: 'Floodlight Counter',
  fls: 'Floodlight Sales',
  ua: 'Universal Analytics (legacy)',
  ga: 'Classic Google Analytics (legacy)',
  gcs: 'Google Surveys Website Satisfaction',
  ts: 'Google Trusted Stores',
  gfpa: 'Google Flights Price Accuracy',
  cloud_retail: 'Cloud Retail',
  zone: 'Zone',
  // Generic tags (web).
  html: 'Custom HTML',
  img: 'Custom Image',
  // Built-in third-party tags (web).
  bzi: 'LinkedIn Insight',
  pntr: 'Pinterest Tag',
  twitter_website_tag: 'Twitter Universal Website Tag',
  crto: 'Criteo OneTag',
  qcm: 'Quantcast Measure',
  qca: 'Quantcast Advertise',
  baut: 'Microsoft Advertising Universal Event Tracking',
  hjtc: 'Hotjar Tracking Code',
  asp: 'AdRoll Smart Pixel',
  cegg: 'Crazy Egg',
  mf: 'Mouseflow',
  qpx: 'Quora Pixel',
  awc: 'AWIN Conversion',
  awj: 'AWIN Journey',
  csm: 'comScore Unified Digital Measurement',
  ljs: 'Lytics JS Tag',
  // Server tags.
  sgtmgaaw: 'GA4 (Server)',
  sgtmadsct: 'Google Ads Conversion Tracking (Server)',
  sgtmadsremarket: 'Google Ads Remarketing (Server)',
  sgtmawud: 'Google Ads User-Provided Data Event (Server)',
  sgtmadsect: 'Google Ads Enhanced Conversions (Server, legacy)',
  sgtmadscl: 'Conversion Linker (Server)',
  sgtmflc: 'Floodlight Counter (Server)',
  sgtmfls: 'Floodlight Sales (Server)',
  sgtmua: 'Universal Analytics (Server, legacy)',
  sgtm_cloud_retail: 'Cloud Retail (Server)',
  sgtmmpfg: 'Measurement Protocol (Server)',
  http_request: 'HTTP Request (Server)'
};

/** Tag types that belong to Google products. */
const GOOGLE_TAG_TYPES = [
  'googtag', 'dest_ga', 'dest_aw', 'dest_dc', 'dest_mc', 'dest_other', 'gaawe',
  'gaawc', 'awct', 'sp', 'awud', 'awcc', 'gclidw', 'flc', 'fls', 'ua', 'ga',
  'gcs', 'ts', 'gfpa', 'cloud_retail', 'sgtmgaaw', 'sgtmadsct',
  'sgtmadsremarket', 'sgtmawud', 'sgtmadsect', 'sgtmadscl', 'sgtmflc',
  'sgtmfls', 'sgtmua', 'sgtm_cloud_retail', 'sgtmmpfg'
];

/** Variable types. */
const VARIABLE_TYPE_NAMES = {
  v: 'Data Layer Variable',
  c: 'Constant',
  jsm: 'Custom JavaScript',
  k: '1st-Party Cookie',
  u: 'URL',
  f: 'HTTP Referrer',
  j: 'JavaScript Variable',
  d: 'DOM Element',
  aev: 'Auto-Event Variable',
  vis: 'Element Visibility',
  r: 'Random Number',
  e: 'Custom Event',
  smm: 'Lookup Table',
  remm: 'RegEx Table',
  gtes: 'Google Tag: Event Settings',
  gtcs: 'Google Tag: Configuration Settings',
  awec: 'User-Provided Data',
  ctv: 'Container Version Number',
  dbg: 'Debug Mode',
  cid: 'Container ID',
  uv: 'Undefined Value',
  ev: 'Environment Name',
  gas: 'Google Analytics Settings (legacy)',
  gaupc: 'Google Analytics User Properties (legacy)',
  // Server variables.
  ed: 'Event Data',
  rh: 'Request Header',
  qp: 'Query Parameter',
  qs: 'Query String',
  rm: 'Request Method',
  rp: 'Request Path',
  cn: 'Client Name',
  sgtmk: 'Cookie Value',
  ip: 'Remote IP Address',
  vr: 'Visitor Region',
  fs: 'Firestore Lookup'
};

/** Trigger types. */
const TRIGGER_TYPE_NAMES = {
  pageview: 'Page View',
  domReady: 'DOM Ready',
  windowLoaded: 'Window Loaded',
  customEvent: 'Custom Event',
  triggerGroup: 'Trigger Group',
  init: 'Initialization',
  consentInit: 'Consent Initialization',
  formSubmission: 'Form Submission',
  click: 'Click - All Elements',
  linkClick: 'Click - Just Links',
  jsError: 'JavaScript Error',
  historyChange: 'History Change',
  timer: 'Timer',
  youTubeVideo: 'YouTube Video',
  scrollDepth: 'Scroll Depth',
  elementVisibility: 'Element Visibility',
  serverPageview: 'Page View (Server)',
  always: 'Custom (Server)',
  ampClick: 'AMP Click',
  ampTimer: 'AMP Timer',
  ampScroll: 'AMP Scroll',
  ampVisibility: 'AMP Visibility'
};

/** Client types (server containers). */
const CLIENT_TYPE_NAMES = {
  gaaw_client: 'Google Analytics: GA4 (Web)',
  gtm_client: 'Google Tag Manager: Web Container',
  mpaw_client: 'Measurement Protocol (GA4)',
  ga4f_client: 'Google Analytics: GA4 (App)',
  ua_client: 'Universal Analytics (legacy)',
  mpv1_client: 'Measurement Protocol (UA, legacy)'
};

/** Transformation types (server containers). */
const TRANSFORMATION_TYPE_NAMES = {
  tf_allow_params: 'Allow parameters',
  tf_exclude_params: 'Exclude parameters',
  tf_augment_event: 'Augment event'
};

/** Google tag setting types (gtag_config). */
const GTAG_CONFIG_TYPE_NAMES = {
  ogt_auto_events: 'Automatic event detection',
  ogt_cross_domain: 'Configure your domains',
  ogt_1p_data_v2: 'User-provided data collection',
  ogt_ip_mark: 'Define internal traffic',
  ogt_referral_exclusion: 'List unwanted referrals',
  ogt_session_timeout: 'Adjust session timeout',
  ogt_cookie_settings: 'Override cookie settings',
  ogt_dtc: 'Data transmission controls',
  ogt_consent_defaults: 'Consent defaults',
  ogt_dma: 'Digital Markets Act settings',
  ogt_data_extract: 'Data extraction',
  ogt_ga_send: 'Google Analytics sending',
  ogt_google_signals: 'Google signals (read-only)',
  ogt_ads_datatos: 'Ads data terms (read-only)',
  ogt_ga_datatos: 'Analytics data terms (read-only)',
  ogt_ga_gam_link: 'Ad Manager link (read-only)'
};

/**
 * Returns a readable name for a type code.
 * @param {!Object<string, string>} names Map of type code to name.
 * @param {string} type
 * @param {?WorkspaceData=} data Used for custom template names.
 * @return {string}
 */
function getTypeLabel(names, type, data) {
  if (!type) return '';
  if (names[type]) return names[type];
  if (/^cvt_/.test(type) && data) {
    const templateName = data.getTemplateName(type);
    if (templateName) return templateName;
  }
  return type;
}

/**
 * Resolves a type typed in a cell (readable name or type code) to a type code.
 * @param {!Object<string, string>} names Map of type code to name.
 * @param {*} value
 * @param {?WorkspaceData=} data Used for custom template names.
 * @return {string} The type code, or '' if the value is empty.
 */
function resolveTypeLabel(names, value, data) {
  const text = toText(value);
  if (!text) return '';
  if (names[text]) return text;
  const lower = text.toLowerCase();
  const code = Object.keys(names).find(key =>
    names[key].toLowerCase() === lower || key.toLowerCase() === lower);
  if (code) return code;
  if (data) {
    const templateType = data.getTemplateTypeByName(text);
    if (templateType) return templateType;
  }
  return text;
}
