#!/usr/bin/env node

/**
 * Playwright Click Reception Validator
 * 
 * Validates that ALL clickable elements successfully receive a click event.
 * - For LINKS: Must have href AND must navigate (or open new tab)
 * - For BUTTONS/UI elements: Click success only
 * - Supports custom elements (bolt-*, etc.)
 * - Supports Shadow DOM traversal
 * - SPA-friendly: Detects content changes even without URL change
 * - Generates HTML + Excel reports
 * - Batch processing of multiple URLs from CSV
 * 
 * Usage: 
 *   Single URL:  node validator.js https://example.com
 *   Batch mode:  node validator.js (reads urls.csv)
 *   With headed: node validator.js --headed
 * 
 * Install: npm install playwright xlsx
 */

const { chromium } = require('playwright');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

// ======================== CONFIGURATION ========================
const CONFIG = {
  retries: 2,
  retryDelayMs: 500,
  navigationTimeout: 60000,
  elementTimeout: 5000,
  screenshotOnFailure: true,
  reportsDir: './click-reports',
  viewport: { width: 1280, height: 720 },
  clickWaitMs: 3000,              // Increased for SPA navigation
  maxElementsPerUrl: 500,
  traverseShadowDOM: true,
  delayBetweenUrls: 3000,
  restoreWaitMs: 1500,
  networkIdleTimeout: 5000        // New: timeout for network idle
};

// ======================== UTILITY FUNCTIONS ========================

function cssEscape(value) {
  if (!value) return '';
  return String(value).replace(/([^a-zA-Z0-9\u00A0-\uFFFF_-])/g, '\\$1');
}

function cssAttrEscape(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/[&<>]/g, (m) => {
    if (m === '&') return '&amp;';
    if (m === '<') return '&lt;';
    if (m === '>') return '&gt;';
    return m;
  });
}

function generateSelector(elementInfo) {
  if (elementInfo.id) return `#${cssEscape(elementInfo.id)}`;
  if (elementInfo.testId) return `[data-testid="${elementInfo.testId}"]`;
  
  let selector = elementInfo.tagName.toLowerCase();
  if (elementInfo.className && elementInfo.className.trim()) {
    const classes = elementInfo.className
      .trim()
      .split(/\s+/)
      .filter(c => c && /^[a-zA-Z_][a-zA-Z0-9_-]*$/.test(c))
      .slice(0, 2);
    if (classes.length) selector += '.' + classes.join('.');
  }
  
  return selector;
}

// ======================== ENHANCED ELEMENT DISCOVERY ========================

async function discoverClickableElements(page) {
  return await page.evaluate((traverseShadowDOM) => {
    const clickableSelectors = [
      'a', 'button', 
      'input[type="button"]', 'input[type="submit"]', 'input[type="reset"]', 'input[type="image"]',
      '[role="button"]', '[role="link"]', '[role="menuitem"]', '[role="tab"]',
      '[onclick]', '[onmousedown]',
      '[tabindex]:not([tabindex="-1"])',
      'details > summary', 'area[href]',
      'bolt-button', 'bolt-link', 'bolt-icon[clickable]',
      '[is="bolt-button"]', '[is="bolt-link"]',
      '*[class*="bolt-button"]', '*[class*="bolt-link"]',
      'button[class*="bolt"]', 'a[class*="bolt"]',
      '*-button', '*-link', '*-clickable',
      '[part*="button"]', '[part*="link"]'
    ];
    
    const discovered = [];
    const seen = new Map();
    
    const getDomPath = (el) => {
      const path = [];
      let node = el;
      while (node && node !== document.body && node !== document.documentElement) {
        let selector = node.tagName.toLowerCase();
        if (node.id) {
          selector += `#${node.id}`;
          path.unshift(selector);
          break;
        }
        if (node.className && typeof node.className === 'string') {
          const classes = node.className.trim().split(/\s+/).slice(0, 2);
          if (classes.length) selector += '.' + classes.join('.');
        }
        const parent = node.parentNode;
        if (parent) {
          const siblings = Array.from(parent.children).filter(c => c.tagName === node.tagName);
          if (siblings.length > 1) {
            const index = siblings.indexOf(node) + 1;
            selector += `:nth-of-type(${index})`;
          }
        }
        path.unshift(selector);
        node = node.parentNode;
      }
      return path.join(' > ');
    };
    
    const traverseTree = (root) => {
      let elements = [];
      try {
        elements = Array.from(root.querySelectorAll(clickableSelectors.join(',')));
      } catch (e) {
        elements = Array.from(root.querySelectorAll('a, button, [role="button"], [onclick], bolt-button, [class*="bolt-button"]'));
      }
      
      for (const el of elements) {
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        const isVisible = style.display !== 'none' && 
                          style.visibility !== 'hidden' && 
                          style.opacity !== '0' &&
                          rect.width > 0 && 
                          rect.height > 0;
        
        const isDisabled = el.disabled === true || 
                          el.hasAttribute('disabled') || 
                          el.getAttribute('aria-disabled') === 'true';
        
        if (!isVisible || isDisabled) continue;
        
        let tagName = el.tagName.toLowerCase();
        let type = tagName;
        let text = '';
        
        if (tagName.includes('-')) {
          text = el.innerText || el.textContent || '';
          text = text.trim();
          if (!text) {
            text = el.getAttribute('aria-label') || el.getAttribute('title') || '';
          }
        } else if (tagName === 'input') {
          if (el.type === 'button' || el.type === 'submit' || el.type === 'reset') {
            text = el.value || '';
          }
        } else if (tagName === 'img' && el.alt) {
          text = el.alt;
        } else {
          text = (el.innerText || el.textContent || '').trim();
        }
        
        const ariaLabel = el.getAttribute('aria-label') || '';
        const title = el.getAttribute('title') || '';
        const displayText = (text || ariaLabel || title || tagName).substring(0, 200);
        
        let href = el.getAttribute('href');
        if (href === '' || href === null) {
          href = null;
        }
        
        const domPath = getDomPath(el);
        const positionKey = `${Math.round(rect.x / 10)}_${Math.round(rect.y / 10)}`;
        const fingerprint = `${tagName}::${displayText.substring(0, 80)}::${domPath.substring(0, 100)}::${positionKey}`;
        
        if (seen.has(fingerprint)) {
          continue;
        }
        
        const elementInfo = {
          tagName: tagName,
          type: type,
          text: displayText.substring(0, 100),
          fullText: text.substring(0, 300),
          ariaLabel: ariaLabel,
          id: el.id || '',
          name: el.name || '',
          className: el.className || '',
          href: href,
          role: el.getAttribute('role') || '',
          target: el.getAttribute('target') || '',
          testId: el.getAttribute('data-testid') || '',
          domPath: domPath,
          x: Math.round(rect.x),
          y: Math.round(rect.y),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
          isVisible: true,
          isCustomElement: tagName.includes('-'),
          hasMissingHref: tagName === 'a' && !href
        };
        
        seen.set(fingerprint, { element: elementInfo, count: 1 });
        discovered.push(elementInfo);
      }
      
      if (traverseShadowDOM) {
        const allElements = Array.from(root.querySelectorAll('*'));
        for (const el of allElements) {
          if (el.shadowRoot) {
            traverseTree(el.shadowRoot);
          }
        }
      }
    };
    
    traverseTree(document);
    
    discovered.sort((a, b) => {
      if (Math.abs(a.y - b.y) < 50) return a.x - b.x;
      return a.y - b.y;
    });
    
    return discovered;
  }, CONFIG.traverseShadowDOM);
}

// ======================== CLICK VALIDATION (FIXED FOR SPA) ========================

async function validateClickReception(page, context, elementInfo, originalUrl) {
  const startTime = Date.now();
  let screenshotPath = null;
  let clickSucceeded = false;
  let clickMethod = 'standard';
  let clickError = null;
  
  const displayText = (elementInfo.text || elementInfo.ariaLabel || elementInfo.id || elementInfo.tagName || '(no text)').substring(0, 80);
  const selector = generateSelector(elementInfo);
  
  const isValidHref = elementInfo.href && 
                      elementInfo.href !== '' &&
                      !elementInfo.href.startsWith('#') &&
                      !elementInfo.href.toLowerCase().startsWith('javascript:') &&
                      !elementInfo.href.startsWith('mailto:') &&
                      !elementInfo.href.startsWith('tel:');
  
  const isNavigationLink = elementInfo.tagName === 'a' && isValidHref;
  const isMissingHrefLink = elementInfo.tagName === 'a' && !elementInfo.href;
  const isAnchorLink = elementInfo.tagName === 'a' && elementInfo.href && elementInfo.href.startsWith('#');
  
  try {
    // Close any stray popups
    const pages = context.pages();
    for (const p of pages) {
      if (p !== page && !p.isClosed()) {
        await p.close().catch(() => {});
      }
    }
    
    // Wait for page to be stable
    await page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
    
    // SPECIAL CHECK: Missing href link - immediate fail
    if (isMissingHrefLink) {
      throw new Error(`❌ BROKEN LINK: Missing href attribute - "${displayText}" does nothing when clicked`);
    }
    
    let locator = null;
    let findAttempt = 0;
    const maxFindAttempts = 3;
    
    while (!locator && findAttempt < maxFindAttempts) {
      findAttempt++;
      
      // Strategy 1: By exact text
      if (elementInfo.text && elementInfo.text.length > 1) {
        try {
          let textLocator = page.getByText(elementInfo.text, { exact: true });
          let count = await textLocator.count().catch(() => 0);
          
          if (count === 0) {
            textLocator = page.getByText(elementInfo.text, { exact: false });
            count = await textLocator.count().catch(() => 0);
          }
          
          if (count > 0) {
            if (count > 1 && elementInfo.x && elementInfo.y) {
              for (let i = 0; i < Math.min(count, 10); i++) {
                const box = await textLocator.nth(i).boundingBox().catch(() => null);
                if (box && Math.abs(box.x - elementInfo.x) < 50 && Math.abs(box.y - elementInfo.y) < 50) {
                  locator = textLocator.nth(i);
                  break;
                }
              }
            }
            if (!locator) locator = textLocator.first();
            if (locator && await locator.isVisible({ timeout: 1000 }).catch(() => false)) break;
          }
        } catch { /* continue */ }
      }
      
      // Strategy 2: By href (for valid navigation links)
      if (isNavigationLink && elementInfo.href && !locator) {
        try {
          const hrefLocator = page.locator(`a[href="${cssAttrEscape(elementInfo.href)}"]`);
          if (await hrefLocator.first().isVisible({ timeout: 1000 }).catch(() => false)) {
            locator = hrefLocator.first();
            break;
          }
        } catch { /* continue */ }
      }
      
      // Strategy 3: For custom elements
      if (elementInfo.isCustomElement && !locator) {
        try {
          const customLocator = page.locator(elementInfo.tagName);
          if (await customLocator.first().isVisible({ timeout: 1000 }).catch(() => false)) {
            locator = customLocator.first();
            break;
          }
        } catch { /* continue */ }
      }
      
      // Strategy 4: By CSS selector
      if (!locator) {
        try {
          const cssLocator = page.locator(selector);
          if (await cssLocator.first().isVisible({ timeout: 1000 }).catch(() => false)) {
            locator = cssLocator.first();
            break;
          }
        } catch { /* continue */ }
      }
      
      if (!locator && findAttempt < maxFindAttempts) {
        await delay(800);
        if (findAttempt === 2) {
          await page.waitForLoadState('networkidle', { timeout: 5000 }).catch(() => {});
        }
      }
    }
    
    if (!locator) {
      throw new Error(`Element not found: "${displayText}"`);
    }
    
    // Get href from actual DOM element to verify
    const actualHref = await locator.getAttribute('href').catch(() => null);
    const actualTarget = await locator.getAttribute('target').catch(() => null);
    
    // Double-check: if this is an <a> tag with missing href, fail it
    const tagName = await locator.evaluate(el => el.tagName.toLowerCase()).catch(() => null);
    if (tagName === 'a' && (!actualHref || actualHref.trim() === '')) {
      throw new Error(`❌ BROKEN LINK: Missing href attribute - "${displayText}" does nothing when clicked`);
    }
    
    await locator.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => {});
    await delay(200);
    
    const isStillVisible = await locator.isVisible({ timeout: 1000 }).catch(() => false);
    if (!isStillVisible) {
      throw new Error(`Element became hidden before click: "${displayText}"`);
    }
    
    // ========== CAPTURE STATE BEFORE CLICK FOR SPA DETECTION ==========
    const urlBeforeClick = page.url();
    
    // Capture content signature for SPA without URL change
    const beforeContentSignature = await page.evaluate(() => {
      // Try multiple selectors for main content
      const contentSelectors = [
        'main', '[role="main"]', '.content', 'article', 
        '.main-content', '#main-content', '.page-content',
        '.container-fluid', '.bolt-container'
      ];
      
      let mainContent = null;
      for (const sel of contentSelectors) {
        mainContent = document.querySelector(sel);
        if (mainContent) break;
      }
      
      if (!mainContent) {
        // Fallback: get body but exclude header/footer
        const body = document.body;
        const header = document.querySelector('header, nav, [role="banner"]');
        const footer = document.querySelector('footer, [role="contentinfo"]');
        
        let content = body ? body.innerText : '';
        if (header) content = content.replace(header.innerText, '');
        if (footer) content = content.replace(footer.innerText, '');
        
        return `${content.length}:${content.length > 0 ? Math.floor(content.length / 100) : 0}`;
      }
      
      const text = mainContent.innerText || '';
      const linksCount = mainContent.querySelectorAll('a, button').length;
      const headingsCount = mainContent.querySelectorAll('h1, h2, h3').length;
      
      // Create a signature that changes when content updates
      return `${text.length}:${linksCount}:${headingsCount}`;
    }).catch(() => '');
    
    let popupOpened = false;
    let popupUrl = null;
    
    const isTargetBlank = actualTarget === '_blank' || elementInfo.target === '_blank';
    let popupPromise = null;
    if (isTargetBlank) {
      popupPromise = context.waitForEvent('page', { timeout: 8000 }).catch(() => null);
    }
    
    // Attempt click with retries
    for (let attempt = 1; attempt <= CONFIG.retries + 1; attempt++) {
      try {
        await locator.click({ timeout: CONFIG.elementTimeout });
        clickSucceeded = true;
        clickMethod = 'standard';
        break;
      } catch (clickErr) {
        clickError = clickErr.message;
        
        if (attempt === CONFIG.retries + 1) {
          try {
            await locator.click({ timeout: CONFIG.elementTimeout, force: true });
            clickSucceeded = true;
            clickMethod = 'force';
            break;
          } catch (forceErr) {
            try {
              await locator.evaluate(el => {
                el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                if (typeof el.click === 'function') el.click();
              });
              clickSucceeded = true;
              clickMethod = 'synthetic';
              break;
            } catch (syntheticErr) {
              throw new Error(`Click failed: ${clickErr.message}`);
            }
          }
        } else {
          await delay(CONFIG.retryDelayMs);
        }
      }
    }
    
    if (!clickSucceeded) {
      throw new Error(clickError || 'Click failed');
    }
    
    // ========== IMPROVED SPA NAVIGATION DETECTION ==========
    let navigationOccurred = false;
    let finalUrl = urlBeforeClick;
    let contentChanged = false;
    
    // Check for popup (new tab)
    if (popupPromise) {
      const popup = await popupPromise;
      if (popup && !popup.isClosed()) {
        popupOpened = true;
        await popup.waitForLoadState('domcontentloaded', { timeout: 10000 }).catch(() => {});
        popupUrl = popup.url();
        await popup.close().catch(() => {});
      }
    }
    
    // Check for same-window navigation (with SPA support)
    if (!popupOpened) {
      // Wait longer for SPA navigation
      await page.waitForTimeout(CONFIG.clickWaitMs);
      
      // Wait for network to be quiet (SPA data loading)
      await page.waitForLoadState('networkidle', { timeout: CONFIG.networkIdleTimeout }).catch(() => {});
      
      finalUrl = page.url();
      navigationOccurred = finalUrl !== urlBeforeClick;
      
      // If URL didn't change, check if content changed (SPA without URL change)
      if (!navigationOccurred && beforeContentSignature) {
        try {
          const afterContentSignature = await page.evaluate(() => {
            const contentSelectors = [
              'main', '[role="main"]', '.content', 'article', 
              '.main-content', '#main-content', '.page-content',
              '.container-fluid', '.bolt-container'
            ];
            
            let mainContent = null;
            for (const sel of contentSelectors) {
              mainContent = document.querySelector(sel);
              if (mainContent) break;
            }
            
            if (!mainContent) {
              const body = document.body;
              const header = document.querySelector('header, nav, [role="banner"]');
              const footer = document.querySelector('footer, [role="contentinfo"]');
              
              let content = body ? body.innerText : '';
              if (header) content = content.replace(header.innerText, '');
              if (footer) content = content.replace(footer.innerText, '');
              
              return `${content.length}:${content.length > 0 ? Math.floor(content.length / 100) : 0}`;
            }
            
            const text = mainContent.innerText || '';
            const linksCount = mainContent.querySelectorAll('a, button').length;
            const headingsCount = mainContent.querySelectorAll('h1, h2, h3').length;
            
            return `${text.length}:${linksCount}:${headingsCount}`;
          }).catch(() => '');
          
          contentChanged = beforeContentSignature !== afterContentSignature;
          
          if (contentChanged) {
            navigationOccurred = true;
          }
        } catch { /* ignore */ }
      }
    }
    
    // ========== VALIDATION LOGIC ==========
    
    // For navigation links (with valid href) - MUST navigate or open popup
    if (isNavigationLink && !isAnchorLink) {
      if (isTargetBlank) {
        // Should open new tab
        if (popupOpened && popupUrl && popupUrl !== 'about:blank') {
          return {
            success: true,
            status: 'pass',
            elementType: elementInfo.type,
            elementText: displayText,
            selector: selector,
            clickMethod: clickMethod,
            screenshotPath: null,
            errorMessage: null,
            durationMs: Date.now() - startTime,
            timestamp: new Date().toISOString(),
            domPath: elementInfo.domPath,
            isCustomElement: elementInfo.isCustomElement,
            href: elementInfo.href,
            validationDetail: `✅ New tab opened: ${popupUrl.substring(0, 80)}`
          };
        } else {
          throw new Error(`❌ BROKEN LINK: target="_blank" did not open a new tab: "${displayText}"`);
        }
      } else {
        // Should navigate in same window
        if (navigationOccurred) {
          const destUrl = finalUrl !== urlBeforeClick ? finalUrl : '(SPA content updated)';
          return {
            success: true,
            status: 'pass',
            elementType: elementInfo.type,
            elementText: displayText,
            selector: selector,
            clickMethod: clickMethod,
            screenshotPath: null,
            errorMessage: null,
            durationMs: Date.now() - startTime,
            timestamp: new Date().toISOString(),
            domPath: elementInfo.domPath,
            isCustomElement: elementInfo.isCustomElement,
            href: elementInfo.href,
            validationDetail: `✅ Navigated to: ${destUrl.substring(0, 80)}`
          };
        } else {
          throw new Error(`❌ BROKEN LINK: Navigation link did not navigate: "${displayText}" (href: ${elementInfo.href || 'missing'})`);
        }
      }
    }
    
    // For anchor links (#) - just need click success
    if (isAnchorLink) {
      return {
        success: true,
        status: 'pass',
        elementType: elementInfo.type,
        elementText: displayText,
        selector: selector,
        clickMethod: clickMethod,
        screenshotPath: null,
        errorMessage: null,
        durationMs: Date.now() - startTime,
        timestamp: new Date().toISOString(),
        domPath: elementInfo.domPath,
        isCustomElement: elementInfo.isCustomElement,
        href: elementInfo.href,
        validationDetail: '✅ Anchor link clicked (page scroll)'
      };
    }
    
    // For non-link elements (buttons, divs with role="button", etc.) - click success is enough
    return {
      success: true,
      status: 'pass',
      elementType: elementInfo.type,
      elementText: displayText,
      selector: selector,
      clickMethod: clickMethod,
      screenshotPath: null,
      errorMessage: null,
      durationMs: Date.now() - startTime,
      timestamp: new Date().toISOString(),
      domPath: elementInfo.domPath,
      isCustomElement: elementInfo.isCustomElement,
      href: elementInfo.href,
      validationDetail: '✅ Click received (UI element)'
    };
    
  } catch (error) {
    // Take screenshot for failures (except navigation interruptions)
    const isNavInterruption = error.message && error.message.includes('interrupted by another navigation');
    const isMissingHref = error.message && error.message.includes('Missing href') || error.message.includes('BROKEN LINK');
    
    if (CONFIG.screenshotOnFailure && !isNavInterruption) {
      try {
        const screenshotDir = path.join(CONFIG.reportsDir, 'screenshots');
        await fs.mkdir(screenshotDir, { recursive: true });
        const hash = crypto.createHash('md5').update(`${selector}_${Date.now()}`).digest('hex').substring(0, 8);
        screenshotPath = path.join(screenshotDir, `failure_${hash}.png`);
        await page.screenshot({ path: screenshotPath, fullPage: false });
      } catch { /* ignore */ }
    }
    
    let errorMessage = error.message || String(error);
    
    // Enhance error message for broken links
    if ((isMissingHrefLink || errorMessage.includes('no href')) && elementInfo.tagName === 'a') {
      errorMessage = `❌ BROKEN LINK: Missing href attribute - "${displayText}" does nothing when clicked`;
    }
    
    return {
      success: false,
      status: 'fail',
      elementType: elementInfo.type,
      elementText: displayText,
      selector: selector,
      errorMessage: errorMessage,
      clickMethod: clickMethod,
      screenshotPath: screenshotPath,
      durationMs: Date.now() - startTime,
      timestamp: new Date().toISOString(),
      domPath: elementInfo.domPath,
      isCustomElement: elementInfo.isCustomElement,
      href: elementInfo.href,
      validationDetail: '❌ FAILED'
    };
  }
}

// ======================== PAGE STATE HELPERS ========================

async function dismissCookieBanner(page) {
  const dismissSelectors = [
    'button:has-text("Accept")',
    'button:has-text("Accept All")',
    'button:has-text("Accept Cookies")',
    'button:has-text("Allow All")',
    'button:has-text("I Accept")',
    'button:has-text("Agree")',
    'button:has-text("OK")',
    'button:has-text("Close")',
    '#onetrust-accept-btn-handler',
    '.onetrust-accept-btn-handler',
    '[aria-label*="Accept"]',
    '[aria-label*="Close"]',
    'bolt-button:has-text("Accept")'
  ];

  for (const sel of dismissSelectors) {
    try {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 1500 })) {
        await btn.click({ timeout: 3000, force: true });
        await page.waitForTimeout(500);
        return true;
      }
    } catch { /* try next */ }
  }
  return false;
}

async function restorePageState(page, originalUrl) {
  if (page.url() === originalUrl) {
    await page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
    await delay(CONFIG.restoreWaitMs);
    return;
  }
  
  await page.waitForLoadState('networkidle', { timeout: 5000 }).catch(() => {});
  
  const wentBack = await page.goBack({ waitUntil: 'domcontentloaded', timeout: 10000 })
    .then(() => page.url() === originalUrl)
    .catch(() => false);
  
  if (!wentBack || page.url() !== originalUrl) {
    await page.goto(originalUrl, { 
      waitUntil: 'domcontentloaded', 
      timeout: CONFIG.navigationTimeout 
    }).catch(() => {});
  }
  
  await page.waitForLoadState('domcontentloaded', { timeout: 5000 }).catch(() => {});
  await delay(CONFIG.restoreWaitMs);
  await dismissCookieBanner(page);
}

// ======================== SINGLE URL VALIDATION ========================

async function runValidation(targetUrl, options = {}) {
  const headed = options.headed || false;
  const startTime = Date.now();
  let browser = null;
  let context = null;
  let page = null;
  const allResults = [];
  
  try {
    await fs.mkdir(CONFIG.reportsDir, { recursive: true });
    
    console.log(`  🌐 Launching browser...`);
    browser = await chromium.launch({ headless: !headed });
    context = await browser.newContext({
      viewport: CONFIG.viewport,
      acceptDownloads: false
    });
    page = await context.newPage();
    
    console.log(`  📄 Loading ${targetUrl}...`);
    await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: CONFIG.navigationTimeout });
    await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);
    const originalUrl = page.url();
    console.log(`  ✅ Page loaded: ${originalUrl}`);
    
    await dismissCookieBanner(page);
    
    console.log(`  🔍 Discovering clickable elements...`);
    const elements = await discoverClickableElements(page);
    console.log(`  📊 Found ${elements.length} clickable elements`);
    
    // Count broken links (missing href)
    const missingHrefLinks = elements.filter(e => e.hasMissingHref);
    if (missingHrefLinks.length > 0) {
      console.log(`  ⚠️ Detected ${missingHrefLinks.length} link(s) with missing href attribute (will FAIL)`);
    }
    
    const elementsToTest = elements.slice(0, CONFIG.maxElementsPerUrl);
    const customCount = elementsToTest.filter(e => e.isCustomElement).length;
    if (customCount > 0) {
      console.log(`  ⚡ Detected ${customCount} custom element(s) (bolt-*, etc.)`);
    }
    
    console.log(`  🖱️ Testing click reception...\n`);
    
    let processedCount = 0;
    
    for (let i = 0; i < elementsToTest.length; i++) {
      const element = elementsToTest[i];
      processedCount++;
      const percent = ((processedCount / elementsToTest.length) * 100).toFixed(1);
      const elementLabel = (element.text || element.ariaLabel || element.id || element.tagName || '(no text)').substring(0, 40);
      const customTag = element.isCustomElement ? ' ⚡' : '';
      const brokenTag = element.hasMissingHref ? ' 💔' : '';
      
      process.stdout.write(`    [${processedCount}/${elementsToTest.length}] (${percent}%) Testing: ${element.type} - "${elementLabel}"${customTag}${brokenTag}... `);
      
      try {
        await restorePageState(page, originalUrl);
        
        const result = await validateClickReception(page, context, element, originalUrl);
        allResults.push(result);
        
        const statusIcon = result.status === 'pass' ? '✅' : '❌';
        console.log(`${statusIcon} ${result.status.toUpperCase()} (${result.durationMs}ms)`);
        
        if (result.status === 'fail' && result.errorMessage) {
          const shortMsg = result.errorMessage.substring(0, 100);
          console.log(`       ⚠️ ${shortMsg}`);
        }
        
        await delay(200);
        
      } catch (err) {
        console.log(`❌ ERROR: ${err.message.substring(0, 80)}`);
        allResults.push({
          status: 'fail',
          elementType: element.type,
          elementText: element.text || '(no text)',
          selector: generateSelector(element),
          errorMessage: `Unexpected error: ${err.message}`,
          durationMs: 0,
          timestamp: new Date().toISOString(),
          domPath: element.domPath,
          isCustomElement: element.isCustomElement,
          href: element.href,
          validationDetail: '❌ ERROR'
        });
      }
    }
    
    const endTime = Date.now();
    const passed = allResults.filter(r => r.status === 'pass').length;
    const failed = allResults.filter(r => r.status === 'fail').length;
    const brokenLinksFound = allResults.filter(r => r.errorMessage && r.errorMessage.includes('BROKEN LINK')).length;
    
    console.log(`\n  📈 Summary: ${passed} passed, ${failed} failed (${((passed / allResults.length) * 100).toFixed(1)}% success rate)`);
    if (brokenLinksFound > 0) {
      console.log(`  💔 Broken links detected: ${brokenLinksFound}`);
    }
    
    return { 
      success: true, 
      results: allResults, 
      stats: { total: allResults.length, passed, failed, passRate: (passed / allResults.length) * 100, brokenLinks: brokenLinksFound },
      url: targetUrl,
      originalUrl: originalUrl,
      duration: endTime - startTime
    };
    
  } catch (error) {
    console.error(`  💥 Error: ${error.message}`);
    return { success: false, error: error.message };
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

// ======================== EXCEL REPORT GENERATION ========================

function generateExcelReport(allUrlResults, startTime, endTime) {
  const workbook = XLSX.utils.book_new();
  
  let totalElements = 0, totalPassed = 0, totalFailed = 0, totalBrokenLinks = 0;
  for (const urlResult of allUrlResults) {
    if (urlResult.stats) {
      totalElements += urlResult.stats.total;
      totalPassed += urlResult.stats.passed;
      totalFailed += urlResult.stats.failed;
      totalBrokenLinks += urlResult.stats.brokenLinks || 0;
    }
  }
  
  // Summary Sheet
  const summaryData = [
    ['Click Reception Validation Report'],
    ['Generated:', new Date().toLocaleString()],
    ['Total Duration:', `${((endTime - startTime) / 1000).toFixed(1)}s`],
    [],
    ['Overall Statistics', '', ''],
    ['URLs Tested', allUrlResults.length, ''],
    ['Total Clickable Elements', totalElements, ''],
    ['✅ Passed (Click Received)', totalPassed, ''],
    ['❌ Failed (Click Failed)', totalFailed, ''],
    ['💔 Broken Links (Missing href)', totalBrokenLinks, ''],
    ['Overall Success Rate', totalElements > 0 ? `${((totalPassed / totalElements) * 100).toFixed(1)}%` : 'N/A', ''],
    [],
    ['Per-URL Summary', '', '', '', ''],
    ['URL', 'Total Elements', 'Passed', 'Failed', 'Broken Links', 'Success Rate', 'Duration (s)']
  ];
  
  for (const urlResult of allUrlResults) {
    summaryData.push([
      urlResult.url,
      urlResult.stats.total,
      urlResult.stats.passed,
      urlResult.stats.failed,
      urlResult.stats.brokenLinks || 0,
      `${urlResult.stats.passRate.toFixed(1)}%`,
      (urlResult.duration / 1000).toFixed(1)
    ]);
  }
  
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  summarySheet['!cols'] = [{wch:60}, {wch:15}, {wch:15}, {wch:15}, {wch:15}, {wch:15}, {wch:15}];
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
  
  // Detailed Results Sheet
  const detailedData = [
    ['Status', 'URL', 'Element Type', 'Element Text', 'Selector', 'Click Method', 'Duration (ms)', 'Href', 'Validation Detail', 'Error Message', 'DOM Path', 'Custom Element', 'Timestamp']
  ];
  
  for (const urlResult of allUrlResults) {
    for (const result of urlResult.results) {
      detailedData.push([
        result.status.toUpperCase(),
        urlResult.url,
        result.elementType,
        result.elementText,
        result.selector,
        result.clickMethod || '-',
        result.durationMs || 0,
        result.href || '(missing)',
        result.validationDetail || '',
        result.errorMessage || '',
        result.domPath || '',
        result.isCustomElement ? 'Yes' : 'No',
        result.timestamp
      ]);
    }
  }
  
  const detailedSheet = XLSX.utils.aoa_to_sheet(detailedData);
  detailedSheet['!cols'] = [{wch:8}, {wch:55}, {wch:15}, {wch:35}, {wch:30}, {wch:12}, {wch:12}, {wch:20}, {wch:40}, {wch:50}, {wch:50}, {wch:12}, {wch:22}];
  XLSX.utils.book_append_sheet(workbook, detailedSheet, 'Detailed Results');
  
  // Failures Only Sheet
  const failuresData = [
    ['Status', 'URL', 'Element Type', 'Element Text', 'Selector', 'Href', 'Error Message', 'Duration (ms)']
  ];
  
  for (const urlResult of allUrlResults) {
    for (const result of urlResult.results) {
      if (result.status === 'fail') {
        failuresData.push([
          result.status.toUpperCase(),
          urlResult.url,
          result.elementType,
          result.elementText,
          result.selector,
          result.href || '(missing)',
          result.errorMessage || '',
          result.durationMs || 0
        ]);
      }
    }
  }
  
  const failuresSheet = XLSX.utils.aoa_to_sheet(failuresData);
  failuresSheet['!cols'] = [{wch:8}, {wch:55}, {wch:15}, {wch:35}, {wch:30}, {wch:20}, {wch:50}, {wch:12}];
  XLSX.utils.book_append_sheet(workbook, failuresSheet, 'Failures Only');
  
  // Broken Links Sheet
  const brokenLinksData = [
    ['URL', 'Element Text', 'Selector', 'Href Status', 'Error Message']
  ];
  
  for (const urlResult of allUrlResults) {
    for (const result of urlResult.results) {
      if (result.errorMessage && result.errorMessage.includes('BROKEN LINK')) {
        brokenLinksData.push([
          urlResult.url,
          result.elementText,
          result.selector,
          result.href ? `has href: ${result.href}` : 'MISSING HREF',
          result.errorMessage
        ]);
      }
    }
  }
  
  if (brokenLinksData.length > 1) {
    const brokenLinksSheet = XLSX.utils.aoa_to_sheet(brokenLinksData);
    brokenLinksSheet['!cols'] = [{wch:55}, {wch:35}, {wch:30}, {wch:25}, {wch:50}];
    XLSX.utils.book_append_sheet(workbook, brokenLinksSheet, 'Broken Links');
  }
  
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const excelPath = path.join(CONFIG.reportsDir, `click_reception_report_${timestamp}.xlsx`);
  XLSX.writeFile(workbook, excelPath);
  
  return excelPath;
}

// ======================== BEAUTIFUL HTML REPORT ========================

async function generateBeautifulHtmlReport(allUrlResults, startTime, endTime) {
  const totalUrls = allUrlResults.length;
  let totalElements = 0;
  let totalPassed = 0;
  let totalFailed = 0;
  let totalBrokenLinks = 0;
  
  for (const urlResult of allUrlResults) {
    if (urlResult.stats) {
      totalElements += urlResult.stats.total;
      totalPassed += urlResult.stats.passed;
      totalFailed += urlResult.stats.failed;
      totalBrokenLinks += urlResult.stats.brokenLinks || 0;
    }
  }
  
  const overallPassRate = totalElements > 0 ? ((totalPassed / totalElements) * 100).toFixed(1) : 0;
  
  let urlSectionsHtml = '';
  
  for (const urlResult of allUrlResults) {
    const urlShort = urlResult.url.length > 80 ? urlResult.url.substring(0, 77) + '...' : urlResult.url;
    const successRate = urlResult.stats.total > 0 ? ((urlResult.stats.passed / urlResult.stats.total) * 100).toFixed(1) : 0;
    const successColor = successRate >= 90 ? '#10b981' : (successRate >= 70 ? '#f59e0b' : '#ef4444');
    
    let tableRows = '';
    for (const result of urlResult.results) {
      const statusClass = result.status;
      const statusIcon = result.status === 'pass' ? '✅' : '❌';
      const elementTypeClass = result.isCustomElement ? 'custom-element' : '';
      const isBrokenLink = result.errorMessage && result.errorMessage.includes('BROKEN LINK');
      const errorDisplay = result.errorMessage ? 
        `<div class="error-tooltip ${isBrokenLink ? 'broken-link' : ''}" title="${escapeHtml(result.errorMessage)}">${isBrokenLink ? '💔 ' : '⚠️ '}${escapeHtml(result.errorMessage.substring(0, 60))}${result.errorMessage.length > 60 ? '…' : ''}</div>` : 
        '<span class="no-error">—</span>';
      
      const validationDisplay = result.validationDetail ? 
        `<div class="validation-detail" title="${escapeHtml(result.validationDetail)}">📌 ${escapeHtml(result.validationDetail.substring(0, 50))}${result.validationDetail.length > 50 ? '…' : ''}</div>` : '';
      
      tableRows += `
        <tr class="${statusClass}-row">
          <td class="status-cell"><span class="status-badge ${statusClass}">${statusIcon} ${result.status.toUpperCase()}</span></td>
          <td><span class="element-type ${elementTypeClass}">${escapeHtml(result.elementType || '?')}</span></td>
          <td class="text-cell" title="${escapeHtml(result.elementText || '')}">${escapeHtml((result.elementText || '-').substring(0, 50))}${(result.elementText || '').length > 50 ? '…' : ''}</td>
          <td class="method-cell">${result.clickMethod || '-'}</td>
          <td class="duration-cell">${result.durationMs || 0}ms</td>
          <td class="error-cell">${errorDisplay}${validationDisplay}</td>
        </tr>
      `;
    }
    
    urlSectionsHtml += `
      <div class="url-section">
        <div class="url-header" onclick="toggleSection(this)">
          <div class="url-title">
            <span class="toggle-icon">▶</span>
            <span class="url-icon">🌐</span>
            <span class="url-text">${escapeHtml(urlShort)}</span>
          </div>
          <div class="url-stats">
            <span class="stat-badge total">📊 ${urlResult.stats.total}</span>
            <span class="stat-badge passed">✅ ${urlResult.stats.passed}</span>
            <span class="stat-badge failed">❌ ${urlResult.stats.failed}</span>
            ${urlResult.stats.brokenLinks > 0 ? `<span class="stat-badge broken">💔 ${urlResult.stats.brokenLinks}</span>` : ''}
            <span class="stat-badge rate" style="background: ${successColor}20; color: ${successColor}">📈 ${successRate}%</span>
            <span class="stat-badge duration">⏱️ ${(urlResult.duration / 1000).toFixed(1)}s</span>
          </div>
        </div>
        <div class="url-content" style="display: none;">
          <div class="url-detail">
            <div class="detail-item"><strong>Full URL:</strong> <a href="${escapeHtml(urlResult.url)}" target="_blank">${escapeHtml(urlResult.url)}</a></div>
            <div class="detail-item"><strong>Original URL:</strong> ${escapeHtml(urlResult.originalUrl || urlResult.url)}</div>
            <div class="detail-item"><strong>Test completed:</strong> ${new Date(urlResult.results[0]?.timestamp || Date.now()).toLocaleString()}</div>
          </div>
          <div class="table-wrapper">
            <table class="results-table">
              <thead>
                <tr><th>Status</th><th>Type</th><th>Element Text</th><th>Click Method</th><th>Duration</th><th>Error / Details</th></tr>
              </thead>
              <tbody>${tableRows}</tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }
  
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Click Reception Validation Report</title>
    <style>
        * { box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 24px; background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%); color: #1a202c; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #1e3c72 100%); color: white; padding: 32px; border-radius: 20px; margin-bottom: 28px; }
        .header h1 { margin: 0 0 12px 0; font-size: 32px; display: flex; align-items: center; gap: 12px; }
        .header .subtitle { font-size: 14px; opacity: 0.85; margin-top: 8px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 28px; }
        .stat-card { background: white; border-radius: 20px; padding: 24px 20px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.05); transition: transform 0.2s; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); }
        .stat-value { font-size: 42px; font-weight: 800; margin-bottom: 8px; }
        .stat-label { color: #64748b; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
        .stat-pass .stat-value { color: #10b981; }
        .stat-fail .stat-value { color: #ef4444; }
        .stat-broken .stat-value { color: #f59e0b; }
        .stat-rate .stat-value { color: #3b82f6; }
        .url-section { background: white; border-radius: 16px; margin-bottom: 20px; overflow: hidden; border: 1px solid #e2e8f0; }
        .url-header { display: flex; justify-content: space-between; align-items: center; padding: 16px 20px; background: #f8fafc; cursor: pointer; border-bottom: 1px solid #e2e8f0; }
        .url-header:hover { background: #f1f5f9; }
        .url-title { display: flex; align-items: center; gap: 12px; font-weight: 600; }
        .toggle-icon { font-size: 12px; color: #64748b; transition: transform 0.2s; }
        .url-section.expanded .toggle-icon { transform: rotate(90deg); }
        .url-text { font-family: monospace; font-size: 13px; color: #1e293b; }
        .url-stats { display: flex; gap: 12px; flex-wrap: wrap; }
        .stat-badge { padding: 4px 12px; border-radius: 30px; font-size: 12px; font-weight: 600; }
        .stat-badge.total { background: #e2e8f0; color: #475569; }
        .stat-badge.passed { background: #d1fae5; color: #065f46; }
        .stat-badge.failed { background: #fee2e2; color: #991b1b; }
        .stat-badge.broken { background: #fef3c7; color: #92400e; }
        .url-detail { padding: 20px; background: #fefce8; border-bottom: 1px solid #e2e8f0; font-size: 13px; }
        .detail-item { margin-bottom: 8px; }
        .detail-item a { color: #3b82f6; text-decoration: none; }
        .table-wrapper { overflow-x: auto; }
        .results-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .results-table th { background: #f1f5f9; padding: 12px; text-align: left; font-weight: 600; border-bottom: 2px solid #e2e8f0; }
        .results-table td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }
        .results-table tr:hover { background: #f8fafc; }
        .status-badge { display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px; border-radius: 30px; font-size: 11px; font-weight: 700; }
        .status-badge.pass { background: #d1fae5; color: #065f46; }
        .status-badge.fail { background: #fee2e2; color: #991b1b; }
        .element-type { font-family: monospace; font-size: 11px; background: #f1f5f9; padding: 4px 8px; border-radius: 8px; display: inline-block; }
        .element-type.custom-element { background: #e0e7ff; color: #3730a3; }
        .text-cell { max-width: 250px; word-break: break-word; }
        .error-tooltip { color: #dc2626; font-size: 11px; cursor: help; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 250px; }
        .error-tooltip.broken-link { color: #f59e0b; }
        .validation-detail { font-size: 10px; color: #10b981; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .no-error { color: #94a3b8; }
        .footer { margin-top: 32px; text-align: center; font-size: 12px; color: #94a3b8; padding: 20px; border-top: 1px solid #e2e8f0; }
        .scroll-top { position: fixed; bottom: 30px; right: 30px; background: #1e3c72; color: white; width: 48px; height: 48px; border-radius: 50%; display: flex; align-items: center; justify-content: center; cursor: pointer; opacity: 0; transition: opacity 0.3s; border: none; font-size: 20px; }
        .scroll-top.visible { opacity: 1; }
        @media (max-width: 768px) { body { padding: 12px; } .url-header { flex-direction: column; align-items: flex-start; gap: 12px; } .text-cell { max-width: 150px; } }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1><span>🖱️</span> Click Reception Validation Report</h1>
        <div class="subtitle">Validating clickable elements: Links MUST navigate, UI elements only need click success</div>
        <div class="subtitle">📅 ${new Date(startTime).toLocaleString()} → ${new Date(endTime).toLocaleString()} (${((endTime - startTime) / 1000).toFixed(1)}s total)</div>
    </div>
    
    <div class="stats-grid">
        <div class="stat-card"><div class="stat-value">${totalUrls}</div><div class="stat-label">🌐 URLs Tested</div></div>
        <div class="stat-card"><div class="stat-value">${totalElements}</div><div class="stat-label">Clickable Elements</div></div>
        <div class="stat-card stat-pass"><div class="stat-value">${totalPassed}</div><div class="stat-label">✅ Click Received</div></div>
        <div class="stat-card stat-fail"><div class="stat-value">${totalFailed}</div><div class="stat-label">❌ Click Failed</div></div>
        ${totalBrokenLinks > 0 ? `<div class="stat-card"><div class="stat-value">${totalBrokenLinks}</div><div class="stat-label">💔 Broken Links</div></div>` : ''}
        <div class="stat-card stat-rate"><div class="stat-value">${overallPassRate}%</div><div class="stat-label">Success Rate</div></div>
    </div>
    
    <div style="margin-bottom: 20px;">
        <button onclick="expandAll()" style="background: #e2e8f0; border: none; padding: 8px 16px; border-radius: 8px; margin-right: 10px; cursor: pointer;">📂 Expand All</button>
        <button onclick="collapseAll()" style="background: #e2e8f0; border: none; padding: 8px 16px; border-radius: 8px; cursor: pointer;">📁 Collapse All</button>
    </div>
    
    ${urlSectionsHtml}
    
    <div class="footer">
        <p>🔍 Click Reception Validator | Shadow DOM: Enabled | Custom Elements: Supported | SPA Navigation: Detected via content change</p>
        <p>✅ <strong>Pass</strong> = Click success (for links: navigation or content change required) | ❌ <strong>Fail</strong> = Click failed or broken link (missing href)</p>
        <p>Report generated: ${new Date().toISOString()}</p>
    </div>
</div>

<button class="scroll-top" onclick="scrollToTop()" id="scrollTopBtn">↑</button>

<script>
    function toggleSection(header) {
        const section = header.closest('.url-section');
        section.classList.toggle('expanded');
        const content = section.querySelector('.url-content');
        const icon = header.querySelector('.toggle-icon');
        if (section.classList.contains('expanded')) {
            content.style.display = 'block';
            icon.textContent = '▼';
        } else {
            content.style.display = 'none';
            icon.textContent = '▶';
        }
    }
    function expandAll() { document.querySelectorAll('.url-section').forEach(s => { s.classList.add('expanded'); const c = s.querySelector('.url-content'); const i = s.querySelector('.toggle-icon'); if(c) c.style.display = 'block'; if(i) i.textContent = '▼'; }); }
    function collapseAll() { document.querySelectorAll('.url-section').forEach(s => { s.classList.remove('expanded'); const c = s.querySelector('.url-content'); const i = s.querySelector('.toggle-icon'); if(c) c.style.display = 'none'; if(i) i.textContent = '▶'; }); }
    window.addEventListener('scroll', function() { document.getElementById('scrollTopBtn').classList.toggle('visible', window.scrollY > 300); });
    function scrollToTop() { window.scrollTo({ top: 0, behavior: 'smooth' }); }
    const firstSection = document.querySelector('.url-section');
    if (firstSection) { firstSection.classList.add('expanded'); const c = firstSection.querySelector('.url-content'); const i = firstSection.querySelector('.toggle-icon'); if(c) c.style.display = 'block'; if(i) i.textContent = '▼'; }
</script>
</body>
</html>`;

  const reportPath = path.join(CONFIG.reportsDir, `click_reception_report_${new Date().toISOString().replace(/[:.]/g, '-')}.html`);
  await fs.writeFile(reportPath, html);
  return reportPath;
}

// ======================== CSV PARSING ========================

function parseCSV(csvPath) {
  try {
    const content = fsSync.readFileSync(csvPath, 'utf-8');
    const lines = content.split('\n').filter(line => line.trim() && !line.startsWith('#'));
    
    if (lines.length < 2) throw new Error('CSV file has no data rows');
    
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
    const urlIndex = headers.findIndex(h => h === 'url');
    
    if (urlIndex === -1) throw new Error('CSV must have a "url" column');
    
    const urls = [];
    for (let i = 1; i < lines.length; i++) {
      const values = lines[i].split(',').map(v => v.trim());
      if (values[urlIndex] && (values[urlIndex].startsWith('http://') || values[urlIndex].startsWith('https://'))) {
        urls.push({ url: values[urlIndex], index: i });
      }
    }
    return urls;
  } catch (error) {
    throw new Error(`Failed to parse CSV: ${error.message}`);
  }
}

// ======================== MAIN ENTRY POINT ========================

async function main() {
  const args = process.argv.slice(2);
  const headed = args.includes('--headed');
  const singleUrl = args.find(arg => arg.startsWith('http://') || arg.startsWith('https://'));
  
  if (headed) console.log('👁 Running in headed mode (browser visible)\n');
  
  // Single URL mode
  if (singleUrl) {
    console.log(`🎯 Validating single URL: ${singleUrl}\n`);
    const result = await runValidation(singleUrl, { headed });
    if (!result.success) process.exit(1);
    
    const allUrlResults = [result];
    const overallStartTime = Date.now() - result.duration;
    const overallEndTime = Date.now();
    
    const htmlPath = await generateBeautifulHtmlReport(allUrlResults, overallStartTime, overallEndTime);
    const excelPath = generateExcelReport(allUrlResults, overallStartTime, overallEndTime);
    
    console.log(`\n📄 HTML Report: ${htmlPath}`);
    console.log(`📊 Excel Report: ${excelPath}`);
    console.log(`\n✅ Done!`);
    process.exit(0);
  }
  
  // CSV batch mode
  const csvPath = path.join(__dirname, 'urls.csv');
  
  if (!fsSync.existsSync(csvPath)) {
    console.error('❌ Error: urls.csv file not found.');
    console.error('Usage: node validator.js (reads urls.csv) or node validator.js https://example.com');
    process.exit(1);
  }
  
  console.log('📄 Reading URLs from urls.csv...\n');
  
  let urlEntries;
  try {
    urlEntries = parseCSV(csvPath);
  } catch (error) {
    console.error(`❌ ${error.message}`);
    process.exit(1);
  }
  
  if (urlEntries.length === 0) {
    console.error('❌ No valid URLs found in urls.csv');
    process.exit(1);
  }
  
  console.log(`✓ Found ${urlEntries.length} URL(s) to validate`);
  console.log('═'.repeat(60));
  
  const allUrlResults = [];
  const overallStartTime = Date.now();
  let successCount = 0;
  
  for (let i = 0; i < urlEntries.length; i++) {
    const { url: targetUrl } = urlEntries[i];
    
    console.log(`\n[${i + 1}/${urlEntries.length}] URL: ${targetUrl}`);
    console.log('─'.repeat(60));
    
    let validatedUrl = targetUrl;
    try {
      if (!validatedUrl.startsWith('http://') && !validatedUrl.startsWith('https://')) {
        validatedUrl = 'https://' + validatedUrl;
      }
      new URL(validatedUrl);
    } catch (err) {
      console.error(`  ❌ Invalid URL: ${targetUrl}`);
      continue;
    }
    
    const result = await runValidation(validatedUrl, { headed });
    
    if (result.success) {
      successCount++;
      allUrlResults.push(result);
      console.log(`  ✅ Completed: ${result.stats.passed}/${result.stats.total} passed (${result.stats.passRate.toFixed(1)}%)`);
      if (result.stats.brokenLinks > 0) {
        console.log(`  💔 Broken links found: ${result.stats.brokenLinks}`);
      }
    } else {
      console.error(`  ❌ Validation failed: ${result.error}`);
    }
    
    if (i < urlEntries.length - 1) {
      console.log(`\n  ⏳ Waiting ${CONFIG.delayBetweenUrls}ms before next URL...`);
      await delay(CONFIG.delayBetweenUrls);
    }
  }
  
  const overallEndTime = Date.now();
  
  console.log('\n' + '═'.repeat(60));
  console.log('📊 BATCH VALIDATION SUMMARY');
  console.log('═'.repeat(60));
  console.log(`✅ URLs processed:   ${successCount}/${urlEntries.length}`);
  
  let totalElements = 0, totalPassed = 0, totalFailed = 0, totalBrokenLinks = 0;
  for (const r of allUrlResults) {
    totalElements += r.stats.total;
    totalPassed += r.stats.passed;
    totalFailed += r.stats.failed;
    totalBrokenLinks += r.stats.brokenLinks || 0;
  }
  console.log(`📊 Total elements:   ${totalElements}`);
  console.log(`✅ Total passed:     ${totalPassed}`);
  console.log(`❌ Total failed:     ${totalFailed}`);
  if (totalBrokenLinks > 0) console.log(`💔 Broken links:     ${totalBrokenLinks}`);
  console.log(`📈 Overall success:  ${totalElements > 0 ? ((totalPassed / totalElements) * 100).toFixed(1) : 0}%`);
  console.log(`⏱️ Total duration:   ${((overallEndTime - overallStartTime) / 1000).toFixed(1)}s`);
  
  // Generate reports
  if (allUrlResults.length > 0) {
    console.log(`\n📊 Generating reports...`);
    const htmlPath = await generateBeautifulHtmlReport(allUrlResults, overallStartTime, overallEndTime);
    const excelPath = generateExcelReport(allUrlResults, overallStartTime, overallEndTime);
    console.log(`📄 HTML Report: ${htmlPath}`);
    console.log(`📊 Excel Report: ${excelPath}`);
  }
  
  console.log('\n✅ Batch validation complete!');
  
  if (totalBrokenLinks > 0) {
    console.log(`\n💔 WARNING: Found ${totalBrokenLinks} broken link(s) with missing href attributes!`);
    process.exit(1);
  }
  
  if (totalFailed > 0) {
    console.log(`\n⚠️ Found ${totalFailed} failed clicks across all URLs`);
    process.exit(1);
  }
  
  process.exit(0);
}

if (require.main === module) {
  main().catch(error => {
    console.error('Unhandled error:', error);
    process.exit(1);
  });
}

module.exports = { runValidation, discoverClickableElements, validateClickReception };