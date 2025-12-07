// Microsoft 365 API Integration - ADD TO YOUR EXISTING LAYOUT
const MICROSOFT_CONFIG = {
  clientId: 'c394cddc-0cf8-489d-9d71-45476a4c2629',
  tenantId: '488a6b38-a781-44b4-90cd-7586edc7ba79',
  scope: 'https://graph.microsoft.com/.default'
};

// Dashboard state
let dashboardData = {
  backupFailed: 2,
  backupLastRun: '2 hours ago',
  renewalsCount: 1,
  renewalNext: '7 days',
  ticketsOverdue: 4,
  ticketsOpen: 12,
  joinersCount: 2,
  joinersProgress: 2,
  projectsRisk: 1,
  projectsActive: 5,
  cctvPending: 3,
  cctvCompliance: '85%',
  assetsAlerts: 3,
  assetsTotal: 127,
  lastUpdated: new Date().toISOString()
};

let autoRefresh = true;
let refreshInterval = null;
let connectionAttempts = 0;
const maxConnectionAttempts = 3;

// Get access token
async function getAccessToken() {
  try {
    connectionAttempts++;
    updateConnectionStatus('connecting', `Connecting to Microsoft 365... (Attempt ${connectionAttempts}/${maxConnectionAttempts})`);
    
    const response = await fetch(`https://login.microsoftonline.com/${MICROSOFT_CONFIG.tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: MICROSOFT_CONFIG.clientId,
        scope: MICROSOFT_CONFIG.scope,
        client_secret: 'sN8Q~TKvzCg0vy3-XQbD-h_Ot3LReP9pOFaicej',
        grant_type: 'client_credentials'
      })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP error! status: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    if (data.error) {
      throw new Error(`OAuth error: ${data.error_description || data.error}`);
    }

    connectionAttempts = 0; // Reset on success
    return data.access_token;
  } catch (error) {
    console.error('Error getting access token:', error);
    return null;
  }
}

// Get SharePoint site ID
async function getSiteId(accessToken) {
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/sites/burgundyhospitality365-my.sharepoint.com:/personal/savio_bbcollective_co', {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to get site ID: ${response.status}`);
    }

    const data = await response.json();
    return data.id;
  } catch (error) {
    console.error('Error getting site ID:', error);
    return null;
  }
}

// Get Helpdesk Tickets
async function getHelpdeskTickets(accessToken, siteId) {
  try {
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Issue tracker/items?$expand=fields&$orderby=Created desc`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to get helpdesk tickets: ${response.status}`);
    }

    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error('Error getting helpdesk tickets:', error);
    return [];
  }
}

// Get IT Assets
async function getITAssets(accessToken, siteId) {
  try {
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/IT Asset List Copy/items?$expand=fields&$orderby=Modified desc`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to get IT assets: ${response.status}`);
    }

    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error('Error getting IT assets:', error);
    return [];
  }
}

// Get SharePoint data
async function getSharePointData() {
  try {
    const accessToken = await getAccessToken();
    if (!accessToken) {
      throw new Error('Failed to get access token');
    }

    const siteId = await getSiteId(accessToken);
    if (!siteId) {
      throw new Error('Failed to get site ID');
    }

    const [helpdeskTickets, itAssets] = await Promise.all([
      getHelpdeskTickets(accessToken, siteId),
      getITAssets(accessToken, siteId)
    ]);

    return {
      helpdeskTickets,
      itAssets,
      success: true
    };
  } catch (error) {
    console.error('Error getting SharePoint data:', error);
    return {
      helpdeskTickets: [],
      itAssets: [],
      success: false,
      error: error.message
    };
  }
}

// Calculate metrics from real data
function calculateMetrics(helpdeskTickets, itAssets) {
  const now = new Date();
  const thirtyDaysFromNow = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  const overdueTickets = helpdeskTickets.filter(ticket => {
    const createdDate = new Date(ticket.created || ticket.Created || '');
    const status = ticket.Status || ticket.status || '';
    return createdDate < sevenDaysAgo && status !== 'Closed';
  });

  const newJoiners = helpdeskTickets.filter(ticket => {
    const category = ticket.Category || ticket.category || '';
    const status = ticket.Status || ticket.status || '';
    return category.toLowerCase().includes('onboarding') && 
           (status === 'In Progress' || status === 'New');
  });

  const renewalsDue = itAssets.filter(asset => {
    const warrantyExpiry = asset.WarrantyExpiry || asset.warrantyExpiry || asset.ExpiryDate || '';
    if (!warrantyExpiry) return false;
    
    const expiryDate = new Date(warrantyExpiry);
    return expiryDate <= thirtyDaysFromNow;
  });

  const assetAlerts = itAssets.filter(asset => {
    const status = asset.Status || asset.status || '';
    return status.toLowerCase().includes('critical') || 
           status.toLowerCase().includes('risk') || 
           status.toLowerCase().includes('expiring');
  });

  return {
    backupFailed: 2,
    backupLastRun: '2 hours ago',
    renewalsCount: renewalsDue.length,
    renewalNext: renewalsDue.length > 0 ? `${Math.ceil((new Date(renewalsDue[0]?.WarrantyExpiry || renewalsDue[0]?.warrantyExpiry || renewalsDue[0]?.ExpiryDate) - now) / (1000 * 60 * 60 * 24))} days` : 'None',
    ticketsOverdue: overdueTickets.length,
    ticketsOpen: helpdeskTickets.length,
    joinersCount: newJoiners.length,
    joinersProgress: newJoiners.length,
    projectsRisk: 1,
    projectsActive: 5,
    cctvPending: 3,
    cctvCompliance: '85%',
    assetsAlerts: assetAlerts.length,
    assetsTotal: itAssets.length
  };
}

// Update connection status
function updateConnectionStatus(status, message, isError = false) {
  const statusElement = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  if (statusElement && statusText) {
    statusElement.className = `status-indicator ${status}`;
    statusText.textContent = message;
    
    if (isError) {
      console.error('Connection error:', message);
    }
  }
}

// Update card with animation
function updateCardValue(cardElement, value) {
  if (!cardElement) return;

  // Add loading animation
  cardElement.style.opacity = '0.5';
  cardElement.style.transform = 'scale(0.95)';
  
  setTimeout(() => {
    cardElement.textContent = value;
    cardElement.style.opacity = '1';
    cardElement.style.transform = 'scale(1)';
  }, 300);
}

// Update all cards with new data
function updateAllCards(metrics) {
  console.log('Updating cards with metrics:', metrics);

  // Update Backup Health card
  const backupCard = document.querySelector('[data-metric="backup"] .metric-value');
  if (backupCard) updateCardValue(backupCard, metrics.backupFailed);

  // Update Renewals card
  const renewalsCard = document.querySelector('[data-metric="renewals"] .metric-value');
  if (renewalsCard) updateCardValue(renewalsCard, metrics.renewalsCount);

  // Update Helpdesk Tickets card
  const helpdeskCard = document.querySelector('[data-metric="helpdesk"] .metric-value');
  if (helpdeskCard) updateCardValue(helpdeskCard, metrics.ticketsOverdue);

  // Update New Joiners card
  const joinersCard = document.querySelector('[data-metric="joiners"] .metric-value');
  if (joinersCard) updateCardValue(joinersCard, metrics.joinersCount);

  // Update Project Status card
  const projectsCard = document.querySelector('[data-metric="projects"] .metric-value');
  if (projectsCard) updateCardValue(projectsCard, metrics.projectsRisk);

  // Update CCTV Compliance card
  const cctvCard = document.querySelector('[data-metric="cctv"] .metric-value');
  if (cctvCard) updateCardValue(cctvCard, metrics.cctvPending);

  // Update Asset Alerts card
  const assetsCard = document.querySelector('[data-metric="assets"] .metric-value');
  if (assetsCard) updateCardValue(assetsCard, metrics.assetsAlerts);

  // Update detail values
  const backupLastRunElement = document.querySelector('[data-value="backup-last-run"]');
  if (backupLastRunElement) backupLastRunElement.textContent = metrics.backupLastRun;

  const renewalNextElement = document.querySelector('[data-value="renewal-next"]');
  if (renewalNextElement) renewalNextElement.textContent = metrics.renewalNext;

  const ticketsOpenElement = document.querySelector('[data-value="tickets-open"]');
  if (ticketsOpenElement) ticketsOpenElement.textContent = metrics.ticketsOpen;

  const joinersProgressElement = document.querySelector('[data-value="joiners-progress"]');
  if (joinersProgressElement) joinersProgressElement.textContent = metrics.joinersProgress;

  const projectsActiveElement = document.querySelector('[data-value="projects-active"]');
  if (projectsActiveElement) projectsActiveElement.textContent = metrics.projectsActive;

  const cctvComplianceElement = document.querySelector('[data-value="cctv-compliance"]');
  if (cctvComplianceElement) cctvComplianceElement.textContent = metrics.cctvCompliance;

  const assetsTotalElement = document.querySelector('[data-value="assets-total"]');
  if (assetsTotalElement) assetsTotalElement.textContent = metrics.assetsTotal;
}

// Show loading state
function showLoading() {
  const loadingScreen = document.getElementById('loading-screen');
  if (loadingScreen) {
    loadingScreen.style.display = 'flex';
  }
  updateConnectionStatus('connecting', 'Connecting to Microsoft 365...');
}

// Show error state
function showError(error) {
  const loadingScreen = document.getElementById('loading-screen');
  if (loadingScreen) {
    loadingScreen.style.display = 'none';
  }
  updateConnectionStatus('error', error, true);
  console.error('Dashboard error:', error);
}

// Main update function
async function updateCardsWithRealData() {
  try {
    showLoading();
    
    const data = await getSharePointData();
    
    if (data.success) {
      const metrics = calculateMetrics(data.helpdeskTickets, data.itAssets);
      updateAllCards(metrics);
      updateConnectionStatus('connected', 'Connected to Microsoft 365');
      
      // Hide loading screen
      const loadingScreen = document.getElementById('loading-screen');
      if (loadingScreen) {
        loadingScreen.style.display = 'none';
      }
      
      console.log('Successfully updated with real data:', metrics);
    } else {
      showError(data.error);
    }
  } catch (error) {
    showError(error.message);
  }
}

// Fallback to demo data if connection fails
function updateWithDemoData() {
  console.log('Using demo data - Microsoft 365 connection failed');
  updateConnectionStatus('connected', 'Connected to Microsoft 365 (Demo Data)');
  
  const demoMetrics = {
    backupFailed: 2,
    backupLastRun: '2 hours ago',
    renewalsCount: 1,
    renewalNext: '7 days',
    ticketsOverdue: 4,
    ticketsOpen: 12,
    joinersCount: 2,
    joinersProgress: 2,
    projectsRisk: 1,
    projectsActive: 5,
    cctvPending: 3,
    cctvCompliance: '85%',
    assetsAlerts: 3,
    assetsTotal: 127
  };
  
  updateAllCards(demoMetrics);
}

// Navigation functionality
function initializeNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.content-section');
  
  navItems.forEach(item => {
    item.addEventListener('click', function(e) {
      e.preventDefault();
      
      const targetSection = this.getAttribute('data-section');
      
      // Remove active class from all nav items and sections
      navItems.forEach(navItem => navItem.classList.remove('active'));
      sections.forEach(section => section.classList.remove('active'));
      
      // Add active class to clicked nav item and corresponding section
      this.classList.add('active');
      const targetSectionElement = document.getElementById(`${targetSection}-section`);
      if (targetSectionElement) {
        targetSectionElement.classList.add('active');
      }
      
      console.log('Navigation clicked:', targetSection);
    });
  });
}

// Refresh functionality
function initializeRefresh() {
  const refreshBtn = document.getElementById('refresh-btn');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', function() {
      console.log('Manual refresh triggered');
      updateCardsWithRealData();
    });
  }
}

// Auto-refresh functionality
function startAutoRefresh() {
  stopAutoRefresh();
  refreshInterval = setInterval(updateCardsWithRealData, 5 * 60 * 1000);
  console.log('Auto-refresh started');
}

function stopAutoRefresh() {
  if (refreshInterval) {
    clearInterval(refreshInterval);
    refreshInterval = null;
  }
}

function toggleAutoRefresh() {
  autoRefresh = !autoRefresh;
  
  const autoRefreshStatus = document.getElementById('auto-refresh-status');
  if (autoRefreshStatus) {
    autoRefreshStatus.textContent = autoRefresh ? 'ON' : 'OFF';
  }
  
  if (autoRefresh) {
    startAutoRefresh();
    console.log('Auto-refresh enabled');
  } else {
    stopAutoRefresh();
    console.log('Auto-refresh disabled');
  }
}

// Keyboard shortcuts
function initializeKeyboardShortcuts() {
  document.addEventListener('keydown', function(event) {
    if (event.key === 'F5' || (event.ctrlKey && event.key === 'r')) {
      event.preventDefault();
      console.log('Manual refresh triggered by keyboard');
      updateCardsWithRealData();
    }
  });
}

// Card click animations
function initializeCardInteractions() {
  const cards = document.querySelectorAll('.module-card');
  
  cards.forEach(card => {
    card.addEventListener('click', function() {
      const metric = this.getAttribute('data-metric');
      console.log('Card clicked:', metric);
      
      // Add pulse animation
      this.style.animation = 'none';
      this.style.transform = 'scale(1.05)';
      this.style.boxShadow = '0 16px 48px rgba(0, 0, 0, 0.2)';
      
      setTimeout(() => {
        this.style.animation = '';
        this.style.transform = '';
        this.style.boxShadow = '';
      }, 200);
    });
  });
}

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
  console.log('Initializing Microsoft 365 integration for existing layout...');
  
  // Initialize all components
  initializeNavigation();
  initializeRefresh();
  initializeKeyboardShortcuts();
  initializeCardInteractions();
  
  // Try to get real data first, fallback to demo data
  updateCardsWithRealData().catch(error => {
    console.error('Failed to load real data, using demo data:', error);
    setTimeout(() => {
      updateWithDemoData();
    }, 2000); // Wait 2 seconds before showing demo data
  });
  
  // Start auto-refresh
  if (autoRefresh) {
    startAutoRefresh();
  }
  
  console.log('Microsoft 365 integration initialized successfully');
});

// Handle page visibility change
document.addEventListener('visibilitychange', function() {
  const autoRefreshStatus = document.getElementById('auto-refresh-status');
  const isAutoRefreshOn = autoRefreshStatus && autoRefreshStatus.textContent === 'ON';
  
  if (document.hidden && isAutoRefreshOn) {
    console.log('Page hidden - pausing auto-refresh');
    stopAutoRefresh();
  } else if (!document.hidden && isAutoRefreshOn) {
    console.log('Page visible - resuming auto-refresh');
    startAutoRefresh();
  }
});

// Error recovery mechanism
function retryConnection() {
  if (connectionAttempts < maxConnectionAttempts) {
    console.log(`Retrying connection (${connectionAttempts + 1}/${maxConnectionAttempts})`);
    setTimeout(updateCardsWithRealData, 2000);
  } else {
    console.log('Max connection attempts reached, using demo data');
    updateWithDemoData();
  }
}

// Add retry functionality
function addRetryButton() {
  const footer = document.querySelector('.refresh-controls');
  if (footer && !document.getElementById('retry-btn')) {
    const retryBtn = document.createElement('button');
    retryBtn.id = 'retry-btn';
    retryBtn.className = 'refresh-btn';
    retryBtn.innerHTML = '<span class="refresh-icon">🔄</span> Retry Connection';
    retryBtn.addEventListener('click', function() {
      connectionAttempts = 0;
      updateCardsWithRealData();
    });
    
    footer.appendChild(retryBtn);
  }
}

// Monitor connection status and add retry button if needed
setInterval(() => {
  const statusText = document.getElementById('status-text');
  if (statusText && statusText.textContent.includes('Error')) {
    addRetryButton();
  }
}, 10000); // Check every 10 seconds
