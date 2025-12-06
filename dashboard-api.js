// Microsoft 365 API Integration for Professional Dashboard
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

// Get access token
async function getAccessToken() {
  try {
    const response = await fetch(`https://login.microsoftonline.com/${MICROSOFT_CONFIG.tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: MICROSOFT_CONFIG.clientId,
        scope: MICROSOFT_CONFIG.scope,
        client_secret: 'b30d2e1d-e17a-4f4b-8ad6-54100446d5d8',
        grant_type: 'client_credentials'
      })
    });

    if (!response.ok) {
      throw new Error(`Failed to get access token: ${response.status}`);
    }

    const data = await response.json();
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
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Issue tracker/items?$expand=fields`, {
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
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/IT Asset List Copy/items?$expand=fields`, {
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
    backupFailed: 2, // Would come from backup logs
    backupLastRun: '2 hours ago',
    renewalsCount: renewalsDue.length,
    renewalNext: renewalsDue.length > 0 ? '7 days' : 'None',
    ticketsOverdue: overdueTickets.length,
    ticketsOpen: helpdeskTickets.length,
    joinersCount: newJoiners.length,
    joinersProgress: newJoiners.length,
    projectsRisk: 1, // Would come from project tracking
    projectsActive: 5,
    cctvPending: 3, // Would come from CCTV tracking
    cctvCompliance: '85%',
    assetsAlerts: assetAlerts.length,
    assetsTotal: itAssets.length,
    source: helpdeskTickets.length > 0 ? 'Microsoft 365' : 'Demo Data'
  };
}

// Update dashboard with data
function updateDashboard(data) {
  // Hide loading screen
  const loadingScreen = document.getElementById('loading-screen');
  if (loadingScreen) {
    loadingScreen.style.display = 'none';
  }

  // Update Backup Health
  updateMetricValue('backup-failed', data.backupFailed);
  updateMetricValue('backup-last-run', data.backupLastRun);

  // Update Renewals
  updateMetricValue('renewals-count', data.renewalsCount);
  updateMetricValue('renewal-next', data.renewalNext);

  // Update Helpdesk Tickets
  updateMetricValue('tickets-overdue', data.ticketsOverdue);
  updateMetricValue('tickets-open', data.ticketsOpen);

  // Update New Joiners
  updateMetricValue('joiners-count', data.joinersCount);
  updateMetricValue('joiners-progress', data.joinersProgress);

  // Update Project Status
  updateMetricValue('projects-risk', data.projectsRisk);
  updateMetricValue('projects-active', data.projectsActive);

  // Update CCTV Compliance
  updateMetricValue('cctv-pending', data.cctvPending);
  updateMetricValue('cctv-compliance', data.cctvCompliance);

  // Update Asset Alerts
  updateMetricValue('assets-alerts', data.assetsAlerts);
  updateMetricValue('assets-total', data.assetsTotal);

  // Update connection status
  updateConnectionStatus('connected', data.source);
}

// Update metric value with animation
function updateMetricValue(elementId, value) {
  const element = document.getElementById(elementId);
  if (element) {
    element.classList.add('loading');
    element.textContent = value;
    setTimeout(() => {
      element.classList.remove('loading');
    }, 500);
  }
}

// Update connection status
function updateConnectionStatus(status, source = 'Demo') {
  const statusIndicator = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  if (statusIndicator && statusText) {
    statusIndicator.classList.remove('connecting', 'connected', 'error');
    statusIndicator.classList.add(status);
    
    if (status === 'connected') {
      statusText.textContent = `Connected to ${source}`;
    } else if (status === 'error') {
      statusText.textContent = 'Connection Error';
    } else {
      statusText.textContent = 'Connecting...';
    }
  }
}

// Show loading state
function showLoading() {
  const loadingScreen = document.getElementById('loading-screen');
  if (loadingScreen) {
    loadingScreen.style.display = 'flex';
  }
  updateConnectionStatus('connecting');
}

// Show error state
function showError(error) {
  const loadingScreen = document.getElementById('loading-screen');
  if (loadingScreen) {
    loadingScreen.style.display = 'none';
  }
  updateConnectionStatus('error');
  console.error('Dashboard error:', error);
}

// Main function to fetch and update dashboard
async function updateDashboardWithRealData() {
  try {
    showLoading();
    
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

    const metrics = calculateMetrics(helpdeskTickets, itAssets);
    updateDashboard(metrics);

    console.log('Dashboard updated with real data:', metrics);
  } catch (error) {
    console.error('Error updating dashboard:', error);
    showError(error.message);
  }
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
  const autoRefreshToggle = document.getElementById('auto-refresh-toggle');
  const autoRefreshStatus = document.getElementById('auto-refresh-status');
  
  if (refreshBtn) {
    refreshBtn.addEventListener('click', function() {
      console.log('Manual refresh triggered');
      updateDashboardWithRealData();
    });
  }
  
  if (autoRefreshToggle && autoRefreshStatus) {
    autoRefreshToggle.addEventListener('click', function() {
      autoRefresh = !autoRefresh;
      autoRefreshStatus.textContent = autoRefresh ? 'ON' : 'OFF';
      
      if (autoRefresh) {
        startAutoRefresh();
        console.log('Auto-refresh enabled');
      } else {
        stopAutoRefresh();
        console.log('Auto-refresh disabled');
      }
    });
  }
}

// Auto-refresh functions
function startAutoRefresh() {
  stopAutoRefresh(); // Clear any existing interval
  refreshInterval = setInterval(updateDashboardWithRealData, 5 * 60 * 1000);
}

function stopAutoRefresh() {
  if (refreshInterval) {
    clearInterval(refreshInterval);
    refreshInterval = null;
  }
}

// Keyboard shortcuts
function initializeKeyboardShortcuts() {
  document.addEventListener('keydown', function(event) {
    if (event.key === 'F5' || (event.ctrlKey && event.key === 'r')) {
      event.preventDefault();
      console.log('Manual refresh triggered by keyboard');
      updateDashboardWithRealData();
    }
  });
}

// Card click interactions
function initializeCardInteractions() {
  const cards = document.querySelectorAll('.module-card');
  
  cards.forEach(card => {
    card.addEventListener('click', function() {
      const metric = this.getAttribute('data-metric');
      console.log('Card clicked:', metric);
      
      // Add pulse animation
      this.style.animation = 'none';
      setTimeout(() => {
        this.style.animation = '';
      }, 10);
    });
  });
}

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
  console.log('Initializing IT Dashboard...');
  
  // Initialize all components
  initializeNavigation();
  initializeRefresh();
  initializeKeyboardShortcuts();
  initializeCardInteractions();
  
  // Start with demo data first, then try to get real data
  updateDashboard(dashboardData);
  
  // Try to get real data
  updateDashboardWithRealData();
  
  // Start auto-refresh if enabled
  if (autoRefresh) {
    startAutoRefresh();
  }
  
  console.log('IT Dashboard initialized successfully');
});

// Handle page visibility change
document.addEventListener('visibilitychange', function() {
  if (document.hidden) {
    stopAutoRefresh();
  } else {
    if (autoRefresh) {
      startAutoRefresh();
    }
  }
});
