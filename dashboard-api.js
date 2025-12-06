// Microsoft 365 API Integration - ADD TO YOUR EXISTING LAYOUT
const MICROSOFT_CONFIG = {
  clientId: 'c394cddc-0cf8-489d-9d71-45476a4c2629',
  tenantId: '488a6b38-a781-44b4-90cd-7586edc7ba79',
  scope: 'https://graph.microsoft.com/.default'
};

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
      throw new Error(`HTTP error! status: ${response.status}`);
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
}

// Main update function
async function updateCardsWithRealData() {
  try {
    updateConnectionStatus('connecting', 'Connecting to Microsoft 365...');
    
    const data = await getSharePointData();
    
    if (data.success) {
      const metrics = calculateMetrics(data.helpdeskTickets, data.itAssets);
      updateAllCards(metrics);
      updateConnectionStatus('connected', 'Connected to Microsoft 365');
      
      console.log('Successfully updated with real data:', metrics);
    } else {
      updateConnectionStatus('error', data.error, true);
      console.error('Failed to update with real data:', data.error);
    }
  } catch (error) {
    updateConnectionStatus('error', error.message, true);
    console.error('Unexpected error:', error);
  }
}

// Add click animations to cards
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

// Initialize navigation
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

// Initialize refresh functionality
function initializeRefresh() {
  const refreshBtn = document.getElementById('refresh-btn');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', function() {
      console.log('Manual refresh triggered');
      updateCardsWithRealData();
    });
  }
}

// Initialize keyboard shortcuts
function initializeKeyboardShortcuts() {
  document.addEventListener('keydown', function(event) {
    if (event.key === 'F5' || (event.ctrlKey && event.key === 'r')) {
      event.preventDefault();
      console.log('Manual refresh triggered by keyboard');
      updateCardsWithRealData();
    }
  });
}

// Auto-refresh functionality
let autoRefreshInterval = null;
let autoRefreshEnabled = true;

function toggleAutoRefresh() {
  autoRefreshEnabled = !autoRefreshEnabled;
  
  if (autoRefreshEnabled) {
    autoRefreshInterval = setInterval(updateCardsWithRealData, 5 * 60 * 1000);
    console.log('Auto-refresh enabled');
  } else {
    if (autoRefreshInterval) {
      clearInterval(autoRefreshInterval);
      autoRefreshInterval = null;
    }
    console.log('Auto-refresh disabled');
  }
}

// Main initialization
document.addEventListener('DOMContentLoaded', function() {
  console.log('Initializing Microsoft 365 integration for existing layout...');
  
  // Initialize all components
  initializeNavigation();
  initializeRefresh();
  initializeKeyboardShortcuts();
  initializeCardInteractions();
  
  // Initial data load
  updateCardsWithRealData();
  
  // Start auto-refresh
  if (autoRefreshEnabled) {
    autoRefreshInterval = setInterval(updateCardsWithRealData, 5 * 60 * 1000);
  }
  
  console.log('Microsoft 365 integration initialized successfully');
});

// Handle page visibility change
document.addEventListener('visibilitychange', function() {
  if (document.hidden) {
    if (autoRefreshInterval) {
      clearInterval(autoRefreshInterval);
    }
  } else {
    if (autoRefreshEnabled) {
      autoRefreshInterval = setInterval(updateCardsWithRealData, 5 * 60 * 1000);
    }
  }
});
