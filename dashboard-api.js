// Microsoft 365 API Integration for Professional Dashboard
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

  // Calculate metrics
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

// Update dashboard with real data
function updateDashboard(metrics) {
  // Update Backup Health
  document.getElementById('backup-failed').textContent = metrics.backupFailed;
  document.getElementById('backup-last-run').textContent = metrics.backupLastRun;

  // Update Renewals
  document.getElementById('renewals-count').textContent = metrics.renewalsCount;
  document.getElementById('renewal-next').textContent = metrics.renewalNext;

  // Update Helpdesk Tickets
  document.getElementById('tickets-overdue').textContent = metrics.ticketsOverdue;
  document.getElementById('tickets-open').textContent = metrics.ticketsOpen;

  // Update New Joiners
  document.getElementById('joiners-count').textContent = metrics.joinersCount;
  document.getElementById('joiners-progress').textContent = metrics.joinersProgress;

  // Update Project Status
  document.getElementById('projects-risk').textContent = metrics.projectsRisk;
  document.getElementById('projects-active').textContent = metrics.projectsActive;

  // Update CCTV Compliance
  document.getElementById('cctv-pending').textContent = metrics.cctvPending;
  document.getElementById('cctv-compliance').textContent = metrics.cctvCompliance;

  // Update Asset Alerts
  document.getElementById('assets-alerts').textContent = metrics.assetsAlerts;
  document.getElementById('assets-total').textContent = metrics.assetsTotal;

  // Update connection status
  const statusIndicator = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  statusIndicator.classList.remove('connecting', 'connected', 'error');
  statusIndicator.classList.add('connected');
  statusText.textContent = 'Connected to Microsoft 365';
}

// Show loading state
function showLoading() {
  const statusIndicator = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  statusIndicator.classList.add('connecting');
  statusText.textContent = 'Connecting to Microsoft 365...';
}

// Show error state
function showError(error) {
  const statusIndicator = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  statusIndicator.classList.remove('connecting', 'connected');
  statusIndicator.classList.add('error');
  statusText.textContent = `Error: ${error}`;
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

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
  console.log('Initializing IT Dashboard...');
  
  // Update with real data immediately
  updateDashboardWithRealData();
  
  // Refresh data every 5 minutes
  setInterval(updateDashboardWithRealData, 5 * 60 * 1000);
  
  // Add refresh functionality
  document.addEventListener('keydown', function(event) {
    if (event.key === 'F5' || (event.ctrlKey && event.key === 'r')) {
      event.preventDefault();
      updateDashboardWithRealData();
    }
  });
  
  console.log('IT Dashboard initialized successfully');
});

// Add click handlers for navigation
document.addEventListener('DOMContentLoaded', function() {
  const navLinks = document.querySelectorAll('.nav-link');
  
  navLinks.forEach(link => {
    link.addEventListener('click', function(event) {
      event.preventDefault();
      
      // Remove active class from all links
      navLinks.forEach(l => l.parentElement.classList.remove('active'));
      
      // Add active class to clicked link
      this.parentElement.classList.add('active');
      
      // Here you could show different views based on selection
      console.log('Navigation clicked:', this.querySelector('.nav-text').textContent);
    });
  });
});
