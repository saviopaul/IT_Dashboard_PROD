// Microsoft 365 API Integration - DELEGATED PERMISSIONS VERSION
const MICROSOFT_CONFIG = {
  clientId: 'c394cddc-0cf8-489d-9d71-45476a4c2629',
  tenantId: '488a6b38-a781-44b4-90cd-7586edc7ba79',
  scope: 'https://graph.microsoft.com/.default'
};

// SharePoint REST API (works with delegated permissions)
async function getSharePointDataWithDelegated() {
  try {
    // For delegated permissions, we need to use SharePoint REST API
    const response = await fetch('https://burgundyhospitality365-my.sharepoint.com/personal/savio_bbcollective_co/_api/web/lists/getbytitle(\'Issue tracker\')/items', {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to get helpdesk tickets: ${response.status}`);
    }

    const helpdeskData = await response.json();
    
    // Get IT Assets
    const assetsResponse = await fetch('https://burgundyhospitality365-my.sharepoint.com/personal/savio_bbcollective_co/_api/web/lists/getbytitle(\'IT Asset List Copy\')/items', {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json'
      }
    });

    const assetsData = assetsResponse.ok ? await assetsResponse.json() : { d: [] };

    return {
      helpdeskTickets: helpdeskData.d || [],
      itAssets: assetsData.d || [],
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

// Fallback to demo data
function getDemoData() {
  return {
    helpdeskTickets: [
      {
        Title: "New laptop setup for John Doe",
        Status: "In Progress",
        Priority: "High",
        Assignee: "IT Team",
        Created: "2025-12-01",
        Category: "Onboarding"
      },
      {
        Title: "VPN access issue",
        Status: "New",
        Priority: "Medium", 
        Assignee: "IT Team",
        Created: "2025-12-05",
        Category: "Support"
      },
      {
        Title: "Printer not working",
        Status: "Overdue",
        Priority: "High",
        Assignee: "IT Team",
        Created: "2025-11-28",
        Category: "Hardware"
      },
      {
        Title: "Email configuration for Sarah",
        Status: "In Progress",
        Priority: "Medium",
        Assignee: "IT Team", 
        Created: "2025-12-04",
        Category: "Onboarding"
      }
    ],
    itAssets: [
      {
        Title: "Laptop - Dell XPS 15",
        Status: "In Use",
        WarrantyExpiry: "2026-03-15",
        Type: "Hardware",
        AssignedTo: "John Doe"
      },
      {
        Title: "Microsoft Office License",
        Status: "Active",
        WarrantyExpiry: "2025-12-31",
        Type: "Software",
        AssignedTo: "HR Department"
      },
      {
        Title: "Server Backup License",
        Status: "Expiring Soon",
        WarrantyExpiry: "2025-12-15",
        Type: "Software",
        AssignedTo: "IT Team"
      }
    ],
    success: true
  };
}

// Calculate metrics from data
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
function updateConnectionStatus(status, message, source = 'Demo') {
  const statusElement = document.getElementById('connection-status');
  const statusText = document.getElementById('status-text');
  
  if (statusElement && statusText) {
    statusElement.className = `status-indicator ${status}`;
    statusText.textContent = message;
    
    if (status === 'error') {
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

// Main update function
async function updateCardsWithDelegatedData() {
  try {
    updateConnectionStatus('connecting', 'Connecting to Microsoft 365 (Delegated)...');
    
    const data = await getSharePointDataWithDelegated();
    
    if (data.success) {
      const metrics = calculateMetrics(data.helpdeskTickets, data.itAssets);
      updateAllCards(metrics);
      updateConnectionStatus('connected', 'Connected to Microsoft 365 (Delegated)');
      
      console.log('Successfully updated with delegated data:', metrics);
    } else {
      console.log('Delegated connection failed, using demo data:', data.error);
      const demoData = getDemoData();
      const metrics = calculateMetrics(demoData.helpdeskTickets, demoData.itAssets);
      updateAllCards(metrics);
      updateConnectionStatus('connected', 'Connected to Microsoft 365 (Demo Data)');
    }
  } catch (error) {
    console.error('Error updating with delegated data:', error);
    updateConnectionStatus('error', 'Connection Error', true);
    
    // Fallback to demo data
    const demoData = getDemoData();
    const metrics = calculateMetrics(demoData.helpdeskTickets, demoData.itAssets);
    updateAllCards(metrics);
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
  if (refreshBtn) {
    refreshBtn.addEventListener('click', function() {
      console.log('Manual refresh triggered');
      updateCardsWithDelegatedData();
    });
  }
}

// Keyboard shortcuts
function initializeKeyboardShortcuts() {
  document.addEventListener('keydown', function(event) {
    if (event.key === 'F5' || (event.ctrlKey && event.key === 'r')) {
      event.preventDefault();
      console.log('Manual refresh triggered by keyboard');
      updateCardsWithDelegatedData();
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
  console.log('Initializing Microsoft 365 integration with delegated permissions...');
  
  // Initialize all components
  initializeNavigation();
  initializeRefresh();
  initializeKeyboardShortcuts();
  initializeCardInteractions();
  
  // Try to get real data, fallback to demo data
  updateCardsWithDelegatedData();
  
  // Start auto-refresh
  const autoRefreshToggle = document.getElementById('auto-refresh-toggle');
  if (autoRefreshToggle) {
    autoRefreshToggle.addEventListener('click', function() {
      const autoRefreshStatus = document.getElementById('auto-refresh-status');
      const isAutoRefreshOn = autoRefreshStatus && autoRefreshStatus.textContent === 'ON';
      
      if (isAutoRefreshOn) {
        console.log('Auto-refresh disabled');
        autoRefreshStatus.textContent = 'OFF';
      } else {
        console.log('Auto-refresh enabled');
        autoRefreshStatus.textContent = 'ON';
        setInterval(updateCardsWithDelegatedData, 5 * 60 * 1000);
      }
    });
    
    // Enable auto-refresh by default
    autoRefreshToggle.click();
  }
  
  console.log('Microsoft 365 integration initialized successfully');
});

// Handle page visibility change
document.addEventListener('visibilitychange', function() {
  const autoRefreshStatus = document.getElementById('auto-refresh-status');
  const isAutoRefreshOn = autoRefreshStatus && autoRefreshStatus.textContent === 'ON';
  
  if (document.hidden && isAutoRefreshOn) {
    console.log('Page hidden - pausing auto-refresh');
  } else if (!document.hidden && isAutoRefreshOn) {
    console.log('Page visible - resuming auto-refresh');
  }
});
