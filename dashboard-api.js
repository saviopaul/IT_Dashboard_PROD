// Microsoft 365 API Configuration
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

    const data = await response.json();
    return data.access_token;
  } catch (error) {
    console.error('Error getting access token:', error);
    return null;
  }
}

// Get SharePoint data
async function getSharePointData(accessToken) {
  try {
    // Get site ID first
    const siteResponse = await fetch('https://graph.microsoft.com/v1.0/sites/burgundyhospitality365-my.sharepoint.com:/personal/savio_bbcollective_co', {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const siteData = await siteResponse.json();
    const siteId = siteData.id;

    // Get helpdesk tickets
    const helpdeskResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Issue tracker/items?$expand=fields`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    // Get IT assets
    const assetsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/IT Asset List Copy/items?$expand=fields`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const [helpdeskData, assetsData] = await Promise.all([
      helpdeskResponse.ok ? helpdeskResponse.json() : { value: [] },
      assetsResponse.ok ? assetsResponse.json() : { value: [] }
    ]);

    return {
      helpdeskTickets: helpdeskData.value || [],
      itAssets: assetsData.value || []
    };
  } catch (error) {
    console.error('Error getting SharePoint data:', error);
    return { helpdeskTickets: [], itAssets: [] };
  }
}

// Update dashboard with real data
async function updateDashboard() {
  try {
    const accessToken = await getAccessToken();
    if (!accessToken) {
      console.error('Failed to get access token');
      return;
    }

    const data = await getSharePointData(accessToken);
    
    // Update metrics cards
    updateMetrics(data);
    
    // Update recent tickets
    updateRecentTickets(data.helpdeskTickets);
    
  } catch (error) {
    console.error('Error updating dashboard:', error);
  }
}

// Update metrics cards
function updateMetrics(data) {
  const now = new Date();
  const thirtyDaysFromNow = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  const overdueTickets = data.helpdeskTickets.filter(ticket => {
    const createdDate = new Date(ticket.created || ticket.Created || '');
    const status = ticket.Status || ticket.status || '';
    return createdDate < sevenDaysAgo && status !== 'Closed';
  });

  const newJoiners = data.helpdeskTickets.filter(ticket => {
    const category = ticket.Category || ticket.category || '';
    const status = ticket.Status || ticket.status || '';
    return category.toLowerCase().includes('onboarding') && 
           (status === 'In Progress' || status === 'New');
  });

  const renewalsDue = data.itAssets.filter(asset => {
    const warrantyExpiry = asset.WarrantyExpiry || asset.warrantyExpiry || asset.ExpiryDate || '';
    if (!warrantyExpiry) return false;
    
    const expiryDate = new Date(warrantyExpiry);
    return expiryDate <= thirtyDaysFromNow;
  });

  const assetAlerts = data.itAssets.filter(asset => {
    const status = asset.Status || asset.status || '';
    return status.toLowerCase().includes('critical') || 
           status.toLowerCase().includes('risk') || 
           status.toLowerCase().includes('expiring');
  });

  // Update dashboard elements
  document.getElementById('overdue-tickets').textContent = overdueTickets.length;
  document.getElementById('new-joiners').textContent = newJoiners.length;
  document.getElementById('renewals-due').textContent = renewalsDue.length;
  document.getElementById('asset-alerts').textContent = assetAlerts.length;
  document.getElementById('total-tickets').textContent = data.helpdeskTickets.length;
  document.getElementById('total-assets').textContent = data.itAssets.length;
}

// Update recent tickets
function updateRecentTickets(tickets) {
  const recentTicketsContainer = document.getElementById('recent-tickets');
  if (!recentTicketsContainer) return;

  const ticketsHtml = tickets.slice(0, 5).map((ticket, index) => `
    <div class="ticket-item">
      <div class="ticket-info">
        <h4>${ticket.Title || ticket.title}</h4>
        <div class="ticket-meta">
          <span class="ticket-category">${ticket.Category || ticket.category || 'General'}</span>
          <span class="ticket-assignee">${ticket.Assignee || ticket.assignee || 'Unassigned'}</span>
        </div>
      </div>
      <div class="ticket-status">
        <span class="status-badge ${getStatusClass(ticket.Status || ticket.status)}">
          ${ticket.Status || ticket.status || 'Unknown'}
        </span>
        <p class="ticket-date">${new Date(ticket.Created || ticket.created || ticket.Created).toLocaleDateString()}</p>
      </div>
    </div>
  `).join('');

  recentTicketsContainer.innerHTML = ticketsHtml;
}

// Get status CSS class
function getStatusClass(status) {
  const statusLower = (status || '').toLowerCase();
  if (statusLower.includes('closed') || statusLower.includes('complete')) return 'status-closed';
  if (statusLower.includes('overdue') || statusLower.includes('critical')) return 'status-overdue';
  if (statusLower.includes('progress') || statusLower.includes('new')) return 'status-progress';
  return 'status-default';
}

// Initialize dashboard
document.addEventListener('DOMContentLoaded', function() {
  updateDashboard();
  
  // Refresh data every 5 minutes
  setInterval(updateDashboard, 5 * 60 * 1000);
});
