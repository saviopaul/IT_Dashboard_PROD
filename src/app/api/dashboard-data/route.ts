import { NextResponse } from 'next/server';

export async function GET() {
  try {
    // Microsoft 365 API Configuration
    const clientId = process.env.MICROSOFT_CLIENT_ID;
    const tenantId = process.env.MICROSOFT_TENANT_ID;
    const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;

    if (!clientId || !tenantId || !clientSecret) {
      throw new Error('Missing Microsoft 365 configuration');
    }

    console.log('Starting Microsoft 365 authentication...');

    // Get access token
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        client_id: clientId,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: clientSecret,
        grant_type: 'client_credentials'
      })
    });

    if (!tokenResponse.ok) {
      const errorText = await tokenResponse.text();
      console.error('Token error:', tokenResponse.status, errorText);
      throw new Error(`Failed to get access token: ${tokenResponse.status} - ${errorText}`);
    }

    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;

    console.log('Access token obtained successfully');

    // Get SharePoint site ID
    const siteResponse = await fetch('https://graph.microsoft.com/v1.0/sites/burgundyhospitality365-my.sharepoint.com:/personal/savio_bbcollective_co', {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!siteResponse.ok) {
      const errorText = await siteResponse.text();
      console.error('Site error:', siteResponse.status, errorText);
      throw new Error(`Failed to get site ID: ${siteResponse.status} - ${errorText}`);
    }

    const siteData = await siteResponse.json();
    const siteId = siteData.id;

    console.log('Site ID obtained:', siteId);

    // Get Helpdesk Tickets
    const helpdeskResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Issue tracker/items?$expand=fields`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    // Get IT Assets
    const assetsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/IT Asset List Copy/items?$expand=fields`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const [helpdeskData, assetsData] = await Promise.all([
      helpdeskResponse.ok ? helpdeskResponse.json() : { value: [] },
      assetsResponse.ok ? assetsResponse.json() : { value: [] }
    ]);

    console.log('Data fetched successfully:', {
      helpdeskTicketsCount: helpdeskData.value?.length || 0,
      itAssetsCount: assetsData.value?.length || 0
    });

    const dashboardData = {
      helpdeskTickets: helpdeskData.value || [],
      itAssets: assetsData.value || [],
      lastUpdated: new Date().toISOString(),
      source: 'Microsoft 365',
      success: true
    };

    return NextResponse.json(dashboardData);

  } catch (error) {
    console.error('Dashboard API Error:', error);
    return NextResponse.json(
      { 
        error: error.message,
        helpdeskTickets: [],
        itAssets: [],
        lastUpdated: new Date().toISOString(),
        source: 'Error'
      },
      { status: 500 }
    );
  }
}
