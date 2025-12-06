'use client';

import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { AlertTriangle, Users, Wrench, FileText, RefreshCw } from 'lucide-react';

const Dashboard = () => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [lastRefresh, setLastRefresh] = useState(null);

  const fetchRealData = async () => {
    try {
      setLoading(true);
      setError(null);
      
      console.log('Fetching Microsoft 365 data...');
      const response = await fetch('/api/dashboard-data');
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const result = await response.json();
      console.log('Data received:', result);
      
      if (result.error) {
        throw new Error(result.error);
      }
      
      setData(result);
      setLastRefresh(new Date());
      setLoading(false);
    } catch (error) {
      console.error('Error fetching dashboard data:', error);
      setError(error.message);
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchRealData();
    
    // Refresh data every 5 minutes
    const interval = setInterval(fetchRealData, 5 * 60 * 1000);
    
    return () => clearInterval(interval);
  }, []);

  // Calculate metrics from real data
  const getMetrics = () => {
    if (!data || !data.helpdeskTickets || !data.itAssets) return null;

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

    return {
      overdueTickets: overdueTickets.length,
      newJoiners: newJoiners.length,
      renewalsDue: renewalsDue.length,
      assetAlerts: assetAlerts.length,
      totalTickets: data.helpdeskTickets.length,
      totalAssets: data.itAssets.length
    };
  };

  const metrics = getMetrics();

  const getStatusColor = (status) => {
    const statusLower = (status || '').toLowerCase();
    if (statusLower.includes('closed') || statusLower.includes('complete')) return 'bg-green-100 text-green-800';
    if (statusLower.includes('overdue') || statusLower.includes('critical')) return 'bg-red-100 text-red-800';
    if (statusLower.includes('progress') || statusLower.includes('new')) return 'bg-blue-100 text-blue-800';
    return 'bg-gray-100 text-gray-800';
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-center">
          <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-blue-600 mx-auto"></div>
          <p className="mt-4 text-lg">Loading Microsoft 365 data...</p>
          <p className="text-sm text-gray-600">Connecting to your SharePoint Lists</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-center max-w-md">
          <AlertTriangle className="h-16 w-16 text-red-500 mx-auto" />
          <p className="mt-4 text-lg text-red-600">Error connecting to Microsoft 365</p>
          <p className="text-sm text-gray-600 mb-4">{error}</p>
          <button
            onClick={fetchRealData}
            className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center mx-auto"
          >
            <RefreshCw className="h-4 w-4 mr-2" />
            Retry Connection
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="mb-8 flex justify-between items-center">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">IT Command Center Dashboard</h1>
            <p className="text-gray-600">Real-time data from Microsoft 365</p>
            {data.source && (
              <Badge variant="outline" className="mt-2">
                📊 Connected to {data.source}
              </Badge>
            )}
          </div>
          <div className="text-right">
            <button
              onClick={fetchRealData}
              className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center"
            >
              <RefreshCw className="h-4 w-4 mr-2" />
              Refresh
            </button>
            {lastRefresh && (
              <p className="text-xs text-gray-500 mt-2">
                Last updated: {lastRefresh.toLocaleTimeString()}
              </p>
            )}
          </div>
        </div>

        {/* Metrics Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
          {/* Helpdesk Tickets */}
          <Card className="bg-white shadow-sm">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center text-lg">
                <FileText className="h-5 w-5 mr-2 text-blue-600" />
                Helpdesk Tickets
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-blue-600">{metrics.overdueTickets}</div>
              <p className="text-sm text-gray-600">Overdue tickets</p>
              <div className="mt-2">
                <Badge variant="secondary" className="text-xs">
                  {metrics.totalTickets} total tickets
                </Badge>
              </div>
            </CardContent>
          </Card>

          {/* New Joiners */}
          <Card className="bg-white shadow-sm">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center text-lg">
                <Users className="h-5 w-5 mr-2 text-green-600" />
                New Joiners
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-green-600">{metrics.newJoiners}</div>
              <p className="text-sm text-gray-600">Onboarding in progress</p>
              <div className="mt-2">
                <Badge variant="secondary" className="text-xs bg-green-100 text-green-800">
                  Active onboarding
                </Badge>
              </div>
            </CardContent>
          </Card>

          {/* Renewals Due */}
          <Card className="bg-white shadow-sm">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center text-lg">
                <AlertTriangle className="h-5 w-5 mr-2 text-orange-600" />
                Renewals Due
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-orange-600">{metrics.renewalsDue}</div>
              <p className="text-sm text-gray-600">Items expiring soon</p>
              <div className="mt-2">
                <Badge variant="secondary" className="text-xs bg-orange-100 text-orange-800">
                  Next 30 days
                </Badge>
              </div>
            </CardContent>
          </Card>

          {/* Asset Alerts */}
          <Card className="bg-white shadow-sm">
            <CardHeader className="pb-3">
              <CardTitle className="flex items-center text-lg">
                <Wrench className="h-5 w-5 mr-2 text-red-600" />
                Asset Alerts
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-2xl font-bold text-red-600">{metrics.assetAlerts}</div>
              <p className="text-sm text-gray-600">Assets at risk</p>
              <div className="mt-2">
                <Badge variant="secondary" className="text-xs bg-red-100 text-red-800">
                  {metrics.totalAssets} total assets
                </Badge>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Recent Helpdesk Tickets */}
        <Card className="bg-white shadow-sm">
          <CardHeader>
            <CardTitle className="text-lg">Recent Helpdesk Tickets</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-3">
              {data.helpdeskTickets && data.helpdeskTickets.length > 0 ? (
                data.helpdeskTickets.slice(0, 5).map((ticket, index) => (
                  <div key={index} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
                    <div className="flex-1">
                      <h4 className="font-medium text-gray-900">{ticket.Title || ticket.title}</h4>
                      <div className="flex items-center space-x-2 mt-1">
                        <Badge variant="outline" className="text-xs">
                          {ticket.Category || ticket.category || 'General'}
                        </Badge>
                        <span className="text-xs text-gray-500">
                          {ticket.Assignee || ticket.assignee || 'Unassigned'}
                        </span>
                      </div>
                    </div>
                    <div className="text-right">
                      <Badge 
                        variant="secondary" 
                        className={`text-xs ${getStatusColor(ticket.Status || ticket.status)}`}
                      >
                        {ticket.Status || ticket.status || 'Unknown'}
                      </Badge>
                      <p className="text-xs text-gray-500 mt-1">
                        {new Date(ticket.Created || ticket.created || ticket.Created).toLocaleDateString()}
                      </p>
                    </div>
                  </div>
                ))
              ) : (
                <div className="text-center py-8">
                  <FileText className="h-12 w-12 text-gray-400 mx-auto" />
                  <p className="text-gray-600 mt-2">No helpdesk tickets found</p>
                </div>
              )}
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default Dashboard;
