﻿using System;
using System.Linq;
using Microsoft.Online.SharePoint.TenantAdministration;
using SharePointPnP.PowerShell.Commands.Enums;
using Resources = SharePointPnP.PowerShell.Commands.Properties.Resources;

namespace SharePointPnP.PowerShell.Commands.Base
{
    public abstract class PnPAdminCmdlet : PnPCmdlet
    {
        private Tenant _tenant;
        private Uri _baseUri;

        public Tenant Tenant
        {
            get
            {
                if (_tenant == null)
                {
                    _tenant = new Tenant(ClientContext);

                }
                return _tenant;
            }
        }

        public Uri BaseUri => _baseUri;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
#warning TODO TEST

            CurrentConnection.CacheContext();

            if (CurrentConnection.TenantAdminUrl != null && CurrentConnection.ConnectionType == ConnectionType.O365)
            {
                var uri = new Uri(CurrentConnection.Url);
                var uriParts = uri.Host.Split('.');
                if (uriParts[0].ToLower().EndsWith("-admin"))
                {
                    _baseUri =
                        new Uri(
                            $"{uri.Scheme}://{uriParts[0].ToLower().Replace("-admin", "")}.{string.Join(".", uriParts.Skip(1))}{(!uri.IsDefaultPort ? ":" + uri.Port : "")}");
                }
                else
                {
                    _baseUri = new Uri($"{uri.Scheme}://{uri.Authority}");
                }
#warning thinkof
                //SPOnlineConnection.CurrentConnection.CloneContext(SPOnlineConnection.CurrentConnection.TenantAdminUrl);
                CurrentConnection.CloneContext(CurrentConnection.TenantAdminUrl);
            }
            else
            {
                Uri uri = new Uri(ClientContext.Url);
                var uriParts = uri.Host.Split('.');
                if (!uriParts[0].EndsWith("-admin") &&
                    CurrentConnection.ConnectionType == ConnectionType.O365)
                {
                    _baseUri = new Uri($"{uri.Scheme}://{uri.Authority}");

                    var adminUrl = $"https://{uriParts[0]}-admin.{string.Join(".", uriParts.Skip(1))}";

                    CurrentConnection.Context =
                        CurrentConnection.CloneContext(adminUrl);
                }
                else if(CurrentConnection.ConnectionType == ConnectionType.TenantAdmin)
                {
                    _baseUri =
                       new Uri(
                           $"{uri.Scheme}://{uriParts[0].ToLower().Replace("-admin", "")}.{string.Join(".", uriParts.Skip(1))}{(!uri.IsDefaultPort ? ":" + uri.Port : "")}");

                }
            }
        }

        protected override void EndProcessing()
        {
#warning thinkof
            //SPOnlineConnection.CurrentConnection.RestoreCachedContext(SPOnlineConnection.CurrentConnection.Url);
            CurrentConnection.RestoreCachedContext(CurrentConnection.Url);
        }
    }
}
