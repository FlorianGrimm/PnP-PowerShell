﻿using System;
using System.Management.Automation;
using System.Threading;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.PowerShell.Commands.Base;
using Resources = SharePointPnP.PowerShell.Commands.Properties.Resources;
using SharePointPnP.PowerShell.CmdletHelpAttributes;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics;

namespace SharePointPnP.PowerShell.Commands
{
    public class PnPCmdlet : PSCmdlet
    {
        public SPOnlineConnection CurrentConnection => Connection ?? SPOnlineConnection.CurrentConnection;

        public ClientContext ClientContext => _ClientContext ?? CurrentConnection?.Context;

        private ClientContext _ClientContext;

        [Parameter(Mandatory = false, HelpMessage = "Optional connection to be used by the cmdlet. Retrieve the value for this parameter by either specifying -ReturnConnection on Connect-PnPOnline or by executing Get-PnPConnection.")] // do not remove '#!#99'
        [PnPParameter(Order = 99)]
        public SPOnlineConnection Connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (Connection == null)
            {
                Connection = SPOnlineConnection.CurrentConnection;
            }
            CurrentConnection?.TelemetryClient.TrackEvent(MyInvocation.MyCommand.Name);
            _ClientContext = CurrentConnection?.Context;

            if (MyInvocation.InvocationName.ToUpper().IndexOf("-SPO", StringComparison.Ordinal) > -1)
            {
                WriteWarning($"PnP Cmdlets starting with the SPO Prefix will be deprecated in the June 2017 release. Please update your scripts and use {MyInvocation.MyCommand.Name} instead.");
            }
            if (Connection == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }
            if (ClientContext == null)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }
            if (Connection.ConnectionMethod == Model.ConnectionMethod.GraphDeviceLogin)
            {
                throw new InvalidOperationException(Resources.NoConnection);
            }
        }

        protected virtual void ExecuteCmdlet()
        { }

        protected override void ProcessRecord()
        {
            var connection = CurrentConnection;
            try
            {
                if (connection.MinimalHealthScore != -1)
                {
                    int healthScore = Utility.GetHealthScore(connection.Url);
                    if (healthScore <= connection.MinimalHealthScore)
                    {
                        ExecuteCmdlet();
                    }
                    else
                    {
                        if (connection.RetryCount != -1)
                        {
                            int retry = 1;
                            while (retry <= connection.RetryCount)
                            {
                                WriteWarning(string.Format(Resources.Retry0ServerNotHealthyWaiting1seconds, retry, connection.RetryWait, healthScore));
                                Thread.Sleep(connection.RetryWait * 1000);
                                healthScore = Utility.GetHealthScore(connection.Url);
                                if (healthScore <= connection.MinimalHealthScore)
                                {
                                    var tag = connection.PnPVersionTag + ":" + MyInvocation.MyCommand.Name.Replace("SPO", "");
                                    if (tag.Length > 32)
                                    {
                                        tag = tag.Substring(0, 32);
                                    }
                                    ClientContext.ClientTag = tag;

                                    ExecuteCmdlet();
                                    break;
                                }
                                retry++;
                            }
                            if (retry > connection.RetryCount)
                            {
                                ThrowTerminatingError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.LimitsExceeded, null));
                            }
                        }
                        else
                        {
                            ThrowTerminatingError(new ErrorRecord(new Exception(Resources.HealthScoreNotSufficient), "HALT", ErrorCategory.LimitsExceeded, null));
                        }
                    }
                }
                else
                {
                    var tag = connection.PnPVersionTag + ":" + MyInvocation.MyCommand.Name.Replace("SPO", "");
                    if (tag.Length > 32)
                    {
                        tag = tag.Substring(0, 32);
                    }
                    ClientContext.ClientTag = tag;

                    ExecuteCmdlet();
                }
            }
            catch (System.Management.Automation.PipelineStoppedException)
            {
                //swallow pipeline stopped exception
            }
            catch (Exception ex)
            {
                //SPOnlineConnection.CurrentConnection.RestoreCachedContext(SPOnlineConnection.CurrentConnection.Url);
                connection.RestoreCachedContext(connection.Url);
                WriteError(new ErrorRecord(ex, "EXCEPTION", ErrorCategory.WriteError, null));
            }
        }

        //protected override void EndProcessing()
        //{
        //    base.EndProcessing();
        //}
    }
}
