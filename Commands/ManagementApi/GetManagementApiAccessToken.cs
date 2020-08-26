﻿#if !ONPREMISES
using System;
using System.Management.Automation;
using PnP.PowerShell.CmdletHelpAttributes;
using PnP.PowerShell.Commands.Base;
using PnP.PowerShell.Commands.Model;

namespace PnP.PowerShell.Commands.ManagementApi
{
    [Cmdlet(VerbsCommon.Get, "PnPManagementApiAccessToken")]
    [CmdletHelp("Gets an access token for the Office 365 Management API",
        Category = CmdletHelpCategory.ManagementApi,
        SupportedPlatform = CmdletSupportedPlatform.Online)]
    [CmdletExample(
       Code = "PS:> Get-PnPManagementApiAccessToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret)",
       Remarks = "Retrieves access token for the Office 365 Management API",
       SortOrder = 1)]
    [Obsolete("Connect using Connect-PnPOnline -ClientId -ClientSecret -AADDomain instead to set up a connection with which Office 365 Management API cmdlets can be executed")]
    public class GetManagementApiAccessToken : BasePSCmdlet
    {
        [Parameter(Mandatory = true, HelpMessage = "The Tenant ID to connect to the Office 365 Management API")]
        public string TenantId;

        [Parameter(Mandatory = true, HelpMessage = "The App\\Client ID of the app which gives you access to the Office 365 Management API")]
        public string ClientId;

        [Parameter(Mandatory = true, HelpMessage = "The Client Secret of the app which gives you access to the Office 365 Management API")]
        public string ClientSecret;

        protected override void ExecuteCmdlet()
        {
            var officeManagementApiToken = OfficeManagementApiToken.AcquireApplicationToken(TenantId, ClientId, ClientSecret, PnPConnection.CurrentConnection.AzureEnvironment);
            WriteObject(officeManagementApiToken.AccessToken);
        }
    }
}
#endif