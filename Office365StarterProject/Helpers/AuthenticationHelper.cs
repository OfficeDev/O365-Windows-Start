// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OAuth;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.CoreServices;
using Office365StarterProject.Helpers;
using System;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Storage;

namespace Office365StarterProject
{
    /// <summary>
    /// Provides clients for the different service endpoints.
    /// </summary>
    internal static class AuthenticationHelper
    {
        // The ClientID is added as a resource in App.xaml when you register the app with Office 365. 
        // As a convenience, we load that value into a variable called ClientID. This way the variable 
        // will always be in sync with whatever client id is added to App.xaml.
        private static readonly string ClientID = App.Current.Resources["ida:ClientID"].ToString();
        private static Uri _returnUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();


        // Properties used for communicating with your Windows Azure AD tenant.
        // The AuthorizationUri is added as a resource in App.xaml when you regiter the app with 
        // Office 365. As a convenience, we load that value into a variable called _commonAuthority, adding _common to this Url to signify
        // multi-tenancy. This way it will always be in sync with whatever value is added to App.xaml.
        private static readonly string CommonAuthority = App.Current.Resources["ida:AuthorizationUri"].ToString() + @"/Common";
        private static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        private const string DiscoveryResourceId = "https://api.office.com/discovery/";

        //Static variables store the clients so that we don't have to create them more than once.
        private static ActiveDirectoryClient _graphClient = null;
        private static OutlookServicesClient _outlookClient = null;
        private static SharePointClient _sharePointClient = null;

        private static ApplicationDataContainer _settings = ApplicationData.Current.LocalSettings;

        //Property for storing and returning the authority used by the last authentication.
        //This value is populated when the user connects to the service and made null when the user signs out.
        private static string LastAuthority
        {
            get
            {
                if (_settings.Values.ContainsKey("LastAuthority") && _settings.Values["LastAuthority"] != null)
                {
                    return _settings.Values["LastAuthority"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["LastAuthority"] = value;
            }
        }
        
        //Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient constructor.
        //This value is populated when the user connects to the service and made null when the user signs out.
        static internal string TenantId
        {
            get
            {
                if (_settings.Values.ContainsKey("TenantId") && _settings.Values["TenantId"] != null)
                {
                    return _settings.Values["TenantId"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["TenantId"] = value;
            }
        }

        // Property for storing the logged-in user so that we can display user properties later.
        //This value is populated when the user connects to the service and made null when the user signs out.
        static internal string LoggedInUser
        {
            get
            {
                if (_settings.Values.ContainsKey("LoggedInUser") && _settings.Values["LoggedInUser"] != null)
                {
                    return _settings.Values["LoggedInUser"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["LoggedInUser"] = value;
            }
        }

        //Property for storing the authentication context.
        public static AuthenticationContext _authenticationContext { get; set; }

        /// <summary>
        /// Checks that a Graph client is available.
        /// </summary>
        /// <returns>The Graph client.</returns>
        public static async Task<ActiveDirectoryClient> EnsureGraphClientCreatedAsync()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_graphClient != null)
            {
                return _graphClient;
            }
            else
            {
                // Active Directory service endpoints
                const string AadServiceResourceId = "https://graph.windows.net/";
                Uri AadServiceEndpointUri = new Uri("https://graph.windows.net/");

                try
                {
                    //First, look for the authority used during the last authentication.
                    //If that value is not populated, use _commonAuthority.
                    string authority = null;
                    if (String.IsNullOrEmpty(LastAuthority))
                    {
                        authority = CommonAuthority;
                    }
                    else
                    {
                        authority = LastAuthority;
                    }

                    // Create an AuthenticationContext using this authority.
                    _authenticationContext = new AuthenticationContext(authority);

                    // Set the value of _authenticationContext.UseCorporateNetwork to true so that you 
                    // can use this app inside a corporate intranet. If the value of UseCorporateNetwork 
                    // is true, you also need to add the Enterprise Authentication, Private Networks, and
                    // Shared User Certificates capabilities in the Package.appxmanifest file.
                    _authenticationContext.UseCorporateNetwork = true;

                    var token = await GetTokenHelperAsync(_authenticationContext, AadServiceResourceId);

                    // Check the token
                    if (String.IsNullOrEmpty(token))
                    {
                        // User cancelled sign-in
                        return null;
                    }
                    else
                    {
                        // Create our ActiveDirectory client.
                        _graphClient = new ActiveDirectoryClient(
                            new Uri(AadServiceEndpointUri, TenantId),
                            async () => await GetTokenHelperAsync(_authenticationContext, AadServiceResourceId));

                        return _graphClient;
                    }


                }

                catch (Exception e)
                {
                    MessageDialogHelper.DisplayException(e as Exception);

                    // Argument exception
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
            }
        }

        /// <summary>
        /// Checks that an OutlookServicesClient object is available. 
        /// </summary>
        /// <returns>The OutlookServicesClient object. </returns>
        public static async Task<OutlookServicesClient> EnsureOutlookClientCreatedAsync()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_outlookClient != null)
            {
                return _outlookClient;
            }
            else
            {
                try
                {
                    // Now get the capability that you are interested in.
                    CapabilityDiscoveryResult result = await GetDiscoveryCapabilityResultAsync("Calendar");

                    _outlookClient = new OutlookServicesClient(
                        result.ServiceEndpointUri,
                        async () => await GetTokenHelperAsync(_authenticationContext, result.ServiceResourceId));

                    return _outlookClient;
                }
                // The following is a list of all exceptions you should consider handling in your app.
                // In the case of this sample, the exceptions are handled by returning null upstream. 
                catch (DiscoveryFailedException dfe)
                {
                    MessageDialogHelper.DisplayException(dfe as Exception);

                    // Discovery failed.
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
                catch (MissingConfigurationValueException mcve)
                {
                    MessageDialogHelper.DisplayException(mcve);

                    // Connected services not added correctly, or permissions not set correctly.
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
                catch (AuthenticationFailedException afe)
                {
                    MessageDialogHelper.DisplayException(afe);

                    // Failed to authenticate the user
                    _authenticationContext.TokenCache.Clear();
                    return null;

                }
                catch (ArgumentException ae)
                {
                    MessageDialogHelper.DisplayException(ae as Exception);
                    // Argument exception
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
            }
        }

        /// <summary>
        /// Checks that a SharePoint client is available to the client.
        /// </summary>
        /// <returns>The SharePoint Online client.</returns>
        public static async Task<SharePointClient> EnsureSharePointClientCreatedAsync()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_sharePointClient != null)
            {
                return _sharePointClient;
            }
            else
            {
                try
                {

                    // Now get the capability that you are interested in.
                    CapabilityDiscoveryResult result = await GetDiscoveryCapabilityResultAsync("MyFiles");

                    _sharePointClient = new SharePointClient(
                        result.ServiceEndpointUri,
                        async () => await GetTokenHelperAsync(_authenticationContext, result.ServiceResourceId));

                    return _sharePointClient;
                }
                catch (DiscoveryFailedException dfe)
                {
                    MessageDialogHelper.DisplayException(dfe as Exception);

                    // Discovery failed.
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
                catch (MissingConfigurationValueException mcve)
                {
                    MessageDialogHelper.DisplayException(mcve);

                    // Connected services not added correctly, or permissions not set correctly.
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
                catch (AuthenticationFailedException afe)
                {
                    MessageDialogHelper.DisplayException(afe);

                    // Failed to authenticate the user
                    _authenticationContext.TokenCache.Clear();
                    return null;

                }
                catch (ArgumentException ae)
                {
                    MessageDialogHelper.DisplayException(ae as Exception);
                    // Argument exception
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
            }
        }


        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static async Task SignOutAsync()
        {
            if (string.IsNullOrEmpty(LoggedInUser))
            {
                return;
            }

            await _authenticationContext.LogoutAsync(LoggedInUser);
            _authenticationContext.TokenCache.Clear();
            //Clean up all existing clients
            _graphClient = null;
            _outlookClient = null;
            _sharePointClient = null;
            //Clear stored values from last authentication. Leave value for LoggedInUser so that we can try again if logout fails.
            _settings.Values["TenantId"] = null;
            _settings.Values["LastAuthority"] = null;

        }

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        private static async Task<string> GetTokenHelperAsync(AuthenticationContext context, string resourceId)
        {
            string accessToken = null;
            AuthenticationResult result = null;

            result = await context.AcquireTokenAsync(resourceId, ClientID, _returnUri);

            if (result.Status == AuthenticationStatus.Success)
            {
                accessToken = result.AccessToken;
                //Store values for logged-in user, tenant id, and authority, so that
                //they can be re-used if the user re-opens the app without disconnecting.
                _settings.Values["LoggedInUser"] = result.UserInfo.UniqueId;
                _settings.Values["TenantId"] = result.TenantId;
                _settings.Values["LastAuthority"] = context.Authority;

                return accessToken;
            }
            else
            {
                return null;
            }
        }

        //Discovery service methods

        public static async Task<DiscoveryServiceCache> CreateAndSaveDiscoveryServiceCacheAsync()
        {
            DiscoveryServiceCache discoveryCache = null;

            var discoveryClient = new DiscoveryClient(DiscoveryServiceEndpointUri,
                                async () => await GetTokenHelperAsync(_authenticationContext, DiscoveryResourceId));

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilitiesAsync();

            discoveryCache = await DiscoveryServiceCache.CreateAndSaveAsync(LoggedInUser, discoveryCapabilityResult);

            return discoveryCache;
        }

        public static async Task<CapabilityDiscoveryResult> GetDiscoveryCapabilityResultAsync(string capability)
        {
            var cacheResult = await DiscoveryServiceCache.LoadAsync();

            CapabilityDiscoveryResult discoveryCapabilityResult = null;

            if (cacheResult != null && cacheResult.DiscoveryInfoForServices.ContainsKey(capability))
            {
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];

                if (LoggedInUser != cacheResult.UserId)
                {
                    // cache is for another user
                    cacheResult = null;
                }
            }

            if (cacheResult == null)
            {
                cacheResult = await CreateAndSaveDiscoveryServiceCacheAsync();
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];
            }

            return discoveryCapabilityResult;
        }

    }
}
//********************************************************* 
// 
//O365-APIs-Start-Windows, https://github.com/OfficeDev/O365-APIs-Start-Windows
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 
