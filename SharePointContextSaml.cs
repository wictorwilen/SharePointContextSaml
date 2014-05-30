using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Configuration;

namespace Replace.This.With.Your.Custom.Namespace
{

    /// <summary>
    /// TokenHelper.cs extensions
    /// </summary>
    public static partial class TokenHelper
    {
        /// <summary>
        /// Identity Claim Type options
        /// </summary>
        public enum IdentityClaimType
        {
            SMTP,
            UPN,
            SIP
        }

        /// <summary>
        /// Claim provider types
        /// </summary>
        public enum ClaimProviderType
        {
            SAML,
            FBA //NOTE: Not tested at all as of now
        }

        private static readonly string TrustedProviderName = WebConfigurationManager.AppSettings.Get("spsaml:TrustedProviderName");
        private static readonly string MembershipProviderName = WebConfigurationManager.AppSettings.Get("spsaml:MembershipProviderName");
        public static readonly IdentityClaimType DefaultIdentityClaimType = (IdentityClaimType)Enum.Parse(typeof(IdentityClaimType), WebConfigurationManager.AppSettings.Get("spsaml:IdentityClaimType"));
        public static readonly ClaimProviderType DefaultClaimProviderType = (ClaimProviderType)Enum.Parse(typeof(ClaimProviderType), WebConfigurationManager.AppSettings.Get("spsaml:ClaimProviderType"));


        private const string CLAIM_TYPE_EMAIL = "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress";
        private const string CLAIM_TYPE_UPN = "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn";
        private const string CLAIM_TYPE_SIP = "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/sip";

        private const string CLAIMS_ID_TYPE_EMAIL = "smtp";
        private const string CLAIMS_ID_TYPE_UPN = "upn";
        private const string CLAIMS_ID_TYPE_SIP = "sip";

        private const int TokenLifetimeMinutes = 1000000;

        //simple class used to hold instance variables for ID claim values
        private class ClaimsUserIdClaim
        {
            public string ClaimsIdClaimType { get; set; }
            public string ClaimsIdClaimValue { get; set; }
        }


        /// <summary>
        /// Retrieves an S2S client context with an access token signed by the application's private certificate on 
        /// behalf of the specified IPrincipal and intended for application at the targetApplicationUri using the 
        /// targetRealm. If no Realm is specified in web.config, an auth challenge will be issued to the 
        /// targetApplicationUri to discover it.
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <param name="UserPrincipal">Identity of the user on whose behalf to create the access token; use HttpContext.Current.User</param>
        /// <param name="SamlIdentityClaimType">The claim type that is used as the identity claim for the user</param>
        /// <param name="IdentityClaimProviderType">The type of identity provider being used</param>
        /// <returns>A ClientContext using an access token with an audience of the target application</returns>
        public static ClientContext GetS2SClientContextWithClaimsIdentity(
            Uri targetApplicationUri,
            ClaimsIdentity identity,
            IdentityClaimType UserIdentityClaimType,
            ClaimProviderType IdentityClaimProviderType,
            bool UseAppOnlyClaim)
        {
            //get the identity claim info first


            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;


            string accessToken = GetS2SClaimsAccessTokenWithClaims(
                targetApplicationUri,
                identity,
                UserIdentityClaimType,
                IdentityClaimProviderType,
                UseAppOnlyClaim);

            return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }



        public static string GetS2SClaimsAccessTokenWithClaims(
            Uri targetApplicationUri,
            ClaimsIdentity identity,
            IdentityClaimType UserIdentityClaimType,
            ClaimProviderType IdentityClaimProviderType,
            bool UseAppOnlyClaim)
        {
            //get the identity claim info first
            TokenHelper.ClaimsUserIdClaim id = null;

            if (IdentityClaimProviderType == ClaimProviderType.SAML)
                id = RetrieveIdentityForSamlClaimsUser(identity, UserIdentityClaimType);
            else
            {
                id = RetrieveIdentityForFbaClaimsUser(identity, UserIdentityClaimType);
            }

            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithClaimsIdentity(identity, UserIdentityClaimType, id, IdentityClaimProviderType) : null;

            return IssueToken(
                ClientId,
                IssuerId,
                realm,
                SharePointPrincipal,
                realm,
                targetApplicationUri.Authority,
                true,
                claims,
                UseAppOnlyClaim,
                id.ClaimsIdClaimType != CLAIMS_ID_TYPE_UPN,
                id.ClaimsIdClaimType,
                id.ClaimsIdClaimValue);
        }

        private static string IssueToken(
           string sourceApplication,
           string issuerApplication,
           string sourceRealm,
           string targetApplication,
           string targetRealm,
           string targetApplicationHostName,
           bool trustedForDelegation,
           IEnumerable<JsonWebTokenClaim> claims,
           bool appOnly = false,
           bool addSamlClaim = false,
           string samlClaimType = "",
           string samlClaimValue = "")
        {
            if (null == SigningCredentials)
            {
                throw new InvalidOperationException("SigningCredentials was not initialized");
            }

            #region Actor token

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<JsonWebTokenClaim> actorClaims = new List<JsonWebTokenClaim>();
            actorClaims.Add(new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid));
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new JsonWebTokenClaim(TokenHelper.TrustedForImpersonationClaimType, "true"));
            }

            //****************************************************************************
            //SPSAML

            //if (samlClaimType == SAML_ID_CLAIM_TYPE_UPN)
            //{
            //    addSamlClaim = true;
            //    samlClaimType = SAML_ID_CLAIM_TYPE_SIP;
            //    samlClaimValue = "bluto2@toys.com";
            //}

            if (addSamlClaim)
                actorClaims.Add(new JsonWebTokenClaim(samlClaimType, samlClaimValue));
            //actorClaims.Add(new JsonWebTokenClaim("smtp", "speschka@vbtoys.com"));
            //****************************************************************************

            // Create token
            JsonWebSecurityToken actorToken = new JsonWebSecurityToken(
                issuer: issuer,
                audience: audience,
                validFrom: DateTime.UtcNow,
                validTo: DateTime.UtcNow.AddMinutes(TokenLifetimeMinutes),
                signingCredentials: SigningCredentials,
                claims: actorClaims);

            string actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);

            if (appOnly)
            {
                // App-only token is the same as actor token for delegated case
                return actorTokenString;
            }

            #endregion Actor token

            #region Outer token

            List<JsonWebTokenClaim> outerClaims = null == claims ? new List<JsonWebTokenClaim>() : new List<JsonWebTokenClaim>(claims);
            outerClaims.Add(new JsonWebTokenClaim(ActorTokenClaimType, actorTokenString));

            //****************************************************************************
            //SPSAML
            if (addSamlClaim)
                outerClaims.Add(new JsonWebTokenClaim(samlClaimType, samlClaimValue));
            //****************************************************************************

            JsonWebSecurityToken jsonToken = new JsonWebSecurityToken(
                nameid, // outer token issuer should match actor token nameid
                audience,
                DateTime.UtcNow,
                DateTime.UtcNow.AddMinutes(10),
                outerClaims);

            string accessToken = new JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken);

            #endregion Outer token

            return accessToken;
        }


        private static TokenHelper.ClaimsUserIdClaim RetrieveIdentityForSamlClaimsUser(ClaimsIdentity identity, IdentityClaimType SamlIdentityClaimType)
        {

            TokenHelper.ClaimsUserIdClaim id = new ClaimsUserIdClaim();

            try
            {
                if (identity.IsAuthenticated)
                {
                    //get the claim type we're looking for
                    string claimType = CLAIM_TYPE_EMAIL;
                    id.ClaimsIdClaimType = CLAIMS_ID_TYPE_EMAIL;

                    //since the vast majority of the time the id claim is email, we'll start out with that
                    //as our default position and only check if that isn't the case
                    if (SamlIdentityClaimType != IdentityClaimType.SMTP)
                    {
                        switch (SamlIdentityClaimType)
                        {
                            case IdentityClaimType.UPN:
                                claimType = CLAIM_TYPE_UPN;
                                id.ClaimsIdClaimType = CLAIMS_ID_TYPE_UPN;
                                break;
                            default:
                                claimType = CLAIM_TYPE_SIP;
                                id.ClaimsIdClaimType = CLAIMS_ID_TYPE_SIP;
                                break;
                        }
                    }

                    foreach (Claim claim in identity.Claims)
                    {
                        if (claim.Type == claimType)
                        {
                            id.ClaimsIdClaimValue = claim.Value;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //not going to do anything here; could look for a missing identity claim but instead will just
                //return an empty string
                Debug.WriteLine(ex.Message);
            }

            return id;
        }

        
        //this is an implementation based on using role claims for SMTP and SIP where the 
        //claim value starts with "SMTP:" or "SIP:".  There isn't a standard way to implement
        //this so you can choose whatever method you want and then update this method appropriately
        private static TokenHelper.ClaimsUserIdClaim RetrieveIdentityForFbaClaimsUser(ClaimsIdentity identity, IdentityClaimType SamlIdentityClaimType)
        {
            throw new NotImplementedException();
        }

        private static JsonWebTokenClaim[] GetClaimsWithClaimsIdentity(ClaimsIdentity indentity, IdentityClaimType SamlIdentityClaimType, TokenHelper.ClaimsUserIdClaim id, ClaimProviderType IdentityClaimProviderType)
        {

            //if an identity claim was not found, then exit
            if (string.IsNullOrEmpty(id.ClaimsIdClaimValue))
                return null;

            Hashtable claimSet = new Hashtable();

            //you always need nii claim, so add that
            claimSet.Add("nii", "temp");

            //set up the nii claim and then add the smtp or sip claim separately
            if (IdentityClaimProviderType == ClaimProviderType.SAML)
                claimSet["nii"] = "trusted:" + TrustedProviderName.ToLower();  //was urn:office:idp:trusted:, but this does not seem to align with what SPIdentityClaimMapper uses
            else
                claimSet["nii"] = "urn:office:idp:forms:" + MembershipProviderName.ToLower();

            //plug in UPN claim if we're using that
            if (id.ClaimsIdClaimType == CLAIMS_ID_TYPE_UPN)
                claimSet.Add("upn", id.ClaimsIdClaimValue.ToLower());

            //now create the JsonWebTokenClaim array
            List<JsonWebTokenClaim> claimList = new List<JsonWebTokenClaim>();

            foreach (string key in claimSet.Keys)
            {
                claimList.Add(new JsonWebTokenClaim(key, (string)claimSet[key]));
            }

            return claimList.ToArray();
        }
    }


    #region HighTrust with SAML

    /// <summary>
    /// Encapsulates all the information from SharePoint in HighTrust mode.
    /// </summary>
    public class SharePointHighTrustSamlContext : SharePointContext
    {
        private readonly ClaimsIdentity logonUserIdentity;

        /// <summary>
        /// The Windows identity for the current user.
        /// </summary>
        public ClaimsIdentity LogonUserIdentity
        {
            get { return this.logonUserIdentity; }
        }

        public override string UserAccessTokenForSPHost
        {
            get
            {
                if (this.SPHostUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.userAccessTokenForSPHost,
                                            () => TokenHelper.GetS2SClaimsAccessTokenWithClaims(
                                                this.SPHostUrl,
                                                this.LogonUserIdentity,
                                                TokenHelper.DefaultIdentityClaimType,
                                                TokenHelper.DefaultClaimProviderType,
                                                false
                                                ));
            }
        }

        public override string UserAccessTokenForSPAppWeb
        {
            get
            {
                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.userAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetS2SClaimsAccessTokenWithClaims(
                                                this.SPAppWebUrl,
                                                this.LogonUserIdentity,
                                                TokenHelper.DefaultIdentityClaimType,
                                                TokenHelper.DefaultClaimProviderType,
                                                false
                                                ));
            }
        }

        public override string AppOnlyAccessTokenForSPHost
        {
            get
            {
                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPHost,
                                            () => TokenHelper.GetS2SClaimsAccessTokenWithClaims(
                                                this.SPAppWebUrl,
                                                null,
                                                TokenHelper.DefaultIdentityClaimType,
                                                TokenHelper.DefaultClaimProviderType,
                                                false));
            }
        }

        public override string AppOnlyAccessTokenForSPAppWeb
        {
            get
            {

                if (this.SPAppWebUrl == null)
                {
                    return null;
                }

                return GetAccessTokenString(ref this.appOnlyAccessTokenForSPAppWeb,
                                            () => TokenHelper.GetS2SClaimsAccessTokenWithClaims(
                                                this.SPAppWebUrl,
                                                null,
                                                TokenHelper.DefaultIdentityClaimType,
                                                TokenHelper.DefaultClaimProviderType,
                                                false));

            }
        }

        public SharePointHighTrustSamlContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, ClaimsIdentity logonUserIdentity)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (logonUserIdentity == null)
            {
                throw new ArgumentNullException("logonUserIdentity");
            }

            this.logonUserIdentity = logonUserIdentity;
        }

        /// <summary>
        /// Ensures the access token is valid and returns it.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        /// <returns>The access token string.</returns>
        private static string GetAccessTokenString(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
        {
            RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

            return IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
        }

        /// <summary>
        /// Renews the access token if it is not valid.
        /// </summary>
        /// <param name="accessToken">The access token to renew.</param>
        /// <param name="tokenRenewalHandler">The token renewal handler.</param>
        private static void RenewAccessTokenIfNeeded(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
        {
            if (IsAccessTokenValid(accessToken))
            {
                return;
            }

            DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);

            if (TokenHelper.HighTrustAccessTokenLifetime > AccessTokenLifetimeTolerance)
            {
                // Make the access token get renewed a bit earlier than the time when it expires
                // so that the calls to SharePoint with it will have enough time to complete successfully.
                expiresOn -= AccessTokenLifetimeTolerance;
            }

            accessToken = Tuple.Create(tokenRenewalHandler(), expiresOn);
        }
    }

    /// <summary>
    /// Default provider for SharePointHighTrustContext.
    /// </summary>
    public class SharePointHighTrustSamlContextProvider : SharePointHighTrustContextProvider
    {

        protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
        {
            ClaimsIdentity logonUserIdentity = HttpContext.Current.User.Identity as ClaimsIdentity;
            //ClaimsIdentity logonUserIdentity = httpRequest.LogonUserIdentity;
            if (logonUserIdentity == null || !logonUserIdentity.IsAuthenticated )
            {
                return null;
            }


            return new SharePointHighTrustSamlContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity);
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            SharePointHighTrustSamlContext spHighTrustContext = spContext as SharePointHighTrustSamlContext;

            if (spHighTrustContext != null)
            {
                Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
                ClaimsIdentity logonUserIdentity = httpContext.Request.LogonUserIdentity;

                return spHostUrl == spHighTrustContext.SPHostUrl &&
                       logonUserIdentity != null &&
                       logonUserIdentity.IsAuthenticated;
            }

            return false;
        }
    }

    #endregion HighTrust
}