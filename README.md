SharePointContextSaml
=====================

SAML Extensions for the SharePoint App Context helper


How to use
---------------------

### 1. Create a SharePoin 2013 Provider Hosted App
Create a SharePoint 2013 Provider Hosted and High Trust app using Visual Studio 2013 and the ASP.NET MVC Template

### 2. Add SharePointContextSaml.cs to the web project
Download and add the SharePointContextSaml.cs file into your web project. 
Modify the namespace of the SharePointContextSaml.cs file to match the default namespace of your web project.

### 3. Modify TokenHelper.cs
In order for SharePointContextSaml.cs to be able to extend the default TokenHelper class TokenHelper.cs must be modified so the class declaration of TokenHelper has the partial keyword:
```csharp
public static partial class TokenHelper {...}`
```
### 4. Modify SharePointContext.cs
The static constructor of the default SharePointContext class has to be modified to use the new SharePointContextSaml provider. Locate the static SharePointContextProvider constructor and modify it so it looks like this:
```csharp
static SharePointContextProvider()
    {
    if (!TokenHelper.IsHighTrustApp())
    {
        SharePointContextProvider.current = new SharePointAcsContextProvider();
    }
    else
    {
        if (HttpContext.Current.User.Identity.GetType() == typeof(ClaimsIdentity)) {
            SharePointContextProvider.current = new SharePointHighTrustSamlContextProvider();
        } else {
            SharePointContextProvider.current = new SharePointHighTrustContextProvider();
        }
    }
}
```
### 5. Modify web.config
The extension classes in SharePointContextSaml needs to know if federated (SAML) or Forms based authentication are used. The appSetting `spsaml:ClaimProviderType` can have the value of `SAML` or `FBA`.
You also have to specify the name of the trusted provider added to SharePoint, the `spsaml:TrustedProviderName` app setting is used for that.
Finally we need to specify which claims is used as an identifier, this is done using the `spsaml:IdentityClaimType` setting, which can have one of the following values `SMTP` (e-mail), `SIP` or `UPN`.
The appSettings section should have settings like this after your modifications:
```xml
<add key="spsaml:ClaimProviderType" value="SAML"/>
<add key="spsaml:TrustedProviderName" value="High Trust SAML Demo"/>
<add key="spsaml:IdentityClaimType" value="SMTP"/>
```

### 6. Done
You don't have to modify any of your app code to now leverage SAML Claims in your High Trust SharePoint 2013 app.


More Information
---------------------
For more information this [blog post](http://www.wictorwilen.se/sharepoint-2013-with-saml-claims-and-provider-hosted-apps) by [Wictor Wilén](http://www.wictorwilen.se).


