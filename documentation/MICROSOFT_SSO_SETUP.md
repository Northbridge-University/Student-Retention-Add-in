# Microsoft SSO Setup Guide

This guide will help you configure Microsoft Single Sign-On (SSO) for the Student Retention Add-in.

## Overview

The add-in now supports **intelligent authentication** with automatic fallback:

1. **Primary**: Microsoft SSO (Azure AD authentication)
2. **Fallback**: Hardcoded demo users (for development/testing)

The system automatically attempts SSO first and falls back to demo accounts if SSO is not configured or fails.

## Current Configuration Status

🟡 **SSO is currently in FALLBACK MODE**

The add-in will use demo accounts until you complete the Azure AD setup below.

## Quick Toggle Configuration

You can control SSO behavior in `/react/src/config/ssoConfig.js`:

```javascript
export const SSOConfig = {
  // Always use fallback (demo users) - set to false to enable SSO attempts
  FORCE_FALLBACK_MODE: false,

  // Automatically fallback to demo users if SSO fails
  ENABLE_SSO_FALLBACK: true,

  // Show Microsoft SSO option in fallback mode
  SHOW_SSO_OPTION: true,
};
```

**For Development/Testing:**
- Set `FORCE_FALLBACK_MODE: true` to always use demo users
- Set `FORCE_FALLBACK_MODE: false` to enable SSO with automatic fallback

## Azure AD App Registration

To enable Microsoft SSO, you need to register an app in Azure AD:

### Step 1: Register the Application

1. Go to [Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)
2. Click **"+ New registration"**
3. Fill in the details:
   - **Name**: Student Retention Add-in
   - **Supported account types**: Accounts in this organizational directory only (Single tenant)
   - **Redirect URI**: Leave blank for now
4. Click **Register**

### Step 2: Configure API Permissions

1. In your app registration, go to **"API permissions"**
2. Click **"+ Add a permission"**
3. Select **"Microsoft Graph"**
4. Select **"Delegated permissions"**
5. Add the following **required** permissions:
   - `openid` - **Required** for SSO authentication
   - `profile` - **Required** for SSO authentication
   - `User.Read` - Read user profile information
6. Click **"Add permissions"**
7. Click **"Grant admin consent"** (requires admin privileges)

**⚠️ Important:** Without `openid` and `profile` permissions, SSO will not work!

### Step 3: Configure Redirect URIs

1. Go to **"Authentication"** in your app registration
2. Click **"+ Add a platform"**
3. Select **"Single-page application"**
4. Add these redirect URIs:
   ```
   https://northbridge-university.github.io/Student-Retention-Add-in/react/dist/index.html
   https://localhost:3000
   ```
5. Check these boxes under **Implicit grant and hybrid flows**:
   - ✅ Access tokens
   - ✅ ID tokens
6. Click **Save**

### Step 4: Note Your Application (Client) ID

1. Go to **"Overview"** in your app registration
2. Copy the **"Application (client) ID"** - you'll need this next

### Step 5: Update the Manifest

1. Open `/manifest.xml` in your project
2. Find the `<WebApplicationInfo>` section **at the end of VersionOverrides** (must be the last child element)
3. Replace the placeholders with your actual values:

```xml
<VersionOverrides>
  <Hosts>
    <!-- ... -->
  </Hosts>
  <Resources>
    <!-- ... -->
  </Resources>

  <!-- WebApplicationInfo MUST be the last child of VersionOverrides -->
  <WebApplicationInfo>
    <Id>YOUR_CLIENT_ID_HERE</Id>
    <Resource>api://northbridge-university.github.io/YOUR_CLIENT_ID_HERE</Resource>
    <Scopes>
      <Scope>openid</Scope>
      <Scope>profile</Scope>
      <Scope>User.Read</Scope>
    </Scopes>
  </WebApplicationInfo>
</VersionOverrides>
```

**Replace:**
- `YOUR_CLIENT_ID_HERE` with your Application (client) ID from Step 4
- Both occurrences must match

**⚠️ CRITICAL:**
- `openid` scope is **required** for SSO to work
- WebApplicationInfo **must** be the last child element of VersionOverrides
- If placed incorrectly, your add-in will be rejected from AppSource

### Step 6: Expose an API (Optional but Recommended)

1. In your app registration, go to **"Expose an API"**
2. Click **"+ Add a scope"**
3. For Application ID URI, use: `api://northbridge-university.github.io/YOUR_CLIENT_ID_HERE`
4. Click **Save and continue**
5. Fill in the scope details:
   - **Scope name**: `access_as_user`
   - **Who can consent**: Admins and users
   - **Admin consent display name**: Access the add-in
   - **Admin consent description**: Allows Office to access the add-in on behalf of the user
   - **User consent display name**: Access your data
   - **User consent description**: Allows the add-in to access your data
6. Click **Add scope**

### Step 7: Configure Pre-authorized Applications (Office)

1. Still in **"Expose an API"**
2. Click **"+ Add a client application"**
3. Add these Office client IDs one by one:

   **Office Desktop (Windows/Mac):**
   ```
   d3590ed6-52b3-4102-aeff-aad2292ab01c
   ```

   **Office Online (Standard):**
   ```
   bc59ab01-8403-45c6-8796-ac3ef710b3e3
   ```

   **Office Online (SharePoint Context):**
   ```
   93d53678-613d-4013-afc1-62e9e444a0a5
   ```

   **Outlook:**
   ```
   57fb890c-0dab-4253-a5e0-7188c88b2bb4
   ```

   **⚠️ Important:** If using Office through SharePoint, you MUST include `93d53678-613d-4013-afc1-62e9e444a0a5`

4. For each, check the scope `access_as_user`
5. Click **Add application**

## Testing SSO

### Option 1: Enable SSO with Fallback (Recommended)

1. In `/react/src/config/ssoConfig.js`:
   ```javascript
   FORCE_FALLBACK_MODE: false,
   ENABLE_SSO_FALLBACK: true,
   ```

2. Build and deploy your add-in
3. The system will:
   - Attempt SSO silently on load
   - Show demo accounts if SSO fails
   - Allow switching between SSO and demo accounts

### Option 2: SSO Only (No Fallback)

1. In `/react/src/config/ssoConfig.js`:
   ```javascript
   FORCE_FALLBACK_MODE: false,
   ENABLE_SSO_FALLBACK: false,
   ```

2. Build and deploy your add-in
3. Only Microsoft SSO will be available

### Option 3: Fallback Only (Current Mode)

1. In `/react/src/config/ssoConfig.js`:
   ```javascript
   FORCE_FALLBACK_MODE: true,
   ```

2. Only demo accounts will be available
3. Great for development without Azure AD setup

## Troubleshooting

### Error: "Office SSO API is not available"
- **Cause**: Office.js not loaded or running outside Office
- **Solution**: Ensure you're running the add-in inside Excel

### Error: "13xxx" codes
- **Cause**: Azure AD configuration issues
- **Common Issues**:
  - `13001`: User not signed in - enable `allowSignInPrompt`
  - `13002`: User aborted sign-in
  - `13003`: User type not supported
  - `13006`: App not trusted - check manifest configuration
  - `13012`: Manifest error - verify WebApplicationInfo

### SSO Always Falls Back to Demo Accounts
1. Check `manifest.xml` has correct Client ID
2. Verify Azure AD app has correct permissions
3. Check browser console for error messages
4. Ensure you're signed in to Microsoft 365

### Token Doesn't Contain User Name
- Ensure `User.Read` and `profile` permissions are granted
- Check that admin consent was granted in Azure AD

## Demo Accounts

When using fallback mode, these demo accounts are available:

- **Angel Baez** - Dean of Academic Affairs
- **Angel Coronel** - Associate Dean of Academic Affairs
- **Darlen Gutierrez** - Student Services Coordinator
- **Victor Blanco** - Student Services Coordinator
- **Kelvin Saliers** - Full Time Instructor
- **Yasser Rojas** - Full Time Instructor

You can modify these in `/react/src/components/utility/SSOtemp.jsx`.

## Security Best Practices

1. **Never commit Azure AD credentials** to version control
2. Use different app registrations for dev/staging/prod
3. Limit API permissions to only what's needed
4. Regularly review and rotate client secrets (if using)
5. Monitor authentication logs in Azure AD

## Advanced Configuration

### Custom Token Claims

To add custom claims to the JWT token:

1. Go to **Token configuration** in Azure AD
2. Click **+ Add optional claim**
3. Select **ID** token type
4. Add claims like `email`, `family_name`, `given_name`

### Multi-Tenant Support

To support users from other organizations:

1. In app registration, change **Supported account types** to:
   - "Accounts in any organizational directory"
2. Update manifest `<Resource>` to use multi-tenant format

## Resources

- [Office Add-ins SSO Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins)
- [Azure AD App Registration](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Troubleshoot SSO in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins)

## Support

If you encounter issues:
1. Check browser console for detailed error messages
2. Review Azure AD sign-in logs
3. Verify all configuration steps were completed
4. Try fallback mode to isolate SSO-specific issues

---

**Last Updated**: 2025-12-26
**Version**: 2.0.0
