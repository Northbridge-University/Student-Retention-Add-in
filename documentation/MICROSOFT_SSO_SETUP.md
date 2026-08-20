# Microsoft SSO Setup Guide

This guide covers how Microsoft Single Sign-On (SSO) is configured for the Student Retention Add-in,
and how to stand up or re-create the Azure AD app registration behind it.

## Overview

The add-in authenticates with **Microsoft SSO as the primary path**, with an explicit **Guest role**
available as a secondary option:

1. **Primary**: Microsoft SSO (Azure AD / Entra authentication), attempted silently on load
2. **Secondary**: "Continue as Guest" - a deliberately permission-limited role

Guest is a designed role, not a debugging leftover. Users who choose it are gated throughout the app
by `isGuest` / `user === 'Guest'` checks - see `TemplatesModal.jsx` (cannot edit or save templates),
`PersonalizedEmail.jsx` (send options hidden), and `App.jsx` (no email claim available, falls back to
the bare display name). Guests receive no access token.

## Current Configuration Status

**SSO is LIVE.** The Azure AD app registration exists in the Northbridge tenant and `manifest.xml`
carries a real Application (client) ID:

```
71f37f39-a330-413a-be61-0baa5ce03ea3
```

Production `SourceLocation` is
`https://Northbridge-University.github.io/Student-Retention-Add-in/react/dist/index.html`.

The setup steps further down are retained as reference for re-creating or reviewing the registration.
They are **not** outstanding work.

## Configuration Flags

Behavior is controlled in `/react/src/config/ssoConfig.js`:

```javascript
export const SSOConfig = {
  ENABLE_SSO_FALLBACK: true,
  FORCE_FALLBACK_MODE: false,
  SSO_RETRY_ATTEMPTS: 2,
  SHOW_SSO_OPTION: true,
  SSO_TIMEOUT: 10000,
};
```

What these actually do:

| Flag | Effect |
| --- | --- |
| `ENABLE_SSO_FALLBACK` | **Display only.** Its single consumer is a branch in `SSO.jsx` deciding whether to surface 13xxx configuration errors to the user. When `true`, config errors are suppressed from the UI. It grants no access and gates no identity. Setting it to `false` shows raw Azure error codes to end users. |
| `FORCE_FALLBACK_MODE` | When `true`, skips the silent SSO attempt entirely and goes straight to the sign-in screen. Useful for local development. |
| `SSO_RETRY_ATTEMPTS` | Retry count before giving up on the silent attempt. |
| `SHOW_SSO_OPTION` | Whether the "Sign in with Microsoft" button is rendered. |
| `SSO_TIMEOUT` | Milliseconds before the silent SSO attempt is abandoned. |

**Note:** none of these flags control whether the Guest option appears. The "Continue as Guest" button
in `SSO.jsx` is rendered unconditionally. Removing or gating Guest access is a code change to
`SSO.jsx`, not a config change.

## Azure AD App Registration

Reference for reviewing the existing registration, or creating a replacement.

### Step 1: Register the Application

1. Go to [Azure Portal - App Registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)
2. Click **"+ New registration"**
3. Fill in the details:
   - **Name**: Student Retention Add-in
   - **Supported account types**: see the note below
   - **Redirect URI**: Leave blank for now
4. Click **Register**

**Supported account types - review this setting.** The current registration is set to
*All Microsoft account users* (multitenant, including personal Microsoft accounts). For a tool that
surfaces gradebook data on named at-risk students, confirm this breadth is intentional - for example
for cross-campus or sister-school access - or narrow it to
*Accounts in this organizational directory only (Single tenant)*.

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

**Important:** Without `openid` and `profile` permissions, SSO will not work.

The live registration uses delegated permissions only: `openid`, `profile`, `offline_access`,
`User.Read`, `User.ReadBasic.All`, `access_as_user`. Nothing high-privilege.

### Step 3: Configure Redirect URIs

1. Go to **"Authentication"** in your app registration
2. Click **"+ Add a platform"**
3. Select **"Single-page application"**
4. Add these redirect URIs:
   ```
   https://Northbridge-University.github.io/Student-Retention-Add-in/react/dist/index.html
   https://localhost:3000
   ```
5. Check these boxes under **Implicit grant and hybrid flows**:
   - Access tokens
   - ID tokens
6. Click **Save**

This is a public SPA client. There is **no client secret in the production authentication path** -
the add-in calls `Office.auth.getAccessToken` client-side. Any client secrets on the registration
should be reviewed and retired if unused.

### Step 4: Note Your Application (Client) ID

1. Go to **"Overview"** in your app registration
2. Copy the **"Application (client) ID"**

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
    <Resource>api://Northbridge-University.github.io/YOUR_CLIENT_ID_HERE</Resource>
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

**Critical:**
- `openid` scope is **required** for SSO to work
- WebApplicationInfo **must** be the last child element of VersionOverrides
- If placed incorrectly, your add-in will be rejected from AppSource

### Step 6: Expose an API

1. In your app registration, go to **"Expose an API"**
2. Click **"+ Add a scope"**
3. For Application ID URI, use: `api://Northbridge-University.github.io/YOUR_CLIENT_ID_HERE`
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

   **Important:** If using Office through SharePoint, you MUST include `93d53678-613d-4013-afc1-62e9e444a0a5`

4. For each, check the scope `access_as_user`
5. Click **Add application**

## Local Development

To skip the silent SSO attempt while working locally, set `FORCE_FALLBACK_MODE: true` in
`/react/src/config/ssoConfig.js`. This goes straight to the sign-in screen, where "Continue as Guest"
is available without an Azure AD round trip.

Do not commit `FORCE_FALLBACK_MODE: true`.

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

Note that with `ENABLE_SSO_FALLBACK: true` these codes are suppressed from the UI. Check the browser
console to see them.

### Users keep landing on the sign-in screen
1. Check `manifest.xml` has the correct Client ID
2. Verify the Azure AD app has the correct permissions and that admin consent was granted
3. Check the browser console for 13xxx error codes
4. Ensure the user is signed in to Microsoft 365

### Token Doesn't Contain User Name
- Ensure `User.Read` and `profile` permissions are granted
- Check that admin consent was granted in Azure AD

## Guest Role

Selecting "Continue as Guest" sets the display name to `Guest` and provides **no access token**.
Guests are restricted by `isGuest` checks across the app:

- `TemplatesModal.jsx` - cannot save, edit or delete email templates
- `PersonalizedEmail.jsx` - send options are hidden
- `App.jsx` - no email claim is cached, so features keyed on user email are unavailable

If Guest access should be removed, that is a change to the button in `SSO.jsx`. It is not
controllable from `ssoConfig.js`.

## Security Notes

1. Never commit Azure AD credentials to version control
2. Use different app registrations for dev/staging/prod
3. Limit API permissions to only what's needed
4. Review any client secrets on the registration - the production auth path is a public SPA client and does not use one
5. Monitor authentication logs in Azure AD
6. Re-confirm the *Supported account types* breadth described in Step 1

## Resources

- [Office Add-ins SSO Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins)
- [Azure AD App Registration](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Troubleshoot SSO in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins)

---

**Last Updated**: 2026-08-20
**Version**: 2.1.0
