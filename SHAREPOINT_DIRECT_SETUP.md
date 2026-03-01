# Where to Get Info for Direct SharePoint (TDR App)

To enable **direct** TDR file pick from SharePoint (no download-then-upload), the app needs an **Azure AD app registration**. You or your **IT / Microsoft 365 admin** can create it and get the three values below.

---

## Who Can Do This

- **Option A:** You have access to **Azure Portal** (portal.azure.com) or **Microsoft Entra admin center** (entra.microsoft.com) → you can do the steps yourself.
- **Option B:** You don’t have access → ask your **IT team** or **Microsoft 365 / Azure admin** to:
  1. Create an app registration (steps below).
  2. Give you: **Tenant ID**, **Client ID**, and **Client secret** (see “Where to find each value” at the end).

---

## Step-by-Step: Create the App Registration

### 1. Open Azure / Entra

- Go to **https://portal.azure.com** (or **https://entra.microsoft.com**).
- Sign in with your work account (the one that can access your SharePoint).

### 2. Go to App Registrations

- In Azure: left menu → **Microsoft Entra ID** (or **Azure Active Directory**) → **App registrations**.
- In Entra: left menu → **Applications** → **App registrations**.
- Click **+ New registration**.

### 3. Register the App

- **Name:** e.g. `TDR Streamlit SharePoint` (any name is fine).
- **Supported account types:** choose **“Accounts in this organizational directory only (Single tenant)”**.
- **Redirect URI:** leave as **Web** and the box **empty** (we use client secret, not browser login).
- Click **Register**.

### 4. Note Tenant ID and Client ID

On the app’s **Overview** page you’ll see:

- **Application (client) ID** → this is your **Client ID**.
- **Directory (tenant) ID** → this is your **Tenant ID**.

Copy both; you’ll need them for the TDR app.

### 5. Create a Client Secret

- In the left menu of the app, click **Certificates & secrets**.
- Under **Client secrets**, click **+ New client secret**.
- **Description:** e.g. `TDR app`.
- **Expires:** e.g. 24 months (you can create a new one before it expires).
- Click **Add**.
- **Important:** Copy the **Value** of the new secret **immediately**. This is your **Client secret**. You won’t see it again; if you lose it, create a new secret.

### 6. Add Microsoft Graph Permissions

- In the left menu, click **API permissions**.
- Click **+ Add a permission**.
- Choose **Microsoft Graph**.
- Choose **Application permissions** (not “Delegated”).
- Search and add:
  - **Sites.Read.All** (read SharePoint sites).
- Click **Add permissions**.
- Then click **Grant admin consent for [Your org]** (so the app can actually use the permission). An admin may need to do this step.

### 7. (If Needed) Grant Access to the SharePoint Site

- If your org restricts app access per site, an admin may need to grant this app access to:
  - **Site:** `https://amdocs.sharepoint.com/sites/USCCTesting_Offshore`
  - **Folder:** Release & OffCycle → UTMO - Migration (Data creation) → Data Creation → R2 Data

Otherwise, **Sites.Read.All** with admin consent is often enough for the app to read that site.

---

## Where to Find Each Value (Summary)

| What you need   | Where to get it |
|-----------------|------------------|
| **Tenant ID**   | Azure/Entra → App registrations → your app → **Overview** → “Directory (tenant) ID”. |
| **Client ID**   | Same **Overview** page → “Application (client) ID”. |
| **Client secret** | Same app → **Certificates & secrets** → create **New client secret** → copy the **Value** once (it’s hidden after you leave the page). |

---

## Setting Secrets in Streamlit Cloud

In your Streamlit Cloud app:

1. Open the app → **Settings** (⚙️) → **Secrets**.
2. Add these three keys (values from Azure/Entra):

```toml
AZURE_TENANT_ID = "your-directory-tenant-id"
AZURE_CLIENT_ID = "your-application-client-id"
AZURE_CLIENT_SECRET = "your-client-secret-value"
```

3. Save. The app will use them only when **TDR file from** is **SharePoint**; then you’ll see a dropdown to pick a file directly from the R2 Data folder (no download-then-upload).

**Important:** Never commit the client secret to Git. Use only Streamlit secrets or environment variables.

---

## What to Give the TDR App / Developer

Share these **three** values (securely, e.g. over a secure channel or password manager):

1. **Tenant ID**
2. **Client ID**
3. **Client secret**

The **client secret** will be stored only in **Streamlit Cloud secrets** (or env vars), not in the code or in Git.

---

## If You Don’t Have Azure Access

- Send this file (**SHAREPOINT_DIRECT_SETUP.md**) or a short note to your **IT / Microsoft 365 admin** and ask them to:
  - Create an app registration as above.
  - Add **Sites.Read.All** (Application) and grant admin consent.
  - Provide you with **Tenant ID**, **Client ID**, and **Client secret** so the TDR Streamlit app can read from the SharePoint folder.

Once you have those three, you can plug them into the TDR app (via Streamlit secrets) and direct SharePoint file pick can be enabled.
