# Deployment Guide for Render

## Step 1: Deploy Backend to Render

1. Go to [dashboard.render.com](https://dashboard.render.com) and sign in
2. Click **New +** → **Web Service**
3. Connect your GitHub account and select repository: `pjojoa/02-ProJectBIMetric`
4. Configure the service:
   - **Name**: `aps-auth-service` (or your preferred name)
   - **Root Directory**: `services/ssa-auth-app`
   - **Environment**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Plan**: Free

5. Add Environment Variables:
   - Click **Add Environment Variable**
   - Add `APS_CLIENT_ID` = `HWD9WiGYiNwqzIJ0FHjiAlTIpcGbFkbsb7d3PGOZZ3kvGAx2`
   - Add `APS_CLIENT_SECRET` = `Lx0mKUIMqK5hRgVe9BYSNsvNcLgT3zwIl6SYuVTm9Thjymwg5ZFVqRxeHBbunWSe`

6. Click **Create Web Service**

7. Wait for deployment to complete (5-10 minutes)

8. Copy your service URL (e.g., `https://aps-auth-service.onrender.com`)

## Step 2: Update Visual Configuration

The visual is already configured to work with a configurable endpoint. You don't need to repackage it.

In Power BI Desktop or Power BI Service:

1. Add the visual to your report
2. Select the visual
3. Go to **Format** pane → **Viewer Runtime**
4. Update **Access Token Endpoint** to: `https://YOUR-RENDER-URL.onrender.com/token`
   (Replace `YOUR-RENDER-URL` with your actual Render service URL)
5. Configure other settings as needed:
   - **Viewer Environment**: `AutodeskProduction2` (default) or `AutodeskProduction`
   - **Region**: `US` (default) or `EMEA`

## Step 3: Test in Power BI Service

1. Publish your report to Power BI Service
2. Open the report in Power BI Service
3. The visual should now load models using the Render-hosted token service

## Troubleshooting

- **Service sleeping**: Free Render services sleep after 15 minutes of inactivity. First load may take 30-60 seconds.
- **CORS errors**: The backend is configured with `origin: "*"` to allow requests from Power BI.
- **Token errors**: Check Render logs to verify the service is running and credentials are correct.
