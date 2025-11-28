# APS PowerBI Tools

Custom Power BI visual integrating Autodesk Platform Services (APS) Viewer.

## Project Structure

- `/services/ssa-auth-app` - Backend token service
- `/visuals/aps-viewer-visual` - Power BI custom visual

## Deployment

### Backend Service (Render)

1. Push this repository to GitHub
2. Create a new Web Service on [Render](https://render.com)
3. Connect your GitHub repository
4. Configure:
   - **Root Directory**: `services/ssa-auth-app`
   - **Build Command**: `npm install`
   - **Start Command**: `node server.js`
   - **Environment Variables**:
     - `APS_CLIENT_ID`: Your Autodesk client ID
     - `APS_CLIENT_SECRET`: Your Autodesk client secret

### Visual Configuration

After deploying the backend:

1. Note the Render URL (e.g., `https://your-app.onrender.com`)
2. In Power BI, configure the visual's **Access Token Endpoint** to: `https://your-app.onrender.com/token`
3. Adjust **Viewer Environment** and **Region** as needed

## Local Development

### Backend
```bash
cd services/ssa-auth-app
npm install
# Create config.json with your credentials
node server.js
```

### Visual
```bash
cd visuals/aps-viewer-visual
npm install
pbiviz start
```
