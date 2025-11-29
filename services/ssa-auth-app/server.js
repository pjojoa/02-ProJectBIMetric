import fs from "fs";
import dotenv from "dotenv";
import fastify from "fastify";
import cors from "@fastify/cors";
import { getClientCredentialsAccessToken } from "./lib/auth.js";

// Load .env file if it exists (local development)
dotenv.config();

// Try to load from config.json if env vars are not set
if (!process.env.APS_CLIENT_ID || !process.env.APS_CLIENT_SECRET) {
    try {
        const configFile = fs.readFileSync("./config.json", "utf8");
        const jsonConfig = JSON.parse(configFile);
        process.env.APS_CLIENT_ID = process.env.APS_CLIENT_ID || jsonConfig.APS_CLIENT_ID;
        process.env.APS_CLIENT_SECRET = process.env.APS_CLIENT_SECRET || jsonConfig.APS_CLIENT_SECRET;
        process.env.PORT = process.env.PORT || jsonConfig.PORT;
    } catch (e) {
        // config.json doesn't exist, that's ok in production
    }
}

const { APS_CLIENT_ID, APS_CLIENT_SECRET, PORT } = process.env;

if (!APS_CLIENT_ID || !APS_CLIENT_SECRET) {
    console.error("Missing required environment variables: APS_CLIENT_ID, APS_CLIENT_SECRET");
    process.exit(1);
}

const SCOPES = ["viewables:read", "data:read"];

const app = fastify({ logger: true });
await app.register(cors, { origin: "*", methods: ["GET"] });

app.get("/", (request, reply) => {
    reply.type("text/html").send(`
        <h1>APS Auth Service</h1>
        <p>Token service running.</p>
        <p>Use <a href="/token">/token</a> to get an access token.</p>
    `);
});

app.get("/token", async (request) => {
    try {
        const tokenData = await getClientCredentialsAccessToken(APS_CLIENT_ID, APS_CLIENT_SECRET, SCOPES);

        // Check for URN in query params for Smart Environment Detection
        const urn = request.query.urn;
        if (urn) {
            try {
                // Ensure URN is URL-safe Base64 for the API call
                let safeUrn = urn.trim();
                if (safeUrn.toLowerCase().startsWith('urn:')) safeUrn = safeUrn.substring(4);

                // If it's not base64 (contains :), encode it
                if (safeUrn.includes(':')) {
                    safeUrn = Buffer.from('urn:' + urn.trim()).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
                } else {
                    // Fix standard base64 to URL-safe
                    safeUrn = safeUrn.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
                }

                const manifestRes = await fetch(`https://developer.api.autodesk.com/modelderivative/v2/designdata/${safeUrn}/manifest`, {
                    headers: { Authorization: `Bearer ${tokenData.access_token}` }
                });

                if (manifestRes.ok) {
                    const manifest = await manifestRes.json();
                    const hasSVF2 = manifest.derivatives?.some(d => d.outputType === 'svf2');
                    tokenData.detected_env = hasSVF2 ? 'AutodeskProduction2' : 'AutodeskProduction';
                    tokenData.detected_region = manifest.region || 'US';
                }
            } catch (e) {
                app.log.warn(`Failed to detect environment for URN ${urn}: ${e.message}`);
            }
        }

        return tokenData;
    } catch (err) {
        app.log.error(err);
        throw new Error("Failed to get access token");
    }
});

try {
    await app.listen({ port: PORT || 3000, host: '0.0.0.0' });
} catch (err) {
    app.log.error(err);
    process.exit(1);
}