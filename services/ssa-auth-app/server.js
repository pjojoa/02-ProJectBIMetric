import fs from "fs";
import dotenv from "dotenv";
import fastify from "fastify";
import cors from "@fastify/cors";
import { getClientCredentialsAccessToken } from "./lib/auth.js";

let config = dotenv.config().parsed || process.env;

if (!config.APS_CLIENT_ID || !config.APS_CLIENT_SECRET) {
    try {
        const configFile = fs.readFileSync("./config.json", "utf8");
        const jsonConfig = JSON.parse(configFile);
        config = { ...config, ...jsonConfig };
    } catch (e) {
        // ignore
    }
}

const { APS_CLIENT_ID, APS_CLIENT_SECRET, PORT } = config;

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

app.get("/token", async () => {
    try {
        return await getClientCredentialsAccessToken(APS_CLIENT_ID, APS_CLIENT_SECRET, SCOPES);
    } catch (err) {
        app.log.error(err);
        throw new Error("Failed to get access token");
    }
});

try {
    await app.listen({ port: PORT || 3000 });
} catch (err) {
    app.log.error(err);
    process.exit(1);
}