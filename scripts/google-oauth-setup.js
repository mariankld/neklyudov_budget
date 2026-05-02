/**
 * One-time helper: opens browser OAuth flow and prints GOOGLE_REFRESH_TOKEN.
 * Prerequisites in .env: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET
 * Add the redirect URI below to Google Cloud Console → Credentials → OAuth client.
 *
 * Usage: node scripts/google-oauth-setup.js
 */
require("dotenv").config();
const http = require("http");
const { URL } = require("url");
const { google } = require("googleapis");

const PORT = Number(process.env.GOOGLE_OAUTH_LOCAL_PORT || 3001);
const REDIRECT_PATH = "/oauth/callback";
const redirectUri =
  process.env.GOOGLE_OAUTH_REDIRECT_URI ||
  `http://127.0.0.1:${PORT}${REDIRECT_PATH}`;

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

function main() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  if (!clientId || !clientSecret) {
    console.error("Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in .env");
    process.exit(1);
  }

  const oauth2Client = new google.auth.OAuth2(
    clientId,
    clientSecret,
    redirectUri
  );

  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: SCOPES,
  });

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url || "/", `http://127.0.0.1:${PORT}`);
      if (url.pathname !== REDIRECT_PATH) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }

      const code = url.searchParams.get("code");
      const err = url.searchParams.get("error");
      if (err) {
        res.writeHead(400);
        res.end(`OAuth error: ${err}`);
        server.close();
        process.exit(1);
        return;
      }

      if (!code) {
        res.writeHead(400);
        res.end("Missing code");
        server.close();
        process.exit(1);
        return;
      }

      const { tokens } = await oauth2Client.getToken(code);
      oauth2Client.setCredentials(tokens);

      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(
        "<p>Success. You can close this tab and return to the terminal.</p>"
      );

      console.log("\nAdd this to your .env:\n");
      if (tokens.refresh_token) {
        console.log(`GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}`);
      } else {
        console.log(
          "(No refresh_token returned — revoke app access in Google Account settings and run again, or ensure prompt=consent and access_type=offline.)"
        );
      }
      console.log(`\nRedirect URI used (must match Console): ${redirectUri}\n`);

      server.close();
      process.exit(0);
    } catch (e) {
      console.error(e);
      res.writeHead(500);
      res.end("Token exchange failed");
      server.close();
      process.exit(1);
    }
  });

  server.listen(PORT, "127.0.0.1", () => {
    console.log(`Open this URL in a browser:\n\n${authUrl}\n`);
    console.log(
      `Listening on ${redirectUri} — add this exact redirect URI in Google Cloud Console if you have not already.\n`
    );
  });
}

main();
