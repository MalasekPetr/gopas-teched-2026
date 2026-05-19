import axios from "axios";

// Acquire an app-only Graph token via the OAuth 2 client_credentials flow.
// Cached in-process until ~1 min before expiry. Demo only — production would
// use MSAL with a certificate, or a Managed Identity, and Key Vault.
let cachedToken: { value: string; expiresAt: number } | null = null;

async function getAppToken(): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresAt > now + 60_000) return cachedToken.value;

  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID!,
    client_secret: process.env.CLIENT_SECRET!,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  cachedToken = { value: data.access_token, expiresAt: now + data.expires_in * 1000 };
  return data.access_token;
}

const BASE = "https://graph.microsoft.com/v1.0";

export async function graphGet<T>(path: string): Promise<T> {
  const { data } = await axios.get<T>(`${BASE}${path}`, {
    headers: { Authorization: `Bearer ${await getAppToken()}` },
  });
  return data;
}

export async function graphPost<T>(path: string, body: unknown): Promise<T> {
  const { data } = await axios.post<T>(`${BASE}${path}`, body, {
    headers: { Authorization: `Bearer ${await getAppToken()}`, "Content-Type": "application/json" },
  });
  return data;
}

export async function graphPatch<T>(path: string, body: unknown): Promise<T> {
  const { data } = await axios.patch<T>(`${BASE}${path}`, body, {
    headers: { Authorization: `Bearer ${await getAppToken()}`, "Content-Type": "application/json" },
  });
  return data;
}
