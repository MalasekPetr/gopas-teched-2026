import axios from "axios";

// STEP 0: Acquire an app-only token via client_credentials.
// Cached in-process until it expires. Demo only — production would use MSAL + key vault.
let cachedToken: { value: string; expiresAt: number } | null = null;

export async function getAppToken(resource: "graph" | "sharepoint"): Promise<string> {
  const scope =
    resource === "graph"
      ? "https://graph.microsoft.com/.default"
      : `${process.env.SHAREPOINT_TENANT}/.default`;

  // Different scopes need different tokens — keep a tiny cache per scope.
  const key = `${resource}:${scope}`;
  const now = Date.now();
  if (cachedToken && cachedToken.value.startsWith(key) && cachedToken.expiresAt > now + 60_000) {
    return cachedToken.value.slice(key.length + 1);
  }

  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID!,
    client_secret: process.env.CLIENT_SECRET!,
    scope,
    grant_type: "client_credentials",
  });

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  cachedToken = {
    value: `${key}:${data.access_token}`,
    expiresAt: now + data.expires_in * 1000,
  };
  return data.access_token;
}

// Thin Graph helper — we only need GET on list items and POST on sendMail.
export async function graphGet<T>(path: string): Promise<T> {
  const token = await getAppToken("graph");
  const { data } = await axios.get<T>(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return data;
}

export async function graphPost<T>(path: string, body: unknown): Promise<T> {
  const token = await getAppToken("graph");
  const { data } = await axios.post<T>(`https://graph.microsoft.com/v1.0${path}`, body, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });
  return data;
}
