import { isGraphConfigured } from "./graph-config";

export { isGraphConfigured };

type TokenCache = { token: string; expiresAt: number };

let cache: TokenCache | null = null;

export async function getGraphAccessToken(): Promise<string> {
  if (!isGraphConfigured()) {
    throw new Error("Microsoft Graph の環境変数が未設定です");
  }

  const now = Date.now();
  if (cache && cache.expiresAt > now + 60_000) {
    return cache.token;
  }

  const tenantId = process.env.AZURE_TENANT_ID!;
  const body = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID!,
    client_secret: process.env.AZURE_CLIENT_SECRET!,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`トークン取得に失敗しました: ${res.status} ${text}`);
  }

  const json = (await res.json()) as { access_token: string; expires_in: number };
  cache = {
    token: json.access_token,
    expiresAt: now + json.expires_in * 1000,
  };
  return json.access_token;
}

export async function graphFetch(path: string, init?: RequestInit): Promise<Response> {
  const token = await getGraphAccessToken();
  const url = path.startsWith("https://")
    ? path
    : `https://graph.microsoft.com/v1.0${path}`;
  return fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...init?.headers,
    },
  });
}
