import { getToken } from "next-auth/jwt";

import { GRAPH_SCOPES } from "@/lib/config";

type StoredToken = {
  accessToken?: string;
  refreshToken?: string;
  accessTokenExpires?: number;
  error?: "RefreshAccessTokenError";
};

function usesSecureAuthCookies() {
  const authUrl = process.env.AUTH_URL ?? process.env.NEXTAUTH_URL ?? "";
  return authUrl.startsWith("https://");
}

async function refreshAccessToken(token: StoredToken) {
  const tenantId = process.env.AZURE_AD_TENANT_ID;
  if (!token.refreshToken || !tenantId) {
    return { ...token, error: "RefreshAccessTokenError" as const };
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: process.env.AZURE_AD_CLIENT_ID ?? "",
        client_secret: process.env.AZURE_AD_CLIENT_SECRET ?? "",
        grant_type: "refresh_token",
        refresh_token: token.refreshToken,
        scope: GRAPH_SCOPES,
      }),
    },
  );

  const payload = await response.json();
  if (!response.ok) {
    return { ...token, error: "RefreshAccessTokenError" as const };
  }

  return {
    ...token,
    accessToken: payload.access_token as string,
    refreshToken: (payload.refresh_token as string | undefined) ?? token.refreshToken,
    accessTokenExpires: Date.now() + Number(payload.expires_in ?? 3600) * 1000,
  };
}

export async function getServerAccessToken(req: Request) {
  const secureCookie = usesSecureAuthCookies();
  const token = ((await getToken({
    req,
    secret: process.env.AUTH_SECRET,
    secureCookie,
  })) ??
    (await getToken({
      req,
      secret: process.env.AUTH_SECRET,
      secureCookie: !secureCookie,
    }))) as StoredToken | null;

  if (!token?.accessToken) {
    return null;
  }

  if (token.accessTokenExpires && Date.now() >= token.accessTokenExpires - 60_000) {
    const refreshed = await refreshAccessToken(token);
    return refreshed.error ? null : refreshed.accessToken ?? null;
  }

  return token.accessToken;
}

export async function requireServerAccessToken(req: Request) {
  const token = await getServerAccessToken(req);
  if (!token) {
    throw new Response(JSON.stringify({ message: "No autenticado" }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }
  return token;
}
