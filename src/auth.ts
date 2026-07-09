import NextAuth from "next-auth";
import MicrosoftEntraID from "next-auth/providers/microsoft-entra-id";

import { GRAPH_SCOPES } from "@/lib/config";

const tenantId = process.env.AZURE_AD_TENANT_ID;

async function refreshAccessToken(token: {
  refreshToken?: string;
  accessToken?: string;
  accessTokenExpires?: number;
}) {
  if (!token.refreshToken || !tenantId) {
    return { ...token, error: "RefreshAccessTokenError" as const };
  }

  try {
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

    const refreshed = await response.json();

    if (!response.ok) {
      throw new Error(refreshed.error_description ?? "Could not refresh access token");
    }

    return {
      ...token,
      accessToken: refreshed.access_token as string,
      refreshToken: (refreshed.refresh_token as string | undefined) ?? token.refreshToken,
      accessTokenExpires: Date.now() + Number(refreshed.expires_in ?? 3600) * 1000,
    };
  } catch {
    return { ...token, error: "RefreshAccessTokenError" as const };
  }
}

export const { handlers, auth, signIn, signOut } = NextAuth({
  trustHost: true,
  session: {
    strategy: "jwt",
    maxAge: 8 * 60 * 60,
  },
  pages: {
    signIn: "/login",
  },
  providers: [
    MicrosoftEntraID({
      clientId: process.env.AZURE_AD_CLIENT_ID,
      clientSecret: process.env.AZURE_AD_CLIENT_SECRET,
      issuer: tenantId ? `https://login.microsoftonline.com/${tenantId}/v2.0` : undefined,
      authorization: {
        params: {
          scope: GRAPH_SCOPES,
          prompt: "select_account",
        },
      },
    }),
  ],
  callbacks: {
    async jwt({ token, account, profile }) {
      if (account) {
        token.accessToken = account.access_token ?? undefined;
        token.refreshToken = account.refresh_token ?? undefined;
        token.accessTokenExpires =
          typeof account.expires_at === "number"
            ? account.expires_at * 1000
            : Date.now() + Number(account.expires_in ?? 3600) * 1000;
        token.sub = token.sub ?? profile?.sub ?? undefined;
      }

      if (token.accessToken && token.accessTokenExpires && Date.now() < token.accessTokenExpires - 60_000) {
        return token;
      }

      return refreshAccessToken(token);
    },
    async session({ session, token }) {
      if (token.sub) session.user.id = token.sub;
      session.error = token.error;
      return session;
    },
  },
});
