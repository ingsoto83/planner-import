import type { ResolvedGraphUser } from "@/types/planner";

import { graphFetch } from "./graph-client";
import { isGraphApiError } from "./graph-errors";

function normalizePrincipal(value: string) {
  return value.trim().toLowerCase();
}

export class UsersService {
  private cache = new Map<string, ResolvedGraphUser | null>();
  private currentUser?: ResolvedGraphUser | null;

  constructor(private readonly token: string) {}

  private async getCurrentUser() {
    if (this.currentUser !== undefined) return this.currentUser;

    try {
      this.currentUser = await graphFetch<ResolvedGraphUser>({
        path: "/me?$select=id,displayName,mail,userPrincipalName",
        token: this.token,
      });
    } catch {
      this.currentUser = null;
    }

    return this.currentUser;
  }

  private async resolveByMailOrUpn(normalized: string) {
    const escaped = normalized.replace(/'/g, "''");
    const filter = encodeURIComponent(`mail eq '${escaped}' or userPrincipalName eq '${escaped}'`);

    try {
      const response = await graphFetch<{ value: ResolvedGraphUser[] }>({
        path: `/users?$select=id,displayName,mail,userPrincipalName&$filter=${filter}&$top=1`,
        token: this.token,
      });
      return response.value[0] ?? null;
    } catch (error) {
      if (isGraphApiError(error) && [400, 403, 404].includes(error.status)) return null;
      throw error;
    }
  }

  async resolveUser(principal: string) {
    const normalized = normalizePrincipal(principal);
    if (!normalized) return null;
    if (this.cache.has(normalized)) return this.cache.get(normalized) ?? null;

    try {
      const user = await graphFetch<ResolvedGraphUser>({
        path: `/users/${encodeURIComponent(normalized)}?$select=id,displayName,mail,userPrincipalName`,
        token: this.token,
      });
      this.cache.set(normalized, user);
      return user;
    } catch (error) {
      if (isGraphApiError(error) && [400, 404].includes(error.status)) {
        const currentUser = await this.getCurrentUser();
        const currentUserMatches =
          normalizePrincipal(currentUser?.mail ?? "") === normalized ||
          normalizePrincipal(currentUser?.userPrincipalName ?? "") === normalized;

        const fallbackUser = currentUserMatches ? currentUser : await this.resolveByMailOrUpn(normalized);
        this.cache.set(normalized, fallbackUser);
        return fallbackUser;
      }
      throw error;
    }
  }
}
