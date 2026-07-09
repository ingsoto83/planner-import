import type { ResolvedGraphUser } from "@/types/planner";

import { graphFetch } from "./graph-client";
import { isGraphApiError } from "./graph-errors";

function normalizePrincipal(value: string) {
  return value.trim().toLowerCase();
}

export class UsersService {
  private cache = new Map<string, ResolvedGraphUser | null>();

  constructor(private readonly token: string) {}

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
      if (isGraphApiError(error) && error.status === 404) {
        this.cache.set(normalized, null);
        return null;
      }
      throw error;
    }
  }
}
