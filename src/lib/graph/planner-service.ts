import type {
  GraphPlannerTask,
  GraphPlannerTaskDetails,
  PlannerBucket,
  PlannerCategoryDescriptions,
  PlannerLabel,
  PlannerLabelKey,
  PlannerPlan,
  PlannerPlanDetails,
} from "@/types/planner";

import { normalizeTextKey } from "@/lib/excel/schema";

import { graphFetch, graphFetchCollection } from "./graph-client";
import { GraphApiError } from "./graph-errors";

const CATEGORY_KEYS = Array.from({ length: 25 }, (_, index) => `category${index + 1}` as PlannerLabelKey);

export function labelsFromDescriptions(categoryDescriptions: PlannerCategoryDescriptions = {}) {
  return CATEGORY_KEYS.flatMap((key) => {
    const name = categoryDescriptions[key]?.trim();
    if (!name) return [];
    return [{ key, name, normalizedName: normalizeTextKey(name) } satisfies PlannerLabel];
  });
}

export function labelMapFromDescriptions(categoryDescriptions: PlannerCategoryDescriptions = {}) {
  return new Map(labelsFromDescriptions(categoryDescriptions).map((label) => [label.normalizedName, label.key]));
}

export class PlannerService {
  constructor(private readonly token: string) {}

  async getMe() {
    return graphFetch<{ id: string; displayName: string; mail?: string; userPrincipalName?: string }>({
      path: "/me?$select=id,displayName,mail,userPrincipalName",
      token: this.token,
    });
  }

  async listPlans() {
    const plans = await graphFetchCollection<PlannerPlan>("/me/planner/plans", this.token);
    return plans.map((plan) => ({ id: plan.id, title: plan.title })).sort((a, b) => a.title.localeCompare(b.title));
  }

  async listBuckets(planId: string) {
    const buckets = await graphFetchCollection<PlannerBucket>(`/planner/plans/${planId}/buckets`, this.token);
    return buckets
      .map((bucket) => ({ id: bucket.id, name: bucket.name, planId: bucket.planId }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  async getPlanDetails(planId: string) {
    return graphFetch<PlannerPlanDetails>({
      path: `/planner/plans/${planId}/details`,
      token: this.token,
    });
  }

  async getLabels(planId: string) {
    const details = await this.getPlanDetails(planId);
    return labelsFromDescriptions(details.categoryDescriptions);
  }

  async ensureLabels(planId: string, missingLabels: string[]) {
    const uniqueMissing = [...new Set(missingLabels.map(normalizeTextKey))].filter(Boolean);
    if (uniqueMissing.length === 0) return this.getPlanDetails(planId);

    let details = await this.getPlanDetails(planId);
    const descriptions = { ...(details.categoryDescriptions ?? {}) };
    const existing = labelMapFromDescriptions(descriptions);
    const created: Record<string, string> = {};

    for (const normalizedLabel of uniqueMissing) {
      if (existing.has(normalizedLabel)) continue;
      const emptyKey = CATEGORY_KEYS.find((key) => !descriptions[key]?.trim());
      if (!emptyKey) {
        throw new GraphApiError({
          status: 400,
          code: "NoAvailablePlannerLabels",
          message: `No hay categorías disponibles para crear la etiqueta "${normalizedLabel}".`,
        });
      }
      descriptions[emptyKey] = missingLabels.find((label) => normalizeTextKey(label) === normalizedLabel)?.trim() ?? normalizedLabel;
      existing.set(normalizedLabel, emptyKey);
      created[emptyKey] = descriptions[emptyKey] ?? "";
    }

    if (Object.keys(created).length === 0) return details;

    const etag = details["@odata.etag"];
    if (!etag) {
      throw new GraphApiError({
        status: 412,
        code: "MissingPlanDetailsEtag",
        message: "No se pudo obtener el ETag de las etiquetas del Plan.",
      });
    }

    try {
      await graphFetch<void>({
        path: `/planner/plans/${planId}/details`,
        method: "PATCH",
        token: this.token,
        headers: {
          "If-Match": etag,
        },
        body: {
          categoryDescriptions: created,
        },
      });
    } catch (error) {
      if (!(error instanceof GraphApiError) || (error.status !== 409 && error.status !== 412)) throw error;
      details = await this.getPlanDetails(planId);
      const refreshedDescriptions = { ...(details.categoryDescriptions ?? {}) };
      const refreshedExisting = labelMapFromDescriptions(refreshedDescriptions);
      const retryPatch: Record<string, string> = {};

      for (const normalizedLabel of uniqueMissing) {
        if (refreshedExisting.has(normalizedLabel)) continue;
        const emptyKey = CATEGORY_KEYS.find((key) => !refreshedDescriptions[key]?.trim());
        if (!emptyKey) throw error;
        const labelName = missingLabels.find((label) => normalizeTextKey(label) === normalizedLabel)?.trim() ?? normalizedLabel;
        retryPatch[emptyKey] = labelName;
        refreshedDescriptions[emptyKey] = labelName;
      }

      if (Object.keys(retryPatch).length > 0) {
        await graphFetch<void>({
          path: `/planner/plans/${planId}/details`,
          method: "PATCH",
          token: this.token,
          headers: {
            "If-Match": details["@odata.etag"] ?? "",
          },
          body: {
            categoryDescriptions: retryPatch,
          },
        });
      }
    }

    return this.getPlanDetails(planId);
  }

  async listBucketTasks(bucketId: string) {
    return graphFetchCollection<GraphPlannerTask>(`/planner/buckets/${bucketId}/tasks`, this.token);
  }

  async createTask(body: Record<string, unknown>) {
    return graphFetch<GraphPlannerTask>({
      path: "/planner/tasks",
      method: "POST",
      token: this.token,
      body,
    });
  }

  async getTaskDetails(taskId: string) {
    return graphFetch<GraphPlannerTaskDetails>({
      path: `/planner/tasks/${taskId}/details`,
      token: this.token,
    });
  }

  async updateTaskDetailsDescription(taskId: string, description: string) {
    const applyPatch = async (details: GraphPlannerTaskDetails) => {
      const etag = details["@odata.etag"];
      if (!etag) {
        throw new GraphApiError({
          status: 412,
          code: "MissingTaskDetailsEtag",
          message: "No se pudo obtener el ETag de los detalles de la tarea.",
        });
      }

      await graphFetch<void>({
        path: `/planner/tasks/${taskId}/details`,
        method: "PATCH",
        token: this.token,
        headers: {
          "If-Match": etag,
        },
        body: {
          description,
          previewType: "description",
        },
      });
    };

    try {
      await applyPatch(await this.getTaskDetails(taskId));
    } catch (error) {
      if (!(error instanceof GraphApiError) || (error.status !== 409 && error.status !== 412)) throw error;
      await applyPatch(await this.getTaskDetails(taskId));
    }
  }
}
