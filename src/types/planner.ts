export type PlannerPlan = {
  id: string;
  title: string;
};

export type PlannerBucket = {
  id: string;
  name: string;
  planId: string;
};

export type PlannerLabelKey = `category${number}`;

export type PlannerLabel = {
  key: PlannerLabelKey;
  name: string;
  normalizedName: string;
};

export type PlannerCategoryDescriptions = Partial<Record<PlannerLabelKey, string | null>>;

export type PlannerPlanDetails = {
  id?: string;
  categoryDescriptions?: PlannerCategoryDescriptions;
  "@odata.etag"?: string;
};

export type ResolvedGraphUser = {
  id: string;
  displayName: string;
  mail?: string | null;
  userPrincipalName: string;
};

export type GraphCollection<T> = {
  value: T[];
  "@odata.nextLink"?: string;
};

export type GraphPlannerTask = {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  dueDateTime?: string | null;
  startDateTime?: string | null;
  assignments?: Record<string, unknown>;
  appliedCategories?: Record<string, boolean>;
  "@odata.etag"?: string;
};

export type GraphPlannerTaskDetails = {
  id: string;
  description?: string;
  previewType?: string;
  "@odata.etag"?: string;
};
