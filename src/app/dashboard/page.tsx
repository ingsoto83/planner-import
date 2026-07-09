import { redirect } from "next/navigation";

import { auth } from "@/auth";
import { ImportDashboard } from "@/components/importer/import-dashboard";

export const dynamic = "force-dynamic";

export default async function DashboardPage() {
  const session = await auth();
  if (!session) redirect("/login");

  return <ImportDashboard initialUser={{ name: session.user?.name, email: session.user?.email }} />;
}
