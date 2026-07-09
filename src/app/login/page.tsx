import { redirect } from "next/navigation";
import Image from "next/image";
import { BarChart3, FileSpreadsheet, ListChecks } from "lucide-react";

import { auth, signIn } from "@/auth";
import { Button } from "@/components/ui/button";

export default async function LoginPage() {
  const session = await auth();
  if (session) redirect("/dashboard");

  return (
    <main className="min-h-screen bg-[#F7FAFC] text-[#172033]">
      <div className="mx-auto grid min-h-screen max-w-6xl items-center gap-10 px-5 py-10 lg:grid-cols-[0.95fr_1.05fr]">
        <section className="space-y-8">
          <div>
            <div className="mb-6 inline-flex rounded-lg bg-[#0D47A1] px-5 py-4 shadow-sm">
              <Image
                src="/gis.png"
                alt="Grupo Industrial Saltillo"
                width={214}
                height={71}
                priority
                className="h-12 w-auto object-contain"
              />
            </div>
            <div className="mb-5 flex w-fit items-center gap-2 rounded-md border border-blue-100 bg-white px-3 py-2 text-sm font-medium text-[#0D47A1] shadow-sm">
              Grupo Industrial Saltillo
            </div>
            <h1 className="text-4xl font-semibold tracking-normal text-[#172033] sm:text-5xl">Planner Importer</h1>
            <p className="mt-5 max-w-xl text-lg leading-8 text-slate-600">
              Importa tareas desde Excel a Microsoft Planner de forma rápida, segura y controlada.
            </p>
          </div>
          <form
            action={async () => {
              "use server";
              await signIn("microsoft-entra-id", { redirectTo: "/dashboard" });
            }}
          >
            <Button type="submit" className="h-11 px-5">
              Iniciar sesión con Microsoft
            </Button>
          </form>
        </section>

        <section className="relative overflow-hidden rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
          <div className="grid gap-4 sm:grid-cols-2">
            <div className="rounded-lg border border-slate-200 bg-[#EAF3FF] p-5">
              <FileSpreadsheet className="mb-8 h-8 w-8 text-[#1565C0]" />
              <p className="text-sm font-semibold">Excel validado</p>
              <p className="mt-2 text-sm text-slate-600">Columnas, responsables, fechas y etiquetas.</p>
            </div>
            <div className="rounded-lg border border-slate-200 bg-white p-5">
              <ListChecks className="mb-8 h-8 w-8 text-[#16A34A]" />
              <p className="text-sm font-semibold">Planner real</p>
              <p className="mt-2 text-sm text-slate-600">Planes y Buckets accesibles para el usuario.</p>
            </div>
            <div className="rounded-lg border border-slate-200 bg-white p-5 sm:col-span-2">
              <div className="mb-5 flex items-center justify-between">
                <p className="text-sm font-semibold">Progreso de importación</p>
                <BarChart3 className="h-5 w-5 text-[#1565C0]" />
              </div>
              <div className="space-y-3">
                <div className="h-3 w-11/12 rounded-full bg-[#1565C0]" />
                <div className="h-3 w-8/12 rounded-full bg-[#EAF3FF]" />
                <div className="h-3 w-10/12 rounded-full bg-slate-100" />
              </div>
            </div>
          </div>
        </section>
      </div>
    </main>
  );
}
