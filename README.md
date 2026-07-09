# Planner Importer

Aplicación web para importar tareas desde archivos Excel (`.xlsx`) hacia Microsoft Planner usando Microsoft Graph API v1.0, permisos delegados y sesión en contexto del usuario autenticado.

## Requisitos

- Node.js 22 o superior.
- pnpm 9 o superior.
- Cuenta corporativa o escolar Microsoft 365 con acceso a Planner Basic.
- App Registration en Microsoft Entra ID.

## Instalación

```bash
pnpm install
cp .env.example .env.local
```

Completa `.env.local` con los valores reales de Entra ID y un `AUTH_SECRET` fuerte.

## Variables de entorno

```env
AZURE_AD_CLIENT_ID=
AZURE_AD_CLIENT_SECRET=
AZURE_AD_TENANT_ID=
AUTH_SECRET=
AUTH_URL=http://localhost:3000
NEXT_PUBLIC_APP_NAME=Planner Importer
APP_TIMEZONE=America/Monterrey
MAX_UPLOAD_SIZE_MB=10
MAX_IMPORT_ROWS=2000
IMPORT_CONCURRENCY=5
IMPORT_MAX_RETRIES=4
```

No subas secretos reales al repositorio.

## Microsoft Entra ID

1. Entra a Microsoft Entra admin center.
2. Ve a **Identity > Applications > App registrations**.
3. Crea una nueva aplicación para cuentas de tu organización.
4. Agrega una Redirect URI tipo Web:

```text
http://localhost:3000/api/auth/callback/microsoft-entra-id
```

Para producción usa:

```text
https://tu-dominio.com/api/auth/callback/microsoft-entra-id
```

5. Crea un Client Secret y copia su valor a `AZURE_AD_CLIENT_SECRET`.
6. Configura permisos delegados de Microsoft Graph:

```text
openid
profile
offline_access
User.Read
User.ReadBasic.All
Tasks.ReadWrite
```

7. Otorga consentimiento de administrador si tu tenant lo requiere.

## Ejecución

```bash
pnpm dev
```

Build de producción:

```bash
pnpm build
pnpm start
```

Verificación local:

```bash
pnpm lint
pnpm test
pnpm build
```

## Excel esperado

La primera hoja debe incluir exactamente estas columnas, con tolerancia a mayúsculas/minúsculas y espacios accidentales:

```text
Titulo
Responsable
fecha de inicio
fecha de vencimiento
tareas
etiqueta
```

- `Titulo`: obligatorio; se mapea a `plannerTask.title`.
- `Responsable`: correo o UPN; soporta múltiples responsables separados por `;` o `,`.
- `fecha de inicio`: acepta Date de Excel, serial Excel, `DD/MM/YYYY` o ISO.
- `fecha de vencimiento`: mismas reglas de fecha.
- `tareas`: se guarda como descripción de la tarea.
- `etiqueta`: se compara contra etiquetas configuradas en el Plan seleccionado.

Las fechas se interpretan en `America/Monterrey` para evitar desplazamientos de día al enviar a Graph.

## Arquitectura

- `src/app`: App Router, páginas y Route Handlers internos.
- `src/components`: UI, layout e importador.
- `src/lib/auth`: lectura segura de tokens server-side.
- `src/lib/graph`: cliente centralizado de Microsoft Graph y servicios Planner/usuarios.
- `src/lib/excel`: parser, schemas y reporte XLSX.
- `src/lib/dates`: parsing y serialización de fechas para Planner.
- `src/services/import-service.ts`: validación, duplicados, concurrencia e importación.
- `src/types`: contratos compartidos.

## Endpoints internos

```text
GET  /api/me
GET  /api/planner/plans
GET  /api/planner/plans/[planId]/buckets
GET  /api/planner/plans/[planId]/labels
POST /api/excel/parse
POST /api/import/validate
POST /api/import/start
POST /api/import/report
```

`/api/import/start` devuelve NDJSON streaming con eventos reales de progreso.

## Seguridad

- Los access tokens y refresh tokens se guardan solo en JWT cifrado de Auth.js.
- La sesión pública no expone tokens al navegador.
- Las llamadas a Microsoft Graph ocurren únicamente en Route Handlers `runtime = "nodejs"`.
- Se validan extensión, MIME, firma ZIP XLSX, tamaño máximo y máximo de filas.
- No se usan Application Permissions ni Power Automate.

## Limitaciones conocidas

- Solo se soportan archivos `.xlsx`; no se ejecutan macros.
- La importación es streaming directo; si el navegador se desconecta, no hay job persistente para reanudar.
- Microsoft Graph Planner API v1.0 no soporta Planner Premium; usa Planes Basic.
- La verificación real end-to-end requiere credenciales Entra válidas y consentimiento de Graph en el tenant.
