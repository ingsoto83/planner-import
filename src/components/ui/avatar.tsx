import * as React from "react";

import { cn } from "@/lib/utils/cn";

export function Avatar({
  src,
  name,
  className,
}: {
  src?: string | null;
  name?: string | null;
  className?: string;
}) {
  const initials =
    name
      ?.split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((part) => part[0]?.toUpperCase())
      .join("") || "U";

  return (
    <div className={cn("flex h-9 w-9 items-center justify-center overflow-hidden rounded-full bg-[#EAF3FF] text-sm font-semibold text-[#0D47A1]", className)}>
      {src ? (
        // eslint-disable-next-line @next/next/no-img-element
        <img src={src} alt="" className="h-full w-full object-cover" />
      ) : (
        initials
      )}
    </div>
  );
}
