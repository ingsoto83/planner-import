import * as React from "react";
import { Check } from "lucide-react";

import { cn } from "@/lib/utils/cn";

export function Checkbox({
  checked,
  className,
  ...props
}: Omit<React.InputHTMLAttributes<HTMLInputElement>, "type">) {
  return (
    <span className={cn("relative inline-flex h-4 w-4 shrink-0", className)}>
      <input
        type="checkbox"
        checked={checked}
        className="peer h-4 w-4 appearance-none rounded border border-slate-300 bg-white transition checked:border-[#1565C0] checked:bg-[#1565C0] focus:outline-none focus:ring-2 focus:ring-[#EAF3FF]"
        {...props}
      />
      <Check className="pointer-events-none absolute left-0.5 top-0.5 hidden h-3 w-3 text-white peer-checked:block" />
    </span>
  );
}
