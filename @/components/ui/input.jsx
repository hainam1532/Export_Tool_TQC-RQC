import * as React from "react"

import { cn } from "@/lib/utils"

const Input = React.forwardRef(({ className, type, ...props }, ref) => {
  return (
    (<input
      type={type}
      className={cn(
        "tailwind.config.jsflex tailwind.config.jsh-10 tailwind.config.jsw-full tailwind.config.jsrounded-md tailwind.config.jsborder tailwind.config.jsborder-slate-200 tailwind.config.jsbg-white tailwind.config.jspx-3 tailwind.config.jspy-2 tailwind.config.jstext-sm tailwind.config.jsring-offset-white file:tailwind.config.jsborder-0 file:tailwind.config.jsbg-transparent file:tailwind.config.jstext-sm file:tailwind.config.jsfont-medium placeholder:tailwind.config.jstext-slate-500 focus-visible:tailwind.config.jsoutline-none focus-visible:tailwind.config.jsring-2 focus-visible:tailwind.config.jsring-slate-950 focus-visible:tailwind.config.jsring-offset-2 disabled:tailwind.config.jscursor-not-allowed disabled:tailwind.config.jsopacity-50 dark:tailwind.config.jsborder-slate-800 dark:tailwind.config.jsbg-slate-950 dark:tailwind.config.jsring-offset-slate-950 dark:placeholder:tailwind.config.jstext-slate-400 dark:focus-visible:tailwind.config.jsring-slate-300",
        className
      )}
      ref={ref}
      {...props} />)
  );
})
Input.displayName = "Input"

export { Input }
