"use client"

import * as React from "react"
import * as CheckboxPrimitive from "@radix-ui/react-checkbox"
import { Check } from "lucide-react"

import { cn } from "@/lib/utils"

const Checkbox = React.forwardRef(({ className, ...props }, ref) => (
  <CheckboxPrimitive.Root
    ref={ref}
    className={cn(
      "tailwind.config.jspeer tailwind.config.jsh-4 tailwind.config.jsw-4 tailwind.config.jsshrink-0 tailwind.config.jsrounded-sm tailwind.config.jsborder tailwind.config.jsborder-slate-200 tailwind.config.jsborder-slate-900 tailwind.config.jsring-offset-white focus-visible:tailwind.config.jsoutline-none focus-visible:tailwind.config.jsring-2 focus-visible:tailwind.config.jsring-slate-950 focus-visible:tailwind.config.jsring-offset-2 disabled:tailwind.config.jscursor-not-allowed disabled:tailwind.config.jsopacity-50 data-[state=checked]:tailwind.config.jsbg-slate-900 data-[state=checked]:tailwind.config.jstext-slate-50 dark:tailwind.config.jsborder-slate-800 dark:tailwind.config.jsborder-slate-50 dark:tailwind.config.jsring-offset-slate-950 dark:focus-visible:tailwind.config.jsring-slate-300 dark:data-[state=checked]:tailwind.config.jsbg-slate-50 dark:data-[state=checked]:tailwind.config.jstext-slate-900",
      className
    )}
    {...props}>
    <CheckboxPrimitive.Indicator
      className={cn(
        "tailwind.config.jsflex tailwind.config.jsitems-center tailwind.config.jsjustify-center tailwind.config.jstext-current"
      )}>
      <Check className="tailwind.config.jsh-4 tailwind.config.jsw-4" />
    </CheckboxPrimitive.Indicator>
  </CheckboxPrimitive.Root>
))
Checkbox.displayName = CheckboxPrimitive.Root.displayName

export { Checkbox }
