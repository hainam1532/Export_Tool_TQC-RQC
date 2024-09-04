import * as React from "react"

import { cn } from "@/lib/utils"

const Card = React.forwardRef(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn(
      "tailwind.config.jsrounded-lg tailwind.config.jsborder tailwind.config.jsborder-slate-200 tailwind.config.jsbg-white tailwind.config.jstext-slate-950 tailwind.config.jsshadow-sm dark:tailwind.config.jsborder-slate-800 dark:tailwind.config.jsbg-slate-950 dark:tailwind.config.jstext-slate-50",
      className
    )}
    {...props} />
))
Card.displayName = "Card"

const CardHeader = React.forwardRef(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn(
      "tailwind.config.jsflex tailwind.config.jsflex-col tailwind.config.jsspace-y-1.5 tailwind.config.jsp-6",
      className
    )}
    {...props} />
))
CardHeader.displayName = "CardHeader"

const CardTitle = React.forwardRef(({ className, ...props }, ref) => (
  <h3
    ref={ref}
    className={cn(
      "tailwind.config.jstext-2xl tailwind.config.jsfont-semibold tailwind.config.jsleading-none tailwind.config.jstracking-tight",
      className
    )}
    {...props} />
))
CardTitle.displayName = "CardTitle"

const CardDescription = React.forwardRef(({ className, ...props }, ref) => (
  <p
    ref={ref}
    className={cn(
      "tailwind.config.jstext-sm tailwind.config.jstext-slate-500 dark:tailwind.config.jstext-slate-400",
      className
    )}
    {...props} />
))
CardDescription.displayName = "CardDescription"

const CardContent = React.forwardRef(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn("tailwind.config.jsp-6 tailwind.config.jspt-0", className)}
    {...props} />
))
CardContent.displayName = "CardContent"

const CardFooter = React.forwardRef(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn(
      "tailwind.config.jsflex tailwind.config.jsitems-center tailwind.config.jsp-6 tailwind.config.jspt-0",
      className
    )}
    {...props} />
))
CardFooter.displayName = "CardFooter"

export { Card, CardHeader, CardFooter, CardTitle, CardDescription, CardContent }
