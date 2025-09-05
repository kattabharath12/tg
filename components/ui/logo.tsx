
import Image from "next/image"
import { cn } from "@/lib/utils"

interface LogoProps {
  className?: string
  size?: "sm" | "md" | "lg" | "xl"
  showText?: boolean
  textClassName?: string
}

const sizeClasses = {
  sm: "h-6 w-6",
  md: "h-8 w-8", 
  lg: "h-12 w-12",
  xl: "h-16 w-16"
}

export function Logo({ 
  className, 
  size = "md", 
  showText = true, 
  textClassName 
}: LogoProps) {
  return (
    <div className={cn("flex items-center space-x-2", className)}>
      <div className={cn("relative", sizeClasses[size])}>
        <Image
          src="/taxgrok-logo.jpeg"
          alt="TaxGrok.AI Logo"
          fill
          className="object-contain rounded-md"
          priority
        />
      </div>
      {showText && (
        <span className={cn(
          "font-bold text-gray-900",
          size === "sm" && "text-lg",
          size === "md" && "text-2xl", 
          size === "lg" && "text-3xl",
          size === "xl" && "text-4xl",
          textClassName
        )}>
          TaxGrok.AI
        </span>
      )}
    </div>
  )
}
