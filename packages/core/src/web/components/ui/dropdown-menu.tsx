"use client"

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as React from "react"
import * as DropdownMenuPrimitive from "@radix-ui/react-dropdown-menu"
import {
  Check,
  ChevronRight,
  Circle,
} from "../office-ribbon/icons/houseIcons";

import { cn } from "../../lib/utils"
import { usePortalHostContainer } from "./portal-host"












const DropdownMenuFocusContext = React.createContext<{
  trigger: React.MutableRefObject<HTMLElement | null>
  interactedOutside: React.MutableRefObject<boolean>
  modal: boolean
} | null>(null)

const DropdownMenu = ({ modal = true, ...props }: React.ComponentProps<typeof DropdownMenuPrimitive.Root>) => {
  const trigger = React.useRef<HTMLElement | null>(null)
  const interactedOutside = React.useRef(false)
  const value = React.useMemo(() => ({ trigger, interactedOutside, modal }), [modal])
  return (
    <DropdownMenuFocusContext.Provider value={value}>
      <DropdownMenuPrimitive.Root modal={modal} {...props} />
    </DropdownMenuFocusContext.Provider>
  )
}

/**
 * Radix opens a dropdown from the trigger's `pointerdown`, and Chromium still
 * delivers *pointer* events to a `<button disabled>` even though it suppresses
 * `mousedown` and `click` on one. Under `asChild` that left every disabled
 * ribbon split-button opening its own menu on a real click: the button face was
 * dead, the trigger behind it was not, and the menu appeared with every item
 * greyed out.
 *
 * Radix already refuses to open — pointer *and* keyboard — when the trigger
 * itself is `disabled`, so carry a single disabled child's own `disabled` up to
 * the trigger instead of re-deciding it at ~200 call sites. An explicit
 * `disabled` on the trigger always wins, and a child that is not disabled is
 * passed through exactly as before.
 */
const DropdownMenuTrigger = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.Trigger>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.Trigger>
>(({ asChild, children, disabled, ...props }, ref) => {
  const childDisabled =
    asChild && React.isValidElement<{ disabled?: unknown }>(children)
      ? Boolean(children.props.disabled)
      : false
  const focus = React.useContext(DropdownMenuFocusContext)
  const setRef = React.useCallback(
    (node: HTMLButtonElement | null) => {
      if (focus) focus.trigger.current = node
      if (typeof ref === "function") ref(node)
      else if (ref) ref.current = node
    },
    [focus, ref],
  )
  return (
    <DropdownMenuPrimitive.Trigger
      {...props}
      ref={setRef}
      asChild={asChild}
      disabled={disabled ?? childDisabled}
    >
      {children}
    </DropdownMenuPrimitive.Trigger>
  )
})
DropdownMenuTrigger.displayName = DropdownMenuPrimitive.Trigger.displayName

const DropdownMenuGroup = DropdownMenuPrimitive.Group

const DropdownMenuPortal = ({ container, ...props }: React.ComponentProps<typeof DropdownMenuPrimitive.Portal>) => {
  const host = usePortalHostContainer(container)
  return host.pending ? null : <DropdownMenuPrimitive.Portal {...props} container={host.container} />
}

const DropdownMenuSub = DropdownMenuPrimitive.Sub

const DropdownMenuRadioGroup = DropdownMenuPrimitive.RadioGroup

const DropdownMenuSubTrigger = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.SubTrigger>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.SubTrigger> & {
    inset?: boolean
  }
>(({ className, inset, children, ...props }, ref) => (
  <DropdownMenuPrimitive.SubTrigger
    ref={ref}
    className={cn(
      "flex cursor-default select-none items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none focus:bg-accent data-[state=open]:bg-accent [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0",
      inset && "pl-8",
      className
    )}
    {...props}
  >
    {children}
    <ChevronRight className="ml-auto" />
  </DropdownMenuPrimitive.SubTrigger>
))
DropdownMenuSubTrigger.displayName =
  DropdownMenuPrimitive.SubTrigger.displayName

/**
 * A submenu is portalled next to its parent menu instead of being rendered
 * inside it. The parent content is a bounded scroll box (`max-h` +
 * `overflow-y-auto overflow-x-hidden`), so a submenu placed inside it was
 * clipped at the parent's edge whenever it opened sideways -- the "menu inside
 * a menu is hidden by the big menu's container" defect that recurred across
 * the chart, table-borders, equation and window menus. Portalling through the
 * same host as `DropdownMenuContent` keeps the submenu in the parent's stacking
 * context (same host, later sibling, same z-index) and lets Radix position it
 * against the viewport rather than the scroll box. Nested submenus portal the
 * same way, one level each.
 */
const DropdownMenuSubContent = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.SubContent>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.SubContent> &
    Pick<React.ComponentProps<typeof DropdownMenuPrimitive.Portal>, "container">
>(({ className, container, ...props }, ref) => (
  <DropdownMenuPortal container={container}>
    <DropdownMenuPrimitive.SubContent
      ref={ref}
      className={cn(
        "z-50 min-w-[8rem] max-h-[var(--radix-dropdown-menu-content-available-height)] overflow-y-auto overflow-x-hidden rounded-md border bg-popover p-1 text-popover-foreground shadow-lg data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0 data-[state=closed]:zoom-out-95 data-[state=open]:zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=left]:slide-in-from-right-2 data-[side=right]:slide-in-from-left-2 data-[side=top]:slide-in-from-bottom-2 data-[state=open]:duration-150 data-[state=closed]:duration-[110ms] data-[state=open]:ease-[cubic-bezier(0.22,1,0.36,1)] data-[state=closed]:ease-[cubic-bezier(0.4,0,1,1)] origin-[--radix-dropdown-menu-content-transform-origin]",
        className
      )}
      {...props}
    />
  </DropdownMenuPortal>
))
DropdownMenuSubContent.displayName =
  DropdownMenuPrimitive.SubContent.displayName

/** Whether a closing menu leaves focus where it is: a field outside the menu
 * (not its trigger) the user has already moved into. */
export function menuCloseLeavesFocus(
  active: Element | null,
  trigger: HTMLElement | null,
  content: Node | null,
): boolean {
  return active instanceof HTMLElement && active !== active.ownerDocument.body && active.isConnected &&
    active !== trigger && !(content && content.contains(active)) &&
    active.matches("input, textarea, select, [contenteditable='true']")
}

const DropdownMenuContent = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.Content>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.Content> &
    Pick<React.ComponentProps<typeof DropdownMenuPrimitive.Portal>, "container">
>(({ className, sideOffset = 4, container, onFocusOutside, onCloseAutoFocus, onInteractOutside, ...props }, ref) => {
  const focus = React.useContext(DropdownMenuFocusContext)
  return (
  <DropdownMenuPortal container={container}>
    <DropdownMenuPrimitive.Content
      ref={ref}
      sideOffset={sideOffset}
      // Focus goes back to the trigger without scrolling (see
      // `DropdownMenuFocusContext`); Radix's own return, a scrolling
      // `focus()`, is pre-empted by the `preventDefault`.
      onCloseAutoFocus={(event) => {
        onCloseAutoFocus?.(event)
        if (event.defaultPrevented || !focus) return
        event.preventDefault()
        // Focus the user already took somewhere else while the menu played
        // its exit (a pane's name double-clicked for renaming the instant
        // Selection Pane opened) stays there: taken back to the trigger, the
        
        const movedOn = menuCloseLeavesFocus(
          typeof document === "undefined" ? null : document.activeElement,
          focus.trigger.current,
          event.currentTarget instanceof Node ? event.currentTarget : null,
        )
        if (!focus.interactedOutside.current && !movedOn) focus.trigger.current?.focus({ preventScroll: true })
        focus.interactedOutside.current = false
      }}
      onInteractOutside={(event) => {
        onInteractOutside?.(event)
        if (!focus) return
        const original = event.detail.originalEvent as PointerEvent | FocusEvent
        const button = "button" in original ? original.button : -1
        const rightClick = button === 2 || (button === 0 && "ctrlKey" in original && original.ctrlKey === true)
        if (!focus.modal || rightClick) focus.interactedOutside.current = true
      }}
      // A non-modal menu dismisses when focus leaves it. Focus landing on a
      // menu TRIGGER is not the user leaving: Radix returns focus to a closed
      // menu's trigger a tick after it closes, and that late return must not
      // take down a menu (the same one or the next) opened in between. A real
      // move to another trigger is a pointer press, which dismisses on its
      // own, and Tab never leaves a menu. Modal menus trap focus and never
      // get here.
      onFocusOutside={(event) => {
        onFocusOutside?.(event)
        const target = event.target
        if (
          !event.defaultPrevented &&
          target instanceof HTMLElement &&
          target.hasAttribute("aria-haspopup")
        ) {
          event.preventDefault()
        }
      }}
      className={cn(
        "z-50 max-h-[var(--radix-dropdown-menu-content-available-height)] min-w-[8rem] overflow-y-auto overflow-x-hidden rounded-md border bg-popover p-1 text-popover-foreground shadow-md data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0 data-[state=closed]:zoom-out-95 data-[state=open]:zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=left]:slide-in-from-right-2 data-[side=right]:slide-in-from-left-2 data-[side=top]:slide-in-from-bottom-2 data-[state=open]:duration-150 data-[state=closed]:duration-[110ms] data-[state=open]:ease-[cubic-bezier(0.22,1,0.36,1)] data-[state=closed]:ease-[cubic-bezier(0.4,0,1,1)] origin-[--radix-dropdown-menu-content-transform-origin]",
        className
      )}
      {...props}
    />
  </DropdownMenuPortal>
  )
})
DropdownMenuContent.displayName = DropdownMenuPrimitive.Content.displayName








function ignoredWhileMenuCloses<E extends Event>(handler: ((event: E) => void) | undefined) {
  return (event: E) => {
    const item = event.currentTarget instanceof Element ? event.currentTarget : null
    if (item?.closest('[role="menu"]')?.getAttribute("data-state") === "closed") {
      event.preventDefault()
      return
    }
    handler?.(event)
  }
}

const DropdownMenuItem = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.Item>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.Item> & {
    inset?: boolean
  }
>(({ className, inset, onSelect, ...props }, ref) => (
  <DropdownMenuPrimitive.Item
    ref={ref}
    onSelect={ignoredWhileMenuCloses(onSelect)}
    className={cn(
      "relative flex cursor-default select-none items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50 [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0",
      inset && "pl-8",
      className
    )}
    {...props}
  />
))
DropdownMenuItem.displayName = DropdownMenuPrimitive.Item.displayName

const DropdownMenuCheckboxItem = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.CheckboxItem>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.CheckboxItem>
>(({ className, children, checked, ...props }, ref) => (
  <DropdownMenuPrimitive.CheckboxItem
    ref={ref}
    className={cn(
      "relative flex cursor-default select-none items-center rounded-sm py-1.5 pl-8 pr-2 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50",
      className
    )}
    checked={checked}
    {...props}
  >
    <span className="absolute left-2 flex h-3.5 w-3.5 items-center justify-center">
      <DropdownMenuPrimitive.ItemIndicator>
        <Check className="h-4 w-4" />
      </DropdownMenuPrimitive.ItemIndicator>
    </span>
    {children}
  </DropdownMenuPrimitive.CheckboxItem>
))
DropdownMenuCheckboxItem.displayName =
  DropdownMenuPrimitive.CheckboxItem.displayName

const DropdownMenuRadioItem = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.RadioItem>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.RadioItem>
>(({ className, children, ...props }, ref) => (
  <DropdownMenuPrimitive.RadioItem
    ref={ref}
    className={cn(
      "relative flex cursor-default select-none items-center rounded-sm py-1.5 pl-8 pr-2 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50",
      className
    )}
    {...props}
  >
    <span className="absolute left-2 flex h-3.5 w-3.5 items-center justify-center">
      <DropdownMenuPrimitive.ItemIndicator>
        <Circle className="h-2 w-2 fill-current" />
      </DropdownMenuPrimitive.ItemIndicator>
    </span>
    {children}
  </DropdownMenuPrimitive.RadioItem>
))
DropdownMenuRadioItem.displayName = DropdownMenuPrimitive.RadioItem.displayName

const DropdownMenuLabel = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.Label>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.Label> & {
    inset?: boolean
  }
>(({ className, inset, ...props }, ref) => (
  <DropdownMenuPrimitive.Label
    ref={ref}
    className={cn(
      "px-2 py-1.5 text-sm font-semibold",
      inset && "pl-8",
      className
    )}
    {...props}
  />
))
DropdownMenuLabel.displayName = DropdownMenuPrimitive.Label.displayName

const DropdownMenuSeparator = React.forwardRef<
  React.ElementRef<typeof DropdownMenuPrimitive.Separator>,
  React.ComponentPropsWithoutRef<typeof DropdownMenuPrimitive.Separator>
>(({ className, ...props }, ref) => (
  <DropdownMenuPrimitive.Separator
    ref={ref}
    className={cn("-mx-1 my-1 h-px bg-muted", className)}
    {...props}
  />
))
DropdownMenuSeparator.displayName = DropdownMenuPrimitive.Separator.displayName

const DropdownMenuShortcut = ({
  className,
  ...props
}: React.HTMLAttributes<HTMLSpanElement>) => {
  return (
    <span
      className={cn("ml-auto text-xs tracking-widest opacity-60", className)}
      {...props}
    />
  )
}
DropdownMenuShortcut.displayName = "DropdownMenuShortcut"

export {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuCheckboxItem,
  DropdownMenuRadioItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
  DropdownMenuGroup,
  DropdownMenuPortal,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuRadioGroup,
}
