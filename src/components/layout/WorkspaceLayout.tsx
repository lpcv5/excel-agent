import { useEffect, useRef } from "react"
import { Outlet } from "react-router"
import type { PanelImperativeHandle } from "react-resizable-panels"
import { LeftSidebar } from "./LeftSidebar"
import { RightSidebar } from "./RightSidebar"
import { useLayout } from "@/hooks/useLayout"
import {
  ResizablePanelGroup,
  ResizablePanel,
  ResizableHandle,
} from "@/components/ui/resizable"

export function WorkspaceLayout() {
  const { leftSidebarOpen, rightSidebarOpen, setLeftSidebarOpen, setRightSidebarOpen } = useLayout()

  const leftPanelRef = useRef<PanelImperativeHandle>(null)
  const rightPanelRef = useRef<PanelImperativeHandle>(null)

  useEffect(() => {
    const panel = leftPanelRef.current
    if (!panel) return
    if (leftSidebarOpen) panel.expand()
    else panel.collapse()
  }, [leftSidebarOpen])

  useEffect(() => {
    const panel = rightPanelRef.current
    if (!panel) return
    if (rightSidebarOpen) panel.expand()
    else panel.collapse()
  }, [rightSidebarOpen])

  return (
    <ResizablePanelGroup
      orientation="horizontal"
      id="workspace-layout"
      className="h-full"
    >
      <ResizablePanel
        panelRef={leftPanelRef}
        defaultSize="20%"
        minSize="14%"
        maxSize="30%"
        collapsible
        collapsedSize="0%"
        onResize={(size) => {
          const collapsed = size.asPercentage === 0
          if (collapsed && leftSidebarOpen) setLeftSidebarOpen(false)
          if (!collapsed && !leftSidebarOpen) setLeftSidebarOpen(true)
        }}
        className="overflow-hidden"
      >
        <LeftSidebar />
      </ResizablePanel>
      <ResizableHandle />
      <ResizablePanel defaultSize="60%" minSize="40%" className="h-full">
        <Outlet />
      </ResizablePanel>
      <ResizableHandle />
      <ResizablePanel
        panelRef={rightPanelRef}
        defaultSize="15%"
        minSize="13%"
        maxSize="24%"
        collapsible
        collapsedSize="0%"
        onResize={(size) => {
          const collapsed = size.asPercentage === 0
          if (collapsed && rightSidebarOpen) setRightSidebarOpen(false)
          if (!collapsed && !rightSidebarOpen) setRightSidebarOpen(true)
        }}
        className="overflow-hidden"
      >
        <RightSidebar />
      </ResizablePanel>
    </ResizablePanelGroup>
  )
}
