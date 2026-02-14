import { PanelLeft, PanelRight, Grid2x2, Settings } from "lucide-react";
import { useLocation } from "react-router";
import { useTranslation } from "react-i18next";
import { ThemeToggle } from "@/components/theme/ThemeToggle";
import { useLayout } from "@/hooks/useLayout";
import { Button } from "@/components/ui/button";
import { Menubar, MenubarContent, MenubarItem, MenubarMenu, MenubarTrigger } from "@/components/ui/menubar";

const isMac = navigator.userAgent.includes("Mac");

export function TitleBar() {
  const { t } = useTranslation();
  const menuItems = [
    { key: "file", label: t("titleBar.menu.file") },
    { key: "edit", label: t("titleBar.menu.edit") },
    { key: "view", label: t("titleBar.menu.view") },
    { key: "tools", label: t("titleBar.menu.tools") },
    { key: "help", label: t("titleBar.menu.help") },
  ];
  const { toggleLeftSidebar, toggleRightSidebar, settingsOpen, openSettings, closeSettings } = useLayout();
  const location = useLocation();
  const inProject = location.pathname.startsWith("/project");

  return (
    <header
      className="flex h-header shrink-0 items-center border-b bg-surface px-2 gap-0.5 select-none"
    >
      {/* macOS traffic lights 占位 */}
      {isMac && <div className="w-[68px] shrink-0" />}

      {/* Logo */}
      <div className="flex items-center gap-1.5 px-2 shrink-0">
        <Grid2x2 className="size-[18px] text-primary" />
        <span className="text-[13px] font-bold text-primary tracking-tight">ExcelAgent</span>
      </div>

      {/* 分隔线 */}
      <div className="w-px h-4 bg-border mx-1 shrink-0" />

      {/* 菜单栏 */}
      <Menubar className="h-auto border-0 bg-transparent p-0 shadow-none gap-px">
        {menuItems.map((item) => (
          <MenubarMenu key={item.key}>
            <MenubarTrigger className="px-2.5 py-1 rounded-sm text-xs text-muted-foreground hover:bg-surface-hover hover:text-foreground data-[state=open]:bg-surface-hover data-[state=open]:text-foreground cursor-default">
              {item.label}
            </MenubarTrigger>
            <MenubarContent>
              <MenubarItem>{t("titleBar.menu.noItems")}</MenubarItem>
            </MenubarContent>
          </MenubarMenu>
        ))}
      </Menubar>

      {/* 弹性空间 */}
      <div className="flex-1" />

      {/* 右侧功能按钮 */}
      <div className="flex items-center gap-0.5 shrink-0">
        <Button variant="ghost" size="icon" className="size-[30px] text-muted-foreground hover:text-primary"
          onClick={toggleLeftSidebar}
        >
          <PanelLeft className="size-4" />
        </Button>
        <Button variant="ghost" size="icon" className="size-[30px] text-muted-foreground hover:text-primary"
          onClick={toggleRightSidebar}
        >
          <PanelRight className="size-4" />
        </Button>
        <Button
          variant="ghost"
          size="icon"
          className={`size-[30px] hover:text-primary ${settingsOpen ? "text-primary" : "text-muted-foreground"}`}
          disabled={!inProject}
          onClick={settingsOpen ? closeSettings : openSettings}
        >
          <Settings className="size-4" />
        </Button>
        <ThemeToggle />
      </div>
    </header>
  );
}
