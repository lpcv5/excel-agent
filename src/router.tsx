import { createMemoryRouter, redirect } from "react-router"
import { AppLayout } from "@/components/layout/AppLayout"
import { WorkspaceLayout } from "@/components/layout/WorkspaceLayout"
import { WelcomeScreen } from "@/components/project/WelcomeScreen"
import { CreateProjectScreen } from "@/components/project/CreateProjectScreen"
import { MainContent } from "@/components/layout/MainContent"
import { projects } from "@/services/api"
import { useProjectStore } from "@/stores/projectStore"
import { useFileTreeStore } from "@/stores/fileTreeStore"

async function rootLoader() {
  const { project } = await projects.getCurrent()
  if (project) {
    useProjectStore.getState().setCurrentProject(project)
    useFileTreeStore.getState().setRootPath(project.path)
    return redirect("/project")
  }
  return null
}

export const router = createMemoryRouter([
  {
    path: "/",
    element: <AppLayout />,
    HydrateFallback: () => null,
    children: [
      { index: true, loader: rootLoader, element: <WelcomeScreen /> },
      { path: "project/new", element: <CreateProjectScreen /> },
      {
        path: "project",
        element: <WorkspaceLayout />,
        children: [
          { index: true, element: <MainContent /> },
        ],
      },
    ],
  },
])
