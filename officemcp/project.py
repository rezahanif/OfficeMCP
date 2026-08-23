"""
Microsoft Project CRUD operations — COM automation.

Uses win32com for live-app operations (Project must be running).
Lazy COM import follows the same pattern as Officer.py.

Security: COM operations limited to the active Project application.
No filesystem write except via Project's own Save/SaveAs.
"""
import os
import sys
from typing import Any


def _com():
    import win32com.client
    return win32com.client


def _pywintypes():
    import pywintypes
    return pywintypes


def _get_project_app():
    """Get or launch Microsoft Project application."""
    com = _com()
    try:
        app = com.GetActiveObject("MSProject.Application")
    except _pywintypes().com_error:
        app = com.Dispatch("MSProject.Application")
    app.Visible = True
    return app


def _resolve(path: str) -> str:
    """Resolve relative path against OfficeMCP root."""
    if os.path.isabs(path):
        return path
    root = os.environ.get("OFFICE_MCP_ROOT") or os.path.join(
        os.path.expanduser("~"), "@OfficeMCP"
    )
    return os.path.join(root, path)


# ---- Project CRUD ----


def msp_create(path: str) -> dict:
    """Create a new blank Project file."""
    try:
        app = _get_project_app()
        proj = app.Projects.Add()
        full_path = _resolve(path)
        proj.SaveAs(full_path)
        return {"success": True, "path": full_path, "name": proj.Name}
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_open(path: str) -> dict:
    """Open an existing .mpp project file."""
    try:
        app = _get_project_app()
        full_path = _resolve(path)
        proj = app.Projects.Open(full_path)
        return {
            "success": True,
            "path": full_path,
            "name": proj.Name,
            "tasks": proj.Tasks.Count,
            "resources": proj.Resources.Count,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_save(path: str | None = None) -> dict:
    """Save the active project. If path provided, SaveAs."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        if path:
            full_path = _resolve(path)
            proj.SaveAs(full_path)
        else:
            proj.Save()
        return {"success": True, "name": proj.Name}
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_close(save: bool = True) -> dict:
    """Close the active project."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        name = proj.Name
        if save:
            proj.Save()
        proj.Close()
        return {"success": True, "name": name}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- Task CRUD ----


def msp_add_task(
    name: str,
    duration: str | None = None,
    start: str | None = None,
    predecessors: str | None = None,
    notes: str | None = None,
) -> dict:
    """Add a task to the active project.

    Args:
        name: Task name
        duration: Duration string (e.g. "3d", "1w", "4h") — default "1d"
        start: Start date string (e.g. "2026-09-01") — default is project start
        predecessors: Predecessor task IDs (e.g. "1,3FS+2d")
        notes: Task notes
    """
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        task = proj.Tasks.Add(name)

        if duration:
            task.Duration = duration
        if start:
            task.Start = start
        if predecessors:
            task.Predecessors = predecessors
        if notes:
            task.Notes = notes

        return {
            "success": True,
            "id": task.UniqueID,
            "name": task.Name,
            "duration": str(task.Duration),
            "start": str(task.Start),
            "finish": str(task.Finish),
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_get_tasks(
    filter_name: str | None = None,
    field: str | None = None,
    max_results: int = 50,
) -> dict:
    """List tasks in the active project.

    Args:
        filter_name: Apply a named filter (e.g. "Incomplete Tasks")
        field: Sort by field name (e.g. "Start", "Finish", "Cost")
        max_results: Maximum tasks to return (default 50)
    """
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        tasks = proj.Tasks

        if filter_name:
            proj.TaskFilter = filter_name

        if field:
            proj.Sort(field, 1)

        result = []
        count = min(tasks.Count, max_results)
        for i in range(1, count + 1):
            t = tasks(i)
            result.append(
                {
                    "id": t.UniqueID,
                    "name": t.Name,
                    "duration": str(t.Duration),
                    "start": str(t.Start),
                    "finish": str(t.Finish),
                    "percent_complete": t.PercentComplete,
                    "cost": t.Cost,
                    "predecessors": t.Predecessors,
                }
            )

        return {"success": True, "count": len(result), "tasks": result}
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_update_task(
    task_id: int,
    name: str | None = None,
    duration: str | None = None,
    start: str | None = None,
    finish: str | None = None,
    percent_complete: int | None = None,
    predecessors: str | None = None,
    notes: str | None = None,
) -> dict:
    """Update an existing task by UniqueID."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        task = proj.Tasks.UniqueID(task_id)

        if name:
            task.Name = name
        if duration:
            task.Duration = duration
        if start:
            task.Start = start
        if finish:
            task.Finish = finish
        if percent_complete is not None:
            task.PercentComplete = percent_complete
        if predecessors:
            task.Predecessors = predecessors
        if notes:
            task.Notes = notes

        return {
            "success": True,
            "id": task.UniqueID,
            "name": task.Name,
            "duration": str(task.Duration),
            "start": str(task.Start),
            "finish": str(task.Finish),
            "percent_complete": task.PercentComplete,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_delete_task(task_id: int) -> dict:
    """Delete a task by UniqueID."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        task = proj.Tasks.UniqueID(task_id)
        name = task.Name
        task.Delete()
        return {"success": True, "deleted": name, "id": task_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- Resource CRUD ----


def msp_add_resource(
    name: str,
    resource_type: str = "work",
    email: str | None = None,
) -> dict:
    """Add a resource to the active project.

    Args:
        name: Resource name
        resource_type: "work", "material", or "cost" (default "work")
        email: Resource email (optional)
    """
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        res = proj.Resources.Add(name)

        type_map = {"work": 1, "material": 2, "cost": 3}
        if resource_type.lower() in type_map:
            res.Type = type_map[resource_type.lower()]

        if email:
            res.EmailAddress = email

        return {
            "success": True,
            "id": res.UniqueID,
            "name": res.Name,
            "type": resource_type,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_get_resources(max_results: int = 50) -> dict:
    """List resources in the active project."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        resources = proj.Resources

        result = []
        count = min(resources.Count, max_results)
        for i in range(1, count + 1):
            r = resources(i)
            result.append(
                {
                    "id": r.UniqueID,
                    "name": r.Name,
                    "type": r.Type,
                    "cost": r.Cost,
                    "work": str(r.Work),
                    "available": r.AvailableTimes,
                }
            )

        return {"success": True, "count": len(result), "resources": result}
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_delete_resource(resource_id: int) -> dict:
    """Delete a resource by UniqueID."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        res = proj.Resources.UniqueID(resource_id)
        name = res.Name
        res.Delete()
        return {"success": True, "deleted": name, "id": resource_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- Assignment ----


def msp_assign_resource(
    task_id: int,
    resource_id: int,
    units: int = 100,
) -> dict:
    """Assign a resource to a task.

    Args:
        task_id: Task UniqueID
        resource_id: Resource UniqueID
        units: Assignment units (100 = full time, 50 = half time)
    """
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        task = proj.Tasks.UniqueID(task_id)
        res = proj.Resources.UniqueID(resource_id)

        assignment = proj.Assignments.Add(
            ResourceID=resource_id,
            TaskID=task_id,
            Units=units,
        )

        return {
            "success": True,
            "task": task.Name,
            "resource": res.Name,
            "units": units,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- Baseline ----


def msp_set_baseline(baseline_number: int = 0) -> dict:
    """Save a baseline for the active project.

    Args:
        baseline_number: 0-10 (0 = Baseline, 1 = Baseline 1, etc.)
    """
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        proj.SaveBaseline(baseline_number)
        return {"success": True, "baseline": baseline_number}
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- Project Info ----


def msp_get_project_info() -> dict:
    """Get info about the active project."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        return {
            "success": True,
            "name": proj.Name,
            "path": proj.FullName,
            "tasks": proj.Tasks.Count,
            "resources": proj.Resources.Count,
            "start": str(proj.Start),
            "finish": str(proj.Finish),
            "calendar": proj.Calendar.Name,
            "saved": proj.Saved,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# ---- View ----


def msp_switch_view(view_name: str) -> dict:
    """Switch to a named view (e.g. Gantt Chart, Resource Sheet)."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        proj.ActivateView(view_name)
        return {"success": True, "view": view_name}
    except Exception as e:
        return {"success": False, "error": str(e)}


def msp_list_views() -> dict:
    """List available views."""
    try:
        app = _get_project_app()
        proj = app.ActiveProject
        views = []
        for i in range(1, proj.Views.Count + 1):
            v = proj.Views(i)
            views.append(v.Name)
        return {"success": True, "views": views}
    except Exception as e:
        return {"success": False, "error": str(e)}
