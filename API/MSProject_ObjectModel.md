# Microsoft Project API — Quick Reference

Source: learn.microsoft.com/en-us/office/vba/api/overview/project

## Application

```python
app = win32com.client.Dispatch("MSProject.Application")
app.Visible = True
proj = app.ActiveProject
```

## Project

```python
proj = app.Projects.Add()                    # New project
proj = app.Projects.Open("path/to/file.mpp") # Open existing
proj.SaveAs("path/to/save.mpp")              # Save as
proj.Save()                                   # Save
proj.Close()                                  # Close
proj.Name                                     # Project name
proj.FullName                                 # Full file path
proj.Start                                    # Project start date
proj.Finish                                   # Project finish date
proj.Tasks.Count                              # Number of tasks
proj.Resources.Count                          # Number of resources
proj.Calendar.Name                            # Active calendar name
```

## Task

```python
task = proj.Tasks.Add("Task Name")            # Add task
task = proj.Tasks.UniqueID(12345)             # Get by ID
task.Name                                      # Read/write name
task.Duration                                  # "3d", "1w" — read/write
task.Start                                     # Date — read/write
task.Finish                                    # Date — read/write
task.Predecessors                              # "1,3FS+2d" — read/write
task.PercentComplete                           # 0-100 — read/write
task.Cost                                      # Read only
task.Work                                      # Read only
task.Notes                                     # Read/write
task.UniqueID                                  # Read only — use for updates
task.Delete()                                  # Delete task
```

## Resource

```python
res = proj.Resources.Add("Resource Name")     # Add resource
res = proj.Resources.UniqueID(12345)          # Get by ID
res.Name                                       # Read/write
res.Type                                       # 1=work, 2=material, 3=cost
res.Cost                                       # Read only
res.Work                                       # Read only
res.EmailAddress                               # Read/write
res.AvailableTimes                             # Read only
res.UniqueID                                   # Read only
res.Delete()                                   # Delete resource
```

## Assignment

```python
asn = proj.Assignments.Add(
    ResourceID=12345,
    TaskID=67890,
    Units=100                                  # 100=full, 50=half
)
```

## Baseline

```python
proj.SaveBaseline(0)   # Baseline (0-10)
proj.SaveBaseline(1)   # Baseline 1
```

## View

```python
proj.ActivateView("Gantt Chart")
proj.ActivateView("Resource Sheet")
proj.ActivateView("Task Usage")
for i in range(1, proj.Views.Count + 1):
    print(proj.Views(i).Name)
```

## Filter

```python
proj.TaskFilter = "Incomplete Tasks"
proj.TaskFilter = "Critical Tasks"
proj.ResourceFilter = "Overallocated Resources"
```

## Sort

```python
proj.Sort("Start")     # Sort by Start date
proj.Sort("Cost")      # Sort by Cost
proj.Sort("Finish")    # Sort by Finish date
```

## Common Duration Strings

| String | Meaning |
|--------|---------|
| "1d" | 1 day |
| "3d" | 3 days |
| "1w" | 1 week |
| "4h" | 4 hours |
| "2ed" | 2 elapsed days |
| "1ew" | 1 elapsed week |

## Common Predecessor Formats

| Format | Meaning |
|--------|---------|
| "1" | Task 1 finish-to-start |
| "1FS" | Explicit finish-to-start |
| "1SS" | Start-to-start |
| "1FF" | Finish-to-finish |
| "1SF" | Start-to-finish |
| "1FS+2d" | FS with 2-day lag |
| "1FS-1d" | FS with 1-day lead |
| "1,3,5" | Multiple predecessors |
