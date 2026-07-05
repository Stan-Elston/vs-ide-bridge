using System.Text.Json.Nodes;
using static VsIdeBridgeService.ArgBuilder;
using static VsIdeBridgeService.SchemaHelpers;

namespace VsIdeBridgeService;

internal static partial class ToolCatalog
{
    private const string ProjectDescriptionProperty = "description";
    private const string ListProjectsTool = "list_projects";
    private const string QueryProjectItemsTool = "query_project_items";
    private const string QueryProjectPropertiesTool = "query_project_properties";

    private static IEnumerable<ToolEntry> ProjectTools()
        =>
        ProjectQueryTools()
            .Concat(ProjectManagementTools())
            .Concat(ProjectPythonTools());

    private static IEnumerable<ToolEntry> ProjectQueryTools()
    {
        yield return BridgeTool(ListProjectsTool,
            "List all projects in the open solution with their names, paths, and metadata. Use project names from the result to target specific projects in other commands like build, query_project_properties, or query_project_items.",
            EmptySchema(), "list-projects", _ => Empty(), Project,
            outputSchema: BuildListProjectsOutputSchema(),
            searchHints: BuildSearchHints(
                workflow: [("build", "Build a specific project by name"), (QueryProjectPropertiesTool, "Read properties for a listed project")],
                related: [(QueryProjectItemsTool, "List files in a project"), ("query_project_references", "List references for a project")]));

        yield return BridgeTool(QueryProjectItemsTool,
            "List items in a project with file paths, kinds, item types, and the project file itself. Use the returned project-file row with read_file or apply_diff when editing .csproj, .vcxproj, .props, or .targets XML.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                Opt(Path, "Optional path filter. Matches project item paths and the project file path."),
                OptInt(Max, "Max items to return (default 200).")),
            "query-project-items",
            a => Build(
                (Project, OptionalString(a, Project)),
                (Path, OptionalString(a, Path)),
                (Max, OptionalText(a, Max))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("read_file", "Read a listed file"), ("add_file_to_project", "Add a missing file to the project")],
                related: [(QueryProjectPropertiesTool, "Read project properties"), ("file_outline", "Get symbol structure of a listed file")]));

        yield return BridgeTool(QueryProjectPropertiesTool,
            "Read MSBuild project properties from one project.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                OptArr("names", "Property names to read.")),
            "query-project-properties",
            a => Build(
                (Project, OptionalString(a, Project)),
                ("names", OptionalStringArray(a, "names"))),
            Project,
            searchHints: BuildSearchHints(
                related: [(QueryProjectItemsTool, "List files in the project"), ("query_project_references", "List project references"), ("build", "Build after reviewing properties")]));

        yield return BridgeTool("query_project_configurations",
            "List project configurations and platforms for one project.",
            ObjectSchema(Req(Project, ProjectDesc)),
            "query-project-configurations",
            a => Build((Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                related: [("set_build_configuration", "Switch to a different configuration"), ("build", "Build with the active configuration")]));

        yield return BridgeTool("query_project_references",
            "List project references for one project.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                OptBool("declared_only", "Return only declared (project-FileArg) references."),
                OptBool("include_framework",
                    "Include framework assembly references (default false).")),
            "query-project-references",
            a => Build(
                (Project, OptionalString(a, Project)),
                BoolArg("declared-only", a, "declared_only", false, true),
                BoolArg("include-framework", a, "include_framework", false, true)),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("scan_project_dependencies", "Get a full dependency health scan")],
                related: [("nuget_add_package", "Add a NuGet package"), (QueryProjectItemsTool, "List source files")]));

        yield return BridgeTool("query_project_outputs",
            "Resolve the primary output artifact and output directory for one project.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                Opt(Configuration, "Build configuration."),
                Opt("target_framework", "Target framework moniker.")),
            "query-project-outputs",
            a => Build(
                (Project, OptionalString(a, Project)),
                (Configuration, OptionalString(a, Configuration)),
                ("target-framework", OptionalString(a, "target_framework"))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("build", "Build the project to produce the output")],
                related: [(QueryProjectPropertiesTool, "Read OutputPath and TargetFramework directly"), ("query_project_configurations", "List available configurations")]));
    }

    private static JsonObject BuildListProjectsOutputSchema()
        => new()
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["count"] = new JsonObject { ["type"] = "integer", [ProjectDescriptionProperty] = "Number of projects found." },
                ["projects"] = new JsonObject
                {
                    ["type"] = "array",
                    [ProjectDescriptionProperty] = "Array of project objects with name, path, and metadata.",
                    ["items"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["properties"] = new JsonObject
                        {
                            ["name"] = new JsonObject { ["type"] = "string", [ProjectDescriptionProperty] = "The display name of the project." },
                            ["uniqueName"] = new JsonObject { ["type"] = "string", [ProjectDescriptionProperty] = "The unique identifier for the project (includes solution folders)." },
                            ["path"] = new JsonObject { ["type"] = "string", [ProjectDescriptionProperty] = "The absolute file system path to the project file." },
                            ["kind"] = new JsonObject { ["type"] = "string", [ProjectDescriptionProperty] = "The project type identifier (e.g., project kind GUID)." },
                            ["isStartup"] = new JsonObject { ["type"] = "boolean", [ProjectDescriptionProperty] = "Whether this is the solution startup project." },
                        },
                        ["required"] = new JsonArray { "name", "uniqueName", "path", "kind", "isStartup" },
                        ["additionalProperties"] = false,
                    },
                },
            },
            ["required"] = new JsonArray { "count", "projects" },
            ["additionalProperties"] = false,
        };

    private static IEnumerable<ToolEntry> ProjectManagementTools()
        => ProjectSolutionTools()
            .Concat(LaunchProfileTools())
            .Concat(ProjectFileTools());

    private static IEnumerable<ToolEntry> ProjectSolutionTools()
    {
        yield return BridgeTool("add_project",
            "Add an existing or new project to the solution.",
            ObjectSchema(
                Req(Project, "Absolute path to the project FileArg."),
                Opt("solution_folder", "Optional solution folder name.")),
            "add-project",
            a => Build(
                (Project, OptionalString(a, Project)),
                ("solution-folder", OptionalString(a, "solution_folder"))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [(ListProjectsTool, "Confirm the project was added"), ("build", "Build after adding")],
                related: [("create_project", "Create a new project instead"), ("remove_project", "Remove a project")]));

        yield return BridgeTool("remove_project",
            "Remove a project from the solution by name or path.",
            ObjectSchema(Req(Project, "Project name or path to remove.")),
            "remove-project",
            a => Build((Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                related: [(ListProjectsTool, "List remaining projects"), ("add_project", "Re-add the project")]));

        yield return BridgeTool("rename_project",
            "Rename a project within the solution. This updates the project name shown by Visual Studio, but does not rename folders or the project file on disk.",
            ObjectSchema(
                Req(Project, "Project name or path to rename."),
                Req("new_name", "New project name to show in the solution.")),
            "rename-project",
            a => Build(
                (Project, OptionalString(a, Project)),
                ("new-name", OptionalString(a, "new_name"))),
            Project,
            searchHints: BuildSearchHints(
                related: [(ListProjectsTool, "Confirm the rename"), (QueryProjectPropertiesTool, "Read project properties")]));

        yield return BridgeTool("create_project",
            "Create a new project and add it to the open solution.",
            ObjectSchema(
                Req("name", "New project name."),
                Opt("template", "Project template name or identifier."),
                Opt("language", "Programming language (e.g. C#, VB, F#)."),
                Opt("directory", "Directory to create the project in."),
                Opt("solution_folder", "Optional solution folder name.")),
            "create-project",
            a => Build(
                ("name", OptionalString(a, "name")),
                ("template", OptionalString(a, "template")),
                ("language", OptionalString(a, "language")),
                ("directory", OptionalString(a, "directory")),
                ("solution-folder", OptionalString(a, "solution_folder"))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [(ListProjectsTool, "Confirm the project was created"), ("build", "Build the new project")],
                related: [("add_project", "Add an existing project instead"), ("add_file_to_project", "Add source files to the new project")]));

        yield return BridgeTool("set_startup_project",
            "Set the solution startup project by name or path.",
            ObjectSchema(Req(Project, ProjectDesc)),
            "set-startup-project",
            a => Build((Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("debug_start", "Start debugging the startup project"), ("build", "Build the startup project")],
                related: [(ListProjectsTool, "List projects to find the right name"), ("list_launch_profiles", "List solution launch profiles")]));
    }

    private static IEnumerable<ToolEntry> LaunchProfileTools()
    {
        yield return BridgeTool("list_launch_profiles",
            "List all solution-level launch profiles from the .slnLaunch file. " +
            "A .slnLaunch file is a JSON array that lives next to the .sln file and defines named multi-project startup configurations - " +
            "the same profiles that appear in the startup dropdown in the VS toolbar. " +
            "Each profile has a Name and a Projects array; each project entry has a Path (relative to the solution), " +
            "an Action (Start, StartWithoutDebugging, or None), and an optional DebugTarget. " +
            "Returns profile names, project counts, and per-project path/action/debugTarget. " +
            "Call this first to discover available profile names before calling set_launch_profile.",
            EmptySchema(), "list-launch-profiles", _ => Empty(), Project,
            searchHints: BuildSearchHints(
                workflow: [("set_launch_profile", "Activate one of the listed profiles")],
                related: [("set_startup_project", "Set a single startup project instead"), (ListProjectsTool, "List projects in the solution")]));

        yield return BridgeTool("set_launch_profile",
            "Activate a named launch profile from the .slnLaunch file, switching VS's startup project selection immediately. " +
            "Supports exact or partial name matching (case-insensitive); partial match must be unambiguous - if multiple profiles match the query an error lists them so you can be more specific. " +
            "Only projects with Action 'Start' or 'StartWithoutDebugging' become startup projects; projects with Action 'None' are listed in the result but not activated. " +
            "Call list_launch_profiles first to see available profile names.",
            ObjectSchema(Req("name", "Launch profile name (or partial match, case-insensitive). Use the exact name from list_launch_profiles to avoid ambiguity.")),
            "set-launch-profile",
            a => Build(("name", OptionalString(a, "name"))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("debug_start", "Start debugging with the activated profile"), ("build", "Build the startup projects")],
                related: [("list_launch_profiles", "List available profiles first"), ("set_startup_project", "Set a single startup project instead")]));
    }

    private static IEnumerable<ToolEntry> ProjectFileTools()
    {
        yield return BridgeTool("add_file_to_project",
            "Add an existing file to a project as a project item. For editing the project file XML itself, use query_project_items to find the project-file row, then read_file and apply_diff.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                Req(FileArg, "Absolute path to the file to add.")),
            "add-file-to-project",
            a => Build(
                (Project, OptionalString(a, Project)),
                (FileArg, OptionalString(a, FileArg))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("read_file", "Read the added file"), ("build", "Build after adding")],
                related: [("remove_file_from_project", "Remove a file from the project"), ("query_project_items", "List current project files")]));

        yield return BridgeTool("remove_file_from_project",
            "Remove a project item from a project without deleting the file from disk. For editing the project file XML itself, use query_project_items to find the project-file row, then read_file and apply_diff.",
            ObjectSchema(
                Req(Project, ProjectDesc),
                Req(FileArg, "File path or project item path to remove.")),
            "remove-file-from-project",
            a => Build(
                (Project, OptionalString(a, Project)),
                (FileArg, OptionalString(a, FileArg))),
            Project,
            searchHints: BuildSearchHints(
                related: [("add_file_to_project", "Re-add the file"), ("query_project_items", "List remaining project files")]));
    }

    private static IEnumerable<ToolEntry> ProjectPythonTools()
    {
        yield return BridgeTool("python_set_project_env",
            "Set the active Python interpreter for the active Python project in Visual Studio.",
            ObjectSchema(
                Req(Path, "Absolute path to the Python interpreter."),
                Opt(Project, "Python project name or path. Defaults to the active project.")),
            "set-python-project-env",
            a => Build(
                (Path, OptionalString(a, Path)),
                (Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                workflow: [("python_list_envs", "Discover available interpreters first")],
                related: [("python_sync_env", "Sync bridge interpreter to VS project")]));

        yield return BridgeTool("python_set_startup_file",
            "Set the startup file for the active Python project.",
            ObjectSchema(
                Req(FileArg, "Path to the Python file to set as startup."),
                Opt(Project, "Python project name or path. Defaults to the active project.")),
            "set-python-startup-file",
            a => Build(
                (FileArg, OptionalString(a, FileArg)),
                (Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                related: [("python_get_startup_file", "Read the current startup file")]));

        yield return BridgeTool("python_get_startup_file",
            "Get the startup file configured for the active Python project.",
            ObjectSchema(
                Opt(Project, "Python project name or path. Defaults to the active project.")),
            "get-python-startup-file",
            a => Build((Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                related: [("python_set_startup_file", "Change the startup file")]));

        yield return BridgeWrapperTool("python_sync_env",
            "Sync the active bridge Python interpreter to the active Python project in Visual Studio.",
            ObjectSchema(
                Opt(Project, "Python project name or path. Defaults to the active project.")),
            "set-python-project-env",
            a => Build(
                (Path, PythonInterpreterState.LoadActiveInterpreterPath()),
                (Project, OptionalString(a, Project))),
            Project,
            searchHints: BuildSearchHints(
                related: [("python_set_project_env", "Set a specific interpreter"), ("python_list_envs", "List available interpreters")]));
    }
}
