# CLAUDE.md - Guidelines for ECATPlugin Revit Add-in

## Build Commands
- Build solution: `msbuild ECATPlugin.csproj /p:Configuration=Debug /p:Platform=x64`
- Clean solution: `msbuild ECATPlugin.csproj /t:Clean`
- Output location: `bin/Debug/ECATPlugin.dll`

## Code Style Guidelines
- **Architecture**: MVVM pattern with WPF for UI components
- **Naming**: 
  - PascalCase for classes, methods, properties, events
  - camelCase for local variables and parameters
  - Private fields use _underscore prefix
  - Interface names start with "I"
- **Error Handling**: Use try-catch blocks around Revit API calls, display errors via TaskDialog.Show
- **UI Conventions**: Follow existing XAML patterns for UI components
- **Documentation**: Add XML comments for public methods and properties

## Project Purpose
- Revit plugin that calculates embodied carbon for construction materials
- Extracts material quantities from Revit models
- Applies carbon factors to calculate embodied carbon assessment
- Implements the ECAT (Embodied Carbon Assessment Tool) methodology