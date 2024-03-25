# Ercot Automation with VBA
This project uses VBA to automatically fetch, parse, and calculate average prices purely from Excel. The ultimate goal is to enable non-technical to be able to take advantage of automation without the need of any tedious overhead that comes with environment set up. This means no dependencies external to Excel such as Python. It is meant to be a reference project on how to accomplish difficult programming tasks within the constraints of the legacy VBA coding experience present at time of writing.

The [RubberDuck](https://rubberduckvba.com/) project is highly recommended for an improved developer experience as well as for using this project.

## Dependencies
This project depends on a few excellent open-source projects and inherently, their dependencies:
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)
  - _Microsoft Scripting Runtime_ (`scrrun.dll`)
- [VBA-UTC](https://github.com/VBA-tools/VBA-UTC/pull/7)
  - _Microsoft VBScript Regular Expressions 5.5_ (`vbscript.dll`)
- [VBA-CSV](https://github.com/sdkn104/VBA-CSV)
- _Microsoft Shell Controls And Automation_ (`shell32.dll`)

All `.dll` references are already added to the given example spreadhseet. To use these open-source projects in a new spreadsheet, they must be added manually in the developer VBA editor window via Tools -> References.

All `.bas` modules are also already imported into the VBA editor.
