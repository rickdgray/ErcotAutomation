# Ercot Automation with VBA
This project uses VBA to automatically fetch, parse, and calculate average prices purely from Excel. It is a reference project to accomplish difficult tasks within the constraints of the legacy VBA coding experience given in Excel.

## Dependencies
This project depends on the excellent [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) project, and inherently it's dependency on the _Microsoft Scripting Runtime_ (`scrrun.dll`) library built into Windows. This project also depends on the _Microsoft Shell Controls And Automation_ (`shell32.dll`) Windows library.

To run this project, a few steps are needed. First, both references must be added in the developer VBA editor window via Tools -> References.

Then, the `JsonConverter.bas` module should be imported into the VBA editor.

The [RubberDuck](https://rubberduckvba.com/) project can be helpful for simplifying the importing of the Json Converter as well as this project.
