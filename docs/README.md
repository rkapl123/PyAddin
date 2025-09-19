PyAddin provides handling Python Scripts started from Excel.

# Using PyAddin

Running an script is simply done by selecting the configured python installation on the dropdown "PythonInstallation" in the Script Addin Ribbon Tab and clicking "run <ScriptDefinition>"
beneath the Sheet-button in the Ribbon group "Run Scripts defined in WB/sheets names". With an activated "script output active/inactive" toggle button the script output is shown in an opened window.
Selecting the Script definition in the ScriptDefinition dropdown highlights the specified definition range.

When running scripts, following is executed:

1. (optional, if defined) the scripts defined inside Excel are written,
2. defined/written scripts are called using the executable located in ExePath/exec (see settings)

When holding the Shift-Key pressed while clicking "run <ScriptDefinition>", a control button is added for the selected script definition if the name of the script definition range is either workbook-wide or on the currently active sheet (and the name is not longer than 31 characters). Using this button the script can be executed in the same way as with the ScriptDefinition dropdown. When the button is added, design mode for buttons is automatically turned on to allow changing properties and size. Once any other GUI element is activated in the PyAddin ribbon, design mode is turned off again.

![Image of screenshot1](https://raw.githubusercontent.com/rkapl123/PyAddin/master/docs/screenshot1.png)

# Defining PyAddin script interactions (ScriptDefinitions)

script interactions (ScriptDefinitions) are defined using a 3 column named range (1st col: definition type, 2nd: definition value, 3rd: (optional) parameters):

The Scriptdefinition range name must start with "PyScript_" and can have a postfix as an additional definition name.
If there is no postfix after "PyScript_", the script is called "MainScript" in the Workbook/Worksheet.

A range name can be at Workbook level or worksheet level.
In the ScriptDefinition dropdowns the worksheet name (for worksheet level names) or the workbook name (for workbook level names) is prepended to the additional postfixed definition name.

So for the 8 definitions (range names) currently defined in the test workbook testRAddin.xlsx, there should be 8 entries in the Scriptdefinition dropdown:

- testScriptAddin.xlsx, (Workbooklevel name, runs as MainScript)
- testScriptAddin.xlsxAnotherDef (Workbooklevel name),
- testScriptAddin.xlsxErrorInDef (Workbooklevel name),
- testScriptAddin.xlsxNewResDiagDir (Workbooklevel name),
- Test_OtherSheet, (name in Test_OtherSheet)
- Test_OtherSheetAnotherDef (name in Test_OtherSheet),
- Test_scriptRngScriptCell (Test_scriptRng) and
- Test_scriptRngScriptRange (Test_scriptRng)

In the 1st column of the Scriptdefinition range are the definition types, possible types are
- `pyInst`: the python installation to be used.
- `pylib`: a python executable (dynamic) library, being able to run the python script in line "script" (or scriptrng/scriptcell). This is only needed for overriding the `PyLib<PyInst>` in the AppSettings in the PyAddin.xll.config file.
- `venv`: path to folders with python (virtual environment) libraries (semicolon separated), in case you need to add them. Only needed when overriding the `PyVenv<PyInst>` in the AppSettings in the PyAddin.xll.config file.
- `envvar`: environment variables to add to the process (each line will be one variable/value entry). Only needed when overriding `EnvironVarName<PyInst>`/`EnvironVarValue<PyInst>` settings in the AppSettings in the PyAddin.xll.config file.
- `dir`: the path where below files (scripts, args, results and diagrams) are stored.
- `script`: full path of an executable script.
- `scriptrng`/`scriptcell` or `skipscript` (Scripts directly within Excel): either ranges, where a script is stored (scriptrng) or directly as a cell value (text content or formula result) in the 2nd column (scriptcell). In case the parameter is `skipscript` then the script execution is skipped (set this dynamically to prevent scripts from running).

Scripts (defined with the `script`, `scriptrng` or `scriptcell` definition types) are executed in sequence of their appearance. Although pylib, venv and dir definitions can appear more than once, only the last definition is used.

In the 2nd column are the definition values as described above.
- For `scriptrng` these are range names referring to the respective ranges to be taken as scriptrng target in the excel workbook.
- The range names that are referred in `scriptrng` types can also be either workbook global (having no ! in front of the name) or worksheet local (having the worksheet name + ! in front of the name). They can also contain extensions for the resulting file name.
- for `pyInst` any python installation available in the dropdown "PythonInstallation". This overrides the selection in the dropdown "PythonInstallation".
- for `pylib` this is the full path for the python executable library and overrides the standard setting `PyLib<PyInst>`.
- for `venv`, an additional path to folders with python modules (semicolon separated), in case you need to add them. This overrides the standard setting `PyVenv<PyInst>`.
- for `envvar`, the name of the environment variable to be added. This overrides the potential standard setting `EnvironVarName<PyInst>`/`EnvironVarValue<PyInst>`.
- for `dir` a path that overrides the current workbook folder.

In the 3rd column are additional parameters as follows
- parent folders for `scriptrng`/`scriptcell` entries. Not existing folders are created automatically, so dynamical paths can be given here.
- for `envvar`, the value of the environment variable to be added.

The definitions are loaded into the ScriptDefinition dropdown either on opening/activating a Workbook with above named areas or by pressing the small dialogBoxLauncher "Show AboutBox" on the Script Addin Ribbon Tab and clicking "refresh ScriptDefinitions":  
![Image of screenshot2](https://raw.githubusercontent.com/rkapl123/PyAddin/master/docs/screenshot2.png)

The mentioned hyperlink to the local help file can be configured in the app config file (PyAddin.xll.config) with key "LocalHelp".
When saving the Workbook the input arguments (definition with arg) defined in the currently selected Scriptdefinition dropdown are stored as well. If nothing is selected, the first Scriptdefinition of the dropdown is chosen.

The error messages are logged to a diagnostic log provided by ExcelDna, which can be accessed by clicking on "show Log". The log level can be set in the `system.diagnostics` section of the app-config file (PyAddin.xll.config):
Either you set the switchValue attribute of the source element to prevent any trace messages being generated at all, or you set the initializeData attribute of the added LogDisplay listener to prevent the generated messages to be shown (below the chosen level)  

You can also run PyAddin in an automated way, simply issue the VBA command `result = Application.Run("executeScript", <ScriptDefinitionName>, <headlessFlag>)`, where `<ScriptDefinitionName>` is the Name of the Script Definition Range and `<headlessFlag>` is a boolean flag indicating whether any user-interaction (as controllable by the Addin) should be avoided, all errors are returned in the `result` of the call.

# Installation of PyAddin and Settings

run Distribution/deployAddin.cmd (this puts PyAddin32.xll/PyAddin64.xll as PyAddin.xll and PyAddin.xll.config into %appdata%\Microsoft\AddIns and starts installAddinInExcel.vbs (setting AddIns("PyAddin.xll").Installed = True in Excel)).

Adapt the settings in PyAddin.xll.config:

```XML
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <configSections>
    <section name="UserSettings" type="System.Configuration.NameValueSectionHandler"/>
  </configSections>
  <UserSettings configSource="PyAddinUser.config"/> : This is a redirection to a user specific config file containing the <appSettings> ... </appSettings> information below (in the same path as PyAddin.xll.config). These settings always override the central appSettings
  <appSettings file="\\Path\to\PyAddinCentral.config"> : This is a redirection to a central config file containing the <appSettings> ... </appSettings> information below (any path). The central config file overrides the settings below.

    <add key="EnvironVarNamePython1" value="PYTHONLIB" />
    <add key="EnvironVarValuePython1" value="C:\Users\rolan\specialLib" />
    <add key="PyLibPython1" value="C:\Users\rolan\anaconda3\python38.dll" />
    <add key="PyVenvPython1 value="C:\Users\rolan\anaconda3\Scripts;C:\Users\rolan\anaconda3\Library\bin;C:\Users\rolan\anaconda3\Library\bin;C:\Users\rolan\anaconda3\Library\usr\bin;C:\Users\rolan\anaconda3\Library\mingw-w64\bin" />
    <add key="presetSheetButtonsCount" value="24"/> : the preset maximum Button Count for Sheets (if you expect more sheets with ScriptDefinitions, you can set it accordingly)
    <add key="DebugAddin" value="True"/> : activate Info messages in Log Display to debug addin.
    <add key="disableSettingsDisplay" value="addin"/> : enter a name here for settings that should not be available for viewing/editing to the user (addin: PyAddin.xll.config, central: PyAddinCentral.config, user: PyAddinUser.config).
    <add key="LocalHelp" value="\\LocalPath\to\LocalHelp.htm" /> : If you download this page to your local site, put it there to have it offline.
    <add key="localUpdateFolder" value="" /> : For updating the Script-Addin Version, you can provide an alternative folder, where the deploy script and the files are maintained for other users.
    <add key="localUpdateMessage" value="New version available in local update folder, start deployAddin.cmd to install it:" /> : For the alternative folder update, you can also provide an alternative message to display.
    <add key="updatesDownloadFolder" value="C:\temp\" /> : You can specify a different download folder here instead of C:\temp\.
    <add key="updatesMajorVersion" value="1.0.0." /> : Usually the versions are numbered 1.0.0.x, in case this is different, the Major Version can be overridden here.
    <add key="updatesUrlBase" value="https://github.com/rkapl123/PyAddin/archive/refs/tags/" /> : Here, the URL base for the update zip packages can be overridden.
  </appSettings>
  <system.diagnostics>
    <sources>
      <source name="ExcelDna.Integration" switchValue="All">
        <listeners>
          <remove name="System.Diagnostics.DefaultTraceListener" />
          <add name="LogDisplay" type="ExcelDna.Logging.LogDisplayTraceListener,ExcelDna.Integration">
            <!-- EventTypeFilter takes a SourceLevel as the initializeData:
                    Off, Critical, Error, Warning (default), Information, Verbose, All -->
            <filter type="System.Diagnostics.EventTypeFilter" initializeData="Warning" />
          </add>
        </listeners>
      </source>
    </sources>
  </system.diagnostics>
</configuration>
```

In the PyAddinUser.config setting file, there are two settings that are persisted by the addin itself, so they should not really be changed:
```XML
  <appSettings>
    <add key="debugScript" value="True"/> : whether the script output is active or inactive
    <add key="selectedPythonInstallation" value="0"/> : the currently selected executable for script execution (with dropdown ScriptExecutable)
  </appSettings>
```

The settings for the scripting executables are structured as follows `<PyInstallationPrefix><PyInst>` and form the selection of available script types in PyAddin.

Following PyInstallationPrefixes are possible:
- ExePath: The Executable Path used for the PyInst
- EnvironVarName: An environment variable name to be added for all processes of PyInst
- EnvironVarValue: The value of the above environment variable
- PathAdd : Additional Path Setting for the PyInst
- ExeArgs : Any additional arguments to the PyInst executable

The minimum requirement for a scripting engine to be regarded as selectable/usable is the ExePath entry. All other ScriptExecPrefixes are optional depending on the requirement of the python installation.

There are three settings files which can be used to create a central setting repository (`<appSettings file="your.Central.Configfile.Path">`) along with a user specific overriding mechanism (`<UserSettings configSource="PyAddinUser.config"/>`) defined in the application config file PyAddin.xll.config. All three settings files can be accessed in the ribbon bar beneaht the dropdown `Settings`.

Additionally you can find an `insert Example` mechanism in this dropdown that adds an example script definition range with the above described definition types and example configs.

# Building

All packages necessary for building are contained, simply open PyAddin.sln and build the solution. The script deployForTest.cmd can be used to deploy the built xll and config to %appdata%\Microsoft\AddIns
