Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks


''' <summary>The main functions for working with ScriptDefinitions (named ranges in Excel) and starting the Script processes (writing input, invoking scripts and retrieving results)</summary>
Public Module PyAddin
    ''' <summary>selected environment number for the fixed selectable python installations</summary>
    Public PyInstallation As Integer
    ''' <summary>library name for executing python scripts</summary>
    Public PyLib As String = ""
    ''' <summary>optional additional virtual environment path for PyLib</summary>
    Public PyVenv As String = ""
    ''' <summary>optional additional environment settings for PyLib</summary>
    Public PyAddEnvironVars As New Dictionary(Of String, String)
    ''' <summary>for PyAddin invocations by executeScript, this is set to true, avoiding a MsgBox</summary>
    Public nonInteractive As Boolean = False
    ''' <summary>collect non interactive error messages here</summary>
    Public nonInteractiveErrMsgs As String
    ''' <summary>Debug the Add-in: write trace messages</summary>
    Public DebugAddin As Boolean
    ''' <summary>The path where the User specific settings (overrides) can be found</summary>
    Public UserSettingsPath As String
    ''' <summary>indicates an error in execution of script, used for non interactive message return</summary>
    Public hadError As Boolean
    ''' <summary></summary>
    Public StdErrMeansError As Boolean
    ''' <summary>the LogDisplay (Diagnostic Display) log source</summary>
    Public theLogDisplaySource As TraceSource

    ''' <summary>the current workbook, used for reference of all script related actions (only one workbook is supported to hold script definitions)</summary>
    Public currWb As Workbook
    ''' <summary>the current script definition range (three columns)</summary>
    Public ScriptDefinitionRange As Range
    ''' <summary></summary>
    Public Scriptcalldefnames As String() = {}
    ''' <summary></summary>
    Public Scriptcalldefs As Range() = {}
    ''' <summary></summary>
    Public ScriptDefsheetColl As Dictionary(Of String, Dictionary(Of String, Range))
    ''' <summary></summary>
    Public ScriptDefsheetMap As Dictionary(Of String, String)
    ''' <summary>reference object for the Add-ins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary></summary>
    Public avoidFurtherMsgBoxes As Boolean
    ''' <summary></summary>
    Public dirglobal As String
    ''' <summary>show the script output for debugging purposes (invisible otherwise)</summary>
    Public debugScript As Boolean
    ''' <summary>needed for workbook save (saves selected ScriptDefinition)</summary>
    Public dropDownSelected As Boolean
    ''' <summary>set to true if warning was issued, this flag indicates that the log button should get an exclamation sign</summary>
    Public WarningIssued As Boolean

    ''' <summary>definitions of current script invocations (scripts, args, results, diags...)</summary>
    Public ScriptDefDic As New Dictionary(Of String, String())
    ''' <summary>currently running scripts to prevent repeated invocations </summary>
    Public ScriptRunDic As New Dictionary(Of Integer, Boolean)

    ''' <summary>startRprocess: started from GUI (button) and accessible from VBA (via Application.Run)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function startScriptprocess() As String
        Dim errStr As String
        avoidFurtherMsgBoxes = False
        ' get the definition range
        errStr = getScriptDefinitions()
        If errStr <> vbNullString Then Return "Failed getting ScriptDefinitions: " + errStr
        Try
            If Not storeScriptRng() Then Return vbNullString
            If Not invokeScripts() Then Return vbNullString
        Catch ex As Exception
            Return "Exception in ScriptDefinitions preparation and execution: " + ex.Message + ex.StackTrace
        End Try
        ' all is OK = return null string
        Return vbNullString
    End Function

    ''' <summary>execute given ScriptDefName, used for VBA call by Application.Run</summary>
    ''' <param name="ScriptDefName">Name of Script Definition</param>
    ''' <param name="headLess">if set to true, ScriptAddin will avoid to issue messages and return messages in exceptions which are returned (headless)</param>
    ''' <returns>empty string on success, error message otherwise</returns>
    <ExcelCommand(Name:="executeScript")>
    Public Function executeScript(ScriptDefName As String, Optional headLess As Boolean = False) As String
        hadError = False : nonInteractive = headLess
        nonInteractiveErrMsgs = "" ' reset non interactive messages
        Try
            PyAddin.ScriptDefinitionRange = ExcelDnaUtil.Application.ActiveWorkbook.Names.Item(ScriptDefName).RefersToRange
        Catch ex As Exception
            nonInteractive = False
            Return "No Script Definition Range (" + ScriptDefName + ") found in Active Workbook: " + ex.Message
        End Try
        LogInfo("Doing Script '" + ScriptDefName + "'.")
        Try
            currWb = ExcelDnaUtil.Application.ActiveWorkbook
            Dim errStr As String = PyAddin.getScriptNames()
            If errStr <> "" Then Throw New Exception("Error in ScriptAddin.getScriptNames: " + errStr)
            errStr = PyAddin.startScriptprocess()
            If errStr <> "" Then Throw New Exception("Error in ScriptAddin.startScriptprocess: " + errStr)
        Catch ex As Exception
            nonInteractive = False
            hadError = True
            Return "Script Definition '" + ScriptDefName + "' execution had following error(s): " + ex.Message
        End Try
        nonInteractive = False
        If hadError Then Return nonInteractiveErrMsgs
        Return "" ' no error, no message
    End Function

    ''' <summary>refresh ScriptNames from Workbook on demand (currently when invoking about box)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function startScriptNamesRefresh() As String
        Dim errStr As String
        If currWb Is Nothing Then Return "No Workbook active to refresh ScriptNames from..."
        ' always reset ScriptDefinitions when refreshing, otherwise this is not being refilled in getRNames
        ScriptDefinitionRange = Nothing
        ' get the defined Script_/R_Addin Names
        errStr = getScriptNames()
        If errStr = "no PyScript Definitions" Then
            Return vbNullString
        ElseIf errStr <> vbNullString Then
            Return "Error while getting Script in startScriptNamesRefresh: " + errStr
        End If
        theRibbon.Invalidate()
        Return vbNullString
    End Function

    ''' <summary>gets defined named ranges for script invocation in the current workbook</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function getScriptNames() As String
        ReDim Preserve Scriptcalldefnames(-1)
        ReDim Preserve Scriptcalldefs(-1)
        ScriptDefsheetColl = New Dictionary(Of String, Dictionary(Of String, Range))
        ScriptDefsheetMap = New Dictionary(Of String, String)
        Dim i As Integer = 0
        For Each namedrange As Name In currWb.Names
            Dim cleanname As String = Replace(namedrange.Name, namedrange.Parent.Name & "!", "")
            If Left(cleanname, 9) = "PyScript_" Then
                Dim prefix As String = Left(cleanname, 7)
                If InStr(namedrange.RefersTo, "#REF!") > 0 Then Return "PyScriptDefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " contains #REF!"
                If namedrange.RefersToRange.Columns.Count <> 3 Then Return "PyScriptDefinitions range " + namedrange.Parent.name + "!" + namedrange.Name + " doesn't have 3 columns !"
                ' final name of entry is without Script_/R_Addin and !
                Dim finalname As String = Replace(Replace(namedrange.Name, prefix, ""), "!", "")
                Dim nodeName As String = Replace(Replace(namedrange.Name, prefix, ""), namedrange.Parent.Name & "!", "")
                If nodeName = "" Then nodeName = "MainScript"
                ' first definition as standard definition (works without selecting a ScriptDefinition)
                If ScriptDefinitionRange Is Nothing Then ScriptDefinitionRange = namedrange.RefersToRange
                If Not InStr(namedrange.Name, "!") > 0 Then
                    finalname = currWb.Name + finalname
                End If
                ReDim Preserve Scriptcalldefnames(Scriptcalldefnames.Length)
                ReDim Preserve Scriptcalldefs(Scriptcalldefs.Length)
                Scriptcalldefnames(Scriptcalldefnames.Length - 1) = finalname
                Scriptcalldefs(Scriptcalldefs.Length - 1) = namedrange.RefersToRange

                Dim scriptColl As Dictionary(Of String, Range)
                If Not ScriptDefsheetColl.ContainsKey(namedrange.Parent.Name) Then
                    ' add to new sheet "menu"
                    scriptColl = New Dictionary(Of String, Range) From {
                        {nodeName, namedrange.RefersToRange}
                    }
                    ScriptDefsheetColl.Add(namedrange.Parent.Name, scriptColl)
                    ScriptDefsheetMap.Add("ID" + i.ToString(), namedrange.Parent.Name)
                    i += 1
                Else
                    ' add ScriptDefinition to existing sheet "menu"
                    scriptColl = ScriptDefsheetColl(namedrange.Parent.Name)
                    scriptColl.Add(nodeName, namedrange.RefersToRange)
                End If
            End If
        Next
        If UBound(Scriptcalldefnames) = -1 Then Return "no PyScript Definitions"
        Return vbNullString
    End Function

    ''' <summary>reset all ScriptDefinition representations</summary>
    Public Sub resetScriptDefinitions()
        ScriptDefDic("args") = {}
        ScriptDefDic("argspaths") = {}
        ScriptDefDic("scripts") = {}
        ScriptDefDic("scriptspaths") = {}
        ScriptDefDic("scriptrng") = {}
        ScriptDefDic("scriptrngpaths") = {}
        dirglobal = vbNullString
    End Sub

    ''' <summary>gets definitions from current selected script invocation range (ScriptDefinitions)</summary>
    ''' <returns>Error message or null string in case of success</returns>
    Public Function getScriptDefinitions() As String
        resetScriptDefinitions()
        Try
            Dim reInitPython As Boolean = False
            If IsNothing(ScriptDefinitionRange) Then Return "No PyScriptDefinitionRange available!"
            For Each defRow As Range In ScriptDefinitionRange.Rows
                Dim deftype As String, defval As String, deffilepath As String
                deftype = LCase(defRow.Cells(1, 1).Value2)
                defval = defRow.Cells(1, 2).Value2
                defval = If(defval = vbNullString, "", defval)
                deffilepath = defRow.Cells(1, 3).Value2
                deffilepath = If(deffilepath = vbNullString, "", deffilepath)
                If (deftype = "pylib") Then
                    If defval <> "" And PyLib <> defval Then
                        PyLib = defval
                        reInitPython = True
                    End If
                ElseIf deftype = "script" Then
                    If defval <> "" Then
                        ReDim Preserve ScriptDefDic("scripts")(ScriptDefDic("scripts").Length)
                        ScriptDefDic("scripts")(ScriptDefDic("scripts").Length - 1) = defval
                        ReDim Preserve ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length)
                        ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length - 1) = deffilepath
                    End If
                ElseIf deftype = "path" And defval <> "" Then
                    If defval <> "" And PyVenv <> defval Then
                        PyVenv = defval
                        reInitPython = True
                    End If
                ElseIf deftype = "envvar" And defval <> "" Then
                    If defval <> "" Then
                        PyAddEnvironVars(defval) = deffilepath
                    End If
                ElseIf deftype = "pyInst" Then
                    If PyInstallations.Contains(defval) And PyInstallation <> defval Then
                        PyInstallation = defval
                        PyLib = fetchSetting("PyLib" + PyInstallation, "")
                        PyVenv = fetchSetting("PyVenv" + PyInstallation, "")
                        reInitPython = True
                        theMenuHandler.selectedPyExecutable = PyInstallations.IndexOf(PyInstallation)
                        ' not really important if not set at startup of addin (timing problem as ribbon is not loaded here)
                        Try : theRibbon.InvalidateControl("pyInstDropDown") : Catch ex As Exception : End Try
                    Else
                        Return "Error in setting type, " + defval + " is not contained in available python installations (check AppSettings for available PyLib<> entries)!"
                    End If
                ElseIf deftype = "arg" Then
                    ReDim Preserve ScriptDefDic("args")(ScriptDefDic("args").Length)
                    ScriptDefDic("args")(ScriptDefDic("args").Length - 1) = defval
                    ReDim Preserve ScriptDefDic("argspaths")(ScriptDefDic("argspaths").Length)
                    ScriptDefDic("argspaths")(ScriptDefDic("argspaths").Length - 1) = deffilepath
                ElseIf deftype = "scriptrng" Or deftype = "scriptcell" Then
                    ReDim Preserve ScriptDefDic("scriptrng")(ScriptDefDic("scriptrng").Length)
                    ScriptDefDic("scriptrng")(ScriptDefDic("scriptrng").Length - 1) = IIf(Right(deftype, 4) = "cell", "=", "") + defval
                    ReDim Preserve ScriptDefDic("scriptrngpaths")(ScriptDefDic("scriptrngpaths").Length)
                    ScriptDefDic("scriptrngpaths")(ScriptDefDic("scriptrngpaths").Length - 1) = deffilepath
                    ' don't set skipscripts here to False as this is done in method storeScriptRng
                ElseIf deftype = "dir" Then
                    dirglobal = defval
                ElseIf deftype <> "" Then
                    Return "Error in getScriptDefinitions: invalid type '" + deftype + "' found in script definition!"
                End If
            Next
            If fetchSetting("EnvironVarName" + PyInstallation, "") <> "" Then
                If Not PyAddEnvironVars.ContainsKey(fetchSetting("EnvironVarName" + PyInstallation, "")) Then
                    PyAddEnvironVars(fetchSetting("EnvironVarName" + PyInstallation, "")) = fetchSetting("EnvironVarValue" + PyInstallation, "")
                End If
            End If
            If PyLib = "" Then Return "Error in getScriptDefinitions: PyLib not defined (check AppSettings for available PyLib<> entries)"
            If ScriptDefDic("scripts").Length = 0 And ScriptDefDic("scriptrng").Length = 0 Then Return "Error in getScriptDefinitions: no script(s) or scriptRng(s) defined in " + ScriptDefinitionRange.Name.Name
            If reInitPython Then PythonCaller.InitPython()
        Catch ex As Exception
            Return "Error in getScriptDefinitions: " + ex.Message
        End Try
        Return vbNullString
    End Function


    ''' <summary>prepare parameter (script, args, results, diags) for usage in invokeScripts, storeArgs, getResults and getDiags</summary>
    ''' <param name="index">index of parameter to be prepared in ScriptDefDic</param>
    ''' <param name="name">name (type) of parameter: scripts, scriptrng, args, results, diags</param>
    ''' <param name="ScriptDataRange">returned Range of data area for scriptrng, args, results and diags</param>
    ''' <param name="returnName">returned name of data file for the parameter: same as range name</param>
    ''' <param name="returnPath">returned path of data file for the parameter</param>
    ''' <param name="ext">extension of filename that should be used for file containing data for that type (e.g. txt for args/results or png for diags)</param>
    ''' <returns>True if success, False otherwise</returns>
    Private Function prepareParam(index As Integer, name As String, ByRef ScriptDataRange As Range, ByRef returnName As String, ByRef returnPath As String, ext As String) As String
        Dim value As String = ScriptDefDic(name)(index)
        If value = "" Then Return "Empty definition value for parameter " + name + ", index: " + index.ToString()
        ' allow for other extensions than txt if defined in ScriptDefDic(name)(index)
        If InStr(value, ".") > 0 Then ext = ""
        ' only for args, results and diags (scripts don't have a target range)
        Dim ScriptDataRangeAddress As String = ""
        If name = "scriptrng" Then
            Try
                ScriptDataRange = currWb.Names.Item(value).RefersToRange
                ScriptDataRangeAddress = ScriptDataRange.Parent.Name + "!" + ScriptDataRange.Address
            Catch ex As Exception
                Return "Error occurred when looking up " + name + " range '" + value + "' in Workbook " + currWb.Name + " (defined correctly ?), " + ex.Message
            End Try
        End If
        ' if arg value refers to a WS Name, cut off WS name prefix for Script file name...
        Dim posWSseparator = InStr(value, "!")
        If posWSseparator > 0 Then
            value = value.Substring(posWSseparator)
        End If
        ' get path of data file, if it is defined
        If ScriptDefDic.ContainsKey(name + "paths") Then
            If Len(ScriptDefDic(name + "paths")(index)) > 0 Then
                returnPath = ScriptDefDic(name + "paths")(index)
            End If
        End If
        returnName = value + ext
        LogInfo("prepared param in index:" + index.ToString() + ",type:" + name + ",returnName:" + returnName + ",returnPath:" + returnPath + IIf(ScriptDataRangeAddress <> "", ",ScriptDataRange: " + ScriptDataRangeAddress, ""))
        Return vbNullString
    End Function

    ''' <summary>creates script files for defined scriptRng ranges</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function storeScriptRng() As Boolean
        Dim scriptRngFilename As String = vbNullString, scriptText = vbNullString
        Dim ScriptDataRange As Range = Nothing
        Dim outputFile As StreamWriter = Nothing

        Dim scriptRngdir As String = dirglobal
        For c As Integer = 0 To ScriptDefDic("scriptrng").Length - 1
            Try
                Dim ErrMsg As String
                ' scriptrng beginning with a "=" is a scriptcell (as defined in getScriptDefinitions) ...
                If Left(ScriptDefDic("scriptrng")(c), 1) = "=" Then
                    scriptText = ScriptDefDic("scriptrng")(c).Substring(1)
                    scriptRngFilename = "ScriptDataRangeRow" + c.ToString() + ".pl"
                Else
                    ErrMsg = prepareParam(c, "scriptrng", ScriptDataRange, scriptRngFilename, scriptRngdir, ".pl")
                    If Len(ErrMsg) > 0 Then
                        If Not PyAddin.UserMsg(ErrMsg) Then Return False
                    End If
                End If

                ' absolute paths begin with \\ or X:\ -> don't prefix with currWB path, else currWBpath\scriptRngdir
                Dim curWbPrefix As String = IIf(Left(scriptRngdir, 2) = "\\" Or Mid(scriptRngdir, 2, 2) = ":\", "", currWb.Path + "\")
                outputFile = New StreamWriter(curWbPrefix + scriptRngdir + "\" + scriptRngFilename, False, Encoding.Default)

                ' reuse the script invocation methods by setting the respective parameters
                ReDim Preserve ScriptDefDic("scripts")(ScriptDefDic("scripts").Length)
                ScriptDefDic("scripts")(ScriptDefDic("scripts").Length - 1) = scriptRngFilename
                ReDim Preserve ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length)
                ScriptDefDic("scriptspaths")(ScriptDefDic("scriptspaths").Length - 1) = scriptRngdir
                ReDim Preserve ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length)
                ScriptDefDic("skipscripts")(ScriptDefDic("skipscripts").Length - 1) = False

                ' write the ScriptDataRange or scriptText (if script directly in cell/formula right next to scriptrng) to file
                If Not IsNothing(scriptText) Then
                    outputFile.WriteLine(scriptText)
                Else
                    Dim i As Integer = 1
                    Do
                        Dim j As Integer = 1
                        Dim writtenLine As String = ""
                        If ScriptDataRange(i, 1).Value2 IsNot Nothing Then
                            Do
                                writtenLine += ScriptDataRange(i, j).Value2
                                j += 1
                            Loop Until j > ScriptDataRange.Columns.Count
                            outputFile.WriteLine(writtenLine)
                        End If
                        i += 1
                    Loop Until i > ScriptDataRange.Rows.Count
                End If
                LogInfo("stored Script to " + curWbPrefix + scriptRngdir + "\" + scriptRngFilename)
            Catch ex As Exception
                If outputFile IsNot Nothing Then outputFile.Close()
                If Not PyAddin.UserMsg("Error occurred when creating script file '" + scriptRngFilename + "', " + ex.Message,, True) Then Return False
            End Try
            If outputFile IsNot Nothing Then outputFile.Close()
        Next
        Return True
    End Function

    Public fullScriptPath As String
    Public script As String
    Public scriptarguments As String
    Public previousDir As String
    Public theScriptOutput As PyOutput

    ''' <summary>invokes current scripts/args/results definition</summary>
    ''' <returns>True if success, False otherwise</returns>
    Public Function invokeScripts() As Boolean
        Dim scriptpath As String
        previousDir = Directory.GetCurrentDirectory()
        scriptpath = dirglobal
        LogInfo("starting " + CStr(ScriptDefDic("scripts").Length - 1) + " scripts")
        ' start script invocation loop as asynchronous thread to allow blocking ShowDialog while not blocking main Excel GUI thread (allows switching the dialog on/off)
        Task.Run(Async Function()
                     Dim ErrMsg As String = ""
                     ' loop through defined scripts, in case you are wondering about scriptrng definitions, for this the scripts dictionary is reused within the invocation of storeScriptRng...
                     For c As Integer = 0 To ScriptDefDic("scripts").Length - 1
                         ' initialize ScriptRunDic entries
                         If Not ScriptRunDic.ContainsKey(c) Then ScriptRunDic.Add(c, False)
                         ' skip script if defined...
                         If ScriptDefDic("skipscripts")(c) Then Continue For
                         ErrMsg = prepareParam(c, "scripts", Nothing, script, scriptpath, "")
                         If Len(ErrMsg) > 0 Then
                             ' allow to ignore preparation errors...
                             If Not PyAddin.UserMsg(ErrMsg) Then Exit For
                             ErrMsg = ""
                         End If

                         ' avoid rerunning same script ...
                         If ScriptRunDic(c) Then
                             If PyAddin.QuestionMsg("Script " + scriptpath + "\" + script + " is already running, start another instance?", MsgBoxStyle.OkCancel, "Script already running", MsgBoxStyle.Exclamation) = MsgBoxResult.Cancel Then Continue For
                         Else
                             ScriptRunDic(c) = True
                         End If
                         ' reflect running state in debug label...
                         PyAddin.theRibbon.InvalidateControl("debug")

                         ' a blank separator indicates additional arguments, separate argument passing because of possible blanks in path -> need quotes around path + scriptname
                         ' assumption: scriptname itself may not have blanks in it.
                         If InStr(script, " ") > 0 Then
                             scriptarguments = script.Substring(InStr(script, " "))
                             script = script.Substring(0, InStr(script, " ") - 1)
                         End If

                         ' absolute paths begin with \\ or X:\ -> don't prefix with currWB path, else currWBpath\scriptpath
                         Dim curWbPrefix As String = IIf(Left(scriptpath, 2) = "\\" Or Mid(scriptpath, 2, 2) = ":\", "", currWb.Path + "\")
                         fullScriptPath = curWbPrefix + scriptpath

                         ' blocking wait for finish of script dialog
                         Await Task.Run(Sub()
                                            theScriptOutput = New PyOutput()
                                            If theScriptOutput.errMsg <> "" Then Exit Sub
                                            theScriptOutput.ShowInTaskbar = False
                                            theScriptOutput.TopMost = True
                                            ' hide script output if not in debug mode
                                            If Not PyAddin.debugScript Then theScriptOutput.Opacity = 0
                                            theScriptOutput.BringToFront()
                                            theScriptOutput.ShowDialog()
                                            ErrMsg = theScriptOutput.errMsg
                                        End Sub)

                         ScriptRunDic(c) = False
                         ' reflect running state in debug label...
                         PyAddin.theRibbon.InvalidateControl("debug")
                     Next
                     ' reset current dir
                     Directory.SetCurrentDirectory(previousDir)
                 End Function)
        Return True
    End Function

    ''' <summary>encapsulates setting fetching (currently ConfigurationManager from DBAddin.xll.config)</summary>
    ''' <param name="Key">registry key to take value from</param>
    ''' <param name="defaultValue">Value that is taken if Key was not found</param>
    ''' <returns>the setting value</returns>
    Public Function fetchSetting(Key As String, defaultValue As String) As String
        Dim UserSettings As NameValueCollection = Nothing
        Dim AddinAppSettings As NameValueCollection = Nothing
        Try : UserSettings = ConfigurationManager.GetSection("UserSettings") : Catch ex As Exception : LogWarn("Error reading UserSettings: " + ex.Message) : End Try
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : LogWarn("Error reading AppSettings: " + ex.Message) : End Try
        ' user specific settings are in UserSettings section in separate file
        If IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key)) Then
            If Not IsNothing(AddinAppSettings) Then
                fetchSetting = AddinAppSettings(Key)
            Else
                fetchSetting = Nothing
            End If
        ElseIf Not (IsNothing(UserSettings) OrElse IsNothing(UserSettings(Key))) Then
            fetchSetting = UserSettings(Key)
        Else
            fetchSetting = Nothing
        End If
        If fetchSetting Is Nothing Then fetchSetting = defaultValue
    End Function

    ''' <summary>change or add a key/value pair in the user settings</summary>
    ''' <param name="theKey">key to change (or add)</param>
    ''' <param name="theValue">value for key</param>
    Public Sub setUserSetting(theKey As String, theValue As String)
        ' check if key exists
        Dim doc As New Xml.XmlDocument()
        doc.Load(UserSettingsPath)
        Dim keyNode As Xml.XmlNode = doc.SelectSingleNode("/UserSettings/add[@key='" + System.Security.SecurityElement.Escape(theKey) + "']")
        If IsNothing(keyNode) Then
            ' if not, add to settings
            Dim nodeRegion As Xml.XmlElement = doc.CreateElement("add")
            nodeRegion.SetAttribute("key", theKey)
            nodeRegion.SetAttribute("value", theValue)
            doc.SelectSingleNode("//UserSettings").AppendChild(nodeRegion)
        Else
            keyNode.Attributes().GetNamedItem("value").InnerText = theValue
        End If
        doc.Save(UserSettingsPath)
        ConfigurationManager.RefreshSection("UserSettings")
    End Sub

    ''' <summary>Msgbox that avoids further Msgboxes (click Yes) or cancels run altogether (click Cancel)</summary>
    ''' <param name="message"></param>
    ''' <returns>True if further Msgboxes should be avoided, False otherwise</returns>
    Public Function UserMsg(message As String, Optional noAvoidChoice As Boolean = False, Optional IsWarning As Boolean = False) As Boolean
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName & "." & theMethod.Name
        WriteToLog(message, If(IsWarning, EventLogEntryType.Warning, EventLogEntryType.Information), caller)
        If nonInteractive Then Return False
        theRibbon.InvalidateControl("showLog")
        If noAvoidChoice Then
            MsgBox(message, MsgBoxStyle.OkOnly + IIf(IsWarning, MsgBoxStyle.Critical, MsgBoxStyle.Information), "PyAddin Message")
            Return False
        Else
            If avoidFurtherMsgBoxes Then Return True
            Dim retval As MsgBoxResult = MsgBox(message + vbCrLf + "Avoid further Messages (Yes/No) or abort ScriptDefinition (Cancel)", MsgBoxStyle.YesNoCancel, "PyAddin Message")
            If retval = MsgBoxResult.Yes Then avoidFurtherMsgBoxes = True
            Return (retval = MsgBoxResult.Yes Or retval = MsgBoxResult.No)
        End If
    End Function

    ''' <summary>ask User (default OKCancel) and log as warning if Critical (logged errors would pop up the trace information window)</summary> 
    ''' <param name="theMessage">the question to be shown/logged</param>
    ''' <param name="questionType">optionally pass question box type, default MsgBoxStyle.OKCancel</param>
    ''' <param name="questionTitle">optionally pass a title for the msgbox instead of default DBAddin Question</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Question</param>
    ''' <returns>choice as MsgBoxResult (Yes, No, OK, Cancel...)</returns>
    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "PyAddin Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(theMessage, If(msgboxIcon = MsgBoxStyle.Critical, EventLogEntryType.Warning, EventLogEntryType.Information), caller) ' to avoid pop up of trace log
        If nonInteractive Then
            If questionType = MsgBoxStyle.OkCancel Then Return MsgBoxResult.Cancel
            If questionType = MsgBoxStyle.YesNo Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.YesNoCancel Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.RetryCancel Then Return MsgBoxResult.Cancel
        End If
        ' tab is not activated BEFORE Msgbox as Excel first has to get into the interaction thread outside this one..
        If theRibbon IsNot Nothing Then theRibbon.ActivateTab("PyaddinTab")
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As EventLogEntryType, caller As String)
        Dim timestamp As Int32 = DateAndTime.Now().Month * 100000000 + DateAndTime.Now().Day * 1000000 + DateAndTime.Now().Hour * 10000 + DateAndTime.Now().Minute * 100 + DateAndTime.Now().Second

        If nonInteractive Then
            ' collect errors and warnings for returning messages in executeScript
            If eEventType = EventLogEntryType.Error Or eEventType = EventLogEntryType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf
            theLogDisplaySource.TraceEvent(TraceEventType.Warning, timestamp, "Non interactive: {0}: {1}", caller, Message)
        Else
            Select Case eEventType
                Case EventLogEntryType.Information
                    theLogDisplaySource.TraceEvent(TraceEventType.Information, timestamp, "{0}: {1}", caller, Message)
                Case EventLogEntryType.Warning
                    theLogDisplaySource.TraceEvent(TraceEventType.Warning, timestamp, "{0}: {1}", caller, Message)
                    WarningIssued = True
                    ' at Add-in Start ribbon has not been loaded so avoid call to it here..
                    If theRibbon IsNot Nothing Then theRibbon.InvalidateControl("showLog")
                Case EventLogEntryType.Error
                    theLogDisplaySource.TraceEvent(TraceEventType.Error, timestamp, "{0}: {1}", caller, Message)
            End Select
        End If
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
            WriteToLog(LogMessage, EventLogEntryType.Information, caller)
        End If
    End Sub

    Public PyInstallations As List(Of String)

    ''' <summary>initialize the PyInstallations list</summary>
    Public Sub initPyInstallations()
        Dim AddinAppSettings As NameValueCollection = Nothing
        Try : AddinAppSettings = ConfigurationManager.AppSettings : Catch ex As Exception : LogWarn("Error reading AppSettings for PyInstallations (PyLib) entries: " + ex.Message) : End Try
        PyInstallations = New List(Of String)
        ' getting User settings might fail (formatting, etc)...
        If Not IsNothing(AddinAppSettings) Then
            For Each key As String In AddinAppSettings.AllKeys
                If LCase(Left(key, 5)) = "pylib" Then PyInstallations.Add(key.Substring(5))
            Next
        End If
        Try : PyAddin.DebugAddin = fetchSetting("DebugAddin", "False") : Catch Ex As Exception : End Try
    End Sub

    Public Sub insertScriptExample()
        If QuestionMsg("Inserting Example Script definition starting in current cell, overwriting 14 rows and 3 columns with example definitions!") = MsgBoxResult.Cancel Then Exit Sub
        Dim retval As String = InputBox("Please provide a range name:", "Range name for the example (empty name exits this)")
        If retval = "" Then
            Exit Sub
        Else
            retval = "PyScript_" + retval
        End If
        Dim curCell As Range = ExcelDnaUtil.Application.ActiveCell
        curCell.Value = "dir"
        curCell.Offset(0, 1).Value = "."

        curCell.Offset(1, 0).Value = "pyInst"
        curCell.Offset(1, 1).Value = "Python1"
        curCell.Offset(1, 2).Value = "n"

        curCell.Offset(2, 0).Value = "script"
        curCell.Offset(2, 1).Value = "yourScript.py"
        curCell.Offset(2, 2).Value = "subfolder\from\Workbook\dir\where\yourScript.py\Is\located"

        curCell.Offset(3, 0).Value = "scriptCell"
        curCell.Offset(3, 1).Value = "# your script code In this cell"
        curCell.Offset(3, 2).Value = "subfolder\from\Workbook\dir\where\tempfile\For\scriptCell\Is\written"

        curCell.Offset(4, 0).Value = "scriptRng"
        curCell.Offset(4, 1).Value = "yourScriptCodeInThisRange"
        curCell.Offset(4, 2).Value = "."

        curCell.Offset(5, 0).Value = "path"
        curCell.Offset(5, 1).Value = "your\additional\folder\To\add\To\the\path"
        curCell.Offset(5, 2).Value = ""

        curCell.Offset(6, 0).Value = "envvar"
        curCell.Offset(6, 1).Value = "yourEnvironmentVar1"
        curCell.Offset(6, 2).Value = "yourEnvironmentVar1Value"

        curCell.Offset(7, 0).Value = "envvar"
        curCell.Offset(7, 1).Value = "yourEnvironmentVar2"
        curCell.Offset(7, 2).Value = "yourEnvironmentVar2Value"

        curCell.Offset(8, 0).Value = "dir"
        curCell.Offset(8, 1).Value = "your\scriptfiles\directory\overriding\the\current\workbook\folder"

        curCell.Offset(9, 0).Value = "pylib"
        curCell.Offset(9, 1).Value = "yourOwnOverridingPython.dll"
        Try
            ExcelDnaUtil.Application.ActiveSheet.Range(curCell, curCell.Offset(13, 2)).Name = retval
        Catch ex As Exception
            UserMsg("Couldn't name example definitions as '" + retval + "': " + ex.Message)
        End Try
        ExcelDnaUtil.Application.ActiveSheet.Range(curCell, curCell.Offset(0, 2)).EntireColumn.AutoFit
        PyAddin.initPyInstallations()
        Dim errStr As String = PyAddin.startScriptNamesRefresh()
        If Len(errStr) > 0 Then PyAddin.UserMsg("refresh Error: " & errStr, True, True)
    End Sub

End Module
