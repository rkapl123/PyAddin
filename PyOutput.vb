Imports System.Windows.Forms
Imports System.IO
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Collections.Generic

''' <summary>Dialog for script invocation and output of standard out and standard err, allowing standard input as well</summary>
Public Class PyOutput
    Public cmd As Process
    Public errMsg As String = ""

    ' mapping for ANSI colors (VT 100 terminal), taken and adapted from https://ss64.com/nt/syntax-ansi.html
    Private ReadOnly fgColDic As New Dictionary(Of String, System.Drawing.Color) From {
{"30", System.Drawing.Color.Black},
{"31", System.Drawing.Color.IndianRed},
{"32", System.Drawing.Color.Green},
{"33", System.Drawing.Color.Yellow},
{"34", System.Drawing.Color.Blue},
{"35", System.Drawing.Color.Magenta},
{"36", System.Drawing.Color.Cyan},
{"37", System.Drawing.Color.LightGray},
{"90", System.Drawing.Color.DarkGray},
{"91", System.Drawing.Color.Red},
{"92", System.Drawing.Color.LightGreen},
{"93", System.Drawing.Color.LightYellow},
{"94", System.Drawing.Color.LightBlue},
{"95", System.Drawing.Color.Magenta},
{"96", System.Drawing.Color.LightCyan},
{"97", System.Drawing.Color.White}}

    Private ReadOnly bgColDic As New Dictionary(Of String, System.Drawing.Color) From {
{"40", System.Drawing.Color.Black},
{"41", System.Drawing.Color.DarkRed},
{"42", System.Drawing.Color.DarkGreen},
{"43", System.Drawing.Color.GreenYellow},
{"44", System.Drawing.Color.DarkBlue},
{"45", System.Drawing.Color.DarkMagenta},
{"46", System.Drawing.Color.DarkCyan},
{"47", System.Drawing.Color.LightGray},
{"100", System.Drawing.Color.DarkGray},
{"101", System.Drawing.Color.PaleVioletRed},
{"102", System.Drawing.Color.LightSeaGreen},
{"103", System.Drawing.Color.LightGoldenrodYellow},
{"104", System.Drawing.Color.LightSkyBlue},
{"105", System.Drawing.Color.Magenta},
{"106", System.Drawing.Color.DarkCyan},
{"107", System.Drawing.Color.White}}

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        Try
            Directory.SetCurrentDirectory(PyAddin.fullScriptPath)
            PythonCaller.CallPythonScript(PyAddin.fullScriptPath + "\" + PyAddin.script)

        Catch ex As Exception
            PyAddin.UserMsg("Error occurred when invoking script '" + PyAddin.fullScriptPath + "\" + PyAddin.script + "', using '" + PyAddin.PyLib + "'" + ex.Message + vbCrLf, True, True)
            Me.errMsg = ex.Message
        End Try

    End Sub

    Private Sub myOutHandler(sender As Object, e As DataReceivedEventArgs)
        If IsNothing(e.Data) Then Exit Sub
        Dim msgtext As String = e.Data
        Dim fgCol As String = "" : Dim bgCol As String = ""
        Dim fgColWin As System.Drawing.Color = System.Drawing.Color.White : Dim bgColWin As System.Drawing.Color = System.Drawing.Color.Black
        If Strings.Left(e.Data, 1) = ChrW(27) Then
            fgCol = Strings.Mid(e.Data, 3, 2)
            If Strings.Mid(e.Data, 5, 1) = ";" Then
                bgCol = Strings.Mid(e.Data, 6, 3)
                bgCol = bgCol.Replace("m", "")
            End If
            msgtext = Strings.Mid(msgtext, Strings.InStr(msgtext, "m") + 1)
            msgtext = Strings.Left(msgtext, Strings.InStr(msgtext, ChrW(27)) - 1)
        End If
        Try
            fgColWin = fgColDic(fgCol)
            bgColWin = bgColDic(bgCol)
        Catch ex As Exception : End Try
        LogInfo("script out: " + msgtext)
        Dim appendAction As Action(Of String, System.Drawing.Color, System.Drawing.Color) = AddressOf appendTxt
        appendAction.Invoke(msgtext + vbCrLf, fgColWin, bgColWin)
    End Sub

    'redirect stdout/err
    'https://stackoverflow.com/questions/37289082/redirect-stdout-stderr-from-python-embedded-in-net-application-to-text-box/76120954#76120954
    Private Sub myErrHandler(sender As Object, e As DataReceivedEventArgs)
        If IsNothing(e.Data) Then Exit Sub
        Dim appendAction As Action(Of String, System.Drawing.Color, System.Drawing.Color) = AddressOf appendTxt
        If PyAddin.StdErrMeansError Then
            Me.errMsg += e.Data
            LogWarn("script error: " + e.Data)
        Else
            ' need this, otherwise appendTxt has synchronisations problems for stderr...
            LogWarn("script error: " + e.Data)
        End If
        appendAction.Invoke(e.Data + vbCrLf, System.Drawing.Color.Red, System.Drawing.Color.Black)
    End Sub

    Private Sub myExitHandler(sender As Object, e As System.EventArgs)
        LogInfo("finished " + PyAddin.fullScriptPath + "\" + PyAddin.script + "', using '" + PyAddin.PyLib + "'")
        ' need this line to wait for stdout/stderr to finish writing...
        Try : cmd.WaitForExit() : Catch ex As Exception
            LogWarn("cmd.WaitForExit exception " + ex.Message)
        End Try

        If PyAddin.debugScript Then
            Dim appendAction As Action(Of String, System.Drawing.Color, System.Drawing.Color) = AddressOf appendTxt
            appendAction.Invoke("Finished script execution, exit code: " + cmd.ExitCode.ToString(), System.Drawing.Color.Yellow, System.Drawing.Color.Black)
        End If
        cmd = Nothing
        If Not PyAddin.debugScript Then Me.Close()
    End Sub

    Private Sub appendTxt(theText As String, textCol As System.Drawing.Color, backCol As System.Drawing.Color)
        Dim pos As Integer = ScriptOutputTextbox.TextLength
        ScriptOutputTextbox.AppendText(theText)
        If textCol <> System.Drawing.Color.White Or backCol <> System.Drawing.Color.Black Then
            ' select text to be colored
            ScriptOutputTextbox.Select(pos, theText.Length)
            ScriptOutputTextbox.SelectionColor = textCol
            ScriptOutputTextbox.SelectionBackColor = backCol
            ' deselect the text
            ScriptOutputTextbox.Select(pos + theText.Length, 0)
            ScriptOutputTextbox.AppendText("")
        End If
    End Sub

    Private Sub ScriptOutput_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If Not IsNothing(cmd) Then
            If e.KeyCode = Keys.Escape Then
                Me.Close() ' cmd cleanup should be resolved in ScriptOutput_Closing
            ElseIf e.KeyCode = Keys.Enter Then
                cmd.StandardInput.WriteLine()
            End If
        End If
    End Sub

    Private Sub ScriptOutput_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If Not IsNothing(cmd) AndAlso Not IsNothing(cmd.StandardInput) Then
            cmd.StandardInput.Write(e.KeyChar)
        End If
    End Sub

    Private Sub ScriptOutput_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not IsNothing(cmd) Then
            LogInfo("terminating " + PyAddin.fullScriptPath + "\" + PyAddin.script + "', using '" + PyAddin.PyLib + "'")
            ' check if process has not exited
            Dim procFinished As Boolean = False
            Try : procFinished = cmd.HasExited : Catch ex As Exception : End Try
            If Not procFinished Then
                If PyAddin.QuestionMsg("Process still running, kill it (cancel leaves this pane open)?", MsgBoxStyle.OkCancel, "Process running", MsgBoxStyle.Exclamation) = MsgBoxResult.Ok Then
                    Try
                        cmd.Kill()
                    Catch ex As Exception
                        LogWarn("cmd.Kill exception: " + ex.Message)
                    End Try
                Else
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

End Class