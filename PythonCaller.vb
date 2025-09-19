Option Strict Off
Imports Python.Runtime

Public Module PythonCaller

    ' initialise python environment with PythonDLL of environment and optionally the path to a virtual python environment
    Public Function InitPython() As Boolean

        ' set python engine and init
        Try
            Environment.SetEnvironmentVariable("PYTHONNET_PYDLL", PyAddin.PyLib, EnvironmentVariableTarget.Process)
            Runtime.PythonDLL = PyAddin.PyLib
            PythonEngine.Initialize()
            PythonEngine.BeginAllowThreads()
        Catch Ex As Exception
            PyAddin.UserMsg(String.Format("Error initialising Python (lib: " + PyAddin.PyLib + "): {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
            Return False
        End Try
        ' activate global interpreter lock in runtime, to append sys.path
        Dim gil As Py.GILState
        Try
            gil = Py.GIL()
        Catch Ex As Exception
            PyAddin.UserMsg(String.Format("Error when getting Python GIL: {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
            Return False
        End Try
        Dim pathToVirtualEnv As String = PyAddin.PyVenv
        Try
            Dim _sys As Object = Py.Import("sys")
            _sys.path.append(pathToVirtualEnv)
        Catch Ex As Exception
            PyAddin.UserMsg(String.Format("Error when setting sys path for " + pathToVirtualEnv + ": {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
            gil.Dispose()
            Return False
        End Try
        ' always unlock gil !
        gil.Dispose()
        Return True
    End Function


    Public Function CallPythonScript(file As String) As Boolean
        Dim resultCall As Boolean = False
        ' Global interpreter lock im runtime aktivieren
        Dim gil As Py.GILState
        Try
            gil = Py.GIL()
        Catch Ex As Exception
            PyAddin.UserMsg(String.Format("Error when getting Python GIL: {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
            Return False
        End Try

        Using scope As Object = Py.CreateScope()

            Dim scriptCompiled As PyObject
            ' python scriptfile als text holen und kompilieren
            Try
                Dim code As String = System.IO.File.ReadAllText(file)
                scriptCompiled = PythonEngine.Compile(code, file)
            Catch Ex As Exception
                PyAddin.UserMsg(String.Format("Error reading/compiling script " + file + ": {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
                gil.Dispose()
                Return resultCall
            End Try

            ' execute compiled python
            Try
                ' Das ist auch der Zeitpunkt, an dem der debugger (wenn aktiv) gestartet wird.
                scope.Execute(scriptCompiled)
                If (PyAddin.debugScript) Then
                    Dim debugListen As PyObject = PythonEngine.Compile("import debugpy" + vbCrLf + "debugpy.listen((""127.0.0.1"",5678), in_process_debug_adapter=True)" + vbCrLf + "print(""waiting for debugger to attach..."")" + vbCrLf + "debugpy.wait_for_client()")
                    scope.Execute(debugListen)
                End If
            Catch Ex As Exception
                PyAddin.UserMsg(String.Format("Error when executing script " + file + ": {0}, Src:{1}, Trace:{2}", Ex.Message, Ex.Source, Ex.StackTrace))
                gil.Dispose()
                Return resultCall
            End Try

        End Using
        gil.Dispose()
        Return resultCall

    End Function



End Module
