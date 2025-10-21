Option Compare Database
Option Explicit

' ============================================================================
'  modAccessObjectRoundtrip
'  - Export objects FROM an external Access DB to a folder.
'  - Import objects FROM a folder INTO an external Access DB (create if missing).
'
'  Naming for round-trip:
'     Form_{Name}.txt
'     Report_{Name}.txt
'     Macro_{Name}.txt
'     Module_{Name}.txt   (SaveAsText flavor; optional on export/import)
'     Module_{Name}.bas   (native std module; optional on export/import)
'     Module_{Name}.cls   (native class module; optional on export/import)
'     Query_{Name}.txt    (SaveAsText flavor)
'     Query_{Name}.sql    (raw SQL; optional on export/import)
'     References.txt      (library manifest; handled inline during import)
'
'  Hardening:
'   - BOM-aware file reader (fixes error 62).
'   - VBProjects is 1-based; guarded access (fixes error 9).
'   - VBOM not required; VB import/export skipped gracefully with log.
'   - Empty .sql skipped with WARN.
'   - Quiet automation; no SaveAllModules call; rely on acQuitSaveAll (no SaveAs prompts).
'   - References round-trip integrated in folder walker.
' ============================================================================


'=============================================================================
' PUBLIC: EXPORT (External DB -> Folder)
'=============================================================================
' ExportModuleText: controls Module_{Name}.txt (SaveAsText)
' ExportVbNative  : controls .bas/.cls native module export
Public Sub ExportAllObjects_External( _
    ByVal SourceDatabasePath As String, _
    ByVal OutputFolderPath As String, _
    Optional ByVal ExportModuleText As Boolean = True, _
    Optional ByVal ExportVbNative As Boolean = False, _
    Optional ByVal ExportQueriesAsSql As Boolean = True, _
    Optional ByVal IncludeSystemQueries As Boolean = False, _
    Optional ByVal IncludeReferences As Boolean = True, _
    Optional ByVal ShowSummary As Boolean = True)

    Dim exported As Long, failed As Long
    Dim accApp As Object ' Access.Application

    If Len(OutputFolderPath) = 0 Then
        MsgBox "Please provide an output folder.", vbExclamation
        Exit Sub
    End If
    EnsureFolderExists OutputFolderPath

    If Len(Dir$(SourceDatabasePath)) = 0 Then
        MsgBox "Source database not found: " & SourceDatabasePath, vbExclamation
        Exit Sub
    End If

    Set accApp = CreateObject("Access.Application")
    accApp.OpenCurrentDatabase SourceDatabasePath
    BeginAutomation accApp

    On Error GoTo CLEANUP

    EnsureExportLogTable_Local

    ExportForms_External accApp, OutputFolderPath, exported, failed
    ExportReports_External accApp, OutputFolderPath, exported, failed
    ExportMacros_External accApp, OutputFolderPath, exported, failed
    ExportModules_External accApp, OutputFolderPath, ExportModuleText, ExportVbNative, exported, failed
    ExportQueries_External accApp, OutputFolderPath, ExportQueriesAsSql, IncludeSystemQueries, exported, failed

    If IncludeReferences Then
        ExportVBAReferences accApp, OutputFolderPath, exported, failed
    End If

    If ShowSummary Then
        MsgBox "External Export complete ? " & SourceDatabasePath & vbCrLf & _
               "Exported: " & exported & vbCrLf & _
               "Failed:   " & failed, vbInformation
    End If

CLEANUP:
    On Error Resume Next
    If Not accApp Is Nothing Then
        ' No SaveAllModules; quit commits VB project silently
        EndAutomation accApp
        accApp.Quit acQuitSaveAll
    End If
    Set accApp = Nothing
End Sub


'=============================================================================
' PUBLIC: IMPORT (Folder -> External DB, create if missing)
'=============================================================================
Public Sub ImportFolderObjects_External( _
    ByVal FolderPath As String, _
    ByVal TargetDatabasePath As String, _
    Optional ByVal CreateIfMissing As Boolean = True, _
    Optional ByVal OverwriteExisting As Boolean = True, _
    Optional ByVal Recurse As Boolean = False, _
    Optional ByVal IncludeReferences As Boolean = True, _
    Optional ByVal ShowSummary As Boolean = True)

    Dim imported As Long, failed As Long
    Dim accApp As Object ' Access.Application
    Dim TargetExisted As Boolean

    If Len(Dir$(FolderPath, vbDirectory)) = 0 Then
        MsgBox "Folder not found: " & FolderPath, vbExclamation
        Exit Sub
    End If

    Set accApp = GetAccessAppForPath(TargetDatabasePath, CreateIfMissing, TargetExisted)
    BeginAutomation accApp

    On Error GoTo CLEANUP

    EnsureImportLogTable_Local

    Dim fso As Object, fld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(FolderPath)

    ProcessFolder_External accApp, fld, OverwriteExisting, Recurse, IncludeReferences, imported, failed

    If ShowSummary Then
        MsgBox "External Import complete ? " & TargetDatabasePath & vbCrLf & _
               "Imported: " & imported & vbCrLf & _
               "Failed:   " & failed, vbInformation
    End If

CLEANUP:
    On Error Resume Next
    If Not accApp Is Nothing Then
        EndAutomation accApp
        accApp.Quit acQuitSaveAll
    End If
    Set accApp = Nothing
End Sub


'=============================================================================
' EXPORT: Per-type workers (External)
'=============================================================================
Private Sub ExportForms_External(ByVal accApp As Object, ByVal Root As String, ByRef exported As Long, ByRef failed As Long)
    Dim ao As Object, p As String
    For Each ao In accApp.CurrentProject.AllForms
        p = BuildPath(Root, "Form_" & SafeName(ao.name) & ".txt")
        On Error GoTo EH
        accApp.SaveAsText acForm, ao.name, p
        LogExport_Local ao.name, "OK", "Form ? " & p
        exported = exported + 1
NextForm:
        On Error GoTo 0
    Next ao
    Exit Sub
EH:
    failed = failed + 1
    LogExport_Local ao.name, "ERROR", Err.Number & " - " & Err.Description
    Resume NextForm
End Sub

Private Sub ExportReports_External(ByVal accApp As Object, ByVal Root As String, ByRef exported As Long, ByRef failed As Long)
    Dim ao As Object, p As String
    For Each ao In accApp.CurrentProject.AllReports
        p = BuildPath(Root, "Report_" & SafeName(ao.name) & ".txt")
        On Error GoTo EH
        accApp.SaveAsText acReport, ao.name, p
        LogExport_Local ao.name, "OK", "Report ? " & p
        exported = exported + 1
NextRpt:
        On Error GoTo 0
    Next ao
    Exit Sub
EH:
    failed = failed + 1
    LogExport_Local ao.name, "ERROR", Err.Number & " - " & Err.Description
    Resume NextRpt
End Sub

Private Sub ExportMacros_External(ByVal accApp As Object, ByVal Root As String, ByRef exported As Long, ByRef failed As Long)
    Dim ao As Object, p As String
    For Each ao In accApp.CurrentProject.AllMacros
        p = BuildPath(Root, "Macro_" & SafeName(ao.name) & ".txt")
        On Error GoTo EH
        accApp.SaveAsText acMacro, ao.name, p
        LogExport_Local ao.name, "OK", "Macro ? " & p
        exported = exported + 1
NextMac:
        On Error GoTo 0
    Next ao
    Exit Sub
EH:
    failed = failed + 1
    LogExport_Local ao.name, "ERROR", Err.Number & " - " & Err.Description
    Resume NextMac
End Sub

Private Sub ExportModules_External(ByVal accApp As Object, ByVal Root As String, ByVal ExportModuleText As Boolean, ByVal ExportVbNative As Boolean, ByRef exported As Long, ByRef failed As Long)
    Dim vbproj As Object, vbcomps As Object, vbcomp As Object
    Set vbproj = GetVBProject(accApp)
    If vbproj Is Nothing Then
        LogExport_Local "(VB)", "SKIP", "VB export skipped (Trust access to VBOM not enabled)"
        Exit Sub
    End If

    Set vbcomps = vbproj.VBComponents

    Dim compName As String, txtPath As String, basPath As String

    For Each vbcomp In vbcomps
        ' Skip form/report "document" modules and MSForms
        If vbcomp.Type = 100 Or vbcomp.Type = 34 Then GoTo NextComp

        compName = vbcomp.name

        If ExportModuleText Then
            txtPath = BuildPath(Root, "Module_" & SafeName(compName) & ".txt")
            On Error GoTo EH
            accApp.SaveAsText acModule, compName, txtPath
            LogExport_Local compName, "OK", "Module(txt) ? " & txtPath
            exported = exported + 1
            On Error GoTo 0
        End If

        If ExportVbNative Then
            On Error GoTo EH
            Select Case vbcomp.Type
                Case 1 ' Std module
                    basPath = BuildPath(Root, "Module_" & SafeName(compName) & ".bas")
                Case 2 ' Class module
                    basPath = BuildPath(Root, "Module_" & SafeName(compName) & ".cls")
                Case Else
                    GoTo NextComp
            End Select
            vbcomp.Export basPath
            LogExport_Local compName, "OK", "VB ? " & basPath
            exported = exported + 1
            On Error GoTo 0
        End If

NextComp:
    Next vbcomp
    Exit Sub
EH:
    failed = failed + 1
    LogExport_Local compName, "ERROR", Err.Number & " - " & Err.Description
    Resume NextComp
End Sub

Private Sub ExportQueries_External(ByVal accApp As Object, ByVal Root As String, ByVal ExportSql As Boolean, ByVal IncludeSystem As Boolean, ByRef exported As Long, ByRef failed As Long)
    Dim db As DAO.Database: Set db = accApp.CurrentDb
    Dim qd As DAO.QueryDef
    Dim txtPath As String, sqlPath As String

    For Each qd In db.QueryDefs
        If ShouldSkipQuery(qd, IncludeSystem) Then GoTo NextQ

        txtPath = BuildPath(Root, "Query_" & SafeName(qd.name) & ".txt")
        On Error GoTo EH
        accApp.SaveAsText acQuery, qd.name, txtPath
        LogExport_Local qd.name, "OK", "Query(txt) ? " & txtPath
        exported = exported + 1
        On Error GoTo 0

        If ExportSql Then
            On Error GoTo EH
            sqlPath = BuildPath(Root, "Query_" & SafeName(qd.name) & ".sql")
            WriteAllText sqlPath, qd.sql
            LogExport_Local qd.name, "OK", "Query(sql) ? " & sqlPath
            exported = exported + 1
            On Error GoTo 0
        End If

NextQ:
    Next qd
    Exit Sub
EH:
    failed = failed + 1
    LogExport_Local qd.name, "ERROR", Err.Number & " - " & Err.Description
    Resume NextQ
End Sub

Private Sub ExportVBAReferences(ByVal accApp As Object, ByVal OutputFolder As String, ByRef exported As Long, ByRef failed As Long)
    On Error Resume Next
    Dim vbproj As Object, ref As Object
    Set vbproj = GetVBProject(accApp)
    If vbproj Is Nothing Then
        LogExport_Local "(References)", "SKIP", "Skipped (VBOM not trusted)"
        Exit Sub
    End If

    Dim refList As String
    For Each ref In vbproj.References
        refList = refList & ref.name & "|" & ref.FullPath & "|" & ref.Major & "." & ref.Minor & vbCrLf
    Next

    WriteAllText BuildPath(OutputFolder, "References.txt"), refList
    LogExport_Local "(References)", "OK", "Exported " & vbproj.References.Count & " references"
    exported = exported + 1
End Sub


'=============================================================================
' IMPORT: Folder walker (External)
' - Handles References.txt inline so it's counted as success.
'=============================================================================
Private Sub ProcessFolder_External( _
    ByVal accApp As Object, _
    ByVal fld As Object, _
    ByVal OverwriteExisting As Boolean, _
    ByVal Recurse As Boolean, _
    ByVal IncludeReferences As Boolean, _
    ByRef imported As Long, _
    ByRef failed As Long)

    Dim f As Object, subFld As Object
    For Each f In fld.Files
        If Not ShouldSkipFile(CStr(f.name)) Then

            ' Handle References.txt inline
            If IncludeReferences And LCase$(CStr(f.name)) = "references.txt" Then
                On Error GoTo RefErr
                ImportVBAReferences accApp, CStr(f.path)
                imported = imported + 1
                GoTo NextFile
RefErr:
                failed = failed + 1
                LogImport_Local CStr(f.path), "ERROR", Err.Number & " - " & Err.Description
                Resume NextFile
            End If

            On Error GoTo HandleErr
            If ImportOneFile_External(accApp, CStr(f.path), OverwriteExisting) Then
                imported = imported + 1
            End If
            On Error GoTo 0
        End If
NextFile:
    Next

    If Recurse Then
        For Each subFld In fld.SubFolders
            ProcessFolder_External accApp, subFld, OverwriteExisting, True, IncludeReferences, imported, failed
        Next
    End If
    Exit Sub

HandleErr:
    failed = failed + 1
    LogImport_Local CStr(f.path), "ERROR", Err.Number & " - " & Err.Description
    Resume NextFile
End Sub

Private Function ImportOneFile_External( _
    ByVal accApp As Object, _
    ByVal FilePath As String, _
    ByVal OverwriteExisting As Boolean) As Boolean

    Dim fName As String, ObjName As String, TypeToken As String, Ext As String
    Dim ObjType As AcObjectType, detected As Boolean

    fName = Dir$(FilePath)
    Ext = LCase$(GetFileExtension(fName))

    detected = ParseTypeAndNameFromFileName(fName, TypeToken, ObjName)
    If detected Then
        ObjType = MapTypeTokenToAcType(TypeToken)
        If ObjType <> acDefault Then
            ImportOneFileCore_External accApp, FilePath, ObjType, ObjName, Ext, OverwriteExisting
            ImportOneFile_External = True
            Exit Function
        End If
    End If

    If Ext = "txt" Then
        If DetectTypeFromTextFile(FilePath, ObjType, ObjName) Then
            If Len(ObjName) = 0 Then ObjName = StripExtension(fName)
            ImportOneFileCore_External accApp, FilePath, ObjType, ObjName, Ext, OverwriteExisting
            ImportOneFile_External = True
            Exit Function
        End If
    End If

    If Ext = "bas" Or Ext = "cls" Then
        ObjName = DeriveNameFromModuleFile(fName)
        ImportVbComponent_External accApp, FilePath, ObjName, OverwriteExisting
        LogImport_Local FilePath, "OK", "Imported VB component: " & ObjName
        ImportOneFile_External = True
        Exit Function
    End If

    If Ext = "sql" Then
        If Not detected Then detected = ParseTypeAndNameFromFileName(fName, TypeToken, ObjName)
        If Not detected Or LCase$(TypeToken) <> "query" Then
            LogImport_Local FilePath, "SKIP", "Expected Query_qryName.sql"
            Exit Function
        End If
        Dim SqlText As String: SqlText = ReadAllText(FilePath)
        If LenB(SqlText) = 0 Then
            LogImport_Local FilePath, "WARN", "Empty SQL file"
            Exit Function
        End If
        CreateOrReplaceQuery_External accApp, ObjName, SqlText, OverwriteExisting
        LogImport_Local FilePath, "OK", "Created/updated Query: " & ObjName
        ImportOneFile_External = True
        Exit Function
    End If

    LogImport_Local FilePath, "SKIP", "Unrecognized file type or naming"
End Function

Private Sub ImportOneFileCore_External( _
    ByVal accApp As Object, _
    ByVal FilePath As String, _
    ByVal ObjType As AcObjectType, _
    ByVal ObjName As String, _
    ByVal Ext As String, _
    ByVal OverwriteExisting As Boolean)

    If ObjectExists_External(accApp, ObjType, ObjName) Then
        If OverwriteExisting Then
            accApp.DoCmd.DeleteObject ObjType, ObjName
        Else
            LogImport_Local FilePath, "SKIP", "Exists: " & ObjName
            Exit Sub
        End If
    End If

    Select Case Ext
        Case "txt"
            accApp.LoadFromText ObjType, ObjName, FilePath
            LogImport_Local FilePath, "OK", "Loaded via LoadFromText as " & ObjName

        Case "bas", "cls"
            ImportVbComponent_External accApp, FilePath, ObjName, OverwriteExisting

        Case "sql"
            CreateOrReplaceQuery_External accApp, ObjName, ReadAllText(FilePath), True

        Case Else
            Err.Raise vbObjectError + 613, , "Unsupported extension: " & Ext
    End Select
End Sub


'=============================================================================
' EXTERNAL DB helpers (existence, queries, VB import, bootstrap)
'=============================================================================
Private Function ObjectExists_External(ByVal accApp As Object, ByVal ObjType As AcObjectType, ByVal ObjName As String) As Boolean
    On Error Resume Next
    ObjectExists_External = (accApp.SysCmd(acSysCmdGetObjectState, ObjType, ObjName) <> 0)
    On Error GoTo 0
End Function

Private Sub CreateOrReplaceQuery_External(ByVal accApp As Object, ByVal QueryName As String, ByVal SqlText As String, ByVal Overwrite As Boolean)
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef

    Set db = accApp.CurrentDb
    On Error Resume Next
    Set qd = db.QueryDefs(QueryName)
    On Error GoTo 0

    If Not qd Is Nothing Then
        If Overwrite Then
            qd.sql = SqlText
            qd.Close
        Else
            Err.Raise vbObjectError + 614, , "Query exists: " & QueryName
        End If
    Else
        Set qd = db.CreateQueryDef(QueryName, SqlText)
        qd.Close
    End If
End Sub

Private Sub ImportVbComponent_External(ByVal accApp As Object, ByVal FilePath As String, ByVal DesiredName As String, ByVal OverwriteExisting As Boolean)
    Dim vbproj As Object, vbcomps As Object, vbcomp As Object

    Set vbproj = GetVBProject(accApp)
    If vbproj Is Nothing Then
        LogImport_Local FilePath, "SKIP", "VB import skipped (Trust access to VBOM not enabled)"
        Exit Sub
    End If

    Set vbcomps = vbproj.VBComponents

    ' Remove existing if requested
    Set vbcomp = Nothing
    On Error Resume Next
    Set vbcomp = vbcomps.Item(DesiredName)
    On Error GoTo 0
    If Not vbcomp Is Nothing Then
        If OverwriteExisting Then
            vbcomps.Remove vbcomp
        Else
            LogImport_Local FilePath, "SKIP", "VB component exists: " & DesiredName
            Exit Sub
        End If
    End If

    vbcomps.Import FilePath

    ' If imported name differs from DesiredName, rename (no SaveAs prompts here)
    On Error Resume Next
    Dim base As String: base = GetFileBaseName(Dir$(FilePath))
    Set vbcomp = vbcomps.Item(base)
    If Not vbcomp Is Nothing Then
        If StrComp(vbcomp.name, DesiredName, vbTextCompare) <> 0 Then
            vbcomp.name = DesiredName
        End If
    End If
    On Error GoTo 0
End Sub

Private Function GetAccessAppForPath(ByVal DbPath As String, ByVal CreateIfMissing As Boolean, ByRef TargetExisted As Boolean) As Object
    Dim accApp As Object: Set accApp = CreateObject("Access.Application")
    TargetExisted = (Len(Dir$(DbPath)) > 0)

    If TargetExisted Then
        accApp.OpenCurrentDatabase DbPath
    Else
        If Not CreateIfMissing Then
            Err.Raise vbObjectError + 620, , "Target database not found and CreateIfMissing=False: " & DbPath
        End If
        accApp.NewCurrentDatabase DbPath
    End If

    Set GetAccessAppForPath = accApp
End Function


'=============================================================================
' VB/VBE helpers (1-based VBProjects & trust guard)
'=============================================================================
Private Function GetVBProject(ByVal accApp As Object) As Object
    On Error GoTo Nope
    Dim vbproj As Object
    Set vbproj = accApp.VBE.VBProjects(1) ' 1-based in Access
    Set GetVBProject = vbproj
    Exit Function
Nope:
    Set GetVBProject = Nothing
End Function


'=============================================================================
' References import (triggered inline when encountering References.txt)
'=============================================================================
Private Sub ImportVBAReferences(ByVal accApp As Object, ByVal RefFilePath As String)
    If Len(Dir$(RefFilePath)) = 0 Then Exit Sub

    Dim vbproj As Object
    Set vbproj = GetVBProject(accApp)
    If vbproj Is Nothing Then
        LogImport_Local RefFilePath, "SKIP", "References import skipped (VBOM not trusted)"
        Exit Sub
    End If

    Dim lines() As String, ln As Variant
    Dim parts() As String, path As String

    lines = Split(ReadAllText(RefFilePath), vbCrLf)
    For Each ln In lines
        If Len(Trim$(CStr(ln))) = 0 Then GoTo NextLn
        parts = Split(CStr(ln), "|")
        If UBound(parts) >= 1 Then
            path = parts(1)
            If Len(Dir$(path)) > 0 Then
                On Error Resume Next
                vbproj.References.AddFromFile path
                If Err.Number <> 0 Then
                    LogImport_Local path, "WARN", "Could not add reference: " & Err.Description
                    Err.Clear
                Else
                    LogImport_Local path, "OK", "Added reference"
                End If
                On Error GoTo 0
            Else
                LogImport_Local path, "WARN", "Missing reference file"
            End If
        End If
NextLn:
    Next
End Sub


'=============================================================================
' Quiet automation (avoid prompts)
'=============================================================================
Private Sub BeginAutomation(ByVal accApp As Object)
    On Error Resume Next
    accApp.Visible = False
    accApp.Echo False
    accApp.DoCmd.SetWarnings False
    accApp.SetOption "Confirm Record Changes", False
    accApp.SetOption "Confirm Action Queries", False
    accApp.SetOption "Confirm Document Deletions", False
    accApp.SetOption "Auto Compact On Close", False
    ' Close any VBE windows to avoid designer prompts
    Dim w As Object
    For Each w In accApp.VBE.Windows
        w.Close
    Next
End Sub

Private Sub EndAutomation(ByVal accApp As Object)
    On Error Resume Next
    accApp.DoCmd.SetWarnings True
    accApp.Echo True
End Sub


'=============================================================================
' Query skip logic (system/temp)
'=============================================================================
Private Function ShouldSkipQuery(ByVal qd As DAO.QueryDef, ByVal IncludeSystem As Boolean) As Boolean
    Dim nm As String: nm = qd.name
    If Left$(nm, 1) = "~" Then ShouldSkipQuery = True: Exit Function
    If Not IncludeSystem Then
        If Left$(nm, 4) = "MSys" Then ShouldSkipQuery = True: Exit Function
    End If
End Function


'=============================================================================
' Logging (tables in the CALLER database)
'=============================================================================
Private Sub EnsureExportLogTable_Local()
    On Error Resume Next
    CurrentDb.Execute _
        "CREATE TABLE ExportLog (" & _
        "   LogID AUTOINCREMENT CONSTRAINT PK_ExportLog PRIMARY KEY," & _
        "   LogTime DATETIME," & _
        "   ObjectName TEXT(128)," & _
        "   Status TEXT(20)," & _
        "   Message TEXT(255)" & _
        ")", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub EnsureImportLogTable_Local()
    On Error Resume Next
    CurrentDb.Execute _
        "CREATE TABLE ImportLog (" & _
        "   LogID AUTOINCREMENT CONSTRAINT PK_ImportLog PRIMARY KEY," & _
        "   LogTime DATETIME," & _
        "   FilePath TEXT(255)," & _
        "   Status TEXT(20)," & _
        "   Message TEXT(255)" & _
        ")", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub LogExport_Local(ByVal ObjectName As String, ByVal Status As String, ByVal Message As String)
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ExportLog (LogTime, ObjectName, Status, Message) VALUES (" & _
                      "Now(), " & Quote(Left$(ObjectName, 128)) & ", " & _
                      Quote(Left$(Status, 20)) & ", " & _
                      Quote(Left$(Message, 255)) & ")", dbFailOnError
    Debug.Print Status & " | " & ObjectName & " | " & Message
    On Error GoTo 0
End Sub

Private Sub LogImport_Local(ByVal FilePath As String, ByVal Status As String, ByVal Message As String)
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ImportLog (LogTime, FilePath, Status, Message) VALUES (" & _
                      "Now(), " & Quote(Left$(FilePath, 255)) & ", " & _
                      Quote(Left$(Status, 20)) & ", " & _
                      Quote(Left$(Message, 255)) & ")", dbFailOnError
    Debug.Print Status & " | " & FilePath & " | " & Message
    On Error GoTo 0
End Sub


'=============================================================================
' IO & parsing utilities (BOM-aware)
'=============================================================================
Private Function ReadAllText(ByVal FilePath As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(FilePath, 1, False, -2) ' ForReading, TristateUseDefault
    ReadAllText = ts.ReadAll
    ts.Close
End Function

Private Sub WriteAllText(ByVal FilePath As String, ByVal Content As String)
    Dim f As Integer: f = FreeFile
    Open FilePath For Output As #f
    Print #f, Content
    Close #f
End Sub

Private Sub EnsureFolderExists(ByVal p As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then fso.CreateFolder p
End Sub

Private Function BuildPath(ByVal folder As String, ByVal file As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        BuildPath = folder & file
    Else
        BuildPath = folder & "\" & file
    End If
End Function

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace$(s, CStr(bad), "_")
    Next
    SafeName = s
End Function

Private Function GetFileExtension(ByVal FileName As String) As String
    Dim p As Long: p = InStrRev(FileName, ".")
    If p > 0 Then GetFileExtension = Mid$(FileName, p + 1)
End Function

Private Function GetFileBaseName(ByVal FileName As String) As String
    Dim p As Long: p = InStrRev(FileName, ".")
    If p > 0 Then
        GetFileBaseName = Left$(FileName, p - 1)
    Else
        GetFileBaseName = FileName
    End If
End Function

Private Function StripExtension(ByVal FileName As String) As String
    StripExtension = GetFileBaseName(FileName)
End Function

Private Function Quote(ByVal s As String) As String
    Quote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function ShouldSkipFile(ByVal FileName As String) As Boolean
    ShouldSkipFile = (Left$(FileName, 2) = "~$" Or Left$(FileName, 1) = "." Or LCase$(Right$(FileName, 4)) = ".bak")
End Function

Private Function ParseTypeAndNameFromFileName( _
    ByVal FileName As String, _
    ByRef TypeToken As String, _
    ByRef ObjName As String) As Boolean

    Dim base As String, us As Long
    base = StripExtension(FileName)
    us = InStr(1, base, "_")
    If us <= 1 Then Exit Function

    TypeToken = Left$(base, us - 1)
    ObjName = Mid$(base, us + 1)
    ParseTypeAndNameFromFileName = (Len(TypeToken) > 0 And Len(ObjName) > 0)
End Function

Private Function MapTypeTokenToAcType(ByVal TypeToken As String) As AcObjectType
    Select Case LCase$(TypeToken)
        Case "form":   MapTypeTokenToAcType = acForm
        Case "report": MapTypeTokenToAcType = acReport
        Case "macro":  MapTypeTokenToAcType = acMacro
        Case "module": MapTypeTokenToAcType = acModule
        Case "query":  MapTypeTokenToAcType = acQuery
        Case Else:     MapTypeTokenToAcType = acDefault
    End Select
End Function

Private Function DetectTypeFromTextFile(ByVal FilePath As String, ByRef ObjType As AcObjectType, ByRef ObjName As String) As Boolean
    Dim first1k As String
    first1k = Left$(ReadAllText(FilePath), 2048)

    If InStr(1, first1k, "Begin Form", vbTextCompare) > 0 Then
        ObjType = acForm
    ElseIf InStr(1, first1k, "Begin Report", vbTextCompare) > 0 Then
        ObjType = acReport
    ElseIf InStr(1, first1k, "Begin Macro", vbTextCompare) > 0 Then
        ObjType = acMacro
    ElseIf InStr(1, first1k, "Module = Begin", vbTextCompare) > 0 Or _
           InStr(1, first1k, "Option Compare", vbTextCompare) > 0 Then
        ObjType = acModule
    ElseIf InStr(1, first1k, "Operation =1  ' CreateQueryDef", vbTextCompare) > 0 Or _
           InStr(1, first1k, "dbText", vbTextCompare) > 0 Then
        ObjType = acQuery
    Else
        DetectTypeFromTextFile = False
        Exit Function
    End If

    ObjName = ExtractNameFromSaveAsText(first1k)
    DetectTypeFromTextFile = True
End Function

Private Function ExtractNameFromSaveAsText(ByVal Chunk As String) As String
    Dim arr() As String
    Dim line As Variant   ' MUST be Variant for For Each over array
    arr = Split(Chunk, vbCrLf)
    For Each line In arr
        If Left$(Trim$(CStr(line)), 5) = "Name =" Then
            ExtractNameFromSaveAsText = Trim$(Mid$(CStr(line), 6))
            Exit Function
        End If
    Next line
End Function

Private Function DeriveNameFromModuleFile(ByVal FileName As String) As String
    Dim token As String, nm As String
    If ParseTypeAndNameFromFileName(FileName, token, nm) Then
        If LCase$(token) = "module" Then
            DeriveNameFromModuleFile = nm
            Exit Function
        End If
    End If
    DeriveNameFromModuleFile = StripExtension(FileName)
End Function


