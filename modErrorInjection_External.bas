
Option Compare Database
Option Explicit

' ==========================================================
' Tailored Error Handler Injector for Access VBA
' - UI components (Forms/Reports): show MsgBox to the user
' - Standard/Class modules: optionally rethrow after logging
' - Central logger tags user/computer/app info
'
' CONFIGURABLE OPTIONS
' ==========================================================
Private Const LOG_TO_DEBUG As Boolean = False
Private Const LOG_TO_TABLE As Boolean = True          ' set True if you created tblErrorLog (DDL below)
Private Const RETHROW_IN_STD_MODULES As Boolean = True ' raise error after logging in non-UI modules
Private Const RETHROW_IN_UI_MODULES As Boolean = False ' usually False—don’t annoy end users

Const cMyModName = "Const cMyModName As String"

Const MyName As String = "ErrorInjection"


' Optional table schema if LOG_TO_TABLE = True
' CREATE TABLE tblErrorLog (
'   LogID AUTOINCREMENT PRIMARY KEY,
'   LoggedAt DATETIME,
'   AppName TEXT(128),
'   ModuleName TEXT(128),
'   ProcName TEXT(128),
'   ErrNumber LONG,
'   ErrDescription TEXT(255),
'   ErrSource TEXT(255),
'   UserName TEXT(128),
'   ComputerName TEXT(128)
' );
' ==========================================================


' Generic entry: run injector against *any* VBIDE.VBProject object you pass in.
' This lets you work with projects from other Office hosts or other Access instances.
Public Sub AddErrorHandlersToVBProject(vbproj As Object)
    On Error GoTo Bail

    Dim vbcomp As Object    ' VBIDE.VBComponent
    Dim cm As Object        ' VBIDE.CodeModule
    Dim insertedCount As Long
    Dim scannedProcs As Long

    ' Basic capability checks
    If vbproj Is Nothing Then Err.Raise 5, , "VBProject is Nothing."
    If vbproj.VBComponents Is Nothing Then Err.Raise 5, , "VBProject has no VBComponents collection."

    For Each vbcomp In vbproj.VBComponents
        If HasCode(vbcomp) Then
            Set cm = vbcomp.CodeModule
            ' Call your existing *safe* injector per component
            insertedCount = insertedCount + InjectIntoComponent_Tailored(vbcomp, cm, scannedProcs)
        End If
    Next vbcomp

    MsgBox "Finished. Procedures scanned: " & scannedProcs & vbCrLf & _
           "Procedures updated (handlers added): " & insertedCount, vbInformation, "Error Handler Injector"
    Exit Sub

Bail:
    MsgBox "AddErrorHandlersToVBProject failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub





Public Sub AddErrorHandlersToProject_Tailored()
    Dim vbproj As Object    ' VBIDE.VBProject
    Dim vbcomp As Object    ' VBIDE.VBComponent
    Dim cm As Object        ' VBIDE.CodeModule
    
    On Error GoTo Bail
    Set vbproj = Application.VBE.ActiveVBProject
    
    Dim insertedCount As Long
    Dim scannedProcs As Long
    
    For Each vbcomp In vbproj.VBComponents
      If HasCode(vbcomp) Then
        If vbcomp.Name <> MyName Then
          Set cm = vbcomp.CodeModule
          insertedCount = insertedCount + InjectIntoComponent_Tailored(vbcomp, cm, scannedProcs)
        End If
      End If
    Next vbcomp
    
    MsgBox "Finished. Procedures scanned: " & scannedProcs & vbCrLf & _
           "Procedures updated (handlers added): " & insertedCount, vbInformation, "Error Handler Injector"
    Exit Sub
Bail:
    MsgBox "AddErrorHandlersToProject_Tailored failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ------------------- Core scanning/injection -------------------

Private Function HasCode(vbcomp As Object) As Boolean
    On Error Resume Next
    HasCode = (vbcomp.CodeModule Is Nothing) = False And (vbcomp.CodeModule.CountOfLines > 0)
End Function

Private Function InjectIntoComponent_Tailored(vbcomp As Object, cm As CodeModule, ByRef scannedProcs As Long) As Long
    Dim totalLines As Long, line As Long
    Dim ProcName As String, procKind As Long
    Dim startLine As Long, bodyLine As Long, procLines As Long, endLine As Long
    Dim updated As Long
    Dim foundMyModName As Boolean
    
    totalLines = cm.CountOfLines
    line = 1
    
    
    Do While line <= cm.CountOfDeclarationLines
      If InStr(cm.Lines(line, 1), cMyModName) <> 0 Then
        foundMyModName = True
        Exit Do
      End If
      line = line + 1
    Loop
    
    If Not foundMyModName Then
      cm.InsertLines cm.CountOfDeclarationLines + 1, cMyModName & " = " & quot(cm.Name)
    End If
    
    line = cm.CountOfDeclarationLines + 1
    totalLines = cm.CountOfLines
    
    Do While line <= totalLines
    
    'Const cMyModName As String = "Form_frmMainDashboard"
    
        ProcName = ""
        On Error Resume Next
        ProcName = cm.ProcOfLine(line, procKind)
        On Error GoTo 0
        
        If Len(ProcName) > 0 Then
            scannedProcs = scannedProcs + 1
            startLine = cm.ProcStartLine(ProcName, procKind)
            'Dim cm1 As Module
            'cm1.ProcStartLine
            'startLine = GetProcFirstTextLine(cm, procName, procKind)
            bodyLine = cm.ProcBodyLine(ProcName, procKind)
            procLines = cm.ProcCountLines(ProcName, procKind)
            endLine = startLine + procLines - 1
            'endLine = startLine + procLines
            
            If Not ProcedureHasErrorHandling(cm, startLine, endLine) Then
                If AddHandlerToProcedure_Tailored(vbcomp, cm, ProcName, procKind, startLine, bodyLine, endLine) Then
                    updated = updated + 1
                    totalLines = cm.CountOfLines
                    line = endLine + 1
                    GoTo NextLoop
                End If
            End If
            line = endLine + 1
        Else
            line = line + 1
        End If
NextLoop:
    Loop
    
    InjectIntoComponent_Tailored = updated
End Function

'procKind should be: vbext_ProcKind
'error on type
Private Function GetProcFirstTextLine(cm As Object, ProcName As String, procKind As Variant) As Long
  Dim startLine As Long
  Dim lineTxt As String
  Dim bLoopMore: bLoopMore = False
  
  startLine = cm.ProcStartLine(ProcName, procKind)
  
  Do
    lineTxt = Trim(cm.Lines(startLine, 1))
    If Left(lineTxt, 1) = "'" Or Len(lineTxt) = 0 Then
      bLoopMore = True
      startLine = startLine + 1
    Else
      bLoopMore = False
    End If
  Loop Until Not bLoopMore

  GetProcFirstTextLine = startLine

End Function


Private Function ProcedureHasErrorHandling(cm As Object, bodyLine As Long, endLine As Long) As Boolean
    Dim i As Long, txt As String
    For i = bodyLine To endLine
        txt = Trim$(cm.Lines(i, 1))
        If Len(txt) > 0 Then
            If Left$(txt, 1) <> "'" Then
                If InStr(1, LCase$(txt), "on error", vbTextCompare) > 0 Then
                    ProcedureHasErrorHandling = True
                    Exit Function
                End If
                If LCase$(Left$(txt, 11)) = "errHandler:" Then
                    ProcedureHasErrorHandling = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function AddHandlerToProcedure_Tailored(vbcomp As Object, cm As Object, ProcName As String, _
                                       procKind As Long, startLine As Long, bodyLine As Long, _
                                       endLine As Long) As Boolean
    On Error GoTo FailFast
    
    Dim declText As String
    'declText = cm.Lines(startLine, 1)
    declText = cm.Lines(bodyLine, 1)
    
    Dim lineCount As Long
    
    Dim exitKeyword As String
    exitKeyword = GetExitKeywordFromDeclaration(declText)
    If Len(exitKeyword) = 0 Then exitKeyword = "Exit Sub"
    
    ' Determine if this is a UI component (Form/Report) by VBComponent.Type and name
    'Dim isUI As Boolean
    'isUI = IsUIComponent(vbComp)
    
    ' Determine behavior flags for injected handler
    Dim doRethrow As Boolean
    
'    If isUI Then
'        doRethrow = RETHROW_IN_UI_MODULES
'    Else
'        doRethrow = RETHROW_IN_STD_MODULES
'    End If
    
    If IsProcedureEvent(ProcName) Then
      doRethrow = False
    Else
      doRethrow = True
    End If
    
    ' 1) Insert "On Error GoTo EH_Handler" at first executable line with existing indentation
    'Dim firstBody As String,
    Dim indent As String
'    firstBody = cm.Lines(bodyLine, 1)
'    indent = Left$(firstBody, Len(firstBody) - Len(LTrim$(firstBody)))
'    'cm.InsertLines bodyLine, indent & "On Error GoTo EH_Handler"
    
    'firstBodyLine returns the procedure declaration, not the actual body
    indent = vbNullString
    
    Dim insertLine As Long
    lineCount = cm.CountOfLines
    insertLine = FindFirstExecutableLine(cm, bodyLine, endLine, ProcName)
    cm.InsertLines insertLine, "  On Error GoTo ErrHandler" & vbCrLf & _
            vbTab & "Dim MyProcName As String: MyProcName = """ & ProcName & """" & vbCrLf

    endLine = endLine + (cm.CountOfLines - lineCount) ' adjust for the number of lines added
    
    ' 2) Insert tailored Exit/Handler block before End Sub/Function/Property
    Dim endStmtLine As Long
    endStmtLine = FindProcedureEndStatementLine(cm, bodyLine, endLine)
    If endStmtLine = 0 Then endStmtLine = endLine
    
    Dim block As String
    block = BuildHandlerBlock(indent, exitKeyword, vbcomp.Name, ProcName, doRethrow)
    
    cm.InsertLines endStmtLine, block
    
    AddHandlerToProcedure_Tailored = True
    Exit Function
    
FailFast:
    Debug.Print "AddHandlerToProcedure_Tailored failed in " & vbcomp.Name & "." & ProcName & ": " & Err.Number & " - " & Err.Description
    AddHandlerToProcedure_Tailored = False
End Function

'Private Function BuildHandlerBlock(ByVal indent As String, ByVal exitKeyword As String, _
'                                   ByVal ModuleName As String, ByVal ProcName As String, _
'                                   ByVal isUI As Boolean, ByVal doRethrow As Boolean) As String
Private Function BuildHandlerBlock(ByVal indent As String, ByVal exitKeyword As String, _
                                   ByVal ModuleName As String, ByVal ProcName As String, _
                                   ByVal doRethrow As Boolean) As String
    Dim s As String
    s = vbCrLf & _
        "ExitHere:" & vbCrLf & _
        vbTab & "On Error Resume Next " & vbCrLf & _
        vbTab & "'Cleanup Objects '" & vbCrLf & _
        vbTab & exitKeyword & vbCrLf & _
        vbCrLf & _
        "errHandler:" & vbCrLf & _
        vbTab & "MyApp.ErrorHandler.HandleError cMyModName, MyProcName, " & CStr(doRethrow) & " , " & CStr(Not doRethrow) & " " & vbCrLf & _
        vbTab & "Resume exitHere "
        'indent & "    HandleVBAError Err, """ & ModuleName & """, """ & ProcName & """, " & CStr(isUI) & vbCrLf
    
    BuildHandlerBlock = s
End Function

Private Function GetExitKeywordFromDeclaration(declText As String) As String
    Dim l As String: l = LCase$(declText)
    If InStr(1, l, " sub ", vbTextCompare) > 0 Or Right$(Trim$(l), 3) = "sub" Then
        GetExitKeywordFromDeclaration = "Exit Sub"
    ElseIf InStr(1, l, " function ", vbTextCompare) > 0 Or Right$(Trim$(l), 8) = "function" Then
        GetExitKeywordFromDeclaration = "Exit Function"
    ElseIf InStr(1, l, " property ", vbTextCompare) > 0 Then
        GetExitKeywordFromDeclaration = "Exit Property"
    Else
        GetExitKeywordFromDeclaration = ""
    End If
End Function

Private Function FindProcedureEndStatementLine(cm As Object, bodyLine As Long, endLine As Long) As Long
    Dim i As Long, t As String, lt As String
    For i = endLine To bodyLine Step -1
        t = cm.Lines(i, 1)
        lt = LCase$(Trim$(t))
        If lt = "end sub" Or lt = "end function" Or lt Like "end property*" Then
            FindProcedureEndStatementLine = i
            Exit Function
        End If
    Next i
    FindProcedureEndStatementLine = 0
End Function

Private Function IsUIComponent(vbcomp As Object) As Boolean
    ' Access forms/reports show up as Type=100 (vbext_ct_Document)
    ' We’ll also check name prefix "Form_" or "Report_"
    On Error Resume Next
    Dim t As Long: t = vbcomp.Type
    Dim n As String: n = vbcomp.Name
    IsUIComponent = (t = 100) And (Left$(n, 5) = "Form_" Or Left$(n, 7) = "Report_")
End Function

' ------------------- Centralized logging -------------------

Public Sub HandleVBAError(ByVal e As ErrObject, ByVal ModuleName As String, ByVal ProcName As String, _
                          Optional ByVal IsUIContext As Boolean = False)
    On Error Resume Next
    
    Dim appName As String
    Dim userName As String
    Dim computerName As String
    
    appName = CurrentProject.Name
    userName = Environ$("USERNAME")
    computerName = Environ$("COMPUTERNAME")
    
    If LOG_TO_DEBUG Then
        Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); _
            " | App=" & appName & _
            " | Module=" & ModuleName & _
            " | Proc=" & ProcName & _
            " | Err#=" & e.Number & _
            " | Desc=" & e.Description & _
            " | Src=" & e.Source & _
            " | User=" & userName & _
            " | PC=" & computerName & _
            " | UI=" & CStr(IsUIContext)
    End If
    
    If LOG_TO_TABLE Then
        Dim sql As String
        sql = "INSERT INTO tblErrorLog(LoggedAt, AppName, ModuleName, ProcName, ErrNumber, ErrDescription, ErrSource, UserName, ComputerName) VALUES(" & _
              "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#," & _
              "'" & Replace(appName, "'", "''") & "'," & _
              "'" & Replace(ModuleName, "'", "''") & "'," & _
              "'" & Replace(ProcName, "'", "''") & "'," & _
              e.Number & "," & _
              "'" & Replace(Left$(e.Description, 255), "'", "''") & "'," & _
              "'" & Replace(Left$(e.Source, 255), "'", "''") & "'," & _
              "'" & Replace(userName, "'", "''") & "'," & _
              "'" & Replace(computerName, "'", "''") & "'" & _
              ")"
        CurrentDb.Execute sql, dbFailOnError
    End If
End Sub


Private Function FindFirstExecutableLine(cm As Object, bodyLine As Long, endLine As Long, ProcName As String) As Long
    ' Scan between start and end lines until the first true executable statement
    ' Skips: line continuations, declarations (Dim, Const, Static, Private, Public, Friend), attributes, comments, blank lines
    
    Dim i As Long
    Dim txt As String, ltxt As String
    Dim priorLineCont As Boolean

    For i = bodyLine + 1 To endLine
        txt = Trim$(cm.Lines(i, 1))
        ltxt = LCase$(txt)
        
        If Left$(LTrim(txt), Len(ProcName)) = ProcName Then
          'proc assigned value on first executable line, use this line
          FindFirstExecutableLine = i
          Exit Function
        End If
        
        
        If Right$(txt, 1) = "_" Then
          priorLineCont = True
          GoTo NextLine
        End If
        
        If priorLineCont Then
          priorLineCont = False
          GoTo NextLine
        End If
        
        If Len(txt) = 0 Then GoTo NextLine
        If Left$(txt, 1) = "'" Then GoTo NextLine
        
        ' Declaration keywords
        If ltxt Like "dim *" Then GoTo NextLine
        If ltxt Like "const *" Then GoTo NextLine
        If ltxt Like "static *" Then GoTo NextLine
        If ltxt Like "private *" Then GoTo NextLine
        If ltxt Like "public *" Then GoTo NextLine
        If ltxt Like "friend *" Then GoTo NextLine
        If ltxt Like "attribute *" Then GoTo NextLine
        
        ' If we reach here, it's executable
        FindFirstExecutableLine = i
        Exit Function
        
NextLine:
    Next i
    
    ' Fallback: just after declaration
    FindFirstExecutableLine = bodyLine + 1
End Function

Public Function GetProcedureEventSuffix(ByVal ProcName As String) As String
    ' List of known event suffixes
    Dim eventNames As Variant
    Dim evt As Variant
    
    eventNames = Array( _
        "Click", "DblClick", "MouseDown", "MouseUp", "MouseMove", _
        "KeyDown", "KeyUp", "KeyPress", _
        "Enter", "Exit", "GotFocus", "LostFocus", _
        "BeforeUpdate", "AfterUpdate", "Change", _
        "NotInList", "DropDown", _
        "Current", "Dirty", "Undo", "BeforeInsert", "AfterInsert", _
        "BeforeDelConfirm", "AfterDelConfirm", "Error", _
        "AttachmentCurrent", "Timer", "Resize", "Open", "Load", "Unload", "Close" _
    )
    
    ' Default return
    GetProcedureEventSuffix = vbNullString
    
    ' Loop through events to check suffix match
    For Each evt In eventNames
        If Right$(ProcName, Len(evt) + 1) = "_" & evt Then
            GetProcedureEventSuffix = evt
            Exit For
        End If
    Next evt
End Function



Public Function IsProcedureEvent(ByVal ProcName As String) As Boolean
    ' List of known event suffixes
    Dim eventNames As Variant
    Dim bEvent As Boolean: bEvent = False
    Dim evt As Variant
    
    eventNames = Array( _
        "Click", "DblClick", "MouseDown", "MouseUp", "MouseMove", _
        "KeyDown", "KeyUp", "KeyPress", _
        "Enter", "Exit", "GotFocus", "LostFocus", _
        "BeforeUpdate", "AfterUpdate", "Change", _
        "NotInList", "DropDown", _
        "Current", "Dirty", "Undo", "BeforeInsert", "AfterInsert", _
        "BeforeDelConfirm", "AfterDelConfirm", "Error", _
        "AttachmentCurrent", "Timer", "Resize", "Open", "Load", "Unload", "Close" _
    )
        
    ' Loop through events to check suffix match
    For Each evt In eventNames
        If Right$(ProcName, Len(evt) + 1) = "_" & evt Then
            bEvent = True
            Exit For
        End If
    Next evt
    
    IsProcedureEvent = bEvent
    
End Function






