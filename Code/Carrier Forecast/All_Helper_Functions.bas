Attribute VB_Name = "All_Helper_Functions"
Option Explicit
'Pauses for x# of milliseconds
'Used for email function to prevent
'All emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Proc  : Function Email
' Date  : 10/11/2012
' Type  : Variant
' Desc  : Sends an email
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body"
'
' TODO  : change attachment to a string array
' TODO  : and loop through for each string it contains
' TODO  : to support multiple attachments.
'
' TODO  : check to make sure files exist before
' TODO  : adding them as attachments
'
' TODO  : add bool to function for delete attached
' TODO  : files after email is sent
'---------------------------------------------------------------------------------------
Function Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As String)
    Dim Mail_Object, Mail_Single As Variant
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    With Mail_Single
        .Subject = Subject
        If Attachment <> "" Then
            'Attachment must contain file path
            .Attachments.Add Attachment
        End If
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
        .Send
    End With
    Sleep 1500
End Function

Sub CleanUp()
    Dim aWorksheets As Variant

    For Each aWorksheets In ThisWorkbook.Sheets
        If aWorksheets.Name <> "Master" And aWorksheets.Name <> "Macro" Then
            RemoveFilter (aWorksheets.Name)
            aWorksheets.Cells.Delete
        End If
    Next
End Sub

Sub Delete(sPath As String)
    On Error Resume Next
    If FileExists(sPath) = True Then
        Kill sPath
    End If
    On Error GoTo 0
End Sub

Sub RemoveFilter(Optional sheet As String)
    On Error Resume Next
    If sheet = "" Then    'If no arg is given clear the activesheet
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilterMode = False
        End If
    Else        'else clear the specified sheet
        If ActiveWorkbook.Worksheets(sheet).AutoFilterMode = True Then
            ActiveWorkbook.Worksheets(sheet).AutoFilterMode = False
        End If
    End If
    On Error GoTo 0
End Sub

Function FileExists(ByVal sLoc As String) As Boolean
    If InStr(Len(sLoc), sLoc, "\") > 0 Then
        Do While (InStr(Len(sLoc), sLoc, "\") > 0)
            sLoc = Left(sLoc, Len(sLoc) - 1)
        Loop
    End If

    If Dir(sLoc, vbDirectory) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub RecMkDir(ByVal sPath)
    Dim sTemp As String
    Dim looplimit As Integer
    Dim i As Integer

    sTemp = sPath

    If InStr(Len(sTemp), sTemp, "\") > 0 Then
        Do While (InStr(Len(sTemp), sTemp, "\") > 0)
            sTemp = Left(sTemp, Len(sTemp) - 1)
        Loop
    End If

    On Error GoTo RECURSIVE:
    Do While FileExists(sPath) <> True
        MkDir sTemp
        sTemp = sPath
    Loop
    On Error GoTo 0
    Exit Sub

RECURSIVE:
    If FileExists(sTemp) = False Then
        For i = Len(sTemp) To 1 Step -1
            If Mid(sTemp, i, 1) = "\" Then
                looplimit = looplimit + 1
                sTemp = Left(sTemp, i - 1)
                Exit For
            End If
        Next i
    End If
    If looplimit = 20 Then Err.Raise 76
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim File As String
    Dim WkbkPath As String


    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)
    codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    If Dir(codeFolder) = "" Then
        RecMkDir codeFolder
    End If
    On Error GoTo 0

    'Remove all previously exported modules
    File = Dir(codeFolder)
    Do While File <> ""
        DeleteFile codeFolder & File
        File = Dir
    Loop

    'Export modules in current project
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = codeFolder & comp.Name & ".bas"
                comp.Export FileName
            Case 2
                FileName = codeFolder & comp.Name & ".cls"
                comp.Export FileName
            Case 3
                FileName = codeFolder & comp.Name & ".frm"
                comp.Export FileName
            Case 100
                If comp.Name = "ThisWorkbook" Then
                    FileName = codeFolder & comp.Name & ".bas"
                    comp.Export FileName
                End If
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String)
    On Error Resume Next
    Kill FileName
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : CombinePaths
' Date : 3/19/2013
' Desc : Adds folders onto the end of a file path
'---------------------------------------------------------------------------------------
Function CombinePaths(ByVal Path1 As String, ByVal Path2 As String) As String
    If Not EndsWith(Path1, "\") Then
        Path1 = Path1 & "\"
    End If
    CombinePaths = Path1 & Path2
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Private Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Result = True
            Exit For
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds references required for helper functions
'---------------------------------------------------------------------------------------
Sub AddReferences()
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean

    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = "{0002E157-0000-0000-C000-000000000046}" And Ref.Major = 5 And Ref.Minor = 3 Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes references required for helper functions
'---------------------------------------------------------------------------------------
Sub RemoveReferences()
    Dim Ref As Variant

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = "{0002E157-0000-0000-C000-000000000046}" And Ref.Major = 5 And Ref.Minor = 3 Then
            Application.VBE.ActiveVBProject.References.Remove Ref
        End If
    Next
End Sub
