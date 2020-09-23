VERSION 5.00
Begin VB.Form frmVBPNewHandler 
   Caption         =   "New Handler for Visual Basic Files"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   3435
      TabIndex        =   0
      Top             =   3795
      Width           =   2700
   End
End
Attribute VB_Name = "frmVBPNewHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub Form_Load()
On Error GoTo E
Dim str As String, sp As String, cmd As String, c As Long, I As Long
Dim cmdPath As String, NewName As String, cmdFileName As String
Dim cmdExt As String, S As String, M As String
    S = Chr$(34)
str = GetStringValue("HKEY_CLASSES_ROOT\.vbp", "")
str = GetStringValue("HKEY_CLASSES_ROOT\" & str & "\Shell\open\Command", "")
        sp = App.Path
            If Right$(sp, 1) <> "\" Then sp = sp & "\"
    cmd = Trim$(Replace$(Command$, Chr(34), ""))
CreateKey "HKEY_CLASSES_ROOT\.vbp\ShellNew"
'SetStringValue "HKEY_CLASSES_ROOT\.vbp\ShellNew", "Command",sp & App.EXEName & ".exe "  & Chr(34) & "%1" & Chr(34) & " /New"
    sp = sp & App.EXEName & ".exe "
Dim buf As String
            buf = Space$(255)
        GetShortPathName sp, buf, 255
            buf = Trim$(TrimNull(buf))
                sp = IIf(buf = "", sp, buf)
SetStringValue "HKEY_CLASSES_ROOT\.vbp\ShellNew", "Command", sp & " %1" & " /New"
SetStringValue "HKEY_CLASSES_ROOT\.vbp\ShellNew", "NullFile", ""

If cmd <> "" Then
        c = InStr(LCase$(cmd), "/new")
    If c Then
        cmd = Left$(cmd, c - 1)
    End If
            c = InStrRev(cmd, "\")
            I = InStrRev(cmd, ".")
        If c Then
            cmdPath = Left$(cmd, c)
            cmdExt = Mid$(cmd, I)
            cmdFileName = Mid$(cmd, c + 1, I - (c + 1))
        End If
       NewName = InputBox("Enter New Filename", , cmdFileName)
            If NewName = "" Then End 'NewName = cmdFileName
                If InStr(NewName, "(") Then
                    NewName = Replace$(NewName, "(", "")
                End If
                If InStr(NewName, ")") Then
                    NewName = Replace$(NewName, ")", "")
                End If
                
        cmd = cmdPath & NewName & cmdExt
        
            Open cmdPath & "frm" & NewName & ".frm" For Output As 1
                    M = Space$(3)
                Print #1, "VERSION 5.00"
                Print #1, "Begin VB.Form " & Replace$("frm" & NewName, " ", "")
                Print #1, M & "Caption = " & S & NewName & S
                Print #1, M & "ClientHeight = 3195"
                Print #1, M & "ClientLeft = 60"
                Print #1, M & "ClientTop = 345"
                Print #1, M & "ClientWidth = 4680"
                Print #1, M & "LinkTopic = " & S & NewName & S
                Print #1, M & "ScaleHeight = 3195"
                Print #1, M & "ScaleWidth = 4680"
                Print #1, M & "StartUpPosition = 3    'Windows Default"
                Print #1, "End"
'                Print #1, "Attribute VB_Name = " & S & NewName & S
                Print #1, "Attribute VB_Name = " & S & Replace$("frm" & NewName, " ", "") & S
                Print #1, "Attribute VB_GlobalNameSpace = False"
                Print #1, "Attribute VB_Creatable = False"
                Print #1, "Attribute VB_PredeclaredId = True"
                Print #1, "Attribute VB_Exposed = False"
                Print #1, "Option Explicit"
            Close 1
                Open cmdPath & "mod" & NewName & ".bas" For Output As 1
                    Print #1, "Attribute VB_Name = " & S & Replace$("mod" & NewName, " ", "") & S
                    Print #1, "Option Explicit"
                Close 1
        Open cmd For Output As 1
            Print #1, "Type=Exe"
            Print #1, "Form=" & "frm" & NewName & ".frm"
            Print #1, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\System32\STDOLE2.TLB#OLE Automation"
            Print #1, "Module=" & Replace$("mod" & NewName, " ", "") & "; " & "mod" & NewName & ".bas"
            Print #1, "IconForm=" & S & Replace$("frm" & NewName, " ", "") & S
            Print #1, "Startup=" & S & Replace$("frm" & NewName, " ", "") & S
            Print #1, "HelpFile=" & S & S
            Print #1, "ExeName32=" & S & NewName & ".exe" & S
            Print #1, "Title=" & S & NewName & S
            Print #1, "Command32=" & S & "" & S
            Print #1, "Name=" & S & NewName & S
            Print #1, "HelpContextID=" & S & "0" & S
            Print #1, "CompatibleMode=" & S & "0" & S
            Print #1, "MajorVer=1"
            Print #1, "MinorVer=0"
            Print #1, "RevisionVer=0"
            Print #1, "AutoIncrementVer=0"
            Print #1, "ServerSupportFiles=0"
            Print #1, "VersionCompanyName=" & S & "VBPHandler" & S ' App.Title & S
            Print #1, "CompilationType=-1"
            Print #1, "OptimizationType=0"
            Print #1, "FavorPentiumPro(tm)=0"
            Print #1, "CodeViewDebugInfo=0"
            Print #1, "NoAliasing=0"
            Print #1, "BoundsCheck=0"
            Print #1, "OverflowCheck=0"
            Print #1, "FlPointCheck=0"
            Print #1, "FDIVCheck=0"
            Print #1, "UnroundedFP=0"
            Print #1, "StartMode=0"
            Print #1, "Unattended=0"
            Print #1, "Retained=0"
            Print #1, "ThreadPerObject=0"
            Print #1, "MaxNumberOfThreads=1"
            Print #1, "DebugStartupOption=0"
            Print #1, ""
            Print #1, "[MS Transaction Server]"
            Print #1, "AutoRefresh=1"
        Close 1
            str = Replace$(str, "%1", cmd)
    If MsgBox("Do you want to Start VisualBasic?", vbQuestion + vbYesNo, "Files created") = vbYes Then
        Shell str, vbNormalFocus
    End If
End If
        End
Exit Sub
E:
MsgBox Err.Description, vbCritical
Close
End
End Sub
