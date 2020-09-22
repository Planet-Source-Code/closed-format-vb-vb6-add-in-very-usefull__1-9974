VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13860
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14190
   _ExtentX        =   25030
   _ExtentY        =   24448
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "My Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Sub Hide()
    
End Sub

Sub Show()
    Dim CurrTab As Integer
    Dim CurrCommand As String
    Dim CurrLine As String
    Dim NextLine As String
    Dim KydStart As Integer
    Dim TabAfter As Boolean
    Dim TabSpace As String
    Dim Count2 As Integer
    Dim LastKeyWord As String
    Dim Count As Long
    'On Error GoTo bad
    For Count = 1 To VBInstance.ActiveCodePane.CodeModule.CountOfLines - 1
foundone:
        CurrLine = VBInstance.ActiveCodePane.CodeModule.Lines(Count, 1)
        NextLine = VBInstance.ActiveCodePane.CodeModule.Lines(Count + 1, 1)
        If Trim(CurrLine) = "" And Trim(NextLine) = "" And Count <= VBInstance.ActiveCodePane.CodeModule.CountOfLines Then
            VBInstance.ActiveCodePane.CodeModule.DeleteLines Count
            GoTo foundone
        End If
        If Right(CurrLine, 1) = "_" Then
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine Count, Left(CurrLine, Len(CurrLine) - 1) & NextLine
            VBInstance.ActiveCodePane.CodeModule.DeleteLines Count + 1
            GoTo foundone
        End If
    Next
    For Count = 1 To VBInstance.ActiveCodePane.CodeModule.CountOfLines
        CurrLine = LCase(Trim(VBInstance.ActiveCodePane.CodeModule.Lines(Count, 1)))
        KydStart = InStr(1, CurrLine, " ")
        If KydStart = 0 Then KydStart = Len(CurrLine)
        CurrCommand = Trim(Mid(CurrLine, 1, KydStart))
        Select Case CurrCommand
            Case Is = "public", "private"
                If Mid(Trim(Mid(CurrLine, Len(CurrCommand) + 1)), 1, Len("sub")) = "sub" Or Mid(Trim(Mid(CurrLine, Len(CurrCommand) + 1)), 1, Len("function")) = "function" Then
                    CurrTab = CurrTab + 1
                    TabAfter = True
                End If
            Case "if"
                If Len(CurrLine) > 6 Then
                    If Mid(CurrLine, Len(CurrLine) - 3, 4) <> "then" Then
                    Else
                        CurrTab = CurrTab + 1
                        TabAfter = True
                    End If
                End If
            Case Is = "while", "do", "select", "for", "sub", "function"
                CurrTab = CurrTab + 1
                TabAfter = True
            Case Is = "end", "wend", "loop", "next"
                If Mid(CurrLine, 1, Len("end if")) = "end if" Or Mid(CurrLine, 1, Len("next")) = "next" Then
                    CurrTab = CurrTab - 1
                    TabAfter = False
                End If
                If Mid(CurrLine, 1, Len("end select")) = "end select" Then
                    CurrTab = CurrTab - 2
                    TabAfter = False
                End If
                If Mid(CurrLine, 1, Len("end function")) = "end function" Or Mid(CurrLine, 1, Len("end sub")) = "end sub" Then
                    CurrTab = 0
                    TabAfter = False
                End If
            Case Is = "else", "case", "elseif"
                If LastKeyWord <> "select" Then
                    CurrTab = CurrTab - 1
                    TabSpace = ""
                    If Len(TabSpace) / 4 <> CurrTab Then
                        For Count2 = 1 To CurrTab
                            TabSpace = TabSpace & vbTab
                        Next
                    End If
                End If
                CurrTab = CurrTab + 1
                TabAfter = True
        End Select
        If TabAfter = True Then
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine Count, TabSpace & Trim(VBInstance.ActiveCodePane.CodeModule.Lines(Count, 1))
            TabSpace = ""
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                Next
            End If
        Else
            TabSpace = ""
            If Len(TabSpace) / 4 <> CurrTab Then
                For Count2 = 1 To CurrTab
                    TabSpace = TabSpace & vbTab
                Next
            End If
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine Count, TabSpace & Trim(VBInstance.ActiveCodePane.CodeModule.Lines(Count, 1))
        End If
        If Len(CurrCommand) > 1 Then
            If Mid(CurrCommand, Len(CurrCommand) - 1, 1) = ":" Then
                VBInstance.ActiveCodePane.CodeModule.ReplaceLine Count, Trim(VBInstance.ActiveCodePane.CodeModule.Lines(Count, 1))
            End If
        End If
        LastKeyWord = CurrCommand
    Next
bad:
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Format")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
    
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
    
End Function
