Attribute VB_Name = "SAPLogon"
Option Explicit

' ### Version       : 2.09
' ### Started       : 07.06.2020
' ### Last Update   : 08.06.2023
' ### Add Referance : C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapfewse.ocx
' ### Dim CollCon As SAPFEWSELib.GuiComponentCollection

Private WshShell As Object

Private SAPApp As SAPFEWSELib.GuiApplication
Private SAPGui As Object
Private connection As SAPFEWSELib.GuiConnection
Public session As SAPFEWSELib.GuiSession

Private SAPLogonPID As Long
Private SAPConnections As ISapCollectionTarget
Private SAPSessions As Object


Private iSapClassicPath As String
Private iSapBCPath As String
Private iSapSCPath As String
Private iSapSystemName As String
Private iSapSystemDescName As String
Private iSapUserLanguage As String
Private iSapUsername As String
Private iSapUserPassword As String
Private iSelectedSapClient As SapClient
Private iConnectedClient As SapClient
Private FirstConnection As Boolean


Public Enum SapConnectionStatus
  Success = 1
  ScriptingError = 2
  AuthError = 3
  ShellStartFailed = 4
End Enum

Public Enum DisconnectType
  ForceClose = 1
  GotoMain = 2
  CloseWindow = 3
  OnlyDisconnect = 4
End Enum

Public Enum SapClient
  Undefined = 1
  Sapshcut = 2
  SapClassic = 3
  SapBC = 4
  [_First] = 3
  [_Last] = 4
End Enum


Private Sub SAPMemoryFlush()
  Set SAPGui = Nothing
  Set SAPApp = Nothing
  Set connection = Nothing
  Set session = Nothing
  Set WshShell = Nothing
  Set SAPConnections = Nothing
  Set SAPSessions = Nothing
  SAPLogonPID = vbNull
  ConnectedClient = Undefined
  FirstConnection = False
End Sub

Public Function SapConnection() As SapConnectionStatus
  
  Dim c As SAPFEWSELib.GuiConnection
  Dim s As SAPFEWSELib.GuiSession
  Dim sc As SapClient
  Dim Status As SapConnectionStatus
  
  
  Call SAPMemoryFlush
  
  'On Error GoTo UnhandledError
  
SAPLogonStartAfter: 'SAP acildiktan sonra geri gel
  
  For sc = SapClient.[_First] To SapClient.[_Last]
    Status = InitSapEngine(sc)
    
    If Status = ScriptingError Then
      SapConnection = ScriptingError
      Exit Function
    ElseIf Status = Success Then
      ConnectedClient = sc
      
      Set SAPConnections = SAPApp.Connections
      
      If SAPConnections.Count > 0 Then
      
        For Each c In SAPConnections
          Set connection = c
          
          Set s = c.Children(0)
          s.CreateSession
          
          Application.Wait Now + TimeValue("0:00:02")
          Set s = Nothing
          
          Set SAPSessions = c.Sessions
          Set session = SAPSessions(c.Sessions.Count - 1)
          
          GoTo ConnectionSuccess
          
        Next
      
      Else
        Exit For
      End If
      
    End If
  Next

  If ConnectedClient = Undefined And connection Is Nothing Then
    GoTo SAPLogonStartAttemp
  ElseIf ConnectedClient <> Undefined And connection Is Nothing Then
    FirstConnection = True
    GoTo CreateFirstConnection
  End If
  
CreateFirstConnection:
  'Mevcut oturum yok yeni oturum yarat
  Set connection = SAPApp.OpenConnection(SapSystemDescName, True)
  Set session = connection.Children(0)
  
  session.LockSessionUI
  
ConnectionSuccess:

  SapConnection = SapConnectionStatus.Success
  WriteLog ("SAP connected.")
  On Error GoTo 0
  Exit Function
  
SAPLogonStartAttemp:
  Call SAPLogonStart
  GoTo SAPLogonStartAfter
  Exit Function
UnhandledError:
  SapConnection = SapConnectionStatus.AuthError
  WriteLog ("SAP connection error. " & Err.Description)
  On Error GoTo 0
End Function

Private Sub SAPLogonStart()

  WriteLog ("SAP shell starting.")
  
  Dim Counter As Byte
  
  SAPLogonPID = Shell(GetSapPath(SelectedSapClient), vbNormalFocus)
  Set WshShell = CreateObject("WScript.Shell")
  
  Do Until WshShell.AppActivate("SAP Logon ", True) Or WshShell.AppActivate("SAP Easy Access", True)
    If Counter > 5 Then
      GoTo StartError
    Else
      Counter = Counter + 1
      Application.Wait Now + TimeValue("0:00:01")
    End If
  Loop
  
  Set WshShell = Nothing
  
  WriteLog ("SAP shell started.")
  
  Exit Sub
  
StartError:
  WriteLog ("SAP shell starting failed")
End Sub

Public Sub Disconnect(ByVal DisType As DisconnectType)
  
  On Error Resume Next
  
  If DisType = ForceClose Then
  
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
    session.findById("wnd[0]").sendVKey 0
    Call tncCore.KillProcess("saplogon.exe")
    Call tncCore.KillProcess("NWBC.exe")
    
  ElseIf DisType = CloseWindow Then
  
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    
  ElseIf DisType = GotoMain Then
  
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
    session.findById("wnd[0]").sendVKey 0
    
    If SAPLogon.ConnectedClient = SapClassic Then
      session.findById("wnd[0]").iconify
    End If
    
  ElseIf DisType = OnlyDisconnect Then
    'asagidaki kod bu ise yariyor
  End If
  
  session.UnlockSessionUI
  
  Call SAPMemoryFlush
  
  On Error GoTo 0
  
  WriteLog "SAP disconnected."

End Sub

Private Function InitSapEngine(ByVal Client As SapClient) As SapConnectionStatus

  InitSapEngine = Success
  
  On Error GoTo OE
  
  If Client = SapClassic Then
    Set SAPGui = GetObject("SAPGUI")
  Else
    Set SAPGui = GetObject("SAPGUISERVER")
  End If
  
  On Error GoTo SE
  Set SAPApp = SAPGui.GetScriptingEngine
  On Error GoTo 0
  
  Exit Function
  
OE:
  InitSapEngine = AuthError
  Exit Function
SE:
  InitSapEngine = ScriptingError
End Function

Property Get SapSystemName() As String
  SapSystemName = iSapSystemName
End Property

Property Let SapSystemName(ByVal Data As String)
  iSapSystemName = Data
End Property

Property Get SapSystemDescName() As String
  SapSystemDescName = iSapSystemDescName
End Property

Property Let SapSystemDescName(ByVal Data As String)
  iSapSystemDescName = Data
End Property

Property Get SapUserLanguage() As String
  SapUserLanguage = iSapUserLanguage
End Property

Property Let SapUserLanguage(ByVal Data As String)
  iSapUserLanguage = Data
End Property

Property Get SapCSPath() As String
  SapCSPath = iSapClassicPath
End Property

Property Let SapCSPath(ByVal Data As String)
  iSapClassicPath = Data
End Property

Property Get SapBCPath() As String
  SapBCPath = iSapBCPath
End Property

Property Let SapBCPath(ByVal Data As String)
  iSapBCPath = Data
End Property

Property Get SapSCPath() As String
  SapSCPath = iSapSCPath
End Property

Property Let SapSCPath(ByVal Data As String)
  iSapSCPath = Data
End Property

Property Get SapUsername() As String
  SapUsername = iSapUsername
End Property

Property Let SapUsername(ByVal Data As String)
  iSapUsername = Data
End Property

Property Get SapUserPassword() As String
  SapUserPassword = iSapUserPassword
End Property

Property Let SapUserPassword(ByVal Data As String)
  iSapUserPassword = Data
End Property

Property Get ConnectedClient() As SapClient
  ConnectedClient = iConnectedClient
End Property

Property Let ConnectedClient(ByVal Client As SapClient)
  iConnectedClient = Client
End Property

Property Get SelectedSapClient() As SapClient
  SelectedSapClient = iSelectedSapClient
End Property

Property Let SelectedSapClient(ByVal Client As SapClient)
  iSelectedSapClient = Client
End Property

Public Function GetSapPath(ByVal ChooseClient As SapClient) As String
  Select Case ChooseClient
    Case SapClassic:
      GetSapPath = SAPLogon.SapCSPath
    Case SapBC
      GetSapPath = SAPLogon.SapBCPath
    Case Sapshcut:
      GetSapPath = SAPLogon.SapSCPath
    Case Else:
      GetSapPath = SAPLogon.SapCSPath
  End Select
End Function

Public Function GetLiveSapClient(ByVal connection As Object) As SapClient
  If connection.Description <> "" Then
    GetLiveSapClient = SapClassic
  Else
    GetLiveSapClient = SapBC
  End If
End Function

Property Get GetLiveSapLanguage()
  GetLiveSapLanguage = session.Info.language
End Property
