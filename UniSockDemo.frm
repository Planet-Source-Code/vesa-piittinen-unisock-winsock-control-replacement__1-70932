VERSION 5.00
Begin VB.Form UniSockDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniSock sample"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "UniSockDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "UniSockDemo.frx":000C
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect client to server"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "UniSockDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Client As UniSock
Attribute Client.VB_VarHelpID = -1
Private WithEvents Server As UniSock
Attribute Server.VB_VarHelpID = -1
Private WithEvents ServerNode As UniSock
Attribute ServerNode.VB_VarHelpID = -1

Private Sub Client_Closing()
    Debug.Print "Client", "Closing"
    ' note: you don't really need to explicitly call CloseSocket, this event only runs before a CloseSocket
    Client.CloseSocket
    ' if I recall correctly, this was one of the biggest nuisances with the Winsock control,
    ' it just sometimes didn't close itself when it was expected to happen
    Command1.Enabled = True
    Command1.SetFocus
End Sub

Private Sub Client_Connect()
    Debug.Print "Client", "Remote IP: " & Client.RemoteHostIP
    ' send some sample data to the server
    Client.SendData Time$ & " This text is sent to server by client."
End Sub

Private Sub Client_DataArrival(ByVal BytesTotal As Long)
    Dim strData As String
    ' show data in the immediate window
    Client.GetData strData
    ' UTF-8 is shown incorrectly (comment Client.Mode in Form_Load to see)
    Text1.Text = strData
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' DIFFERENCE TO WINSOCK CONTROL: errors do NOT stop things on track!
    Debug.Print "Client", Source, Description
End Sub

Private Sub Client_TextArrival(Text As String, ByVal LineChange As UniSockLineChange, ByVal ANSI As Boolean)
    ' UTF-8 is read correctly! (uncomment Client.Mode in Form_Load to see)
    Text1.Text = Text
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    ' just let client connect to server
    Set ServerNode = Nothing
    Cls
    CurrentX = 150
    CurrentY = 150
    If Not Client.Connect("localhost", 5000) Then Debug.Print "Can't connect"
End Sub

Private Sub Form_Load()
    Set Server = New UniSock
    ' bind the server to localhost port 5000 and start listening
    If Not Server.Bind(5000, "localhost") Then Debug.Print "Server", "Bind failed"
    If Not Server.Listen Then Debug.Print "Server", "Listen failed"
    Set Client = New UniSock
    
    ' for some reason text mode is slightly more unstable than binary mode
    'Client.Mode = [Socket Text Mode]
    
    ' make the initial call to command button
    Command1_Click
    ' also notice the text is drawn after the form has appeared, this means asynchronous operation has happened
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' cleaning up is a good thing to do
    Set Client = Nothing
    Set Server = Nothing
    Set ServerNode = Nothing
End Sub

Private Sub Server_ConnectionRequest(ByVal RequestID As Long)
    Debug.Print "Server", "Creating ServerNode"
    Set ServerNode = New UniSock
    ' IMPORTANT: RequestID sent by UniSock is not a socket handle! It does not work with other Winsock implementations.
    ServerNode.Accept RequestID
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' DIFFERENCE TO WINSOCK CONTROL: errors do NOT stop things on track!
    Debug.Print "Server", Source, Description
End Sub

Private Sub ServerNode_Closing()
    ' this event never triggers thanks to manual use of CloseSocket
    Debug.Print "ServerNode", "Closing"
End Sub

Private Sub ServerNode_Connect()
    ' this event triggers when accepting a connection (does not happen with other Winsock implementations?)
    Debug.Print "ServerNode", "Remote IP: " & ServerNode.RemoteHostIP, "Local IP: " & ServerNode.LocalIP
End Sub

Private Sub ServerNode_DataArrival(ByVal BytesTotal As Long)
    Dim strText As String
    ' now print the text to the form
    ServerNode.GetData strText
    Print strText
    ' also send back something to the client: INTERACTION! Didn't see that coming?
    ' note: text is sent as UTF-8!
    ServerNode.SendText "ServerNode sent this to client and quited ÅÄÖ"
    ServerNode.CloseSocket
End Sub

Private Sub ServerNode_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' DIFFERENCE TO WINSOCK CONTROL: errors do NOT stop things on track!
    Debug.Print "ServerNode", Source, Description
End Sub
