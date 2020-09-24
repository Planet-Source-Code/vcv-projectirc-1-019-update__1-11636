VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm Client 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "projectIRC"
   ClientHeight    =   7980
   ClientLeft      =   4440
   ClientTop       =   1815
   ClientWidth     =   10815
   Icon            =   "frmClient_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picTask 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   3
      Top             =   7695
      Width           =   10815
      Begin VB.Shape shpTask 
         BorderColor     =   &H00808080&
         Height          =   270
         Left            =   0
         Top             =   15
         Width           =   1170
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3150
      Top             =   1605
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":307E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient_MDI.frx":5832
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picToolMain 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   0
      Top             =   7275
      Width           =   10815
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   420
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   741
         BandCount       =   1
         ForeColor       =   -2147483628
         FixedOrder      =   -1  'True
         _CBWidth        =   9345
         _CBHeight       =   420
         _Version        =   "6.0.8169"
         MinHeight1      =   24
         Width1          =   619
         UseCoolbarColors1=   0   'False
         NewRow1         =   0   'False
         BandStyle1      =   1
         AllowVertical1  =   0   'False
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   75
            TabIndex        =   2
            Top             =   45
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSWinsockLib.Winsock IDENT 
      Left            =   1155
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   113
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   1785
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Connect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_File_Disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Options 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Quit 
         Caption         =   "&Quit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_View_Status 
         Caption         =   "&Status Window"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Window_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_Window_TileH 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnu_Tile_Vertically 
         Caption         =   "&Tile Vertically"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clrSep = &H808080
Public intActive As Integer, sngLastX As Integer
Public intHover As Integer, bReDraw As Boolean, intLast As Integer

Const buttonShadow = &H80000010
Const buttonHilight = &H80000014
Sub CTCPReply(strNick As String, strReply As String)
    Client.SendData "NOTICE " & strNick & " :" & strAction & strReply & strAction
End Sub

Sub DrawToolbar()
    Dim intSeps As Integer, CenX As Integer, j As Integer, realWidth As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, intBegin As Integer
    Dim intEnd As Integer, strDrawText As String
    bReDraw = True
    realWidth = Client.ScaleWidth / 15
    intSeps = WindowCount - 1
    
    intWidth = realWidth / (intSeps + 1)
    picTask.Cls
    
    picTask.CurrentY = 2
    '* If one window open, draw only single thing..
    If intSeps = 0 Then
        strTitle = GetWindowTitle(1)
        strDrawText = TaskText(intWidth, strTitle)
        CenX = (realWidth - picTask.TextWidth(strDrawText)) / 2
        picTask.CurrentX = CenX
        picTask.ForeColor = vbRed
        picTask.Print strDrawText;
        Exit Sub
    End If
    
    picTask.ForeColor = clrSep
    For i = 1 To intSeps
        picTask.CurrentX = intWidth * i
        For j = 1 To picTask.ScaleHeight - 2 Step 2
            picTask.PSet (picTask.CurrentX, j)
        Next j
    Next i
    
    picTask.ForeColor = vbBlack
    For i = 1 To intSeps + 1
        picTask.CurrentY = 2
        picTask.CurrentX = intWidth * (i - 1)
        intBegin = picTask.CurrentX + 1
        intEnd = intBegin + intWidth - 2
        If intActive = i Then
            picTask.ForeColor = vbRed
            '* Active window, let's draw a inset bevel
            picTask.Line (intBegin, 2)-(intEnd, 2), buttonShadow
            picTask.Line (intBegin, 2)-(intBegin, picTask.ScaleHeight - 2), buttonShadow
            picTask.Line (intEnd, 2)-(intEnd, picTask.ScaleHeight - 2), buttonHilight
            picTask.Line (intBegin, picTask.ScaleHeight - 2)-(intEnd, picTask.ScaleHeight - 2), buttonHilight
            picTask.CurrentY = 3
            If intHover = i Then picTask.Tag = intHover
        ElseIf intHover = i Then
            picTask.Tag = intHover
            picTask.ForeColor = vbBlack
            picTask.Line (intBegin, 2)-(intEnd, 2), buttonHilight
            picTask.Line (intBegin, 2)-(intBegin, picTask.ScaleHeight - 2), buttonHilight
            picTask.Line (intEnd, 2)-(intEnd, picTask.ScaleHeight - 2), buttonShadow
            picTask.Line (intBegin, picTask.ScaleHeight - 2)-(intEnd, picTask.ScaleHeight - 2), buttonShadow
            picTask.CurrentY = 2
        Else
            If picTask.ForeColor <> vbBlack Then picTask.ForeColor = vbBlack
            picTask.CurrentY = 2
        End If
        
        strTitle = GetWindowTitle(Int(i))
        
        picTask.CurrentX = intWidth * (i - 1)
        intBegin = picTask.CurrentX
        intEnd = intBegin + intWidth
        strDrawText = TaskText(intWidth, strTitle)
        picTask.CurrentX = TaskCenter(intWidth, strDrawText) + picTask.CurrentX
        
        
        'picTask.CurrentY = 2
        'picTask.CurrentX = intWidth * (i - 1)
        'picTask.CurrentX = TaskCenter(intWidth, strTitle) + picTask.CurrentX
        'strDrawText = TaskText(intWidth, strTitle)
        picTask.Print strDrawText;
    Next i
    
        
End Sub

Public Sub HandleCTCP(strNick As String, strData As String)
    strData = RightR(strData, 1)
    strData = LeftR(strData, 1)
    
    Dim strCom As String, strParam As String
    Seperate strData, " ", strCom, strParam
    
    Select Case LCase(strCom)
        Case "version"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested your client version"
            CTCPReply strNick, "VERSION projectIRC for Windows"
        Case "ping"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just pinged you"
            CTCPReply strNick, "PING 0 seconds, cause projectIRC is elite"
    End Select
End Sub


Sub interpret(strData As String)
    Dim parsed As ParsedData, AllParams As String, intTemp As Integer
    Dim i As Integer, strChan As String, strTemp As String
    
    strData = Replace(strData, Chr(13), "")
    strData = Replace(strData, Chr(10), "")
    ParseData Replace(strData, Chr(13), ""), parsed
    AllParams = Params(parsed, 1, -1)
    If parsed.strCommand = "" Then Exit Sub
    
    'PutData Status.DataIn, strMyNick & "*****" & parsed.strNick & "~" & parsed.strCommand & "~" & AllParams
    
    Select Case LCase(parsed.strCommand)
        Case "001"
            strMyNick = Params(parsed, 1, 1)
            PutData Status.DataIn, "* " & Params(parsed, 2, -1)
            Status.Update
            Exit Sub
        Case "002", "003", "004", "005"
            PutData Status.DataIn, "* " & Params(parsed, 2, -1)
            Exit Sub
        Case "ping"
            SendData "PONG :" & AllParams
            PutData Status.DataIn, strColor & "03Ping? Pong! [" & AllParams & "]"
            Exit Sub
        Case "join"
            If LCase(parsed.strNick) = LCase(strMyNick) Then
                intTemp = NewChannel(AllParams)
                Client.SendData "MODE " & AllParams
            Else
                intTemp = GetChanIndex(parsed.strParams(1))
                If intTemp = -1 Then Exit Sub
                Channels(intTemp).AddNick parsed.strNick
                PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has joined " & strBold & Channels(intTemp).strName
            End If
            Exit Sub
        Case "privmsg"
            strChan = Params(parsed, 1, 1)
            If Left(strChan, 1) = "#" Then  'privmsg to channel
                intTemp = GetChanIndex(strChan)
                If intTemp <> -1 Then Channels(intTemp).PutText parsed.strNick, Params(parsed, 2, -1)               '
            ElseIf parsed.strNick = strMyNick Then
                If Params(parsed, 2, 2) = strAction & "VERSION" & strAction Then    'version
                    'Client.SendData "CTCP REPLY " & strChan & " VERSION :jIRC for Windows9x"
                    Client.SendData "NOTICE " & parsed.strNick & " :VERSION projectIRC for Win32"
                End If
                GoTo msg
            
            Else    'send to query window
msg:
                strTemp = Params(parsed, 2, -1)
                If Left(strTemp, 1) = strAction Then
                    HandleCTCP parsed.strNick, strTemp
                    Exit Sub
                End If
                
                If QueryExists(parsed.strNick) Then
                    'MsgBox "exists"
                    intTemp = GetQueryIndex(parsed.strNick)
                    If intTemp = -1 Then Exit Sub
                    
                    If Queries(intTemp).strHost <> parsed.strFullHost Then
                        Queries(intTemp).strHost = RightOf(parsed.strFullHost, "!")
                        Queries(intTemp).lblHost = RightOf(parsed.strFullHost, "!")
                        
                    End If
                    Queries(intTemp).Caption = parsed.strNick
                    Queries(intTemp).strNick = parsed.strNick
                    Queries(intTemp).lblNick = parsed.strNick
                    Queries(intTemp).PutText parsed.strNick, strTemp
                Else
                    'MsgBox "doesnt"
                    NewQuery parsed.strNick, parsed.strFullHost
                    intTemp = GetQueryIndex(parsed.strNick)
                    If intTemp = -1 Then Exit Sub
                    Queries(intTemp).Caption = parsed.strNick
                    Queries(intTemp).strNick = parsed.strNick
                    Queries(intTemp).lblNick = parsed.strNick
                    Queries(intTemp).PutText parsed.strNick, strTemp
                End If
            End If
            Exit Sub
        Case "nick"
            If parsed.strNick = strMyNick Then
                strMyNick = Params(parsed, 1, 1)
                PutData Status.DataIn, strColor & "03Your nick is now " & strBold & strMyNick
                ChangeNick parsed.strNick, Params(parsed, 1, -1)
                Status.Update
            Else
                ChangeNick parsed.strNick, Params(parsed, 1, 1)
            End If
            Exit Sub
        Case "part"
            If parsed.strNick = strMyNick Then Exit Sub
            intTemp = GetChanIndex(parsed.strParams(1))
            'MsgBox intTemp & "~" & parsed.strParams(1) & "~"
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).RemoveNick parsed.strNick
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has left " & strBold & Channels(intTemp).strName
            If parsed.strNick = strMyNick Then Unload Channels(intTemp)
            Exit Sub
        Case "353" 'nick list!
            'MsgBox parsed.strParams(3)
            intTemp = GetChanIndex(parsed.strParams(3))
            If intTemp = -1 Then Exit Sub
            Dim strNicks() As String
            strNicks = Split(Params(parsed, 4, -1), " ")
            For i = LBound(strNicks) To UBound(strNicks)
                Channels(intTemp).AddNick strNicks(i)
            Next i
            Exit Sub
        Case "mode"     'set mode
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " sets mode: " & Params(parsed, 2, -1)
            ParseMode Params(parsed, 1, 1), Params(parsed, 2, -1)
            Exit Sub
        Case "quit"     'quit
            NickQuit parsed.strNick, Params(parsed, 1, -1)
            Exit Sub
        Case "kick"     'kick
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & Params(parsed, 2, 2) & strBold & " was kicked from " & strBold & Params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & Params(parsed, 3, -1) & " ]"
            Channels(intTemp).RemoveNick Params(parsed, 2, 2)
            
            '* If user, close channel
            If Params(parsed, 2, 2) = strMyNick Then
                Channels(intTemp).Tag = "NOPART"
                Unload Channels(intTemp)
                PutData Status.DataIn, strColor & "03" & "You were kicked from " & strBold & Params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & Params(parsed, 3, -1) & " ]"
            End If
            
            Exit Sub
        Case "332"  'topic!
            intTemp = GetChanIndex(Params(parsed, 2, 2))
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).rtbTopic.Text = ""
            PutData Channels(intTemp).rtbTopic, Params(parsed, 3, -1)
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.SelLength = 1
            Channels(intTemp).rtbTopic.SelText = ""
            PutData Channels(intTemp).DataIn, strColor & "03Topic is """ & strColor & Params(parsed, 3, -1) & strColor & "03"""
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.Tag = "locked"
            Exit Sub
        Case "topic"    'change in topic!
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).rtbTopic.Text = ""
            PutData Channels(intTemp).rtbTopic, Params(parsed, 2, -1)
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.SelLength = 1
            Channels(intTemp).rtbTopic.SelText = ""
            PutData Channels(intTemp).DataIn, strColor & "03Topic changed by " & strBold & parsed.strNick & strBold & " : " & Params(parsed, 2, -1)
            Exit Sub
        Case "333"  'topic on param2 set by param3, on param4
            intTemp = GetChanIndex(Params(parsed, 2, 2))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03Topic set by " & strBold & Params(parsed, 3, 3) & strBold
            
            'Exit Sub
        Case "366"  'end of names list
            Exit Sub
        Case "324"  'set channel modes
            ParseMode Params(parsed, 2, 2), Params(parsed, 3, -1)
            Exit Sub
        Case "notice"
            On Error Resume Next
            If parsed.strNick = "" Then
                PutData Client.ActiveForm.DataIn, strColor & "05" & Params(parsed, 2, -1)
            Else
                PutData Client.ActiveForm.DataIn, strColor & "05" & strBold & "NOTICE" & strBold & strColor & " " & strBold & parsed.strNick & strBold & ":" & Chr(9) & Params(parsed, 2, -1)
            End If
            
            Exit Sub
        Case "433"  'nick name already in use
            If Params(parsed, 2, 2) = strMyNick Then
                strTemp = strMyNick
                strMyNick = strOtherNick
                strOtherNick = strMyNick
                Client.SendData "NICK " & strMyNick
            End If
        Case "372"  'MOTD
            PutData Status.DataIn, Params(parsed, 2, -1)
            Exit Sub
        Case "375"  'start of MOTD
            PutData Status.DataIn, strColor & "02" & Params(parsed, 2, -1)
            Exit Sub
        Case "376"  'end of MOTD
            Exit Sub
        Case "251", "252", "253", "254", "255", _
            "265", "266" 'server info, users, ops, channels, clients
            PutData Status.DataIn, strColor & "06" & Params(parsed, 2, -1)
            Exit Sub
    End Select
    PutData Status.DataIn, "*** " & strBold & parsed.strCommand & strBold & " " & AllParams ' & " [" & parsed.strFullHost & "]"
End Sub


Sub SendData(strData As String)
    On Error Resume Next
    sock.SendData strData & Chr(10)
End Sub


Private Sub Command1_Click()
    picTask.Cls
    DrawToolbar
End Sub

Private Sub IDENT_ConnectionRequest(ByVal requestID As Long)
    IDENT.Close
    IDENT.Accept requestID
    IDENT.SendData IDENT.LocalPort & ", " & IDENT.RemotePort & " : USERID : UNUX : " & strMyNick & vbCrLf
    
    Dim i As Integer
    For i = 1 To 500
        DoEvents
    Next i
    
    IDENT.Close
End Sub

Private Sub IDENT_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    IDENT.GetData dat, vbString
    
    If dat Like "*, *" Then
        dat = LeftR(dat, 2)
        PutData Status.DataIn, "*** IDENT : " & dat
        dat = dat & " : USERID : UNIX : " & strMyIdent
        Client.SendData dat
        PutData Status.DataIn, "*** IDENT reply : " & dat
        'MsgBox "~" & dat & "~"
        Dim i As Integer
        IDENT.Close
    End If
End Sub

Private Sub IDENT_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, Chr(Color) & "04IDENT Error " & strColor & Description
End Sub

Private Sub MDIForm_Activate()
    DrawToolbar
End Sub

Private Sub MDIForm_Load()
    '* Use this until setting files implemented
    DrawToolbar
    strServer = "irc.otherside.com"
    sngLastX = 1
    strMyNick = "pIRCu"
    strOtherNick = "OtherNick"
    strFullName = "projectIRC User"
    strMyIdent = "projectIRC"
    lngPort = 6667
    DoEvents
    Me.Visible = True
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    intHover = 0
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Client.SendData "QUIT :using projectIRC, closed"
    Dim i As Integer
    For i = 0 To 50
        DoEvents
    Next i
    Cancel = 0
End Sub

Private Sub MDIForm_Resize()
    CoolBar1.Width = (Me.ScaleWidth / 15) + 1
    shpTask.Width = (Me.Width / 15) - 8
'    MsgBox Me.ScaleWidth

    
    'intHover = 0
    'DrawToolbar
    SetWinFocus intActive
    DrawToolbar
    
    
    
End Sub

Private Sub MDIForm_Terminate()
    Client.SendData "QUIT :Client closed, using jIRC"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Client.SendData "QUIT :Client closed, using jIRC"
End Sub

Sub mnu_File_Connect_Click()
    Select Case mnu_File_Connect.Caption
        Case "&Connect"
            '* Connect
            sock.Close
            mnu_File_Connect.Caption = "&Cancel"
            sock.RemoteHost = strServer
            sock.RemotePort = lngPort
            sock.Connect
            IDENT.Close
            On Error Resume Next
            IDENT.Listen
            PutData Status.DataIn, strColor & "02Connecting to " & strBold & strServer & strBold & " port " & strBold & lngPort
        Case "&Cancel"
            '* Cancel
            IDENT.Close
            sock.Close
            mnu_File_Connect.Caption = "&Connect"
            PutData Status.DataIn, strColor & "05Connection attempt cancelled"
    End Select
End Sub


Sub mnu_File_Disconnect_Click()
    mnu_File_Connect.Enabled = True
    mnu_File_Disconnect.Enabled = False
    PutData Status.DataIn, strColor & "05Disconnected from " & sock.RemoteHost
    sock.Close
    Status.lblServer = "not connected"
    Status.Update

End Sub


Private Sub mnu_File_Options_Click()
    Options.Show 1
End Sub

Private Sub mnu_File_Quit_Click()
    Client.SendData "QUIT :Using projectIRC, closed"
    IDENT.Close
    sock.Close
    Dim i As Integer
    For i = 1 To 1000
        DoEvents
    Next i
    Unload Me
End Sub

Private Sub mnu_Help_About_Click()
    About.Show vbModal
End Sub

Private Sub mnu_Tile_Vertically_Click()
    Client.Arrange vbTileVertical
End Sub

Private Sub mnu_View_Status_Click()
    mnu_View_Status.Checked = Not mnu_View_Status.Checked
    Status.Visible = mnu_View_Status.Checked
End Sub

Private Sub mnu_Window_Cascade_Click()
    Client.Arrange vbCascade
End Sub

Private Sub mnu_Window_TileH_Click()
    Client.Arrange vbTileHorizontal
End Sub


Private Sub picTask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim which As Integer, wincnt As Integer, wid As Integer
    wincnt = WindowCount
    
    wid = picTask.ScaleWidth \ wincnt
    which = Int((X \ wid) + 0.5) + 1
    sngLastX = X
    
    '* Which now contains which button was clicked
    intActive = which
    DrawToolbar
    SetWinFocus which
End Sub


Private Sub picTask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim which As Integer, wincnt As Integer, wid As Integer
    wincnt = WindowCount
    
    wid = picTask.ScaleWidth \ wincnt
    which = Int((X \ wid) + 0.5) + 1
    'sngLastX = X
    intHover = which
    If Val(picTask.Tag) = intHover Then Exit Sub
    Client.DrawToolbar
    
    
    'bReDraw = False
End Sub


Private Sub sock_Close()
    PutData Status.DataIn, strColor & "02Disconnected by SERVER from " & strServer
    sock.Close
    IDENT.Close
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False
    Status.lblServer = "not connected"
    Status.Update

End Sub

Private Sub sock_Connect()
    mnu_File_Connect.Enabled = False
    mnu_File_Disconnect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    PutData Status.DataIn, strColor & "03Connected to " & strServer
    
    SendData "PASS password"
    SendData "NICK " & strMyNick
    SendData "USER " & strMyNick & " " & sock.LocalHostName & " irc :" & strFullName
    
    Status.lblServer = sock.RemoteHost
    Status.Update
    
    '* Let's close all open windows
    Dim i As Integer
    
    For i = 1 To intChannels
        Channels(i).Tag = "NOPART"
        Unload Channels(i)
    Next i
    
    For i = 1 To intQueries
        Unload Queries(i)
    Next i
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, AllParams As String
    Dim strData() As String, i As Integer
    
    sock.GetData dat, vbString
    
    '* this'll stay for about half a second
    strData = Split(dat, Chr(10))
    
    For i = LBound(strData) To UBound(strData)
        interpret strData(i)
    Next i
    
End Sub


Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, strColor & "04ERROR : " & Description
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnu_File_Connect_Click
        Case 2
            mnu_File_Disconnect_Click
        Case 3
            mnu_File_Options_Click
    End Select
End Sub


