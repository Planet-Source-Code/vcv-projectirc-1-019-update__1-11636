Attribute VB_Name = "modWindows"

Function GetWindowIndex(strCaption As String)
    Dim i As Integer
    For i = 1 To WindowCount
        If GetWindowTitle(i) = strCaption Then
            GetWindowIndex = i
            Exit Function
        End If
    Next i
    GetWindowIndex = -1
End Function

Function GetWindowTitle(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then GetWindowTitle = "Status": Exit Function
    
    cnt = 2
    For i = 1 To intChannels
        If cnt = intWhich Then
            GetWindowTitle = Channels(i).strName
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        If cnt = intWhich Then
            GetWindowTitle = Queries(i).strNick
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
            'GetWindowTitle = DCCChats(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        If cnt = intWhich Then
            'GetWindowTitle = DCCSends(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
final:
    GetWindowTitle = ""
End Function

Sub SetWinFocus(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then Status.SetFocus: Exit Sub
    
    cnt = 2
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            If cnt = intWhich Then
                Channels(i).SetFocus
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then
            If cnt = intWhich Then
                Queries(i).SetFocus
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    
    'WindowCount = cnt + 1 'add 1 for status window

End Sub

Function TaskCenter(intActual As Integer, strText As String) As Integer
    TaskCenter = (intActual - Client.picTask.TextWidth(strText)) / 2
End Function


Function TaskText(intWidth As Integer, strText As String) As String
    'MsgBox intWidth & ".." & Client.picTask.TextWidth(strText) & ".."
    Dim lastWidth As Integer, i As Integer, strBuf As String
    Dim intTemp As Integer
    
    For i = 1 To Len(strText)
        strBuf = Left(strText, i) & "..."
        intTemp = Client.picTask.TextWidth(strBuf) + 8
        
        If intTemp >= intWidth Then
            TaskText = Left(strText, i - 1) & "..."
            Exit Function
        End If
    Next i
    TaskText = strText
        
End Function

Function WindowCount() As Integer
    Dim cnt As Integer, i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then cnt = cnt + 1
    Next i
    
    
    WindowCount = cnt + 1 'add 1 for status window
    
    '* Add DCC and stuff here

End Function


