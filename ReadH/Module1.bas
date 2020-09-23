Attribute VB_Name = "Module1"
Global Counter As Integer
Global Init As Boolean
Global LineNo As Integer
Global RowNo As Integer
Global TextBox As Integer
Global FirstKey As Boolean
Global OldLine As Integer
Global BinFile As String
Global FilName As String
'Global Const MPath = App.Path = "C:\Magic"
Global ReturnData As String
Global ReturnLength As Integer
Global TestE3 As String
Global Un29 As String
Global RdE3 As String
Global RdE31 As String
Global RdE32 As String
Global RdE33 As String
Global RdE34 As String
Global RdE35 As String
Global RdE36 As String
Global RdE37 As String
Global RdE38 As String
Global RdE39 As String
Global Chk29 As String
Global SafeCheck(16) As String
Global Janitor(8) As String
Global JanitorFlag As Boolean
Global SendEEprom As Boolean
Global ShowPack As Boolean
Global ByteDelay
Global Const ByteMin = &H10000
Global Const ByteMax = &H200000
Global sndList As String
Global OutByte As Integer
Global Ix As Integer
Global Elist As String
Global E3Flag As Boolean
Global BufLen As Integer
Global TempAtr As Integer
Global HoldAtr As Integer
Global OutBuf As String
Global InBuf As String
Global MaxP3Limit As Integer
Global MaxP2Limit As Integer
Global TimeOut As Integer
Global StateChanged As Integer
Global Mask As Integer
Global Const BlockHeader1 = "48 42 00 00 14 "
Global Const BlockFooter1 = "R01"
Global Const BlockHeader2 = "09 10 00 00 24 25 60 B5 03 "
Global Const BlockHeader3 = "02 BB 00 0C "
Global Const BlockFooter2 = "R02"
Global Const Encrypt = "&H49&H53&H4F&H37&H38&H31&H36&H50&H72&H6F&H67"
Global BlockAdd(255) As String
Global BlockSig(255) As String
Global Form4Start As Boolean
Global MagiString As String
Global Start As Integer

Sub Main()
    Dim Form As New Form1
    Dim I As Long
    Dim J As Integer
    Dim Hdl As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim Msg As String
    Dim Hold As String
    
    
    Form4.Show
    Form4.Refresh
    For I = 1 To 42 Step 4
        Hold = Mid$(Encrypt, I, 4)
        MagiString = MagiString + Chr(Hold)
    Next
    
    For J = 1 To 5
        For I = 1 To &H100000
    
        Next
    Next J
    
    Load Form1
    
    For I = 1 To &H100000
    
    Next
    Unload Form4
    Set Form = Nothing
    Form1.Show
End Sub

Public Sub UpdatEEProm()
    Dim I As Long
    Dim Hold1 As Integer
    Dim Hold As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim Packet As String
    
    Form1.XPLList.Clear
    
    I = (Ix * 16) + 32768
    Form1.txtStat.Text = "Writing Block at: " + Hex(I)
    Select Case Hex(I)
        Case "8020" 'Fuse
            Temp = Left$(Elist, 2)
            Temp1 = Mid$(Elist, 4, 2)
            Hold1 = ConvertHex(Temp)
            Hold2 = ConvertHex(Temp1)
            If Hold1 + Hold2 <> 255 Then
                MsgBox "                   If you send this packet to your card" + Chr$(13) + Chr$(10) + "                          your card will be 99'ed.", vbOKOnly, "WARNING == Packet will 99 card."
                Exit Sub
            End If
        Case "84F0" 'IRD No
            Temp = "00 00 00 00"
            Mid$(Elist, 37, 11) = Temp
            
        Case "83D0" 'IRD No
            Temp = "00 00 00 00"
            Mid$(Elist, 1, 11) = Temp
        Case "8590" 'E3 hole
            If E3Flag Then
            
            End If
            
    End Select
    
    Temp = Hex$(I)
    Packet = ""
    If E3Flag Then
        Temp1 = "48 42 00 00 1D"
        Form1.XPLList.AddItem Temp1
        Temp1 = "R01"
        Form1.XPLList.AddItem Temp1
        Packet = "60 D5 02 85 8E E3 14 10 "
        Packet = Packet + Left$(Temp, 2) + " " + Right$(Temp, 2) + " "
        Packet = Packet + Elist + "00 BB 00"
        Form1.XPLList.AddItem Packet
        Temp1 = "R02"
        Form1.XPLList.AddItem Temp1
    Else
        Temp1 = "48 40 00 00 67"
        Form1.XPLList.AddItem Temp1
        Temp1 = "R01"
        Form1.XPLList.AddItem Temp1
        Packet = "09 11 00 00 30 60 00 06 39 00 04 F4 22 33 CF 03"
        Packet = Packet + " 0E 1B 00 CF 03 0E 1B 00 CF 03 0E 1B 00 CF 03 0E"
        Packet = Packet + " 1B 00 CF 03 0E 1B 00 BB 00 12 00 00 00 00 00 00"
        Packet = Packet + " 00 00 00 00 00 00 00 00 00 00 00 00 12 00 00 00"
        Packet = Packet + " 00 00 35 08 13 86 46 13 8A 1C 00 00 00 00 00 60 "
        Packet = Packet + " BB 15 10 " + Left$(Temp, 2) + " " + Right$(Temp, 2) + " "
        Packet = Packet + Elist + " 00 00"
        Form1.XPLList.AddItem Packet
        Temp1 = "R02"
        Form1.XPLList.AddItem Temp1
    End If
    
    SendEEprom = True
    RunXPL
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    
    'Comm.Output = ""
    Form1.lblSNd.Caption = ""
    If E3Flag = True Then
        If Left$(InBuf, 5) = "60 D5" Then
            Form1.txtStat.Text = "Write Block at: " + Hex(I) + " was successful."
        Else
            Form1.txtStat.Text = "Write Block at: " + Hex(I) + " was unsuccessful."
            StartFlag = False
            Temp = Form1.Comm.Input
            SendEEprom = False
            Exit Sub
        End If
    Else
        If Right$(InBuf, 6) = "90 80 " Then
            Form1.txtStat.Text = "Write Block at: " + Hex(I) + " was successful."
        Else
            Form1.txtStat.Text = "Write Block at: " + Hex(I) + " was unsuccessful."
            StartFlag = False
            Temp = Form1.Comm.Input
            SendEEprom = False
            Exit Sub
        End If
                
    End If
    
    Temp = Form1.Comm.Input
    ResetATR
    DontReceive = True
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    
End Sub

Public Sub WriteImage()
    Dim I As Integer
    Dim Ulist As String
    Dim cnt As Integer
    cnt = 0
    
    If Form2.EEPromList.ListCount = 0 Then
        Form1.txtStat.Text = "No EEProm image to write."
        Form2.txtStat.Text = "No EEProm image to write."
        Exit Sub
    End If
   
    For Ix = 2 To 255
        Form2.EEPromList.ListIndex = Ix
        Form2.UpdateList.ListIndex = Ix
        Elist = Form2.EEPromList.Text
        Ulist = Form2.UpdateList.Text
        If Elist <> Ulist Then
            cnt = cnt + 1
            UpdatEEProm
            'rem update the list
            'so we don't rewite it
            'if user selects write again
            Form2.UpdateList.ListIndex = Ix
            Form2.UpdateList.RemoveItem Ix
            Form2.UpdateList.AddItem Elist, Ix
            If SendEEprom = False Then
                Exit Sub
            End If
        End If
    Next Ix
    If cnt = 0 Then
        Form1.txtStat.Text = "No changes to EEProm."
        Form2.txtStat.Text = "No changes to EEProm."
    End If
    
End Sub

'all of XPL is in the XPLList box
'read it verify and execute

Public Sub RunXPL()
    Dim I As Integer
    Dim J As Integer
    Dim Hold As Integer
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim Index As Integer
    Dim Msg As String
    Dim Test As String
    Dim H As Integer
    Dim k As Integer
    Dim L As Integer
    Dim M As Integer
    Dim Z As Integer
    Dim x As Integer
    Dim Q As Integer
    Dim RecData As String
    Dim NumberNeeded As Integer
    Dim DontAdd As Boolean
    Dim RetBytes As Integer
    Dim PacketBytes As Integer
    Dim sendText As String
    Dim PackLen As Integer
    Dim PackDiff As Integer
    Dim cnt As Integer
    
    Msg = ""
    Temp = ""
    Temp1 = ""
    Temp2 = ""
    Position = 0
    InBuf = ""
    Z = 0
    Form1.sendList.Text = ""
    sndList = ""
    If SendEEprom = False Then
        ShowList
    End If
    DontAdd = False
    cnt = 0
Top:
    Test = ""
    Index = Form1.XPLList.ListCount - 1
    For I = Z To Index
        Form1.XPLList.ListIndex = I
        Temp = Form1.XPLList.Text
        
        For J = 1 To Len(Temp)
           'check first byte for ; or
           'or number or rem or space
            Hold = Asc(Mid$(Temp, J, 1))
            Select Case Hold
                Case 39, 59, 96
                    'ok this is remarks get next line
                    J = Len(Temp)
                    DontAdd = True
                    GoTo GetNextLine
                Case 48 To 57
                    'ok it is a number add to string
                    Test = Test + Mid$(Temp, J, 1)
                Case 65 To 70
                    'ok alpha upper
                    Test = Test + Mid$(Temp, J, 1)
                Case 97 To 102
                    'ok alpha lower
                    'make it upper
                    Hold = Hold - 32
                    Test = Test + Chr(Hold)
                Case 32
                    'ok it is a space
                    If (J < Len(Temp)) And (Right$(Test, 1) <> " ") And (J <> 1) Then
                        Test = Test + " "
                    End If
                Case 82, 114
                    'ok this is #bytes to return
                    GoTo GetReturn
                Case 88, 120
                    'ok this is for menu
                    'check how many we need
                    x = J + 3
                    NumberNeeded = 1
                    Do
                        If UCase(Mid$(Temp, x, 1)) = "X" Then
                            NumberNeeded = NumberNeeded + 1
                            x = x + 3
                        Else
                            Exit Do
                        End If
                    Loop While x < Len(Temp)
                    RecData = InputBox(Mid$(Temp, x, Len(Temp) - x), "")
                    If RecData = "" Then
                        GoTo Handler
                    End If
                    Hold = 0
                    For x = 1 To (NumberNeeded * 2) Step 2
                        Hold = Hold + ConvertHex(Mid$(RecData, x, 2))
                        Test = Test + Hex(Hold)
                    Next x

                    J = Len(Temp)
                    DontAdd = False
                    GoTo GetNextLine
                    
                Case Else
                    Msg = "XPL file incorrect format."
                    GoTo Handler
            End Select
        Next J

GetNextLine:
    If (Len(Test) > 0) And (DontAdd = False) Then
        Test = Test + " "
    Else
        DontAdd = False
    End If
    
    Next I

GetReturn:
    If I > Index Then
        Msg = "XPL file incorrect format."
        GoTo Handler
    End If
    
        'get # of bytes to receive
        Form1.XPLList.ListIndex = I
        Temp = Form1.XPLList.Text
        If UCase(Left$(Temp, 1)) <> "R" Then
            Msg = "No receieve value in packet."
            GoTo Handler
        End If
        Temp1 = Right$(Temp, Len(Temp) - 1)
        Hold = ConvertHex(Temp1)
        
        RetBytes = Hold + 2
        
        'now check next line for more receieve bytes
        Do
            I = I + 1
            If I > Index Then
                Exit Do
            End If
            Form1.XPLList.ListIndex = I
            Temp = Form1.XPLList.Text
            If UCase(Left$(Temp, 1)) = "R" Then
                'we have more to receive
                Temp1 = Right$(Temp, Len(Temp) - 1)
                Hold = ConvertHex(Temp1)
                RetBytes = RetBytes + Hold
            Else
                I = I - 1
                Form1.XPLList.ListIndex = I
                Exit Do
            End If
        Loop While (UCase(Left$(Temp, 1) = "R")) And (I < Index)

        For H = 1 To Len(Test) Step 3
            Temp = "&H" + (Mid$(Test, H, 2))
            TempAtr = CInt(Temp)
            If H > Len(Test) - 3 Then
                PacketBytes = TempAtr + 2
            End If
            ConvertAtr
            Temp2 = Temp2 + Chr(HoldAtr)
        Next H
        BufLen = 1
        PackLen = Len(Temp2)
        
        MaxP2Limit = (RetBytes - 2) + PackLen
        MaxP3Limit = RetBytes + 5
        If RetBytes > 256 Then
            InTime = ByteDelay * 25
        Else
            InTime = ByteDelay * 17
        End If
        
        'ok we have a line now send to the card
        If Len(Temp2) > 0 Then
            Form1.PBar.Max = Len(Temp2)
        Else
            InTime = InTime
        End If
        
        J = 1
        M = 1
        For k = 1 To Len(Temp2)
            If k = Len(Temp2) Then
                BufLen = RetBytes
            End If
            Form1.lblSNd.Caption = Mid$(Test, J, 2)
            If M = 48 Then
                sndList = sndList + Chr$(13) + Chr$(10)
            End If
            sndList = sndList + Mid$(Test, J, 3)
            Form1.sendList.Text = sndList
            J = J + 3
            M = M + 3
            If M > 48 Then
                M = 1
            End If
            Form1.PBar.Value = k
            SendChar Mid$(Temp2, k, 1)

        Next k
        sendText = Right$(Form1.sendList.Text, 2)
        If (sendText <> Chr$(13) + Chr$(10)) Then
            sndList = sndList + Chr$(13) + Chr$(10)
            Form1.sendList.Text = sndList
        End If
        
        Temp = ""
        PackDiff = Len(InBuf) - PackLen
        'convert read
        For k = 1 To Len(InBuf)
            TempAtr = Asc(Mid$(InBuf, k, 1))
            ConvertAtr
            If HoldAtr < &H10 Then
                Temp = Temp + "0" + Hex$(HoldAtr) + " "
            Else
                Temp = Temp + Hex$(HoldAtr) + " "
            End If
        Next k

        'Test = Test + " " + Mid$(Temp, 4, 2)
        If (InStr(1, Temp, Left$(Test, 5)) = False) Then
            Msg = "Unable to read EEProm."
            StartFlag = False
            Temp = Form1.Comm.Input
            GoTo Handler
        End If
                
        M = 1
        J = 1
        Test = Right$(Temp, PackDiff * 3)
        For k = 1 To Len(Test) Step 3
            If M = 16 Then
                sndList = sndList + Chr$(13) + Chr$(10)
            End If
            sndList = sndList + Mid$(Test, J, 3)
            Form1.sendList.Text = sndList
            If (PackDiff > 2) And (J = 1) Then
                If PackDiff < 512 Then
                    M = 1
                Else
                    M = M + 1
                End If
                
            Else
                M = M + 1
            End If
            If PackDiff > 500 Then
                If J = 1536 Then
                    M = 1
                End If
            Else
                If J = (PackDiff * 3) - 8 Then
                    M = 1
                End If
            End If
            
            J = J + 3
            If M > 16 Then
                M = 1
            End If
        Next k
        If (Right$(sndList, 2) <> Chr$(13)) Then
            sndList = sndList + Chr$(13) + Chr$(10)
            Form1.sendList.Text = sndList
        End If
        
    Z = I + 1
    
    If Z < Index Then
        If RetBytes > 256 Then
            ResetATR
        End If
        cnt = cnt + 1
        'If cnt = 2 Then
        '    cnt = 0
        '    ResetATR
       ' End If
        
        Temp = Form1.Comm.Input
        Form1.Comm.InputLen = 0
        Temp1 = ""
        Temp2 = ""
        InBuf = ""
        Temp = ""
        
        GoTo Top
    End If
    InTime = ByteDelay * 17
    Temp1 = Form1.Comm.Input
    Form1.Comm.InputLen = 0
    Temp1 = ""
    Temp2 = ""
    'XPLList.Clear
    If SendEEprom = True Then
        InBuf = Temp
        Exit Sub
    End If
    Form1.txtStat.Text = "XPL file complete."
    Exit Sub
        
Handler:
    Form1.txtStat.Text = Msg
    Temp = Form1.Comm.Input
    Form1.Comm.InputLen = 0
    Temp1 = ""
    Temp2 = ""
    InBuf = ""
    Temp = ""
    'XPLList.Clear
    InTime = ByteDelay * 17
End Sub

Public Sub ShowList()

    Form1.lblATR.Visible = False
    Form1.txtATR.Visible = False
    Form1.Label(10).Caption = "HIDE PACKET RESULTS"
    Form1.Label(10).Enabled = True
    Form1.sendList.Visible = True
    Form1.XPLList.Visible = True
    Form1.Label9.Visible = True
    Form1.Label10.Visible = True
End Sub

Public Sub HideList()
    
    Form1.lblATR.Visible = True
    Form1.txtATR.Visible = True
    Form1.Label(10).Caption = "SHOW PACKET RESULTS"
    Form1.Label(10).Enabled = True
    Form1.sendList.Visible = False
    Form1.XPLList.Visible = False
    Form1.Label9.Visible = False
    Form1.Label10.Visible = False
End Sub

Public Function ConvertHex(s As String) As Integer
    Dim I As Integer
    Dim Temp As String
    Dim Hold As Integer
    
    If Len(s) = 1 Then
        Select Case UCase(s)
            Case "0"
                ConvertHex = 0
            Case "1"
                ConvertHex = 1
            Case "2"
                ConvertHex = 2
            Case "3"
                ConvertHex = 3
            Case "4"
                ConvertHex = 4
            Case "5"
                ConvertHex = 5
            Case "6"
                ConvertHex = 6
            Case "7"
                ConvertHex = 7
            Case "8"
                ConvertHex = 8
            Case "9"
                ConvertHex = 9
            Case "A"
                ConvertHex = 10
            Case "B"
                ConvertHex = 11
            Case "C"
                ConvertHex = 12
            Case "D"
                ConvertHex = 13
            Case "E"
                ConvertHex = 14
            Case "F"
                ConvertHex = 15
            Case Else
                ConvertHex = 0
            Exit Function
        End Select
    ElseIf Len(s) = 2 Then
        Select Case UCase(Left$(s, 1))
            Case "0"
                Hold = 0
            Case "1"
                Hold = 1 * 16
            Case "2"
                Hold = 2 * 16
            Case "3"
                Hold = 3 * 16
            Case "4"
                Hold = 4 * 16
            Case "5"
                Hold = 5 * 16
            Case "6"
                Hold = 6 * 16
            Case "7"
                Hold = 7 * 16
            Case "8"
                Hold = 8 * 16
            Case "9"
                Hold = 9 * 16
            Case "A"
                Hold = 10 * 16
            Case "B"
                Hold = 11 * 16
            Case "C"
                Hold = 12 * 16
            Case "D"
                Hold = 13 * 16
            Case "E"
                Hold = 14 * 16
            Case "F"
                Hold = 15 * 16
            Case Else
                ConvertHex = 0
                Exit Function
        End Select
        Select Case UCase(Right$(s, 1))
            Case "0"
                ConvertHex = Hold
            Case "1"
                ConvertHex = Hold + 1
            Case "2"
                ConvertHex = Hold + 2
            Case "3"
                ConvertHex = Hold + 3
            Case "4"
                ConvertHex = Hold + 4
            Case "5"
                ConvertHex = Hold + 5
            Case "6"
                ConvertHex = Hold + 6
            Case "7"
                ConvertHex = Hold + 7
            Case "8"
                ConvertHex = Hold + 8
            Case "9"
                ConvertHex = Hold + 9
            Case "A"
                ConvertHex = Hold + 10
            Case "B"
                ConvertHex = Hold + 11
            Case "C"
                ConvertHex = Hold + 12
            Case "D"
                ConvertHex = Hold + 13
            Case "E"
                ConvertHex = Hold + 14
            Case "F"
                ConvertHex = Hold + 15
            Case Else
                ConvertHex = 0
        End Select
    Else
        ConvertHex = 0
    End If
    
End Function

'*******************************
'
'   Sub to convert data
'   by reversing the bits
'   then inverting all the bits
'
'   Paramaters: TempAtr is byte
'   to convert
'
'*******************************

Public Sub ConvertAtr()
    HoldAtr = 0
Top:
    Select Case TempAtr
        Case Is > 127
            HoldAtr = HoldAtr + 1
            TempAtr = TempAtr - 128
            GoTo Top
        Case Is > 63
            HoldAtr = HoldAtr + 2
            TempAtr = TempAtr - 64
            GoTo Top
        Case Is > 31
            HoldAtr = HoldAtr + 4
            TempAtr = TempAtr - 32
            GoTo Top
        Case Is > 15
            HoldAtr = HoldAtr + 8
            TempAtr = TempAtr - 16
            GoTo Top
        Case Is > 7
            HoldAtr = HoldAtr + 16
            TempAtr = TempAtr - 8
            GoTo Top
        Case Is > 3
            HoldAtr = HoldAtr + 32
            TempAtr = TempAtr - 4
            GoTo Top
        Case Is > 1
            HoldAtr = HoldAtr + 64
            TempAtr = TempAtr - 2
            GoTo Top
        Case Is = 1
            HoldAtr = HoldAtr + 128
            
    End Select
    TempAtr = HoldAtr
    HoldAtr = 255 Xor TempAtr
    
End Sub

'********************************
'
'   Sub to output 1 character
'   to communications port
'
'   Paramaters: s as output byte
'
'********************************

Public Sub SendChar(s As String)
    Dim Temp As String
    
    'clear input buffer
    Form1.Comm.InputLen = 0
    Temp = Form1.Comm.Input
    Form1.Comm.RTSEnable = False
    Form1.Comm.RThreshold = 1
    TimeOut = 0
    Form1.Timer1.Enabled = False
    Form1.Timer1.Interval = 200
    Form1.Timer1.Enabled = True
    Form1.Comm.Output = s
    Do While (TimeOut = 0) Or (StateChanged = 0)
        DoEvents
    Loop
    'StateChanged = 0
    'TimeOut = 0
    'Timer1.Enabled = False
    'Timer1.Interval = 200
    'Temp = Comm.Input
    'Timer1.Enabled = True
    'Do While (TimeOut = 0) Or (StateChanged = 0)
    '    DoEvents
    'Loop
    'Comm.InputLen = 0
    StateChanged = 0
    
End Sub

Public Sub ResetATR()
    Dim Temp As String
    Dim I As Integer
    
    InBuf = ""
    If Form1.Comm.PortOpen = True Then
        Form1.Comm.PortOpen = False
    End If
    Form1.Comm.Settings = "9600,O,8,2"
    Form1.Comm.PortOpen = True
    TimeOut = 0
    Form1.Timer1.Enabled = False
    Form1.Timer1.Interval = 200
    Form1.Comm.RTSEnable = True
    Form1.Comm.DTREnable = False
    Form1.Comm.RThreshold = 1
    Form1.Comm.InputLen = 0
    MaxP2Limit = 13
    MaxP3Limit = 20
    Form1.Timer1.Enabled = True
    Form1.Timer1.Interval = 200
    Do While TimeOut = 0
        DoEvents
    Loop
    
    StateChanged = 0
    TimeOut = 0
    Form1.Timer1.Enabled = False
    Form1.Timer1.Enabled = True
    Form1.Comm.RTSEnable = False
    'rem P3 card
    Do While Len(InBuf) < 16
        If TimeOut = 1 Then
            'rem P2 card
            If Len(InBuf) > 11 Then
                Exit Do
            Else
                GoTo SendEnd
            End If
        End If
        DoEvents
    Loop
    
    For I = 1 To Len(InBuf)
        TempAtr = Asc(Mid$(InBuf, I, 1))
        ConvertAtr
        If HoldAtr < &H10 Then
            Temp = Temp + "0" + Hex$(HoldAtr) + " "
        Else
            Temp = Temp + Hex$(HoldAtr) + " "
        End If
        If I = 3 Then
            Mask = HoldAtr And &HF
        End If
    Next I
    
    Select Case Mask
        Case 1
            Form1.Comm.Settings = "9600,O,8,2"
        Case 2
            Form1.Comm.Settings = "19200,O,8,2"
        Case 3
            Form1.Comm.Settings = "38400,O,8,2"
        Case 5
            Form1.Comm.Settings = "115200,O,8,2"
    End Select
SendEnd:
    StateChanged = 0
    Temp = Form1.Comm.Input
    Form1.Comm.InputLen = 0
    Temp = Form1.Comm.Input
    Form1.Timer1.Enabled = False
End Sub



