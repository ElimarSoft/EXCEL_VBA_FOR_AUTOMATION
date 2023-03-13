'Find a Window with Class & Caption
Private Function FindWindowText(h1 As Long, wClass As String, Text As String)
    
    Dim h2 As Long: h2 = 0
    FindWindowText = apiFindWindowEx(h1, 0, wClass, Text)
    
End Function

'Get text from Window
Private Function GetText(h1 As Long) As String
    Dim Text As String: Text = vbNullString
    Dim Len1 As Long: Len1 = apiSendMessage(h1, WM_GETTEXTLENGTH, 0, 0)
    DoEvents
    If Len1 > 0 Then
        Text = Space(Len1)
        Dim res1 As Long: res1 = apiSendMessage(h1, WM_GETTEXT, Len1 + 1, ByVal Text)
        DoEvents
    End If
    GetText = Text

End Function

'Find a Window with Class and repetition count
Private Function FindWindowMul(h1 As Long, wClass As String, count As Integer)
    
    Dim h2 As Long: h2 = 0
    
    Do
        h2 = apiFindWindowEx(h1, h2, wClass, vbNullString)
        'Debug.Print Hex(h2) + " " + GetText(h2)
        If h2 = 0 Then Exit Do
        If count > 1 Then
            count = count - 1
        Else
            Exit Do
        End If
    Loop

    FindWindowMul = h2

End Function
      'Sets a checkbox checking status with pixel reading
Public Sub CheckBoxSet(h1 As Long, Checked As Boolean)
    
    Dim rect1 As winRect
    apiGetWindowRect h1, rect1
    Dim x As Integer: x = (rect1.Right + rect1.Left) / 2
    Dim y As Integer: y = (rect1.Bottom + rect1.Top) / 2
    
    Dim hdcMemDC As Long: hdcMemDC = GetWindowDC(h1)
    Dim Color1 As Long: Color1 = GetPixel(hdcMemDC, 5, 10)
    'Debug.Print Hex(Color1)
    If (Checked And Color1 = &HFFFFFF) Or _
        (Not Checked And Color1 <> &HFFFFFF) Then
        apiSendMessage h1, WM_KEYDOWN, VK_SPACE, 0
        apiSleep (200)
        apiSendMessage h1, WM_KEYUP, VK_SPACE, 0
    End If
End Sub

'Left Click event moving the cursor temporaly
Private Sub LeftClick(x As Integer, y As Integer)
    
    Dim pt1 As POINT
    
    GetCursorPos pt1
    apiSetCursorPos x, y
    apiSleep 50
    apiMouseEvent MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    apiSleep 50
    apiMouseEvent MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    apiSetCursorPos pt1.x, pt1.y

End Sub
      
'Left Click event without moving the cursor     
Private Sub LeftClick2(h1 As Long, x As Integer, y As Integer)

    Dim coord As Long: coord = 256
    coord = coord * 256 * y + x
    
    apiSendMessage h1, WM_LBUTTONUP, MK_LBUTTON, ByVal coord
    SetCapture (h1)
    apiSendMessage h1, WM_LBUTTONUP, 0, ByVal coord
    
    apiSleep (20)
    apiSendMessage h1, WM_LBUTTONDOWN, 0, ByVal coord
    apiSleep (20)
    apiSendMessage h1, WM_LBUTTONUP, 0, ByVal coord

End Sub

'Same for double click
      Private Sub DblClick(h1 As Long, x As Integer, y As Integer)

    Dim coord As Long: coord = 256
    coord = coord * 256 * y + x
    
    apiSendMessage h1, WM_LBUTTONUP, MK_LBUTTON, ByVal coord
    SetCapture (h1)
    apiSendMessage h1, WM_LBUTTONUP, 0, ByVal coord
    apiSleep (100)
    apiSendMessage h1, WM_LBUTTONDBLCLK, 0, ByVal coord

End Sub

'Get listbox items 
Private Function GetListBox(h1 As Long) As String()

   Dim count As Integer: count = apiSendMessage(h1, LB_GETCOUNT, 0, 0)

   Dim ItemText() As String
   Dim Text5 As String
   ReDim ItemText(count)
   
   Dim ItemData As Variant: ItemData = apiSendMessage(h1, LB_GETSEL, 0, 0)
   
   Dim n As Integer
   Dim len5 As Integer
   For n = 1 To count
    len5 = apiSendMessage(h1, LB_GETTEXTLEN, n - 1, 0)
    Text5 = Space(len5)
    Call apiSendMessage(h1, LB_GETTEXT, n - 1, ByVal Text5)
    ItemText(n) = Text5
   Next n

   GetListBox = ItemText

End Function

 'Click a Button
Private Sub Click(h1 As Long)
    apiSendMessage h1, BM_CLICK, 0, 0
End Sub

'Set Text
Private Sub SetText(h1 As Long, Text As String)
    apiSendMessage h1, WM_SETTEXT, 0, ByVal Text
End Sub
        
      
