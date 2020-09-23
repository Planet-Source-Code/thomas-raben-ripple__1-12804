Attribute VB_Name = "ModDeclares"
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Public counter As Integer
Public counter2 As Integer

Public Way As Boolean

Public Sub DumpToWindow(TargetBox As Control)
    Dim Desktop As Long
    
    Desktop = GetDC(GetDesktopWindow)
    
    ww = 1600
    hh = Screen.Height / Screen.TwipsPerPixelY
    
    BitBlt TargetBox.hDC, 0, 0, ww, 200, Desktop, 0, hh - 400, &HCC0020
    'BitBlt TargetBox.hDC, 0, 0, 1600, 200, frmMain.Background.hDC, 0, 0, vbSrcAnd
End Sub

Public Sub Ripple(Source As Control, Dest As Control)
    Dim i As Integer
    Dim x As Double
    
    ww = 1600
    hh = Screen.Height / Screen.TwipsPerPixelY
    
    Dest.Cls
    
    'Dest.Refresh
    For i = 200 To 0 Step -1
        
        If i < 201 Then
            x = Cos(counter2 / (1 + (i / 20))) * ((i / 10))
            BitBlt Dest.hDC, x, i, ww, 1, Source.hDC, 0, 200 - i, vbSrcCopy
        End If
    
        counter = counter + 1
        
        If counter < 192 Then
            counter2 = counter
        End If
        
        If counter > 192 Then
            counter = 0
        End If
        
    Next i
    Dest.Refresh
    

End Sub
