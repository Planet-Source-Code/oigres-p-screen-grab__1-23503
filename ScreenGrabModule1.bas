Attribute VB_Name = "Module1"
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
        (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
        ipic As IPicture) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As _
        Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
        ByVal hObject As Long) As Long


Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
        ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal lScreenDC As Long, ByVal XSrc As Long, _
        ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
        ByVal hdc As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
        lpRect As RECT) As Long
Public Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
        ByVal y As Long, ByVal lp As Long)
Public Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal _
        y As Long)

Public Const debugme As Boolean = False

Public Function savepictureRoutine() As Boolean

    On Error GoTo errDialog
    'start off false
    savepictureRoutine = False

    If MDIForm1.ActiveForm Is Nothing Then Exit Function

    'get filename with commondialog

    MDIForm1.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn _
       Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    MDIForm1.CommonDialog1.CancelError = True
    MDIForm1.CommonDialog1.Filter = "All Files|*.*|JPEG|*.jpg|Bitmaps|*.bmp"
    MDIForm1.CommonDialog1.DefaultExt = "bmp"
    '
    MDIForm1.CommonDialog1.ShowSave

    If debugme = True Then MsgBox MDIForm1.CommonDialog1.FileName

    If MDIForm1.CommonDialog1.FileName <> "" Then

        'savepicure method saves only bitmap (icon if loaded icon file)
        SavePicture MDIForm1.ActiveForm.Picture1.Picture, MDIForm1.CommonDialog1.FileName
        savepictureRoutine = True

    Else

    End If

    Exit Function
errDialog:

End Function

Public Function max(a, b) As Variant
        If a > b Then
        max = a
        Else
        max = b
        End If
        
End Function
