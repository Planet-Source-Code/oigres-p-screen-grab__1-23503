VERSION 5.00
Begin VB.Form frmChild 
   Caption         =   "frmChild"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_IsDirty As Boolean

Property Get IsDirty() As Boolean

    IsDirty = m_IsDirty

End Property

Property Let IsDirty(v As Boolean)

    m_IsDirty = v

End Property

Private Sub Form_Activate()
    If debugme = True Then MsgBox "activate child"
    'adjust menu to enable save or not
    MDIForm1.mnuSave.Enabled = IsDirty
End Sub

Private Sub Form_Load()
    m_IsDirty = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim answer As Integer
    ' process child forms; save picture grab if want to

    If IsDirty Then

        answer = MsgBox("Do you wish to save screen grab?", vbYesNoCancel, "Screen Grab")

        Select Case answer
            Case vbYes

                If debugme = True Then MsgBox "you chose yes"

                'saveroutine
                If savepictureRoutine = True Then

                Else
                    Cancel = True
                End If

                If debugme = True Then MsgBox Me.Picture.Width & ":" & Me.Picture.Height
                If debugme = True Then MsgBox Me.Picture.Width / Screen.TwipsPerPixelX & ":" & Me.Picture.Height / Screen.TwipsPerPixelY

                Dim pwidth, pheight

                With Me
                    pwidth = CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels))
                    pheight = CInt(.ScaleY(.Picture.Height, vbHimetric, _
                       vbPixels))

                    If debugme = True Then MsgBox pwidth & ":" & pheight
                End With

            Case vbNo

                If debugme = True Then MsgBox "you chose no"

            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

Private Sub Picture1_Click()

    Me.Picture1.Picture = Me.Picture1.Image
    MsgBox CInt(Me.ScaleX(Me.Picture1.Picture.Width, vbHimetric, vbPixels))

End Sub

Private Function myfunc() As Boolean

    '
    MsgBox "my function calling"

End Function

