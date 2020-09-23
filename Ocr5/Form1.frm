VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCR5 @ 2004 BY YOVASXP"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Height          =   2790
      Left            =   6120
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.HScrollBar hsTol 
      Height          =   375
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   5
      Top             =   3000
      Value           =   70
      Width           =   2895
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   15
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   113
      MousePointer    =   2  'Cross
      ScaleHeight     =   3000
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdTeach 
      Caption         =   "TEACH"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "LOAD FROM DATA FILE"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "RECOGNIZE"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblLearned 
      Caption         =   "LEARNED : #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOLERANCE : 70%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'THIS MODEL USES A DISPERSED MATRIX IN A VECTOR AND IT CONTAINS THE BLACK POINTS IN PICTURE CONTROL
'WITH THIS MODEL YOU ONLY SAVE THE BLACK POINTS AND THAT MEANS A MEMORY OPTIMIZATION.

'THE FIRST REGISTER IN THE VECTOR CONTAINS INFORMATION:
'ROW = CHARACTER IN ASCII.
'COL = DIMENSION OF VECTOR.

'WITH THIS YOU ONLY SAVE THE COORDS OF BLACK POINTS
Private Type reg
    row As Byte
    col As Byte
End Type

Dim fileName As String 'FILE WITH DATA
Dim mouseDown As Boolean 'IS MOUSE DOWN
Dim tx1 As Long 'FIRST VERTICAL TANGENT AT CHARACTER'S DRAWING
Dim tx2 As Long 'SECOND VERTICAL TANGENT AT CHARACTER'S DRAWING
Dim ty1 As Long 'FIRST HORIZONTAL TANGENT AT CHARACTER'S DRAWING
Dim ty2 As Long 'SECOND HORIZONTAL TANGENT AT CHARACTER'S DRAWING
Dim dm() As reg 'VECTOR THAT CONTAINS THE CHARACTER'S DRAWING BLACK POINTS
Dim dmb() As reg 'VECTOR THAT CONTAINS ALL CHARACTERS LOADED FROM FILE

Private Sub cmdClear_Click()
    pic.Cls
    cmdTeach.Enabled = True
End Sub

Private Sub cmdRec_Click()
    Call recognizeChar
End Sub

Private Sub cmdLoad_Click()
    
    If Not loadFile Then
        Exit Sub
    End If
    
    cmdRec.Enabled = True
    cmdLoad.Enabled = False
End Sub

Private Sub cmdTeach_Click()
        
    If Not saveChar Then
        Exit Sub
    End If
        
    If Not loadFile Then
        Exit Sub
    End If
    
    cmdRec.Enabled = True
    cmdTeach.Enabled = False
    cmdLoad.Enabled = False
End Sub

Private Sub Form_Load()
    fileName = App.Path & "\data.txt"
End Sub

Private Sub hsTol_Change()
    lblTol.Caption = "TOLERANCE : " & CStr(hsTol.Value) & "%"
End Sub

Private Sub hsTol_Scroll()
    Call hsTol_Change
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown = True
    pic.Line (X, Y)-(X + 1, Y + 1), vbBlack
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mouseDown And Button = 1 Then
        pic.Line (X, Y)-(X + 1, Y + 1), vbBlack
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown = False
End Sub

Private Sub img2dm(p As PictureBox)
Dim sectorWidth As Long
Dim sectorHeight As Long
Dim ii As Long
Dim jj As Long
Dim ddm As Byte 'DIMENSION OF VECTOR
    
    sectorWidth = (tx2 - tx1) / 10 + 1
    sectorHeight = (ty2 - ty1) / 10 + 1
    
    ddm = 1
    
    For ii = tx1 To (tx2 - sectorWidth) Step sectorWidth
        For jj = ty1 To (ty2 - sectorHeight) Step sectorHeight
            If p.Point(ii, jj) = vbBlack Then
                
                ReDim Preserve dm(ddm)
                
                dm(ddm).row = (ii - tx1) \ sectorWidth
                dm(ddm).col = (jj - ty1) \ sectorHeight
                
                ddm = ddm + 1
            End If
        Next jj
    Next ii
    
    dm(0).col = ddm - 1
End Sub

Private Function tangents(p As PictureBox) As Boolean
Dim ii As Long
Dim jj As Long
Dim bFirstStep As Boolean 'FLAG TO KNOW IF TANGENT POINTS ARE INITIALIZED
    
    For jj = 0 To p.ScaleHeight Step 100
        For ii = 0 To p.ScaleWidth Step 100
            If p.Point(ii, jj) = vbBlack Then
                If bFirstStep Then
                    If ii < tx1 Then
                        tx1 = ii
                    End If
                    
                    If ii > tx2 Then
                        tx2 = ii
                    End If
                    
                    If jj < ty1 Then
                        ty1 = jj
                    End If
                    
                    If jj > ty2 Then
                        ty2 = jj
                    End If
                Else
                    tx1 = ii
                    tx2 = ii
                    ty1 = jj
                    ty2 = jj
                    
                    bFirstStep = True
                End If
            End If
        Next ii
    Next jj

    If ((tx2 - tx1) <= 500 And (ty2 - ty1) <= 500) Or ((tx2 - tx1) * (ty2 - ty1) = 0) Then
        MsgBox "You must draw something valid!", vbExclamation
    Else
        tangents = True
    End If
End Function

Private Function loadFile() As Boolean
Dim i As Long
Dim ff As Integer
Dim lFile As Long
    
    On Error GoTo err_loadFile
    
    If Len(Dir$(fileName)) = 0 Then
        MsgBox "File not found!", vbExclamation
        Exit Function
    End If
    
    lFile = FileLen(fileName)
    
    If lFile Mod 2 = 1 Then
        MsgBox "File corrupted!", vbExclamation
        Exit Function
    End If
    
    ReDim dmb(lFile / 2 - 1)
    
    ff = FreeFile
    
    Open fileName For Binary Access Read As #ff
        
    Get #ff, , dmb
        
    Close #ff
    
    Call loadLearnedChars
    
    loadFile = True
    
    Exit Function
    
err_loadFile:
    MsgBox Err.Description, vbExclamation
    Close
End Function

Private Function saveChar() As Boolean
Dim i As Byte
Dim tmpStr As String
Dim lFile As Long
Dim ff As Integer
    
    On Error GoTo err_saveChar
    
    If Not tangents(pic) Then
        Exit Function
    End If
        
    Call img2dm(pic)
        
    tmpStr = InputBox("Enter the character that is drawn:")
    
    If StrPtr(tmpStr) = 0 Then
        Exit Function
    End If
    
    If Len(Trim$(tmpStr)) <> 1 Then
        MsgBox "Please enter one character!", vbExclamation
        Exit Function
    End If
    
    dm(0).row = Asc(tmpStr)
    
    If Len(Dir$(fileName)) > 0 Then
        lFile = FileLen(fileName)
    End If
    
    ff = FreeFile
    
    Open fileName For Binary Access Write As #ff
            
    Put #ff, lFile + 1, dm
        
    Close #ff
        
    saveChar = True
    
    Exit Function
        
err_saveChar:
    MsgBox Err.Description, vbExclamation
    Close
End Function

Private Sub recognizeChar()
Dim m_char As String * 1
Dim char As String * 1
Dim m_aprox As Integer
Dim aprox As Integer
Dim dim_a As Integer
Dim dim_b As Integer
Dim i As Long
Dim j As Byte
Dim n As Byte

    If Not tangents(pic) Then
        Exit Sub
    End If
    
    Call img2dm(pic)
    
    dim_b = dm(0).col
    
    Do
        char = Chr$(dmb(i).row)
        dim_a = dmb(i).col
        aprox = 0
        n = 1
        
        For i = i + 1 To i + dim_a
            
            For j = n To dim_b
                
                If dm(j).row = dmb(i).row And dm(j).col = dmb(i).col Then
                    aprox = aprox + 1
                    n = j + 1
                    Exit For
                End If
                                
            Next j
            
        Next i
        
        aprox = ((100 - (dim_a + dim_b - aprox)) + aprox) * 2 - 100
                                
        If m_aprox < aprox Then
            m_char = char
            m_aprox = aprox
            
            If m_aprox = 100 Then
                Exit Do
            End If
        End If
        
        Me.Caption = "COMPARING CHARACTERS : " & CStr(CInt(((i - 1) / UBound(dmb)) * 100)) & "%"
        
    Loop While i < UBound(dmb)
    
    If m_aprox >= hsTol.Value Then
        MsgBox "Match with character : " & m_char & vbCrLf & "With an aprox. of : " & CStr(m_aprox) & "%"
    Else
        MsgBox "Don't match with any of the characters in the data file!"
    End If
End Sub

Public Sub loadLearnedChars()
Dim i As Long
Dim tmpStr As String
Dim tmpChar As String * 1

    lstChars.Clear
    DoEvents
    
    Do
        tmpChar = Chr$(dmb(i).row)
        
        If InStr(1, tmpStr, tmpChar, vbTextCompare) = 0 Then
            lstChars.AddItem Chr$(dmb(i).row)
            tmpStr = tmpStr & tmpChar
        End If
        
        i = i + dmb(i).col + 1
        
    Loop While i < UBound(dmb)
    
    lblLearned.Caption = "LEARNED : " & CStr(lstChars.ListCount)
End Sub
