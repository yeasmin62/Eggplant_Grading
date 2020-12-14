VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "CVS"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   939
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "3.Object Detection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   21
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   12000
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   525
      Left            =   13560
      TabIndex        =   18
      Text            =   "%"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   13560
      TabIndex        =   17
      Text            =   "Pixels"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   13560
      TabIndex        =   16
      Text            =   "Pixels"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5. Disease Detection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6840
      TabIndex        =   15
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   12000
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6.Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   11
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   12000
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   12000
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4.Binary Transformation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2.Background elimination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      FillColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   5400
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1.Noise Reduction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   5040
      Width           =   2535
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   600
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Disease 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Percentage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   9720
      TabIndex        =   14
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Eggplant fruit Grading and Disease Detection"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   13
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Volume of defected area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   9720
      TabIndex        =   10
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Total Volume "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   480
      Left            =   9720
      TabIndex        =   7
      Top             =   1800
      Width           =   2265
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Image After operations"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Original Image "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpenImage 
         Caption         =   "&Open Image"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim red(500, 500), green(500, 500), blue(500, 500), gray(500, 500) As Long
Dim red1(500, 500), green1(500, 500), blue1(500, 500), arr1(500, 500) As Long
Dim avgRed, avgGreen, avgBlue As Long
Dim pixel, pixel1 As Long
Dim p, q As Integer

Sub GetData()

    Dim pointRGB As Long
    Dim m As Integer
    Dim n As Integer
    Dim picwidth As Integer
    Dim picHeight As Integer
    picwidth = frmMain.Picture1.ScaleWidth - 1
    picHeight = frmMain.Picture1.ScaleHeight - 1
    
    'picwidth = ImageCapture.Picture3.ScaleWidth - 1
    'picheight = ImageCapture.Picture3.ScaleHeight - 1
    For m = 0 To picwidth
        For n = 0 To picHeight
            pointRGB = frmMain.Picture1.Point(m, n)
            red(m, n) = pointRGB Mod 256
            green(m, n) = ((pointRGB And &HFF00FF00) / 256&)
            blue(m, n) = (pointRGB And &HFF0000) / (256& * 256&)
        Next n
    Next m
End Sub

'The histogram information is displayed on a separate form (launch from button)
Private Sub cmdDispHistogram_Click()
    'frmHistogram.Show
End Sub


Private Sub Command1_Click()

Picture1.AutoSize = True
Dim a(300, 256) As Long
        Dim B(300, 256) As Long
        Dim ii, jj As Integer
        Dim k, m, L As Integer
        Dim total, avg As Long
        Dim sum(16) As Long
        Dim i, j As Integer
        Dim t, med As Long
        
        PicMain.ScaleMode = 3
        Picture1.ScaleMode = 3
        
        
        For i = 0 To 300
            For j = 0 To 256
                a(i, j) = PicMain.Point(i, j)
             Next j
             Next i
            
            For i = 0 To 298
                For j = 0 To 254
                    total = 0
                    L = 0
                    For k = i To i + 2
                        For m = j To j + 2
                            L = L + 1
                            sum(L) = a(k, m)
                            total = total + sum(L)
                        Next m
                    Next k
        
                    For ii = 1 To L - 1
                        For jj = ii + 1 To L
                            If (sum(ii) > sum(jj)) Then
                                t = sum(ii)
                                sum(ii) = sum(jj)
                                sum(jj) = t
                            End If
                        Next jj
                    Next ii
                    med = sum(5)
        
                    For k = i To i + 2
                        For m = j To j + 2
                            If (a(k, m) < med) Then
                                B(k, m) = med
                            Else: B(k, m) = a(i, j)
                            End If
                        Next m
                    Next k
                Next j
            Next i
            
            For i = 1 To 300
                For j = 1 To 256
                    Picture1.PSet (i, j), B(i, j)
                Next j
            Next i
            Picture1.AutoSize = True
            Label2.Caption = "After Noise Reduction"
        
           'MsgBox ("Done")
           For i = 0 To 256 'width
       For j = 0 To 256  'height
                      
            If Picture1.Point(i, j) > RGB(0, 0, 0) Then
      
                 'Picture4.PSet (i, j), a(i, j)
                 pixel = pixel + 1
       '         End If
            
            End If
                
        Next j
     Next i
   
    'Print "Number of pixel of original image is = "; pixel
    Text1.Text = pixel

End Sub


           

Private Sub Command2_Click()
Call GetData   '** calling GetData function**
    Picture1.ScaleMode = 3
    'Picture3.ScaleMode = 3
         '** convert the psicture to gray scale**
     Dim a, B As Integer
     
     For a = 0 To Picture1.ScaleWidth - 1
       For B = 0 To Picture1.ScaleHeight - 1
            red(a, B) = (red(a, B) + green(a, B) + blue(a, B)) / 3
            green(a, B) = red(a, B)
            blue(a, B) = red(a, B)
            Picture1.PSet (a, B), RGB(red(a, B), green(a, B), blue(a, B))
        Next
       ' ProgressBar1.Value = i * 100 / (Picture1.ScaleWidth - 1)
        'StatusBar1.Panels(1).Text = Str(Int(i * 100 / (Picture1.ScaleWidth - 1))) + "%"
        DoEvents
    Next
End Sub


Private Sub Command3_Click()
Call GetData '** calling GetData function**
Dim TV As Integer
Dim a As Integer
Dim B As Integer
Dim CheckBlack As Integer
    TV = 120 ' TV for threshold value to remove noise
 
 ''** Removin Background noise of the image**
    For a = 2 To Picture1.ScaleWidth - 1
        For B = 2 To Picture1.ScaleHeight - 1

            CheckBlack = 0

            If red(a - 1, B - 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a, B - 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a + 1, B - 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a + 1, B) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a + 1, B + 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a, B + 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a - 1, B + 1) <= TV Then
                CheckBlack = CheckBlack + 1
            End If
            If red(a - 1, B) <= TV Then
                CheckBlack = CheckBlack + 1
            End If

            If CheckBlack >= 5 Then
                red(a, B) = 0
                green(a, B) = 0
                blue(a, B) = 0

                red1(a, B) = 0
                green1(a, B) = 0
                blue1(a, B) = 0
            End If
            Picture1.PSet (a, B), RGB(red(a, B), green(a, B), blue(a, B))
        Next B
    Next a
    Label2.Caption = "After Background elimination"

End Sub

Private Sub Command4_Click()
    Call GetData        '** calling GetData function**
    Picture1.ScaleMode = 3
    'Picture6.ScaleMode = 3
    Dim i, j As Integer
    Dim a, B As Integer
    Dim pixel As Long
    'Code for scanning the image for Object Detection start from the left-Top Corner
    For a = 0 To 255
      For B = 0 To 255
            If red(a, B) > 115 Then 'green(a, b) <= 100 || blue(a, b) <= 100 Then
               Picture1.PSet (a, B), RGB(255, 255, 255)
               
            End If
            
            If red(a, B) <= 115 Then  'Or green(a, b) <= 100 Or blue(a, b) <= 100 Then
               Picture1.PSet (a, B), RGB(0, 0, 0)
               
            End If
                       
            DoEvents
        Next B
    Next a
    
    'End of the Code for scanning line start from the left top corner
    
    For i = 0 To 256 'width
       For j = 0 To 256  'height
                      
            If Picture1.Point(i, j) > RGB(0, 0, 0) Then
      
                 'Picture4.PSet (i, j), a(i, j)
                 pixel1 = pixel1 + 1
       '         End If
            
            End If
                
        Next j
     Next i
     Label2.Caption = "After Binary Transformation"
    Text2.Text = pixel1

End Sub



Private Sub Command5_Click()
p = pixel1 / pixel
q = p * 100
Text3.Text = q
If q <= 2 Then
   Text7.Text = "No Disease"
 End If
If q > 2 Then
   Text7.Text = "Phomopsis Blight"
 End If
     End Sub










Private Sub Command6_Click()

If q <= 2 Then
     MsgBox ("no defect")
End If
If q <= 25 Then
     MsgBox ("Partially defected")
 End If
 If q > 25 And q <= 50 Then
     MsgBox ("Moderately defected")
 End If
 If q > 50 Then
     MsgBox ("Unhealthy")
 End If
 
End Sub

Private Sub Command7_Click()
'ComputeTimeStart = Now
    Call GetData        '** calling GetData function**
    
    
    Dim a, B As Integer
    'Code for scanning the image for Object Detection start from the left-Top Corner
    For a = 0 To Picture1.ScaleWidth - 1
      For B = 0 To Picture1.ScaleHeight - 1
            If red(a, B) <= 160 Then 'Or green(a, b) <= 100 Or blue(a, b) <= 100 Then
               Picture1.PSet (a, B), RGB(0, 0, 0)
            End If
            DoEvents
        Next B
    Next a
    Label2.Caption = "After Object Detection"
    
End Sub

'Automatically initialized
Private Sub Form_Load()
    'Upon loading the form, reset two variables:
    'Luminance is the default histogram source
    lastHistSource = DRAWMETHOD_LUMINANCE
    'Line graph is the default drawing option
    lastHistMethod = DRAWMETHOD_BARS
    
    'frmHistogram.Show 0, Me
End Sub


'Launched from the menu
Private Sub MnuOpenImage_Click()
    'Common dialog interface
    Dim CC As cCommonDialog
    'String returned from the common dialog wrapper
    Dim sFile As String
    Set CC = New cCommonDialog
    'This string contains the filters for loading different kinds of images.  Using
    'this option correctly makes our common dialog box a lot nicer to use.
    Dim cdfStr As String
    cdfStr = "All Compatible Graphics|*.bmp;*.jpg;*.jpeg;*.gif;*.wmf;*.emf;*.ico;*.dib;*.rle|"
    cdfStr = cdfStr & "BMP - Windows Bitmaps only (non-OS2)|*.bmp|DIB - Windows DIBs only (non-OS2)|*.dib|EMF - Enhanced Meta File|*.emf|GIF - Compuserve|*.gif|ICO - Windows Icon|*.ico|JPG - JPEG - JFIF Compliant|*.jpg;*.jpeg|RLE - Windows only (non-Compuserve)|*.rle|WMF - Windows Meta File|*.wmf|All files|*.*"
    'If cancel isn't selected, load a picture from the user-specified file
    If CC.VBGetOpenFileName(sFile, , , , , True, cdfStr, 1, , "Open an image", , frmMain.hWnd, 0) Then
        PicMain.Picture = LoadPicture(sFile)
    End If
    
    'Redraw the histogram
    'frmHistogram.GenerateHistogram

End Sub

