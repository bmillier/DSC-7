VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form DSC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Differential Scanning Calorimeter V7  Feb 12 2016 CIC"
   ClientHeight    =   9180
   ClientLeft      =   885
   ClientTop       =   1230
   ClientWidth     =   11280
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   9480
      Picture         =   "DSC.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   3360
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   39
      Top             =   8760
      Width           =   4215
   End
   Begin VB.TextBox StartTempwind 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      TabIndex        =   35
      Top             =   5040
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   840
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Store Calibration to Instrument"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   33
      Top             =   8160
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Vminlabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Height          =   4095
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   1695
      Begin VB.TextBox Heatsinktemptext 
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox referencetempwind 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox sampletempwind 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox cursYwind 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox cursXwind 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "HeatSink Temp"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1680
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Reference temp"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Sample Temp"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "App. Heat"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Temp."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Redrawbut 
      Caption         =   "ReDraw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox VFSwindow 
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.VScrollBar VertZoomScroll 
      Height          =   975
      Left            =   360
      Max             =   1
      Min             =   50
      TabIndex        =   17
      Top             =   1680
      Value           =   10
      Width           =   375
   End
   Begin VB.TextBox VFSlabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "DSC.frx":0B87
      Top             =   840
      Width           =   135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox FinalTempWind 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton stopbut 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Startbut 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Control Panel"
      Height          =   2535
      Left            =   1800
      TabIndex        =   1
      Top             =   5520
      Width           =   9015
      Begin VB.Frame Frame6 
         Caption         =   "Peltier Cooler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   5040
         TabIndex        =   40
         Top             =   840
         Width           =   3855
         Begin VB.CommandButton CoolerOffButton 
            Caption         =   "Cooler Off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton CoolerOnButton 
            Caption         =   "Cooler on"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Calibration Run"
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1335
         Begin VB.CheckBox CalibrateButton 
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox StatusWind 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Acquisition Mode"
         Height          =   615
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   3255
         Begin VB.OptionButton Rampupdownoption 
            Caption         =   "Ramp Up/Down"
            Height          =   255
            Left            =   1560
            TabIndex        =   37
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Rampupoption 
            Caption         =   "Ramp Up"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ramp Rate"
         Height          =   615
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   3735
         Begin VB.OptionButton Option3 
            Caption         =   "20 C/min"
            Height          =   255
            Left            =   2520
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "10 C/min"
            Height          =   255
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "5 C /min"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   6480
         Max             =   300
         Min             =   51
         TabIndex        =   4
         Top             =   480
         Value           =   250
         Width           =   1815
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   300
         Min             =   20
         TabIndex        =   2
         Top             =   480
         Value           =   50
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
         Caption         =   "STATUS"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Final Temp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Starting Temp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   10
      DTREnable       =   -1  'True
      InputLen        =   1
      BaudRate        =   57600
   End
   Begin VB.PictureBox Graph1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   2040
      ScaleHeight     =   4785
      ScaleWidth      =   8745
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label12 
      Caption         =   "Set for COM10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   47
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      Caption         =   "compiled for use with Excel 2002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   32
      Top             =   8760
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vertical Zoom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sample Temperature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Menu filemenu 
      Caption         =   "File"
      Begin VB.Menu loadsamplerun 
         Caption         =   "Load last run from disk"
      End
      Begin VB.Menu loadfromdiskmenu 
         Caption         =   "Load selected run from disk"
      End
      Begin VB.Menu Savetodiskmenu 
         Caption         =   "Save current run to disk"
      End
      Begin VB.Menu exitprogrammenu 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu Spreadsheetbut 
      Caption         =   "SpreadSheet Functions"
      Begin VB.Menu saveExcelSpreadsheetmenu 
         Caption         =   "Save Current Run  to Excel Spreadsheet"
      End
   End
   Begin VB.Menu processbut 
      Caption         =   "Data Processing"
      Begin VB.Menu OnsetTempmenu 
         Caption         =   "Define Onset Temperature Line"
      End
      Begin VB.Menu drawpeakbaselinemenu 
         Caption         =   "Define Peak Baseline"
      End
      Begin VB.Menu Calculationmenu 
         Caption         =   "Perform UP RAMP Calculations (do onset and Baseline first)"
      End
      Begin VB.Menu calculation2menu 
         Caption         =   "Perform DOWN RAMP Calculations (do onset and baseline first)"
      End
   End
   Begin VB.Menu CalibButton 
      Caption         =   " Calibration (Service Technician ONLY!)"
      Begin VB.Menu SlopeCalibrationBut 
         Caption         =   "Slope Calibration"
      End
      Begin VB.Menu menuSlopeOne 
         Caption         =   "Set Slope Correction to 1.00"
      End
   End
End
Attribute VB_Name = "DSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sample(1 To 700) As Integer
Dim Reference(1 To 700) As Integer
Dim y(1 To 350) As Double ' used by Curvefit routine
Dim x(1 To 350) As Double ' used by curvefit routine
Dim BaselineCollected As Boolean
Dim start As Integer
Dim final As Integer
Dim ramp As Integer
Dim size As Integer
Dim mouseClickCount As Integer
Dim mousemode As Integer
Dim mouseX(0 To 5) As Integer
Dim mouseY(0 To 5) As Integer
Dim x1last, x2last, y1last, y2last As Single
Dim blcoef1, blcoef2, blcoef3 As Single
Dim samplecoef1, samplecoef2, samplecoef3 As Single
Dim referencecoef1, referencecoef2, referencecoef3 As Single
Dim timer1flag As Boolean
Dim timer2flag As Boolean
Dim calibrateoption As Boolean
Dim Rampmode As Integer
Dim fastcool As Boolean
Dim ComPortNumber As Integer











Private Sub calculation2menu_Click()
'DOWN RAMP CALCULATIONS
  ' this calculates the onset temperature and area under the
  'enthalpy peak, from
  ' 1) line drawn by user thru steepest rising edge of peak
  ' 2) line drawn by user defining the baseline under the peak
  ' onset temp line definitions
  ' mouseY(2) is Y co-ord of first point placed by user
  ' mouseX(2) is X co-ord of first point placed by user
  ' mouseY(3) is Y co-ord of second point placed by user
  ' mouseX(3) is X co-ord of second point placed by user
  ' peak baseline definitions
  ' mouseY(4) is Y co-ord of first point placed by user
  ' mouseX(4) is X co-ord of first point placed by user
  ' mouseY(5) is Y co-ord of second point placed by user
  ' mouseX(5) is X co-ord of second point placed by user
  
  Open "calcs" For Output As #1
  Dim X1, Y1, X2, Y2, V1, V2, U1, U2 As Integer
  Dim b1, b2, a1, a2 As Single
  Dim xi, yi As Single
    
  Dim fs, ls, ns As Integer
  Dim bl1, bl2 As Integer
  Dim dy As Single
  Dim integral As Single
  ' swap onset temp co-ordinates, if user drew line from bottom  to top
  If mouseY(2) < mouseY(3) Then
   temp = mouseY(2)
   mouseY(2) = mouseY(3)
   mouseY(3) = temp
   temp = mouseX(2)
   mouseX(2) = mouseX(3)
   mouseX(3) = temp
 End If
  
  Text1 = Str(mouseX(2)) + " " + Str(mouseY(2)) + " " + Str(mouseX(3)) + " " + Str(mouseY(3))
  StatusWind = Str(mouseX(4)) + " " + Str(mouseY(4)) + " " + Str(mouseX(5)) + " " + Str(mouseY(5))
  
  ' define the onset temp line as X1Y1 X2Y2
  X1 = mouseX(2)
  Y1 = mouseY(2)
  X2 = mouseX(3)
  Y2 = mouseY(3)
  ' define the peak baseline as U1V1 U2V2
  U1 = mouseX(4)
  V1 = mouseY(4)
  U2 = mouseX(5)
  V2 = mouseY(5)
  b1 = (Y2 - Y1) / (X2 - X1) ' slope of onset temp line
  b2 = (V2 - V1) / (U2 - U1) ' slope of baseline
  a1 = Y1 - (b1 * X1)   ' intercept of onset temp line
  a2 = V1 - (b2 * U1)   ' intercept of baseline
  xi = -(a1 - a2) / (b1 - b2)
  yi = a1 + (b1 * xi)
  
  If ((X1 - xi) * (xi - -X2) >= 0) And ((U1 - xi) * (xi - U2) >= 0) And ((Y1 - yi) * (yi - Y2) >= 0) And ((V1 - yi) * (yi - V2) >= 0) Then
    onsettempwind = Str(xi)
  Else
   MsgBox ("your onset temp line does not cross baseline!")
   onsettempwind = "???"
  End If
  If mouseX(4) > mouseX(5) Then
    temp = mouseX(5)
    mouseX(5) = mouseX(4)
    mouseX(4) = temp
  End If
  fs = 1 + mouseX(4) - start
  ls = 1 + mouseX(5) - start
  ns = (ls - fs) - 1
  bl1 = Sample(size - fs) - Reference(size - fs)
  bl2 = Sample(size - ls) - Reference(size - ls)
  dy = (bl2 - bl1) / ns
  integral = 0
  Max = 0
   Graph1.ForeColor = RGB(255, 255, 255)
   bl = bl1
   
   Text3 = ls
  For i = (fs) To (ls)
    s = Sample(size - i) - Reference(size - i) - bl
    integral = integral + s
    Graph1.Line (start + i, bl)-(start + i, Sample(size - i) - Reference(size - i))
    Print #1, start + i; bl, Sample(size - i) - Reference(size - i)
    bl = bl + dy
  Next
  Close #1
  integral = Abs(integral)
  a$ = Format$(integral, "#.####e+##")
  MsgBox ("ONSET Temperature = " & Str(xi) & "   ENTHALPY = " & a$)
End Sub

Private Sub Calculationmenu_Click()
' UP RAMP CALCULATIONS
  ' this calculates the onset temperature and area under the
  'enthalpy peak, from
  ' 1) line drawn by user thru steepest rising edge of peak
  ' 2) line drawn by user defining the baseline under the peak
  ' onset temp line definitions
  ' mouseY(2) is Y co-ord of first point placed by user
  ' mouseX(2) is X co-ord of first point placed by user
  ' mouseY(3) is Y co-ord of second point placed by user
  ' mouseX(3) is X co-ord of second point placed by user
  ' peak baseline definitions
  ' mouseY(4) is Y co-ord of first point placed by user
  ' mouseX(4) is X co-ord of first point placed by user
  ' mouseY(5) is Y co-ord of second point placed by user
  ' mouseX(5) is X co-ord of second point placed by user
  Dim X1, Y1, X2, Y2, V1, V2, U1, U2 As Integer
  Dim b1, b2, a1, a2 As Single
  Dim xi, yi As Single
    
  Dim fs, ls, ns As Integer
  Dim bl1, bl2 As Integer
  Dim dy As Single
  Dim integral As Single
  ' swap onset temp co-ordinates, if user drew line from bottom  to top
  If mouseY(2) < mouseY(3) Then
   temp = mouseY(2)
   mouseY(2) = mouseY(3)
   mouseY(3) = temp
   temp = mouseX(2)
   mouseX(2) = mouseX(3)
   mouseX(3) = temp
 End If
 
  
  Text1 = Str(mouseX(2)) + " " + Str(mouseY(2)) + " " + Str(mouseX(3)) + " " + Str(mouseY(3))
  StatusWind = Str(mouseX(4)) + " " + Str(mouseY(4)) + " " + Str(mouseX(5)) + " " + Str(mouseY(5))
  
  ' define the onset temp line as X1Y1 X2Y2
    X1 = mouseX(2)
  Y1 = mouseY(2)
  X2 = mouseX(3)
  Y2 = mouseY(3)
  ' define the peak baseline as U1V1 U2V2
  U1 = mouseX(4)
  V1 = mouseY(4)
  U2 = mouseX(5)
  V2 = mouseY(5)
  b1 = (Y2 - Y1) / (X2 - X1) ' slope of onset temp line
  b2 = (V2 - V1) / (U2 - U1) ' slope of baseline
  a1 = Y1 - (b1 * X1)   ' intercept of onset temp line
  a2 = V1 - (b2 * U1)   ' intercept of baseline
  xi = -(a1 - a2) / (b1 - b2)
  yi = a1 + (b1 * xi)
  
  If ((X1 - xi) * (xi - -X2) >= 0) And ((U1 - xi) * (xi - U2) >= 0) And ((Y1 - yi) * (yi - Y2) >= 0) And ((V1 - yi) * (yi - V2) >= 0) Then
    onsettempwind = Str(xi)
  Else
   MsgBox ("your onset temp line does not cross baseline!")
   onsettempwind = "???"
  End If
  If mouseX(4) > mouseX(5) Then
    temp = mouseX(5)
    mouseX(5) = mouseX(4)
    mouseX(4) = temp
  End If
  
  fs = 1 + mouseX(4) - start
  ls = 1 + mouseX(5) - start
  ns = (ls - fs) - 1
  bl1 = Sample(fs) - Reference(fs)
  bl2 = Sample(ls) - Reference(ls)
  dy = (bl2 - bl1) / ns
  integral = 0
  Max = 0
   Graph1.ForeColor = RGB(255, 255, 255)
   bl = bl1
   Text3 = ls
  For i = (fs) To (ls)
      
    s = Sample(i) - Reference(i) - bl
    integral = integral + s
    Graph1.Line (i + start, bl)-(i + start, Sample(i) - Reference(i))
    bl = bl + dy
  Next
  a$ = Format$(integral, "#.####e+##")
    MsgBox ("ONSET Temperature = " & Str(xi) & "   ENTHALPY = " & a$)
End Sub
Sub calibration2menu()

End Sub
Private Sub CalibrateOption_Click()
 VertZoomScroll.Enabled = False
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
Dim cmdstring As String
Dim tempstr1, tempstr2 As String
Dim fsize As Integer
Open "DSCCurvefit" For Output As #1
StatusWind = "Saving curvefit parameters to instrument"
If Rampmode = 0 Then
 fsize = size
Else
 fsize = size \ 2
End If

' first curve fit the Sample pan data- upward scan

For i = 1 To fsize
  y(i) = Sample(i)
  x(i) = start + i
Next i
Call Curvfit(fsize)
samplecoef1 = blcoef1
samplecoef2 = blcoef2
samplecoef3 = blcoef3

' then curve fit the Reference pan data - upward scan
For i = 1 To fsize
  y(i) = Reference(i)
  x(i) = start + i
Next i
Call Curvfit(fsize)
referencecoef1 = blcoef1
referencecoef2 = blcoef2
referencecoef3 = blcoef3

' Do the download of curvefit parameters for the up ramp ( A command)
cmdstring = "A"
Call SendCmd(cmdstring)

tempstr1 = Str(samplecoef1)
tempstr2 = Format(tempstr1, "00000.00")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

tempstr1 = Str(samplecoef2)
tempstr2 = Format(tempstr1, "0000.000")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

tempstr1 = Str(samplecoef3)
tempstr2 = Format(tempstr1, "000.0000")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

tempstr1 = Str(referencecoef1)
tempstr2 = Format(tempstr1, "00000.00")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

tempstr1 = Str(referencecoef2)
tempstr2 = Format(tempstr1, "0000.000")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

tempstr1 = Str(referencecoef3)
tempstr2 = Format(tempstr1, "000.0000")
Print #1, tempstr2
cmdstring = tempstr2
Call Slowsend(cmdstring)

If Rampmode = 2 Then

   ' then curve fit the Sample pan data- downward scan
   For i = 1 To fsize
     y(i) = Sample(i + fsize)
     x(i) = final - i
   Next i
   Call Curvfit(fsize)
   samplecoef4 = blcoef1
   samplecoef5 = blcoef2
   samplecoef6 = blcoef3

' then curve fit the Reference pan data - downward scan
   For i = 1 To fsize
     y(i) = Reference(i + fsize)
     x(i) = final - i
   Next i
   Call Curvfit(fsize)
   referencecoef4 = blcoef1
   referencecoef5 = blcoef2
   referencecoef6 = blcoef3

' Do the download of curvefit params for the down ramp (B command)
  cmdstring = "B"
  Call SendCmd(cmdstring)

  tempstr1 = Str(samplecoef4)
  tempstr2 = Format(tempstr1, "00000.00")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)

  tempstr1 = Str(samplecoef5)
  tempstr2 = Format(tempstr1, "0000.000")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)

  tempstr1 = Str(samplecoef6)
  tempstr2 = Format(tempstr1, "000.0000")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)

  tempstr1 = Str(referencecoef4)
  tempstr2 = Format(tempstr1, "00000.00")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)

  tempstr1 = Str(referencecoef5)
  tempstr2 = Format(tempstr1, "0000.000")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)

  tempstr1 = Str(referencecoef6)
  tempstr2 = Format(tempstr1, "000.0000")
  Print #1, tempstr2
  cmdstring = tempstr2
  Call Slowsend(cmdstring)
 End If
Close #1
StatusWind = "Calibration parameters sent to instrument"
Command1.Visible = False
 Command1.Enabled = False

End Sub
Sub Slowsend(a$)
Timer1.Enabled = True
Timer1.Interval = 10
timer1flag = False
For i = 1 To Len(a$)
 Do
  DoEvents
 Loop Until timer1flag = True
 b$ = Mid$(a$, i, 1)
 MSComm1.Output = b$
 timer1flag = False
Next
b$ = Chr(13)
MSComm1.Output = b$
End Sub





Private Sub Command2_Click()
End Sub

Private Sub CoolerOffButton_Click()
 timerindex = 0
 fastcool = False
 Picture1.Visible = False
 PeltierWind = "OFF"
 SendCmd ("C")
 SendParam ("0")
 sampletempwind = ""
 referencetempwind = ""
 Heatsinktemptext = ""
 Timer2.Enabled = False
 CoolerOffButton.Enabled = False
 CoolerOnButton.Enabled = True
 
End Sub

Private Sub CoolerOnButton_Click()
 CoolerOnButton.Enabled = False
 CoolerOffButton.Enabled = True
 timerindex = 0
 fastcool = True
 Picture1.Visible = True
 PeltierWind = "ON"
 SendCmd ("C")
 SendParam ("1")
 Timer2.Interval = 2000
 Timer2.Enabled = True
End Sub

Private Sub drawpeakbaselinemenu_Click()
mousemode = 2
x1last = 0
x2last = 0
y1last = 0
y2last = 0
mouseClickCount = 0
Graph1.MousePointer = 2
Graph1.DrawMode = 1
Graph1.ForeColor = RGB(255, 255, 0)
End Sub

Private Sub exitprogrammenu_Click()
 End
End Sub

Private Sub Graph1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim mm As Integer
  If mousemode = 0 Then Exit Sub
  mm = mousemode * 2
  sn = 1 + (Int(x) - start)
  If sn > size Then sn = size
  If sn < 1 Then sn = 1
  mouseX(mm + mouseClickCount) = Int(x)
  mouseY(mm + mouseClickCount) = Int(y)
  Graph1.DrawWidth = 1
  If mouseClickCount = 0 Then
     Graph1.PSet (x1last, y1last)
     Graph1.PSet (x, y)
  End If
  mouseClickCount = mouseClickCount + 1

  If mouseClickCount = 2 Then mouseClickCount = 0
  'Graph1.DrawMode = 1
  Graph1.DrawWidth = 1
End Sub

Private Sub Graph1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim sn As Integer
   If (mousemode = 1) And (mouseClickCount = 1) Then
       Graph1.DrawMode = 7
      Graph1.Line (x1last, y1last)-(x2last, y2last)
      x1last = mouseX(2)
      y1last = mouseY(2)
      x2last = x
      y2last = y
      Graph1.Line (x1last, y1last)-(x2last, y2last)
      x1last = mouseX(2)
      y1last = mouseY(2)
      x2last = x
      y2last = y
   End If
     If (mousemode = 2) And (mouseClickCount = 1) Then
      Graph1.DrawMode = 7
      Graph1.Line (x1last, y1last)-(x2last, y2last)
      x1last = mouseX(4)
      y1last = mouseY(4)
      x2last = x
      y2last = y
      Graph1.Line (x1last, y1last)-(x2last, y2last)
      x1last = mouseX(4)
      y1last = mouseY(4)
      x2last = x
      y2last = y
  End If
 If BaselineCollected = False Then Exit Sub
 
 cursXwind = Int(x)
 sn = 1 + (Int(x) - start)
 If sn > size Then sn = size
 If sn < 1 Then sn = 1
 'If DataOption = True Then
 ' cursYwind = Sample(sn) - Reference(sn)
 'Else
 ' cursYwind = reference(sn)
 'End If
 
  End Sub



Private Sub Integrate_Click()
End Sub

Private Sub loadfromdiskmenu_Click()
 Dim fn, fn2 As String
 Dim a As Integer
   
   CommonDialog2.InitDir = "C:\my documents"
   CommonDialog2.FileName = "*.dss"
   CommonDialog2.Action = 1
   If CommonDialog2.FileName = "*.dss" Then Exit Sub
   fn = CommonDialog2.FileName
   a = InStr(1, fn, ".")
   If a > 0 Then
     fn = Left$(fn, a - 1)
   End If
   fn2 = fn & ".dss"
   Call LoadSample(fn2)
   StatusWind = "LOADED CHOSEN SAMPLE RUN TO DISK"
  
End Sub

Private Sub loadsamplerun_Click()
 LoadSample ("currun")

 End Sub

Private Sub menuSlopeOne_Click()
 Dim slope As Single
 Dim slopeStr As String * 15
 Dim cmdstring As String
 ans = InputBox("This should only be done to remove slope correction (for re-calibration)- Proceed?", "SLOPE CORRECTION =1", "No")
 If ans = "No" Then Exit Sub
 slope = 2#   ' add 1.0 workaround to correct arduino parsefloat bug where it fails on slope <1
 slopeStr = Str(slope)
 slopeStr = Left$(slopeStr, 6)
  Heatsinktemptext = slopeStr
  cmdstring = "T"
  Call SendCmd(cmdstring)
  cmdstring = slopeStr
  Call Slowsend(cmdstring)
  StatusWind = "Slope =1 sent to instrument"
End Sub

Private Sub OnsetTempmenu_Click()
 mousemode = 1
 mouseClickCount = 0
 x1last = 0
 x2last = 0
 y1last = 0
 y2last = 0
  Graph1.MousePointer = 2
  Graph1.ForeColor = RGB(255, 255, 0)
End Sub
Private Sub LoadSample(fn As String)
  Dim temp As Integer
  mousemode = 0
  Open fn For Input As #1
    Input #1, start
    Input #1, final
    Input #1, ramp
    Input #1, size
    Input #1, Rampmode
    For i = 1 To size
     Input #1, Sample(i), Reference(i)
    Next
  Close #1
  If Rampmode = 0 Then
   Rampupoption = True
   Rampupdownoption = False
  End If
  If Rampmode = 2 Then
   Rampupoption = False
   Rampupdownoption = True
  End If
  StartTempwind = Str(start)
  HScroll1.Value = start
  FinalTempWind = Str(final)
  HScroll2.Value = final
  If ramp = 5 Then Option1.Value = True
  If ramp = 10 Then Option2.Value = True
  If ramp = 20 Then Option3.Value = True
  StatusWind = "LAST RUN RELOADED FROM DISK"
  PlotData
  VertZoomScroll.Enabled = True
End Sub
Private Sub PeakIntegrate_Click()
  Dim fs, ls, ns As Integer
  Dim bl1, bl2 As Integer
  Dim dy As Single
  Dim integral As Single
  
  If mouseX(4) > mouseX(5) Then
    temp = mouseX(5)
    mouseX(5) = mouseX(4)
    mouseX(4) = temp
  End If
  
  fs = 1 + mouseX(4) - start
  ls = 1 + mouseX(5) - start
  ns = (ls - fs) - 1
  bl1 = Sample(fs) - Reference(fs)
  bl2 = Sample(ls) - Reference(ls)
  dy = (bl2 - bl1) / ns
  integral = 0
  Max = 0
   Graph1.ForeColor = RGB(0, 0, 0)
   bl = bl1
  For i = (fs) To (ls)
    s = Sample(i) - Reference(i) - bl
    integral = integral + s
    Graph1.Line (i + start - 1, bl)-(i + start - 1, Sample(i) - Reference(i))
    bl = bl + dy
  Next
  a$ = "Integral = " & Format$(integral, "#.#####e+##")
  MsgBox (a$)
  PlotData   ' replot data when message box is removed
  
End Sub

Private Sub Option4_Click()
 
End Sub

Private Sub Rampupdownoption_Click()
Rampmode = 2
End Sub

Private Sub Rampupoption_Click()
Rampmode = 0
End Sub

Private Sub Redrawbut_Click()
 x1last = 0
 x2last = 0
 y1last = 0
 y2last = 0
 PlotData
End Sub

Private Sub saveExcelSpreadsheetmenu_Click()
    ' application workbook, and worksheet objects.
 Dim xlApp As Excel.Application
 Dim xlBook As Excel.Workbook
 Dim xlSheet As Excel.Worksheet
 Dim temp As Integer
   
 CommonDialog1.InitDir = "C:\my documents"
 CommonDialog1.FileName = "*.xls"
 CommonDialog1.Action = 2
 If CommonDialog1.FileName = "*.xls" Then Exit Sub
  
   ' Assign object references to the variables. Use
   ' Add methods to create new workbook and worksheet
   ' objects.
 Set xlApp = New Excel.Application
 xlApp.Caption = "DSC data"
 Set xlBook = xlApp.Workbooks.Add
 Set xlSheet = xlBook.ActiveSheet
 xlSheet.Name = "Data"
 xlSheet.Cells(1, 1).Value = "Temperature"
 xlSheet.Cells(1, 2).Value = "Reference pan"
 xlSheet.Cells(1, 3).Value = "Sample pan"
 xlSheet.Cells(1, 4).Value = "Ramp rate=" & Str(ramp) & " Deg C/min"
      
 If Rampmode = 0 Then
   fsize = size
 Else
   fsize = size \ 2
 End If
  ' Assign the DSC run values  to
   ' Microsoft Excel cells.
 temp = start + 1
 For i = 1 To fsize
  xlSheet.Cells(i + 1, 1).Value = Str(temp)
  xlSheet.Cells(i + 1, 2).Value = Str(Reference(i))
  xlSheet.Cells(i + 1, 3).Value = Str(Sample(i))
  temp = temp + 1
 Next
 If Rampmode = 2 Then
   temp = temp - 2
   For i = 1 To fsize
     xlSheet.Cells(fsize + i + 1, 1).Value = Str(temp)
     xlSheet.Cells(fsize + i + 1, 2).Value = Str(Reference(fsize + i))
     xlSheet.Cells(fsize + i + 1, 3).Value = Str(Sample(fsize + i))
     temp = temp - 1
   Next
 End If
 On Error GoTo 999
 
 xlSheet.SaveAs CommonDialog1.FileName
   ' Close the Workbook
 xlBook.Close
 ' Close Microsoft Excel with the Quit method.

 xlApp.Quit
   ' Release the objects.
 Set xlApp = Nothing
 Set xlBook = Nothing
 Set xlSheet = Nothing
 datasaved = True
 PlotData
Exit Sub
   
999:
   MsgBox ("The chosen file is already in use by Excel")
   xlApp.Quit
   ' Release the objects.
   Set xlApp = Nothing
   Set xlBook = Nothing
   Set xlSheet = Nothing
   PlotData
  


End Sub



Sub SaveSamplerun(fn As String)
  Open fn For Output As #1
    Print #1, start
    Print #1, final
    Print #1, ramp
    Print #1, size
    Print #1, Rampmode
    For i = 1 To size
     Print #1, Sample(i),
     Print #1, Reference(i)
    Next
  Close #1
End Sub
Private Sub Curvfit(n As Integer)

Dim i, j, k, l, m As Integer
Dim v(350), a(350), b(350), c(350), d(350), c2(350), e(350), f(350) As Double
m = 2 ' order of fit is 2
'*****************************************************************
'*      Least squares polynomial fitting subroutine              *
'* ------------------------------------------------------------- *
'* This program least squares fits a polynomial to input data.   *
'* Forsythe orthogonal polynomials are used in the fitting.      *
'* The number of data points is n.                               *
'* The data is fed into in x(i), y(i) pairs.                     *
'* The coefficients are returned in c(i),                        *
'* the smoothed data is returned in v(i),                        *
'* the order of the fit is 2                                     *
'* The standard deviation of the fit is returned in d.           *
'*  e=0, making the fit order m,                                 *
'*****************************************************************
  n1 = m + 1
  V1 = 10000000
  'Initialize the arrays
  For i = 1 To n1
    a(i) = 0
    b(i) = 0
    f(i) = 0
  Next
  For i = 1 To n
    v(i) = 0
    d(i) = 0
  Next
  d1 = Sqr(n)
  w = d1
  For i = 1 To n
    e(i) = 1 / w
  Next
  f1 = d1: a1 = 0
  For i = 1 To n
    a1 = a1 + x(i) * e(i) * e(i)
  Next
  c1 = 0
  For i = 1 To n
    c1 = c1 + y(i) * e(i)
  Next
  b(1) = 1 / f1
  f(1) = b(1) * c1
  For i = 1 To n
    v(i) = v(i) + e(i) * c1
  Next
  m = 1
10 'Save latest results
  For i = 1 To l
    c2(i) = c(i)
  Next
  l2 = l: V2 = v: f2 = f1: a2 = a1: f1 = 0
  For i = 1 To n
    b1 = e(i)
    e(i) = (x(i) - a2) * e(i) - f2 * d(i)
    d(i) = b1
    f1 = f1 + e(i) * e(i)
  Next
  f1 = Sqr(f1)
  For i = 1 To n
    e(i) = e(i) / f1
  Next
  a1 = 0
  For i = 1 To n
    a1 = a1 + x(i) * e(i) * e(i)
  Next
  c1 = 0
  For i = 1 To n
    c1 = c1 + e(i) * y(i)
  Next
  m = m + 1: i = 0:
15 l = m - i: b2 = b(l): d1 = 0
  If l > 1 Then d1 = b(l - 1)
  d1 = d1 - a2 * b(l) - f2 * a(l)
  b(l) = d1 / f1: a(l) = b2: i = i + 1
  If i <> m Then GoTo 15
  For i = 1 To n
    v(i) = v(i) + e(i) * c1
  Next
  For i = 1 To n1
    f(i) = f(i) + b(i) * c1
    c(i) = f(i)
  Next
  vz = 0
  For i = 1 To n
    vz = vz + (v(i) - y(i)) * (v(i) - y(i))
  Next
  'Note the division is by the number of degrees of freedom
  vz = Sqr(vz / (n - l - 1))
  l = m

20 If m = n1 Then GoTo 30
  GoTo 10
30 'Shift the c(i) down, so c(0) is the constant term
  For i = 1 To l
    c(i - 1) = c(i)
  Next
  c(l) = 0
  'l is the order of the polynomial fitted
  l = l - 1: dz = vz
  '  Return
  GoTo 100
50 'Aborted sequence, recover last values
  l = l2: vz = V2
  For i = 1 To l
    c(i) = c2(i)
  Next
  GoTo 30
100

blcoef1 = c(0)
blcoef2 = c(1)
blcoef3 = c(2)
Graph1.ForeColor = RGB(255, 0, 0)
For i = 1 To n
  Graph1.PSet (x(i), v(i))
Next


End Sub




Sub PlotData()
Dim temp As Integer
Graph1.DrawWidth = 1
Call Preparegraph
If Rampmode = 0 Then
 fsize = size
Else
 fsize = size \ 2
End If
  
SP = start + 1
If calibrateoption = True Then
    For i = 1 To fsize
     ' plot sample/reference for upward ramp
     Graph1.ForeColor = RGB(0, 255, 0)
     temp = Sample(i)
     If temp > Graph1.ScaleTop Then
       temp = Graph1.ScaleTop
     End If
     If i = 1 Then
       Graph1.PSet (SP, temp)
       slx = temp
     Else
       Graph1.Line (SP - 1, slx)-(SP, temp)
       slx = temp
     End If
     
     Graph1.ForeColor = RGB(0, 0, 255)
     temp = Reference(i)
     If temp > Graph1.ScaleTop Then
       temp = Graph1.ScaleTop
     End If
     If i = 1 Then
       Graph1.PSet (SP, temp)
       rlx = temp
     Else
       Graph1.Line (SP - 1, rlx)-(SP, temp)
       rlx = temp
     End If
     SP = SP + 1
   Next
  If Rampupdownoption = True Then
     ' plot sample/reference for downward ramp
   SP = SP - 2
   For i = 1 To fsize
     Graph1.ForeColor = RGB(0, 255, 0)
     temp = Sample(fsize + i)
     If temp > Graph1.ScaleTop Then
       temp = Graph1.ScaleTop
     End If
     Graph1.Line (SP + 1, slx)-(SP, temp)
     slx = temp
     Graph1.ForeColor = RGB(0, 0, 255)
     temp = Reference(fsize + i)
     If temp > Graph1.ScaleTop Then
       temp = Graph1.ScaleTop
     End If
     Graph1.Line (SP + 1, rlx)-(SP, temp)
     rlx = temp
     SP = SP - 1
   Next
  End If
Else
   ' plot normal sample run - 1 trace of sample-reference
     For i = 1 To fsize
      Graph1.ForeColor = RGB(0, 0, 255)
      temp = Sample(i) - Reference(i)
      If temp > Graph1.ScaleTop Then
        temp = Graph1.ScaleTop
      End If
      If i = 1 Then
       Graph1.PSet (SP, temp)
      Else
       Graph1.Line -(SP, temp)
      End If
      SP = SP + 1
     Next
   SP = SP - 2
   If Rampupdownoption = True Then
    For i = 1 To fsize
     Graph1.ForeColor = RGB(0, 255, 0)
     temp = Sample(fsize + i) - Reference(fsize + i)
     If temp > Graph1.ScaleTop Then
       temp = Graph1.ScaleTop
     End If
     Graph1.Line -(SP, temp)
     SP = SP - 1
    Next
   End If
  
  
  End If
End Sub


Private Sub Preparegraph()
Dim deltaY As Single

  xmax = Val(FinalTempWind)
  xmin = Val(StartTempwind)
  Xminwind = xmin
  Xmaxwind = xmax
  mouseClickCount = 0
  Graph1.Cls
  Graph1.ScaleLeft = xmin
  Graph1.ScaleWidth = xmax - xmin
  ' when collecting Calibration run use full heating power as F.Scale
  If calibrateoption = True Then
    Graph1.ScaleTop = 25000
    Graph1.ScaleHeight = -25000
  Else  ' data run
    Graph1.ScaleTop = 1000 / VertZoomScroll.Value
    Graph1.ScaleHeight = -2000 / VertZoomScroll.Value
  End If
  Graph1.ForeColor = RGB(0, 0, 0) 'black
  Graph1.Line (xmin, 0)-(xmax, 0)
   ymin = Graph1.ScaleTop + Graph1.ScaleHeight
   VFSlabel = Format(Str(Graph1.ScaleTop), "#######")
   Vminlabel = Format(Str(ymin), "######")
  ' Place the Horizontal Axis Ticks
  j = 1
  hst = 5 - (xmin Mod 5)
  hst = xmin + hst
  hfl = xmax
  ticklen = Graph1.ScaleTop / 100
  For i = hst To hfl Step 5
   If i Mod 10 = 0 Then Tick = 2 * ticklen Else Tick = ticklen
   Graph1.Line (i, ymin)-(i, ymin + Tick)
   j = j + 1
  Next
 Graph1.DrawMode = 13
 displaymode = True
 
End Sub

Private Sub Form_Load()
 On Error Resume Next
 ChDir "C:\my documents"
 If Err.Number <> 0 Then
   MsgBox ("Can't find the C:\my documents folder. You must first create this folder(used to store DSC data)")
   End
 End If
   
 Open "c:\my documents\comport.txt" For Input As #10
 If Err.Number <> 0 Then
    MsgBox ("Can't find the comport.txt file in C:\my documents folder. You must create this file, and place the com port # of the DSC on the first line")
    Close #1
    End
  End If
 Input #10, CommportNumber
 If CommportNumber <> 0 Then
  MSComm1.CommPort = CommportNumber
  Label12.Caption = "Set for COM" & Str(CommportNumber)
  Close #10
 Else
  MsgBox ("Can't find a valid comport # in comport.txt file. Make sure comport.txt file contains the correct com port # of your DSC")
  Close #10
  End
 End If
 
 MSComm1.PortOpen = True
 stopbut.Enabled = False
 VertZoomScroll = 1
 VFSwindow = " X1"
 BaselineCollected = False
 VertZoomScroll.Enabled = False
 mouseClickCount = 0
 mousemode = 0
 Graph1.DrawMode = 1
 calibrateoption = False
 Rampmode = 0 ' default to ramp up only
 HScroll1 = 50
 StartTempwind = HScroll1.Value
 HScroll2 = 250
 FinalTempWind = HScroll2.Value
 Picture1.Visible = False
 CoolerOnButton.Enabled = True
 CoolerOffButton.Enabled = False
 Command1.Enabled = False
End Sub


Private Sub HScroll1_Change()
 StartTempwind = HScroll1.Value
 
End Sub

Private Sub HScroll2_Change()
 FinalTempWind = HScroll2.Value
 End Sub



Private Sub Savetodiskmenu_Click()
 Dim fn, fn2 As String
 Dim a As Integer
   CommonDialog2.InitDir = "C:\my documents"
   CommonDialog2.FileName = "*.dss"
   CommonDialog2.Action = 2
   If CommonDialog2.FileName = "*.dss" Then Exit Sub
   fn = CommonDialog2.FileName
   a = InStr(1, fn, ".")
   If a > 0 Then
     fn = Left$(fn, a - 1)
   End If
   fn2 = fn & ".dss"
   Call SaveSamplerun(fn2)
   StatusWind = "SAMPLE RUN SAVED TO DISK"
  

End Sub

Private Sub SlopeCalibrationBut_Click()
 Dim SampleTop, ReferenceTop As Integer
 Dim index As Integer
 Dim slope As Single
 Dim slopeStr As String * 15
 Dim cmdstring As String
 If Rampmode = 2 Then
    index = (size / 2) - 1
 Else
    index = size - 1
 End If
 sampletempwind = Sample(index)
 referencetempwind = Reference(index)
 slope = Sample(index) / Reference(index)
'slope = 0.9543   'test purposes
  MsgBox ("Slope correction is " & Str(slope))
  slope = slope + 1   ' Arduino 1.65 PARSEFLOAT fail with vals <1 using VB: sub 1 in DSC
  slopeStr = Str(slope)
  slopeStr = Left$(slopeStr, 6)

  ans = InputBox("This should only be done after a run with empty pans has been performed, and if a slope correction is needed- Proceed?", "SLOPE CORRECTION ROUTINE", "No")
  If ans = "No" Then Exit Sub

  Heatsinktemptext = Str(slope - 1)
  cmdstring = "T"
  Call SendCmd(cmdstring)
  cmdstring = " " & slopeStr & Chr(13)
  MSComm1.Output = cmdstring
  StatusWind = "Slope correction sent to instrument"
  
End Sub

Private Sub Startbut_Click()
 Dim c$, d$
 Dim DSCstatus As Integer
 Dim Heat As Long
 Dim SP As Integer
 Dim deltaSample As Integer
 Dim deltaReference As Integer
 Dim heatSample As Integer
 Dim heatReference As Integer
 
 Dim actualtemp As Single
 Dim index As Integer
 Dim T As Integer
 If fastcool = True Then
  Timer2.Enabled = False
  sampletempwind = ""
  referencetempwind = ""
  Heatsinktemptext = ""
 End If
 CoolerOnButton.Enabled = False
 CoolerOffButton.Enabled = False
 If CalibrateButton = Checked Then
  calibrateoption = True
Else
  calibrateoption = False
  SendCmd ("E") 'retrieve the existing curvefit params from DSC's EEPROM into DSC RAM
 End If
 stopbut.Enabled = True
 VertZoomScroll.Value = 1
 mousemode = 0
 index = 1
 StatusWind = ""
 Call Preparegraph
 Graph1.DrawWidth = 1
 start = Val(StartTempwind)
 s = Str(start)
 s = Trim(s)
 SendCmd ("L")
 SendParam (s)
 final = Val(FinalTempWind)
 s = Str(final)
 s = Trim(s)
 SendCmd ("U")
 SendParam (s)
 If Option1 = True Then s = "5"
 If Option2 = True Then s = "10"
 If Option3 = True Then s = "20"
 ramp = Val(s)
 SendCmd ("R")
 SendParam (s)
 SendCmd ("M")
 If Rampupoption = True Then
  s = " 0"
 Else
  s = "2"
 End If
 SendParam (s)
 SendCmd ("G")
 StatusWind = "RAMPING UP TO START TEMPERATURE."
  
  Do
  DoEvents
  c$ = getline()
  Text1 = c$
  ' Extract the Status value
  temp = InStr(1, c$, " ")
  d$ = Left$(c$, temp - 1)
  DSCstatus = Val(d$)
  If DSCstatus = 1 Then
   StatusWind = "STABILIZING"
   End If
  If DSCstatus = 2 Then
   StatusWind = "PERFORMING RUN"
  End If
  If DSCstatus = 3 Then
   StatusWind = " RUN COMPLETE"
   Exit Do
  End If
  If DSCstatus = 4 Then
   StatusWind = " RUN ABORTED"
   Exit Do
  If DSCstatus = 5 Then
   StatusWind = " RAMPING DOWN"
  End If
   If DSCstatus = 9 Then
     MsgBox ("OVER-HEATING- CHECK WATER FLOW")
     Exit Do
   End If
  End If
  
  ' Extract the SP value
  temp2 = InStr(temp + 1, c$, " ")
  Length = temp2 - temp
  d$ = Mid$(c$, temp + 1, Length)
  SP = Val(d$)
  cursXwind = d$
  
  ' Extract the sample pan deviation from setpoint
  temp3 = InStr(temp2 + 1, c$, " ")
  Length = temp3 - temp2
  d$ = Mid$(c$, temp2 + 1, Length)
  deltaSample = Val(d$)
  
  ' Extract the reference pan deviation from setpoint
  temp4 = InStr(temp3 + 1, c$, " ")
  Length = temp4 - temp3
  d$ = Mid$(c$, temp3 + 1, Length)
  deltaReference = Val(d$)
  
  ' extract the sample applied heat
  temp5 = InStr(temp4 + 1, c$, " ")
  Length = temp5 - temp4
  d$ = Mid$(c$, temp4 + 1, Length)
  heatSample = Val(d$)
  Sample(index) = heatSample
  'Text4 = heatSample
  
  ' extract the Reference applied heat
  temp6 = Len(c$)
  Length = temp6 - temp5
  d$ = Mid$(c$, temp5 + 1, Length)
  heatReference = Val(d$)
  Reference(index) = heatReference
  'Text5 = heatReference
  
  If (DSCstatus = 2) Or (DSCstatus = 5) Then
   If calibrateoption = True Then
    ' calibration run graphing- show both sample and reference
     If heatSample > Graph1.ScaleTop Then
       heatSample = Graph1.ScaleTop
     End If
     If heatReferenc > Graph1.ScaleTop Then
       heatReference = Graph1.ScaleTop
     End If
     Graph1.ForeColor = RGB(0, 255, 0)  ' Green for sample
     Graph1.PSet (SP, heatSample)
     Graph1.ForeColor = RGB(0, 0, 255)  ' blue for Reference
     Graph1.PSet (SP, heatReference)
   Else
    ' normal run graphing - show sample-reference
    If DSCstatus = 2 Then
     Graph1.ForeColor = RGB(0, 0, 0) ' black on way up
    Else
     Graph1.ForeColor = RGB(255, 0, 0) ' red on way down
    End If
    If index = 1 Then
     Graph1.PSet (SP, heatSample - heatReference)
    Else
     Graph1.Line -(SP, heatSample - heatReference)
    End If
   End If
   index = index + 1
  End If
Loop

If DSCstatus = 4 Then
 stopbut.Enabled = False
 Startbut.Enabled = True
 StatusWind = "RUN ABORTED"
 CoolerOnButton.Enabled = True
 CoolerOffButton.Enabled = True
 Exit Sub
End If

size = index - 1
If calibrateoption = True Then
  Call SaveSamplerun("calibration")
 StatusWind = " SAVING CALIBRATION DATA TO DISK"
 Command1.Visible = True
 Command1.Enabled = True
Else
  Call SaveSamplerun("currun")
  StatusWind = " SAVING RUN DATA TO DISK"
End If
CoolerOnButton.Enabled = True
CoolerOffButton.Enabled = True
VertZoomScroll.Enabled = True
End Sub

Sub SendCmd(cmd As String)
 MSComm1.Output = cmd
 Do  'wait for parameter to be echoed back
  DoEvents
 Loop Until MSComm1.Input = cmd
 End Sub
Sub SendParam(par As String)
MSComm1.Output = par & Chr$(13)
Do  'wait for parameter to be echoed back
  DoEvents
Loop Until MSComm1.Input = Chr$(10)
End Sub


 
Function getline() As String
 Dim a$, b$
  a$ = ""
  Do
   b$ = MSComm1.Input
   If b$ = Chr$(10) Then Exit Do
   a$ = a$ + b$
  Loop
  getline = a$
 End Function

Private Sub stopbut_Click()
 MSComm1.Output = "Q"
End Sub

Private Sub Timer1_Timer()
 timer1flag = True
End Sub

Private Sub Timer2_Timer()
Dim c$, d$, z$
Dim temp2, temp3 As Integer
 timerindex = timerindex + 1
 If timerindex >= 180 Then  'if fast cool runs for more than 6 minutes without run starting,turn peltier off
  If fastcool = True Then
   fastcool = False
   Picture1.Visible = False
   SendCmd ("C")
   SendParam ("0")
  End If
 End If
 
  SendCmd ("Z")
  c$ = getline()
    
    ' Extract the sample temperature
  d$ = Mid$(c$, 1, 5)
  sampletempwind = d$
  temp2 = InStr(1, c$, " ")
  d$ = Mid$(c$, temp2 + 1, 6)
  referencetempwind = d$
  
  'Extract heatsink temperature
   Length = 4
   d$ = Right$(c$, Length)
     
   ' convert HS thermistor ADC reading to deg C using rough formula derived in Excel
   ' using a 10K thermistor with a beta of 3860
   hstemp = 75 - (Val(d$) / 10)
   Heatsinktemptext = Int(hstemp)
   timer2flag = True
 
End Sub

Private Sub VertZoomScroll_Change()
 VFSwindow = "X " & Str(VertZoomScroll.Value)
End Sub
