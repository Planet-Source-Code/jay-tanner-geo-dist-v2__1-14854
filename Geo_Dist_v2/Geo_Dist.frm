VERSION 5.00
Begin VB.Form Geodesic_Distance_Form 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Geodesic Distance Between Points"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Geo_Dist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4515
   Begin VB.CheckBox Keep_on_Top 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Keep on Top "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   135
      TabIndex        =   41
      Top             =   3420
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.TextBox Lat2_NS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "N"
      Top             =   2205
      Width           =   390
   End
   Begin VB.TextBox Lng2_EW 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   45
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "E"
      Top             =   2205
      Width           =   390
   End
   Begin VB.TextBox Lat1_NS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "N"
      Top             =   855
      Width           =   390
   End
   Begin VB.TextBox Lng1_EW 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   45
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "W"
      Top             =   855
      Width           =   390
   End
   Begin VB.TextBox Lat2_Sec 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      MaxLength       =   4
      TabIndex        =   15
      Text            =   "11.0"
      Top             =   2205
      Width           =   660
   End
   Begin VB.TextBox Lat2_Min 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "50"
      Top             =   2205
      Width           =   390
   End
   Begin VB.TextBox Lng2_Sec 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1395
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "14.0"
      Top             =   2205
      Width           =   660
   End
   Begin VB.TextBox Lng2_Min 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   990
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "20"
      Top             =   2205
      Width           =   390
   End
   Begin VB.TextBox Lat1_Sec 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "17.0"
      Top             =   855
      Width           =   660
   End
   Begin VB.TextBox Lat1_Min 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "55"
      Top             =   855
      Width           =   390
   End
   Begin VB.TextBox Lng1_Sec 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "55.5"
      Top             =   855
      Width           =   660
   End
   Begin VB.TextBox Lng1_Min 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   990
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "03"
      Top             =   855
      Width           =   435
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output Distance Units"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   45
      TabIndex        =   20
      Top             =   2790
      Width           =   4440
      Begin VB.OptionButton km_Option 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Kilometers "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton mi_Option 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Miles "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3330
         TabIndex        =   18
         Top             =   270
         Width           =   990
      End
      Begin VB.OptionButton nmi_Option 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Nautical Miles "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1665
         TabIndex        =   17
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.TextBox Computed_Dist 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3915
      Width           =   2775
   End
   Begin VB.TextBox Lng2_Deg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   450
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "02"
      Top             =   2205
      Width           =   525
   End
   Begin VB.TextBox Lat1_Deg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2925
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "38"
      Top             =   855
      Width           =   390
   End
   Begin VB.TextBox Lng1_Deg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   450
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "77"
      Top             =   855
      Width           =   525
   End
   Begin VB.TextBox Lat2_Deg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2925
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "48"
      Top             =   2205
      Width           =   390
   End
   Begin VB.CommandButton Compute_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Compute"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   45
      TabIndex        =   19
      Top             =   3915
      Width           =   1605
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4545
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4500
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4545
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4545
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3285
      TabIndex        =   40
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Longitude and Latitude of 2nd Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   39
      Top             =   1440
      Width           =   4425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Longitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   810
      TabIndex        =   38
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3285
      TabIndex        =   37
      Top             =   405
      Width           =   915
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Longitude and Latitude of 1st Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   36
      Top             =   90
      Width           =   4425
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deg"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2925
      TabIndex        =   35
      Top             =   1980
      Width           =   390
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Min"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3330
      TabIndex        =   34
      Top             =   1980
      Width           =   405
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sec"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3780
      TabIndex        =   33
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deg"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   450
      TabIndex        =   32
      Top             =   1980
      Width           =   480
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Min"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   990
      TabIndex        =   31
      Top             =   1980
      Width           =   405
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sec"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1440
      TabIndex        =   30
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sec"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3780
      TabIndex        =   29
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Min"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3330
      TabIndex        =   28
      Top             =   630
      Width           =   405
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deg"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2925
      TabIndex        =   27
      Top             =   630
      Width           =   390
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sec"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1485
      TabIndex        =   26
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Min"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1035
      TabIndex        =   25
      Top             =   630
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Distance Between Locations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1710
      TabIndex        =   24
      Top             =   3690
      Width           =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Longitude"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   855
      TabIndex        =   22
      Top             =   405
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deg"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   495
      TabIndex        =   21
      Top             =   630
      Width           =   480
   End
End
Attribute VB_Name = "Geodesic_Distance_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

' A program to compute the geodesic distance between two points
' on the surface of the Earth to within approximately 50 meters.
'
' Function written by Jay Tanner - Based on the mathematical
' method of H. Andoyer.
'
' v2.0
'
' References:
'
' International Earth Rotation Service,
' Annual report for 1996 (Observatorie de Paris, 1997)
'
' Annuaire du Bureau des Longitudes pour 1950 (Paris) pg.145
'
'
'
' What to do when this program starts up
' ======================================
  Private Sub Form_Load()

  Recall_Program_Settings "Geodesic Distance Calculator", _
  Geodesic_Distance_Form, True

  If Keep_on_Top.Value = 1 Then
'    Lock on top
     SetWindowPos Geodesic_Distance_Form.hwnd, -1, 0, 0, 0, 0, 3
  Else
'    Restore to normal
     SetWindowPos Geodesic_Distance_Form.hwnd, -2, 0, 0, 0, 0, 3
  End If

  End Sub

' What to do when the COMPUTE button is clicked
' =============================================
  Private Sub Compute_Button_Click()
  COMPUTE
  End Sub

' What to do when the
' ===============================
  Private Sub Keep_On_Top_Click()

  If Keep_on_Top.Value = 1 Then
     Keep_on_Top.FontBold = True
     SetWindowPos Geodesic_Distance_Form.hwnd, -1, 0, 0, 0, 0, 3
  Else
     Keep_on_Top.FontBold = False
     SetWindowPos Geodesic_Distance_Form.hwnd, -2, 0, 0, 0, 0, 3
     Geodesic_Distance_Form.AutoRedraw = True
  End If

  End Sub
  
' ==================================
' SUB to read the interface data and
' call the computation module.

  Private Sub COMPUTE()

  Dim Long_1  As Double
  Dim Lat_1   As Double
  Dim Long_2  As Double
  Dim Lat_2   As Double

  Dim Q       As Variant

  Dim DUnits  As String

' Adjust appearance of interface settings for neatness
  Lng1_EW = UCase(Lng1_EW): If Lng1_EW <> "E" Then Lng1_EW = "W"
  Lat1_NS = UCase(Lat1_NS): If Lat1_NS <> "S" Then Lat1_NS = "N"
  Lng2_EW = UCase(Lng2_EW): If Lng2_EW <> "E" Then Lng2_EW = "W"
  Lat2_NS = UCase(Lat2_NS): If Lat2_NS <> "S" Then Lat2_NS = "N"

' Add padding zeros to minutes and seconds values
  Lng1_Min = Format(Val(Lng1_Min), "0#")
  Lng1_Sec = Format(Val(Lng1_Sec), "0#.0")
  Lng2_Min = Format(Val(Lng2_Min), "0#")
  Lng2_Sec = Format(Val(Lng2_Sec), "0#.0")
  Lat1_Min = Format(Val(Lat1_Min), "0#")
  Lat1_Sec = Format(Val(Lat1_Sec), "0#.0")
  Lat2_Min = Format(Val(Lat2_Min), "0#")
  Lat2_Sec = Format(Val(Lat2_Sec), "0#.0")

' Read Long, Lat arguments given in degrees, minutes and seconds
' and convert into decimal degrees.
     Q = (Val(Lng1_Deg) * 3600 + Val(Lng1_Min) _
        * 60 + Val(Lng1_Sec)) / 3600

  If UCase(Trim(Lng1_EW)) = "E" Then Q = -Q Else Lng1_EW = "W"
     Long_1 = Q

     Q = (Val(Lat1_Deg) * 3600 + Val(Lat1_Min) _
        * 60 + Val(Lat1_Sec)) / 3600
  If UCase(Trim(Lat1_NS)) = "S" Then Q = -Q Else Lat1_NS = "N"
     Lat_1 = Q

     Q = (Val(Lng2_Deg) * 3600 + Val(Lng2_Min) _
        * 60 + Val(Lng2_Sec)) / 3600
  If UCase(Trim(Lng2_EW)) = "E" Then Q = -Q Else Lng2_EW = "W"
     Long_2 = Q

     Q = (Val(Lat2_Deg) * 3600 + Val(Lat2_Min) _
        * 60 + Val(Lat2_Sec)) / 3600

  If UCase(Trim(Lat2_NS)) = "S" Then Q = -Q Else Lat2_NS = "N"
     Lat_2 = Q

' =================================
' Determine distance units selected
  If km_Option.Value = True Then DUnits = " km"
  If mi_Option.Value = True Then DUnits = " mi"
  If nmi_Option.Value = True Then DUnits = " nmi"

  Q = Geo_Dist_Between(Long_1, Lat_1, Long_2, Lat_2, DUnits)
  
' Attach distance units symbol to computed result
  If km_Option.Value = True Then
     Q = Format(Q, "#,###.#0") & " km  ±0.05"
  End If
  If mi_Option.Value = True Then
     Q = Format(Q, "#,###.#0") & " mi  ±0.03"
  End If
  If nmi_Option.Value = True Then
     Q = Format(Q, "#,###.#0") & " nmi  ±0.03"
  End If
  
' Compute the geodesic surface distance between the coordinates.
  Computed_Dist = Q

' Save interface settings after a computation
  Store_Program_Settings "Geodesic Distance Calculator", _
  Geodesic_Distance_Form, True
  
  End Sub
  
' =====================================
' What do do when this form is unloaded

  Private Sub Form_Unload(Cancel As Integer)
  Store_Program_Settings "Geodesic Distance Calculator", _
  Geodesic_Distance_Form, True
  End Sub

' =================================================
' These SUBs adjust the computations when the units
' of distance are changed.

  Private Sub km_option_Click()
  km_Option.FontBold = True
  mi_Option.FontBold = False
  nmi_Option.FontBold = False
  COMPUTE
  End Sub

  Private Sub mi_option_Click()
  km_Option.FontBold = False
  mi_Option.FontBold = True
  nmi_Option.FontBold = False
  COMPUTE
  End Sub

  Private Sub nmi_option_Click()
  km_Option.FontBold = False
  mi_Option.FontBold = False
  nmi_Option.FontBold = True
  COMPUTE
  End Sub

' ==================================
' Highlight selected values on focus
  Private Sub Lng1_EW_GotFocus()
  Lng1_EW.SelStart = 0
  Lng1_EW.SelLength = Len(Lng1_EW.Text)
  End Sub
  Private Sub Lng1_Deg_GotFocus()
  Lng1_Deg.SelStart = 0
  Lng1_Deg.SelLength = Len(Lng1_Deg.Text)
  End Sub
  Private Sub Lng1_Min_GotFocus()
  Lng1_Min.SelStart = 0
  Lng1_Min.SelLength = Len(Lng1_Min.Text)
  End Sub
  Private Sub Lng1_Sec_GotFocus()
  Lng1_Sec.SelStart = 0
  Lng1_Sec.SelLength = Len(Lng1_Sec.Text)
  End Sub
  Private Sub Lat1_NS_GotFocus()
  Lat1_NS.SelStart = 0
  Lat1_NS.SelLength = Len(Lat1_NS.Text)
  End Sub
  Private Sub Lat1_Deg_GotFocus()
  Lat1_Deg.SelStart = 0
  Lat1_Deg.SelLength = Len(Lat1_Deg.Text)
  End Sub
  Private Sub Lat1_Min_GotFocus()
  Lat1_Min.SelStart = 0
  Lat1_Min.SelLength = Len(Lat1_Min.Text)
  End Sub
  Private Sub Lat1_Sec_GotFocus()
  Lat1_Sec.SelStart = 0
  Lat1_Sec.SelLength = Len(Lat1_Sec.Text)
  End Sub

  Private Sub Lng2_EW_GotFocus()
  Lng2_EW.SelStart = 0
  Lng2_EW.SelLength = Len(Lng2_EW.Text)
  End Sub
  Private Sub Lng2_Deg_GotFocus()
  Lng2_Deg.SelStart = 0
  Lng2_Deg.SelLength = Len(Lng2_Deg.Text)
  End Sub
  Private Sub Lng2_Min_GotFocus()
  Lng2_Min.SelStart = 0
  Lng2_Min.SelLength = Len(Lng2_Min.Text)
  End Sub
  Private Sub Lng2_Sec_GotFocus()
  Lng2_Sec.SelStart = 0
  Lng2_Sec.SelLength = Len(Lng2_Sec.Text)
  End Sub
  Private Sub Lat2_NS_GotFocus()
  Lat2_NS.SelStart = 0
  Lat2_NS.SelLength = Len(Lat2_NS.Text)
  End Sub
  Private Sub Lat2_Deg_GotFocus()
  Lat2_Deg.SelStart = 0
  Lat2_Deg.SelLength = Len(Lat2_Deg.Text)
  End Sub
  Private Sub Lat2_Min_GotFocus()
  Lat2_Min.SelStart = 0
  Lat2_Min.SelLength = Len(Lat2_Min.Text)
  End Sub
  Private Sub Lat2_Sec_GotFocus()
  Lat2_Sec.SelStart = 0
  Lat2_Sec.SelLength = Len(Lat2_Sec.Text)
  End Sub

' This is the core function around which the main program is built
' =================================================================
  Public Function Geo_Dist_Between(Long1, Lat1, Long2, Lat2, Units)
' V2.0
'
' Compute the gedesic distance between two Earth surface
' coordinates to an accuracy of about ±50 meters.  To get
' this degree of accuracy, this function takes into account
' the spheroidal flattening factor of the Earth rather than
' assuming that the Earth is a perfect sphere.
'
' The arguments (Lng1, Lat1, Lng2, Lat2, Units) are the longitudes
' and latitudes of the two locations expressed in degrees and the
' units to be used for the computed distance (km, mi, nmi).
'
' Where:  km = Kilometers,   mi = Miles,   nmi = Nautical miles
'
' Positive longitude is west and negative is east.
'
' The default coordinates in the default interface are for the
' U.S. Naval Observatory in Washington, D.C. and the Paris
' Observatory in France (Observatorie de Paris).
'
' When changed, the settings remain preserved in the registry
' until changed again so they are not lost when the program
' is terminated.
' ===========================================================
' Function written by Jay Tanner - Based on the mathematical
' method of H. Andoyer.
'
' References:
'
' International Earth Rotation Service,
' Annual report for 1996 (Observatorie de Paris, 1997)
'
' Annuaire du Bureau des Longitudes pour 1950 (Paris) pg.145
'
' ==========================================================
' Flattening factor of the earth geoid
  Const ff As Double = 1 / 298.257

' Auxiliary working variables
  Dim C  As Double
  Dim D  As Double
  Dim F  As Double
  Dim G  As Double
  Dim H1 As Double
  Dim H2 As Double
  Dim L  As Double
  Dim O  As Double
  Dim R  As Double
  Dim S  As Double
  Dim W1 As Double
  Dim W2 As Double
  Dim SG As Double
  Dim CG As Double
  Dim SF As Double
  Dim CF As Double
  Dim SL As Double
  Dim CL As Double
  Dim U  As String
  Dim UF As Double
  
'    Read output units (km, mi or nmi)
'    and set units factor for conversion.
     U = LCase(Trim(Units))
    UF = 1
  If U = "mi" Then UF = 1.609344
  If U = "nmi" Then UF = 1.852
  
' Compute auxiliary angles
  F = (Lat1 + Lat2) / 2
  G = (Lat1 - Lat2) / 2
  L = (Long1 - Long2) / 2
  
' Compute sines and cosines of auxiliary angles
  SG = Sin(G * Atn(1) / 45)
  CG = Cos(G * Atn(1) / 45)
  SF = Sin(F * Atn(1) / 45)
  CF = Cos(F * Atn(1) / 45)
  SL = Sin(L * Atn(1) / 45)
  CL = Cos(L * Atn(1) / 45)

  W1 = SG * CL: W1 = W1 * W1
  W2 = CF * SL: W2 = W2 * W2
   S = W1 + W2

  W1 = CG * CL: W1 = W1 * W1
  W2 = SF * SL: W2 = W2 * W2
   C = W1 + W2

   O = Atn(Sqr(S / C))
   R = Sqr(S * C) / O
   D = 2 * O * 6378.14

  H1 = (3 * R - 1) / (2 * C)
  H2 = (3 * R + 1) / (2 * S)

' Compute the angle between the points on a synthetic
' geodesic sphereoid connecting the two points.
  W1 = SF * CG: W1 = W1 * W1 * H1 * ff + 1
  W2 = CF * SG: W2 = W2 * W2 * H2 * ff

' Return the distance between the given locations in
' the units indicated by the units factor UF
  Geo_Dist_Between = D * (W1 - W2) / UF

  End Function


