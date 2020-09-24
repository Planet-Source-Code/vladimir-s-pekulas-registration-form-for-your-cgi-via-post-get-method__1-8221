VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frm_Register 
   Caption         =   "Registration Form ..."
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6675
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   140
      TabIndex        =   59
      Top             =   5520
      Width           =   2050
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1560
         Picture         =   "Form2.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mandatory "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frm1 
      Height          =   6135
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   6120
         Picture         =   "Form2.frx":014A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   56
         Top             =   3720
         Width           =   240
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   6120
         Picture         =   "Form2.frx":0294
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   42
         Top             =   5040
         Width           =   240
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   6120
         Picture         =   "Form2.frx":03DE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   240
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   6120
         Picture         =   "Form2.frx":0528
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         Top             =   937
         Width           =   240
      End
      Begin VB.TextBox txtWebSite 
         Height          =   285
         Left            =   1530
         TabIndex        =   16
         Text            =   "http://"
         Top             =   5475
         Width           =   4455
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   1530
         TabIndex        =   13
         Top             =   4275
         Width           =   1455
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   4530
         TabIndex        =   14
         Top             =   4275
         Width           =   1455
      End
      Begin VB.TextBox txtMail 
         Height          =   285
         Left            =   1530
         TabIndex        =   15
         Top             =   4995
         Width           =   4455
      End
      Begin VB.TextBox txtProv 
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   3675
         Width           =   735
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   2970
         TabIndex        =   12
         Top             =   3675
         Width           =   3015
      End
      Begin VB.TextBox txtAdd1 
         Height          =   285
         Left            =   2250
         TabIndex        =   8
         Top             =   2235
         Width           =   3735
      End
      Begin VB.TextBox txtAdd2 
         Height          =   285
         Left            =   2250
         TabIndex        =   9
         Top             =   2715
         Width           =   3735
      End
      Begin VB.TextBox txtAdd3 
         Height          =   285
         Left            =   2250
         TabIndex        =   10
         Top             =   3195
         Width           =   3735
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   2250
         TabIndex        =   7
         Top             =   1395
         Width           =   3735
      End
      Begin VB.TextBox txtFirst 
         Height          =   285
         Left            =   2250
         TabIndex        =   6
         Top             =   915
         Width           =   3735
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   255
         Left            =   5130
         TabIndex        =   5
         Top             =   435
         Width           =   735
      End
      Begin VB.OptionButton optMs 
         Caption         =   "Ms."
         Height          =   255
         Left            =   3810
         TabIndex        =   3
         Top             =   435
         Width           =   735
      End
      Begin VB.OptionButton optDr 
         Caption         =   "Dr."
         Height          =   255
         Left            =   4530
         TabIndex        =   4
         Top             =   435
         Width           =   735
      End
      Begin VB.OptionButton optMrs 
         Caption         =   "Mrs."
         Height          =   255
         Left            =   2970
         TabIndex        =   2
         Top             =   435
         Width           =   735
      End
      Begin VB.OptionButton optMr 
         Caption         =   "Mr."
         Height          =   255
         Left            =   2250
         TabIndex        =   1
         Top             =   435
         Width           =   735
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "WebSite:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   570
         TabIndex        =   26
         Top             =   5475
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   570
         TabIndex        =   25
         Top             =   4275
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3570
         TabIndex        =   24
         Top             =   4275
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "E@mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   570
         TabIndex        =   23
         Top             =   4995
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   570
         TabIndex        =   22
         Top             =   3675
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2490
         TabIndex        =   21
         Top             =   3675
         Width           =   615
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Home Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   570
         TabIndex        =   20
         Top             =   2235
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   570
         TabIndex        =   19
         Top             =   1395
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   570
         TabIndex        =   18
         Top             =   915
         Width           =   1455
      End
   End
   Begin VB.Timer T 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10000
      Top             =   5040
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   135
      Left            =   0
      TabIndex        =   57
      Top             =   10200
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   -120
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame frm3 
      Height          =   6135
      Left            =   2280
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   4200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect to Database"
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblEnd 
         Caption         =   "Completed Sucessfuly !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   58
         Top             =   2880
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblProc 
         Caption         =   "0"
         Height          =   255
         Left            =   2280
         TabIndex        =   53
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percent Copleted:"
         Height          =   255
         Index           =   17
         Left            =   840
         TabIndex        =   52
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Please Press the button bellow to start the submision of data you"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   840
         TabIndex        =   49
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "have entred."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   840
         TabIndex        =   48
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Thank you for your time."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   840
         TabIndex        =   47
         Top             =   1200
         Width           =   5775
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame frm2 
      Height          =   6135
      Left            =   2280
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox CoEdu 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2685
         Width           =   2775
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   4800
         Picture         =   "Form2.frx":0672
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
         Top             =   3997
         Width           =   240
      End
      Begin VB.CheckBox chPhone 
         Caption         =   "Can we contact you by phone ?"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   5160
         Width           =   2655
      End
      Begin VB.CheckBox chSubs 
         Caption         =   "Subscribe to our Mailing List ?"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   4680
         Width           =   2655
      End
      Begin VB.ComboBox CoProduct 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3960
         Width           =   4095
      End
      Begin VB.ComboBox CoJob 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1965
         Width           =   2775
      End
      Begin VB.ComboBox CoAge 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1245
         Width           =   2775
      End
      Begin VB.ComboBox CoGender 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   525
         Width           =   2775
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Product to Register:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   37
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Education:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   36
         Top             =   2715
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Job:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   34
         Top             =   1995
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   32
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   29
         Top             =   555
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5595
      Left            =   0
      Picture         =   "Form2.frx":07BC
      ScaleHeight     =   5595
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   0
      Width           =   2265
   End
End
Attribute VB_Name = "frm_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# (C) 2000, Vladimir S. Pekulas
'#
'# This app was writen for anybody who needs a registration form
'# for his/her app., and then to be able to either connect to your
'# on-line database, or send the informnation directly to your
'# mailbox.
'# Since this doesn't do much more then that it's at least a time
'# saver.
'#
'# Enyoj.
'#
'# PS: Sorry for my English.
'#     Oh, you can NOT use the main graphics which is on the form.
'#     It's only an example and it's copyrighted by Corel Corp.

Dim Pre_Fix_vb As String
Dim NextGo As Integer '# Should we go to the next window ?
Dim TopY As Integer, LeftX As Integer '# Position of next or "back" window

Private Sub cmdConnect_Click()
On Error Resume Next
T.Enabled = True
'# Connects to the web using "GET" method.
'# If you connect to the script that uses POST then call
'  "CompleteStringForPOSTMethod" !
CopleteGetString
cmdConnect.Enabled = False
cmdBack.Enabled = False
End Sub

Private Sub cmdExit_Click() '# Exit
On Error Resume Next
   If MsgBox("Would you like to Exit Registration ?", vbQuestion + vbYesNo, "Exit ?") = vbYes Then
        'Yes
        End
   Else
        'No
        Exit Sub
   End If
End Sub

Private Sub cmdNext_Click() '# What window's next ?
On Error Resume Next
CheckFields
PreFixValue
If NextGo = 1 Then
    MsgBox "Please Enter Required Fields."
    Exit Sub
Else

'# Frame 1 going to Frame 2
    If frm1.Visible = True Then
       frm1.Visible = False
       frm3.Visible = False
       frm2.Top = TopY
       frm2.Left = LeftX
       frm2.Visible = True
       cmdBack.Enabled = True
       Exit Sub
    End If

'# Do we have all we need ?
If CoProduct.Text = "" Then
    MsgBox "Please Select Product to Register."
    Exit Sub
End If

'# Frame 2 going to Frame 3
    If frm2.Visible = True Then
       frm2.Visible = False
       frm1.Visible = False
       frm3.Top = TopY
       frm3.Left = LeftX
       frm3.Visible = True
       cmdBack.Enabled = True
       cmdNext.Enabled = False
       Exit Sub
    End If
    cmdBack.Enabled = True
End If
End Sub

Private Sub cmdBack_Click() '# What window to show ?
On Error Resume Next
  If frm2.Visible = True Then
       frm2.Visible = False
       frm3.Visible = False
       frm1.Top = TopY
       frm1.Left = LeftX
       frm1.Visible = True
       cmdBack.Enabled = False
       Exit Sub
    End If
'# Frame 2 going to Frame 3
    If frm3.Visible = True Then
       frm3.Visible = False
       frm1.Visible = False
       frm2.Top = TopY
       frm2.Left = LeftX
       frm2.Visible = True
       cmdBack.Enabled = True
       cmdNext.Enabled = True
       Exit Sub
    End If
End Sub



Function CheckFields()      '# Check mandatory fields absolutlly optional ....
On Error Resume Next
If txtFirst.Text = "" Then
    NextGo = 1
Else
    NextGo = 0
End If
'##
If txtLast.Text = "" Then
    NextGo = 1
Else
    NextGo = 0
End If
'##
If txtMail.Text = "" Then
    NextGo = 1
Else
    NextGo = 0
End If
End Function

Private Sub Form_Activate()
On Error Resume Next
'Got focus
optMr.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
'# Variables for frame containers
TopY = frm1.Top
LeftX = frm1.Left
frm1.Visible = True
frm2.Visible = False
frm3.Visible = False

'Combos for Gender
CoGender.AddItem "Female"
CoGender.AddItem "Male"
CoGender.AddItem "Decline"

'Combos for Age
CoAge.AddItem "Less then 18"
CoAge.AddItem "18 to 24"
CoAge.AddItem "24 to 30"
CoAge.AddItem "31 to 39"
CoAge.AddItem "40 to 50"
CoAge.AddItem "Over 50"
CoAge.AddItem "Decline"

'Combos for Job
CoJob.AddItem "Full-Time Student"
CoJob.AddItem "Part-Time Student"
CoJob.AddItem "Full-Time Employed"
CoJob.AddItem "Part-Time Employed"
CoJob.AddItem "Self Employed"
CoJob.AddItem "Retired"
CoJob.AddItem "Other"
CoJob.AddItem "Deline"

'Combos for Education
CoEdu.AddItem "High-School"
CoEdu.AddItem "College Degree"
CoEdu.AddItem "University Degree"
CoEdu.AddItem "Masters Degree"
CoEdu.AddItem "Doctorate"
CoEdu.AddItem "Other"
CoEdu.AddItem "Decline"


'Combos for Products
CoProduct.AddItem "Product 1"
CoProduct.AddItem "Product 2"
CoProduct.AddItem "Product 3"
CoProduct.AddItem "Product 4"
CoProduct.AddItem "Product 5"

End Sub


Function PreFixValue() '# What's the Pre-Fix ?
On Error Resume Next
If optMr.Value = True Then Pre_Fix_vb = "Mr."
If optMrs.Value = True Then Pre_Fix_vb = "Mrs."
If optMs.Value = True Then Pre_Fix_vb = "Ms."
If optDr.Value = True Then Pre_Fix_vb = "Dr."
If optOther.Value = True Then Pre_Fix_vb = "Other"
End Function



Function CopleteGetString()
On Error Resume Next
Dim GetMethod As String, Subscribe As String, Phone As String
' is chSubscribe checked if so then remember it ....
    If chSubs.Value = 1 Then
        Subscribe = "Subscribe"
    Else
        Subscribe = "Do Not Subscribe"
    End If
' Same for Phone .....
    If chPhone.Value = 1 Then
        Phone = "Allowed"
    Else
        Phone = "Not Allowed"
    End If
' Change the recipient= to your email address and of course change the URL
' to your script which is using "GET".
GetMethod = "http://www.Your-Site.com/cgi-bin/FormMail.cgi?recipient=Your@email.com&subject=Visual+Basic+Registration%21&pre-fix=" & Pre_Fix_vb & "&First_Name=" & txtFirst.Text & "&Last_Name=" & txtLast.Text & "&Address1=" & txtAdd1.Text & "&Address2=" & txtAdd2.Text & "&Address3=" & txtAdd3.Text & "&Province=" & txtProv.Text & "&City=" & txtCity.Text & "&Phone=" & txtPhone.Text & "&Fax=" & txtPhone.Text & "&E@mail=" & txtMail.Text & "&WebSite=" & txtWebSite.Text & "&Gender=" & CoGender.Text & "&Age=" & CoAge.Text & "&Job=" & CoJob.Text & "&Product=" & CoProduct.Text & "&Mailing_List=" & Subscribe & "&Phone_Contact=" & Phone & "&Submit=" & "By Registration"
Web.Navigate GetMethod
End Function


Function CompleteStringForPOSTMethod()
On Error Resume Next
' # This is the function for "POST" method
' # The more common method is "GET", but still for some
' # scripts here is it.
Dim strURL As String, strFormData As String
' change URL to your script which uses "POST" method.
    strURL = "http://www.Your-Site.com/cgi-bin/PostMethod.cgi"
' Either use the form.html included with this project or modify
' the line bellow for your form.
    strFormData = "&pre-fix=" & Pre_Fix_vb & "&First_Name=" & txtFirst.Text & "&Last_Name=" & txtLast.Text & "&Address1=" & txtAdd1.Text & "&Address2=" & txtAdd2.Text & "&Address2=" & txtAdd3.Text & "&Province=" & txtProv.Text & "&City=" & txtCity.Text & "&Phone=" & txtPhone.Text & "&Fax=" & txtFax.Text & "&E@mail=" & txtMail.Text & "&WebSite=" & txtWebSite.Text & "&Gender=" & CoGender.Text & "&Age=" & CoAge.Text & "&Job=" & CoJob.Text & "&Product=" & CoProduct.Text
    Inet1.Execute strURL, "Post", strFormData, _
            "Content-Type: application/x-www-form-urlencoded"
End Function

Private Sub T_Timer()
On Error Resume Next
' Just a very dummy prg bar for download ...
' Anyone knows how to control the download/Upload progress ?
If T.Interval = 200 Then
    If Web.Busy = True Then
If prg.Value = 100 Then prg.Value = 0
prg.Value = prg.Value + 5
lblProc.Caption = prg.Value
    Else
    prg.Value = 100
    lblProc.Caption = prg.Value
lblEnd.Visible = True
    Exit Sub
    End If
End If
End Sub
