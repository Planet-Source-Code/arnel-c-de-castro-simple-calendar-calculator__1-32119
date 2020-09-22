VERSION 5.00
Begin VB.Form frmDateCalulator 
   Caption         =   "learn25 (arnel.c.decastro) Date Calculator Program"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   " Calculate Previous Date "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   7575
      Begin VB.TextBox txtDateBck 
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtBackward 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "Get Date >>"
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtReferenceDate 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "No. of days backward:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Reference Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Calculate the number of days before you reach your ""target"" date.  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   7575
      Begin VB.TextBox txtResult 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   2175
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "Calculate >>"
         Height          =   495
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtEndDate 
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtStartDate 
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Result: (No. of days)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Enter ""Target"" Date: (mm/dd/yyyy)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Enter ""Start"" Date: (mm/dd/yyyy)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Calculate the number of days from Jan. 1, 1 up to the date you entered  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      Begin VB.TextBox txtNoOfDays 
         Height          =   375
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Calculate >>"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtDate1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Result:   (No. of days)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Date: (mm/dd/yyyy)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmDateCalulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd1_Click()
    Dim x
    If IsDate(Trim(txtDate1.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtDate1.Text = ""
       txtNoOfDays.Text = ""
       txtDate1.SetFocus
       Exit Sub
    End If
    x = CountNumberOfDaysFromJan1Year1ToDec31YearEntered(GetYear(Trim(txtDate1.Text)) - 1) + CountDaysInAYear(txtDate1.Text)
    txtNoOfDays.Text = Trim(str(x))
End Sub
Private Sub cmd2_Click()
Dim date1, date2 As Long
    If IsDate(Trim(txtStartDate.Text)) = False Or IsDate(Trim(txtEndDate.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtStartDate.Text = ""
       txtEndDate.Text = ""
       txtResult.SetFocus
       Exit Sub
    End If
   date1 = CountNumberOfDaysFromJan1Year1ToDec31YearEntered(GetYear(Trim(txtStartDate.Text)) - 1) + CountDaysInAYear(txtStartDate.Text)
   date2 = CountNumberOfDaysFromJan1Year1ToDec31YearEntered(GetYear(Trim(txtEndDate.Text)) - 1) + CountDaysInAYear(txtEndDate.Text)
   txtResult.Text = Trim(str(date2 - date1))
End Sub

Private Sub cmd3_Click()
Dim x
If Trim(txtBackward.Text) <> "" And Trim(txtReferenceDate.Text) <> "" Then
    txtDateBck.Text = GetDateNdaysAgoFromReferenceDate(txtReferenceDate.Text, txtBackward.Text)
End If
End Sub

Private Sub txtBackward_LostFocus()
If Trim(txtBackward.Text) <> "" Then
   If IsNumeric(Trim(txtBackward.Text)) = False Then
      MsgBox "Invalid Input. ", vbInformation, "Input Error: "
      txtBackward.Text = ""
      txtBackward.SetFocus
      Exit Sub
   Else
     txtBackward.Text = Int(Val(txtBackward.Text))
   End If
End If
End Sub

Private Sub txtDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmd1_Click
End Sub
Private Sub txtDate1_LostFocus()
 If Trim(txtDate1.Text) <> "" Then
    If IsDate(Trim(txtDate1.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtDate1.Text = ""
       txtNoOfDays.Text = ""
       txtDate1.SetFocus
    End If
 End If
End Sub
Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call cmd2_Click
End Sub
Private Sub txtEndDate_LostFocus()
If Trim(txtEndDate.Text) <> "" Then
    If IsDate(Trim(txtEndDate.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtEndDate.Text = ""
       txtResult.Text = ""
       txtEndDate.SetFocus
    End If
 End If
End Sub
Private Sub txtReferenceDate_LostFocus()
If Trim(txtReferenceDate.Text) <> "" Then
    If IsDate(Trim(txtReferenceDate.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtReferenceDate.Text = ""
       txtDateBck.Text = ""
       txtReferenceDate.SetFocus
    End If
 End If
End Sub
Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmd2_Click
End Sub
Private Sub txtStartDate_LostFocus()
If Trim(txtStartDate.Text) <> "" Then
    If IsDate(Trim(txtStartDate.Text)) = False Then
       MsgBox "Invalid Date", vbInformation, "Input Error: "
       txtStartDate.Text = ""
       txtResult.Text = ""
       txtStartDate.SetFocus
    End If
 End If
End Sub
