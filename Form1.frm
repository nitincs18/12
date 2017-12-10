VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton findbtn 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5400
      TabIndex        =   26
      Top             =   1800
      Width           =   1500
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Sitka Subheading"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1320
      TabIndex        =   25
      Text            =   "STUDENT PROFILE"
      Top             =   240
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton uploadbtn 
      Caption         =   "UPLOAD PICTURE"
      Height          =   550
      Left            =   8640
      TabIndex        =   24
      Top             =   4920
      Width           =   1740
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   550
      Left            =   7320
      TabIndex        =   23
      Top             =   6480
      Width           =   1500
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
      Height          =   550
      Left            =   5280
      TabIndex        =   22
      Top             =   6480
      Width           =   1500
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   550
      Left            =   3360
      TabIndex        =   21
      Top             =   6480
      Width           =   1500
   End
   Begin VB.CommandButton addnew 
      Caption         =   "ADD NEW"
      Height          =   550
      Left            =   1680
      TabIndex        =   20
      Top             =   6480
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   7920
      ScaleHeight     =   2715
      ScaleWidth      =   2835
      TabIndex        =   19
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   400
      Left            =   2880
      TabIndex        =   18
      Top             =   5760
      Width           =   2400
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   2880
      TabIndex        =   17
      Top             =   5280
      Width           =   2400
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Text            =   "Select Semester"
      Top             =   4920
      Width           =   2400
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2880
      TabIndex        =   15
      Text            =   "Select Course"
      Top             =   4560
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   14
      Text            =   "Select Department"
      Top             =   4080
      Width           =   2400
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4080
      TabIndex        =   13
      Top             =   3600
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2760
      TabIndex        =   12
      Top             =   3600
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   400
      Left            =   2760
      TabIndex        =   11
      Top             =   3000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      _Version        =   393216
      Format          =   96927745
      CurrentDate     =   43078
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   2760
      TabIndex        =   10
      Top             =   2400
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label Label9 
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   8
      Top             =   5880
      Width           =   1400
   End
   Begin VB.Label Label8 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   7
      Top             =   5400
      Width           =   1400
   End
   Begin VB.Label Label7 
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   6
      Top             =   4920
      Width           =   1400
   End
   Begin VB.Label Label6 
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   5
      Top             =   4560
      Width           =   1400
   End
   Begin VB.Label Label5 
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   4
      Top             =   4080
      Width           =   1400
   End
   Begin VB.Label Label4 
      Caption         =   "Gender "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   3
      Top             =   3600
      Width           =   1400
   End
   Begin VB.Label Label3 
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1400
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   1400
   End
   Begin VB.Label Label1 
      Caption         =   "Roll No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub addnew_Click()
rs.addnew
clear
End Sub
Sub clear()
Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = "10/05/2005"
Option1.Value = False
Option2.Value = False
Combo1.Text = "Select Department"
Combo2.Text = "Select Course"
Combo3.Text = "Select Semester"
Text3.Text = ""
Text4.Text = ""
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Combo1_Click()
Combo2.clear
If Combo1.Text = "Computer Science" Then
Combo2.AddItem "M.C.A"
Combo2.AddItem "B.C.A"
Combo2.AddItem "B.Sc(IT)"
ElseIf Combo1.Text = "Electrical Engineering" Then
Combo2.AddItem "B.TECH (EE)"
Combo2.AddItem "M.TECH (EE)"
ElseIf Combo1.Text = "Civil Engineering" Then
Combo2.AddItem "B.TECH (CE)"
Combo2.AddItem "M.TECH (CE)"
Else
End If
End Sub

Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to delete the Student Profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "Message"
rs.Update
refreshdata
Else
MsgBox "Profile Not Deleted ..!!", vbInformation, "Message"
End If
End Sub
Sub refreshdata()
rs.Close
rs.Open "Select * from ProfileTBL", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub

Private Sub findbtn_Click()
rs.Close
rs.Open "Select * from ProfileTBL where RollNo='" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found ..!!", vbInformation
End If

End Sub
Sub reload()
rs.Close
rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\Database Folder\ProfileDB.mdb;Persist Security Info=False"
rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "Computer Science"
Combo1.AddItem "Electrical Engineering"
Combo1.AddItem "Civil Engineering"
Combo3.AddItem "SEMESTER-I"
Combo3.AddItem "SEMESTER-II"
Combo3.AddItem "SEMESTER-III"
Combo3.AddItem "SEMESTER-IV"
Combo3.AddItem "SEMESTER-V"
Combo3.AddItem "SEMESTER-VI"
Combo3.AddItem "SEMESTER-VII"
Combo3.AddItem "SEMESTER-VIII"

End Sub
Sub display()
Text1.Text = rs!Rollno
Text2.Text = rs!Name
DTPicker1.Value = rs!DOB
If rs!Gender = "MALE" Then
Option1.Value = True
Else
Option2.Value = True
End If
Combo1.Text = rs!Dept
Combo2.Text = rs!Course
Combo3.Text = rs!Semester
Text3.Text = rs!Address
Text4.Text = rs!phone
Picture1.Picture = LoadPicture(rs!photo)
End Sub

Private Sub savebtn_Click()
rs.Fields("RollNo").Value = Text1.Text
rs.Fields("Name").Value = Text2.Text
rs.Fields("DOB").Value = DTPicker1.Value
If Option1.Value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Dept").Value = Combo1.Text
rs.Fields("Course").Value = Combo2.Text
rs.Fields("Semester").Value = Combo3.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("Phone").Value = Text4.Text
rs.Fields("Photo").Value = str
MsgBox "Data is saved successfully ..!!!", vbInformation
rs.Update
End Sub

Private Sub updatebtn_Click()
rs.Fields("RollNo").Value = Text1.Text
rs.Fields("Name").Value = Text2.Text
rs.Fields("DOB").Value = DTPicker1.Value
If Option1.Value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Dept").Value = Combo1.Text
rs.Fields("Course").Value = Combo2.Text
rs.Fields("Semester").Value = Combo3.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("Phone").Value = Text4.Text
MsgBox "Data is updated successfully ..!!!", vbInformation
rs.Update
End Sub

Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub
