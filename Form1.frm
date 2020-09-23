VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmWho 
   Caption         =   "Whois"
   ClientHeight    =   4605
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar mainToolbar 
      Height          =   330
      Left            =   3480
      TabIndex        =   38
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ToolbarImagesYes"
      DisabledImageList=   "ToolbarImagesNo"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2535
      TabIndex        =   34
      Top             =   480
      Width           =   2535
      Begin RichTextLib.RichTextBox MailResponse 
         Height          =   135
         Left            =   1680
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   238
         _Version        =   393217
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":0442
      End
      Begin RichTextLib.RichTextBox txtResponse 
         Height          =   495
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":04C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   855
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   6421
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Handle"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "E-mail"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fax"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label LenLabel 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   4680
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servers"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
      Begin VB.TextBox MainSer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "whois.networksolutions.com"
         Top             =   480
         Width           =   3975
      End
      Begin VB.OptionButton OpM 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Cut 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   3360
         Width           =   4215
      End
      Begin VB.TextBox Ser8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Dom8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Ser7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox Dom7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Dom6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Ser6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   2280
         Width           =   3135
      End
      Begin VB.TextBox Dom5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Ser5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox Dom4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Ser4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Dom3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Ser3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Dom2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Ser2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Dom1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Ser1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.OptionButton Op1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   1080
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton Op2 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Op3 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Op4 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Op5 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   28
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Op6 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Op7 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton Op8 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Text to cut out of the responce"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3150
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Domain"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Main Server"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ToolbarImagesNo 
      Left            =   4680
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0547
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C3F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ToolbarImagesYes 
      Left            =   4680
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":133B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18D7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Host"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   495
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu menSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu menQuit 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu menShow 
      Caption         =   "&Show"
      Begin VB.Menu menresp 
         Caption         =   "Response"
         Checked         =   -1  'True
      End
      Begin VB.Menu menSer 
         Caption         =   "Servers"
      End
      Begin VB.Menu mentxt 
         Caption         =   "txt to Remove"
      End
      Begin VB.Menu space5 
         Caption         =   "-"
      End
      Begin VB.Menu menAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu menOPtions 
      Caption         =   "&Options"
      Begin VB.Menu mail 
         Caption         =   "&Mail Box"
      End
      Begin VB.Menu Empl 
         Caption         =   "&Employe Records"
      End
      Begin VB.Menu menhst 
         Caption         =   "&HST Query"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu menhelp 
         Caption         =   "&Help"
      End
   End
   Begin VB.Menu ListBoxList 
      Caption         =   "ListBoxList"
      Visible         =   0   'False
      Begin VB.Menu EmployRecNEW 
         Caption         =   "Employee record [new]"
      End
      Begin VB.Menu EmployRecSAME 
         Caption         =   "Employee record [same]"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu CopyN 
         Caption         =   "Copy Name"
      End
      Begin VB.Menu CopyH 
         Caption         =   "Copy Handle"
      End
      Begin VB.Menu CopyE 
         Caption         =   "Copy E-mail"
      End
      Begin VB.Menu CopyP 
         Caption         =   "Copy Phone number"
      End
      Begin VB.Menu CopyF 
         Caption         =   "Copy Fax number"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu menViewTXT 
         Caption         =   "View Text"
      End
   End
End
Attribute VB_Name = "FrmWho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ser2use As String
Dim OldData As String
Dim hist As New Collection
Dim ons
Dim MaxMailLen As String

Private Sub CopyE_Click()
Clipboard.SetText ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub CopyF_Click()
Clipboard.SetText ListView1.SelectedItem.SubItems(4)
End Sub

Private Sub CopyH_Click()
Clipboard.SetText ListView1.SelectedItem.SubItems(1)
End Sub

Private Sub CopyN_Click()
Clipboard.SetText ListView1.SelectedItem.Text
End Sub

Private Sub CopyP_Click()
Clipboard.SetText ListView1.SelectedItem.SubItems(3)
End Sub

Private Sub Cut_Change()
If Op1.Value = True Then
    Op1.Tag = Cut
ElseIf Op2.Value = True Then
    Op2.Tag = Cut
ElseIf Op3.Value = True Then
    Op3.Tag = Cut
ElseIf Op4.Value = True Then
    Op4.Tag = Cut
ElseIf Op5.Value = True Then
    Op5.Tag = Cut
ElseIf Op6.Value = True Then
    Op6.Tag = Cut
ElseIf Op7.Value = True Then
    Op7.Tag = Cut
ElseIf Op8.Value = True Then
    Op8.Tag = Cut
ElseIf OpM.Value = True Then
    OpM.Tag = Cut
End If
End Sub

Private Sub empl_Click()
Dim emplo  As String
emplo = InputBox("Check an employee Record" & vbCrLf & "enter the employee code e.g.(BGY139) found in the mail box lookup", "Employee Record")
If emplo <> "" Then
    Call go("HANDLE " & emplo, "", True)
End If
End Sub

Private Sub EmployRecNEW_Click()
Dim newWho As New FrmWho
Dim dat  As String
dat = ListView1.SelectedItem.SubItems(1)
If dat = "" Then Exit Sub
dat = Replace(dat, "(", "")
dat = Replace(dat, ")", "")
newWho.go "Handle " & dat, "", True
newWho.Show
End Sub

Private Sub EmployRecSAME_Click()
Dim dat  As String
dat = ListView1.SelectedItem.SubItems(1)
If dat = "" Then Exit Sub
dat = Replace(dat, "(", "")
dat = Replace(dat, ")", "")
Call go("Handle " & dat, "", True)
End Sub

Private Sub Form_Load()
MainSer = GetSetting(App.Title, "Settings", "MainServer", "whois.networksolutions.com")
Dom1 = GetSetting(App.Title, "Settings", "Domain1", ".com")
Dom2 = GetSetting(App.Title, "Settings", "Domain2", "")
Dom3 = GetSetting(App.Title, "Settings", "Domain3", "")
Dom4 = GetSetting(App.Title, "Settings", "Domain4", "")
Dom5 = GetSetting(App.Title, "Settings", "Domain5", "")
Dom6 = GetSetting(App.Title, "Settings", "Domain6", "")
Dom7 = GetSetting(App.Title, "Settings", "Domain7", "")
Dom8 = GetSetting(App.Title, "Settings", "Domain8", "")
Ser1 = GetSetting(App.Title, "Settings", "Server1", "whois.networksolutions.com")
Ser2 = GetSetting(App.Title, "Settings", "Server2", "")
Ser3 = GetSetting(App.Title, "Settings", "Server3", "")
Ser4 = GetSetting(App.Title, "Settings", "Server4", "")
Ser5 = GetSetting(App.Title, "Settings", "Server5", "")
Ser6 = GetSetting(App.Title, "Settings", "Server6", "")
Ser7 = GetSetting(App.Title, "Settings", "Server7", "")
Ser8 = GetSetting(App.Title, "Settings", "Server8", "")
OpM.Tag = GetSetting(App.Title, "Settings", "Cutm", "The Data in Network Solutions' WHOIS database is provided by Network" & vbCrLf & _
"Solutions for information purposes, and to assist persons in obtaining" & vbCrLf & _
"information about or related to a domain name registration record." & vbCrLf & _
"Network Solutions does not guarantee its accuracy.  By submitting a" & vbCrLf & _
"WHOIS query, you agree that you will use this Data only for lawful" & vbCrLf & _
"purposes and that, under no circumstances will you use this Data to:" & vbCrLf & _
"(1) allow, enable, or otherwise support the transmission of mass" & vbCrLf & _
"unsolicited, commercial advertising or solicitations via e-mail" & vbCrLf & _
"(spam); or  (2) enable high volume, automated, electronic processes" & vbCrLf & _
"that apply to Network Solutions (or its systems).  Network Solutions" & vbCrLf & _
"reserves the right to modify these terms at any time.  By submitting" & vbCrLf & _
"this query, you agree to abide by this policy.")
Op1.Tag = GetSetting(App.Title, "Settings", "Cut1", OpM.Tag)
Op2.Tag = GetSetting(App.Title, "Settings", "Cut2", "")
Op3.Tag = GetSetting(App.Title, "Settings", "Cut3", "")
Op4.Tag = GetSetting(App.Title, "Settings", "Cut4", "")
Op5.Tag = GetSetting(App.Title, "Settings", "Cut5", "")
Op6.Tag = GetSetting(App.Title, "Settings", "Cut6", "")
Op7.Tag = GetSetting(App.Title, "Settings", "Cut7", "")
Op8.Tag = GetSetting(App.Title, "Settings", "Cut8", "")
Call Op1_Click
ons = 0


   txtSearch = Command$
   If Trim(txtSearch) <> "" Then Call go(txtSearch, "", True)
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 3795 Then Me.Width = 3795
If Me.Height < 4440 Then Me.Height = 4440
txtSearch.Width = Me.ScaleWidth - txtSearch.Left - 120 - mainToolbar.Width - 60
mainToolbar.Left = txtSearch.Left + txtSearch.Width + 60
Picture1.Width = Me.ScaleWidth - 240
Picture1.Height = Me.ScaleHeight - Picture1.Top - 120
ListView1.Width = Picture1.Width
ListView1.Height = Picture1.Height
txtResponse.Width = Picture1.Width
txtResponse.Height = Picture1.Height
Frame1.Height = txtResponse.Height
Frame1.Width = txtResponse.Width
Frame1.Top = Picture1.Top
MainSer.Width = Frame1.Width - 240 - OpM.Width
Ser1.Width = Frame1.Width - Ser1.Left - 375
Op1.Left = Ser1.Left + Ser1.Width
OpM.Left = Op1.Left
Ser2.Width = Ser1.Width
Op2.Left = Ser2.Left + Ser2.Width
Ser3.Width = Ser1.Width
Op3.Left = Ser3.Left + Ser3.Width
Ser4.Width = Ser1.Width
Op4.Left = Ser4.Left + Ser4.Width
Ser5.Width = Ser1.Width
Op5.Left = Ser5.Left + Ser5.Width
Ser6.Width = Ser1.Width
Op6.Left = Ser6.Left + Ser6.Width
Ser7.Width = Ser1.Width
Op7.Left = Ser7.Left + Ser7.Width
Ser8.Width = Ser1.Width
Op8.Left = Ser8.Left + Ser8.Width
Cut.Height = Frame1.Height - Cut.Top - 120
Cut.Width = MainSer.Width + OpM.Width
SizeListView
'listview shit
End Sub
Sub SizeListView()
ListView1.ColumnHeaders(1).Width = ListView1.Width - _
ListView1.ColumnHeaders(2).Width - _
ListView1.ColumnHeaders(3).Width - _
ListView1.ColumnHeaders(4).Width - _
ListView1.ColumnHeaders(5).Width - 350
End Sub
Public Sub go(host As String, server As String, UseHis)

Dim lFirstArg  As String
Dim lSecondArg As String
Dim lPos      As Integer

If LCase(Left(host, 7)) = "http://" Then host = Right(host, Len(host) - 7)
If LCase(Left(host, 4)) = "www." Then host = Right(host, Len(host) - 4)
txtSearch = host

        '#\/##\/##\/##\/##\/#History shit#\/##\/##\/##\/##\/#
        If UseHis = True Then
            If ons = 0 Then
                hist.Add host
                ons = ons + 1
            ElseIf hist(ons) <> host Then
                hist.Add host, , , ons
                ons = ons + 1
                If hist.Count <> ons Then
                    deleting = hist.Count
                    Do Until deleting = ons
                        hist.Remove (deleting)
                        deleting = deleting - 1
                    Loop
                    mainToolbar.Buttons(3).Enabled = False 'dissable falward
                    mainToolbar.Buttons(3).ToolTipText = "Folward"
                End If
                mainToolbar.Buttons(2).Enabled = True 'enable back
                mainToolbar.Buttons(2).ToolTipText = "Back [" & hist(ons - 1) & "]"
            End If
        End If
        '#^###^###^###^###^#History shit#^###^###^###^###^#

    lPos = InStr(1, txtSearch, " ")
    If lPos = 0 Then
        ' If ipos is zero then there was only one argument found
        lFirstArg = Trim(txtSearch)
    Else
        lFirstArg = Mid(txtSearch, 1, lPos - 1)
        ' Now parse the command line for the second argument
        lSecondArg = Mid(txtSearch, lPos + 1, Len(txtSearch))
    End If
    host = lFirstArg
    server = lSecondArg
    txtSearch.Locked = True
    mainToolbar.Buttons(1).Enabled = False
    txtResponse = ""
    OldData = ""
    ListView1.ListItems.Clear
    MaxMailLen = 1440
    If Left(host, 1) = "@" Or Right(host, 1) = "@" Then
        txtResponse.Visible = False
        ListView1.Visible = True
    Else
        txtResponse.Visible = True
        ListView1.Visible = False
    End If
    Winsock.Close
    Winsock.LocalPort = 0
    Me.Caption = "whois [connecting]"
    If server = "" Then
        If Dom1.Text <> "" And Right(txtSearch.Text, Len(Dom1.Text)) = Dom1.Text Then
            Winsock.Connect Ser1, 43
            Ser2use = 1
        ElseIf Dom2.Text <> "" And Right(txtSearch.Text, Len(Dom2.Text)) = Dom2.Text Then
            Winsock.Connect Ser2, 43
            Ser2use = 2
        ElseIf Dom3.Text <> "" And Right(txtSearch.Text, Len(Dom3.Text)) = Dom3.Text Then
            Winsock.Connect Ser3, 43
            Ser2use = 3
        ElseIf Dom4.Text <> "" And Right(txtSearch.Text, Len(Dom4.Text)) = Dom4.Text Then
            Winsock.Connect Ser4, 43
            Ser2use = 4
        ElseIf Dom5.Text <> "" And Right(txtSearch.Text, Len(Dom5.Text)) = Dom5.Text Then
            Winsock.Connect Ser5, 43
            Ser2use = 5
        ElseIf Dom6.Text <> "" And Right(txtSearch.Text, Len(Dom6.Text)) = Dom6.Text Then
            Winsock.Connect Ser6, 43
            Ser2use = 6
        ElseIf Dom7.Text <> "" And Right(txtSearch.Text, Len(Dom7.Text)) = Dom7.Text Then
            Winsock.Connect Ser7, 43
            Ser2use = 7
        ElseIf Dom8.Text <> "" And Right(txtSearch.Text, Len(Dom8.Text)) = Dom8.Text Then
            Winsock.Connect Ser8, 43
            Ser2use = 8
        Else
            Ser2use = 0
            Winsock.Connect MainSer, 43
        End If
    Else
        'do server
        Winsock.Connect server, 43
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.Title, "Settings", "MainServer", MainSer
SaveSetting App.Title, "Settings", "Domain1", Dom1
SaveSetting App.Title, "Settings", "Domain2", Dom2
SaveSetting App.Title, "Settings", "Domain3", Dom3
SaveSetting App.Title, "Settings", "Domain4", Dom4
SaveSetting App.Title, "Settings", "Domain5", Dom5
SaveSetting App.Title, "Settings", "Domain6", Dom6
SaveSetting App.Title, "Settings", "Domain7", Dom7
SaveSetting App.Title, "Settings", "Domain8", Dom8
SaveSetting App.Title, "Settings", "Server1", Ser1
SaveSetting App.Title, "Settings", "Server2", Ser2
SaveSetting App.Title, "Settings", "Server3", Ser3
SaveSetting App.Title, "Settings", "Server4", Ser4
SaveSetting App.Title, "Settings", "Server5", Ser5
SaveSetting App.Title, "Settings", "Server6", Ser6
SaveSetting App.Title, "Settings", "Server7", Ser7
SaveSetting App.Title, "Settings", "Server8", Ser8
SaveSetting App.Title, "Settings", "Cutm", OpM.Tag
SaveSetting App.Title, "Settings", "Cut1", Op1.Tag
SaveSetting App.Title, "Settings", "Cut2", Op2.Tag
SaveSetting App.Title, "Settings", "Cut3", Op3.Tag
SaveSetting App.Title, "Settings", "Cut4", Op4.Tag
SaveSetting App.Title, "Settings", "Cut5", Op5.Tag
SaveSetting App.Title, "Settings", "Cut6", Op6.Tag
SaveSetting App.Title, "Settings", "Cut7", Op7.Tag
SaveSetting App.Title, "Settings", "Cut8", Op8.Tag
End Sub

Private Sub ListView1_DblClick()
EmployRecNEW_Click
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  PopupMenu ListBoxList
End If
End Sub

Private Sub mail_Click()
Dim mailBo As String
If Left(txtSearch, 1) = "@" Then
    mailBo = InputBox("Check the e-mail address on that server (normaly the first 50)", "Mail Box", Right(txtSearch, Len(txtSearch) - 1))
Else
    mailBo = InputBox("Check the e-mail address on that server (normaly the first 50)", "Mail Box", txtSearch)
End If
If mailBo <> "" Then
    Call go("@" & mailBo, "", True)
End If
End Sub

Private Sub mainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Go"
            Call go(txtSearch, "", True)
        Case "Forward"
            ons = ons + 1
            txtSearch.Text = hist(ons)
            mainToolbar.Buttons(2).Enabled = True 'enable back
            If ons = hist.Count Then
                mainToolbar.Buttons(3).Enabled = False 'dissable folward
                mainToolbar.Buttons(3).ToolTipText = "Folward"
            Else
                mainToolbar.Buttons(3).ToolTipText = "Folward [" & hist(ons + 1) & "]"
            End If
            mainToolbar.Buttons(2).ToolTipText = "Back [" & hist(ons - 1) & "]"
            Call go(txtSearch, "", False)
        Case "Back"
            ons = ons - 1
            txtSearch.Text = hist(ons)
            mainToolbar.Buttons(3).Enabled = True 'enable folward
            If ons = 1 Then
                mainToolbar.Buttons(2).Enabled = False 'dissable back
                mainToolbar.Buttons(2).ToolTipText = "Back"
            Else
                mainToolbar.Buttons(2).ToolTipText = "Back [" & hist(ons - 1) & "]"
            End If
            mainToolbar.Buttons(3).ToolTipText = "Folward [" & hist(ons + 1) & "]"
            Call go(txtSearch, "", False)
    End Select
End Sub

Private Sub menAbout_Click()
MsgBox "Programed by Syphen" & vbCrLf & "Syphen_2k@hotmail.com" & vbCrLf & "www.hackuk.net", vbOKOnly, "About"
If MsgBox("do you want to to go to the site?", vbYesNo, "?") = vbYes Then
On Error Resume Next
Dim Web_WWW As Long
Dim WebPage As String
WebPage = "http://www.hackuk.net"
Web_WWW = ShellExecute(Me.hwnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)
End If
End Sub

Private Sub menhelp_Click()
Call go("help", "whois.networksolutions.com", True)
End Sub


Private Sub menhst_Click()
Dim hst  As String
hst = InputBox("Search for info on a server, enter server handle e.g. NS99999", "Mail Box")
If hst <> "" Then
    Call go("SERVER " & hst, "", True)
End If
End Sub

Private Sub menQuit_Click()
Unload Me
End Sub

Private Sub menresp_Click()
If menresp.Checked = False Then
    menresp.Checked = True
    menSer.Checked = False
    Picture1.Visible = True
    Frame1.Visible = False
End If
End Sub

Private Sub menSave_Click()
CommonDialog1.Filter = "Text File (*.txt)|*.txt"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtResponse.Text
    Close #1
End If
End Sub

Private Sub menSer_Click()
If menSer.Checked = False Then
    menSer.Checked = True
    menresp.Checked = False
    Picture1.Visible = False
    Frame1.Visible = True
End If
End Sub

Private Sub mentxt_Click()
Shell "C:\winnt\notepad.exe " & App.Path & "/T2CUT.txt", vbNormalFocus
End Sub

Private Sub menViewTXT_Click()
txtResponse.Visible = True
ListView1.Visible = False
End Sub

Private Sub Op1_Click()
Cut = Op1.Tag
End Sub
Private Sub Op2_Click()
Cut = Op2.Tag
End Sub
Private Sub Op3_Click()
Cut = Op3.Tag
End Sub
Private Sub Op4_Click()
Cut = Op4.Tag
End Sub
Private Sub Op5_Click()
Cut = Op5.Tag
End Sub
Private Sub Op6_Click()
Cut = Op6.Tag
End Sub
Private Sub Op7_Click()
Cut = Op7.Tag
End Sub
Private Sub Op8_Click()
Cut = Op8.Tag
End Sub

Private Sub OpM_Click()
Cut = OpM.Tag
End Sub

Private Sub txtResponse_Change()
txtResponse = ClearUnwantedTxt(txtResponse)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call go(txtSearch, "", True)
End Sub

Private Sub Winsock_Close()
Me.Caption = "whois"
If ListView1.Visible = True And txtSearch.Locked = True Then
    MailResponse.Text = txtResponse.Text
    With MailResponse
        Dim ATStart As String
        'check there is a mailbox
        .Find "No match for ", 0
        If .SelText = "No match for " Then
                ListView1.ListItems.Add , , "No mailbox found"
                txtSearch.Locked = False
                mainToolbar.Buttons(1).Enabled = True
                Exit Sub
        End If
        .Text = Replace(.Text, vbCrLf & Chr(9) & Chr(9) & Chr(9) & Chr(9), "*")
        .Find txtSearch.Text, 0
        Do Until LCase(.SelText) <> LCase(txtSearch) And .SelText <> UCase(txtSearch)
            ATStart = .SelStart + .SelLength
            .SelStart = ATStart
            .SelLength = 1
            If .SelText = "." Or Right(txtSearch, 1) = "@" Then
                .Find vbCrLf, ATStart
                ATend = .SelStart
                Dim TABs As String
                Dim SPACEs As String
                .Find Chr(9), ATStart, ATend
                If .SelText = Chr(9) Then TABs = .SelStart
                .Find " ", ATStart, ATend
                If .SelText = " " Then SPACEs = .SelStart
                If SPACEs < TABs Then 'space is closer
                    .SelText = "*"
                ElseIf SPACEs > TABs Then 'tab is closer
                    .Find Chr(9), ATStart, ATend
                    .SelText = "*"
                End If
            Else
                .SelText = "*"
            End If
            .Find txtSearch, ATStart
        Loop
        zString = .Text
        zString = Replace(zString, vbCrLf & Chr(9) & Chr(9) & Chr(9) & Chr(9), "*")
        zString = Replace(zString, Chr(9) & Chr(9), "*")
        zString = Replace(zString, Chr(9) & Chr(9) & Chr(9), "*")
        zString = Replace(zString, "(FAX)", "*")
        zString = Replace(zString, "fax:", "*")
        zString = Replace(zString, Chr(9), "")
        zString = Replace(zString, " ", "")
        zString = Replace(zString, ",", ", ")
        zString = Replace(zString, vbCrLf & "*", "*")
        zString = Replace(zString, "**", "*")
        zString = Replace(zString, "**", "*")
        .Text = zString
        Dim Sstart As String
        Dim Send As String
        Dim LineInfo As String
        Sstart = 0
        .Find vbCrLf, 0
        Do Until .SelText <> vbCrLf
            Send = .SelStart
            .SelStart = Sstart
            .SelLength = Send - Sstart
            LineInfo = .SelText
            LineInfo = Replace(LineInfo, "(", "*(", , 1)
            LineInfo = Replace(LineInfo, ")", ")*", , 1)
            LineInfo = Replace(LineInfo, "***", "*")
            LineInfo = Replace(LineInfo, "**", "*")
On Error Resume Next
            LenLabel.Font = ListView1.Font
            LenLabel.Caption = Split(LineInfo, "*")(2)
            If LenLabel.Width > MaxMailLen Then
                MaxMailLen = LenLabel.Width
                ListView1.ColumnHeaders(3).Width = MaxMailLen + 250
                SizeListView
            End If
            Dim bleh As ListItem
            Set bleh = ListView1.ListItems.Add(, , Split(LineInfo, "*")(0))
            For i = 1 To 4
                bleh.SubItems(i) = Split(LineInfo, "*")(i)
            Next i
            Sstart = .SelStart + .SelLength + Len(vbCrLf)
            .Find vbCrLf, Sstart
        Loop
    End With
End If
txtSearch.Locked = False
mainToolbar.Buttons(1).Enabled = True
End Sub

Private Sub Winsock_Connect()
Me.Caption = "whois [downloading]"

lPos = InStr(1, txtSearch, " ")
If lPos = 0 Then
    Winsock.SendData Trim(txtSearch) & vbCrLf
Else
    Winsock.SendData Mid(txtSearch, 1, lPos - 1) & vbCrLf
End If

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String
Winsock.GetData strData
strData = Replace(strData, Chr(10), vbCrLf)
OldData = OldData & strData
OldData = ClearUnwantedTxt(OldData)
txtResponse = OldData
End Sub

Function ClearUnwantedTxt(Data)
On Error GoTo err1
If Ser2use = 1 Then
    Data = Replace(Data, Op1.Tag, "")
ElseIf Ser2use = 2 Then
    Data = Replace(Data, Op2.Tag, "")
ElseIf Ser2use = 3 Then
    Data = Replace(Data, Op3.Tag, "")
ElseIf Ser2use = 4 Then
    Data = Replace(Data, Op4.Tag, "")
ElseIf Ser2use = 5 Then
    Data = Replace(Data, Op5.Tag, "")
ElseIf Ser2use = 6 Then
    Data = Replace(Data, Op6.Tag, "")
ElseIf Ser2use = 7 Then
    Data = Replace(Data, Op7.Tag, "")
ElseIf Ser2use = 8 Then
    Data = Replace(Data, Op8.Tag, "")
ElseIf Ser2use = 0 Then
    Data = Replace(Data, OpM.Tag, "")
End If
err1:
Data = Replace(Data, "The ARIN Registration Services Host contains ONLY Internet", "")
Data = Replace(Data, "Network Information: Networks, ASN's, and related POC's.", "")

Data = Replace(Data, "Please use the whois server at rs.internic.net for DOMAIN related", "")
Data = Replace(Data, "Information and whois.nic.mil for NIPRNET Information.", "")
Data = Replace(Data, "To single out one record, look it up with ""!xxx"", where xxx is the", "")
Data = Replace(Data, "handle, shown in parenthesis following the name, which comes first.", "")
Data = Replace(Data, "Aborting search 50 records found .....", "")

ClearUnwantedTxt = Data
End Function

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Description <> "The connection is reset by remote side" Then
    MsgBox Description, vbOKOnly, "Winsock Error"
    txtSearch.Locked = False
    mainToolbar.Buttons(1).Enabled = True
    Me.Caption = "whois"
End If
End Sub
