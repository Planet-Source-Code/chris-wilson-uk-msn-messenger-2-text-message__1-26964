VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "msn2sms by Chris Wilson"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2232
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":327E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":61A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Test SMS"
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPrefix 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Text            =   "447751"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtNumber 
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Text            =   "219715"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtICQPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtICQNumber 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      Caption         =   "show settings"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Unban"
      Height          =   315
      Left            =   5520
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ban"
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   690
      Left            =   3840
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "send incomming messages"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "send user online status's"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7080
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SMS Log"
         Object.Width           =   5786
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "msn2sms loading ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Mobile number"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ICQ password"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ICQ number"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "please wait"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Msg As MsgrObject
Attribute Msg.VB_VarHelpID = -1

Dim SMSPrefix As String
Dim SMSNumber As String
Dim SMSMessage As String

Private Const fH = 4725
Dim fH2 As Long

Dim ICQNumber As String
Dim ICQPassword As String
Dim MessagesLeft As Integer

Dim SpecialCode As String




Private Sub SendSMS(textMessage As String)
If ICQNumber = "" Then Exit Sub
If ICQPassword = "" Then Exit Sub

Dim aceSTring As String
aceSTring = textMessage
aceSTring = RemoveString(aceSTring, "*** ", "")


ICQNumber = txtICQNumber
ICQPassword = txtICQPassword

SMSPrefix = txtPrefix
SMSNumber = txtNumber

SMSMessage = textMessage

ListView1.ListItems.Add 1, , aceSTring, , 4

lblStatus = "Sending SMS Notification..."
'send the message to the phone number you want
 ret = Inet1.OpenURL("http://web.icq.com/sms/send_history/1,,,00.html?target=msghistory&prefix=+" + SMSPrefix + "&carrier=aaa&tophone=" + SMSNumber + "&msg=" + SMSMessage)

lblStatus = "Waiting for messenger event"



If textMessage = "msn2sms test message" Then
ListView1.ListItems(1).SmallIcon = 2

Exit Sub
End If


If textMessage = "*** You have new hotmail" Then
ListView1.ListItems(1).SmallIcon = 7
Exit Sub
End If


If Mid(textMessage, 1, 3) = "***" Then
ListView1.ListItems(1).SmallIcon = 1
Exit Sub
End If

ListView1.ListItems(1).SmallIcon = 3



End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
fH2 = Form1.Height
Form1.Height = fH
Exit Sub
End If

If Check3.Value = 0 Then
Form1.Height = fH2
Exit Sub
End If


End Sub

Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
List1.AddItem Text1
Text1 = ""
List1.TopIndex = List1.ListCount - 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
ICQNumber = txtICQNumber
ICQPassword = txtICQPassword

lblStatus = "Connecting to SMS centre..."
ListView1.ListItems.Add 1, , "Connecting to SMS centre", , 8
'opens the registry page and say your are online
ret = Inet1.OpenURL("http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + ICQNumber + "&uPassword=" + ICQPassword)
ListView1.ListItems.Item(1).Text = "Connected to SMS centre"
ListView1.ListItems(1).SmallIcon = 9

lblStatus = "Testing SMS Sending..."
SendSMS "msn2sms test message"
End Sub

Private Sub Form_Load()
Form1.Show
Set Msg = New MsgrObject

txtPrefix = GetSetting("MSN2SMS", "Settings", "Prefix", "447123")
txtNumber = GetSetting("MSN2SMS", "Settings", "Number", "123456")
txtICQNumber = GetSetting("MSN2SMS", "Settings", "ICQNumber")
txtICQPassword = GetSetting("MSN2SMS", "Settings", "ICQPassword")


ICQNumber = txtICQNumber
ICQPassword = txtICQPassword

Dim TempSTring As String
TempSTring = Msg.LocalLogonName
TempSTring = Mid(TempSTring, 1, InStr(1, TempSTring, "@") - 1)
TempSTring = RemoveString(TempSTring, "_", " ")

Label6 = "Welcome " & TempSTring


Dim BanCount As Integer
Dim TheX As Integer

BanCount = GetSetting("MSN2SMS", "Bans", "Count", 0)
TheX = -1
If BanCount = 0 Then GoTo 15

Do
TheX = TheX + 1
List1.AddItem GetSetting("MSN2SMS", "Bans", "Ban " & TheX)
If TheX = BanCount - 1 Then GoTo 15
DoEvents
Loop


15 If ICQNumber = "" Then Exit Sub

ListView1.ListItems.Add 1, , "Connecting to SMS centre", , 8
lblStatus = "Connecting to SMS centre..."
'opens the registry page and say your are online
ret = Inet1.OpenURL("http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + ICQNumber + "&uPassword=" + ICQPassword)
ListView1.ListItems.Item(1).Text = "Connected to SMS centre"
ListView1.ListItems(1).SmallIcon = 9

lblStatus = "Waiting for messenger event"
MessagesLeft = 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If List1.ListCount = 0 Then
Debug.Print "BANS TO ADD: " & List1.ListCount
DeleteSetting "MSN2SMS", "Bans"
Exit Sub
End If

DeleteSetting "MSN2SMS", "Bans"

SaveSetting "MSN2SMS", "Bans", "Count", List1.ListCount
Debug.Print "BANS TO ADD: " & List1.ListCount

Dim TheX As Integer
TheX = -1

Do
TheX = TheX + 1
SaveSetting "MSN2SMS", "Bans", "Ban " & TheX, List1.List(TheX)
Debug.Print "BAN ADDED TO REGISTRY: " & List1.List(TheX)
If TheX = List1.ListCount - 1 Then Exit Sub
DoEvents
Loop
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
If State = 11 Then ListView1.ListItems(1).SmallIcon = 6
End Sub

Private Sub Msg_OnTextReceived(ByVal pIMSession As IMsgrIMSession, ByVal pSourceUser As IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
On Error Resume Next

If InStr(1, bstrMsgHeader, "TypingUser") Then Exit Sub
If Not Check2.Value = 1 Then Exit Sub
Dim TheMessage As String

Dim TheX As Integer
On Error GoTo 10
Do
If InStr(1, pSourceUser.EmailAddress, List1.List(TheX)) Then Exit Sub
TheX = TheX + 1
If TheX >= List1.ListCount Then GoTo 10
DoEvents
Loop

10

SendSMS Mid(pSourceUser.EmailAddress & ": " & bstrMsgText, 1, 160)

End Sub

Private Sub Msg_OnUserStateChanged(ByVal pUser As IMsgrUser, ByVal mPrevState As MSTATE, pfEnableDefault As Boolean)
If Not Check1.Value = 1 Then Exit Sub


Dim TheMessage As String

Dim TheX As Integer
Do
If List1.ListCount = 0 Then GoTo 10
If InStr(1, pUser.EmailAddress, List1.List(TheX)) Then Exit Sub
5 TheX = TheX + 1: Debug.Print " X = X + 1"
If TheX >= List1.ListCount Then GoTo 10
DoEvents
Loop

10

If pUser.State = MSTATE_ONLINE Then SendSMS "*** " & pUser.EmailAddress & " has come online"
If pUser.State = MSTATE_OFFLINE Then SendSMS "*** " & pUser.EmailAddress & " has gone offline"
End Sub

Private Sub txtICQNumber_Change()
SaveSetting "MSN2SMS", "Settings", "ICQNumber", txtICQNumber
End Sub

Private Sub txtICQPassword_Change()
SaveSetting "MSN2SMS", "Settings", "ICQPassword", txtICQPassword
End Sub

Private Sub txtNumber_Change()
SaveSetting "MSN2SMS", "Settings", "Number", txtNumber
End Sub

Private Sub txtPrefix_Change()
SaveSetting "MSN2SMS", "Settings", "Prefix", txtPrefix
End Sub

Public Function RemoveString(Entire As String, Word As String, Replace As String) As String
    Dim I As Integer
    I = 1
    Dim LeftPart
    Do While True
        I = InStr(1, Entire, Word)
        If I = 0 Then
            Exit Do
        Else
            LeftPart = Left(Entire, I - 1)
            Entire = LeftPart & Replace & Right(Entire, Len(Entire) - Len(Word) - Len(LeftPart))
        End If
    Loop
    
   RemoveString = Entire
   
End Function

Private Sub Msg_OnUnreadEmailChanged(ByVal MFOLDER As MFOLDER, ByVal cUnreadEmail As Long, pfEnableDefault As Boolean)
If Not Check1.Value = 1 Then Exit Sub
SendSMS "*** You have new hotmail"
End Sub

Private Sub Msg_OnLogonResult(ByVal hr As Long, ByVal pService As IMsgrService)
If pService.Status = MSS_LOGGED_ON Then

ListView1.ListItems.Add 1, , "Connecting to SMS centre", , 8
lblStatus = "Connecting to SMS centre..."
'opens the registry page and say your are online
ret = Inet1.OpenURL("http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + ICQNumber + "&uPassword=" + ICQPassword)
ListView1.ListItems.Item(1).Text = "Connected to SMS centre"
ListView1.ListItems(1).SmallIcon = 9

lblStatus = "Waiting for messenger event"
End If


End Sub
