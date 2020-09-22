VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   9840
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox txtMessages 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "frmMain.frx":0000
      Top             =   3480
      Width           =   7935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800080&
      Caption         =   "10"
      Height          =   372
      Index           =   0
      Left            =   6960
      TabIndex        =   15
      Top             =   2280
      Value           =   -1  'True
      Width           =   492
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800080&
      Caption         =   "25"
      Height          =   372
      Index           =   1
      Left            =   7560
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800080&
      Caption         =   "50"
      Height          =   372
      Index           =   2
      Left            =   8160
      TabIndex        =   13
      Top             =   2280
      Width           =   492
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800080&
      Caption         =   "No Limit"
      Height          =   372
      Index           =   3
      Left            =   8640
      TabIndex        =   12
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00800080&
      Caption         =   "All Types"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00800080&
      Caption         =   "Gif/Jpeg"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   2280
      Value           =   1  'Checked
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800080&
      Caption         =   "Text/Html"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1092
   End
   Begin VB.TextBox txtDir 
      Height          =   288
      Left            =   4440
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
   End
   Begin VB.DirListBox Dir1 
      Height          =   6165
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtWebsite 
      Height          =   288
      Left            =   4440
      TabIndex        =   0
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Begin Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   7200
      X2              =   10320
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   5400
      X2              =   2520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Maximum No of Files to be downloaded"
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
      Left            =   6720
      TabIndex        =   11
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Select Types of files to be downloaded"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2520
      X2              =   10440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "Directory Name:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9240
      Picture         =   "frmMain.frx":0004
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "                            This application will connect to a website and download the source files."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "    WebSite Replicator"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800080&
      Caption         =   "WebSite Address:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Boolean
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAcessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nServerPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, byValReferer As String, ByVal lpszAcceptTypes As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal lpszheaders As String, ByVal dwHeadersLenght As Long, ByVal lpOptional As String, ByVal dwOptionalLength As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long) As Boolean
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal address As String, ByVal headers As String, ByVal headlenght As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Dim url(400) As String
Dim levu(1000) As String
Dim xz As Integer
Dim oo As Integer
Dim opt1 As Boolean
Dim opt2 As Boolean
Dim opt3 As Boolean
Dim o As Integer
Dim ooo As Integer
Dim levl(1000) As String
Dim strDurl As String
Dim exitproc As Boolean
Dim msize As Long
Dim b As Boolean
Dim f As Boolean
Dim files As Integer
Dim hInternet As Long
Dim hConnect As Long
Dim strServer As String
Dim iPort As Integer
Dim bRes As Boolean
Dim lFlags As Long
Dim hRequest As Long
Dim strURL As String
Dim strBuffer As String * 1
Dim strDir As String
Dim strFile As String
Dim strMurl As String
Dim appdir As String
Dim files1 As Integer
Const INTERNET_FLAG_NO_COOKIES = &H80000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const INTERNET_SERVICE_HTTP = 3
Private Sub cmdConnect_Click()
cmdHangup.Enabled = True
b = InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0)
End Sub

Private Sub cmdHangup_Click()
f = InternetAutodialHangup(0)

End Sub

Private Sub cmdBack_Click()
Load frmUrl
frmUrl.Text1.Text = txtWebsite.Text
frmUrl.Text2.Text = txtDir.Text
Unload Me
frmUrl.Show

End Sub

Private Sub Label6_Click()
If Option1(0).Value = True Then frmMain.Text1.Text = 10
If Option1(1).Value = True Then frmMain.Text1.Text = 25
If Option1(2).Value = True Then frmMain.Text1.Text = 50
If Option1(3).Value = True Then frmMain.Text1.Text = 200
txtMessages.Text = ""

On Error Resume Next
cmdStart.Enabled = False
exitproc = False
Gif89a1.Visible = True
Gif89a1.FileName = App.Path & "\mov1.gif"
xz = 0
o = 1
oo = 1
ooo = 1
Dim a As Integer
Dim c As Integer
Dim er As Integer
Dim br As Integer
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim s As String
files1 = Text1.Text
If Check1.Value = 1 Then opt1 = True Else opt1 = False
If Check2.Value = 1 Then opt2 = True Else opt2 = False
If Check3.Value = 1 Then opt3 = True Else opt3 = False
appdir = txtDir.Text
br = Len(appdir)
er = InStrRev(appdir, "\")
If Not fso.folderexists(appdir) Then
MsgBox "Invalid destination directory"
Exit Sub
End If
If br = er Then appdir = Left(appdir, br)
stryyy = txtWebsite.Text
files = 0
c = Len(stryyy)
a = InStr(stryyy, "/")
If a = 0 Then stryyy = stryyy & "/"
a = InStr(stryyy, "/")
strServer = Left(stryyy, a - 1)
strURL = Right(stryyy, c - a + 1)
strTryurl = strURL
er = InStr(strTryurl, ".htm")
If er = 0 Then er = InStr(strTryurl, ".asp")
If er = 0 Then
a = InStrRev(strTryurl, "/")
If Not a = Len(strTryurl) Then strTryurl = strTryurl & "/"
a = InStrRev(strTryurl, "/")
c = Len(strTryurl)
strMurl = Left(strTryurl, a - 1)
Call getsize
Call urltry
Else
a = InStrRev(strTryurl, "/")
c = Len(strTryurl)
strMurl = Left(strTryurl, a - 1)
End If
iPort = 80
Call process(strServer, strURL)
Call stripurl
txtMessages.Text = txtMessages.Text & vbCrLf & " Starting to download links in file"
txtMessages.SelStart = Len(txtMessages.Text)
Call dotry
For jj = 1 To o
Call level1(url(jj))
Next jj
Call downlevel
For jj = 1 To oo
Call level2(levu(jj))
Next jj
Call downlevel2
MsgBox "Finished downloading"
Command1.Caption = "Exit"
Set frmMain = Nothing
Set frmstart = Nothing
Set frmUrl = Nothing
cmdStart.Enabled = True
End Sub




Private Sub download(strSServer As String, strUURL As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim sServer As String
Dim sUrl As String
Dim x As String
Dim y As String
Dim z, f
If files > files1 Then Exit Sub
iPort = 80
sServer = strSServer
sUrl = strUURL
iFlags = INTERNET_FLAG_NO_COOKIES
iFlags = iFlags Or INTERNET_FLAG_NO_CACHE_WRITE
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.fileexists(appdir & sUrl) Then Exit Sub
hInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
If hInternet <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Open Successfull"
hConnect = InternetConnect(hInternet, sServer, iPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)
If hConnect <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Connect Succesfull"
hRequest = HttpOpenRequest(hConnect, "GET", sUrl, "HTTP/1.0", vbNullString, vbNullString, iFlags, 0)
If hRequest <> 0 Then txtMessages.Text = txtMessages.Text & vbCrLf & "Http Open Request succesfull"
bRes = HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0)
If bRes = True Then txtMessages.Text = txtMessages.Text & vbCrLf & "Request successfull"
strDir = Dir(appdir & sUrl)
If Len(strDir) > 0 Then
Kill appdir & sUrl
End If
iFile = FreeFile()
Call makedire(sUrl)
Open appdir & sUrl For Binary Access Write As iFile
Do
bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
If lBytesRead > 0 Then
Put iFile, , strBuffer
End If
Loop While lBytesRead > 0
Close iFile
files = files + 1
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished downloading " & sServer & sUrl
txtMessages.SelStart = Len(txtMessages.Text)
DoEvents
If exitproc = True Then Unload Me
End Sub


Private Sub makedire(strYZ As String)
If exitproc = True Then Exit Sub
On Error Resume Next
strYZZ = strYZ
Dim sty As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim b As Integer
Dim a As Integer
Dim x(10) As Integer
b = 0
a = InStr(strYZZ, "/")
c = Len(strYZZ)
stree = strYZZ
x(0) = 0
While a <> 0
b = b + 1
x(b) = x(b - 1) + a
strYZZ = Right(strYZZ, c - a)
c = Len(strYZZ)
a = InStr(strYZZ, "/")
Wend
For s = 1 To b
stre = Left(stree, x(s))

y = appdir & stre
txtMessages.Text = txtMessages.Text & vbCrLf & "Creating local sub directory " & appdir & stre
txtMessages.SelStart = Len(txtMessages.Text)
If Not fso.folderexists(y) Then MkDir (y)
Next s
DoEvents
If exitproc = True Then Unload Me
End Sub

 
 

Private Sub subfiles(strBserver As String, strBurl As String)
On Error Resume Next
If exitproc = True Then Exit Sub
Dim aer As Integer
Dim ber As Integer
Dim iFile As Integer
Dim strTry5 As String
Dim strbburl As String
strbburl = strBurl
If strbburl = "" Then Exit Sub
Dim strTry6 As String
strTry6 = ""
iFile = 1
Dim strCheck As String
Dim strTry3 As String
strTry3 = "src=" & Chr(34)
strTry4 = Chr(34)
Dim strtry9 As String
strtry9 = "SRC=" & Chr(34)
Open appdir & strbburl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1

ans = InStr(strCheck, strTry3)
If ans = 0 Then ans = InStr(strCheck, strtry9)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strTry3) - ans)
cns = InStr(strCheck, strTry4)
If cns > 0 Then
strTry5 = Left(strCheck, cns - 1)
aer = InStr(strTry5, "/")
If aer <> 0 Then
ber = InStr(strTry5, "../")
If ber <> 0 Then
strTry6 = Right(strTry5, Len(strTry5) - ber + 1)
GoTo 10
Else:
strTry6 = strMurl & "/" & strTry5
GoTo 10
End If
End If
strTry6 = strMurl & "/" & strTry5
10:
If aer = 1 Then strTry6 = strTry5
Dim mz As Integer
Dim mx As Integer
Dim my As Integer
Dim mw As Integer
Dim ms As Integer
Dim mt As Integer
If opt2 = False Then ms = 0 Else ms = InStr(strTry6, ".gif")
If opt2 = False Then mt = 0 Else mt = InStr(strTry6, ".jpg")
If opt3 = True Then
ms = 1
mt = 1
End If
Dim mlo As Integer
mlo = InStr(strTry6, ".htm")
If mlo <> 0 Then
url(o) = strTry6
o = o + 1
End If
mz = InStr(strTry6, ".co")
mx = InStr(strTry6, ".net")
my = InStr(strTry6, ".org")
mw = InStr(strTry6, ".edu")
If mz = 0 And my = 0 And mx = 0 And mw = 0 And (mt <> 0 Or ms <> 0) Then
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading File " & strServer & strTry6
txtMessages.SelStart = Len(txtMessages.Text)
Call download(strServer, strTry6)
End If
End If

ans = InStr(strCheck, strTry3)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile

End Sub

Private Sub stripurl()
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim strseek99 As String
Dim g As Integer
Dim mep As Integer
Dim mpp As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
txtMessages.SelStart = Len(txtMessages.Text)
strTry = Chr(34)
strSeek = "href=" & Chr(34)
strseek99 = "HREF=" & Chr(34)
iFile = FreeFile()
Open appdir & strURL For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 Then
mep = InStr(strtry1, "../")
mpp = InStr(strtry1, "./")
If mep <> 0 Then
url(o) = strMurl & strtry1
ElseIf mpp <> 0 Then url(o) = strMurl & strtry1
Else: url(o) = strMurl & "/" & strtry1
End If
o = o + 1
End If
End If
ans = InStr(strCheck, strSeek)
DoEvents
If exitproc = True Then Unload Me
Wend



DoEvents
If exitproc = True Then Unload Me
Loop
Close iFile
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"
txtMessages.SelStart = Len(txtMessages.Text)
End Sub

Private Sub Command4_Click()
Call stripurl
End Sub
Private Sub process(strsrv As String, stru As String)
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strDserv As String

strDserv = strsrv
strDurl = stru
txtMessages.Text = txtMessages & vbCrLf & "Starting to download " & strDserv & strDurl
txtMessages.SelStart = Len(txtMessages.Text)
Call download(strDserv, strDurl)
Call background(appdir & strDurl)
txtMessages.Text = txtMessages.Text & vbCrLf & "Downloading image files"
txtMessages.SelStart = Len(txtMessages.Text)
Call subfiles(strDserv, strDurl)
txtMessages.Text = ""
End Sub

Private Sub dotry()
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer

For jj = 1 To o
DoEvents
If exitproc = True Then Unload Me

If url(jj) = "" Then Exit Sub
Call process(strServer, url(jj))
Next jj
End Sub

Private Sub Command1_Click()
exitproc = True
Set frmMain = Nothing
Unload Me

End Sub






Private Sub getsize()
hInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
If hInternet <> 0 Then hRequest = InternetOpenUrl(hInternet, "http://" & txtWebsite.Text, vbNullString, 0, INTERNET_FLAG_NO_AUTO_REDIRECT, 0)
Open appdir & "/temp.log" For Binary Access Write As 1
Do
bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
If lBytesRead > 0 Then
Put #1, , strBuffer
End If
Loop While lBytesRead > 0
Close #1
Dim fso, y, f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(appdir & "\temp.log")
y = f.Size
msize = y
End Sub

Private Sub background(strFilename As String)
Dim check As String
Dim muck As String
muck = "background=" & Chr(34)
muck1 = "BACKGROUND=" & Chr(34)
Dim a As Integer
Dim b As Integer
Open strFilename For Input As 1
Do While Not EOF(1)
Input #1, check
a = InStr(check, muck)
If a <> 0 Then
check = Right(check, Len(check) - 11 - a)
muck = Chr(34)
b = InStr(check, muck)
check = Left(check, b - 1)
cdz = InStr(check, "http://")
If cdz <> 0 Then Call download(strServer, strDurl)
Close #1
Exit Sub
End If
gr = InStr(check, muck1)
If gf <> 0 Then
check = Right(check, Len(check) - 11 - a)
muck1 = Chr(34)
b = InStr(check, muck1)
check = Left(check, b - 1)
cdz = InStr(check, "http://")
If cdz <> 0 Then Call download(strServer, strDurl)
Close #1
Exit Sub
End If
Loop
Close #1
End Sub

Private Sub Form_Load()
txtMessages.Text = "PRESS BEGIN DOWNLOADING TO BEGIN THE WEBSITE REPLICATION PROCESS."
'txtWebsite.Enabled = False
'txtDir.Enabled = False
End Sub

Private Sub urltry()
Call download(strServer, strMurl & "/index.htm")
Dim fso, f, s, t, y
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(appdir & strMurl & "/index.htm")
s = f.Size
If s = msize Then
strURL = strMurl & "/index.htm"
Exit Sub
End If
Call download(strServer, strMurl & "/index.html")
Set f = fso.GetFile(appdir & strMurl & "/index.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/index.html"
Kill (appdir & strMurl & "/index.htm")
Exit Sub
End If
Call download(strServer, strMurl & "/default.htm")
Set f = fso.GetFile(appdir & strMurl & "/default.htm")
t = f.Size
If t = msize Then
strURL = strMurl & "/default.htm"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Exit Sub
End If
Call download(strServer, strMurl & "/default.html")
Set f = fso.GetFile(appdir & strMurl & "/default.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/default.html"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Exit Sub
End If

Call download(strServer, strMurl & "/start.htm")
Set f = fso.GetFile(appdir & strMurl & "/start.htm")
t = f.Size
If t = msize Then
strURL = strMurl & "/start.htm"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Kill (appdir & strMurl & "/default.html")
Exit Sub
End If
Call download(strServer, strMurl & "/start.html")
Set f = fso.GetFile(appdir & strMurl & "/start.html")
t = f.Size
If t = msize Then
strURL = strMurl & "/start.html"
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Kill (appdir & strMurl & "/default.html")
Kill (appdir & strMurl & "/start.htm")
Exit Sub
End If
Kill (appdir & strMurl & "/index.htm")
Kill (appdir & strMurl & "/index.html")
Kill (appdir & strMurl & "/default.htm")
Kill (appdir & strMurl & "/default.html")
Kill (appdir & strMurl & "/start.htm")
Kill (appdir & strMurl & "/start.html")
Call download(strServer, strMurl & "/index.asp")
Set f = fso.GetFile(appdir & strMurl & "/index.asp")
s = f.Size
If s = msize Then
strURL = strMurl & "/index.asp"
Exit Sub
End If
Call download(strServer, strMurl & "/default.asp")
Set f = fso.GetFile(appdir & strMurl & "/default.asp")
t = f.Size
If t = msize Then
strURL = strMurl & "/default.asp"
Kill (appdir & strMurl & "/index.asp")
Exit Sub
End If
Call download(strServer, strMurl & "/start.asp")
Set f = fso.GetFile(appdir & strMurl & "/start.asp")
t = f.Size
If t = msize Then
strURL = strMurl & "/start.asp"
Kill (appdir & strMurl & "/index.asp")
Kill (appdir & strMurl & "/default.asp")
Exit Sub
End If
Kill (appdir & strMurl & "/index.asp")
Kill (appdir & strMurl & "/default.asp")
Kill (appdir & strMurl & "/start.asp")
MsgBox "Starting file not found"
End Sub

Private Sub level1(uurl As String)
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim strseek99 As String
Dim mep As Integer
Dim mpp As Integer
Dim g As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
txtMessages.SelStart = Len(txtMessages.Text)
strTry = Chr(34)
strSeek = "href=" & Chr(34)
strseek99 = "HREF=" & Chr(34)
iFile = FreeFile()
If uurl = "" Then Exit Sub
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim asd As Boolean
asd = fso.fileexists(appdir & uurl)
If asd = True Then
Open appdir & uurl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
If Not strCheck = "" Then
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
Dim bee As Integer
bee = InStr(strtry1, ".htm")
If bee = 0 Then bee = InStr(strtry1, ".asp")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 And bee <> 0 Then
mep = InStr(strtry1, "../")
mpp = InStr(strtry1, "./")
If mep <> 0 Then
levu(oo) = strMurl & strtry1
ElseIf mpp <> 0 Then levu(oo) = strMurl & strtry1
Else: levu(oo) = strMurl & "/" & strtry1
End If
oo = oo + 1
End If
End If
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
DoEvents
If exitproc = True Then Unload Me
Wend
If exitproc = True Then Unload Me
End If
Loop
Close iFile
End If
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"
txtMessages.SelStart = Len(txtMessages.Text)

End Sub
Private Sub downlevel()
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer

For jj = 1 To oo
DoEvents
If exitproc = True Then Unload Me

If levu(jj) = "" Then Exit Sub

Call process(strServer, levu(jj))
Next jj
End Sub
Private Sub level2(uuurl As String)
If exitproc = True Then Exit Sub
On Error Resume Next
Dim strSeek As String
Dim strCheck As String
Dim strSearch As String
Dim e As Integer
Dim h As Integer
Dim x As Boolean
Dim y As String
Dim ans, bns, cns
h = 0
Dim c As Integer
Dim d As Integer
Dim strTry As String
Dim strTry2 As String
Dim strTry3 As String
Dim strTry4 As String
Dim strTry5 As String
Dim strTry7 As String
Dim strseek99 As String
Dim mep As Integer
Dim mpp As Integer
Dim g As Integer
txtMessages.Text = txtMessages.Text & vbCrLf & "Finding downloadable links in url file "
txtMessages.SelStart = Len(txtMessages.Text)
strTry = Chr(34)
strSeek = "href=" & Chr(34)
strseek99 = "HREF=" & Chr(34)
iFile = FreeFile()
If uuurl = "" Then Exit Sub
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim asd As Boolean
asd = fso.fileexists(appdir & uurl)
If asd = True Then
Open appdir & uurl For Input As iFile
Do While Not EOF(iFile)
Input #iFile, strCheck
If Not strCheck = "" Then
bns = Len(strCheck)
bns = bns + 1
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
While ans <> 0
bns = Len(strCheck)
bns = bns + 1
h = h + 1
strCheck = Right(strCheck, bns - Len(strSeek) - ans)
cns = InStr(strCheck, strTry)
If cns > 0 Then
strtry1 = Left(strCheck, cns - 1)
c = InStr(strtry1, "http://")
d = InStr(strtry1, "#")
e = InStr(strtry1, "mailto:")
g = InStr(strtry1, "ftp:")
po = InStr(strtry1, "=")
pe = InStr(strtry1, ".com")
Dim bee As Integer
bee = InStr(strtry1, ".htm")
If bee = 0 Then bee = InStr(strtry1, ".asp")
If c = 0 And d = 0 And e = 0 And g = 0 And po = 0 And pe = 0 And bee <> 0 Then
mep = InStr(strtry1, "../")
mpp = InStr(strtry1, "./")
If mep <> 0 Then
levl(ooo) = strMurl & strtry1
ElseIf mpp <> 0 Then levl(ooo) = strMurl & strtry1
Else: levl(ooo) = strMurl & "/" & strtry1
End If
ooo = ooo + 1
End If
End If
ans = InStr(strCheck, strSeek)
If ans = 0 Then ans = InStr(strCheck, strseek99)
DoEvents
If exitproc = True Then Unload Me
Wend
If exitproc = True Then Unload Me
End If
Loop
Close iFile
End If
txtMessages.Text = txtMessages.Text & vbCrLf & "Finished finding links in the url"
txtMessages.SelStart = Len(txtMessages.Text)

End Sub
Private Sub downlevel2()
If files > files1 Then Exit Sub
If exitproc = True Then Exit Sub
On Error Resume Next
Dim jj As Integer

For jj = 1 To ooo
DoEvents
If exitproc = True Then Unload Me

If levl(jj) = "" Then Exit Sub

Call process(strServer, levl(jj))
Next jj
End Sub
