VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Ebay AuctionWatcher"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView List1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Auction Item"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bids"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   979
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ends"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Last Updated:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Menu mnuAuctions 
      Caption         =   "Auctions"
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update..."
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove..."
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToPage 
         Caption         =   "Go To Page..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AuctionName As String
Dim AuctionBids As Integer
Dim AuctionPrice As String
Dim AIN As String
Dim AIN2 As String
Dim AuctionTime As String

Private Sub List1_ItemClick(ByVal Item As ComctlLib.ListItem)
Label1.Caption = "Last Updated: " & Time
AIN2 = Item.Key
AIN = Right(AIN2, Len(AIN2) - 2)
GetAuctionDetails

End Sub

Private Sub mnuAdd_Click()
On Error GoTo ErrHndlr
'Add an auction
AIN = InputBox("Auction Item Number")
AIN2 = "A" & AIN
List1.ListItems.Add , AIN2
List1.ListItems(AIN2).Text = "[Loading Ebay Page]"
'Get Auction Details
GetAuctionDetails

Exit Sub
ErrHndlr:
MsgBox "Auction all ready added.", vbExclamation, "Error"






End Sub

Sub GetAuctionDetails()
'Set Last Updated Label
Label1.Caption = "Last Updated: " & Time
'Get the auction information
Winsock1.Connect "cgi.ebay.com", 80
'Next block of code is in winsock1_connect()

End Sub

Private Sub Winsock1_Close()
'Server completed Send.
RichTextBox1.SaveFile "C:\debug.html"
Call AnalyzeData

End Sub

Private Sub Winsock1_Connect()
'You just connected. Now, send ebay the request

DoEvents
Winsock1.SendData "GET /ws/eBayISAPI.dll?ViewItem&item=" & AIN & " HTTP/1.1" & vbNewLine & _
"Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & vbNewLine & _
"Accept -Language: en -us" & vbNewLine & _
"Accept -Encoding: gzip , deflate" & vbNewLine & _
"User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbNewLine & _
"Host: cgi.ebay.com" & vbNewLine & _
"Connection: Keep -Alive" & vbNewLine & vbNewLine

RichTextBox1.Text = ""
'next block of code is in winsock1_dataarival


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Data, vbString, bytesTotal

RichTextBox1.Text = RichTextBox1.Text & Data 'put the auction page together


End Sub

Sub AnalyzeData()
'On Error GoTo errhndlr2
'Get Auction Name
aryAuctionName = Split(RichTextBox1.Text, ") -", -1, vbTextCompare)
aryAuctionName2 = Split(aryAuctionName(1), "</title>", -1, vbTextCompare)
aryAuctionName3 = Split(aryAuctionName2(0), vbNewLine, -1, vbTextCompare)
AuctionName = aryAuctionName3(1)

'Get Bids
arybids = Split(RichTextBox1.Text, "var itemNumBids =", -1, vbTextCompare)
arybids2 = Split(arybids(1), """", -1, vbTextCompare)
AuctionBids = arybids2(1)



'Get Auction Price
aryPrice = Split(RichTextBox1.Text, "US $", -1, vbTextCompare)
aryPrice2 = Split(aryPrice(1), "</b>", -1, vbTextCompare)
AuctionPrice = aryPrice2(0)



'Time Remaining
arytime = Split(RichTextBox1.Text, "(Ends", -1, vbTextCompare)
aryTime2 = Split(arytime(1), vbNewLine, -1, vbTextCompare)
AuctionTime = aryTime2(1)





List1.ListItems(AIN2).Text = AuctionName
List1.ListItems(AIN2).SubItems(1) = AuctionBids
List1.ListItems(AIN2).SubItems(2) = "$ " & AuctionPrice
List1.ListItems(AIN2).SubItems(3) = AuctionTime
DoEvents
Winsock1.Close


Exit Sub
errhndlr2:
MsgBox "Invalid Ebay Item Number", vbExclamation, "Error"
List1.ListItems.Remove (AIN2)
Winsock1.Close

End Sub
