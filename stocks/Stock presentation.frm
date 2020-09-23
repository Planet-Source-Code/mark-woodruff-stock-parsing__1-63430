VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "WINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Stock Presentation By Mark Woodruff"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2055
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   6360
      Width           =   4815
   End
   Begin RichTextLib.RichTextBox Stocks 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Stock presentation.frx":0000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Stock"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "yhoo"
      Top             =   1920
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   6000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Stock presentation.frx":0082
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Information"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Stocks.Text = ""
Winsock1.Close
Winsock1.Connect "finance.yahoo.com", "80"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Winsock1_Connect()
RichTextBox1.Text = ""
Winsock1.SendData "GET http://finance.yahoo.com/q?s=" & Text1.Text & " HTTP/1.0" & vbCrLf & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Html As String
On Error Resume Next
Winsock1.GetData Html
RichTextBox1.Text = RichTextBox1.Text & Html
If InStr(Html, "Invalid Ticker Symbol ") Then
Stocks.Text = "Invalid Ticker Symbol Please Try A Different One"
Exit Sub
End If
If InStr(Html, "</html>") Then
If InStr(RichTextBox1.Text, "Summary for ") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Summary for ") + 12)
Stocks.Text = Stocks.Text & "Company: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, " - ") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Last Trade:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Last Trade:") + 11)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "<b>") + 3)
Stocks.Text = Stocks.Text & "Last Trade: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</b") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Trade Time:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Trade Time:") + 11)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Trade Time: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Change:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Change:") + 7)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, " alt=") + 6)
Stocks.Text = Stocks.Text & "Stock Change: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, ">") - 2)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "#") + 10)
Stocks.Text = Stocks.Text & "  " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</b") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Prev Close:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Prev Close:") + 11)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Previous Close: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Open:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Open:") + 11)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Open: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Bid:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Bid:") + 4)
Text2.Text = Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</tr>") - 1)
If InStr(Text2.Text, "N/A") Then
Stocks.Text = Stocks.Text & "Bid: Not Available" & vbCrLf
Else
Dim BID As String
Text2.Text = Mid(Text2.Text, InStr(Text2.Text, "tabledata1") + 12)
BID = Left(Text2.Text, InStr(Text2.Text, "<sm") - 1)
Text2.Text = Mid(Text2.Text, InStr(Text2.Text, "<small> ") + 8)
Stocks.Text = Stocks.Text & "Bid: " & BID & " " & Left(Text2.Text, InStr(Text2.Text, "</s") - 1) & vbCrLf
End If
End If
If InStr(RichTextBox1.Text, "Ask:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Ask:") + 4)
Text2.Text = Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</tr>") - 1)
If InStr(Text2.Text, "N/A") Then
Stocks.Text = Stocks.Text & "Ask: Not Available" & vbCrLf
Else
Dim Ask As String
Text2.Text = Mid(Text2.Text, InStr(Text2.Text, "tabledata1") + 12)
Ask = Left(Text2.Text, InStr(Text2.Text, "<sm") - 1)
Text2.Text = Mid(Text2.Text, InStr(Text2.Text, "<small> ") + 8)
Stocks.Text = Stocks.Text & "Ask: " & Ask & " " & Left(Text2.Text, InStr(Text2.Text, "</s") - 1) & vbCrLf
End If
End If
If InStr(RichTextBox1.Text, "1y Target Est:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "1y Target Est:") + 14)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "1y Target Est: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Day's Range:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Day's Range:") + 12)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Day's Range: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "52wk Range:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "52wk Range:") + 11)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "52wk Range: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Volume:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Volume:") + 7)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Volume: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Avg Vol (3m):") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Avg Vol (3m):") + 7)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Avg Vol (3m): " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Market Cap:") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Market Cap:") + 7)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Market Cap: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "P/E") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "P/E") + 3)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "P/E (ttm): " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "EPS") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "EPS") + 3)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "EPS (ttm): " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
If InStr(RichTextBox1.Text, "Div") Then
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "Div") + 3)
RichTextBox1.Text = Mid(RichTextBox1.Text, InStr(RichTextBox1.Text, "tabledata1") + 12)
Stocks.Text = Stocks.Text & "Div & Yield: " & Left(RichTextBox1.Text, InStr(RichTextBox1.Text, "</t") - 1) & vbCrLf
End If
Winsock1.Close
WebBrowser1.Navigate2 "http://ichart.finance.yahoo.com/t?s=" & Text1.Text
End If
End Sub
