VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   450
      Left            =   8055
      TabIndex        =   2
      Top             =   4860
      Width           =   1000
   End
   Begin VB.CommandButton cmdConString 
      Caption         =   "Build Conn String"
      Height          =   450
      Left            =   9090
      TabIndex        =   1
      Top             =   4860
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4650
      Left            =   3465
      TabIndex        =   0
      Top             =   90
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   8202
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTables 
      Height          =   4740
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   8361
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImgLst1"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLst1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0454
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":09BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0A80
            Key             =   "CompRoot"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":0DA4
            Key             =   "book"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1204
            Key             =   "open1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1658
            Key             =   "history"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrid.frx":1F00
            Key             =   "close"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************
'   Refereneces set to:
'   ActiveX Data Objects 2.5 lib
'   OLE DB service component 1.0 lib
'   ADO Example Dale Cebula 2000
'
'   Demonstrates how to use ADO to retrieve database tables into a treeview
'   then uses the node click event of the treeview to retrieve the recordset
'   and populate the data bound grid. Should work with SQL Server as well as MS Access
'*****************************************************************************

    Private cDocmd As New clsADO
'   Declare new nodes
    Private newNode As Node        '   tvTables
'   Variables to hold the text on the selected node
    Private SelectedNode$
    Private reckey%

'   Sub to populate Grid
Private Sub PopulateGrid()
Dim adoRs As New ADODB.Recordset
Dim adocn As New ADODB.Connection

    adocn.CursorLocation = adUseClient
    adocn.Open cDocmd.ConnectionString$

'   Open the connection
'   Square brackets are use here to open tables with spaces between their names
    adoRs.Open "[" & SelectedNode$ & "]", adocn, adOpenStatic
'   set the grid DS to the active connection
    Set DataGrid1.DataSource = adoRs
    Set adoRs = Nothing: Set adocn = Nothing
End Sub

Private Sub cmdConString_Click()
On Error Resume Next
Dim msDLink As New DataLinks
'   This sub uses MS Datalinks to build a connection string that is used to
'   poulate the treeview

'   Choose the provider Jet OLE 3 is Access '97
'   OLE 4.0 is Access 2000
'   Then select the database that you want to open and click ok.
    
    cDocmd.ConnectionString = msDLink.PromptNew
'   Populate the treeview with the tables
    PopulateTreeView
End Sub

Private Sub PopulateTreeView()
Dim TableCount%
Dim strID$
Screen.MousePointer = vbHourglass

If cDocmd.ConnectionString <> "" Then
'   Add the first Nodes to the Treeview
    Set newNode = tvTables.Nodes.Add(, tvwFirst, "Root1", "DataBase Tables", ImgLst1.ListImages("CompRoot").Key)
'   Expand root nodes
    newNode.Expanded = True

'  Retrieve the tables from the underlying database - Call the DBCommands Object
'  1 argument opens the recordset open
    cDocmd.GetTables , cDocmd.ConnectionString, , , 1
    
    For TableCount% = 1 To cDocmd.TableNames.Count
'  Build Key
        reckey% = reckey% + 1
        strID$ = "Key"
        strID$ = strID$ + CStr(reckey%)
'   Add each table to the treeview
        Set newNode = tvTables.Nodes.Add("Root1", tvwChild, strID$, cDocmd.TableNames.Item(TableCount%), ImgLst1.ListImages("close").Key, _
        ImgLst1.ListImages("open").Key)
    Next TableCount%

'   Remove the tables from the collection
    For TableCount% = TableCount% - 1 To 1 Step -1
        cDocmd.TableNames.Remove (TableCount%)
    Next TableCount%
End If

    Screen.MousePointer = vbDefault

End Sub


Private Sub tvTables_NodeClick(ByVal Node As MSComctlLib.Node)
'   Initialise variables
    SelectedNode = Node.Text

    If Node <> "DataBase Tables" Then
        PopulateGrid
    End If
End Sub
