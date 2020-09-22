VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Anele Mbanga's ListView Row Column Backcolor Example"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column 4"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' set back colors for the specified rows and columns
    SetLIBackColor ListView1, 4, 4, vbCyan      ' row 4 column 4
    SetLIBackColor ListView1, 2, 2, vbMagenta   ' row 2 column 2

End Sub

Private Sub Form_Load()
    ' subclass the listview using the handle of the form
    ' if you are using the listview in a user control, pass the handle of the usercontrol in the
    ' user control initialize sub
    g_addProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    
    ' add a few listitems for practise
    Dim lstItm As MSComctlLib.ListItem
    
    Set lstItm = ListView1.ListItems.Add(, , "Item 1")
    lstItm.SubItems(1) = "Item 1"
    lstItm.SubItems(2) = "Item 1"
    lstItm.SubItems(3) = "Item 1"
    
    Set lstItm = ListView1.ListItems.Add(, , "Item 2")
    lstItm.SubItems(1) = "Item 2"
    lstItm.SubItems(2) = "Item 2"
    lstItm.SubItems(3) = "Item 2"
    
    Set lstItm = ListView1.ListItems.Add(, , "Item 3")
    lstItm.SubItems(1) = "Item 3"
    lstItm.SubItems(2) = "Item 3"
    lstItm.SubItems(3) = "Item 3"
    
    Set lstItm = ListView1.ListItems.Add(, , "Item 4")
    lstItm.SubItems(1) = "Item 4"
    lstItm.SubItems(2) = "Item 4"
    lstItm.SubItems(3) = "Item 4"
    
    ReDim Preserve clr(ListView1.ListItems.Count, ListView1.ColumnHeaders.Count)
    'Initialise the subclassing
    g_MaxItems = ListView1.ListItems.Count - 1
    g_MaxColumns = ListView1.ColumnHeaders.Count
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' unsubclass the listview
    ' if the listview is inside a usercontrol, put this in a terminate event
    SetWindowLong hWnd, GWL_WNDPROC, g_addProcOld
    
End Sub

Private Sub ListView1_Click()
    ' when a user clicks a listview item
    ' on the selected row, change column 2 and column 3 to be green and red
    Dim lstItm As MSComctlLib.ListItem
    Set lstItm = ListView1.SelectedItem
    
    If TypeName(lstItm) = "Nothing" Then Exit Sub
    SetLIBackColor ListView1, ListView1.SelectedItem.Index, 2, vbGreen
    SetLIBackColor ListView1, ListView1.SelectedItem.Index, 3, vbRed
End Sub
