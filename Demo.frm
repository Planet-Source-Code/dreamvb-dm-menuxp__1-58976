VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DM CoolMenu XP"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMenuArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   990
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   0
      Top             =   4815
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtEd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "Demo.frx":0000
      Top             =   465
      Width           =   7065
   End
   Begin VB.Label lbMenu 
      AutoSize        =   -1  'True
      Caption         =   "&Edit"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Tag             =   "1"
      Top             =   90
      Width           =   270
   End
   Begin VB.Label lbMenu 
      AutoSize        =   -1  'True
      Caption         =   "&File"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Tag             =   "1"
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DM CoolMenu XP
' Well anyway I am back agian. after makeing my frame XP control I desided to
' have a go at a custom menu well this took a little longer than expeted almost an hour and a half to make
' tho it still does look quite good.
' as always there are some bugs. mostly when moveing over items in the menu it flashes a little.
' but apart from that it not to bad.
' well as always do what you want with the code.
' if you update or fix some things lets of know. I like to see what you done with it.

' Thanks.
' Ben jones vbdream2k@yahoo.com

Private Type MenuItems
    Index As Integer
    ItemCaption As String
    ItemKey As Variant
End Type

Private Type MyMenu
    MenuCount As Integer
    Menus() As MenuItems
End Type

Dim Y_Pos As Integer, ICurrentMenuIndex As Integer, IsDown As Boolean
Dim ICurrentMenuIndexCount As Integer

Private TmpKeys() As Variant
Private TmpMenuCaption() As String

Private DmMenu As MyMenu

Private Sub LoadMenu()
    ' File Menu
    MenuAddItem 0, "New Project", "M_NEW"
    MenuAddItem 0, "Open Project...", "M_OPEN"
    MenuAddItem 0, "Save project", "M_SAVE"
    MenuAddItem 0, "Save As...", "M_SAVE_AS"
    MenuAddItem 0, "Print", "M_PRINT"
    MenuAddItem 0, "Print Setup", "M_PRINT_SETUP"
    MenuAddItem 0, "Exit", "M_EXIT"
    'Edit Menu
    MenuAddItem 1, "Cut", "M_CUT"
    MenuAddItem 1, "Copy", "M_COPY"
    MenuAddItem 1, "Paste", "M_PASTE"
    MenuAddItem 1, "Delete.", "M_DELETE"
    MenuAddItem 1, "Select All", "M_SEL_ALL"
    MenuAddItem 1, "Find & Replace", "M_REPLACE"
    MenuAddItem 1, "Insert", "M_INSERT"
End Sub

Private Sub MenuClick(MenuKey As Variant, MenuIndex As Integer, MenuCaption As String)
    
    Select Case UCase(MenuKey)
        Case "M_NEW"
            MsgBox "you clicked new"
        Case "M_CUT"
            txtEd.SelText = ""
        Case "M_SEL_ALL"
            txtEd.SelStart = 0
            txtEd.SelLength = Len(txtEd.Text)
            txtEd.SetFocus
        Case "M_EXIT"
            MsgBox "Good Bye"
            ResetMenu
            End
    End Select
    
End Sub

Private Sub DrawMenu(MnuIndex As Integer, PicMnuBox As PictureBox)
Dim X As Integer, Y As Integer, Z As String
Dim TmpArr() As String, nWidth As Long, nHeight As Long
Dim MenuCount As Integer
    
    Dim Cnt As Integer
    Cnt = -1
    Erase TmpArr
    Erase TmpKeys()

    ReDim Preserve TmpKeys(0)
    ReDim Preserve TmpArr(0)
    ReDim Preserve TmpMenuCaption(0)
    
    PicMnuBox.Cls
    MenuCount = 0
    X = 0
    ' little bit of the code that copys all the menu items for the selected menu index
    ' into an array we use this later to sort the menu see other code below

    For X = 0 To UBound(DmMenu.Menus)
        If Val(MnuIndex) = DmMenu.Menus(X).Index Then
            Cnt = Cnt + 1
            ReDim Preserve TmpArr(Cnt)
            ReDim Preserve TmpMenuCaption(Cnt)
            ReDim Preserve TmpKeys(Cnt)
            
            TmpMenuCaption(Cnt) = DmMenu.Menus(X).ItemCaption
            TmpKeys(Cnt) = DmMenu.Menus(X).ItemKey
            TmpArr(Cnt) = DmMenu.Menus(X).ItemCaption
            
        End If
    Next
    
    ICurrentMenuIndexCount = UBound(TmpMenuCaption)
    
    ' This little bit of code I used to find the size of the menu
    ' what it does is looks thought each menu item and finds the
    ' caption with the most text in it

    For X = 0 To UBound(TmpArr) - 1
        For Y = X + 1 To UBound(TmpArr)
            If Len(TmpArr(X)) > Len(TmpArr(Y)) Then
                Z = TmpArr(Y)
                TmpArr(Y) = TmpArr(X)
                TmpArr(X) = Z
            End If
        Next
    Next
    
    PicMnuBox.Visible = True ' Show the menu
    
    nWidth = PicMnuBox.TextWidth(TmpArr(UBound(TmpArr))) * 2 ' Get the width for our menu
    'the above sets the width of the menu form the info we found out from the sort code
    nHeight = PicMnuBox.TextHeight("Xz") ' Set the height of the menu
    MenuCount = UBound(TmpArr) + 1 ' Set the count to the number of items in the temp array
    
    PicMnuBox.Width = nWidth ' Update menu width
    PicMnuBox.Height = (MenuCount * 20) + nHeight ' Update the menus height
    DrawMenuOutline ' Draw the menu outline
    
    On Error Resume Next
    
    For X = 0 To UBound(TmpMenuCaption) + 1
        ' this nise block of code does all the menu stuff
        If (IsDown And Y_Pos = X) Then ' Are we on a menu item index and is mouse down
            PicMnuBox.ForeColor = vbBlue ' update menu forcolor
            PicMnuBox.FontBold = True   ' turn on bold
            n_pos = 6 + (Y_Pos * 20)    ' used as an offset to draw a small frame for the menu item
            PicMenuArea.Line (4, n_pos)-(PicMenuArea.ScaleWidth - 5, n_pos + 20), &HA1A1A1, B ' draw the outline
            PicMenuArea.Line (5, n_pos + 1)-(PicMenuArea.ScaleWidth - 6, n_pos + 19), &HE6E6E6, BF ' draw the inside the cell
            PicMnuBox.CurrentX = 13 ' we to start to print our caption from
            PicMnuBox.CurrentY = (10 + (X * 20)) ' Position of the text in the right place
            PicMnuBox.Print TmpMenuCaption(X) ' set the menu items caption
            PicMnuBox.Refresh
        Else
            ' Almost save as above
            PicMnuBox.FontBold = False
            PicMnuBox.ForeColor = vbBlack
            PicMnuBox.CurrentX = 13
            PicMnuBox.CurrentY = (10 + (X * 20))
            PicMnuBox.Print TmpMenuCaption(X)
            PicMnuBox.Refresh
        End If
    Next
    ' Position the menu we it ment to show
    PicMnuBox.Left = lbMenu(ICurrentMenuIndex).Left - 2
    PicMnuBox.Top = lbMenu(ICurrentMenuIndex).Top + lbMenu(ICurrentMenuIndex).Height
    
    X = 0: Y = 0: Z = ""
    nWidth = 0: nHeight = 0
    
End Sub

Private Sub MenuAddItem(Index As Integer, MenuCaption As String, MenuKey As Variant)
    ' All this does is add information for the menu eg Caption, Key, Menu Index
    ReDim Preserve DmMenu.Menus(DmMenu.MenuCount)
    DmMenu.Menus(DmMenu.MenuCount).Index = Index
    DmMenu.Menus(DmMenu.MenuCount).ItemCaption = MenuCaption
    DmMenu.Menus(DmMenu.MenuCount).ItemKey = MenuKey
    DmMenu.MenuCount = DmMenu.MenuCount + 1
End Sub

Sub ResetMenu()
    DmMenu.MenuCount = 0
    Erase DmMenu.Menus()
    ReDim Preserve DmMenu.Menus(0)
End Sub

Sub DrawMenuOutline()
    ' this sub Draws the outline for the menu
    PicMenuArea.Line (0, 0)-(PicMenuArea.ScaleWidth - 1, PicMenuArea.ScaleHeight - 1), &HA5A5A5, B
End Sub

Private Sub Form_Load()
    ResetMenu ' Reset menu
    LoadMenu  ' Load in some menu items
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsDown Then
        Y_Pos = -1: DrawMenu ICurrentMenuIndex, PicMenuArea
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicMenuArea_MouseUp Button, Shift, X, Y
End Sub

Private Sub lbMenu_Click(Index As Integer)
    ICurrentMenuIndex = Index
    DrawMenu Index, PicMenuArea
End Sub

Private Sub PicMenuArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Y_Pos = (Y \ 20): IsDown = True
    If (Y_Pos > ICurrentMenuIndexCount) Then IsDown = False: Exit Sub 'is Y_Pos more than our menu count
    DrawMenu ICurrentMenuIndex, PicMenuArea ' Call DrawMenu
    MenuClick TmpKeys(Y_Pos), ICurrentMenuIndex, TmpMenuCaption(Y_Pos) ' Call MenuClick
End Sub

Private Sub PicMenuArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Y_Pos = (Y \ 20): IsDown = True
    If (Y_Pos > ICurrentMenuIndexCount) Then IsDown = False: Exit Sub
    DrawMenu ICurrentMenuIndex, PicMenuArea ' Call DrawMenu
End Sub

Private Sub PicMenuArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IsDown = False
    PicMenuArea.Visible = False
End Sub

Private Sub txtEd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicMenuArea_MouseUp Button, Shift, X, Y
End Sub
