VERSION 5.00
Begin VB.UserControl TextBox 
   BackColor       =   &H00FFC0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ToolboxBitmap   =   "TextBox.ctx":0000
   Begin VB.TextBox TextBox 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MiTextBox Control v. 1.1.5
' Inspired in some Controls not founds in Internet.   =)
' Proyect Start - jul/2003
' Actual Revision - Aug/05/2003
' Comments, sugestions, etc. are welcome.
' You can set the Text Alignment (Left, Right, Center)
' You can set the control Appearance (Flat, 3D) and Color in Normal and OnFocus states.
' You can set the type of entry allowed (Numbers, Numbers w/ simbols, Characters, (Y) or (N), Alphanumeric or Anything)
' You can set to play beep on bad entries.
' You can set Auto Upper Case.
' You can set Auto Select Text.
' Undo Text with Escape key pressing.
' Controls navigation with keys:
'                               Next TextBox: right arrow, down arrow and Enter.
'                               Previous TextBox: left arrow, up arrow.
' Support all the VB TextBox methods and properties
' Written by José Luis Farías.
' Chile 1446 - Salto - Uruguay - CP 50.000
' JoseloFarias[at]adinet.com.uy
' ¡¡¡Vamo' arriba Uruguay, carajo!!!
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
' ¡PLEASE!, if you use this Control sendme your Name and Country
' And if you like, emailme a program copy (source code if better)
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*

Option Explicit 'is better

'********************************************************************************
'***    PUBLIC ENUMERATIONS
'********************************************************************************

Enum ctrAlignment
    [Align Left] = 0
    [Align Right] = 1
    [Centered] = 2
End Enum

Enum ctrlAppearance
    [Flat] = 0
    [3D] = 1
End Enum

Enum ctrlBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Enum ctrlLinkMode
    [None] = 0
    [Automatic] = 1
    [Manual] = 2
    [Notify] = 3
End Enum

Enum ctrlMousePointer
    [Default] = 0
    [arrow] = 1
    [Cross] = 2
    [I Beam] = 3
    [Icon] = 4
    [Size] = 5
    [Size NE SW] = 6
    [Size N S] = 7
    [Size NW SE] = 8
    [Size W E] = 9
    [Up Arrow] = 10
    [Hourglass] = 11
    [No Drop] = 12
    [Arrow And Hourglass] = 13
    [Arrow And Question] = 14
    [Size All] = 15
    [Custom] = 99
End Enum

Enum ctrlOLEDragMode
    Manual = 0
    Automatic = 1
End Enum

Enum ctrlOLEDropMode
    None = 0
    Manual = 1
    Automatic = 2
End Enum

Enum ctrlEntryType
    [Numeric] = 0
    [Numbers & Signs] = 1
    [Character] = 2
    [(Y)es or (N)o] = 3
    [AlphaNumeric] = 4
    [Anything] = 5
End Enum

'********************************************************************************
'***    DECLARATION OF PRIVATE VARIABLES
'********************************************************************************

Private mBackColor_Normal As OLE_COLOR
Private mBackColor_OnGotFocus As OLE_COLOR
Private mAutoSelectText As Boolean
Private mBeepOnBadType As Boolean
Private mAutoUpperCase As Boolean
Private mMouseDown As Boolean
Private mSelectOnClick  As Boolean
Private mUseDefaultText As Boolean
Private mAppearance_Normal As Byte
Private mAppearance_OnGotFocus As Byte
Private mEntryType As Byte
Private mDefaultText As String
Private mUndoText As String

'********************************************************************************
'***    DECLARATION OF PUBLIC EVENTS
'********************************************************************************

Event Change()
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

'********************************************************************************
'***    GET'S AND LET'S
'********************************************************************************

Public Property Get Alignment() As ctrAlignment
Attribute Alignment.VB_Description = "Returns/sets the text alignmnent"
    Alignment = TextBox.Alignment
End Property
Public Property Let Alignment(ByVal Value As ctrAlignment)
    TextBox.Alignment = Value
    PropertyChanged "Alignment"
End Property
Public Property Get Appearance_Normal() As ctrlAppearance
    Appearance_Normal = mAppearance_Normal
End Property
Public Property Let Appearance_Normal(ByVal Value As ctrlAppearance)
    mAppearance_Normal = Value
    TextBox.Appearance = Value
    PropertyChanged "Appearance_Normal"
End Property
Public Property Get Appearance_OnGotFocus() As ctrlAppearance
    Appearance_OnGotFocus = mAppearance_OnGotFocus
End Property
Public Property Let Appearance_OnGotFocus(ByVal Value2 As ctrlAppearance)
    mAppearance_OnGotFocus = Value2
    PropertyChanged "Appearance_OnGotFocus"
End Property
Public Property Get AutoUpperCase() As Boolean
    AutoUpperCase = mAutoUpperCase
End Property
Public Property Let AutoUpperCase(ByVal Value As Boolean)
    mAutoUpperCase = Value
    PropertyChanged "AutoUpperCase"
End Property
Public Property Get AutoSelect() As Boolean
    AutoSelect = mAutoSelectText
End Property
Public Property Let AutoSelect(ByVal Value As Boolean)
    mAutoSelectText = Value
    PropertyChanged "AutoSelect"
End Property
Public Property Get BeepOnBadType() As Boolean
    BeepOnBadType = mBeepOnBadType
End Property
Public Property Let BeepOnBadType(ByVal Value As Boolean)
    mBeepOnBadType = Value
    PropertyChanged "BeepOnBadType"
End Property
Public Property Get BackColor_Normal() As OLE_COLOR
    BackColor_Normal = mBackColor_Normal
End Property
Public Property Let BackColor_Normal(ByVal Value As OLE_COLOR)
    mBackColor_Normal = Value
    TextBox.BackColor = Value
    PropertyChanged "BackColor_Normal"
End Property
Public Property Get BackColor_OnGotFocus() As OLE_COLOR
    BackColor_OnGotFocus = mBackColor_OnGotFocus
End Property
Public Property Let BackColor_OnGotFocus(ByVal Value2 As OLE_COLOR)
Attribute BackColor_OnGotFocus.VB_Description = "This is the color that the textbox will change to when it receives focus."
Attribute BackColor_OnGotFocus.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    mBackColor_OnGotFocus = Value2
    PropertyChanged "BackColor_OnGotFocus"
End Property
Public Property Get BorderStyle() As ctrlBorderStyle
    BorderStyle = TextBox.BorderStyle
End Property
Public Property Let BorderStyle(ByVal Value As ctrlBorderStyle)
    TextBox.BorderStyle = Value
    PropertyChanged "BorderStyle"
End Property
Public Property Get CausesValidation() As Boolean
    CausesValidation = TextBox.CausesValidation
End Property
Public Property Let CausesValidation(ByVal Value As Boolean)
    TextBox.CausesValidation = Value
    PropertyChanged "CausesValidation"
End Property
Public Property Get DataField() As String
Attribute DataField.VB_MemberFlags = "103c"
    DataField = TextBox.DataField
End Property
Public Property Let DataField(ByVal Value As String)
    TextBox.DataField = Value
    PropertyChanged "DataField"
End Property
Public Property Get DefaultText() As String
Attribute DefaultText.VB_ProcData.VB_Invoke_Property = ";Text"
    DefaultText = mDefaultText
End Property
Public Property Let DefaultText(ByVal Value As String)
    mDefaultText = Value
    PropertyChanged "DefaultText"
End Property
Public Property Get Enabled() As Boolean
    Enabled = TextBox.Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    TextBox.Enabled = Value
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
   Set Font = TextBox.Font
End Property
Public Property Set Font(ByVal Value As Font)
   Set TextBox.Font = Value
   PropertyChanged "Font"
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
   FontBold = TextBox.FontBold
End Property
Public Property Let FontBold(ByVal Value As Boolean)
   TextBox.FontBold = Value
   PropertyChanged "FontBold"
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
   FontItalic = TextBox.FontItalic
End Property
Public Property Let FontItalic(ByVal Value As Boolean)
   TextBox.FontItalic = Value
   PropertyChanged "FontItalic"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
   FontName = TextBox.FontName
End Property
Public Property Let FontName(ByVal Value As String)
   TextBox.FontName = Value
   PropertyChanged "FontName"
End Property
Public Property Get FontSize() As Integer
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
   FontSize = TextBox.FontSize
End Property
Public Property Let FontSize(ByVal Value As Integer)
   TextBox.FontSize = Value
   PropertyChanged "FontSize"
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethru.VB_MemberFlags = "400"
   FontStrikethru = TextBox.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal Value As Boolean)
   TextBox.FontStrikethru = Value
   PropertyChanged "FontStrikethru"
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
   FontUnderline = TextBox.FontUnderline
End Property
Public Property Let FontUnderline(ByVal Value As Boolean)
   TextBox.FontUnderline = Value
   PropertyChanged "FontUnderline"
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = TextBox.ForeColor
End Property
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
   TextBox.ForeColor = Value
   PropertyChanged "ForeColor"
End Property
Public Property Get HideSelection() As Boolean
   HideSelection = TextBox.HideSelection
End Property
Public Property Get hWnd() As Long
   hWnd = TextBox.hWnd
End Property
Public Property Get Index() As Long
   Index = TextBox.Index
End Property
Public Property Get Left() As Long
   Left = TextBox.Left
End Property
Public Property Let Left(ByVal Value As Long)
   TextBox.Left = Value
   PropertyChanged "Left"
End Property
Public Property Get LinkItem() As String
   LinkItem = TextBox.LinkItem
End Property
Public Property Let LinkItem(ByVal Value As String)
   TextBox.LinkItem = Value
   PropertyChanged "LinkItem"
End Property
Public Property Get Locked() As Boolean
   Locked = TextBox.Locked
End Property
Public Property Let Locked(ByVal Value As Boolean)
   TextBox.Locked = Value
   PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
   MaxLength = TextBox.MaxLength
End Property
Public Property Let MaxLength(ByVal Value As Long)
   TextBox.MaxLength = Value
   PropertyChanged "MaxLength"
End Property
Public Property Get MousePointer() As ctrlMousePointer
   MousePointer = TextBox.MousePointer
End Property
Public Property Let MousePointer(ByVal Value As ctrlMousePointer)
   TextBox.MousePointer = Value
   PropertyChanged "MousePointer"
End Property
Public Property Get MultiLine() As Boolean
   MultiLine = TextBox.MultiLine
End Property
Public Property Get OLEDragMode() As ctrlOLEDragMode
   OLEDragMode = TextBox.OLEDragMode
End Property
Public Property Let OLEDragMode(ByVal Value As ctrlOLEDragMode)
   TextBox.OLEDragMode = Value
   PropertyChanged "OLEDragMode"
End Property
Public Property Get OLEDropMode() As ctrlOLEDropMode
   OLEDropMode = TextBox.OLEDropMode
End Property
Public Property Let OLEDropMode(ByVal Value As ctrlOLEDropMode)
   TextBox.OLEDropMode = Value
   PropertyChanged "OLEDropMode"
End Property
Public Property Get PasswordChar() As String
   PasswordChar = TextBox.PasswordChar
End Property
Public Property Let PasswordChar(ByVal Value As String)
   TextBox.PasswordChar = Value
   PropertyChanged "PasswordChar"
End Property
Public Property Get RightToLeft() As Boolean
   RightToLeft = TextBox.RightToLeft
End Property
Public Property Let RightToLeft(ByVal Value As Boolean)
   TextBox.RightToLeft = Value
   PropertyChanged "RightToLeft"
End Property
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
    Text = TextBox.Text
End Property
Public Property Let Text(ByVal NewText As String)
    TextBox.Text() = NewText
    PropertyChanged "Text"
End Property
Public Property Get Text_AsDefaultProperty() As String
    Text_AsDefaultProperty = TextBox.Text
End Property
Public Property Let Text_AsDefaultProperty(ByVal NewText As String)
Attribute Text_AsDefaultProperty.VB_ProcData.VB_Invoke_PropertyPut = ";Text"
Attribute Text_AsDefaultProperty.VB_UserMemId = 0
Attribute Text_AsDefaultProperty.VB_MemberFlags = "40"
    TextBox.Text() = NewText
    PropertyChanged "Text_AsDefaultProperty"
End Property
Public Property Get EntryType() As ctrlEntryType
    EntryType = mEntryType
End Property
Public Property Let EntryType(ByVal Value As ctrlEntryType)
    mEntryType = Value
    PropertyChanged "EntryType"
End Property
Public Property Get UseDefaultText() As Boolean
    UseDefaultText = mUseDefaultText
End Property
Public Property Let UseDefaultText(ByVal Value As Boolean)
    mUseDefaultText = Value
    PropertyChanged "UseDefaultText"
End Property
Private Sub TextBox_Change()
    RaiseEvent Change
End Sub
Private Sub TextBox_Click()
'Used for Text AutoSelection
    RaiseEvent Click
      With TextBox
        If mAutoSelectText And mSelectOnClick Then
            .SelStart = 0
            .SelLength = Len(.Text)
            mSelectOnClick = False
        End If
    End With
End Sub
Private Sub TextBox_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub TextBox_GotFocus()
'Used for Appearance change
    TextBox.Appearance = mAppearance_OnGotFocus
'AutoSelectText
    With TextBox
        If mAutoSelectText Then
            .SelStart = 0
            .SelLength = Len(.Text)
            mSelectOnClick = False
        End If
'Remember actual text
        mUndoText = .Text
        .BackColor = BackColor_OnGotFocus
    End With
End Sub
'Provide a control navigation with arrows keys and Enter
Private Sub textbox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
        Case 39, 40, 13  'Next Control: right arrow, down arrow and Enter
            SendKeys "{Tab}"
        Case 37, 38 'Previous Control: left and up arrows
            SendKeys "+{Tab}"
    End Select
End Sub
Private Sub TextBox_KeyPress(KeyAscii As Integer)
'AutoUpperCase, UseDefaultText, UndoText, and input validation
    RaiseEvent KeyPress(KeyAscii)
    If mAutoUpperCase Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    If KeyAscii = vbKeyEscape Then
        TextBox.Text = mUndoText
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
        If mUseDefaultText Then
            mUndoText = mDefaultText
        End If
        Exit Sub
'Key Codes: Copy = 3, Tab = 9, Paste = 22, Cut = 24, Undo = 26
    ElseIf KeyAscii = vbKeyBack Or KeyAscii = 24 Or KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 26 Then
        Exit Sub
    End If
    Dim sTemp As String
    sTemp = Chr$(KeyAscii)
    Select Case mEntryType
        Case 0 'Numbers only
            Select Case sTemp
                Case "0" To "9"
                    KeyAscii = Asc(sTemp)
                Case Else
                    KeyAscii = 0
            End Select

        Case 1 'Numbers, minus sign, commas and points
            Select Case sTemp
                Case "0" To "9", "-", ".", ","
                    KeyAscii = Asc(sTemp)
                Case Else
                    KeyAscii = 0
            End Select
        Case 2 'Characters and Space (with spanish specials)
            Select Case sTemp
                Case "a" To "z", "A" To "Z", "á", "Á", "é", "É", "í", "Í", "ó", "Ó", "ú", "Ú", "ñ", "Ñ", " "
                    KeyAscii = Asc(sTemp)
                Case Else
                    KeyAscii = 0
            End Select
        Case 3 'Y or N
            Select Case sTemp
                Case "Y", "N", "y", "n"
                    KeyAscii = Asc(sTemp)
                Case Else
                    KeyAscii = 0
            End Select
        Case 4 'Alphanumeric
            Select Case sTemp
                Case "0" To "9", "-", ".", ","
                    KeyAscii = Asc(sTemp)
                Case "a" To "z", "A" To "Z", "á", "Á", "é", "É", "í", "Í", "ó", "Ó", "ú", "Ú", "ñ", "Ñ", " "
                    KeyAscii = Asc(sTemp)
                Case Else
                    KeyAscii = 0
            End Select
        Case 5 'Anything
    End Select
'Beep on bad type
    If mBeepOnBadType And KeyAscii = 0 Then Beep
End Sub
Private Sub TextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub TextBox_LostFocus()
    TextBox.BackColor = mBackColor_Normal
    TextBox.Appearance = mAppearance_Normal
    mSelectOnClick = True
End Sub
Private Sub TextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    mMouseDown = True
End Sub
Private Sub TextBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub TextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    mMouseDown = False
End Sub
Private Sub TextBox_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub
Private Sub TextBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub
Private Sub TextBox_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub
Private Sub TextBox_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub
Private Sub TextBox_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub
Private Sub TextBox_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub
Private Sub UserControl_InitProperties()
'Initial Properties of the new created Controls
    TextBox.Appearance = 0  'Flat
    mAutoSelectText = True  'AutoSelect
    mAutoUpperCase = False  'No AutoUpperCase
    mBackColor_Normal = vbWhite 'White ;-)
    TextBox.BackColor = mBackColor_Normal   'Set TextBox BackColor as the Control BackColor
    mBackColor_OnGotFocus = &HFFC0C0    'Some type of blue
    mEntryType = 5  'Anything
    mDefaultText = ""   'Default text is a empty string
    mUseDefaultText = True  'Use default text
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mAutoSelectText = .ReadProperty("AutoSelect", True)
        mAutoUpperCase = .ReadProperty("AutoUpperCase", False)
        mBeepOnBadType = .ReadProperty("BeepOnBadType", False)
        mBackColor_Normal = .ReadProperty("BackColor_Normal", vbWhite)
        mBackColor_OnGotFocus = .ReadProperty("BackColor_OnGotFocus", vbCyan)
        mAppearance_Normal = .ReadProperty("Appearance_Normal", 0)
        mAppearance_OnGotFocus = .ReadProperty("Appearance_OnGotFocus", 0)
        mDefaultText = .ReadProperty("DefaultText", "")
        mUseDefaultText = .ReadProperty("UseDefaultText", True)
        mEntryType = .ReadProperty("EntryType", 5)
    End With
    With TextBox
        .Alignment = PropBag.ReadProperty("Alignment", 0)
        .Appearance = PropBag.ReadProperty("Appearance", 0)
        .BackColor = PropBag.ReadProperty("BackColor_Normal", vbWhite)
        .BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
        .CausesValidation = PropBag.ReadProperty("CausesValidation", True)
        .DataField = PropBag.ReadProperty("DataField", "")
        .Enabled = PropBag.ReadProperty("Enabled", True)
        .FontBold = PropBag.ReadProperty("FontBold", False)
        .FontItalic = PropBag.ReadProperty("FontItalic", False)
        .FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
        .FontSize = PropBag.ReadProperty("FontSize", 8)
        .FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
        .FontUnderline = PropBag.ReadProperty("FontUnderline", False)
        .ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
        .Left = PropBag.ReadProperty("Left", 0)
        .Locked = PropBag.ReadProperty("Locked", False)
        .MaxLength = PropBag.ReadProperty("MaxLength", 0)
        .MousePointer = PropBag.ReadProperty("MousePointer", ctrlMousePointer.Default)
        .OLEDragMode = PropBag.ReadProperty("OLEDragMode", False)
        .OLEDropMode = PropBag.ReadProperty("OLEDropMode", False)
        .PasswordChar = PropBag.ReadProperty("PasswordChar", "")
        .RightToLeft = PropBag.ReadProperty("RightToLeft", False)
        .Text = PropBag.ReadProperty("Text", "")
   End With
   If mUseDefaultText Then TextBox.Text = mDefaultText
End Sub
Private Sub UserControl_Resize()
    TextBox.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("AutoSelect", mAutoSelectText, True)
        Call .WriteProperty("BeepOnBadType", mBeepOnBadType, False)
        Call .WriteProperty("AutoUpperCase", mAutoUpperCase, False)
        Call .WriteProperty("BackColor_Normal", mBackColor_Normal, vbWhite)
        Call .WriteProperty("BackColor_OnGotFocus", mBackColor_OnGotFocus, vbCyan)
        Call .WriteProperty("Appearance_Normal", mAppearance_Normal, 0)
        Call .WriteProperty("Appearance_OnGotFocus", mAppearance_OnGotFocus, 0)
        Call .WriteProperty("DefaultText", mDefaultText, "")
        Call .WriteProperty("EntryType", mEntryType, 5)
        Call .WriteProperty("UseDefaultText", mUseDefaultText, True)
    End With
    With TextBox
        Call PropBag.WriteProperty("Alignment", .Alignment, 0)
        Call PropBag.WriteProperty("Appearance", .Appearance, 0)
        Call PropBag.WriteProperty("BorderStyle", .BorderStyle, 1)
        Call PropBag.WriteProperty("CausesValidation", .CausesValidation, True)
        Call PropBag.WriteProperty("DataField", .DataField, "")
        Call PropBag.WriteProperty("Enabled", .Enabled, True)
        Call PropBag.WriteProperty("FontBold", .FontBold, False)
        Call PropBag.WriteProperty("FontItalic", .FontItalic, False)
        Call PropBag.WriteProperty("FontName", .FontName, "MS Sans Serif")
        Call PropBag.WriteProperty("FontSize", .FontSize, 8)
        Call PropBag.WriteProperty("FontStrikethru", .FontStrikethru, False)
        Call PropBag.WriteProperty("FontUnderline", .FontUnderline, False)
        Call PropBag.WriteProperty("ForeColor", .ForeColor, vbBlack)
        If IsArray(TextBox) Then Call PropBag.WriteProperty("Index", .Index)
        Call PropBag.WriteProperty("Left", .Left, 0)
        Call PropBag.WriteProperty("Locked", .Locked, False)
        Call PropBag.WriteProperty("MaxLength", .MaxLength, 0)
        Call PropBag.WriteProperty("MousePointer", .MousePointer, ctrlMousePointer.Default)
        Call PropBag.WriteProperty("MultiLine", .MultiLine, False)
        Call PropBag.WriteProperty("OLEDragMode", .OLEDragMode, ctrlOLEDragMode.Manual)
        Call PropBag.WriteProperty("OLEDropMode", .OLEDropMode, ctrlOLEDropMode.None)
        Call PropBag.WriteProperty("PasswordChar", .PasswordChar, "")
        Call PropBag.WriteProperty("RightToLeft", .RightToLeft, False)
        Call PropBag.WriteProperty("Text", .Text, "")
    End With
End Sub
