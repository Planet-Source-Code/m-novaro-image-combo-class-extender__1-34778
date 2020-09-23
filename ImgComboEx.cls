VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImgComboEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private blnAuto As Boolean 'Keeps the autocomplete functions from
                           'triggering the Change event

Private WithEvents myCombo As ImageCombo
Attribute myCombo.VB_VarHelpID = -1

Private m_OnlyPresentItems As Boolean
Private m_bItemFound As Boolean
Private m_bShowWithKeyDown As Boolean
Private m_bSelectOnEntry As Boolean

Private eventGenerated As Boolean

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_LIMITTEXT = &H141
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETLBTEXTLEN = &H149
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17 'win95/98 only
Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Const SM_CXVSCROLL = 2
Private Const CB_GETEXTENDEDUI = &H156
Private Const CB_SETEXTENDEDUI = &H155
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF&
Private Const EM_SETMARGINS = &HD3&
Private Const EM_GETMARGINS = &HD4&

Private Type SIZE
  cx As Long
  cy As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function ReleaseDC Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal hDc As Long) As Long
   
Private Declare Function GetDC Lib "user32" _
  (ByVal hwnd As Long) As Long
  
Private Declare Function GetSystemMetrics Lib "user32" _
  (ByVal nIndex As Long) As Long
  
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DrawText Lib "user32" _
   Alias "DrawTextA" _
  (ByVal hDc As Long, _
   ByVal lpStr As String, _
   ByVal nCount As Long, _
   lpRect As RECT, _
   ByVal wFormat As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDc As Long, ByVal hObject As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" _
    Alias "GetTextExtentPoint32A" _
   (ByVal hDc As Long, _
    ByVal lpsz As String, _
    ByVal cbString As Long, _
    lpSize As SIZE) As Long

Private Declare Function GetStockObject Lib "gdi32" _
  (ByVal nIndex As Long) As Long
  
Private Declare Function DeleteObject Lib "gdi32" _
  (ByVal hObject As Long) As Long
  

Public Event ItemAccepted(theItem As String, bItemFound As Boolean)
Property Let AllowOnlyPresentItems(allow As Boolean)
    m_OnlyPresentItems = allow
End Property

Property Get AllowOnlyPresentItems() As Boolean
    AllowOnlyPresentItems = m_OnlyPresentItems
End Property

Public Sub SetRefToCombo(theCombo As ImageCombo)
    Set myCombo = theCombo
End Sub

Private Sub Class_Initialize()
    m_OnlyPresentItems = False
    ' Metto questo a true per default. Infatti, se non fosse così, l'effetto autocompletamento
    ' permette solo di selezionare un elemento alla volta.
    m_bSelectOnEntry = True
    eventGenerated = True
End Sub

Private Sub myCombo_Click()
    ' Nota: in questo caso, l'evento va generato comunque
    Debug.Print "Evento generato myCombo_Click" & myCombo.Text; " "; True
    eventGenerated = True
    RaiseEvent ItemAccepted(myCombo.SelectedItem.Text, True)
End Sub

Private Sub myCombo_GotFocus()
    eventGenerated = True
    If m_bSelectOnEntry = True Then
        myCombo.SelStart = 0
        myCombo.SelLength = Len(myCombo.Text)
    End If
End Sub

Private Sub myCombo_KeyPress(KeyAscii As Integer)
    On Error GoTo myH
    Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
    Dim found As Boolean

'    If Chr(KeyAscii) = "a" Then Set myCombo = Nothing  ' Err.Raise 91

    eventGenerated = False
    iStart = myCombo.SelStart
    If KeyAscii = vbKeyBack Then
        If myCombo.SelStart <= 1 Then
            myCombo.Text = ""
            KeyAscii = 0
            Exit Sub
        End If
        If myCombo.SelLength = 0 Then
            strPart = Left(myCombo.Text, Len(myCombo.Text) - 1)
        Else
            strPart = Left$(myCombo.Text, myCombo.SelStart - 1)
        End If
    ElseIf KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        myCombo_LostFocus
        myCombo.Refresh
        Exit Sub
    ElseIf KeyAscii < 32 Or (KeyAscii > 127 And KeyAscii < 161) Then
        Exit Sub
    Else
        If myCombo.SelLength = 0 Then
            strPart = myCombo.Text & Chr$(KeyAscii)
        Else
            strPart = Left$(myCombo.Text, myCombo.SelStart) & Chr$(KeyAscii)
        End If
    End If

    For iLoop = 1 To myCombo.ComboItems.Count
        'compare each item to the part the user has typed,
        '"complete" with the first good match
        strItem = UCase$(myCombo.ComboItems.Item(iLoop))
        If strItem Like UCase$(strPart & "*") Then
            'partial match but not the whole thing.
            '(if whole thing, nothing to complete!)
            Set myCombo.SelectedItem = myCombo.ComboItems.Item(iLoop)
            myCombo.SelStart = Len(strPart)
            myCombo.SelLength = Len(myCombo.Text) - myCombo.SelStart
            KeyAscii = 0

            found = True
            Exit For
        End If
    Next iLoop
    m_bItemFound = found
    If Not found Then
        If m_OnlyPresentItems Then
            KeyAscii = 0
        Else
            'niente
        End If
    End If
    blnAuto = False

    Debug.Print strPart & " " & myCombo.SelStart & " " & myCombo.SelLength & " " & found
    Exit Sub

myH:
    MsgBox "ERR " & Err.Number & Err.Description
End Sub

Private Sub myCombo_LostFocus()
    Dim iLoop As Integer

    Debug.Print "Combo lost focus"

    'Match capitalization if item entered is one on the list
    m_bItemFound = False
    If myCombo.Text <> "" Then
        For iLoop = 1 To myCombo.ComboItems.Count
            If UCase$(myCombo.ComboItems.Item(iLoop)) = UCase$(myCombo.Text) Then
                blnAuto = True
                myCombo.ComboItems.Item(iLoop).Selected = True
                m_bItemFound = True
                blnAuto = False
                Exit For
            End If
        Next iLoop
    End If
    If Not eventGenerated Then
       Debug.Print "Evento generato myCombo_LostFocus" & myCombo.Text; " "; m_bItemFound
       eventGenerated = True
       RaiseEvent ItemAccepted(myCombo.Text, m_bItemFound)
    End If
End Sub

Public Sub ShowDropDown()
    SendMessageAsLong myCombo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub

Public Sub HideDropDown()
    SendMessageAsLong myCombo.hwnd, CB_SHOWDROPDOWN, False, 0
End Sub

Public Property Get ItemFound() As Boolean
    ItemFound = m_bItemFound
End Property

Public Property Let ItemFound(found As Boolean)
    ' Chiama questa procedura per essere sicuro che l'item sia nella lista
    ' Quindi non basta mettere l'itemFound a true, ma occorre davvero
    ' aggiungere l'item al combo perchè questo funzioni
    myCombo_LostFocus
End Property

Function GetDroppedState() As Boolean
    GetDroppedState = SendMessageAsLong(myCombo.hwnd, CB_GETDROPPEDSTATE, 0, 0)
End Function

Sub SetMaxLength(ByVal MaxLength As Long)
    SendMessageAsLong myCombo.hwnd, CB_LIMITTEXT, MaxLength, 0
End Sub

Public Property Get SelectOnEntry() As Boolean
    SelectOnEntry = m_bSelectOnEntry
End Property

Public Property Let SelectOnEntry(ByVal bSelectOnEntry As Boolean)
    m_bSelectOnEntry = bSelectOnEntry
End Property

Public Sub AddItemSorted(newItem As String, theKey As Long, Optional IconKey As String, Optional isSelected As Boolean = True)
    Dim idx As Integer

    ' Questo serve per aggiungere l'elemento in modo ordinato nella lista
    ' Cerco la posizione corretta
    For idx = 1 To myCombo.ComboItems.Count
        If StrComp(myCombo.Text, myCombo.ComboItems.Item(idx).Text, vbTextCompare) < 0 Then
            Exit For
        End If
    Next idx

    If IconKey <> "" Then
        myCombo.ComboItems.add idx, CStr(theKey) & "_", newItem, IconKey
    Else
        myCombo.ComboItems.add idx, CStr(theKey) & "_", newItem
    End If
    myCombo.Refresh
    If isSelected Then
        myCombo.ComboItems.Item(idx).Selected = True 'Questo serve per far vedere subito l'icona.
    End If
    m_bItemFound = True
End Sub

Property Let DropDownWidth(newVal As Long)

    ' Resize the dropdown portion of the combo box using SendMessage
    Call SendMessage(myCombo.hwnd, CB_SETDROPPEDWIDTH, newVal, ByVal 0)
End Property

Property Get DropDownWidth() As Long

    DropDownWidth = SendMessage(myCombo.hwnd, CB_GETDROPPEDWIDTH, 0, ByVal 0)
End Property

Public Sub AutosizeDropDown()
   Dim cwidth As Long
   Dim i As Long
   Dim NumOfChars As Long
   Dim LongestComboItem As Long
   Dim avgCharWidth As Long
   Dim NewDropDownWidth As Long

  'loop through the combo entries, using SendMessage
  'with CB_GETLBTEXTLEN to determine the longest item
  'in the dropdown portion of the combo

   For i = 0 To myCombo.ComboItems.Count - 1

      NumOfChars = SendMessage(myCombo.hwnd, CB_GETLBTEXTLEN, i, ByVal 0)
      If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars

    Next

  'get the average size of the characters using the
  'GetFontDialogUnits API. Because a dummy string is
  'used in GetFontDialogUnits, avgCharWidth is an
  'approximation based on that string.
   avgCharWidth = GetFontDialogUnits()

  'compute the size the dropdown needs to be to accommodate
  'the longest string. Here I subtract 2 because I find that
  'on my system, using the dummy string in GetFontDialogUnits,
  'the width is just a bit too wide.
   NewDropDownWidth = (LongestComboItem - 2) * avgCharWidth

  'resize the dropdown portion of the combo box
   Call SendMessage(myCombo.hwnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, ByVal 0)

End Sub

Private Function GetFontDialogUnits() As Long
   Dim hFont As Long
   Dim hFontOld As Long
   Dim r As Long
   Dim avgWidth As Long
   Dim hDc As Long
   Dim tmp As String
   Dim sz As SIZE
   
   'get the hdc to the main window
   hDc = GetDC(myCombo.Container.hwnd)
   
  'with the current font attributes, select the font
   hFont = GetStockObject(ANSI_VAR_FONT)
   hFontOld = SelectObject(hDc, hFont&)

  'get its length, then calculate the average character width
   tmp = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
   GetTextExtentPoint32 hDc, tmp, 52, sz
   avgWidth = (sz.cx \ 52)
   
  're-select the previous font & delete the hDc
   SelectObject hDc, hFontOld
   DeleteObject (hFont)
   ReleaseDC myCombo.Container.hwnd, hDc

  'return the average character width
   GetFontDialogUnits = avgWidth

End Function


Property Let DropWithDownKey(ByVal bState As Boolean)
    ' Set whether combo box drops down using the Down Arrow or not:
    SendMessageLong myCombo.hwnd, CB_SETEXTENDEDUI, Abs(bState), 0
End Property

Public Property Get DropWithDownKey() As Boolean
    ' Get whether combo box drops down using the Down Arrow or not:
    DropWithDownKey = (SendMessageLong(myCombo.hwnd, CB_GETEXTENDEDUI, 0, 0) <> 0)
End Property
