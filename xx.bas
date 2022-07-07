Attribute VB_Name = "Module1"
'Begin Code

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_TEXT As Long = &H1
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETITEM As Long = (LVM_FIRST + 6)

Public hwndInventario As Long
Public hwndHechizos As Long
Private Type LVITEM
   mask As Long
   iItem As Long
   iSubItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   lParam As Long
   iIndent As Long
End Type

Public Declare Function EnumWindows Lib "User32" _
  (ByVal lpEnumFunc As Long, _
   ByVal lParam As Long) As Long
   
Public Declare Function EnumChildWindows Lib "User32" _
  (ByVal hWndParent As Long, _
   ByVal lpEnumFunc As Long, _
   ByVal lParam As Long) As Long

Private Declare Function GetWindowTextLength Lib "User32" _
    Alias "GetWindowTextLengthA" _
   (ByVal hwnd As Long) As Long
'' SACAR
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
'' SACAR
Private Declare Function GetWindowText Lib "User32" _
    Alias "GetWindowTextA" _
   (ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Private Declare Function GetClassName Lib "User32" _
    Alias "GetClassNameA" _
   (ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Public Declare Function IsWindowVisible Lib "User32" _
   (ByVal hwnd As Long) As Long
   
Private Declare Function GetParent Lib "User32" _
   (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "User32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Public Function EnumWindowProc(ByVal hwnd As Long, _
                               ByVal lParam As Long) As Long
   
  'working vars
   Dim nSize As Long
   Dim sTitle As String
   Dim sClass As String
   
   Dim sIDType As String
   Dim itmX As ListItem
   Dim nodX As Node
   
  'eliminate windows that are not top-level.
   If GetParent(hwnd) = 0& And _
      IsWindowVisible(hwnd) Then
      
     'get the window title / class name
      sTitle = GetWindowIdentification(hwnd, sIDType, sClass)

     'add to the listview
      Set itmX = frmHwnd.ListView1.ListItems.add(Text:=sTitle, key:=CStr(hwnd) & "h")
      'itmX.SmallIcon = Form1.ImageList1.ListImages("parent").Key
      itmX.SubItems(1) = CStr(hwnd)
      itmX.SubItems(2) = sIDType
      itmX.SubItems(3) = sClass
      
   End If
   
  'To continue enumeration, return True
  'To stop enumeration return False (0).
  'When 1 is returned, enumeration continues
  'until there are no more windows left.
   EnumWindowProc = 1
   
End Function


Private Function GetWindowIdentification(ByVal hwnd As Long, _
                                         sIDType As String, _
                                         sClass As String) As String

   Dim nSize As Long
   Dim sTitle As String

  'get the size of the string required
  'to hold the window title
   nSize = GetWindowTextLength(hwnd)
   
  'if the return is 0, there is no title
   If nSize > 0 Then
   
      sTitle = Space$(nSize + 1)
      Call GetWindowText(hwnd, sTitle, nSize + 1)
      sIDType = "title"
      
      sClass = Space$(64)
      Call GetClassName(hwnd, sClass, 64)
   
   Else
   
     'no title, so get the class name instead
      sTitle = Space$(64)
      Call GetClassName(hwnd, sTitle, 64)
      sClass = sTitle
      sIDType = "class"
   
   End If
   
   GetWindowIdentification = TrimNull(sTitle)

End Function


Public Function EnumChildProc(ByVal hwnd As Long, _
                              ByVal lParam As Long) As Long
   
  'working vars
   Dim sTitle As String
   Dim sClass As String
   Dim sIDType As String
   Dim itmX As ListItem

  'get the window title / class name
   sTitle = GetWindowIdentification(hwnd, sIDType, sClass)

  'add to the listview
   Set itmX = frmHwnd.ListView1.ListItems.add(, , sTitle)
   itmX.SubItems(1) = hwnd
   itmX.SubItems(2) = sIDType
   itmX.SubItems(3) = sClass
      
   Listview_IndentItem frmHwnd.ListView1.hwnd, CLng(itmX.index), 1
   
   EnumChildProc = 1
   
End Function


Private Function TrimNull(startstr As String) As String

  Dim pos As Integer

  pos = InStr(startstr, Chr$(0))
  
  If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
  End If
  
 'if this far, there was
 'no Chr$(0), so return the string
  TrimNull = startstr
  
End Function


Private Sub Listview_IndentItem(hwnd As Long, _
                                nItem As Long, _
                                nIndent As Long)

   Dim LV As LVITEM

  'if nIndent indicates that indentation
  'is requested nItem is the item to indent
   If nIndent > 0 Then
      
      With LV
        .mask = LVIF_INDENT
        .iItem = nItem - 1 '0-based
        .iIndent = nIndent
      End With
      
      Call SendMessage(hwnd, LVM_SETITEM, 0&, LV)
      
   End If
       
End Sub




