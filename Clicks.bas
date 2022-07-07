Attribute VB_Name = "MDL_Macros"
'Begin Code

Const KEYEVENTF_KEYUP = &H2
Const KEYEVENTF_EXTENDEDKEY = &H1
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
Const HWND_TOPMOST = -1
Public Type POINTAPI
        x As Long
        y As Long
End Type
Sub Macros(ByVal Boton As String)
Dim pt As POINTAPI

Select Case Boton

          '  Case "Hechizos"
           '     ClickIzquierdo frmCoord.h_x, frmCoord.h_y
    
          '  Case "Inventario"
          '      ClickIzquierdo frmCoord.i_x, frmCoord.i_y
          '  Case "Aimbot"
          '      ClickIzquierdo frmCoord.i_x, frmCoord.i_y
    
           ' Case "Lanzar"
           '     ClickIzquierdo frmCoord.l_x, frmCoord.l_y
    
         '   Case "Paralizar"
            '    ClickIzquierdo frmCoord.p_x, frmCoord.p_y
    
          '  Case "Especial"
            '    ClickIzquierdo frmCoord.e_x, frmCoord.e_y
    
          '  Case "AutoRemo"
              '  ClickIzquierdo frmCoord.h_x, frmCoord.h_y
              '  ClickIzquierdo frmCoord.r_x, frmCoord.r_y
              '  ClickIzquierdo frmCoord.l_x, frmCoord.l_y
                  '  Sleep (100)
                'ClickIzquierdo frmCoord.c_x, frmCoord.c_y
    

        Case "VolverCor"
            SetCursorPos xM, yM
            
        Case "BuscarCor"
            GetCursorPos pt
            xM = pt.x
            yM = pt.y
        
        Case "ClickCor"
          Call ClickIzquierdo(xM, yM)
End Select
   
End Sub


