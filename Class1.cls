VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Begin Code

'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08/03/07
' Created and Implemented by Rapsodius
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

'Customizes User Keys.
'
'@author Rapsodius
'@version 1.1.0
'@date 20080413

'03/08/2007 - Rapsodius
'   - First Release
'
'13/04/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Removed lots of unwanted buggy features keeping it clean. Revised documentation and coding standards.

Option Explicit

'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 9

'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\BindKeys.bin"

'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyToggleCombatMode Key for toggling combat mode
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyToggleFPS Key for toggling the FPS counter
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKey_ActivateCheat = 1
    mKey_AutoLanzar
    mKey_AutoAim
    mKey_AutoRemo
    mKey_HechizoUno
    mKey_HechizoDos
    mKey_LanzarUno
    mKey_OcultarCheat
    mKey_AutoPotas
End Enum

'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid
Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
Select Case KeyCode
    Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
        ReadableName = Chr$(KeyCode)
        
    Case KeyCodeConstants.vbKey0 To KeyCodeConstants.vbKey9
        ReadableName = Chr$(KeyCode)
    
    Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
        ReadableName = "NPad " & KeyCode - KeyCodeConstants.vbKeyNumpad0 '& " (teclado numérico)"
    
    Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
        ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
    
    Case KeyCodeConstants.vbKeyMultiply
        ReadableName = "NumPad *"
    
    Case KeyCodeConstants.vbKeyAdd
        ReadableName = "NumPad +"
    
    Case KeyCodeConstants.vbKeySeparator
        'Note: Separator appears in MSDN as the Enter key for
        'the numpad, but MS recognizes it as the normal
        'Enter key
        ReadableName = "Intro"
    
    Case KeyCodeConstants.vbKeySubtract
        ReadableName = "NumPad -"
    
    Case KeyCodeConstants.vbKeyDecimal
        ReadableName = "NumPad ."
        
    Case KeyCodeConstants.vbKeyTab
        ReadableName = "Tab"
        
    Case KeyCodeConstants.vbKeyDivide
        ReadableName = "NumPad /"
    
    Case KeyCodeConstants.vbKeyShift
        ReadableName = "Shift"
    
    Case KeyCodeConstants.vbKeyControl
        ReadableName = "Ctrl"
    
    Case KeyCodeConstants.vbKeyMenu 'Alt
        ReadableName = "Alt"
    
    Case KeyCodeConstants.vbKeyPause
        ReadableName = "Pausa"
    
    Case KeyCodeConstants.vbKeyCapital
        ReadableName = "Blq Mayús"
    
    Case KeyCodeConstants.vbKeyEscape
        ReadableName = "Esc"
    
    Case KeyCodeConstants.vbKeyPageUp
        ReadableName = "Repág"
    
    Case KeyCodeConstants.vbKeyPageDown
        ReadableName = "Avpág"
    
    Case KeyCodeConstants.vbKeyEnd
        ReadableName = "Fin"
    
    Case KeyCodeConstants.vbKeyHome
        ReadableName = "Inicio"
    
    Case KeyCodeConstants.vbKeyLeft
        ReadableName = "Izquierda"
    
    Case KeyCodeConstants.vbKeyUp
        ReadableName = "Arriba"
    
    Case KeyCodeConstants.vbKeyRight
        ReadableName = "Derecha"
    
    Case KeyCodeConstants.vbKeyDown
        ReadableName = "Abajo"
    
    Case KeyCodeConstants.vbKeyInsert
        ReadableName = "Insert"
    
    Case KeyCodeConstants.vbKeyNumlock
        ReadableName = "Bloq Num"
    
    Case KeyCodeConstants.vbKeyScrollLock
        ReadableName = "Bloq Despl"
    
    Case KeyCodeConstants.vbKeySpace
        ReadableName = "SpaceBar"
    
    Case KeyCodeConstants.vbKeyBack
        ReadableName = "Borrar"
    
    Case KeyCodeConstants.vbKeyReturn
        ReadableName = "Enter"
    
    Case KeyCodeConstants.vbKeyDelete
        ReadableName = "Supr"
        
    Case KeyCodeConstants.vbKeyLButton
        ReadableName = "Mouse1"
        
    Case KeyCodeConstants.vbKeyMButton
        ReadableName = "Mouse3"
        
    Case KeyCodeConstants.vbKeyRButton
        ReadableName = "Mouse2"
    
    Case 186
        ReadableName = "`"
    
    Case 187
        ReadableName = "+"
        
    Case 188
        ReadableName = ","
        
    Case 189
        ReadableName = "-"
        
    Case 190
        ReadableName = "."
        
    Case 191
        ReadableName = "ç"
        
    Case 219
        ReadableName = "'"
        
    Case 220
        ReadableName = "º"
        
    Case 221
        ReadableName = "¡"
        
    Case 222
        ReadableName = "´"
        
    Case Else
        'In all other cases, key is considered not valid.
        ReadableName = Constants.vbNullString
End Select
'MsgBox KeyCode
End Function

'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.
Public Sub LoadCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/21/07
'
'***************************************************
Dim i As Long
Dim FileNum As Integer
FileNum = FreeFile()

'Inits keyboard config with default configurations
Call LoadDefaults
    
Open App.Path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum

For i = 1 To NUM_CUSTOM_KEYS
    Get FileNum, , mappedKeys(i)
    If LenB(ReadableName(mappedKeys(i))) = 0 Then
        Close FileNum
        Call Err.Raise(ERROR_NO_KEY_FILE)
        Exit Sub
    End If
Next i

Close FileNum
End Sub

'Saves custom keys to a file
Public Sub SaveCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
Dim FileNum As Integer
Dim i As Long
FileNum = FreeFile()

Open App.Path & KEY_CONFIG_FILE For Binary Access Write As FileNum

For i = 1 To NUM_CUSTOM_KEYS
    Put FileNum, , mappedKeys(i)
Next i

Close #FileNum
End Sub

'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index
Public Property Get BindedKey(ByVal index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
If index < 1 Or index > NUM_CUSTOM_KEYS Then Exit Property
BindedKey = mappedKeys(index)
End Property

'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code
Public Property Let BindedKey(ByVal index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 08/05/07
'
'***************************************************
If index < 1 Or index > NUM_CUSTOM_KEYS Then Exit Property

If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
    Exit Property                      'Rejects the new assignment
End If

If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
    Exit Property
End If

mappedKeys(index) = NewVal
End Property

'Loads default keys
Public Sub LoadDefaults()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************


mappedKeys(eKeyType.mKey_ActivateCheat) = KeyCodeConstants.vbKeyInsert
mappedKeys(eKeyType.mKey_AutoLanzar) = KeyCodeConstants.vbKey9
mappedKeys(eKeyType.mKey_AutoAim) = KeyCodeConstants.vbKey1
mappedKeys(eKeyType.mKey_AutoRemo) = KeyCodeConstants.vbKeyTab
mappedKeys(eKeyType.mKey_HechizoUno) = KeyCodeConstants.vbKeyQ
mappedKeys(eKeyType.mKey_HechizoDos) = KeyCodeConstants.vbKeyE
mappedKeys(eKeyType.mKey_LanzarUno) = KeyCodeConstants.vbKey2
mappedKeys(eKeyType.mKey_OcultarCheat) = KeyCodeConstants.vbKeyF12
mappedKeys(eKeyType.mKey_AutoPotas) = KeyCodeConstants.vbKeyF1
End Sub

'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise
Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
Dim Counter As Long

KeyAssigned = False

For Counter = 1 To NUM_CUSTOM_KEYS
    If mappedKeys(Counter) = KeyCode Then
        KeyAssigned = True
        Exit For
    End If
Next Counter
End Function

'Retrieves the number of customizable keys
'
'@return The number of customizable keys
Public Property Get Count() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
Count = NUM_CUSTOM_KEYS
End Property

'Constructor. Loads Keys from file
Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
On Error GoTo ErrHand:
Call LoadCustomKeys
Exit Sub

ErrHand:
If Err.Number = ERROR_NO_KEY_FILE Then
    Call LoadDefaults
    Call SaveCustomKeys
    Resume Next
End If
End Sub

'Destructor. Saves Keys to file
Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
'Call SaveCustomKeys
End Sub

