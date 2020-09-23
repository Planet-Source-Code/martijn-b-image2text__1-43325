Attribute VB_Name = "mFontSize"
Option Explicit
Public Type SizeAPI
  Width   As Long
  Height  As Long
End Type
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Public Function GetTextSize(oFont As IFont, ByVal sText As String) As SizeAPI

  'Returns the Width and Height of sText in Pixels.
  Dim lRet        As Long
  Dim hDC         As Long
  Dim hOldFont    As Long
  Dim tSize       As SizeAPI
  On Error GoTo LocalError
  'Create a Device Context to draw text
  hDC = CreateDC("DISPLAY", vbNullString, vbNullString, vbNullString)

  If hDC <> 0 Then

    'Select the font into the DC
    hOldFont = SelectObject(hDC, oFont.hFont)
    'Get the TextWidth
    lRet = GetTextExtentPoint32(hDC, sText, Len(sText), tSize)
    'De-select the font and delete the DC
    lRet = SelectObject(hDC, hOldFont)
    lRet = DeleteDC(hDC)

  End If

  'Return the TextWidth
  GetTextSize.Width = tSize.Width
  GetTextSize.Height = tSize.Height
NormalExit:
  Exit Function
LocalError:
  MsgBox Err.Description, vbExclamation, "GetTextSize"
  Resume NormalExit

End Function

