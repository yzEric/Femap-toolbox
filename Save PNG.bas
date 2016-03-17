'______________________________________________________________
'  Name:       Save PNG
'  Author:     Eric LE GAL > https://github.com/yzEric
'  Version:    1.0 
'  Date:       17/03/2016
'  Languages:  WinWrap
'  Purpose:    Make capture of femap model and save it in png format in the model directory 
'______________________________________________________________
'
'
' The MIT License (MIT)
'
' Copyright (c) 2016 Eric LE GAL
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'______________________________________________________________



Option Explicit    '-> All variables must be declared
Option Base 0      '-> First index of Arrays is 0


Sub Main()
  
  Dim femapMod As femap.model, rc As Long
  Set femapMod = feFemap() 

  femapMod.feAppMessage( FCM_NORMAL , "Save PNG - Copyrigth (c) 2016 Eric LE GAL - See source file for more informations about the MIT license (MIT)")

  Dim modelPath As String, pictPath As String
  modelPath = femapMod.ModelName
  If modelPath = "" Then
     femapMod.feAppMessage( FCM_WARNING , "Model must be save before." ) : Exit Sub
  End If

  pictPath = modelPath +"_.PNG"   


  rc= femapMod.feFilePictureSave( False, False, FPM_PNG, pictPath )
      If AssertRC( femapMod, rc, "Unable to save preview, " + pictPath ) Then Exit Sub

  If femapMod.Info_Version <10.0 Then Call femapMod.feAppMessage(FCM_NORMAL, "Picture Save to Image Complete.  Image File: " + pictPath )

End Sub





Function AssertRC( femapMod As femap.model, rc As Long, msg As String )As Boolean
  If rc=FE_OK Then Exit Function

  Dim info As String
  Select Case rc
  Case FE_TOO_SMALL
    info = "Too small"
  Case FE_FAIL
    info = "Fail"
  Case FE_BAD_TYPE
    info = "Bad type"
  Case FE_CANCEL
    info = "Cancel"
  Case FE_BAD_DATA
    info = "Bad data"
  Case FE_INVALID
    info = "Invalid"
  Case FE_NO_MEMORY
    info = "No memory"
  Case FE_NOT_EXIST
    info = "Not Exist"
  Case FE_NEGATIVE_MASS_VOLUME
    info = "Negative mass volume"
  Case FE_SECURITY
    info = "Security"
  Case FE_NO_FILENAME
    info = "No file name"
  Case FE_NOT_AVAILABLE
    info = "Not available"
  Case Else
    info = "Unkonw return code: "+CStr(rc)
  End Select

  info = "Return code is '" +info+ "'"

  On Error Resume Next

  Call femapMod.feAppMessage(FCM_ERROR, msg )
  Call femapMod.feAppMessage(FCM_ERROR, info )

  Call MsgBox( msg + vbCrLf + info )
  AssertRC = True
End Function
