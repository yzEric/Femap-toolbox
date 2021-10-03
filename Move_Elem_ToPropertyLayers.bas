'______________________________________________________________
'  Name:       Move Elements to property layers
'  Author:     Eric LE GAL > https://github.com/yzEric
'  Version:    1.0
'  Date:       04/10/2021
'  Languages:  WinWrap
'  Purpose:    Moves elements to the layers defined for their properties
'______________________________________________________________
'
'
' The MIT License (MIT)
'
' Copyright (c) 2021 Eric LE GAL
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

  Dim femapMod As femap.model,  aSet As femap.Set , rc As Long
  Set femapMod = feFemap() 
  Call femapMod.feAppMessage(FCM_NORMAL, "Move Elements to property layers  - Copyright (c) 2021 Eric LE GAL - MIT License (MIT)" )
  
  Set aSet= femapMod.feSet


'--- Select properties to use
  rc= aSet.Select( FT_PROP, True, "Select properties to use" )
      If rc=FE_NOT_EXIST Then
        Call femapMod.feAppMessage(FCM_NORMAL, "Model must contain at least one property" )
      	Exit Sub
      End If
      If rc=FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select properties" ) Then Exit Sub
 

  Dim count As Long, IDs As Variant, i As Long
  count = aSet.Count
  If count = 0 Then
    Call femapMod.feAppMessage(FCM_NORMAL, "You must select at least one property" )
    Exit Sub
  End If


  rc= aSet.GetArray( count, IDs )
      If AssertRC( femapMod, rc, "Unable to get ID list of properties" ) Then Exit Sub

  Dim Prop As femap.Prop, propID As Long, layerID As Long 
  Set Prop = femapMod.feProp

  For i = 0 To count-1
     propID = IDs(i)
     rc = Prop.Get(propID)
       If AssertRC( femapMod, rc, "Unable to data of property ID:" + CStr(propID) ) Then Exit Sub
          
     aSet.Reset
	 aSet.AddRule(propID, FGD_ELEM_BYPROP)
     
     layerID = Prop.layer
     rc = femapMod.feModifyLayer(FT_ELEM, aSet.ID, layerID)
       If AssertRC( femapMod, rc, "Unable to move elements of property ID:" + CStr(propID)+" in layer ID: "+CStr(layerID) ) Then Exit Sub

  Next

  Call femapMod.feViewRedraw(0)
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
    info = "Unkonw (" + CStr(rc) + ")"
  End Select
  info = "Return code is '" +info+ "'"

  On Error Resume Next

  Call femapMod.feAppMessage(FCM_ERROR, msg )
  Call femapMod.feAppMessage(FCM_ERROR, info )

  Call MsgBox( msg + vbCrLf + info )
  AssertRC = True
End Function
