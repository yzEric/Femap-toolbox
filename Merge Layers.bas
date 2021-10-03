'______________________________________________________________
'  Name:       Merge layers
'  Author:     Eric LE GAL > https://github.com/yzEric
'  Version:    1.0
'  Date:       03/10/2021
'  Languages:  WinWrap
'  Purpose:    Merge layers
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
  Call femapMod.feAppMessage(FCM_NORMAL, "Merge layers  - Copyright (c) 2021 Eric LE GAL - MIT License (MIT)" )
  
  Set aSet= femapMod.feSet


'--- Select destination layer
  Dim destLayerID As Long
  rc= aSet.SelectID( FT_LAYER, "Select destination layer", destLayerID )
      If rc=FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select layer" ) Then Exit Sub


'--- Select layers to Merge
  rc= aSet.Select( FT_LAYER, True, "Select layers to Merge with layer " + CStr(destLayerID) )
      If rc=FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select layers" ) Then Exit Sub

  rc= aSet.Remove( destLayerID )
      If AssertRC( femapMod, rc, "Unable to remove destination layer from list of layers to delete" ) Then Exit Sub

  Dim count As Long, layerIDs As Variant, i_layer As Long
  count = aSet.Count
  If count = 0 Then
    Call femapMod.feAppMessage(FCM_NORMAL, "You must select at least one layer different from destination layer" )
    Exit Sub
  End If


  rc= aSet.GetArray( count, layerIDs )
      If AssertRC( femapMod, rc, "Unable to get list of layer IDs" ) Then Exit Sub


'--- Delete layers
  rc= femapMod.feDelete(FT_LAYER, aSet.ID )
    If AssertRC( femapMod, rc, "Unable to delete layers" ) Then Exit Sub


'--- Renumber destination layer to catch content of deleted layers
  Dim PreviousID As Long, NewID As Long
  PreviousID = destLayerID

  For i_layer = 0 To count-1
     NewID= layerIDs(i_layer)

     rc= femapMod.feRenumber(FT_LAYER , -1*PreviousID , NewID)
       If AssertRC( femapMod, rc, "Unable To renumber a layer "+CStr(PreviousID)+" To "+CStr(NewID) ) Then Exit Sub

     PreviousID= NewID
  Next


  rc= femapMod.feRenumber(FT_LAYER , -1*PreviousID , destLayerID)
        If AssertRC( femapMod, rc, "Unable To renumber a layer "+CStr(PreviousID)+" To "+CStr(destLayerID) ) Then Exit Sub


  Call femapMod.feAppMessage(FCM_NORMAL, CStr(count)+ " layer" + IIf(count=1,"","s")+" removed"  )
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
