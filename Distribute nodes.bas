'______________________________________________________________
'  Name:       Distribute nodes
'  Author:     Eric LE GAL > https://github.com/yzEric
'  Version:    1.0
'  Date:       05/10/2021
'  Languages:  WinWrap
'  Purpose:
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
  Call femapMod.feAppMessage(FCM_NORMAL, "Distribute nodes  - Copyright (c) 2021 Eric LE GAL - MIT License (MIT)" )
  
  Set aSet= femapMod.feSet


'--- Select layers to Merge
  rc= aSet.Select( FT_NODE, True, "Select nodes to use" )
      If rc=FE_NOT_EXIST Then
        Call femapMod.feAppMessage(FCM_NORMAL, "Model must containt few nodes" )      
        Exit Sub
      End If      
      If rc=FE_CANCEL Then Exit Sub
      If AssertRC( femapMod, rc, "Unable to select nodes" ) Then Exit Sub

  If aSet.Count < 2 Then Exit Sub


  Dim Node As femap.Node, numNode As Long, nodes_ID As Variant, XYZ As Variant
  Set Node = femapMod.feNode
  rc = Node.GetCoordArray( aSet.ID, numNode, nodes_ID, XYZ )
    If AssertRC( femapMod, rc, "Unable to get nodes coordinates" ) Then Exit Sub 


  Dim sortIndex() As Long, Last As Long, i As Long, u As Long
  sortIndex =CountsToIndex( getCountingSort_lin3D( XYZ ) )
  Last = UBound(sortIndex)


  Dim CoeffA As Double, CoeffB As Double, xyzA As Variant, xyzB As Variant
  i =sortIndex(0)
' Call femapMod.feAppMessage( FCM_NORMAL, "First "+CStr( nodes_ID( i ) )   )
  u = i*3
  xyzA = Array( XYZ(u) , XYZ(u+1), XYZ(u+2) )

  i =sortIndex(Last)
' Call femapMod.feAppMessage( FCM_NORMAL, "Last "+CStr( nodes_ID( i ) )   )
  u = i*3
  xyzB = Array( XYZ(u) , XYZ(u+1), XYZ(u+2) )


  For i=1 To Last-1

    CoeffB = i /Last  :  CoeffA = 1 - CoeffB

     u = sortIndex(i)
 '  Call femapMod.feAppMessage( FCM_NORMAL, CStr( nodes_ID( u ) )   )
    u = u*3
    XYZ(u+0)= CoeffA * xyzA(0) + CoeffB * xyzB(0)
    XYZ(u+1)= CoeffA * xyzA(1) + CoeffB * xyzB(1)
    XYZ(u+2)= CoeffA * xyzA(2) + CoeffB * xyzB(2)

  Next i

  rc =Node.PutCoordArray( numNode, nodes_ID, XYZ )
    If AssertRC( femapMod, rc, "Unable to update nodes coordinates" ) Then Exit Sub

  Call femapMod.feViewRegenerate(0)

End Sub




Function CountsToIndex( counts() As Long ) As Long()

  Dim i As Long, Last As Long
  Last = UBound(counts)
  ReDim Index(Last) As Long
  For i=0 To Last
    Index( counts(i) ) = i
  Next i

   CountsToIndex= Index
End Function



Function getCountingSort_lin3D( XYZ As Variant) As Long()
  Dim MaxLoc As Long, MaxI As Long, maxLength As Double
  Dim iLoc As Long, Index As Long, iLength As Double

  Dim lastloc As Long
  lastloc = (UBound(XYZ)+1) /3 -1

  
  '--- start from first node and search most distant node
  For iLoc= 1 To lastloc
    Index=Index+3
    iLength = Sqr( (XYZ(0)-XYZ(Index))^2 + (XYZ(1)-XYZ(Index+1))^2 + (XYZ(2)-XYZ(Index+2))^2 )
    
    If iLength >= maxLength Then 
      MaxLoc = iLoc
      maxLength = iLength
    End If
  Next
  
  
  '--- Get all Distance
  ReDim Distances( lastloc ) As Double
  MaxI = MaxLoc * 3
  Index = 0
  For iLoc= 0 To lastloc
    Distances(iLoc) = Sqr( (XYZ(MaxI)-XYZ(Index))^2 + (XYZ(MaxI+1)-XYZ(Index+1))^2 + (XYZ(MaxI+2)-XYZ(Index+2))^2 )
    Index=Index+3
  Next
    
  getCountingSort_lin3D = CountingSort(Distances)
End Function




Function CountingSort(values As Variant) As Long()
  Dim i As Long, j As Long, lastIndex As Long
  lastIndex = UBound(values)
  ReDim Count(lastIndex) As Long
  For i = lastIndex To 0 Step -1
    For j = i - 1 To 0 Step -1
      If values(i) < values(j) Then
        Count(j) = Count(j) + 1
      Else
        Count(i) = Count(i) + 1
      End If
    Next j
  Next i

  CountingSort = Count
End Function



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
