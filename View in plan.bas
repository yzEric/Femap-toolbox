'______________________________________________________________
'  Name:       View in plan
'  Author:     Eric LE GAL > https://github.com/yzEric
'  Version:    1.0
'  Date:       08/03/2016
'  Languages:  WinWrap
'  Purpose:    Adjust orientation of view to fit nearest plan 
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

  Dim femapMod As femap.model
  Set femapMod = feFemap() 

  Dim viewID As Long, aView As femap.View
  Set aView = femapMod.feView

  Call femapMod.feAppGetActiveView( viewID )
  Call aView.Get(viewID)

  aView.rotation(0) = CLng( aView.rotation(0) /90 ) *90
  aView.rotation(1) = CLng( aView.rotation(1) /90 ) *90
  aView.rotation(2) = CLng( aView.rotation(2) /90 ) *90

  Call aView.Put( 0 )
  Call femapMod.feViewRegenerate( 0 )

End Sub
