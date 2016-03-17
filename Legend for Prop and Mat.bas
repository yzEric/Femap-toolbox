'______________________________________________________________
'    Name:       Show properties / materials Legends
'    Author:     E. LE GAL
'    Version:    1.0
'    Date:       08/03/2016
'    Languages:  WinWrap
'    Purpose: Show legends of properties or materials
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



Option Explicit
Option Base 0

Dim femapMod As femap.model
Dim mySet As femap.Set
Dim mySet_isProp As Boolean


Dim opt_textColor As Long
Dim opt_font As Long
Dim opt_border As Boolean
Dim opt_groupByProp As Boolean
Dim opt_selectAll As Boolean
Dim opt_infoProp_ID As Boolean
Dim opt_infoProp_title As Boolean
Dim opt_infoProp_type As Boolean
Dim opt_infoProp_thickness As Boolean
Dim opt_infoMat_ID As Boolean
Dim opt_infoMat_title As Boolean
Dim opt_infoMat_type As Boolean
Dim opt_position_onLeft As Boolean
Dim opt_position_Verti As Long
Dim opt_position_offsetV As Long
Dim opt_position_offsetH As Long
Dim opt_position_spacing As Long
Dim opt_visibility_All As Boolean
Dim opt_visibility_viewID As Long

Const positionV_top=0
Const positionV_center=1
Const positionV_bottom=2

Const kLayerTitle="Legend_Prop_Mat"


'-------------------------------------------------------------------------


Sub Main

'--- Link to femap
    Set femapMod = feFemap()
    femapMod.feAppMessage( FCM_NORMAL , "Legend for properties and materials , Copyrigth (c) 2016 Eric LE GAL - See source file for more informations about the MIT license (MIT)")

'--- get femap font list
    Dim fonts() As String
    fonts=getFontList

'--- get list of views in model
    Dim views() As String
    views=getViewsList


'--- UserDialog
	Begin Dialog UserDialog 560,336,"Setup Legend",.dialogFunct ' %GRID:10,4,1,1
		GroupBox 10,0,550,68,"",.GroupBox_Main
		text 20,16,40,16,"Font",.Text5,1
		DropListBox 70,12,190,20,fonts(),.DropListBox_font

		text 270,20,140,16,"Text and border color",.Text4,1
		TextBox 420,16,50,20,.TextBox_color
		PushButton 480,16,70,20,"Palette",.PushButton_color

		CheckBox 460,44,80,16,"Border",.CheckBox_border

		GroupBox 10,72,220,76,"Colors and groups by",.GroupBox_colorsBy
		OptionGroup .Group_byPropOrMat
			OptionButton 20,92,90,20,"Properties",.OptionButton_byProp
			OptionButton 20,112,90,20,"Materials",.OptionButton_byMat
		CheckBox 130,96,90,12,"Select All",.CheckBox_selectAll
		PushButton 130,116,80,20,"Select",.PushButton_Select

		GroupBox 10,156,220,96,"Properties info",.GroupBox_propInfo
		CheckBox 30,172,70,16,"ID",.CheckBox_propInfo_ID
		CheckBox 30,192,80,16,"Title",.CheckBox_propInfo_title
		CheckBox 30,212,120,16,"Type",.CheckBox_propInfo_type
		CheckBox 30,232,120,16,"Plate thickness",.CheckBox_propInfo_Th

		GroupBox 10,256,220,80,"Material info",.GroupBox_matInfo
		CheckBox 30,272,120,16,"ID",.CheckBox_matInfo_ID
		CheckBox 30,292,110,16,"Title",.CheckBox_matInfo_title
		CheckBox 30,312,70,16,"Type",.CheckBox_matInfo_type

		GroupBox 240,72,320,132,"Position",.GroupBox_position
		OptionGroup .GroupPositionH
			OptionButton 260,88,90,16,"On left",.OptionButton_OnLeft
			OptionButton 260,108,100,16,"On right",.OptionButton_OnRight
		TextBox 460,92,80,20,.TextBox_offsetH
		text 380,96,70,16,"offset %",.Text1,1

		OptionGroup .GroupPositionV
			OptionButton 260,140,80,16,"On top",.OptionButton_OnTop
			OptionButton 260,156,70,16,"Center",.OptionButton_Fill
			OptionButton 260,176,100,16,"On bottom",.OptionButton_OnBottom
		TextBox 460,148,80,20,.TextBox_offsetV
		text 390,152,60,16,"offset %",.Text2,1
		TextBox 460,176,80,20,.TextBox_spacing
		text 390,180,60,12,"spacing",.Text6,1

		GroupBox 240,216,320,80,"Visibility",.GroupBox_visibility
		OptionGroup .Group_visibilityView
			OptionButton 270,228,230,16,"All views",.OptionButton_allViews
			OptionButton 270,244,200,20,"Single View",.OptionButton_singleView
		DropListBox 330,268,220,20,views(),.DropListBox_view

		PushButton 240,304,70,24,"Clear",.PushButton_Clear
		PushButton 320,304,70,24,"Exit",.PushButtonExit
		PushButton 400,304,70,24,"Apply",.PushButton_Show
		OKButton 480,304,70,24,.PushButton_OK


	End Dialog

'--- Set default values
	Call options_iniDefaultValues

'--- Try to read last saved user data
    Call options_readUserdata

'--- Ini user dialog with data
	Dim dlg As UserDialog
	dlg.TextBox_color = CStr(  opt_textColor )
	dlg.DropListBox_font = IndexOf( fonts ,	opt_font )
    dlg.CheckBox_border = opt_border
    dlg.Group_byPropOrMat =IIf( opt_groupByProp  , 0 , 1 )
    dlg.CheckBox_selectAll = opt_selectAll

    dlg.CheckBox_propInfo_ID = opt_infoProp_ID
    dlg.CheckBox_propInfo_title = opt_infoProp_title
    dlg.CheckBox_propInfo_type = opt_infoProp_type
    dlg.CheckBox_propInfo_th = opt_infoProp_thickness

    dlg.CheckBox_matInfo_ID = opt_infoMat_ID
    dlg.CheckBox_matInfo_title = opt_infoMat_title
    dlg.CheckBox_matInfo_type = opt_infoMat_type

    dlg.GroupPositionH = IIf( opt_position_onLeft , 0 , 1 )
    dlg.GroupPositionV = opt_position_Verti
    dlg.TextBox_offsetH = CStr( opt_position_offsetH )
    dlg.TextBox_offsetV = CStr( opt_position_offsetV )
    dlg.TextBox_spacing = CStr( opt_position_spacing )

    dlg.DropListBox_view = IndexOf(  views , opt_visibility_viewID )
    dlg.Group_visibilityView = IIf( opt_visibility_All , 0 , 1 )

'--- Show user dialog
	Dialog dlg

End Sub


'-------------------------------------------------------------------------


Private Function dialogFunct(DlgItem$, Action%, SuppValue?) As Boolean

   ' Usefull for debug
   '   femapMod.feAppMessage( FCM_COMMAND , "item: "+DlgItem+" action: "+CStr(Action)+" SuppValue: "+CStr(SuppValue))

    Dim v As Variant

	Select Case Action%

	Case 1 ' Dialog box initialization
        v = DlgValue( "Group_byPropOrMat" )
        DlgEnable( "CheckBox_propInfo_ID" , v<>1 )
        DlgEnable( "CheckBox_propInfo_title" , v<>1 )
        DlgEnable( "CheckBox_propInfo_type" , v<>1 )
        DlgEnable( "CheckBox_propInfo_Th" , v<>1 )


        v = DlgValue( "GroupPositionV" )
        DlgEnable( "TextBox_offsetV" , v<>1 )

        v = DlgValue( "Group_visibilityView" )
        DlgEnable( "DropListBox_view" , v<>0 )

	Case 2 ' Value changing or button pressed
		Select Case DlgItem

        Case "PushButton_color"
           Call dialog_color
           dialogFunct = True

        Case "PushButton_Clear"
           dialog_removeAllLegend( True, True )
           dialogFunct = True

        Case "Group_byPropOrMat"
        	DlgEnable( "CheckBox_propInfo_ID" , SuppValue<>1 )
        	DlgEnable( "CheckBox_propInfo_title" , SuppValue<>1 )
        	DlgEnable( "CheckBox_propInfo_type" , SuppValue<>1 )
        	DlgEnable( "CheckBox_propInfo_Th" , SuppValue<>1 )

        Case "GroupPositionV"
        	DlgEnable( "TextBox_offsetV" , SuppValue<>1 )

        Case "Group_visibilityView"
        	DlgEnable( "DropListBox_view" , SuppValue<>0 )

        Case "PushButton_Select"
             Call selectEntitiesToShow
            dialogFunct =True

        Case "PushButtonExit"
            'Nothing todo

        Case "PushButton_Show" , "PushButton_OK"
            If options_LoadAllFromDialog Then
               dialogFunct =True
               Exit Function
            End If
        	Call dialog_removeAllLegend( False, False )
        	Call dialog_showLegends
        	Call options_SaveUserdata
            If DlgItem="PushButton_Show" Then dialogFunct = True
		End Select

	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		'Wait .1 : dialogFunct = True
	Case 6 ' Function key
	End Select
End Function


'-------------------------------------------------------------------------


Sub dialog_showLegends()
  Dim rc As Long
  Dim entType As Long, entName As String
  entType = IIf( opt_groupByProp , FT_PROP , FT_MATL )
  entName =IIf( opt_groupByProp , "properties" , "materials" )

   If mySet Is Nothing Then Set mySet = femapMod.feSet
   If mySet_isProp <> opt_groupByProp Then mySet.Clear

   If opt_selectAll Then
        rc= mySet.AddAll( entType )
               If AssertRC( femapMod, rc, "Unable to select all "+entName ) Then Exit Sub
        mySet_isProp= ( entType= FT_PROP )

   ElseIf mySet.Count =0 Then
   	  mySet_isProp =opt_groupByProp
      rc= mySet.Select( entType , True , "Select "+entName+" to use in legend" )
             If rc= FE_CANCEL Then Exit Sub
             If AssertRC( femapMod, rc, "Unable to select  "+entName ) Then Exit Sub

   End If
   If mySet.Count = 0 Then Exit Sub


  Dim boxText As femap.text, boxColor As femap.text, itemID As Long, posVertic As Long, posStep As Long
  Dim aProp As femap.Prop, aMat As femap.Matl , aColor As Long, info As String
  Set aProp = femapMod.feProp
  Set aMat = femapMod.feMatl

  Dim layerID As Long, aLayer As femap.layer
  layerID = getLayerByTitle( kLayerTitle )
  If layerID=-1 Then
    Set aLayer = femapMod.feLayer
    layerID= aLayer.NextEmptyID
    aLayer.title = kLayerTitle
    rc= aLayer.Put( layerID )
        If AssertRC( femapMod, rc, "Unable to create a layer to store legend") Then Exit Sub
    femapMod.feAppMessage( FCM_NORMAL, "Create layer "+CStr(layerID)+".."+kLayerTitle )
  End If

'-----------------------------------------
  Set boxText = femapMod.feText
  boxText.layer =layerID
  boxText.color =opt_textColor
  boxText.FontNumber = opt_font
  boxText.DrawBorder = False

If opt_position_onLeft Then
  boxText.TextPosition(0) =opt_position_offsetH +2
  boxText.HorzJustify = 1
Else
  boxText.TextPosition(0) = 100-opt_position_offsetH -2
  boxText.HorzJustify =2
End If

  boxText.AllViews = opt_visibility_All
  boxText.VisibleView = IIf( opt_visibility_All , 0 , opt_visibility_viewID )
  boxText.ModelPosition = False


'-----------------------------------------
  Set boxColor = femapMod.feText

  boxColor.layer =layerID
  boxColor.text = "   "
'  boxColor.color =opt_textColor
'  boxColor.FontNumber = opt_font
  boxColor.DrawBorder = True

  If opt_position_onLeft Then
       boxColor.TextPosition(0) =opt_position_offsetH
	  'boxColor.HorzJustify = 1
  Else
      boxColor.TextPosition(0) = 100-opt_position_offsetH
     'boxColor.HorzJustify =2
  End If

  boxColor.AllViews = opt_visibility_All
  boxColor.VisibleView = IIf( opt_visibility_All , 0 , opt_visibility_viewID )
  boxColor.ModelPosition = False


If opt_position_Verti = positionV_top Then
   posVertic = opt_position_offsetV
   posStep = opt_position_spacing
ElseIf opt_position_Verti = positionV_bottom Then
   posVertic = 100- (opt_position_spacing * mySet.Count )
   posStep = opt_position_spacing
Else
  posStep = opt_position_spacing
  posVertic = (100- posStep*mySet.Count)/2
End If

'-----------------------------------------

  itemID =mySet.First

  Do Until itemID = 0
      info = ""

  	  If opt_groupByProp Then
           rc= aProp.Get( itemID )
               If AssertRC( femapMod, rc, "Unable to get data of property "+CStr(itemID) ) Then Exit Sub

           If opt_infoProp_ID Then info = info + IIf( Len(info)<>0 , " - propID: " , "propID: " ) + CStr( aProp.ID )
           If opt_infoProp_title Then info = info + IIf( Len(info)<>0 , " - " , "" ) + CStr( aProp.title )
           If opt_infoProp_type Then info = info + IIf( Len(info)<>0 , " - " , "" ) + propTypeAsString( aProp.type )
           If opt_infoProp_thickness Then
                 If is2DElementType( aProp.type) Then info = info + IIf( Len(info)<>0 , " - Th=" , "Th=" ) + CStr( aProp.pval(0) )
          End If

           itemID = aProp.matlID
  	  End If


      If itemID >0 Then
        rc= aMat.Get( itemID ) 
           If AssertRC( femapMod, rc, "Unable to get data of material "+CStr(itemID) ) Then Exit Sub

        If opt_infoMat_ID Then  info = info + IIf( Len(info)<>0 , " - matID: " , "matID: " ) + CStr( aMat.ID )
        If opt_infoMat_title Then  info = info + IIf( Len(info)<>0 , " - " , "" ) + CStr( aMat.title )
        If opt_infoMat_type Then info = info + IIf( Len(info)<>0 , " - " , "" ) + matTypeAsString( aMat.type )
     End If

  	  If opt_groupByProp Then aColor =aProp.color Else aColor=aMat.color

	  boxText.text = info
      boxColor.BackColor =aColor

      If opt_border Then
         boxColor.BorderColor =opt_textColor
      Else
         boxColor.BorderColor =aColor
      End If


	  boxText.TextPosition(1) =posVertic
	  boxColor.TextPosition(1) =posVertic
      posVertic = posVertic + posStep

	  rc= boxText.Put( boxText.NextEmptyID )
          If AssertRC( femapMod, rc, "Creation of text legend produce an errror" ) Then Exit Sub


	  rc= boxColor.Put( boxColor.NextEmptyID )
          If AssertRC( femapMod, rc, "Creation of text legend produce an errror" ) Then Exit Sub

      itemID = mySet.Next

  Loop

   Call femapMod.feViewRegenerate(0)
End Sub


'-------------------------------------------------------------------------


Function is2DElementType( elemType As Long ) As Boolean
  Select Case elemType
  Case FET_L_SHEAR, FET_P_SHEAR
  Case FET_L_MEMBRANE, FET_P_MEMBRANE
  Case FET_L_BENDING, FET_P_BENDING
  Case FET_L_PLATE, FET_P_PLATE
  Case FET_L_LAMINATE_PLATE, FET_P_LAMINATE_PLATE
  Case FET_L_PLANE_STRAIN, FET_P_PLANE_STRAIN
  Case FET_L_AXISYM_SHELL, FET_P_AXISYM_SHELL
  Case FET_L_PLOT_PLATE
  Case Else
    Exit Function
  End Select

  is2DElementType= True
End Function


'-------------------------------------------------------------------------


Function IndexOf( titles() As String, searchID As Long )As Long
   Dim i As Long, last As Long , ID As Long, anObject As femap.View
   Set anObject = femapMod.feView

   last = UBound( titles )
   For i=0 To last
   	    ID = anObject.ParseTitleID( titles(i) )
        If ID =searchID Then
         IndexOf = i
         Exit Function
     End If
   Next
   IndexOf = -1
End Function


'-------------------------------------------------------------------------


Sub selectEntitiesToShow()
   Dim rc As Long
   If mySet Is Nothing Then Set mySet = femapMod.feSet
   mySet.Clear

   rc= mySet.Select( IIf( opt_groupByProp , FT_PROP , FT_MATL ) , True , "Select "+IIf( opt_groupByProp , "properties" , "materials" )+" to use in legend" )
          If rc= FE_CANCEL Then Exit Sub
          If AssertRC( femapMod, rc, "Unable to select "+IIf( opt_groupByProp , "properties" , "materials" ) ) Then Exit Sub

   DlgValue( "CheckBox_selectAll" , False )
   mySet_isProp =opt_groupByProp

End Sub


'-------------------------------------------------------------------------


Function matTypeAsString( T As Long ) As String
   Select Case T
   Case FMT_ANISOTROPIC_2D
     matTypeAsString="Anisotropic 2d"
   Case FMT_ANISOTROPIC_3D
     matTypeAsString="Anisotropic 3d"
   Case FMT_FLUID
     matTypeAsString="Fluid"
   Case FMT_GENERAL
     matTypeAsString="General"
   Case FMT_HYPERELASTIC
     matTypeAsString="Hyperelastic"
   Case FMT_ISOTROPIC
     matTypeAsString="Isotropic"
   Case FMT_ORTHOTROPIC_2D
     matTypeAsString="Orthotropic 2d"
   Case FMT_ORTHOTROPIC_3D
     matTypeAsString="Orthotropic 3d"

	Case Else
        matTypeAsString = "*"
	End Select

End Function


'-------------------------------------------------------------------------


Function propTypeAsString( T As Long ) As String
  'zElementType
  Select Case T

'--- Line Elements ---
  Case FET_NONE
   propTypeAsString = "None"
  Case FET_L_ROD
   propTypeAsString = "Rod"
  Case FET_L_TUBE
   propTypeAsString = "Tube"
  Case FET_L_CURVED_TUBE
   propTypeAsString = "Curved Tube"
  Case FET_L_BAR
   propTypeAsString = "Bar"
  Case FET_L_BEAM, FET_P_BEAM
   propTypeAsString = "Beam"
  Case FET_L_LINK
   propTypeAsString = "Link"
  Case FET_L_CURVED_BEAM
   propTypeAsString = "Curved Beam"
  Case FET_L_SPRING
   propTypeAsString = "Spring/Damper"
  Case FET_L_DOF_SPRING
   propTypeAsString = "DOF Spring"
  Case FET_L_GAP
   propTypeAsString = "Gap"
  Case FET_L_PLOT
   propTypeAsString = "Plot Only"


'--- Plane Elements ---
  Case FET_L_SHEAR, FET_P_SHEAR
   propTypeAsString = "Shear"
  Case FET_L_MEMBRANE, FET_P_MEMBRANE
   propTypeAsString = "Membrane"
  Case FET_L_BENDING, FET_P_BENDING
   propTypeAsString = "Bending"
  Case FET_L_PLATE, FET_P_PLATE
   propTypeAsString = "Plate"
  Case FET_L_LAMINATE_PLATE, FET_P_LAMINATE_PLATE
   propTypeAsString = "Laminate"
  Case FET_L_PLANE_STRAIN, FET_P_PLANE_STRAIN
   propTypeAsString = "Palne strain"
  Case FET_L_AXISYM_SHELL, FET_P_AXISYM_SHELL
   propTypeAsString = "Axisymmetric Shell" 
  Case FET_L_PLOT_PLATE
   propTypeAsString = "Plot Only"

'--- Volume Elements ---
  Case FET_L_AXISYM, FET_P_AXISYM
   propTypeAsString = "Axisymmetric"
  Case FET_L_SOLID, FET_P_SOLID
   propTypeAsString = "Solid"

'--- Other Elements ---
  Case FET_L_MASS
   propTypeAsString = "Mass"
  Case FET_L_MASS_MATRIX
   propTypeAsString = "Mass Matrix"
  Case FET_L_RIGID
   propTypeAsString = "Rigid"
  Case FET_L_STIFF_MATRIX
   propTypeAsString = "Stiffness Matrix
  Case FET_L_SLIDE_LINE
   propTypeAsString = "Slide line"
  Case FET_L_WELD
   propTypeAsString = "Spring"
  Case FET_L_CONTACT
   propTypeAsString = "Contact"

  Case Else
   propTypeAsString = "*"
  End Select

End Function


'-------------------------------------------------------------------------


Sub dialog_removeAllLegend( verboseMsg As Boolean, Regenerate As Boolean)

   Dim layerID As Long
   layerID = getLayerByTitle( kLayerTitle )
   If layerID=-1 Then
     If verboseMsg Then femapMod.feAppMessage( FCM_NORMAL, "Layer "+kLayerTitle+" doesn't exist in model, no legend to delete")
     Exit Sub
   End If

   If femapMod.Info_Count( FT_TEXT ) = 0 Then
     If verboseMsg Then femapMod.feAppMessage( FCM_HIGHLIGHT , "No text in model" )
     Exit Sub
   End If

   Dim textSet As femap.Set , textItem As femap.text, rc As Long, msg As String
   Set textSet  =femapMod.feSet
   Set textItem = femapMod.feText

   rc = textItem.First
   While rc=FE_OK
   	      If textItem.layer = layerID Then textSet.Add( textItem.ID )
          rc = textItem.Next
   Wend

   If textSet.Count = 0 Then
     If verboseMsg Then femapMod.feAppMessage( FCM_HIGHLIGHT , "No text find in layer ID: "+CStr(layerID) )
     Exit Sub
   End If

   rc= femapMod.feDelete( FT_TEXT , textSet.ID )
       If AssertRC( femapMod,rc, "Unable to delete text legend, commmand feDelete produce an error" ) Then Exit Sub

   If verboseMsg Then femapMod.feAppMessage(FCM_NORMAL , "Legends removed" )
   If Regenerate Then Call femapMod.feViewRegenerate(0)
End Sub
'-------------------------------------------------------------------------


Sub dialog_color()
   Dim colorID As Long
   colorID =femapMod.Info_Color( FT_TEXT )
   Call parseTextAsULng( DlgText( "TextBox_color" ) , colorID, "" )

   Call femapMod.feAppColorPalette( colorID, colorID )
   DlgText( "TextBox_color" , CStr( colorID ) )
End Sub


'-------------------------------------------------------------------------


Function getFontList() As String()

	Dim out(76) As String , index As Long,iNames As Long, iSizes As Long
    Dim names As Variant , sizes As Variant
    names=Array( "MS Sans Serif" , "MS Serif","Courier", "Modern", "Script", "Arial", "Times New Roman", "Courier New", "MS UI Gothic","MS Mincho","MS Shell Dlg")
    sizes = Array( 8,10,12,14,18,24,32)

    For iNames = 0 To 10
      For iSizes = 0 To 6
         out(index)=CStr(index)+".."+CStr(sizes(iSizes))+"pt "+names(iNames)
         index = index+1
      Next
    Next

    getFontList = out
End Function


'-------------------------------------------------------------------------


Function getViewsList() As Variant
	Dim View As femap.View, count As Long, IDs As Variant, Titles As Variant, i As Long
	Set View = femapMod.feView

    Call View.GetTitleIDList( True , 0 , 0 , count, IDs, Titles )
     If count=0 Then Exit Function

     ReDim out( count-1 ) As String

     For i=0 To count-1
     	out(i)=Titles(i)
     Next

      getViewsList= out
End Function


'-------------------------------------------------------------------------


Function parseTextAsULng( sData As String , ByRef value As Long, Optional errorMsg As String ="" )As Boolean
' try to parse text as an unsigned long
   If Not IsNumeric( sData ) Then
   	 If errorMsg<>"" Then
        femapMod.feAppMessage( FCM_HIGHLIGHT , errorMsg )
        MsgBox( errorMsg , vbInformation )
     End If
     parseTextAsULng = False
   ElseIf CLng(sData)<0 Then
   	 If errorMsg<>"" Then
        femapMod.feAppMessage( FCM_HIGHLIGHT , errorMsg )
        MsgBox( errorMsg , vbInformation )
     End If
     parseTextAsULng = False
   Else
   	 value = CLng( sData )
     parseTextAsULng = True
   End If

End Function


'-------------------------------------------------------------------------


Function getLayerByTitle( title As String ) As Long
   getLayerByTitle= -1

   Dim rc As Long, IDs As Variant, titles As Variant, aLayer As femap.layer, layerCount As Long, iL As Long
   Set aLayer = femapMod.feLayer
   
   rc= aLayer.GetTitleList (0, 0, layerCount, IDs, titles)
       If AssertRC( femapMod, rc, "Unable to get list of layers" ) Then Exit Function

   For iL=0 To layerCount-1
     If titles( iL )=title Then
       getLayerByTitle= IDs(iL) : Exit Function
     End If
   Next

End Function


'-------------------------------------------------------------------------


Function options_LoadAllFromDialog(  ) As Boolean
   Dim v As Variant , findError As Boolean

   If Not parseTextAsULng( DlgText( "TextBox_color" ) , opt_textColor , "Error : verify color ID" ) Then findError=True

   opt_font = CLng( DlgValue( "DropListBox_font" ) )

   opt_border = DlgValue( "CheckBox_border" )
   opt_groupByProp = ( DlgValue("Group_byPropOrMat")=0 )
   opt_selectAll = ( DlgValue("CheckBox_selectAll") )
   opt_infoProp_ID = DlgValue("CheckBox_propInfo_ID")
   opt_infoProp_title = DlgValue("CheckBox_propInfo_title")
   opt_infoProp_type = DlgValue("CheckBox_propInfo_type")
   opt_infoProp_thickness = DlgValue("CheckBox_propInfo_th")
   opt_infoMat_ID = DlgValue("CheckBox_matInfo_ID")
   opt_infoMat_title = DlgValue("CheckBox_matInfo_title")
   opt_infoMat_type = DlgValue("CheckBox_matInfo_type")
   opt_position_onLeft = (DlgValue( "GroupPositionH" )=0)
   opt_position_Verti = DlgValue("GroupPositionV")
   opt_position_spacing = DlgText("TextBox_spacing")

   If Not parseTextAsULng( DlgText( "TextBox_offsetH" ) , opt_position_offsetH , "Error : verify horizontal offset" ) Then findError=True
   If Not parseTextAsULng( DlgText( "TextBox_offsetV" ) , opt_position_offsetV , "Error : verify vertical offset" ) Then findError=True

   opt_visibility_All = ( DlgValue("Group_visibilityView")=0)
   If opt_visibility_All Then
      opt_visibility_viewID = 0
   Else
       Call parseTextAsULng( DlgValue( "DropListBox_view" ) , opt_visibility_viewID , "Error : verify color ID" )
   End If

   options_LoadAllFromDialog =findError
End Function


'-------------------------------------------------------------------------


Sub options_iniDefaultValues

   opt_textColor = femapMod.Info_Color( FT_TEXT )
   opt_font = 0 'Todo get test font define in F6 parameters
   opt_border = False
   opt_groupByProp = True
   opt_selectAll = ( femapMod.Info_ActiveID( FT_PROP ) <= 10 )
   opt_infoProp_ID = True
   opt_infoProp_title = True
   opt_infoProp_type = False
   opt_infoProp_thickness = True
   opt_infoMat_ID = False
   opt_infoMat_title = False
   opt_infoMat_type = False
   opt_position_onLeft = True
   opt_position_Verti  =0
   opt_position_offsetV =3
   opt_position_offsetH = 3
   opt_position_spacing = 5
   opt_visibility_All = True
   opt_visibility_viewID =femapMod.Info_ActiveID( FT_VIEW )

End Sub


'-------------------------------------------------------------------------


Sub options_readUserdata

 Dim UData As femap.UserData
 Set UData = femapMod.feUserData

 If UData.GetTitle( "yz.PropMatLegend.v1" )<>FE_OK Then Exit Sub

 UData.ReadLong( opt_textColor )
 UData.ReadLong( opt_font )
 UData.ReadBool( opt_border )
 UData.ReadBool( opt_groupByProp )
 UData.ReadBool( opt_selectAll )
 UData.ReadBool( opt_infoProp_ID )
 UData.ReadBool( opt_infoProp_title )
 UData.ReadBool( opt_infoProp_type )
 UData.ReadBool( opt_infoProp_thickness )
 UData.ReadBool( opt_infoMat_ID )
 UData.ReadBool( opt_infoMat_title )
 UData.ReadBool( opt_infoMat_type )
 UData.ReadBool( opt_position_onLeft )
 UData.ReadLong( opt_position_Verti )
 UData.ReadLong( opt_position_offsetV )
 UData.ReadLong( opt_position_offsetH )
 UData.ReadLong( opt_position_spacing )
 UData.ReadBool( opt_visibility_All )
 UData.ReadLong( opt_visibility_viewID )

End Sub


'-------------------------------------------------------------------------


Sub options_SaveUserdata

 Dim UData As femap.UserData, rc As Long
 Set UData = femapMod.feUserData

 UData.WriteLong( opt_textColor )
 UData.WriteLong( opt_font )
 UData.WriteBool( opt_border )
 UData.WriteBool( opt_groupByProp )
 UData.WriteBool( opt_selectAll )
 UData.WriteBool( opt_infoProp_ID )
 UData.WriteBool( opt_infoProp_title )
 UData.WriteBool( opt_infoProp_type )
 UData.WriteBool( opt_infoProp_thickness )
 UData.WriteBool( opt_infoMat_ID )
 UData.WriteBool( opt_infoMat_title )
 UData.WriteBool( opt_infoMat_type )
 UData.WriteBool( opt_position_onLeft )
 UData.WriteLong( opt_position_Verti )
 UData.WriteLong( opt_position_offsetV )
 UData.WriteLong( opt_position_offsetH )
 UData.WriteLong( opt_position_spacing )
 UData.WriteBool( opt_visibility_All )
 UData.WriteLong( opt_visibility_viewID )

  rc= UData.PutTitle( "yz.PropMatLegend.v1" )
      call AssertRC( femapMod, rc, "Unable to save legend settings in model" )

End Sub


'-------------------------------------------------------------------------


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
