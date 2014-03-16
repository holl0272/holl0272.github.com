<%Option Explicit
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version:	2.00.001		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		August 18, 2003											*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clssfDesign
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private pstrdsgnALTBgnd1
Private pstrdsgnALTBgnd2
Private pstrdsgnALTColor1
Private pstrdsgnALTColor2
Private pstrdsgnALTFontColor1
Private pstrdsgnALTFontColor2
Private pstrdsgnALTFontFace1
Private pstrdsgnALTFontFace2
Private pstrdsgnALTFontSize1
Private pstrdsgnALTFontSize2
Private pstrdsgnBannerColor
Private pstrdsgnBannerImage
Private pstrdsgnBGColor0
Private pstrdsgnBGColor1
Private pstrdsgnBGColor2
Private pstrdsgnBGColor3
Private pstrdsgnBGColor4
Private pstrdsgnBGColor5
Private pstrdsgnBGColor7
Private pstrdsgnBgnd1
Private pstrdsgnBgnd2
Private pstrdsgnBgnd3
Private pstrdsgnBgnd4
Private pstrdsgnBgnd5
Private pstrdsgnBgnd7
Private pstrdsgnBTN01
Private pstrdsgnBTN02
Private pstrdsgnBTN03
Private pstrdsgnBTN04
Private pstrdsgnBTN05
Private pstrdsgnBTN06
Private pstrdsgnBTN07
Private pstrdsgnBTN08
Private pstrdsgnBTN09
Private pstrdsgnBTN10
Private pstrdsgnBTN11
Private pstrdsgnBTN12
Private pstrdsgnBTN13
Private pstrdsgnBTN14
Private pstrdsgnBTN15
Private pstrdsgnBTN16
Private pstrdsgnBTN17
Private pstrdsgnBTN18
Private pstrdsgnBTN19
Private pstrdsgnBTN20
Private pstrdsgnBTN21
Private pstrdsgnBTN22
Private pstrdsgnBTN23
Private pstrdsgnBTN24
Private pstrdsgnDescription
Private pstrdsgnFontColor1
Private pstrdsgnFontColor2
Private pstrdsgnFontColor3
Private pstrdsgnFontColor4
Private pstrdsgnFontColor5
Private pstrdsgnFontColor7
Private pstrdsgnFontFace1
Private pstrdsgnFontFace2
Private pstrdsgnFontFace3
Private pstrdsgnFontFace4
Private pstrdsgnFontFace5
Private pstrdsgnFontFace7
Private pstrdsgnFontSize1
Private pstrdsgnFontSize2
Private pstrdsgnFontSize3
Private pstrdsgnFontSize4
Private pstrdsgnFontSize5
Private pstrdsgnFontSize7
Private pstrdsgnForm
Private pstrdsgnGenALink
Private pstrdsgnGenBgColor
Private pstrdsgnGenBgnd
Private pstrdsgnGenLink
Private pstrdsgnGenVLink
Private plngdsgnID
Private plngdsgnIsActive
Private pstrdsgnName
Private pstrdsgnTableWidth
Private pstrdsgnThumbnail

Private pstrFormColor
Private pstrFormFontFace
Private pstrFormFontSize

Public Property Get FormColor()
    FormColor = pstrFormColor
End Property
Public Property Get FormFontFace()
    FormFontFace = pstrFormFontFace
End Property
Public Property Get FormFontSize()
    FormFontSize = pstrFormFontSize
End Property

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
End Sub

Private Sub class_Terminate()
    On Error Resume Next
    Set pRS = Nothing
End Sub

'***********************************************************************************************

Public Property Let Recordset(oRS)
    set pRS = oRS
End Property

Public Property Get Recordset()
    set Recordset = pRS
End Property


Public Property Get Message()
    Message = pstrMessage
End Property

Public Sub OutputMessage()

Dim i
Dim aError

    aError = Split(pstrMessage, cstrDelimeter)
    For i = 0 To UBound(aError)
        If pblnError Then
            Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"
        Else
            Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"
        End If
    Next 'i

End Sub 'OutputMessage


Public Property Get dsgnALTBgnd1()
    dsgnALTBgnd1 = pstrdsgnALTBgnd1
End Property

Public Property Get dsgnALTBgnd2()
    dsgnALTBgnd2 = pstrdsgnALTBgnd2
End Property

Public Property Get dsgnALTColor1()
    dsgnALTColor1 = pstrdsgnALTColor1
End Property

Public Property Get dsgnALTColor2()
    dsgnALTColor2 = pstrdsgnALTColor2
End Property

Public Property Get dsgnALTFontColor1()
    dsgnALTFontColor1 = pstrdsgnALTFontColor1
End Property

Public Property Get dsgnALTFontColor2()
    dsgnALTFontColor2 = pstrdsgnALTFontColor2
End Property

Public Property Get dsgnALTFontFace1()
    dsgnALTFontFace1 = pstrdsgnALTFontFace1
End Property

Public Property Get dsgnALTFontFace2()
    dsgnALTFontFace2 = pstrdsgnALTFontFace2
End Property

Public Property Get dsgnALTFontSize1()
    dsgnALTFontSize1 = pstrdsgnALTFontSize1
End Property

Public Property Get dsgnALTFontSize2()
    dsgnALTFontSize2 = pstrdsgnALTFontSize2
End Property

Public Property Get dsgnBannerColor()
    dsgnBannerColor = pstrdsgnBannerColor
End Property

Public Property Get dsgnBannerImage()
    dsgnBannerImage = pstrdsgnBannerImage
End Property

Public Property Get dsgnBGColor0()
    dsgnBGColor0 = pstrdsgnBGColor0
End Property

Public Property Get dsgnBGColor1()
    dsgnBGColor1 = pstrdsgnBGColor1
End Property

Public Property Get dsgnBGColor2()
    dsgnBGColor2 = pstrdsgnBGColor2
End Property

Public Property Get dsgnBGColor3()
    dsgnBGColor3 = pstrdsgnBGColor3
End Property

Public Property Get dsgnBGColor4()
    dsgnBGColor4 = pstrdsgnBGColor4
End Property

Public Property Get dsgnBGColor5()
    dsgnBGColor5 = pstrdsgnBGColor5
End Property

Public Property Get dsgnBGColor7()
    dsgnBGColor7 = pstrdsgnBGColor7
End Property

Public Property Get dsgnBgnd1()
    dsgnBgnd1 = pstrdsgnBgnd1
End Property

Public Property Get dsgnBgnd2()
    dsgnBgnd2 = pstrdsgnBgnd2
End Property

Public Property Get dsgnBgnd3()
    dsgnBgnd3 = pstrdsgnBgnd3
End Property

Public Property Get dsgnBgnd4()
    dsgnBgnd4 = pstrdsgnBgnd4
End Property

Public Property Get dsgnBgnd5()
    dsgnBgnd5 = pstrdsgnBgnd5
End Property

Public Property Get dsgnBgnd7()
    dsgnBgnd7 = pstrdsgnBgnd7
End Property

Public Property Get dsgnBTN01()
    dsgnBTN01 = pstrdsgnBTN01
End Property

Public Property Get dsgnBTN02()
    dsgnBTN02 = pstrdsgnBTN02
End Property

Public Property Get dsgnBTN03()
    dsgnBTN03 = pstrdsgnBTN03
End Property

Public Property Get dsgnBTN04()
    dsgnBTN04 = pstrdsgnBTN04
End Property

Public Property Get dsgnBTN05()
    dsgnBTN05 = pstrdsgnBTN05
End Property

Public Property Get dsgnBTN06()
    dsgnBTN06 = pstrdsgnBTN06
End Property

Public Property Get dsgnBTN07()
    dsgnBTN07 = pstrdsgnBTN07
End Property

Public Property Get dsgnBTN08()
    dsgnBTN08 = pstrdsgnBTN08
End Property

Public Property Get dsgnBTN09()
    dsgnBTN09 = pstrdsgnBTN09
End Property

Public Property Get dsgnBTN10()
    dsgnBTN10 = pstrdsgnBTN10
End Property

Public Property Get dsgnBTN11()
    dsgnBTN11 = pstrdsgnBTN11
End Property

Public Property Get dsgnBTN12()
    dsgnBTN12 = pstrdsgnBTN12
End Property

Public Property Get dsgnBTN13()
    dsgnBTN13 = pstrdsgnBTN13
End Property

Public Property Get dsgnBTN14()
    dsgnBTN14 = pstrdsgnBTN14
End Property

Public Property Get dsgnBTN15()
    dsgnBTN15 = pstrdsgnBTN15
End Property

Public Property Get dsgnBTN16()
    dsgnBTN16 = pstrdsgnBTN16
End Property

Public Property Get dsgnBTN17()
    dsgnBTN17 = pstrdsgnBTN17
End Property

Public Property Get dsgnBTN18()
    dsgnBTN18 = pstrdsgnBTN18
End Property

Public Property Get dsgnBTN19()
    dsgnBTN19 = pstrdsgnBTN19
End Property

Public Property Get dsgnBTN20()
    dsgnBTN20 = pstrdsgnBTN20
End Property

Public Property Get dsgnBTN21()
    dsgnBTN21 = pstrdsgnBTN21
End Property

Public Property Get dsgnBTN22()
    dsgnBTN22 = pstrdsgnBTN22
End Property

Public Property Get dsgnBTN23()
    dsgnBTN23 = pstrdsgnBTN23
End Property

Public Property Get dsgnBTN24()
    dsgnBTN24 = pstrdsgnBTN24
End Property

Public Property Get dsgnDescription()
    dsgnDescription = pstrdsgnDescription
End Property

Public Property Get dsgnFontColor1()
    dsgnFontColor1 = pstrdsgnFontColor1
End Property

Public Property Get dsgnFontColor2()
    dsgnFontColor2 = pstrdsgnFontColor2
End Property

Public Property Get dsgnFontColor3()
    dsgnFontColor3 = pstrdsgnFontColor3
End Property

Public Property Get dsgnFontColor4()
    dsgnFontColor4 = pstrdsgnFontColor4
End Property

Public Property Get dsgnFontColor5()
    dsgnFontColor5 = pstrdsgnFontColor5
End Property

Public Property Get dsgnFontColor7()
    dsgnFontColor7 = pstrdsgnFontColor7
End Property

Public Property Get dsgnFontFace1()
    dsgnFontFace1 = pstrdsgnFontFace1
End Property

Public Property Get dsgnFontFace2()
    dsgnFontFace2 = pstrdsgnFontFace2
End Property

Public Property Get dsgnFontFace3()
    dsgnFontFace3 = pstrdsgnFontFace3
End Property

Public Property Get dsgnFontFace4()
    dsgnFontFace4 = pstrdsgnFontFace4
End Property

Public Property Get dsgnFontFace5()
    dsgnFontFace5 = pstrdsgnFontFace5
End Property

Public Property Get dsgnFontFace7()
    dsgnFontFace7 = pstrdsgnFontFace7
End Property

Public Property Get dsgnFontSize1()
    dsgnFontSize1 = pstrdsgnFontSize1
End Property

Public Property Get dsgnFontSize2()
    dsgnFontSize2 = pstrdsgnFontSize2
End Property

Public Property Get dsgnFontSize3()
    dsgnFontSize3 = pstrdsgnFontSize3
End Property

Public Property Get dsgnFontSize4()
    dsgnFontSize4 = pstrdsgnFontSize4
End Property

Public Property Get dsgnFontSize5()
    dsgnFontSize5 = pstrdsgnFontSize5
End Property

Public Property Get dsgnFontSize7()
    dsgnFontSize7 = pstrdsgnFontSize7
End Property

Public Property Get dsgnForm()
    dsgnForm = pstrdsgnForm
End Property

Public Property Get dsgnGenALink()
    dsgnGenALink = pstrdsgnGenALink
End Property

Public Property Get dsgnGenBgColor()
    dsgnGenBgColor = pstrdsgnGenBgColor
End Property

Public Property Get dsgnGenBgnd()
    dsgnGenBgnd = pstrdsgnGenBgnd
End Property

Public Property Get dsgnGenLink()
    dsgnGenLink = pstrdsgnGenLink
End Property

Public Property Get dsgnGenVLink()
    dsgnGenVLink = pstrdsgnGenVLink
End Property

Public Property Get dsgnID()
    dsgnID = plngdsgnID
End Property

Public Property Get dsgnIsActive()
    dsgnIsActive = plngdsgnIsActive
End Property

Public Property Get dsgnName()
    dsgnName = pstrdsgnName
End Property

Public Property Get dsgnTableWidth()
    dsgnTableWidth = pstrdsgnTableWidth
End Property

Public Property Get dsgnThumbnail()
    dsgnThumbnail = pstrdsgnThumbnail
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    pstrdsgnALTBgnd1 = trim(rs("dsgnALTBgnd1"))
    pstrdsgnALTBgnd2 = trim(rs("dsgnALTBgnd2"))
    pstrdsgnALTColor1 = trim(rs("dsgnALTColor1"))
    pstrdsgnALTColor2 = trim(rs("dsgnALTColor2"))
    pstrdsgnALTFontColor1 = trim(rs("dsgnALTFontColor1"))
    pstrdsgnALTFontColor2 = trim(rs("dsgnALTFontColor2"))
    pstrdsgnALTFontFace1 = trim(rs("dsgnALTFontFace1"))
    pstrdsgnALTFontFace2 = trim(rs("dsgnALTFontFace2"))
    pstrdsgnALTFontSize1 = trim(rs("dsgnALTFontSize1"))
    pstrdsgnALTFontSize2 = trim(rs("dsgnALTFontSize2"))
    pstrdsgnBannerColor = trim(rs("dsgnBannerColor"))
    pstrdsgnBannerImage = trim(rs("dsgnBannerImage"))
    pstrdsgnBGColor0 = trim(rs("dsgnBGColor0"))
    pstrdsgnBGColor1 = trim(rs("dsgnBGColor1"))
    pstrdsgnBGColor2 = trim(rs("dsgnBGColor2"))
    pstrdsgnBGColor3 = trim(rs("dsgnBGColor3"))
    pstrdsgnBGColor4 = trim(rs("dsgnBGColor4"))
    pstrdsgnBGColor5 = trim(rs("dsgnBGColor5"))
    pstrdsgnBGColor7 = trim(rs("dsgnBGColor7"))
    pstrdsgnBgnd1 = trim(rs("dsgnBgnd1"))
    pstrdsgnBgnd2 = trim(rs("dsgnBgnd2"))
    pstrdsgnBgnd3 = trim(rs("dsgnBgnd3"))
    pstrdsgnBgnd4 = trim(rs("dsgnBgnd4"))
    pstrdsgnBgnd5 = trim(rs("dsgnBgnd5"))
    pstrdsgnBgnd7 = trim(rs("dsgnBgnd7"))
    pstrdsgnBTN01 = trim(rs("dsgnBTN01"))
    pstrdsgnBTN02 = trim(rs("dsgnBTN02"))
    pstrdsgnBTN03 = trim(rs("dsgnBTN03"))
    pstrdsgnBTN04 = trim(rs("dsgnBTN04"))
    pstrdsgnBTN05 = trim(rs("dsgnBTN05"))
    pstrdsgnBTN06 = trim(rs("dsgnBTN06"))
    pstrdsgnBTN07 = trim(rs("dsgnBTN07"))
    pstrdsgnBTN08 = trim(rs("dsgnBTN08"))
    pstrdsgnBTN09 = trim(rs("dsgnBTN09"))
    pstrdsgnBTN10 = trim(rs("dsgnBTN10"))
    pstrdsgnBTN11 = trim(rs("dsgnBTN11"))
    pstrdsgnBTN12 = trim(rs("dsgnBTN12"))
    pstrdsgnBTN13 = trim(rs("dsgnBTN13"))
    pstrdsgnBTN14 = trim(rs("dsgnBTN14"))
    pstrdsgnBTN15 = trim(rs("dsgnBTN15"))
    pstrdsgnBTN16 = trim(rs("dsgnBTN16"))
    pstrdsgnBTN17 = trim(rs("dsgnBTN17"))
    pstrdsgnBTN18 = trim(rs("dsgnBTN18"))
    pstrdsgnBTN19 = trim(rs("dsgnBTN19"))
    pstrdsgnBTN20 = trim(rs("dsgnBTN20"))
    pstrdsgnBTN21 = trim(rs("dsgnBTN21"))
    pstrdsgnBTN22 = trim(rs("dsgnBTN22"))
    pstrdsgnBTN23 = trim(rs("dsgnBTN23"))
    pstrdsgnBTN24 = trim(rs("dsgnBTN24"))
    pstrdsgnDescription = trim(rs("dsgnDescription"))
    pstrdsgnFontColor1 = trim(rs("dsgnFontColor1"))
    pstrdsgnFontColor2 = trim(rs("dsgnFontColor2"))
    pstrdsgnFontColor3 = trim(rs("dsgnFontColor3"))
    pstrdsgnFontColor4 = trim(rs("dsgnFontColor4"))
    pstrdsgnFontColor5 = trim(rs("dsgnFontColor5"))
    pstrdsgnFontColor7 = trim(rs("dsgnFontColor7"))
    pstrdsgnFontFace1 = trim(rs("dsgnFontFace1"))
    pstrdsgnFontFace2 = trim(rs("dsgnFontFace2"))
    pstrdsgnFontFace3 = trim(rs("dsgnFontFace3"))
    pstrdsgnFontFace4 = trim(rs("dsgnFontFace4"))
    pstrdsgnFontFace5 = trim(rs("dsgnFontFace5"))
    pstrdsgnFontFace7 = trim(rs("dsgnFontFace7"))
    pstrdsgnFontSize1 = trim(rs("dsgnFontSize1"))
    pstrdsgnFontSize2 = trim(rs("dsgnFontSize2"))
    pstrdsgnFontSize3 = trim(rs("dsgnFontSize3"))
    pstrdsgnFontSize4 = trim(rs("dsgnFontSize4"))
    pstrdsgnFontSize5 = trim(rs("dsgnFontSize5"))
    pstrdsgnFontSize7 = trim(rs("dsgnFontSize7"))
    pstrdsgnForm = trim(rs("dsgnForm"))
    pstrdsgnGenALink = trim(rs("dsgnGenALink"))
    pstrdsgnGenBgColor = trim(rs("dsgnGenBgColor"))
    pstrdsgnGenBgnd = trim(rs("dsgnGenBgnd"))
    pstrdsgnGenLink = trim(rs("dsgnGenLink"))
    pstrdsgnGenVLink = trim(rs("dsgnGenVLink"))
    plngdsgnID = trim(rs("dsgnID"))
    plngdsgnIsActive = trim(rs("dsgnIsActive"))
    pstrdsgnName = trim(rs("dsgnName"))
    pstrdsgnTableWidth = trim(rs("dsgnTableWidth"))
    pstrdsgnThumbnail = trim(rs("dsgnThumbnail"))

Dim plngPos1
Dim plngPos2
Dim pstrTemp

'Split Form Element


	plngPos1 = instr(1,pstrdsgnForm,"BACKGROUND-COLOR:")
	If plngPos1 > 0 Then
		plngPos2 = instr(plngPos1,pstrdsgnForm,";")
		If plngPos2 > 0 Then pstrFormColor = trim(mid(pstrdsgnForm,plngPos1+17,plngPos2-plngPos1-17))
	End If
		
	plngPos1 = instr(1,pstrdsgnForm,"FONT-FAMILY:")
	If plngPos1 > 0 Then
		plngPos2 = instr(plngPos1,pstrdsgnForm,";")
		If plngPos2 > 0 Then pstrFormFontFace = trim(mid(pstrdsgnForm,plngPos1+12,plngPos2-plngPos1-12))
	End If
		
	plngPos1 = instr(1,pstrdsgnForm,"FONT-SIZE:")
	If plngPos1 > 0 Then
		plngPos2 = instr(plngPos1,pstrdsgnForm,"pt;")
		If plngPos2 > 0 Then pstrFormFontSize = trim(mid(pstrdsgnForm,plngPos1+10,plngPos2-plngPos1-10))
	End If
		
'	Const C_FORMDESIGN     = "BACKGROUND-COLOR: #ffffcc; FONT-FAMILY: Verdana; FONT-SIZE: 8pt"
	
'debugprint "pstrFormColor", pstrFormColor
'debugprint "pstrFormFontFace", pstrFormFontFace
'debugprint "pstrFormFontSize", pstrFormFontSize

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
'        pstrdsgnForm = Trim(.Item("dsgnForm"))
		pstrFormColor = Trim(.Item("dsgnALTBgnd1"))
		pstrFormFontFace = Trim(.Item("dsgnALTBgnd1"))
		pstrFormFontSize = Trim(.Item("dsgnALTBgnd1"))

		pstrdsgnForm = "BACKGROUND-COLOR: " & pstrFormColor _
					 & "; FONT-FAMILY: " & pstrFormFontFace _
					 & "; FONT-SIZE: " & pstrFormFontSize & "pt;"

        pstrdsgnALTBgnd1 = Trim(.Item("dsgnALTBgnd1"))
        pstrdsgnALTBgnd2 = Trim(.Item("dsgnALTBgnd2"))
        pstrdsgnALTColor1 = Trim(.Item("dsgnALTColor1"))
        pstrdsgnALTColor2 = Trim(.Item("dsgnALTColor2"))
        pstrdsgnALTFontColor1 = Trim(.Item("dsgnALTFontColor1"))
        pstrdsgnALTFontColor2 = Trim(.Item("dsgnALTFontColor2"))
        pstrdsgnALTFontFace1 = Trim(.Item("dsgnALTFontFace1"))
        pstrdsgnALTFontFace2 = Trim(.Item("dsgnALTFontFace2"))
        pstrdsgnALTFontSize1 = Trim(.Item("dsgnALTFontSize1"))
        pstrdsgnALTFontSize2 = Trim(.Item("dsgnALTFontSize2"))
        pstrdsgnBannerColor = Trim(.Item("dsgnBannerColor"))
        pstrdsgnBannerImage = Trim(.Item("dsgnBannerImage"))
        pstrdsgnBGColor0 = Trim(.Item("dsgnBGColor0"))
        pstrdsgnBGColor1 = Trim(.Item("dsgnBGColor1"))
        pstrdsgnBGColor2 = Trim(.Item("dsgnBGColor2"))
        pstrdsgnBGColor3 = Trim(.Item("dsgnBGColor3"))
        pstrdsgnBGColor4 = Trim(.Item("dsgnBGColor4"))
        pstrdsgnBGColor5 = Trim(.Item("dsgnBGColor5"))
        pstrdsgnBGColor7 = Trim(.Item("dsgnBGColor7"))
        pstrdsgnBgnd1 = Trim(.Item("dsgnBgnd1"))
        pstrdsgnBgnd2 = Trim(.Item("dsgnBgnd2"))
        pstrdsgnBgnd3 = Trim(.Item("dsgnBgnd3"))
        pstrdsgnBgnd4 = Trim(.Item("dsgnBgnd4"))
        pstrdsgnBgnd5 = Trim(.Item("dsgnBgnd5"))
        pstrdsgnBgnd7 = Trim(.Item("dsgnBgnd7"))
        pstrdsgnBTN01 = Trim(.Item("dsgnBTN01"))
        pstrdsgnBTN02 = Trim(.Item("dsgnBTN02"))
        pstrdsgnBTN03 = Trim(.Item("dsgnBTN03"))
        pstrdsgnBTN04 = Trim(.Item("dsgnBTN04"))
        pstrdsgnBTN05 = Trim(.Item("dsgnBTN05"))
        pstrdsgnBTN06 = Trim(.Item("dsgnBTN06"))
        pstrdsgnBTN07 = Trim(.Item("dsgnBTN07"))
        pstrdsgnBTN08 = Trim(.Item("dsgnBTN08"))
        pstrdsgnBTN09 = Trim(.Item("dsgnBTN09"))
        pstrdsgnBTN10 = Trim(.Item("dsgnBTN10"))
        pstrdsgnBTN11 = Trim(.Item("dsgnBTN11"))
        pstrdsgnBTN12 = Trim(.Item("dsgnBTN12"))
        pstrdsgnBTN13 = Trim(.Item("dsgnBTN13"))
        pstrdsgnBTN14 = Trim(.Item("dsgnBTN14"))
        pstrdsgnBTN15 = Trim(.Item("dsgnBTN15"))
        pstrdsgnBTN16 = Trim(.Item("dsgnBTN16"))
        pstrdsgnBTN17 = Trim(.Item("dsgnBTN17"))
        pstrdsgnBTN18 = Trim(.Item("dsgnBTN18"))
        pstrdsgnBTN19 = Trim(.Item("dsgnBTN19"))
        pstrdsgnBTN20 = Trim(.Item("dsgnBTN20"))
        pstrdsgnBTN21 = Trim(.Item("dsgnBTN21"))
        pstrdsgnBTN22 = Trim(.Item("dsgnBTN22"))
        pstrdsgnBTN23 = Trim(.Item("dsgnBTN23"))
        pstrdsgnBTN24 = Trim(.Item("dsgnBTN24"))
        pstrdsgnDescription = Trim(.Item("dsgnDescription"))
        pstrdsgnFontColor1 = Trim(.Item("dsgnFontColor1"))
        pstrdsgnFontColor2 = Trim(.Item("dsgnFontColor2"))
        pstrdsgnFontColor3 = Trim(.Item("dsgnFontColor3"))
        pstrdsgnFontColor4 = Trim(.Item("dsgnFontColor4"))
        pstrdsgnFontColor5 = Trim(.Item("dsgnFontColor5"))
        pstrdsgnFontColor7 = Trim(.Item("dsgnFontColor7"))
        pstrdsgnFontFace1 = Trim(.Item("dsgnFontFace1"))
        pstrdsgnFontFace2 = Trim(.Item("dsgnFontFace2"))
        pstrdsgnFontFace3 = Trim(.Item("dsgnFontFace3"))
        pstrdsgnFontFace4 = Trim(.Item("dsgnFontFace4"))
        pstrdsgnFontFace5 = Trim(.Item("dsgnFontFace5"))
        pstrdsgnFontFace7 = Trim(.Item("dsgnFontFace7"))
        pstrdsgnFontSize1 = Trim(.Item("dsgnFontSize1"))
        pstrdsgnFontSize2 = Trim(.Item("dsgnFontSize2"))
        pstrdsgnFontSize3 = Trim(.Item("dsgnFontSize3"))
        pstrdsgnFontSize4 = Trim(.Item("dsgnFontSize4"))
        pstrdsgnFontSize5 = Trim(.Item("dsgnFontSize5"))
        pstrdsgnFontSize7 = Trim(.Item("dsgnFontSize7"))
        pstrdsgnGenALink = Trim(.Item("dsgnGenALink"))
        pstrdsgnGenBgColor = Trim(.Item("dsgnGenBgColor"))
        pstrdsgnGenBgnd = Trim(.Item("dsgnGenBgnd"))
        pstrdsgnGenLink = Trim(.Item("dsgnGenLink"))
        pstrdsgnGenVLink = Trim(.Item("dsgnGenVLink"))
        plngdsgnID = Trim(.Item("dsgnID"))
        plngdsgnIsActive = Trim(.Item("dsgnIsActive"))
        pstrdsgnName = Trim(.Item("dsgnName"))
        pstrdsgnTableWidth = Trim(.Item("dsgnTableWidth"))
        pstrdsgnThumbnail = Trim(.Item("dsgnThumbnail"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

	Find = False
    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "dsgnID=" & lngID
            Else
                .MoveLast
            End If
            If Not .EOF Then 
				LoadValues (pRS)
				Find = True
			End If
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function LoadAll()

'On Error Resume Next

    Set pRS = GetRS("Select * from sfDesign Order By dsgnIsActive Desc, dsgnName")
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(lngdsgnID)

Dim sql

'On Error Resume Next

    sql = "Delete from sfDesign where dsgnID = " & lngdsgnID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Update()

Dim sql
Dim rs
Dim strErrorMessage
Dim blnAdd

'On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
        If Len(plngdsgnID) = 0 Then plngdsgnID = 0

        sql = "Select * from sfDesign where dsgnID = " & plngdsgnID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("dsgnALTBgnd1") = pstrdsgnALTBgnd1
        rs("dsgnALTBgnd2") = pstrdsgnALTBgnd2
        rs("dsgnALTColor1") = pstrdsgnALTColor1
        rs("dsgnALTColor2") = pstrdsgnALTColor2
        rs("dsgnALTFontColor1") = pstrdsgnALTFontColor1
        rs("dsgnALTFontColor2") = pstrdsgnALTFontColor2
        rs("dsgnALTFontFace1") = pstrdsgnALTFontFace1
        rs("dsgnALTFontFace2") = pstrdsgnALTFontFace2
        rs("dsgnALTFontSize1") = pstrdsgnALTFontSize1
        rs("dsgnALTFontSize2") = pstrdsgnALTFontSize2
        rs("dsgnBannerColor") = pstrdsgnBannerColor
        rs("dsgnBannerImage") = pstrdsgnBannerImage
        rs("dsgnBGColor0") = pstrdsgnBGColor0
        rs("dsgnBGColor1") = pstrdsgnBGColor1
        rs("dsgnBGColor2") = pstrdsgnBGColor2
        rs("dsgnBGColor3") = pstrdsgnBGColor3
        rs("dsgnBGColor4") = pstrdsgnBGColor4
        rs("dsgnBGColor5") = pstrdsgnBGColor5
        rs("dsgnBGColor7") = pstrdsgnBGColor7
        rs("dsgnBgnd1") = pstrdsgnBgnd1
        rs("dsgnBgnd2") = pstrdsgnBgnd2
        rs("dsgnBgnd3") = pstrdsgnBgnd3
        rs("dsgnBgnd4") = pstrdsgnBgnd4
        rs("dsgnBgnd5") = pstrdsgnBgnd5
        rs("dsgnBgnd7") = pstrdsgnBgnd7
        rs("dsgnBTN01") = pstrdsgnBTN01
        rs("dsgnBTN02") = pstrdsgnBTN02
        rs("dsgnBTN03") = pstrdsgnBTN03
        rs("dsgnBTN04") = pstrdsgnBTN04
        rs("dsgnBTN05") = pstrdsgnBTN05
        rs("dsgnBTN06") = pstrdsgnBTN06
        rs("dsgnBTN07") = pstrdsgnBTN07
        rs("dsgnBTN08") = pstrdsgnBTN08
        rs("dsgnBTN09") = pstrdsgnBTN09
        rs("dsgnBTN10") = pstrdsgnBTN10
        rs("dsgnBTN11") = pstrdsgnBTN11
        rs("dsgnBTN12") = pstrdsgnBTN12
        rs("dsgnBTN13") = pstrdsgnBTN13
        rs("dsgnBTN14") = pstrdsgnBTN14
        rs("dsgnBTN15") = pstrdsgnBTN15
        rs("dsgnBTN16") = pstrdsgnBTN16
        rs("dsgnBTN17") = pstrdsgnBTN17
        rs("dsgnBTN18") = pstrdsgnBTN18
        rs("dsgnBTN19") = pstrdsgnBTN19
        rs("dsgnBTN20") = pstrdsgnBTN20
        rs("dsgnBTN21") = pstrdsgnBTN21
        rs("dsgnBTN22") = pstrdsgnBTN22
        rs("dsgnBTN23") = pstrdsgnBTN23
        rs("dsgnBTN24") = pstrdsgnBTN24
        rs("dsgnDescription") = pstrdsgnDescription
        rs("dsgnFontColor1") = pstrdsgnFontColor1
        rs("dsgnFontColor2") = pstrdsgnFontColor2
        rs("dsgnFontColor3") = pstrdsgnFontColor3
        rs("dsgnFontColor4") = pstrdsgnFontColor4
        rs("dsgnFontColor5") = pstrdsgnFontColor5
        rs("dsgnFontColor7") = pstrdsgnFontColor7
        rs("dsgnFontFace1") = pstrdsgnFontFace1
        rs("dsgnFontFace2") = pstrdsgnFontFace2
        rs("dsgnFontFace3") = pstrdsgnFontFace3
        rs("dsgnFontFace4") = pstrdsgnFontFace4
        rs("dsgnFontFace5") = pstrdsgnFontFace5
        rs("dsgnFontFace7") = pstrdsgnFontFace7
        rs("dsgnFontSize1") = pstrdsgnFontSize1
        rs("dsgnFontSize2") = pstrdsgnFontSize2
        rs("dsgnFontSize3") = pstrdsgnFontSize3
        rs("dsgnFontSize4") = pstrdsgnFontSize4
        rs("dsgnFontSize5") = pstrdsgnFontSize5
        rs("dsgnFontSize7") = pstrdsgnFontSize7
        rs("dsgnForm") = pstrdsgnForm
        rs("dsgnGenALink") = pstrdsgnGenALink
        rs("dsgnGenBgColor") = pstrdsgnGenBgColor
        rs("dsgnGenBgnd") = pstrdsgnGenBgnd
        rs("dsgnGenLink") = pstrdsgnGenLink
        rs("dsgnGenVLink") = pstrdsgnGenVLink
        rs("dsgnIsActive") = ((plngdsgnIsActive = "on") * -1)
        rs("dsgnName") = pstrdsgnName
        rs("dsgnTableWidth") = pstrdsgnTableWidth
        rs("dsgnThumbnail") = pstrdsgnThumbnail

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngdsgnID = rs("dsgnID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The record was successfully added."
            Else
                pstrMessage = "The record was successfully updated."
            End If
        Else
            pblnError = True
        End If
    Else
        pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************

Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim pstrTitle, pstrURL, pstrAbbr

	With Response
	
    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<tr><td>"
    .Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
    .Write "<table class='tbl' id='tblSummary' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "  <tr class='tblhdr'>"

    .Write "<TH align='center'>Name</TH>"
    .Write "<TH align='center'>Description</TH>"
    .Write "<TH align='center'>Active</TH></TR>"
    If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrAbbr = Trim(prs("dsgnID"))
 			pstrTitle = "Click to edit " & prs("dsgnName") & "."
			pstrURL = "sfDesignAdmin.asp?Action=View&dsgnID=" & pstrAbbr
 
 			if pstrAbbr = plngdsgnID then
        		.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
				.Write "<TD>" & prs("dsgnName") & "</TD>" & vbCrLf
			else
				if cBool(pRS("dsgnIsActive")) then
					.Write "<TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				else
					.Write "<TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
        		end if
				.Write "<TD><a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prs("dsgnName") & "</a></TD>" & vbCrLf
        	end if

            Response.Write "<TD>" & pstrdsgnDescription & "&nbsp;</TD>" & vbCrLf
			if cBool(pRS("dsgnIsActive")) then
				.Write "<TD><b>Active</b></TD></TR>" & vbCrLf
			else
				.Write "<TD><a href='sfDesignAdmin.asp?Action=Activate&dsgnID=" & prs("dsgnID") & "'>Inactive</a></TD></TR>" & vbCrLf
       		end if
            prs.MoveNext
        Next
    Else
        Response.Write "<TR><TD><h3>There are no Designs</h3></TD></TR>"
    End If
    .Write "</TABLE>"
    .Write "</DIV>"
    .Write "</TD></TR></TABLE>"
    End With

End Sub      'OutputSummary

'***********************************************************************************************

Public Sub WriteDesignToFile()

'On Error Resume Next

Dim fso, MyFile
Dim p_strFile, p_strSSLFile

	p_strFile = mstrBasePath & "SFLib/incDesign.asp"
	p_strSSLFile = mstrBasePath & "SSL/SFLib/incDesign.asp"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(p_strFile, True)

With MyFile
	.WriteLine "<" & "%"
	.WriteLine "'******************************************************************"
	.WriteLine "' "
	.WriteLine "' Created With Sandshot Software's WebStore Manager for Storefront "
	.WriteLine "' "
	.WriteLine "' Design constants used for StoreFront "
	.WriteLine "' "
	.WriteLine "'*******************************************************************"
	.WriteLine ""
	.WriteLine "	'---- Content ----"
	.WriteLine "	Const C_FONTCOLOR4     = " & chr(34) & pstrdsgnFontColor4 & chr(34)
	.WriteLine "	Const C_FONTSIZE4      = " & chr(34) & pstrdsgnFontSize4 & chr(34)
	.WriteLine "	Const C_FONTFACE4      = " & chr(34) & pstrdsgnFontFace4 & chr(34)
	.WriteLine ""
	.WriteLine "	'---- Button Design Images ----"
	.WriteLine "	Const C_BTN01          = " & chr(34) & pstrdsgnBTN01 & chr(34) & " '--- Go"
	.WriteLine "	Const C_BTN02          = " & chr(34) & pstrdsgnBTN02 & chr(34) & " '--- Save to Cart"
	.WriteLine "	Const C_BTN03          = " & chr(34) & pstrdsgnBTN03 & chr(34) & " '--- Add to Cart"
	.WriteLine "	Const C_BTN04          = " & chr(34) & pstrdsgnBTN04 & chr(34) & " '--- Continue Search"
	.WriteLine "	Const C_BTN05          = " & chr(34) & pstrdsgnBTN05 & chr(34) & " '--- Checkout"
	.WriteLine "	Const C_BTN06          = " & chr(34) & pstrdsgnBTN06 & chr(34) & " '--- Delete"
	.WriteLine "	Const C_BTN07          = " & chr(34) & pstrdsgnBTN07 & chr(34) & " '--- Save"
	.WriteLine "	Const C_BTN08          = " & chr(34) & pstrdsgnBTN08 & chr(34) & " '--- View Saved Cart"
	.WriteLine "	Const C_BTN09          = " & chr(34) & pstrdsgnBTN09 & chr(34) & " '--- Return to Shop"
	.WriteLine "	Const C_BTN10          = " & chr(34) & pstrdsgnBTN10 & chr(34) & " '--- Shopping Cart (Order)"
	.WriteLine "	Const C_BTN11          = " & chr(34) & pstrdsgnBTN11 & chr(34) & " '--- Change Cart"
	.WriteLine "	Const C_BTN12          = " & chr(34) & pstrdsgnBTN12 & chr(34) & " '--- Sign Up"
	.WriteLine "	Const C_BTN13          = " & chr(34) & pstrdsgnBTN13 & chr(34) & " '--- Shopping Cart"
	.WriteLine "	Const C_BTN14          = " & chr(34) & pstrdsgnBTN14 & chr(34) & " '--- Recalculate"
	.WriteLine "	Const C_BTN15          = " & chr(34) & pstrdsgnBTN15 & chr(34) & " '--- Help"
	.WriteLine "	Const C_BTN16          = " & chr(34) & pstrdsgnBTN16 & chr(34) & " '--- Login"
	.WriteLine "	Const C_BTN17          = " & chr(34) & pstrdsgnBTN17 & chr(34) & " '--- Forgot Password"
	.WriteLine "	Const C_BTN18          = " & chr(34) & pstrdsgnBTN18 & chr(34) & " '--- Submit"
	.WriteLine "	Const C_BTN19          = " & chr(34) & pstrdsgnBTN19 & chr(34) & " '--- New Account"
	.WriteLine "	Const C_BTN20          = " & chr(34) & pstrdsgnBTN20 & chr(34) & " '--- Verify"
	.WriteLine "	Const C_BTN21          = " & chr(34) & pstrdsgnBTN21 & chr(34) & " '--- Search"
	.WriteLine "	Const C_BTN22          = " & chr(34) & pstrdsgnBTN22 & chr(34) & " '--- Add to Cart (Small)"
	.WriteLine "	Const C_BTN23          = " & chr(34) & pstrdsgnBTN23 & chr(34) & " '--- Clear Shipping Info"
	.WriteLine "	Const C_BTN24          = " & chr(34) & pstrdsgnBTN24 & chr(34) & " '--- Email A Friend"
	.WriteLine "%" & ">"

	.Close
End With

	fso.CopyFile p_strFile,p_strSSLFile

	Set fso = Nothing
	Set MyFile = Nothing

	'Call WriteDesignToCSSFile
	
End Sub      'WriteDesignToFile

'***********************************************************************************************

Public Sub WriteDesignToCSSFile()

'On Error Resume Next

Dim fso, MyFile
Dim p_strFile, p_strSSLFile

	p_strFile = mstrBasePath & "sfcss.css"
	p_strSSLFile = mstrBasePath & "SSL/sfcss.css"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(p_strFile, True)

With MyFile

	.WriteLine ".Footer {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize7) & "pt " & pstrdsgnFontFace7 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor7 & ";"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".AltFont1 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnALTFontSize1) & "pt " & pstrdsgnALTFontFace1 & ";"
	.WriteLine "  color: " & pstrdsgnALTFontColor1 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".AltFont2 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnALTFontSize2) & "pt " & pstrdsgnALTFontFace2 & ";"
	.WriteLine "  color: " & pstrdsgnALTFontColor2 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".Content_Small {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4 - 1) & "pt " & pstrdsgnFontFace4 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".Content_Large {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4 + 2) & "pt " & pstrdsgnFontFace4 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".ECheck {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4-1) & "pt " & pstrdsgnFontFace4 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".ECheck2 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4-2) & "pt " & pstrdsgnFontFace4 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".Error {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize5) & "pt " & pstrdsgnFontFace5 & ";"
	.WriteLine "  color: #FF0000;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".Middle_Top_Banner_Small {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize2) & "pt " & pstrdsgnFontFace2 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor2 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".ContentBar_Small {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize5) & "pt " & pstrdsgnFontFace5 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor5 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".TopBanner_Large {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnALTFontSize1) & "pt " & pstrdsgnALTFontFace1 & ";"
	.WriteLine "  color: " & pstrdsgnALTFontColor1 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine "body {"
	.WriteLine "  background-color: " & pstrdsgnGenBgColor & ";"
	If Len(pstrdsgnGenBgnd) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnGenBgnd & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine "body, table, td, p {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4) & "pt " & pstrdsgnFontFace4 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine "h1 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize4) & "pt " & pstrdsgnFontFace2 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor4 & ";"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdbackgrnd {"
	.WriteLine "  background-color: " & pstrdsgnBGColor0 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdTopBanner {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize1) & "pt " & pstrdsgnFontFace1 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor1 & ";"
	.WriteLine "  background-color: " & pstrdsgnBgnd1 & ";"
	If Len(pstrdsgnBgnd1) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd1 & Chr(34) & ");"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdTopBanner2 {"
	.WriteLine "  background-color: " & pstrdsgnBannerColor & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdMiddleTopBanner {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize2) & "pt " & pstrdsgnFontFace2 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor2 & ";"
	.WriteLine "  background-color: " & pstrdsgnBgnd2 & ";"
	If Len(pstrdsgnBgnd2) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd2 & Chr(34) & ");"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdBottomTopBanner {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize3) & "pt " & pstrdsgnFontFace3 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor3 & ";"
	.WriteLine "  background-color: " & pstrdsgnBgnd3 & ";"
	If Len(pstrdsgnBgnd3) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd3 & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdBottomTopBanner2 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize3) & "pt " & pstrdsgnFontFace3 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor3 & ";"
	.WriteLine "  background-color: " & pstrdsgnBgnd3 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdContent {"
	.WriteLine "  background-color: " & pstrdsgnBgnd4 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdContent2 {"
	.WriteLine "  background-color: " & pstrdsgnBgnd4 & ";"
	If Len(pstrdsgnBgnd4) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd4 & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdContent3 {"
	.WriteLine "  background-color: " & pstrdsgnBgnd4 & ";"
	If Len(pstrdsgnBgnd5) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd5 & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdContentBar {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize5) & "pt " & pstrdsgnFontFace5 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor5 & ";"
	.WriteLine "  background-color: " & pstrdsgnBgnd5 & ";"
	If Len(pstrdsgnBgnd5) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnBgnd5 & Chr(34) & ");"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdFooter {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnFontSize7) & "pt " & pstrdsgnFontFace7 & ";"
	.WriteLine "  color: " & pstrdsgnFontColor7 & ";"
	.WriteLine "  font-weight: bold;"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdAltFont1 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnALTFontSize1) & "pt " & pstrdsgnALTFontFace1 & ";"
	.WriteLine "  color: " & pstrdsgnALTFontColor1 & ";"
	.WriteLine "  background-color: " & pstrdsgnALTColor1 & ";"
	If Len(pstrdsgnALTBgnd1) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnALTBgnd1 & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdAltFont2 {"
	.WriteLine "  font: " & FontSizeValue(pstrdsgnALTFontSize2) & "pt " & pstrdsgnALTFontFace2 & ";"
	.WriteLine "  color: " & pstrdsgnALTFontColor2 & ";"
	.WriteLine "  background-color: " & pstrdsgnALTColor2 & ";"
	If Len(pstrdsgnALTBgnd2) > 0 Then .WriteLine "  background: url(" & Chr(34) & pstrdsgnALTBgnd2 & Chr(34) & ");"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdAltBG1 {"
	.WriteLine "  background-color: " & pstrdsgnALTColor1 & ";"
	.WriteLine "}"
	.WriteLine ""
	
	.WriteLine ".tdAltBG2 {"
	.WriteLine "  background-color: " & pstrdsgnALTColor2 & ";"
	.WriteLine "}"

	.Close
End With

	fso.CopyFile p_strFile,p_strSSLFile

	Set fso = Nothing
	Set MyFile = Nothing

End Sub      'WriteDesignToCSSFile

'***********************************************************************************************

Function FontSizeValue(bytValue)

	If bytValue < 0 Then
		FontSizeValue = "7.5"
	Else
		Select Case CStr(bytValue)
			Case "0": FontSizeValue = "7.5"
			Case "1": FontSizeValue = "7.5"
			Case "2": FontSizeValue = "10"
			Case "3": FontSizeValue = "12"
			Case "4": FontSizeValue = "13.5"
			Case "5": FontSizeValue = "18"
			Case "6": FontSizeValue = "24"
			Case Else: FontSizeValue = "37"
		End Select
	End If
	
End Function

'***********************************************************************************************

Function Activate(dsgnID)

	cnn.Execute "Update sfDesign Set dsgnIsActive='0'",,128
	cnn.Execute "Update sfDesign Set dsgnIsActive='1' where dsgnID=" & dsgnID,,128

End Function

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrdsgnName) = 0 Then
        strError = strError & "Please enter a design name." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues
End Class   'clssfDesign

Function SetImagePath(strImage)

	If len(trim(strImage)) > 0 Then
		SetImagePath = mstrBaseHRef & strImage
	Else
		SetImagePath = "images/NoImage.gif"
	End If

End Function

mstrPageTitle = "StoreFront Design Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   Connection: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfDesign
Dim mstrShow
Dim mrsColor
Dim mrsFontFace
Dim mrsFontSize
Dim mvntID

Set mrsColor = GetRS("Select slctvalColor,slctvalColorCode from sfSelectValues where slctvalColor<>'' Order by slctvalColor")
Set mrsFontFace = GetRS("Select slctvalFontType from sfSelectValues where slctvalFontType<>'' Order by slctvalFontType")
Set mrsFontSize = GetRS("Select slctvalFontSize from sfSelectValues where slctvalFontSize<>'' Order by slctvalFontSize")

    mvntID = Request.QueryString("dsgnID")
    If len(mvntID) = 0 Then mvntID = Request.Form("dsgnID")

    mstrShow = Request.QueryString("Show")
    If Len(mstrShow) = 0 Then mstrShow = Request.Form("Show")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclssfDesign = New clssfDesign
    With mclssfDesign
    
		Select Case mAction
		    Case "New", "Update"
		        .Update
		        If (lCase(.dsgnIsActive) = "on") Then
					.Activate mvntID
					If .LoadAll Then 
						.Find mvntID
						.WriteDesignToFile
					End If
				Else	
					If .LoadAll Then .Find mvntID
				End If
		    Case "Delete"
		        .Delete mvntID
		        .LoadAll
		    Case "View"
		        If .LoadAll Then .Find mvntID
		    Case "Activate"
		        .Activate mvntID
		        If .LoadAll Then 
					.Find mvntID
					.WriteDesignToFile
				End If
		    Case Else
		        mclssfDesign.LoadAll
		End Select
    
Call WriteHeader("body_onload();",True)
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var theKeyField;
var strDetailTitle = "<%= .dsgnName %> Details";
var strBase;
var strNoImage;

function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = theDataForm.dsgnID;
	strBase = theDataForm.strBaseHRef.value;
	strNoImage = strBase + "ssl/admin/ssadmin/images/NoImage.gif";
	SetDesign();
}

function SetColor(theItem,strColor)
{
if (strColor == "#XXXXXX")
	{ theItem = ""; }
else
{
alert(strColor);
theItem = strColor; } 
}

function SafeColor(strColor)
{
if (strColor == "#XXXXXX")
	{ return(""); }
else
	{ return(strColor); } 
}

function SetImage(strTarget, strSource)
{
var strImage = eval("theDataForm." + strSource).value;

	if (strImage == "")
	{ 
		document.all(strTarget).style.background = "";
		document.all("img" + strSource).src = strNoImage;
	}else{
		document.all(strTarget).style.background = "URL(" + strBase + strImage + ")";
		document.all("img" + strSource).src = strBase + eval("theDataForm." + strSource).value;
	} 
}

function SetSectionSettings(strSource,intNumber)
{
var theItem = document.all(strSource);

	//theItem.style.backgroundColor = SafeColor(eval("theDataForm.dsgnBGColor" + intNumber + ".value"));
	//SetImage(strSource,"dsgnBgnd" + intNumber);
	theItem.style.color = SafeColor(eval("theDataForm.dsgnFontColor" + intNumber + ".value"));
	theItem.style.fontSize = eval("theDataForm.dsgnFontSize" + intNumber + ".value") + "ex";
	theItem.style.fontFamily = eval("theDataForm.dsgnFontFace" + intNumber + ".value");
}

function UpdateDesign()
{
SetDesign();
}

function SetDesign()
{
//var strBase = theDataForm.strBaseHRef.value;
//var strImage = strBase + theDataForm.dsgnGenBgnd.value;

SetSectionSettings("tdContent",4);		//     '---- Content ----


//     '---- General Settings ----"

var strSource;
var strImage;

for (var i=1; i<24; i++)
{
	if (i<10)
	{
		strSource = "dsgnBTN0" + i
	}else{
		strSource = "dsgnBTN" + i
	}
	strImage = eval("theDataForm." + strSource).value;
	
		if (strImage == "")
		{ 
			document.all("img" + strSource).src = strNoImage;
		}else{
			document.all("img" + strSource).src = strBase + strImage;
		} 
}

blnReset = false;
return(true);
}

function SetDefaults(theForm)
{
    theForm.dsgnBTN01.value = "";
    theForm.dsgnBTN02.value = "";
    theForm.dsgnBTN03.value = "";
    theForm.dsgnBTN04.value = "";
    theForm.dsgnBTN05.value = "";
    theForm.dsgnBTN06.value = "";
    theForm.dsgnBTN07.value = "";
    theForm.dsgnBTN08.value = "";
    theForm.dsgnBTN09.value = "";
    theForm.dsgnBTN10.value = "";
    theForm.dsgnBTN11.value = "";
    theForm.dsgnBTN12.value = "";
    theForm.dsgnBTN13.value = "";
    theForm.dsgnBTN14.value = "";
    theForm.dsgnBTN15.value = "";
    theForm.dsgnBTN16.value = "";
    theForm.dsgnBTN17.value = "";
    theForm.dsgnBTN18.value = "";
    theForm.dsgnBTN19.value = "";
    theForm.dsgnBTN20.value = "";
    theForm.dsgnBTN21.value = "";
    theForm.dsgnBTN22.value = "";
    theForm.dsgnBTN23.value = "";
    theForm.dsgnBTN24.value = "";
    theForm.dsgnFontColor4.value = "";
    theForm.dsgnFontFace4.value = "";
    theForm.dsgnFontSize4.value = "";

//    theForm.dsgnForm.value = "";
    theForm.dsgnGenALink.value = "";
    theForm.dsgnID.value = "";
    theForm.dsgnIsActive.value = "";
    theForm.dsgnName.value = "";
    theForm.dsgnThumbnail.value = "";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add Design";
    theForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + theForm.PromoTitle.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "Delete";
    theForm.submit();
    return(true);
    }
    Else
    {
    return(false);
    }
}

var blnReset = false;

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
    blnReset = true;
    return true;
}

function ValidInput(theForm)
{
  if (theForm.dsgnName.value=="")
  {
	alert("Please enter a design name.");
	theForm.dsgnName.focus();
	theForm.dsgnName.select();
	return(false);
   }
   {
    return(true);
   }
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theDataForm.Action.value = "View";
	theDataForm.submit();
	return false;
}

function HideSections()
{
     document.all("tblContent").style.display = "none";
     document.all("tblButton").style.display = "none";
}

function DisplaySection(strSection)
{
var pstrSection = "tbl" + strSection;

  frmData.Show.value = strSection;
  HideSections();
  document.all(pstrSection).style.display = "";

return(false);
}

var gobjImage;
var gblnSwitch;

function SelectImage(theImage)
{
	gblnSwitch = true;
	gobjImage = theImage;
	document.frmData.tempFile.click();
	return false;
}

function ProcessPath(theFile)
{
var pstrFilePath = theFile.value;
var pstrBaseHRef = document.frmData.strBaseHRef.value;
var pstrBasePath = document.frmData.strBasePath.value;
var pstrHREF;
var pstrItem;
var xyz = "\\";

	if (gblnSwitch)
	{
	gobjImage.src = pstrFilePath;
	pstrItem = gobjImage.name.replace("img","");
	pstrHREF = pstrFilePath.replace(pstrBasePath,"");
	eval("document.frmData." + pstrItem).value = pstrHREF.replace(xyz,"/");
	gblnSwitch = false;
	theFile.value = "";
	UpdateDesign();
	}

}

//-->
</SCRIPT>

<BODY onload="body_onload();">
<CENTER>
<DIV class="pagetitle "><%= mstrPageTitle %></DIV>
<%= .OutputMessage %>
<%= .OutputSummary %>

<FORM action='sfDesignAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<INPUT type=hidden id=dsgnID name=dsgnID value=<%= .dsgnID %>>
<INPUT type=hidden id=Action name=Action value='Update'>
<INPUT type=hidden id=Show name=Show Value=''>
<INPUT type=hidden id=strBaseHRef name=strBaseHRef Value='<%= mstrBaseHRef %>'>
<INPUT type=hidden id=strBasePath name=strBasePath Value='<%= mstrBasePath %>'>
<SPAN id=spantempFile style="display:none">
<INPUT type=file id=tempFile name=tempFile onchange="ProcessPath(this);">
</SPAN>
<TABLE class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" ID="Table1">
<TR>
<TD colspan=2>
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" ID="Table2">
<TR class='tblhdr'>
<TH colspan="2" align=center><SPAN id="spanDetailTitle"><%= .dsgnName %> Detail</SPAN></TH>
      </TR>
     <TR>
        <TD class="Label"><LABEL id=lbldsgnName for=dsgnName>Name:&nbsp;</LABEL></TD>
        <TD><INPUT id=dsgnName name=dsgnName Value='<%= .dsgnName %>' maxlength=50 size=50></TD>
      </TR>
       <TR>
        <TD class="Label"><LABEL id=lbldsgnDescription for=dsgnDescription>Description:&nbsp;</LABEL></TD>
        <TD><INPUT id=dsgnDescription name=dsgnDescription Value='<%= .dsgnDescription %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;</TD>
        <TD><INPUT type=checkbox id=dsgnIsActive name=dsgnIsActive <% If (.dsgnIsActive=1) Then Response.Write "Checked" %>>&nbsp;<LABEL id=lbldsgnIsActive for=dsgnIsActive>Is Active</LABEL></TD>
      </TR>
<!--
      <TR>
        <TD colspan="2"><hr></TD>
      </TR>
      <TR>
        <TD class="Label"><LABEL id=lblstrBaseHRef for=strBaseHRef>Website Base URL:&nbsp;</LABEL></TD>
        <TD><INPUT id=strBaseHRef name=strBaseHRef Value='<%= mstrBaseHRef %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD class="Label"><LABEL id=lblstrBasePath for=strBasePath>File Path for Base URL:&nbsp;</LABEL></TD>
        <TD><INPUT id=strBasePath name=strBasePath Value='<%= mstrBasePath %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD colspan="2"><hr></TD>
      </TR>
-->
</TABLE>
</TD>
</TR>
<TR>
<TD width=25% align=center valign=top><!-- Start Reference Table -->
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="0" rules="none" id="tblReference">
<TR><TD>
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id='tdReference'>
<SPAN align=center title="View General Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">
<TR>
<TD width=100% align=center>
General Settings
</TD>
</TR>
<TR>
<TD width=100% align=center>
<DIV id=divLink>Hyperlink</DIV>
<DIV id=divvLink>Visited Hyperlink</DIV>
<DIV id=divaLink>Active Hyperlink</DIV>
</TD>
</TR>
</SPAN>
  <TR>
    <TD width="100%">
      <TABLE border="1" width="100%" cellspacing="0" ID="Table3">
        <TR>
          <TD align=center width="100%" id="tdTop" title="View Top Banner Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Top Banner</TD>
        </TR>
        <TR>
          <TD align=center width="100%" id="tdMiddle" title="View Middle Banner Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Middle Banner</TD>
        </TR>
        <TR>
          <TD align=center width="100%" id="tdBottom" title="View Bottom Banner Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Bottom Banner</TD>
        </TR>
        <TR><TD align=center>
        <TABLE width=100% border=0 ID="Table4">
        <TR>
          <TD align=center width="100%" id="tdContent" style='cursor:hand;' onclick="return DisplaySection('Content');" title="View Content Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Content</TD>
        </TR>
        <TR>
        <TD align=center>
        <TABLE width=95% ID="Table5">
        <TR>
          <TD align=center style='border-style:solid; border-color:black; border-left-width:1pt; border-right-width:1pt; border-bottom-width:1pt; border-top-width:1pt;' width="100%" id="tdContentBar" title="View Content Bar Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Content Bar</TD>
        </TR>
        <TR>
          <TD align=center style='border-style:solid; border-color:black; border-left-width:1pt; border-right-width:1pt; border-bottom-width:1pt; border-top-width:1pt;' width="100%" id="tdAlt1" title="View Alternating Color 1 Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Content Alternating Color 1</TD>
        </TR>
        <TR>
          <TD align=center style='border-style:solid; border-color:black; border-left-width:1pt; border-right-width:1pt; border-bottom-width:1pt; border-top-width:1pt;' width="100%" id="tdAlt2" title="View Alternating Color 2 Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Content Alternating Color 2</TD>
        </TR>
        </TD></TR>
        </TABLE>
        <TR>
			<TD align=center width="100%" id="tdContent2" onclick="return DisplaySection('Content');" title="View Content Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">&nbsp;</TD>
		</TR>
        </TABLE>
        <TR>
          <TD align=center width="100%" id="tdFooter" title="View Footer Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');">Footer</TD>
        </TR>
      </table>
    </TD>
  </TR>
</table>
</td></tr>
  <TR>
	<TD align=center width="100%" id="tdButtons" onclick="return DisplaySection('Button');" title="View Button Settings" onmouseover="DisplayTitle(this);" onmouseout="ClearTitle('');"><SPAN style="color:whitesmoke; background-color:steelblue; border: lightsteelblue outset;">Buttons</div></TD>
  </TR>
</table>
</TD><!-- End Reference Table -->
<TD width=75% valign=top>
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id='tblContent'>
      <TR class='tblhdr'>
        <TH colspan="2">Content Settings</TH>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnFontColor4 for=dsgnFontColor4>Font Color</LABEL></TD>
        <TD>
			<SELECT  onchange="UpdateDesign();" size="1"  id=dsgnFontColor4 name=dsgnFontColor4>
			<% Call MakeCombo(mrsColor,"slctvalColor","slctvalColorCode",.dsgnFontColor4) %>
			</SELECT>
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnFontSize4 for=dsgnFontSize4>Font Size</LABEL></TD>
        <TD>
			<SELECT  onchange="UpdateDesign();" size="1"  id=dsgnFontSize4 name=dsgnFontSize4>
			<% Call MakeCombo(mrsFontSize,"slctvalFontSize","slctvalFontSize",.dsgnFontSize4) %>
			</SELECT>
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnFontFace4 for=dsgnFontFace4>Font Face</LABEL></TD>
         <TD>
			<SELECT  onchange="UpdateDesign();" size="1"  id=dsgnFontFace4 name=dsgnFontFace4>
			<% Call MakeCombo(mrsFontFace,"slctvalFontType","slctvalFontType",.dsgnFontFace4) %>
			</SELECT>
		</TD>
      </TR>
</TABLE>
<TABLE class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id='tblButton'>
      <TR>
        <TH colspan="2">Button Settings</TH>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN01 for=dsgnBTN01>Go</LABEL></TD>
        <TD>
			 <INPUT onchange="UpdateDesign();" id=dsgnBTN01 name=dsgnBTN01 Value='<%= .dsgnBTN01 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN01 id=imgdsgnBTN01 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN01) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Go' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN02 for=dsgnBTN02>Save to Cart</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN02 name=dsgnBTN02 Value='<%= .dsgnBTN02 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN02 id=imgdsgnBTN02 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN02) %>" 
				 onclick="return SelectImage(this);" 
				 tiave to Cart' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN03 for=dsgnBTN03>Add to Cart</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN03 name=dsgnBTN03 Value='<%= .dsgnBTN03 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN03 id=imgdsgnBTN03 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN03) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Add to Cart' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN04 for=dsgnBTN04>Continue Search</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN04 name=dsgnBTN04 Value='<%= .dsgnBTN04 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN04 id=imgdsgnBTN04 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN04) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Continue Search' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN05 for=dsgnBTN05>Checkout</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN05 name=dsgnBTN05 Value='<%= .dsgnBTN05 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN05 id=imgdsgnBTN05 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN05) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Checkout' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN06 for=dsgnBTN06>Delete</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN06 name=dsgnBTN06 Value='<%= .dsgnBTN06 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN06 id=imgdsgnBTN06 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN06) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Delete' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN07 for=dsgnBTN07>Save</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN07 name=dsgnBTN07 Value='<%= .dsgnBTN07 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN07 id=imgdsgnBTN07 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN07) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Save' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN08 for=dsgnBTN08>View Saved Cart</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN08 name=dsgnBTN08 Value='<%= .dsgnBTN08 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN08 id=imgdsgnBTN08 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN08) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'View Saved Cart' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN09 for=dsgnBTN09>Return to Shop</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN09 name=dsgnBTN09 Value='<%= .dsgnBTN09 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN09 id=imgdsgnBTN09 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN09) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Return to Shop' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN10 for=dsgnBTN10>Shopping Cart (Order)</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN10 name=dsgnBTN10 Value='<%= .dsgnBTN10 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN10 id=imgdsgnBTN10 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN10) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Shopping Cart (Order)' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN11 for=dsgnBTN11>Change Cart</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN11 name=dsgnBTN11 Value='<%= .dsgnBTN11 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN11 id=imgdsgnBTN11 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN11) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Change Cart' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN12 for=dsgnBTN12>Sign Up</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN12 name=dsgnBTN12 Value='<%= .dsgnBTN12 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN12 id=imgdsgnBTN12 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN12) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Sign Up' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN13 for=dsgnBTN13>Shopping Cart</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN13 name=dsgnBTN13 Value='<%= .dsgnBTN13 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN13 id=imgdsgnBTN13 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN13) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Shopping Cart' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN14 for=dsgnBTN14>Recalculate</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN14 name=dsgnBTN14 Value='<%= .dsgnBTN14 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN14 id=imgdsgnBTN14 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN14) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Recalculate' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN15 for=dsgnBTN15>Help</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN15 name=dsgnBTN15 Value='<%= .dsgnBTN15 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN15 id=imgdsgnBTN15 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN15) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Help' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN16 for=dsgnBTN16>Login</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN16 name=dsgnBTN16 Value='<%= .dsgnBTN16 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN16 id=imgdsgnBTN16 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN16) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Login' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN17 for=dsgnBTN17>Forgot Password</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN17 name=dsgnBTN17 Value='<%= .dsgnBTN17 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN17 id=imgdsgnBTN17 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN17) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Forgot Password' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN18 for=dsgnBTN18>Submit</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN18 name=dsgnBTN18 Value='<%= .dsgnBTN18 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN18 id=imgdsgnBTN18 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN18) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Submit' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN19 for=dsgnBTN19>New Account</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN19 name=dsgnBTN19 Value='<%= .dsgnBTN19 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN19 id=imgdsgnBTN19 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN19) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'New Account' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN20 for=dsgnBTN20>Verify</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN20 name=dsgnBTN20 Value='<%= .dsgnBTN20 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN20 id=imgdsgnBTN20 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN20) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Verify' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN21 for=dsgnBTN21>Search</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN21 name=dsgnBTN21 Value='<%= .dsgnBTN21 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN21 id=imgdsgnBTN21 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN21) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Search' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN22 for=dsgnBTN22>Add to Cart (Small)</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN22 name=dsgnBTN22 Value='<%= .dsgnBTN22 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN22 id=imgdsgnBTN22 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN22) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Add to Cart (Small)' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN23 for=dsgnBTN23>Clear Shipping Info</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN23 name=dsgnBTN23 Value='<%= .dsgnBTN23 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN23 id=imgdsgnBTN23 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN23) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Clear Shipping Info' button">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lbldsgnBTN24 for=dsgnBTN24>Email A Friend</LABEL></TD>
        <TD> <INPUT onchange="UpdateDesign();" id=dsgnBTN24 name=dsgnBTN24 Value='<%= .dsgnBTN24 %>' maxlength=200 size=60>
			<IMG name=imgdsgnBTN24 id=imgdsgnBTN24 border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.dsgnBTN24) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit 'Email A Friend' button">
		</TD>
      </TR>
</TABLE>
<TABLE border=0 cellPadding=1 cellSpacing=1 width='100%' ID="Table6">
  <TR>
    <TD align=center>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick(this)'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='btnReset_onclick(this);' onblur='UpdateDesign();'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</TD>
</TR>
</TABLE>
</FORM>
</CENTER>
</BODY>
<%
    End With

	If len(mstrShow)>0 then 
		Response.Write "<Script>DisplaySection(" & chr(34) & mstrShow & chr(34) & ");</script>"
	else
		Response.Write "<Script>DisplaySection(" & chr(34) & "Content" & chr(34) & ");</script>"
	end if

    Set mclssfDesign = Nothing
	Set mrsColor = Nothing
	Set mrsFontFace = Nothing
	Set mrsFontSize = Nothing
    cnn.close
    Set cnn = Nothing
	
    Response.Flush
%>