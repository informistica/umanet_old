<%
'	Author: Danilo Cicognani
'	Script: resize.asp
'	Version: 2.01
'	Date: 24/08/2007
'	Use: Resizes an image GIF, JPG or PNG using the library AspImage or Microsoft Office Web Components
'	Copyright (c) 2007 Danilo Cicognani
'	https://www.ciconet.it
'	You can use, modify and redistribute this file, but you must mantain this copyright notice

Private Function Mult(lsb, msb)
	'Sum less significant byte to more significant byte
	Mult = lsb + (msb * CLng(256))
End Function

'Determine image size
Sub ImageSize(FileName, Width, Height)
	Const ForReading = 1
	Const BUFFERSIZE = 65535
	Dim bBuf(65535)
	'Determine image type
	j = InStrRev(FileName, ".")
	fileExt = right(FileName, Len(FileName) - j + 1)
	'Read image file
	Dim fso, f
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(FileName, ForReading)
	i = 0
	Do While (Not f.AtEndOfStream) And i < BUFFERSIZE
		bBuf(i) = f.Read(1)
		i = i + 1
	Loop
	Set f = Nothing
	Set fso = Nothing
	Select Case fileExt
		Case ".gif"	'If the image is a GIF
			Width = Mult(Asc(bBuf(6)), Asc(bBuf(7)))
			Height = Mult(Asc(bBuf(8)), Asc(bBuf(9)))
		Case ".png"	'If the image is a PNG
			Width = Mult(Asc(bBuf(19)), Asc(bBuf(18)))
			Height = Mult(Asc(bBuf(23)), Asc(bBuf(22)))
		Case ".jpg", ".jpeg"	'If the image is a JPG
			Do
				' loop through looking for the byte sequence FF,D8,FF
				' which marks the begining of a JPEG file
				' lPos will be left at the postion of the start
				If (Asc(bBuf(lPos)) = &HFF And Asc(bBuf(lPos + 1)) = &HD8 And Asc(bBuf(lPos + 2)) = &HFF) Or (lPos >= BUFFERSIZE - 10) Then Exit Do
				' move our pointer up
				lPos = lPos + 1
			Loop
			lPos = lPos + 2
			If lPos >= BUFFERSIZE - 10 Then Exit Sub
			Do
				' loop through the markers until we find the one 
				'starting with FF,C0 which is the block containing the 
				'image information
				Do
					' loop until we find the beginning of the next marker
					If Asc(bBuf(lPos)) = &HFF And Asc(bBuf(lPos + 1)) <> &HFF Then Exit Do
					lPos = lPos + 1
					If lPos >= BUFFERSIZE - 10 Then Exit Sub
				Loop
				' move pointer up
				lPos = lPos + 1
				Select Case Asc(bBuf(lPos))
					Case &HC0, &HC1, &HC2, &HC3, &HC5, &HC6, &HC7, &HC9, &HCA, &HCB, &HCD, &HCF
						' we found the right block
						Exit Do
				End Select
				' otherwise keep looking
				lPos = lPos + Mult(Asc(bBuf(lPos + 2)), Asc(bBuf(lPos + 1)))
				' check for end of buffer
				If lPos >= BUFFERSIZE - 10 Then Exit Sub
			Loop
			Width = Mult(Asc(bBuf(lPos + 7)), Asc(bBuf(lPos + 6)))
			Height = Mult(Asc(bBuf(lPos + 5)), Asc(bBuf(lPos + 4)))
	End Select
End Sub

'Resizes the image FileName, to file OutFileName, to OutFormat, to Width and Height specified, preserving aspect ratio
Sub ResizeImage(FileName, OutFileName, OutFormat, Width, Height)
	On Error Resume Next
	Err = 0
	Set Image = Server.CreateObject("AspImage.Image")
	If Err = 0 Then
		'AscImage available: use it
		'Load image
		Image.LoadImage FileName

		'Get the original file size
		OriginalWidth = Image.MaxX
		OriginalHeight = Image.MaxY

		'Calculate the scaled size
		ScaledWidth = Width
		ScaledHeight = Height
		If OriginalWidth < Width Then
			If OriginalHeight < Height Then
				ScaledWidth = OriginalWidth
				ScaledHeight = OriginalHeight
			Else
				ScaledWidth = CInt(OriginalWidth / (OriginalHeight / Height))
			End If
		ElseIf OriginalHeight < Height Then
			ScaledHeight = CInt(OriginalHeight / (OriginalWidth / Width))
		Else
			If OriginalWidth / Width > OriginalHeight / Height Then
				ScaledHeight = CInt(OriginalHeight / (OriginalWidth / Width))
			Else
				ScaledWidth = CInt(OriginalWidth / (OriginalHeight / Height))
			End If
		End If

		'export the picture to a file
		Image.ResizeR ScaledWidth, ScaledHeight
		Image.FileName = OutFileName
		Select Case OutFormat
			Case "JPG"
				Image.ImageFormat = 1
			Case "PNG"
				Image.ImageFormat = 3
			Case "GIF"
				Image.ImageFormat = 5
		End Select
		Image.SaveImage
	Else
		'AspImage unavailable, use OWC
		Dim Chs, chConstants

		'Create an OWC chart object
		Set Chs = CreateObject("OWC10.ChartSpace")
		Set chConstants = Chs.Constants

		'Set background of the chart
		Chs.Interior.SetTextured FileName, chConstants.chStretchPlot, , chConstants.chAllFaces

		'Se the border of the chart (unlikely it's impossible to not have a border)
		Chs.border.color = RGB(255, 255, 255)
		Chs.Border.Weight = chConstants.owcLineWeightHairline

		'Get the original file size
		OriginalWidth = 0
		OriginalHeight = 0
		ImageSize FileName, OriginalWidth, OriginalHeight

		'Calculate the scaled size
		ScaledWidth = Width
		ScaledHeight = Height
		If OriginalWidth < Width Then
			If OriginalHeight < Height Then
				ScaledWidth = OriginalWidth
				ScaledHeight = OriginalHeight
			Else
				ScaledWidth = CInt(OriginalWidth / (OriginalHeight / Height))
			End If
		ElseIf OriginalHeight < Height Then
			ScaledHeight = CInt(OriginalHeight / (OriginalWidth / Width))
		Else
			If OriginalWidth / Width > OriginalHeight / Height Then
				ScaledHeight = CInt(OriginalHeight / (OriginalWidth / Width))
			Else
				ScaledWidth = CInt(OriginalWidth / (OriginalHeight / Height))
			End If
		End If

		'export the picture to a file
		Chs.ExportPicture OutFileName, OutFormat, ScaledWidth, ScaledHeight

		Set Chs = Nothing
		Set chConstants = Nothing
	End If
	Set Image = Nothing
End Sub

'Example
'ResizeImage server.mappath(".") & "/test.png", server.mappath(".") & "/output.jpg", "JPG", 100, 150
%>