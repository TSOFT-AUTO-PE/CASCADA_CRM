﻿Option Explicit
Dim var1, varfila
Dim str_tipoDocumento
Dim str_nroDocumento
Dim CantFilas
Dim Iterator
Dim Rol
Dim Comp
Dim str_DniContacto
Dim NroDig
str_tipoDocumento= DataTable("e_TipoDocumento", "Alta_Producto_Fija")
str_nroDocumento=DataTable("e_NumDocumento",  "Alta_Producto_Fija")
'str_DniContacto=DataTable("e_DniContacto",  "Alta_Producto_Fija")

Call Busqueda_Cliente()

If str_tipoDocumento= "ACUERDO" Then
	Call SeleccionarAcuerdo()
End If

If not str_tipoDocumento="SUSCRIPCION" Then
	Call SeleccionarContacto()
else
	Call seleccionarActivo
End If

Sub Busqueda_Cliente()
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").Close
		End If
		
		JavaWindow("Ejecutivo de interacción").JavaButton("Find-Caller").Click
		
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Cliente").Exist) = False
			wait 1
		Wend
		wait 1
		If str_tipoDocumento="ACUERDO" then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Acuerdo de Facturación").Set
		elseif not str_tipoDocumento="SUSCRIPCION" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Cliente").Set
		End If
 @@ hightlight id_;_5642214_;_script infofile_;_ZIP::ssf3.xml_;_
		wait 1
		Select Case str_tipoDocumento
			Case "RUC"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo ID Compańía:").Select "RUC"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("TextFieldNative$1").SetFocus
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("TextFieldNative$1").Set str_nroDocumento
			Case "DNI"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "DNI" @@ hightlight id_;_11636504_;_script infofile_;_ZIP::ssf8.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento @@ hightlight id_;_31606068_;_script infofile_;_ZIP::ssf9.xml_;_
			Case "Pasaporte"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "Pasaporte" @@ hightlight id_;_11636504_;_script infofile_;_ZIP::ssf8.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento
			Case "CE"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaList("Tipo de documento").Select "CE" @@ hightlight id_;_11636504_;_script infofile_;_ZIP::ssf8.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Numero de Documento").Set str_nroDocumento
			Case "IDCLIENTE"	
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("ID del Cliente:").Set str_nroDocumento
			Case "SUSCRIPCION"
				
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaRadioButton("Suscripción").Set
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("Número de Suscripción:").Set str_nroDocumento
			Case "ACUERDO"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaEdit("ID del Acuerdo de Facturación:").Set str_nroDocumento
				

		End Select @@ hightlight id_;_5678565_;_script infofile_;_ZIP::ssf5.xml_;_
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaCliente.png", True	
		imagenToWord "Carga del Cliente", RutaEvidencias() & "CargaCliente.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Encontrar Comunicante").JavaButton("Buscar ahora").Click @@ hightlight id_;_6577223_;_script infofile_;_ZIP::ssf6.xml_;_
		wait 2
		
End Sub
Sub SeleccionarContacto()
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").Exist(2) Then
 	var1 =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
	 	If (var1= "0 Registros") or (var1= "-- Registros") Then
	 		Reporter.ReportEvent micFail,"Fallido", "Nose se encuentra el RUC:"&DataTable("e_NumDocumento",  "Alta_Producto_Fija")
			ExitTest
	 	else
	 		Select Case ucase(DataTable("e_DocContactoAsociado",  "Alta_Producto_Fija"))
						
				Case "DNI"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaList("ComboBoxNative$1").Select "DNI"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaEdit("TextFieldNative$1").Set DataTable("e_NumContactoAsociado", "Alta_Producto_Fija")
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Buscar ahora").Click
					wait 5
					Dim r, t
					r = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
					While r = "0 Registros"
						r = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
						Wait 1
						t = t + 1
						If (t >= 50) Then
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No se encontro el documento asociado"
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorContacto.png", True
							imagenToWord "Error al buscar contacto", RutaEvidencias() & "ErrorContacto.png"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Wend
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow ("#0")
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContactoEncontrado.png", True
					imagenToWord "Contacto encontrado", RutaEvidencias() & "ContactoEncontrado.png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
							
				Case "Pasaporte"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaList("ComboBoxNative$1").Select "Pasaporte"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaEdit("TextFieldNative$1").Set DataTable("e_NumContactoAsociado", "Alta_Producto_Fija")
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Buscar ahora").Click
					wait 5
					Dim s, f
					s = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
					While s = "0 Registros"
						s = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
						Wait 1
						f = f + 1
						If (f >= 50) Then
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No se encontro el documento asociado"
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorContacto.png", True
							imagenToWord "Error al buscar contacto", RutaEvidencias() & "ErrorContacto.png"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Wend
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow ("#0")
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContactoEncontrado.png", True
					imagenToWord "Contacto encontrado", RutaEvidencias() & "ContactoEncontrado.png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
							
				Case "CE"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaList("ComboBoxNative$1").Select "CE"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaEdit("TextFieldNative$1").Set DataTable("e_NumContactoAsociado", "Alta_Producto_Fija")
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Buscar ahora").Click
					Dim z , l
					z = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
					While z = "0 Registros"
						z = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("0 Registros").GetROProperty("text")
						Wait 1
						l = l + 1
						If (l >= 50) Then
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							DataTable("s_Detalle", dtLocalSheet) = "No se encontro el documento asociado"
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorContacto.png", True
							imagenToWord "Error al buscar contacto", RutaEvidencias() & "ErrorContacto.png"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							ExitActionIteration
						End If
					Wend
					wait 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaTable("SearchJTable").SelectRow ("#0")
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContactoEncontrado.png", True
					imagenToWord "Contacto encontrado", RutaEvidencias() & "ContactoEncontrado.png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Cliente").JavaButton("Seleccionar").Click
				End Select
	 	End If
	End If
End Sub
Sub seleccionarActivo
	Dim i, row, estado, tipoIDCompania, rowActive, nifRow
	i=0
	While not i=10
		i=i+1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").Exist(1) Then
			i=10
			row=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetROProperty("rows")
			For rowActive = 0 To row-1 Step 1
				estado=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,11)
				tipoIDCompania=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,12)
				Rol = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").GetCellData(rowActive,8)
				If (estado="Activo") or (estado="Suspendido") Then
				
					Select Case tipoIDCompania
					
						Case "RUC"
							If Rol = "Titular" Then				
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
								imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
								rowActive=row
							End If
						Case "NIF"
							nifRow=rowActive
						Case "SUSCRIPCION"
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
							imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
							rowActive=row
					End Select
				
'					If tipoIDCompania="RUC" Then
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
'						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
'						imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
'						rowActive=row
'					ElseIf tipoIDCompania="NIF" Then
'						nifRow=rowActive
'						
'					End If
					If rowActive = row-1 Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaTable("SearchJTable").SelectRow "#"&rowActive
						JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
						imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click
					End If
					
				End If	
			Next
			
		End If
	Wend
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaNumeroActivo.png", True	
'	imagenToWord "Carga de números activos", RutaEvidencias() & "CargaNumeroActivo.png"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Suscripci").JavaButton("Seleccionar").Click @@ hightlight id_;_6577223_;_script infofile_;_ZIP::ssf6.xml_;_
	wait 2
End Sub
Sub SeleccionarAcuerdo()
	
	wait 2
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").Exist = True Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").JavaTable("SearchJTable").Exist = True Then
			CantFilas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").JavaTable("SearchJTable").GetROProperty("rows")
			wait 2
			For Iterator = 0 To CantFilas-1 Step 1
				Rol = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").JavaTable("SearchJTable").GetCellData (Iterator, "#7")	
				If Rol = "Facturación" Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").JavaTable("SearchJTable").SelectRow "#"&Iterator
					wait 2
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CargaAcuerdoFact.png", True	
					imagenToWord "Carga Acuerdo de Facturación", RutaEvidencias() & "CargaAcuerdoFact.png"
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Búsqueda: Contacto y Acuerdo").JavaButton("Seleccionar").Click
					
					Exit For
					
				End If
			Next
		End If
	End If

End Sub


