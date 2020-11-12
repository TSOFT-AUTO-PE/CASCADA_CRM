Dim var1, t, Num_Iter, filas

Dim str_departamento
Dim str_provincia
Dim str_distrito
Dim str_nombreVia
Dim str_DepNum
Dim str_Busca
Dim str_Categoria
Dim str_valEstadoOrden
Dim str_tipoalta
Dim str_motivo_alta
Dim intStartTime, intStopTime
Dim tiempo

intStartTime = Timer

str_departamento	=	DataTable("e_Departamento", dtLocalSheet)
str_provincia		=	DataTable("e_Provincia", dtLocalSheet) 
str_tipoalta		=	DataTable("e_TipodeAlta", dtLocalSheet)
str_motivo_alta		=	DataTable("e_MotivoAlta", dtLocalSheet)
str_distrito        =	DataTable("e_Distrito", dtLocalSheet) @@ hightlight id_;_31334378_;_script infofile_;_ZIP::ssf8.xml_;_
str_nombreVia 		=	DataTable("e_NombreVia", dtLocalSheet)
str_DepNum 			=	DataTable("e_DepNum", dtLocalSheet)
str_Busca 			=	DataTable("e_Busca", dtLocalSheet)
str_Categoria		= 	DataTable("e_Categoria", dtLocalSheet)
Num_Iter 			= 	Environment.Value("ActionIteration")
 
Call SeleccionarTipoAlta()
Call FlujoWIC()
Call Direccion()
Call SeleccionarPlanTarifario()
Call ParametrosAlta()
Call RecursosAltas()
Call NegociarVisita()
Call NegociarDistribucion()
'Call GeneracionOrden()
'''Si falla la WIC2, habilitar los siguientes 2 "CALL" y comentar Call GeneracionOrden()
Call InspectorSmart()
Call EnviarOrden()

Sub SeleccionarTipoAlta()
'Buscar Cliente
RunAction "Buscar_Cliente_2", oneIteration


	Select Case DataTable("e_TipodeAlta", dtLocalsheet)
		Case "Alta Producto Fija + Mono"
			Dim c
			c = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			While c = 0
				wait 1
				c = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			Wend
			Call Carga()
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").Select "Presencial" @@ hightlight id_;_7021600_;_script infofile_;_ZIP::ssf118.xml_;_
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Medio:").Select "Presencial" @@ hightlight id_;_14214434_;_script infofile_;_ZIP::ssf119.xml_;_
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Comunicación:").Select "Entrante" @@ hightlight id_;_1838632_;_script infofile_;_ZIP::ssf120.xml_;_
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#9"
			wait 1
	End	Select
End Sub




Sub FlujoWIC()
	'Llamar a la WIC1
	If DataTable("e_WIC_ValidaCli", dtLocalsheet)="SI" Then
RunAction "WIC1", oneIteration
	End If

End  Sub

Sub Direccion()
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaEdit("Dirección").Exist = false
	wait 1
Wend
Dim x
x = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaEdit("Dirección").GetROProperty("text")
If (x = "") Then
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Lookup-notValidated").Click
wait 6
Dim filasT
filasT = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaTable("SearchJTable").GetROProperty("rows")	
If filasT >0 Then
	dim Iterator , direccion	
	For Iterator = filasT-1 To 0 step -1	    
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaTable("SearchJTable").SelectRow ("#"&Iterator)		
		direccion = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaTable("SearchJTable").GetCellData("#"&Iterator, "#3")	    
		If (direccion = "") Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaButton("Crear").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_10").JavaList("Departamento:").Exist = false
	wait 1
Wend
Call Carga()
wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_10").JavaList("Departamento:").Select DataTable("e_Departamento", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Provincia:").Set DataTable("e_Provincia", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Lookup-notValidated_2").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Distrito:").Exist = false
wait 1
wend
Call Carga()
wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Distrito:").Set DataTable("e_Distrito", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Lookup-notValidated_3").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Nombre de Vía:").Exist = false
	wait 1
wend
Call Carga()
wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Nombre de Vía:").Set DataTable("e_NombreVia", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Número:").Set DataTable("e_DepNum", dtLocalSheet)
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Direccion.png", True
	imagenToWord "Direccion", RutaEvidencias() & Ciclo&"Direccion.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Validar").Click
Call Carga()
wait 7
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Guardar").Click	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Direccion.png", True
	imagenToWord "Validando dirección", RutaEvidencias() & "Direccion.png"
Call Carga()
wait 2
		else
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaButton("Seleccionar").Click
		End If
	Next
Else
Dim wtt
		wtt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaButton("Crear").GetROProperty("enabled")
		While wtt = 0
			wait 1
		wtt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaButton("Crear").GetROProperty("enabled")
		Wend
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_6").JavaButton("Crear").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_10").JavaList("Departamento:").Exist = false
	wait 1
Wend
Call Carga()
wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_10").JavaList("Departamento:").Select DataTable("e_Departamento", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Provincia:").Set DataTable("e_Provincia", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Lookup-notValidated_2").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Distrito:").Exist = false
	wait 1
wend
Call Carga()
wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Distrito:").Set DataTable("e_Distrito", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Lookup-notValidated_3").Click
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Nombre de Vía:").Exist = false
wait 1
wend
Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Nombre de Vía:").Set DataTable("e_NombreVia", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaEdit("Número:").Set DataTable("e_DepNum", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Validar").Click
Call Carga()
wait 7
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Guardar").Click	
	'imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Direccion.png", True
	imagenToWord "Validando dirección", RutaEvidencias() & "Direccion.png"
Call Carga()
wait 2
End If
End If
'Validacion de direccion no encontrada
If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Falló la validación de").Exist = True Then
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ProblemaValidacion.png", True
	imagenToWord "Direccion Validada", RutaEvidencias() & Ciclo&"ProblemaValidacion.png"
	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Lookup-notValidated_3").Click
	Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Validar").Click
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Guardar").Exist = false
		wait 1
	Wend
	Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Guardar").Click
End if
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").Exist = False
	wait 1
Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").Click
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Direccion_Validada.png", True
	imagenToWord "Direccion Validada", RutaEvidencias() & Ciclo&"Direccion_Validada.png"
End Sub

Sub SeleccionarPlanTarifario()
'Validar mensaje de problema 
If  JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist = True Then
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Problema.png", True
	imagenToWord "Direccion Validada", RutaEvidencias() & Ciclo&"Problema.png"
	JavaWindow("Server no disponible").JavaDialog("Problema").JavaButton("OK").Click
	ExitActionIteration
End If
'Esperar a que comboBox sea visible
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").Exist = False
wait 1	
Wend
'Selecciona Plan tarifario
Select Case DataTable("e_Categoria", dtLocalSheet)
'Selecciona Mono
	Case "Monos"
	'wait 4
	Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 6000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").Select DataTable("e_Categoria", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 4000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaEdit("Buscar por LOB").Set DataTable("e_Busca", dtLocalSheet)
	Call Carga()
		'wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Buscar").Click
	Call Carga()
		'wait 2
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"PlanTarifario.png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() & Ciclo&"PlanTarifario.png"
	'Esperar a que elemento sea visible
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
		wait 1
	Wend
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		wait 1
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("No existen ofertas elegibles").GetROProperty("text")
		DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)=varsap
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png", True
		imagenToWord "ErrorPlanTarifario", RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Cerrar").Click
		wait 1
		ExitActionIteration
	End If
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Set "ON"
	Call Carga()
	Dim aa
		aa = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		While aa = 0
			wait 1
		aa = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").Click
	End	Select
End Sub

Sub ParametrosAlta()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist = False
	wait 1	
	Wend
	Call Carga()
		wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select "Pedido de Cliente"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_motivo_alta
	Call Carga()
		wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png", True
	imagenToWord "Atributos del Producto", RutaEvidencias() &Num_Iter&"_"&"Atributos del Producto"&".png"
	'Validar a todos si el boton esta
	If 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Aplicar a Todos").Exist = True Then
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Aplicar a Todos").Click
	End If
	'Valida mensaje
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ParametroAlta.png", True
	imagenToWord "Parametros de Alta", RutaEvidencias() & Ciclo&"ParametroAlta.png"
	Dim e
		e = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").GetROProperty("enabled")
		While e = 0
			wait 1
		e = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	'Mensaje de validacion
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		ExitActionIteration
	End If
End Sub


Sub RecursosAltas()
	Select Case DataTable("e_RecursosAltas", dtLocalsheet)
	
		'Selecciona producto de Mono Voz línea movistar
		Case "Producto Mono Voz" 'Mono línea movistar / control
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend
		'Accesorios @@ hightlight id_;_18051293_;_script infofile_;_ZIP::ssf492.xml_;_
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
		wait 1
		Wend @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf140.xml_;_
		Dim filasTT
		filasTT = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filasTT >=0 Then
		dim  trio	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#2")		
		trio = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#2", "#2")	
		'wait 7
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Set "ON"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaButton("Agregar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		wait 2
		End If @@ hightlight id_;_8694380_;_script infofile_;_ZIP::ssf493.xml_;_
		'producto de equipo @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf500.xml_;_
			Call Carga()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#1" @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf136.xml_;_
			Call Carga()
			wait 1
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
			End If
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
			Call Carga()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
			'Error Interno
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			'Imagen
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
			imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
			ExitActionIteration
			End If
			Call Carga()
			'Numero Telefono
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").Exist = false
			wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#2" @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf150.xml_;_
			Call Carga()
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTab("0.00").Select "Asignación de número" @@ hightlight id_;_6191933_;_script infofile_;_ZIP::ssf151.xml_;_
			Call Carga()
			wait 1
			Dim ii1
			ii1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
			While ii1 = 0
			wait 1
			ii1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").Click @@ hightlight id_;_24514005_;_script infofile_;_ZIP::ssf152.xml_;_
			'wait 4
			Dim ii2
			ii2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
			While ii2 = 0
			wait 1
			ii2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").Click @@ hightlight id_;_25618982_;_script infofile_;_ZIP::ssf153.xml_;_
			Call Carga()
			wait 1
			'Guardar numero
			Dim n6
			n6 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("SearchJTable").GetCellData("#0", "#1")
			'MsgBox n6 <<GUARDAR NUMERO>>
			DataTable("s_NumeroAsignado", dtLocalSheet) = n6
			'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
			'POSIBLEMENTE NO ES PARTE DE 
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Guardar").Click @@ hightlight id_;_4265608_;_script infofile_;_ZIP::ssf154.xml_;_
			'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
			Call Carga()
			'wait 1
			'EEDDIITTAARR
			'REGRESAR A CONFIGURACION
			'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaObject("StyleAuxTabbedPaneUI$Scrollabl").Click
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTab("0.00").Select "Configuración"
			Call Carga()
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click @@ hightlight id_;_21995782_;_script infofile_;_ZIP::ssf539.xml_;_
			Call Carga()
			WAIT 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
			Call Carga()
			WAIT 1
			'Imagen
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
			imagenToWord "Recursos Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
		End  Select
		'SIGUIENTE
		Dim w
		w = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Siguiente >").GetROProperty("enabled")
		While w = 0
			wait 1
		w = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Siguiente >").Click
		tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar visita'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaEdit("Teléfono 1").Exist(1)
End Sub

Sub Carga()
'Metodo para cargar elementos	
RunAction "Carga", oneIteration
End Sub


Sub NegociarVisita()
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaEdit("Teléfono 1").Exist = false
	wait 1
Wend
'wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaEdit("Teléfono 1").Set DataTable("e_NumVisita", dtLocalSheet)
	Dim j
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Obtener franjas horarias").GetROProperty("enabled")
		While j = 0
			wait 1
		j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Obtener franjas horarias").GetROProperty("enabled")
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Obtener franjas horarias").Click
	'Espera a que elemento sea visible
	'While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Desde Fecha").Exist = false 'or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist = false) 'then
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Desde Fecha").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = false
		wait 1
	Wend
	if JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Error interno del sistema,").Exist = true then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInternoSistema.png", True
		imagenToWord "Error Interno del sistema", RutaEvidencias() & Ciclo&"ErrorInternoSistema.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitActionIteration
	End If
	Call Carga()
	wait 2
	'wait 5
	'Carga de velocidad de internet
		Dim tab3
		tab3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Desde Fecha").GetROProperty("rows")
		While tab3 = 0 or tab3 =""
			wait 1
		tab3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Desde Fecha").GetROProperty("rows")
		Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Desde Fecha").SelectRow "#0"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaTable("Hasta Fecha").SelectRow "#0" @@ hightlight id_;_4551549_;_script infofile_;_ZIP::ssf126.xml_;_
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Seleccione Cita").Click
	Call Carga() @@ hightlight id_;_18631972_;_script infofile_;_ZIP::ssf127.xml_;_
	wait 2
	'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Visita.png", True
		imagenToWord "Visita", RutaEvidencias() & Ciclo&"Visita.png"
		Dim cc
		cc = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Siguiente >").GetROProperty("enabled")
		While cc = 0
			wait 1
		cc = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Visita (Orden").JavaButton("Siguiente >").Click
		
			tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar distribución'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaStaticText("Acuerdo de Facturación(st)").Exist(1)
End Sub

Sub NegociarDistribucion()
'Espera a que elemento sea visible
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist = false
	wait 1
	Wend
	'Combobox y while = similar
	Dim v
		v = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").GetROProperty("enabled")
		While v = 0
			wait 1
		v = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").GetROProperty("enabled")
		Wend
	Call Carga()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set "ON"
		Dim vb
		vb = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").GetROProperty("enabled")
		While vb = 0
			wait 1
		vb = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").GetROProperty("enabled")
		Wend
	Call Carga()
	wait 5
	'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Distribucion.png", True
		imagenToWord "Distribución", RutaEvidencias() & Ciclo&"Distribucion.png"
		Dim wt
		wt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").GetROProperty("enabled")
		While wt = 0
			wait 1
		wt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
	Call Carga()
	wait 12
	'Error Interno
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
		imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
		'
		JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		ExitActionIteration
		End If
		'Mensaje de validacion
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png", True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"Mensaje de Validación"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		ExitActionIteration
		End If
	'wait 7
	tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Generar Orden'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist(1)


'			''Cuando falla la wic2 (Puedes comentarlo si no quieres el aviso)
'	Dim x1
'	x1 = "¡RECORDATORIO! SE GUARDARÁ LA ORDEN"
'	MsgBox x1
'		''flujo para guardar
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Guardar").Click
'	Call Carga()
'	'Imagen
'	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Resumen_Orden.png", True
'	imagenToWord "Resumen_Orden", RutaEvidencias() & Ciclo&"Resumen_Orden.png"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Close
'	JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Guardar").Click
'	Call Carga()
End Sub

Sub EnviarOrden()
	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png" , True
		imagenToWord "Envio de Orden", RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Enviar orden").Click
	End If
	'Bucle que espera el envío de la orden
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaEdit("TextAreaNative$1").Exist) = False
		wait 1
	Wend

	Dim correcto, orden ,x
	orden = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaEdit("TextAreaNative$1").GetROProperty("text")
	correcto = InStr(orden, "correctamente")
	If correcto <> "0" Then
		Reporter.ReportEvent micPass, "La orden se envio correctamente", "PASS"
	End If
	DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM(replace(replace(orden,"La orden",""),"se envio correctamente.","")))
	wait 1 
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click	
End Sub


Sub InspectorSmart()
	'Antes de guardar
	Dim ordenSmart,hhh,ddd
	ordenSmart = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").GetROProperty("text") 
	hhh = replace(ordenSmart,"Resumen de la orden (Orden ","")
	ddd = left(hhh,7)
	DataTable("s_Nro_Orden_Smart", dtLocalSheet) = ddd
	wait 1 
	''Guardar
	Call CuandoFallaWIC2_Guardar()
	'Despues de guardar
	'>>>>>Ver orden
		tiempo=0
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
		Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden_Smart", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		Call Carga()
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").SelectRow "#0"
		Call Carga()

	'Inspector smart
	JavaWindow("Ejecutivo de interacción").JavaMenu("Ayuda").Select
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Ayuda").JavaMenu("Inspector Smart").Select
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTree("TreeNative$SmartJTree").Exist = False
		wait 1
	Wend
	
	Dim enabled
	enabled = "Habilitar ENABLED"
	MsgBox enabled
'	Call Carga()
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTree("TreeNative$SmartJTree").Expand "Node;Controls;Botón"
'	Call Carga()
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTree("TreeNative$SmartJTree").Select "Node;Controls;Botón;btnNextStep"
'	
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTab("Propiedades").Select "Propiedades"
	'''Habilitar boton Enabled
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").SelectRow "#8" 
	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").GetCellData"#8","#1"
	
'	'<<<<<<<<<<<<<<<
'	dim Iterator , filas, j,h
'	filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").GetROProperty("rows")
'	For Iterator = 0 To filas-1 step 1
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").SelectRow ("#"&Iterator)
'	j = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").GetCellData("#"&Iterator, "#1")
'	h = Instr(1,j,"Enabled")
'	If h <> 0 Then
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").GetROProperty("SelectedRow").DoubleClickCell "#"&Iterator, "#8", "LEFT"
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").SelectRow "#8" 
'	'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").JavaTable("SearchJTable").SetCellData "#"&Iterator, "#8"',str_tipoSIM
'	'JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "IngresoTipoSIM.png", True
'	'imagenToWord "Ingresamos Tipo SIM",RutaEvidencias() & "IngresoTipoSIM.png"
'	Exit for
'	End If
'	Next
'	'>>>>>>>>>>>>
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Smart Inspector V3.0").Close
	'Enviar orden
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable").Exist = False
		wait 1
	Wend

End Sub


Sub CuandoFallaWIC2_Guardar()
	'Cuando falla la wic2
'	Dim x1
'	x1 = "¡RECORDATORIO! GUARDAR"
'	MsgBox x1
		'flujo para guardar
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Guardar").Click
	Call Carga()
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Resumen_Orden.png", True
	imagenToWord "Resumen_Orden", RutaEvidencias() & Ciclo&"Resumen_Orden.png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Close
	JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Guardar").Click
	Call Carga()
End Sub

Sub GeneracionOrden()
	Dim x1
	x1 = "¡RECORDATORIO! Usted realizará los siguientes 3 pasos DESPUÉS de 'VALIDACIÓN NO BIOMETRICA' (Click en ACEPTAR)"
	MsgBox x1
	
Dim tiempo
	tiempo = 0
	Do
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
			If DataTable("e_WIC_ContrCli",dtLocalSheet)="SI" Then
       'Llamar a la WIC 2				
RunAction "WIC2", oneIteration
				Exit Do
			End If
			wait 3
		Wend
		
		'Flujo continua
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(3) Then
			wait 3
			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
	   	 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	   	 	wait 2
		End  If
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
			wait 2
		End If
		wait 1
			
			If tiempo>=180 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se a cargado el contrato correctamente"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"Contrato Exitoso","Se a cargado el contrato correctamente"
			End If
	wait 2
	Loop While Not (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist or (var1 = "Contratos no Generados") or (var1 = "0"))
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Generación de Orden"&".png" , True
	imagenToWord "Generación de Orden", RutaEvidencias() &Num_Iter&"_"&"Generación de Orden"&".png"
	'Mensaje de 
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "GenerarContrato_"&Num_Iter&".png", True
		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").GetROProperty("enabled") <> "0" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		End If
		
	End If
	'MsgBox H
	'Bucle que espera "Enviar orden"
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist) = False
		Wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Enviar orden-"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png", True
			imagenToWord "No se habilitó el botón -Enviar orden_"&Num_Iter, RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	Wait 1

	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png" , True
		imagenToWord "Envio de Orden", RutaEvidencias() &Num_Iter&"_"&"Envio de Orden"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		'JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Enviar orden").Click
	End If
	'Bucle que espera el envío de la orden
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaEdit("TextAreaNative$1").Exist) = False
		wait 1
	Wend

	Dim correcto, orden ,x
	orden = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaEdit("TextAreaNative$1").GetROProperty("text")
	correcto = InStr(orden, "correctamente")
	If correcto <> "0" Then
		Reporter.ReportEvent micPass, "La orden se envio correctamente", "PASS"
	End If
	DataTable("s_Nro_Orden", dtLocalSheet) = RTRIM(LTRIM(replace(replace(orden,"La orden",""),"se envio correctamente.","")))
	wait 1 
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden Generada"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click	
End Sub







 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf537.xml_;_
