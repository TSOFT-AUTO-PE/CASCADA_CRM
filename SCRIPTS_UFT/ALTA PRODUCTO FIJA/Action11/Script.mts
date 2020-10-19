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
Call GeneracionOrden()

Sub SeleccionarTipoAlta()
'Buscar Cliente
RunAction "Buscar_Cliente_2", oneIteration


	Select Case DataTable("e_TipodeAlta", dtLocalsheet)
		Case "Alta Producto Fija + Trio"
			Dim a
			a = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			While a = 0
				wait 1
				a = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			Wend
			Call Carga()
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").Select "Otros"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Medio:").Select "Otros Medios"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Comunicación:").Select "Saliente"
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#9"
			wait 1
		Case "Alta Producto Fija + Duo"
			Dim b
			b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			While b = 0
				wait 1
				b = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").GetROProperty("enabled")
			Wend
			Call Carga()
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Tipo:").Select "Presencial"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Medio:").Select "Presencial"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_2").JavaList("Comunicación:").Select "Entrante"
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Panel_Interaccion_"&Num_Iter&".png", True
			imagenToWord "Panel de Interacción", RutaEvidencias() &"Panel_Interaccion_"&Num_Iter&".png"
			JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#9"
			wait 1
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
	'Esperar a que elemento sea visible @@ hightlight id_;_9816496_;_script infofile_;_ZIP::ssf529.xml_;_
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Exist = False
'		wait 1
'	Wend
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
	'wait 2
	Dim aa
		aa = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		While aa = 0
			wait 1
		aa = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").Click
'Selecciona Duos
	Case "Dúos"
	Call Carga()
		wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 6000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").Select DataTable("e_Categoria", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 4000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaEdit("Buscar por LOB").Set DataTable("e_Busca", dtLocalSheet)
	Call Carga()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Buscar").Click
	Call Carga()
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"PlanTarifario.png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() & Ciclo&"PlanTarifario.png"
	'Esperar a que elemento sea visible
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))   = False
		wait 1
	Wend
	'Valida si no hay ofertas
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		'wait 1
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("No existen ofertas elegibles").GetROProperty("text")
		DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)=varsap
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png", True
		imagenToWord "ErrorPlanTarifario", RutaEvidencias() &Num_Iter&"_"&"ErrorPlanTarifario"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Cerrar").Click
		wait 1
		ExitActionIteration
	End If
	Call Carga()
		wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Set "ON"
	wait 2
	Dim bb
		bb = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		While bb = 0
			wait 1
		bb = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").Click
'Selecciona Trio
	Case "Tríos"
	Call Carga()
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 6000
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").Select DataTable("e_Categoria", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 4000
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaEdit("Buscar por LOB").Set DataTable("e_Busca", dtLocalSheet)
	Call Carga()
		wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Buscar").Click
	Call Carga()
		wait 1
	'Imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"PlanTarifario.png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() & Ciclo&"PlanTarifario.png"
	'Espera a que elemento sea visible
	While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))   = False
		wait 1
	Wend
	'Valida si no hay odertas
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
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija").JavaButton("Cerrar").Click
		wait 1
		ExitActionIteration
	End If @@ hightlight id_;_27080509_;_script infofile_;_ZIP::ssf1.xml_;_
	Call Carga()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaCheckBox("Seleccionar_2").Set "ON"
	Call Carga()
	Dim e
		e = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
		While e = 0
			wait 1
		e = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nueva Oferta Linea Fija_11").JavaButton("Siguiente >").GetROProperty("enabled")
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
	
		'Selecciona producto de Mono TV
		Case "Producto Mono TV"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend
		Dim filasT
		filasT = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filasT >=0 Then
		dim  tv	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#17")		
		tv = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#17", "#2")	
		'Bloque FULL HD
'		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Exist = false
'			wait 1
'		Wend
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Exist = false
			wait 1
		Wend
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Set "Bloque Full HD" '"Bloque Estelar"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Set "ON"
		'Validacion cuando el bloque ya ha sido seleccionado
		If  JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaStaticText("BLOQUE FULL HD 60 CANALES").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png", True
		    imagenToWord "Mensaje de validación de canales", RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Cancelar").Click
			ExitActionIteration
		End If
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Agregar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_6").JavaButton("Validar").Click
		Call Carga()
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_6").JavaButton("Calcular").Click
		'wait 7
		Call Carga()
		wait 1
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
		imagenToWord "Recursos Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
		End if
		
		'Selecciona producto de Mono Internet
		'Prueba Internet naked
		Case "Producto Mono Internet"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#0"
		Dim filasT4
		filasT4 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filasT4 >=0 Then
		dim  duoII2	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#7")		
		duoII2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#7", "#2")	
		'wait 7
		'validacion -No disponible
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").Exist = True Then
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"NoDisponible.png", True
		imagenToWord "No disponible para la region", RutaEvidencias() & Ciclo&"NoDisponible.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitActionIteration
		End If
		'Error velocidad de internet
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
		imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
		ExitActionIteration
		End If
		'Esperar a que cargue la tabla
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		'Carga de velocidad de internet
		Dim tabb
		tabb = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		While tabb = 0 or tabb =""
			wait 1
		tabb = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").SelectRow "#1" @@ hightlight id_;_9881409_;_script infofile_;_ZIP::ssf333.xml_;_
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		Call Carga()
		wait 1
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad de Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaButton("Aceptar").Click @@ hightlight id_;_19697227_;_script infofile_;_ZIP::ssf404.xml_;_
		'wait 5
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		End if
		'Equipo
			Call Carga()
			'wait 1
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
			'MsgBox n2
			DataTable("s_NumeroAsignado", dtLocalSheet) = n6
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Guardar").Click @@ hightlight id_;_4265608_;_script infofile_;_ZIP::ssf154.xml_;_
			Call Carga()
			'wait 1
			'Imagen
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
			imagenToWord "Recursos Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
		
		
		'Selecciona producto de Duo TV + BB
		Case "Producto Duo TV y BB"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf140.xml_;_
		Dim filas
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filas >=0 Then
		dim  duos	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#7")
		duos = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#7", "#2")	
		'wait 7
		'validacion -No disponible
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").Exist = True Then
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"NoDisponible.png", True
		imagenToWord "No disponible para la region", RutaEvidencias() & Ciclo&"NoDisponible.png"
		'JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaStaticText("No hay velocidades disponibles").Click
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitActionIteration
		End If
		'Error velocidad de internet
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
		imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
		ExitActionIteration
		End If
		'Esperar a que cargue la tabla
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		'Carga de velocidad de internet
		Dim tab
		tab = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		While tab = 0 or tab =""
			wait 1
		tab = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").SelectRow "#1" @@ hightlight id_;_9881409_;_script infofile_;_ZIP::ssf333.xml_;_
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		Call Carga()
		wait 1
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad de Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaButton("Aceptar").Click @@ hightlight id_;_19697227_;_script infofile_;_ZIP::ssf404.xml_;_
		'wait 5
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		Call Carga()
		
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		'wait 1
		End if
			'Equipo
			Call Carga()
			'wait 1
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
			'wait 1
		'TV
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#2"
		wait 4
		Dim f11
		f11 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fi11 >=0 Then
		dim  tv4	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#9"	
		tv4 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#9", "#2")	
		'wait 5
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Exist = false
			wait 1
		Wend
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Set "Bloque Full HD" '"Bloque Estelar"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Set "ON"
		'Validacion cuando el bloque ya ha sido seleccionado
		If  JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaStaticText("BLOQUE FULL HD 60 CANALES").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png", True
		    imagenToWord "Mensaje de validación de canales", RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Cancelar").Click
			ExitActionIteration
		End If
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Agregar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 2
		'Valida mensaje 
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
            JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacion.png", True
		    imagenToWord "Mensaje de validación", RutaEvidencias() & Ciclo&"mensajeValidacion.png"
		    ExitActionIteration
		End If		
			'Imagen
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
			imagenToWord "Recursos Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
			Call Carga()
	End if
		
		
		'Selecciona producto de Duo TV + VOZ
		Case "Producto Duo TV y VOZ" 
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf140.xml_;_
		Dim fill1
		fill1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fill1 >=0 Then
		dim  du	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#2"	
		du = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#2", "#2")	
		'wait 7
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Set "ON"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaButton("Agregar").Click
		'wait 3
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 1
		End if
			'Equipo
			Call Carga()
			'wait 1
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
		'TV
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#2"
		Call Carga()
		wait 1
		Dim fil12
		fil12 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fil12 >=0 Then
		dim  tv11	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#17"	
		tv11 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#17", "#2")	
		'wait 5
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Exist = false
			wait 1
		Wend
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Set "Bloque Full HD" '"Bloque Estelar"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Set "ON"
		'Validacion cuando el bloque ya ha sido seleccionado
		If  JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaStaticText("BLOQUE FULL HD 60 CANALES").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png", True
		    imagenToWord "Mensaje de validación de canales", RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Cancelar").Click
			ExitActionIteration
		End If
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Agregar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Validar").Click
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 2
		'Valida mensaje 
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
            JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacion.png", True
		    imagenToWord "Mensaje de validación", RutaEvidencias() & Ciclo&"mensajeValidacion.png"
		    ExitActionIteration
		End If		
		'Telefono
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").Exist = false
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#3" @@ hightlight id_;_14896642_;_script infofile_;_ZIP::ssf390.xml_;_
		'wait 5
		Call Carga()
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTab("0.00").Select "Asignación de número"
		wait 3
		Dim dt
		dt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
		While dt = 0
			wait 1
		dt = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").Click
		'wait 4
		Dim de
		de = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
		While de = 0
			wait 1
		de = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").Click
		wait 4 @@ hightlight id_;_5880905_;_script infofile_;_ZIP::ssf282.xml_;_
		'Guardar numero @@ hightlight id_;_13261133_;_script infofile_;_ZIP::ssf400.xml_;_
		Dim n5
		n5 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("SearchJTable").GetCellData("#0", "#1")
		'MsgBox n3
		DataTable("s_NumeroAsignado", dtLocalSheet) = n5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Guardar").Click
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
		imagenToWord "Recurso Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
		End if
		
		
	'Selecciona producto de Duo BB + VOZ
		Case "Producto Duo BB y VOZ"
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
			wait 1
		Wend @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf140.xml_;_
		Dim fila2
		fila2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fila2 >=0 Then
		dim  duo2	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#2"	
		duo2 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#2", "#2")	
		'wait 7
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaCheckBox("Seleccionar").Set "ON"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_12").JavaButton("Agregar").Click
		'wait 3
		Call Carga()
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 1
		End if
		'Internet'
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#1"
		Dim filasT3
		filasT3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filasT3 >=0 Then
		dim  duoII	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#7")		
		duoII = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#7", "#2")	
		'wait 7
		'validacion -No disponible
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").Exist = True Then
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"NoDisponible.png", True
		imagenToWord "No disponible para la region", RutaEvidencias() & Ciclo&"NoDisponible.png"
		'JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaStaticText("No hay velocidades disponibles").Click
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitActionIteration
		End If
		'Error velocidad de internet
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
		imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
		ExitActionIteration
		End If
		'Esperar a que cargue la tabla
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		'Carga de velocidad de internet
		Dim tab10
		tab10 = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		While tab10 = 0 or tab10 =""
			wait 1
		tab10 = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").SelectRow "#1" @@ hightlight id_;_9881409_;_script infofile_;_ZIP::ssf333.xml_;_
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		Call Carga()
		wait 1
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad de Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaButton("Aceptar").Click
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		'wait 1
		End if
			'Equipo
			Call Carga()
			'wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#2" @@ hightlight id_;_70411_;_script infofile_;_ZIP::ssf136.xml_;_
			Call Carga()
			wait 1
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
			End If
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
			wait 1
			'Numero Telefono
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").Exist = false
			wait 1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#3" @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf150.xml_;_
			Call Carga()
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTab("0.00").Select "Asignación de número" @@ hightlight id_;_6191933_;_script infofile_;_ZIP::ssf151.xml_;_
			Call Carga()
			wait 1
			Dim phone
			phone = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
			While phone = 0
			wait 1
			phone = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").Click @@ hightlight id_;_24514005_;_script infofile_;_ZIP::ssf152.xml_;_
			'wait 4
			Dim phoneN
			phoneN = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
			While phoneN = 0
			wait 1
			phoneN = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").Click @@ hightlight id_;_25618982_;_script infofile_;_ZIP::ssf153.xml_;_
			Call Carga()
			wait 1
			'Guardar numero
			Dim n7
			n7 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("SearchJTable").GetCellData("#0", "#1")
			'MsgBox n2
			DataTable("s_NumeroAsignado", dtLocalSheet) = n6
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Guardar").Click @@ hightlight id_;_4265608_;_script infofile_;_ZIP::ssf154.xml_;_
			Call Carga()
			'wait 1
			'Imagen
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
			imagenToWord "Recursos Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
			Call Carga()
	
	
		'Selecciona producto trio	
		Case "Producto Trio"
			'Accesorios
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").Exist = false
		wait 1
		Wend
		Dim filaT3
		filaT3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If filaT3 >=0 Then
		dim  trio3	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow ("#2")		
		trio3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#2", "#2")	
		'wait 7
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_10").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_10").JavaCheckBox("Seleccionar").Set "ON"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_10").JavaButton("Agregar").Click
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		'wait 7
		Call Carga()
		wait 2
		End If
		'Internet
		wait 2			 
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#1" @@ hightlight id_;_19794266_;_script infofile_;_ZIP::ssf335.xml_;_
		Dim fi
		fi = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fi >=0 Then
		dim  trioI @@ hightlight id_;_22840093_;_script infofile_;_ZIP::ssf260.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#7" @@ hightlight id_;_31605953_;_script infofile_;_ZIP::ssf264.xml_;_
		trioI = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#7", "#2")	
		'wait 7
		'validacion -No disponible
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").Exist = True Then
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"NoDisponible.png", True
		imagenToWord "No disponible para la region", RutaEvidencias() & Ciclo&"NoDisponible.png"
		'JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaStaticText("No hay velocidades disponibles").Click
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaDialog("Mensaje").JavaButton("OK").Click
		ExitActionIteration
		End If
		'velocidad internet
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").Exist = false
			wait 1
		Wend
		'Carga de velocidad de internet
		Dim tab2
		tab2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		While tab2 = 0 or tab2 =""
			wait 1
		tab2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").GetROProperty("rows")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaTable("SearchJTable").SelectRow "#1" @@ hightlight id_;_9881409_;_script infofile_;_ZIP::ssf333.xml_;_
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"Velocidad_Internet.png", True
		imagenToWord "Velocidad Internet", RutaEvidencias() & Ciclo&"Velocidad_Internet.png"
		'JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaButton("Aceptar").Click
		Call Carga()
		'wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_11").JavaButton("Aceptar").Click
		'Error velocidad de internet
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf286.xml_;_
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"ErrorInterno.png", True
		imagenToWord "Error Interno", RutaEvidencias() & Ciclo&"ErrorInterno.png"
		ExitActionIteration
		End If
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 2
		End If
		'Equipo
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#2"
		wait 3 @@ hightlight id_;_25280307_;_script infofile_;_ZIP::ssf232.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 2
		'TV
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#3"
		wait 4
		Dim fil
		fil = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").GetROProperty("rows")	
		If fil >=0 Then
		dim  tv1	    
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").SelectRow "#9"		
		tv1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar atributos:").DoubleClickCell("#9", "#2")	
		'wait 5
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Exist = false
			wait 1
		Wend
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaEdit("TextFieldNative$1").Set "Bloque Full HD" '"Bloque Estelar"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Buscar").Click
		While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Exist = false
			wait 1
		Wend
		Call Carga()
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaCheckBox("Seleccionar").Set "ON"
		'Validacion cuando el bloque ya ha sido seleccionado
		If  JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaStaticText("BLOQUE FULL HD 60 CANALES").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaDialog("Mensaje").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png", True
		    imagenToWord "Mensaje de validación de canales", RutaEvidencias() & Ciclo&"mensajeValidacionCanales.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Cancelar").Click
			ExitActionIteration
		End If
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración_5").JavaButton("Agregar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Calcular").Click
		Call Carga()
		wait 2
		'Valida mensaje 
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
            JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"mensajeValidacion.png", True
		    imagenToWord "Mensaje de validación", RutaEvidencias() & Ciclo&"mensajeValidacion.png"
		    ExitActionIteration
		End If		
		'Telefono
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").Exist = false
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("Mostrar:").SelectRow "#4"
		Call Carga()
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTab("0.00").Select "Asignación de número"
		wait 3
		Dim d
		d = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
		While d = 0
			wait 1
		d = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Proponer números").Click
		Call Carga()
		Dim dn1
		dn1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
		While dn1 = 0
			wait 1
		dn1 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").GetROProperty("enabled")
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Distribuir Número").Click
		wait 4 @@ hightlight id_;_5880905_;_script infofile_;_ZIP::ssf282.xml_;_
		'Guardar numero @@ hightlight id_;_13261133_;_script infofile_;_ZIP::ssf400.xml_;_
		Dim n3
		n3 = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaTable("SearchJTable").GetCellData("#0", "#1")
		'MsgBox n3
		DataTable("s_NumeroAsignado", dtLocalSheet) = n3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración_9").JavaButton("Guardar").Click
		'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & Ciclo&"RecursosAlta.png", True
		imagenToWord "Recurso Alta", RutaEvidencias() & Ciclo&"RecursosAlta.png"
		'Cuando no carga
		End if
		End  Select
		
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
		'End If	
	'End	Select
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


		'Cuando falla la wic2
'	Dim x1
'	x1 = "¡RECORDATORIO! GUARDAR"
'	MsgBox x1
		''flujo para guardar
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Guardar").Click @@ hightlight id_;_18537769_;_script infofile_;_ZIP::ssf531.xml_;_
'	Call Carga()
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Close @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf532.xml_;_
'	JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Guardar").Click
'	Call Carga()
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
