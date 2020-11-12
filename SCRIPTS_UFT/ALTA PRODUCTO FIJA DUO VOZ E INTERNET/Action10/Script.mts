
While Window("Ejecutivo de interacción").InsightObject("InsightObject_01").Exist = false
	wait 1
Wend

If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = True Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Incidencia.png", True
	imagenToWord "Aviso: Actualmente mos encontramos en incidencia", RutaEvidencias() & "Incidencia.png"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
Else
	Window("Ejecutivo de interacción").InsightObject("InsightObject_01").Click
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratacionValidacion.png", True
	imagenToWord "Contratación y Validación de Identidad", RutaEvidencias() & "ContratacionValidacion.png"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_02").Exist = false
		wait 1
	Wend
	'imagen
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "noBiometrica.png", True
	imagenToWord "Validación no biometrica", RutaEvidencias() & "noBiometrica.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_02").Click
	'Validacion - Ultimo cuestionario mal
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Exist = True Then
		Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
		imagenToWord "Se produjo un error", RutaEvidencias() & "ErrorServicio.png"
	ExitActionIteration
	End If
	'____________________________________
	'Flujo Manual
	Dim x
	x = "Realice los 3 pasos manualmente, hasta antes de PULSAR el botón 'CONTINUAR' (CLICK DESPUES DE LOS 3 PASOS)"
	MsgBox x
	'____________________________________
	
	'Continua el flujo despues de los 3 pasos
	
'	'Validación
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_04").Exist = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "CuestionarioAprobado.png", True
	imagenToWord "Cuestionario Aprobado", RutaEvidencias() & "CuestionarioAprobado.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_04").Click
	wait 2
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_05").Exist = false
		wait 1
	Wend
    Window("Ejecutivo de interacción").InsightObject("InsightObject_05").Click
    
	'Prueba
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	'Descargar contratos
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_08").Exist = false
	wait 1
	Dim b
		b=b+1
			If b>=120 Then
				Call NextStep()
			End If	
	Wend
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_08").Click
	wait 2
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_013").Exist = false
	wait 1
	Wend
	Window("Ejecutivo de interacción").InsightObject("InsightObject_013").Click
	
	'Declaracion Jurada
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_015").Exist = false
	wait 1
	Wend
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_015").Click
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_011").Click
	
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_016").Exist = True Then
		'Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click
		Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
		imagenToWord "Se produjo un error", RutaEvidencias() & "ErrorServicio.png"
	ExitActionIteration
	End If
	
	'Validacion final
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_003").Exist = False
		wait 1
	Wend
	
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_003").Exist = True Then
	Window("Ejecutivo de interacción").InsightObject("InsightObject_003").Click
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "validacion.png", True
	imagenToWord "Validacion y Descarga de documentos completa", RutaEvidencias() & "validacion.png"
	wait 3
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_002").Click
	
End If



