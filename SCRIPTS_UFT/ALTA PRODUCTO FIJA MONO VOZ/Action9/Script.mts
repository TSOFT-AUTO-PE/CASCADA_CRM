

If ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "POSTPAGO" Then
	Call Potspago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PREPAGO" Then
	Call Prepago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PORTABILIDAD" Then
	Call Portabilidad()
End If



Sub Potspago()
While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
wait 6
 Window("Ejecutivo de interacción").InsightObject("InsightObject").Click @@ hightlight id_;_421_;_script infofile_;_ZIP::ssf1.xml_;_
 While Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = false
 	wait 1
 Wend
 wait 1
 Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click @@ hightlight id_;_423_;_script infofile_;_ZIP::ssf2.xml_;_
While  (Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Exist or Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Exist) = false
	wait 1
Wend
wait 1
If Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Exist = true Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
	imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
	ExitActionIteration
End If
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ScoreCalculado.png", True
imagenToWord "Score Calculado", RutaEvidencias() & "ScoreCalculado.png"
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN 2}"
If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = True Then
	Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click @@ hightlight id_;_443_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 3
	'else
	'wait 3
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN 2}"

'Continuar con el flujo
if Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = false Then
	wait 1
End if
wait 1
If Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Exist = true Then
	Call Validacion()
End If
Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
wait 5
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
'Validacion 
If Window("Ejecutivo de interacción").InsightObject("InsightObject_00").Exist = true Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
	imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
	ExitActionIteration
End If
	
Else
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist = True Then
	Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
	End If
End If
'wait 3
'Set shell = CreateObject("Wscript.Shell") 
'shell.SendKeys "{PGDN 2}"
'
''Continuar con el flujo
'if Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = false Then
'	wait 1
'End if
'wait 1
'If Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Exist = true Then
'	Call Validacion()
'End If
'Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
'wait 5
'Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
'imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
'Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
''Validacion 
'If Window("Ejecutivo de interacción").InsightObject("InsightObject_00").Exist = true Then
'	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
'	imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
'	ExitActionIteration
'End If

End Sub

Sub Validacion()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Click
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
	wait 1
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "prueba"
	wait 1
	Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Type micCtrlDwn + micAltDwn + "q" + micCtrlUp + micAltUp
	wait 2
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "gmail.com"
	wait 1

End Sub

