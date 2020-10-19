Option Explicit

Dim str_tipoDocumento
Dim str_identificadorPersona
Dim str_departamento
Dim str_provincia
Dim str_distrito
Dim str_nombreVia
Dim str_numeroCasa
Dim str_tipoCliente
Dim str_SubTipoCliente
Dim str_Nacionalidad

str_tipoDocumento		 =	DataTable("e_TipoDocumento", dtLocalSheet)
str_identificadorPersona =	DataTable("e_IdentificacionPersona", dtLocalSheet)
str_departamento 		 =	DataTable("e_Departamento", dtLocalSheet)
str_provincia			 =	DataTable("e_Provincia", dtLocalSheet)
str_distrito 			 =	DataTable("e_Distrito", dtLocalSheet)
str_nombreVia 			 =	DataTable("e_NombreVia", dtLocalSheet)
str_numeroCasa			 =	DataTable("e_NumeroCasa", dtLocalSheet)
str_tipoCliente			 =	DataTable("e_TipoCliente", dtLocalSheet)
str_SubTipoCliente		 =	DataTable("e_SubTipoCliente", dtLocalSheet)
str_Nacionalidad		 =	DataTable("e_Nacionalidad", dtLocalSheet)


Call SeleccionarDetalleContacto()
Call IngresaDireccion()
Call DetalleCliente()

Sub SeleccionarDetalleContacto()
While JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").Exist = False
	wait 1
Wend
JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").Select 
JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Cliente").Select 
'Selecciona datos
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaList("Tipo de documento").Exist = False
	wait 1
Wend
Call Carga()
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaList("Tipo de documento").Select str_tipoDocumento
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente").JavaEdit("Identificación de la persona").Set str_identificadorPersona

'Si se habilita campo de Nacionalidad
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").GetROProperty("enabled") <> "0" Then
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Nacionalidad:").Select str_Nacionalidad
	End If
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar").Click
wait 1
'Si el cliente ya ha sido creado @@ hightlight id_;_31560954_;_script infofile_;_ZIP::ssf36.xml_;_
If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist = True Then
	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click @@ hightlight id_;_8788895_;_script infofile_;_ZIP::ssf27.xml_;_
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"ClienteYaExiste.png", True
	imagenToWord "ClienteYaExiste", RutaEvidencias() &"ClienteYaExiste.png"
	'Cerrar ventana
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Cerrar").Click
	wait 1
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").Exist = True Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Guardar el formulario").JavaButton("Descartar").Click
	End If
	ExitActionIteration
End If
Wait 2
If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist = True Then
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Problema.png", True
	imagenToWord "Problema", RutaEvidencias() &"Problema.png"
	ExitTest
End If
'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"DetalleContacto.png", True
		imagenToWord "DetalleContacto", RutaEvidencias() &"DetalleContacto.png"
End Sub

Sub Carga()
''Metodo para cargar elementos	
RunAction "CargaElemento", oneIteration
End Sub


''''''''' @@ hightlight id_;_11055441_;_script infofile_;_ZIP::ssf12.xml_;_
Sub IngresaDireccion()
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Exist = false
		wait 1	
	Wend
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Departamento:").Select str_departamento @@ script infofile_;_ZIP::ssf20.xml_;_
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Provincia:").Set str_provincia
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated").Click
Call Carga()
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Distrito:").Set str_distrito
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Lookup-notValidated_2").Click
Call Carga()
wait 1
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Nombre de Vía:").Set str_nombreVia
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaEdit("Número:").Set str_numeroCasa
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Validar_2").Click
Call Carga()
wait 1
'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"IngresaDireccion.png", True
		imagenToWord "IngresaDireccion", RutaEvidencias() &"IngresaDireccion.png"
END SUB		 
			 

Sub DetalleCliente()
While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Cliente:").Exist = False
	wait 1
Wend
Call Carga()
wait 1
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Tipo de Cliente:").Select str_tipoCliente
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaList("Subtipo de Cliente:").Select str_SubTipoCliente
JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Dar de alta al cliente*").JavaButton("Guardar").Click
'Imagen
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"DetalleCliente.png", True
		imagenToWord "DetalleCliente", RutaEvidencias() &"DetalleCliente.png"
Call Carga()
		
END SUB

