' PARAMETROS
Dim datos, nombre, tipo_via, nombre_via, numero, cod_postal, provincia, municipio, num_justificante, importe

' datos(0): NIF
' datos(1): Número de solicitud
datos = Split(readAllParameterFile(Parameter("Param_File")))

nombre = "Creada Deprova Persona"
tipo_via = "AVINGUDA"
nombre_via = "GENERALITAT"
numero = 1
cod_postal = "08630"
provincia = "BARCELONA"
municipio = "ABRERA"


' MAIN
' Acceso a Portal

IniciaPortalURL URL_BEGUDES_ENSUCRADES


'Login

IniciarSesion
TerminarCarga

With Browser("B").Page("P")
	' Introducir los datos del Sujeto Pasivo
	.WebEdit("NIF").Set datos(0)
	.WebEdit("Cognoms i nom").Set nombre
	.WebList("Tipus de via").Select tipo_via
	.WebEdit("Nom de la via").Set nombre_via
	.WebEdit("Núm").Set numero
	.WebEdit("Codi postal").Set cod_postal
	.WebList("Província").Select provincia
	Wait 2
	.WebList("Municipi").Select municipio
	Wait 2
	
	' Completar la autoliquidación
	.WebButton("Emplenar autoliquidació").Click
	
	' Elegir trimestre
	.WebList("Trimestre").Select 1

	' Si aparece un pop-up advirtiendo de que estamos fuera de plazo, continúa
	If .WebElement("Període fora de termini").Exist(5) Then
		.WebButton("Continuar").Click
	End If
	
	' Base imponible para cada tipo de bebida azucarada
	.WebEdit("BaseImposableA").Set 100		' > 8 gramos de azucar
	.WebEdit("BaseImposableB").Set 200		' Entre 5 y 8 gramos
	
	' Pagar por transferencia de Cuenta Corriente
	.WebRadioGroup("Tramitació").Select "P"
	
	' Seleccionar Pagador
	.WebList("Pagador").Select 1

	RellenarCuenta
	
	' Confirmar y avanzar
	.WebButton("Validar").Click
	.WebButton("Continuar_2").Click
	.WebButton("Signar, pagar i presentar").WaitProperty "visible", true, 45000
	.WebButton("Signar, pagar i presentar").Click
	
	' Espera a que cargue la página
	.WebElement("Núm. de justificant").WaitProperty "visible", true, 45000
	
	' Validaciones
	Validar .WebElement("NIF"), "NIF", datos(0)
	Validar .WebElement("Cognoms i nom"), "Cognoms i nom", nombre
	ValidarCalle
	Validar .WebElement("Codi postal"), "Codi postal", cod_postal
	Validar .WebElement("Província"), "Província", provincia
	Validar .WebElement("Municipi"), "Municipi", municipio
	ValidarAutoliquidacion
	
	' Guardar el número de justificante y el importe
	num_justificante = Split(.WebElement("Núm. de justificant").GetROProperty("innertext"))(3)
	importe = .WebTable("Autoliquidació").GetCellData(10, 4)
	writeParameterFile "Begudes_Ensucrades_01-01.txt", num_justificante & " " & importe
End With

'@Description	Valida que un campo tiene el texto esperado.
'@Documentation	Valida que un campo tiene el texto esperado.
Sub Validar(ByRef obj, ByVal name, ByVal text)
'<obj>:		Objeto a validar
'<name>:	Título de la validación.
'<text>:	String con el texto esperado en el objeto.

	Reporter.ReportEvent IIf(obj.GetROProperty("innertext") = text, micPass, micFail), name, "Valor esperat: " & text & vbCrLf & "Valor trobat: " & obj.GetROProperty("innertext")
End Sub

'@Description	Valida que la información de la dirección del Sujeto Pasivo sean las introducidas.
'@Documentation	Valida que la información de la dirección del Sujeto Pasivo sean las introducidas.
Sub ValidarCalle()
	Dim text, value, temp
	text = Trim(Browser("B").Page("P").WebElement("Adreça").GetROProperty("innertext"))
	
	' Tipo de vía
	value = Trim(Split(text)(0))
	Reporter.ReportEvent IIf(value = tipo_via, micPass, micFail), "Tipus de via", "Valor esperat: " & tipo_via & vbCrLf & "Valor trobat: " & value
	
	' Nombre de la vía
	value = Trim(Split(Split(text, ",")(0), " ", 2)(1))
	Reporter.ReportEvent IIf(value = nombre_via, micPass, micFail), "Nom de la via", "Valor esperat: " & nombre_via & vbCrLf & "Valor trobat: " & value
	
	' Número
	temp = Split(text)
	value = CInt(temp(Ubound(temp)))
	Reporter.ReportEvent IIf(value = numero, micPass, micFail), "Número de la via", "Valor esperat: " & numero & vbCrLf & "Valor trobat: " & value
End Sub

'@Description	Valida que los datos de la autoliquidación sean los introducidos.
'@Documentation	Valida que los datos de la autoliquidación sean los introducidos.
Sub ValidarAutoliquidacion()
	Dim value
	With Browser("B").Page("P")
		' Tipo A (> 8 gramos de azucar)
		value = .WebTable("Autoliquidació").GetCellData(2, 4)
		Reporter.ReportEvent IIf(value = 100, micPass, micFail), "Base imposable 1a", "Valor esperat: " & 100 & vbCrLf & "Valor trobat: " & value
		
		' Tipo B (5-8 gramos de azucar)
		value = .WebTable("Autoliquidació").GetCellData(2, 8)
		Reporter.ReportEvent IIf(value = 200, micPass, micFail), "Base imposable 1b", "Valor esperat: " & 200 & vbCrLf & "Valor trobat: " & value
	End With
End Sub

