' PARAMETROS
Dim datos, nif, num_sol, num_liq, importe

' NIF y número de la solicitud generada en [BEGUDES ENSUCRADES]-00-02
' - datos(0): NIF
' - datos(1): Número de solicitud
datos = Split(readAllParameterFile(Parameter("fileName-00-02")))
nif = datos(0)
num_sol = datos(1)

' Número de la liquidación e importe subidos en [BEGUDES ENSUCRADES]-01-01
' - datos(0): Número de liquidación
' - datos(1): Importe
datos = Split(readParameterFile(Parameter("fileName-01-01")))
num_liq = datos(0)
importe = datos(1)


' PROGRAMA

' Accede a Espriu en PRE
Dialog("SAP Logon 740").WinListView("SysListView32").Activate "Espriu-PRE"

' Login
Login

With SAPGuiSession("Session").SAPGuiWindow("SAP")
	' Ejecuta la transacción SE16
	.SAPGuiOKCode("OKCode").Set "SE16"
	.SAPGuiButton("Continuar").Click
	
	' Informa la tabla DFMCA_RETURN
	.SAPGuiEdit("Tabla").Set "DFMCA_RETURN"
	.SAPGuiButton("Continuar").Click
	
	' Informa el número de liquidación
	.SAPGuiEdit("Justificant autoliq.").Set num_liq
	.SAPGuiButton("Executar").Click
	
	' Visualiza el detalle del registro encontrado
	.SAPGuiCheckBox("SAPGuiCheckBox").Set "ON"
	.SAPGuiButton("Visualitzar").Click
	
	' Valida los datos del registro: NIF, periodo, importe y modelo
	Validar .SAPGuiEdit("(TAXPAYER ID)"), "NIF", nif
	Validar .SAPGuiEdit("(PERIOD KEY)"), "Periode", ObtenerPeriodo()
	Validar .SAPGuiEdit("Import"), "Import", importe
	Validar .SAPGuiEdit("Model"), "Model", "520"	

	' Salir
	.Close

End With

SAPGuiSession("Session").SAPGuiWindow("Sortir del sistema").SAPGuiButton("Sí").Click
Dialog("SAP Logon 740").Close


' FUNCIONES

'@Description	Inicia sesión en Espriu con los datos pasados a la prueba mediante los parámetros "usuari" y "contrasenya".
'@Documentation	Inicia sesión en Espriu con los datos pasados a la prueba mediante los parámetros "usuari" y "contrasenya".
Sub Login()
	With SAPGuiSession("Session").SAPGuiWindow("SAP")
		' A veces no identifica el campo del usuario a la primera
		If not .SAPGuiEdit("Usuari").Exist(2) Then		' Si no encuentra el campo del usuario
			.SAPGuiEdit("Usuari").RefreshObject			' Refresca el objeto
		End If
		
		' Rellena el campo del usuario
		.SAPGuiEdit("Usuari").SetFocus
		.SAPGuiEdit("Usuari").Set Parameter("usuari")
		
		' Rellena el campo de la contraseña
		.SAPGuiEdit("Contras.").SetFocus
		.SAPGuiEdit("Contras.").Set Parameter("contrasenya")
		
		' Accede al sistema
		.SAPGuiButton("Continuar").Click
	End With
End Sub

'@Description	Valida que un campo tiene el texto esperado.
'@Documentation	Valida que un campo tiene el texto esperado.
Sub Validar(ByRef obj, ByVal name, ByVal text)
'<obj>:		Objeto a validar
'<name>:	Título de la validación.
'<text>:	String con el texto esperado en el objeto.

	Reporter.ReportEvent IIf(obj.GetROProperty("value") = text, micPass, micFail), name, "Valor esperat: " & text & vbCrLf & "Valor trobat: " & obj.GetROProperty("value")
End Sub

'@Description	Genera el código del periodo previo al trimestre actual.
'@Documentation	Genera el código del periodo previo al trimestre actual.
Function ObtenerPeriodo()
'return:	Código del trimestre previo al periodo actual. Formato <aa>T<t>, donde <aa> son las 2 últimas cifras del año y <t> es el número del trimestre.

	Dim fecha : fecha = DateAdd("m", -3, Date)		' Fecha del trimestre anterior
	
	ObtenerPeriodo = CStr(Year(fecha) mod 100) & "T" & CStr(Fix((Month(fecha) - 1)/12) + 1)
End Function


