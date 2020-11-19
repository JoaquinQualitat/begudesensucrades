Dim resultado
 @@ hightlight id_;_Browser("B").Page("Entrada al sistema 2").SAPEdit("Sistema:")_;_script infofile_;_ZIP::ssf45.xml_;_

'Login
Call LoginSAPWeb ()
 @@ hightlight id_;_Browser("IEinicio").Page("Entrada al sistema").SAPButton("Acceder al sistema")_;_script infofile_;_ZIP::ssf28.xml_;_
With Browser("B").Page("P")
	'.SAPNWBC("nav_1").Navigate "Base dades" + vbCrLf + "fiscals" @@ hightlight id_;_Browser("IEinicio").Page("Actualització especificacions").SAPNWBC("nav 1")_;_script infofile_;_ZIP::ssf29.xml_;_
	
	With .Frame("Frame")
		If Browser("B").Page("P").WebElement("Base dadesfiscals").Exist Then
			Do
				Browser("B").Page("P").WebElement("WebElement").Click
			Loop While Browser("B").Page("P").WebElement("Base dadesfiscals").getROProperty("Visible") = False
			
				
		'	Loop While Browser("B").Page("P").WebElement("Base dadesfiscals").Visible = False
			Browser("B").Page("P").WebElement("Base dadesfiscals").Click

		End If
		.Link("Cerca i alta de persones").Click @@ hightlight id_;_Browser("IEinicio").Page("Base dades fiscals").Frame("iFrameId 1558023688839").Link("Cerca i alta de persones")_;_script infofile_;_ZIP::ssf30.xml_;_
		
		' Busqueda del NIF 12345678Z
		resultado = BuscarNif("12345678Z")
		Print resultado
		If resultado = 3 Then
			Reporter.ReportEvent micPass, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 12345678Z."
		Else
			Reporter.ReportEvent micFail, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 12345678Z." &_
														   vbCrLf & "S'esperaven 3 resultats."
		End If
		
		' Busqueda del NIF 99900000Y
		resultado = BuscarNif("99900000Y")
		Print resultado
		If resultado = 1 Then
			Reporter.ReportEvent micPass, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 99900000Y."
		Else
			Reporter.ReportEvent micFail, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 99900000Y." &_
														   vbCrLf & "S'esperava 1 resultat."
		End If
		
		' Busqueda del NIF 99900000Z
		resultado = BuscarNif("99900000Z")
		Print resultado
		If resultado = 0 Then
			Reporter.ReportEvent micPass, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 99900000Z."
		Else
			Reporter.ReportEvent micFail, "Búsqueda NIF", "S'han trobat " & resultado & " coincidències amb el NIF 99900000Z." &_
														   vbCrLf & "S'esperaven 0 resultats."
		End If
		
	End With
End With @@ hightlight id_;_Browser("IEinicio").Page("Cerca i alta de persones").Frame("iFrameId 1558023692205").SAPButton("Cercar")_;_script infofile_;_ZIP::ssf35.xml_;_

Browser("B").Page("Cerca i alta de persones").Link("Finalitzar la sessió").Click @@ hightlight id_;_Browser("B").Page("Cerca i alta de persones").Link("Finalitzar la sessió")_;_script infofile_;_ZIP::ssf42.xml_;_
Browser("B").Page("Cerca i alta de persones").Frame("iframeLogoffMsgDialog").WebElement("D'acord").Click @@ hightlight id_;_Browser("B").Page("Cerca i alta de persones").Frame("iframeLogoffMsgDialog").WebElement("D'acord")_;_script infofile_;_ZIP::ssf43.xml_;_
Browser("B").Page("Finalitzar la sessió").Sync @@ hightlight id_;_Browser("B").Page("Finalitzar la sessió")_;_script infofile_;_ZIP::ssf44.xml_;_
Browser("B").CloseAllTabs


' FUNCIONES

'@Description	Inicia sesión en la base de Base de Datos Fiscales. Abre el explorador si no hay ninguno abierto. Si ya hay
'				una sesión iniciada solo lleva a la página de inicio.
'@Documentation	Inicia sesión en la base de Base de Datos Fiscales. Abre el explorador si no hay ninguno abierto. Si ya hay
'				una sesión iniciada solo lleva a la página de inicio.
Sub Login ()
	With Browser("B")
		' Si no existe el navegador, abre uno nuevo (Internet Explorer)
		If not .Exist(0) Then
		'	SystemUtil.Run Environment ("BrowserIE"), Environment (Environment ("Entorn"))
			SystemUtil.Run Environment ("BrowserIE"), Environment ("BDF_" & (Environment ("Entorn")))
		End If
	
		'Selecciona el entorno y accede a la URL correspondiente
		'.Navigate Environment (Environment ("Entorn"))
		.Navigate Environment ("BDF_" & (Environment ("Entorn")))
		
		'Login
		With .Page ("Entrada al sistema")
			.Sync
			'Solo se pasa por la pantalla de Login si no hay una sesión activa
			'If .SAPEdit("Usuario").Exist(1) Then
				.SAPEdit("Usuario").Set "OQSP"    'pass: Init.123a @@ hightlight id_;_Browser("B").Page("Entrada al sistema").SAPEdit("Clave de acceso: *")_;_script infofile_;_ZIP::ssf36.xml_;_
				'.SAPEdit("Clave de acceso").SetSecure "5dbac3360fb0546db73af8b37aae8169bc9aa224a0c715b750315e67" 
				.SAPEdit("Clave de acceso").SetSecure "5dbac3360fb0546db73af8b37aae8169bc9aa224a0c715b750315e67" 
				
'				SelectByIndex .SAPList("Idioma"), 2
				SelectByValue .SAPList("Idioma"), "CA"

				.SAPButton("Acceder al sistema").Click
			'End If
		End With
		
		.Page ("micclass:=Page").Sync
	End With
End Sub

'@Description	Busca el NIF indicado en la Base de Datos Fiscales. Devuelve el número de resultados encontrados.
'@Documentation	Busca el NIF <nif> en la Base de Datos Fiscales. Devuelve el número de resultados encontrados.
Function BuscarNIF(ByVal nif)
'<nif>:		NIF a buscar en la Base de Datos Fiscales.
'return:	Número de resultados encontrados en la base de datos.

	With Browser("B").Page("P").Frame("Frame")
		' Introduce el NIF
		.SAPEdit("SAPEdit").Set nif
		' Lanza la búsqueda
		.SAPButton("Cercar").Click @@ hightlight id_;_Browser("IEinicio").Page("Cerca i alta de persones").Frame("iFrameId 1558023692205").SAPEdit("SAPEdit")_;_script infofile_;_ZIP::ssf31.xml_;_
		Wait 5
		
		' Devuelve el número de resultados
		BuscarNIF = CInt(Split(.WebElement("Resultat de cerca").GetROProperty("innertext"), " ", 5)(3))
	End With
End Function

'@Description	Selecciona un elemento del SAPList mediante su índice.
'@Documentation	Selecciona un elemento del <sap_list> mediante su índice <index>.
Sub SelectByIndex(ByRef sap_list, ByVal index)
'<sap_list>:	SAPList sobre el que se realiza la acción.
'<index>:		Índice del elemento a seleccionar.

	' Se hace focus sobre la SAPList
	sap_list.Object.Focus
	' Se acciona el evento de "Click" de la SAPList
	sap_list.FireEvent "click"
	
	' Se selecciona el elemento con el índice <index>
	Browser("B").Page("Entrada al sistema").WebElement("html id:=SL1-key-" & index).FireEvent "click"
End Sub

'@Description	Selecciona un elemento del SAPList por su valor.
'@Documentation	Selecciona un elemento del <sap_list> por su valor <value>.
Sub SelectByValue(ByRef sap_list, ByVal value)
'<sap_list>:	SAPList sobre el que se realiza la acción.
'<value>:		Valor del elemento a seleccionar.

	Dim elem, i
	
	' Se hace focus sobre la SAPList
	sap_list.Object.Focus
	' Se acciona el evento de "Click" de la SAPList
	sap_list.FireEvent "click"
	
	With Browser("B").Page("Entrada al sistema")
		' Se genera una lista con todos los elementos de la SAPList
		elem = Split(Trim(.WebTable("html id:=SL1-tab").GetROProperty("text")))
		
		' Por cada elemento de la SAPList
		For i = 0 To UBound(elem)
			' Si el valor buscado está en el elemento
			If InStr(elem(i), value) Then
				' Se selecciona el elemento
				.WebElement("html id:=SL1-key-" & (i+1)).Click
				Exit For
			End If
		Next
	End With
	Wait 1
End Sub

