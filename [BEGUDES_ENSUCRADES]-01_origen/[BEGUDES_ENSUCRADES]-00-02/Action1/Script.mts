' PARAMETROS

' Datos de la persona
Dim nombre, apellido_1, apellido_2, nif

nombre = "Persona"
apellido_1 = "Creada"
apellido_2 = "Deprova"
nif = GenerarNIF

' Dato de la dirección
Dim direccion, pais, provincia, municipio, via, calle, cod_postal

direccion = "Residència habitual"
pais = "Espanya"
provincia = "Barcelona"
municipio = "Abrera"
via = "Avinguda"
calle = "GENERALITAT"
cod_postal = "08630"


' TEST

' Login
'Login
Call LoginSAPWeb @@ hightlight id_;_Browser("IEinicio").Page("Entrada al sistema").SAPButton("Acceder al sistema")_;_script infofile_;_ZIP::ssf28.xml_;_
With Browser("B").Page("P")
	' Navega a la Base de Datos Fiscales
	.SAPNWBC("nav_1").Navigate "Base dades" + vbCrLf + "fiscals" @@ hightlight id_;_Browser("IEinicio").Page("Actualització especificacions").SAPNWBC("nav 1")_;_script infofile_;_ZIP::ssf29.xml_;_
	
	With .Frame("Frame")
		' Accede a la página de Búsqueda y Alta de Personas
		.Link("Cerca i alta de persones").Click @@ hightlight id_;_Browser("IEinicio").Page("Base dades fiscals").Frame("iFrameId 1558023688839").Link("Cerca i alta de persones")_;_script infofile_;_ZIP::ssf30.xml_;_
		
		' Crea una nueva entrada de persona
		.SAPButton("Nou").Click
		CrearPersona
	End With
End With @@ hightlight id_;_Browser("IEinicio").Page("Cerca i alta de persones").Frame("iFrameId 1558023692205").SAPButton("Cercar")_;_script infofile_;_ZIP::ssf35.xml_;_


' FUNCIONES

'@Description	Inicia sesión en la base de Base de Datos Fiscales. Abre el explorador si no hay ninguno abierto. Si ya hay
'				una sesión iniciada solo lleva a la página de inicio.
'@Documentation	Inicia sesión en la base de Base de Datos Fiscales. Abre el explorador si no hay ninguno abierto. Si ya hay
'				una sesión iniciada solo lleva a la página de inicio.
Sub Login ()
	With Browser("B")
		' Si no existe el navegador, abre uno nuevo (Internet Explorer)
		If not .Exist(0) Then
			SystemUtil.Run Environment ("BrowserIE"), Environment ("BDF_" & (Environment ("Entorn")))
			'SystemUtil.Run Environment ("BrowserIE"), Environment (Environment ("Entorn"))
		End If
	
		'Selecciona el entorno y accede a la URL correspondiente
		.Navigate Environment ("BDF_" & (Environment ("Entorn")))
		'.Navigate Environment (Environment ("Entorn"))
		
		'Login
		With .Page ("Entrada al sistema")
			.Sync
			'Solo se pasa por la pantalla de Login si no hay una sesión activa
			If .SAPEdit("Usuario").Exist(1) Then
				'.SAPEdit("Usuario").Set "52010432A" @@ hightlight id_;_Browser("IEinicio").Page("Entrada al sistema").SAPEdit("Usuario: *")_;_script infofile_;_ZIP::ssf24.xml_;_
				.SAPEdit("Usuario").Set "OQSP"
				'.SAPEdit("Clave de acceso").SetSecure "5dbac3360fb0546db73af8b37aae8169bc9aa224a0c715b750315e67"
				.SAPEdit("Clave de acceso").SetSecure "5ed782c1a6c3eb222870d619783d034e7201ebc9217eb1c7"
'				SelectByIndex .SAPList("Idioma"), 2
				SelectByValue .SAPList("Idioma"), "CA" @@ hightlight id_;_Browser("B").Page("Entrada al sistema").SAPList("Idioma:")_;_script infofile_;_ZIP::ssf45.xml_;_
 @@ hightlight id_;_Browser("B").Page("Entrada al sistema").SAPButton("Catalán - 3 de 5 elementos")_;_script infofile_;_ZIP::ssf50.xml_;_
				.SAPButton("Acceder al sistema").Click
			End If
		End With
		
		.Page ("micclass:=Page").Sync
	End With
End Sub

'@Description	Crea una nueva persona en la Base de Datos Fiscales rellenando y validando todos los campos obligatorios.
'				Toma los valores de los campos de las variables globales.
'@Documentation	Crea una nueva persona en la Base de Datos Fiscales rellenando y validando todos los campos obligatorios.
'				Toma los valores de los campos de las variables globales:
'				- De la persona: <nombre>, <apellido_1>, <apellido_2>, <nif>
'				- De la dirección: <direccion>, <pais>, <provincia>, <municipio>, <via>, <calle>, <cod_postal>
Sub CrearPersona()
	Dim dir_esperada, dir_creada

	With Browser("Persona").Page("P").Frame("Frame")
		' Rellena los campos personales
		.SAPEdit("Nom").Set nombre
		.SAPEdit("Primer cognom").Set apellido_1
		.SAPEdit("Segon cognom").Set apellido_2
		.SAPEdit("NIF").Set nif
		
		' Activa el Checkbox "Indicador sincronizable"
		.SAPCheckBox("Indicador sincronitzable").Set "ON"
		
		' Cambia a la pestaña de las direcciones
		'.WebElement("Adreces").Click		
		Browser("Persona:").Page("Persona:").Frame("content_frame").WebElement("AdrecesAdreces-").Click
		Wait 2
		
		' Crea una nueva dirección asociada a la persona
		'.SAPButton("Nou").Click
		'Browser("Persona").Page("P").WebElement("WebElement").Click
		Browser("Persona").Page("P").Frame("content_frame_2").SAPUIButton("Nou").Click

		CrearDireccion
		
		' Validación de la dirección
		dir_esperada = 	via & " " & calle & ", " & municipio & ", " & cod_postal & ", " & provincia & ", " & pais		'Avinguda GENERALITAT, Abrera, 08630, Barcelona, Espanya
		dir_creada = .WebElement("Adreça nova").GetROProperty("innertext")
		If dir_esperada = dir_creada Then
			Reporter.ReportEvent micPass, "Adreça", "Adreça creada correctament:" & vbCrLf & dir_creada
		Else
			Reporter.ReportEvent micPass, "Adreça", "Adreça esperada:" & vbCrLf & "- " & dir_esperada & vbCrLf &_
													"Adreça creada:" & vbCrLf & "- " & dir_creada
		End If
		
		' Verifica los datos personales
		.SAPButton("Verificar").Click
		
		If .WebElement("Verificar").Exist(5) Then
			If .WebElement("Verificar").GetROProperty("innertext") = "No s'ha trobat cap error" Then				' Si no hay ningún error
				Reporter.ReportEvent micPass, "Verificar", .WebElement("Verificar").GetROProperty("innertext")		' OK
			ElseIf InStr(.WebElement("Verificar").GetROProperty("innertext"), "Ja existeix") Then					' Si el NIF está repetido
				Reporter.ReportEvent micWarning, "Verificar", .WebElement("Verificar").GetROProperty("innertext")	' Señala un Warning
				CambiarNIF																							' Cambia el NIF a otro aleatorio
			Else																									' En otro caso
				Reporter.ReportEvent micFail, "Verificar", .WebElement("Verificar").GetROProperty("innertext")		' Reporta el error
			End If
		End If
		
		' Solicita la aceptación de la nueva persona
		.SAPButton("Sol·licitar").Click
		
		' Valida que la solicitud se ha generado correctamente
		If .WebElement("Sol·licitud").Exist(5) Then
			Reporter.ReportEvent micPass, "Sol·licitud", .WebElement("Sol·licitud").GetROProperty("innertext")
			
			' Guarda el NIF y el número de la solicitud generada en el fichero de parámetros
			writeParameterFile Parameter("Parameter_File"), nif & " " & Split(.WebElement("Sol·licitud").GetROProperty("innertext"))(6)
		Else
			Reporter.ReportEvent micPass, "Sol·licitud", .WebElement("Sol·licitud").GetROProperty("innertext")
		End If
	End With
End Sub

'@Description	Crea una nueva dirección asociada a la persona en la Base de Datos Fiscales rellenando todos los campos obligatorios.
'				Toma los valores de los campos de las variables globales.
'@Documentation	Crea una nueva dirección asociada a la persona en la Base de Datos Fiscales rellenando todos los campos obligatorios.
'				Toma los valores de los campos de las variables globales: <direccion>, <pais>, <provincia>, <municipio>, <via>,
'				<calle> y <cod_postal>.
Sub CrearDireccion()
	With Browser("Persona").Page("P").SAPFrame("MDG")
		
		' Selecciona el tipo de dirección de la lista
		FakeListSelect .SAPEdit("Adreça"), direccion
		' Selecciona el país de la lista
		FakeListSelect .SAPEdit("País"), pais
		Wait 1
		' Introduce la provincia
		.SAPEdit("Província").SetAndEnter provincia
		Wait 1
		' Introduce el municipio
		.SAPEdit("Municipi").SetAndEnter municipio
		Wait 1
		' Selecciona el tipo de via de la lista
		FakeListSelect .SAPEdit("Tipus via"), via
		Wait 1
		' Introduce el nombre de la vía
		.SAPEdit("Nom via").SetAndEnter calle
		Wait 1
		' Introduce el código postal
		.SAPEdit("Codi postal").SetAndEnter cod_postal
		Wait 1
		
		' Confirma los datos introducidos
		.SAPButton("D'acord").Click
	End With
End Sub

'@Description	Cambia y verifica el NIF introducido hasta que se selecciona uno que no exista o se supere el número máximo de intentos.
'@Documentation	Cambia y verifica el NIF introducido hasta que se selecciona uno que no exista o se superen los <MAX_TRY> intentos.
Sub CambiarNIF()
	Const MAX_TRY = 50
	
	Dim intento : intento = 0

	With Browser("Persona").Page("P").Frame("Frame")
		' Cambia a la pestaña de los datos de la persona
		.WebElement("Persona").Click
		
		' Mientras se siga mostrando el mensaje de que el NIF ya existe, o se supere el máximo de intentos:
		Do
			Wait 1
			' Genera un nuevo NIF
			nif = GenerarNIF
			' Introduce el NIF
			.SAPEdit("NIF").Set nif
			.SAPEdit("NIF").Click
			
			' Verifica los datos
			.SAPButton("Verificar").Click
			Wait 5
			
			' Incrementa el número de intentos realizados
			intento = intento + 1
		Loop While InStr(.WebElement("Verificar").GetROProperty("innertext"), "Ja existeix") and intento < MAX_TRY
		
		' Valida que los datos son correctos
		If .WebElement("Verificar").GetROProperty("innertext") = "No s'ha trobat cap error" Then
			Reporter.ReportEvent micPass, "Verificar", .WebElement("Verificar").GetROProperty("innertext") & vbCrLf & "NIF: " & nif
		Else
			Reporter.ReportEvent micFail, "Verificar", .WebElement("Verificar").GetROProperty("innertext")
		End If
	End With
End Sub

'@Description	Simula el Select de las listas para los objetos SAPEdit que imitan a un SAPList sin serlo. Se utilizan acciones físicas,
'				por lo que la ventana debe estar en primer plano.
'@Documentation	Simula el Select de las listas para los objetos SAPEdit que imitan a un SAPList sin serlo. Se utilizan acciones físicas,
'				por lo que la ventana debe estar en primer plano.
Sub FakeListSelect(ByRef list, ByVal value)
'<list>:	SAPEdit que imita el estilo de un SAPList.
'<value>:	Valor que se desea seleccionar de la lista desplegable.

	list.Object.Focus
	list.FireEvent "click"
	
	Browser("Persona").Page("P").SAPFrame("MDG").WebElement("innertext:=" & value & " ").Object.Focus
	Browser("Persona").Page("P").SAPFrame("MDG").WebElement("innertext:=" & value & " ").ClickFisico
End Sub

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

