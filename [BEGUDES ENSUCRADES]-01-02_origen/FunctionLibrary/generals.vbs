'generals.vbs
'Historic: Versio - Data - Persona - Comentari
'	1.00	08/05/2009	Marc Gurt			Creació
'	1.01	14/05/2013	Alejandro Díaz		Se convierte configureFromEnv en algo práctico para el programador
'	1.02	21/05/2013	Alejandro Díaz		Se llama a configureFromEnv desde LoginUsuari, pero solo una vez
'	1.03	19/08/2013	Alejandro Díaz		Se añaden funciones que antes puse en ComunGaudi.qfl, aunque mejoradas
'											FechaNormal, Fecha, Hoy y HoraNormal
'	1.04	29/10/2013	Alejandro Díaz		Se añade la función EnteroSuperior ()
'	1.05	10/01/2014	Alejandro Díaz		Se añade la función SubArray ()
'	1.06	14/01/2014	Alejandro Díaz		Se añaden GuardarFichero (), NombrePaso () SiguientePaso (), GuardarEnLineas () y LeerLineas ()
'	1.07	12/02/2014	Assumpta Burriel	Se añade CopiarPDFenNotepad y BuscarString
'	1.08	24/02/2014	Alejandro Díaz		Se modifica y mueve NIF_aleatorio () desde GaudiGenerals.vbs
'											Se añaden CalcularNIF () y EsNIFvalido ()
'	1.09	25/02/2014	Alejandro Díaz		Se modifica InformeResultado porque WebElement (X).Exist (0) puede devolver un True que es 1 y no -1
'	1.10	04/03/2014	Alejandro Díaz		Se añade IsAnEmptyArray()
'	1.11	14/05/2014	Alejandro Díaz		Se añade a ConfigureFromEnv el tratamiento de ImageCaptureForTestResults
'	1.12	22/05/20N114	Alejandro Díaz		Se modifica ConfigureFromEnv para no tener que cambiar en la variable "entorn" y también "URL_ENV"
'	1.13	27/05/2014	Alejandro Díaz		Se traslada AsociarRepositorio() desde GaudiGenerals.vbs
'	1.14	23/06/2014	Alejandro Díaz		Se añade el procedimiento Tabular()

'	1.15	11/08/2014	Alejandro Díaz		------------ Migración a QC 11.00, UFT 11.50 y IE 8 ------------
'
'	1.16	18/08/2014	Alejandro Díaz		Se añade AceptarCertificado() para poder operar en UFT 11.50 sobre IE 6, 8...
'											En cada versión de QTP hay una forma diferente de referirse a la "carpeta" en que se guardan
'											los repositorios. Se añade RUTA_REPOSITORIOS, que se inicializa en LocalizarRepositorios ()
'	1.17	18/08/2014	Alejandro Díaz		Se adapta a cualquier versión de Internet Explorer: AceptarCertificado () y RUTA_EXPLORER
'	1.18	27/08/2014	Alejandro Díaz		QC 11 / UFT 11.50 no parecen capaces de reconocer el IE abierto por un test previo.
'											Así que de momento lo vamos a configurar para que cada test cierre el IE al acabar.
'	1.19	16/09/2014	Alejandro Díaz		Se mejora la selección de la versión de IExplorer 32 ó 64 bits, en función de lo que haya
'	1.20	14/10/2014	Alejandro Díaz		Sincronización del diálogo de impresión de IE8, más lento, quizá por lo complejo.
'	1.21	14/01/2015	Monica Piñol & 		Se añade InformeResultadoMultiple
'						Francesc Alonso
'	1.22	16/01/2015	Alejandro Díaz		Se añade la función ParameterIsEmpty para facilitar la programación de
'											subprogramas con parámetros optativos.
'	1.23	24/02/2015	Alejandro Díaz		La configuración de cómo se hace click ha de ser diferente para los check boxes de los grid ExtJS
'											y para los edit boxes de los combos ExtJS de Gaudí. Hay que cambiar la configuración antes del click.
'											Esto se debe a la migración de Gaudí a WL12
'	1.24	13/05/2015	Alejandro Díaz		Se añade la función ExtraerVariables ()
'	1.25	26/08/2015	Alejandro Díaz		Se trae de GaudiGenerals.vbs la definición de EsteUFT, que aquí se usa más
'	1.26	26/08/2015	Alejandro Díaz		------------ Migración a ALM 12.20, UFT 12.02 ------------
'	1.27	14/09/2015	Alejandro Díaz		Se añade Coinciden () para hacer un código más limpio en los tests
'	1.28	15/09/2015	Alejandro Díaz		Se añade IndiceWebList () para generalizar la búsqueda de opciones en las listas
'	1.29	23/09/2015	Alejandro Díaz		Se añade GetContenidoPDF() como apoyo para la comprobación de los informes contables de Ingressos
'	1.30	06/10/2015	Alejandro Díaz		En una librería, la inicialización de variables e incluso su declaración no tiene porqué
'											interpretarse antes de ejecutarse las llamadas a las funciones que las emplean.
'											Como UltimoRepositorio y RepositoriosCargados (10), que pasan a tratarse de modo dinámico
'	1.31	20/11/2015	Alejandro Díaz		Conchita pasa a ser pública para que la use cualquier otra librería y script.
'	1.32	03/03/2016	Alejandro Díaz		Se mejora WaitObject para minimizar consultas y ajustar tiempos
'	1.33	04/07/2016	Alejandro Diaz		Se añade GuardarPDF () para guardar el contenido de los documentos mostrados en Acrobat Reader
'	1.34	30/08/2016	Alejandro Díaz		Se trae MostrarMensaje () desde GaudiGenerals.vbs para usarla en los tests de Portal.
'											Y se le añade el sinónimo Traza ()
'	1.35	05/09/2016	Alejandro Díaz		Se añade Existe para usar la propiedad Exist () sin que dé problemas
'	1.36	07/09/2016	Alejandro Díaz		Se añade EscaparER () para escapar una cadena que se vaya a usar como expresión regular
'	1.37n	13/10/2016	Alejandro Díaz		Se añaden LazyAnd, LazyOr, LazyAnd_M y LazyOr_M para evaluación perezosa.
'											No funcionan porque desde esta librería no se tiene acceso a las variables declaradas fuera.
'											Pero en v1.39 se consigue apañar
'	1.38	05/12/2016	J.M. Loscertales	Se extrae a EtiquetaPDF () la lectura de la etiqueta desde el PDF,
'											usado en la acción IncorporarDocumento y en los tests de Inspecció, con accesos diferentes.
'											Se añade ContenidoTablaWeb () para el análisis de las tablas web
'	1.39	09/02/2017	Alejandro Díaz		Ahora sí que funcionan LazyAnd, LazyOr, LazyAnd_M y LazyOr_M.
'											Es preciso incluir sus declaraciones en una cadena a usar en el script / librería cliente
'											mediante la sentencia Execute.
'											AVISO: Esta sentencia impide depurar.
'											DeclaraPaquete permite emplear Execute de forma transparente.
'											Pero no funciona con objetos, solo con variables normales.
'	1.40	17/02/2017	Alejandro Díaz		Aseguramos que InformeResultado() y compañía siempre muestren sus mensajes.
'											Es neceario tras modificar WaitCarregantRepoGenerals() con v1.86.
'	1.41	10/03/2017	Alejandro Díaz		Coinciden() ya solo registra los problemas, en lugar de todos los resultados, haciendo los
'											informes de resultados más legibles y escuetos. Se añade Pantallazo().
'	1.42	19/04/2017	Alejandro Díaz		Se pasa EtiquetaPDF() a GaudiGenerals.vbs y se modifica GetContenidoPDF() para asegurarse
'											de que se lee la etiqueta, porque en 28.1 el PDF tarda en mostrarse y no da tiempo a cogerla.
'	1.43	26/05/2017	Alejandro Díaz		EsteUFT pasa a ser una función
'	1.44	06/10/2017	Marc Gurt			Ejecucion totalmente desasistida no interactua con botones hacemos close de ventana
'	1.44	17/12/2018	José Luis Pérez		Añadida la función GenerarNIF ().
'
' Buscar *** para posibles mejoras
' Buscar ### para código comentado pendiente de borrado



'***************************************************************************************************************
'************************************************** PUBLIC **************************************************
'***************************************************************************************************************
Dim Existe																				' Para asignar el resultado de la propiedad Exist

Const t120 = 120  'Segons d'espera
Dim tWait  'Segons d'espera
Const tWaitExist = 5  'Segons d'espera per validar si Exist o No
tWait = 120  'Segons d'espera
'**************************************** UTILITATS DIVERSES QTP ****************************************
'@Description Espera que un objecte estigui disponible.
'@Documentation Espera <iSeconds> seg. a que <objObject> estigui disponible
Public Function WaitObject(objObject, iSeconds)
'objObject: objecte pel que espera
'iSeconds: segons que ha d'esperar que aparegui l'objecte. Cada iSecondsWait segons (fixat a 2) prova si el troba
'Retorna valor boleà: True si l'ha trobat i False si no l'ha trobat.

	WaitObject = objObject.Exist (iSeconds)												' Es preciso porque Exist es propiedad y no función

' *** v1.32
'		No sirve, porque si objObject no existe al llegar luego, cuando existe, no sabe ponerse al día.
'		Pero quizá si se le pasase una colección con las propiedades...
'	Dim Total, Fraccion
'
'	MercuryTimers.Timer ("x").Start
'	Total = iSeconds * 1000
'	Fraccion = iSeconds * 100															' Hacemos 10 esperas
'	Do
'		If objObject.Exist (Fraccion) Then
'			WaitObject = True
'			Exit do
'		End If
'		If MercuryTimers ("x").ElapsedTime > Total Then
'			WaitObject = False
'			Exit do
'		End If
'	Loop
'	MercuryTimers ("x").Stop

' ### v1.32
'	Const iSecondsWait=2
'	iCounter=iSeconds/iSecondsWait
'	Do until( objObject.Exist(iSecondsWait) )
'		iCounter = iCounter-1
'		if (iCounter<1) then Exit Do
'	Loop
'	waitObject=objObject.Exist(0)
End Function

'@Description Espera que el valor d'un camp de text canvii
'@Documentation Espera <iSeconds> seg. a que en <objObject> canvii el valor <iniValue>
Public Function waitValueChange( objObject, iniValue, iSeconds)
'objObject: objecte camp de text.
'iniValue: valor inicial que ha de canviar
'iSeconds: segons que ha d'esperar que aparegui l'objecte. Cada iSecondsWait segons (fixat a 2) prova si el troba.
'Retorna valor boleà: True si ha canviat i False si no.
	Dim iCounter
	Const iSecondsWaitwvc = 2
	iCounter = iSeconds/iSecondsWaitwvc
	Do until( objObject.Object.value<>iniValue )
		wait(iSecondsWaitwvc)
		iCounter = iCounter-1
		if (iCounter<1) then Exit Do
	Loop
	waitValueChange = (objObject.Object.value<>iniValue)
End Function

'@Description Retorna timeStamp, en base vTime, en el format indicat
'@Documentation Retorna timeStamp basat en <vTime>, del tipus <iType>
Public Function TimeStamp(vTime, iType)
'vTime: temps a partir del que obtenir el timeStamp. Habitualment s'indicara la funcio Now
'iType: tipus de timeStamp retornat
' 1 : YYYYMMDDhhmmss
' 2 : YYMMDDhhmmss
' 3:  YYMMDD
' 4:  YYMMDDhhmm
' 5:  MMDDhhmm

	Dim Partes (5)																		' Los distintos componentes a usar
	Dim Lista																			' La lista que los contiene

	If 2 > iType or iType > 5 Then														' Primero montamos las partes (que necesitamos)
		Partes (0) = Year (vTime)
	ElseIf iType <> 5 Then
		Partes (0) = Right (CStr (Year (vTime)), 2)
	End If
	Partes (1) = DosDigitos (Month (vTime))
	Partes (2) = DosDigitos (Day (vTime))
	If iType <> 3 Then
		Partes (3) = DosDigitos (Hour (vTime))
		Partes (4) = DosDigitos (Minute (vTime))
	End If
	If 2 >= iType or iType > 5 Then
		Partes (5) = DosDigitos (Second (vTime))
	End If
	
	Select Case iType																	' Y después las unimos
		Case 3																			' YYMMDD
			Lista = Array (Partes (0), Partes (1), Partes (2))
		Case 4																			' YYMMDDhhmm
			Lista = Array (Partes (0), Partes (1), Partes (2), Partes (3), Partes (4))
		Case 5																			' MMDDhhmm
			Lista = Array (Partes (1), Partes (2), Partes (3), Partes (4))
		Case else																		' Con (-inf, 2] U [6, +inf) va todo (YY)YYMMDDhhmmss
			Lista = Partes
	End Select
	TimeStamp = Join (Lista, "")


' ### Lo que dejó HP, con mucho copia y pega
'	Select Case iType
'		Case 1		timeStamp=year(vTime) & right("0" & month(vTime),2) & right("0" & day(vTime),2) & right("0" & hour(vTime),2) & right("0" & minute(vTime),2) & right("0" & second(vTime),2)
'		Case 2		timeStamp=right(Cstr(year(vTime)),2) & right("0" & month(vTime),2) & right("0" & day(vTime),2) & right("0" & hour(vTime),2) & right("0" & minute(vTime),2) & right("0" & second(vTime),2)
'		Case 3		timeStamp=right(Cstr(year(vTime)),2) & right("0" & month(vTime),2) & right("0" & day(vTime),2) 
'		Case 4		timeStamp=right(Cstr(year(vTime)),2) & right("0" & month(vTime),2) & right("0" & day(vTime),2) & right("0" & hour(vTime),2) & right("0" & minute(vTime),2)
'		Case 5		timeStamp=right("0" & month(vTime),2) & right("0" & day(vTime),2) & right("0" & hour(vTime),2) & right("0" & minute(vTime),2)
'		Case Else	timeStamp=year(vTime) & right("0" & month(vTime),2) & right("0" & day(vTime),2) & right("0" & hour(vTime),2) & right("0" & minute(vTime),2) & right("0" & second(vTime),2)
'	End Select
End Function

Public Sub parameterADataTable( sParameter )
'Si QTP te el paràmetre informat, el volca a la DataTable Local. ATENCIO: parametre i columna de Datatable han de tenir el mateix nom
	If Len(Parameter(sParameter)) <> 0 Then DataTable(sParameter, dtLocalSheet) = Parameter(sParameter)
End Sub

Private EstaConfigurado																	' Para configurarlo solo una vez, en el login o fuera
Private ElUnicoUFT																		' Única instancia accesible con EsteUFT()

' v1.25	Es global para ahorrarnos tener que crearla una y otra vez y para que esté accesible a cada instancia de cParrilla
' ###Dim EsteUFT
Public Function EsteUFT()

	If IsEmpty (ElUnicoUFT) Then
		Set ElUnicoUFT = getObject ("", "QuickTest.Application")
	End If
	Set EsteUFT = ElUnicoUFT
	
End Function

'@Description Configura variables del QTP a partir del XML d'Environment
'@Documentation Configura variables del QTP a partir del XML d'Environment
Public Sub configureFromEnv()

	If EstaConfigurado Then																' v1.02: HP llamaba a esta funció solo en algunos tests
		Exit sub																		' Ahora se llama siempre que se hace login,
	End If																				' pero solo una vez

	Dim XML_ENTORNO : XML_ENTORNO = QTP_FILES & "variableEnvironment.xml"
' ###
'	Const XML_ENTORNO = "C:\QTP_FILES\variableEnvironment.xml"

' ### v1.25
'	Dim qtpApp
	Dim i, n

	EstaConfigurado = True
' ### v1.43
'	If IsEmpty (EsteUFT) Then															' v1.25
'		Set EsteUFT = getObject ("", "QuickTest.Application")
'	End If
' ###
'	Set qtpApp = getObject ("", "QuickTest.Application")
	EsteUFT.Test.Environment.LoadFromFile (XML_ENTORNO)
	Environment ("URL_ENV") = Environment (Environment ("Entorn"))						' v1.12
	EsteUFT.Options.Web.RunMouseByEvents = True											' v1.23	Todo con eventos web. Los checkboxes de los grids
' v1.11		En general, lo que haya en el fichero de configuración.
	EsteUFT.Options.Run.ImageCaptureForTestResults = Environment ("ImageCaptureForTestResults")
' ***	La línea siguiente es una alternativa a la edición del fichero XML, útil quizá cuando se esté desarrollando.
'	qtpApp.Options.Run.ImageCaptureForTestResults = "Always"
	With EsteUFT.Test.Settings
		.Run.ObjectSyncTimeOut = CLng (Environment ("ObjectSyncTimeOut"))
		.Run.OnError = Environment("SettingsRunOnError")  '"Dialog"															' Para que cuando no se use QC, podamos depurar
		.Run.DisableSmartIdentification = (Environment ("DisableSmartIdentification") = "True")
		.Web.BrowserNavigationTimeout = CLng (Environment ("BrowserNavigationTimeout"))
		With .Launchers ("Web")
			.Active = False																' Queremos usar sesiones ya abiertas para lo cual
' ***	La cadena mostrada depende de cada máquina.
'		No la especificamos porque en el CTTI todo será "Microsoft Internet Explorer" de 32 bits en W7 32 bits
'			.Browser = "IE"
			.CloseOnExit = True
' ###	v1.18	Mientras no encuentre la forma de que funcione lo de abajo, tendremos que cerrar
'			.CloseOnExit = False														' necesitamos que no se cierre el navegador al acabar
		End with
	End with
	AsociarRepositorio "Generals"
' ###
'	AsociarRepositorio RUTA_REPOSITORIOS & "\Generals.tsr"
' ###
'	On error resume next																' *** Puede que no se llame Action1 o que ya esté
'	n = qtpApp.Test.Actions.Count
'	For i = 1 to n
'' ###
''		qtpApp.Test.Actions (i).ObjectRepositories.Add "[QualityCenter] Subject\.Repositoris\Generals.tsr"
'		qtpApp.Test.Actions (i).ObjectRepositories.Add RUTA_REPOSITORIOS & "\Generals.tsr"
'	Next
'' ###
''	qtpApp.Test.Actions ("Action1").ObjectRepositories.Add "[QualityCenter] Subject\.Repositoris\Generals.tsr"
'	On error goto 0
' ### v1.25
'	Set qtpApp = Nothing

' ### Lo que dejó HP v1.00
'	Dim qtpApp
'	Set qtpApp = getObject("","QuickTest.Application")
'	qtpApp.Test.Settings.Run.ObjectSyncTimeOut = CLng( Environment ("ObjectSyncTimeOut") )
'	If (Environment ("DisableSmartIdentification")="True") Then
'		qtpApp.Test.Settings.Run.DisableSmartIdentification = True
'	Else
'		qtpApp.Test.Settings.Run.DisableSmartIdentification = False
'	End If
'	qtpApp.Test.Settings.Run.OnError = Environment("SettingsRunOnError")
'	qtpApp.Test.Settings.Web.BrowserNavigationTimeout = CLng( Environment ("BrowserNavigationTimeout") )
'	qtpApp.Test.Settings.Launchers("Web").Active = True
'	qtpApp.Test.Settings.Launchers("Web").Browser = "IE"
'	'qtpApp.Test.Settings.Launchers("Web").Address = "http://gtgaudi.int.serveis.atc.gencat.cat/gt_web/index.jsp"
'	qtpApp.Test.Settings.Launchers("Web").CloseOnExit = True
'	Set qtpApp = Nothing

End Sub



'@Description Retorna True si l'addin ExtJS esta actiu
'@Documentation Valida si l'addin ExtJS esta actiu
Public Function ExtJSActive()
' ### v1.25
'	Dim qtpApp, addins
	Dim Addins
	ExtJSActive = False

' ### v1.43
'	If IsEmpty (EsteUFT) Then															' v1.25
'		Set EsteUFT = getObject ("", "QuickTest.Application")
'	End If
' ###
'	Set qtpApp = getObject("","QuickTest.Application")
	Set addins = EsteUFT.Addins
	For i = 1 to addins.Count
		If addins.Item (i).Name = "ExtJS" Then
			ExtJSActive = addins.Item (i).Status = "Active"
			Exit For
		End If
	Next
	Set addins = Nothing
' ### v1.25
'	Set qtApp = Nothing
End Function

'@Description Retorna els segons que hi ha entre els temps indicats
'@Documentation Retorna els segons entre els temps <tstart> i <tend>
Public Function secDifference( ByVal tstart, ByVal tend)
	secDifference = CInt( (tend-tstart)*24*60*60 )
End Function

'@Description Retorna <sText> amb la primera mayúscula i les següents minúscoles
'@Documentation Retorna <sText> amb la primera mayúscula i les següents minúscoles
Function capitalCase( sText )
   capitalCase = UCase( Left( sText, 1) ) & LCase( Right( sText, Len(sText)-1 ) )
End Function

'@Description Retorna texts concatenats per separador <sSeparator>. Si sTextini es "" només torna sTextAdd
'@Documentation Retorna texts concatenats per separador <sSeparator>. Si sTextini es "" només torna sTextAdd
Function addCodingRow( sTextini, sTextAdd, sSeparator )
	If ( Len(sTextini)=0 ) Then
		addCodingRow = sTextAdd
	Else
		addCodingRow = sTextini & sSeparator & sTextAdd
	End If
End Function

'@Description A la variable <sText> li afageix el paràmetre de nom <sParamName> i valor <sParamValue>
'@Documentation A la variable <sText> li afageix el paràmetre de nom <sParamName> i valor <sParamValue>
Function setParam( ByRef sText, ByVal sParamName, ByVal sParamValue )
'Els diferents paràmetres emmagatzemats es separan per ;
'sParamName i sParamValue son String afageix el Param=Valor. Si son Arrays, afageix tots els Param=Valor dels Arrays
	If Len(sText)=0 Then  'Si la variable està en blanc l'inicialitza amb ;
		sText =  ";"
	End If
	If VarType(sParamName) = vbString Then
		sText = sText & sParamName & "=" & sParamValue & ";"
	Else  'es Array
		Dim iPos
		For iPos = 0 To Ubound( sParamName )
			sText = sText & sParamName(iPos) & "=" & sParamValue(iPos) & ";"
		Next
	End If
	setParam = sText
End Function

'@Description A la variable <sText> li afageix el paràmetre multivalor de nom <sParamName> i valors l'array <sParamValue>
'@Documentation A la variable <sText> li afageix el paràmetre multivalor de nom <sParamName> i valors l'array <sParamValue>
Function setParamMulti( ByRef sText, ByVal sParamName, ByVal sParamValue )
'Els diferents paràmetres emmagatzemats es separan per ;
'sParamName i sParamValue son String afageix el Param=Valor. Si son Arrays, afageix tots els Param=Valor dels Arrays
	If Len(sText)=0 Then  'Si la variable està en blanc l'inicialitza amb ;
		sText =  ";"
	End If
	Dim iPos
	For iPos = 0 To Ubound( sParamValue )
		sText = sText & sParamName & "=" & sParamValue(iPos) & ";"
	Next
	setParamMulti = sText
End Function

'@Description Donada la cadena de paràmetres <sText> retorna el valor del paràmetre de nom <sParamName>
'@Documentation Donada la cadena de paràmetres <sText> retorna el valor del paràmetre de nom <sParamName>
Function getParam( ByVal sText, ByVal sParamName )
	Dim iPosini, iPosFin, iLenParamName
	iPosini = Instr( sText, ";" & sParamName & "=" )
	If ( iPosini>0 ) Then
		iLenParamName = Len(sParamName) + 2  '+ 2 pel ; i el =
		iPosFin = Instr( iPosini+1, sText, ";" )
		getParam = Mid( sText, iPosini + iLenParamName, iPosFin - (iPosini + iLenParamName) )
	Else  'NO hi ha el paràmetre
		getParam = ""
	End If
End Function

'@Description Donada la cadena de paràmetres <sText> reeplaça el valor del paràmetre de nom <sParamName>
'@Documentation Donada la cadena de paràmetres <sText> reeplaça el valor del paràmetre de nom <sParamName>
Function replaceParam( ByRef sText, ByVal sParamName, ByVal sNewParamValue )
	Dim iPosini, iPosFin, iLenParamName
	iPosini = Instr( sText, ";" & sParamName & "=" )
	If ( iPosini>0 ) Then
		iLenParamName = Len(sParamName) + 1
		iPosFin = Instr( iPosini+1, sText, ";" )
		sText = Left( sText, iPosini + iLenParamName ) & sNewParamValue & Mid( sText, iPosFin )
	Else  'si NO hi ha el paràmetre l'afageix
		setParam sText, sParamName, sNewParamValue
	End If
	replaceParam = sText
End Function

'@Description Donada la cadena de paràmetres <sText> retorna array dels valors del paràmetre de nom <sParamName>
'@Documentation Donada la cadena de paràmetres <sText> retorna array dels valors del paràmetre de nom <sParamName>
Function getParamMulti( ByVal sText, ByVal sParamName )
	Dim iPosini, iPosFin, iLenParamName, iPosAr
	ReDim arParamMulti(0)
	iPosAr = 0
	arParamMulti(0) = ""
	Do while (True)
		iPosini = Instr( sText, ";" & sParamName & "=" )
		If ( iPosini>0 ) Then
			iLenParamName = Len(sParamName) + 2  '+ 2 pel ; i el =
			iPosFin = Instr( iPosini+1, sText, ";" )
			ReDim Preserve arParamMulti( iPosAr )
			arParamMulti( iPosAr ) = Mid( sText, iPosini + iLenParamName, iPosFin - (iPosini + iLenParamName) )
			iPosAr = iPosAr + 1
			sText = Mid( sText, iPosini + iLenParamName )
		Else  'NO hi ha el paràmetre
			Exit Do
		End If
	Loop
	getParamMulti = arParamMulti
End Function

'@Description Retorna el primer valor entre parentesi treient-li tots els espais
'@Documentation Retorna el primer valor entre parentesi treient-li tots els espais
Function getValueParentheses( ByVal sText )
	Dim pini, pfin
	getValueParentheses = ""
	pini = instr( sText, "(" )
	If ( pini>0 ) Then
		pfin = instr( pini, sText, ")" )
		If ( pfin>0 ) Then
			getValueParentheses = Replace( Mid( sText, pini+1, pfin-(pini+1) ), " ", "" )
		End If
	End If
End Function

'**************************************** TREBALL AMB FITXERS ****************************************

'@Description Escriu, al fitxer indicat el text indicat. Utilitza variable global sDirParameterFile
'@Documentation Escriu <sText> al fitxer <sFileName> (en /PARAMETERS)
Public Sub writeParameterFile( sFileName, sText )
'sFileName: nom del fitxer que contindra el text
'sText: text a incloure al fitxer
	writeFile sDirParameterFile & sFileName, sText, 2
End Sub

'@Description Llegeix la primera linia del fitxer de text indicat. Utilitza variable global sDirParameterFile
'@Documentation Llegeix la primera linia del fitxer <sFileName> (en /PARAMETERS)
Public Function readParameterFile( sFileName )
'sFileName: nom del fitxer de text del que llegira
'Retorna el text llegit del fitxer indicat
	readParameterFile = readFile( sDirParameterFile & sFileName, True )
End Function

'@Description Llegeix tot el fitxer de text indicat. Utilitza variable global sDirParameterFile
'@Documentation Llegeix tot el fitxer <sFileName> (en /PARAMETERS)
Public Function readAllParameterFile( sFileName )
'sFileName: nom del fitxer de text del que llegira
'Retorna el text llegit del fitxer indicat
	readAllParameterFile = readFile( sDirParameterFile & sFileName, False )
End Function

'@Description Llegeix tot el fitxer de text indicat. Cal nom complet (Unitat:Directori\Fitxer)
'@Documentation Llegeix tot el fitxer <sFilePath>
Public Function readAllFile( sFilePath )
'sFilePath: nom complet del fitxer de text que llegira
'Retorna el text llegit del fitxer indicat
	readAllFile = readFile( sFilePath, False )
End Function

'@Description Afegeix al fitxer indicat el text indicat. Utilitza variable global sDirPreTestDataFile
'@Documentation Afegeix <sText> al fitxer <sFileName> (en /PRETESTDATA)
Public Sub addPreTestDataFile( sFileName, sText )
'sFileName: nom del fitxer que s'afegirà el text
'sText: text a afegir al fitxer
	writeFile sDirPreTestDataFile & sFileName, sText, 8
End Sub

'@Description Retorna la primera línia del fitxer indicat i esborra la línia del fitxer. Utilitza variable global sDirPreTestDataFile
'@Documentation Retorna la primera línia del fitxer (en /PRETESTDATA) indicat i esborra la línia del fitxer
Public Function readPreTestDataFile( sFileName )
'sFileName: nom del fitxer que es llegeix el text
	readPreTestDataFile = readFile( sDirPreTestDataFile & sFileName, True )
	deleteLineFromFilePriv sDirPreTestDataFile & sFileName, 1
End Function

'@Description Llegeix un dels parametres (separats per ;) del text contingut al fitxer de text indicat. Utilitza variable global sDirParameterFile
'@Documentation Llegeix el paràmetre <sParameter> (separats per ;) del text contingut al fitxer <sFileName> (en /PARAMETERS)
Public Function readOneParameterFromFile( sFileName , sParameter)
'sFileName: nom del fitxer de text del que llegira el text
'sParameter: numero de parametre en la linea
'Retorna el paràmetre (separats per ;) del text contingut al fitxer
	Dim readParametersInFile, parametersList
	readParametersInFile = readParameterFile( sFileName )
	parametersList = Split( readParametersInFile, ";", -1, 1)
	readOneParameterFromFile = ParametersList(CInt(sParameter) - 1)
End Function

'@Description Esborra una linia del fitxer indicat. Utilitza variable global sDirParameterFile
'@Documentation Esborra la linia <iline> del contingut del fitxer <sFileName> (en /PARAMETERS)
Public Function deleteLineFromFile ( sFileName, iline)
'sFileName: nom del fitxer
'iline: linia a esborrar
	deleteLineFromFilePriv sDirParameterFile & sFileName, iline
End Function

'@Description Esborra una linia del fitxer indicat
'@Documentation Esborra la linia <iline> del contingut del fitxer <sFileName>
Private Function deleteLineFromFilePriv ( sFileName, iline)
'sFileName: directori i nom del fitxer
'iline: linia a esborrar
	Const ForReading = 1
	Const ForWriting = 2
	Dim objFSO, objTextFile, strContents, arrLines, i
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile( sFileName, ForReading)
	strContents = objTextFile.ReadAll
	objTextFile.Close
	arrLines = Split(strContents, vbCrLf)
	Set objTextFile = objFSO.OpenTextFile( sFileName, ForWriting)
	For i = 0 to UBound(arrLines)
		If i <> iline - 1 Then
			objTextFile.WriteLine arrLines(i)
		End If
	Next
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO = Nothing
End Function

'@Description Afegeix al final del fitxer indicat el text indicat. Si el fitxer NO existeix el crea. Utilitza variable global sDirParameterFile
'@Documentation Afegeix al final del fitxer <sFileName> el text <sText>
Public Sub addToLogFile( sFileName, sText)
'sFileName: nom del fitxer de text al que s'afegirà el text
'sText: text a afegir al fitxer
	Dim objFSO, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile( sDirParameterFile & sFileName, 8, True)  '8: mode afegir / True: el crea si no existeix
	objTextFile.WriteLine( sText ) 
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO = Nothing
End Sub

'@Description	Comprueba si existe el fichero indicado. Utiliza la variable global sDirParameterFile.
'@Documentation	Comprueba si existe el fichero  <sFileName> (en /PARAMETERS).
Public Function fileExist(sFileName)
'<sFileName>:	Nombre del fichero que se verificará.

	Dim fileSystemObj : Set fileSystemObj = createobject("Scripting.FileSystemObject")
	
	fileExist = fileSystemObj.FileExists(sDirParameterFile & sFileName)
End Function

Public Function readLines (sFileName)
'	Dim fs, f
'
'	Set FS = CreateObject ("Scripting.FileSystemObject")
'	Set f = fs.OpenTextFile (Fichero, ForReading, True)
'	LeerLineas = Split (f.ReadAll, vbCrLf)
'	f.Close
'	Set f = Nothing
'	Set fs = Nothing
	readLines = LeerLineas(sDirParameterFile & sFileName)
End Function
'**************************************** VALIDACIONS ****************************************

'@Description Valida si el valor actual es igual a l'esperat
'@Documentation Valida si <title> te valor <valueActual> igual a <valueExpected>
Public Sub validationCompValues( title, valueExpected, valueActual )
	If ( valueExpected=valueActual ) Then
		Reporter.ReportEvent micPass, title, title & " te el valor esperat -->  (Esperat) = " & valueExpected
	Else
		Reporter.ReportEvent micFail, title, title & " NO te el valor esperat --> (Esperat) = " & valueExpected & " pero (Obtingut) = " & valueActual
	End If
End Sub

'@Description Check que existeixen els elements de l'array. Retorna True si es troben tots
'@Documentation Check que existeixen els elements <arElements>
Public Function checkElementsExist(ByVal arElements)
	Dim valExist
	checkElementsExist = True
	For Each iElement in arElements
	   valExist = iElement.Exist(tWaitExist)
	   checkElementsExist = checkElementsExist And valExist
		If ( valExist ) Then
			Reporter.ReportEvent micPass, iElement.ToString, "Element trobat: " & iElement.ToString
		Else
			Reporter.ReportEvent micFail, iElement.ToString, "Element NO trobat: " & iElement.ToString
		End If
   Next
End Function

'@Description Check que NO existeixen els elements de l'array. Retorna True si NO s'ha trobat cap d'ells
'@Documentation Check que NO existeixen els elements <arElements>
Public Function checkElementsNoExist(ByVal arElements)
	Dim valExist
	checkElementsNoExist = True
	For Each iElement in arElements
	   valExist = iElement.Exist(tWaitExist)
	   checkElementsNoExist = checkElementsNoExist And Not(valExist)
		If ( valExist ) Then
			Reporter.ReportEvent micFail, iElement.ToString, "Element SI trobat però NO hauria de ser-hi: " & iElement.ToString
		Else
			Reporter.ReportEvent micPass, iElement.ToString, "Element NO trobat: " & iElement.ToString
		End If
   Next
End Function


'**************************************** NOVES FUNCIONS DE OBJECTE FRAME ****************************************

'@Description Retorna el contingut assenyalat per l'etiqueta (en RegExp)
'@Documentation Retorna el contingut assenyalat per <sLabel> (RegExp)
Public Function getValueLabelH( ByRef test_object, ByVal sLabel, ByVal posH )
'En una taula HTML on l'etiqueta i el valor estan en cel·les (TD) contigües horitzontals
'sLabel: ATENCIO es Expressio Regular --> Cal escapar (, ), ., ...
'posH: posicions horitzontals que cal anar a la dreta, normalment és 1
	Dim objDesc, objTD
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "TD"
	objDesc("innertext").Value = sLabel
	objDesc("index").Value = 0
	Set objTD = test_object.WebElement( objDesc)
	If objTD.Exist Then
		Dim objTDObject, i
		Set objTDObject = objTD.Object
		For i = 1 to posH
			Set objTDObject = objTDObject.nextSibling
			If ( objTDObject is Nothing ) Then Exit For
		Next
		If ( objTDObject is Nothing ) Then
			getValueLabelH =  ""
		Else
			getValueLabelH =  objTDObject.innerText
		End If
	Else
		Reporter.ReportEvent micFail, sLabel & " - NO TROBADA", "No s'ha trobat l'etiqueta '" & sLabel & "'"
		getValueLabelH = ""
	End If
	Set objTD = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "getValueLabelH", "getValueLabelH"

'@Description Check el contingut assenalat per l'etiqueta (en RegExp)
'@Documentation Valida que el contingut assenyalat per <sLabel> (RegExp) es <sValue>
Public Function checkValueLabelH( ByRef test_object, ByVal sLabel, ByVal sValue, ByVal posH )
'En una taula HTML on l'etiqueta i el valor estan en cel·les (TD) contigues horitzontals
'sLabel: ATENCIO ee Expressio Regular --> Cal escapar (, ), ., ...
'posH: posicions horitzontals que cal anar a la dreta, normalment és 1
	Dim sValueNow
	sValueNow = getValueLabelH(test_object, sLabel, posH)
	If ( sValueNow=sValue ) Then
		Reporter.ReportEvent micPass, sLabel, "El valor de '" & sLabel & "' es CORRECTA = '" & sValue & "'"
	Else
		Reporter.ReportEvent micFail, sLabel, "El valor de '" & sLabel & "' es ERRONI. Valor esperat = '" & sValue & "'  /  Valor trobat = '" & sValueNow & "'."
	End If
End Function
RegisterUserFunc "Frame", "checkValueLabelH", "checkValueLabelH"

'@Description Retornar el valor de la columna <retColumn> trobat en la linia de expressions de cerca <arLabels>
'@Documentation Retornar el valor de la columna <retColumn> trobat en la linia de expressions de cerca <arLabels>
Public Function getLineColumnByText( ByRef test_object, ByVal arLabels, ByVal posLabelSearch, ByVal retColumn )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
'posLabelSearch: element de l'array que a usar per la primera cerca. Inici per 0. Si dona igual posar 0.
'retColumn: s'indica que retornar. Sempre produeix Pass/Fail al report. Ha de ser un numero de 0 a Ubound(arLabels)-1
	getLineColumnByText = checkLine( test_object, arLabels, posLabelSearch, retColumn )
End Function
RegisterUserFunc "Frame", "getLineColumnByText", "getLineColumnByText"

'@Description Check si existeix la línia de dades introduïda com array d'expressions regulars
'@Documentation Check si existeix la línia de dades <arLabels>
Public Function checkLineTextExist( ByRef test_object, ByVal arLabels )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
	checkLineTextExist = checkLine( test_object, arLabels, 0, -1 )
End Function
RegisterUserFunc "Frame", "checkLineTextExist", "checkLineTextExist"

'@Description Check si existeix la línia de dades introduïda com array d'expressions regulars fent una sola comprovació
'@Documentation Check si existeix la línia de dades <arLabels>
Public Function checkLineTextExistFast( ByRef test_object, ByVal arLabels )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
	Dim tWaitini
	tWaitini = tWait
	tWait = 0
	checkLineTextExistFast = checkLine( test_object, arLabels, 0, -1 )
	tWait = tWaitini
End Function
RegisterUserFunc "Frame", "checkLineTextExistFast", "checkLineTextExistFast"

'@Description Check si hi ha una taula amb columnes <arLabels> amb valors <arValues> (fila just a sota)
'@Documentation Check si hi ha una taula amb columnes <arLabels> amb valors <arValues> (fila just a sota)
Public Function checkTableValues( ByRef test_object, ByVal arLabels, ByVal arValues )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
	Dim objTD, objTRValues, objTDValue, regExpLabel, regExpValue, iPos, sResum, micValue, sTemp

	Set objTD = checkLine( test_object, arLabels, 0, -3 )
	If Not( objTD is Nothing ) Then
		Set objTRValues = objTD.parentNode.nextSibling
		If Not( objTRValues is Nothing ) Then
			Set objTDValue = objTRValues.firstChild
			Set regExpLabel = New RegExp
			Set regExpValue = New RegExp
			iPos = 0
			sResum = "LLEGENDA: col = 'Columna_Trobada' [Columna_Sol·licitada]  -->  valor = 'Valor_Trobat' [Valor_Sol·licitat]" & Chr(13)
			micValue = micPass
			Do until (objTD is Nothing)
				regExpLabel.Pattern = arLabels(iPos)
				regExpValue.Pattern = arValues(iPos)
				If ( regExpLabel.Test( objTD.innerText ) ) Then
					If ( regExpValue.Test( objTDValue.innerText ) ) Then
						sResum = sResum & "OK: col = '" & arLabels(iPos) & "' [" & arLabels(iPos) & "]  -->  valor = '" & objTDValue.innerText & "' [" & arValues(iPos) & "]" & Chr(13)
					Else
						micValue = micFail
						sResum = sResum & "ERROR: col = '" & arLabels(iPos) & "' [" & arLabels(iPos) & "]  -->  valor = '" & objTDValue.innerText & "' [" & arValues(iPos) & "]" & Chr(13)
					End If
					iPos = iPos + 1
				End If
				Set objTD = objTD.nextSibling
				Set objTDValue = objTDValue.nextSibling
			Loop
		Else
			micValue = micFail
			sResum = "No s'ha trobat la fila de valors."
			sTemp = test_object.CheckProperty( "name", "ErrorForçat", 1 )
		End If
	Else
		micValue = micFail
		sResum = "No s'ha trobat la fila de capçaleres."
	End If
	Reporter.ReportEvent micValue, "Valors columnes", sResum
End Function
RegisterUserFunc "Frame", "checkTableValues", "checkTableValues"

'@Description DEPRECATED Check si existeix la línia de dades introduïda com array d'expressions regulars
'@Documentation Check si existeix la línia de dades <arLabels>
Public Function checkLineExist( ByRef test_object, ByVal arLabels, ByVal posLabelSearch, ByVal retValues )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
'posLabelSearch: element de l'array que a d'usar per la primera cerca. Inici per 0. Si dona igual posar 0.
'retValues: s'indica que retornar i si hi ha Pass/Fail al report:
'			-2 -> retorna True/False. Treu Pass/Fail al report. Verifica amb innerHTML de les cel·les.
'			-1 -> retorna True/False. Treu Pass/Fail al report. Verifica amb innerText de les cel·les.
'			Num -> El valor trobat per la posició Num segons arLabels. Treu Pass/Fail al report
'			1000+Num -> El valor trobat per la posició Num segons arLabels. Treu Done al report
'Exemple: checkLineExist( Array("\(A\.[0-9]\)", "HP", "PD- PISOS.*", "150"), 2, -1 )
'				checkLineExist( Array("\(A\.[0-9]\)", "HP", "PD- PISOS.*", "^150,00$"), 2, 0 )	
'				checkLineExist( Array("\(A\.[0-9]\)", "HP", ".*", "PD- PISOS.*", "^150,00$"), 2 ) --> Amb ".*"  s'indica que entre "HP" i "PD-" ha d'haver-hi almenys una columna
	checkLineExist = checkLine( test_object, arLabels, posLabelSearch, retValues )
End Function
RegisterUserFunc "Frame", "checkLineExist", "checkLineExist"

'@Description Prem el Link amb el text (RegExp) indicat
'@Documentation Prem el Link amb el text <sText> (RegExp)
Public Function clickLinkByText( ByRef test_object, ByVal sText )
'sText: com Expresssió Regular
	Dim objDesc, objLink
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "A"
	objDesc("innertext").Value = sText
	objDesc("index").Value = 0
	Set objLink = test_object.Link( objDesc )
	If objLink.Exist Then
		clickLinkByText = True
		objLink.Click
	Else
		Reporter.ReportEvent micFail, sText & " - Link NO TROBAT", "No s'ha trobat el Link amb text '" & sText & "'"
		clickLinkByText = False
	End If
	Set objLink = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "clickLinkByText", "clickLinkByText"

'@Description Prem l'imatge <sImgFile> que es troba en la linía amb expressió regular <sText>
'@Documentation Prem l'imatge <sImgFile> que es troba en la linía amb expressió regular <sText>
Public Function clickImageByLineText( ByRef test_object, ByVal imgFile, ByVal sText )
'sImgFile:	1) Nom fitxer de l'imatge. Ex: lupa.gif
'				2) Array amb nom fitxer de l'imatge i núm (la 1ª és 0) de imatge en la línia. Ex: Array( "lupa.gif", 1) = La segona lupa.gif de la línia
'sText: com Expresssió Regular
	Dim objDesc, objTR, objDesc2, objImg, sImgFile, iImgFile
	If ( vartype(imgFile)=8 ) Then 'Texto
		sImgFile = imgFile
		iImgFile = 0
	Else 'Array
		sImgFile = imgFile(0)
		iImgFile = imgFile(1)
	End If
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "TR"
	objDesc("innertext").Value = sText
	objDesc("index").Value = 0
	Set objTR = test_object.WebElement( objDesc )
	If objTR.Exist Then
		Set objDesc2 = Description.Create()
		objDesc2("html tag").Value = "IMG"
		objDesc2("file name").Value = sImgFile
		objDesc2("index").Value = iImgFile
		Set objImg = objTR.Image( objDesc2 )
		If objImg.Exist Then
			clickImageByLineText = True
			objImg.Click
		Else
			Reporter.ReportEvent micFail, sImgFile & " - Imatge NO TROBADA", "No s'ha trobat la imatge  '" & sImgFile & "'(" & Cstr(iImgFile) & ") colindant al text '" & sText & "'"
			clickImageByLineText = False
		End If
	Else
		Reporter.ReportEvent micFail, sImgFile & " - Imatge NO TROBADA", "No trobat el text colindant  '" & sText & "' a la imatge '" & sImgFile & "'(" & Cstr(iImgFile) & ")"
		clickImageByLineText = False
	End If
	Set objImg = Nothing
	Set objDesc2 = Nothing
	Set objTR = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "clickImageByLineText", "clickImageByLineText"

'@Description Prem el botó <btn> que es troba en la linía amb expressió regular <sText>
'@Documentation Prem el botó <btn> que es troba en la linía amb expressió regular <sText>
Public Function clickButtonByLineText( ByRef test_object, ByVal btn, ByVal sText )
'btn:	1) Número de botó en la línia amb qualsevol text. El primer es 0.
'		  2) Array amb text del botó i núm (la 1ª és 0) de botó en la línia. Ex: Array( "Baixar", 0) = El primer botó amb text "Baixar" de la línia
'sText: com Expresssió Regular
	Dim objDesc, objTR, objDesc2, objBtn, sBtn, iBtn
	If ( vartype(btn)=2 ) Then 'Numero
		sBtn = ".*"
		iBtn = btn
	Else 'Array
		sBtn = btn(0)
		iBtn = btn(1)
	End If
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "TR"
	objDesc("innertext").Value = sText
	objDesc("index").Value = 0
	Set objTR = test_object.WebElement( objDesc )
	If objTR.Exist Then
		Set objDesc2 = Description.Create()
		objDesc2("html tag").Value = "INPUT"
		objDesc2("type").Value = "button"
		objDesc2("value").Value = sBtn
		objDesc2("index").Value = iBtn
		Set objBtn = objTR.WebButton( objDesc2 )
		If objBtn.Exist Then
			clickButtonByLineText = True
			objBtn.Click
		Else
			Reporter.ReportEvent micFail, "Botó NO TROBAT", "No s'ha trobat el botó de text '" & sBtn & "' i número  '" & iBtn & "' colindant al text '" & sText & "'"
			clickButtonByLineText = False
		End If
	Else
		Reporter.ReportEvent micFail, "Botó NO TROBAT", "No trobat el text colindant  '" & sText & "' al botó"
		clickButtonByLineText = False
	End If
	Set objBtn = Nothing
	Set objDesc2 = Nothing
	Set objTR = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "clickButtonByLineText", "clickButtonByLineText"

'@Description Selecciona el CheckBox número <iNumCheckBox> (primer 0) que es troba en la linía de dades <sText> (RegExp)
'@Documentation Selecciona el CheckBox número <iNumCheckBox> (primer 0) que es troba en la linía de dades <sText> (RegExp)
Public Function selectCheckBoxByLineText( ByRef test_object, ByVal iNumCheckBox, ByVal sText )
'iCheckBox: número de CheckBox en la línia. El primer es 0.
'sText: com Expresssió Regular
	Dim objDesc, objTR, objDesc2, objCB
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "TR"
	objDesc("innertext").Value = sText
	objDesc("index").Value = 0
	Set objTR = test_object.WebElement( objDesc )
	If objTR.Exist Then
		Set objDesc2 = Description.Create()
		objDesc2("html tag").Value = "INPUT"
		objDesc2("type").Value = "checkbox"
		objDesc2("index").Value = iNumCheckBox
		Set objCB = objTR.WebCheckBox( objDesc2 )
		If objCB.Exist Then
			selectCheckBoxByLineText = True
			objCB.Set "ON"
		Else
			Reporter.ReportEvent micFail, "CheckBox NO TROBAT", "No s'ha trobat el CheckBox número  '" & iNumButton & "' colindant al text '" & sText & "'"
			selectCheckBoxByLineText = False
		End If
	Else
		Reporter.ReportEvent micFail, "CheckBox NO TROBAT", "No trobat el text colindant  '" & sText & "' al CheckBox"
		selectCheckBoxByLineText = False
	End If
	Set objCB = Nothing
	Set objDesc2 = Nothing
	Set objTR = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "selectCheckBoxByLineText", "selectCheckBoxByLineText"

'@Description Selecciona el Radio número <iNumRadio> (només 0) que es troba en la linía de dades <sText> (RegExp)
'@Documentation Selecciona el Radio número <iNumRadio> (només 0) que es troba en la linía de dades <sText> (RegExp)
Public Function SelectRadioByLineText( ByRef test_object, ByVal iNumRadio, ByVal sText )
'iNumRadio: número de Radio en la línia. El primer es 0. NOMËS CODIFICAT PER OPCIÓ 0.
'sText: com Expresssió Regular
	Dim objDesc, objTR, objDesc2, objRadio
	Set objDesc = Description.Create()
	objDesc("html tag").Value = "TR"
	objDesc("innertext").Value = sText
	objDesc("index").Value = 0
	Set objTR = test_object.WebElement( objDesc )
	If objTR.Exist Then
		Dim sTR_HTML, iPosTypeRadio, iPosFinTag, iPosValue, iPosValueFin, sRadioValue
		sRadioValue = ""
		'A partir del outerHTML de la fila
		sTR_HTML = objTR.GetROProperty("outerhtml")
		'Obté la posició del type=radio
		iPosTypeRadio = InStr( sTR_HTML, "type=radio" )  '<INPUT type=radio value=128298812 name=radioEper>
		If ( iPosTypeRadio>0 ) Then  'Si troba el "type=radio"
			iPosFinTag = InStr( iPosTypeRadio, sTR_HTML, ">" )  'Busca final del Tag <INPUT type=radio ...>
			iPosValue = InStr( iPosTypeRadio, sTR_HTML, "value=" )  'Busca el "value=" del type=radio
			If( iPosValue<iPosFinTag ) Then  'El value trobat és vàlid si està abans del final de Tag <INPUT type=radio ...>
				iPosValueFin = InStr( iPosValue, sTR_HTML, " " )  'Suposa que després de value=NNNN hi ha un espai
				sRadioValue = Mid( sTR_HTML, iPosValue + 6, iPosValueFin - (iPosValue +6) )  'Recupera en Value, entre "value=" i l'espai
			End If
		End If
		'Msgbox sValue
		Set objDesc2 = Description.Create()
		objDesc2("html tag").Value = "INPUT"
		objDesc2("type").Value = "radio"
		objDesc2("index").Value = iNumRadio
		Set objRadio = objTR.WebRadioGroup( objDesc2 )
		If objRadio.Exist And Len(sRadioValue)<>0Then
			SelectRadioByLineText = True
			'Msgbox objRadio.GetROProperty("value")
			'Msgbox objRadio.GetROProperty("checked"), 0, "Checked"
			'Msgbox objRadio.GetROProperty("all items")
			'objRadio.Select objRadio.GetROProperty("value")
			objRadio.Select sRadioValue
		Else
			Reporter.ReportEvent micFail, "Radio NO TROBAT", "No s'ha trobat el Radio número  '" & iNumButton & "' colindant al text '" & sText & "'"
			SelectRadioByLineText = False
		End If
	Else
		Reporter.ReportEvent micFail, "Radio NO TROBAT", "No trobat el text colindant  '" & sText & "' al Radio"
		SelectRadioByLineText = False
	End If
	Set objRadio = Nothing
	Set objDesc2 = Nothing
	Set objTR = Nothing
	Set objDesc = Nothing
End Function
RegisterUserFunc "Frame", "SelectRadioByLineText", "SelectRadioByLineText"

'@Description: Passem un nombre amb decimals informats. si després de la coma hi ha 00 elimina la coma i els zeros i els punts dels milers
'@Description Devuelve una cadena con la fecha que se le pase normalizada según ISO 8601.
' Concretamente, la fecha del calendario en representación reducida en su representación completa (YYYYMMDD).
'@Documentation http://es.wikipedia.org/wiki/ISO_8601
' Antes se estaban extrayendo subcadenas de la cadena que representa now. Pero esto funciona o no según la máquina.
' Parámetro:
'	- LaFecha: Cualquier valor que exprese una fecha en un formato inteligible
Public Function xifraSeseComaSiNoDecimals(xifra)
	Dim splitXifra, stringXifraSenseMilers

	stringXifraSenseMilers = Split (xifra,".")
	
	For Iterator = 0 To Ubound(stringXifraSenseMilers) Step 1
		finalString = finalstring&stringXifraSenseMilers(Iterator)
	Next
	
	splitXifra = split(finalString,",")
	If splitXifra(1) = "00" Then
		xifraSeseComaSiNoDecimals = splitXifra(0)
	Else
		xifraSeseComaSiNoDecimals = Replace (finalString,",",".")
	End If

End Function

' Exponemos una variable ¿o constante? privada porque es de interés público
Public Function DirectorioParametros ()

	DirectorioParametros = sDirParameterFile

End Function



'*****************************************************************************************************************
'************************************************** PRIVATE **************************************************
'*****************************************************************************************************************
Const QTP_FILES = "C:\QTP_FILES\"
Dim sDirParameterFile, sDirPreTestDataFile
sDirParameterFile = QTP_FILES & "PARAMETERS\"
sDirPreTestDataFile = QTP_FILES & "PRETESTDATA\"
' ### Se necesitan estos directorios fuera de esta librería
'Private sDirParameterFile, sDirPreTestDataFile
'sDirParameterFile = "C:\QTP_FILES\PARAMETERS\"
'sDirPreTestDataFile = "C:\QTP_FILES\PRETESTDATA\"



Private Sub writeFile( sFileName, sText, iomode )
'sFileName: nom COMPLERT (Drive: i Path) del fitxer de text que contindra el text
'sText: text a incloure al fitxer
'iomode: valors possibles 2 = Open a file for writing   /   8 = Open a file and write to the end of the file
	Dim objFSO, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile( sFileName, iomode, True )
	objTextFile.WriteLine( sText )
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO = Nothing
End Sub

Private Function readFile( sFileName, bLine )
'sFileName: nom COMPLERT (Drive: i Path) del fitxer de text del que llegira el text
'bLine: True - llegeix la primera linia    /    False - llegeix tot el fitxer
'Retorna el text llegit del fitxer indicat
	Dim objFSO, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile( sFileName, 1 )
	If bLine Then
		readFile = objTextFile.ReadLine
	Else
		readFile = objTextFile.ReadAll
	End If
	objTextFile.Close
	Set objTextFile = Nothing
	Set objFSO = Nothing
End Function

Private Function checkLine( ByRef test_object, ByVal arLabels, ByVal posLabelSearch, ByVal retValues )
'arLabel: ATENCIO contingut en Expressio Regular --> Cal escapar \ ^ $ * + ? . ( ) [ ] { } | ...
	'La coincidencia és en el mateix ordre de l'array, però no cal indicar totes les columnes de dades
	'Si un dels valors es precedix de _SEARCH_, s'usarà aquest per la primera cerca, independentment de l'indicat a posLabelSearch. INDICAR-NE NOMÉS UN. Es per poder-ho configurar en l'array
	'Si hi ha valors precedits de _RETURN_, es retornaran per la funció. Se'n poden indicar tants com es vulgui
	'Si hi ha valors precedits de _CASEIN_, es comparara patró amb Case Insensitive. Se'n poden indicar tants com es vulgui
	'Ordre paraules clau: sempre _SEARCH_ primer -->  _SEARCH__RETURN__CASEIN_
'posLabelSearch: element de l'array que a d'usar per la primera cerca. Inici per 0. Si dona igual posar 0.
'retValues: s'indica que retornar i si hi ha Pass/Fail al report:
'			-3 -> retorna el primer TD de la fila que compleix
'			-2 -> retorna True/False. Treu Pass/Fail al report. Verifica amb innerHTML de les cel·les.
'			-1 -> retorna True/False. Treu Pass/Fail al report. Verifica amb innerText de les cel·les.
'			Num -> El valor trobat per la posició Num segons arLabels. Treu Pass/Fail al report
'			1000+Num -> El valor trobat per la posició Num segons arLabels. Treu Done al report
'Exemple: checkLine( Array("\(A\.[0-9]\)", "HP", "PD- PISOS.*", "150"), 2, -1 )
'				checkLine( Array("\(A\.[0-9]\)", "HP", "PD- PISOS.*", "^150,00$"), 2, 0 )	
'				checkLine( Array("\(A\.[0-9]\)", "HP", ".*", "PD- PISOS.*", "^150,00$"), 2 ) --> Amb ".*"  s'indica que entre "HP" i "PD-" ha d'haver-hi almenys una columna
	Dim objDesc, ListObjTD, maxTD, iTD, objTD, maxArLabels, iArLabels, objTDObject, objTDObject2, objTDFirst, regEx, reportResult, tstart, bWait, sTemp
	ReDim arTDs(0)
	Dim maxArTDs, iArTDs, iColRETURN
	maxArLabels = Ubound(arLabels)
	ReDim arLabelsFound(maxArLabels)
	ReDim arLabelsReturn(maxArLabels)
'	ReDim arLabelsCaseIn(maxArLabels)
	iColRETURN = 0
	'Comprova si un dels valors es precedix de _SEARCH_ o de _RETURN_
' ***	Está fallando en [COMPENSACIONS]-01-01: No coinciden los tipos: 'arLabels(...)'
'		Cuando haya tiempo, ya veremos si se usa y cómo corregirlo.
	For iArTDs = 0 To maxArLabels
		If ( Left( arLabels(iArTDs), 8) = "_SEARCH_" ) Then
			arLabels(iArTDs) = Mid( arLabels(iArTDs), 9 )
			posLabelSearch = iArTDs
		End If
		If ( Left( arLabels(iArTDs), 8) = "_RETURN_" ) Then
			arLabels(iArTDs) = Mid( arLabels(iArTDs), 9 )
			arLabelsReturn( iColRETURN ) = iArTDs  'En acabat arLabelsReturn conté les columnes que caldrà retornar
			iColRETURN = iColRETURN + 1
		End If
'		If ( Left( arLabels(iArTDs), 8) = "_CASEIN_" ) Then
'			arLabels(iArTDs) = Mid( arLabels(iArTDs), 9 )
'			arLabelsCaseIn(iArTDs)=True
'		Else
'			arLabelsCaseIn(iArTDs)=False
'		End If
	Next
' ***
	'Defineix objDesc per cerca TOTS els elements TD que contenen el Text/HTML indicat per la primera cerca
	Set objDesc = Description.Create()
	objDesc("micclass").Value = "WebElement"
	objDesc("html tag").Value = "TD"
	If ( retValues=-2 ) Then
		objDesc("innerhtml").Value = arLabels(posLabelSearch)
	Else
		objDesc("innertext").Value = arLabels(posLabelSearch)
	End If
	Set regEx = New RegExp
	bWait = False
	tstart = Now
	Do
		If (bWait) Then 'Excepte la primera vegada després espera 5 segons entre recomprobacions
			wait 5
		End If
		'Cerquem tots els TD's amb contingut indicat en posLabelSearch de arLabels
		iArLabels = 0
		Set ListObjTD = test_object.ChildObjects(objDesc)
		maxTD = ListObjTD.Count()
		'Recorrem tots els trobats per verificar si algun compleix tots els valors
		For iTD = 0 to maxTD-1
			Set objTD = ListObjTD(iTD)
			'Obtenim el primer element de la fila on està el TD localitzat
			Set objTDObject = objTD.Object
			Set objTDObject2 = objTDObject.previousSibling
			Do until (objTDObject2 is Nothing)
				Set objTDObject = objTDObject2
				' Set objTDObject2 = objTDObject.previousSibling
			Loop
			Set objTDFirst = objTDObject
			'Volquem tota la fila en array per falicitar verificació
			iArTDs = 0
			Do until (objTDObject is Nothing)
				ReDim Preserve arTDs(iArTDs)
				If ( retValues=-2 ) Then
					arTDs(iArTDs) = Trim(objTDObject.innerHTML)
				Else
					arTDs(iArTDs) = Trim(objTDObject.innerText)
				End if
				iArTDs = iArTDs + 1
				Set objTDObject = objTDObject.nextSibling
			Loop
			maxArTDs = ubound(arTDs)
			'Verifiquem si tota arLabels està continguda en l'array obtingut
			iArLabels = 0
			regEx.Pattern = arLabels(iArLabels)
			For iArTDs = 0 to maxArTDs
				If ( regEx.Test( arTDs(iArTDs) ) ) Then
					arLabelsFound(iArLabels) = arTDs(iArTDs)  'Guardo valor de coincidència
					iArLabels = iArLabels + 1
					If ( iArLabels>maxArLabels ) Then
						'Si es més gran es que ja s'han trobat totes les coincidències
						Exit For
					End If
					regEx.Pattern = arLabels(iArLabels)
				End If
			Next
			If ( iArLabels>maxArLabels ) Then
				'Si es més gran es que ja s'han trobat totes les coincidències, no cal verificar altres TDs
				Exit For
			End If
		Next
		bWait = True
	Loop Until (iArLabels>maxArLabels Or secDifference( tstart, Now) > tWait )  'Repeteix si no s'ha trobat fins exhaurir tWait
	ReDim arColReturn(0)
	If ( iArLabels>maxArLabels ) Then 'Si es més gran es que ja s'han trobat totes les coincidències
		If ( retValues<0 And iColRETURN=0 ) Then  'retValues = -1 o -2 AND no hi ha columnes configurades per ser retornades, iColRETURN = 0
			 If (retValues=-3) Then
				set checkLine = objTDFirst
			Else
				checkLine = True
			End If
			reportResult = micPass
		Else
		'arLabelsReturn( iColRETURN )
			If ( iColRETURN=0 ) Then  'No hi ha columnes configurades per ser retornades, iColRETURN = 0
				If ( retValues<1000) Then
					checkLine =  arLabelsFound( retValues )
				Else
					checkLine =  arLabelsFound( retValues-1000 )
				End If
			Else
				For iArTDs = 0 To iColRETURN-1
					ReDim Preserve arColReturn(iArTDs)
					arColReturn(iArTDs) = arLabelsFound( arLabelsReturn(iArTDs) )
				Next
				checkLine = arColReturn
			End If
			If ( retValues<1000) Then
				reportResult = micPass
			Else
				reportResult = micDone
			End If
        End If
		'Si es més gran es que s'han trobat totes les coincidències
' ***	Está fallando en [COMPENSACIONS]-01-01: No coinciden los tipos: 'arLabels(...)'
'		Cuando haya tiempo, ya veremos si se usa y cómo corregirlo.
		Reporter.ReportEvent reportResult, arLabels(posLabelSearch), "S'ha trobat la fila de dades [" & join( arLabels, "]  [" ) & "]" & Chr(13) & Chr(13) & _
																		 " en la fila [" & join( arTDs, "]  [" ) & "]" & Chr(13) & Chr(13) & _
																		 " amb valors trobats [" & join( arLabelsFound, "]  [" ) & "]"
' ***
	Else  'NO s'ha trobat
		If ( retValues<1000) Then
			reportResult = micFail
		Else
			reportResult = micDone
		End If
		Reporter.ReportEvent reportResult, arLabels(posLabelSearch), "NO s'ha trobat la fila de dades [" & join( arLabels, "]  [" ) & "]"
		sTemp = test_object.CheckProperty( "name", "ErrorForçat", 1 )
		If ( iColRETURN=0 ) Then  'No hi ha columnes configurades per ser retornades, iColRETURN = 0
			If (retValues=-3) Then
				Set checkLine = Nothing
			Else
				checkLine = False
			End If
		Else
			For iArTDs = 0 To iColRETURN-1
				ReDim Preserve arColReturn(iArTDs)
				arColReturn(iArTDs) = ""
			Next
			checkLine = arColReturn
		End If
	End If
	Set regEx = Nothing
	Set objTD = Nothing
	Set ListObjTD = Nothing
	Set objDesc = Nothing
End Function

' v1.03 FechaNormal, Fecha, Hoy y HoraNormal vienen de ComunGaudi.qfl, pero se mejoran aquí.
Const SIN_SEPARADOR = ""

'@Description Devuelve una cadena con la fecha que se le pase normalizada según ISO 8601.
' Concretamente, la fecha del calendario en representación reducida en su representación completa (YYYYMMDD).
'@Documentation http://es.wikipedia.org/wiki/ISO_8601
' Antes se estaban extrayendo subcadenas de la cadena que representa now. Pero esto funciona o no según la máquina.
' Parámetro:
'	- LaFecha: Cualquier valor que exprese una fecha en un formato inteligible
Public Function FechaNormal (ByRef LaFecha)

	FechaNormal = Join (Array (Year (LaFecha), DosDigitos (Month (LaFecha)), DosDigitos (Day (LaFecha))), SIN_SEPARADOR)
' ###
'	Dim x																				' Para consultar solo una vez cada parte
'
'	FechaNormal = year (LaFecha)
'	x = month (LaFecha)
'	If 10 > x Then
'		FechaNormal = FechaNormal & "0"
'	End If
'	FechaNormal = FechaNormal & x
'	x = day (LaFecha)
'	If 10 > x Then
'		FechaNormal = FechaNormal & "0"
'	End If
'	FechaNormal = FechaNormal & x

End Function



'@Description Devuelve una cadena con la fecha que se le pase en el formato habitual (DD/MM/YYYY).
' Parámetro:
'	- LaFecha: Cualquier valor que exprese una fecha en un formato inteligible
Public Function Fecha (ByRef LaFecha)

	Fecha = Join (Array (DosDigitos (Day (LaFecha)), DosDigitos (Month (LaFecha)), Year (LaFecha)), "/")
' ###
'	Const SEPARADOR = "/"
'	Dim x																			   	' Para consultar solo una vez cada parte
'
'	x = day (LaFecha)
'	If 10 > x Then
'		Fecha = "0" & x
'	Else
'		Fecha = x
'	End If
'	Fecha = Fecha & SEPARADOR
'	x = month (LaFecha)
'	If 10 > x Then
'		Fecha = Fecha & "0"
'	End If
'	Fecha = Fecha & x & SEPARADOR & year (LaFecha)

End Function



' Devuelve la fecha actual en formato DD/MM/YYYY.
Public Function Hoy ()

	Hoy = Fecha (Now)

End Function



'@Description Devuelve una cadena con la hora que se le pase normalizada según ISO 8601. Concretamente la hora en representación reducida.
'@Documentation http://es.wikipedia.org/wiki/ISO_8601
' Antes se estaban extrayendo subcadenas de la cadena que representa now. Pero esto funciona o no según la máquina.
' Parámetro:
'	- LaHora: Cualquier valor que exprese una hora en un formato inteligible
Public Function HoraNormal (ByRef LaHora)

	HoraNormal = Join (Array (DosDigitos (hour (LaHora)), DosDigitos (minute (LaHora)), DosDigitos (second (LaHora))), SIN_SEPARADOR)
' ###
'	Dim x																				' Para consultar solo una vez cada parte
'
'	x = hour (LaHora)
'	If 10 > x Then
'		HoraNormal = "0" & x
'	Else
'		HoraNormal = x
'	End If
'	x = minute (LaHora)
'	If 10 > x Then
'		HoraNormal = HoraNormal & "0"
'	End If
'	HoraNormal = HoraNormal & x
'	x = second (LaHora)
'	If 10 > x Then
'		HoraNormal = HoraNormal & "0"
'	End If
'	HoraNormal = HoraNormal & x

End Function



' Es una versión previa de NumericString, más general. Pero esta es más rápida.
Private Function DosDigitos (ByRef Dato)

	If 10 > Dato Then
		DosDigitos = "0" & Dato
	Else
		DosDigitos = Dato
	End If

End Function



' Proporciona el primer día hábil igual o posterior al proporcionado como parámetro.
Public Function PrimerDiaHabil (ByRef Fecha)

'	Dim sa, 

'	PrimerDiaHabil = 

End Function



'@Description Devuelve una cadena con el número indicado con los ceros no significativos necesarios para llegar al número de dígitos indicado
' Parámetros:
'	- Dato: El número, en formato cadena o numérico
'	- NumDigitos: Cuántos dígitos queremos que tenga la cadena resultante.
' Si el dato no es numérico, devolverá Null.
' Si el dato tiene más dígitos que los deseados, devolverá Null.
' En ambos casos informará en el informe de resultados.
Public Function NumericString (Dato, NumDigitos)

	Dim Numero, CuantosFaltan

	NumericString = Null
	If IsNumeric (Dato) Then
		Numero = CStr (Dato)
		CuantosFaltan = NumDigitos - Len (Numero) 
		If CuantosFaltan > 0 Then
			NumericString = String (CuantosFaltan, "0") & Numero
		ElseIf CuantosFaltan = 0 Then
			NumericString = Numero
		Else
			Reporter.ReportEvent micWarning, "¿Error de programación?",_
								 "NumericString () ha recibido un numero demasiado grande:" & vbCR &_
								 "Dato = " & CStr (Dato) & vbCR & "NumDigitos = " & NumDigitos
		End If
	Else
		Reporter.ReportEvent micWarning, "¿Error de programación?", "NumericString () ha recibido un dato no numérico: " & CStr (Dato)
	End If

End Function



'@description Gestiona la respuesta del sistema operativo a las órdenes de impresión de las aplicaciones.
' Espera las diferentes respuestas posibles _conocidas_, de forma simultánea pero minimizando el tiempo de espera.
'			Además, ahora contemplamos también el fallo de Open Office
' ***	Habría que ver la forma de recuperarse al detectar el fallo del Open Office,
'		que afectará a todos los tests que se ejecuten.
Public Sub ResponderImpresion ()

	Dim ERROR_OPEN_OFFICE
	Dim Dialogo, Mensaje, Estado, Espera
	
	ERROR_OPEN_OFFICE = Join (Array ("S'ha produït un error al processar la petició d'impressió",_
									 "Error:  No hay instancias Open Office disponibles" _
									 ), vbLF)											' Las ctes VBScript no admiten concatenaciones
	Estado = Null
	Espera = 0
	Do
		Wait 10
		If Browser ("B").Dialog ("nativeclass:=#32770").Exist (10) Then					' *** Usa Browser ("B"), definido en generals.tsr	>:/
			Set Dialogo = Browser ("B").Dialog ("nativeclass:=#32770")
			Titulo = Dialogo.GetROProperty ("regexpwndtitle")
			If Titulo = "Imprimir" Then										   			' Lo normal, el diálogo para imprimir
				With Dialogo.WinButton ("text:=Cancelar")
					.WaitProperty "visible", True, 60									' v1.20 A veces el diálogo tarda demasiado
					'.Click																' v1.44
					Dialogo.Close
				End With
			ElseIf Titulo = "Microsoft Internet Explorer" OR _
				   Titulo = "Mensaje de página web" Then								' IE6 o IE8
				Mensaje = Dialogo.Static ("window id:=65535").GetROProperty ("text")	' Cuando no hay instancia disponible de Open Office o
				If Mensaje = ERROR_OPEN_OFFICE Then
					Estado = micDone
				Else
					Estado = micWarning													' o cualquier otro evento inesperado
				End If
				Dialogo.WinButton ("text:=Aceptar").Click
			Else
				Mensaje = ""
				Estado = micWarning
			End If
			If not IsNull (Estado) Then
				Reporter.ReportEvent Estado , "Impresión",_
									 Join (Array ("La aplicación respondió el mensaje ", Titulo & ": ", Mensaje), vbCr)
			End If
			Exit do
		ElseIf Window ("Adobe Reader").Dialog ("Imprimir").Exist (10) Then				' Esto ¿era? necesario en la factoría de HP en León
			Window ("Adobe Reader").Dialog ("Imprimir").WinButton ("Cancelar").Click
			If Window("Adobe Reader").Exist (0) Then
				Window("Adobe Reader").Close
			End If
			Exit do
		End if
		Espera = Espera + 10
	Loop until Espera > 300

End Sub



' Devuelve el primer entero superior o igual a un número.
Public Function EnteroSuperior (ByRef x)

	Dim ParteEntera : ParteEntera = Fix (x)

	If x = ParteEntera Then
		EnteroSuperior = x
	Else
		EnteroSuperior = ParteEntera + 1
	End If

End Function



Private UltimoPaso																		' Para que el usuario no tenga que ir contando sus pasos
Private NombrePasos																		' y desentenderse incluso del nombre de cada paso
Private PasosNoCensados																	' Para usar NombrePasos si intercalamos otros textos



' Decide cuál es el siguiente paso.
Private Function SiguientePaso ()

	If IsEmpty (UltimoPaso) Then
		UltimoPaso = 1
	Else
		UltimoPaso = UltimoPaso + 1
	End If
'	If isEmpty (PasosNoCensados) Then
'		PasosNoCensados = 0
'	End if
	SiguientePaso = UltimoPaso

End Function


Dim ContadorTemporal : ContadorTemporal = 1

' Forma el título del paso con un número y un texto.
' Se apoya en las variables privadas UltimoPaso, que es un contador, y NombrePasos, que es una lista de cadenas con el nombre de los pasos.
' Parámetros:
'	- Paso: Una cadena con un breve texto descriptivo. O un número. O Null. O incluso.
Private Function NombrePaso (Paso)

	Const SEPARADOR = ". "
	Dim SinTexto

'	NombrePaso = "No funciona el contador " & ContadorTemporal
'	ContadorTemporal = ContadorTemporal + 1

	SinTexto = True
	If IsEmpty (Paso) Then
		Paso = SiguientePaso
	ElseIf IsNull (Paso) Then
		Paso = SiguientePaso
	ElseIf IsNumeric (Paso) Then
		PasosNoCensados = PasosNoCensados + UltimoPaso - Paso
		UltimoPaso = Paso
	Else
		Paso = SiguientePaso & SEPARADOR & Paso
		SinTexto = False
	End If
	If SinTexto Then
		If not IsEmpty (NombrePasos) Then
			Paso = Paso & SEPARADOR & NombrePasos (UltimoPaso - PasosNoCensados - 1)
		Else
			Paso = "Paso " & Paso
		End If
	Else																				' Para poder intercalar textos ajenos a NombrePasos
		PasosNoCensados = PasosNoCensados + 1											' sin salirnos luego del array
	End If
	NombrePaso = Paso

End Function



'
Public Sub NombrarPasos (ByRef Vector)

	Const PASO = "Error del programador"

' LOG COMMENT Reporter.Filter = rfEnableErrorsAndWarnings														' v1.40
	PasosNoCensados = 0
	If IsEmpty (Vector) Then
		Reporter.ReportEvent micFail, PASO, "No has dado valor al parámetro de NombrarPasos ()"
	ElseIf IsNull (Vector) Then
		Reporter.ReportEvent micFail, PASO, "El parámetro de NombrarPasos () no puede ser Null."
	ElseIf VarType (Vector) <> vbArray + vbVariant Then
		Reporter.ReportEvent micFail, PASO, "El parámetro de NombrarPasos () debe ser un vector de cadenas."
	Else
		NombrePasos = Vector
		Exit Sub
	End If
	ExitTest

End Sub



Dim FICHERO_IMAGEN : FICHERO_IMAGEN = QTP_FILES & "Comprobante.png"

'@Description Guarda en el informe de resultados del script los valores esperado y encontrado y una copia de la pantalla,
' según ambos valores coincidan o no. Devuelve True solo si el valor encontrado es el esperado.
' Parámetros:
'	- Ventana:				El objeto de la ventana que estamos probando. Objeto, no su id ni su título.
'	- ResultadoEsperado:	Valor que se espera como resultado de una operación en la aplicación.
'	- ResultadoEncontrado:	Valor que se ha encontrado en la aplicación. Conviene que este y el anterior sean valores mostrados por la propia
'							aplicación y no valores procesados posteriormente por el script.
'	- Paso:					Número del paso dentro del caso de prueba. O un mensaje que describa el paso.
'							Si es Null, se emplea el texto introducido con NombrarPasos()
Public Function InformeResultado (Ventana, ResultadoEsperado, ResultadoEncontrado, Paso)

	Const VALOR_ESPERADO   = "Esperado: "
	Const VALOR_ENCONTRADO = "Encontrado: "
	Dim Detalles

' ###
'	If IsNumeric (Paso) Then
'		Paso = "Paso " & Paso
'	End If

' v1.09	Le pueden llegar True distintos de -1, procedentes de WebElement ().Exist (0).
'		Y al compararlos con un True puesto tal cual, 1 != -1, así que se iba por el Else.
' LOG COMMENT Reporter.Filter = rfEnableErrorsAndWarnings													' v1.40
	If VarType (ResultadoEsperado) = vbBoolean Then										' v1.09
		ResultadoEncontrado = CBool (CInt (ResultadoEncontrado))
	End If
	If ResultadoEsperado = ResultadoEncontrado Then				 						' Si todo va bien, no es precisa imagen alguna
		Reporter.ReportEvent micPass, NombrePaso (Paso), VALOR_ENCONTRADO & ResultadoEncontrado
		InformeResultado = True
	Else
		Detalles = VALOR_ESPERADO & ResultadoEsperado & vbCrLf & vbCrLf & VALOR_ENCONTRADO & ResultadoEncontrado

		If IsNull (Ventana) Then
			Set Ventana = Desktop														' Quizá se quiera ver toda la pantalla
		End if
		If Ventana is Browser ("B") Then												' Copia toda la página, incluso lo oculto
			Pantallazo																	' y es más ligero que una imagen
			Reporter.ReportEvent micFail, NombrePaso (Paso), Detalles
		Else
			Ventana.CaptureBitmap FICHERO_IMAGEN, True
			Reporter.ReportEvent micFail, NombrePaso (Paso), Detalles, FICHERO_IMAGEN
		End If
' ###
'		If Ventana is Browser ("B") Then													' Copia de toda la página
'			Pantallazo
'			Reporter.ReportEvent micFail, NombrePaso (Paso), Detalles
'		Else
'			If IsNull (Ventana) Then
'				Desktop.CaptureBitmap FICHERO_IMAGEN, True								' si algo falla, es mejor ver toda la pantalla
'			Else
'				Ventana.CaptureBitmap FICHERO_IMAGEN, True								' salvo que sepamos claramente qué nos interesa
'			End If
'			Reporter.ReportEvent micFail, NombrePaso (Paso), Detalles, FICHERO_IMAGEN
'		End If
		InformeResultado = False
		ExitTest
	End if

End function



Private Test : Test = 1



'@Description Modificación de InformeResultado para varios test que se ejecutan en una sólo script.
'PRECONDICIÓN: Inicializar el nombre de los pasos usando NombrarPasos antes de llamar InformeResultadoMultiple por primera vez.
'Una vez inicializado NombrarPasos, se pueden intercalar pasos no definidos en NombrarPasos.
'POSTCONDICIÓN: Si un test falla o finaliza se debe pasar al siguiente (si hay).
Function InformeResultadoMultiple (ResultadoEsperado, ResultadoEncontrado, Paso)

	Const VALOR_ESPERADO   = "Esperado: "
	Const VALOR_ENCONTRADO = "Encontrado: "

' LOG COMMENT Reporter.Filter = rfEnableErrorsAndWarnings														' v1.40
	If VarType (ResultadoEsperado) = vbBoolean Then										
		ResultadoEncontrado = CBool (CInt (ResultadoEncontrado))
	End If
	If ResultadoEsperado = ResultadoEncontrado Then				 						' Si todo va bien, la ventana es la prueba pero

		Reporter.ReportEvent micPass, Test& " - " & NombrePaso(Paso),  VALOR_ENCONTRADO & ResultadoEncontrado ' *** & vbCrLf & INSERCION_FICHERO
		InformeResultadoMultiple = True
	Else
		Desktop.CaptureBitmap FICHERO_IMAGEN, True										' si algo falla, es mejor ver toda la pantalla
		Reporter.ReportEvent micFail, Test & " - " & NombrePaso(Paso), VALOR_ESPERADO & ResultadoEsperado & vbCrLf & vbCrLf &_
							VALOR_ENCONTRADO & ResultadoEncontrado, FICHERO_IMAGEN
		InformeResultadoMultiple = False
	End if

	Select Case True
	Case Not(InformeResultadoMultiple), ((UBound(NombrePasos) + 1) + PasosNoCensados = UltimoPaso)
		UltimoPaso = 0
		PasosNoCensados = 0
		Test = Test + 1
	End Select

End Function

'
'@Description Guarda en el informe de resultados del script una copia de la ventana indicada, junto con un texto explicativo
' Parámetros:
'	- Ventana:	El objeto de la ventana que estamos probando. Objeto, no su id ni su título.
'	- Mensaje:	Un texto que explica qué es lo que vemos o hacemos en esa ventana.
'	- Paso:		Número del paso dentro del caso de prueba.
'				Si es Null, se emplea el texto introducido con NombrarPasos()
Public sub InformePaso (Ventana, Mensaje, Paso)

	Dim Texto

'LOG COMMENT Reporter.Filter = rfEnableErrorsAndWarnings														' v1.40
' *** Que se guarden diferentes ficheros y sean accesibles incluso desde QC
	Ventana.CaptureBitmap FICHERO_IMAGEN, True

' ###
'	If IsNumeric (Paso) Then
'		Texto = "Paso " & Paso
'	Else
'		Texto = Paso
'	End If
	Reporter.ReportEvent micDone, NombrePaso (Paso), Mensaje, FICHERO_IMAGEN

End Sub



'@Description Comprueba que un objeto mostrado en pantalla tenga el valor esperado
' Deja un rastro en el informe de resultado.
' Parametros:
'	- Objeto: Un objeto del interfaz de usuario (no su nombre en el repositorio) que se buscará en la aplicación para averiguar su valor.
'			  También acepta una cadena, que comparará tal cual.
'	- Valor:  El que esperamos que tenga la aplicación
' Devuelve True solo cuando lo que muestra la aplicación coincida con lo esperado.
Function Coinciden (Objeto, Valor)

	Dim Resumen
	Dim EnPantalla, Resultado

	Resumen = "Comparando " & Valor
	If IsObject (Objeto) Then															' Es un objeto y lo vemos en pantalla
		If Objeto.Exist (0) Then
			Select Case Objeto.GetROProperty ("micclass")
				Case "WebElement"
					EnPantalla = Objeto.GetROProperty ("innertext")
				Case "WebList"
					EnPantalla = Objeto.GetROProperty ("selection")
				Case "WebEdit"
					EnPantalla = Objeto.GetROProperty ("value")
				Case "Link", "Static"
					EnPantalla = Objeto.GetROProperty ("text")
				Case else
					Reporter.ReportEvent micFail, "ERROR DE PROGRAMACIÓN",_
										 "Este tipo de objeto aún no ha sido integrado en Coinciden(): " & Objeto.ToString
					ExitTest
			End Select
			If Valor = "" and IsNumeric (EnPantalla) Then											' 
				Resultado = CDbl (EnPantalla) = 0
			Else
				Resultado = Valor = EnPantalla
			End If
'			If Resultado Then															' micDone porque a veces tendrán que ser diferentes
'				Reporter.ReportEvent micDone, Resumen, "El valor coincide"
'			Else
			If not Resultado Then
				CopiaPantalla Objeto
				Reporter.ReportEvent micWarning, Resumen,_
									 "Los valores no coinciden: """ & Valor & """ != """ & EnPantalla & """" & vbCR &_
									 "Copia de pantalla en el nodo superior"
			End If
			Coinciden = Resultado
		Else
			Desktop.CaptureBitmap FICHERO_IMAGEN, True
			Reporter.ReportEvent micFail, "Coinciden()",_
								 "Ahora mismo la aplicación no muestra el objeto: " & Objeto.ToString, FICHERO_IMAGEN
			Coinciden = False
		End If
	Else
		Select Case VarType (Objeto)
			Case vbString, vbSingle
				If Valor = Objeto Then
'					Reporter.ReportEvent micDone, Resumen, "El valor coincide"
					Coinciden = True
				Else
					Reporter.ReportEvent micWarning, Resumen,_
										 "Los valores no coinciden: """ & Valor & """ != """ & Objeto & """" & vbCR &_
										 "Copia de pantalla en el nodo superior"
					Coinciden = False
				End If
			Case else
				Reporter.ReportEvent micFail, "ERROR DE PROGRAMACIÓN",_
									 "Coinciden() ha recibido algo que no es un objeto: " & Objeto.ToString
				ExitTest
		End Select
	End If

End Function



'@Description CopiaPantalla sin objeto.
Public Sub Pantallazo ()

	CopiaPantalla Null

'	Dim CapturaAnterior
'
'	On error resume next
'		If Browser ("B").Exist (0) Then
'			CapturaAnterior = EsteUFT.Options.Run.ImageCaptureForTestResults
'			EsteUFT.Options.Run.ImageCaptureForTestResults = "Always"
'			Browser ("PANTALLAZO").Page ("PANTALLA").Frame ("cabecera").WebElement ("Instancia").Click
'			EsteUFT.Options.Run.ImageCaptureForTestResults = CapturaAnterior
'		Else
'			Desktop.CaptureBitmap FICHERO_IMAGEN, True
'			Reporter.ReportEvent micDone, "PANTALLA", "No había un Browser (""B"")", FICHERO_IMAGEN
'		End If
'	On error goto 0

End Sub



'@Description En aplicaciones web, registra en el informe de resultados una copia de la página completa.
' Es una alternativa indirecta a hacer un pantallazo PNG de tan solo la parte visible.
' Parámetros:
'	- Objeto: El del repositorio en que se hará click derecho para que se marque en el informe.
'			  Se puede usar Null si no se sabe dónde marcar, en cuyo caso usará la instancia que soporta la sesión.
Public Sub CopiaPantalla (Objeto)

	Dim CapturaAnterior

	On error resume next
		If Browser ("B").Exist (0) Then
			CapturaAnterior = EsteUFT.Options.Run.ImageCaptureForTestResults
			EsteUFT.Options.Run.ImageCaptureForTestResults = "Always"
			If IsNull (Objeto) Then														' Solo apto para Gaudí
				Browser ("PANTALLAZO").Page ("PANTALLAZO").Frame ("cabecera").WebElement ("Instancia").Click
			ElseIf Objeto.Exist (0) Then
			' Nada, porque Exist() ya queda registrado con una imagen en que el objeto está recuadrado
'				Objeto.Click 1, 1, micRightBtn
'				Conchita.SendKeys "{ESC}"
			Else
				Browser ("PANTALLAZO").Page ("PANTALLAZO").Frame ("cabecera").WebElement ("Instancia").Click
			End If
			EsteUFT.Options.Run.ImageCaptureForTestResults = CapturaAnterior
		Else
			Desktop.CaptureBitmap FICHERO_IMAGEN, True
			Reporter.ReportEvent micDone, "PANTALLAZO", "No había un Browser (""B"")", FICHERO_IMAGEN
		End If
	On error goto 0
	
End Sub


'
'	OBSOLETO: La función integrada Mid() hace lo mismo más eficientemente
'
' Devuelve un array que es el subconjunto de otro array.
' Parámetros:
'	- Origen: El array del que queremos un subconjunto
'	- Inicio: Indice del primer elemento que queremos copiar. Debería ser superior a 0 y menor que el final del array.
'	- Fin:	  Indice del último elemento que queremos copiar. Debería ser superior a 1 y menor que el final del array.
' Lo ideal es que 0 = LBound (Origen) <= Inicio <= Fin <= UBound (Origen).
' Si esto se incumple, la función corrige los parámetros para darle sentido, de modo que nunca falle. It's foolproof code.
Public Function SubArray (ByRef Origen, Inicio, Fin)

	Dim Destino ()																		' El array subconjunto
	Dim i, j, f

	f = UBound (Origen)
	If Inicio > Fin Then
		i = Inicio
		Inicio = Fin
		Fin = i
	End If
	If 0 > Inicio Then
		Inicio = 0
	End If
	If Fin > f Then
		Fin = f
	End If
	j = 0
	ReDim Destino (Fin - Inicio)
	For i = Inicio to Fin
		Destino (j) = Origen (i)
		j = j + 1
	Next


End Function



' Interactúa con el diálogo estándar de guardado de ficheros que Internet Explorer (o cualquier otro programa) puede abrir.
Public Sub DialogoGuardarComo (ByRef RutaCompleta)

	Dim Mensaje

' ### Esto es porque hace falta darle dos veces al botón Guardar.
'Dim x : x = 0
	With Dialog ("Descarga de archivos")
		Do
' ###Wait 0,500
			.Dialog ("Descarga de archivo").WinButton ("Guardar").Click
' ###x = x + 1
		Loop until .Dialog ("Guardar como").Exist (1)
' ###Reporter.ReportEvent micDone, "guardando", x
		With .Dialog ("Guardar como")
			.WinEdit ("Nombre:").Set RutaCompleta
			.WinButton ("Guardar").Click
			If .Dialog("Aviso").Exist(1) Then
				Mensaje = .Dialog("Aviso").Static ("Mensaje").GetROProperty ("text")
				If Mensaje = RutaCompleta & " ya existe." & vbLF & "¿Desea reemplazarlo?" Then
					.Dialog ("Aviso").WinButton ("Sí").Click
				Else
					InformeResultado Dialog ("Descarga de archivos"), "", Mensaje, "Guardando fichero"
				End If
			End If
		End with
	End with

End Sub



Public Const ForReading		= 1															' Esto tendría que definirlo el propio VBScript
Public Const ForWriting		= 2															' Además, en la ayuda de QTP las dos primeras constantes
Public Const ForAppending	= 8															' tienen los valores intercambiados



' Guarda cada una de las cadenas de un vector en una línea diferente de un fichero.
' Parámetros:
'	- Cadenas: Array con las cadenas de texto
'	- Fichero: Nombre del fichero en que se guardarán
Public Sub GuardarEnLineas (ByRef Cadenas, ByRef Fichero)

	Dim fs, f
	Dim i, u

	Set FS = CreateObject ("Scripting.FileSystemObject")
	Set f = fs.OpenTextFile (Fichero, ForWriting, True)
	u = UBound (Cadenas) - 1
	For i = 0 to u
		f.WriteLine Cadenas (i)
	Next
	f.Write Cadenas (i)																	' El último sin CR+LF para el split de LeerLineas
	f.Close
	Set f = Nothing
	Set fs = Nothing

End Sub



'
Public Function LeerLineas (ByRef Fichero)
	Dim fs, f

	Set FS = CreateObject ("Scripting.FileSystemObject")
	Set f = fs.OpenTextFile (Fichero, ForReading, True)
	LeerLineas = Split (f.ReadAll, vbCrLf)
	f.Close
	Set f = Nothing
	Set fs = Nothing
End Function



' Devuelve el NIF correspondiente al número pasado como parámetro.
Public Function CalcularNIF (ByRef DNI)

	If IsNumeric (DNI) Then	
		CalcularNIF = DNI & Mid ("TRWAGMYFPDXBNJZSQVHLCKE", (DNI mod 23) + 1, 1)
	Else
		Reporter.ReportEvent micFail, "CalcularNIF ()", "Para calcular el NIF se ha de emplear un número positivo de hasta 8 dígitos." & vbCR &_
							 "Has introducido como parámetro: """ & DNI & """."
		ExitTest
	End If

End Function

'Description calculem una quantitat de dies laborables i retornem la data concreta. 0- quantitat de díes, 1- ,data d'origen
Public Function calcularDataDiesLaborables(dadesCalcularDiesLaborables)

diaSolicitat = getDataAvui
diesLaborablesRestants = 0
comptador = 0
If UBound(dadesCalcularDiesLaborables) = 1 Then
	diaSolicitat = dadesCalcularDiesLaborables(1)
End If

do While diesLaborablesRestants <> dadesCalcularDiesLaborables(0)
	comptador = comptador + 1
	diaAnalitzat = Weekday(DateAdd("d",comptador,diaSolicitat),2)
	If (diaAnalitzat <> 6 and diaAnalitzat <> 7) Then
		diesLaborablesRestants = diesLaborablesRestants + 1
	End If		
	If comptador = 100 Then
		Exit do
	End If
loop

calcularDataDiesLaborables = DateAdd("d",comptador,diaSolicitat)
End Function 

Public Function diaLaborableOElMesProxim(data)

diaLaborableOElMesProxim = calcularDataDiesLaborables(Array(1,DateAdd("d",-1,data)))

End Function
	
'
Public Function EsNIFvalido (ByRef NIF)

	Dim Longitud, Letra, Numero

	EsNIFvalido = False
	If VarType (NIF) = VbString Then													' Que llegue una cadena
		Longitud = Len (NIF)
		If Longitud >= 2 Then
			If EsUnaLetra (Left (NIF, 1)) Then											' Puede ser de una persona jurídica, el antiguo CIF
				EsNIFvalido = Longitud >= 3												' *** Habría que verificar el carácter de control final
			Else																		' Puede ser el NIF de una persona física
				Numero = Left (NIF, Longitud - 1)
				If IsNumeric (Numero) Then
					EsNIFvalido = Right (CalcularNIF (CLng (Numero)), 1) = UCase (Right (NIF, 1))
				End If
			End if
		End If
	End If

' Esto funciona, pero solo para NIF de personas físicas
'	Const CODIGO_A = 65, CODIGO_Z = 90
'	
'	Dim Longitud, Letra, Codigo, Numero
'
'	If IsEmpty (NIF) Then																' Que llegue una cadena
'		EsNIFvalido = False
'	ElseIf IsNull (NIF) Then
'		EsNIFvalido = False
'	Else
'		EsNIFvalido = False
'		If VarType (NIF) = vbString Then
'			Longitud = Len (NIF)
'			If 9 >= Longitud Then														' Que sea del dominio de los NIF (nº + letra)
'				If Longitud >= 2 Then
'					Letra = UCase (Right (NIF, 1))
'					Codigo = Asc (Letra)
'					If Codigo >= CODIGO_A Then
'						If CODIGO_Z >= Codigo Then
'							Numero = Left (NIF, Longitud - 1)
'							If IsNumeric (Numero) Then
'								If Right (CalcularNIF (CLng (Numero)), 1) = Letra Then	' Y que se correspondan el número y la letra
'									EsNIFvalido = True
'								End If
'							End If
'						End If							
'					End If
'				End If
'			End if
'		End If
'	End If

End Function



' Devuelve un NIF generado aleatoriamente.
Public Function NIF_aleatorio ()

' ###	Dim DNI

	Randomize
	NIF_aleatorio = CalcularNIF (Int (1 + Rnd * 99999998))
' ### Alguna vez ha generado la misma secuencia todas las veces
'	NIF_aleatorio = CalcularNIF (RandomNumber (1, 99999999))

' ***	¿Por qué solo desde el 80 millones en adelante?
'	DNI = RandomNumber (80000000, 99999999)
'	NIF_aleatorio = DNI & Mid ("TRWAGMYFPDXBNJZSQVHLCKE", (DNI mod 23) + 1, 1)

End Function

Public Function calculDataFiVoluntaria(data)
	
	dataSeparada = Split(data,"/")
	If dataSeparada(0) < 16 Then
		dataSeparada(0) = 20
		dataReconstruida = dataSeparada(0)&"/"&dataSeparada(1)&"/"&dataSeparada(2)
		dataCompleta = dateAdd("m",1,dataReconstruida)
		calculDataFiVoluntaria = diaLaborableOElMesProxim(dataCompleta)
	Else
		dataSeparada(0) = 5
		dataReconstruida = dataSeparada(0)&"/"&dataSeparada(1)&"/"&dataSeparada(2)
		 dataCompleta = dateAdd("m",2,dataReconstruida)
		 calculDataFiVoluntaria = diaLaborableOElMesProxim(dataCompleta)
	End If
	
End Function


Public Function EsUnaLetra (x)

	Const CODIGO_A = 65, CODIGO_Z = 90
	Dim Codigo

	EsUnaLetra = False
	If VarType (x) = VbString Then														' Que llegue una cadena
		If Len (x) = 1 Then																' con un solo carácter
			Codigo = Asc (UCase (x))
			If Codigo >= CODIGO_A Then
				If CODIGO_Z >= Codigo Then
					EsUnaLetra = True
				End If
			End If
		End If
	End if

End Function


' Devuelve True solo cuando el objeto es un array vacío, es decir, sin ningún elemento.
Public Function IsAnEmptyArray (ByRef Objeto)

	If IsEmpty (Objeto) Then
		IsAnEmptyArray = False
	ElseIf IsNull (Objeto) Then
		IsAnEmptyArray = False
	ElseIf IsArray (Objeto) Then
		IsAnEmptyArray = UBound (Objeto) = -1
	Else
		IsAnEmptyArray = False
	End If

End Function



Public Function CopiarPDFenNotepad (ByRef aActivex, ByRef Pathname, ByRef filename)
   'Copia el contenido de un .pdf a un fichero(Crl+c, Ctrl+v). 
   'aActivex es el activeX del .pdf
   'Si se quiere abrir un notepad nuevo, el pathname estará vacio y el filename será "Sin título".
   'Si se quiere abrir un notepad existente, el pathname será el directorio donde se encuentra y el filename el nombre del fichero.
   'Call CopiarPDFenNotepad (Browser("B").Page("P").Frame("impresoPdf").ActiveX("Adobe PDF Reader"), "", "Sin título")

Dim DESCRIPCION, filetxt

	DESCRIPCION = "text:="& filename & ".txt - Bloc de notas"  
	filetxt = Pathname & "\" & filename 

	Call waitobject (aActivex.WinObject("AVPageView"),60)
	
	aActivex.WinObject("AVPageView").Type micCtrlDwn + "a" + micCtrlUp
	aActivex.WinObject("AVPageView").Type micCtrlDwn + "c" + micCtrlUp

	systemutil.run "C:\Windows\system32\notepad.exe",filetxt,"","open"

	If Window (DESCRIPCION).Exist (0) Then	 		' Es un mensaje informativo, accesorio, prescindible
		Window (DESCRIPCION).Type micCtrlDwn + "e" + micCtrlUp ' Seleccionamos todo

		Window (DESCRIPCION).Type micDel    

		Window (DESCRIPCION).Type micCtrlDwn + "v" + micCtrlUp

		Window (DESCRIPCION).Type micCtrlDwn + "s" + micCtrlUp

		Window (DESCRIPCION).Activate
		Window (DESCRIPCION).Close
		Window("Bloc de notas").Dialog("Bloc de notas").Activate
		Window("Bloc de notas").Dialog("Bloc de notas").WinButton("Sí").Click
	
	End If 
 
	waitCarregantRepoGenerals

End Function

Function BuscarString(ByRef Filename,ByRef stext)
 'Busca el String stext en en fichero (con el path incluido) que le pasamos.
 'Guarda en la x el valor de la fila anterior, ya que es la fecha actual que se està ejecutando el test.
 'Cuando encuentra un string igual al parecido en la variable x, tiene el valor anterior, 
 'que es el núm de expediente, y lo devuelve al script inicial
Dim x

	Set oRegEx = CreateObject("VBScript.RegExp")
	oRegEx.Pattern = stext' Enter pattern/string you want to search
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.OpenTextFile(Filename,ForReading) ' .OpenTextFile("C:\QTPSchools\abhi.txt", ForReading)
	Do Until oFile.AtEndOfStream
	
		strSearchString = oFile.ReadLine
		Set colMatches = oRegEx.Execute(strSearchString) 
		
		If colMatches.Count > 0 Then
		   Exit Do
		End If
		x=oFile.ReadLine
	Loop
	oFile.Close
	BuscarString = x

End Function



Private Portapapeles																	' El objeto que nos da acceso al portapapeles
Set Portapapeles = CreateObject ("Mercury.Clipboard")									' Accederemos varias veces y es mejor reutilizarlo


' Extrae el contenido de un documento en formato PDF mostrado en pantalla.
' Parámetros:
'	- Objeto: el objeto de clase WinObject que contenga el documento PDF
' Devuelve un vector de cadenas, una por cada línea de las que componen el documento PDF.
' IMPORTANTE: El contenido de este vector casi nunca conserva el orden mostrado en pantalla, siendo preciso un análisis individual.
Function GetContenidoPDF (Objeto)

	Dim Texto
	Do
		With Objeto																		' Seleccionamos y copiamos el texto del PDF y
			.Click
			.Type micCtrlDwn + "a" + micCtrlUp
			.Type micCtrlDwn + "c" + micCtrlUp
		End With
		Texto = Portapapeles.GetText
		If Texto = "" Then																' A veces el PDF no se muestra a tiempo
			Wait 3
		Else
			Exit do
		End If
	Loop																				' lo partimos en trozos, que no serán líneas
	GetContenidoPDF = Split (Texto, vbCRLF)												' tal y como las que se ven en pantalla

End Function
RegisterUserFunc "WinObject", "ContenidoPDF", "GetContenidoPDF"							' Para incrementar la usabilidad




' Guarda el contenido de un PDF mostrado en el correspondiente WinObject a un fichero con el nombre del test actual y una marca temporal
' Parámetros:
'	- Objeto:	Un objeto de clase WinObject que contiene el texto del PDF
'	- Ruta:		Una cadena con la ruta completa en que se desea guardar el fichero.
'				Null o "" si se quiere guardar en C:\QTP_Files\GuardarPDF
Sub GuardarPDF (Objeto, Ruta)

	Dim Intentos																		' A veces no aparece a la primera
	Dim Fichero																			' El nombre del fichero que guardamos

	If IsNull (Ruta) Then
		Ruta = ""
	End If
	If Ruta = "" Then
		Ruta = QTP_FILES & "GuardarPDF"
	End If
	Fichero = Join (Array (Ruta, "\", Environment ("TestName"), "_", FechaNormal (Now), HoraNormal (Now), ".pdf"), "")
	Intentos = 10
	Do
		Objeto.Click 5, 5
		Objeto.Type micCtrlDwn + micShiftDwn + "S" + micShiftUp + micCtrlUp
		If Window ("Adobe Reader").Dialog ("Guardar como").Exist (5) Then	
			Exit do
		Else
			If Intentos = 0 Then
				InformeResultado Desktop, True, False, "No aparece el PDF"
			Else
				Intentos = Intentos - 1
				MostrarMensaje "Otro intento de guardar el PDF"
			End If
		End If
	Loop
	If Intentos = 0 Then
		InformeResultado Desktop, True, False, "No aparece el PDF"
	Else
		With Window ("Adobe Reader").Dialog ("Guardar como")
			.WinEdit ("Nombre").Set Fichero
			.WinButton ("Guardar").Click
		End With
	End If

End Sub
RegisterUserFunc "WinObject", "GuardarPDF", "GuardarPDF"



' Permite aceptar posibles "nuevos" certificados debidos a fallos en la configuración de Internet Explorer.
' Trata los diálogos que muestra IE 6 y las páginas de IE 8 y 9
Sub AceptarCertificado ()

	Const ESPERA = 1

	With Browser ("B")
		If .Exist (ESPERA) Then
			If Environment ("ProductVer") = "9.5" Then
				If .Dialog ("Alerta de seguridad").Exist (0) Then								' IE 6 por QTP 9
					.Dialog ("Alerta de seguridad").WinButton ("Sí").Click
					Exit Sub
				End If
			Else
				With .Page ("Error certificado")												' Es igual que "P". Existe por organización
					If .Exist (0) Then															' IE 8 y 9 por UFT 11.50
						If .WebElement ("Problema").Exist (0) Then
							.Link ("Ir").Click
						End If
					End If
				End With
			End If
		End If
	End With

End Sub



'Private RepositorioSinCargar	: RepositorioSinCargar = True									' Para cargar solo una vez Generals.tsr
'Private Const QC_PATH = "[QualityCenter] Subject\.Repositoris\"								' v1.38n (gg.vbs) Se carga cualquier repositorio
' ###	v1.30
'Private RepositoriosCargados (10)
Private RepositoriosCargados																	' v1.30 Lo crearemos cuando haga falta
Private UltimoRepositorio
Private RUTA_REPOSITORIOS																		' v1.44	Para cualquier versión
Private Const EXTENSION = ".tsr"
Private RUTA_EXPLORER																			' *** Esto es mejor configurarlo en el xml



' Devuelve la ruta completa a Internet Explorer.
' Está pensada para Portal. Gaudí no la necesita.
Function RutaExplorer ()

	If IsEmpty (RUTA_EXPLORER) Then
		LocalizarRepositorios
	End If
	RutaExplorer = RUTA_EXPLORER
	
End Function

' Gaudi emplea repositorios de objetos externos a las acciones.
' En función de la versión de QTP / UFT, la carpeta que los contiene se llama de una u otra forma.
Sub LocalizarRepositorios ()

	' v1.19 Localización estándar de las dos versiones posibles de IE
	Const IE_NATIVO  = "C:\Program Files\Internet Explorer\iexplore.exe"
	Const IE_32_BITS = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"			' En Windows 64 bits pueden instalarse los dos

	Dim fs

	Set fs = CreateObject ("Scripting.FileSystemObject")								' Usamos el IE de 32 bits SIEMPRE
	If fs.FileExists (IE_32_BITS) Then
		RUTA_EXPLORER = IE_32_BITS
	Else
		RUTA_EXPLORER = IE_NATIVO														' o el que haya (que puede ser de 32 ó de 64)
	End If
	Set fs = Nothing
	If RUTA_REPOSITORIOS = "" Then														' *** Definido en GaudiGenerals.vbs
		RUTA_REPOSITORIOS = RUTA_RECURSOS & ".Repositoris\"								' v1.26 En UFT 11.53 y 12.02
	End If
' ### v1.26 Ya sabemos que no habrá más de una versión de UFT ejecutando pruebas y que UFT 11.53 y 12.02 usan las mismas referencias
''		*** v1.17
''		RUTA_EXPLORER  = "C:\Archivos de programa\Internet Explorer\iexplore.exe"		' El único IE en Windows 32 bits y IE de 64 bits
'		Select Case Environment ("ProductVer")
'		Case "9.5"
'			RUTA_REPOSITORIOS = "[QualityCenter] Subject\.Repositoris\"
'			Exit Sub
'		Case "11.00"
'			RUTA_REPOSITORIOS = "[QualityCenter\Resources]"
''			*** v1.17
''			If InStr (Environment ("ProductDir"), " (x86)") <> 0 Then						' W7 64 bits, pero usamos IE 32 bits porque
''				RUTA_EXPLORER = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"		' QTP 11 solo funciona con IE 32 bits
''			End If
'		Case "11.50", "11.53", "12.02"
'			RUTA_REPOSITORIOS = "[ALM\Resources]"
'' ###
''		Case "11.53"
''			RUTA_REPOSITORIOS = "[ALM\Resources]"
'		Case Else
'			Reporter.ReportEvent micFail, "PENDIENTE DE PROGRAMAR",_
'								 "Se está usando una versión de QTP / UFT que requiere actualizar el procedimiento Localizarrepositorios ()."
'			ExitTest
'		End Select
'		RUTA_REPOSITORIOS = RUTA_REPOSITORIOS & " Resources\Asset_Upgrade_Tool_1\.Repositoris\"	' Algo más legible llevaría a cambiar 200+ tests
'	End If

End Sub




Private Function RepositorioSinCargar (ByVal Modulo)

	Dim i, Final : Final = UltimoRepositorio - 1

	If IsEmpty (RepositoriosCargados) Then												' v1.30 Al asociar 2 repositorios (ING + COMPT),
		RepositoriosCargados = Array (Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
	End If																				' se llega aquí con RepositoriosCargados sin declarar
	For i = 0 to Final
		If RepositoriosCargados (i) = Modulo Then
			RepositorioSinCargar = False
			Exit function
		End If
	Next
	RepositorioSinCargar = True

End Function



'@Description Asocia un repositorio de objetos a las acciones del test en ejecución.
' Parámetros:
'	- Modulo: El nombre del módulo cuyo repositorio nos interesa, sin la extensión ".tsr" ni la ruta en QC
Public Sub AsociarRepositorio (ByVal Modulo)

' ### v1.25
'	Dim QTP, Acciones, Repositorios, i
	Dim Acciones, Repositorios, i
	Dim ElRepositorio																	' v1.38n (en GaudiGenerals.vbs)

	LocalizarRepositorios
	If IsEmpty (UltimoRepositorio) Then													' v1.30
		UltimoRepositorio = 0
	End If
' v1.38n
'	If RepositorioSinCargar Then
	If RepositorioSinCargar (Modulo) Then
		ElRepositorio = RUTA_REPOSITORIOS & Modulo & EXTENSION
' ### v1.43
'		If IsEmpty (EsteUFT) Then														' v1.25 Si no lo tenía, añade al test el repositorio
'			Set EsteUFT = getObject ("", "QuickTest.Application")
'		End If
' ###
'		Set QTP = getObject ("", "QuickTest.Application")								' Si no lo tenía, añade al test el repositorio
		Set Acciones = EsteUFT.Test.Actions												' v1.18 de cada acción del test
		For i = 1 to Acciones.Count
			If Acciones (i).Type <> "External" Then
				Set Repositorios = Acciones (i).ObjectRepositories
				If IsEmpty (Repositorios) Then
					Repositorios.Add ElRepositorio
					RepositoriosCargados (UltimoRepositorio) = Modulo
					UltimoRepositorio = UltimoRepositorio + 1
				ElseIf Repositorios.Find (ElRepositorio) = -1 Then
					On error resume next
						Repositorios.Add ElRepositorio
						If Err <> 0 Then
							Reporter.ReportEvent micDone, "LoginUsuari ()",_
												 "No se ha podido añadir el repositorio de objetos " & ElRepositorio &_
												 " a la acción " & Acciones (i).Name & "." & vbCR &_
												 "El error VBScript es " & Err & ": " & Err.Description
							Err.Clear
						Else
							RepositoriosCargados (UltimoRepositorio) = Modulo
							UltimoRepositorio = UltimoRepositorio + 1
						End If
					On error goto 0
				End If
			End If
		Next
' v1.38n
'		RepositorioSinCargar = False
	End If

End Sub



Public Conchita : Set Conchita = CreateObject ("WScript.Shell")							' Para Tabular() y para quien quiera usarla fuera
' ***
'Private BarraTareas



' Pulsa el tabulador para que un objeto pierda el foco y la página web se recargue o lo que haga falta.
Public Function Tabular ()

'	MsgBox "AAAAACHTUNG !!!", vbOkOnly, "Intervención manual del operador"
	Conchita.SendKeys "{TAB}"
'	If MsgBox ("¿Parar?", vbYesNo, "Intervención manual del operador") = vbYes Then
'		ExitTest
'	End If
'	If Environment ("TestName") <> "[TRANSMISSIONS]-01-03" Then								' Ese parece no necesitarlo
'		MsgBox "Con el foco en el último campo modificado, pulsa el tabulador para que se recargue algún otro campo del formulario",_
'			   vbOkOnly, "Intervención manual del operador"
'	End If
' ***
'	If IsEmpty (BarraTareas) Then
'		Set BarraTareas = Window ("nativeclass:=Shell_TrayWnd")
'	End If
'	BarraTareas.Click

' ### A veces funciona. A veces no.
'	If Environment("Entorn") = "EVO" Then
'		MsgBox "Con el foco en el último campo modificado, pulsa el tabulador para que se recargue algún otro campo del formulario",_
'			   vbOkOnly, "Intervención manual del operador"
'	Else
'		Conchita.SendKeys "{TAB}"
'	End If

' ### No siempre sale la ruleta. Y depende de una función ajena a la librería
'	Tabular = WaitCarregantRepoGenerals

End Function



' Pulsa el tabulador sobre un campo para que, al perder el foco, se recargue la página web.
' La tabulación se hará hasta cumplir una condición determinada o llegar a un límite de intentos.
' Parámetros:
'	- Campo: Objeto de clase WebEdit, WebList, etc sobre el que se tiene que tabular
'	- Objeto: El objeto que se revisa para cumplir la condición que se ha de cumplir para considerar que la recarga se ha producido
'	- Propiedad: La propiedad de ese objeto que se comprobará.
'	- Valor: El valor de la propiedad del objeto que determinará que la condición se cumple.
Sub TabularHastaQue (Campo, Objeto, Propiedad, Valor)

	Dim Intentos : Intentos = 10
	Dim Anterior

	If VarType (Campo) <> vbObject Then
		Reporter.ReportEvent micFail, "Error de programación en TabularHastaQue ()",_
							 "Se intenta usar algo que NO es un objeto. VarType (Campo) = " & VarType (Campo)
		ExitTest
	End If
	Do
		On error resume next															' Al recargar la página, el objeto es otro y da error
			Campo.Object.Focus
		On error goto 0
		Conchita.SendKeys "{TAB}"
'		Wait 2																			' Según qué páginas tardan una eternidad en recargar
		WaitCarregantRepoGenerals
		If Objeto.Exist (0) Then
			If Objeto.GetROProperty (Propiedad) = Valor Then Exit Sub
		End If
		Intentos = Intentos - 1
	Loop until Intentos = 0

End Sub


' Pulsa el tabulador sobre un campo para que, al perder el foco, se recargue la página web.
' La tabulación se hará hasta cumplir una condición determinada o llegar a un límite de intentos.
' Parámetros:
'	- Campo: Objeto de clase WebEdit, WebList, etc sobre el que se tiene que tabular
'	- Objeto: El objeto que se revisa para cumplir la condición que se ha de cumplir para considerar que la recarga se ha producido
Sub TabularHastaQueExista (Campo, Objeto)

	Dim Intentos : Intentos = 10
	
	If VarType (Campo) <> vbObject Then
		Reporter.ReportEvent micFail, "Error de programación en TabularHastaQue ()",_
							 "Se intenta usar algo que NO es un objeto. VarType (Campo) = " & VarType (Campo)
		ExitTest
	End If
	Do
		On error resume next															' Al recargar la página, el objeto es otro y da error
			Campo.Object.Focus
		On error goto 0
		Conchita.SendKeys "{TAB}"
		Wait 2																			' Según qué páginas tardan una eternidad en recargar
		WaitCarregantRepoGenerals
		If Objeto.Exist (1) Then
			Exit Sub
		End If
		Intentos = Intentos - 1
	Loop until Intentos = 0

End Sub



' Devuelve True solo cuando el objeto está "vacío", es nulo o una cadena vacía.
' Resulta útil cuando nuestros subprograma tienen parámetros opcionales
Public Function ParameterIsEmpty (ByRef Objeto)

	Select Case VarType (Objeto)
		Case vbEmpty, vbNull
			ParameterIsEmpty = True
		Case vbString
			ParameterIsEmpty = Objeto = ""
		Case else
			ParameterIsEmpty = False
	End Select

End Function



' Dado un texto en que aparecen cadenas conocidas (etiquetas) y valores desconocidos intercalados, devuelve estos en un vector de cadenas.
' Parámetros:
'	- Texto: Cadena con el texto completo que queremos analizar
'	- Constantes: Un vector de cadenas con los textos conocidos, en el orden exacto en que aparecen en el texto
' ***	Está pensada para textos que contengan n pares (cte, var).
'		¿Hay que ampliarla para los casos de que haya una variable al principio o una constante al final?
Public Function ExtraerVariables (Texto, Constantes)

	Dim Coordenadas ()																	' Cada inicio de constante y el de la variable siguiente
	Dim Variables ()																	' Las variables que vamos encontrando
	Dim NUM_CONSTANTES : NUM_CONSTANTES = UBound (Constantes)							' UBound es muy lenta
	Dim t, c, i																			' Indices varios

	ReDim Coordenadas (NUM_CONSTANTES, 1)
	ReDim Variables (NUM_CONSTANTES)													' *** Todo en Gaudí parece empezar por una constante
	t = 1
	i = 0
	For each c in Constantes
		Coordenadas (i, 0) = InStr (t, Texto, c)										' Buscamos dónde empieza cada constante y
		If Coordenadas (i, 0) = 0 Then
			Reporter.ReportEvent micFail, "ERROR DE PROGRAMACIÓN",_
								 "ExtraerVariables () está buscando una constante que no está en el texto proporcionado:" & vbCR &_
								 "Se busca: """ & c & """ en el texto:" & vbCR & Texto
			ExitTest
		End If
		Coordenadas (i, 1) = Coordenadas (i, 0) + Len (c)								' dónde empieza la variable que le sigue
		t = Coordenadas (i, 1)
		i = i + 1
	Next
' ***	De momento todo en Gaudí parece empezar por una constante
'	If Coordenadas (0, 0) = 1 Then														' Puede empezar con una constante o con una variable
'		ReDim Variables (NUM_CONSTANTES)
'	Else
'		ReDim Variables (NUM_CONSTANTES + 1)
'	End If
	Variables (NUM_CONSTANTES) = Mid (Texto, Coordenadas (NUM_CONSTANTES, 1))			' La última variable, hasta el final
	For i = NUM_CONSTANTES - 1 To 0 Step -1
		Variables (i) = Mid (Texto, Coordenadas (i, 1), Coordenadas (i + 1, 0) - Coordenadas (i, 1))
	Next
	ExtraerVariables = Variables

End Function



' @Description Averigua cuál de los elementos de una lista desplegable contiene una clave concreta en su comienzo.
' Esta función nace para generalizar IndiceEntidad(), de GaudiGenerals.vbs, y emplearla en modelos de impuestos, tarifas, etc
' Parámetros:
'	- Lista:  La lista desplegable, un objeto de clase WebList
'	- Clave: Una cadena que esté al comienzo del elemento de la lista que nos interesa
' Devuelve el índice del elemento dentro de la WebList, un entero mayor o igual que 0.
' En caso de que el elemento no esté en la lista o de error en los parámetros, devuelve -1.
Public Function IndiceWebList (Lista, Clave)

	Dim i, NumElementos																	' Para recorrer la lista
	Dim Longitud																		' *** ¿Todas las claves miden lo mismo?

	If Lista.GetROProperty ("micclass") = "WebList" Then
		Longitud = Len (Clave)
		NumElementos = Lista.GetROProperty ("items count")
		For i = 1 to NumElementos
			If Left (Lista.GetItem (i), Longitud) = Clave Then
				IndiceWebList = i - 1													' GetItem empieza en 1. Select, en 0
				Exit function
			End If
		Next
	End If
	IndiceWebList = -1																	' No lo encuentra, el 0 no es válido y fallará

End Function



' @Description Selecciona un elemento de la lista desplegable que al comienzo tenga una cadena concreta.
' Parámetros:
'	- Lista: La lista desplegable, un objeto de clase WebList
'	- Clave: Una cadena que esté al comienzo del elemento de la lista que nos interesa
' En caso de que no haya en la lista ningún valor que empiece por clave, no hará nada.
Public Sub WebList_Elige (Lista, Clave)

	Dim i

	Traza "Elige " & Clave
	EsperaCarrega Lista
	i = IndiceWebList (Lista, Clave)
	If i > -1 Then
		Lista.Select i
	End If

End Sub

RegisterUserFunc "WebList", "Elige", "WebList_Elige"									' Para incrementar la usabilidad



'@Description Ejecuta SIN PODER DEPURAR los elementos incluidos en la cadena que se pase como parámetro
' No sirve si los subprogramas reciben como parámetros objetos de una clase que hayamos declarado.
' Solo es util para variables normales.
'Sub DeclaraPaquete (s)
'	Execute s
'End Sub

' v1.39
' Esta es la declaración de las cuatro funciones comentadas más abajo.
' Para usarlas es preciso usar DeclaraPaquete o ExecuteGlobal PACK_EVALUACION_PEREZOSA
' según vayamos a usar o no objetos de clases declaradas por nosotros mismos.
'
' Para qué sirve la evaluación perezosa:
'	- Para evaluar el menor número necesario de términos de una expresión lógica compuesta.
'	- A veces solo se puede evaluar un término si se cumplen los precedentes.
' Para qué no sirven estas funciones:
'	- Trabajan en el contexto globlal. Como tal, tienen acceso, por ejemplo, a los objetos del repositorio,
'	  pero NO tienen acceso a atributos privados de una clase
' Se acompaña la función Entrecomillas() para componer las cadenas de los operandos sin evaluarlos antes de tiempo.

'Public PACK_EVALUACION_PEREZOSA : PACK_EVALUACION_PEREZOSA = Join (Array (_
'	"Function LazyAND (x, y): If VarType (x) = vbString Then: x = Eval (x): End If:",_
'	"If x Then: LazyAnd = Eval (y): Else: LazyAnd = False: End If: End Function:",_
'	"Function LazyOR (x, y): If VarType (x) = vbString Then: x = Eval (x): End If:",_
'	"If x Then: LazyOR = True: Else: LazyOR = Eval (y): End If: End Function:",_
'	"Function LazyAnd_M (Condiciones): Dim c: For each c in Condiciones:",_
'	"If not Eval (c) Then: LazyAnd_M = False: Exit Function: End If: Next: LazyAnd_M = True: End Function:",_
'	"Function LazyOr_M (Condiciones): Dim c: For each c in Condiciones:",_
'	"If Eval (c) Then: LazyOr_M = True: Exit Function: End If: Next: LazyOr_M = False: End Function:"))


' ***	v1.37n Equivale a lo declarado justo arriba.
'		Lo conservamos por si fuese preciso modificarlo, tener algo más cómodo.
' Evaluación perezosa de la conjunción
' Parámetros:
'	- x: Una primera expresión lógica de cumplimiento necesario
'	- y: Una cadena conteniendo la segunda expresión lógica de obligado cumplimiento
Function LazyAND (x, y)

	If VarType (x) = vbString Then
		x = Eval (x) 
	End If
	If x Then
		LazyAnd = Eval (y)
	Else
		LazyAnd = False
	End If

End Function



' Evaluación perezosa de la disyunción
' Parámetros:
'	- x: Una primera expresión lógica suficiente
'	- y: Una cadena conteniendo la segunda expresión suficiente
Function LazyOR (x, y)

	If VarType (x) = vbString Then
		x = Eval (x) 
	End If
	If x Then
		LazyOR = True
	Else
		LazyOR = Eval (y)
	End If

End Function



' Evaluación perezosa de la conjunción de un número indeterminado de predicados
' Parámetros:
'	- Condiciones: Un array con las necesarias condiciones a evaluar, como Array ("x = ""a""", "Fichero.EsFinFichero", "1 <> ""1""")
Function LazyAnd_M (Condiciones)

	Dim c

	For each c in Condiciones
		If not Eval (c) Then
			LazyAnd_M = False
			Exit Function
		End If
	Next
	LazyAnd_M = True

End Function



' Evaluación perezosa de la disyunción de un número indeterminado de predicados
' Parámetros:
'	- Condiciones: Un array con las condiciones suficientes a evaluar, como Array ("x = ""a""", "Fichero.EsFinFichero", "1 <> ""1""")
Function LazyOr_M (Condiciones)

	Dim c

	For each c in Condiciones
		If Eval (c) Then
			LazyOr_M = True
			Exit Function
		End If
	Next
	LazyOr_M = False

End Function



Private Const COMILLAS = """"																	' Para meter comillas en las cadenas

'@Description Entrecomilla dos operandos y los concatena con un operador de comparación entre medias
' Esta pensada para su uso en Eval
' Parámetros:
'	- o1, o2: Dos operandos a comparar
'	- op: Cadena con el operador de comparación a aplicar
Function Entrecomillas (o1, op, o2)

'	If VarType (x) = vbString Then
'		' *** Para algo del estilo de Browser ("
'	End If

	Entrecomillas = Join (Array (COMILLAS, o1, COMILLAS, op, COMILLAS, o2, COMILLAS), "")
	
End Function



' Muestra un mensaje en pantalla para indicar al operador que QTP inicia un periodo prolongado de inactividad
' a la espera de que el sistema probado realice los cambios necesarios en la información.
' *** Habría que mostrar el mensaje de una forma más elegante. Pero el tiempo apremia.
Public Sub Traza (Cadena)

	Print Join (Array (Time, Cadena))

End Sub

' El nombre original de Traza era este
Public Sub MostrarMensaje (Cadena)

	Traza Cadena

End Sub



' Devuelve la cadena resultante de escapar en otra cadena los caracteres con semántica al definir expresiones regulares.
' Parámetros:
'	- Cadena: Texto que se quiere procesar.
Function EscaparER (Cadena)

	Dim Buscados, NumBuscados
	Dim Resultado, i, c

	'			123456789.123456
	Buscados = "\^$*+?.()|{},[]-"
	NumBuscados = Len (Buscados)
	Resultado = Cadena
	For i = 1 to NumBuscados
		c = Mid (Buscados, i, 1)
		Resultado = Replace (Resultado, c, "\" & c)
	Next
	EscaparER = Resultado

End Function



'@Description Muestra en el panel Output el contenido de todos los elementos (rellenos) de una tabla web
Sub ContenidoTablaWeb (table)

	Dim filas, cols, i, j, x
	
	Print Left (table.GetROProperty ("innertext"), 30)
	filas = table.RowCount
	For i  = 1 To filas
		cols = table.ColumnCount(i)
		For j  = 1 To cols
			x = table.GetCellData (i, j)
			If x <> "" Then
				Print Join (Array (i, ", ", j, ": ", x), "")
			End If
'		     print "Fila: " & i & " Columna: " & j & " Contenido: " & table.GetCellData(i,j)
		Next
	Next

End sub

'Funcions afegides a partir de juny del 2017

Public Function validacioTextPopup(textPopup)

logger Array("Info","Confirmem que apareix un popup amb el text: '"&textPopup(0)&"'.")

IF Browser("micclass:=Browser").Dialog("micclass:=Dialog").Static("window id:=65535").GetROProperty("regexpwndtitle") <> textPopup(0) Then
	logger Array("KO"," El text del popup no coincideix amb l'esperat. S'espera: '"&textPopup(0)&"'. S'obté: '"&Browser("micclass:=Browser").Dialog("micclass:=Dialog").Static("window id:=65535").GetROProperty("regexpwndtitle")&"'.")
Else
	logger Array("OK","El text del popup es el correcte: '"&textPopup(0)&"'.")
End If
WaitCarregantRepoGenerals
	
End Function



'@Description Pasa l'any, mes, dia, hora minut i segon (10 xifres)
Public Function nombreAleatori
'S'usa per donar un nom únic a les imatges usades al log
dtYear = Right(String(2, "0") & Year(Now), 2)
dtMonth = Right(String(2, "0") & Month(Now), 2)
dtDay = Right(String(2, "0") & Day(Now), 2)
dtHour = Right(String(2, "0") & Hour(Now), 2)
dtMin = Right(String(2, "0") & Minute(Now), 2)
dtSec = Right(String(2, "0") & Second(Now), 2)
nombreAleatori =dtYear&dtMonth&dtDay&dtHour&dtMin&dtSec

End Function


'@Description Revisa el títol de la pàgina.0º Descripció input, 1º Identificador del WebEdit, 2º Valor a comparar ,3º Browser(Opcional)
Public Function validacioContingutInput(dadesContingutInput)
	Dim webElementData
	Dim browserData
	
	WebElementData = dadesContingutInput(1)
	browserData = "micclass:=Browser"
	
	If Ubound(dadesContingutInput)=3 Then
		browserData = dadesContingutInput(3)
	End If
	
	If NOT  Browser(browserData).Page("micclass:=Page").WebEdit(WebElementData).GetROProperty("value") = dadesContingutInput(2) Then
		logger Array("KO",  "El contingut del webEdit '"&dadesContingutInput(0)&"' no es l'esperat. S'espera: "&dadesContingutInput(2)&" i s'obté: "&Browser(browserData).Page("micclass:=Page").WebEdit(WebElementData).GetROProperty("value")&".")
	else
		logger Array("OK",  "El contingut del webEdit '"&dadesContingutInput(0)&"' es el correcte")
	End If
	
End Function

'@Description Valida el contingut d'un WebEdit. 0º Descripcio pel log. 1º Contingut a validar.
Public Function validacioContingutInputDos(ByRef Web_Edit, dadesContingutInputDos)
	
	If NOT Web_Edit.GetROProperty("value") = dadesContingutInputDos(1) Then
		logger Array("KO",  "El contingut del webEdit '"&dadesContingutInputDos(0)&"' no es l'esperat. S'espera: "&dadesContingutInputDos(1)&" i s'obté: "&Web_Edit.GetROProperty("value")&".")
	else
		logger Array("OK",  "El contingut del webEdit '"&dadesContingutInputDos(0)&"' es el correcte")
	End If
	
End Function
RegisterUserFunc "WebEdit","sgtCheckContent","validacioContingutInputDos"

'Funcions d'omplir camps
'@Description Escriu a un WebEdit. 1º Descripcio pel log. 2ºLocalitzador. 3º Valor a escriure. 4º Localitzador Finestra Browser (opcional).
Public Function formulariCompletarWebEdit (dadesWebEdit)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebEdit
	' 2. Valor a introduir al webEdit	
	' 3. Localitzador Finestra Browser
	
	If UBound(dadesWebEdit) = 2 Then
		logger Array("Info", "S'omple l'input '"&dadesWebEdit(0)&"' amb el text: '"&dadesWebEdit(2)&"'.")
		Browser("micclass:=Browser").Page("micclass:=Page").WebEdit(dadesWebEdit(1)).Set dadesWebEdit(2)
	ElseIf UBound(dadesWebEdit) = 3 Then
		logger Array("Info", "S'omple l'input '"&dadesWebEdit(0)&"' amb el text: '"&dadesWebEdit(2)&"'.",dadesWebEdit(3))
		Browser(dadesWebEdit(3)).Page("micclass:=Page").WebEdit(dadesWebEdit(1)).Set dadesWebEdit(2)
	End If

End Function

'@Description Escriu a un WebEdit. 0º Descripcio pel log 1º Valor a escriure.
Public Function formulariCompletarWebEditDos(ByRef Web_Edit, dadesWebEditDos)
	' 0. Nom entendible de l'objecte pel report
	' 1. Valor a introduir al webEdit	
		
	logger Array("Info", "S'omple l'input '"&dadesWebEditDos(0)&"' amb el text: '"&dadesWebEditDos(1)&"'.")
	Web_Edit.Set dadesWebEditDos(1)
	
End Function
RegisterUserFunc "WebEdit","sgtSet","formulariCompletarWebEditDos"



Public Function formulariSeleccionaRadioGroup (dadesRadioGroup)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebEdit
	' 2. Valor a Seleccionar al radioButton
	logger Array("Info", "Es selecciona la opció '"&dadesRadioGroup(2)&"' del radio selector: '"&dadesRadioGroup(0)&"'.")
	Browser("micclass:=Browser").Page("micclass:=Page").WebRadioGroup(dadesRadioGroup(1)).Select dadesRadioGroup(2)
End Function

Public Function formulariClicarWebElement (dadesElement)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebEdit
	' 2. Localitzador Finestra Browser (Opcional)
	
	If UBound(dadesElement) = 1 Then
		logger Array("Info", "Es fa click al  Element '"&dadesElement(0)&"'.")
		Browser("micclass:=Browser").Page("micclass:=Page").WebElement(dadesElement(1)).Click
	ElseIf UBound(dadesElement) = 2 Then 
		logger Array("Info", "Es fa click al  Element '"&dadesElement(0)&"'.",dadesElement(2))
		Browser(dadesElement(2)).Page("micclass:=Page").WebElement(dadesElement(1)).Click
	End If
End Function

Public Function validacioRadioButtonMarcat(dadesValidacioRadioGroup)
	Dim webRadioButtonData
	Dim browserData
	
	browserData = "micclass:=Browser"
	
	If Ubound(dadesValidacioRadioGroup)=3 Then
		browserData = dadesValidacioRadioGroup(3)
	End If
	
	If NOT  Browser(browserData).Page("micclass:=Page").WebRadioGroup(dadesValidacioRadioGroup(1)).GetROProperty("checked") = 1 Then
		logger Array("KO", "S'esperava que el radio button '"&dadesValidacioRadioGroup(0)&"' estigues marcat s'obté: "&Browser(browserData).Page("micclass:=Page").WebRadioGroup(webRadioGroupData).GetROProperty("checked")&".")
	else
		logger Array("OK", "El radio button '"&dadesValidacioRadioGroup(0)&"' està marcat")
	End If
	
End Function


'@Description Revisda si un RadioButton està marcat. 0ºDescripcio pel log.
Public Function validacioRadioButtonMarcatDos(ByRef Web_RadioButton, dadesValidacioRadioGroupDos)
	If NOT  Web_RadioButton.GetROProperty("checked") = 1 Then
		logger Array("KO", "S'esperava que el radio button '"&dadesValidacioRadioGroupDos(0)&"' estigues marcat s'obté: "&Web_RadioButton.GetROProperty("checked")&".")
	else
		logger Array("OK", "El radio button '"&dadesValidacioRadioGroup(0)&"' està marcat")
	End If
	
End Function
RegisterUserFunc "WebRadioButton","sgtCheckStatus","validacioRadioButtonMarcatDos"


Public Function formulariBuidarWebEdit (dadesBuidarWebEdit)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebEdit
	
	logger Array("Info", "Es buida l'input '"&dadesBuidarWebEdit(0)&"'")
	Browser("micclass:=Browser").Page("micclass:=Page").WebEdit(dadesBuidarWebEdit(1)).Set ""

End Function

' @Description Buida un WebElement. 0º Descripció pel log.
Public Function formulariBuidarWebEditDos(ByRef Web_Edit, dadesBuidarWebEditDos)
	' 0. Nom entendible de l'objecte pel report
	
	logger Array("Info", "Es buida l'input '"&dadesBuidarWebEditDos(0)&"'")
	Web_Edit.Set ""

End Function
RegisterUserFunc "WebEdit","sgtClear","formulariBuidarWebEditDos"


Public Function formulariAcceptarWinButton()

	logger Array("Info","Apareix un popup de confirmació. Es prem el botó acceptar")
	Browser("micclass:=Browser").Dialog("text:=Mensaje de página web").WinButton("text:=Aceptar").Click
	
End Function


'@Description Aceptar WinButton d'un popup
Public Function formulariAcceptarWinButtonDos(ByRef Win_Button)

	logger Array("Info","Apareix un popup de confirmació. Es prem el botó acceptar")
	Win_Button.WinButton("text:=Aceptar").Click
	
End Function
RegisterUserFunc "Dialog","sgtAccept","formulariAcceptarWinButtonDos"


Public Function formulariCheckTextIAcceptarWinButton(dadesWinButton)
	logger Array("Info","Apareix un popup. Revisem que el text que hi apareix sigui: '"&dadesWinButton(0)&"'.")
	textPopup = Browser("micclass:=Browser").Dialog("text:=Mensaje de página web").Static("window id:=65535").GetROProperty("text")	
	If textPopup = dadesWinButton(0) Then
		logger Array ("OK","El text del popup coincideix")
	Else
		logger Array ("KO","El text del popup no coincideix")
	End If
	Browser("micclass:=Browser").Dialog("text:=Mensaje de página web").WinButton("text:=Aceptar").Click
End Function


'@Description Revisa el text d'un popup i fa click a Acceptar. 0º Text del popup.
Public Function formulariCheckTextIAcceptarWinButtonDos(ByRef Win_Dialog, dadesWinButtonDos)

	logger Array("Info","Apareix un popup. Revisem que el text que hi apareix sigui: '"&dadesWinButtonDos(0)&"'.")
	textPopup = Win_Dialog.Static("window id:=65535").GetROProperty("text")	
	If textPopup = dadesWinButtonDos(0) Then
		logger Array ("OK","El text del popup coincideix")
	Else
		logger Array ("KO","El text del popup no coincideix")
	End If
	Win_Dialog.WinButton("text:=Aceptar").Click
	
End Function
RegisterUserFunc "Dialog","sgtCheckTextAccept","formulariCheckTextIAcceptarWinButtonDos"


Public Function formulariCompletarWebCheckbox (dadesCheckbox)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebEdit
	' 2. Marcat o desmarcat del checkbox	
	logger Array("Info", "Es marca el checkbox '"&dadesCheckbox(0)&"' com a '"&dadesCheckbox(2)&"'.")
	Browser("micclass:=Browser").Page("micclass:=Page").WebCheckBox(dadesCheckbox(1)).Set dadesCheckbox(2)
	
End Function


'@Description Marca o Desmarca un checkbox. 0º Descripció pel log, 1ºMarcar o desmarcar chebox (ON o OFF)
Public Function formulariCompletarWebCheckboxDos(ByRef Web_CheckBox, dadesCheckboxDos)
	' 0. Nom entendible de l'objecte pel report
	' 1. Marcat o desmarcat del checkbox	
	
	logger Array("Info", "Es marca el checkbox '"&dadesCheckboxDos(0)&"' com a '"&dadesCheckboxDos(1)&"'.")
	Web_CheckBox.Set dadesCheckboxDos(1)
	
End Function
RegisterUserFunc "WebCheckBox","sgtSet","formulariCompletarWebCheckboxDos"


Public Function formulariSeleccionarWebList (dadesWebList)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebList
	' 2. Objecte a escollir
	logger Array("Info", "Es selecciona la opció '"&dadesWebList(2)&"' de la llista '"&dadesWebList(0)&"'.")
	
	If IsArray(dadesWebList(1)) = True Then
			Browser("micclass:=Browser").Page("micclass:=Page").WebList(dadesWebList(1)(0),dadesWebList(1)(1)).Select dadesWebList(2)
	Else
			Browser("micclass:=Browser").Page("micclass:=Page").WebList(dadesWebList(1)).Select dadesWebList(2)
	End If	
	
End Function


'@Description Selecciona un element d'un WebList. 0º Descripció pel log. 1º Element a seleccionar
Public Function formulariSeleccionarWebListDos (ByRef Web_List, dadesWebListDos)
	' 0. Nom entendible de l'objecte pel report
	' 1. Objecte a escollir
	logger Array("Info", "Es selecciona la opció '"&dadesWebListDos(1)&"' de la llista '"&dadesWebListDos(0)&"'.")
	Web_List.Select dadesWebListDos(1)
	
End Function
RegisterUserFunc "WebList","sgtSelect","formulariSeleccionarWebListDos"


'@Description  Selecciona la primera opció d'un weblist SAP
Public Function formulariSAPSeleccionarWebList (dadesSAPWebList)
'Només selecciona la primera opció del llistat
'0 Nombre de vegades que es pulsa fletxa abaix
'1 Localitzador Finestra Browser (Opcional)

	If UBound(dadesSAPWebList) = 1 Then
		logger Array("Info","Seleccionem la '"&dadesSAPWebList(0)&"' opció de la llista",dadesSAPWebList(1))
	End If
	
	Dim WshShell
	Dim comptador
	Set WshShell = CreateObject("WScript.Shell") 
	For comptador = 1 To dadesSAPWebList(0) Step 1
		WshShell.SendKeys "{DOWN}"
		wait(1)
	Next	
	WshShell.SendKeys "{ENTER}"
End Function

Public Function SendKeys(KeysSent)
	Set WshShell = CreateObject("WScript.Shell") 
	WshShell.SendKeys KeysSent
End Function

Public Function formulariClickBoto (dadesClickBoto)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebButton
	If UBound(dadesClickBoto) = 1 Then
		logger Array("Info", "Es fa click al botó '"&dadesClickBoto(0)&"'.")
		If IsArray(dadesClickBoto(1)) = True Then
			Browser("micclass:=Browser").Page("micclass:=Page").WebButton(dadesClickBoto(1)(0),dadesClickBoto(1)(1)).Click
		Else
			Browser("micclass:=Browser").Page("micclass:=Page").WebButton(dadesClickBoto(1)).Click
		End If	
	
	ElseIf UBound(dadesClickBoto) = 2 Then	
		logger Array("Info", "Es fa click al botó '"&dadesClickBoto(0)&"'.")
		If IsArray(dadesClickBoto(1)) = True Then
			Browser(dadesClickBoto(2)).Page("micclass:=Page").WebButton(dadesClickBoto(1)(0),dadesClickBoto(1)(1)).Click
		Else
			Browser(dadesClickBoto(2)).Page("micclass:=Page").WebButton(dadesClickBoto(1)).Click
		End If	
	End If
	
End Function


Public Function formulariClickBotoDos(ByRef Web_Button, dadesClickBotoDos)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebButton

	logger Array("Info", "Es fa click al botó '"&dadesClickBotoDos(0)&"'.")
	Web_Button.Click

End Function
RegisterUserFunc "WebButton","sgtClick","formulariClickBotoDos"


Public Function formulariSAPClickBoto (dadesSAPClickBoto)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del WebButton
	' 2. Localitzador Finestra Browser (Opcional)
	
	If UBound(dadesSAPClickBoto) = 1 Then
		logger Array("Info", "Es fa click al botó '"&dadesSAPClickBoto(0)&"'.")
		If IsArray(dadesSAPClickBoto(1)) = True Then
			Browser("micclass:=Browser").Page("micclass:=Page").SAPButton(dadesSAPClickBoto(1)(0),dadesSAPClickBoto(1)(1)).Click
		Else
			Browser("micclass:=Browser").Page("micclass:=Page").SAPButton(dadesSAPClickBoto(1)).Click
		End If	
	
	ElseIf UBound(dadesSAPClickBoto) = 2 Then	
		logger Array("Info", "Es fa click al botó '"&dadesSAPClickBoto(0)&"'.",dadesSAPClickBoto(2))
		If IsArray(dadesSAPClickBoto(1)) = True Then
			Browser(dadesSAPClickBoto(2)).Page("micclass:=Page").SAPButton(dadesSAPClickBoto(1)(0),dadesSAPClickBoto(1)(1)).Click
		Else
			Browser(dadesSAPClickBoto(2)).Page("micclass:=Page").SAPButton(dadesSAPClickBoto(1)).Click
		End If	
	End If
	
End Function


'@Description Fa click a un botó SAP. 0º Descripció pel log
Public Function formulariSAPClickBotoDos(ByRef WebSap_Button, dadesSAPClickBotoDos)
	' 0. Nom entendible de l'objecte pel report
	logger Array("Info", "Es fa click al botó '"&dadesSAPClickBotoDos(0)&"'.")
	WebSap_Button.Click
	
End Function
RegisterUserFunc "SAPButton","sgtClick","formulariSAPClickBotoDos"


Public Function formulariClickLink (dadesClickLink)
	' 0. Nom entendible de l'objecte pel report
	' 1. Localitzador del Link
	' 2. Localitzador del Frame (Necessari en cas de voler usar Xpath a Gaudí)(Opcional)
		'
	logger Array("Info", "Es fa click al enllaç '"&dadesClickLink(0)&"'.")
	
	If IsArray(dadesClickLink(1))Then
		Browser("micclass:=Browser").Page("micclass:=Page").Link(dadesClickLink(1)(0),dadesClickLink(1)(1)).Click
	ElseIf UBound(dadesClickLink) = 1 Then
		Browser("micclass:=Browser").Page("micclass:=Page").Link(dadesClickLink(1)).Click
	Else
		If UBound(dadesClickLink) = 2 Then
			Browser("micclass:=Browser").Page("micclass:=Page").Frame(dadesClickLink(2)).Link(dadesClickLink(1)).Click
		End If

	End If
	
End Function

'@Description Fa click a un enllaç. 0º Descripció pel log
Public Function formulariClickLinkDos(ByRef Web_Link, dadesClickLinkDos)
	' 0. Nom entendible de l'objecte pel report
	logger Array("Info", "Es fa click al enllaç '"&dadesClickLinkDos(0)&"'.")
	Web_Link.Click
	
End Function
RegisterUserFunc "Link","sgtClick","formulariClickLinkDos"

'Validacions de formularis / cams

'@Description Revisa el títol de la pàgina. 1º Text del títol, 2º Browser(Opcional)
Public Function validacioRevisarTitolPagina(dadesTitolPagina)
	Dim webElementData
	Dim browserData
	
	WebElementData = "innertext:="&dadesTitolPagina(0)
	browserData = "micclass:=Browser"
		
	If Ubound(dadesTitolPagina)=1 Then
		browserData = dadesTitolPagina(1)
	End If
	

	logger Array("Info", "Revisem que el títol de la pàgina actual sigui '"&dadesTitolPagina(0)&"'")
	If Browser(browserData).Page("micclass:=Page").WebElement(WebElementData).Exist Then
		logger Array("OK",  "El títol es el correcte")
	else
		logger Array("KO",  "El títol no es el correcte")
	End If
End Function


'@Description Revisa el títol de la pàgina.0º Descripció element, 1º Identificador del element, 2º Valor a comparar ,3º Frame(Opcional)
Public Function validacioValorElement(dadesValorElement)

	Dim valorWebElement
	Dim browserData
	Dim frameData
	Dim valorElement
	
	WebElementData = dadesValorElement(1)
	browserData = "micclass:=Browser"
	frameData = Null
	
	If Ubound(dadesValorElement)=3 Then
		frameData = dadesValorElement(3)
	End If
	
	If frameData = Null Then
		valorElement = Browser(browserData).Page("micclass:=Page").WebElement(WebElementData).GetROProperty("innertext")
	Else
		valorElement = Browser(browserData).Page("micclass:=Page").Frame(frameData).WebElement(WebElementData).GetROProperty("innertext")
	End If
	
	If NOT  valorElement = dadesValorElement(2) Then
		logger Array("KO",  "El contingut del webEdit '"&dadesValorElement(0)&"' no es l'esperat. S'espera: "&dadesValorElement(2)&" i s'obté: "&valorElement&".")
	else
		logger Array("OK",  "El contingut del webEdit '"&dadesValorElement(0)&"' es el correcte")
	End If
	
End Function


'@Description Valida el valor d'un WebElement. 0º Descripció pel log. 1º Valor de l'element.
Public Function validacioValorElementDos(ByRef Web_Element, dadesValorElementDos)

	Dim valorWebElement
	
	valorElement = Web_Element.GetROProperty("innertext")

	If NOT  valorElement = dadesValorElementDos(1) Then
		logger Array("KO",  "El contingut del webEdit '"&dadesValorElementDos(0)&"' no es l'esperat. S'espera: "&dadesValorElementDos(1)&" i s'obté: "&valorElement&".")
	else
		logger Array("OK",  "El contingut del webEdit '"&dadesValorElementDos(0)&"' es el correcte")
	End If
	
End Function
RegisterUserFunc "WebElement","sgtCheckValue","validacioValorElementDos"

'@Description Fa click a un WebElement. 0º Descripció pel log.
Public Function clickWebElement(ByRef Web_Element, dadesWebElement)
	logger Array("Info", "El Es fa click al element: '"&dadesWebElement(0)&"'.")
	Web_Element.Click
End Function
RegisterUserFunc "WebElement","sgtClick","clickWebElement"

'@Description Revisa el valor d'un WebEdit.0º Descripció element, 1º Identificador del element, 2º Valor a comparar ,3º Frame(Opcional)
Public Function validacioValorWebEdit(dadesValorWebEdit)

	Dim valorWebElement
	Dim browserData
	Dim frameData
	Dim valorElement
	
	WebElementData = dadesValorWebEdit(1)
	browserData = "micclass:=Browser"
	frameData = 0
	
	If Ubound(dadesValorWebEdit)=3 Then
		frameData = dadesValorWebEdit(3)
	End If
	
	If frameData = 0 Then
		valorWebEdit = Browser(browserData).Page("micclass:=Page").WebEdit(dadesValorWebEdit(1)).GetROProperty("value")
	Else
		valorWebEdit = Browser(browserData).Page("micclass:=Page").Frame(frameData).WebEdit(dadesValorWebEdit(1)).GetROProperty("value")
	End If
	
	If NOT  valorWebEdit = dadesValorWebEdit(2) Then
		logger Array("KO",  "El contingut del webEdit '"&dadesValorWebEdit(0)&"' no es l'esperat. S'espera: "&dadesValorWebEdit(2)&" i s'obté: "&valorWebEdit&".")
	else
		logger Array("OK",  "El contingut del webEdit '"&dadesValorWebEdit(0)&"' es el correcte")
	End If
	
End Function


Public Function revisarTitolNavegador(Text)
	Dim valorTitol
	logger Array("Info", "Revisem que el títol de la pestanya del navegador actual sigui '"&Text&"'")
	valorTitol = Browser("micclass:=Browser").Page("micclass:=Page").GetROProperty("title")
	
	If valorTitol = Text Then
		logger Array("OK",  "El títol es el correcte "&Text)
	else
		logger Array("KO",  "El títol no es el correcte "&Text)
	End If
End Function


Public Function validarTaulaResultatsCerca (dadesValidarTaulaResultatsCerca)
Dim estatValidarTaulaResultatsCerca
logger Array ("Info", "Cerquem la fila de dades'"&dadesValidarTaulaResultatsCerca(0)&","&dadesValidarTaulaResultatsCerca(1)&","&dadesValidarTaulaResultatsCerca(2)&","&dadesValidarTaulaResultatsCerca(3)&","&dadesValidarTaulaResultatsCerca(4)&","_
		&dadesValidarTaulaResultatsCerca(5)&","&dadesValidarTaulaResultatsCerca(6)&","&dadesValidarTaulaResultatsCerca(7)&"'.")
		
Browser("micclass:=Browser").Page("micclass:=Page").Frame("name:=workflow").CheckLineaTextExisteix Array(dadesValidarTaulaResultatsCerca(0),dadesValidarTaulaResultatsCerca(1),_
			dadesValidarTaulaResultatsCerca(2),dadesValidarTaulaResultatsCerca(3),dadesValidarTaulaResultatsCerca(4),dadesValidarTaulaResultatsCerca(5),dadesValidarTaulaResultatsCerca(6),dadesValidarTaulaResultatsCerca(7))
End Function


Public Function validarCercaAmbResultats(nif)
	waitCarregantRepoGenerals
	IF (Browser("micclass:=Browser").Page("micclass:=Page").Link("innertext:="&nif(0)).Exist) Then
		validarCercaAmbResultats = true
	Else
		validarCercaAmbResultats = false
	End IF
End Function


'@Description Funció que retorna la data d'avui en format DD/MM/YYYY
Public Function getDataAvui

Dim dtMonth
Dim dtDay


dtMonth = Right(String(2, "0") & Month(date), 2)
dtDay = Right(String(2, "0") & Day(date), 2)
getDataAvui =dtDay&"/"&dtMonth&"/"&YEAR(Now)
End Function

'@Description Funció que et retorna el text d'un webelement. Array 0-Localitzador del frame. 1-Localitzador del element
Public Function DadesObtindreTextDelWebElement (dadesDelWebElement)
' 0 Localitzador del Frame
' 1 Localitzador del Element
	DadesObtindreTextDelWebElement = Browser("micclass:=Browser").Page("micclass:=Page").Frame(dadesDelWebElement(0)).WebElement(dadesDelWebElement(1)).GetROProperty("innertext")
End Function

'@Description Revisem si un element apareix per pantalla. Array 0-Nom pel log, 1-Localitzador del WebElement
Public Function chkExisteix(dadesChkExisteix)
	logger Array("Info","Revisem que l'element '"&dadesChkExisteix(0)&"' existeix.")
	
	If Browser("micclass:=Browser").Page("micclass:=Page").WebElement(dadesChkExisteix(1)).Exist Then
		logger Array("OK","L'element apareix per pantalla")
	Else
		logger Array("KO","NO s'ha localitzat l'element per pantalla")
	
	End If
	
End Function


'@Description Funcio que revisa que un WebElement tingui el valor indicat. Array 0-Nom pel log. 1-Localitzador del WebElement. 2- Valor a comprovar
Public Function chkValorCamp(dadesChkValorCamp)
Dim valorElement
valorElement = Browser("micclass:=Browser").Page("micclass:=Page").WebElement(dadesChkValorCamp(1)).GetROProperty ("Value")

If valorElement = Empty Then
	valorElement = Browser("micclass:=Browser").Page("micclass:=Page").WebElement(dadesChkValorCamp(1)).GetROProperty ("innertext")
End If

If valorElement = CStr(dadesChkValorCamp(2)) Then
	logger Array("OK","El valor per el camp '"&dadesChkValorCamp(0)&"' es el correcte. ("&dadesChkValorCamp(2)&")")
Else
	logger Array("KO","El valor per el camp '"&dadesChkValorCamp(0)&"' es incorrecte. S'esperava '"&dadesChkValorCamp(2)&"' i es troba '"&valorElement&"')")
End If
	
End Function

'@Description Revisa el valor seleccionat d'un weblist. 0º Descipció pel log. 1º Valor a revisar
Public Function chkValorWebList(ByRef Web_List,dadesChkValorCamp)
Dim valorElement
valorElement = Web_List.GetROProperty("Value")

If valorElement = CStr(dadesChkValorCamp(1)) Then
	logger Array("OK","El valor per el camp '"&dadesChkValorCamp(0)&"' es el correcte. ("&dadesChkValorCamp(1)&")")
Else
	logger Array("KO","El valor per el camp '"&dadesChkValorCamp(0)&"' es incorrecte. S'esperava '"&dadesChkValorCamp(1)&"' i es troba '"&valorElement&"')")
End If
	
End Function
RegisterUserFunc "WebList", "sgtValorWebList", "chkValorWebList" 

Public Function LoginSAPWeb ()
	With Browser("B")
		' Si no existe el navegador, abre uno nuevo (Internet Explorer)
		If not .Exist(0) Then
			SystemUtil.Run Environment ("BrowserIE"), Environment ("BDF_" & (Environment ("Entorn")))
		End If
	
		'Selecciona el entorno y accede a la URL correspondiente
		.Navigate Environment ("BDF_" & (Environment ("Entorn")))
		'.Navigate Environment (Environment ("Entorn"))
		'Login
		With .Page ("Entrada al sistema")
			.Sync
			'Solo se pasa por la pantalla de Login si no hay una sesión activa
			If .SAPEdit("Usuario").Exist(1) Then				
				'.SAPEdit("Usuario").Set "52010432A"
				.SAPEdit("Usuario").Set "OQSP"
				'.SAPEdit("Clave de acceso").SetSecure "5dbac3360fb0546db73af8b37aae8169bc9aa224a0c715b750315e67"
				.SAPEdit("Clave de acceso").SetSecure "5ed782c1a6c3eb222870d619783d034e7201ebc9217eb1c7"
'				SelectByIndex .SAPList("Idioma"), 2
				'SelectByValue .SAPList("Idioma"), "CA"				
				'Browser("B").Page("Entrada al sistema").SAPButton("Idioma - Español - Cuadro").Click
				'Browser("B").Page("Entrada al sistema").SAPList("Idioma:").Select "Catalán"
				Browser("B").Page("Entrada al sistema").SAPList("Idioma:").Select "#2"
				'Browser("B").Page("Entrada al sistema").SAPButton("Catalán - 3 de 5 elementos").Click

				.SAPButton("Acceder al sistema").Click
			End If
		End With
		
		.Page ("micclass:=Page").Sync
	End With
End Function



'@Description Funcio de login pel SAP web.
Public Function LoginBDF()	

	revisarTitolNavegador "Entrada al sistema"
	logger Array("info","Procedim a fer el login al sistema")
		

	formulariCompletarWebEdit Array("Usuari","name:=sap-user","52010432A") 'modificat
	formulariCompletarWebEdit Array("Contrasenya","name:=sap-password",InputBox("Paraula de pas de 52010432A", "Paraula de pas"))	'modificat

	logger Array("info","Seleccionem l'idioma Català")
	
	Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=sap-language-dropdown-btn").Click
	
	Wait (1)
	Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=SL1-key-2").Click
	Wait (1)
	
	'Browser("micclass:=Browser").Page("micclass:=Page").Link("name:=Acceder al sistema").sgtClick Array("Acceder al Sistema")
	
	formulariSAPClickBoto Array("Accedir","html id:=LOGON_BUTTON")
	
	If Browser("micclass:=Browser").Page("micclass:=Page").SAPTable("name:=Accessos existents").Exist Then
		formulariSAPClickBoto Array("Continuar",Array("innertext:=Continuar","Location:=0"))
	End If
	
	If Browser("micclass:=Browser").Dialog("regexpwndtitle:=Mensaje de página web").Exist Then
		formulariAcceptarWinButton
	End If
End Function

'@Description Funcio de login pel SAP web.
Public Function LoginBDF2(dni, password)	

	revisarTitolNavegador "Entrada al sistema"
	logger Array("info","Procedim a fer el login al sistema")
	formulariCompletarWebEdit Array("Usuari","name:=sap-user",dni) 'modificat
	formulariCompletarWebEdit Array("Contrasenya","name:=sap-password",password) 'modificat
	'formulariCompletarWebEdit Array("Contrasenya","name:=sap-password",InputBox("Paraula de pas de 46653452E", "Paraula de pas"))
	logger Array("info","Seleccionem l'idioma Català")
	Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=sap-language-dropdown-btn").Click
	Wait (1)
	Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=SL1-key-2").Click
	Wait (1)
	
	'Browser("micclass:=Browser").Page("micclass:=Page").Link("name:=Acceder al sistema").sgtClick Array("Acceder al Sistema")
	
	'formulariSAPClickBoto Array("Acceder al sistema","html id:=LOGON_BUTTON")
	Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=LOGON_BUTTON").Click
	wait (3)
	If Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=SESSION_QUERY_CONTINUE_BUTTON").Exist Then
		Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=SESSION_QUERY_CONTINUE_BUTTON").Click
		'formulariSAPClickBoto Array("Continuar",Array("innertext:=Continuar","Location:=0"))
	End If
	
'	If Browser("micclass:=Browser").Dialog("regexpwndtitle:=Mensaje de página web").Exist Then
'		formulariAcceptarWinButton
'	End If
End Function

' @Description Revisa si hi ha un procés en execució al sistema i el mata. 0º Descipció pel log. 1º Nom del procés (PE: chrome.exe)
Public Function checkMatarProces(proces)

	Dim wShell, exec
	Dim str
	
	Set wShell = CreateObject( "WScript.Shell" )
	Set exec = wShell.Exec("TASKLIST /FI ""IMAGENAME eq "&proces(1)&"""")
	str = exec.StdOut.ReadAll
	
	If InStr(str,"no hay tareas") Then 
		logger Array("Info","No hi ha cap procés "&proces(0))
	Else 
		Set exec = wShell.Exec("TASKKILL /F /im "&proces(1))
		logger Array("Info","Matem els processos "&proces(0)&" existents")
	End If
	
End Function

Public Function waitCarregantBaseDeDadesFiscal(byRef Browser)
	Do While Browser.Page("micclass:=Page").Frame("name:=content_frame").WebElement("xpath:=//div[@class='lsLoadImg']").Exist(3)
			contador = contador+1
			wait 5
			If contador >= 12 Then
				Exit Do
				logger Array("Info","S'ha esperat dos minuts pero encara segueix carregant")
			End If
			
			logger Array("Info","Apareix el gif de càrrega. Esperem")
	loop	
End Function
RegisterUserFunc "Browser", "waitCarregantBaseDeDadesFiscal", "waitCarregantBaseDeDadesFiscal"

'@Description	
'@Documentation 
Public Function GenerarNIF ()
	Dim t, nif_num, letra
	Dim letras_nif, ceros
	
	letras_nif = Array("T", "R", "W", "A", "G", "M", "Y", "F", "P", "D", "X", "B", "N", "J", "Z", "S", "Q", "V", "H", "L", "C", "K", "E")
	ceros = "00000000"
	
'	s = Right(timeStamp(now, 2), 8)
	
'	nif_num = CLng(s)
	
	t = now
	nif_num = Month(t) * 31 + Day(t)
	nif_num = nif_num * 24 + Hour(t)
	nif_num = nif_num * 60 * 60 + Minute(t) * 60 + Second(t)
	
	letra = letras_nif(nif_num Mod 23)
	nif_num =  CStr(nif_num)
	
	GenerarNIF = Right(ceros, 8 - Len(nif_num)) & nif_num & letra
	
	Erase letras_nif
End Function

''@Description	
''@Documentation 
'Public Sub SetWithKeys (ByRef web_edit, ByVal str)
''<web_edit>	
''<str>		
'
'	Dim last_char
'	Dim WshShell
'	
'	Set WshShell = CreateObject("WScript.Shell")
'	
'	last_char = Right(str, 1)
'	str = Left(str, Len(str) - 1)
'	
'	web_edit.Set str
'	web_edit.Click
'	WshShell.SendKeys "(" & last_char & ")"
'	
'	Set last_char = Nothing
'	Set WshShell = Nothing
'End Sub
'RegisterUserFunc "WebEdit", "SetWithKeys", "SetWithKeys"

'@Description	
'@Documentation 
Public Sub SetWithKeys (ByRef web_edit, ByVal str)
'<web_edit>	
'<str>		

	Dim i
	Dim WshShell
	
	Set WshShell = CreateObject("WScript.Shell")
	
	web_edit.Click
	
	For i = 1 To Len(str)
		WshShell.SendKeys "(" & Mid(str, i, 1) & ")"
	Next
	
	Set WshShell = Nothing
End Sub
RegisterUserFunc "WebEdit", "SetWithKeys", "SetWithKeys"
RegisterUserFunc "WinEdit", "SetWithKeys", "SetWithKeys"

'@Description	
'@Documentation 
Public Sub SetAndEnter (ByRef web_edit, ByVal str)
'<web_edit>	
'<str>		

	Dim WshShell
	
	Set WshShell = CreateObject("WScript.Shell")
	
	web_edit.Set str
	web_edit.Click
	WshShell.SendKeys "{ENTER}"
	
	Set WshShell = Nothing
End Sub
RegisterUserFunc "WebEdit", "SetAndEnter", "SetAndEnter"
RegisterUserFunc "WinEdit", "SetAndEnter", "SetAndEnter"
RegisterUserFunc "SAPEdit", "SetAndEnter", "SetAndEnter"

'@Description	Realiza un click físico sobre el objeto pasado como parámetro. ¡¡Atención!! Cuando se habilita "ReplayType = 2" se hace un click físico sobre la aplicación,
'				en la ubicación señalada. Si hay algo tapando el objetivo puede fallar o tener un comportamiento impredecible.
'@Documentation Realzia un click físico sobre <obj>.  ¡¡Atención!! Cuando se habilita "ReplayType = 2" se hace un click físico sobre la aplicación,	en la ubicación señalada.
'				Si hay algo tapando el objetivo puede fallar o tener un comportamiento impredecible.
Public Sub ClickFisico (ByRef obj)
'<obj>	Objeto sobre el que se realiza el click físico.
	Dim old
	old = Setting.WebPackage("ReplayType")
	Setting.WebPackage("ReplayType") = 2	' Enable Replay Type
	obj.Click
	Setting.WebPackage("ReplayType") = old   ' Disable Replay Type
	Set old = Nothing
End Sub
RegisterUserFunc "WebElement", "ClickFisico", "ClickFisico"
RegisterUserFunc "WebEdit", "ClickFisico", "ClickFisico"

Public Function CheckObjExist(ByRef obj, ByRef obj_name, ByVal timeout)
	Dim res
	If timeout = Nothing Then
		timeout = 0
	End If
	
	res = .WebElement("").Exist(timeout)
	If res Then
		Reporter.ReportEvent micPass, obj_name, "L'objecte '" + obj_name + "' existeix."
	Else
		Reporter.ReportEvent micFail, obj_name, "L'objecte '" + obj_name + "' NO existeix."
	End If
	CheckObjExist = res
End Function

'@Description	
'@Documentation	
Public Function ArrayFind (ByRef arr, ByVal elem)
	For ArrayFind = 0 To UBound(arr)
		If arr(ArrayFind) = elem Then
			Exit Function
		End If		
	Next
	ArrayFind = -1
	
End Function

'@Description	Evalúa el primer parámetro y devuelve el segundo en caso de validarse (True) o el tercero en caso de que no (False).
'				Permite realizar evaluaciones en línea.
'@Documentation	Evalúa <bClause> y devuelve <sTrue> en caso de validarse (True) o <sFalse> en caso de que no (False).
'				Permite realizar evaluaciones en línea.
Public Function IIf(bClause, sTrue, sFalse)
'<bClause>:	Cláusula a evaluar.
'<sTrue>:	Respuesta en caso de <bClause> = True
'<sFalse>:	Respeusta en caso de <bClasue> = False
'return:	Devuelve <sTrue> si bClause es True, o <sFalse> si bClause es False

    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

'@Description	Valida si el NIF pasado es de una persona física o de una entidad/organización.
'@Documentation	Valida si el NIF <nif> es de una persona física o de una entidad/organización.
Public Function EsPersona(ByVal nif)
'<nif>:		Número de identificación. A de ser un NIF, NIE o CIF.
'return:	Devuelve True si <nif> es un NIF o NIE o False si es un CIF.

	Dim oRE
	Set oRE = new RegExp
	
	oRE.Pattern = "^[xyzXYZ\d]\d{7}\w"
	
	EsPersona = oRE.Test(nif)
	
	Set oRE = Nothing
End Function

'@Description	Importa una Hoja de un fichero .xlsx al DataTable.
'@Documentation	Importal la Hoja <dSourceSheetName> del fichero <dFileName> a la Hoka <dDestinationSheetName> del DataTable.
Function ImportSheetFromXLSX(dFileName, dSourceSheetName, dDestinationSheetName)
'<dFileName>:				Ruta del fichero XLSX a importar.
'<dSourceSheetName>:		Nombre de la Hoja de <dFileName> a importar.
'<dDestinationSheetName>:	Nombre de la Hoja del DataTable en la que cargar la Hoja importada.
'return:					Referencia a la Hoja del DataTable <dDestinationSheetName>.

	Dim ExcelApp
	Dim ExcelFile
	Dim ExcelSheet
	Dim sRowCount
	Dim sColumnCount
	Dim sRowIndex
	Dim sColumnIndex
	Dim sColumnValue
	
	Set ExcelApp = CreateObject("Excel.Application")
	Set ExcelFile = ExcelApp.WorkBooks.Open (dFileName)
	Set ExcelSheet = ExcelApp.WorkSheets(dSourceSheetName)
	
	Set qSheet = DataTable.GetSheet(dDestinationSheetName)
	
	sColumnCount = ExcelSheet.UsedRange.Columns.Count
	sRowCount = ExcelSheet.UsedRange.rows.count
	
	For sColumnIndex = 1 to sColumnCount
		
		sColumnValue = ExcelSheet.Cells(1, sColumnIndex)
		sColumnValue = Replace(sColumnValue, " ", "_")
		
		If sColumnValue = "" Then
			sColumnValue = "NoColumn" & sColumnIndex
		End If
		
		Set qColumn = qSheet.AddParameter (sColumnValue,"")
		
		For sRowIndex = 2 to sRowCount
			sRowValue = ExcelSheet.Cells(sRowIndex, sColumnIndex)
			qColumn.ValueByRow(sRowIndex - 1) = sRowValue
		Next
		
	Next
	Set ImportSheetFromXLSX = qSheet
	ExcelFile.Close
	ExcelApp.Quit
	
	Set ExcelApp = Nothing
	Set ExcelFile = Nothing
	Set ExcelSheet = Nothing
	Set qSheet = Nothing
End Function

'@Description	Devuelve la representación en formato de moneda (X.XXX,YY) del valor indicado.
'@Documentation	Devuelve la representación en formato de moneda (X.XXX,YY) de <valor>.
Function FormatoMoneda(ByVal valor)
'<valor>:	Valor numérico que se quiere obtener en formato de moneda.
'return:	String con la representaión del número <valor> en formato de moneda (X.XXX,YY)

	Dim temp, valor_str, index
	
	valor_str = CStr(CLng(valor * 100))
	temp = "," & Right(valor_str, 2)
	index = Len(valor_str) - 1
	
	Do While index > 4
		index = index - 3
		temp = "." & Mid(valor_str, index, 3) & temp
	Loop
	
	FormatoMoneda = Left(valor_str, index - 1) & temp
End Function

'@Description	Devuelve el String repetido el número de veces indicado.
'@Documentation	Devuelve el String <str> repetido <rep> veces.
Function RepeatString(ByVal str, ByVal rep)
'<str>:		Cadena que a repetir.
'<rep>:		Número de repeticiones del String <str>.
'return:	String <str> repetido <rep> veces.
	Dim i
	
	RepeatString = ""
	
	For i = 1 To rep
		RepeatString = RepeatString & str
	Next
End Function

''' *** No hay tiempo
'''Private Const LOCALIZACION = "Localización"												' Menos cadenas en memoria
'''
'''
'''
'''' Carga el data table con los datos del test en cuestión
'''Sub PrepararLocalizacion ()
'''
'''	DataTable.AddSheet LOCALIZACION
'''	DataTable.ImportSheet QTP_FILES & LOCALIZACION & ".xls", LOCALIZACION, LOCALIZACION
'''	Do
'''		If DataTable ("Test", LOCALIZACION) = Environment ("TestName") Then
'''			
'''		End If
'''	Loop until DataTable.GetCurrentRow = 1
'''	Data
'''
'''End Sub
'''
'''
'''
'''Class cDireccion
'''
'''	Private Provincia, CP, Municipio, Via, TipoVia										' Los componentes básicos de cualquier dirección
'''
'''End Class
