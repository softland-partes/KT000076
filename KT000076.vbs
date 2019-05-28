'--------------------- Variables Globales -------------
Public oConn 'Conexión a la base
Dim oFSO, oReadFolder, oReadFile, oFilestream 'Lectura de Archivos
Dim cNombre_Empresa , sEsMulti, sEsMain, sCodemp'Empresa destino
Dim sQItem, sQValues  'Query a la tabla conversión de Archivos
Dim resultSetItem ,resultSetValue 'Resulset de La Query a la Tabla de conversión
Dim IContador, mensajeError, FechaFTP
Dim aListValue , aItemValue
Dim aListKey , aItemKey, identi
Dim fOrigen , fCodigo 'Ruta de archivos y codigo de la interfaz
Dim aParents, ValorRecup, noExisteArtCod, m_sCodProvincia
Dim dKeyValue, dFieldKey, dFieldTable, dFieldOrden, dValues, dValidaciones 'Diccionarios creados con el JSON
'------------------------------------------------------
m_sCodProvincia = "JURISD"
m_sListPrecio = "LISPRE"
cNombre_Empresa = "HUINCA"
sPathEjecutable = "C:\WinSCP\Winscp.com"
mensajeError = ""

FechaFTP ="2019-04-16" ' TIMESTAMP  : hoy
Const cNombre_Wizard_IN = "INIMTREXWIZ"    'Wizard que se intenta ejecutar
Const cTipoObjeto_IN = 6                   'Tipo del objeto que se intenta crear. A un Wizard le corresponde un 6


OpenConnection
ProcesoEmpresas
FTPDownload

Do
	 bFlag = CheckProcess(sPathEjecutable)
Loop Until bFlag = False

IContador = 0
'Array para array multidimesional
'Valores
aListValue = Array()
aItemValue = Array()
'Tags(Nodos)
aListKey = Array()
aItemKey = Array()

noExisteArtCod = False
'Array creado para guardar los predecesores de las claves hojas
aParents = Array()

Set dTablasMadres = CreateObject("Scripting.Dictionary")

'USR_INMTAI_CLAARC = Campo en la tabla SAR
'USR_INMTAE_ORIGEN = Carpeta que contiene los JSON.
sQItem = "Select USR_INMTAE_CODIGO codigo, USR_INMTAE_ORIGEN origen, USR_INMTAI_CLAARC columna from USR_INMTAI " & _
	 	"INNER JOIN USR_INMTAE ON USR_INMTAI_CODIGO = USR_INMTAE_CODIGO order by USR_INMTAI_AORDEN asc "

Set resultSetItem = oConn.Execute(CStr(sQItem))

fOrigen = resultSetItem("origen").Value
fCodigo= resultSetItem("codigo").Value

CreardValues(fCodigo)  'Diccionario de claves.


Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oReadFolder = oFSO.GetFolder(fOrigen)

'Crea diccionarios con los datos de la tabal conversion de archivos.
dTableJsonCompose

For Each oReadFile In oReadFolder.Files
	Set oFilestream = oReadFile.OpenAsTextStream(1,-2)
	Set dKeyValue = CreateObject("Scripting.Dictionary")
	identi = ""
	sTabla ="XXX"
	'Crea una estructura tipo dicionario multiple con los VALORES del JSON
	creaStructJsonData oFilestream, dKeyValue, aParents

  IcontadorFielTable = 0
  IcontadorFieldKey = -1

  aKeyMaximo = 0
  'Array para campos de tabla
   aTabla = Array()

   For Each key in dFieldKey.Keys
    'aTabla es un array que me va a indicar que campos de mi tabla voy a insertar
    'Hay que tener en cuenta que cada vez que cambie de tabla el array se va a Poner en 0
    if stabla = "XXX" then
      stabla = dFieldTable.Item(key)
    End If
    IcontadorFieldKey = IcontadorFieldKey + 1
    if stabla = dFieldTable.Item(key) then
      MaxValue(key)
      Redim Preserve aTabla(IcontadorFielTable)
      aTabla(IcontadorFielTable) = key
      IcontadorFielTable  = IcontadorFielTable  + 1
    End If

    If stabla <> dFieldTable.Item(key) or UBound(dFieldKey.Keys) = IcontadorFieldKey then
		keyBuffer = key
		InsertTable(aKeyMaximo)
		stabla = dFieldTable.Item(keyBuffer)
		aTabla = Array()
		IcontadorFielTable  = 0
		Redim Preserve aTabla(IcontadorFielTable)
		aTabla(IcontadorFielTable) = keyBuffer
		IcontadorFielTable  = IcontadorFielTable  +  1
    End If
   Next

	Init
	oFilestream.close
	procesarOrdenes(oReadFile.name)
Next
ProcesoInterfaz("USR_CL")
ProcesoInterfaz("USR_FC")
CloseConnection

Sub CreardValues(codigo)
	Dim Icon
	Icont = 0
	'Busco todas los valores de mi tablas que son claves en el archivo
	sQValues = "Select * FROM USR_INMTAV INNER JOIN USR_INMTAE ON USR_INMTAV_CODIGO = '"&codigo&"'"
	Set resultSetValue = oConn.Execute(CStr(sQValues))

  Set dValues = CreateObject("Scripting.Dictionary")
	Do While Not resultSetValue.EOF
		valor = resultSetValue("USR_INMTAV_AVALOR").value
    valid = resultSetValue("USR_INMTAV_AVALID").value
    if dValues.Exists(UCase(valor)) = false  then
      dValues.Add UCase(valor) , valid
  		Icont = Icont + 1
    End if
		resultSetValue.MoveNext
	Loop
End Sub
Sub dTableJsonCompose()
	Dim clave, aTabla , valor ,Icontador, rd2, linea, parraf
	Icontador = 0
	valor = ""
  parraf = ""
	cont = 0
   Set dFieldKey = CreateObject("Scripting.Dictionary")
   Set dFieldTable = CreateObject("Scripting.Dictionary")
   Set dFieldOrden= CreateObject("Scripting.Dictionary")
	 parraf = "Diccionario estructura" & vbCRLF

   sQValue = "Select USR_INMTAV_AVALOR val, USR_INMTAV_TIPSEP sep  FROM USR_INMTAV WHERE USR_INMTAV_CLAARC = "
   Do While Not resultSetItem.EOF
    clave = resultSetItem ("columna").Value
    'Rearmo la query segùn la clave
    Set rd2 = oConn.Execute(CStr(sQValue&Chr(39)&CStr(clave)&Chr(39)))

    aTabla = SPLIT(clave,"_")
    sNombreTabla = "SAR_"&aTabla(1)

		Do While Not rd2.EOF

			val = rd2("val").Value
			valor = valor & val

			ArrayMultidimensionalValores val, aListKey, aItemKey,IContador
			cont = cont + 1
			rd2.MoveNext
		Loop

    dFieldKey.Add clave , valor
    dFieldTable.Add clave , sNombreTabla
    dFieldOrden.Add clave , IContador
		'-------------ARMANDO LOG-------------'
    linea = "Orden: " & IContador &_
            " Tabla: "& sNombreTabla &_
            " Clave: " & clave &_
            " Valor: " & valor  & vbCRLF
    parraf = parraf & linea
		'----------------LOG-------------'
    IContador = IContador + 1
    valor = ""
    resultSetItem.MoveNext
   Loop
   grabarLog_Archivo(parraf)

End Sub
Sub creaStructJsonData(oFilestream, dKeyValue, aParents)

	Do While Not oFilestream.AtEndOfStream
		sLineaActual = oFilestream.Readline
			'Guardo la clave y el valor en sClave y sValor
			arrLinea = Split(sLineaActual,":")

			val = Trim(Replace(Replace(CStr(arrLinea(0)),vbTab,""),Chr(34),""))
			isParent  = searchParents(arrLinea, sLineaActual, aParents, val)
			if isParent = 0 then
			If UBound(aParents)<0 then
					 Predecesor =  ""
				Else
						Predecesor = JOIN(aParents,".")&"."
			End If
			crearIdenti val, Predecesor, sLineaActual
			sClave = UCase(ClaveOK(val, Predecesor))
			if sClave <> "F" then
				sValor = ObtenerValor(sClave, sLineaActual, Predecesor & val )

				 sClave = UCase(Predecesor&sClave)
				 if dKeyValue.Exists(sClave) <> true then
					 dKeyValue.add sClave, IContador
					 IContador = IContador + 1
				 End if
				 ArrayMultidimensionalValores sValor, aListValue, aItemValue, dKeyValue.Item(sClave)
			End If
		End if
	Loop
End Sub
Sub InsertTable(max)

	'aValueInsert: Array de valores para insertar
	'aKeyInsert: Array de Campo de tabla para el Insert
	'aKeyChange: Traigo de una lista que tiene array de valores para cada campo el que necesito
	'max: es un numero que me indica cuantas inserciones tengo que hacer.
	'Declaro los array sin dimension para que sean dinamicos y poder Redimensionarlo
	'dependiendo la cantidad de Datos que venga de mi tabla parametro

	Dim aValueInsert,sKeyChange,IContador, aKeyInsert ,aux, mensajeError
	aValueInsert = Array()
	aKeyInsert = Array()
	Dim resultFordate

	For i = 0 to (max-1)

		Redim Preserve aValueInsert(i)
		IContador = 0

		For each Field in aTabla  'Recorro los campos de la tabla SAR
		 	Redim Preserve aKeyInsert(IContador)   'Voy creando un array con los datos que quiero insertar
			aKeyChange = aListKey(dFieldOrden.Item(Field))
			NewKey = dFieldKey.Item(Field)

			contValueMax = 0
			for Each key in aKeyChange
				if dKeyValue.Exists(Key) = true then
						arrayItemValues = aListValue(dKeyValue.Item(Key))
						itemValue = arrayItemValues(i)
						if contValueMax > 0 then
							itemValue = " " + itemValue
						end if
				elseIf Key = "SAR_COUNT" then
					itemValue = i + 1
				elseIf Key = "ERRMSG" then
					itemValue = mensajeError
				elseIf Key = "STATUS" AND mensajeError <> "" then
					itemValue = "E"
				elseIf Key = "STATUS" AND mensajeError = "" then
						itemValue = "N"
				elseIf Key = "CODEMP" then
					itemValue = sCodemp
				elseIf Key = "IDENTI" then
					itemValue = identi
				else
					itemValue = Key
				End If

				NewKey = Replace(NewKey,key,itemValue)
				contValueMax = contValueMax + 1
			Next
			errmsg = catchError(NewKey)
			if errmsg <>"" Then
				mensajeError = mensajeError & errmsg &" | "
			End if

      stabla = dFieldTable.Item(Field)
      aKeyInsert(IContador) = convert(stabla, Field, NewKey,mensajeError)
      IContador = IContador + 1
		Next

		aValueInsert(i) = "("&JOIN(aKeyInsert,",")&")"
	Next
	if Instr(stabla,"SAR")<=0 Then
		stabla = "SAR_"&stabla
	End if
	sInsert="Insert into "&CStr(stabla)&"("& JOIN(aTabla,",") &")"&" Values "& JOIN(aValueInsert,",")
	oConn.Execute(CStr(sInsert))

End Sub
Function FTPDownload()
  Set oFTPScriptShell = CreateObject("WScript.Shell")
  oFTPScriptShell.Run "c://Winscp/Winscp.com /command "& chr(34) &"open ftp://softland:Ttd7r?30@cocinaconvalentino.com.ar/"&chr(34)&_
      			    " "&chr(34)&"get output/orders/order_1014.json C:\temp\Ordenes\" & chr(34) &_
	    	  	 	    " "&chr(34)& "mv output/orders/order_1014.json  /input/*.json "&chr(34)&"  exit"
	    	  	 	    '" "&chr(34)&"get output/orders/*>=%%"&FechaFTP&"#yyyy-mm-dd%% C:\temp\Ordenes\" &chr(34)&_
	    	  	 	    '" "&chr(34)&"mv output/orders/*<=%%"&FechaFTP&"#yyyy-mm-dd%%  /input/*.json" "&chr(34)&"  exit"




  Set oFTPScriptShell = Nothing
End function
Function CheckProcess(ProcessPath)
	Dim strComputer, objWMIService, colProcesses, WshShell, Tab, ProcessName
	 strComputer = "."
	 Tab = Split(ProcessPath,"\")
	 ProcessName = Tab(UBound(Tab))
	 ProcessName = Replace(ProcessName, Chr(34), "")
	 Set objWMIService = GetObject("winmgmts:" _
	 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	 Set colProcesses = objWMIService.ExecQuery _
	 ("Select * from Win32_Process Where Name = '" & ProcessName & "'")

	 If colProcesses.Count = 0 Then
		 CheckProcess = False
	 Else
	 	 CheckProcess = True
	 End If
End Function
Sub ProcesoInterfaz(CodTra)

    Set oWizard = oApplication.Companies(sCodemp).GetObject(cNombre_Wizard_IN, cTipoObjeto_IN, "RUN_FOR_SCRIPT")

    With oWizard.Steps(1).Table
        .Fields("VIRT_CODTRA").Value = CodTra
        .Fields("VIRT_TIPIMP").Value = "T"
        oWizard.Finish
    End With

    Set oWizard = Nothing
End Sub
Sub ProcesoEmpresas()
	Dim RdCodEmp,sSqlCodemp
	sSqlCodemp = "Select ISNULL(ISMAIN,'N') MAIN,	ISNULL(ISMULTI,'N') MULTI, GRTEMP_CODEMP CODEMP "&_
			   "From CWSGCORE.DBO.CWOMCOMPANIES "&_
	  		   "inner join GRTEMP on GRTEMP_ISMAIN = ISMAIN "&_
			   "Where NAME='HUINCA'"
				 '"case When ISNULL(ISMULTI,'N')  = 'S' "&_
				' "then (SELECT GRTEMP_CODEMP FROM GRTEMP WHERE GRTEMP_ISMAIN = 'N') "&_
				 '"else '"&cNombre_Empresa&"' END CODEMP "&_
				' "From CWSGCORE.DBO.CWOMCOMPANIES Where NAME='"&cNombre_Empresa &"' "
	Set RdCodEmp = oConn.Execute(CStr(sSqlCodemp))
  If Not RdCodEmp.EOF Then
    sEsMulti = RdCodEmp("MULTI").Value
    sEsMain = RdCodEmp("MAIN").Value
	sCodemp = RdCodEmp("CODEMP").Value
  End If
  RdCodEmp.Close
  Set RdCodEmp = Nothing
End Sub

'----------------------------------------------------------------------
'-------------------- FUNCIONES AUXILIARES ----------------------------
'----------------------------------------------------------------------

Sub ArrayMultidimensionalValores(val,lista,item,max)
	Dim ultimo, ultimoLista
	ultimoLista = Ubound(lista)

	if ultimoLista  <> -1 and ultimoLista >= max then
		valor =max
		arrItem = lista(valor)
		ultimoItem = Ubound(arrItem)+1
		Redim preserve arrItem(ultimoItem)
		arrItem(ultimoItem) = val
		lista(valor) = arrItem
	else
		if ultimoLista = -1 then
			ultimo = 0
		else
			ultimo  = ultimoLista + 1
		End if

		Redim preserve lista(ultimo)
		Redim preserve item(0)
		item(0)= val
		lista(ultimo) = item
	End If
End Sub
Function searchParents(arrLinea, sLineaActual, aParents, val)
	Dim cantidadAux, flag
	Dim ObjRegEx, ColMatches
	Set ObjRegEx = CreateObject("VBScript.RegExp")
	ObjRegEx.Global = True
	cantidadAux = 0
	flag = 0
	if UBound(arrLinea)>0 then
		if Trim(arrLinea(1)) = "{" or Trim(arrLinea(1)) = "[" then
			LastIndex = UBound(aParents)
			if LastIndex <0 then
				Redim Preserve aParents(0)
				aParents(0) = val
			else
				Redim Preserve aParents(LastIndex + 1)
	  			aParents(LastIndex + 1) = val
	  		End If
	  		flag = 1

		End If
	Else
		slinea = EliminarTabCadena(sLineaActual)
		ObjRegEx.Pattern = "},{"
	 	Set ColMatches =  ObjRegEx.Execute(slinea)
	 	if colMatches.Count = 0 and UBound(aParents)>-1 then

			ObjRegEx.Pattern = "}"
			Set ColMatches =  ObjRegEx.Execute(slinea)
			cantidadAux = colMatches.Count
			ObjRegEx.Pattern = "]"
			Set ColMatches =  ObjRegEx.Execute(slinea)
			if colMatches.Count > 0 then
				cantidadAux = colMatches.Count
			End If
			if cantidadAux>0 then
				flag = 1
				for i = 1 to cantidadAux

				LastIndex = UBound(aParents)

			 		if LastIndex = 0 then
				 	 	 aParents = Array()
				 	 else
						Redim Preserve aParents(LastIndex-1)
			  		End If
		  		Next
			End If
		End If
	End If
	searchParents = flag
End Function
Function EliminarTabCadena(linea)
	 Do While InStr(1,linea,vbTab)>0
	 	 linea = Replace(linea,vbTab,"")
	 Loop
	 EliminarTabCadena = Trim(linea)
End Function
Function ObtenerValor(clave, sLineaActual, dValor)
	Dim ObjRegEx ,valor , ColMatches ,sLinea
	Set ObjRegEx = CreateObject("VBScript.RegExp")
	ObjRegEx.Global = True 'Esto hace que no solo busque en la primera que encuentre
	ObjRegEx.IgnoreCase = true  ' no sensible a las mayúsculas
	ObjRegEx.Pattern = CStr(Chr(34)&clave&Chr(34) &":")
	sLinea = EliminarTabCadena(sLineaActual)

	Set ColMatches =  ObjRegEx.Execute(sLineaActual)
	if colMatches.Count>0 then
		valor = Right(sLinea,len(sLinea) - ColMatches(0).length)
    valor = Trim(Replace(Replace(valor,",",""),Chr(34),""))
    valor = corregirCaracteresSpeciales(valor)
    ValorRecup = valor

    valid = dValues.Item(UCase(dValor))
    if valid <> "" Then
      creardValidaciones
      valor = dValidaciones.item(LCase(valid))
    End if

		ObtenerValor = valor
	else
		ObtenerValor =""
	End if
End Function
Function ClaveOK(val, Predecesor)
	if dValues.Exists(UCase(CStr(Predecesor&val))) then
		ClaveOK = val
	else
		ClaveOK = "f"
	End if
End Function
Function convert(Table,Column,newKey,mensajeError)
	 Dim ValorCastear
		Sql = "Select systypes.name as Tipo from syscolumns "& _
					"inner join systypes ON systypes.xtype = syscolumns.xtype "& _
					"and  syscolumns.name = '"&Column&"' inner join sysobjects " & _
					"on syscolumns.id = sysobjects.id and sysobjects.name = '"& Table &"'"

		Set rd3 = oConn.Execute(CStr(Sql))

		TipoDeDato = rd3("Tipo").Value
		if newKey ="" and (TipoDeDato = "numeric" or TipoDeDato = "int") Then
			newKey = 0
		End if
		Select case TipoDeDato
			case "varchar"
				ValorCasteado ="'"&CStr(newKey)&"'"
			Case "numeric"
				ValorCasteado = CDbl(newKey)
			Case "int"
				ValorCasteado = CInt(newKey)
			Case Else
				ValorCasteado = "'"&CStr(newKey)&"'"
		End Select
	convert = ValorCasteado
End Function
Function catchError(valor)
	if valor <> "" Then
		Dim ObjRegEx , ColMatches ,sLinea
		Set ObjRegEx = CreateObject("VBScript.RegExp")
		ObjRegEx.Global = True 'Esto hace que no solo busque en la primera que encuentre
		ObjRegEx.IgnoreCase = false  ' no sensible a las mayúsculas
		ObjRegEx.Pattern = "^\[ERRMSG\]"

		Set ColMatches =  ObjRegEx.Execute(valor)
		if colMatches.Count>0 then
			mensaje = Replace(valor,"[ERRMSG]","")
			valor = ""
			catchError = mensaje
		End if

	End If

End Function
Function corregirCaracteresSpeciales(sValor)
	sValor = Replace(sValor,"Ã³","ó")
	sValor = Replace(sValor,"Âº","º")
	sValor = Replace(sValor,"Ãº","ú")
	sValor = Replace(sValor,"Â°","°")
	sValor = Replace(sValor,"Ãº","ú")
	sValor = Replace(sValor,"Ã±","ñ")
	sValor = Replace(sValor,"Ã¡","á")
	sValor = Replace(sValor,"Ã‰","É")
	sValor = Replace(sValor,"Ã","í")
	corregirCaracteresSpeciales = sValor
End Function
Sub crearIdenti(val,Predecesor, sLineaActual)
	Dim valor, valor2
	valor = ""
	valor2 = ""
	if Ucase(Predecesor&val) = "OBJECT.DATE" Then
 		valor	= ObtenerValor(val, sLineaActual, Predecesor&val)
		valor = Replace(Replace(Replace(valor,"-","")," ",""),":","")
	End if
	' if UCase(Predecesor&val) = "OBJECT.ORDERID" Then
	' 	valor2	= ObtenerValor(val, sLineaActual, Predecesor&val)
	' End if
	' identi = identi & valor2
	valor = Replace(Replace(Replace(Replace(Replace(valor,"/","")," ",""),":",""),"PM","8077"),"AM","6577")
	identi = identi & valor
End Sub
Sub MaxValue(key)
	jsonField = dFieldKey.Item(key)

	if dKeyValue.Exists(jsonField) = true then
		ordenList = dKeyValue.Item(jsonField)
		IKmax = UBound(aListValue(ordenList))+1 'Cuanto valores hay para ese campo en mi JSON
		if (IKmax)> aKeyMaximo then
		 	aKeyMaximo = IKmax
		End If
	End If
End Sub
Sub Init()
	aListValue = Array()
	aItemValue = Array()
	Icontador = 0
	dKeyValue.RemoveAll
End Sub
Sub OpenConnection()
    Set oConn = CreateObject("ADODB.Connection")
    DBProperties.CompanyName = cNombre_Empresa
    oConn.Provider = "sqloledb"
    oConn.Properties("Data Source").Value = DBProperties.Server
    oConn.Properties("Initial Catalog").Value = DBProperties.Database
    oConn.Properties("User ID").Value = DBProperties.User
    oConn.Properties("Password").Value = DBProperties.Password
    oConn.Open
End Sub
Sub CloseConnection()
    oConn.Close
    Set oConn = Nothing
End Sub
Sub procesarOrdenes(archivo)
	Dim sFileUri,  oFile, FileFrom, FileTo
	Dim path
	path =  "C:\Temp\Procesados\"
	Set oFile = CreateObject("Scripting.FileSystemObject")
	If not oFile.FolderExists(path) then
			oFile.CreateFolder(path)
	End If
	FileFrom = "C:\Temp\Ordenes\" & archivo
	FileTo = path & archivo
	oFile.MoveFile FileFrom , FileTo
	Set oFile = Nothing
End Sub
Sub grabarLog_Archivo(pDato)
    Dim sFileUri,  oFile, File
    Dim path
    path =  "C:\log\"
    Set oFile = CreateObject("Scripting.FileSystemObject")

    If not oFile.FolderExists(path) then
        oFile.CreateFolder(path)
    End If

    sFileUri = path &"File_"& Replace(Replace(CStr(Date), "/", "-"), ":", ".") + ".log"

    If oFile.FileExists(sFileUri) Then
        Set File = oFile.OpenTextFile(sFileUri, 8, False)
    Else
        Set File = oFile.CreateTextFile(sFileUri)
    End If

    File.Write (CStr(Now) & " - " + pDato + vbCRLF)
    File.Close
End Sub

'----------------------------------------------------------------------
'----------------------- VALIDACIONES ---------------------------------
'----------------------------------------------------------------------
Sub creardValidaciones()
  Set dValidaciones = CreateObject("Scripting.Dictionary")
  'dValidaciones.Add "validate1", validate1
  dValidaciones.Add "validate2" , validate2
  dValidaciones.Add "validate3" , validate3
  dValidaciones.Add "validate4" , validate4
  dValidaciones.Add "validate5" , validate5
  dValidaciones.Add "validate6" , validate6
	dValidaciones.Add "validate7" , validate7
End Sub
Function validate1()
	sQuery = " SELECT COUNT(INTEQE_CODEQU) EXISTE FROM INTEQE "&_
					 " WHERE "&_
					 " INTEQE_CODIGO = '" & m_sListPrecio & "' and INTEQE_CODI01 = '" & ValorRecup & "'"
	Set resultSet = oConn.Execute(CStr(sQuery))
	existe = resultSet("EXISTE").value
	if existe = 0 Then
		validate1 = "[ERRMSG]No existe equivalencia para lista de precio "&ValorRecup
	else
		sQuery = " SELECT INTEQE_CODEQU INTEQE_CODEQU FROM INTEQE "&_
						 " WHERE "&_
						 " INTEQE_CODIGO = '" & m_sListPrecio & "' and INTEQE_CODI01 = '" & ValorRecup & "'"
		Set resultSet = oConn.Execute(CStr(sQuery))
		sCodProv = resultSet("INTEQE_CODEQU").value
		validate1 = ValorRecup
	End if
End Function
Function validate2()
  sValor = ValorRecup
  if (sValor<> "" or sValor<>"NULL") and Len(sValor)>4 then
  	 sValor = Mid(sValor,2,4)
  End if
  sQuery ="SELECT COUNT(GRTPAC_CODPOS) EXISTE FROM GRTPAC WHERE GRTPAC_CODPOS = '" & sValor &"'"
  Set resultSet = oConn.Execute(CStr(sQuery))
  sCodPos = resultSet("EXISTE").value
  if sCodPos = 0 Then
      validate2 = 1000
  else
      validate2 = sValor
  End if
End Function
Function validate3()
  'ValorRecup = Replace(Replace(Replace(Replace(Replace(ValorRecup,"ó","o"),"í","i"),"ú","u"),"á","a"),"é","e")
  sQuery = " SELECT COUNT(INTEQE_CODEQU) EXISTE FROM INTEQE "&_
           " WHERE "&_
           " INTEQE_CODIGO = '" & m_sCodProvincia & "' and INTEQE_CODI01 = '" & ValorRecup & "'"
  Set resultSet = oConn.Execute(CStr(sQuery))
  existe = resultSet("EXISTE").value
  if existe = 0 Then
    validate3 = "[ERRMSG]No existe equivalencia para "&ValorRecup
  else
    sQuery = " SELECT INTEQE_CODEQU INTEQE_CODEQU FROM INTEQE "&_
             " WHERE "&_
             " INTEQE_CODIGO = '" & m_sCodProvincia & "' and INTEQE_CODI01 = '" & ValorRecup & "'"
    Set resultSet = oConn.Execute(CStr(sQuery))
    sCodProv = resultSet("INTEQE_CODEQU").value
    validate3 = sCodProv
  End if
End Function
Function validate4()
   sQuery ="SELECT COUNT(STMPDH_ARTCOD) EXISTE FROM STMPDH WHERE STMPDH_ARTCOD = '" & ValorRecup & "'"
   Set resultSet = oConn.Execute(CStr(sQuery))
   sArtcod = resultSet("EXISTE").value

   if sArtcod = 0 Then
       validate4 = "9999"
			 noExisteArtCod = True
   else
       validate4 = ValorRecup
   End if
End Function
Function validate5()
  if noExisteArtCod = True Then
    validate5 = ValorRecup
  End if
End Function
Function validate6()
  if noExisteArtCod = True Then
    validate6 = "O"
	else
		validate6 = "V"
  End if
End Function
Function validate7()
  sQuery ="SELECT COUNT(VTMCLH_NROCTA) EXISTE FROM VTMCLH "&_
          "INNER jOIN VTMCLC ON VTMCLC_NROCTA = VTMCLH_NROCTA "&_
          "WHERE "&_
          "(VTMCLH_DIREML = '"& ValorRecup &"' "&_
          "OR VTMCLC_DIREML = '"& ValorRecup &"') "
  Set resultSet = oConn.Execute(CStr(sQuery))
  existe = resultSet("EXISTE")
  if existe  = 0 Then
      validate7 = ""
  else
  sQuery ="SELECT VTMCLH_NROCTA NROCTA FROM VTMCLH "&_
          "INNER jOIN VTMCLC ON VTMCLC_NROCTA = VTMCLH_NROCTA "&_
          "WHERE "&_
          "(VTMCLH_DIREML = '"& ValorRecup &"' "&_
          "OR VTMCLC_DIREML = '"& ValorRecup &"') "
  Set resultSet = oConn.Execute(CStr(sQuery))
      sNrocta = resultSet("NROCTA").value
      validate7 = sNrocta
  End if
End Function
