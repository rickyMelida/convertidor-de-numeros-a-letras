REM  *****  BASIC  *****
Global uni As Variant 
Global Unidades As Variant, Decenas As Variant, Centenas As Variant, NumerosRedondos As Variant, Mil As Variant, Millon As Variant
Global indiceUnidad As Variant, indiceDecena As Variant, indiceCentena As Variant, indiceUMil as Variant, cantidadDeDigitos As Variant

Sub CargarDatos
	Unidades = Array("Uno", "Dos", "Tres", "Cuatro", "Cinco", "Seis", "Siete", "Ocho", "Nueve")
	Decenas = Array("Dieci", "Veinti", "Treita", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa")
	Centenas = Array("Ciento", "Docientos", "Trecientos", "Cuatrocientos", "Quinientos", "Seiscientos", "Setecientos", "Ochocientos", "Novecientos")
	
	Mil = "Mil"
	Millon = Array("Millon", "Millones")
	
	NumerosRedondos = Array("Cero", "Diez", "Veinte", "Treita", "Cuarenta", "Cincuenta", "Sesenta", "Setenta", "Ochenta", "Noventa")
	
End Sub

Sub Main

End Sub


Function CONVERTIRALETRAS(numero As Long) As string
	Dim cantidadDeDigitos As Integer
	cantidadDeDigitos = Len(numero)
	
	Select Case cantidadDeDigitos
	Case 1
		CONVERTIRALETRAS = EsNumeroDeUnDigito(numero)
	Case 2
		CONVERTIRALETRAS = EsNumeroDeDosDigitos(numero)
	Case 3
		CONVERTIRALETRAS = EsNumeroDeTresDigitos(numero)
	Case 4
		CONVERTIRALETRAS = EsNumeroDeCuatroDigitos(numero)
	Case 5
		CONVERTIRALETRAS = EsNumeroDeCincoDigitos(numero)
	Case 6
		CONVERTIRALETRAS = EsNumeroDeSeisDigitos(numero)
	Case 7
		CONVERTIRALETRAS = EsNumeroDeSieteDigitos(numero)
	Case 8
		CONVERTIRALETRAS = EsNumeroDeOchoDigitos(numero)
	Case 9
		CONVERTIRALETRAS = EsNumeroDeNueveDigitos(numero)
	Case Else
		CONVERTIRALETRAS = "Numero demasiado alto."
	
	End Select
End Function






'-------------------------------------------------------------------------------------------- Decenas Exclusivas
Function DecenasEXclusivas(numero As Variant) As string
	Dim DecenasDeDiezExclusivas As Variant, indiceDecena As Variant
	
	DecenasDeDiezExclusivas = Array("Diez", "Once", "Doce", "Trece", "Catorce", "Quince", "Veinte")
	indiceDecena = Right(numero, 1)
	DecenasEXclusivas = DecenasDeDiezExclusivas(CInt(indiceDecena))
End Function



'-------------------------------------------------------------------------------------------- Comprobar Decenas Exclusivas
Function esDecenasExclusivo(numero as Variant) As Boolean
	 esDecenasExclusivo = numero > 10 And numero <= 15 

End Function



'-------------------------------------------------------------------------------------------- Coversion de Numeros de Un Digitos
Function EsNumeroDeUnDigito(numero As Integer) As String
	CargarDatos()
	indiceUnidad = numero
	EsNumeroDeUnDigito = Unidades(indiceUnidad-1)
End Function



'-------------------------------------------------------------------------------------------- Coversion de Numero de Dos Digitos
Function EsNumeroDeDosDigitos(numero As Integer) As String
	CargarDatos()
	indiceUnidad = CInt(Right(CStr(numero), 1))
    indiceDecena = CInt(Left(CStr(numero), 1))
    
    '--- 12		
    If esDecenasExclusivo(numero) then
    	EsNumeroDeDosDigitos = DecenasEXclusivas(numero)
    	exit Function
    End If	
    
    '--- 20
    If indiceUnidad = 0 then
    	EsNumeroDeDosDigitos = NumerosRedondos(indiceDecena)
    	exit Function
    End If

	'--- 52
    If numero > 30 And numero < 100 then
    	EsNumeroDeDosDigitos = Decenas(indiceDecena-1) +  " y "+ Unidades(indiceUnidad-1)
    '--- 21
    Else 
		EsNumeroDeDosDigitos = Decenas(indiceDecena-1) + " " + Unidades(indiceUnidad-1)
    End If
End Function




'-------------------------------------------------------------------------------------------- Coversion de Numero de Tres Digitos
Function EsNumeroDeTresDigitos(numero As Long) As String
	CargarDatos()
	indiceUnidad = CInt(Right(CStr(numero), 1))
    indiceDecena = CInt(Mid(numero, 2, 1))
    indiceCentena = CInt( Left(CStr(numero), 1) )
    
    '--- 100
    If numero = 100 then
    	EsNumeroDeTresDigitos = "Cien"
    	Exit Function
    End If
    		
    '---200		
    If indiceDecena = 0 And indiceUnidad = 0 then
    	EsNumeroDeTresDigitos = Centenas(indiceCentena-1)
    	Exit Function
    End If 
    
    '---102
    If indiceDecena = 0 then
    	EsNumeroDeTresDigitos = Centenas(indiceCentena-1) + " " + Unidades(indiceUnidad-1)
    	exit Function
    End If
    
    '---120
    If indiceUnidad = 0 then
    	EsNumeroDeTresDigitos = Centenas(indiceCentena-1) + " " + NumerosRedondos(indiceDecena-1)
    	exit Function
    End If    
    
    EsNumeroDeTresDigitos = Centenas(indiceCentena-1) + " " + Decenas(indiceDecena-1) +  " y "+ Unidades(indiceUnidad-1) 
	
End Function




'-------------------------------------------------------------------------------------------- Coversion de Numero de Cuatro Digitos
Function EsNumeroDeCuatroDigitos(numero As Integer) As String
	Dim nCentena As Integer
	CargarDatos()
	
	nCentena = CInt(Right(CStr(numero), 3))
	indiceUMil = CInt(Left(CStr(numero), 1))
	
	'--- 8000
	If nCentena = 000 And indiceUMil <> 1 then 
		EsNumeroDeCuatroDigitos = Unidades(indiceUMil-1) + " " + Mil
		exit Function
	End If
	
	If indiceUMil = 1 then
		'--- 1000
		If nCentena = 000 then
			EsNumeroDeCuatroDigitos = Mil
			exit Function
		'--- 1526
		Else
			EsNumeroDeCuatroDigitos = Mil + " " +EsNumeroDeTresDigitos(nCentena)
			exit Function
		EndIf
	End If
	
	EsNumeroDeCuatroDigitos = Unidades(indiceUMil-1) + " " + Mil + " " +EsNumeroDeTresDigitos(nCentena)
End Function 


'-------------------------------------------------------------------------------------------- Coversion de Numero de Cinco Digitos
Function EsNumeroDeCincoDigitos(numero As Long) As String
	Dim indiceDecenaMil As Integer, numeroMil As Long
	
	CargarDatos()
	
	numeroMil = CInt(Right(CStr(numero), 4))
	indiceDecenaMil = CInt(Left(CStr(numero), 1))
	
	'---90000
	If numeroMil = 0000 And indiceDecenaMil <> 1 then 
		EsNumeroDeCincoDigitos = NumerosRedondos(indiceDecenaMil) + " " + Mil
		exit Function
	End If
	
	'--- 10000
	If numeroMil = 0000 And indiceDecenaMil = 1 then
		EsNumeroDeCincoDigitos = "Diez" + Mil
		exit Function
	End If
	
	
	Dim primerosDosNumeros As Integer, tresUltimosNumeros As Long
	primerosDosNumeros = CInt(Left(CStr(numero), 2))
	tresUltimosNumeros = CInt(Right(CStr(numero), 3))
		 
	EsNumeroDeCincoDigitos = EsNumeroDeDosDigitos(primerosDosNumeros) + " " + Mil + " " + EsNumeroDeTresDigitos(tresUltimosNumeros)
	
End Function


'-------------------------------------------------------------------------------------------- Coversion de Numero de Seis Digitos
Function EsNumeroDeSeisDigitos(numero As Long) As String
	CargarDatos()
	
	Dim primerosTresNumeros As Integer, ultimosTresNumeros As Long
	
	primerosTresNumeros = CInt(Left(CStr(numero), 3))
	ultimosTresNumeros = CInt(Right(CStr(numero), 3))
	
	
	'---100000
	If primerosTresNumeros = 100 And ultimosTresNumeros = 000 then
		EsNumeroDeSeisDigitos = "Cien" + " " +Mil
		Exit Function
	End If
	
	'---800000
	If primerosTresNumeros <> 100 And ultimosTresNumeros = 000 then 
		EsNumeroDeSeisDigitos = Centenas(CInt(Left( CStr(numero),1)) -1) + Mil
		Exit Function
	End If
	
	EsNumeroDeSeisDigitos = EsNumeroDeTresDigitos(primerosTresNumeros) + " "  + Mil + " " + EsNumeroDeTresDigitos(ultimosTresNumeros)
	

End Function



'-------------------------------------------------------------------------------------------- Coversion de Numero de Siete Digitos
Function EsNumeroDeSieteDigitos(numero As Long) As String 
	CargarDatos()
	Dim primerNumero As Integer, seisUltimosNumeros As Long
	
	primerNumero = CInt(Left(CStr(numero), 1))
	seisUltimosNumeros = CLng(Right(CStr(numero), 6))
	
	If primerNumero = 1 And seisUltimosNumeros = 000000 then
		EsNumeroDeSieteDigitos = "Un " + Millon(primerNumero-1)
		exit Function
	End If
	
	If primerNumero <> 1 And seisUltimosNumeros = 000000 then
		EsNumeroDeSieteDigitos = EsNumeroDeUnDigito(primerNumero) + " " +Millon(1)
		exit Function
	End If
	
	If primerNumero = 1 then
		EsNumeroDeSieteDigitos = primerNumero + " " + Millon(primerNumero-1) + " " + EsNumeroDeSeisDigitos(seisUltimosNumeros)
		exit Function 
	End If
	
	EsNumeroDeSieteDigitos = EsNumeroDeUnDigito(primerNumero) + " " + Millon(1) + " " + EsNumeroDeSeisDigitos(seisUltimosNumeros)
	
End Function


'-------------------------------------------------------------------------------------------- Coversion de Numero de Ocho Digitos
Function EsNumeroDeOchoDigitos(numero As Long) As String
	CargarDatos()
	Dim primerosDosNumero As Integer, seisUltimosNumeros As Long
	
	primerosDosNumero = CInt(Left(CStr(numero), 2))
	seisUltimosNumeros = CLng(Right(CStr(numero), 6))
	
	If primerosDosNumero = 10 And seisUltimosNumeros = 000000 then
		EsNumeroDeOchoDigitos = EsNumeroDeDosDigitos(primerosDosNumero) + " " +Millon(0)
		exit Function
	End If
	
	If primerosDosNumero <> 10 And seisUltimosNumeros = 000000 then
		EsNumeroDeOchoDigitos = EsNumeroDeDosDigitos(primerosDosNumero) + " " +Millon(1)
		exit Function
	End If	
	
	
	EsNumeroDeOchoDigitos = EsNumeroDeDosDigitos(primerosDosNumero) + " " + Millon(1) + " " + EsNumeroDeSeisDigitos(seisUltimosNumeros)
	
End Function


'-------------------------------------------------------------------------------------------- Coversion de Numero de Nueve Digitos
Function EsNumeroDeNueveDigitos(numero As Long) As String
	CargarDatos()
	Dim primerosTresNumero As Integer, seisUltimosNumeros As Long
	
	primerosTresNumero = CInt(Left(CStr(numero), 3))
	seisUltimosNumeros = CLng(Right(CStr(numero), 6))
	
	'-- 100.000.000
	If primerosTresNumero = 100 And seisUltimosNumeros = 000000 then
		EsNumeroDeNueveDigitos = EsNumeroDeTresDigitos(primerosTresNumero) + " " +Millon(0)
		exit Function
	End If
	
	'--- 300.000.000
	If primerosTresNumero <> 100 And seisUltimosNumeros = 000000 then
		EsNumeroDeNueveDigitos = EsNumeroDeTresDigitos(primerosTresNumero) + " " +Millon(1)
		exit Function
	End If
	
	
	'--- 120.965.789
	EsNumeroDeNueveDigitos = EsNumeroDeTresDigitos(primerosTresNumero) + " " + Millon(1) + " " + EsNumeroDeSeisDigitos(seisUltimosNumeros)
End Function
