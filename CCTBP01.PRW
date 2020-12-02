#INCLUDE "rwmake.ch"
#Include "Protheus.ch"
#INCLUDE 'FONT.CH'
#INCLUDE 'COLORS.CH'
#INCLUDE "topconn.ch"

/*
+-------------+---------+---------+--------------------------------
------+
I Programa: | CCTBAP01I Autor: | José Alves - Anadi Consultoria Novembro/2016 |
+-------------+---------+------------------------------------------
------+
| Descrição: | Rotina de Importação de Lançamentos Contábeis	I
+------------------------------------------------------------------
------+
I Uso:	|
Coopmil
+-------------+----------------------------------------------------
----+
*/
//FONTE 06
//PG01
User Function CCTBP01()

	Local _cPath := GetTempPath()
	Private _cLote 	:= ""
	Private _lRet 	:= .F.
	Private nLinCont:= 0 
	Private aDados 	:= {}
	Private _aHeader:= {}
	Private _aErros := {}
	Private _nHdl := 0, _cEOL := CHR(13) + CHR(10)

	If Select("SX2") == 0 
		RpcSetType(3)
		RpcSetEnv("01","01",,,"CTB",GetEnvServer(), {"SA1"})
	Endif

	_cLote := GetMV("MV__LOTIMP")

	//Verifica se a empresa/Filial permite importação de lançamentos contábeis
	If !("**" $ AllTrim(GetMv("MV__CTBP2G"))) .And. !(cFilAnt $ Alltrim(GetMv("MV__CTBP2G")))
		MsgStop("Empresa/Filial não habilitada para importação de lançamentos contábeis")
		Return
	EndIF
	MsgRun("Lendo Arquivo Texto", "Aguarde...",{|| LerArq()})

	If _lRet
		MsgRun("Ajustando Informações","Aguarde...",{|| AjustaDados(nLinCont+1)})
		MsgRun("Importando Lançamentos Contábeis","Aguarde...",{|| ImportaDados(nLinCont+1)}) 
		Msglnfo("Importação Finalizada")
	Else
		Alert("Layout em desacordo com o arquivo a ser importado!")
	Endif
	
	//Gravação do arquivo de Log 
	If Len(_aErros) > 0 
		If Empty(_cPath)
			_cPath := AllTrim(cGetFile("Arquivos txt|*.txt",'Selecione o Diretório Destino', 1,'C:\TEMP\',.F., nOR(GETF_LOCALHARD, GETF_RETDIRECTORY),.F.,.T.))
			While Empty(_cPath)
				MsgStop("0 diretorio destino nâo foi informado.","ATENCAO")
			EndDo
		Endif

		//PG02
		_cPath := _cPath + "LOG_IMPORT_LANCTOS_CONTABEIS_" + DTOS(Date()) + "_" + StrTran(time(),":","") + ".txt"

		//Criação do Arquivo
		If (_nHdl := MsfCreate(_cPath,0))==	-1

			//Se nao conseguiu criar o arquivo, exibe tela com LOG
			CCTBP01LOG()
		Else

			For nx := 1 To Len(_aErros)
				If fWrite(_nHdl, _aErros[nx][1] + " " + _aErros[nx][2] + " " + _aErroa[nx][3] + " " + _aErros[nx][4] + " ---- " + _aErros[nx][5] + cEOL) < 1 
					fClose(_nHdl)

					//Se não conseguiu gravar no arquivo, exibe tela com log 
					CCTBPOILOG()
					Return 
				Endif 
			Next nx

			If File(_cPath)
				ShellExecute("Open",_cPath,"","C:\", 1 ) 
			Endif 
		Endif 
	Endif
Return

Static Function LerArq()

	local cFile 	:= cGetFile("*.txt|*.txt","Selecione o Arquivo a ser lido",0,,.t., GETF_LOCALHARD+GETF_LOCALFLOPPY+GETF_NETWORKDRIVE)
	Local cLinha 	:= ""
	Local nLinha 	:= 0 
	Local nTamLinha := 0
	Local nTamArq	:= 0 
	Local Low 		:= 0

	cArqImpor := cFile

	//Valida arquivo 

	If !file (cArqImpor)
		Aviso("Arquivo", "Arquivo não selecionado ou invalido.", {"Sair"},1)
		Return
	Else
		//+--------------*--------------------------------------------------------+
		//| Abertura do arquivo texto	|
		//+-----------------------------------------------------------------------+
		nHdl := fOpen(cArqImpor)

		If nHdl== -1
			IF FERROR()== 516
				ALERT("Feche o arquivo.")
			EndIF 
		Endif
		//+------------------------------------------------------------------------
		//| Verifica se foi possivel abrir o arquivo	l
		//+------------------------------------------------------------------------

		If nHdl == -1
			cMsg := "O arquivo de nome " + cArqImpor + " nao pode ser aberto! Verifique os parâmetros."
			MsgAlert(cMsg,"Atencao!")
			//PG02
			Return
		Endif

		// +-----------------------
		//I Posiciona no Inicio do ARQUIVO
		//+------------------------
		FSEEK(nhdl,0,0)

		//+-------------------------------
		//I Traz o Tamanho do Arquivo TXT
		//+-------------------------------
		nTamArq:=FSEEK(nHdl,0,2)
		// +.............................
		//I Posicona novamemte no inicio
		//+--............................
		FSEEK(nHdl,0,0)
		//+.....................................-....................
		//| Fecha o Arquivo
		//+----------------------------------------------------------
		fClose(nHdl)
		FT_FUse(cArqImpor) //abre o arquivo
		FT_FGOTOP()	//posiciona na primeira linha do arquivo
		nTamLinha := Len(FT_FREADLN()) //Ve o tamanho da linha 
		FT_FGOTOP()
		//+----------------------------------------------------------
		//| Verifica quantas linhas tem o arquivo
		//+----------------------------------------------------------
		nLinhas := nTamArq/nTamLinha

		ProcRegua(nLinhas)

		aDados:={}

		While !FT_FEOF() //Ler todo o arquivo enquanto não for o final dele

			IncProc('Importando...')
			clinha := FT_FREADLN()
			aadd(aDados,Separa(cLinha,CHR(09),.T.))
			FT_FSKIP()
		EndDo
		FT_FUse()
		fClose(nHdl)
	Endif

	ProcRegua(len(aDados))

	_aHeader := BuscaLayout()

	_nCont:= 0
	_lRet := .T.

	If Len(_aHeader) > 1
		For _I := 1 to len(aDados)
			For Y := 1 To Len(_aHeader)
				_C1 := UPPER(LIMPATXT(Alltrim(aDados[_I][Y]),""))
				_C2	:= UPPER(LIMPATXT(Alltrim(_aHeader[Y][1]),""))
				If _C1 == _C2 
					_nCont++
				Endif
			Next 
			_Low := 0
			//PG03
			For I := 1 To Len(_aHeader)
				If !Empty(alltrim(_aHeader[1][2]))
					_low++
				EndIF
			Next
			If _nCont == _low
				_lRet:= .t.
				Exit
			Else
				_nCont :=0
				_lRet:= .f.
			EndIF
		Next
	EndIF

	nLinCont := _I 

Return _lRet 

Static Function BuscaLayout()

	cQuery := ""
	aDadosH := {}


	DbSelectArea("SZ2")
	SZ2->(DbSetOrder(3))
	If DbSeek(xFilial("SZ2") + "CCTBAP01")

		IF Select("TRB") > 0 
			DbSelectArea("TRB")
			TRB->(DbCloseArea())
		Endif

		cQuery := "SELECT * FROM " + RetSqlName("SZ3") + " WHERE Z3_CODIGO = '" + SZ2->Z2_CODIGO + "' AND Z3_POSCOL <>'' AND D_E_L_E_T_ = '' ORDER BY Z3_POSCOL"
		TcQuery cQuery New Alias "TRB"

		DbSelectArea("TRB")
		TRB->(DbGoTop())

		While !TRB->(Eof())
			aAdd(aDadosH,{TRB->Z3_NMCOL,TRB->Z3_POSCOL,TRB->Z3_CPDEST, TRB->Z3_FORMULA})
			TRB->(DbSkip())
		EndDo
	Else
		Alert("Layout não cadastrado para a importação de lançamentos contábeis CCTBAP01") 
	Endif
	*/
Return aDadosH


Static Function AjustaDados(_Con)

	Local aArrayl 	:= aDados
	Local _nPosData	:=	ASCAN(_aHeader,	{ |x| AllTrim(x[3])	==	"CT2_DATA"})	
	Local _nPoaDeb 	:=	ASCAN(_aHeader,	{ |x| AllTrim(x[3])	==	"CT2_DEBITO"})
	Local _nPosCre 	:=	ASCAN(_aHeader,	{ |x| AllTrim(x[3])	==	"CT2_CREDIT"})
	Local _nPosHis 	:=	ASCAN(_aHeader,	{ |x| AllTrim(x[3])	==	"CT2_HIST"})
	Local _nPosVlr	:=	ASCAN(_aHeader,	{ |x| AllTrim(x[3])	==	"CT2_VALOR"})

	//PAG05	
	For K := _Con To Len(aArrayL)
		aArrayL[K] [_nPosData] 	:= DtoS(CtoD(Alltrim(aArrayL[K][_nPosData]))) 
		aArrayL[K] [_nPosDeb] 	:= Alltrim(StrTran(StrTran (aArrayL[K][_nPosDeb],".",""),"-",""))
		aArrayL[K] [_nPosCre] 	:= Alltrim(StrTran(StrTran (aArrayL[K][_nPosCre],".",""),"-",""))
		aArrayL[K] [_nPosHis] 	:= UPPER(LIMPATXT(aArrayL[k][_nPosHis],""))
		_nValor	:= Replace(aArrayL[K][_nPosVlr],".","")
		_nValor	:= Replace(_nValor," ," ,".")		
		aArrayL[K][_nPosVlr] := _nValor
	Next

	aDados := aArrayL

Return

Static Function ImportaDados(nNumLin)

	Local aDia := {} 
	Local nTotDeb := 0 
	Local nTotCred := 0

	Local dDtDia := ""
	Local dDtDiaAnt := ""
	Local _nPosData := ASCAN(_aHeader, {|x| AllTrim (x[3]) == "CT2_DATA" })
	Local dDtDiaAnt := aDados[nNumLin][_nPosData]

	For I := nNumLin To Len (aDados)
		dDtDia := aDados[I][_nPosData]
		If dDtDia == dDtDiaAnt 
			aAdd(aDia,aDados[I])
		Else
			ValidaDia(aDia)
			dDtDiaAnt := aDados[I][_nPosData]
			aDia := {}
			aAdd(aDia,aDados[I])
		EndIf
	Next

	//Para Validar o Ultimo Dia ou Caso o Arquivo Tenha Somente 1 Dia 
	ValidaDia (aDia)

Return


Static Function ValidaDia(_aDia)

	Local aArray := _aDia 
	Local aGrava := {}

	Local _nPoeData	:= ASCAN(_aHeader, { |x| AllTrim(x[3]) == "CT2_DATA"})
	Local _nPosDeb 	:= ASCAN(_aHeader, { |x| AllTrim(x[3]) == "CT2_DEBITO"}) 
	Local _nPosCre	:= ASCAN(_aHeader, { |x| AllTrim(x[3]) == "CT2_CREDIT"})
	Local _nPosHis 	:= ASCAN(_aHeader, { |x| AllTrim(x[3]) == "CT2_HIST"})
	Local nPosVlr	:= ASCAN(_aHeader, { |x| AllTrim(x[3]) == "CT2_VALOR"})


	//Validações

	For J := l To Len(aArray)

		lErro := .F.

		//PG06
		//Valida se existe Conta Debito
		DbSelectArea("CT1")
		
		If !Empty(Alltrim(aArray[J][_nPosDeb]))
			If DbSeek(xFilial("CT1") + PADR(Alltrim(aArray[J] [_nPosDeb]) ,Len (CT1->CT1_CONTA)))
				//Valida se conta está bloqueada 
				If CT1->CT1_BLOQ = "1"
					aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J] [_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta credito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: " + aArray[J][_nPosVlr],"CONTA DEBITO BLOQUEADA PARA LANÇAMENTOS"})
					lErro := .T.
				EndIF
			Else
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J] [_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta Crédito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: " + aArray[J][_nPosVlr],"CONTA DEBITO NAO EXISTE"}) 
				lErro := .T.
			EndIf
		EndIf

/*
		//Valida se existe Conta Credito 
		If !Empty (Alltrim(aArray[J] [_nPosCre]))
			If DbSeek(xFilial("CT1") + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)))
				//Valida se conta está bloqueada 
				If CT1->CT1_BLOQ = "1"
					aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J][_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J] [_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta Crédito:" + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: "+ aArray[J][_nPosVlr],"CONTA CREDITO BLOQUEADA PARA LANÇAMENTOS"}) 
					lErro := .T.
				EndIf
			Else
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J] [_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb)),Len(CT1->CT1_CONTA)),"Conta Crédito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: "+ aArray[J][_nPosVlr],"CONTA CREDITO NAO EXISTE"})
				lErro := .T.
			EndIf
		EndIf
*/
		If !Empty(Alltrim(aArray[J] [_nPosDeb])) .And. !Empty(Alltrim(aArray[J] [_nPosCre]))
			//Valida se contas Partida Dobrada sao diferentes
			If PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)) == PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA))
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J] [_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta Crédito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: " + aArray[J][_nPosVlr],"ERRO CONTAS IGUAIS EM LANÇAMENTO DE PARTIDA DOBRADA"}) 
				lErro := .T.
			EndIf
		EndIf
		//Valida se ja existe o lançamento em partida dobrada
		If !Empty(Alltrim(aArray[J][_nPosDeb])) .And. !Empty(Alltrim(aArray[J][_nPosCre])) 
			DbSelectArea("CT2")
			CT2->(DbSetOrder(17))
			If DbSeek(xFilial("CT2") + aArray[J][_nPosData] + PADR(Alltrim(_cLote),Len(CT2->CT2_LOTE)) + PADR(Alltrim(aArray[J][_nPosDeb]), Len(CT1->CT1_CONTA)) + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA))+ aArray[J][_nPosVlrl])
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J][_nPosData])),"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta Crédito: "+PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: "+ aArray[J][_nPosVlr],"LANÇAMENTO PARTIDA DOBRADA JÁ IMPORTADO ANTERIORMENTE"}) 
				lErro := .T.
			EndIf
		EndIf
	//PG07
	//Valida se ja existe o lançamento a debito
		If !Empty(Alltrim(aArray[J][_nPosDeb])) .And. Empty(Alltrim(aArray[J][_nPosCre])) 
			DbSelectArea("CT2")
			CT2->(DbSetOrder(18))

			IF DbSeek(xFilial("CT2") + aArray[J][_nPosData] + PADR(Alltrim(_cLote), Len(CT2->CT2_LOTE)) + PADR(Alltrim(aArray[J] [_nPosDeb]),Len(CT1->CT1_CONTA))+ aArray[J][_nPosVlr])
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J][_nPosData]))/"Conta Debito: " + PADR(Alltrim(aArray[J][_nPosDeb]),Len(CT1->CT1_CONTA)),"Conta Crédito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: "+ aArray[J][_nPosVlr],"LANÇAMENTO A DEBITO JÁ IMPORTADO ANTERIORMENTE"}) 
				lErro := .T.
			Endif
		Endif

	//Valida se-ja existe o lançamento a credito
		If !Empty(Alltrim(aArray[J][_nPosCre])) .And. Empty(Alltrim(aArray[J][_nPosDeb])) 
			DbSelectArea("CT2")
			CT2->(DbSetOrder(19))
			If DbSeek(xFilial ("CT2")+aArray[J][_nPosData] + PADR(Alltrim(_cLote), Len(CT2->CT2_LOTE)) + PADR(Alltrim(aArray[J][_nPosCre]),Len(CTl->CT1_CONTA))+ aArray[J][_nPosVlr]) 
				aAdd(_aErros,{"Data Lançamento:" + DtoC(StoD(aArray[J] [_nPosData])),"Conta Debito:	" + PADR(Alltrim (aArray[J] [_nPosDeb]), Len(CT1->CT1_CONTA)), "Conta Crédito: " + PADR(Alltrim(aArray[J][_nPosCre]),Len(CT1->CT1_CONTA)),"Valor: " + aArray[J][_nPosVlr],"LANÇAMENTO A CREDITO JÁ IMPORTADO ANTERIORMENTE"}) 
				lErro := .T.
			Endif
		Endif


		IF !lErro
			aAdd(aGrava,aArray[J]) 
		EndIf
	Next
	
	If Len(aGrava) > 0 
		GravaDia(aGrava) 
	Endif

Return

Static Function GravaDia(a_Dia)

	Local aCab := {}
	Local altens := {} 
	Local IMaisLote := .F. 
	Local nLinha := 0 
	Local nDoc := 1
	Local cDifDC :=	GetMV("MV_CONTSB")					
	Local _cTp := ""							
	Local nPosData	:= ASCAN (_aHeader,	{|x| AllTrim(x[3]) == "CT2_DATA"})
	Local nPosDeb	:= ASCAN (_aHeader,	{|x| AllTrim(x[3]) == "CT2_DEBITO"})
	Local _nPosCre	:= ASCAN (_aHeader,	{|X| AllTrim(x[3]) == "CT2_CREDIT"})
	Local nPosHis	:= ASCAN (_aHeader,	{|X| AllTrim(x[3]) == "CT2_HIST"})
	Local _nPosVlr	:= ASCAN (_aHeader,	{|X| AllTrim(x[3]) == "CT2_VALOR"})
	Local _aDados := {}
	Local _cErroExec := ""

	Private lMsErroAuto := .f.

	DbSelectArea("SZ2")
	SZ2->(DbSetOrder(3))
	DbSeek(xFilial("SZ2") + "CCTBAP01")
	If Select("TRBA") > 0
		DbSelectArea("TRBA")
		DbCloseArea ()
	Endif


	_cQuery := "SELECT * FROM " + RetSqlName("SZ3") +" WHERE Z3_CODIGO = '" + SZ2->Z2_CODIGO + "' AND D_E_L_E_T_ = '' ORDER BY Z3_POSCOL
	TcQuery cQuery New Alias "TRBA"

	//PG08
	DbSelectArea("TRBA")
	TRBA->(DbGoTop())


	While !TRBA->(EOF())
		aADD(_aDados,{TRBA->Z3_NMCOL,TRBA->Z3_POSCOL,TRBA->Z3_CPDEST,TRBA->Z3_FORMULA})
		TRBA->(DbSkip())
	EndDo

	PutMV("MV_CONTSB","S")

	aAdd(aCab, {'DDATALANC'		,StoD(a_Dia[1][_nPosData])			,NIL})
	aAdd(aCab, {'CLOTE'			,GetMV("MV__LOTIMP")				,NIL})    //CLAUDIO DUVIDA
	aAdd(aCab, {'CSUBLOTE'		,'001'								,NIL})
	aAdd(aCab, {'CDOC'			,StrZero(nDoc,6)					,NIL})
	aAdd(aCab, {'CPADRAO'		,''									,NIL})
	aAdd(aCab, {'NTOTINF'		,0									,NIL})
	aAdd(aCab, {'NTOTINFLOT'	,0									,NIL})

	nLinha := 0

	For K := 1 To Len(a_Dia)
		//nCont++
		If nLinha > 800
			MSExecAuto({|x,y,z| CTBA102(x,y, z)}, aCab ,altens, 3)

			If lMsErroAuto
				lMsErroAuto := .F.
				_cErroExec := Mostraerro("C:\temp\LOG\ERROINTEGRACAO.log") 

				aAdd(_aErros,{"","","","",_cErroExec})
			EndIF

			aCab := {} 
			altens := {} 
			nDoc++ 
			nLinha := 0

			aAdd(aCab,	{'DDATALANC'	,StoD(a_Dia[1][_nPosData])	,NIL})
			aAdd(aCab,	{'CLOTE'		,GetMV("MV__LOTIMP")		,NIL})
			aAdd(aCab,	{'CSUBLOTE'		,'001'						,NIL})
			aAdd(aCab,	{'CDOC'			,StrZero(nDoc,6)			,NIL})
			aAdd(aCab,	{'CPADRAO'		,''							,NIL})
			aAdd(aCab,	{'NTOTINF'		,0							,NIL})
			aAdd(aCab,	{'NTOTINFLOT'	,0							,NIL})
		EndIf

		nLinha++

		If !Empty(Alltrim(a_Dia[K][_nPosDeb])) .And. !Empty(Alltrim(a_Dia[K][_nPosCre]))
			_cTp := "3"
		Elseif !Empty(Alltrim(a_Dia[K][_nPosDeb])).And. Empty(Alltrim(a_Dia[K][ nPosCre])) 
			_cTp := "1"
		Elseif Empty(Alltrim(a_Dia[K][_nPosDeb])).And. !Empty(Alltrim(a_Dia[K][ nPosCre])) 
			_cTp := "2"
		Endif

		//_AHEADER[X][2] -> POSICAO 
		// AHEADER[X][3] -> CAMPO QUE IRA RECEBER 
		//_aHEADER[X][4] -> FORMULA

		//PAREI AQUI CLAUDIO

		//PG09
		aAdd(altens,{	{'CT2_LINHA'	,StrZero(nLinha,3)				,NIL},;
		{'CT2_DC'		,cTp							,NIL},;
		{'CT2_DEBITO'	,Alltrim(a_Dia[K][_nPosDeb])	,NIL},;
		{'CT2_CREDIT'	,Alltrim(a_Dia[K][_nPosCre])	,NIL},;
		{'CT2_VALOR'	,Val(a_Dia[K][_nPosVlr])		,NIL},;
		{'CT2_HIST'		,Alltrim(a_Dia[K][_nPosHis])	,NIL}})

		For F := 1 To Len(_aDados)
			_nPos := ASCAN(aItens[nLinha], {|x| AllTrim(x[1]) == Alltrim(_aDados[F][3])}) 
			IF _nPos < 1
				aAdd(aItens[nLinha],{_aDados[F][3],	IIF(!Empty(_aDados[F][2]),Alltrim(a_Dia[K][Val(_aDados[F][2])]),&(_aDados[F][4])),NIL})	
			Endif
		Next

		//ANALISA A NECESSIDADE DE INCLUIR LINHAS DE CONTINUAÇÃO DE HISTORICO 

		dbSelectArea("CT5")
		nTamCt5Hist := iif(GetNewpar("MV_CT5HIST",.T.), TamSx3("CT2_HIST")[1], Len(Alltrim(a_Dia[K][_nPosHis])))

		cHistorico := Alltrim( a_Dia[K][_nPosHis])

		nLen := Len(CriaVar("CT2_HIST"))

		nSeqHis := 1 
		cSeqLanc := nLinha

		If Len(cHistorico) > nLen
			For nCoUnt:= nLen+1 To Len(cHistorico) Step nLen 
				cHist := Substr(cHistorico,nCoUnt,nLen)
				nLinha++
				nSeqHis++

				aAdd(altens,{ 	{'CT2_FILIAL'	,xFilial("CT2")			,NIL},;
				{'CT2_LINHA'	,StrZero(nLinha,3)		,NIL},;
				{'CT2 DC'		,"4"					,NIL},;
				{'CT2_ORIGEM'	,'CCTBAP01'				,NIL},;
				{'CT2 FLAG'		,.F.					,NIL},;
				{'CT2_SEQLAN'	,StrZero(cSeqLanc, 3)	,NIL},;
				{'CT2_SEQHIS'	,StrZero(nSeqHis, 3)	,NIL},;
				{'CT2_HIST'		,cHist					,NIL}})
			Next

		Endif
	Next

	MSExecAuto({|x, y,z| CTBA102(x,y,z)}, aCab, aItens, 3)
	If lMsErroAuto
		lMsErroAuto := .F.
		_cErroExec := Mostraerro("C:\temp\LOG\ERROINTEGRACAO.log")
		aAdd(_aErros,{"","","","",_cErroExec})
	EndIf

	PutMV("MV_CONTSB",cDifDC) 

Return

//PG10 
Static Function LIMPATXT( cCpoLmp, cNaoLimpa )

	Local cAcentos := "çÇáéíóúÁÉÍÓÚâêîôûÂÊÎÔÛàèìòùÀÈÌÒÙäëïöüÄËÏÖÜãõÃÕ!âè'/èiÒÕ•i§ÕU£úSfi%$^"
	Local cAcSubst := "cCCcaeiouAEIOUaeiouAEIOUaoAOAAAAaaaaaaEEeeeIiiOOooooooUuuuNnPScoiooaq"
	Local cCaraPont:= "/\'I0UaeÍOUAEI0^aeÍ0U^'I0Uao^0'^^aaaaaaE^'eeeIii^0oooooo^uuu^nPSC0^’OOaq"
	local nI := 0 
	Local nPos	:= 0
	Local  aTira	:=	{}
	Local  cTexto	:=	""


	cCpoLmp := AllTrim( cCpoLmp )
	cCpoLmp := NoAcento( AnsiToOem( cCpoLmp ) )

	aTira := StrToArray (cNaoLimpa, ";")

	For I := 1 To Len (aTira) 
		cTexto 		:= aTira[I]
		cCaraPont 	:= Replace(cCaraPont,cTexto,"")
	Next


	//TROCA ACENTOS

	For nI := 1 To Len( cCpoLmp )

		If ( nPos := At( SubStr( cCpoLmp, nI, 1 ), cAcentos ) ) > 0
			cCpoLmp := SubStr(cCpoLmp, 1, nI - 1 ) + SubStr( cAcSubst, nPos, 1 ) + SubStr( cCpoLmp, nI + 1 )
		EndIF
	Next nI

	//Tira Caracteres de pontuacao
	For nI := 1 To Len(cCpoLmp)
		IF (nPos := At(SubStr(cCpoLmp, nI, 1), cCaraPont ) ) > 0
			cCpoLmp := SubStr(cCpoLmp, 1, nI - 1) + '#' + SubStr(cCpoLmp, nI + 1)
		EndIF 
	Next nI

	cCpoLmp := StrTran( cCpoLmp,	"" )
	cCpoLmp := StrTran ( cCpoLmp, "#",	)

Return( cCpoLmp )

/*
+-------------+-------------+-------+-------------------------------------------+--------+
| I	^rama:	I	CCTBPOILOG	|	Autor:	|	Jorge Henrique Alves - Anadi Consultoria	|	Data:	|
Dezembro/2016 |
+-------------+-------------+-------+--------------------------------------------+------+.
--------+
I Descrição: I Exibe Log de inconsistências, quando não for possivel gerar o arquivo	I
+-------------+—--------------------------------------------------------------------------
( Uso:	I
Coopmil	|
*/

Static Function CCTBPOILOG()

	Local oButton1, oMultiGe1 
	Local cMultiGe1 := ""
	Local nx := 1 
	Local oSay1, oSay2, _oDlg1, _oFont1

	For nx := 1 To Len(_aErros)
		cMultiGe1 += "*" + _aErros[nx][1] + " " + _aErros[nx][2] + " " + _aErros[nx][3] + " " +	_aErros[nx][4] + " ---- " + _aErros[nx][5] + _cEOL
	Next nx

	DEFINE FONT  _oFont1 Name "Tahoma" Size 0, -11 BOLD
	DEFINE MSDIALOG _oDlg1 TITLE "LOG - Importacao Lancamentos Contabeis" FROM 000, 000 TO 500, 700 OF _oDlg1 PIXEL 
	@ 005, 002 SAY oSay1 PROMPT "Nao foi possivel gerar o arquivo de LOG" SIZE 150, 007 OF _oDlg1 COLORS 255, 16777215 FONT _oFont1 PIXEL
	@ 015, 002 SAY oSay2 PROMPT "Abaixo inconsistencias na importacao de lancamentos contabeis" SIZE 277, 007 OF _oDlg1 COLORS 255, 16777215 FONT _oFont1 PIXEL 
	@ 030, 002 GET oMultiGe1 VAR cMultiGe1 OF _oDlg1 MULTILINE SIZE 340, 195 READONLY HSCROLL PIXEL
	@ 235, 305 BUTTON oButton1 PROMPT "OK" SIZE 037, 012 OF _oDlg1 PIXEL ACTION _oDlg1:End() 

	ACTIVATE MSDIALOG _oDlg1 CENTERED
	
	
Return
