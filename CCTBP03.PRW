#INCLUDE "protheus.CH"
#INCLUDE "RPTDEF.CH" 
#INCLUDE "TBICONN.CH"
#INCLUDE "TOPCONN.Ch"

Static _cPerg := PADR ("CCTBP03", Len (SX1->X1_GRUPO))

//FONTE07
//PAGINA 1
User Function CCTBP03()

	Local _lContinua := .t.
	Private _cFile := _xcTab := "", _cRotina := PADR("CCTBP03",TamSX3("Z2_ROTINA")[1]), _nHdl := 0
	

	//Verifica se existe layout de importação cadastrado 
	DbSelectArea("SZ2")
	DbSetOrder(3)
	IF !DbSeek (xFilial ("SZ2") + _cRotina)
		MsgStop("Layout não cadastrado para a rotina " + _cRotina,"Atenção")
		Return
	ElseIf SZ2->Z2_TIPO != "2"
		MsgStop("O layout possui TIPO diferente do esperado, o layout deve ser cadastrado como TXT com Posições"/"Atenção")
		Return
	EndIf

	//Verifica se a empresa/Filial permite exportação de lançamentos
	If !("**" $ Alltrim(GetMv("MV__CTBP3F"))) .And. !(cFilAnt $ Alltrim(GetMv("MV__CTBP3F")))
		MsgStop("Empresa/Filial não habilitada para exportação de lançamentos contábeis")
		Return
	EndIf


	//Parâmetros para geração do arquivo
	CCTBP03X1()
		Pergunte(_cPerg,.f.)
		If !Pergunte(_cPerg,.t.)
		Return 
	EndIf
	
	


	//Pasta para salvar o arquivo
	//_cFile := cGetFile ("*.txt|*.txt","Selecione o Diretório Destino",1,"C:\Temp\",.f.,GETF_LOCALHARD+GETF_LOCALFLOPPY+GETF_NETWORKDRIVE,.F., .T. ) 
	_cFile := AllTrim(cGetFile("Arquivos txt|*.txt","Selecione o Diretório Destino1", 1,"C:\TEMP\",.F., nOR(GETF_LOCALHARD, GETF_RETDIRECTORY),.F.,.T. ))

	If Empty(_cFile)
		MsgStop("Local para salvar o arquivo não foi informado","Atenção")
		Return
	EndIf

	_cFile := _cFile + mv_par06 + "_" + DTOS( Date()) + "_" + StrTran( Time(),":","" ) + ".txt"

	If (_nHdl := MsfCreate(_cFile,0)) == -1
		MsgStop("Não foi possivel gerar o arquivo " + _cFile,"Atenção")
		Return
	EndIf

	_xcTab := GetNextAlias() //Alias da tabela temporária com CT2 Filtrado 
	MsAguarde ({|| _lContinua := CCTBP03A()} , "Processando...", "Filtrando os registros...")

	//PAGINA 2

	If _lContinua
		MsAguarde({|| CCTBP03B()} , "Processando...", "Gerando o arquivo...")
	EndIF

	If Select(_xcTab) > 0 
		DbSelectArea(_xcTab)
		DbCloseArea()
	EndIf

	fClose(_nHdl)
Return

/*
+--------------+-----------+--------------------------------------------------------------------+
I Programa: | CCTBP03A | Autor: | CLAUDIO AMBROSINI - Ethosx Consultoria | Data: 13/09/2019     |
+--------------+-----------+-----------------------------------------------------------+--------+
-------1-
I Descrição: | Filtra os registros da CT2	|
+----------------------------------------------------------------------------------------
-------+
I Uso:	|
Coopmil	I
---------------------------------------------------------------------------------------
--+
*/

Static Function CCTBP03A()

	Local _lRet := .t., _cLote := _cSQL:= ""
	Pergunte(_cPerg,.f.)

	If !Empty(mv_par03)
		_cLote := Alltrim(Posicione("SX5",1,xFilial("SX5") + "09" + mv_par03, "X5_DESCRI"))
		_cLote := StrZero(Val(_cLote) ,TamSX3("CT2_LOTE")[1])

		If Empty(_cLote)
			MsgStop( "O lote informado é inválido. Verifique os parâmetros","ATENCAO" )
			Return .F.
		Endif
	Endif

	_cSQL := " Select CT2.CT2_FILIAL, CT2.R_E_C_N_O_ CT2_NREC From " + RetSqlName("CT2") + " CT2 "	
	_cSQL += " Where" 
	If mv_par05 == 1	 
		_cSQL += " CT2.CT2_FILIAL = '" + xFilial("CT2") + "' And "		
	EndIf					
	_cSQL += " CT2.CT2_DATA Between '" + DTOS(mv_par01) + "' And '" + DTOS(mv_par02) + "' And "
	If !Empty(mv_par03)					
		_cSQL += "CT2.CT2_LOTE = '"	+ _cLote + "' And "			
	EndIf
	_cSQL += " CT2.D_E_L_E_T_ = '' And CT2.CT2_VALOR > 0 And CT2.CT2__IMP = '' "
	_cSQL += " Order By CT2.CT2_FILIAL,CT2.R_E_C_N_O_ "

	If Select(_xcTab) > 0 
		DbSelectArea(_xcTab)
		DbCloseArea()
	EndIf

	DbUseArea(.T.,"TOPCONN",TcGenQry(,,_cSQL),_xcTab,.T.,.T.)

	DbSelectArea(_xcTab)
	DbGoTop()

	If Eof()
		_lRet:= .f.
		MsgStop("Não foram encontrados lançamentos contábeis para os parâmetros informados","Atenção") 
	Endif

Return _lRet 

//PAGINA 3

Static Function CCTBP03B()

	Local _cSQL := _cLinha := "", _cEOL := CHR(13) + CHR(10), _lErro := .F.
	Private _nReg := 0, _cSZ3 := "",_cCntCred:= _cCntDeb:= _cTpMov := "", _nSequen:=0
	

	//Busca Layout para exportação do arquivo
	_cSZ3 := GetNextAlias()

	_cSQL := " SELECT Z3_TAMANHO,Z3_POSINI,Z3_FORMULA FROM " + RetSqlName("SZ3") + " Z3 "
	_cSQL += " WHERE Z3_FILIAL = '" + xFilial ("SZ3") + "' and Z3_CODIGO = '" + SZ2->Z2_CODIGO + "' And Z3.D_E_L_E_T_ = '' "
	_cSQL += " Order By Z3_POSINI"

	If Select(_cSZ3) > 0
		DbSelectArea(_cSZ3)
		DbCloseArea()
	EndIf
	DbUseArea(.T.,"TOPCONN",TcGenQry(,,_cSQL),_cSZ3,.T.,.T.)

	DbSelectArea(_xcTab)
	DbGoTop()

	While (_xcTab)->(!Eof())
		DbSelectArea("CT2")
		CT2->(DbGoTo((_xcTab)->CT2_NREC))
		_cLinha := _cCntCred := _cCntDeb := _cTpMov := ""

		//Verifica a existência da máscara contábil If !Empty(CT2->CT2_DEBITO)
		If !Empty(CT2->CT2_DEBITO)
			If !CCTBP03MASC(CT2->CT2_DEBITO)
				Exit
			Endif
		EndIf


		If !Empty(CT2->CT2_CREDIT)
			If !CCTBP03MASC(CT2->CT2_CREDIT)
				Exit
			Endif
		EndIf
		DbSelectArea("CT2")
		CT2->(DbGoTo((_xcTab)->CT2_NREC))
		
		//Se não está habilitado envio de partida dobrada, quebra os lançamentos em 2 linhas 
		If !GetMv("MV__CTBP3D") //ZEROU PARA O PROXIMO
			If !Empty(CT2->CT2_DEBITO)
				_cCntCred := ""
				_cCntDeb := CT2->CT2_DEBITO

				//Preenche o arquivo com as informações do layout
				//PAGINA 4

				DbSelectArea(_cSZ3)
				(_cSZ3)->(DbGoTop())
				While (_cSZ3)->(!Eof())
					
					_cLinha += PADR (& ((_cSZ3)->Z3_FORMULA),(_cSZ3)->Z3_TAMANHO)
					//DbSelectArea(_cSZ3)//ambrosini
					(_cSZ3)->(DbSkip())

				EndDo

				If fWrite(_nHdl, _cLinha + _cEOL) < 1 
					_lErro := .t.
					Exit
				Endif
			EndIf

			_cLinha := _cCntCred := _cCntDeb := _cTpMov := ""
			If !Empty(CT2->CT2_CREDIT)
				_cCntCred := CT2->CT2_CREDIT 
				_cCntDeb := ""

				//Preenche o arquivo com as informações do layout 
				DbSelectArea(_cSZ3)
				(_cSZ3)->(DbGoTop())
				While (_cSZ3)->(!Eof ())

					_cLinha += PADR(&((_cSZ3)->Z3_FORMULA),(_cSZ3)->Z3_TAMANHO)
					//DbSelectArea(_cSZ3)
					(_cSZ3)->(DbSkip())
				EndDo


				If fWrite(_nHdl, _cLinha + _cEOL) < 1 
					_lErro := .t.
					Exit
				EndIf
			EndIf
		Else
			//Caso permitido uso de partida dobrada, haverá uma coluna para credito e outra para debito
			DbSelectArea(_cSZ3)
			(_cSZ3)->(DbGoTop())
			While (_cSZ3)->(!Eof ())

				_cLinha += PADR(&((_cSZ3)->Z3_FORMULA),(_cSZ3)->Z3_TAMANHO)
				//DbSelectArea(_cSZ3)
				(_cSZ3)->(DbSkip())
			EndDo

			If fWrite(_nHdl, _cLinha + _cEOL) < 1 
				_lErro := .t.
				Exit
			EndIf
		EndIf

		//DbSelectArea(_xcTab)
		(_xcTab)->(DbSkip())

	EndDo

	If _lErro
		Msglnfo("Erro durante a gravação do arquivo")
	EndIf

	If Select(_cSZ3) > 0 
		DbSelectArea(_cSZ3)
		DbCloseArea()
	Endif

	//PAGINA 5
	DbSelectArea(_xcTab)

Return

//ADICIONADO POR CLAUDIO AMBROSINI PARA COLOCAR FILIAL CORRESPONDENTE
USer Function MudaFilial(cFilCt2)

Local aArea := GetArea()
Local cFilnew := alltrim(cFilCt2)

SX5->(DbSetOrder(1))
If SX5->(DbSeek(xFilial("SX5")+"ZV"))

		WHILE SX5->(!EOF()) .AND. SX5->X5_TABELA == "ZV" 
		
	            If ALLTRIM(SubStr(SX5->X5_DESCRI,1,2)) == cFilnew
					cFilnew := ALLTRIM(SubStr(SX5->X5_DESCRI,4,5))
					EXIT
				EndIf
			SX5->(DbSkip())
		EndDo
Else
	Alert("Não encontrou!")
EndIf

RestArea(aArea)

Return cFilnew


User Function CCTBP03M(_cConta)

	Local _cCntMasc:= _cSep:= _cMasc:= _cCtaSup:= ""

	//Se a conta SUperÍOr for analítica/ ela será gravada no TXT invés da conta do lançamento contábil

	DbSelectArea("CT1")
	DbSetOrder(1)
	DbSeek(xFilial("CT1") + _cConta)

	/*If !Empty(CTl—>CT1_CTASUP)
	_cCtaSup := CT1->CT1_CTASUP
	If Posicione ("CTl", l,xFilial ("CTl") + _cCtaSup,"CTl CLASSE") == "2"
	_cConta := _cCtaSup Endlf Endlf*/

	If !Empty(CT1->CT1__CTASBR)
		_cConta := CT1->CTl__CTASBR 
	EndIf

	_cMasc:= RetMasCtb(StrZero(Len(Alltrim(_cConta)),TamSX3("CTM_CODIGO")[1]), @_cSep ) 
	_cCntMasc := MascaraCTB (_cConta,_cMasc, ,_cSep)

Return _cCntMasc

/* -1	 —					

1 Programa:	1	CCTBP03H | Autor:	| CLAUDIO AMBROSINI - ETHOSX Consultoria | Data: |
setembro 2019 i-	1		
i			
I Descrição:	1	Monta o histórico	do lançamento a ser considerado no
arquivo +			1	
-----+
I Uso: Coopmil
*/

User Function CCTBP03H()
	Local _aArea := GetArea(), _cHist:= _cSQL:= _cCT2:= ""

	_cHist:= CT2->CT2_HIST

	//Busca complemento de histórico 

	_cCT2 := GetNextAlias()

	_cSQL := "Select CT2.CT2_HIST,CT2.CT2_SEQHIS From " + RetSqlName ("CT2") + " CT2 "
	_cSQL += "Where CT2.CT2_FILIAL = '" + CT2->CT2_FILIAL + "' And CT2.CT2_DATA = '" + DTOS(CT2->CT2_DATA) + "' And "
	_cSQL += "CT2.CT2_LOTE = '" + CT2->CT2_LOTE + "' And CT2.CT2_SBLOTE = '" + CT2->CT2_SBLOTE + "' And CT2.CT2_DOC = '" + CT2->CT2_DOC + "' And "
	_cSQL += "CT2.CT2_DC = '4' And CT2.CT2_HIST <> '' And CT2.CT2_SEQLAN = '" + CT2->CT2_SEQLAN + "' And CT2.D_E_L_E_T_ = '' "
	_cSQL += "Order By CT2.CT2_SEQHIS "

	//pagina 6
	If Select(_cCT2) > 0 
		DbSelectArea (_cCT2)
		DbCloseArea()
	EndIf

	DbUseArea( .T.,"TOPCONN",TcGenQry(,,_cSQL),_cCT2,.T.,.T.)
	DbSelectArea(_cCT2)
	DbgoTop()

	While !Eof()
		_cHist += (_cCT2)->CT2_HIST

		DbSelectArea(_cCT2)
		DbSkip()
	EndDo

	_cHist:= SubStr(_cHist,1,(_cSZ3)->Z3_TAMANHO) //Tamanho máximo permitido no layout

	RestArea(_aArea)
Return _cHist

Static Function CCTBP03MASC(_cConta)
	Local _lRet := .t., _cCodMasc := _cMsg := ""
	_cCodMasc := StrZero(Len(Alltrim(_cConta)),TamSX3("CTM_CODIGO")[1])
	DbSelectArea("CTM")
	DbSetOrder(1)

	If !DbSeek(xFilial("CTM") + _cCodMasc)//duvida
		_cMsg := "Nào localizada máscara contábil" + _cCodMasc + " Corrija o cadastro e gere o arquivo novamente"
		_lRet := .f.
	EndIf

	If !Empty(_cMsg)
		MsgStop(_cMsg,"Cadastrar Máscara Contábil")
	EndIf

Return _lRet

/*
+	+. 	+ | Programa: I setembro 2019	CCTBP03S 1 Autor: |claudio ambrosini - ethosx Consultoria 1	1 Data: |
+ — — — — — —		
Descrição: | 'quivo 1	Incrementa a variável utilizada para controlar quantidade de	registros no
+ | Uso: 1		
Coopmil +	+-		1
-------+
*/

User Function CCTBP03S()
	_nReg++
Return _nReg

/*
+------------+------------+--------+------------------------------------------+--------+
-------+
I	Programa:	I	CCTBP03X1	|	Autor:	|	CLAUDIO AMBROSINI - ethosx Consultoria	I	Data:	|
SETEMBRO 2019 |
+------------+------------+--------+------------------------------------------+--------+
—,——
Descrição: | Criação dos
-6-

CMy^«hl,|,'D,.pl,.,'cll.ntt.to.pmlia. WOJEWtt* 3QUC1TACQE3 DE MUDANÇA TODOS OS PROJETOSIE.p.clnc.;l. Iiptl» ,	d. 20, 7
parâmetros	"	"	j----------------------—
+----------------------------------------------------------------------------
| Uso:	I
Coopmil	/
*/

Static Function CCTBP03X1()

	DbSelectArea("SX1") 
	DbSetOrder(1)
	If !DbSeek(_cPerg + "01",.f.)
		While !RecLock("SX1",.t.)
		Enddo
		SX1->X1_GRUPO := _cPerg 
		SX1->X1_ORDEM := "01"
		SX1->X1_PERGUNT := "Data De?"
		SX1->X1_VARIAVL := "mv_ch1"
		SX1->X1_TIPO := "D"
		SX1->X1_TAMANHO := TamSX3("CT2_DATA")[1]
		SX1->X1_GSC := "G"
		SX1->X1_VAR01 := "mv_par01"
		MsUnLock ()
	EndIF



	IF !DbSeek(_cPerg + "02",.f.) 
		While !RecLock("SX1", .t.)
		Enddo
		SX1->X1_GRUPO := _cPerg 
		SX1->X1_ORDEM := "02"
		SX1->X1_PERGUNT:= "Data Ate?"
		SX1->X1_VARIAVL:= "mv_ch2"
		SX1->X1_TIPO := "D"
		SX1->X1_TAMANHO:= TamSX3("CT2_DATA")[1] 
		SX1->X1_GSC := "G"
		SX1->X1_VAR01 = "mv_par02"
		MsUnLock()
	Endif


	If  !DbSeek (_cPerg + "03",.f.)
	While !RecLock ("SX1",.t.)
	Enddo 
	SX1->X1_GRUPO := _cPerg
	SX1->X1_ORDEM := "03"
	SX1->X1_PERGUNT := "Lote (Modulo)?"
	SX1->X1_VARIAVL := "mv_ch3"
	SX1->X1_TIPO := "C"
	SX1->X1_TAMANHO := TamSX3("X5_CHAVE")[1] 
	SX1->X1_GSC := "G"
	SX1->X1_VAR01 := "mv_par03"
	SX1->X1_F3 := "09"
	MsUnLock()
	EndIf

	If 	!DbSeek(_cPerg + "04",.F.)
	While !RecLock("SX1",.t.)
	Enddo
	SX1->X1_GRUPO 	:= _cPerg 
	SX1->X1_ORDEM 	:= "04"
	SX1->X1_PERGUNT := "Cod. Hist. SISBR?"
	SX1->X1_VARIAVL := "mv_ch4"
	SX1->X1_TIPO	:= "C"
	SX1->X1_TAMANHO := 5
	SX1->X1_GSC		:= "G"
	SX1->X1_VAR01	:= "mv_par04"
	MsUnLock ()
	EndIF


	If !DbSeek(_cPerg + "05",.f.) 
	While !RecLock("SX1",.t.) 
	Enddo
	SX1->X1_GRUPO	:= _cPerg
	SX1->X1_ORDEM	:= "05"
	SX1->X1_PERGUNT := "Considerar Lançamentos?"
	SX1->X1_VARIAVL := "mv_ch5"
	SX1->X1_TIPO	:= "N"
	SX1->X1_TAMANHO	:= 1
	SX1->X1_GSC		:= "C"
	SX1->X1_VAR01	:= "mv_par05"
	SX1->X1_DEF01	:= "Filial Logada"
	SX1->X1_DEF02	:= "Empresa Logada"
	SX1->X1_PRESEL	:= 1
	MsUnLock()	
	EndIf	


	If !DbSeek(_cPerg + "06",.f.)	
	While !RecLock("SX1",.t.)	
	Enddo	
	SX1->X1_GRUPO	:= _cPerg
	SX1->X1_ORDEM	:= "06"
	SX1->X1_PERGUNT	:= "Cod. Cooperativa?"
	SX1->X1_VARIAVL	:= "mv_ch6"
	SX1->X1_TIPO	:= "C"
	SX1->X1_TAMANHO	:= 4
	SX1->X1_GSC		:= "G"
	SX1->X1_VAR01	:= "mv_par06"
	MsUnLock()	
	EndIF
	
/*
	If 	!DbSeek(_cPerg + "07",.f.) 
	While !RecLock("SX1", .t.) 
	Enddo
	SX1->X1_GRUPO	:= _cPerg
	SX1->X1_ORDEM	:= "07"
	SX1->X1_PERGUNT	:= "Côd. PAC?"
	SX1->X1_VARIAVL	:= "mv_ch7"
	SX1->X1_TIPO	:= "C"
	SX1->X1_TAMANHO	:= 2
	SX1->X1_GSC		:= "G"
	SX1->X1_VAR01	:= "MV_PAR07"
	SX1->X1_CNT01	:= cFilant
	MsUnLock()	
	EndIf
*/	
Return