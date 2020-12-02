#INCLUDE "FileIO.CH"
#INCLUDE "PROTHEUS.CH"    

//--------------------------------------------------
/*/{Protheus.doc} CCTBP02
Importação de Arquivo de Plano de Contas Coopmil

@author André Brito
@since 07/08/2020
@version P12.1.17
 
@return 
/*/
//--------------------------------------------------

User Function CCTBP02()

	Local aRet			:= {}
	Local aAreaCT1		:= CT1->(GetArea())
	Local aAreaCVD		:= CVD->(GetArea())
	Local aCfg			:= {}
	Local cCampos		:= ""
	Local lContinua		:= .T.
	Local oModCT1Imp	:= Nil

	Private oProcess   

	SaveInter()

	oModCT1Imp	:= FWLoadModel("CTBA020")

	aCfg := { { "CT1", cCampos, {|| FWMVCRotAuto(oModCT1Imp, "CT1", 3, { {"CT1MASTER",xAutoCab} }, , .T.) } }, {"CT1",,} }

	If ParamBox({	{6,"Selecione Arquivo",PadR("",150),"",,"", 90 ,.T.,"Importação de Plano de Contas","",GETF_LOCALHARD+GETF_LOCALFLOPPY+GETF_NETWORKDRIVE}},;
					"Importar Estrutura de Plano de Contas",@aRet) 

		oProcess:= MsNewProcess():New( {|lEnd| CoopImpCSV( lEnd, oProcess, aRet[1], aCfg )} )
		oProcess:Activate()

	EndIf

	oModCT1Imp:Destroy()
	oModCT1Imp := Nil

	RestInter()

RestArea(aAreaCVD)
RestArea(aAreaCT1)

Return Nil

RestInter()

Return .T.


//--------------------------------------------------
/*/{Protheus.doc} CoopImpCSV
Importa registros do arquivo para tabela CT1

@author André Brito
@since 02/12/2019
@version P12.1.17
 
@return 
/*/
//--------------------------------------------------

Static Function CoopImpCSV(lEnd, oProcess, cArq, aCfg , lProc)

Local cLinha      := ""
Local lPrim       := .T.
Local aCampos     := {}
Local aDados      := {}
Local aProds      := {}
Local i           := 0
Local x           := 0
Local lMsErroAuto := .F.
Local nAtual      := 0
Local nTotal      := 0
Local nTot2       := 0
Local nNumProd    := 0
Local aCab        := {}
Local oModel      := Nil
Local aArea       := GetArea()
Local aAreaSb1    := GetArea()
Local lRet        := .T.
Local cCofins     := ""
Local cPis        := ""
Local aProdutos   := {}
Local nInclu      := 0
Local nOpcAuto    := 0
Local cLog        := ""
Local cConta      := ""
Local nX
Local oCT1
Local aLog

Private aErro := {}
 
If !File(cArq)
	MsgStop("O arquivo "  + cArq + " não foi encontrado. A importação será abortada!","ATENCAO")
	Return
EndIf
 
FT_FUSE(cArq)
ProcRegua(FT_FLASTREC())
FT_FGOTOP()

nTot2 := FT_FLASTREC()
oProcess:SetRegua1(nTot2)

While !FT_FEOF()
	

	oProcess:IncRegua1("Totais de contas lidas: " + cValToChar(nNumProd))

	nNumProd := nNumProd + 1

	cLinha := FT_FREADLN()
 
	If lPrim
		aCampos := Separa(cLinha,";",.T.)
		lPrim := .F.
	Else
		/*If cLinha $ ";;;;;;;;;;;;;;;;;"
			Exit
		EndIf*/
		AADD(aDados,Separa(cLinha,";",.T.))
	EndIf
 
	FT_FSKIP()
EndDo

nTotal := Len(aDados)

Count To nTotal

oProcess:SetRegua2(nTotal)

Begin Transaction

For i := 1 to Len(aDados)
	
	DbSelectArea("CT1")
	DbSetOrder(1)
	
	If !Empty (aDados [i][6]) .And. Alltrim(aDados [i][6])!= "CONTA CONTÁBIL"
		
		oProcess:IncRegua2("Incluindo conta contabil: " + aDados[i][6])
		
		cConta := StrTran( Alltrim(aDados [i][6]), ".", "" )
		cConta := StrTran( cConta, "-", "" )
		
		If Alltrim(aDados [i][12]) == "Saldo Devedor"
			cNormal := "2"
		Else
			cNormal := "1"
		EndIf
		
		If DbSeek(xFilial("CT1") + cConta)
			RecLock("CT1", .F.)
				CT1->CT1_FILIAL  := xFilial("CT1")
				CT1->CT1_CONTA   := cConta
				CT1->CT1_DESC01  := Alltrim(aDados [i][12])
				CT1->CT1_RES     := Alltrim(aDados [i][11])
				CT1->CT1_DTBLIN  := CTOD(aDados [i][23])
				CT1->CT1_NORMAL  := cNormal
			MsUnLock
		Else
			RecLock("CT1", .T.)
				CT1->CT1_FILIAL  := xFilial("CT1")
				CT1->CT1_CONTA   := cConta
				CT1->CT1_DESC01  := Alltrim(aDados [i][12])
				CT1->CT1_RES     := Alltrim(aDados [i][11])
				CT1->CT1_DTBLIN  := CTOD(aDados [i][23])
				CT1->CT1_CLASSE  := "2"
				CT1->CT1_NORMAL  := cNormal
			MsUnLock
		EndIf

	EndIf

Next

End Transaction

//U_CoopMsg(Len(aDados), nInclu, aProdutos)
 
FT_FUSE()

If lRet
	ApMsgInfo("Importação das contas contabeis concluída com sucesso!","SUCESSO")
Else
	ApMsgInfo("Aconteceram erros na sua importação, verifique!","Conferência")
EndIf

RestArea(aArea)

Return( lRet )

User Function CoopMsg(nDados, nInclu, aProdutos)

	Local lRetMens             := .F.
	Local oDlgMens
	Local oBtnOk, cTxtConf     := ""
	Local oBtnCnc, cTxtCancel  := ""
	Local oBtnSlv
	Local oFntTxt              := TFont():New("Verdana",,-011,,.F.,,,,,.F.,.F.)
	Local oMsg
	Local nIni                 := 1
	Local nFim                 := 50
	Local cMsg                 := ""
	Local cTitulo              := "Contas importadas"
	Local cQuebra              := CRLF + CRLF
	Local nTipo                := 1 // 1=Ok; 2= Confirmar e Cancelar
	Local lEdit                := .F.
    Local nX                   := 0

    cMsg  := "Total de contas processadas: " + Alltrim(Str(nDados)) + CRLF
    cMsg  += "Total de contas inclusas: " + Alltrim(Str(nInclu))
    
	cTexto := "Função   - " + FunName()       + CRLF
	cTexto += "Usuário  - " + cUserName       + CRLF
	cTexto += "Data     - " + dToC(dDataBase) + CRLF
	cTexto += "Hora     - " + Time()          + CRLF
	cTexto += "Mensagem - " + cTitulo + cQuebra  + cMsg + " " + cQuebra
	cTexto += CRLF

	If nInclu != nDados
		cTexto += "Registros não inclusos:" + CRLF + CRLF
	EndIf

	For nX := 1 To Len(aProdutos)
		cTexto += "Código Produto: " + Alltrim(aProdutos[nX][1]) + CRLF
		cTexto += "Descrição Pedido: " + Alltrim(aProdutos[nX][2]) + CRLF
	Next
    
    //Definindo os textos dos botões
	If(nTipo == 1)
		cTxtConf:='Ok'
	Else
		cTxtConf:='Confirmar'
		cTxtCancel:='Cancelar'
	EndIf
 
    //Criando a janela centralizada com os botões
	DEFINE MSDIALOG oDlgMens TITLE cTitulo FROM 000, 000  TO 300, 400 COLORS 0, 16777215 PIXEL
        //Get com o Log
	@ 002, 004 GET oMsg VAR cTexto OF oDlgMens MULTILINE SIZE 191, 121 FONT oFntTxt COLORS 0, 16777215 HSCROLL PIXEL
	If !lEdit
		oMsg:lReadOnly := .T.
	EndIf
         
        //Se for Tipo 1, cria somente o botão OK
	If (nTipo==1)
		@ 127, 144 BUTTON oBtnOk  PROMPT cTxtConf   SIZE 051, 019 ACTION (lRetMens:=.T., oDlgMens:End()) OF oDlgMens PIXEL
         
        //Senão, cria os botões OK e Cancelar
	ElseIf(nTipo==2)
		@ 127, 144 BUTTON oBtnOk  PROMPT cTxtConf   SIZE 051, 009 ACTION (lRetMens:=.T., oDlgMens:End()) OF oDlgMens PIXEL
		@ 137, 144 BUTTON oBtnCnc PROMPT cTxtCancel SIZE 051, 009 ACTION (lRetMens:=.F., oDlgMens:End()) OF oDlgMens PIXEL
	EndIf
         
        //Botão de Salvar em Txt
	@ 127, 004 BUTTON oBtnSlv PROMPT "&Salvar em .txt" SIZE 051, 019 ACTION (SalvaArq(cMsg, cTitulo, Alltrim(Str(nDados)),cTexto)) OF oDlgMens PIXEL
	ACTIVATE MSDIALOG oDlgMens CENTERED
 
Return lRetMens

//--------------------------------------------------
 
Static Function SalvaArq(cMsg, cTitulo, cQtdDados, cTxt)

	Local cFileNom :='\x_arq_'+dToS(Date())+StrTran(Time(),":")+".txt"
	Local cQuebra  := CRLF + "+=======================================================================+" + CRLF
	Local lOk      := .T.
	Local cTexto   := ""
     
    //Pegando o caminho do arquivo
	cFileNom := cGetFile( "Arquivo TXT *.txt | *.txt", "Arquivo .txt...",,'',.T., GETF_LOCALHARD)
 
    //Se o nome não estiver em branco    
	If !Empty(cFileNom)
        //Teste de existência do diretório
		If ! ExistDir(SubStr(cFileNom,1,RAt('\',cFileNom)))
			Alert("Diretório não existe:" + CRLF + SubStr(cFileNom, 1, RAt('\',cFileNom)) + "!")
			Return
		EndIf

		cTexto := cTxt
         
        //Testando se o arquivo já existe
		If File(cFileNom)
			lOk := MsgYesNo("Arquivo já existe, deseja substituir?", "Atenção")
		EndIf
         
		If lOk
			MemoWrite(cFileNom, cTexto)
			MsgInfo("Arquivo Gerado com Sucesso:"+CRLF+cFileNom,"Atenção")
		EndIf
	EndIf
Return
