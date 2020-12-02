#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
 
Static cTitulo := "Cadastro de Processamento de Filiais"
 
User Function AtuFilAtf()
    
Local aArea   := GetArea()
Local oBrowse

Private aSN1    := {}

oBrowse := FWMBrowse():New()
 
oBrowse:SetAlias("ZZ0")
 
oBrowse:SetDescription(cTitulo)
     
//Legendas
oBrowse:AddLegend( "ZZ0->ZZ0_FLAG != '1'", "RED"  ,  "Não Processado" )
oBrowse:AddLegend( "ZZ0->ZZ0_FLAG == '1'", "GREEN",  "Processado" )
 
//Ativa a Browse
oBrowse:Activate()
 
RestArea(aArea)

Return Nil
 
/*---------------------------------------------------------------------*
 | Func:  MenuDef                                                      |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação do menu MVC                                          |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function MenuDef()

Local aRot := {}
 
//Adicionando opções
ADD OPTION aRot TITLE 'Incluir'    ACTION 'VIEWDEF.AtuFilAtf' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION X
ADD OPTION aRot TITLE 'Processar'  ACTION 'U_GerFils()' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Alterar'    ACTION 'VIEWDEF.AtuFilAtf' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 3
ADD OPTION aRot TITLE 'Visualizar' ACTION 'VIEWDEF.AtuFilAtf' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 1
ADD OPTION aRot TITLE 'Excluir'    ACTION 'VIEWDEF.AtuFilAtf' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
 
Return aRot
 
/*---------------------------------------------------------------------*
 | Func:  ModelDef                                                     |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação do modelo de dados MVC                               |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ModelDef()

Local oModel     := Nil
Local oStPai     := FWFormStruct(1, 'ZZ0')
Local oStFilho   := FWFormStruct(1, 'ZZ1')
Local aRel    := {}
 
//Criando o modelo e os relacionamentos
oModel := MPFormModel():New('PROCFIL')
oModel:AddFields('ZZ0MASTER',/*cOwner*/,oStPai)
oModel:AddGrid('ZZ1DETAIL','ZZ0MASTER',oStFilho,/*bLinePre*/, /*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
 
//Fazendo o relacionamento entre o Pai e Filho
//aAdd(aSB1Rel, {'B1_FILIAL',    'BM_FILIAL'} )
oModel:SetRelation("ZZ1DETAIL", ;       
 					{{"ZZ1_FILIAL",'xFilial("ZZ1")'},;        
					{"ZZ1_LOTE","ZZ0_LOTE"  }}, ;       
					ZZ1->(IndexKey(1)))  
 
oModel:GetModel('ZZ1DETAIL'):SetUniqueLine({"ZZ1_FILIAL","ZZ1_LOTE","ZZ1_CCUSTO"})    //Não repetir informações ou combinações {"CAMPO1","CAMPO2","CAMPOX"}
oModel:SetPrimaryKey({})
 
//Setando as descrições
oModel:SetDescription("Processamento Filiais de Ativo")
oModel:GetModel('ZZ0MASTER'):SetDescription('Cód. de Processamento')
oModel:GetModel('ZZ1DETAIL'):SetDescription('Centro de Custo x Filial')

Return oModel
 
/*---------------------------------------------------------------------*
 | Func:  ViewDef                                                      |
 | Autor: André Brito                                                |
 | Data:  17/08/2015                                                   |
 | Desc:  Criação da visão MVC                                         |
 | Obs.:  /                                                            |
 *---------------------------------------------------------------------*/
 
Static Function ViewDef()
    Local oView      := Nil
    Local oModel     := FWLoadModel('AtuFilAtf')
    Local oStPai     := FWFormStruct(2, 'ZZ0')
    Local oStFilho   := FWFormStruct(2, 'ZZ1')
     
    //Criando a View
    oView := FWFormView():New()
    oView:SetModel(oModel)
     
    //Adicionando os campos do cabeçalho e o grid dos filhos
    oView:AddField('VIEW_ZZ0',oStPai,'ZZ0MASTER')
    oView:AddGrid('VIEW_ZZ1',oStFilho,'ZZ1DETAIL')
     
    //Setando o dimensionamento de tamanho
    oView:CreateHorizontalBox('CABEC',30)
    oView:CreateHorizontalBox('GRID',70)
     
    //Amarrando a view com as box
    oView:SetOwnerView('VIEW_ZZ0','CABEC')
    oView:SetOwnerView('VIEW_ZZ1','GRID')
     
    //Habilitando título
    oView:EnableTitleView('VIEW_ZZ0','Cod. Processamento')
    oView:EnableTitleView('VIEW_ZZ1','Centro de Custo x Filial')
    
Return oView

User Function GerFils()

Local nRadio := 0
Local aItems := {}
Local oRadio
Local nSel := 0
Local aBotoes    := {}
Local bConfir := {||GerAtf(oRadio, nSel) }
Local bClosed := {|| oDlg:DeActivate()}

aItems := {'SN1 / SN3 - Cad. Ativo/Saldos e Valores','SN4 - Movimentações Ativo'}

Aadd(aBotoes, {"", "Processar", bConfir, , , .T., .F.}) // 'Confirmar'
Aadd(aBotoes, {"", "Fechar", bClosed, , , .T., .F.}) // 'Confirmar'

oDlg := FWDialogModal():New()
oDlg:SetBackground(.T.) // .T. -> escurece o fundo da janela
oDlg:SetTitle("Selecione a tabela do ativo desejada")//Tipo de Atividade
oDlg:SetEscClose(.T.)//permite fechar a tela com o ESC
oDlg:SetSize(120,160) //cria a tela maximizada (chamar sempre antes do CreateDialog)
oDlg:EnableFormBar(.T.)

oDlg:CreateDialog() //cria a janela (cria os paineis)
oPanel := oDlg:getPanelMain()
oDlg:createFormBar()//cria barra de botoes

oRadio := TRadMenu():New (05,03,aItems,, oPanel ,,,,,,,,110,12,,,,.T.)
oRadio:bSetGet := {|u|Iif (PCount()==0,nSel,nSel:=u)}

oDlg:addButtons(aBotoes)
oDlg:activate()
lRetorno := IIF(oDlg:getButtonSelected() > 0, .T., .F.)//pegando a resposta di usuario na tela

Return

Static Function GerAtf(oRadio,nSel)

If nSel == 1
	MsgRun("Processando filiais de ativos ...","Aguarde...",{|| GerFilAtf(oRadio, nSel) })
ElseIf nSel == 2
	MsgRun("Processando filiais das movimentações de ativos ...","Aguarde...",{|| GerFilAtf(oRadio, nSel) })
EndIf
Return

Static Function GerFilAtf(oRadio,nSel)

Local lRetAtf    := .F.
Local lRetSal    := .F.

If ZZ0_FLAG == "1"
	If MSGYESNO( "Lote já processado, deseja continuar?", "Lote Processado" )
		If nSel == 1
			lRetAtf := AtuFiliais(ZZ0_LOTE, nSel)
		EndIf
		
		If nSel == 2
			lRetSal := AtuFiliais(ZZ0_LOTE, nSel)
		EndIf
	Else
		Return
	EndIf
EndIf

If !lRetAtf .And.!lRetSal
	lRetAtf := AtuFiliais(ZZ0_LOTE, nSel)
EndIf

If lRetSal
	RecLock("ZZ0",.F.)
		ZZ0->ZZ0_FLAG = "1"
	MsUnLock
	
	MsgInfo( "Filiais gravadas na tabela SN4 com sucesso!", "Concluido" )
	
	aSN1 := {}
EndIf

If lRetAtf
	RecLock("ZZ0",.F.)
		ZZ0->ZZ0_FLAG = "1"
	MsUnLock
	
	MsgInfo( "Filiais gravadas nas tabelas SN1 / SN2 / SN3 com sucesso!", "Concluido" )
	
	If MSGYESNO( "Deseja visualizar os ativos que tiveram suas filiais alteradas?", "Visualizar Log" )
		MsgRun("Gerando Planilha Excel...","Aguarde...",{|| RelFil(aSN1) })
	EndIf
	
	aSN1 := {}
	
EndIf

Return 

//--------------------------------------------------------------------------------

Static Function AtuFiliais(cLote, nSel)

Local aArea   := GetArea()
Local lRet    := .F.
Local lRetorno:= .F.
Local cQuery  := ""
Local cAliAux := GetNextAlias()
Local nTotal  := 0
Local nAtual  := 0
Local x       := 0

cQuery := "SELECT * FROM "
//If nSel == 1
	cQuery += RetSqlName("ZZ1") + "WHERE ZZ1_LOTE = '" + Alltrim(cLote) + "' AND D_E_L_E_T_ = ' ' ORDER BY ZZ1_CCUSTO "
//Else
	//cQuery += RetSqlName("ZZ1") + "WHERE ZZ1_LOTE = '" + Alltrim(cLote) + "' AND ZZ1_DATA = ' ' AND D_E_L_E_T_ = ' ' ORDER BY ZZ1_CCUSTO "
//EndIf
	
cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

(cAliAux)->(dbGotop())

Do While !(cAliAux)->(Eof())
	
	If nSel == 1
		lRet := AtuSal((cAliAux)->ZZ1_CCUSTO,(cAliAux)->ZZ1_FILDES)
		lRetorno := .T.
	EndIf
	
	If nSel == 2
		lRet := AtuSN4((cAliAux)->ZZ1_CCUSTO,(cAliAux)->ZZ1_FILDES)
		AtuZZ1((cAliAux)->ZZ1_LOTE, (cAliAux)->ZZ1_CCUSTO)
		lRetorno := .T.
	EndIf

	AtuSN5((cAliAux)->ZZ1_CCUSTO,(cAliAux)->ZZ1_FILDES,(cAliAux)->ZZ1_CTA)
	AtuSN6((cAliAux)->ZZ1_CCUSTO,(cAliAux)->ZZ1_FILDES,(cAliAux)->ZZ1_CTA)
	
	(cAliAux)->(dbskip())		

EndDo

DbCloseArea(cAliAux)


RestArea(aArea)

Return lRetorno

//--------------------------------------------------------------------------------

Static Function AtuSal(cCusto,cFilDes)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()
Local aCods     := {}

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SN3") + " SN3 "
cQuery += "WHERE N3_CUSTBEM = '" + cCusto + "' AND N3_FILIAL = ' ' AND D_E_L_E_T_ = ' ' "
	
cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTop,.T.,.T.)

(cAliasTop)->(dbGotop())

Begin Transaction 

Do While !(cAliasTop)->(Eof())
	
	SN3->(DbGoto((cAliasTop)->R_E_C_N_O_))
	
	aAdd(aCods,{(cAliasTop)->N3_FILIAL,(cAliasTop)->N3_CBASE })
	
	lRet  := .T.
	
	RecLock("SN3",.F.)
	
	SN3->N3_FILIAL := cFilDes
		
	SN3->(MsUnLock())	

	(cAliasTop)->(dbskip())		

EndDo

AtuSn1(aCods, cFilDes)
AtuSn2(aCods, cFilDes)

End Transaction

DbCloseArea(cAliasTop)

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuSN4(cCusto,cFilDes)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTp  := GetNextAlias()

//DbSelectArea("SN4")
//DbSetOrder(1)

cQuery := "SELECT R_E_C_N_O_ FROM "
cQuery += RetSqlName("SN4") + " SN4 "
cQuery += "WHERE N4_CCUSTO = '" + cCusto + "' AND N4_FILIAL = ' ' AND D_E_L_E_T_ = ' ' "
	
cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTp,.T.,.T.)

(cAliasTp)->(dbGotop())

Begin Transaction

Do While !(cAliasTp)->(Eof())
	
	SN4->(DbGoto((cAliasTp)->R_E_C_N_O_))
	
	lRet := .T.
	
	RecLock("SN4",.F.)
	
	SN4->N4_FILIAL := cFilDes
		
	MsUnLock()	

	(cAliasTp)->(DbSkip())			

EndDo

End Transaction

DbCloseArea(cAliasTp)

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuSN1(aCodBem, cFilDes)

Local aArea     := GetArea()
Local lRet      := .F.
Local x         := 0

DbSelectArea("SN1")
DbSetOrder(1)

For x := 1 To Len(aCodBem)
	If SN1->( DbSeek(aCodBem[x][1] + aCodBem[x][2]))
		lRet := .T.
		
		aAdd(aSN1,Recno())

		RecLock("SN1", .F.)
			SN1->N1_FILIAL := cFilDes
		SN1->(MsUnLock())
	EndIf
Next

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuSN2(aCodBem, cFilDes)

Local aArea     := GetArea()
Local lRet      := .F.
Local x         := 0

DbSelectArea("SN2")
DbSetOrder(1)

For x := 1 To Len(aCodBem)
	If SN2->( DbSeek(aCodBem[x][1] + aCodBem[x][2]))
		lRet := .T.
		RecLock("SN2", .F.)
			SN2->N2_FILIAL := cFilDes
		SN2->(MsUnLock())
	EndIf
Next

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuZZ1(cLote, cCusto)

Local aArea     := GetArea()
Local lRet      := .F.
Local x         := 0
Local cQuery    := ""
Local cAliZZ1   := GetNextAlias()

DbSelectArea("ZZ1")
DbSetOrder(1)


cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("ZZ1") + " ZZ1 "
cQuery += "WHERE ZZ1_CCUSTO = '" + cCusto + "' AND ZZ1_LOTE = '" + cLote + " ' AND D_E_L_E_T_ = ' ' "
	
cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliZZ1,.T.,.T.)

(cAliZZ1)->(dbGotop())

Do While !(cAliZZ1)->(Eof())
	
	ZZ1->(DbGoto((cAliZZ1)->R_E_C_N_O_))
	
	RecLock("ZZ1",.F.)
	
	ZZ1->ZZ1_DATA := dDataBase
		
	ZZ1->(MsUnLock())	

	(cAliZZ1)->(dbskip())		

EndDo

DbCloseArea(cAliZZ1)

RestArea(aArea)

Return 

//--------------------------------------------------------------------------------

Static Function RelFil(aSN1)

Local aArea     := GetArea()
Local lRet      := .F.
Local oExcelApp := Nil
Local cPath     := "C:\Temp"
Local cArquivo  := "AtivosFiliais.XLS"
Local x         := 0
Local oExcel
Local oExcelApp
Local _oPlan
Local aColunas   := {}
Local aLocais    := {}
 
oBrush1  := TBrush():New(, RGB(193,205,205))

If !ApOleClient('MsExcel')

    MsgAlert("Falha ao abrir Excel!")
    //Return

EndIf

oExcel  := FWMSExcel():New()
cAba    := "Bens que tiveram filiais inclusas"
cTabela := "Ativo Fixo - Coopmil"

// Criação de nova aba 
oExcel:AddworkSheet(cAba)

// Criação de tabela
oExcel:AddTable (cAba,cTabela)

// Criação de colunas 

oExcel:AddColumn(cAba,cTabela,"ATIVO FILIAL"     ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ATIVO FIXO"       ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"DESCRIÇÃO"        ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"CHAPA"            ,2,1,.F.)

DbSelectArea("SN1")
DbSetOrder(1)

For x := 1 to Len(aSN1)
	DbGoto(aSN1[X])
	oExcel:AddRow(cAba,cTabela, { FWFilialName(,SN1->N1_FILIAL,1) ,;
                                  SN1->N1_CBASE   ,; 
                                  SN1->N1_DESCRIC ,; 
                                  SN1->N1_CHAPA })
	
Next

If !Empty(oExcel:aWorkSheet)

    oExcel:Activate()
    oExcel:GetXMLFile(cArquivo)
 
    CpyS2T("\SYSTEM\"+cArquivo, cPath)

    oExcelApp := MsExcel():New()
    oExcelApp:WorkBooks:Open(cPath + "\" + cArquivo) // Abre a planilha
	oExcelApp:SetVisible(.T.)
	
EndIf

RestArea(aArea)

Return
//--------------------------------------------------------------------------------

Static Function AtuSN5(cCusto,cFilDes,cConta)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()

Default cFilDes := FWCodFil()

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SN5") + " SN5 "
cQuery += "WHERE N5_CONTA = '" + cConta + "' " 
cQuery += "AND D_E_L_E_T_ = ' '"

cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTop,.T.,.T.)

(cAliasTop)->(dbGotop())

Do While !(cAliasTop)->(Eof())
	
	SN5->(DbGoto((cAliasTop)->R_E_C_N_O_))
	
	lRet  := .T.
	
	RecLock("SN5",.F.)
	
	SN5->N5_FILIAL:= cFilDes
	
	SN5->(MsUnLock())	

	(cAliasTop)->(dbskip())		

EndDo

DbCloseArea(cAliasTop)

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuSN6(cCusto,cFilDes,cConta)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()

Default cFilDes := FWCodFil()

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SN6") + " SN6 "
cQuery += "WHERE N6_CCUSTO = '" + cCusto + "' " 
cQuery += "AND D_E_L_E_T_ = ' '"

cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTop,.T.,.T.)

(cAliasTop)->(dbGotop())

Do While !(cAliasTop)->(Eof())
	
	SN6->(DbGoto((cAliasTop)->R_E_C_N_O_))
	
	lRet  := .T.
	
	RecLock("SN6",.F.)
	
	SN6->N6_FILIAL:= cFilDes
	
	SN6->(MsUnLock())	

	(cAliasTop)->(dbskip())		

EndDo

DbCloseArea(cAliasTop)

RestArea(aArea)

Return lRet