#Include 'Protheus.ch'
#Include 'FWMVCDef.ch'
 
Static cTitulo := "Cadastro de Processamento de Contas Contábeis"
 
User Function AtuCtaAtf()
    
Local aArea   := GetArea()
Local oBrowse

Private aSN3  := {}
 
oBrowse := FWMBrowse():New()
 
oBrowse:SetAlias("ZZ2")
 
oBrowse:SetDescription(cTitulo)
     
//Legendas
oBrowse:AddLegend( "ZZ2->ZZ2_FLAG != '1'", "RED"  ,  "Não Processado" )
oBrowse:AddLegend( "ZZ2->ZZ2_FLAG == '1'", "GREEN",  "Processado" )
 
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
ADD OPTION aRot TITLE 'Incluir'    ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION X
ADD OPTION aRot TITLE 'Visualizar' ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_VIEW   ACCESS 0 //OPERATION 1
ADD OPTION aRot TITLE 'Alterar'    ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_UPDATE ACCESS 0 //OPERATION 3
ADD OPTION aRot TITLE 'Processar'  ACTION 'U_GerCtas()' OPERATION MODEL_OPERATION_INSERT ACCESS 0 //OPERATION 4
ADD OPTION aRot TITLE 'Excluir'    ACTION 'VIEWDEF.AtuCtaAtf' OPERATION MODEL_OPERATION_DELETE ACCESS 0 //OPERATION 5
 
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
Local oStPai     := FWFormStruct(1, 'ZZ2')
Local oStFilho   := FWFormStruct(1, 'ZZ3')
Local aRel    := {}
 
//Criando o modelo e os relacionamentos
oModel := MPFormModel():New('PROCCTA')
oModel:AddFields('ZZ2MASTER',/*cOwner*/,oStPai)
oModel:AddGrid('ZZ3DETAIL','ZZ2MASTER',oStFilho,/*bLinePre*/, /*bLinePost*/,/*bPre - Grid Inteiro*/,/*bPos - Grid Inteiro*/,/*bLoad - Carga do modelo manualmente*/)
 
//Fazendo o relacionamento entre o Pai e Filho
//aAdd(aSB1Rel, {'B1_FILIAL',    'BM_FILIAL'} )
oModel:SetRelation("ZZ3DETAIL", ;       
 					{{"ZZ3_FILIAL",'xFilial("ZZ3")'},;        
					{"ZZ3_LOTE","ZZ2_LOTE"  }}, ;       
					ZZ3->(IndexKey(1)))  
 
oModel:GetModel('ZZ3DETAIL'):SetUniqueLine({"ZZ3_FILIAL","ZZ3_LOTE","ZZ3_GRUPO"})    //Não repetir informações ou combinações {"CAMPO1","CAMPO2","CAMPOX"}
oModel:SetPrimaryKey({})
 
//Setando as descrições
oModel:SetDescription("Processamento Filiais de Ativo")
oModel:GetModel('ZZ2MASTER'):SetDescription('Cód. de Processamento')
oModel:GetModel('ZZ3DETAIL'):SetDescription('Ativos x Grupos')

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
    Local oModel     := FWLoadModel('AtuCtaAtf')
    Local oStPai     := FWFormStruct(2, 'ZZ2')
    Local oStFilho   := FWFormStruct(2, 'ZZ3')
     
    //Criando a View
    oView := FWFormView():New()
    oView:SetModel(oModel)
     
    //Adicionando os campos do cabeçalho e o grid dos filhos
    oView:AddField('VIEW_ZZ2',oStPai,'ZZ2MASTER')
    oView:AddGrid('VIEW_ZZ3',oStFilho,'ZZ3DETAIL')
     
    //Setando o dimensionamento de tamanho
    oView:CreateHorizontalBox('CABEC',30)
    oView:CreateHorizontalBox('GRID',70)
     
    //Amarrando a view com as box
    oView:SetOwnerView('VIEW_ZZ2','CABEC')
    oView:SetOwnerView('VIEW_ZZ3','GRID')
     
    //Habilitando título
    oView:EnableTitleView('VIEW_ZZ2','Cód. de Processamento')
    oView:EnableTitleView('VIEW_ZZ3','Ativos x Grupos')
    
Return oView

User Function GerCtas()

RptStatus({|| GerCtaAtf()}, "Aguarde...", "Gravando contas nos ativos...")

Return

Static Function GerCtaAtf()

Local lRet    := .F.


If ZZ2_FLAG == "1"
	If MSGYESNO( "Lote já processado, deseja continuar?", "Lote Processado" )
		lRet := AtuContas(ZZ2_LOTE)
	Else
		Return
	EndIf
EndIf

If !lRet
	lRet := AtuContas(ZZ2_LOTE)
EndIf

If lRet
	RecLock("ZZ2",.F.)
		ZZ2->ZZ2_FLAG = "1"
	MsUnLock
	
	MsgInfo( "Entidades contábeis gravadas nos ativos com sucesso!", "Concluido" )
EndIf 

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuContas(cLote)

Local aArea   := GetArea()
Local lRet    := .F.
Local cQuery  := ""
Local cAliAux := GetNextAlias()
Local lAtuSn3 := .F.
Local lAtuSn5 := .F.
Local lAtuSn6 := .F.
Local nTotal  := 0
Local nAtual  := 0
Local x       := 0

cQuery := "SELECT * FROM "
cQuery += RetSqlName("ZZ3") + "WHERE ZZ3_LOTE = '" + Alltrim(cLote) + "' AND D_E_L_E_T_ = ' ' "
	
cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)


Count To nTotal
SetRegua(nTotal)
    
(cAliAux)->(dbGotop())


Do While !(cAliAux)->(Eof())
	
	nAtual++
    IncRegua()
    
	lAtuSn3 := AtuCtaSN3((cAliAux)->ZZ3_ATFINI,(cAliAux)->ZZ3_ATFFIM,(cAliAux)->ZZ3_GRUPO)

	(cAliAux)->(dbskip())		

EndDo

DbCloseArea(cAliAux)

If lAtuSn3 
	lRet := .T.
	If MSGYESNO( "Deseja visualizar os ativos que tiveram suas contas alteradas?", "Visualizar Log" )
		MsgRun("Gerando Planilha Excel","Aguarde...",{|| RelCta(aSN3) })
	EndIf
EndIf

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function AtuCtaSN3(cAtfIni,cAtfFim,cGrupo)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()
Local cContaBem := BuscaConta(cGrupo)
Local cContaDes := BuscaDesp(cGrupo)
Local cContAcum := BuscaAcum(cGrupo)

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SN3") + " SN3 "
cQuery += "WHERE N3_CBASE BETWEEN '" + cAtfIni    + "' AND '" + cAtfFim  + "' " 
cQuery += "AND D_E_L_E_T_ = ' '"

cQuery := ChangeQuery(cQuery) 

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTop,.T.,.T.)

(cAliasTop)->(dbGotop())

Do While !(cAliasTop)->(Eof())
	
	SN3->(DbGoto((cAliasTop)->R_E_C_N_O_))
	
	lRet  := .T.
	
	aAdd(aSN3,(cAliasTop)->R_E_C_N_O_)
	
	RecLock("SN3",.F.)
	
	SN3->N3_CCONTAB:= cContaBem
	SN3->N3_CDEPREC:= cContaDes
	SN3->N3_CCDEPR := cContAcum
		
	SN3->(MsUnLock())	

	(cAliasTop)->(dbskip())		

EndDo

DbCloseArea(cAliasTop)

RestArea(aArea)

Return lRet

//--------------------------------------------------------------------------------

Static Function BuscaConta(cGrupo)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()
Local cContaBem := ""

DbSelectArea("SNG")
DbSetOrder(1)

If SNG->( DbSeek(xFilial("SNG")+ cGrupo))
	cContaBem := SNG->NG_CCONTAB
EndIf

RestArea(aArea)

Return cContaBem

//--------------------------------------------------------------------------------

Static Function BuscaDesp(cGrupo)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()
Local cContaBem := ""

DbSelectArea("SNG")
DbSetOrder(1)

If SNG->( DbSeek(xFilial("SNG")+ cGrupo))
	cContaBem := SNG->NG_CDEPREC
EndIf

RestArea(aArea)

Return cContaBem

//--------------------------------------------------------------------------------

Static Function BuscaAcum(cGrupo)

Local aArea     := GetArea()
Local lRet      := .F.
Local cQuery    := ""
Local cAliasTop := GetNextAlias()
Local cContaBem := ""

DbSelectArea("SNG")
DbSetOrder(1)

If SNG->( DbSeek(xFilial("SNG")+ cGrupo))
	cContaBem := SNG->NG_CCDEPR
EndIf

RestArea(aArea)

Return cContaBem

//--------------------------------------------------------------------------------

Static Function RelCta(aSN3)

Local aArea     := GetArea()
Local lRet      := .F.
Local oExcelApp := Nil
Local cPath     := "C:\Temp"
Local cArquivo  := "AtivosContas.XLS"
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
cAba    := "Registros alterados"
cTabela := "Ativo Fixo - Coopmil"

// Criação de nova aba 
oExcel:AddworkSheet(cAba)

// Criação de tabela
oExcel:AddTable (cAba,cTabela)

// Criação de colunas 

oExcel:AddColumn(cAba,cTabela,"ATIVO FILIAL"         ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"ATIVO FIXO"           ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"CENTRO CUSTO"         ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"CONTA DO BEM"         ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"CONTA DEPRECIACAO"    ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"CONTA DEP. ACUMULADA" ,2,1,.F.) 


DbSelectArea("SN3")
DbSetOrder(1)

For x := 1 to Len(aSN3)
	DbGoto(aSN3[X])
	oExcel:AddRow(cAba,cTabela, { FWFilialName(,SN3->N3_FILIAL,1) ,;
                                  SN3->N3_CBASE   ,; 
                                  SN3->N3_CUSTBEM ,; 
                                  SN3->N3_CCONTAB ,;
                                  SN3->N3_CDEPREC ,;
                                  SN3->N3_CCDEPR })
	
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