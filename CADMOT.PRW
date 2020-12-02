#Include 'Protheus.ch'
#Include 'FWMVCDEF.ch'

User Function CadMot()

	Local oBrowse

	oBrowse := FWMBrowse():New()
	oBrowse:SetAlias('Z01')
	oBrowse:SetDescription('Cadastro Motivos - COOPMIL')
	oBrowse:Activate()
	
Return

Static Function MenuDef()

	Local aRotina := {}

	ADD OPTION aRotina TITLE 'Visualizar'  ACTION 'VIEWDEF.CADMOT' OPERATION 2 ACCESS 0
	ADD OPTION aRotina TITLE 'Incluir'     ACTION 'VIEWDEF.CADMOT' OPERATION 3 ACCESS 0
	ADD OPTION aRotina TITLE 'Alterar'     ACTION 'VIEWDEF.CADMOT' OPERATION 4 ACCESS 0
	ADD OPTION aRotina TITLE 'Excluir'     ACTION 'VIEWDEF.CADMOT' OPERATION 5 ACCESS 0
	ADD OPTION aRotina TITLE 'Copiar'      ACTION 'VIEWDEF.CADMOT' OPERATION 9 ACCESS 0

Return aRotina

Static Function ModelDef()

	Local oModel

	Local oStruZ01 	:= FWFormStruct(1, "Z01", {|cCpo| ALLTRIM(cCpo)$"Z01_COD,Z01_MOTIVO" })
	Local oStruZ02  := FWFormStruct(1, "Z01", {|cCpo| ALLTRIM(cCpo)$"Z01_CODSUB,Z01_SUBMOT" })
  
	oModel := MPFormModel():New( 'INC_GRADE', /*bPreValidacao*/, /*bPosValidacao*/, /*bCommit*/, /*bCancel*/ )

	oModel:SetDescription("Cadastro Motivos - COOPMIL")

	oModel:addFields('Z01MASTER',,oStruZ01)

	oModel:addGrid('Z01DETAIL','Z01MASTER',oStruZ02)

	oModel:SetPrimaryKey({'Z01_CODSUB'})
 						
Return oModel

Static Function ViewDef()

	Local oModel := ModelDef()
	Local oView
	Local oStrZB5 	:= FWFormStruct(2, "Z01", {|cCpo| ALLTRIM(cCpo)$"Z01_COD,Z01_MOTIVO" })
	Local oStrZB6   := FWFormStruct(2, 'Z01', {|cCpo| ALLTRIM(cCpo)$"Z01_CODSUB,Z01_SUBMOT" } )
    
	oStrZB6:RemoveField( 'Z01_COD' )

	oView := FWFormView():New()
	oView:SetModel(oModel)
	oView:AddField('FORM_CAB' , oStrZB5,'Z01MASTER' )
	oView:AddGrid('FORM_GRID' , oStrZB6,'Z01DETAIL')

	oView:CreateHorizontalBox( 'BOX_FORM_CAB', 30)
	oView:CreateHorizontalBox( 'BOX_FORM_GRID', 70)

	oView:SetOwnerView('FORM_GRID','BOX_FORM_GRID')
	oView:SetOwnerView('FORM_CAB','BOX_FORM_CAB')

	oView:AddIncrementField( 'FORM_GRID', 'Z01_CODSUB' )
 	 	
Return oView