#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TCBROWSE.CH"
#INClUDE "FWMVCDEF.CH"
#INCLUDE "FWBROWSE.CH"

/*
Òezembro/2016* |C-CADRDÍ>* Autor: | Silverio Bastos - Anadi Soluções
I Descrição: | Cadastro De Para
I Uso:	|
Coopmil
I Data: | +--------+.
*/

//FONTE02
//pg1

User Function C_CADRDP()

	PRIVATE _cFil 		:= "Z7_FLAG=='X'"
	PRIVATE cCadastro 	:= "Cadastro De Para" , cString := "SZ7"
	PRIVATE aRotina 	:= { {"Pesquisar","AxPesqui" ,0,1},;
	{ "Manutenção"	, "U_C_CADDPF()",0,3},;
	{"Executar De Para" ,"U_C_CADDPC()", 0,3}}

	DbSelectArea(cString)
	DbSetOrder(1)
	DbGoTop()
	SET FILTER TO Z7_FLAG == 'X'


	mBrowse(6,1,22,75,cString,,,,,,,,,,,,,,)
	SET FILTER TO 

Return

/*
Programa	1	C_CADDPF	I Autor | Silverio Bastos	| Data | Dezembro/2016
Qp^^-.ricao	1	Cadastro	De Para	
— Uso	1	Coopmil		
*/

User Function C_CADDPF()

	Private _aArea := GetArea()

	Processa ({|| C_CADDPA(SZ7->Z7_CTADE) },"Processando, por favor aguarde...")
	RestArea(_aArea)

Return
/*
Programa	|C CADDPA 1 Autor | Silverio Bastos	| Data | Dezembro/2016
Uso	| Coopmil	
*/

Static Function C_CADDPA(_cCampox)
	//pg2

	Local oButton5 
	Local oButton6 
	Local oButton7
	Local oFont1 := TFont():New("MS Serif",,026,,.T.,,,,,.F.,.F.)
	Local oSay1
	Local oSay2
	Local oSay3
	Local oSay4
	Local oSay5
	Local oSay6
	Local oSay7
	Local oSay8
	Local oSay9
	Local oSay10
	Local oSayll
	Local oSayl2
	Local oSayl3
	Local AColsex2 := {}
	Local _cQryTRF := ""
	Private oEnchoic1 
	Private oMSNewGe1 
	Private oDlgCx 
	Private aHeadery
	Private aAcolsy 


	DEFINE MSDIALOG oDlgCx TITLE "Cadastro De Para" FROM 000, 000 TO 550, 1300 COLORS 0, 16777215 PIXEL
	//rr///>F.,.F.)

	fMSNewGe1(_cCampox) 
	aColsx := oMSNewGe1:Acols

	//oMSNewGel rbLDbClick := { | | U_FFCX351B () }	
	//oMSNewGel:oBrowse:bLDblClick := { | | oDlgCx:End () ,U_FFCX351B (oMSNewGel :oBrowse:nAt) }

	@ 250,026 BUTTON oButton7 PROMPT "Executar De Para" SIZE 060, 012 OF oDlgCx ACTION(U_C_CADDPB(oMSNewGe1:AHeader, oMSNewGe1:Acols), U_C_CADDPC(),oDlgCx:End()) PIXEL 
	@ 250, 166 BUTTON oButton5 PROMPT "Gravar" SIZE 040, 012 OF oDlgCx ACTION(U_C_CADDPB(oMSNewGe1:AHeader, oMSNewGe1:Acols), oDlgCx:End()) PIXEL 
	@ 250,266 BUTTON oButton6 PROMPT "Sair" SIZE 040, 012 OF oDlgCx ACTION oDlgCx:End() PIXEL

	ACTIVATE MSDIALOG oDlgCx CENTERED

Return


Static Function fMSNewGe1(_cCampox)

	Local nX := 0
	Local aHeaderEx := {}
	Local aFieldFill := {}
	Local aFields :={}
	Local aFieldTit := {}
	Local aAlterFields := {"Z7_CTADE","Z7_CTAATE"}
	Local _cQrySL1 := ""
	Public aColsEx := {}

	_cQrySL1 := " SELECT Z7_FILIAL, Z7_CTADE, Z7_DSCDE, Z7_CTAATE, Z7_DSCATE, Z7_HORA, Z7_DATA, " 
	_cQrySL1 += " Z7_USER, Z7_FLAG "
	_cQrySL1 += " From " + RetSqlName("SZ7") + " SZ7 "
	_cQrySL1 += " Where SZ7.Z7_FILIAL = '" + xFilial("SZ7") + "' And SZ7.Z7_FLAG = 'X' AND SZ7.D_E_L_E_T_ = '' "

	If Select("SZ4Q") > 0
		DbSelectArea("SZ4Q")
		DbCloseArea()
	EndIf

	TcQuery _cQrySL1 New Alias "SZ4Q"

	//pg3
	DbSelectArea("SZ4Q")
	dbGoTop()

	While !Eof()
		__cData 	:= date()
		__Hora		:= time()
		__Ctade		:= SZ4Q->Z7_CTADE
		__DscDe		:= SZ4Q->Z7_DSCDE	
		__CtaATE	:= SZ4Q->Z7_CTAATE
		__DscATE	:= SZ4Q->Z7_DSCATE
		__Flag		:= SZ4Q->Z7_FLAG

		Aadd(aColsEx, {__Ctade, __CtaATE,__DscDe,__DscAte,__Flag, .F.})
		dbskip ()
	Enddo


	aFields		:= {"Z7_CTADE","Z7_DSCDE","Z7_CTAATE", "Z7_DSCATE","Z7_FLAG"}
	aFieldTit 	:= {"Conta De","Desc. Cta De","Conta Ate","Desc. Cta Ate","Flag"}

	// Define field properties 
	DbSelectArea("SX3")
	SX3->(DbSetOrder(2))
	For nX := 1 to Len (aFields)
		If SX3->(DbSeek(aFields[nX]))
			Aadd(aHeaderEx,{aFieldTit[nx],SX3->X3_CAMPO,SX3->X3_PICTURE,SX3->X3_TAMANHO,SX3->X3_DECIMAL,SX3->X3_VALID,;
			SX3->X3_USADO,SX3->X3_TIPO,SX3->X3_F3,SX3->X3_CONTEXT,SX3->X3_CBOX, SX3->X3_RELACAO})
		Endif 
	Next nX
	oMSNewGe1 := MsNewGetDados():New( 002, 002, 500, 1200,GD_INSERT+GD_DELETE+GD_UPDATE,;
	"AllwaysTrue", "AllwaysTrue","",aAlterFields,,,"AllwaysTrue",,"AllwaysTrue", oDlgCx,aHeaderEx, aColsEx)
Return(aColSex)
/*
Programa	1	C_CADDPB | Autor	| Silverio Bastos	I Data | Dezembro/2016
Descricao	1	Rotina Gravação	SZ7	
Uso	1	Coopmil		
*/

User Function C_CADDPB(aHeadery,aAcolsy) 

	Local ixv		:= 0
	Local aColSex1	:= aAcolsy
	Local _cQryCo1 	:= ''

	For ixv := 1 to Len(aColSex1)
		If aColSex1[ixv][5] == 'X' 
			dbSelectArea("SZ7") 
			dbSetOrder(1)
			If !aAcolsy[ixv][Len(aHeadery) + 1]
				If !dbSeek(xFilial("SZ7")+aColSex1[ixv][1],.f.)
					__lRetx1 := .T.
				Else
					__lRetx1 := .F.
				Endif
				
				While !RecLock("SZ7",__lRetx1)
				Enddo

				SZ7->Z7_Filial	:= xFilial("SZ7") 
				SZ7->Z7_CTADE 	:= aColSex1[ixv][1]
				//pg4
				SZ7->Z7_DSCDE	:= aColSex1[ixv][3]
				SZ7->Z7_CTAATE	:= aColSex1[ixv][2]
				SZ7->Z7_DSCATE	:= aColSex1[ixv][4]
				SZ7->Z7_HORA	:= TIME()
				SZ7->Z7_DATA	:= DATE()
				SZ7->Z7_USER	:= USRRETNAME(__CUSERID)
				SZ7->Z7_FLAG	:= aColSex1[ixv][5]
				MsUnLock ()
			Else
				If dbSeek(xFilial("SZ7") + aColSex1[ixv][1],.f.) 
					While !RecLock("SZ7",.F.)
					Enddo 
					DBDelete ()
					MsUnLock()
				Endif 
			Endif 
		Endif 
	Next ixv
RETURN


/*	
Programa	1	C_CADDPC | Autor	1	Silverio Bastos	I Data | Dezembro/2016
fôjj^ricao	1	Rotina Execução	De	Para	
Uso	1	Coopmil			
*/

USER FUNCTION C_CADDPC()

	Local _cQrySZ7 := ""
	Local _cTimeex := TIME()
	Local _cDataex := dtos(DATE())
	Local _cUserex := USRRETNAME (__CUSERID)
	Public aColsEx := {}


	_cQrySZ7 := " Select Z7_FILIAL, Z7_CTADE, Z7_DSCDE, Z7_CTAATE, Z7_DSCATE, Z7_HORA, Z7_DATA, "
	_cQrySZ7 += " Z7_USER, Z7_FLAG, SZ7.R_E_C_N_O_ Z7RECNO "
	_cQrySZ7 += " From " + RetSqlName ("SZ7") + " SZ7 "
	_cQrySZ7 += " Where SZ7.Z7_FILIAL = '" + xFilial ("SZ7") + "' And SZ7.Z7_FLAG = 'X' And SZ7.D_E_L_E_T_ = '' And Z7_CTAATE <>'' "

	If Select("SZ7Q") > 0
		DbSelectArea("SZ7Q")
		DBCloseArea()
	Endif

	TcQuery _cQrySZ7 New Alias "SZ7Q"

	DbSelectArea("SZ7Q")
	DbGoTop()


	While !Eof ()
		// Verificar os campos na SZ6 
		dbSelectArea("SZ6") 
		dbSetOrder(1) 
		dbGotop()
		While !Eof()
			_cTab	:= SZ6->Z6_TABELA 
			_cCpo 	:= SZ6->Z6_CAMPO

			// Update no campo da tabela do Cadastro Referencial SZ6
			_cQRY := " Update " + RetSqlName(___cTab) + " SET " + ___cCpo + " =  '" + S27Q->Z7_CTAATE + "' WHERE " + cCpo+ "= 1"+SZ7Q->Z7_CTADE+"' AND D_E_L_E_T_ = ""
			TCSQLEXEC(_cQry)
			//pg5
			dbSelectArea("SZ6") 
			dbSkip ()
		Enddo

		//Atualizar a tabela SZ7 
		DBSelectArea("SZ7")
		DbGoTo(SZ7Q->Z7RECNO)
		If SZ7->(Recno()) == SZ7Q->Z7RECNO
			While !Reclock("SZ7",.f.)
			EndDo

			SZ7->Z7_FLAG := " "
			SZ7->Z7_DTEXEC := Date()
			SZ7->Z7_HREXEC := Time()
			SZ7->Z7_USEXEC := _cUserex 
			SZ7->(MsUnlock())
		Endif

		DbSelectArea("SZ7Q")
		dbSkip ()

	Enddo

	/*
	DbSelectArea ("SZ7Q")
	DbGoTop()
	Update na tabela SZ7 - campo Z7_FLAG _cQRY := "Update "+RetSqlName("SZ?")+" SET Z7_FLAG = '' Z7_HREXEC = '"+ _cTimeex+"', Z7_USEXEC = 1"+_cUserex+"' D_E_L_E_T_ = ' 1 ”
	TCSQLEXEC(_cQry)
	Endif
	Z7_DTEXEC = ' "+_cDataex+" WHERE Z7 FLAG = 'X' AND
	*/

	MsgInfo("Termino de Processamento !!!")

RETURN





