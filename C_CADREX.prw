#INCLUDE "RWMAKE.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TCBROWSE.CH"
#INCLUDE "FWMVCDEF.CH"
#INCLUDE "FWBROWSE.CH"

 /*
I Programa: | C_CADREX| Autor: | CLAUDIO AMBROSINI OUTUBRO 2019
I Descrição: l Cadastro Referencial+
I Uso:      I     |
Coopmil          
*/
//FONTE01
User Function C_CADREX()

      Private _cFil           := ""
      Private cCadastro       := "Cadastro Referencial", cString := "SZ6"
      Private aRotina         := { {"Pesquisar","AxPesqui",0,1},{"Manutenção" ,"U_C_CADREF()"   ,0,3}}

      DbSelectArea (cString)
      DbSetOrder(1)
      DbGoTop()

      mBrowse (6,1,22,75,cString,,,,,,)

Return
 
/*
I Programa: | C_CADREX| Autor: | CLAUDIO AMBROSINI OUTUBRO 2019
I Descrição: l Cadastro Referencial+
I Uso:      I     |
Coopmil          
*/
 
User Function C_CADREF()

      Local _aArea     := GetArea()
      Private _cTabelax := ""

      Processa({|| C_CADREA(SZ6->Z6_TABELA)},"Processando, por favor aguarde...")
      RestArea(_aArea)
Return

/*
I Programa: | C_CADREX| Autor: | CLAUDIO AMBROSINI OUTUBRO 2019
I Descrição: l Cadastro Referencial+
I Uso:      I     |
Coopmil          
*/
 
Static Function C_CADREA(_cTabelax)

 
      Local oButton5
      Local oButton6
      Local oButton7
      Local oCxBol
      Local nCxBol      := nCxBol1 := 0
      Local oCxChq
      Local nCxChq      := nCxChq1 := 0
      Local oCxCrd
      Local nCxCrd      := nCxCrd1 := 0
      Local oCxDeb
      Local nCxDeb      := nCxDeb1 := 0  
      Local oCxDep
      Local nCxDep      := nCxDep1 := 0  
      Local oCxEnt
      Local nCxEnt      := nCxEnt1 := 0
      Local oCxTrf
      Local nCxTrf      := nCxTrf1 := 0
      Local oCxDin
      Local nCxDin      := nCxDin1 := 0
      Local oCxobs
      Local cCxObs      := space(100)
      Local oCxOut
      Local nCxOut      := nCxOut1 := 0
      Local oCxSrg
      Local nCxSrg      := 0
      Local oCxTrc
      Local oCxTot      := nCxEnt1 := 0
      Local nCxTot      := 0
      Local oFont1      := TFont():New("MS Serif",,026,,.T.,,,,,.F.,.F.)
      Local oSay1
      Local oSay2
      Local oSay3
      Local oSay4
      Local oSay5
      Local oSay6
      Local oSay7
      Local oSay8
      Local oSay9
      Local oSayl0
      Local oSayll
      Local oSayl2
      Local oSayl3
      Local AColsex2    := {}
      Local _cQryTRF    := ""

      Private _nRetCx2:= .F.
      Private oEnchoicl
      Private oMSNewGe1
      Private oDlgCx
      Private aHeadery
      Private aAcolsy


      DEFINE MSDIALOG oDlgCx TITLE "Cadastro Referencial" FROM 000, 000 TO 480, 630 COLORS 0, 16777215 PIXEL

      fMSNewGel(_cTabelax)
      aColsx := oMSNewGe1:Acols

      //oMSNewGelrbLDbClick := {II U_FFCX351B() }
      //oMSNewGel:oBrowse:bLDblClick := {II oDlgCx:End(),U_FFCX351B(oMSNewGel:oBrowse:nAt)}

      @ 224, 166 BUTTON oButton5 PROMPT "Gravar" SIZE 040, 012 OF oDlgCx ACTION(U_C_CADREB(oMSNewGe1:AHeader,oMSNewGe1:Acols),oDlgCx:End()) PIXEL
      @ 224, 266 BUTTON oButton6 PROMPT "Sair" SIZE 040, 012 OF oDlgCx ACTION oDlgCx:End() PIXEL

      ACTIVATE MSDIALOG ODlgCx CENTERED

Return(_nRetCx2)

 

Static Function fMSNewGel(_cTabelax)

      Local nX
      Local aHeaderEx 	:= {}
      Local aFieldFill 	:= {}
      Local aFields 	:= {}
      Local aFieldTit 	:= {}
      Local aAlterFields:= {"Z6_TABELA","Z6_CAMPO"}
      Local _cQrySLl    := ""
      Local _cLinok     := "U_VALIDCMP()"
      Public aColsEx	:= {}
 

      _cQrySLl := "SELECT Z6_FILIAL, Z6_TABELA, Z6_DSCTAB, Z6_CAMPO, Z6_HORA, Z6_DATA,"
      _cQrySLl += "Z6_USER "
      _cQrySLl += " From " + RetSqlName("SZ6") + " SZ6 "
      _cQrySLl += " Where SZ6.Z6_FILIAL =  '" + xFilial("SZ6") + "'  And SZ6.Z6_TABELA = '" + _cTabelax + "'  And SZ6.D_E_L_E_T_ = '' "
 

      If Select ("SZ4Q") > 0
            DbSelectArea("SZ4Q")
            DbCloseArea()
      EndIf

      TcQuery _cQrySLl New Alias "SZ4Q"
      DbSelectArea("SZ4Q")
      DbGotop()

      While !EOF()
            __Data          := date()
            __Hora          := Time()
            __Tabela    	:= SZ4Q->Z6_TABELA
            __DscTab    	:= SZ4Q->Z6_DSCTAB
            __Campo         := SZ4Q->Z6_CAMPO

            Aadd(aColsex, {__tabela, __DscTab, __Campo, .F.})
            dbSkip()
      EndDo

      aFields           :=    {"Z6_TABELA","Z6_DSCTAB","Z6_CAMPO"}
      aFieldTit   :=    {"Tabela","Desc.Tab","Campo"}


      // Define field properties
      DbSelectArea("SX3")
      SX3->(DbSetOrder(2))


      For nX := 1 to Len (aFields)

            If SX3->(DbSeek(aFields[nX]))
                  Aadd(aHeaderEx,{aFieldTit[nx],SX3->X3_CAMPO,SX3->X3_PICTURE,SX3->X3_TAMANHO, SX3->X3_DECIMAL,SX3->X3_VALID,;
                                    SX3->X3_USADO,SX3->X3_TIPO,SX3->X3_F3,SX3->X3_CONTEXT,SX3->X3_CBOX,SX3->X3_RELACAO})
            Endif
      Next nX

      //110, 004, 218, 307
      oMSNewGe1 := MsNewGetDados() :New ( 004, 004, 218, 307,GD_INSERT+GD_DELETE+GD_UPDATE,;
      _cLinok, "AllwaysTrue","",aAlterFields,,,"AllwaysTrue",,"AllwaysTrue", oDlgCx, aHeaderEx,aColsEx )

Return(aColSex)

/*
+                 +     +    
| Programa  I     C_CADREB    |     Autor |     Silverio Bastos
+     +     +     +    
| Descricao | Rotina Gravação SZ6
+-( ‘            
| èjj I Coopmil
*/

User Function C_CADREB(aHeadery,aAcolsy)

      Local ixv         := 0
      Local aColSexl    := aAcolsy
      Local _cQryCol    := ""
    
      For ixv := 1 to Len(aColSexl)

            dbSelectArea("SZ6")
            dbSetOrder(1)
            If !aAcolsy[ixv][Len(aHeadery) + 1]

                  If !dbSeek(xFilial("SZ6") + aColSexl[ixv][1] + aColSexl[ixv][3],.f.)
                        While lRecLock("SZ6",.T.)
                        Enddo

                        SZ6->Z6_Filial    := xFilial("SZ6")
                        SZ6->Z6_TABELA    := aColSexl[ixv][1]    
                        SZ6->Z6_DSCTAB    := aColSexl[ixv][2]
                        SZ6->Z6_CAMPO     := aColSexl[ixv][3]
                        SZ6->Z6_HORA      := time()
                        SZ6->Z6_DATA      := Date()
                        SZ6->Z6_USER      := USRRETNAME(__CUSERID)
                        MsUnLock()
                  Endif
            Else
                  If dbSeek(xFilial("SZ6")+aColSexl[ixv][1]+aColSexl[ixv][3],.f.)
                        While !RecLock("SZ6",.F.)
                        EndDo
                        DBDelete()
                        MsUnLock()
                  Endif
            Endif
      Next ixv

      //MsgAlert("Gravação Concluida !!!")

Return

/*
| Programa: I C_CTBCX3| Autor: | Silverio Bastos - Anadi Consultoria      | Data: |
Dezembro/2016 I
descrição: I Consulta padrão para campos do SX3, de acordo com a tabela recebida por parâmetro I
| Uso:      I
Coopmil     I
*/


User Function C_CTBCX3(_cAliasX3)

      Local lRet := .f.
      Local cFiltro := "SX3->X3_CONTEXT != 'V' .AND. SX3->X3_ARQUIVO =='" + _cAliasX3 + "'"
      Local oDlg, oBrowse,oMainPanel,oPanelBtn,oBtnOK,oBtnCan,oColumnl,oColumn2,oColumn3,oColumn4
      Local cTitulo := MvParDef := _cRet := "", nx := 0
      //Local _cField := ''//claudio
      Default _cAliasX3 := ""

      IF Empty(_cAliasX3)
            MsgStop("Alias da Tabela não informado.","Atenção")
            return(_cField)
      EndIF

      //Define MsDialog oDlg From 0, 0, to 390, 515 Title "Campos da Tabela" + _cAliasSx3 pixel Of oMainWnd
      Define MsDialog oDlg TITLE  "Campos da Tabela" + _cAliasSx3   From 0,0 To 390,515 OF oMainWnd PIXEL
 
      @00, 00 MsPanel oMainPanel Size 250, 80
      oMainPanel:Align := CONTROL_ALIGN_ALLCLIENT

      @00, 00 MsPanel oPanelBtn Size 250, 15
      oPanelBtn:Align := CONTROL_ALIGN_BOTTOM

      Define FwBrowse oBrowse DATA TABLE ALIAS "SX3" NO CONFIG NO REPORT DOUBLECLICK {|| lRet := .T., oDlg:End()} NO LOCATE Of oMainPanel

      ADD COLUMN oColumnl DATA {|| SX3->X3_CAMPO}    Title "Campo"      Size Len(SX3->X3_CAMPO) Of oBrowse 
      ADD COLUMN oColumn2 DATA {|| SX3->X3_TITULO} Title "Titulo"      Size Len(SX3->X3_TITULO) Of oBrowse
      ADD COLUMN oColumn3 DATA {|| SX3->X3_DESCRIC} Title "Descricao" Size Len(SX3->X3_DESCRIC) Of oBrowse
      
      oBrowse:SetFilterDefault (cFiltro)
      oBrowse:Activate()


      Define SButton oBtnOK From 02, 02 Type 1 Enable Of oPanelBtn ONSTOP "Confirmar" Action (lRet := .T., oDlg:End())
      Define SButton oBtnCan From 02, 32 Type 2 Enable Of oPanelBtn ONSTOP "Cancelar" Action (lRet := .F., oDlg:End())

      Activate MsDialog oDlg Centered


Return .t.

 /*
| Programa: I VALIDCMP| Autor: | Silverio Bastos - Anadi Consultoria      | Data: |
Dezembro/2016 I
descrição: I VERIFICA SE O CAMPO EXISTE NA TABELA DE REFERENCIA
| Uso:      I
Coopmil     I

*/

User Function U_VALIDCMP()

      Local _aXArea     := GetArea()
      Local _cParam     := ""
      Local _nPos       := 0
      Local _lrety      := .T.
      Local _xTabela    := SPACE(3)
      Local _oTela      := oMSNewGel

      _cParam                 := oTela:aCols[_oTela:nAt][3]
      __Tabela          := oTela:aCols[_oTela:nAt][1]
      _nPos             := ASCAN(_oTela:aCols,{ |x| x[3]== _cParam })


      dbSQlectArea("SX3")
      dbSetorder(2)
      If !dbSeek(_cParam,.f.)
            MsgAlert("Campo " + _cParam +" não existe !!!")
            _lrety := .F.
      Else
            xTabela := SX3->X3_ARQUIVO
            IF _xTabela <> _Tabela
                  MsgAlert("Campo" + _cParam + "nao existe nesta tabela" + _tabela + "!!!")
                  _lrety := .F.
            EndIF
      EndIf

      //verifica se o campo jA existe na acols atual desconsiderando a linha posiciona

      IF _lRety .AND. _nPos > 0 .AND. _nPos != _oTela:nAt
            MsgAlert("Campo " + _cParam+ " já informado neste Grid !!!")
            _lRety := .F.
      Endif

      RestArea(_aXArea)

Return (_lrety)
