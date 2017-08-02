/*/
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦Programa  ¦ ExpFWExc ¦ Autor ¦ Wagner Cabrera        ¦ Data¦16/02/2016 ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Descricao ¦ Rotina de Exportação para Excel a Partir do FWMBrowse	  ¦¦¦
¦¦¦          ¦ Como Utilizar:				                              ¦¦¦
¦¦¦          ¦ U_ExpFWExc(FWMBrowse,'ARQTRAB','COLUNA_DE_QUEBRA','TITULO')¦¦¦
¦¦¦          ¦								                              ¦¦¦
¦¦¦          ¦ As linhas serão pintadas com as cores da Legenda do Browse ¦¦¦
¦¦¦          ¦								                              ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Uso       ¦ GERAL                                                      ¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
/*/

#include "protheus.ch"
#include "AppExcel.ch"

user function ExpFWExc(_OBROWSE, cArq, cQUEBRASOMA, _cTitulo) 
	PROCESSA({|| ExpFWExc(_OBROWSE, cArq, cQUEBRASOMA, _cTitulo)})
return

static function ExpFWExc(_OBROWSE, cArq, cQUEBRASOMA, _cTitulo)
	Local oExcel := AppExcel():NEW()
	Local oCellCab := AppExcCell():New()
	Local oCellTitulo := AppExcCell():New()
	Local oCellTit2 := AppExcCell():New()
	
	Local oCellCol := AppExcCell():New()
	Local oCellData := AppExcCell():New()
	Local oCellNumero := AppExcCell():New()
	
	Local oCellDefault := AppExcCell():New()
	Local oCellNDefault := AppExcCell():New()
	Local oCellDDefault := AppExcCell():New()	
	
	Local oCellRED := AppExcCell():New()
	Local oCellNRED := AppExcCell():New()
	Local oCellDRED := AppExcCell():New()
	
	Local oCellGREEN := AppExcCell():New()
	Local oCellNGREEN := AppExcCell():New()
	Local oCellDGREEN := AppExcCell():New()
	
	Local oCellYELLOW := AppExcCell():New()
	Local oCellNYELLOW := AppExcCell():New()
	Local oCellDYELLOW := AppExcCell():New()
	
	Local oCellBLUE := AppExcCell():New()
	Local oCellNBLUE := AppExcCell():New()
	Local oCellDBLUE := AppExcCell():New()
	
	Local oCellWHITE := AppExcCell():New()
	Local oCellNWHITE := AppExcCell():New()
	Local oCellDWHITE := AppExcCell():New()
	
	Local oCellGRAY := AppExcCell():New()
	Local oCellNGRAY := AppExcCell():New()
	Local oCellDGRAY := AppExcCell():New()
	
	Local oCellORANGE := AppExcCell():New()
	Local oCellNORANGE := AppExcCell():New()
	Local oCellDORANGE := AppExcCell():New()
	
	Local oCellBROWN := AppExcCell():New()
	Local oCellNBROWN := AppExcCell():New()
	Local oCellDBROWN := AppExcCell():New()
	
	Local oCellPINK := AppExcCell():New()
	Local oCellNPINK := AppExcCell():New()
	Local oCellDPINK := AppExcCell():New()
	
	Local oCellBLACK := AppExcCell():New()
	Local oCellNBLACK := AppExcCell():New()
	Local oCellDBLACK := AppExcCell():New()
	
	Local oCellVIOLET := AppExcCell():New()
	Local oCellNVIOLET := AppExcCell():New()
	Local oCellDVIOLET := AppExcCell():New()
	
	Local oCellHGREEN := AppExcCell():New()
	Local oCellNHGREEN := AppExcCell():New()
	Local oCellDHGREEN := AppExcCell():New()
	
	Local oCellLBLUE := AppExcCell():New()
	Local oCellNLBLUE := AppExcCell():New()
	Local oCellDLBLUE := AppExcCell():New()
	
	//Configuração das fontes
	Local oFont := AppExcFont():New("Arial","10","#000000") 
	Local oFontCab := AppExcFont():New("Arial","12","#000000") 
	Local oFontTitulo := AppExcFont():New("Cambria","18","#363636") 
	Local oFontBranco := AppExcFont():New("Arial","10","#FFFFFF") 
	
	Local nLinha := 1
	Local nPosTitulo := 1
	Local nPosField := 2
	Local nPosTipo := 3
	Local aCores := {}
	Local aSoma := {}
	Local aSomaTotal := {}
	Local CULTIMAQUEBRA := ''
	
	Local bTotal  := .F.
	Local nI	  := 0
	Local bExpres := .F.
	
	//Seta Negrito
	oFontTitulo:SetBold(.T.)	
	
	//personalização das celulas
	oCellCab:SetCellColor("#87CEFA")
	oCellCab:SetABorders(.T.)       
	oCellCab:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellCab:SetFont(oFontCab)
	
	oCellTitulo:SetVertAlign( VERTICAL_ALIGN_CENTER ) //ALINHA NO CENTRO
	oCellTitulo:SetHorzAlign( HORIZONTAL_ALIGN_CENTER ) //ALINHA NO CENTRO
	oCellTitulo:SetFont(oFontTitulo)
	
	oCellTit2:SetFont(oFont)		
	
	oCellDefault:SetABorders(.T.)
	oCellDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDefault:SetFont(oFont)
	
	oCellNDefault:SetABorders(.T.)
	oCellNDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNDefault:SetFont(oFont)
	
	oCellDDefault:SetABorders(.T.)
	oCellDDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDDefault:SetFont(oFont)		
	
	oCellRED:SetABorders(.T.)  
	oCellRED:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellRED:SetCellColor("#FF7F50")
	oCellRED:SetFont(oFontBranco)
	
	oCellNRED:SetABorders(.T.)  
	oCellNRED:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNRED:SetCellColor("#FF7F50")
	oCellNRED:SetFont(oFontBranco)
	
	oCellDRED:SetABorders(.T.)  
	oCellDRED:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDRED:SetCellColor("#FF7F50")
	oCellDRED:SetFont(oFontBranco)		

	oCellGREEN:SetABorders(.T.)   
	oCellGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS) 
	oCellGREEN:SetCellColor("#3CB371")
	oCellGREEN:SetFont(oFont)
	
	oCellNGREEN:SetABorders(.T.)   
	oCellNGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS) 
	oCellNGREEN:SetCellColor("#3CB371")
	oCellNGREEN:SetFont(oFont)
	
	oCellDGREEN:SetABorders(.T.)   
	oCellDGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS) 
	oCellDGREEN:SetCellColor("#3CB371")
	oCellDGREEN:SetFont(oFont)		

	oCellYELLOW:SetABorders(.T.)  
	oCellYELLOW:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellYELLOW:SetCellColor("#FFFF00")
	oCellYELLOW:SetFont(oFont)
	
	oCellNYELLOW:SetABorders(.T.)  
	oCellNYELLOW:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNYELLOW:SetCellColor("#FFFF00")
	oCellNYELLOW:SetFont(oFont)

	oCellDYELLOW:SetABorders(.T.)  
	oCellDYELLOW:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDYELLOW:SetCellColor("#FFFF00")
	oCellDYELLOW:SetFont(oFont)		

	oCellBLUE:SetABorders(.T.)  
	oCellBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellBLUE:SetCellColor("#6495ED")
	oCellBLUE:SetFont(oFontBranco)
	
	oCellNBLUE:SetABorders(.T.)  
	oCellNBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNBLUE:SetCellColor("#6495ED")
	oCellNBLUE:SetFont(oFontBranco)
		
	oCellDBLUE:SetABorders(.T.)  
	oCellDBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDBLUE:SetCellColor("#6495ED")
	oCellDBLUE:SetFont(oFontBranco)		
	
	oCellWHITE:SetABorders(.T.)  
	oCellWHITE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellWHITE:SetFont(oFont)
	
	oCellNWHITE:SetABorders(.T.)  
	oCellNWHITE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNWHITE:SetFont(oFont)
	
	oCellDWHITE:SetABorders(.T.)  
	oCellDWHITE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDWHITE:SetFont(oFont)		

	oCellGRAY:SetABorders(.T.)  
	oCellGRAY:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellGRAY:SetCellColor("#808080")
	oCellGRAY:SetFont(oFont)
	
	oCellNGRAY:SetABorders(.T.)  
	oCellNGRAY:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNGRAY:SetCellColor("#808080")
	oCellNGRAY:SetFont(oFont)
	
	oCellDGRAY:SetABorders(.T.)  
	oCellDGRAY:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDGRAY:SetCellColor("#808080")
	oCellDGRAY:SetFont(oFont)				

	oCellORANGE:SetABorders(.T.)  
	oCellORANGE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellORANGE:SetCellColor("#FFA500")
	oCellORANGE:SetFont(oFont)
	
	oCellNORANGE:SetABorders(.T.)  
	oCellNORANGE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNORANGE:SetCellColor("#FFA500")
	oCellNORANGE:SetFont(oFont)	
	
	oCellDORANGE:SetABorders(.T.)  
	oCellDORANGE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDORANGE:SetCellColor("#FFA500")
	oCellDORANGE:SetFont(oFont)		

	oCellBROWN:SetABorders(.T.)  
	oCellBROWN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellBROWN:SetCellColor("#8B4513")
	oCellBROWN:SetFont(oFontBranco)
	
	oCellNBROWN:SetABorders(.T.)  
	oCellNBROWN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNBROWN:SetCellColor("#8B4513")
	oCellNBROWN:SetFont(oFontBranco)
	
	oCellDBROWN:SetABorders(.T.)  
	oCellDBROWN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDBROWN:SetCellColor("#8B4513")
	oCellDBROWN:SetFont(oFontBranco)		

	oCellPINK:SetABorders(.T.)  
	oCellPINK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellPINK:SetCellColor("#FF1493")
	oCellPINK:SetFont(oFont)
	
	oCellNPINK:SetABorders(.T.)  
	oCellNPINK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNPINK:SetCellColor("#FF1493")
	oCellNPINK:SetFont(oFont)
	
	oCellDPINK:SetABorders(.T.)  
	oCellDPINK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDPINK:SetCellColor("#FF1493")
	oCellDPINK:SetFont(oFont)		

	oCellBLACK:SetABorders(.T.)  
	oCellBLACK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellBLACK:SetCellColor("#000000")
	oCellBLACK:SetFont(oFontBranco)
	
	oCellNBLACK:SetABorders(.T.)  
	oCellNBLACK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNBLACK:SetCellColor("#000000")
	oCellNBLACK:SetFont(oFontBranco)
	
	oCellDBLACK:SetABorders(.T.)  
	oCellDBLACK:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDBLACK:SetCellColor("#000000")
	oCellDBLACK:SetFont(oFontBranco)		
	
	oCellVIOLET:SetABorders(.T.)  
	oCellVIOLET:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellVIOLET:SetCellColor("#EE82EE")
	oCellVIOLET:SetFont(oFont)
	
	oCellNVIOLET:SetABorders(.T.)  
	oCellNVIOLET:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNVIOLET:SetCellColor("#EE82EE")
	oCellNVIOLET:SetFont(oFont)
	
	oCellDVIOLET:SetABorders(.T.)  
	oCellDVIOLET:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDVIOLET:SetCellColor("#EE82EE")
	oCellDVIOLET:SetFont(oFont)		

	oCellHGREEN:SetABorders(.T.)  
	oCellHGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellHGREEN:SetCellColor("#006400")
	oCellHGREEN:SetFont(oFontBranco)
	
	oCellNHGREEN:SetABorders(.T.)  
	oCellNHGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNHGREEN:SetCellColor("#006400")
	oCellNHGREEN:SetFont(oFontBranco)
	
	oCellDHGREEN:SetABorders(.T.)  
	oCellDHGREEN:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDHGREEN:SetCellColor("#006400")
	oCellDHGREEN:SetFont(oFontBranco)		

	oCellLBLUE:SetABorders(.T.)  
	oCellLBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellLBLUE:SetCellColor("#ADD8E6")	
	oCellLBLUE:SetFont(oFont)
	
	oCellNLBLUE:SetABorders(.T.)  
	oCellNLBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNLBLUE:SetCellColor("#ADD8E6")	
	oCellNLBLUE:SetFont(oFont)
	
	oCellDLBLUE:SetABorders(.T.)  
	oCellDLBLUE:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDLBLUE:SetCellColor("#ADD8E6")	
	oCellDLBLUE:SetFont(oFont)		

	//Titulo do Relatório
	oExcel:Merge(nLinha,1,len(_OBROWSE:AFIELDS)-1,0,_cTitulo, oCellTitulo)
	nLinha++
	oExcel:Merge(nLinha,1,len(_OBROWSE:AFIELDS)-1,0,'Data do Relatório: ' + DTOC(DATE()) + ' ' + Time() + ' - Usuário: ' + cUserName, oCellTit2)
	nLinha++
	
	//Cabeçalho
	nLinha++
	For nField := 1 To Len(_OBROWSE:AFIELDS)
		oExcel:AddCell(nLinha,nField, _OBROWSE:AFIELDS[nField][nPosTitulo], oCellCab)
	Next nField
	
	//Carrega Cores de Legenda
	if len(_OBROWSE:ALEGENDS)>0
		for i := 1 To Len(_OBROWSE:ALEGENDS[1][2]:ALEGEND)
			AADD(aCores, {substr(_OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1], 1, AT('=', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1])-1), ; 
							         _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1], ;
	 				       			 _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][3]})
								
			/*AADD(aCores, {substr(_OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1], 1, AT('=', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1])-1), ; 
				          substr(_OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1], AT('=', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1])+1, len(_OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1])), ;
				       			 _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][3]})
			IF bExpres == .F.
				bExpres := AT('>', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1]) > 0 .OR. AT('<', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1]) > 0 .OR. AT('=', _OBROWSE:ALEGENDS[1][2]:ALEGEND[i][1]) > 0
			ENDIF*/				       			 
		Next i	
	endif
	
	(cArq)->(DbGoTop())
	
	while (cArq)->(!EOF())
		nLinha++
		
		oCellCol := oCellDefault 
		oCellData := oCellDDefault 	 
		oCellNumero := oCellNDefault
			
		For i:=1 To Len(aCores)
			//if IIF(bExpres == .F., ("'"+Alltrim((cArq)->&(aCores[i][1]))+"'"=aCores[i][2]), &(aCores[i][2]))
			if &(aCores[i][2])
				if aCores[i][3]='BR_VERMELHO'
					oCellCol := oCellRED
					oCellData := oCellDRED 	 
					oCellNumero := oCellNRED
					EXIT
				endif
				
				if aCores[i][3]='BR_VERDE'
					oCellCol := oCellGREEN
					oCellData  := oCellDGREEN
					oCellNumero := oCellNGREEN
					EXIT  	 
				endif						
						
				if aCores[i][3]='BR_AMARELO'
					oCellCol := oCellYELLOW
					oCellData  := oCellDYELLOW
					oCellNumero := oCellNYELLOW
					EXIT  	 
				endif

				if aCores[i][3]='BR_AZUL'
					oCellCol := oCellBLUE
					oCellData  := oCellDBLUE
					oCellNumero := oCellNBLUE
					EXIT  	 
				endif
					
				if aCores[i][3]='BR_BRANCO'
					oCellCol := oCellWHITE
					oCellData  := oCellDWHITE
					oCellNumero := oCellNWHITE
					EXIT  	 
				endif		
						
				if aCores[i][3]='BR_CINZA'
					oCellCol := oCellGRAY
					oCellData := oCellDGRAY
					oCellNumero := oCellNGRAY  
					EXIT	 
				endif											

				if aCores[i][3]='BR_LARANJA'
					oCellCol := oCellORANGE
					oCellData := oCellDORANGE
					oCellNumero := oCellNORANGE 
					EXIT 	 
				endif	

				if aCores[i][3]='BR_MARROM'
					oCellCol := oCellBROWN
					oCellData := oCellDBROWN
					oCellNumero := oCellNBROWN 
					EXIT 	 
				endif	
						
				if aCores[i][3]='BR_PINK'
					oCellCol := oCellPINK
					oCellData := oCellDPINK
					oCellNumero := oCellNPINK  	
					EXIT 
				endif			
						
				if aCores[i][3]='BR_PRETO'
					oCellCol := oCellBLACK
					oCellData := oCellDBLACK
					oCellNumero := oCellNBLACK 
					EXIT 	 
				endif		
						
				if aCores[i][3]='BR_VIOLETA'
					oCellCol := oCellVIOLET
					oCellData := oCellDVIOLET
					oCellNumero := oCellNVIOLET
					EXIT
				endif																

				if aCores[i][3]='BR_VERDE_ESCURO'
					oCellCol := oCellHGREEN
					oCellData := oCellDHGREEN
					oCellNumero := oCellNHGREEN  
					EXIT	 
				endif	

				if aCores[i][3]='BR_AZUL_CLARO'
					oCellCol := oCellLBLUE
					oCellData := oCellDLBLUE
					oCellNumero := oCellNLBLUE
					EXIT  	 
				endif
			endif
		Next i	
		
		For nField := 1 To Len(_OBROWSE:AFIELDS)
			if (_OBROWSE:AFIELDS[nField][nPosTipo]='D')
				oCellData:SetFormat(4)
			
				oExcel:AddCell(nLinha,nField, If(!Empty((cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])),(cArq)->&(_OBROWSE:AFIELDS[nField][nPosField]),""), oCellData)
			else 
				if (_OBROWSE:AFIELDS[nField][nPosTipo]='N')
					oCellNumero:SetFormat(5)
					oExcel:AddCell(nLinha,nField,(cArq)->&(_OBROWSE:AFIELDS[nField][nPosField]), oCellNumero)
					
					//SOMA DAS COLUNAS VALORES
					nPOSSoma := ASCAN(aSoma,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
					
					FOR nI := 1 TO LEN(_OBROWSE:AFIELDS)
						IF (_OBROWSE:AFIELDS[nField][nPosTipo] = 'C')
							bTotal := AT("*TOTAL*", UPPER((cArq)->&(_OBROWSE:AFIELDS[nI][nPosField]))) > 0
						ENDIF	
						IF bTotal
							EXIT
						ENDIF	
					NEXT					
					if (nPOSSoma>0)
						aSoma[nPOSSoma, 2] += (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])
					else 
						AADD(aSoma, {_OBROWSE:AFIELDS[nField][nPosField], (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])})
					endif
					
					nPOSSoma := ASCAN(aSomaTotal,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
					IF !bTotal
						if (nPOSSoma>0)
							aSomaTotal[nPOSSoma, 2] += (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])	
						else 
							AADD(aSomaTotal, {_OBROWSE:AFIELDS[nField][nPosField], (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])})
						endif
						bTotal := .F.					
					ENDIF	
				else
					oCellCol:SetFormat(0)
					oExcel:AddCell(nLinha,nField,(cArq)->&(_OBROWSE:AFIELDS[nField][nPosField]), oCellCol)
				endif
			endif
			
			if (_OBROWSE:AFIELDS[nField][nPosField]=cQUEBRASOMA)
				if !Empty((cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])) 
					cULTIMAQUEBRA := (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])
				endif	
			endif
		Next nField
	
		(cArq)->(dbskip())
		
		if ((cArq)->(EOF()) .or. (cULTIMAQUEBRA<>(cArq)->&(cQUEBRASOMA))) .and. (len(aSoma)>0) .and. !Empty(cQUEBRASOMA)
			lImpSoma := .T.
			nLinha++
			
			oCellCol := oCellDefault 
			oCellNumero := oCellNDefault			
			
			For nField := 1 To Len(_OBROWSE:AFIELDS)
				nPOSSoma := ASCAN(aSoma,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
				
				if (nPOSSoma>0)
					oCellNumero:SetFormat(5)
					oExcel:AddCell(nLinha,nField, aSoma[nPOSSoma,2], oCellNumero)	
				else
					oCellCol:SetFormat(0)
					oExcel:AddCell(nLinha,nField,if(lImpSoma,'TOTALIZAÇÃO--->',''), oCellCol)
					lImpSoma := .F.					
				endif
						
			Next nField
			
			aSoma := {}		
		endif				
	enddo
	
	if len(aSomaTotal)>0
		nLinha++
		lImpSoma := .T.	
		
		oCellCol := oCellDefault 
		oCellNumero := oCellNDefault			
			
		For nField := 1 To Len(_OBROWSE:AFIELDS)
			nPOSSoma := ASCAN(aSomaTotal,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
				
			if (nPOSSoma>0)
				oCellNumero:SetFormat(5)
				oExcel:AddCell(nLinha,nField, aSomaTotal[nPOSSoma,2], oCellNumero)	
			else
				oCellCol:SetFormat(0)
				oExcel:AddCell(nLinha,nField,if(lImpSoma,'TOTAL GERAL--->',''), oCellCol)
				lImpSoma := .F.					
			endif
					
		Next nField		
	endif
	
	(cArq)->(DbGoTop())
	
	//Monta Excel
	oExcel:Make()
		
	//apresenta a planilha gerada
	oExcel:OpenXML()
	oExcel:Destroy()
return

/*/
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦Programa  ¦ ExpMGExc ¦ Autor ¦ Wagner Cabrera        ¦ Data¦18/05/2016 ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Descricao ¦ Rotina de Exportação para Excel a Partir do MsNewGetDados  ¦¦¦
¦¦¦          ¦ Como Utilizar:				                              ¦¦¦
¦¦¦          ¦ U_ExpFWExc(MsNewGetDados,'COLUNA_DE_QUEBRA','TITULO')	  ¦¦¦
¦¦¦          ¦								                              ¦¦¦
¦¦¦          ¦ As linhas serão pintadas com as cores da Legenda do Browse ¦¦¦
¦¦¦          ¦								                              ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Uso       ¦ GERAL                                                      ¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
/*/

user function ExpMGExc(_OBROWSE, cQUEBRASOMA, _cTitulo) 
	PROCESSA({|| ExpMGExc(_OBROWSE, cQUEBRASOMA, _cTitulo)})
return

static function ExpMGExc(_OBROWSE, cQUEBRASOMA, _cTitulo)
	Local oExcel := AppExcel():NEW()
	Local oCellCab := AppExcCell():New()
	Local oCellTitulo := AppExcCell():New()
	Local oCellTit2 := AppExcCell():New()
	
	Local oCellCol := AppExcCell():New()
	Local oCellData := AppExcCell():New()
	Local oCellNumero := AppExcCell():New()
	
	Local oCellDefault := AppExcCell():New()
	Local oCellNDefault := AppExcCell():New()
	Local oCellDDefault := AppExcCell():New()	
	
	//Configuração das fontes
	Local oFont := AppExcFont():New("Arial","10","#000000") 
	Local oFontCab := AppExcFont():New("Arial","12","#000000") 
	Local oFontTitulo := AppExcFont():New("Cambria","18","#363636") 
	Local oFontBranco := AppExcFont():New("Arial","10","#FFFFFF") 
	
	Local nLinha := 1
	Local nPosTitulo := 1
	Local nPosField := 2
	Local nPosTipo := 8
	Local aCores := {}
	Local aSoma := {}
	Local aSomaTotal := {}
	Local CULTIMAQUEBRA := ''
	
	//Seta Negrito
	oFontTitulo:SetBold(.T.)	
	
	//personalização das celulas
	oCellCab:SetCellColor("#87CEFA")
	oCellCab:SetABorders(.T.)       
	oCellCab:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellCab:SetFont(oFontCab)
	
	oCellTitulo:SetVertAlign( VERTICAL_ALIGN_CENTER ) //ALINHA NO CENTRO
	oCellTitulo:SetHorzAlign( HORIZONTAL_ALIGN_CENTER ) //ALINHA NO CENTRO
	oCellTitulo:SetFont(oFontTitulo)
	
	oCellTit2:SetFont(oFont)		
	
	oCellDefault:SetABorders(.T.)
	oCellDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDefault:SetFont(oFont)
	
	oCellNDefault:SetABorders(.T.)
	oCellNDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellNDefault:SetFont(oFont)
	
	oCellDDefault:SetABorders(.T.)
	oCellDDefault:SetALineBorders(BORDER_LINE_CONTINUOUS)
	oCellDDefault:SetFont(oFont)		
	

	//Titulo do Relatório
	oExcel:Merge(nLinha,1,len(_OBROWSE:AHEADER)-1,0,_cTitulo, oCellTitulo)
	nLinha++
	oExcel:Merge(nLinha,1,len(_OBROWSE:AHEADER)-1,0,'Data do Relatório: ' + DTOC(DATE()) + ' ' + Time() + ' - Usuário: ' + cUserName, oCellTit2)
	nLinha++
	
	//Cabeçalho
	nLinha++
	For nField := 1 To Len(_OBROWSE:AHEADER)
		oExcel:AddCell(nLinha,nField, _OBROWSE:AHEADER[nField][nPosTitulo], oCellCab)
	Next nField
	
	//(cArq)->(DbGoTop())
	
	//while (cArq)->(!EOF())
	for nCols := 1 to Len(_OBROWSE:ACOLS)
		nLinha++
		
		oCellCol := oCellDefault 
		oCellData := oCellDDefault 	 
		oCellNumero := oCellNDefault
			
		For nField := 1 To (Len(_OBROWSE:AHEADER))
			if (_OBROWSE:AHEADER[nField][nPosTipo]='D')
				oCellData:SetFormat(4)
			
				oExcel:AddCell(nLinha,nField, If(!Empty(_OBROWSE:ACOLS[nCols][nField]),_OBROWSE:ACOLS[nCols][nField],""), oCellData)
			else 
				if (_OBROWSE:AHEADER[nField][nPosTipo]='N')
					oCellNumero:SetFormat(5)
					oExcel:AddCell(nLinha,nField,_OBROWSE:ACOLS[nCols][nField], oCellNumero)
					
					//SOMA DAS COLUNAS VALORES
					/*nPOSSoma := ASCAN(aSoma,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
					if (nPOSSoma>0)
						aSoma[nPOSSoma, 2] += (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])
					else 
						AADD(aSoma, {_OBROWSE:AFIELDS[nField][nPosField], (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])})
					endif
					
					nPOSSoma := ASCAN(aSomaTotal,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
					if (nPOSSoma>0)
						aSomaTotal[nPOSSoma, 2] += (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])	
					else 
						AADD(aSomaTotal, {_OBROWSE:AFIELDS[nField][nPosField], (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])})
					endif*/					
				else
					oCellCol:SetFormat(0)
					oExcel:AddCell(nLinha,nField,_OBROWSE:ACOLS[nCols][nField], oCellCol)
				endif
			endif
			
			/*if (_OBROWSE:AFIELDS[nField][nPosField]=cQUEBRASOMA)
				if !Empty((cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])) 
					cULTIMAQUEBRA := (cArq)->&(_OBROWSE:AFIELDS[nField][nPosField])
				endif	
			endif*/
		Next nField
	
		/*(cArq)->(dbskip())
		
		if ((cArq)->(EOF()) .or. (cULTIMAQUEBRA<>(cArq)->&(cQUEBRASOMA))) .and. (len(aSoma)>0) .and. !Empty(cQUEBRASOMA)
			lImpSoma := .T.
			nLinha++
			
			oCellCol := oCellDefault 
			oCellNumero := oCellNDefault			
			
			For nField := 1 To Len(_OBROWSE:AFIELDS)
				nPOSSoma := ASCAN(aSoma,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
				
				if (nPOSSoma>0)
					oCellNumero:SetFormat(5)
					oExcel:AddCell(nLinha,nField, aSoma[nPOSSoma,2], oCellNumero)	
				else
					oCellCol:SetFormat(0)
					oExcel:AddCell(nLinha,nField,if(lImpSoma,'TOTALIZAÇÃO--->',''), oCellCol)
					lImpSoma := .F.					
				endif
						
			Next nField
			
			aSoma := {}		
		endif		*/		
	Next nCols
	//enddo
	
	/*if len(aSomaTotal)>0
		nLinha++
		lImpSoma := .T.	
		
		oCellCol := oCellDefault 
		oCellNumero := oCellNDefault			
			
		For nField := 1 To Len(_OBROWSE:AFIELDS)
			nPOSSoma := ASCAN(aSomaTotal,{|X| RTRIM(X[1]) == _OBROWSE:AFIELDS[nField][nPosField]})
				
			if (nPOSSoma>0)
				oCellNumero:SetFormat(5)
				oExcel:AddCell(nLinha,nField, aSomaTotal[nPOSSoma,2], oCellNumero)	
			else
				oCellCol:SetFormat(0)
				oExcel:AddCell(nLinha,nField,if(lImpSoma,'TOTAL GERAL--->',''), oCellCol)
				lImpSoma := .F.					
			endif
					
		Next nField		
	endif*/
	
	//(cArq)->(DbGoTop())
	
	//Monta Excel
	oExcel:Make()
		
	//apresenta a planilha gerada
	oExcel:OpenXML()
	oExcel:Destroy()
return
