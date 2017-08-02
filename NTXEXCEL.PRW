#include "protheus.ch"
#include "topconn.ch"
#Include "AppExcel.ch"

/*/
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦Programa  ¦ NTXEXCEL  ¦ Autor ¦ Wagner Cabrera       ¦ Data¦16/02/2016 ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Descricao ¦ Exportador padrão de excel por SQL						  ¦¦¦
¦¦¦          ¦					                               			  ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Uso       ¦		                                                      ¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
/*/
User Function NTXEXCEL(aVetTitulo, aVetDetalhe, aVetCPOTit, aVetCPODet, cQuery, cAba, cTituloRel, cCampAgr, aVetParam, cNomeArq, cLocalArq)
	PROCESSA({|| EXPEXCEL(aVetTitulo, aVetDetalhe, aVetCPOTit, aVetCPODet, cQuery, cAba, cTituloRel, cCampAgr, aVetParam, cNomeArq, cLocalArq)})
Return	

//1 - Configurações das colunas Titulo/2 - Configurações das colunas Detalhe/3 - Titulos Coluna/4 - Campos da SX3 para detalhes/5 - SQL/7 - Nome Aba/8 - Titulo Relatorio/9 - Campo que agrupa  
Static Function EXPEXCEL(aVetTitulo, aVetDetalhe, aVetCPOTit, aVetCPODet, cQuery, cAba, cTituloRel, cCampAgr, aVetParam, cNomeArq, cLocalArq)
	Local nCont			:= 0
	Local nLinha		:= 1
	Local nPosI			:= 0
	Local I				:= 0
	Local cVlrAgr		:= ""
	Local aVetTotais	:= {}
	Local aVetTotParc	:= {}
	Local oExcel 		:= nil
	Local oCellTitulo	:= nil
	Local oCellTotal	:= nil
	Local oCellDetalheC	:= nil
	Local oCellDetalheN	:= nil
	Local oCellDetalheD	:= nil
	Local oCellAux		:= nil
	Local oCellNoxtox	:= nil
	Local oCellTit2 	:= AppExcCell():New()
	Local oCellSubTot	:= AppExcCell():New()
	Local oCellDefault 	:= AppExcCell():New()
	Local oFontNoxtox	:= nil
	Local oFontTitulo	:= nil
	Local oFontDetalhe	:= nil
	Local oFontTotal	:= nil
	Local oFontAux		:= nil
	Local oFont 		:= AppExcFont():New("Arial","10","#000000")
			
	TCQUERY cQuery NEW Alias 'TABTMP'
	//dbUseArea(.T., 'TOPCONN', TCGenQry(,,cQuery), 'TABTMP', .F., .T.)		
	DBSELECTAREA('TABTMP')
	
	IF !TABTMP->(EOF())
		oExcel 			:= AppExcel():NEW(cAba)
		oCellTitulo 	:= AppExcCell():NEW()
		oCellDetalheC	:= AppExcCell():NEW()
		oCellDetalheN	:= AppExcCell():NEW()
		oCellDetalheD	:= AppExcCell():NEW()
		oCellTotal		:= AppExcCell():NEW()
		oCellNoxtox		:= AppExcCell():NEW()
		
		if !(cNomeArq = NIL) 
			if !Empty(cNomeArq)
				oExcel:setFileName(cNomeArq)	
			endif 
		endif
		
		if !(cLocalArq = NIL) 
			if !Empty(cLocalArq)
				oExcel:setDestPath(cLocalArq)	
			endif 
		else
			cLocalArq := ''
		endif		
		
		//1 - COR FONTE/2 - TIPO/3 - TAM/4 - NEGRITO/5 - SUBLINHADO/6 - ITALICO/7 - TIPO LINHA/8 - COR BORDA/9 - MOSTRAR BORDA/10 - CAMPO TOTALIZA
		
		//Configuração das fontes
		oFontTitulo	:= AppExcFont():New(aVetTitulo[2],aVetTitulo[3],aVetTitulo[1])
		oFontTitulo:SetBold(aVetTitulo[4])
		oFontTitulo:SetUnderLine(aVetTitulo[5])
		oFontTitulo:SetItalic(aVetTitulo[6])
		//----------------------------------------------------------------------------
		oFontDetalhe := AppExcFont():New(aVetDetalhe[2],aVetDetalhe[3],aVetDetalhe[1])
		oFontDetalhe:SetBold(aVetDetalhe[4])
		oFontDetalhe:SetUnderLine(aVetDetalhe[5])
		oFontDetalhe:SetItalic(aVetDetalhe[6]) 	 
		//----------------------------------------------------------------------------
		oFontNoxtox := AppExcFont():New("Arial","13","#0000FF")
		oFontNoxtox:SetBold(.T.)
		
		//----------------------------------------------------------------------------	
		//personalização das celulas
		oCellTitulo:SetCellColor(aVetTitulo[8])
		oCellTitulo:SetABorders(aVetTitulo[9])
		IF aVetTitulo[7] <> NIL
			oCellTitulo:SetALineBorders(aVetTitulo[7])
		ENDIF	
		oCellTitulo:SetFont(oFontTitulo)
		//----------------------------------------------------------------------------
		oCellDetalheC:SetCellColor(aVetDetalhe[8])
		oCellDetalheC:SetABorders(aVetDetalhe[9])
		IF aVetDetalhe[7] <> NIL
			oCellDetalheC:SetALineBorders(aVetDetalhe[7])
		ENDIF	
		oCellDetalheC:SetFont(oFontDetalhe)
		//----------------------------------------------------------------------------
		oCellDetalheD:SetFormat(4)
		oCellDetalheD:SetCellColor(aVetDetalhe[8])
		oCellDetalheD:SetABorders(aVetDetalhe[9])
		IF aVetDetalhe[7] <> NIL
			oCellDetalheD:SetALineBorders(aVetDetalhe[7])
		ENDIF	
		oCellDetalheD:SetFont(oFontDetalhe)
		//----------------------------------------------------------------------------
		oCellDetalheN:SetFormat(5)
		oCellDetalheN:SetCellColor(aVetDetalhe[8])
		oCellDetalheN:SetABorders(aVetDetalhe[9])
		IF aVetDetalhe[7] <> NIL
			oCellDetalheN:SetALineBorders(aVetDetalhe[7])
		ENDIF	
		oCellDetalheN:SetFont(oFontDetalhe)
		
		oCellDefault:SetCellColor(aVetDetalhe[8])
		oCellDefault:SetABorders(aVetDetalhe[9])
		IF aVetDetalhe[7] <> NIL
			oCellDefault:SetALineBorders(aVetDetalhe[7])
		ENDIF	
		oCellDefault:SetFont(oFontDetalhe)
		//----------------------------------------------------------------------------
		oCellNoxtox:SetCellColor("#DCDCDC")
		oCellNoxtox:SetABorders(.T.)
		oCellNoxtox:SetHorzAlign( HORIZONTAL_ALIGN_CENTER ) //ALINHA NO CENTRO
		oCellNoxtox:SetALineBorders(2)
		oCellNoxtox:SetFont(oFontNoxtox)
		
		oCellTit2:SetFont(oFont)
		
		oCellSubTot:SetFormat(5)
		oCellSubTot:SetABorders(.T.)
		oCellSubTot:SetALineBorders(BORDER_LINE_CONTINUOUS)
		oCellSubTot:SetFont(oFontTitulo)
		//----------------------------------------------------------------------------
		//Titulo do Relatório
		oExcel:Merge(nLinha,1,LEN(aVetCPOTit)-1,0, cTituloRel, oCellNoxtox)
		++nLinha
		oExcel:Merge(nLinha,1,LEN(aVetCPOTit)-1,0,'Data do Relatório: ' + DTOC(DATE()) + ' ' + Time() + ' - Usuário: ' + cUserName, oCellTit2)
		++nLinha		
		
		FOR nCont := 1 TO LEN(aVetCPOTit)
			IF VALTYPE(aVetCPOTit[nCont]) == "A"
				oFontAux	:= nil
				oFontAux	:= AppExcFont():New(aVetCPOTit[nCont, 2, 2],aVetCPOTit[nCont, 2, 3],aVetCPOTit[nCont, 2, 1])
				oFontAux:SetBold(aVetCPOTit[nCont, 2, 4])
				oFontAux:SetUnderLine(aVetCPOTit[nCont, 2, 5])
				oFontAux:SetItalic(aVetCPOTit[nCont, 2, 6]) 	 
				
				oCellAux := nil
				oCellAux := AppExcCell():NEW()
				oCellAux:SetCellColor(aVetCPOTit[nCont, 2, 8])
				oCellAux:SetABorders(aVetCPOTit[nCont, 2, 9])
				IF aVetCPOTit[nCont, 2, 7] <> NIL
					oCellAux:SetALineBorders(aVetCPOTit[nCont, 2, 7])
				ENDIF	
				oCellAux:SetFont(oFontAux)
				oExcel:AddCell(nLinha, nCont, aVetCPOTit[nCont, 1], oCellAux)
			ELSE				
				oExcel:AddCell(nLinha, nCont, aVetCPOTit[nCont], oCellTitulo)
			ENDIF	
		NEXT aVetCPOTit
		
		IF cCampAgr <> "" 
			cVlrAgr := TABTMP->&(cCampAgr)
		ENDIF	
		
		WHILE !TABTMP->(EOF())
			oCellDetalheC := oCellDefault
			oCellDetalheD := oCellDefault
			oCellDetalheN := oCellDefault
			
			++nLinha
			IF cVlrAgr <> TABTMP->&(cCampAgr) .AND. cCampAgr <> ""
				oExcel:AddCell(nLinha, 1, "Sub-Total:", oCellSubTot)
				oExcel:AddCell(nLinha, 2, cVlrAgr, oCellSubTot)	
				nPosI := 3
				FOR I := 1 TO LEN(aVetTotParc)
					FOR Y := nPosI TO LEN(aVetCPODet)
						IF aVetTotParc[I, 1] == nPosI
							oExcel:AddCell(nLinha, aVetTotParc[I, 1], aVetTotParc[I, 2], oCellSubTot)
							++nPosI
							EXIT
						ELSE
							oExcel:AddCell(nLinha, nPosI, " ", oCellSubTot)
							++nPosI
						ENDIF	
					NEXT //aScan(aVetTotParc, {|x| AllTrim(x[1]) == nPosI})
				NEXT
				aVetTotParc := {}   	
				++nLinha
			ENDIF

			FOR nCont := 1 TO LEN(aVetCPODet)
				IF VALTYPE(aVetCPODet[nCont]) == "A"
					oFontAux	:= nil
					oFontAux	:= AppExcFont():New(aVetCPODet[nCont, 2, 2],aVetCPODet[nCont, 2, 3],aVetCPODet[nCont, 2, 1])
					oFontAux:SetBold(aVetCPODet[nCont, 2, 4])
					oFontAux:SetUnderLine(aVetCPODet[nCont, 2, 5])
					oFontAux:SetItalic(aVetCPODet[nCont, 2, 6]) 	 
					
					oCellAux := nil
					oCellAux := AppExcCell():NEW()
					oCellAux:SetCellColor(aVetCPODet[nCont, 2, 8])
					oCellAux:SetABorders(aVetCPODet[nCont, 2, 9])
					IF aVetCPODet[nCont, 2, 7] <> NIL
						oCellAux:SetALineBorders(aVetCPODet[nCont, 2, 7])
					ENDIF	
					oCellAux:SetFont(oFontAux)
					
					//dbSelectArea("SX3")
					//dbsetorder(2)
					//IF dbseek(aVetCPODet[nCont, 1])
						IF VALTYPE(TABTMP->&(aVetCPODet[nCont, 1])) == "C"//X3_TIPO == "C"
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont, 1]), oCellAux)
						ELSEIF VALTYPE(TABTMP->&(aVetCPODet[nCont, 1])) == "N"//X3_TIPO == "N"
							oCellAux:SetFormat(5)
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont, 1]), oCellAux)
							
							nPosI := 0
							FOR I := 1 TO LEN(aVetTotais)
								IF aVetTotais[I, 1] == nCont
									nPosI := I
								ENDIF
							NEXT   	
							
							IF nPosI > 0 
								aVetTotais[nPosI, 2] += TABTMP->&(aVetCPODet[nCont, 1]) 	
							ELSE
								AADD(aVetTotais, {nCont, TABTMP->&(aVetCPODet[nCont, 1])})
							ENDIF
							
							nPosI := 0
							FOR I := 1 TO LEN(aVetTotParc)
								IF aVetTotParc[I, 1] == nCont
									nPosI := I
								ENDIF
							NEXT   	
							
							IF nPosI > 0 
								aVetTotParc[nPosI, 2] += TABTMP->&(aVetCPODet[nCont, 1]) 	
							ELSE
								AADD(aVetTotParc, {nCont, TABTMP->&(aVetCPODet[nCont, 1])})
							ENDIF										
											
						ELSEIF VALTYPE(TABTMP->&(aVetCPODet[nCont, 1])) == "D"//X3_TIPO == "D"//DATA TEM QUE CONVERTER NO SQL
							oCellAux:SetFormat(4)
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont, 1]), oCellAux)
						ENDIF
					//ELSE
						//oExcel:AddCell(nLinha, nCont, "CAMPO NÃO ENCONTRADO SX3", oCellAux)		
					//ENDIF
				ELSE
					//dbSelectArea("SX3")
					//dbsetorder(2)
					//IF dbseek(aVetCPODet[nCont])
						IF VALTYPE(TABTMP->&(aVetCPODet[nCont])) == "C"//X3_TIPO == "C"
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont]), oCellDetalheC)
						ELSEIF VALTYPE(TABTMP->&(aVetCPODet[nCont])) == "N"//X3_TIPO == "N"
							oCellDetalheN:SetFormat(5)
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont]), oCellDetalheN)
							
							nPosI := 0
							FOR I := 1 TO LEN(aVetTotais)
								IF aVetTotais[I, 1] == nCont
									nPosI := I
								ENDIF
							NEXT   	
							
							IF nPosI > 0 
								aVetTotais[nPosI, 2] += TABTMP->&(aVetCPODet[nCont]) 	
							ELSE
								AADD(aVetTotais, {nCont, TABTMP->&(aVetCPODet[nCont])})
							ENDIF
							
							nPosI := 0
							FOR I := 1 TO LEN(aVetTotParc)
								IF aVetTotParc[I, 1] == nCont
									nPosI := I
								ENDIF
							NEXT   	
							
							IF nPosI > 0 
								aVetTotParc[nPosI, 2] += TABTMP->&(aVetCPODet[nCont]) 	
							ELSE
								AADD(aVetTotParc, {nCont, TABTMP->&(aVetCPODet[nCont])})
							ENDIF										
								
						ELSEIF VALTYPE(TABTMP->&(aVetCPODet[nCont])) == "D"//X3_TIPO == "D"//DATA TEM QUE CONVERTER NO SQL
							oCellDetalheD:SetFormat(4)
							oExcel:AddCell(nLinha, nCont, TABTMP->&(aVetCPODet[nCont]), oCellDetalheD)
						ENDIF
					//ELSE
						//oExcel:AddCell(nLinha, nCont, "CAMPO NÃO ENCONTRADO SX3", oCellDetalheC)		
					//ENDIF	
				ENDIF	
			NEXT aVetCPODet
			IF cCampAgr <> ""
				cVlrAgr := TABTMP->&(cCampAgr)
			ENDIF			
			TABTMP->(DBSKIP())	
		ENDDO
		
		IF cCampAgr <> ""
			++nLinha
			oExcel:AddCell(nLinha, 1, "Sub-Total:", oCellSubTot)
			oExcel:AddCell(nLinha, 2, cVlrAgr     , oCellSubTot)	
				
			nPosI := 3
			FOR I := 1 TO LEN(aVetTotParc)
				FOR Y := nPosI TO LEN(aVetCPODet)
					IF aVetTotParc[I, 1] == nPosI
						oExcel:AddCell(nLinha, aVetTotParc[I, 1], aVetTotParc[I, 2], oCellSubTot)
						++nPosI
						EXIT
					ELSE
						oExcel:AddCell(nLinha, nPosI, " ", oCellSubTot)
						++nPosI
					ENDIF	
				NEXT 
			NEXT			
			aVetTotParc := {}   	
		ENDIF		
		
		++nLinha
		oExcel:AddCell(nLinha, 1, "Total Geral: ", oCellSubTot)
		/*FOR I := 1 TO LEN(aVetTotais)
			oExcel:AddCell(nLinha, aVetTotais[I, 1], aVetTotais[I, 2], oCellSubTot)
		NEXT*/
		
		nPosI := 2
		FOR I := 1 TO LEN(aVetTotais)
			FOR Y := nPosI TO LEN(aVetCPODet)
				IF aVetTotais[I, 1] == nPosI
					oExcel:AddCell(nLinha, aVetTotais[I, 1], aVetTotais[I, 2], oCellSubTot)
					++nPosI
					EXIT
				ELSE
					oExcel:AddCell(nLinha, nPosI, " ", oCellSubTot)
					++nPosI
				ENDIF	
			NEXT 
		NEXT			
		TABTMP->(DBCLOSEAREA())
		
		//PARAMETROS
		IF !aVetParam == NIL
			oExcel:AddSheet("Parametros")
			nLinha := 1
			oExcel:Merge(nLinha,1,2,0, "Parametros", oCellNoxtox)
			FOR I := 1 TO LEN(aVetParam)
				++nLinha
				oExcel:AddCell(nLinha, 1, aVetParam[I, 1], oCellSubTot)
				oExcel:AddCell(nLinha, 2, aVetParam[I, 2], oCellSubTot)
				oExcel:AddCell(nLinha, 3, aVetParam[I, 3], oCellDefault)
			NEXT				
		ENDIF	
		
		oExcel:Make()
		
		//apresenta a planilha gerada
		if Empty(cLocalArq)
			oExcel:OpenXML()
		endif
		           
		oExcel:Destroy()
		
	ELSE
		TABTMP->(DBCLOSEAREA())
		if Empty(cLocalArq)
			ALERT("Não há dados para o relatório.")
		endif
	ENDIF
Return