/*/
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦Programa  ¦ RELFINCLI ¦ Autor ¦ Wagner Cabrera       ¦ Data¦10/02/2016 ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Descricao ¦ Relatório de Dividas de Clientes 						  ¦¦¦
¦¦¦          ¦								                              ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Uso       ¦ SIGAFIN                                                    ¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
/*/

#include "protheus.ch"
#include "topconn.ch"

user function RELFINCLI
	local aCampos := {}
	local aCamposArq := {}
	local cQuery := ''
	local aSeek := {}
	local aValSaldo := {}
	local aValFat := {}
	local nTOT_FATUR := 0
	local nTOT_PROJET := 0
	local nTOT_REAIS := 0
	local nTOT_DOLAR := 0
	local nTOT_PAGTO := 0
	local nTOT_PAGT2 := 0
	local nTOT_CORRE := 0
	local nTOT_DESCO := 0
	local nTOT_JUROS := 0
	local nTOT_MULTA := 0
	local cChave := ''
	local cChaveCli := ''
	
	Private aRotina := {}
	Private CCADASTRO := 'Dividas de Clientes '
	Private oBrowse
	Private aColunas := {}

	if PERGUNTE(VALIDPERG(),.t.)
		if MV_PAR17 == 1
			CCADASTRO += '- ABERTOS'
		else
			CCADASTRO += '- RECEBIDOS'
		endif
		
		AADD(aRotina,{'Legenda',    		  "U_LEGDIV", 0, 3})
		AADD(aRotina,{'Exportar P/Excel',     "U_EXPDIV", 0, 3})
	
		aCampos := {}
		
		
		AADD(aCampos, "RTV")
		if MV_PAR25 == 2
			AADD(aCampos, "VNDRESP")			
		endif
		AADD(aCampos, "E1_UREGIAO")
		AADD(aCampos, "E1_CLIENTE")
		AADD(aCampos, "E1_LOJA")
		AADD(aCampos, "E1_NOMCLI")
		AADD(aCampos, "A1_CGCSM")
		AADD(aCampos, "E1_UCDGRP")
		AADD(aCampos, "E1_MSFIL")
		AADD(aCampos, "E1_NUM")
		AADD(aCampos, "E1_PARCELA")
		AADD(aCampos, "E1_TIPO")
		AADD(aCampos, "E1_EMISSAO")
		AADD(aCampos, "E1_VENCORI")
		AADD(aCampos, "E1_VENCTO")
		AADD(aCampos, "BANCO")
		AADD(aCampos, "A6_NREDUZ")
		AADD(aCampos, "E1_MOEDA")
		AADD(aCampos, "VL_FATUR")
		AADD(aCampos, "E1_TXMOEDA")
		AADD(aCampos, "SD_REAIS")
		AADD(aCampos, "SD_DOLAR")
		AADD(aCampos, "SD_CORRIRE")
		AADD(aCampos, "SD_CORRIDO")
		AADD(aCampos, "E1_HIST")
		AADD(aCampos, "STATUSLIB")
		AADD(aCampos, "DIVRS")
		AADD(aCampos, "NOMEADV")
		AADD(aCampos, "SITCART")
		  			
		if MV_PAR17 == 2
			AADD(aCampos, "DTPAGAMENT")
			AADD(aCampos, "DTBANCO")
			AADD(aCampos, "VLPAGAMENT")
			AADD(aCampos, "MOEDA2")
			AADD(aCampos, "VLCORRECAO")
			AADD(aCampos, "VLDESC")
			AADD(aCampos, "VLJUROS")
			AADD(aCampos, "VLMULTA")
			AADD(aCampos, "HISPGTO")
		endif
		
		AADD(aCampos, "SITUACAO")
		AADD(aCampos, "PENDDUP")
		AADD(aCampos, "NTXOBS")
		
		//if MV_PAR27 ==2
		AADD(aCampos, "A1_UDTVCHC")
		AADD(aCampos, "A1_UVLHIPO")
		//endif
			        
		DBSELECTAREA("SX3")
		DBSETORDER(2)
		for I := 1 to len(aCampos)
			if SX3->(DbSeek(aCampos[I]))
				Aadd(aColunas,{X3_TITULO,X3_CAMPO,X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
				Aadd(aCamposArq,{X3_CAMPO, X3_TIPO, X3_TAMANHO,X3_DECIMAL})
			else
				if aCampos[I]="STATUSLIB"
					SX3->(DbSeek("X5_DESCRI"))
					Aadd(aColunas,{"Status Lib.","STATUSLIB",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"STATUSLIB",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
				
				if aCampos[I]="SITCART"
					SX3->(DbSeek("X5_DESCRI"))
					Aadd(aColunas,{"Sit.Carteira","SITCART",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SITCART",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif				
				
				if aCampos[I]="NTXOBS"
					Aadd(aColunas,{"Observação","NTXOBS",'C',,,2047,0})
					Aadd(aCamposArq,{"NTXOBS",'C',2047,0})
				endif	
				
				if aCampos[I]="A1_CGCSM"
					SX3->(DbSeek("A1_CGC"))
					Aadd(aColunas,{"CNPJ/CPF","A1_CGCSM",X3_TIPO,,,18,X3_DECIMAL})
					Aadd(aCamposArq,{"A1_CGCSM",X3_TIPO,18,X3_DECIMAL})
				endif								
							
				
				if aCampos[I]="BANCO"
					SX3->(DbSeek("E1_PORTADO"))
					Aadd(aColunas,{"Banco/Caixa","BANCO",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"BANCO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
				
				if aCampos[I]="NOMEADV"
					SX3->(DbSeek("Z18_NOME"))
					Aadd(aColunas,{"Advogado","NOMEADV",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"NOMEADV",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif				
				
				if aCampos[I]="VL_FATUR"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Faturado R$","VL_FATUR",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"VL_FATUR",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
				
				if aCampos[I]="SD_REAIS"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Saldo R$","SD_REAIS",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SD_REAIS",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
				
				if aCampos[I]="SD_DOLAR"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Saldo US$","SD_DOLAR",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SD_DOLAR",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
				
				if aCampos[I]="SD_PROJET"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Saldo R$ Proj US$ " + Transform(MV_PAR18,"@E 999.99"),"SD_PROJET",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SD_PROJET",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif
																								
				if aCampos[I]="RTV"
					SX3->(DbSeek("A3_NOME"))
					Aadd(aColunas,{"RTV","RTV",X3_TIPO,X3_PICTURE,,40,X3_DECIMAL})
					Aadd(aCamposArq,{"RTV",X3_TIPO,40,X3_DECIMAL})
				endif

				if aCampos[I]="VNDRESP"
					SX3->(DbSeek("A3_NOME"))
					Aadd(aColunas,{"Vend.Responsavel","VNDRESP",X3_TIPO,X3_PICTURE,,40,X3_DECIMAL})
					Aadd(aCamposArq,{"VNDRESP",X3_TIPO,40,X3_DECIMAL})
				endif
				
				if aCampos[I]="SITUACAO"
					Aadd(aColunas,{"Situação","SITUACAO",'C',,,7,0})
					Aadd(aCamposArq,{"SITUACAO",'C',7,0})
				endif
				
				if aCampos[I]="DIVRS"
					Aadd(aColunas,{"Divida em Reais","DIVRS",'C',,,3,0})
					Aadd(aCamposArq,{"DIVRS",'C',3,0})
				endif
				
				if aCampos[I]="SD_CORRIRE"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Valor Corrigido R$","SD_CORRIRE",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SD_CORRIRE",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif	
				
				if aCampos[I]="SD_CORRIDO"
					SX3->(DbSeek("E1_VALOR"))
					Aadd(aColunas,{"Valor Corrigido US$","SD_CORRIDO",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
					Aadd(aCamposArq,{"SD_CORRIDO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
				endif							
				
				if aCampos[I]="PENDDUP"
					Aadd(aColunas,{"Pend.Duplicata","PENDDUP",'C',,,18,0})
					Aadd(aCamposArq,{"PENDDUP",'C',18,0})
				endif
				
				if MV_PAR17 == 2
				
					if aCampos[I]="E5_BCO"
						SX3->(DbSeek("E5_BANCO"))
						Aadd(aColunas,{"Banco/Caixa","E5_BCO",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"E5_BCO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif					
				
					if aCampos[I]="DTPAGAMENT"
						SX3->(DbSeek("E5_DATA"))
						Aadd(aColunas,{"Data da Baixa","DTPAGAMENT",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"DTPAGAMENT",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="DTBANCO"
						SX3->(DbSeek("E5_DATA"))
						Aadd(aColunas,{"Data do Pgto","DTBANCO",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"DTBANCO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="VLPAGAMENT"
						SX3->(DbSeek("E5_VALOR"))
						Aadd(aColunas,{"Valor Pagto","VLPAGAMENT",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"VLPAGAMENT",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
				
					if aCampos[I]="HISPGTO"
						SX3->(DbSeek("E5_HISTOR"))
						Aadd(aColunas,{"Hist.Pagto","HISPGTO",X3_TIPO,X3_PICTURE,,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"HISPGTO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="MOEDA2"
						SX3->(DbSeek("E5_VLMOED2"))
						Aadd(aColunas,{"Valor Pagto US$","MOEDA2",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"MOEDA2",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="VLCORRECAO"
						SX3->(DbSeek("E5_VLCORRE"))
						Aadd(aColunas,{"Variação Cambial","VLCORRECAO",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"VLCORRECAO",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="VLDESC"
						SX3->(DbSeek("E5_VLDESCO"))
						Aadd(aColunas,{"Valor Desconto","VLDESC",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"VLDESC",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
					
					if aCampos[I]="VLJUROS"
						SX3->(DbSeek("E5_VLJUROS"))
						Aadd(aColunas,{"Valor Juros","VLJUROS",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"VLJUROS",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif

					if aCampos[I]="VLMULTA"
						SX3->(DbSeek("E5_VLMULTA"))
						Aadd(aColunas,{"Valor Multa","VLMULTA",X3_TIPO,X3_PICTURE,2,X3_TAMANHO,X3_DECIMAL})
						Aadd(aCamposArq,{"VLMULTA",X3_TIPO,X3_TAMANHO,X3_DECIMAL})
					endif
				endif
			endif
		next
			         	
		// Arquivo para Trabalho Temporario
		cArqTrab:= CriaTrab(aCamposArq)
		dbUseArea(.T.,,cArqTrab,'ATRB',.F.,.F.)
		
		cInd1 := Left( cArqTrab, 7 ) + "1"
		cInd2 := Left( cArqTrab, 7 ) + "2"
		cInd3 := Left( cArqTrab, 7 ) + "3"
	
		IndRegua( "ATRB", cInd1, "E1_NOMCLI+DTOS(E1_VENCTO)", , , "Criando índices (Prefixo + Ordem)...")
		IndRegua( "ATRB", cInd2, "E1_VENCTO", , , "Criando índices (Prefixo + Ordem)...")
		IndRegua( "ATRB", cInd3, "E1_UREGIAO+RTV+E1_NOMCLI+DTOS(E1_VENCTO)", , , "Criando índices (Prefixo + Ordem)...")
		
		dbClearIndex()
		// Agrega a lista dos índices da tabela (arquivo).
		dbSetIndex(cInd1 + OrdBagExt())
		dbSetIndex(cInd2 + OrdBagExt())
		dbSetIndex(cInd3 + OrdBagExt()) 
		AAdd(aSeek,{"Nome do Cliente",{{"","C",60,0,"Nome do Cliente","@!"}}})
		AAdd(aSeek,{"Vencimento",{{"","D",10,0,"Vencimento",""}}})
		AAdd(aSeek,{"Região de Venda + RTV",{{"","C",11,0,"Região de Venda + RTV","@!"}}})
		
		// Inicia Query
		cQuery += 'SELECT SE1.R_E_C_N_O_ RECNO,E1_CLIENTE, E1_LOJA, SA1.A1_NOME E1_NOMCLI, SA1.A1_GRPVEN E1_UCDGRP, E1_MSFIL, E1_NUM, E1_PARCELA, E1_TIPO, E1_EMISSAO, '
		cQuery += '		  E1_VENCORI, E1_PORTADO, E1_AGEDEP, E1_CONTA, E1_MOEDA, E1_VALOR, E1_VLCRUZ, E1_SALDO, E1_HIST, SA1.A1_CGC, '
		
		if MV_PAR27 == 1
			cQuery += 'E1_VENCREA E1_VENCTO,'
		else
			cQuery += 'E1_VENCTO,'
		endif
		
		cQuery += "		  SX5.X5_DESCRI STATUSLIB, E1_UREGIAO, "
		cQuery += "       case when (CAST(E1_VENCTO AS DATE)<=(GETDATE()) and (E1_SALDO>0)) then 'VENCIDO' ELSE case when E1_SALDO=0 then 'PAGO' else 'ABERTO' END END as SITUACAO, "
		cQuery += "       E1_TXMOEDA, ISNULL(CONVERT(VARCHAR(2047), CONVERT(VARBINARY(2047), E1_UOBS)),'') NTXOBS, SA1HIPO.A1_UDTVCHC, SA1HIPO.A1_UVLHIPO "
		
		if MV_PAR24 == 1
			cQuery += " ,SA3.A3_COD+'-'+SA3.A3_NOME RTV "       	
		else
			cQuery += " ,SA3VENDA.A3_COD+'-'+SA3VENDA.A3_NOME RTV "
		endif
		
		cQuery += " ,SA3RESP.A3_COD+'-'+SA3RESP.A3_NOME VNDRESP "
		
		cQuery += " ,Z18.Z18_NOME NOMEADV, CASE WHEN SE1.E1_UDIVRS='S' THEN 'SIM' ELSE 'NÃO' END DIVRS, SX5SIT.X5_DESCRI SITCART "
		
		if MV_PAR17 == 2
			cQuery += " ,SE5.E5_DATA DTPAGAMENT, SE5.E5_VALOR VLPAGAMENT, SE5.E5_HISTOR HISPGTO, SE5.E5_VLMOED2 MOEDA2, SE5.E5_VLCORRE VLCORRECAO, SE5.E5_VLDESCO VLDESC, SE5.E5_VLJUROS VLJUROS, SE5.E5_VLMULTA VLMULTA, SE5.E5_TXMOEDA, SE5.E5_BANCO, SE5.E5_DOCUMEN, SE5.E5_FORNADT, SE5.E5_LOJAADT"//, SE5RA.E5_DATA E5RA_DATA, SE5RA.E5_BANCO E5RA_BC"
		endif
		
		//VERIFICA PENDENCIA DE DUPLICATA
		cQuery += " , CASE WHEN (SELECT COUNT(*) FROM SE1010 SE1PEN "
		cQuery += "				 JOIN Z10010 on Z10_COD=SE1PEN.E1_USTDP "
		cQuery += "				 WHERE SE1.E1_CLIENTE=SE1PEN.E1_CLIENTE AND SE1.E1_LOJA=SE1PEN.E1_LOJA AND SE1PEN.E1_SALDO>0 AND Z10_BLOQ='S')>0 THEN "
		cQuery += "     CASE WHEN (CAST(E1_VENCTO AS DATE)<=(GETDATE()) and (E1_SALDO>0)) then 'HÁ PENDÊNCIAS/VENC' else 'HÁ PENDÊNCIAS' END else 'SEM PENDÊNCIAS' END AS PENDDUP "
		
		cQuery += " FROM "+RETSQLNAME("SE1") + " SE1 "
		cQuery += " JOIN "+RETSQLNAME("SA1") + " SA1 on SA1.A1_COD=SE1.E1_CLIENTE AND SA1.A1_LOJA=SE1.E1_LOJA AND SA1.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("SA1") + " SA1HIPO on SA1HIPO.A1_COD=SE1.E1_CLIENTE AND SA1HIPO.A1_UVLHIPO>0 AND SA1HIPO.D_E_L_E_T_=' '"

		if MV_PAR17 == 2
			cQuery += ' JOIN '+RETSQLNAME("SE5") + ' SE5 on SE5.E5_PREFIXO = SE1.E1_PREFIXO '
			cQuery += " 								AND SE5.E5_NUMERO = SE1.E1_NUM "
			cQuery += " 								AND SE5.E5_PARCELA = SE1.E1_PARCELA "
			cQuery += " 								AND SE5.E5_TIPO = SE1.E1_TIPO "
			cQuery += " 								AND SE5.E5_CLIFOR = SE1.E1_CLIENTE "
			cQuery += " 								AND SE5.E5_LOJA = SE1.E1_LOJA "
			cQuery += " 								AND SE5.E5_FILIAL = SE1.E1_FILIAL "
			cQuery += " 								AND SE5.E5_TIPODOC not in ('CM', 'JR', 'RA', 'DC', 'MT', 'ES', "+ if(MV_PAR15<>'NCC',"'NCC',", "")+" 'RA') " //Correção - Juros - RA - Desconto - Multa - Estorno
			cQuery += " 								AND SE5.E5_SITUACA <> 'C' "
			cQuery += " 								AND SE5.D_E_L_E_T_=' '"
		endif

		cQuery += " LEFT JOIN "+RETSQLNAME("SX5") + " SX5 on SX5.X5_TABELA='99' AND SX5.X5_FILIAL='"+cFilAnt+"' AND SX5.X5_CHAVE=E1_UAVAL AND SX5.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("SA3") + " SA3 on SE1.E1_UREGIAO=SA3.A3_REGIAO AND SA3.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("SA3") + " SA3RESP on SA3RESP.A3_COD=SA1.A1_UVADM AND SA3RESP.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("SA3") + " SA3VENDA on SA3VENDA.A3_COD=SE1.E1_URTV AND SA3VENDA.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("Z18") + " Z18 on Z18.Z18_COD=SE1.E1_UCDADV AND Z18.D_E_L_E_T_=' '"
		cQuery += " LEFT JOIN "+RETSQLNAME("SX5") + " SX5SIT on SX5SIT.X5_TABELA='07' AND SX5SIT.X5_FILIAL='"+cFilAnt+"' AND SX5SIT.X5_CHAVE=E1_SITUACA AND SX5SIT.D_E_L_E_T_=' '"
		
		//cQuery += " LEFT JOIN "+RETSQLNAME("SA6") + " SA6 on A6_COD=" + IF(MV_PAR17 == 2, "SE5.E5_BANCO","E1_PORTADO") + " AND SA6.D_E_L_E_T_=' '"
		cQuery += " WHERE SE1.D_E_L_E_T_=' ' "
		
		//Desconsiderar recebimentos antecipados e devolução
		cQuery += " AND SE1.E1_TIPO not in ('RA'"+ if(MV_PAR15<>'NCC',",'NCC'", "")+ IF(MV_PAR17 == 2, ",'CRR', 'VDR'", "") +")"
		
		if MV_PAR17 == 1
			cQuery += " AND E1_SALDO>0 "
		else
			if MV_PAR17 == 2
				//cQuery += " AND E1_SALDO=0 " Retirado para mostrar as baixas parciais
				
				//Colocado para não considerar os estornos de baixas
				cQuery += " AND SE5.E5_PREFIXO+SE5.E5_NUMERO+SE5.E5_PARCELA+SE5.E5_TIPO+SE5.E5_CLIFOR+SE5.E5_LOJA+SE5.E5_FILIAL+SE5.E5_SEQ not in (  "
				cQuery += "   SELECT SE5C.E5_PREFIXO+SE5C.E5_NUMERO+SE5C.E5_PARCELA+SE5C.E5_TIPO+SE5C.E5_CLIFOR+SE5C.E5_LOJA+SE5C.E5_FILIAL+SE5C.E5_SEQ "
				cQuery += "   FROM "+RETSQLNAME("SE5") + " SE5C "
				cQuery += "     WHERE SE5C.E5_PREFIXO=SE5.E5_PREFIXO AND SE5C.E5_NUMERO=SE5.E5_NUMERO AND SE5C.E5_PARCELA=SE5.E5_PARCELA "
				cQuery += "       AND SE5C.E5_TIPO=SE5.E5_TIPO AND SE5C.E5_CLIFOR=SE5.E5_CLIFOR AND SE5C.E5_LOJA=SE5.E5_LOJA "
				cQuery += "       AND SE5C.E5_FILIAL=SE5.E5_FILIAL AND SE5C.E5_SEQ=SE5.E5_SEQ AND SE5C.E5_TIPODOC='ES' "
				cQuery += "       AND SE5C.D_E_L_E_T_=' '"
				cQuery += " ) "
				//--------------------------------------------------
			endif
		endif
		
		if MV_PAR27 == 1
			cQuery += "AND SE1.E1_VENCREA >= '"+DTOS(MV_PAR01)+"' "
			cQuery += "AND SE1.E1_VENCREA <= '"+DTOS(MV_PAR02)+"' "
		else
			//If !Empty(MV_PAR01)
			cQuery += "AND SE1.E1_VENCTO >= '"+DTOS(MV_PAR01)+"' "
			//EndIf
			
			//If !Empty(MV_PAR02)
			cQuery += "AND SE1.E1_VENCTO <= '"+DTOS(MV_PAR02)+"' "
			//EndIf
		endif 
		
		
		if MV_PAR17 == 1
			//
		else
			if MV_PAR17 == 2
				//If !Empty(MV_PAR03)
				cQuery += "AND SE5.E5_DATA >= '"+DTOS(MV_PAR03)+"' "
				//EndIf
		
				//If !Empty(MV_PAR04)
				cQuery += "AND SE5.E5_DATA <= '"+DTOS(MV_PAR04)+"' "
				//EndIf
			Endif
		Endif
		
		//If !Empty(MV_PAR05)
		cQuery += "AND SA1.A1_GRPVEN >= '"+MV_PAR05+"' "
		//EndIf
		
		//If !Empty(MV_PAR06)
		cQuery += "AND SA1.A1_GRPVEN <= '"+MV_PAR06+"' "
		//EndIf
		
		//If !Empty(MV_PAR07)
		cQuery += "AND SE1.E1_CLIENTE >= '"+MV_PAR07+"' "
		//EndIf
		
		//If !Empty(MV_PAR08)
		cQuery += "AND SE1.E1_CLIENTE <= '"+MV_PAR08+"' "
		//EndIf	
		
		//If !Empty(MV_PAR09)
		cQuery += "AND SE1.E1_URTV >= '"+MV_PAR09+"' "
		//EndIf
		
		//If !Empty(MV_PAR10)
		cQuery += "AND SE1.E1_URTV <= '"+MV_PAR10+"' "
		//EndIf	
		
		//If !Empty(MV_PAR11)
		cQuery += "AND SE1.E1_UREGIAO >= '"+MV_PAR11+"' "
		//EndIf
		
		//If !Empty(MV_PAR12)
		cQuery += "AND SE1.E1_UREGIAO <= '"+MV_PAR12+"' "
		//EndIf
		
		//If !Empty(MV_PAR13)
		cQuery += "AND SE1.E1_PREFIXO >= '"+MV_PAR13+"' "
		//EndIf
		
		//If !Empty(MV_PAR14)
		cQuery += "AND SE1.E1_PREFIXO <= '"+MV_PAR14+"' "
		//EndIf	
		
		//If !Empty(MV_PAR15)
		cQuery += "AND SE1.E1_TIPO >= '"+MV_PAR15+"' "
		//EndIf
		
		//If !Empty(MV_PAR16)
		cQuery += "AND SE1.E1_TIPO <= '"+MV_PAR16+"' "
		//EndIf		
		
		cQuery += "AND SE1.E1_EMISSAO >= '"+DTOS(MV_PAR19)+"' "
		cQuery += "AND SE1.E1_EMISSAO <= '"+DTOS(MV_PAR20)+"' "
		
		if !empty(MV_PAR21) .and. str(MV_PAR21)<>'0' 
			cQuery += "AND SE1.E1_MOEDA = "+alltrim(str(MV_PAR21))+" "
		endif
		
		if !empty(MV_PAR22) 
			cQuery += "AND SE1.E1_UCDADV = '"+alltrim(MV_PAR22)+"' "
		endif	
		
		if MV_PAR23 == 2
			cQuery += "AND SE1.E1_MOTPROR = '"+'EST'+"' "			
		endif
		
		if MV_PAR23 == 3
			cQuery += "AND SE1.E1_MOTPROR = '"+'COB'+"' "			
		endif	
		
		if !empty(MV_PAR26)
			cQuery += "AND SE1.E1_SITUACA IN ( " + RetSituacao() + "') "  
		endif
		
		cQuery += "ORDER BY SE1.E1_CLIENTE "	
		
		TCQUERY cQuery NEW ALIAS 'TRB'
		TcSetField("TRB","E1_EMISSAO","D",08,0)
		TcSetField("TRB","E1_VENCTO","D",08,0)
		TcSetField("TRB","E1_VENCORI","D",08,0)
		
		if MV_PAR17 == 2
			TcSetField("TRB","DTPAGAMENT","D",08,0)
			//			
		Endif
		
		//if MV_PAR27 ==2
		TcSetField("TRB","A1_UDTVCHC","D",08,0)
		//endif
		
		TRB->(DBGOTOP())
		
		SM2->(DBSETORDER(1))
		SE1->(DBSETORDER(1))
		SA6->(DBSETORDER(1))
				
		if !TRB->(eof())
			while !TRB->(eof())
				reclock('ATRB',.T.)
				ATRB->RTV     	  := TRB->RTV
				if MV_PAR25 == 2
					ATRB->VNDRESP     := TRB->VNDRESP
				endif
				ATRB->E1_UREGIAO  := TRB->E1_UREGIAO
				ATRB->E1_CLIENTE  := TRB->E1_CLIENTE
				ATRB->E1_LOJA     := TRB->E1_LOJA
				ATRB->E1_NOMCLI   := TRB->E1_NOMCLI
				
				if len(alltrim(TRB->A1_CGC))<=11
					ATRB->A1_CGCSM:=  TRANSFORM(alltrim(TRB->A1_CGC),"@R 999.999.999-99")
				else
					ATRB->A1_CGCSM:= Transform(alltrim(TRB->A1_CGC), "@R 99.999.999/9999-99")
				endif
				
				ATRB->E1_UCDGRP   := TRB->E1_UCDGRP
				ATRB->E1_MSFIL    := TRB->E1_MSFIL
				ATRB->E1_NUM  	  := TRB->E1_NUM
				ATRB->E1_PARCELA  := TRB->E1_PARCELA
				ATRB->E1_TIPO     := TRB->E1_TIPO
				ATRB->E1_EMISSAO  := TRB->E1_EMISSAO
				ATRB->E1_VENCORI  := TRB->E1_VENCORI
				ATRB->E1_VENCTO   := TRB->E1_VENCTO
				ATRB->E1_MOEDA    := TRB->E1_MOEDA
				ATRB->E1_HIST     := TRB->E1_HIST
				ATRB->STATUSLIB   := TRB->STATUSLIB
				ATRB->SITUACAO    := TRB->SITUACAO
				ATRB->PENDDUP     := TRB->PENDDUP
				ATRB->NTXOBS      := TRB->NTXOBS
				ATRB->A6_NREDUZ   := ''
				ATRB->NOMEADV     := TRB->NOMEADV
				ATRB->DIVRS       := TRB->DIVRS
				ATRB->SITCART     := TRB->SITCART
				
				if cChaveCli<>ATRB->E1_CLIENTE
					ATRB->A1_UDTVCHC  := TRB->A1_UDTVCHC
					ATRB->A1_UVLHIPO  := TRB->A1_UVLHIPO	
					cChaveCli := ATRB->E1_CLIENTE
				endif
				
				//if MV_PAR27 ==2
				//ATRB->A1_UDTVCHC  := TRB->A1_UDTVCHC
				//ATRB->A1_UVLHIPO  := TRB->A1_UVLHIPO
				//endif 
				
				if MV_PAR17 == 2
					ATRB->DTPAGAMENT := TRB->DTPAGAMENT
					ATRB->DTBANCO 	 := TRB->DTPAGAMENT
					ATRB->VLPAGAMENT := TRB->VLPAGAMENT
					ATRB->HISPGTO    := TRB->HISPGTO
					ATRB->BANCO  	 := TRB->E5_BANCO
					
					if /*(ATRB->E1_MOEDA=1) .and.*/ (TRB->E5_TXMOEDA<=1)
						aValFat := FuncaMoeda(TRB->E1_VENCTO,TRB->MOEDA2,TRB->E1_MOEDA)
						ATRB->MOEDA2 := aValFat[4]
					else
						ATRB->MOEDA2 := TRB->MOEDA2
					endif
					
					ATRB->VLCORRECAO := TRB->VLCORRECAO
					ATRB->VLDESC 	 := TRB->VLDESC
					ATRB->VLJUROS 	 := TRB->VLJUROS
					ATRB->VLMULTA 	 := TRB->VLMULTA
					
					nTOT_PAGTO       += ATRB->VLPAGAMENT
					nTOT_PAGT2       += ATRB->MOEDA2
					nTOT_CORRE       += ATRB->VLCORRECAO
					nTOT_DESCO       += ATRB->VLDESC
					nTOT_JUROS       += ATRB->VLJUROS
					nTOT_MULTA       += ATRB->VLMULTA
					
					//Função para Encontrar Banco do RA
					
					if Empty(ATRB->BANCO) .and. ('RA'$TRB->E5_DOCUMEN)
						cQuery2 := "SELECT SE5RA.E5_BANCO E5RA_BC, SE5RA.E5_DATA E5RA_DATA "
						cQuery2 += " FROM "+RETSQLNAME("SE5") +" SE5RA " 
						cQuery2 += " WHERE SE5RA.D_E_L_E_T_=' '"
						cQuery2 += "  AND SE5RA.E5_PREFIXO='" + SUBSTR(TRB->E5_DOCUMEN,1,3) +"'"
						cQuery2 += "  AND SE5RA.E5_NUMERO='" + SUBSTR(TRB->E5_DOCUMEN,4,9) +"'"
						cQuery2 += "  AND SE5RA.E5_PARCELA='" + SUBSTR(TRB->E5_DOCUMEN,13,3) +"'"
						cQuery2 += "  AND SE5RA.E5_TIPO='" + SUBSTR(TRB->E5_DOCUMEN,16,3) +"'"
						cQuery2 += "  AND SE5RA.E5_CLIFOR='" + TRB->E5_FORNADT +"'"
						cQuery2 += "  AND SE5RA.E5_LOJA='" + TRB->E5_LOJAADT +"'"
						cQuery2 += "  AND SE5RA.E5_TIPODOC='" + 'RA' +"'"	
						
						TCQUERY cQuery2 NEW ALIAS 'TR2'
						TcSetField("TR2","E5RA_DATA","D",08,0)
						
						if !TR2->(Eof())
							ATRB->BANCO := TR2->E5RA_BC
							ATRB->DTBANCO := TR2->E5RA_DATA
						endif
						TR2->(DBCLOSEAREA())
						DBSELECTAREA('TRB')
					endif	
					
					if !Empty(ATRB->BANCO)
						if SA6->(DBSEEK(xFilial('SA6')+ATRB->BANCO))
							ATRB->A6_NREDUZ   := SA6->A6_NOME	
						endif
					else
						ATRB->BANCO := TRB->E1_PORTADO
						if !Empty(ATRB->BANCO)
							if SA6->(DBSEEK(xFilial('SA6')+ATRB->BANCO))
								ATRB->A6_NREDUZ   := SA6->A6_NOME	
							endif				
						endif						
					endif 
		
				else
					ATRB->BANCO := TRB->E1_PORTADO
					
					if !Empty(ATRB->BANCO)
						if SA6->(DBSEEK(xFilial('SA6')+ATRB->BANCO+TRB->E1_AGEDEP+TRB->E1_CONTA))
							ATRB->A6_NREDUZ   := SA6->A6_NOME	
						endif				
					endif	
				endif
				
				//Removido a Pedido de Usuário
				//aValFat := FuncaMoeda(TRB->E1_VENCTO,TRB->E1_VALOR,TRB->E1_MOEDA)
				//ATRB->VL_FATUR := aValFat[1]
				ATRB->VL_FATUR := TRB->E1_VLCRUZ
								
				//CALCULA VALOR DE MOEDAS
			
				ATRB->SD_REAIS := 0
				ATRB->SD_DOLAR := 0
				
				if (TRB->E1_TXMOEDA > 0)//(ATRB->SD_DOLAR=0) .or. (ATRB->SD_REAIS=0)
					//SE NÃO ENCONTRAR MOEDA NA DATA DO VENCIMENTO, ATUALIZA COM A MOEDA DA DATA BASE
					aValSaldo := FuncaMoeda(dDataBase,TRB->E1_SALDO,TRB->E1_MOEDA)	
					
					ATRB->E1_TXMOEDA := TRB->E1_TXMOEDA
				
					if (ATRB->SD_REAIS=0) 
						if (ATRB->E1_MOEDA=1)
							ATRB->SD_REAIS := TRB->E1_SALDO
						else
							ATRB->SD_REAIS := TRB->E1_SALDO * TRB->E1_TXMOEDA
						endif 
					endif
				
					if (ATRB->SD_DOLAR=0)
						if (ATRB->E1_MOEDA=1)
							ATRB->SD_DOLAR := aValSaldo[4]
						else
							ATRB->SD_DOLAR := TRB->E1_SALDO 
						endif	
					endif
				else
					//CALCULA VALOR DE MOEDAS
					aValSaldo := FuncaMoeda(TRB->E1_VENCTO,TRB->E1_SALDO,TRB->E1_MOEDA)		
					
					ATRB->SD_REAIS := aValSaldo[1]
					ATRB->SD_DOLAR := aValSaldo[4]	
					
					ATRB->E1_TXMOEDA := round(aValSaldo[1] / aValSaldo[4],4)	
												
				endif			
				
				SE1->(DBGOTO(TRB->RECNO))				

				ATRB->SD_CORRIRE := ATRB->SD_REAIS + fa070Juros(1)//REAIS
				ATRB->SD_CORRIDO := ATRB->SD_DOLAR + fa070Juros(4)//DOLLAR
				
				if MV_PAR18 == 2
					if ATRB->E1_MOEDA=1
						ATRB->SD_DOLAR := 0	
						ATRB->SD_CORRIDO := 0					
					else
						ATRB->SD_REAIS := 0
						ATRB->SD_CORRIRE := 0
					endif	
				endif
				
				if (cChave = (TRB->E1_NUM + TRB->E1_PARCELA + TRB->E1_TIPO)) .and. (MV_PAR17 == 2)
					ATRB->SD_REAIS := 0
					ATRB->SD_DOLAR := 0	
					ATRB->VL_FATUR := 0	
				endif 
				
				//nTOT_FATUR += ATRB->VL_FATUR
				//nTOT_PROJET += ATRB->SD_PROJET
				//nTOT_REAIS += ATRB->SD_REAIS
				//nTOT_DOLAR += ATRB->SD_DOLAR
				
				ATRB->(msunlock())
				
				cChave := TRB->E1_NUM + TRB->E1_PARCELA + TRB->E1_TIPO 
				
				TRB->(dbSkip())
			enddo
			
			//TOTALIZADORES 
			/*reclock('ATRB',.T.)
			ATRB->E1_NOMCLI   := 'ZZZZZZZZZ'
			ATRB->E1_VENCTO   := STOD('29991231')
			ATRB->RTV     := 'TOTAIS--------------------------------->'
			ATRB->VL_FATUR := nTOT_FATUR
			ATRB->SD_PROJET := nTOT_PROJET
			ATRB->SD_REAIS := nTOT_REAIS
			ATRB->SD_DOLAR := nTOT_DOLAR
				
			if MV_PAR17 == 2
				ATRB->VL_FATUR := 0
				ATRB->SD_PROJET := 0
				ATRB->SD_REAIS := 0
				ATRB->SD_DOLAR := 0
				ATRB->VLPAGAMENT := nTOT_PAGTO
				ATRB->MOEDA2 := nTOT_PAGT2
				ATRB->VLCORRECAO := nTOT_CORRE
				ATRB->VLDESC := nTOT_DESCO
				ATRB->VLJUROS := nTOT_JUROS
				ATRB->VLMULTA := nTOT_MULTA
			endif
					
			ATRB->(msunlock())*/
			//-------------
			
			DbSelectArea('ATRB')
			
			//Browse
			oBrowse:=FWMBrowse():New()
			oBrowse:SetUseCaseFilter(.T.)
			oBrowse:SetUseFilter(.T.)
			oBrowse:ForceQuitButton(.T.)
			oBrowse:SetWalkThru(.F.)
			oBrowse:SetProfileID( '1' )
			oBrowse:SetDescription(CCADASTRO)
			oBrowse:SetAlias('ATRB')
			oBrowse:SetFields(aColunas)
			oBrowse:SetSeek(.T.,aSeek)
			
			oBrowse:AddLegend("PENDDUP='"+'HÁ PENDÊNCIAS/VENC'+"'","BR_VERMELHO","Vencidos + Pendência Aceite/Aval")
			oBrowse:AddLegend("SITUACAO='"+'VENCIDO'+"'","BR_AZUL","Vencidos")
			oBrowse:AddLegend("PENDDUP='"+'HÁ PENDÊNCIAS'+"'","BR_LARANJA","Pendência Aceite/Aval")
			oBrowse:AddLegend("SITUACAO='"+'PAGO'+"'","BR_BRANCO","Pago")
			
			oBrowse:DisableDetails()
			oBrowse:Activate()
			
		else
			Alert('Não existe dados para serem demonstrados')
		endif
		
		TRB->(DBCLOSEAREA())
		ATRB->(DBCLOSEAREA())
		ERASE(cInd1 + OrdBagExt())
		ERASE(cInd2 + OrdBagExt())
		FErase(cArqTrab)
	endif
return


STATIC FUNCTION VALIDPERG
	LOCAL AREGS := {}
	LOCAL I,J
	LOCAL CPERG := PADR('RELFINCLI',10)
	
	DBSELECTAREA("SX1")
	DBSETORDER(1)
	
	// GRUPO/ORDEM/PERGUNTA/VARIAVEL/TIPO/TAMANHO/DECIMAL/PRESEL/GSC/VALID/VAR01/DEF01/CNT01/VAR02/DEF02/CNT02/VAR03/DEF03/CNT03/VAR04/DEF04/CNT04/VAR05/DEF05/CNT05
	AADD(AREGS,{CPERG,"01","Vencimento De       ?","","",  "MV_CH01","D",08,0,0,"G","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"02","Vencimento Ate      ?","","",  "MV_CH02","D",08,0,0,"G","","MV_PAR02","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"03","Pagamento De        ?","","",  "MV_CH03","D",08,0,0,"G","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"04","Pagamento Ate       ?","","",  "MV_CH04","D",08,0,0,"G","","MV_PAR04","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"05","Grupo Cliente De    ?","","",  "MV_CH05","C",06,0,0,"G","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","","","ACY",""})
	AADD(AREGS,{CPERG,"06","Grupo Cliente Ate   ?","","",  "MV_CH06","C",06,0,0,"G","","MV_PAR06","","","","","","","","","","","","","","","","","","","","","","","","","ACY",""})
	AADD(AREGS,{CPERG,"07","Cliente De          ?","","",  "MV_CH07","C",09,0,0,"G","","MV_PAR07","","","","","","","","","","","","","","","","","","","","","","","","","CLI",""})
	AADD(AREGS,{CPERG,"08","Cliente Ate         ?","","",  "MV_CH08","C",09,0,0,"G","","MV_PAR08","","","","","","","","","","","","","","","","","","","","","","","","","CLI",""})
	AADD(AREGS,{CPERG,"09","Vendedor De         ?","","",  "MV_CH09","C",06,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","","SA3",""})
	AADD(AREGS,{CPERG,"10","Vendedor Ate        ?","","",  "MV_CH10","C",06,0,0,"G","","MV_PAR10","","","","","","","","","","","","","","","","","","","","","","","","","SA3",""})
	AADD(AREGS,{CPERG,"11","Regiao de Venda De  ?","","",  "MV_CH11","C",05,0,0,"G","","MV_PAR11","","","","","","","","","","","","","","","","","","","","","","","","","Z02",""})
	AADD(AREGS,{CPERG,"12","Regiao de Venda Até ?","","",  "MV_CH12","C",05,0,0,"G","","MV_PAR12","","","","","","","","","","","","","","","","","","","","","","","","","Z02",""})
	AADD(AREGS,{CPERG,"13","Prefixo De          ?","","",  "MV_CH13","C",03,0,0,"G","","MV_PAR13","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"14","Prefixo Ate         ?","","",  "MV_CH14","C",03,0,0,"G","","MV_PAR14","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"15","Tipo De             ?","","",  "MV_CH15","C",03,0,0,"G","","MV_PAR15","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"16","Tipo Ate            ?","","",  "MV_CH16","C",03,0,0,"G","","MV_PAR16","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"17","Tipo do Relatório   ?","?","?","MV_CH17","C",18,0,0,"C","","MV_PAR17","Titulos Abertos","","","","","Titulos Recebidos","","","","",""/*Todos*/,"","","","","","","","","","","","","", "","","","","",""})
	AADD(AREGS,{CPERG,"18","Demonstr. de Moeda  ?","?","?","MV_CH18","C",18,0,0,"C","","MV_PAR18","Normal","","","","","Zerar Moeda Dif.","","","","",""/*Todos*/,"","","","","","","","","","","","","", "","","","","",""})
	AADD(AREGS,{CPERG,"19","Emissao De          ?","","",  "MV_CH19","D",08,0,0,"G","","MV_PAR19","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"20","Emissao Ate         ?","","",  "MV_CH20","D",08,0,0,"G","","MV_PAR20","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"21","Moeda               ?","","",  "MV_CH21","N",02,0,0,"G","","MV_PAR21","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	AADD(AREGS,{CPERG,"22","Advogado            ?","","",  "MV_CH22","C",04,0,0,"G","","MV_PAR22","","","","","","","","","","","","","","","","","","","","","","","","","Z18",""})
	AADD(AREGS,{CPERG,"23","Motivo Prorrogação  ?","?","?","MV_CH23","C",18,0,0,"C","","MV_PAR23","Todos","","","","","Estoque","","","","","Cobrança","","","","","","","","","","","","","", "","","","","",""})
	AADD(AREGS,{CPERG,"24","Vendedor            ?","?","?","MV_CH24","C",18,0,0,"C","","MV_PAR24","Atual","","","","","Da Venda","","","","",""/*Todos*/,"","","","","","","","","","","","","", "","","","","",""})
	AADD(AREGS,{CPERG,"25","Mostra Vend Respons ?","?","?","MV_CH25","C",18,0,0,"C","","MV_PAR25","Não","","","","","Sim","","","","",""/*Todos*/,"","","","","","","","","","","","","", "","","","","",""})
	AADD(AREGS,{CPERG,"26","Situação		    ?","?","?","MV_CH26","C",99,0,0,"C","U_REDIVSIT()","MV_PAR26","","","","","","","","","","","","","","","","","","","","","","","","", "",""})
	AADD(AREGS,{CPERG,"27","Demonstr. Vencimento?","?","?","MV_CH27","C",18,0,0,"C","","MV_PAR27","Real","","","","","Original","","","","",""/*Todos*/,"","","","","","","","","","","","","", "","","","","",""})		

	FOR I:=1 TO LEN(AREGS)
		IF !DBSEEK(CPERG+AREGS[I,2])
			RECLOCK("SX1",.T.)
			FOR J:=1 TO FCOUNT()
				IF J <= LEN(AREGS[I])
					FIELDPUT(J,AREGS[I,J])
				ENDIF
			NEXT
			MSUNLOCK()
		ENDIF
	NEXT
return(CPERG)

User Function LEGDIV()
	Private aLegenda:= {}
	aadd(aLegenda,{'BR_AZUL'    ,'Duplicatas Vencidas'})
	aadd(aLegenda,{'BR_LARANJA' ,'Pendência de Assinatura de Duplicata'})
	aadd(aLegenda,{'BR_VERMELHO','Duplicatas Vencidas + Pendência de Assinatura de Duplicata'})
	aadd(aLegenda,{'BR_BRANCO'  ,'Pago'})
	BrwLegenda('Legenda de Divida de Clientes','Legenda',aLegenda)
return

User Function EXPDIV()
	U_ExpFWExc(oBrowse, 'ATRB', 'E1_CLIENTE', CCADASTRO)	
return

user function REDIVSIT(l1Elem, lTipoRet)
	Local cTitulo	:= ""
	Local MvPar
	Local MvParDef	:= ""
	Local nCont		:= 0
	Private aSit	:= {}
	
	l1Elem := If (l1Elem = Nil , .F. , .T.)

	DEFAULT lTipoRet := .T.

	IF lTipoRet
		MvPar := &(Alltrim(ReadVar()))		 // Carrega Nome da Variavel do Get em Questao
		mvRet := Alltrim(ReadVar())			 // Iguala Nome da Variavel ao Nome variavel de Retorno
	EndIF

	//cQry := "SELECT * FROM " + RETSQLNAME("SX5") + " WHERE D_E_L_E_T_ = '' AND X5_TABELA='07' AND X5_FILIAL='"+xFilial('SX5')+"' ORDER BY X5_CHAVE"
	cQry := "SELECT DISTINCT X5_CHAVE, X5_DESCRI from "+RETSQLNAME('SE1')+ " SE1 "
	cQry += "JOIN "+RETSQLNAME('SX5') + " SX5 on SX5.D_E_L_E_T_ = '' AND SX5.X5_TABELA='07' AND SX5.X5_FILIAL='"+xFilial('SX5')+"' AND SX5.X5_CHAVE=SE1.E1_SITUACA "
	cQry += "WHERE SE1.E1_SALDO>0 AND SE1.D_E_L_E_T_ = ''"
	cQry += "ORDER BY X5_CHAVE" 
	
	TCQUERY cQry NEW ALIAS "TRB"	
	
	dbselectarea("TRB")
	
	MvParDef := ""
	WHILE !(TRB->(EOF()))
		++nCont
		AADD(aSit, ALLTRIM(TRB->X5_CHAVE) + " - " + TRB->X5_DESCRI)
		MvParDef += ALLTRIM(TRB->X5_CHAVE)
		TRB->(DBSKIP())  
	ENDDO
	TRB->(DBCLOSEAREA())
	
	cTitulo := "Filtro Situação"
	
	IF lTipoRet
		IF f_Opcoes(@MvPar, cTitulo, @aSit, @MvParDef, NIL, NIL, l1Elem, 1, 99)  // Chama funcao f_Opcoes
			&MvRet := StrTran(mvpar, '*', '')                                                                          // Devolve Resultado
		EndIF
	EndIF
Return(IF( lTipoRet, .T., MvParDef))

Static Function RetSituacao()
	Local cReturn := ""
	Local nPos	  := 0
	Local nI   	  := 0
	
	nPos := 1
	FOR nI := 1 TO INT(LEN(MV_PAR26))
		IF cReturn == ""
			cReturn += "'"
		ELSE
			cReturn += "', '"
		ENDIF		
		cReturn += SUBSTR(MV_PAR26, nPos, 1) 
		nPos++
	NEXT	
Return cReturn



