#include "protheus.ch"
#include "topconn.ch"
#Include "AppExcel.ch"

/*/
_____________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦Programa  ¦ EXEMPLOEXCELSQL ¦ Autor ¦ Wagner Cabrera ¦ Data¦02/08/2016 ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Descricao ¦ Exportador padrão de excel por SQL						  ¦¦¦
¦¦¦          ¦					                               			  ¦¦¦
¦¦¦----------+------------------------------------------------------------¦¦¦
¦¦¦Uso       ¦		                                                      ¦¦¦
¦¦+-----------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
/*/

user function EXEMPLOEXCELSQL()
	Local cAgrupa		:= ""
	Local aVetTitulo	:= {}
	Local aVetDetalhe 	:= {}
	Local aVetCPOTit 	:= {}
	Local aVetCPODet 	:= {}	
	Local aVetParam 	:= {} //Passar parametro com Desc e conteudo
	Local cTitulo		:= "EXEMPLO exportação com SQL"
	
	//Array com Formtação do Titulo
	aVetTitulo 	:= {"#000000", "Arial", "12", .T., .F., .F., BORDER_LINE_CONTINUOUS, "", .T.}
	//Array com Formtação dos detalhes
	aVetDetalhe	:= {"#000000", "Arial", "10", .F., .F., .F., BORDER_LINE_CONTINUOUS, "", .T.}
	
	//Array com Campos
	aVetCPODet 	:= {"B1_COD", "B1_DESC"}
	//Array com Descrição dos Campos
	aVetCPOTit	:= {"Cod.Produto", "Descrição"}

	//SELECT COM CAMPOS A SEREM EXPORTADOS 
	cQuery := "SELECT B1_COD, B1_DESC FROM "+ RETSQLNAME("SB1") + " SB1 WHERE SB1.D_E_L_E_T_=' ' "
	
	//Chamada da funcao
	U_NTXEXCEL(aVetTitulo, aVetDetalhe, aVetCPOTit, aVetCPODet, cQuery, "Resumo", cTitulo, cAgrupa, aVetParam)	
return