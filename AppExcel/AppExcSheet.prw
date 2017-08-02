/*
Copyright 2015 AppSoft - Fabrica de Software

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

#Include "Totvs.ch"         

#DEFINE ROW_ID 		1  
#DEFINE ROW_OBJ  	2
            

/*/{Protheus.doc} AppExcSheet
Classe de gerenciamento das abas na planilha Excel, esta classe � utilizada e manipulada pela classe AppExcel
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type class
/*/
CLASS AppExcSheet From LongClassName
	DATA cName 			AS String  HIDDEN
	DATA aRows 			AS Array   HIDDEN      
	DATA nMaxColumn 	AS INTEGER HIDDEN     
	DATA nMaxRow      	AS INTEGER HIDDEN       
	DATA oSheetOptions	AS Objtect HIDDEN    
	
	//Class Properties
	DATA cClassName AS STRING HIDDEN  
	                     
	METHOD New(cNamePar) CONSTRUCTOR            
	METHOD AddCell(nRow,nCol,xContent,oStyle,cFormula)
	METHOD SetName(cNamePar)
	METHOD GetName()            
	METHOD GetColumnCount()
	METHOD GetRowCount()
	METHOD OrderRows() 	        
	METHOD OrderSheet()   
	METHOD AddIndexRows()
	METHOD RowToString( nRow )               
	
	METHOD HasOptions()               
	METHOD OptionsToString()
	METHOD SetHorzFrozen( nRows )			
	METHOD SetVertFrozen( nRows )    
	
	//Class Properties
	METHOD Destroy()  
	METHOD ClassName()
	
ENDCLASS               
                                                  

/*/{Protheus.doc} AppExcSheet:New
M�todo construtor da classe AppExcSheet
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@param [cSheetName], caractere, nome que ser� apresentada na aba (Sheet) do Excel
@type method
/*/
METHOD New(cNamePar) CLASS AppExcSheet
	DEFAULT cNamePar := "Sheet"    
	
	::cClassName 	:= "AppExcSheet"

	::cName 	 	:= cNamePar
	::aRows	 	 	:= {}
	::nMaxColumn 	:= 1
	::nMaxRow	 	:= 1	
	     
	::oSheetOptions	:= AppExcSheetOptions():New()
RETURN                     
        
                              
/*/{Protheus.doc} AppExcSheet:AddCell
M�todo com os tratamentos para adicionar uma nova c�lula na aba, sendo de tipo v�riavel, podendo ter: estilo, f�rmula ou jun��es com outras c�lulas
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param nRow				,inteiro	, n�mero da linha a qual a c�lula ser� atribuida
@param nCol				,inteiro	, n�mero da coluna a qual a c�lula ser� atribuida
@param xContent			,indefinido	, conteudo da c�lula, podendo ser: num�rico, string ou data
@param [oStyle]			,objeto		, objeto AppExcCell com a defini��o (estilo) da c�lula
@param [cFormula]		,String		, Formula no padr�o Excel, ex.: "=RC[-3]+RC[-2]+RC[-1]"
@param [nMergeAcross]	,inteiro	, n�mero de c�lulas a frente da refer�ncia que ser�o mescladas
@param [nMergeDown]		,inteiro	, n�mero de c�lulas a abaixo da refer�ncia que ser�o mescladas
/*/
METHOD AddCell(nRow, nCol, xContent, oStyle, cFormula, nMergeAcross, nMergeDown) CLASS AppExcSheet
	Local oCell := nil
	Local nPos  := 0                    
	
	DEFAULT oStyle  		:= nil      
	DEFAULT cFormula  		:= nil
	DEFAULT nMergeAcross	:= 0
	DEFAULT nMergeDown     	:= 0
	                      
	If nRow > 0 .And. nCol > 0      
		oCell := AppExcCellProperties():New()
	
		oCell:AddRow( nRow )
		oCell:AddCol( nCol )
		oCell:AddContent( xContent ) 
		oCell:AddStyle( oStyle )           
		oCell:AddFormula( cFormula )          
		
		If nMergeAcross > 0 .Or. nMergeDown > 0
			oCell:SetMerged( .T. )
			
			If nMergeAcross > 0
				oCell:SetMergeAcross( nMergeAcross ) 
			EndIf			
			                                        
			If nMergeDown > 0
				oCell:SetMergeDown( nMergeDown ) 
			EndIf			
			
		EndIf    
		
		
		nPos := aScan( ::aRows, { |x| x[ROW_ID] == nRow } )
		                      		
		If ::nMaxRow < nRow + nMergeDown	
			::nMaxRow := nRow + nMergeDown	
		EndIf   
		
		
		If ::nMaxColumn < nCol + nMergeAcross
			::nMaxColumn := nCol + nMergeAcross		
		EndIf      
		
		If nPos == 0
			aAdd(::aRows,{nRow, AppExcRow():New(nRow) })
			nPos := len(::aRows)            
		EndIf
		
		::aRows[nPos][ROW_OBJ]:AddCell(oCell)
		
	EndIf
		
RETURN
        


/*/{Protheus.doc} AppExcSheet:SetName
Altera o nome de apresenta��o da aba
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param cNameSheet, caractere, nome da aba (Sheet) que ser� adicionada na planilha
/*/
METHOD SetName(cNamePar) CLASS AppExcSheet
	::cName := cNamePar
RETURN                          

                       
/*/{Protheus.doc} AppExcSheet:GetName
Retorna o nome definido na aba
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD GetName() CLASS AppExcSheet
RETURN ::cName                                                             

/*/{Protheus.doc} AppExcSheet:GetColumnCount
Retorna a quantidade de colunas que a aba possui
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD GetColumnCount() CLASS AppExcSheet
RETURN ::nMaxColumn                             


/*/{Protheus.doc} AppExcSheet:GetRowCount
Retorna a quantidade de linhas que a aba possui
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD GetRowCount() CLASS AppExcSheet
RETURN ::nMaxRow
   
                                            
/*/{Protheus.doc} AppExcSheet:OrderRows
Ordena as linhas adicionadas a aba
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD OrderRows() CLASS AppExcSheet  
	Local nX := 0

	aSort(::aRows,,,{|x,y| x[ROW_ID] < y[ROW_ID]}  )
	
	For nX := 1 to len(::aRows)
		::aRows[nX][ROW_OBJ]:OrderCells()
	Next
RETURN

                             
/*/{Protheus.doc} AppExcSheet:OrderSheet
Ordena as linhas e abas da planilha
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD OrderSheet() CLASS AppExcSheet
	::OrderRows()
	::AddIndexRows()                        
RETURN

                                      
/*/{Protheus.doc} AppExcSheet:AddIndexRows
Adiciona indices as c�lulas, para evitar ter de adicionar c�lulas vazias na planilha gerada
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0
@type method
/*/
METHOD AddIndexRows() CLASS AppExcSheet
	Local nX 		:= 0
	Local nIndex	:= 0    
	
	If len(::aRows) > 0
		If ::aRows[1][ROW_ID] > 1
			::aRows[1][ROW_OBJ]:SetIndex( ::aRows[1][ROW_ID] )
		EndIf        
	
		For nX := 1 to len(::aRows) - 1   
			nIndex := ::aRows[nX + 1][ROW_ID] - ::aRows[nX][ROW_ID] 
			
			If nIndex > 1                                                                                             
				::aRows[nX + 1][ROW_OBJ]:SetIndex( ::aRows[nX][ROW_ID] + nIndex)	
			EndIf	
			
		Next
	EndIf	
	
RETURN
      

/*/{Protheus.doc} AppExcSheet:HasOptions
Verificar se a worksheet possui op��es personalizadas
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                              
@type method
/*/
METHOD HasOptions() CLASS AppExcSheet
RETURN ::oSheetOptions:HasOptions()
                               
   
/*/{Protheus.doc} AppExcSheet:OptionsToString
Retorna a serializa��o das op��es da worksheet no padr�o XML Excel
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                              
@type method
/*/
METHOD OptionsToString() CLASS AppExcSheet
RETURN ::oSheetOptions:OptionsToString()

                                           
/*/{Protheus.doc} AppExcSheet:SetHorzFrozen
Congela as linhas superiores na rolagem de tela
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                 
@type method             
@param nRows, inteiro, numero de linhas que ser�o congeladas
@example
	oExcel:SetHorzFrozen( 2 )
/*/
METHOD SetHorzFrozen( nRows ) CLASS AppExcSheet
      
	::oSheetOptions:SetHorzFrozen( nRows ) 

RETURN
      
                    
/*/{Protheus.doc} AppExcSheet:SetVertFrozen
Congela as linhas laterais na rolagem de tela
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                 
@type method             
@param nCols, inteiro, numero de colunas que ser�o congeladas
@example
	oExcel:SetVertFrozen( 2 )
/*/
METHOD SetVertFrozen( nCols ) CLASS AppExcSheet
      
	::oSheetOptions:SetVertFrozen( nCols ) 

RETURN

       
/*/{Protheus.doc} AppExcSheet:RowToString
Serializa��o da linha para o padr�o XML Excel
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
@param nRow, inteiro, n�mero da linha que ser� serializada
@return caractere, String contendo a linha especificada no padr�o XML
/*/
METHOD RowToString( nRow ) CLASS AppExcSheet
	Local cRow 		:= ""   
	Local nX   		:= 1 
	Local aCells    := {}
	
	Local nPos := aScan(::aRows, {|x| x[ROW_ID] == nRow}  )
		
		                                              
	If nPos > 0            
		cRow += Space(9)+ ::aRows[nPos][ROW_OBJ]:GetAssinature() + CRLF		
                
		For nX := 1 to ::aRows[nPos][ROW_OBJ]:GetSize()                     
			cRow += Space(12) + ::aRows[nPos][ROW_OBJ]:CellToString( nX ) + CRLF                                
		Next  
		
		cRow += Space(9)+'</Row>' + CRLF
			
	EndIf

	
RETURN cRow                                    

  
/*/{Protheus.doc} AppExcSheet:ClassName
M�todo respons�vel por retornar o nome da classe
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
@return caractere, retorna o nome da classe
/*/
METHOD ClassName() CLASS AppExcSheet
RETURN ::cClassName
              
  
/*/{Protheus.doc} AppExcSheet:Destroy
M�todo destrutor do objeto, respons�vel pela desaloca��o da mem�ria
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
/*/
METHOD Destroy() CLASS AppExcSheet
	Local nX := 0
	
	For nX := 1 to len(::aRows)
		::aRows[nX][ROW_OBJ]:Destroy()
	Next                           
	
	::oSheetOptions:Destroy()
	  
	FreeObj(self)
Return