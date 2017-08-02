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
#Include "AppExcel.ch"

/*/{Protheus.doc} AppExcCell
Classe fornecedora de m�todo para gerenciameto de forma��o de c�lulas em Excel
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                   
@type class
/*/
CLASS AppExcCell FROM LongClassName

	//Identifier
	DATA cID	AS STRING HIDDEN  
	
	//Borders
	DATA oBorderTop 	AS OBJECT HIDDEN              
	DATA oBorderBottom 	AS OBJECT HIDDEN              
	DATA oBorderLeft 	AS OBJECT HIDDEN              
	DATA oBorderRight 	AS OBJECT HIDDEN              
	DATA oBorderDLeft 	AS OBJECT HIDDEN              
	DATA oBorderDRight 	AS OBJECT HIDDEN              
    
	//Font
	DATA oFont	AS OBJECT HIDDEN
	                          
	//Interior
	DATA cCellColor	AS STRING HIDDEN            
	            
	//Format
	DATA oFormat AS OBJECT HIDDEN                           
	
	//Align
	DATA oAlign AS OBJECT HIDDEN                           
	    	                                    
	//Class Properties
	DATA cClassName AS STRING HIDDEN
	
	//QUEBRAR TEXTO AUTOMATICAMENTE
	DATA bQuebra AS BOOLEAN
	
	//Constructor                       
	METHOD New() CONSTRUCTOR    
	
	//Identifier
	METHOD SetId(cIdPar)    
	METHOD GetId()    
           
    //Borders           
	METHOD SetBorder(nTypeBorder, bSetBorder)    
	METHOD SetABorders(bSetBorder, bDiagonal)                 
	METHOD SetLineBorder(nTypeLine, nTypeLine)    
	METHOD SetALineBorders(nTypeLine, bDiagonal)  
	METHOD SetLineWeigth(nTypeBorder, nWeigth)       
	METHOD SetALineWeigth(nWeigth, bDiagonal)
	METHOD SetLineColor(nTypeBorder, cColor)       
	METHOD SetALineColor(cColor, bDiagonal)
	                                             
	METHOD GetBorder(nTypeBorder)    
	METHOD GetLineBorder(nTypeBorder)                  
	METHOD GetLineWeigth(nTypeBorder)            
	METHOD GetLineColor(nTypeBorder)          

	//Interior
	METHOD SetCellColor(cColor) 
	METHOD GetCellColor()
	           
	//Font
	METHOD SetFont(oFontPar)      
	METHOD GetFont()      
	
	//Format
	METHOD SetFormat( nFormat )	
	METHOD GetFormat()	
	                
	//Align
	METHOD SetVertAlign( nAlign )
	METHOD SetHorzAlign( nAlign )
	               
	METHOD GetVertAlign( )       
	METHOD GetHorzAlign( )
	                         
	//Manipulation
	METHOD Clone( oCellFather )	
	
	//To String
	METHOD CellToString()
	
	//Class Properties
	METHOD Destroy()  
	METHOD ClassName()	 
	
ENDCLASS

  
/*/{Protheus.doc} AppExcCell:New
M�todo construtor da classe AppExcCell
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                
@type constructor
/*/
METHOD New() CLASS AppExcCell
	::cClassName	:= "AppExcCell"

 	::oBorderTop 		:= AppExcCellBorder():New( BORDER_POSITION_TOP )
	::oBorderBottom 	:= AppExcCellBorder():New( BORDER_POSITION_BOTTOM ) 
	::oBorderLeft 		:= AppExcCellBorder():New( BORDER_POSITION_LEFT )
	::oBorderRight 		:= AppExcCellBorder():New( BORDER_POSITION_RIGHT )
	::oBorderDLeft 		:= AppExcCellBorder():New( BORDER_POSITION_DIAGONAL_LEFT )
	::oBorderDRight 	:= AppExcCellBorder():New( BORDER_POSITION_DIAGONAL_RIGHT )
	     
	::oFormat 		:= AppExcCellFormat():New()	 
	::oAlign 		:= AppExcCellAlign():New()	
	::oFont 		:= nil
RETURN       



/*/{Protheus.doc} AppExcCell:SetId
M�todo manipulador da propriedade cID
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0         
@type method                                 
@param cIdPar, character, c�digo do estilo que ser� atribuido         
/*/
METHOD SetId(cIdPar) CLASS AppExcCell
	::cID := cIdPar
RETURN


/*/{Protheus.doc} AppExcCell:GetId
M�todo de acesso da propriedade cID
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                             
@type method                              
@return caractere, identificador do objeto
/*/
METHOD GetId() CLASS AppExcCell   
RETURN ::cID  


/*/{Protheus.doc} AppExcCell:SetBorder
M�todo para ativar/desativar bordas nas c�lulas
@author anderson.toledo
@since 18/02/2014
@version 1.0
@type method
@param nTypeBorder, num�rico, identificador da borda podendo ser: |ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li| |li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li| |li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li| |li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li| |li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li| |li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|
@param bSetBorder, booleano, indica se a borda deve ser ativada ou n�o

/*/
METHOD SetBorder(nTypeBorder, bSetBorder ) CLASS AppExcCell     
                                               
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			::oBorderTop:SetBorder( bSetBorder )
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			::oBorderBottom:SetBorder( bSetBorder )
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			::oBorderLeft:SetBorder( bSetBorder )
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			::oBorderRight:SetBorder( bSetBorder )
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			::oBorderDLeft:SetBorder( bSetBorder )
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			::oBorderDRight:SetBorder( bSetBorder )
	End Case
	
RETURN



/*/{Protheus.doc} AppExcCell:SetABorders
M�todo para ativar/desativar todas as bordas da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param bSetBorder, l�gico, indica se a borda deve ser ativada ou n�o
@param bDiagonal, l�gico, indica se as bordas diagonais devem ser consideradas
/*/
METHOD SetABorders(bSetBorder, bDiagonal) CLASS AppExcCell  
	DEFAULT bDiagonal := .F.                                          
                    
	::oBorderTop:SetBorder( bSetBorder )
	::oBorderBottom:SetBorder( bSetBorder )
	::oBorderLeft:SetBorder( bSetBorder )
	::oBorderRight:SetBorder( bSetBorder )
	
	If bDiagonal
		::oBorderDLeft:SetBorder( bSetBorder )
		::oBorderDRight:SetBorder( bSetBorder )
	EndIf	
RETURN	            
           


/*/{Protheus.doc} AppExcCell:SetLineBorder
M�todo alterar o tipo da linha na borda
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                            
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|                                                                            
@param nTypeLine, inteiro, indica o tipo da linha podendo ser:|ul||li|BORDER_LINE_CONTINUOUS 		-> "Continuous", linha continua|/li||li|BORDER_LINE_DOT   			-> "Dot", linha pontilhada|/li||li|BORDER_LINE_DASHDOT 		-> "DashDot", linha intercalada pontilhada/tracejada|/li||li|BORDER_LINE_DASHDOTDOT      -> "DashDotDot", linha intercalada pontilhada/tracejada/tracejada|/li||li|BORDER_LINE_SLANTDASHDOT  	-> "SlantDashDot", linha intercalada pontilhada/tracejada inclinada|/li|        |li|BORDER_LINE_DOUBLE			-> "Double", linha dupla|/li||/ul|              
/*/
METHOD SetLineBorder(nTypeBorder, nTypeLine) CLASS AppExcCell  
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			::oBorderTop:SetLineStyle( nTypeLine )
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			::oBorderBottom:SetLineStyle( nTypeLine )
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			::oBorderLeft:SetLineStyle( nTypeLine )
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			::oBorderRight:SetLineStyle( nTypeLine )
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			::oBorderDLeft:SetLineStyle( nTypeLine )
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			::oBorderDRight:SetLineStyle( nTypeLine )
	End Case
RETURN      


/*/{Protheus.doc} AppExcCell:SetALineBorders
M�todo alterar o tipo da linha de todas as bordas
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                           
@type method
@param nTypeLine, inteiro, indica o tipo da linha podendo ser:|ul||li|BORDER_LINE_CONTINUOUS 		-> "Continuous", linha continua|/li||li|BORDER_LINE_DOT   			-> "Dot", linha pontilhada|/li||li|BORDER_LINE_DASHDOT 		-> "DashDot", linha intercalada pontilhada/tracejada|/li||li|BORDER_LINE_DASHDOTDOT      -> "DashDotDot", linha intercalada pontilhada/tracejada/tracejada|/li||li|BORDER_LINE_SLANTDASHDOT  	-> "SlantDashDot", linha intercalada pontilhada/tracejada inclinada|/li||li|BORDER_LINE_DOUBLE			-> "Double", linha dupla|/li||/ul|
@param bDiagonal, l�gico, indica se as bordas diagonais devem ser consideradas
/*/
METHOD SetALineBorders(nTypeLine, bDiagonal) CLASS AppExcCell  
 	DEFAULT bDiagonal := .F.   
	
	::oBorderTop:SetLineStyle( nTypeLine )
	::oBorderBottom:SetLineStyle( nTypeLine )
	::oBorderLeft:SetLineStyle( nTypeLine )
	::oBorderRight:SetLineStyle( nTypeLine )
	
	If bDiagonal
		::oBorderDLeft:SetLineStyle( nTypeLine )
		::oBorderDRight:SetLineStyle( nTypeLine )
	EndIf
RETURN
   
                                                        

/*/{Protheus.doc} AppExcCell:SetLineWeigth
M�todo alterar a espessura da linha
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|
@param nWeigth,inteiro,indica a esperrura da linha podendo variar de 0 a 3
/*/
METHOD SetLineWeigth(nTypeBorder, nWeigth) CLASS AppExcCell       
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			::oBorderTop:SetWeight( nWeigth )
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			::oBorderBottom:SetWeight( nWeigth )
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			::oBorderLeft:SetWeight( nWeigth )
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			::oBorderRight:SetWeight( nWeigth )
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			::oBorderDLeft:SetWeight( nWeigth )
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			::oBorderDRight:SetWeight( nWeigth )
	End Case
RETURN

/*/{Protheus.doc} AppExcCell:SetALineWeigth
M�todo alterar o tipo da linha de todas as bordas
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                           
@type method
@param nWeigth,inteiro,indica a esperrura da linha podendo variar de 0 a 3
@param bDiagonal, l�gico, indica se as bordas diagonais devem ser consideradas
/*/
METHOD SetALineWeigth(nWeigth, bDiagonal) CLASS AppExcCell
 	DEFAULT bDiagonal := .F.   
	
	::oBorderTop:SetWeight( nWeigth )
	::oBorderBottom:SetWeight( nWeigth )
	::oBorderLeft:SetWeight( nWeigth )
	::oBorderRight:SetWeight( nWeigth )
	
	If bDiagonal
		::oBorderDLeft:SetWeight( nWeigth )
		::oBorderDRight:SetWeight( nWeigth )
	EndIf
RETURN             
                                                                          


/*/{Protheus.doc} AppExcCell:SetLineColor
M�todo para alterar a cor da linha de uma borda
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                           
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li|                                                                            |/ul|
@param cColor, caractere, cor da linha em padr�o hexadecimal ex.: #000000 (preto)
/*/
METHOD SetLineColor(nTypeBorder, cColor) CLASS AppExcCell       
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			::oBorderTop:SetColor( cColor )
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			::oBorderBottom:SetColor( cColor )
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			::oBorderLeft:SetColor( cColor )
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			::oBorderRight:SetColor( cColor )
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			::oBorderDLeft:SetColor( cColor )
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			::oBorderDRight:SetColor( cColor )
	End Case
RETURN

                                                                                 
/*/{Protheus.doc} AppExcCell:SetALineColor
M�todo para alterar a cor da linha de todas as bordas
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                                  
@type method 
@param cColor, caractere, cor da linha em padr�o hexadecimal ex.: #000000 (preto)
@param bDiagonal, l�gico, indica se as bordas diagonais devem ser consideradas               
/*/
METHOD SetALineColor(cColor, bDiagonal) CLASS AppExcCell
 	DEFAULT bDiagonal := .F.   
	
	::oBorderTop:SetColor( cColor )
	::oBorderBottom:SetColor( cColor )
	::oBorderLeft:SetColor( cColor )
	::oBorderRight:SetColor( cColor )
	
	If bDiagonal
		::oBorderDLeft:SetColor( cColor )
		::oBorderDRight:SetColor( cColor )
	EndIf
RETURN
           
   

/*/{Protheus.doc} AppExcCell:SetCellColor
M�todo para alterar a cor de fundo de uma c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                                               
@type method
@param cColor, caractere, cor de funda da c�lula em padr�o hexadecimal ex.: #000000 (preto)
/*/
METHOD SetCellColor(cColor) CLASS AppExcCell  
	::cCellColor := cColor
RETURN                                                                         


/*/{Protheus.doc} AppExcCell:SetFont
M�todo para alterar a fonte utilizada na c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                           
@type method
@param oFontPar, objeto, objeto da classe AppExcCell
@See
AppExcFont
/*/
METHOD SetFont(oFontPar) CLASS AppExcCell  
	::oFont := oFontPar
RETURN      

      
/*/{Protheus.doc} AppExcCell:SetFormat
M�todo para alterar o formato da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0         
@type method                                                                      
@param nFormat, inteiro, c�digo do formato a ser utilizado, os formatos suportados s�o:|ul||li|NUMBER_CURRENCY_REAL		-> Numero formato moeda em R$|/li||li|NUMBER_CURRENCY_RED_REAL    -> Numero formato moeda em R$, negativos em vermelho|/li||/ul|                                               
@See
AppExcFormat
/*/
METHOD SetFormat( nFormat ) CLASS AppExcCell                                           
	::oFormat:SetFormat( nFormat )	                                         	                                              
RETURN                                        
         
/*/{Protheus.doc} AppExcCell:SetVertAlign
M�todo para alterar o alinhamento vertical da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0      
@type method                  
@param nAlign, inteiro, c�digo do tipo do alinhamento, sendo:|ul||li|VERTICAL_ALIGN_TOP   	 -> define o alinhamento vertical como "acima"|/li||li|VERTICAL_ALIGN_CENTER    -> define o alinhamento vertical como "centralizado"|/li||li|VERTICAL_ALIGN_BOTTOM	 -> define o alinhamento vertical como "abaixo"|/li||/ul|
@See
AppExcAlign
/*/
METHOD SetVertAlign( nAlign ) CLASS AppExcCell                                                  
	::oAlign:SetVertAlign( nAlign )
RETURN
        
/*/{Protheus.doc} AppExcCell:SetHorzAlign
M�todo para alterar o alinhamento horizontal da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param nAlign, inteiro, c�digo do tipo do alinhamento, sendo:|ul||li|HORIZONTAL_ALIGN_RIGHT   	-> define o alinhamento "a direira"|/li||li|HORIZONTAL_ALIGN_CENTER    	-> define o alinhamento "centralizado"|/li||li|HORIZONTAL_ALIGN_LEFT	 	-> define o alinhamento "a esquerda"|/li||/ul|
@See
AppExcAlign
/*/
METHOD SetHorzAlign( nAlign ) CLASS AppExcCell  
	::oAlign:SetHorzAlign( nAlign )
RETURN


/*/{Protheus.doc} AppExcCell:GetBorder
M�todo de acesso para verificar se determina borda est� ativada
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|                                               
@return l�gico, indica se a borda informada est� ativa
/*/
METHOD GetBorder(nTypeBorder) CLASS AppExcCell    
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			Return ::oBorderTop:GetBorder()
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			Return ::oBorderBottom:GetBorder()
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			Return ::oBorderLeft:GetBorder()
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			Return ::oBorderRight:GetBorder()
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			Return ::oBorderDLeft:GetBorder()
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			Return ::oBorderDRight:GetBorder()
	End Case
RETURN                                                             


/*/{Protheus.doc} AppExcCell:GetLineBorder
M�todo de acesso para verificar o tipo da linha em uma determinada borda
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|
@return inteiro, indica o c�digo do tipo da linha utilizada
/*/
METHOD GetLineBorder(nTypeBorder) CLASS AppExcCell  
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			Return ::oBorderTop:GetLineStyle()
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			Return ::oBorderBottom:GetLineStyle()
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			Return ::oBorderLeft:GetLineStyle()
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			Return ::oBorderRight:GetLineStyle()
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			Return ::oBorderDLeft:GetLineStyle()
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			Return ::oBorderDRight:GetLineStyle()
	End Case
RETURN                        
     
/*/{Protheus.doc} AppExcCell:GetLineWeigth
M�todo de acesso para verificar a espessura da linha em uma determinada borda
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                             
@type method 
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|
@return inteiro, indica a espessura da borda especificada
/*/
METHOD GetLineWeigth(nTypeBorder) CLASS AppExcCell       
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			Return ::oBorderTop:GetWeight()
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			Return ::oBorderBottom:GetWeight()
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			Return ::oBorderLeft:GetWeight()
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			Return ::oBorderRight:GetWeight()
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			Return ::oBorderDLeft:GetWeight()
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			Return ::oBorderDRight:GetWeight()
	End Case
RETURN           
     

/*/{Protheus.doc} AppExcCell:GetLineColor
M�todo de acesso para verificar a cor da linha em uma determinada borda
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                                                             
@type method
@param nTypeBorder, inteiro, identificador da borda podendo ser:|ul||li|BORDER_POSITION_TOP 			-> borda superior da c�lula|/li||li|BORDER_POSITION_BOTTOM 			-> borda inferior da c�lula|/li||li|BORDER_POSITION_LEFT 			-> borda esquerda da c�lula|/li||li|BORDER_POSITION_RIGHT 			-> borda direita da c�lula|/li||li|BORDER_POSITION_DIAGONAL_LEFT 	-> borda diagonal da esquerda para direita|/li||li|BORDER_POSITION_DIAGONAL_RIGHT 	-> borda diagonal da direita para a esquerda|/li||/ul|                                               
@return caractere, string contendo a cor da c�lula em padr�o hexadecimal
/*/
METHOD GetLineColor(nTypeBorder) CLASS AppExcCell       
	Do Case
		Case BORDER_POSITION_TOP == nTypeBorder  	
			Return ::oBorderTop:GetColor()
		Case BORDER_POSITION_BOTTOM == nTypeBorder  	
			Return ::oBorderBottom:GetColor()
		Case BORDER_POSITION_LEFT == nTypeBorder  	
			Return ::oBorderLeft:GetColor()
		Case BORDER_POSITION_RIGHT == nTypeBorder  	
			Return ::oBorderRight:GetColor()
		Case BORDER_POSITION_DIAGONAL_LEFT == nTypeBorder  	
			Return ::oBorderDLeft:GetColor()
		Case BORDER_POSITION_DIAGONAL_RIGHT == nTypeBorder  	 
			Return ::oBorderDRight:GetColor()
	End Case
RETURN    


/*/{Protheus.doc} AppExcCell:GetCellColor
M�todo de acesso para obter a cor de fundo da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@return character, string contendo a cor de fundo da c�lula em padr�o hexadecimal
/*/
METHOD GetCellColor(cteste) CLASS AppExcCell                 
RETURN ::cCellColor    

/*/{Protheus.doc} AppExcCell:GetFont
M�todo de acesso para obter a fonte utilizada na c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@return objeto, objeto da classe AppExcCell utilizada na c�lula
/*/
METHOD GetFont() CLASS AppExcCell  
RETURN ::oFont
   
/*/{Protheus.doc} AppExcCell:GetFormat
M�todo de acesso para obter o formato da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@return inteiro, retorna o identificador da forma��o da c�lula
/*/
METHOD GetFormat( ) CLASS AppExcCell                                           
RETURN ::oFormat:GetFormat( )                                  


/*/{Protheus.doc} AppExcCell:GetVertAlign
M�todo de acesso obter o alinhamento vertical da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@return inteiro, identificador do alinhamento vertical da c�lula
/*/
METHOD GetVertAlign( ) CLASS AppExcCell                                                  
RETURN ::oAlign:GetVertAlign( )

/*/{Protheus.doc} AppExcCell:GetHorzAlign
M�todo de acesso obter o alinhamento horizontal da c�lula
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                        
@type method
@return inteiro, identificador do alinhamento horizontal da c�lula
/*/
METHOD GetHorzAlign( nAlign ) CLASS AppExcCell  
RETURN ::oAlign:GetHorzAlign( )
                                                                     
/*/{Protheus.doc} AppExcCell:Clone
M�todo para copiar todos atributos de um objeto da classe AppExcCell, evitando duplicidade na cria��o do script AdvPl da planilha
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0                           
@type method
@param oCellFather, objeto, objeto da classe AppExcCell
@see
AppExcCell
/*/
METHOD Clone( oCellFather ) CLASS AppExcCell  

    //Borders
	::SetBorder(BORDER_POSITION_LEFT			, oCellFather:GetBorder(BORDER_POSITION_LEFT) )   
	::SetBorder(BORDER_POSITION_RIGHT			, oCellFather:GetBorder(BORDER_POSITION_RIGHT) )   
	::SetBorder(BORDER_POSITION_TOP				, oCellFather:GetBorder(BORDER_POSITION_TOP) )   
	::SetBorder(BORDER_POSITION_BOTTOM			, oCellFather:GetBorder(BORDER_POSITION_BOTTOM) )   
	::SetBorder(BORDER_POSITION_DIAGONAL_LEFT	, oCellFather:GetBorder(BORDER_POSITION_DIAGONAL_LEFT) )   
	::SetBorder(BORDER_POSITION_DIAGONAL_RIGHT	, oCellFather:GetBorder(BORDER_POSITION_DIAGONAL_RIGHT) )   
	
	::SetLineBorder(BORDER_POSITION_LEFT			, oCellFather:GetLineBorder(BORDER_POSITION_LEFT))    
	::SetLineBorder(BORDER_POSITION_RIGHT			, oCellFather:GetLineBorder(BORDER_POSITION_RIGHT))    
	::SetLineBorder(BORDER_POSITION_TOP				, oCellFather:GetLineBorder(BORDER_POSITION_TOP))    
	::SetLineBorder(BORDER_POSITION_BOTTOM			, oCellFather:GetLineBorder(BORDER_POSITION_BOTTOM))    
	::SetLineBorder(BORDER_POSITION_DIAGONAL_LEFT	, oCellFather:GetLineBorder(BORDER_POSITION_DIAGONAL_LEFT))    
	::SetLineBorder(BORDER_POSITION_DIAGONAL_RIGHT	, oCellFather:GetLineBorder(BORDER_POSITION_DIAGONAL_RIGHT))
			
	::SetLineWeigth(BORDER_POSITION_LEFT			, oCellFather:GetLineWeigth(BORDER_POSITION_LEFT))       
	::SetLineWeigth(BORDER_POSITION_RIGHT			, oCellFather:GetLineWeigth(BORDER_POSITION_RIGHT))       
	::SetLineWeigth(BORDER_POSITION_TOP				, oCellFather:GetLineWeigth(BORDER_POSITION_TOP))       
	::SetLineWeigth(BORDER_POSITION_BOTTOM			, oCellFather:GetLineWeigth(BORDER_POSITION_BOTTOM))       
	::SetLineWeigth(BORDER_POSITION_DIAGONAL_LEFT	, oCellFather:GetLineWeigth(BORDER_POSITION_DIAGONAL_LEFT))       
	::SetLineWeigth(BORDER_POSITION_DIAGONAL_RIGHT	, oCellFather:GetLineWeigth(BORDER_POSITION_DIAGONAL_RIGHT))
			
	::SetLineColor(BORDER_POSITION_LEFT				, oCellFather:GetLineColor(BORDER_POSITION_LEFT))       
	::SetLineColor(BORDER_POSITION_RIGHT			, oCellFather:GetLineColor(BORDER_POSITION_RIGHT))       
	::SetLineColor(BORDER_POSITION_TOP				, oCellFather:GetLineColor(BORDER_POSITION_TOP))       	
	::SetLineColor(BORDER_POSITION_BOTTOM			, oCellFather:GetLineColor(BORDER_POSITION_BOTTOM))       	
	::SetLineColor(BORDER_POSITION_DIAGONAL_LEFT	, oCellFather:GetLineColor(BORDER_POSITION_DIAGONAL_LEFT))       	
	::SetLineColor(BORDER_POSITION_DIAGONAL_RIGHT	, oCellFather:GetLineColor(BORDER_POSITION_DIAGONAL_RIGHT))       	

                             
	//Interior
	::SetCellColor( oCellFather:GetCellColor() ) 
	       
	//Font
	::SetFont( oCellFather:GetFont() )      
	
	//Format
	::SetFormat( oCellFather:GetFormat() )	
		                
	//Align
	::SetVertAlign( oCellFather:GetVertAlign() )
	::SetHorzAlign( oCellFather:GetHorzAlign() )
	               
RETURN
              
/*/{Protheus.doc} AppExcCell:CellToString
Serializa��o da c�lula para o padr�o XML Excel
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
@param nRow, inteiro, n�mero da linha que ser� serializada
@return caractere, String contendo a c�lula especificada no padr�o XML
/*/
METHOD CellToString() CLASS AppExcCell  
	Local cCell := ""                                        
     
    cCell := Space(6)+'<Style ss:ID="'+::cID+'">' + CRLF
            
    ::oAlign:bQuebra := ::bQuebra 
                              
    If ::oAlign:HasAlign()
		cCell += ::oAlign:AlignToString() + CRLF                                                                                       
    EndIf
    
    If ::oFormat:HasFormat()
  		cCell += ::oFormat:FormatToString() + CRLF                                                             
    EndIf                        
                                                                            
    
    If (::oBorderTop:HasBorder() .Or. ::oBorderBottom:HasBorder() .Or. ::oBorderLeft:HasBorder() .Or.;
     	::oBorderRight:HasBorder() .Or. ::oBorderDLeft:HasBorder() .Or. ::oBorderDRight:HasBorder() )
		
		cCell += Space(9)+'<Borders>' + CRLF                                   
			If ::oBorderTop:HasBorder()
				cCell += ::oBorderTop:BorderToString() + CRLF   
			EndIf
			If ::oBorderBottom:HasBorder()
				cCell += ::oBorderBottom:BorderToString() + CRLF   
			EndIf
			If ::oBorderLeft:HasBorder()
				cCell += ::oBorderLeft:BorderToString() + CRLF   
			EndIf
			If ::oBorderRight:HasBorder()
				cCell += ::oBorderRight:BorderToString() + CRLF   
			EndIf						
			If ::oBorderDLeft:HasBorder()
				cCell += ::oBorderDLeft:BorderToString() + CRLF   
			EndIf			
			If ::oBorderDRight:HasBorder()
				cCell += ::oBorderDRight:BorderToString() + CRLF   
			EndIf			
	    cCell += Space(9)+'</Borders>' + CRLF                                                                                         
    EndIf    
    
    If !Empty(::cCellColor)
        cCell += Space(9)+'<Interior ss:Color="'+::cCellColor+'" ss:Pattern="Solid"/>' + CRLF             
    EndIf                 
    
    If ValType(::oFont) == "O"
  		cCell += Space(9)+::oFont:FontToString() + CRLF  
    EndIf                 
    
    cCell += Space(6)+'</Style>' + CRLF                           
    
RETURN cCell

/*/{Protheus.doc} AppExcCell:ClassName
M�todo respons�velpor retornar o nome da classe
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
@return caractere, retorna o nome da classe
/*/
METHOD ClassName() CLASS AppExcCell 
RETURN ::cClassName 
                                
/*/{Protheus.doc} AppExcCell:Destroy
M�todo destrutor do objeto, respons�vel pela desaloca��o da mem�ria
@author Anderson Toledo - anderson@appsoft.com.br
@since 18/02/2014
@version 1.0     
@type method
/*/
METHOD Destroy() CLASS AppExcCell 
	::oBorderTop:Destroy()
	::oBorderBottom:Destroy()
	::oBorderLeft:Destroy()
	::oBorderRight:Destroy()
	::oBorderDLeft:Destroy()
	::oBorderDRight:Destroy()
	::oFormat:Destroy()
	::oAlign:Destroy()
    
	If ValType( ::oFont ) == "O"
		::oFont:Destroy()
	EndIf	
                                                           
	FreeObj(self)                                
RETURN                                          