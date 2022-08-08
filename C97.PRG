********C97.部品出貨毛利分析報表*****
SET EXCL OFF
CLOSE ALL
CLEAR 
DIMENSION REZRO[12]
STORE 0 TO REZRO
AREA1='YCO'
UKEF=''
FLG='0'
FCH=''
CHK=''
YEAR_STR=''			&&列印報表之資料年度
IF ! USED('A01')
   SELE 0
   USE A01 INDE A01
ELSE
   SELE A01
ENDIF      
SET ORDER TO A01
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'C97'	
********************
C97FORM=CREATEOBJECT("TKC97")
C97FORM.SHOW       
DEFINE CLASS TKC97 AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='C97.請選擇您所需之分析報表'
  LEFT=INT_015*120
  TOP=INT_015*230    
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*130
  WIDTH=INT_015*600
  FONTSIZE=INT_015*9
  MAXBUTTON=.F.
  MINBUTTON=.F.    
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKC97'
ADD OBJECT DATA_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*25,;
      TOP=INT_015*14,;
      HEIGHT=INT_015*100,;    
      WIDTH=INT_015*165,;      
      BUTTONCOUNT=5,;
      OPTION1.CAPTION='客戶別出貨毛利分析',;
      OPTION1.TOP=INT_015*3,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='產品別出貨毛利分析',;
      OPTION2.TOP=INT_015*23,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.AUTOSIZE=.T.,;
      OPTION3.CAPTION='業務員A別出貨毛利分析',;
      OPTION3.TOP=INT_015*43,;
      OPTION3.LEFT=INT_015*4,;
      OPTION3.FONTSIZE=INT_015*9,;
      OPTION3.AUTOSIZE=.T.,;
      OPTION4.CAPTION='業務員B別出貨毛利分析',;
      OPTION4.TOP=INT_015*63,;
      OPTION4.LEFT=INT_015*4,;
      OPTION4.FONTSIZE=INT_015*9,;
      OPTION4.AUTOSIZE=.T.,;      
      OPTION5.CAPTION='業務助理別出貨毛利分析',;
      OPTION5.TOP=INT_015*83,;
      OPTION5.LEFT=INT_015*4,;
      OPTION5.FONTSIZE=INT_015*9,;
      OPTION5.AUTOSIZE=.T.,;
      NAME='DATA_OPTION'       
  ADD OBJECT CHS2 AS OPTIONGROUP WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*43,;
      HEIGHT=INT_015*40,;    
      WIDTH=INT_015*100,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='毛利率%分析',;    
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='毛利分析' ,;         
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.AUTOSIZE=.T.,;        
      NAME='CHS2'                     
  ADD OBJECT CHS1 AS OPTIONGROUP WITH;
      LEFT=INT_015*310,;
      TOP=INT_015*33,;
      HEIGHT=INT_015*60,;    
      WIDTH=INT_015*60,;      
      BUTTONCOUNT=3,;
      OPTION1.CAPTION='部品',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='製品',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;      
      OPTION3.CAPTION='全部',;
      OPTION3.FONTSIZE=INT_015*9,;
      OPTION3.LEFT=INT_015*4,;
      OPTION3.TOP=INT_015*43,;
      OPTION3.AUTOSIZE=.T.,;     
      NAME='CHS1'               
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*380,;
      TOP=INT_015*55,;
      AUTOSIZE=.T.,;
      CAPTION='分析資料之年份'      
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*465,;
      TOP=INT_015*50,;
      WIDTH=INT_015*50          
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*530,;     
      TOP=INT_015*30,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*530,;
      TOP=INT_015*71,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='離開此作業!快速鍵->ALT+X',;
      NAME='CMND2'  
 PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')
 ENDPROC         
  PROC MTH_LIST.INIT        
       IF !USED('A23')
	    SELECT  0
	    USE A23
       ELSE 
	    SELECT  A23
       ENDIF
       SET ORDER TO A231    
       SELECT  LEFT(DTOS(A23.F01),4) FROM A23 ORDER BY A23.F01  INTO CURSOR A23A
       SELECT A23A.EXP_1 FROM A23A GROUP BY A23A.EXP_1 ORDER BY A23A.EXP_1 INTO CURSOR A23B  
       SELECT A23B  
       GO TOP
       WITH THIS
            K=0
            S=0
            DO WHILE !EOF()                
                   K=K+1
                  .ADDITEM(A23B.EXP_1)
                   IF LEFT(DTOS(DATE()),4)=A23B.EXP_1
                       S=K
                   ENDIF       
                   SKIP
            ENDDO   
            .VALUE=S
       ENDWITH   
       SELECT A23
       USE
       SELECT A23A
       USE 
       SELECT A23B
       USE
  ENDPROC     
***********************
PROCEDURE CMND1.CLICK
     YEAR_STR=THISFORM.MTH_LIST.DISPLAYVALUE
     DO CASE
            CASE THISFORM.CHS1.VALUE=1
                       UKEF='Y'
            CASE THISFORM.CHS1.VALUE=2
                        UKEF='N'
            CASE THISFORM.CHS1.VALUE=3
                        UKEF='Z'
     ENDCASE                           
     DO CASE 
           CASE THISFORM.CHS2.VALUE=1       
                      FLG='0'
           CASE THISFORM.CHS2.VALUE=2
                      FLG='1'
      ENDCASE     
      DO CASE 
            CASE THISFORM.DATA_OPTION.VALUE=1
	     		IF FLG='0'
	     		     CREATE CURSOR YCO1;
	    		     (F01 C(5),F04 C(8),JAN N(11,6),FEB N(11,6),MAR N(11,6),APR N(11,6),MAY N(11,6),JUN N(11,6),JUL N(11,6),AUG N(11,6),SEP N(11,6),OCT N(11,6),NOV N(11,6),DEC N(11,6))
	    		ELSE
	     		     CREATE CURSOR YCO1;
	    		     (F01 C(5),F04 C(8),JAN N(11),FEB N(11),MAR N(11),APR N(11),MAY N(11),JUN N(11),JUL N(11),AUG N(11),SEP N(11),OCT N(11),NOV N(11),DEC N(11))	    			 
	    		 ENDIF  
	     		 INDE ON F01 TAG YCO1_1
	     		 SELECT YCO1
	     		 SET ORDER TO YCO1_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 G='G11'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F09 C(8),F79 C(43))
		             H='C10_TEMP'
		             CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))	      
		             G='G11_TEMP'          
		         ELSE
		              IF !USED('&H')
 		 		         SELECT 0
		 	             USE (H) 
		 	          ELSE
		 		         SELECT H
		 	          ENDIF       		  
		              IF !USED('&G')
 		 		         SELECT 0
		 	             USE (G) 
		 	          ELSE
		 		         SELECT G
		 	          ENDIF       			 	                     
		         ENDIF	      	     	
		         *********
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR			         	 
		 	    FOR  I = 1 TO 12
		 	          SELECT F02,F09 FROM &H ORDER BY F02 GROUP BY F02 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO1
		 	                 SEEK C01_TEMP.F02
		 	                 IF !FOUND()
		 	                     SELECT YCO1
		 	                     APPEND BLANK 
		 	                     REPLACE YCO1.F01 WITH C01_TEMP.F02
		 	                     REPLACE YCO1.F04 WITH C01_TEMP.F09
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE	
				  ****** 		     
*		      		  SELECT  &H..F02,SUM(&H..F07*&H..F08),SUM(B01.F37*&H..F08) FROM &H,B01 WHERE &H..F08>0 AND B01.F01=&H..F03 AND B01.F97<>UKEF ORDER BY &H..F02 GROUP BY &H..F02  INTO CURSOR C10_1 
		      		  SELECT  &H..F02,SUM(&H..F07*&H..F08),SUM(&G..F21*&H..F08),B01.F97 FROM &H,&G,B01 WHERE B01.F01=&H..F03 AND B01.F97<>UKEF AND &G..F02= &H..F79 ORDER BY &H..F02 GROUP BY &H..F02  INTO CURSOR C10_1  &&&H..F08>0 AND 
		     		  SELECT  C10_1
		      		  GO TOP
		      		  DO WHILE !EOF()
		             		 SELECT  YCO1
		             		 SEEK C10_1.F02
		                  	 IF  FOUND()
		                       	      T=FIELD(I+2)
		                       	      IF FLG='0'
                       			          IF C10_1.SUM_EXP_2<>0
                      			               REPL &T WITH ROUND((C10_1.SUM_EXP_2-C10_1.SUM_EXP_3)/C10_1.SUM_EXP_2*100,6)
                      			          ENDIF   
                      			      ELSE
                      			          REPLACE &T WITH ROUND(C10_1.SUM_EXP_2-C10_1.SUM_EXP_3,0)
                      			      ENDIF     
                      			      IF &T<>0
                      			          REZRO[I]=REZRO[I]+1
                      			      ENDIF    
                  		         ENDIF          
            			  	 SELE C10_1
            				 SKIP
     				   ENDDO      
     				  SELECT &H
		     		  USE
		     		  SELECT &G
		     		  USE
		     		  SELE C10_1
		     		  USE
		     		  S=VAL(RIGHT(H,6))+1
		    		  H='C10'+ALLTRIM(STR(S))
		    		  G='G11'+ALLTRIM(STR(S))
		    		  F=ALLTRIM(H)+'.DBF'
				  IF FILE('&F')=.F.
				        CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F09 C(8),F79 C(43))
				        H='C10_TEMP'
				        CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
				        G='G11_TEMP'
				  ELSE
 	                   IF !USED('&H')
			 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
			 		     SELECT H
			 	        ENDIF  				   				        
 	                   IF !USED('&G')
			 		     SELECT 0
			 	             USE (G) 
			 	        ELSE
			 		     SELECT G
			 	        ENDIF  					 	        
				  ENDIF			    		    
		     	 ENDFOR
		     	 WAIT CLEAR
		     	 ****
		     	 SELECT  YCO1
		     	 DELE FOR JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC=0
		 	 IF FLG='0'
  		 	     SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC)/ALIB(JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC)  SUM_TOL FROM YCO1 INTO CURSOR YCO READWRITE 
		 	 ELSE
                 SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC) SUM_TOL FROM YCO1  INTO CURSOR YCO READWRITE 
		     ENDIF    
  			 SELECT YCO1
  			 USE
  			 SELECT YCO		         
		 	 INDEX ON F01 TAG YCO_1
		 	 SET ORDER TO YCO_1
		      	 GO TOP
		     	IF _TALLY < 0 
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	ELSE
			      DO C97A		     	     
		        ENDIF			     	
            CASE THISFORM.DATA_OPTION.VALUE=2
			IF !USED('B01')
		  	     SELE 0
		   	     USE B01 INDE B01
			ELSE
		   	     SELE B01
			ENDIF
			SET ORDER TO B011                
	     		IF FLG='0'
  	     		    CREATE TABLE YCO1;
	    		    (F01 C(43),F02 C(40),JAN N(11,6),FEB N(11,6),MAR N(11,6),APR N(11,6),MAY N(11,6),JUN N(11,6),JUL N(11,6),AUG N(11,6),SEP N(11,6),OCT N(11,6),NOV N(11,6),DEC N(11,6))
	    		ELSE
  	     		    CREATE TABLE YCO1;
	    		    (F01 C(43),F02 C(40),JAN N(11),FEB N(11),MAR N(11),APR N(11),MAY N(11),JUN N(11),JUL N(11),AUG N(11),SEP N(11),OCT N(11),NOV N(11),DEC N(11))	    			 
	    		 ENDIF   
	     		 INDE ON F01 TAG YCO1_1
	     		 SELECT YCO1
	     		 SET ORDER TO YCO1_1	     		 
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 G='G11'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F03 C(43),F07 N(14,6),F08 N(11),F79 C(43))
		             H='C10_TEMP'	          
		             CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
		             G='G11_TEMP'      
		         ELSE
		              IF !USED('&H')
		 		         SELECT 0
		 	             USE (H) 
		 	          ELSE
		 		        SELECT H
		 	          ENDIF       		             
		              IF !USED('&G')
		 		         SELECT 0
		 	             USE (G) 
		 	          ELSE
		 		        SELECT G
		 	          ENDIF       			 	          
		         ENDIF	      	     	
		         *********
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR			         	 
		 	FOR  I = 1 TO 12
		 	          SELECT F03 FROM &H ORDER BY F03 GROUP BY F03 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO1
		 	                 SEEK C01_TEMP.F03
		 	                 IF !FOUND()
		 	                     SELECT YCO1
		 	                     APPEND BLANK 
		 	                     REPLACE YCO1.F01 WITH C01_TEMP.F03
		 	                     REPLACE YCO1.F02 WITH IIF(SEEK(C01_TEMP.F03,'B01'),B01.F02,'')
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 	      			 	
      				  SELE  &H..F03,SUM(&H..F07*&H..F08),SUM(&G..F21*&H..F08),B01.F97 FROM &H,&G,B01 WHERE &H..F07*&H..F08<>0 AND &G..F02=&H..F79 AND B01.F01=&H..F03 AND B01.F97<>UKEF ORDER BY &H..F03 GROUP BY &H..F03 INTO CURSOR C10_1 &&&H..F08>0 AND 
     				  SELE C10_1
      				  GO TOP
      				  DO WHILE !EOF()
             				 SELE YCO1
             				 SEEK C10_1.F03
                  			 IF   FOUND()
                       			       T=FIELD(I+2)
                       			       IF FLG='0'
                       			            IF C10_1.SUM_EXP_2<>0
                      			                REPL &T WITH  ROUND((C10_1.SUM_EXP_2-C10_1.SUM_EXP_3)/C10_1.SUM_EXP_2*100,6)
                      			            ENDIF    
                      			        ELSE
                      			            REPL &T WITH  C10_1.SUM_EXP_2-C10_1.SUM_EXP_3
                      			        ENDIF    
                      			       IF &T<>0
                       			           REZRO[I]=REZRO[I]+1
                      			       ENDIF                          			        
                  			 ENDIF          
            			  	 SELE C10_1
            				 SKIP
     				  ENDDO      
     				  SELECT &H
		     		  USE
		     		  SELECT &G
		     		  USE
		     		  SELE C10_1
		     		  USE
		     		  S=VAL(RIGHT(H,6))+1
		    		  H='C10'+ALLTRIM(STR(S))
		    		  G='G11'+ALLTRIM(STR(S))
		    		  F=ALLTRIM(H)+'.DBF'
				  IF FILE('&F')=.F.
				         CREATE CURSOR C10_TEMP (F03 C(43),F07 N(14,6),F08 N(11),F79 C(43))
				         H='C10_TEMP'
				         CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
				         G='G11_TEMP'
				  ELSE
		                IF !USED('&H')
			 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
			 		     SELECT H
			 	        ENDIF  				   				        
		                IF !USED('&G')
			 		     SELECT 0
			 	             USE (G) 
			 	        ELSE
			 		     SELECT G
			 	        ENDIF  						 	        
				  ENDIF			    		    
		     	 ENDFOR
		     	 WAIT CLEAR
		     	 SELECT B01
		     	 USE
     			 SELECT  YCO1
     			 DELE FOR JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC=0
	 		 IF FLG='0'
  			     SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC)/ALIB(JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC) SUM_TOL FROM YCO1  INTO CURSOR YCO READWRITE   
  			 ELSE
  			     SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC) SUM_TOL FROM YCO1   INTO CURSOR YCO READWRITE   				  
  			 ENDIF    
  			 SELECT YCO1
  			 USE
  			 SELECT YCO
	 		 INDEX ON F01 TAG YCO_1
	 		 SET ORDER TO YCO_1
     			 GO TOP
		     	IF _TALLY < 0 
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	ELSE
			      DO C97B	     	     
		        ENDIF		            
     	    CASE THISFORM.DATA_OPTION.VALUE=3 OR THISFORM.DATA_OPTION.VALUE=4	&&業務AB
	     		 IF FLG='0'	     			    
    	     		    CREATE CURSOR  YCO1;
	    		    (F18 C(5),F033 C(8),JAN N(11,6),FEB N(11,6),MAR N(11,6),APR N(11,6),MAY N(11,6),JUN N(11,6),JUL N(11,6),AUG N(11,6),SEP N(11,6),OCT N(11,6),NOV N(11,6),DEC N(11,6))
	    		 ELSE
    	     		    CREATE CURSOR  YCO1;	    			
	    		    (F18 C(5),F033 C(8),JAN N(11),FEB N(11),MAR N(11),APR N(11),MAY N(11),JUN N(11),JUL N(11),AUG N(11),SEP N(11),OCT N(11),NOV N(11),DEC N(11))	    			
	    		 ENDIF    
   			 INDEX ON F18 TAG YCO1_1
	     		 SELECT YCO1
	     		 SET ORDER TO YCO1_1   			
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 G='G11'+ALLTRIM(S)	      		 
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F10 C(5),F11 C(8),F12 C(5),F13 C(8),F79 C(43))
		             H='C10_TEMP'	                
		             CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
		             G='G11_TEMP'
		         ELSE
		              IF !USED('&H')
		 		          SELECT 0
		 	              USE (H) 
		 	          ELSE
		 		          SELECT H
   		 		      ENDIF       		             
		              IF !USED('&G')
		 		          SELECT 0
		 	              USE (G) 
		 	          ELSE
		 		          SELECT G
   		 		      ENDIF       	   		 		      
		         ENDIF	      	     	
		         *********
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR			         	 
		 	FOR  I = 1 TO 12
		 	          IF THISFORM.DATA_OPTION.VALUE=3 
		 	               SELECT F10,F11 FROM &H ORDER BY F10 GROUP BY F10 INTO CURSOR C01_TEMP
		 	          ELSE
		 	               SELECT F12,F13 FROM &H ORDER BY F12 GROUP BY F12 INTO CURSOR C01_TEMP
		 	          ENDIF     
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 IF THISFORM.DATA_OPTION.VALUE=3 
		 	                      SELECT YCO1
		 	                      SEEK C01_TEMP.F10
		 	                      IF !FOUND()
		 	                          SELECT YCO1
		 	                          APPEND BLANK 
		 	                          REPLACE YCO1.F18 WITH C01_TEMP.F10
		 	                          REPLACE YCO1.F033 WITH C01_TEMP.F11	 	                     
		 	                      ENDIF    
		 	                 ELSE
		 	                      SELECT YCO1
		 	                      SEEK C01_TEMP.F12
		 	                      IF !FOUND()
		 	                          SELECT YCO1
		 	                          APPEND BLANK 
		 	                          REPLACE YCO1.F18 WITH C01_TEMP.F12
		 	                          REPLACE YCO1.F033 WITH C01_TEMP.F13	 	                     
		 	                      ENDIF  		 	                      
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	         ****** 	  	
		 	         IF THISFORM.DATA_OPTION.VALUE=3 		 	
                                      SELECT &H..f10,SUM(&H..F07*&H..f08),SUM(&G..F21*&H..f08),B01.F97 FROM &H,&G,B01 GROUP BY &H..f10 ORDER BY &H..f10 WHERE  &G..F02=&H..F79 AND B01.f01=&H..f03 AND B01.F97<>UKEF INTO CURSOR C10_1 &&&AND H..f08>0 
	     			 ELSE
                                      SELECT &H..f12,SUM(&H..F07*&H..f08),SUM(&G..F21*&H..f08) FROM &H,&G,B01 GROUP BY &H..f12 ORDER BY &H..f12 WHERE  &G..F02=&H..F79 AND B01.f01=&H..f03 AND B01.F97<>UKEF INTO CURSOR C10_1 &&AND &H..f08>0  			 
	     			 ENDIF
	     			 SELE C10_1
	      			 GO TOP
	      			 DO WHILE !EOF()
	             			SELECT  YCO1
	             			SEEK IIF(THISFORM.DATA_OPTION.VALUE=3,C10_1.F10,C10_1.F12)
	                  		IF  FOUND()	                  			 
	                       		      T=FIELD(I+2)
	                       		      IF FLG='0'
	                       		           IF C10_1.SUM_EXP_2<>0
   	                      		               REPLACE  &T WITH ROUND((C10_1.SUM_EXP_2-C10_1.SUM_EXP_3)/C10_1.SUM_EXP_2*100,6)
   	                      		           ENDIF   
 	                      		      ELSE
 	                      		           REPLACE  &T WITH C10_1.SUM_EXP_2-C10_1.SUM_EXP_3
 	                      		      ENDIF    
                      			      IF &T<>0
                       			          REZRO[I]=REZRO[I]+1 	                      			       
                       			      ENDIF    
	                  		ENDIF          
	            			SELE C10_1
	            			SKIP
	     			 ENDDO      
     				  SELECT &H
		     		  USE
		     		  SELECT &G
		     		  USE
		     		  SELE C10_1
		     		  USE
		     		  S=VAL(RIGHT(H,6))+1
		    		  H='C10'+ALLTRIM(STR(S))
		    		  G='G11'+ALLTRIM(STR(S))
		    		  F=ALLTRIM(H)+'.DBF'
				  IF FILE('&F')=.F.
				       CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F10 C(5),F11 C(8),F12 C(5),F13 C(8),F79 C(43))
				        H='C10_TEMP'
				        CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
				        G='G11_TEMP'
				  ELSE
		                IF !USED('&H')
				 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
				 		     SELECT H
			 	        ENDIF  				   				        
		                IF !USED('&G')
				 		     SELECT 0
			 	             USE (G) 
			 	        ELSE
				 		     SELECT G
			 	        ENDIF  						 	        
				  ENDIF			    		    
		     	 ENDFOR
		     	 WAIT CLEAR
	     		 SELECT  YCO1
	     		DELE FOR JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC=0
		 	IF FLG='0'
  		 	    SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC)/ALIB(JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC) SUM_TOL FROM YCO1  INTO CURSOR YCO READWRITE 
  		 	ELSE
  		 	    SELECT  *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC) SUM_TOL FROM YCO1  INTO CURSOR YCO  READWRITE 		 			  
  		 	ENDIF    
  		 	SELECT YCO1
  		 	USE
  		 	SELECT YCO
		 	INDEX ON F18 TAG YCO_1		
		 	SET ORDER TO YCO_1		 
		     	GO TOP
		     	IF _TALLY < 0 
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	ELSE
		     	       IF THISFORM.DATA_OPTION.VALUE=3 	
			           DO C97C	     	     
			       ELSE
			           DO C97D
			       ENDIF    
		        ENDIF	    
     	    CASE THISFORM.DATA_OPTION.VALUE=5
	     		IF FLG='0'
  	     		    CREATE CURSOR YCO1;
	    		   (F23 C(5),F231 C(8),JAN N(11,6),FEB N(11,6),MAR N(11,6),APR N(11,6),MAY N(11,6),JUN N(11,6),JUL N(11,6),AUG N(11,6),SEP N(11,6),OCT N(11,6),NOV N(11,6),DEC N(11,6))
	    		ELSE
  	     		    CREATE CURSOR YCO1;
	    		   (F23 C(5),F231 C(8),JAN N(11),FEB N(11),MAR N(11),APR N(11),MAY N(11),JUN N(11),JUL N(11),AUG N(11),SEP N(11),OCT N(11),NOV N(11),DEC N(11))	    			
	    		ENDIF   
	     		INDE ON F23 TAG YCO1_1
                        SET ORDER TO YCO1_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 G='G11'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F14 C(5),F15 C(8),F79 C(43))
		             H='C10_TEMP'	                
		             CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
		             G='G11_TEMP'
		         ELSE
		              IF !USED('&H')
				 		   SELECT 0
			 	           USE (H) 
		 		      ELSE
		 			   SELECT H
		 	      	ENDIF       		             
		              IF !USED('&G')
				 		   SELECT 0
			 	           USE (G) 
		 		      ELSE
		 			   SELECT G
		 	      	ENDIF       			 	      	
		         ENDIF	      	     	
		         *********
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR			         	 
		 	FOR  I = 1 TO 12                        
		 	          SELECT F14,F15 FROM &H ORDER BY F14 GROUP BY F14 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO1
		 	                 SEEK C01_TEMP.F14
		 	                 IF !FOUND()
		 	                     SELECT YCO1
		 	                     APPEND BLANK 
		 	                     REPLACE YCO1.F23 WITH C01_TEMP.F14
		 	                     REPLACE YCO1.F231 WITH C01_TEMP.F15
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 	                          
                      SELECT &H..F14,SUM(&H..F07*&H..f08),SUM(&G..F21*&H..f08),B01.F97 FROM &H,&G,B01 GROUP BY &H..f14 ORDER BY &H..f14 WHERE  &G..F02=&H..F79 AND B01.f01=&H..f03 AND B01.F97<>UKEF INTO CURSOR C10_1 &&AND &H..f08>0 
	     			  SELECT  C10_1
	      			  GO TOP
	      			  DO WHILE !EOF()
	             		 	 SELECT  YCO1
	             			 SEEK C10_1.F14
	                  		 IF  FOUND()
	                       		      T=FIELD(I+2)
	                       		      IF FLG='0'
	                       		           IF C10_1.SUM_EXP_2<>0
       	                      		                REPL &T WITH ROUND((C10_1.SUM_EXP_2-C10_1.SUM_EXP_3)/C10_1.SUM_EXP_2*100,6)  	                      			           
       	                      		           ENDIF    
  	                      		       ELSE
  	                      		           REPL &T WITH C10_1.SUM_EXP_2-C10_1.SUM_EXP_3
  	                      		       ENDIF    
                      			       IF &T<>0
                       			            REZRO[I]=REZRO[I]+1 	                      			       
                       			       ENDIF      	                      			       
	                  		  ENDIF          
	            			  SELECT  C10_1
	            		          SKIP
	     			   ENDDO      
     				  SELECT &H
		     		  USE
		     		  SELECT &G
		     		  USE
		     		  SELE C10_1
		     		  USE
		     		  S=VAL(RIGHT(H,6))+1
		    		  H='C10'+ALLTRIM(STR(S))
		    		  G='G11'+ALLTRIM(STR(S))		    		  
		    		  F=ALLTRIM(H)+'.DBF'
				  IF FILE('&F')=.F.
				       CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F14 C(5),F15 C(8),F79 C(43))
				        H='C10_TEMP'
				        CREATE CURSOR G11_TEMP (F02 C(43),F21 N(14,6))
				        G='G11_TEMP'
				  ELSE
		                IF !USED('&H')
				 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
				 		     SELECT H
			 	        ENDIF  				   				        
		                IF !USED('&G')
				 		     SELECT 0
			 	             USE (G) 
			 	        ELSE
				 		     SELECT G
			 	        ENDIF  						 	        
				  ENDIF			    		    
		     	 ENDFOR
		     	 WAIT CLEAR
	     		 SELECT  YCO1
	     		 DELE FOR JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC=0
		 	 IF FLG='0'
  		 	      SELECT *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC)/ALIB(JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC) SUM_TOL ORDER BY F23 FROM YCO1 INTO CURSOR YCO READWRITE 
                         ELSE
  		 	      SELECT *,(JAN+FEB+MAR+APR+MAY+JUN+JUL+AUG+SEP+OCT+NOV+DEC) SUM_TOL ORDER BY F23 FROM YCO1 INTO CURSOR YCO   READWRITE                                      
                         ENDIF  		 		 
  			 SELECT YCO1
  			 USE
  			 SELECT YCO                         	     
		 	 INDEX ON F23  TAG YCO_1
		 	 SET ORDER TO YCO_1	
		     	 GO TOP
		         IF _TALLY < 0 
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	 ELSE
			      DO C97E	     	     
		         ENDIF         	       
           ENDCASE 
           THISFORM.CMND2.CLICK
 ENDPROC      
 **************************      
 PROCEDURE CMND2.CLICK       
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
 ENDPROC      
ENDDEFINE
**********
FUNCTION ALIB
PARAMETERS A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12
OUR=IIF(A1=0,0,1)+IIF(A2=0,0,1)+IIF(A3=0,0,1)+IIF(A4=0,0,1)+IIF(A5=0,0,1)+IIF(A6=0,0,1)+IIF(A7=0,0,1)+IIF(A8=0,0,1)+IIF(A9=0,0,1)+IIF(A10=0,0,1)+IIF(A11=0,0,1)+IIF(A12=0,0,1)
RETURN OUR


 
