********C99.年度出貨金額分析報表*****
SET EXCL OFF
CLOSE ALL
CLEAR 
AREA1='YCO'
UKEF=''
FLG='0'
FCH=''
CHK=''
KEY_CK=1
YEAR_STR=''			&&列印報表之資料年度
ACC=0	   			&&
TOLMOY=0   			&&
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
SEEK SYS_OPER+'C99'	
********************
C99FORM=CREATEOBJECT("TKC99")
C99FORM.SHOW       
DEFINE CLASS TKC99 AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='C99.請選擇您所需之分析報表'
  LEFT=INT_015*120
  TOP=INT_015*230  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*130
  WIDTH=INT_015*500
  FONTSIZE=INT_015*9
  MAXBUTTON=.F.
  MINBUTTON=.F.  
  CONTROLBOX=.F.
  CLOSABLE=.F.
  BORDERSTYLE=1
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKC99'
ADD OBJECT DATA_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*25,;
      TOP=INT_015*14,;
      HEIGHT=INT_015*100,;    
      WIDTH=INT_015*165,;      
      BUTTONCOUNT=5,;
      OPTION1.CAPTION='客戶別年度出貨分析',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='產品別年度出貨分析',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;
      OPTION3.CAPTION='業務員A別年度出貨分析',;
      OPTION3.FONTSIZE=INT_015*9,;
      OPTION3.LEFT=INT_015*4,;
      OPTION3.TOP=INT_015*43,;
      OPTION3.AUTOSIZE=.T.,;
      OPTION4.CAPTION='業務員B別年度出貨分析',;
      OPTION4.FONTSIZE=INT_015*9,;
      OPTION4.LEFT=INT_015*4,;
      OPTION4.TOP=INT_015*63,;
      OPTION4.AUTOSIZE=.T.,;
      OPTION5.CAPTION='業務助理別年度出貨分析',;
      OPTION5.FONTSIZE=INT_015*9,;
      OPTION5.LEFT=INT_015*4,;
      OPTION5.TOP=INT_015*83,;
      OPTION5.AUTOSIZE=.T.,;      
      NAME='DATA_OPTION'       
ADD OBJECT OUSK_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*30,;
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
      NAME='OUSK_OPTION'             
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*270,;
      TOP=INT_015*55,;
      AUTOSIZE=.T.,;
      CAPTION='分析資料之年份'
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*360,;
      TOP=INT_015*50,;
      WIDTH=INT_015*50          
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*430,;     
      TOP=INT_015*30,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*430,;
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
PROCEDURE CMND1.CLICK
     YEAR_STR=THISFORM.MTH_LIST.DISPLAYVALUE
      DO CASE
            CASE THISFORM.OUSK_OPTION.VALUE=1           
                        UKEF='N'
            CASE THISFORM.OUSK_OPTION.VALUE=2
                        UKEF='Y'
            CASE THISFORM.OUSK_OPTION.VALUE=3 
                        UKEF='YN'       
      ENDCASE                             
      DO CASE 
            CASE THISFORM.DATA_OPTION.VALUE=1
                 IF !USED('A01')
                    SELECT 0
                    USE A01
                 ELSE
                    SELECT A01
                 ENDIF
                 SET ORDER TO 1      
                 IF !USED('C01')
                    SELECT 0
                    USE C01
                 ELSE
                    SELECT C01
                 ENDIF
                 SET ORDER TO C011                  
	     		CREATE CURSOR YCO;
	    		(F01 C(5),F05 C(8),GRD C(1),SALES C(8),ASSIS C(8),JAN N(17,6),FEB N(17,6),MAR N(17,6),APR N(17,6),MAY N(17,6),JUN N(17,6),JUL N(17,6),AUG N(17,6),SEP N(17,6),OCT N(17,6),NOV N(17,6),DEC N(17,6),SUM_TOL N(17,6),;
	    		      JAN_AOT N(11),FEB_AOT N(11),MAR_AOT N(11),APR_AOT N(11),MAY_AOT N(11),JUN_AOT N(11),JUL_AOT N(11),AUG_AOT N(11),SEP_AOT N(11),OCT_AOT N(11),NOV_AOT N(11),DEC_AOT N(11),SUM_TO_AOT N(17),F44 C(5))
	     		 INDE ON F01 TAG YCO_1
	     		 INDEX ON SALES+GRD+F01 TAG YCO_2
	     		 INDEX ON ASSIS+GRD+F01 TAG YCO_3
	     		 INDE ON SUM_TOL DESC TAG YCO_4
	     		 INDE ON SUM_TO_AOT DESC TAG YCO_5
	     		 SET ORDER TO YCO_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F09 C(8))
		             H='C10_TEMP'	  
		         ELSE
		              IF !USED('&H')
		 		   SELECT 0
		 	           USE (H) 
		 	      ELSE
		 		   SELECT H
		 	      ENDIF                   
		         ENDIF	      		
		         *********
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR		         	 
		 	 FOR  I = 1 TO 12
		 	          SELECT F02,F09 FROM &H ORDER BY F02 GROUP BY F02 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO
		 	                 SEEK C01_TEMP.F02
		 	                 IF !FOUND()
		 	                     SELECT YCO
		 	                     APPEND BLANK 
		 	                     REPLACE YCO.F01 WITH C01_TEMP.F02
		 	                     REPLACE YCO.F05 WITH C01_TEMP.F09
		 	                     REPLACE YCO.GRD WITH IIF(SEEK(YCO.F01,'C01'),C01.F03,'')
		 	                     REPLACE YCO.SALES WITH IIF(SEEK(IIF(SEEK(YCO.F01,'C01'),C01.F33,SPACE(5)),'A01'),A01.F03,'')
		 	                     REPLACE YCO.ASSIS WITH IIF(SEEK(IIF(SEEK(YCO.F01,'C01'),C01.F23,SPACE(5)),'A01'),A01.F03,'')
		 	                     REPLACE YCO.F44 WITH IIF(SEEK(YCO.F01,'C01'),C01.F44,SPACE(5))
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ******
		      		  SELECT   F02,SUM(F07*F08),SUM(F08) FROM &H ORDER BY F02 GROUP BY F02 INTO CURSOR C10_1 WHERE F03=ANY(SELE F01 FROM B01 WHERE F97$UKEF)
		     		  SELECT  C10_1
		      		  GO TOP
		      		  DO WHILE !EOF()
		             		 SELECT  YCO
		             		 SEEK C10_1.F02
		                	 IF  FOUND()
		                  	      T=FIELD(I+5)
		                      	      REPLACE  &T WITH C10_1.SUM_EXP_2
		                      	      REPLACE  SUM_TOL WITH YCO.SUM_TOL+C10_1.SUM_EXP_2
		                      	      T=FIELD(I+18)
		                      	      REPLACE  &T WITH C10_1.SUM_F08
		                      	      REPLACE  SUM_TO_AOT WITH YCO.SUM_TO_AOT+C10_1.SUM_F08	                      	       
		                  	 ENDIF          
		            		 SELE C10_1
		            		 SKIP
		     		   ENDDO      
		     		   SELECT &H
		     		   USE
		     		   SELE C10_1
		     		   USE
		     		   S=VAL(RIGHT(H,6))+1
		    		   H='C10'+ALLTRIM(STR(S))
		    		   F=ALLTRIM(H)+'.DBF'
				   IF FILE('&F')=.F.
				        CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F09 C(8))
				        H='C10_TEMP'
				   ELSE
			                IF !USED('&H')
			 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
			 		     SELECT H
			 	        ENDIF  				   				        
				   ENDIF		    				  
		     	 ENDFOR
		     	 WAIT CLEAR
		     	 SELECT  YCO
		     	 DELETE  FOR SUM_TOL+SUM_TO_AOT=0
		     	 CALC SUM(YCO.SUM_TOL) TO TOLMOY
		     	 IF TOLMOY=0
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	 ELSE
		     	      SELECT YCO
			          DO C99A	     	     
		     	 ENDIF	 	            
             CASE THISFORM.DATA_OPTION.VALUE=2
   			      IF !USED('B01')
		  	         SELE 0
		   	         USE B01 INDE B01
			      ELSE
		   	         SELE B01
			     ENDIF
			     SET ORDER TO B011        
	     		CREATE CURSOR YCO;
	    		(F01 C(43),F02 C(40),JAN N(17,6),FEB N(17,6),MAR N(17,6),APR N(17,6),MAY N(17,6),JUN N(17,6),JUL N(17,6),AUG N(17,6),SEP N(17,6),OCT N(17,6),NOV N(17,6),DEC N(17,6),SUM_TOL N(17,6),;
	    		      JAN_AOT N(11),FEB_AOT N(11),MAR_AOT N(11),APR_AOT N(11),MAY_AOT N(11),JUN_AOT N(11),JUL_AOT N(11),AUG_AOT N(11),SEP_AOT N(11),OCT_AOT N(11),NOV_AOT N(11),DEC_AOT N(11),SUM_TO_AOT N(17))
	     		 INDE ON F01 TAG YCO_1
	     		 INDE ON SUM_TOL DESC TAG YCO_2
	     		 INDE ON SUM_TO_AOT DESC TAG YCO_3
	     		 SET ORDER TO YCO_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F03 C(43),F07 N(14,6),F08 N(11))
		             H='C10_TEMP'
		         ELSE
		              IF !USED('&H')
		 		   SELECT 0
		 	           USE (H) 
		 	      ELSE
		 		   SELECT H
		 	      ENDIF       		             
		         ENDIF	            
		         WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR	
 			 FOR  I = 1 TO 12
		 	          SELECT F03 FROM &H ORDER BY F03 GROUP BY F03 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO
		 	                 SEEK C01_TEMP.F03
		 	                 IF !FOUND()
		 	                     SELECT YCO
		 	                     APPEND BLANK 
		 	                     REPLACE YCO.F01 WITH C01_TEMP.F03
		 	                     REPLACE YCO.F02 WITH IIF(SEEK(C01_TEMP.F03,'B01'),B01.F02,'')
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 			 
      				  SELECT   F03,SUM(F07*F08),SUM(F08) FROM &H ORDER BY F03 GROUP BY F03 INTO CURSOR C10_1 WHERE F03 =ANY(SELE F01 FROM B01 WHERE F97$UKEF)
     				  SELECT  C10_1
      				  GO TOP
      				  DO WHILE !EOF()
             				 SELECT  YCO
             				 SEEK C10_1.F03
                  			 IF   FOUND()
                  			       T=FIELD(I+2)
                      			       REPLACE  &T WITH C10_1.SUM_EXP_2
                      			       REPLACE  SUM_TOL WITH YCO.SUM_TOL+C10_1.SUM_EXP_2
		                      	       T=FIELD(I+15)
		                      	       REPLACE  &T WITH C10_1.SUM_F08
		                      	       REPLACE  SUM_TO_AOT WITH YCO.SUM_TO_AOT+C10_1.SUM_F08
                  			 ENDIF          
            			  	 SELE C10_1
            				 SKIP
     				  ENDDO      
     				  SELE C10_1
     				  USE
		     		  S=VAL(RIGHT(H,6))+1
		    		  H='C10'+ALLTRIM(STR(S))
		    		  F=ALLTRIM(H)+'.DBF'
				  IF FILE('&F')=.F.
				       CREATE CURSOR C10_TEMP (F03 C(43),F07 N(14,6),F08 N(11))
				       H='C10_TEMP'
				  ELSE
			                IF !USED('&H')
			 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
			 		     SELECT H
			 	        ENDIF  					       
				  ENDIF	
     			 ENDFOR
     			 WAIT CLEAR
     			 SELECT B01
     			 USE
     			 SELE YCO
     			 DELE FOR SUM_TOL+SUM_TO_AOT=0
		     	 CALC SUM(YCO.SUM_TOL) TO TOLMOY	
		     	 IF TOLMOY=0
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	 ELSE
		     	      SELECT YCO
			      DO C99B	     	     
		     	 ENDIF	 	
     	      CASE THISFORM.DATA_OPTION.VALUE=3
	     		CREATE CURSOR YCO;
	    		(F18 C(5),F033 C(8),JAN N(17,6),FEB N(17,6),MAR N(17,6),APR N(17,6),MAY N(17,6),JUN N(17,6),JUL N(17,6),AUG N(17,6),SEP N(17,6),OCT N(17,6),NOV N(17,6),DEC N(17,6),SUM_TOL N(17,6),;
	    		      JAN_AOT N(11),FEB_AOT N(11),MAR_AOT N(11),APR_AOT N(11),MAY_AOT N(11),JUN_AOT N(11),JUL_AOT N(11),AUG_AOT N(11),SEP_AOT N(11),OCT_AOT N(11),NOV_AOT N(11),DEC_AOT N(11),SUM_TO_AOT N(17))
	     		 INDE ON F18 TAG YCO_1
	     		 INDE ON SUM_TOL DESC TAG YCO_2
	     		 INDE ON SUM_TO_AOT DESC TAG YCO_3
	     		 SET ORDER TO YCO_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F10 C(5),F11 C(8))
		             H='C10_TEMP'
		         ELSE
		              IF !USED('&H')
		 		   SELECT 0
		 	           USE (H) 
		 	      ELSE
		 		   SELECT H
		 	      ENDIF       		             
		         ENDIF	                  
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR	
		 	FOR  I = 1 TO 12
		 	          SELECT F10,F11 FROM &H ORDER BY F10 GROUP BY F10 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO
		 	                 SEEK C01_TEMP.F10
		 	                 IF !FOUND()
		 	                     SELECT YCO
		 	                     APPEND BLANK 
		 	                     REPLACE YCO.F18 WITH C01_TEMP.F10
		 	                     REPLACE YCO.F033 WITH C01_TEMP.F11
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 			 	
	      		         SELECT  F10,SUM(F07*F08),SUM(F08) FROM &H ORDER BY F10 GROUP BY F10 WHERE F03=ANY(SELE F01 FROM B01 WHERE F97$UKEF)  INTO CURSOR C10_1 
	     			 SELE C10_1
	      			 GO TOP
	      			 DO WHILE !EOF()
	             		        SELECT  YCO
	             		        SEEK C10_1.F10
		                  	IF   FOUND()
		                      	      T=FIELD(I+2)
		                      	      REPLACE  &T WITH C10_1.SUM_EXP_2
		                      	      REPLACE  SUM_TOL WITH YCO.SUM_TOL+C10_1.SUM_EXP_2
		                      	      T=FIELD(I+15)
		                      	      REPLACE  &T WITH C10_1.SUM_F08
		                      	      REPLACE  SUM_TO_AOT WITH YCO.SUM_TO_AOT+C10_1.SUM_F08
		                      ENDIF
		            	      SELE C10_1
		            	      SKIP
		                ENDDO      
		     		SELECT &H
		     		USE
		     		SELE C10_1
		     		USE
		     		S=VAL(RIGHT(H,6))+1
		    		H='C10'+ALLTRIM(STR(S))
		    		F=ALLTRIM(H)+'.DBF'
				IF FILE('&F')=.F.
				    CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F10 C(5),F11 C(8))
				    H='C10_TEMP'
				ELSE
			            IF !USED('&H')
			 	        SELECT 0
			 	        USE (H) 
			 	    ELSE
			 	        SELECT H
			 	    ENDIF  					    
				ENDIF	
	     		 ENDFOR
	     		 WAIT CLEAR
	     		 SELECT  YCO
	     	         DELE FOR SUM_TOL+SUM_TO_AOT=0
		     	 CALC SUM(YCO.SUM_TOL) TO TOLMOY        	      
		     	 IF TOLMOY=0
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	 ELSE
		     	      SELECT YCO
			      DO C99C		     	     
		     	 ENDIF	     	      
     	      CASE THISFORM.DATA_OPTION.VALUE=4
	     		CREATE CURSOR YCO;
	    		(F18 C(5),F033 C(8),JAN N(17,6),FEB N(17,6),MAR N(17,6),APR N(17,6),MAY N(17,6),JUN N(17,6),JUL N(17,6),AUG N(17,6),SEP N(17,6),OCT N(17,6),NOV N(17,6),DEC N(17,6),SUM_TOL N(17,6),;
	    		      JAN_AOT N(11),FEB_AOT N(11),MAR_AOT N(11),APR_AOT N(11),MAY_AOT N(11),JUN_AOT N(11),JUL_AOT N(11),AUG_AOT N(11),SEP_AOT N(11),OCT_AOT N(11),NOV_AOT N(11),DEC_AOT N(11),SUM_TO_AOT N(17))
	     		 INDE ON F18 TAG YCO_1
	     		 INDE ON SUM_TOL DESC TAG YCO_2
	     		 INDE ON SUM_TO_AOT DESC TAG YCO_3
	     		 SET ORDER TO YCO_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F12 C(5),F13 C(8))
		             H='C10_TEMP'
		         ELSE
		              IF !USED('&H')
		 		   SELECT 0
		 	           USE (H) 
		 	      ELSE
		 		   SELECT H
		 	      ENDIF       		             
		         ENDIF	                  
		        WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR	
		 	FOR  I = 1 TO 12
		 	          SELECT F12,F13 FROM &H ORDER BY F12 GROUP BY F12 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO
		 	                 SEEK C01_TEMP.F12
		 	                 IF !FOUND()
		 	                     SELECT YCO
		 	                     APPEND BLANK 
		 	                     REPLACE YCO.F18 WITH C01_TEMP.F12
		 	                     REPLACE YCO.F033 WITH C01_TEMP.F13
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 			 	
	      		         SELECT  F12,SUM(F07*F08),SUM(F08) FROM &H ORDER BY F12 GROUP BY F12 WHERE F03=ANY(SELE F01 FROM B01 WHERE F97$UKEF)  INTO CURSOR C10_1 
	     			 SELE C10_1
	      			 GO TOP
	      			 DO WHILE !EOF()
	             		        SELECT  YCO
	             		        SEEK C10_1.F12
		                  	IF   FOUND()
		                      	      T=FIELD(I+2)
		                      	      REPLACE  &T WITH C10_1.SUM_EXP_2
		                      	      REPLACE  SUM_TOL WITH YCO.SUM_TOL+C10_1.SUM_EXP_2
		                      	      T=FIELD(I+15)
		                      	      REPLACE  &T WITH C10_1.SUM_F08
		                      	      REPLACE  SUM_TO_AOT WITH YCO.SUM_TO_AOT+C10_1.SUM_F08
		                      ENDIF
		            	      SELE C10_1
		            	      SKIP
		                ENDDO      
		     		SELECT &H
		     		USE
		     		SELE C10_1
		     		USE
		     		S=VAL(RIGHT(H,6))+1
		    		H='C10'+ALLTRIM(STR(S))
		    		F=ALLTRIM(H)+'.DBF'
				IF FILE('&F')=.F.
				    CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F12 C(5),F13 C(8))
				    H='C10_TEMP'
				ELSE
			            IF !USED('&H')
			 	        SELECT 0
			 	        USE (H) 
			 	    ELSE
			 		 SELECT H
			 	    ENDIF  					    
				ENDIF	
	     		 ENDFOR
	     		 WAIT CLEAR
	     		 SELECT  YCO
	     	         DELE FOR SUM_TOL+SUM_TO_AOT=0
		     	 CALC SUM(YCO.SUM_TOL) TO TOLMOY        	      
		     	 IF TOLMOY=0
             	              =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	              THISFORM.MTH_LIST.SETFOCUS
                               RETURN		     	 
		     	 ELSE
		     	      SELECT YCO
			      DO C99D		     	     
		     	 ENDIF	          	      
     	      CASE THISFORM.DATA_OPTION.VALUE=5  
	     		CREATE CURSOR YCO;
	    		(F23 C(5),F231 C(8),JAN N(17,6),FEB N(17,6),MAR N(17,6),APR N(17,6),MAY N(17,6),JUN N(17,6),JUL N(17,6),AUG N(17,6),SEP N(17,6),OCT N(17,6),NOV N(17,6),DEC N(17,6),SUM_TOL N(17,6),;
	    		      JAN_AOT N(11),FEB_AOT N(11),MAR_AOT N(11),APR_AOT N(11),MAY_AOT N(11),JUN_AOT N(11),JUL_AOT N(11),AUG_AOT N(11),SEP_AOT N(11),OCT_AOT N(11),NOV_AOT N(11),DEC_AOT N(11),SUM_TO_AOT N(17))
	     		 INDE ON F23 TAG YCO_1
	     		 INDE ON SUM_TOL DESC TAG YCO_2
	     		 INDE ON SUM_TO_AOT DESC TAG YCO_3
	     		 SET ORDER TO YCO_1
	     		 S=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'01'
	      		 H='C10'+ALLTRIM(S)
	      		 F=ALLTRIM(H)+'.DBF'
		       	 IF FILE('&F')=.F.
		       	     CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F14 C(5),F15 C(8))
		             H='C10_TEMP'
		         ELSE
		              IF !USED('&H')
		 		   SELECT 0
		 	           USE (H) 
		 	      ELSE
		 		   SELECT H
		 	      ENDIF       		             
		         ENDIF	      		
		         WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR		 
		 	 FOR  I = 1 TO 12
		 	          SELECT F14,F15 FROM &H ORDER BY F14 GROUP BY F14 INTO CURSOR C01_TEMP
		 	          SELECT C01_TEMP
		 	          GO TOP
		 	          DO WHILE !EOF()
		 	                 SELECT YCO
		 	                 SEEK C01_TEMP.F14
		 	                 IF !FOUND()
		 	                     SELECT YCO
		 	                     APPEND BLANK 
		 	                     REPLACE YCO.F23 WITH C01_TEMP.F14
		 	                     REPLACE YCO.F231 WITH C01_TEMP.F15
		 	                 ENDIF
		 	                 SELECT C01_TEMP
		 	                 SKIP
		 	          ENDDO
		 	          SELECT C01_TEMP
		 	          USE
		 	          ****** 			 	 
	      		          SELECT   F14,SUM(F07*F08),SUM(F08) FROM &H ORDER BY F14 GROUP BY F14 WHERE F03=ANY(SELE F01 FROM B01 WHERE F97$UKEF)  INTO CURSOR C10_1 
	     			  SELE C10_1
	      			  GO TOP
	      		          DO WHILE !EOF()
	             			 SELECT  YCO
	             		         SEEK C10_1.F14
	                  		 IF   FOUND()
	                      		       T=FIELD(I+2)
	                      		       REPLACE  &T WITH C10_1.SUM_EXP_2
	                      	               REPLACE  SUM_TOL WITH YCO.SUM_TOL+C10_1.SUM_EXP_2 
		                      	      T=FIELD(I+15)
		                      	      REPLACE  &T WITH C10_1.SUM_F08
		                      	      REPLACE  SUM_TO_AOT WITH YCO.SUM_TO_AOT+C10_1.SUM_F08	 	                      	                   		                       	   	                      	                       
	                  		  ENDIF          
	            			  SELE C10_1
	            			  SKIP
	     		           ENDDO      
		     		   SELECT &H
		     		   USE
		     		   SELE C10_1
		     		   USE
		     		   S=VAL(RIGHT(H,6))+1
		    		   H='C10'+ALLTRIM(STR(S))
		    		   F=ALLTRIM(H)+'.DBF'
				   IF FILE('&F')=.F.
				        CREATE CURSOR C10_TEMP (F02 C(5),F03 C(43),F07 N(14,6),F08 N(11),F14 C(5),F15 C(8))
				        H='C10_TEMP'
				   ELSE
			                IF !USED('&H')
			 		     SELECT 0
			 	             USE (H) 
			 	        ELSE
			 		     SELECT H
			 	        ENDIF  					        
				   ENDIF	
	     	       ENDFOR
	     	       WAIT CLEAR
	     	       SELECT  YCO
		       DELE FOR SUM_TOL+SUM_TO_AOT=0
		       CALC SUM(YCO.SUM_TOL) TO TOLMOY	      	      
		       IF TOLMOY=0
             	            =MESSAGEBOX('此年份無資料,請重新輸入!',0+64,'提示視窗')
            	            THISFORM.MTH_LIST.SETFOCUS
                             RETURN		     	 
		       ELSE
		     	     SELECT YCO
			     DO C99E		     	     
		      ENDIF	        	            
     ENDCASE      
    THISFORM.CMND2.CLICK	 
 ENDPROC      
 **************************      
 PROCEDURE CMND2.CLICK       
      IF !USED('A05')
          SELECT 0
          USE A05
      ELSE
         SELECT A05
      ENDIF
      SET ORDER TO 1
      SEEK sys_oper+'C99'
      IF FOUND()
        DELETE 
      ENDIF         
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
 ENDPROC      
ENDDEFINE



 
