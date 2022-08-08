           
			    					                        
                    IF THISFORM.OPTIONGROUP_CHOICE.VALUE=1	&&建立基本檔
                       SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='0' OR A20.F02='9' ORDER BY A20.F01 INTO CURSOR A20_TEMP
        		       SELECT A20_TEMP
		               GO TOP
		               DO WHILE !EOF()
		                  TableName=ALLTRIM(A20_TEMP.F01)	&&資料檔名
		                  IF FILE(TableName+'.DBF')=.F.
		                     SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                 ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                 SELECT A22_TEMP
			                 GO TOP
			                 IF _TALLY>0
			                    FieldName=ALLTRIM(A22_TEMP.F01)	&&欄位名稱
			                    TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
			                    CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
			                    SELECT A22_TEMP
			                    SKIP
			                    DO WHILE !EOF()		&&建立欄位
			                       FieldName=ALLTRIM(A22_TEMP.F01)
			                       TypeName=ALLTRIM(A22_TEMP.F02)
			                       ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
			                       SELECT A22_TEMP
			                       SKIP
			                    ENDDO 
			                    SELECT A22_TEMP
			                    USE
			                    SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
			                    SELECT A21_TEMP
			                    GO TOP
			                    DO WHILE !EOF()		&&建立索引
			                       THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+'  '+ALLTRIM(A20_TEMP.F07) 
			                       THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
			                       IndexName=ALLTRIM(A21_TEMP.F02)	&&索引名稱
			                       IndexKey=ALLTRIM(A21_TEMP.F03)	        &&索引內容
			                       IF A21_TEMP.F04='1'		&&順向排序
			                          SELECT &TableName
			                          INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
			                       ELSE					&&逆向排序
			                          SELECT &TableName
			                          INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
			                       ENDIF   
			                       SELECT A21_TEMP
			                       SKIP
			                    ENDDO
			                    SELECT &TableName
			                    USE
			                    F5=TableName+'.BAK'
			                    IF FILE('&F5')=.T.
			                       DELETE FILE  &F5
			                    ENDIF
			                    SELECT A21_TEMP
			                    USE 
			                 ENDIF
		                  ENDIF  
		                  SELECT A20_TEMP
		                  SKIP
                       ENDDO
		               SELECT A20_TEMP
		               USE       
		               		                IF !USED('C00')
		                    SELECT 0
		                    USE C00
		                ELSE
		                   SELECT C00
		                ENDIF   
		                SET ORDER TO 1
		                SEEK INT_011
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH INT_011
		                   REPLACE F02 WITH 1
		                   REPLACE F03 WITH DTCX(DATE())+' SET UP'
		                ENDIF
		                USE		                
		                IF !USED('D00')
		                   SELECT 0
		                   USE D00
		                ELSE
		                   SELECT D00
		                ENDIF   
		                SET ORDER TO 1
		                SEEK INT_011
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH INT_011
		                   REPLACE F02 WITH 1
		                   REPLACE F03 WITH DTCX(DATE())+'  SET UP'
		                ENDIF          
		                USE		                		                 
		                IF !USED('A09')
		                   SELECT 0
		                   USE A09
		                ELSE
		                   SELECT A09
		                ENDIF
		                SET ORDER TO 1
		                SEEK 'C000 '
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH 'C000'
		                   REPLACE F06 WITH '客戶基本主檔編碼'
		                   REPLACE F08 WITH 'C0000' 
		                ENDIF
		                SEEK 'D000 '
		                IF !FOUND()                    
		                   APPEND BLANK
		                   REPLACE F01 WITH 'D000'
		                   REPLACE F06 WITH '廠商基本主檔編碼'
		                   REPLACE F08 WITH 'V0000'            
		                ENDIF
		                SEEK 'E000 '
		                IF !FOUND()         
		                   APPEND BLANK
		                   REPLACE F01 WITH 'E000'
		                   REPLACE F06 WITH '協力廠商基本主檔編碼'
		                   REPLACE F08 WITH 'W0000'                  		                   
		                ENDIF  
		                USE 
		               =MESSAGEBOX('基本檔已建立完畢!',0+64,'提示訊息視窗')
		               THISFORM.CMND_TEMP.CLICK
		                ****
                    ELSE		&&建立月份檔
			           IF FILE('A23.DBF')<>.T.
                          =MESSAGEBOX('請先建立基本檔後在建立月份檔!',0+64,'提示訊息視窗') 
                          THISFORM.CMND_TEMP.CLICK    
      			       ELSE     
      			          IF !USED('A23')
			                 SELECT  0
			                 USE A23 INDE A23
			              ELSE 
			                 SELECT  A23
			              ENDIF
			              SET ORDER TO A231
			              GO TOP
			              IF CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01')<A23.F01 AND !EMPTY(A23.F01)	
                             =MESSAGEBOX('系統月份檔之起始日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故不得小於此日期,請重新輸入!',0+64,'提示訊息視窗') 
                             THISFORM.CMND_TEMP.CLICK
				             THISFORM.TXTF01.SETFOCUS
				             RETURN	 			                   
			              ENDIF
			              ***			             
			              SELECT A23
			              GO BOTTOM 
			              IF MONTH(A23.F01) = 12
			                 IF YEAR(A23.F01) + 1 <> YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))
			                    =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次年1月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                    THISFORM.CMND_TEMP.CLICK
				                THISFORM.TXTF01.SETFOCUS
				                RETURN	
				             ELSE
				                IF MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01')) <> 1
			                       =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次年1月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                       THISFORM.CMND_TEMP.CLICK
				                   THISFORM.TXTF01.SETFOCUS
				                   RETURN					                   
				                ENDIF
			                 ENDIF 			                 
			              ELSE
			                 IF YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))=YEAR(A23.F01) AND !EMPTY(A23.F01)
			                    IF (MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))-MONTH(A23.F01))<>1
			                        =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                        THISFORM.CMND_TEMP.CLICK
				                    THISFORM.TXTF01.SETFOCUS
				                    RETURN	
			                    ENDIF
			                 ELSE
			                    IF YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))=(YEAR(A23.F01)+1) AND !EMPTY(A23.F01)
			                       IF (MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))<>1 AND MONTH(A23.F01)<>12)=.F.
			                          =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                          THISFORM.CMND_TEMP.CLICK
				                      THISFORM.TXTF01.SETFOCUS
				                      RETURN	
			                       ENDIF		                    
			                    ELSE
			                       IF !EMPTY(A23.F01)
			                          =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                          THISFORM.CMND_TEMP.CLICK
				                      THISFORM.TXTF01.SETFOCUS
				                      RETURN
				                   ENDIF    
			                    ENDIF
			                 ENDIF   
			              ENDIF	             
			              *****
                          SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='1'  ORDER BY A20.F01 INTO CURSOR A20_TEMP		      
			              IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			                 DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			              ELSE
			                 MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			                 DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			              ENDIF   			      
			              DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		                  DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)
			                 SELECT A23
			                 SEEK DTOS(CTOD(DateFile+'/01'))
			                 IF !FOUND()
			                    APPEND BLANK
			                    REPLACE  F01 WITH CTOD(DateFile+'/01')
			                 ENDIF		
			                 ***	            
			                 SELECT A20_TEMP
			                 GO TOP
			                 DO WHILE !EOF()
			                    TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                    IF FILE(TableName+'.DBF')=.F.
			                       SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                       ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                       SELECT A22_TEMP
				                   GO TOP
				                   IF _TALLY>0
				                      FieldName=ALLTRIM(A22_TEMP.F01)    &&欄位名稱
				                      TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
				                      CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
				                      SELECT A22_TEMP
				                      SKIP
				                      DO WHILE !EOF()		&&建立欄位
				                         FieldName=ALLTRIM(A22_TEMP.F01)
				                         TypeName=ALLTRIM(A22_TEMP.F02)
				                         ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
				                         SELECT A22_TEMP
				                         SKIP
				                      ENDDO 
				                      SELECT A22_TEMP
				                      USE
				                      ***
				                      SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
				                      SELECT A21_TEMP
				                      GO TOP
				                      DO WHILE !EOF()		&&建立索引
				                         THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
				                         THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
				                         IndexName=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))		&&索引名稱
				                         IndexKey=ALLTRIM(A21_TEMP.F03)	       &&索引內容
				                         IF A21_TEMP.F04='1'		&&順向排序
				                            SELECT &TableName
				                            INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
				                         ELSE						&&逆向排序
				                            SELECT &TableName
				                            INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
				                         ENDIF   
				                         SELECT A21_TEMP
				                         SKIP
				                      ENDDO
				                      ***帶入各項暫存檔至該月份檔
							          IF  LEFT(TableName,3) = 'C13'  AND FILE('C1C.DBF') = .T.	&&應收帳款對帳單檔
							              IF !USED('C1C') 
								             SELE 0
								             USE C1C
								          ELSE
								             SELE C1C
								          ENDIF
							              SELECT &TableName
							              APPEND FROM C1C
							              SELECT C1C
							              DELETE ALL
							              USE
							          ENDIF
							          IF LEFT(TableName,3) = 'D19'  AND FILE('D1I.DBF') = .T.	&&應付帳款對帳單檔
							             IF !USED('D1I') 
								           SELE 0
								           USE D1I
								         ELSE
								           SELE D1I
								         ENDIF
							             SELECT &TableName
							             APPEND FROM D1I
							             SELECT D1I
							             DELETE ALL
							             USE
							          ENDIF	
							          IF LEFT(TableName,3) = 'P19'  AND FILE('P1I.DBF') = .T.	&&外購應付帳款對帳單檔
							             IF !USED('P1I') 
								            SELE 0
								            USE P1I
								         ELSE
								            SELE P1I
								         ENDIF
							             SELECT &TableName
							             APPEND FROM P1I
							             SELECT P1I
							             DELETE ALL
							             USE
							          ENDIF
							          IF LEFT(TableName,3) = 'K24'  AND FILE('K2D.DBF') = .T. 	&&發票表頭檔
							             IF !USED('K2D')
								            SELE 0
								            USE K2D
								         ELSE
								            SELE K2D
								         ENDIF
							             SELECT &TableName
							             APPEND FROM K2D
							             SELECT K2D
							             DELETE ALL
							             USE
							          ENDIF
							          IF LEFT(TableName,3) = 'K25' AND FILE('K2E.DBF') = .T.	&&發票明細檔
							             IF !USED('K2E') 
								            SELE 0
								            USE K2E
								         ELSE
								            SELE K2E
								         ENDIF
							             SELECT &TableName
							             APPEND FROM K2E
							             SELECT K2E
							             DELETE ALL
							             USE
							          ENDIF							       							       							       						       
				                      ***
				                      SELECT &TableName
				                      USE
					                  F5=TableName+'.BAK'
					                  IF FILE('&F5')=.T.
					                     DELETE FILE  &F5
					                  ENDIF				                              
				                      SELECT A21_TEMP
				                      USE 
				                   ENDIF
			                    ENDIF  
			                    SELECT A20_TEMP
			                    SKIP
	                         ENDDO
	                         IF (VAL(RIGHT(DateFile,2))+1)>12
			                    DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                 ELSE
			                    MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                    DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                 ENDIF  
			              ENDDO
			              SELECT A20_TEMP
			              USE 
***********************************************建立奇月檔
                         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='3'  ORDER BY A20.F01 INTO CURSOR A20_TEMP		      
			             IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			                 DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			             ELSE
			                 MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			                 DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			             ENDIF   	
			             DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
			             DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)     			                
		                    SELECT A20_TEMP
			                GO TOP
			                DO WHILE !EOF()
			                   IF VAL(ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6)))%2<>0 
			                      TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                   
			                      IF FILE(TableName+'.DBF')=.F.
			                         SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                         ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                         SELECT A22_TEMP
				                     GO TOP
				                     IF _TALLY>0
				                        FieldName=ALLTRIM(A22_TEMP.F01)    &&欄位名稱
				                        TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
				                        CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
				                        SELECT A22_TEMP
				                        SKIP
				                        DO WHILE !EOF()		&&建立欄位
				                           FieldName=ALLTRIM(A22_TEMP.F01)
				                           TypeName=ALLTRIM(A22_TEMP.F02)
				                           ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
				                           SELECT A22_TEMP
				                           SKIP
				                        ENDDO 
				                        SELECT A22_TEMP
				                        USE
				                        ***
				                        SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
				                        SELECT A21_TEMP
				                        GO TOP
				                        DO WHILE !EOF()		&&建立索引
				                           THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
				                           THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
				                           IndexName=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))		&&索引名稱
				                           IndexKey=ALLTRIM(A21_TEMP.F03)	       &&索引內容
				                           IF A21_TEMP.F04='1'		&&順向排序
				                              SELECT &TableName
				                              INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
				                           ELSE						&&逆向排序
				                              SELECT &TableName
				                              INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
				                           ENDIF   
				                           SELECT A21_TEMP
				                           SKIP
				                        ENDDO			  				                        
			                            SELECT &TableName
				                        USE
					                    F5=TableName+'.BAK'
					                    IF FILE('&F5')=.T.
					                       DELETE FILE  &F5
					                    ENDIF				                              
				                        SELECT A21_TEMP
				                        USE 				                        
				                     ENDIF  
				                  ENDIF
				               ENDIF   	 
			                   SELECT A20_TEMP
			                   SKIP
	                        ENDDO	                            	
	                        IF (VAL(RIGHT(DateFile,2))+1)>12
			                   DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                ELSE
			                   MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                   DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                ENDIF  
			             ENDDO
			             SELECT A20_TEMP
			             USE                   			                	                
			               
***************************************************			              
			              SELECT A23
			              USE     				              		              
			              =MESSAGEBOX('月份檔已建立完畢!',0+64,'提示訊息視窗') 
			              THISFORM.RELEASE				      
			            ENDIF
			       **********
                     ENDIF       
                 