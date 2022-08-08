***************A27.資料檔處理****************
CLOSE ALL
CLEAR 
*************************建立表單*************************    
A27FORM=CREATEOBJECT("TKA27")
A27FORM.SHOW  
DEFINE CLASS TKA27 AS FORM
       CAPTION='A27.資料檔處理'
       AUTOCENTER=.T.
       CONTROLBOX=.F.
       BORDERSTYLE=3  
       MAXBUTTON=.F.
       MINBUTTON=.F.
       MOVABLE=.F.
       CLOSABLE=.F.
       CONTROLCOUNT=57
       FONTSIZE=INT_015*9
       HEIGHT=INT_015*300
       SHOWTIPS=.T.
       SHOWWINDOW=1
       WIDTH=INT_015*425
       WINDOWTYPE=1
       NAME='TKA27'
*********************************
       ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
           LEFT=INT_015*24,;
           TOP=INT_015*24,;
           HEIGHT=INT_015*37,;
           WIDTH=INT_015*130,;
           FONTSIZE=INT_015*11,;
           CAPTION='\<N.產生新月份檔',;
           TOOLTIPTEXT='產生新月份檔!快速鍵->ALT+N',;
           NAME='CMND1'
       ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
           LEFT=INT_015*24,;
           TOP=INT_015*96,;
           HEIGHT=INT_015*37,;
           WIDTH=INT_015*130,;
           FONTSIZE=INT_015*11,;
           CAPTION='\<D.移除月份資料',;
           TOOLTIPTEXT='移除月份資料!快速鍵->ALT+D',;
           NAME='CMND2'      
       ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
           LEFT=INT_015*264,;
           TOP=INT_015*24,;
           HEIGHT=INT_015*37,;
           WIDTH=INT_015*130,;
           FONTSIZE=INT_015*11,;
           CAPTION='\<I.安裝基本資料',;
           TOOLTIPTEXT='安裝基本資料!快速鍵->ALT+I',;
           NAME='CMND3'  
       ADD OBJECT CMND4 AS COMMANDBUTTON WITH;
           LEFT=INT_015*264,;
           TOP=INT_015*96,;
           HEIGHT=INT_015*37,;
           WIDTH=INT_015*130,;
           FONTSIZE=INT_015*11,;
           CAPTION='\<B.移除基本資料',;
           TOOLTIPTEXT='移除基本資料!快速鍵->ALT+B',;
           NAME='CMND4' 
       ADD OBJECT CMND5 AS COMMANDBUTTON WITH;
           LEFT=INT_015*264,;
           TOP=INT_015*180,;
           HEIGHT=INT_015*37,;
           WIDTH=INT_015*85,;
           FONTSIZE=INT_015*11,;
           CAPTION='\<X.離開',;
           TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
           NAME='CMND5'                 
       ADD OBJECT LABEL1 AS LABEL WITH;
           LEFT=INT_015*180,;
           TOP=INT_015*200,;
           AUTOSIZE=.T.,;
           CAPTION=''
       ADD OBJECT LABEL2 AS LABEL WITH;      
           LEFT=INT_015*180,;
           TOP=INT_015*240,;
           AUTOSIZE=.T.,;
           CAPTION=''
 *********************               
       PROCEDURE INIT
            THISFORM.SETALL('HEIGHT',17*INT_015,'LABEL')
            THISFORM.SETALL('FONTSIZE',INT_015*11,'LABEL')         
       ENDPROC     
*******************產生新月份檔
       PROCEDURE CMND1.CLICK
       		MTH=RIGHT(LEFT(DTOC(DATE()),5),2)
            THISFORM.CMND1.ENABLED=.F. 
            THISFORM.CMND2.ENABLED=.F. 
            THISFORM.CMND3.ENABLED=.F. 
            THISFORM.CMND4.ENABLED=.F. 
            THISFORM.CMND5.ENABLED=.F.        
            OPEN DATABASE TKDB
            SELE 0
            IF !USED('A23')
               USE A23
            ELSE
               SELE A23
            ENDIF      
            SET ORDER TO 1
            LOCATE FOR LEFT(DTOS(F01),6)=LEFT(DTOS(DATE()),6)
            IF !FOUND()
               APPEND BLANK
               REPL F01 WITH DATE()-DAY(DATE())+1
            ENDIF
                  SELE 0
                  IF !USED('A20')
                     USE A20
                  ELSE
                     SELE A20
                  ENDIF      
                  SET ORDER TO 1
                  SELE 0
                  IF !USED('A22')
                     USE A22
                  ELSE
                     SELE A22
                  ENDIF      
                  SET ORDER TO 1
                  SELE * FROM A20 WHERE F02='1' INTO CURSOR A20A ORDER BY F01
                  MICAU1=VAL(LEFT(DTOS(DATE()),4))                             &&先從TKDB.DBC移出三個月前的月份檔
                  MICAU2=VAL(SUBSTR(DTOS(DATE()),5,2))
                  MICAU=IIF(MICAU2<4,ALLTRIM(STR(MICAU1-1))+ALLTRIM(STR(MICAU2+9)),ALLTRIM(STR(MICAU1))+PADL(ALLTRIM(STR(MICAU2-3)),2,'0'))
                  GO TOP
                  DO WHILE !EOF()
                     HK=ALLTRIM(F01)+MICAU 
                     IF INDBC(HK,'TABLE')      
                        THISFORM.LABEL1.CAPTION='先移除'+HK
                        REMOVE TABLE &HK 
                     ENDIF   
                     SKIP
                  ENDDO            
                  SELE A22.F01,A22.F02,A22.F03,A22.F04,A20A.F01,A20A.F07 FROM A22,A20A WHERE A22.F05=A20A.F01; &&再產生月份檔
                       INTO CURSOR AAA ORDER BY A20A.F01,A22.F01
                  SELE AAA
                  GO TOP
                  DO WHILE !EOF()
                     FILE_NAME=ALLTRIM(F01_B)
                     FIELD_NAME=''
                     DO WHILE F01_B=FILE_NAME
                        FIELD_NAME=FIELD_NAME+F01_A+' '+F02+'('+ALLTRIM(STR(F03))+IIF(F02='N',','+ALLTRIM(STR(F04)),'')+')'+','
                        SKIP
                     ENDDO
                     CREATE_FIELD=LEFT(FIELD_NAME,LEN(FIELD_NAME)-1)
                     CREATE_FILE=ALLTRIM(FILE_NAME)+LEFT(DTOS(DATE()),6)
                     IF !INDBC(CREATE_FILE,'TABLE')
                        THISFORM.LABEL1.CAPTION='產生'+CREATE_FILE
                        CREATE TABLE &CREATE_FILE;
                        (&CREATE_FIELD)
                        S=1
                        DO WHILE S<10
                           L='F'+ALLTRIM(STR(S+9))
                           SELE F01 FROM A22 WHERE ALLTRIM(A22.F05)=LEFT(CREATE_FILE,3) AND LEN(ALLTRIM(A22.&L))>0 INTO ARRAY TMP1 ORDER BY A22.&L
                           IF _TALLY>0
                              KEY_PASS=''
                              FOR I= 1 TO ALEN(TMP1)
*!*	                                  KEY_PASS=KEY_PASS+IIF(TYPE('&TMP1(I)')='D','DTOS('+TMP1(I)+')',TMP1(I))+'+'
					KEY_PASS=KEY_PASS+IIF(TYPE('&TMP1(I)')='D','DTOS('+TMP1(I)+')',IIF(TYPE('&TMP1(I)')='N','ALLTRIM(STR('+TMP1(I)+'))',TMP1(I)))+'+'
                              ENDFOR    
                              KEY_ACT=LEFT(KEY_PASS,LEN(KEY_PASS)-1)        
                              THISFORM.LABEL2.CAPTION=KEY_ACT
                              INDE ON &KEY_ACT TAG (STUFF(CREATE_FILE,4,0,ALLTRIM( STR(S))))
                           ENDIF 
                           RELEASE ALL LIKE TMP1*
                           S=S+1                                          
                        ENDDO     
                     ENDIF   
                     SELE AAA
                  ENDDO      
                  CLOSE TABLE ALL    
                  C13='C13'+LEFT(DTOS(DATE()),6)   &&應收帳款對帳單檔
                  IF !USED('&C13')
                     SELE 0
                     USE (C13) ALIA C13
                  ELSE
                     SELE C13
                  ENDIF
                  SET ORDER TO 1
                  D19='D19'+LEFT(DTOS(DATE()),6)   &&應付帳款對帳單檔
                  IF !USED('&D19')
                     SELE 0
                     USE (D19) ALIA D19
                  ELSE
                     SELE D19
                  ENDIF
                  SET ORDER TO 1
                  P19='P19'+LEFT(DTOS(DATE()),6)   &&外購應付帳款對帳單檔
                  IF !USED('&P19')
                     SELE 0
                     USE (P19) ALIA P19
                  ELSE
                     SELE P19
                  ENDIF
                  SET ORDER TO 1
                  K24='K24'+LEFT(DTOS(DATE()),6)   &&發票表頭檔
                  IF !USED('&K24')
                     SELE 0
                     USE (K24) ALIA K24
                  ELSE
                     SELE K24
                  ENDIF
                  SET ORDER TO 1                    
                  K25='K25'+LEFT(DTOS(DATE()),6)   &&發票明細檔
                  IF !USED('&K25')
                     SELE 0
                     USE (K25) ALIA K25
                  ELSE
                     SELE K25
                  ENDIF
                  SET ORDER TO 1    
                  IF !USED('C1C')
                     SELE 0
                     USE C1C
                  ELSE
                     SELE C1C
                  ENDIF
                  SET ORDER TO 1
                  IF !USED('D1I')
                     SELE 0
                     USE D1I
                  ELSE 
                     SELE D1I
                  ENDIF
                  SET ORDER TO 1
                  IF !USED('P1I')
                     SELE 0
                     USE P1I
                  ELSE
                    SELE P1I
                  ENDIF
                  SET ORDER TO 1
                  IF !USED('K2D')
                     SELE 0
                     USE K2D
                  ELSE
                     SELE K2D
                  ENDIF
                  SET ORDER TO 1                  
                  IF !USED('K2E')
                     SELE 0
                     USE K2E
                  ELSE
                     SELE K2E
                  ENDIF
                  SET ORDER TO 1
                  SELE C13
                  APPEND FROM C1C
                  SELE D19
                  APPEND FROM D1I
                  SELE P19
                  APPEND FROM P1I
                  SELE K24
                  APPEND FROM K2D
                  SELE K25
                  APPEND FROM K2E
                  SELE C1C
                  DELE ALL
                  SELE D1I 
                  DELE ALL
                  SELE P1I
                  DELE ALL
                  SELE K2D                                
                  DELE ALL                   
                  SELE K2E                                
                  DELE ALL   
                  CLOSE TABLE ALL
                  =MESSAGEBOX(MTH+'月份之月份檔已產生完畢!',0+48,'提示訊息視窗')                           	               
            THISFORM.LABEL1.CAPTION=''
            THISFORM.LABEL2.CAPTION=''
            THISFORM.CMND1.ENABLED=.T. 
            THISFORM.CMND2.ENABLED=.T. 
            THISFORM.CMND3.ENABLED=.T. 
            THISFORM.CMND4.ENABLED=.T. 
            THISFORM.CMND5.ENABLED=.T.           
       ENDPROC
 ******************* 移除月份資料
       PROCEDURE CMND2.CLICK
       		 MTH=RIGHT(LEFT(DTOC(DATE()),5),2)
             THISFORM.LABEL2.CAPTION='' 
             THISFORM.CMND1.ENABLED=.F. 
             THISFORM.CMND2.ENABLED=.F. 
             THISFORM.CMND3.ENABLED=.F. 
             THISFORM.CMND4.ENABLED=.F. 
             THISFORM.CMND5.ENABLED=.F.       
             OPEN DATABASE TKDB
             SELE * FROM A20 WHERE F02='1' ORDER BY F01 INTO CURSOR A20A
             GO TOP
             DO WHILE !EOF()
                HK=ALLTRIM(F01)+LEFT(DTOS(DATE()),6)
                IF INDBC(HK,'TABLE')
                   THISFORM.LABEL1.CAPTION=HK
                   REMOVE TABLE &HK 
                ENDIF   
                SKIP
             ENDDO
             =MESSAGEBOX(MTH+'月份之月份檔已移除完畢!',0+64,'提示訊息視窗')
             THISFORM.LABEL1.CAPTION=''
             THISFORM.LABEL2.CAPTION=''
             THISFORM.CMND1.ENABLED=.T. 
             THISFORM.CMND2.ENABLED=.T. 
             THISFORM.CMND3.ENABLED=.T. 
             THISFORM.CMND4.ENABLED=.T. 
             THISFORM.CMND5.ENABLED=.T.         
       ENDPROC
******************* 安裝基本資料     
       PROCEDURE CMND3.CLICK
             THISFORM.CMND1.ENABLED=.F. 
             THISFORM.CMND2.ENABLED=.F. 
             THISFORM.CMND3.ENABLED=.F. 
             THISFORM.CMND4.ENABLED=.F. 
             THISFORM.CMND5.ENABLED=.F.        
             OPEN DATABASE TKDB
             SELE 0
             IF !USED('A20')
                USE A20
             ELSE
                SELE A20
             ENDIF      
             SET ORDER TO A20
             SELE 0
             IF !USED('A22')
                USE A22
             ELSE
                SELE A22
             ENDIF      
             SET ORDER TO A22
             SELE *FROM A20 WHERE F02='0' OR F02='9' INTO CURSOR A20A ORDER BY F01
             SELE A22.F01,A22.F02,A22.F03,A22.F04,A20A.F01,A20A.F07 FROM A22,A20A WHERE A22.F05=A20A.F01;
             INTO CURSOR AAA ORDER BY A20A.F01,A22.F01
             SELE AAA
             GO TOP
             DO WHILE !EOF()
                FILE_NAME=F01_B
                FIELD_NAME=''
                DO WHILE F01_B=FILE_NAME
                   FIELD_NAME=FIELD_NAME+F01_A+' '+F02+'('+ALLTRIM(STR(F03))+IIF(F02='N',','+ALLTRIM(STR(F04)),'')+')'+','
                   SKIP
                ENDDO
                CREATE_FIELD=LEFT(FIELD_NAME,LEN(FIELD_NAME)-1)
                CREATE_FILE=ALLTRIM(FILE_NAME)
                IF !INDBC(CREATE_FILE,'TABLE')
                   THISFORM.LABEL1.CAPTION=CREATE_FILE+'  '+IIF(SEEK(CREATE_FILE,'A20'),A20.F07,'') 
                   CREATE TABLE &CREATE_FILE; 
                   (&CREATE_FIELD)
                   S=1
                   DO WHILE S<10
                      L='F'+ALLTRIM(STR(S+9))
                      SELE F01 FROM A22 WHERE ALLTRIM(A22.F05)=LEFT(CREATE_FILE,3) AND LEN(ALLTRIM(A22.&L))>0 INTO ARRAY TMP1 ORDER BY A22.&L
                      IF _TALLY>0
                         KEY_PASS=''
                         V=0
                         FOR I= 1 TO ALEN(TMP1)
*!*	                             KEY_PASS=KEY_PASS+IIF(TYPE('&TMP1(I)')='D','DTOS('+TMP1(I)+')',TMP1(I))+'+'
				KEY_PASS=KEY_PASS+IIF(TYPE('&TMP1(I)')='D','DTOS('+TMP1(I)+')',IIF(TYPE('&TMP1(I)')='N','ALLTRIM(STR('+TMP1(I)+'))',TMP1(I)))+'+'
                             V=V+VAL(IIF(SEEK(FILE_NAME+TMP1(I),'A22'),A22.&L,''))
                         ENDFOR    
                         KEY_ACT=LEFT(KEY_PASS,LEN(KEY_PASS)-1)
                         THISFORM.LABEL2.CAPTION=KEY_ACT+'  '+IIF(SEEK(KEY_ACT,'A20'),A20.F07,'')
                         IF V>0
                            INDE ON &KEY_ACT TAG (CREATE_FILE+ALLTRIM(STR(S))) ASCEND
                         ELSE
                            INDE ON &KEY_ACT TAG (CREATE_FILE+ALLTRIM(STR(S))) DESCEND
                         ENDIF   
                      ENDIF   
                      RELEASE ALL LIKE TMP1*
                      S=S+1             
                   ENDDO
                ENDIF   
                SELE AAA
             ENDDO      
             IF !USED('C00')
                SELE 0
                USE C00
             ELSE
                SELE C00
             ENDIF
             SET ORDER TO 1        
             SEEK INT_011
             IF !FOUND()           
                APPEND BLANK
                REPLACE F01 WITH INT_011
                REPLACE F02 WITH 1
             ENDIF      
             IF !USED('D00')
                SELE 0
                USE D00
             ELSE
                SELE D00
             ENDIF
             SET ORDER TO 1                   
             SEEK INT_011
             IF !FOUND()
                APPEND BLANK
                REPLACE F01 WITH INT_011
                REPLACE F02 WITH 1
             ENDIF   
             IF !USED('A09')
                SELE 0
                USE A09
             ELSE
                SELE A09
             ENDIF
             SET ORDER TO 1
             SEEK 'C000 '
             IF !FOUND()
                APPEND BLANK
                REPL F01 WITH 'C000'
                REPL F06 WITH '客戶主檔編號'
                REPL F08 WITH 'C0000'
             ENDIF
             SEEK 'D000 '   
             IF !FOUND()
                APPEND BLANK
                REPL F01 WITH 'D000'
                REPL F06 WITH '廠商主檔編號'
                REPL F08 WITH 'V0000'
             ENDIF
             SEEK 'E000 '
             IF !FOUND()   
                APPEND BLANK
                REPL F01 WITH 'E000'
                REPL F06 WITH '協力廠商主檔編號'
                REPL F08 WITH 'W0000' 
             ENDIF   
             CLOSE ALL
             =MESSAGEBOX('基本資料已安裝完畢!',0+64,'提示訊息視窗')
             THISFORM.LABEL1.CAPTION=''
             THISFORM.LABEL2.CAPTION=''
             THISFORM.CMND1.ENABLED=.T. 
             THISFORM.CMND2.ENABLED=.T. 
             THISFORM.CMND3.ENABLED=.T. 
             THISFORM.CMND4.ENABLED=.T. 
             THISFORM.CMND5.ENABLED=.T.             
       ENDPROC
 ******************* 移除基本資料     
       PROCEDURE CMND4.CLICK 
             THISFORM.LABEL2.CAPTION='' 
             THISFORM.CMND1.ENABLED=.F. 
             THISFORM.CMND2.ENABLED=.F. 
             THISFORM.CMND3.ENABLED=.F. 
             THISFORM.CMND4.ENABLED=.F. 
             THISFORM.CMND5.ENABLED=.F.     
             if messagebox('基本資料將會完全被刪除!您是否確定?',4+48+256,'請確認') = 6            	             
                OPEN DATABASE TKDB
                SELE 0
                IF !USED('A20')
                   USE A20
                ELSE
                   SELE A20
                ENDIF      
                SET ORDER TO A20
                SELE 0
                IF !USED('A22')
                   USE A22
                ELSE
                   SELE A22
                ENDIF      
                SET ORDER TO A22
                SELE * FROM A20 WHERE F02='0' ORDER BY F01 INTO CURSOR A20A
                GO TOP
                DO WHILE !EOF()
                   HK=ALLTRIM(F01)
                   IF INDBC(HK,'TABLE')
                      THISFORM.LABEL1.CAPTION=HK+'  '+IIF(SEEK(HK,'A20'),A20.F07,'')
                      REMOVE TABLE &HK 
                   ENDIF   
                   SKIP
                ENDDO
                =MESSAGEBOX('基本資料已移除完畢!',0+64,'提示訊息視窗')
             ENDIF   
             THISFORM.LABEL1.CAPTION=''
             THISFORM.LABEL2.CAPTION=''
             THISFORM.CMND1.ENABLED=.T. 
             THISFORM.CMND2.ENABLED=.T. 
             THISFORM.CMND3.ENABLED=.T.
             THISFORM.CMND4.ENABLED=.T. 
             THISFORM.CMND5.ENABLED=.T.            
        ENDPROC
 ******************* 離開          
        PROCEDURE CMND5.CLICK       
             CLOSE TABLE ALL
             THISFORM.RELEASE 
             RETURN  
        ENDPROC      
ENDDEFINE
*******************