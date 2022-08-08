****程式名稱 B27 庫存相關單據重算
CLOSE ALL
CLEAR
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'B27'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('B01')
   SELE 0
   USE B01
ELSE
   SELE B01
ENDIF
SET ORDER TO 1   

IF !USED('C01')
   SELE 0
   USE C01
ELSE
   SELE C01   
ENDIF
SET ORDER TO 1
IF !USED('D01')
   SELE 0
   USE D01
ELSE 
   SELE D01
ENDIF
SET ORDER TO 1
IF !USED('A14')
   SELE 0
   USE A14
ELSE
   SELE A14
ENDIF      
SET ORDER TO 1  
*********
TKB27FORM=CREATEOBJECT("TKB27")
TKB27FORM.SHOW       
DEFINE CLASS TKB27 AS FORM
  CAPTION='B27.庫存相關單據重算'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*400
  WIDTH=INT_015*450
  AUTOCENTER=.T.
  CLOSABLE=.F.    
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKB27'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*250,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入'
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='單據編號'     
  ADD OBJECT LBL21 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION='計算料號'
  ADD OBJECT LBL31 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT LBL4 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION='寫入筆數'            
  ADD OBJECT LBL41 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION=''              
  ADD OBJECT TXT1 AS TEXTBOX WITH;
      LEFT=INT_015*160,;
      TOP=INT_015*245,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*80,;  
      MAXLENGTH=7,;
      NAME='TXT1'    
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*120,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認後開始執行重算單據',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<C.放棄',;
      TOOLTIPTEXT='放棄執行,並離開此畫面!快速鍵->ALT+C'
      NAME='CMND2'
 ********     
      PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.LBL1.CAPTION=THISFORM.LBL1.CAPTION+IIF(INT_012=1,'民國年/月','西洋年/月')
	     THISFORM.TXT1.VALUE=LEFT(DTCX(DATE()),7-INT_012)
      ENDPROC
 **********    
      PROCEDURE CMND1.CLICK
         IF EMPTY(CTOD(THISFORM.TXT1.VALUE+'/01'))
            =MESSAGEBOX('無此月份資料!',0+48,'提示訊息視窗')
            THISFORM.TXT1.SETFOCUS
            RETURN         
         ENDIF
         IF !USED('A23')
            SELE 0
            USE A23
         ELSE
           SELE A23   
         ENDIF
         SET ORDER TO 1         
         SEEK DTOS(CTOD(THISFORM.TXT1.VALUE+'/01'))
         IF !FOUND()
            =MESSAGEBOX('無此月份資料!',0+48,'提示訊息視窗')
            THISFORM.TXT1.SETFOCUS
            RETURN
         ENDIF

         B25='B25'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)         
         IF !USED('&B25')
            SELE 0
            USE (B25) ALIA B25
         ELSE
            SELE B25
         ENDIF
         SET ORDER TO 1        
*         SELE F01,F02,F12,F15 FROM B25 ORDER BY F01,F02 INTO CURSOR B25BK
         SELE B25            
         DELE ALL
         B26='B26'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B26')
            SELE 0
            USE (B26) ALIA B26
         ELSE
            SELE B26
         ENDIF
         SET ORDER TO 1                                 
         DELE ALL
         THISFORM.LBL21.CAPTION='B02 進貨單'
         B02='B02'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B02')
            SELE 0
            USE (B02) ALIA B02
         ELSE
            SELE B02
         ENDIF
         SET ORDER TO 1          
         GO TOP
         S=0
         DO WHILE !EOF()
            IF B02.F13<>'Y' AND B02.F10='2' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK
               REPL F01 WITH B02.F03
               REPL F02 WITH B02.F05
               REPL F03 WITH B02.F02
               REPL F04 WITH (B02.F04+B02.F15)
               REPL F06 WITH 'A'
               REPL F07 WITH B02.F01
               REPL F08 WITH '採'+B02.F09+IIF(!EMPTY(B02.F02),'請'+B02.F08,'')
               S=S+1               
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))                              
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))  
             ENDIF  
             SELE B02
             SKIP         
         ENDDO 
         SELE F05,F03,SUM(F04),SUM(F15) FROM B02 WHERE F10='2' AND F13<>'Y' AND;
         F03=ANY(SELE F01 FROM B01 WHERE !EMPTY(F98)) GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B0B
         SELE B02
         USE
         SELE B0B
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0B.F05+B0B.F03
            IF FOUND()
               REPL F04 WITH F04+(B0B.SUM_F04+B0B.SUM_F15)
            ELSE
               APPEND BLANK
               REPL F01 WITH B0B.F05
               REPL F02 WITH B0B.F03
               REPL F04 WITH (B0B.SUM_F04+B0B.SUM_F15)               
            ENDIF       
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))            
            SELE B0B
            SKIP
         ENDDO
         SELE B0B
         USE         
         THISFORM.LBL21.CAPTION='H02 進貨檢驗單'
         H02='H02'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&H02')
            SELE 0
            USE (H02) ALIA H02
         ELSE
            SELE H02
         ENDIF
         SET ORDER TO 1          
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F13='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK
               REPL F01 WITH H02.F03
               REPL F02 WITH H02.F17
               REPL F03 WITH H02.F02
               REPL F04 WITH H02.F09
               REPL F06 WITH 'N'
               REPL F07 WITH H02.F01
               REPL F08 WITH '進'+H02.F05+'採'+H02.F06
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))            
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                                       
            ENDIF                  
            SELE H02
            SKIP         
         ENDDO 
         SELE F17,F03,SUM(F09) FROM H02 WHERE F13='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR H0B
         SELE H02
         USE
         SELE H0B
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK H0B.F17+H0B.F03
            IF FOUND()
               REPL F04 WITH F04+H0B.SUM_F09
            ELSE
               APPEND BLANK
               REPL F01 WITH H0B.F17
               REPL F02 WITH H0B.F03
               REPL F04 WITH H0B.SUM_F09
            ENDIF       
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))            
            SELE H0B
            SKIP
         ENDDO
         SELE H0B
         USE                  
         
         THISFORM.LBL21.CAPTION='P12 外購進貨單'
         P12='P12'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&P12')
            SELE 0
            USE (P12) ALIA P12
         ELSE
            SELE P12
         ENDIF
         SET ORDER TO 1          
         GO TOP
         S=0
         DO WHILE !EOF()
            IF P12.F13<>'Y' AND P12.F10='2' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK
               REPL F01 WITH P12.F03
               REPL F02 WITH P12.F05
               REPL F03 WITH P12.F02
               REPL F04 WITH (P12.F04+P12.F15)
               REPL F06 WITH 'A'
               REPL F07 WITH P12.F01
               REPL F08 WITH '採'+P12.F09+IIF(!EMPTY(P12.F08),'請'+P12.F08,'')
               S=S+1               
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))      
               SELECT B25
               SEEK P12.F05+P12.F03
               IF FOUND()
                  REPLACE F04 WITH F04+(P12.F04+P12.F15)
               ELSE
                  APPEND BLANK
                  REPL F01 WITH P12.F05
                  REPL F02 WITH P12.F03
                  REPL F04 WITH (P12.F04+P12.F15)                 
               ENDIF         
               THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))      
             ENDIF  
             SELE P12
             SKIP         
         ENDDO 
*!*	         SELE F05,F03,SUM(F04),SUM(F15) FROM P12 WHERE F10='2' AND F13<>'Y' AND;
*!*	         F03=ANY(SELE F01 FROM B01 WHERE !EMPTY(F98)) GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR P1B        
*!*	         SELE P12
*!*	         USE
*!*	         SELE P1B
*!*	         GO TOP
*!*	         S=0
*!*	         DO WHILE !EOF()
*!*	            SELE B25
*!*	            SEEK P1B.F05+P1B.F03
*!*	            IF FOUND()
*!*	               REPL F04 WITH F04+(P1B.SUM_F04+P1B.SUM_F15)
*!*	            ELSE
*!*	               APPEND BLANK
*!*	               REPL F01 WITH P1B.F05
*!*	               REPL F02 WITH P1B.F03
*!*	               REPL F04 WITH (P1B.SUM_F04+P1B.SUM_F15)               
*!*	            ENDIF       
*!*	            S=S+1
*!*	            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))
*!*	            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))            
*!*	            SELE P1B
*!*	            SKIP
*!*	         ENDDO
*!*	         SELE P1B
*!*	         USE                           
         THISFORM.LBL21.CAPTION='B03 進貨退出單'
         B03='B03'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B03')
            SELE 0
            USE (B03) ALIA B03
         ELSE
            SELE B03
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F10='2' AND F24='1' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))  &&非折讓及虛擬料號才做運算
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK
               REPL F01 WITH B03.F03
               REPL F02 WITH B03.F05
               REPL F03 WITH B03.F02
               REPL F04 WITH (B03.F04+B03.F17)*(-1)
               REPL F06 WITH 'B'
               REPL F07 WITH B03.F01
               REPL F08 WITH '採'+B03.F09    
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B03
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04),SUM(F17) FROM B03 WHERE F10='2' AND F24='1' AND;
         F03=ANY(SELE F01 FROM B01 WHERE !EMPTY(F98))  GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B0C
         SELE B03
         USE 
         SELE B0C
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0C.F05+B0C.F03
            IF FOUND()
               REPL F05 WITH F05+B0C.SUM_F04+B0C.SUM_F17
            ELSE
               APPEND BLANK
               REPL F01 WITH B0C.F05
               REPL F02 WITH B0C.F03
               REPL F05 WITH B0C.SUM_F04+B0C.SUM_F17
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F05))   
            SELE B0C
            SKIP 
         ENDDO
         SELE B0C
         USE
         THISFORM.LBL21.CAPTION='B04 出貨單'         
         B04='B04'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B04')
            SELE 0
            USE (B04) ALIA B04
         ELSE
            SELE B04
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F10='2' AND !EMPTY(IIF(SEEK(F79,'B01'),B01.F98,''))
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK
               REPL F01 WITH B04.F79
               REPL F02 WITH B04.F05
               REPL F03 WITH B04.F02
               REPL F04 WITH (B04.F04+B04.F13)*(-1)
               REPL F06 WITH 'C'
               REPL F07 WITH B04.F01
               REPL F08 WITH '訂'+B04.F07+' '+IIF(SEEK(B04.F06,'C01'),C01.F05,'')                 
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
             ENDIF  
            SELE B04
            SKIP                               
         ENDDO   
         SELE F05,F79,SUM(F04),SUM(F13) FROM B04 WHERE F10='2' AND;
         F79=ANY(SELE F01 FROM B01 WHERE !EMPTY(F98)) GROUP BY F05,F79 ORDER BY F05,F79 INTO CURSOR B0D
         SELE B04
         USE 
         SELE B0D
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0D.F05+B0D.F79
            IF FOUND()
               REPL F06 WITH F06+B0D.SUM_F04+B0D.SUM_F13
            ELSE
               APPEND BLANK
               REPL F01 WITH B0D.F05
               REPL F02 WITH B0D.F79
               REPL F06 WITH B0D.SUM_F04+B0D.SUM_F13
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F06))   
            SELE B0D
            SKIP 
         ENDDO
         SELE B0D
         USE                  
         THISFORM.LBL21.CAPTION='B05 出貨退回單'         
         B05='B05'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B05')
            SELE 0
            USE (B05) ALIA B05
         ELSE
            SELE B05
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F12='2' AND F24='1' AND !EMPTY(IIF(SEEK(F79,'B01'),B01.F98,''))
               SELE B26                                            &&庫存異動紀錄             
               APPEND BLANK
               REPL F01 WITH B05.F79
               REPL F02 WITH B05.F05
               REPL F03 WITH B05.F02
               REPL F04 WITH (B05.F04+B05.F15)
               REPL F06 WITH 'D'
               REPL F07 WITH B05.F01
               REPL F08 WITH '訂'+B05.F11                              
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B05
            SKIP                               
         ENDDO   
         SELE F05,F79,SUM(F04),SUM(F15) FROM B05 WHERE F12='2' AND F24='1' AND;
         F79=ANY(SELE F01 FROM B01 WHERE !EMPTY(F98)) GROUP BY F05,F79 ORDER BY F05,F79 INTO CURSOR B0E
         SELE B05
         USE 
         SELE B0E
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0E.F05+B0E.F79
            IF FOUND()
               REPL F07 WITH F07+B0E.SUM_F04+B0E.SUM_F15
            ELSE
               APPEND BLANK
               REPL F01 WITH B0E.F05
               REPL F02 WITH B0E.F79
               REPL F07 WITH B0E.SUM_F04+B0E.SUM_F15
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F07))   
            SELE B0E
            SKIP 
         ENDDO
         SELE B0E
         USE                           
         THISFORM.LBL21.CAPTION='B06 移轉單'         
         B06='B06'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B06')
            SELE 0
            USE (B06) ALIA B06
         ELSE
            SELE B06
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F07='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&轉出  
               REPL F01 WITH B06.F79                             
               REPL F02 WITH B06.F05
               REPL F03 WITH B06.F02
               REPL F04 WITH B06.F04*(-1)
               REPL F06 WITH 'E'
               REPL F07 WITH B06.F01
               REPL F08 WITH '移轉至'+B06.F10+IIF(SEEK(B06.F10,'A14'),A14.F02,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                
               APPEND BLANK                                               &&轉入   
               REPL F01 WITH B06.F79
               REPL F02 WITH B06.F10
               REPL F03 WITH B06.F02
               REPL F04 WITH B06.F04
               REPL F06 WITH 'E'
               REPL F07 WITH B06.F01
               REPL F08 WITH '從'+B06.F05+IIF(SEEK(B06.F05,'A14'),A14.F02,'')+'轉入'
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B06
            SKIP                               
         ENDDO   
         SELE F05,F79,SUM(F04) FROM B06 WHERE F07='2' GROUP BY F05,F79 ORDER BY F05,F79 INTO CURSOR B0F_1
         SELE B0F_1
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0F_1.F05+B0F_1.F79
            IF FOUND()
               REPL F09 WITH F09+B0F_1.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B0F_1.F05
               REPL F02 WITH B0F_1.F79
               REPL F09 WITH B0F_1.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F09))   
            SELE B0F_1
            SKIP 
         ENDDO
         SELE B0F_1
         USE 
         SELE F10,F79,SUM(F04) FROM B06 WHERE F07='2' GROUP BY F10,F79 ORDER BY F05,F79 INTO CURSOR B0F_2
         SELE B0F_2
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0F_2.F10+B0F_2.F79
            IF FOUND()
               REPL F08 WITH F08+B0F_2.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B0F_2.F10
               REPL F02 WITH B0F_2.F79
               REPL F08 WITH B0F_2.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F08))   
            SELE B0F_2
            SKIP 
         ENDDO
         SELE B06
         USE                   
         SELE B0F_2
         USE                                             
         THISFORM.LBL21.CAPTION='B07 委外入庫單' 
         B07='B07'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B07')
            SELE 0
            USE (B07) ALIA B07
         ELSE
            SELE B07
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F10='2' AND F13='1'  AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))          
               SELE B26
               APPEND BLANK                                               &&轉入   
               REPLACE  F01 WITH B07.F03
               REPLACE  F02 WITH B07.F05
               REPLACE  F03 WITH B07.F02
               REPLACE  F04 WITH B07.F04
               REPLACE  F06 WITH 'M'
               REPLACE  F07 WITH B07.F01
               REPLACE  F08 WITH '委'+B07.F15+'製'+B07.F09
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B07
            SKIP                               
         ENDDO        
         SELECT F14,F03,SUM(F04) FROM B07 WHERE F10='2' AND F13='1' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,'')) GROUP BY F14,F13 ORDER BY F14,F03 INTO CURSOR B0G_1
         SELE B0G_1
         GO TOP
         S=0
         DO WHILE !EOF()                              &&產出及轉出
            SELE B25
            SEEK B0G_1.F14+B0G_1.F03
            IF !FOUND()
                APPEND BLANK
                REPL F01 WITH B0G_1.F14
                REPL F02 WITH B0G_1.F03            
            ENDIF
            REPL F09 WITH F09+B0G_1.SUM_F04
            REPL F10 WITH F10+B0G_1.SUM_F04
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F10))   
            SELE B0G_1
            SKIP 
         ENDDO
         SELE B0G_1
         USE                                                    
         SELE F05,F03,SUM(F04) FROM B07 WHERE F10='2' AND F13='1' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B0G_2
         SELE B07
         USE 
         SELE B0G_2                          &&轉入
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0G_2.F05+B0G_2.F03
            IF !FOUND()
               APPEND BLANK
               REPL F01 WITH B0G_2.F05
               REPL F02 WITH B0G_2.F03            
             ENDIF
             REPL F08 WITH F08+B0G_2.SUM_F04
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F08))   
            SELE B0G_2
            SKIP 
         ENDDO
         SELE B0G_2
         USE                                                                                    
         THISFORM.LBL21.CAPTION='B08 製令領料單'         
         B08='B08'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B08')
            SELE 0
            USE (B08) ALIA B08
         ELSE
            SELE B08
         ENDIF
         SET ORDER TO 1               
         SELECT F01,F02,F04,F07,F08,SUM(F12) FROM B08 WHERE F13='2' GROUP BY F01,F08 ORDER BY F01,F08,F02 INTO CURSOR B0H
          SELECT B0H              
         GO TOP
         S=0
         DO WHILE !EOF()
                FOR I=1 TO 2
                    SELE B26                                            &&庫存異動紀錄
                    APPEND BLANK                                               &&耗用  
                    REPLACE  F01 WITH B0H.F08                               
                    REPLACE  F02 WITH IIF(I=1,B0H.F07,B0H.F04)
                    REPLACE  F03 WITH B0H.F02
                    REPLACE  F04 WITH B0H.SUM_F12*(-1)^(I)
                    REPLACE  F06 WITH 'P'
                    REPLACE  F07 WITH B0H.F01
                    IF I=1
                        REPLACE  F08 WITH IIF(SEEK(B0H.F04,'A14'),'發料給'+A14.F02+'生產','')    
                    ELSE    
                        REPLACE  F08 WITH IIF(SEEK(B0H.F07,'A14'),'向'+A14.F02+'領用上線','')                                    
                    ENDIF    
                    S=S+1
                    THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                    THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
                ENDFOR    
                SELE B0H
                SKIP                               
         ENDDO   
         SELE B0H
         USE          
         SELE B08
         GO TOP
         S=0
         DO WHILE !EOF()
                 IF F13='2'   
                      SELE B25                                             &&分部門庫存月報表 
                      SEEK B08.F07+B08.F08
                      IF  !FOUND()                                                     &&發料部門報表
                           APPEND BLANK
                           REPLACE  F01 WITH B08.F07
                           REPLACE  F02 WITH B08.F08
                      ENDIF   
                       REPLACE  F09 WITH F09+B08.F12 
                      SEEK B08.F04+B08.F08
                      IF  !FOUND()                                                     &&領料部門報表
                           APPEND BLANK
                           REPLACE  F01 WITH B08.F04
                           REPLACE  F02 WITH B08.F08
                      ENDIF   
                       REPLACE  F08 WITH F08+B08.F12 
                      S=S+1
                     THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
                     THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F14))   
                 ENDIF    
                 SELE B08
                SKIP 
         ENDDO
         SELE B08
         USE                                     
         THISFORM.LBL21.CAPTION='B09 盤差調整單'         
         B09='B09'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B09')
            SELE 0
            USE (B09) ALIA B09
         ELSE
            SELE B09
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F11='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                              
               REPL F01 WITH B09.F03                               
               REPL F02 WITH B09.F05
               REPL F03 WITH B09.F02
               REPL F04 WITH B09.F04
               REPL F06 WITH 'H'
               REPL F07 WITH B09.F01
               REPL F08 WITH B09.F07
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B09
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B09 WHERE F11='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B0I
         SELE B09
         USE 
         SELE B0I
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B0I.F05+B0I.F03
            IF FOUND()
               REPL F13 WITH F13+B0I.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B0I.F05
               REPL F02 WITH B0I.F03
               REPL F13 WITH B0I.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F13))   
            SELE B0I
            SKIP 
         ENDDO
         SELE B0I
         USE              
         THISFORM.LBL21.CAPTION='B10 報廢單'         
         B10='B10'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B10')
            SELE 0
            USE (B10) ALIA B10
         ELSE
            SELE B10
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F09='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&報廢
               REPL F01 WITH B10.F03                               
               REPL F02 WITH B10.F05
               REPL F03 WITH B10.F02
               REPL F04 WITH B10.F04*(-1)
               REPL F06 WITH 'I'
               REPL F07 WITH B10.F01
               REPL F08 WITH B10.F07
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B10
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B10 WHERE F09='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B1Z
         SELE B10
         USE 
         SELE B1Z
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B1Z.F05+B1Z.F03
            IF FOUND()
               REPL F14 WITH F14+B1Z.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B1Z.F05
               REPL F02 WITH B1Z.F03
               REPL F14 WITH B1Z.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F14))   
            SELE B1Z
            SKIP 
         ENDDO
         SELE B1Z
         USE              
         THISFORM.LBL21.CAPTION='生產耗料紀錄'                  
         E14='E14'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)                  &&製令用料耗用紀錄         
         IF !USED('&E14')
            SELE 0
            USE (E14) ALIA E14
         ELSE
            SELE E14
         ENDIF
         SET ORDER TO 3 
         GO TOP          
         DO WHILE !EOF()
                SELECT B26
                APPEND BLANK
                REPLACE F01 WITH E14.F02
                REPLACE F02 WITH E14.F07
                REPLACE F03 WITH E14.F03
                REPLACE F04 WITH E14.F05*(-1)
                REPLACE F06 WITH SUBSTR(E14.F04,2,1)
                REPLACE F07 WITH E14.F04
                REPLACE F08 WITH E14.F01
                S=S+1
                THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                    
                SELECT B25
                SEEK E14.F07+E14.F02
                IF !FOUND()
                    APPEND BLANK
                    REPLACE F01 WITH E14.F07
                    REPLACE F02 WITH E14.F02
                ENDIF 
                REPLACE F11 WITH F11+E14.F05
                THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
                THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F11))                   
                SELECT E14
                SKIP                
         ENDDO
         THISFORM.LBL21.CAPTION='B33 製令用料耗損單'         
         B16='B16'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B16')
            SELE 0
            USE (B16) ALIA B16
         ELSE
            SELE B16
         ENDIF
         SET ORDER TO 1               
         SELECT F01,F02,F04,F07,F08,SUM(F12) FROM B16 WHERE F13='2' GROUP BY F01,F08 ORDER BY F01,F08,F02 INTO CURSOR B16F
          SELECT B16F              
         GO TOP
         S=0
         DO WHILE !EOF()
                    SELE B26                                            &&庫存異動紀錄
                    APPEND BLANK                                               &&耗用  
                    REPLACE  F01 WITH B16F.F08                               
                    REPLACE  F02 WITH B16F.F07
                    REPLACE  F03 WITH B16F.F02
                    REPLACE  F04 WITH B16F.SUM_F12*(-1)
                    REPLACE  F06 WITH 'K'
                    REPLACE  F07 WITH B16F.F01
                    REPLACE  F08 WITH IIF(SEEK(B16F.F04,'A14'),'發料給'+A14.F02+'生產','')                    
                    S=S+1
                    THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                    THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
                SELE B16F
                SKIP                               
         ENDDO   
         SELE B16F
         USE          
         SELE B16
         GO TOP
         S=0
         DO WHILE !EOF()
                 IF F13='2'
                     SELECT E14
                     SEEK B16.F04+B16.F05+B16.F08
                     IF FOUND()
                         UKEF=F08
                     ELSE
                         UKEF=.F.
                     ENDIF             
                    SELE B25                                             &&分部門庫存月報表 
                        SEEK B16.F07+B16.F08                      &&轉出              
                        IF  !FOUND()                                                     &&轉出部門報表
                             APPEND BLANK
                             REPLACE  F01 WITH B16.F07
                             REPLACE  F02 WITH B16.F08
                        ENDIF   
                    IF B16.F07<>B16.F04                           &&若耗料與發料為不同部門                            
                        REPLACE F09 WITH F09+B16.F12
                        SEEK B16.F04+B16.F08              &&轉入及耗損
                        IF !FOUND()
                            APPEND BLANK
                             REPLACE F01 WITH B16.F04
                             REPLACE F02 WITH B16.F08
                        ENDIF
                        REPLACE F08 WITH F08+B16.F12
                    ENDIF    
                    IF UKEF=.T.         &&若是客提物料則記入轉出數量
                       REPLACE  F09 WITH F09+B16.F12 
                    ELSE                    &&若非客提物料則記入產耗     
                       REPLACE F14 WITH F14+B16.F12 
                    ENDIF                                                                                          
                      S=S+1
                     THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
                     THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F14))   
                 ENDIF    
                 SELE B16
                SKIP 
         ENDDO
         SELE B16
         USE                            
         SELECT E14
         USE
         THISFORM.LBL21.CAPTION='B34 退回入倉單'         
         B18='B18'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B18')
             SELECT  0
             USE (B18) ALIAS   B18
         ELSE
             SELECT  B18
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
                IF F10='2'
                    SELE B26                                            &&庫存異動紀錄
                    APPEND BLANK                                               &&轉入  
                    REPLACE  F01 WITH B18.F07                    
                    REPLACE  F02 WITH B18.F05                               
                    REPLACE  F03 WITH B18.F02
                    REPLACE  F04 WITH B18.F09
                    REPLACE  F06 WITH 'L'
                    REPLACE  F07 WITH B18.F01
                    S=S+1
                   THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                   THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               ENDIF   
              SELECT  B18
              SKIP                               
         ENDDO   
         SELECT  F05,F07,SUM(F09) FROM B18 WHERE F10='2' GROUP BY F05,F07 ORDER BY F05,F07 INTO CURSOR B1H
         SELECT  B18
         USE 
         SELE B1H
         GO TOP
         S=0
         DO WHILE !EOF()
            SELECT  B25
            SEEK B1H.F05+B1H.F07
            IF FOUND()
                REPLACE  F11 WITH F11+B1H.SUM_F09*(-1)
            ELSE
               APPEND BLANK
               REPLACE  F01 WITH B1H.F05
               REPLACE  F02 WITH B1H.F07
               REPLACE  F11 WITH B1H.SUM_F09*(-1)
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F11))   
            SELECT  B1H
            SKIP 
         ENDDO
         SELECT  B1H
         USE                                    
         THISFORM.LBL21.CAPTION='B40 額外生產耗料單'         
         B40='B40'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B40')
            SELE 0
            USE (B40) ALIA B40
         ELSE
            SELE B40
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F11='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&產耗
               REPL F01 WITH B40.F03                               
               REPL F02 WITH B40.F05
               REPL F03 WITH B40.F02
               REPL F04 WITH B40.F04*(-1)
               REPL F06 WITH 'Q'
               REPL F07 WITH B40.F01
               REPL F08 WITH B40.F07
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               SELECT B25
               SEEK B40.F05+B40.F03
               IF !FOUND()
                    APPEND BLANK
                    REPLACE F01 WITH B40.F05
                    REPLACE F02 WITH B40.F03
               ENDIF
               IF B40.F07='N'
                   REPLACE F11 WITH F11+B40.F04
               ELSE
                   REPLACE F09 WITH F09+B40.F04          
                ENDIF   
                THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
                THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F14))                   
            ENDIF   
            SELE B40
            SKIP                               
         ENDDO   
          SELECT B40
          USE
         THISFORM.LBL21.CAPTION='B55 完工入庫單'         
         B55='B55'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B55')
            SELE 0
            USE (B55) ALIA B55
         ELSE
            SELE B55
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F13='2' AND !EMPTY(IIF(SEEK(F08,'B01'),B01.F98,''))
               SELE B26
               APPEND BLANK                                               &&轉入   
               REPLACE  F01 WITH B55.F08
               REPLACE  F02 WITH B55.F04
               REPLACE  F03 WITH B55.F02
               REPLACE  F04 WITH B55.F12
               REPLACE  F06 WITH 'T'
               REPLACE  F07 WITH B55.F01   
               REPLACE  F08 WITH '製'+B55.F05               
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))   
               IF B55.F09='2'   &&若為重工
                   SELE B26
                   APPEND BLANK                                               &&轉出   
                   REPLACE  F01 WITH B55.F08
                   REPLACE  F02 WITH B55.F07
                   REPLACE  F03 WITH B55.F02
                   REPLACE  F04 WITH B55.F12*(-1)
                   REPLACE  F06 WITH 'T'
                   REPLACE  F07 WITH B55.F01   
                   REPLACE  F08 WITH '製'+B55.F05+'重工'               
                   S=S+1
                  THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                  THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                       
               ENDIF 
            ENDIF   
            SELE B55
            SKIP                               
         ENDDO   
         SELE F07,F08,SUM(F12) FROM B55 WHERE F13='2' AND F09<>'2'  AND !EMPTY(IIF(SEEK(F08,'B01'),B01.F98,''))  GROUP BY F07,F08 ORDER BY F07,F08 INTO CURSOR B5E_1
         SELE B5E_1
         GO TOP
         S=0
         DO WHILE !EOF()                              &&產出及轉出
            SELE B25
            SEEK B5E_1.F07+B5E_1.F08
            IF FOUND()
               REPL F09 WITH F09+B5E_1.SUM_F12
               REPL F10 WITH F10+B5E_1.SUM_F12
            ELSE
               APPEND BLANK
               REPL F01 WITH B5E_1.F07
               REPL F02 WITH B5E_1.F08
               REPL F09 WITH B5E_1.SUM_F12
               REPL F10 WITH B5E_1.SUM_F12
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F10))   
            SELE B5E_1
            SKIP 
         ENDDO
         SELE B5E_1
         USE                                                    
         SELE F07,F08,SUM(F12) FROM B55 WHERE F13='2' AND F09='2' AND !EMPTY(IIF(SEEK(F08,'B01'),B01.F98,'')) GROUP BY F07,F08 ORDER BY F07,F08 INTO CURSOR B5E_3
         SELE B5E_3
         GO TOP
         S=0
         DO WHILE !EOF()                              &&重工轉出
            SELE B25
            SEEK B5E_3.F07+B5E_3.F08
            IF !FOUND()               
               APPEND BLANK
               REPL F01 WITH B5E_3.F07
               REPL F02 WITH B5E_3.F08            
            ENDIF
            REPLACE F09 WITH F09+B5E_3.SUM_F12  &&轉出
            REPLACE F10 WITH F10+B5E_3.SUM_F12  &&產出
            REPLACE F11 WITH F11+B5E_3.SUM_F12  &&產耗 
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F10))   
            SELE B5E_3
            SKIP 
         ENDDO
         SELE B5E_3
         USE                                                             
         SELE F04,F08,SUM(F12) FROM B55 WHERE F13='2' AND !EMPTY(IIF(SEEK(F08,'B01'),B01.F98,'')) GROUP BY F04,F08 ORDER BY F04,F08 INTO CURSOR B5E_2
         SELE B55
         USE 
         SELE B5E_2                          &&轉入
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B5E_2.F04+B5E_2.F08
            IF FOUND()
               REPL F08 WITH F08+B5E_2.SUM_F12
            ELSE
               APPEND BLANK
               REPL F01 WITH B5E_2.F04
               REPL F02 WITH B5E_2.F08
               REPL F08 WITH B5E_2.SUM_F12
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F08))   
            SELE B5E_2
            SKIP 
         ENDDO
         SELE B5E_2
         USE                      
                                                          
         THISFORM.LBL21.CAPTION='B56 製品退出單' 
         B56='B56'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B56')
            SELE 0
            USE (B56) ALIA B56
         ELSE
            SELE B56
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F10='2' AND F13='1' AND F24<>'2' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))          
               SELE B26
               APPEND BLANK                                               &&轉出   
               REPLACE  F01 WITH B56.F03
               REPLACE  F02 WITH B56.F05
               REPLACE  F03 WITH B56.F02
               REPLACE  F04 WITH B56.F04*(-1)
               REPLACE  F06 WITH 'R'
               REPLACE  F07 WITH B56.F01
               REPLACE  F08 WITH '製退'+B56.F09+'重'+B56.F25
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               IF LEN(ALLTRIM(B56.F14))=4   &&若為自製
                  APPEND BLANK                                               &&轉入   
                  REPLACE  F01 WITH B56.F03
                  REPLACE  F02 WITH B56.F14
                  REPLACE  F03 WITH B56.F02
                  REPLACE  F04 WITH B56.F04 
                  REPLACE  F06 WITH 'R'
                  REPLACE  F07 WITH B56.F01
                  REPLACE  F08 WITH '製退'+B56.F09+'重'+B56.F25
                  S=S+1
                  THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
                  THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                      
               ENDIF
            ENDIF   
            SELE B56
            SKIP                               
         ENDDO        
         SELECT F14,F03,SUM(F04) FROM B56 WHERE F10='2' AND F13='1' AND F24<>'2' AND LEN(ALLTRIM(F14))=4 AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,'')) GROUP BY F14,F13 ORDER BY F14,F03 INTO CURSOR B5F_1
         SELE B5F_1
         GO TOP
         S=0
         DO WHILE !EOF()                              &&轉入
            SELE B25
            SEEK B5F_1.F14+B5F_1.F03
            IF !FOUND()
                APPEND BLANK
                REPL F01 WITH B5F_1.F14
                REPL F02 WITH B5F_1.F03            
            ENDIF
            REPLACE  F08 WITH F08+B5F_1.SUM_F04
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F08))   
            SELE B5F_1
            SKIP 
         ENDDO
         SELE B5F_1
         USE                                                    
         SELE F05,F03,SUM(F04) FROM B56 WHERE F10='2' AND F13='1' AND F24<>'2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B5F_2
         SELE B56
         USE 
         SELE B5F_2                          &&轉出
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B5F_2.F05+B5F_2.F03
            IF !FOUND()
               APPEND BLANK
               REPL F01 WITH B5F_2.F05
               REPL F02 WITH B5F_2.F03            
             ENDIF
             REPLACE  F09 WITH F09+B5F_2.SUM_F04
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F09))   
            SELE B5F_2
            SKIP 
         ENDDO
         SELE B5F_2
         USE                                              
         THISFORM.LBL21.CAPTION='B61 借入單'         
         B61='B61'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B61')
            SELE 0
            USE (B61) ALIA B61
         ELSE
            SELE B61
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F12='2'
               SELE B26                                            
               APPEND BLANK                   &&庫存異動紀錄---收料部門借入
               REPL F01 WITH B61.F03
               REPL F02 WITH B61.F05
               REPL F03 WITH B61.F02
               REPL F04 WITH B61.F04
               REPL F06 WITH 'U'
               REPL F07 WITH B61.F01
               REPL F08 WITH '借自'+SPACE(2)+B61.F10+SPACE(1)+IIF(SEEK(B61.F10,'D01'),D01.F04,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               APPEND BLANK                  &&庫存異動紀錄---借出入倉借出
               REPL F01 WITH B61.F03
               REPL F02 WITH LEFT(INT_190,5)
               REPL F03 WITH B61.F02
               REPL F04 WITH B61.F04*(-1)
               REPL F06 WITH 'U'
               REPL F07 WITH B61.F01
               REPL F08 WITH '借自'+SPACE(2)+B61.F10+SPACE(1)+IIF(SEEK(B61.F10,'D01'),D01.F04,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                   
            ENDIF   
            SELE B61
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B61 WHERE F12='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B6A
         SELE B61
         USE 
         SELE B6A
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B6A.F05+B6A.F03                               &&借入部門
            IF FOUND()
               REPL F18 WITH F18+B6A.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B6A.F05
               REPL F02 WITH B6A.F03
               REPL F18 WITH B6A.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F18))   
            SEEK LEFT(INT_190,5)+B6A.F03                               &&借出入倉借出
            IF FOUND()
               REPL F20 WITH F20+B6A.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH LEFT(INT_190,5)
               REPL F02 WITH B6A.F03
               REPL F20 WITH B6A.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F20))               
            SELE B6A
            SKIP 
         ENDDO
         SELE B6A
         USE                                                             
         THISFORM.LBL21.CAPTION='B62 借入還出單'         
         B62='B62'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B62')
            SELE 0
            USE (B62) ALIA B62
         ELSE
            SELE B62
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F12='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                             &&借料部門歸還廠商
               REPL F01 WITH B62.F03
               REPL F02 WITH B62.F05
               REPL F03 WITH B62.F02
               REPL F04 WITH (B62.F04)*(-1)
               REPL F06 WITH 'V'
               REPL F07 WITH B62.F01
               REPL F08 WITH '還出 '+B62.F08+' '+IIF(SEEK(B62.F10,'D01'),D01.F04,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               APPEND BLANK                                             &&借出入倉收料
               REPL F01 WITH B62.F03
               REPL F02 WITH LEFT(INT_190,5)
               REPL F03 WITH B62.F02
               REPL F04 WITH (B62.F04)
               REPL F06 WITH 'V'
               REPL F07 WITH B62.F01
               REPL F08 WITH '還出 '+B62.F08+' '+IIF(SEEK(B62.F10,'D01'),D01.F04,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                   
            ENDIF   
            SELE B62
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B62 WHERE F12='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B6B
         SELE B62
         USE 
         SELE B6B
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK B6B.F05+B6B.F03                         &&借料部門還料
            IF FOUND()
               REPL F19 WITH F19+B6B.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B6B.F05
               REPL F02 WITH B6B.F03
               REPL F19 WITH B6B.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F19))   
            SEEK LEFT(INT_190,5)+B6B.F03                          &&借出入倉收料
            IF FOUND()
               REPL F21 WITH F21+B6B.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH LEFT(INT_190,5)
               REPL F02 WITH B6B.F03
               REPL F21 WITH B6B.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F21))               
            SELE B6B
            SKIP 
         ENDDO
         SELE B6B
         USE                                                                                         
         THISFORM.LBL21.CAPTION='B63 借出單'         
         B63='B63'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B63')
            SELE 0
            USE (B63) ALIA B63
         ELSE
            SELE B63
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F12='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&借出部門
               REPL F01 WITH B63.F03
               REPL F02 WITH B63.F05
               REPL F03 WITH B63.F02
               REPL F04 WITH (B63.F04)*(-1)
               REPL F06 WITH 'W'
               REPL F07 WITH B63.F01
               REPL F08 WITH '借至'+SPACE(2)+B63.F10+SPACE(1)+IIF(SEEK(B63.F10,'C01'),C01.F05, '')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               APPEND BLANK                                               &&借出入倉借入
               REPL F01 WITH B63.F03
               REPL F02 WITH LEFT(INT_190,5)
               REPL F03 WITH B63.F02
               REPL F04 WITH (B63.F04)
               REPL F06 WITH 'W'
               REPL F07 WITH B63.F01
               REPL F08 WITH '借至'+SPACE(2)+B63.F10+SPACE(1)+IIF(SEEK(B63.F10,'C01'),C01.F05, '')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                   
            ENDIF   
            SELE B63
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B63 WHERE F12='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B6C
         SELE B63
         USE 
         SELE B6C
         GO TOP
         S=0
         DO WHILE !EOF()                     &&庫存月報表
            SELE B25
            SEEK B6C.F05+B6C.F03                &&借出部門
            IF FOUND()
               REPL F20 WITH F20+B6C.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B6C.F05
               REPL F02 WITH B6C.F03
               REPL F20 WITH B6C.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F20))   
            SEEK LEFT(INT_190,5)+B6C.F03                &&借出入倉借入
            IF FOUND()
               REPL F18 WITH F18+B6C.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH LEFT(INT_190,5)
               REPL F02 WITH B6C.F03
               REPL F18 WITH B6C.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F18))               
            SELE B6C
            SKIP 
         ENDDO
         SELE B6C
         USE                                                                                                     
         THISFORM.LBL21.CAPTION='B64 借出還入單'         
         B64='B64'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&B64')
            SELE 0
            USE (B64) ALIA B64
         ELSE
            SELE B64
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F12='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                              &&收料部門
               REPL F01 WITH B64.F03
               REPL F02 WITH B64.F05
               REPL F03 WITH B64.F02
               REPL F04 WITH (B64.F04)
               REPL F06 WITH 'X'
               REPL F07 WITH B64.F01
               REPL F08 WITH '還入 '+B64.F08+' '+IIF(SEEK(B64.F10,'C01'),C01.F05, '')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
               APPEND BLANK                                              &&借出入倉還料
               REPL F01 WITH B64.F03
               REPL F02 WITH LEFT(INT_190,5)
               REPL F03 WITH B64.F02
               REPL F04 WITH (B64.F04)*(-1)
               REPL F06 WITH 'X'
               REPL F07 WITH B64.F01
               REPL F08 WITH '還入 '+B64.F08+' '+IIF(SEEK(B64.F10,'C01'),C01.F05, '')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))                   
            ENDIF   
            SELE B64
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM B64 WHERE F12='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR B6D
         SELE B64
         USE 
         SELE B6D
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25                          &&庫存月報表
            SEEK B6D.F05+B6D.F03                    &&收料部門
            IF FOUND()
               REPL F21 WITH F21+B6D.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH B6D.F05
               REPL F02 WITH B6D.F03
               REPL F21 WITH B6D.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F21))   
            SEEK LEFT(INT_190,5)+B6D.F03                    &&借出入倉還料
            IF FOUND()
               REPL F19 WITH F19+B6D.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH LEFT(INT_190,5)
               REPL F02 WITH B6D.F03
               REPL F19 WITH B6D.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F19))               
            SELE B6D
            SKIP 
         ENDDO
         SELE B6D
         USE                      
             
         THISFORM.LBL21.CAPTION='C38 樣品領用單'         
         C38='C38'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&C38')
            SELE 0
            USE (C38) ALIA C38
         ELSE
            SELE C38
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F07='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&損耗  
               REPL F01 WITH C38.F79                              
               REPL F02 WITH C38.F05
               REPL F03 WITH C38.F02
               REPL F04 WITH C38.F04*(-1)
               REPL F06 WITH 'Z'
               REPL F07 WITH C38.F01           
               REPL F08 WITH C38.F10+SPACE(2)+IIF(SEEK(C38.F10,'C01'),C01.F05,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE C38
            SKIP                               
         ENDDO   
         SELE F05,F79,SUM(F04) FROM C38 WHERE F07='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR C3H
         SELE C38
         USE 
         SELE C3H
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK C3H.F05+C3H.F79
            IF FOUND()
               REPL F17 WITH F17+C3H.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH C3H.F05
               REPL F02 WITH C3H.F79
               REPL F17 WITH C3H.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F17))   
            SELE C3H
            SKIP 
         ENDDO
         SELE C3H
         USE                                                                                                               
         
         THISFORM.LBL21.CAPTION='C40 客供品進貨單'         
         C40='C40'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&C40')
            SELE 0
            USE (C40) ALIA C40
         ELSE
            SELE C40
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F09='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&損耗  
               REPL F01 WITH C40.F03                               
               REPL F02 WITH C40.F06
               REPL F03 WITH C40.F02
               REPL F04 WITH C40.F04
               REPL F06 WITH 'G'
               REPL F07 WITH C40.F01           
               REPL F08 WITH C40.F05+SPACE(2)+IIF(SEEK(C40.F05,'C01'),C01.F05,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE C40
            SKIP                               
         ENDDO   
         SELE F06,F03,SUM(F04) FROM C40 WHERE F09='2' GROUP BY F06,F03 ORDER BY F06,F03 INTO CURSOR C4J
         SELE C40
         USE 
         SELE C4J
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK C4J.F06+C4J.F03
            IF FOUND()
               REPL F04 WITH F04+C4J.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH C4J.F06
               REPL F02 WITH C4J.F03
               REPL F04 WITH C4J.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))   
            SELE C4J
            SKIP 
         ENDDO
         SELE C4J
         USE                                                      
         
         THISFORM.LBL21.CAPTION='C41 客供品退貨單'         
         C41='C41'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&C41')
            SELE 0
            USE (C41) ALIA C41
         ELSE
            SELE C41
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F09='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&進退 
               REPL F01 WITH C41.F03                               
               REPL F02 WITH C41.F05
               REPL F03 WITH C41.F02
               REPL F04 WITH C41.F04*(-1)
               REPL F06 WITH 'J'
               REPL F07 WITH C41.F01           
               REPL F08 WITH C41.F06+SPACE(2)+IIF(SEEK(C41.F06,'C01'),C01.F05,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE C41
            SKIP                               
         ENDDO   
         SELE F05,F03,SUM(F04) FROM C41 WHERE F09='2' GROUP BY F05,F03 ORDER BY F05,F03 INTO CURSOR C4A
         SELE C41
         USE 
         SELE C4A
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK C4A.F05+C4A.F03
            IF FOUND()
               REPL F05 WITH F05+C4A.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH C4A.F05
               REPL F02 WITH C4A.F03
               REPL F05 WITH C4A.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))   
            SELE C4A
            SKIP 
         ENDDO
         SELE C4A
         USE                                                        
         
         THISFORM.LBL21.CAPTION='D39 樣品申請單(廠商)'         
         D39='D39'+LEFT(DTOS(CTOD(THISFORM.TXT1.VALUE+'/01')),6)
         IF !USED('&D39')
            SELE 0
            USE (D39) ALIA D39
         ELSE
            SELE D39
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F07='2'
               SELE B26                                            &&庫存異動紀錄
               APPEND BLANK                                               &&轉入  
               REPL F01 WITH D39.F03                               
               REPL F02 WITH D39.F10
               REPL F03 WITH D39.F02
               REPL F04 WITH D39.F04
               REPL F06 WITH 'F'
               REPL F07 WITH D39.F01           
               REPL F08 WITH '進貨廠商 '+D39.F05+SPACE(2)+IIF(SEEK(D39.F05,'D01'),D01.F04,'')
               S=S+1
               THISFORM.LBL41.CAPTION='B26 庫存異動紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+THISFORM.TXT1.VALUE+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE D39
            SKIP                               
         ENDDO   
         SELE F10,F03,SUM(F04) FROM D39 WHERE F07='2' GROUP BY F10,F03 ORDER BY F10,F03 INTO CURSOR D3H
         SELE D39
         USE 
         SELE D3H
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE B25
            SEEK D3H.F10+D3H.F03
            IF FOUND()
               REPL F04 WITH F04+D3H.SUM_F04
            ELSE
               APPEND BLANK
               REPL F01 WITH D3H.F10
               REPL F02 WITH D3H.F03
               REPL F04 WITH D3H.SUM_F04
            ENDIF
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))            
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F04))   
            SELE D3H
            SKIP 
         ENDDO
         SELE D3H
         USE                                                                                                          
         SELE B26
         USE                  
*****************************************************************  
         SELE A23
         SKIP -1       
         IF !BOF()
            BK=VAL(LEFT(DTOS(A23.F01),6))
            BK25='B25'+ALLTRIM(STR(BK))
            IF !USED('&BK25')
               SELE 0
               USE (BK25) ALIA BK25
            ELSE
               SELE BK25   
            ENDIF
            SET ORDER TO 1   
            THISFORM.LBL21.CAPTION='B25 月報表期初庫存數量回存'         
            S=0
            SELE BK25         
            GO TOP
            DO WHILE !EOF()
               SELE B25
               SEEK BK25.F01+BK25.F02
               IF FOUND()
                  REPL F03 WITH BK25.F15
               ELSE
                  APPEND BLANK
                  REPL F01 WITH BK25.F01
                  REPL F02 WITH BK25.F02
                  REPL F03 WITH BK25.F15
               ENDIF
               S=S+1
               THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F03))   
               THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))                  
               SELE BK25
               SKIP                          
            ENDDO
            SELE BK25
            USE
         ENDIF
         THISFORM.LBL21.CAPTION='B25 月報表期末庫存數量重算'         
         SELE B25
         S=0
         GO TOP         
         DO WHILE !EOF()
            REPL F15 WITH F03+F04-F05-F06+F07+F08-F09+F10-F11+F13-F14-F17+F18-F19-F20+F21
            S=S+2
            THISFORM.LBL31.CAPTION=F01+' '+F02+' '+ALLTRIM(STR(F15))
            THISFORM.LBL41.CAPTION='B25 庫存月報表'+'  '+ALLTRIM(STR(S))                    
            SKIP
         ENDDO
         SELE B25
         USE
         THISFORM.CMND2.CLICK
      ENDPROC   
  ****    
      PROCEDURE CMND2.CLICK        
            IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'B27'
           IF FOUND()
              DELETE 
           ENDIF        
           CLOSE TABLE ALL
           THISFORM.RELEASE 
           RETURN  
      ENDPROC      
ENDDEFINE            
*****