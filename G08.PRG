****程式名稱 G08成本計算
CLOSE ALL
CLEAR
SET DECIMALS TO 9
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'G08'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
  SELE A14
ENDIF
SET ORDER TO 1
IF !USED('A23')
   SELE 0
   USE A23
ELSE
   SELE A23
ENDIF
SET ORDER TO A231
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
TKG08FORM=CREATEOBJECT("TKG08")
TKG08FORM.SHOW       
DEFINE CLASS TKG08 AS FORM
  CAPTION='G08.月加權平均成本計算'
*!*	  AUTOCENTER=.T.  
  FONTSIZE=INT_015*9
  TOP=INT_015*90
  LEFT=INT_015*150    
  HEIGHT=INT_015*400
  WIDTH=INT_015*450
  CONTROLBOX=.F.
  BORDERSTYLE=2
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.  
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKG08'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*250,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入成本計算月份'
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
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*160,;
      TOP=INT_015*245,;
      WIDTH=INT_015*80,;      
      HEIGHT=INT_015*30,;    
      FONTSIZE=INT_015*12
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*50,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*120,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.執行成本計算',;
      TOOLTIPTEXT='開始執行重算單據',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*135,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<V.執行成本反結轉',;
      TOOLTIPTEXT='開始執行反結轉',;
      NAME='CMND2'      
  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*330,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<C.離開',;
      TOOLTIPTEXT='離開,並離開此畫面!快速鍵->ALT+C'
      NAME='CMND3'
 ********     
      PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.MTH_LIST.SETFOCUS
      ENDPROC
 **********    
      PROCEDURE CMND1.CLICK
      IF MESSAGEBOX('是否確定要執行成本計算作業',4+32+256,'請確認') = 6 
          SELECT A23
         SEEK DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01'))
         IF !FOUND()
            =MESSAGEBOX('無此月份資料!',0+48,'提示訊息視窗')
            THISFORM.MTH_LIST.SETFOCUS
            RETURN
         ELSE
               REPLACE F04 WITH 0
               REPLACE F05 WITH 0
               REPLACE F06 WITH 0
               REPLACE F10 WITH 0  
         ENDIF
          OMT1=VAL(RIGHT(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE),2))+1   &&計算本月最後一天為幾號
          IF OMT1>12
              OMT1=OMT1%12
              OMT2=VAL(LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),4))+1
           ELSE
              OMT2=VAL(LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),4))  
          ENDIF
          OMT3=RIGHT(DTOS(CTOD(ALLTRIM(STR(OMT2))+'/'+PADL(ALLTRIM(STR(OMT1)),2,'0')+'/01')-1),2)
  
        G06='G06'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G06')
            SELE 0
            USE (G06) ALIA G06
         ELSE
            SELE G06
         ENDIF
         SET ORDER TO 1                     
         DELETE FOR LEN(ALLTRIM(F02))=5
         REPLACE F05 WITH 0 ALL   &&每點平均分攤直接人工
         REPLACE F07 WITH 0 ALL   &&總有效工點
         REPLACE F08 WITH 0 ALL   &&每工點平均分攤製費
         REPLACE F09 WITH 0 ALL   &&已分攤成本工點 
         REPLACE F10 WITH 0 ALL   &&委外工資總額
         REPLACE F11 WITH 0 ALL   &&每點平均分攤委外工資
         REPLACE F12 WITH 0 ALL   &&間接材料費用
        G08='G08'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G08')
            SELE 0
            USE (G08) ALIA G08
         ELSE
            SELE G08
         ENDIF
         SET ORDER TO 1            
         DELE ALL                                   
        G10='G10'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G10')
            SELE 0
            USE (G10) ALIA G10
         ELSE
            SELE G10
         ENDIF
         SET ORDER TO 1            
         DELE ALL                          
        G13='G13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G13')
            SELE 0
            USE (G13) ALIA G13
         ELSE
            SELE G13
         ENDIF
         SET ORDER TO 1            
         DELE ALL                                   
         THISFORM.LBL21.CAPTION='B25 月報表期末庫存數量回存'         
         S=0         
        G11='G11'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G11')
            SELE 0
            USE (G11) ALIA G11
         ELSE
            SELE G11
         ENDIF
         SET ORDER TO 1            
         DELE ALL        
         B25='B25'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B25')
            SELE 0
            USE (B25) ALIA B25
         ELSE
            SELE B25
         ENDIF
         SET ORDER TO 1  
         SELE F02,SUM(F15) FROM B25 GROUP BY F02 ORDER BY F02 WHERE F01=ANY(SELE F01 FROM A14 WHERE F13='*') INTO CURSOR B25BK            
         SELE B25BK
         GO TOP
         S=0
         DO WHILE !EOF()
            SELE G11
            APPEND BLANK
            REPL F02 WITH B25BK.F02
            
            REPL F11 WITH B25BK.SUM_F15
            S=S+1
            THISFORM.LBL41.CAPTION='B25 庫存月報表匯入期初期末數量'+'  '+ALLTRIM(STR(S))
            THISFORM.LBL31.CAPTION=F02+' '+ALLTRIM(STR(F11))              
            SELE B25BK
            SKIP
         ENDDO 
         SELE B25
         USE
         SELE B25BK 
         USE
         G1=VAL(RIGHT(ALLTRIM(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)),2))-1
         IF G1=0
            G1=12
            G2=YEAR(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))-1
         ELSE
            G2=YEAR(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))
         ENDIF
         HX=ALLTRIM(STR(G2))+PADL(ALLTRIM(STR(G1)),2,'0')
         G0='G11'+HX
         IF !USED('&G0')
            SELE 0
            USE (G0) ALIA G0
         ELSE
            SELE G0
         ENDIF
         SET ORDER TO 1      
         SELE G11
         GO TOP
         DO WHILE !EOF()
                SELE G0
                SEEK G11.F02
                IF FOUND()
                    SELE G11
                    REPLACE F03 WITH G0.F11
                    REPLACE  F13 WITH G0.F21
                    REPLACE  F23 WITH F03*F13
                    THISFORM.LBL41.CAPTION=G0+' 匯入前期期末單價與期末數量'+'  '+ALLTRIM(STR(S))
                    THISFORM.LBL31.CAPTION=F02+' '+ALLTRIM(STR(F21))   
                ENDIF                             
                SELE G11
                SKIP
         ENDDO
         SELECT G0
         USE
         G3='G13'+HX
         IF !USED('&G3')
            SELE 0
            USE (G3) ALIA G3
         ELSE
            SELE G3
         ENDIF
         SET ORDER TO 1      
         GO TOP
          S=0    
         DO WHILE !EOF()
                 IF F02=0
                     SELECT G13
                     SEEK G3.F01
                     IF !FOUND()
                         APPEND BLANK
                         REPLACE F01 WITH G3.F01
                     ENDIF
                     REPLACE F02 WITH F02+G3.F02
                     REPLACE F03 WITH F03+G3.F03
                     REPLACE F04 WITH F04+G3.F04
                     REPLACE F05 WITH F05+G3.F05
                     REPLACE F06 WITH F06+G3.F06    
                     THISFORM.LBL41.CAPTION=G3+' 匯入前期在製品'+'  '+ALLTRIM(STR(S))
                     THISFORM.LBL31.CAPTION=F01+' '+ALLTRIM(STR(F02))
                 ENDIF                             
            SELE G3
            SKIP
         ENDDO                  
         SELECT G3
         USE
         G15='G15'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&G15')
            SELE 0
            USE (G15) ALIA G15
         ELSE
            SELE G15
         ENDIF
         SET ORDER TO 1                                 
         DELE ALL
         SELE G11
         GO TOP
         S=0
         DO WHILE !EOF()
                IF F03<>0
                    SELE G15
                    APPEND BLANK
                    REPLACE  F01 WITH G11.F02               
                    REPLACE  F04 WITH G11.F03
                    REPLACE  F05 WITH G11.F13
                    REPLACE  F07 WITH SPACE(2)+'上月結轉'
                    THISFORM.LBL41.CAPTION='G15 上月結轉數量及單價'+'  '+ALLTRIM(STR(S))
                    THISFORM.LBL31.CAPTION=F01+' '+ALLTRIM(STR(F05))                      
                ENDIF   
                SELE G11                
                SKIP            
         ENDDO
         THISFORM.LBL21.CAPTION='G06 其他進貨費用'
         G07='G07'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&G07')
            SELE 0
            USE (G07) ALIA G07
         ELSE
            SELE G07
         ENDIF
         SET ORDER TO 1          
         GO TOP
         S=0
         DO WHILE !EOF()
                SELE G11
                SEEK G07.F05
                IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH G07.F05
                ENDIF
                REPLACE  F24 WITH F24+G07.F09                       
                S=S+1               
                THISFORM.LBL41.CAPTION='G06 其他進貨費用'+'  '+ALLTRIM(STR(S))                              
                THISFORM.LBL31.CAPTION=F02+' '+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+' '+ALLTRIM(STR(F24))  
               SELECT G15
               APPEND BLANK
               REPLACE F01 WITH G07.F05
               REPLACE F02 WITH OMT3
               REPLACE F03 WITH '1'
               REPLACE F05 WITH G07.F09
               REPLACE F07 WITH '進'+G07.F01                    
                SELECT A23
                REPLACE F06 WITH F06+G07.F09   
                SELE G07
                SKIP         
         ENDDO          

         THISFORM.LBL21.CAPTION='B02 進貨單'
         B02='B02'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
                IF B02.F10='2' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,'')) 
                     SELE G11
                     SEEK B02.F03
                     IF !FOUND()
                         APPEND BLANK                         
                         REPLACE F02 WITH B02.F03
                     ENDIF    
                     REPLACE F04 WITH F04+(B02.F04+B02.F15)   &&進貨數量
                     REPLACE  F24 WITH F24+ROUND(IIF(B02.F18=1,B02.F04*B02.F17,B02.F04*B02.F17*B02.F18),INT_068)                
                     SELE G15                                            &&存貨明細紀錄
                     APPEND BLANK
                     REPLACE  F01 WITH B02.F03
                     REPLACE  F02 WITH B02.F02
                     REPLACE  F04 WITH (B02.F04+B02.F15)
                     REPLACE  F05 WITH ROUND((IIF(B02.F18=1,B02.F04*B02.F17,B02.F04*B02.F17*B02.F18)+IIF(SEEK(B02.F01+B02.F03+B02.F09,'G07'),G07.F09,0))/(B02.F04+B02.F15),6)  &&單價
                     REPLACE  F03 WITH 'A'
                     REPLACE  F07 WITH '進'+B02.F01+'採'+B02.F09+IIF(!EMPTY(B02.F02),'請'+B02.F08,'')
                     S=S+1               
                    THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))                              
                    THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))  
                     SELECT A23
                     REPLACE F06 WITH F06+ROUND(IIF(B02.F18=1,B02.F04*B02.F17,B02.F04*B02.F17*B02.F18),INT_068)                     
                ENDIF  
                SELE B02
                SKIP         
         ENDDO 
         SELECT B02
         USE 
         SELECT G07
         USE
         THISFORM.LBL21.CAPTION='P16 JW其他進貨費用'
         P17='P17'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&P17')
            SELE 0
            USE (P17) ALIA P17
         ELSE
            SELE P17
         ENDIF
         SET ORDER TO 1          
         GO TOP
         S=0
         DO WHILE !EOF()
            IF P17.F09<>0    &&測試免除金額為零者
               SELE G11
               SEEK P17.F05
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH P17.F05   
               ENDIF
               REPL F24 WITH F24+P17.F09                  
               S=S+1               
               THISFORM.LBL41.CAPTION='P16 JW其他進貨費用'+'  '+ALLTRIM(STR(S))                              
               THISFORM.LBL31.CAPTION=F02+' '+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+' '+ALLTRIM(STR(F24))  
               SELECT G15
               APPEND BLANK
               REPLACE F01 WITH P17.F05
               REPLACE F02 WITH OMT3
               REPLACE F03 WITH '1'
               REPLACE F05 WITH P17.F09
               REPLACE F07 WITH '送'+P17.F01+SPACE(2)+'進'+P17.F10
               SELECT A23
               REPLACE F06 WITH F06+P17.F09     
             ENDIF            
             SELE P17
             SKIP         
         ENDDO        
         SELECT P17           
         SET ORDER TO 2
         THISFORM.LBL21.CAPTION='P12 外購進貨單'
         P12='P12'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
                IF P12.F10='2' AND !EMPTY(IIF(SEEK(P12.F03,'B01'),B01.F98,''))
                    SELE G11
                    SEEK P12.F03
                    IF !FOUND()
                         APPEND BLANK
                         REPLACE F02 WITH P12.F03
                    ENDIF                         
                    REPLACE F04 WITH F04+(P12.F04+P12.F15)
                    REPLACE  F24 WITH F24+ROUND(IIF(P12.F18=1,P12.F04*P12.F17,P12.F04*P12.F17*P12.F18),INT_068)
                    SELE G15                                            &&存貨明細紀錄
                    APPEND BLANK
                    REPLACE  F01 WITH P12.F03
                    REPLACE  F02 WITH P12.F02
                    REPLACE  F04 WITH (P12.F04+P12.F15)
                    REPLACE  F05 WITH ROUND((IIF(P12.F18=1,P12.F04*P12.F17,P12.F04*P12.F17*P12.F18)+IIF(SEEK(P12.F01+P12.F03+P12.F09,'P17'),P17.F09,0))/(P12.F04+P12.F15),6)             
                    REPLACE  F03 WITH 'A'
                    REPLACE  F07 WITH '進'+P12.F01+'採'+P12.F09+IIF(!EMPTY(P12.F08),'請'+P12.F08,'')
                    S=S+1               
                    THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))
                    THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))               
                    SELECT A23
                    REPLACE F06 WITH F06+ROUND(IIF(P12.F18=1,P12.F04*P12.F17,P12.F04*P12.F17*P12.F18),INT_068)                    
                ENDIF  
                SELE P12
                SKIP         
         ENDDO 
         SELECT P12
         USE         
         SELECT P17
         USE 
******************************************
         THISFORM.LBL21.CAPTION='D39 樣品申請單(廠商)'
         D39='D39'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
            IF D39.F07='2' AND IIF(SEEK(D39.F10,'A14'),A14.F13,'')='*'  AND !EMPTY(IIF(SEEK(D39.F03,'B01'),B01.F98,''))
               SELE G11
               SEEK D39.F03
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH D39.F03
               ENDIF              
               REPLACE F04 WITH F04+D39.F04     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK
               REPL F01 WITH D39.F03
               REPL F02 WITH D39.F02
               REPL F04 WITH D39.F04
               REPL F05 WITH 0
               REPL F03 WITH 'F'
               REPL F07 WITH '樣'+D39.F01+'進貨廠商 '+D39.F05+SPACE(2)+IIF(SEEK(D39.F05,'D01'),D01.F04,'')
               S=S+1               
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
             ENDIF  
             SELE D39
             SKIP         
         ENDDO 
         SELECT D39
         USE
         B06='B06'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B06')
            SELE 0
            USE (B06) ALIA B06
         ELSE
            SELE B06
         ENDIF
         SET ORDER TO 1                        
         THISFORM.LBL21.CAPTION='B06 移轉單視為樣入者'         
         SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ')  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13='*' ) INTO CURSOR B06T1 ORDER BY F01,F03 NOWAIT &&由非成本部門轉入成本部門視為樣入        
         IF _TALLY>0
            GO TOP
            S=0
            DO WHILE !EOF()
               SELE G11
               SEEK B06T1.F03
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH B06T1.F03
               ENDIF              
               REPLACE F04 WITH F04+B06T1.F04     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK
               REPL F01 WITH B06T1.F03
               REPL F02 WITH B06T1.F02
               REPL F04 WITH B06T1.F04
               REPL F05 WITH 0
               REPL F03 WITH 'E'
               REPL F07 WITH '移'+B06T1.F01+' 自'+B06T1.F05+SPACE(2)+IIF(SEEK(B06T1.F05,'A14'),A14.F02,'')+'轉入'
               S=S+1               
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))     
               SELE B06T1
               SKIP         
            ENDDO 
            USE         
         ENDIF         
         B08='B08'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B08')
            SELE 0
            USE (B08) ALIA B08
         ELSE
            SELE B08
         ENDIF
         SET ORDER TO 1                        
         THISFORM.LBL21.CAPTION='B08 製令領料單視為樣入者'               
         SELECT * FROM B08 WHERE F12>0 AND F13='2' AND F07 =ANY (SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ')  AND F04 =ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B08T ORDER BY F01,F08 NOWAIT &&從非成本部門領料視為樣入         
         IF _TALLY>0
            GO TOP
            S=0
            DO WHILE !EOF()
               SELE G11
               SEEK B08T.F08
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH B08T.F08
               ENDIF              
               REPLACE F04 WITH F04+B08T.F12     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK
               REPL F01 WITH B08T.F08
               REPL F02 WITH B08T.F02
               REPL F04 WITH B08T.F12
               REPL F05 WITH 0
               REPL F03 WITH 'P'
               REPL F07 WITH '領'+B08T.F01+' 自'+B08T.F07+SPACE(2)+IIF(SEEK(B08T.F07,'A14'),A14.F02,'')+'領料'
               S=S+1               
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))     
               SELE B08T
               SKIP         
            ENDDO 
            USE                                     
         ENDIF
         SELECT B08
         USE
******************************************         
         THISFORM.LBL21.CAPTION='B03 進貨退出單'
         B03='B03'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
                IF F10='2' AND !EMPTY(IIF(SEEK(F03,'B01'),B01.F98,''))  &&非虛擬料號才做運算
                    SELE G11
                    SEEK B03.F03
                    IF !FOUND()
                         APPEND BLANK
                         REPLACE F02 WITH B03.F03
                    ENDIF                         
                    REPLACE F04 WITH F04+IIF(B03.F24='2',0,(B03.F04+B03.F17)*(-1))      &&若為折讓則退貨數量計為零
                    REPLACE F24 WITH F24+ROUND(IIF(B03.F18=1,B03.F04*B03.F16,B03.F04*B03.F16*B03.F18),INT_068)*(-1)
                    SELE G15                                            &&存貨明細紀錄
                    APPEND BLANK
                    REPLACE  F01 WITH B03.F03
                    REPLACE  F02 WITH B03.F02
                    REPLACE  F04 WITH IIF(B03.F24='2',0,(B03.F04+B03.F17)*(-1))      &&若為折讓則退貨數量計為零
                    REPLACE  F05 WITH IIF(B03.F24='2',IIF(B03.F18=1,B03.F04*B03.F16,B03.F04*B03.F16*B03.F18)*(-1),ROUND(IIF(B03.F18=1,B03.F04*B03.F16,B03.F04*B03.F16*B03.F18)/(B03.F04+B03.F17),6))
                    REPLACE  F03 WITH 'B'
                    REPLACE  F07 WITH IIF(B03.F24='2','折','退')+B03.F01+'採'+B03.F09
                    S=S+1
                    THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                    THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
                    SELECT A23
                    REPLACE F06 WITH F06+ROUND(IIF(B03.F18=1,B03.F04*B03.F16,B03.F04*B03.F16*B03.F18),INT_068)*(-1)                    
                ENDIF   
                SELE B03
                SKIP                               
         ENDDO   
         SELECT B03
         USE
         THISFORM.LBL21.CAPTION='B04/B05 出貨/出貨退回單'         
         C13='C13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)         
         IF !USED('&C13')
            SELE 0
            USE (C13) ALIA C13
         ELSE
            SELE C13
         ENDIF
         SET ORDER TO 4                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF !EMPTY(IIF(SEEK(C13.F79,'B01'),B01.F98,'')) 
               SELE G11
               SEEK C13.F79
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH C13.F79
               ENDIF       
               REPLACE F05 WITH F05+IIF(C13.F18='03',0,C13.F06)  &&若為折讓則退貨數量計為零       
               REPLACE F25 WITH F25+ROUND(C13.F08/IIF(C13.F16='02',1+INT_002/100,1),INT_069)&&稅含者要刨掉
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK
               REPL F01 WITH C13.F79
               REPL F02 WITH RIGHT(DTOC(C13.F01),2)
               REPL F04 WITH IIF(C13.F18='03',0,C13.F06*(-1))  &&若為折讓則退貨數量計為零 
               REPL F05 WITH IIF(C13.F18='03',C13.F08,ROUND(C13.F08/C13.F06/IIF(C13.F16='02',1+INT_002/100,1),6))
               REPL F03 WITH SUBSTR(C13.F02,2,1)
               REPL F07 WITH IIF(F03='C','出','退')+C13.F02+'訂'+C13.F04+' '+IIF(SEEK(C13.F03,'C01'),C01.F05,'')
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF
            SELE C13
            SKIP                               
         ENDDO   
         SELECT C13
         USE

         **********計算當出貨退回原出貨當月時之成本***
         THISFORM.LBL21.CAPTION='B05 出貨退回單成本計算'       
         B05='B05'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B05')
            SELE 0
            USE (B05) ALIA B05
         ELSE
            SELE B05
         ENDIF
         SET ORDER TO 1                                     
         S=0    
         SELECT F79,F04,F08,F15 FROM B05 WHERE CTOD(ALLTRIM(F08)+'/01')<CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01') AND F12='2' AND F24<>'2' ORDER BY F08 INTO CURSOR ZEN NOWAIT         
         IF _TALLY>0
            SELECT ZEN 
            GO TOP
            HP=SPACE(7)
            KP='ZZZZZZZZZ'
            DO WHILE !EOF()
               IF ZEN.F08<>HP
                  IF FILE('&KB'+'.DBF')
                     SELECT &KP
                     USE                 
                  ENDIF   
               ENDIF   
               KP='G11'+LEFT(DTOS(CTOD(ALLTRIM(ZEN.F08)+'/01')),6)
               
                   IF !USED('&KP')
                     SELECT 0
                     USE &KP 
                  ELSE
                     SELECT &KP
                  ENDIF
                  SET ORDER TO 1
                  SELECT G11
                  SEEK ZEN.F79
                  IF FOUND()             
                     REPLACE F33 WITH F33+ZEN.F04+ZEN.F15 
                     REPLACE F35 WITH F35+ZEN.F04*IIF(SEEK(ZEN.F79,'&KP'),&KP..F21,0)            
                     S=S+1
                     THISFORM.LBL41.CAPTION='寫回 G11'+'  '+ALLTRIM(STR(S))
                     THISFORM.LBL31.CAPTION=ZEN.F79+'成本單價'+TRANSFORM(F35/F33,'@R 9999999999.999999')                     
                  ENDIF    
                       
                  HP=ZEN.F08
               SELECT ZEN
               SKIP       
            ENDDO
         ENDIF   
****************************************         
         SELECT B05         
         USE
        SELE G11
        GO TOP
        DO WHILE !EOF()
            IF F03+F04+F33<>0
               IF F23=0 AND F24=0 
                  REPLACE F21 WITH ROUND((F23+F24+F35)/(F03+F04+F33),6) &&計算期末單價   
               ELSE
                  REPLACE F21 WITH ROUND((F23+F24)/(F03+F04),6)
               ENDIF              
            ELSE
               REPLACE F21 WITH 0 
            ENDIF                  
            &&REPLACE  F31 WITH ROUND(F11*F21,6)                 &&計算期末小計
            SKIP 
         ENDDO
         THISFORM.LBL21.CAPTION='B06 移轉單視為他耗者'
         SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*')  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ') INTO CURSOR B06T2 ORDER BY F01,F79 NOWAIT  &&由成本部門轉入非成本部門視為樣出   
         IF _TALLY>0
            GO TOP
            S=0
            DO WHILE !EOF()              
               SELECT G11
               SEEK B06T2.F79
               IF !FOUND()
                    APPEND BLANK
                    REPLACE F02 WITH B06T2.F79
               ENDIF
               REPLACE F10 WITH F10+B06T2.F04     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK                                               &&損耗  
               REPL F01 WITH B06T2.F79                              
               REPL F02 WITH B06T2.F02
               REPL F04 WITH B06T2.F04*(-1)
               REPL F03 WITH 'E'
               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)               
               REPL F07 WITH '移'+B06T2.F01+' 轉出到  '+B06T2.F10+SPACE(2)+IIF(SEEK(B06T2.F10,'A14'),A14.F02,'')
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))                
               SELE B06T2
               SKIP                               
           ENDDO                       
           USE                     
         ENDIF
         SELECT B06
         USE
         THISFORM.LBL21.CAPTION='B09 盤差調整單'         
         B09='B09'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
            IF F11='2'  AND IIF(SEEK(B09.F05,'A14'),A14.F13,'')='*'  AND !EMPTY(IIF(SEEK(B09.F03,'B01'),B01.F98,''))
               SELECT G11
               SEEK B09.F03
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH B09.F03
               ENDIF
               REPLACE F09 WITH F09+B09.F04    
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK                                              
               REPL F01 WITH B09.F03                               
               REPL F02 WITH B09.F02
               REPL F04 WITH B09.F04
               REPL F03 WITH 'H'
               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)
               REPL F07 WITH '盤'+B09.F01+'  '+B09.F07
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B09
            SKIP                               
         ENDDO   
         USE
         THISFORM.LBL21.CAPTION='B10 報廢單'         
         B10='B10'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
            IF F09='2' AND IIF(SEEK(B10.F05,'A14'),A14.F13,'')='*'  AND !EMPTY(IIF(SEEK(B10.F03,'B01'),B01.F98,''))
                SELECT G11
                SEEK B10.F03
                IF !FOUND()
                     APPEND BLANK
                     REPLACE F02 WITH B10.F03
                ENDIF                
                REPLACE F10 WITH F10+B10.F04     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK                                               &&報廢
               REPL F01 WITH B10.F03                               
               REPL F02 WITH B10.F02
               REPL F04 WITH B10.F04*(-1)
               REPL F03 WITH 'I'
               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)               
               REPL F07 WITH '廢'+B10.F01+'  '+B10.F07
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE B10
            SKIP                               
         ENDDO   
         USE
         THISFORM.LBL21.CAPTION='B33 製令耗料單'         
         B16='B16'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B16')
            SELE 0
            USE (B16) ALIA B16
         ELSE
            SELE B16
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         DO WHILE !EOF()
            IF F13='2' AND IIF(SEEK(B16.F07,'A14'),A14.F13,'')='*' 
               SELECT G11
               SEEK B16.F08
               IF !FOUND()
                   APPEND BLANK
                   REPLACE F02 WITH B16.F08
                ENDIF
                REPLACE F07 WITH F07+B16.F12   
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK                                               &&轉出  
               REPL F01 WITH B16.F08                               
               REPL F02 WITH B16.F02
               REPL F04 WITH B16.F12*(-1)
               REPL F03 WITH 'K'
               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)               
               REPL F07 WITH '領'+B16.F01+'製'+B16.F05+'耗用'
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
               SELECT G10
               SEEK B16.F09+B16.F08
               IF FOUND()
                    REPLACE F04 WITH F04+B16.F12
               ELSE                        
                    APPEND BLANK
                    REPLACE F01 WITH B16.F09
                    REPLACE F02 WITH B16.F08
                    REPLACE F04 WITH B16.F12
               ENDIF                        
            ENDIF   
            SELE B16
            SKIP                               
         ENDDO   
         USE
*!*	         THISFORM.LBL21.CAPTION='B34 退回入倉單'         
*!*	         B18='B18'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
*!*	         IF !USED('&B18')
*!*	            SELE 0
*!*	            USE (B18) ALIA B18
*!*	         ELSE
*!*	            SELE B18
*!*	         ENDIF
*!*	         SET ORDER TO 1                              
*!*	         GO TOP
*!*	         S=0
*!*	         DO WHILE !EOF()
*!*	            IF F10='2'
*!*	               SELE G15                                            &&存貨明細紀錄
*!*	               APPEND BLANK                                               &&轉出  
*!*	               REPL F01 WITH B18.F07                             
*!*	               REPL F02 WITH B18.F02
*!*	               REPL F04 WITH B18.F09
*!*	               REPL F03 WITH 'L'
*!*	               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)               
*!*	               REPL F07 WITH '退'+B18.F01
*!*	               S=S+1
*!*	               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
*!*	               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
*!*	            ENDIF   
*!*	            SELE B18
*!*	            SKIP                               
*!*	         ENDDO            
*!*	         USE
         THISFORM.LBL21.CAPTION='B40 額外生產耗料單'         
         B40='B40'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
            IF F11='2' AND IIF(SEEK(B40.F05,'A14'),A14.F13,'')='*'
                SELECT G11
                SEEK B40.F03
                IF !FOUND()
                     APPEND BLANK
                     REPLACE F02 WITH B40.F03
                ENDIF
                REPLACE F10 WITH F10+B40.F04     
               SELE G15
               APPEND BLANK                                               &&轉入   
               REPL F01 WITH B40.F03
               REPL F02 WITH B40.F02
               REPL F04 WITH B40.F04*(-1)
               REPLACE F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)   
               REPL F03 WITH 'Q'
               REPL F07 WITH '耗'+B40.F01   
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
               
               IF IIF(SEEK(B40.F05,'A14'),A14.F05,'')='*'
                  SELECT G06
                  SEEK B40.F05
                  IF !FOUND()
                     APPEND BLANK
                     REPLACE F02 WITH B40.F05
                  ENDIF    
                  REPLACE F12 WITH F12+G15.F05*B40.F04
               ENDIF   
            ENDIF   
            SELE B40
            SKIP                               
         ENDDO            
         USE
         THISFORM.LBL21.CAPTION='C38 樣品領用單'         
         C38='C38'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
            IF F07='2' AND IIF(SEEK(F05,'A14'),A14.F13,'')='*'
               SELECT G11
               SEEK C38.F79
               IF !FOUND()
                    APPEND BLANK
                    REPLACE F02 WITH C38.F79
               ENDIF
               REPLACE F10 WITH F10+C38.F04     
               SELE G15                                            &&存貨明細紀錄
               APPEND BLANK                                               &&損耗  
               REPL F01 WITH C38.F79                              
               REPL F02 WITH C38.F02
               REPL F04 WITH C38.F04*(-1)
               REPL F03 WITH 'Z'
               REPL F05 WITH IIF(SEEK(F01,'G11'),G11.F21,0)               
               REPL F07 WITH '樣'+C38.F01+C38.F10+SPACE(2)+IIF(SEEK(C38.F10,'C01'),C01.F05,'')
               S=S+1
               THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
            ENDIF   
            SELE C38
            SKIP                               
         ENDDO                       
         USE         
         IF !USED('E00')
              SELECT 0
              USE E00
          ELSE
              SELECT E00
          ENDIF
          SET ORDER TO 1        
         THISFORM.LBL21.CAPTION='E13 製程完工紀錄'         
         E13='E13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&E13')
            SELE 0
            USE (E13) ALIA E13
         ELSE
            SELE E13
         ENDIF
         SET ORDER TO 1         
         THISFORM.LBL21.CAPTION='E14 製令耗料紀錄'         
         E14='E14'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&E14')
            SELE 0
            USE (E14) ALIA E14
         ELSE
            SELE E14
         ENDIF
         SET ORDER TO 1                                                                                                     
         THISFORM.LBL21.CAPTION='B55 完工入庫單'         
         B55='B55'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
                IF F13='2'    &&如果已過帳才做計算
                    IF F09<>'2'  &&如果不是重工
                        SELECT G11
                        SEEK B55.F08
                        IF !FOUND()
                            APPEND BLANK
                            REPLACE F02 WITH B55.F08
                       ENDIF
                       REPLACE F06 WITH F06+B55.F12     
                        SELE G15
                        APPEND BLANK                                               &&轉入   
                        REPLACE  F01 WITH B55.F08
                        REPLACE  F02 WITH B55.F02
                        REPLACE  F04 WITH B55.F12
                        REPLACE  F03 WITH 'T'
                        REPLACE  F07 WITH '完'+B55.F01+'製'+B55.F05   
                        S=S+1
                        THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                        THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
                    ENDIF    
                        SELECT G08
                        SEEK B55.F01+B55.F05+B55.F07
                        IF FOUND()
                            REPLACE F03 WITH F03+IIF(B55.F09='2',0,B55.F12)
                         ELSE
                            APPEND BLANK
                            REPLACE F01 WITH B55.F08
                            REPLACE F02 WITH B55.F05
                            REPLACE F03 WITH IIF(B55.F09='2',0,B55.F12)
                            REPLACE F05 WITH B55.F07
                        ENDIF    
                        IF INT_066='1'      &&若以標準工時計算工點
                            SELECT E13
                            SEEK B55.F01+B55.F05
                            IF FOUND()
                                 DO WHILE F04+F01=B55.F01+B55.F05
                                        SELECT G08
                                        REPLACE F04 WITH F04+E13.F05*IIF(SEEK(E13.F02,'E00'),E00.F02,1)+IIF(SEEK(E13.F02,'E00'),E00.F06,1)
                                        SELECT G06
                                        SEEK B55.F07
                                        IF FOUND()
                                            REPLACE F07 WITH F07+E13.F05*IIF(SEEK(E13.F02,'E00'),E00.F02,1)+IIF(SEEK(E13.F02,'E00'),E00.F06,1)
                                        ENDIF
                                        SELECT E13
                                        SKIP
                                 ENDDO      
                            ENDIF                     
                        ELSE                 &&若以實際工時計算工點
                             SELECT G08
                             REPLACE F04 WITH F04+B55.F10
                            SELECT G06
                            SEEK B55.F07
                            IF FOUND()
                                REPLACE F07 WITH F07+B55.F10
                            ENDIF
                        ENDIF    
               ENDIF
               SELECT E14
               SEEK B55.F01+B55.F05      
               IF FOUND()          &&若有此耗料紀錄     
                   DO WHILE F04+F01=B55.F01+B55.F05        
	                       IF F08<>.T.        &&且不為客供品   
                              SELECT G11
                              SEEK E14.F02
                              IF !FOUND()
                                   APPEND BLANK
                                   REPLACE F02 WITH E14.F02
                              ENDIF
                              REPLACE F07 WITH F07+E14.F05     
                              SELECT G15
                              APPEND BLANK
                              REPLACE F01 WITH E14.F02
                              REPLACE F02 WITH E14.F03
                              REPLACE F03 WITH 'T'
                              REPLACE F04 WITH E14.F05*(-1)
                              REPLACE F05 WITH IIF(SEEK(E14.F02,'G11'),G11.F21,0)
                              REPLACE F07 WITH '完'+B55.F01+'製'+E14.F01+'耗用'                      
                              THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                              THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))                                  
                               SELECT G10
                               SEEK B55.F08+E14.F02
                               IF !FOUND()
                                   APPEND BLANK 
                                   REPLACE F01 WITH B55.F08
                                   REPLACE F02 WITH E14.F02                                          
                              ENDIF
                              REPLACE F03 WITH F03+E14.F05
                         ELSE     &&若為客供品亦將之加入直接材料耗料檔以便佇列排定成本計算優先順序
                                 SELECT G10
                                 SEEK B55.F08+SPACE(43)
                                 IF !FOUND()
                                      APPEND BLANK
                                      REPLACE F01 WITH B55.F08
                                 ENDIF                              
                         ENDIF     
                        SELECT E14
                        SKIP
                  ENDDO
              ELSE                 &&若無耗料紀錄亦將之加入直接材料耗料檔以便佇列排定成本計算優先順序
                    SELECT G10
                    SEEK B55.F08+SPACE(43)
                    IF !FOUND()
                        APPEND BLANK
                        REPLACE F01 WITH B55.F08
                    ENDIF                         
              ENDIF                   
               SELE B55
               SKIP                               
         ENDDO   
         USE         
         THISFORM.LBL21.CAPTION='B07 委外入庫單'         
         B07='B07'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
         IF !USED('&B07')
            SELE 0
            USE (B07) ALIA B07
         ELSE
            SELE B07
         ENDIF
         SET ORDER TO 1                              
         GO TOP
         S=0
         JKS=''
         
         DO WHILE !EOF()
               IF F10='2' AND !EMPTY(IIF(SEEK(B07.F03,'B01'),B01.F98,''))
                    
                        SELECT G11
                        SEEK B07.F03
                        IF !FOUND()
                             APPEND BLANK
                             REPLACE F02 WITH B07.F03
                        ENDIF
                        IF SEEK(B07.F01+B07.F09,'E14')  &&如果有耗料
                           REPLACE F06 WITH F06+IIF(B07.F13='1',B07.F04,0)
                        ELSE
                           REPLACE F04 WITH F04+IIF(B07.F13='1',B07.F04,0)
                           REPLACE F24 WITH F24+ROUND(IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18),INT_070)
                           IF F23=0 AND F24=0 AND F26=0
                               REPLACE F21 WITH ROUND((F23+F24+F26+F35)/(F03+F04+F06+F33),6)
                           ELSE
                               IF F03+F04+F06<>0
                                  REPLACE F21 WITH ROUND((F23+F24+F26)/(F03+F04+F06),6)
                               ELSE
                                  IF F03>0
                                     REPLACE F21 WITH F23/F03
                                  ELSE
                                     IF F04>0
                                        REPLACE F21 WITH F24/F04
                                     ELSE
                                        IF F06>0
                                           REPLACE F21 WITH F26/F06   
                                        ENDIF   
                                     ENDIF      
                                  ENDIF                                          
                               ENDIF   
                           ENDIF    
                           &&REPLACE F31 WITH F11*F21
                        ENDIF       
                        
                           SELE G15                                            &&存貨明細紀錄
                           APPEND BLANK                                              
                           REPLACE  F01 WITH B07.F03                               
                           REPLACE  F02 WITH B07.F02
                           REPLACE  F04 WITH IIF(B07.F13='1',B07.F04,0)
                           REPLACE  F03 WITH 'M'
                           REPLACE  F05 WITH ROUND((IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18))/(IIF(B07.F13='1',B07.F04,1)),6)             
                           REPLACE  F07 WITH '委'+B07.F01+'外'+B07.F15+' 製'+B07.F09
                           S=S+1
                           THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                           THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
                       
                    SELECT E14
                    SEEK B07.F01+B07.F09                    
                    IF FOUND()       &&若耗料紀錄有此料號
                       SELECT G08
                       SEEK B07.F03+B07.F09+B07.F14
                       IF FOUND()
                          REPLACE F03 WITH F03+IIF(B07.F13='1',B07.F04,0)                                            
                       ELSE
                          APPEND BLANK
                          REPLACE F01 WITH B07.F03
                          REPLACE F02 WITH B07.F09
                          REPLACE F03 WITH IIF(B07.F13='1',B07.F04,0)
                          REPLACE F05 WITH B07.F14
                       ENDIF
                       SELECT E13
                       SEEK B07.F01+B07.F09
                       IF FOUND()
                          DO WHILE F04+F01=B07.F01+B07.F09
                                SELECT G08
                                REPLACE F04 WITH F04+E13.F05*IIF(SEEK(E13.F02,'E00'),E00.F02,1)+IIF(SEEK(E13.F02,'E00'),E00.F06,1)
                                SELECT G06
                                SEEK B07.F14
                                IF !FOUND()
                                    APPEND BLANK
                                    REPLACE F02 WITH B07.F14
                               ENDIF
                               REPLACE F07 WITH F07+E13.F05*IIF(SEEK(E13.F02,'E00'),E00.F02,1)+IIF(SEEK(E13.F02,'E00'),E00.F06,1)
                               SELECT E13
                               SKIP
                          ENDDO
                       ENDIF
                       IF JKS<>B07.F01+B07.F09 &&若同一張入庫單製令未重複
                          SELECT E14      
                       
                          DO WHILE F04+F01=B07.F01+B07.F09
                                IF F08<>.T.      &&且不為客供品
                                   
                                     SELECT G11
                                     SEEK E14.F02
                                     IF !FOUND()
                                          APPEND BLANK
                                          REPLACE F02 WITH E14.F02
                                     ENDIF
                                     REPLACE F07 WITH F07+E14.F05     
                                     SELECT G15
                                     APPEND BLANK
                                     REPLACE F01 WITH E14.F02
                                     REPLACE F02 WITH E14.F03
                                     REPLACE F03 WITH 'M'
                                     REPLACE F04 WITH E14.F05*(-1)
                                     REPLACE F05 WITH IIF(SEEK(E14.F02,'G11'),G11.F21,0)
                                     REPLACE F07 WITH '委'+B07.F01+'製'+E14.F01+'耗用'                    
                                     THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                                     THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))                                                                  
                                     SELECT G10
                                     SEEK B07.F03+E14.F02
                                     IF !FOUND()
                                         APPEND BLANK 
                                         REPLACE F01 WITH B07.F03
                                         REPLACE F02 WITH E14.F02                                          
                                     ENDIF
                                     REPLACE F03 WITH F03+E14.F05               
                                 ELSE         &&若為客供品將之加入直接材料耗料檔以便佇列排定成本計算優先順序
                                      SELECT G10
                                      SEEK B07.F03+SPACE(43)
                                      IF !FOUND()
                                         APPEND BLANK
                                         REPLACE F01 WITH B07.F03    
                                      ENDIF                                             
                                 ENDIF    
                                 SELECT E14
                                 SKIP
                           ENDDO
                         ENDIF
                         SELECT G06
                         SEEK B07.F14
                         IF !FOUND()
                             APPEND BLANK
                            REPLACE F02 WITH B07.F14
                         ENDIF            
                         REPLACE F10 WITH F10+ROUND(IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18),INT_070)
                         SELECT G13
                         SEEK B07.F03
                         IF !FOUND()
                            APPEND BLANK
                            REPLACE F01 WITH B07.F03
                         ENDIF
                         REPLACE F06 WITH F06+ROUND(IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18),INT_070)
                        
                    ENDIF
                    SELECT A23
                    IF SEEK(B07.F01+B07.F09,'E14')
                       REPLACE F10 WITH F10+ROUND(IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18),INT_070) &&如果有耗料算委外工資
                    ELSE 
                       REPLACE F06 WITH F06+ROUND(IIF(B07.F18=1,B07.F04*B07.F17,B07.F04*B07.F17*B07.F18),INT_070)  &&否則算進貨金額
                    ENDIF    
               ENDIF              
               SELE B07
               JKS=B07.F01+B07.F09 
               SKIP                
                             
         ENDDO            
         USE         
         SELECT E13
         USE
         THISFORM.LBL21.CAPTION='B56 製品退出單'         
         B56='B56'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
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
                IF F10='2' AND LEN(ALLTRIM(F14))>4 &&若為委外
                    
                        SELECT G11    &&以入貨退出方式處理
                        SEEK B56.F03
                        IF !FOUND()
                             APPEND BLANK
                             REPLACE F02 WITH B56.F03
                        ENDIF                        
                        
                           *REPLACE F06 WITH F06+IIF(B56.F24<>'2',B56.F04*(-1),0)    &&若為折讓則不減數量
***********************************************************************************8                       
                          REPLACE F04 WITH F04+IIF(B56.F13='1'AND B56.F24='1',B56.F04*(-1),0)  &&若為整組且為退貨則扣數量否則不做數量加減
                           REPLACE F24 WITH F24+ROUND(IIF(B56.F18=1,B56.F04*B56.F17,B56.F04*B56.F17*B56.F18),INT_070)*(-1)
                           IF F23=0 AND F24=0 AND F26=0
                               REPLACE F21 WITH ROUND((F23+F24+F26+F35)/(F03+F04+F06+F33),6)
                           ELSE
                               IF F03+F04+F06<>0
                                  REPLACE F21 WITH ROUND((F23+F24+F26)/(F03+F04+F06),6)
                               ELSE
                                  IF F03>0
                                     REPLACE F21 WITH F23/F03
                                  ELSE
                                     IF F04>0
                                        REPLACE F21 WITH F24/F04
                                     ELSE
                                        IF F06>0
                                           REPLACE F21 WITH F26/F06   
                                        ENDIF   
                                     ENDIF      
                                  ENDIF                                          
                               ENDIF   
                           ENDIF                        
***********************                       
*!*	                    IF B56.F13='1'      &&若是整組加工才丟入存貨明細紀錄    
                        SELE G15                                            &&存貨明細紀錄
                        APPEND BLANK                                              
                        REPLACE  F01 WITH B56.F03                               
                        REPLACE  F02 WITH B56.F02
                        REPLACE  F04 WITH IIF(B56.F13='1' AND B56.F24='1',B56.F04*(-1),0)  &&若整組退且不為折讓則整組退回才加入數量
                        REPLACE  F03 WITH 'R'
                        REPLACE  F05 WITH ROUND((IIF(B56.F18=1,B56.F04*B56.F17,B56.F04*B56.F17*B56.F18))/B56.F04,6)             
                        REPLACE  F07 WITH IIF(B56.F24<>'2','退','折')+B56.F01+'外'+B56.F15+' 製'+B56.F09
                        S=S+1
                        THISFORM.LBL41.CAPTION='G15 存貨明細紀錄'+'  '+ALLTRIM(STR(S))
                        THISFORM.LBL31.CAPTION=F02+' '+F01+' '+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/'+F03+' '+ALLTRIM(STR(F04))    
*!*	                    ENDIF
*!*	                    SELECT G08
*!*	                    SEEK B56.F03+B56.F09+B56.F14
*!*	                    IF FOUND()
*!*	                        REPLACE F03 WITH F03+IIF(B56.F13='1' AND B56.F24<>'2',B56.F04*(-1),0)                                            
*!*	                    ELSE
*!*	                        APPEND BLANK
*!*	                        REPLACE F01 WITH B56.F03
*!*	                        REPLACE F02 WITH B56.F09
*!*	                        REPLACE F03 WITH IIF(B56.F13='1' AND B56.F24<>'2',B56.F04*(-1),0)
*!*	                        REPLACE F05 WITH B56.F14
*!*	                    ENDIF                            
*!*	                       SELECT G10    &&耗料紀錄沒有者亦加入直接材料耗料檔以便佇列排定成本計算優先順序
*!*	                       SEEK B56.F03+SPACE(43)
*!*	                       IF !FOUND()
*!*	                            APPEND BLANK
*!*	                            REPLACE F01 WITH B56.F03    
*!*	                        ENDIF    
*!*	                    SELECT G06
*!*	                    SEEK B56.F14
*!*	                    IF !FOUND()
*!*	                        APPEND BLANK
*!*	                        REPLACE F02 WITH B56.F14
*!*	                   ENDIF            
*!*	                   REPLACE F10 WITH F10+ROUND(IIF(B56.F18=1,B56.F04*B56.F17,B56.F04*B56.F17*B56.F18),INT_070)*(-1)
*!*	                   SELECT G13
*!*	                   SEEK B56.F03
*!*	                   IF !FOUND()
*!*	                       APPEND BLANK
*!*	                       REPLACE F01 WITH B56.F03
*!*	                   ENDIF
*!*	                   REPLACE F06 WITH F06+ROUND(IIF(B56.F18=1,B56.F04*B56.F17,B56.F04*B56.F17*B56.F18),INT_070)*(-1)
                   SELECT A23
                   REPLACE F06 WITH F06+ROUND(IIF(B56.F18=1,B56.F04*B56.F17,B56.F04*B56.F17*B56.F18),INT_070)*(-1)
                ENDIF              
                SELE B56
                SKIP                               
         ENDDO            
         USE           
        SELE G11
        GO TOP
        DO WHILE !EOF()
               REPLACE  F20 WITH F21                     &&計算損耗單價
               REPLACE  F30 WITH ROUND(F10*F20,6)                 &&計算損耗小計
               REPLACE  F19 WITH F21                     &&計算盤差單價
               REPLACE  F29 WITH ROUND(F09*F19,6)                 &&計算盤差小計
               REPLACE  F17 WITH F21                     &&計算產耗單價
               REPLACE  F27 WITH ROUND(F07*F17,6)                 &&計算產耗小計
               REPLACE  F34 WITH IIF(F35=0,0,ROUND(F35/F33,6))       &&計算出貨退回成本單價               
               REPLACE  F14 WITH ROUND(IIF(F24=0 .OR. F04=0,0,F24/F04),6)    &&計算進貨單價
               REPLACE  F15 WITH ROUND(IIF(F25=0 .OR. F05=0,0,F25/F05 ),6)   &&計算出貨單價
               THISFORM.LBL41.CAPTION='G11 計算平均單價'+'  '+ALLTRIM(STR(S))
               THISFORM.LBL31.CAPTION=F02+' '+ALLTRIM(STR(F25))                          
               SKIP
         ENDDO            
         
         THISFORM.LBL21.CAPTION='計算直接人工及製造費用'   
              THISFORM.LBL41.CAPTION=''          
         SELECT G06
         GO TOP
         DO WHILE !EOF()
                REPLACE F05 WITH F03/IIF(F07=0,1,F07)
                REPLACE F08 WITH (F04+F12)/IIF(F07=0,1,F07)
                REPLACE F11 WITH F10/IIF(F07=0,1,F07)
                SELECT A23
                REPLACE F04 WITH F04+G06.F03
                REPLACE F05 WITH F05+(G06.F04+G06.F12)
                SELECT G06
                SKIP
          ENDDO
          SELECT G08
          GO TOP
          DO WHILE !EOF()
                 SELECT G13
                                        SEEK G08.F01
                                        IF FOUND() 
                                            REPLACE F02 WITH F02+G08.F03                  
                                            IF LEN(ALLTRIM(G08.F05))=4    
                                                REPLACE F04 WITH F04+G08.F04*IIF(SEEK(G08.F05,'G06'),G06.F05,0)    &&直接人工
                                            ENDIF    
                                            REPLACE F05 WITH F05+G08.F04*IIF(SEEK(G08.F05,'G06'),G06.F08,0)         &&製造費用                           
                                        ELSE    
                                            APPEND BLANK
                                            REPLACE F01 WITH G08.F01
                                            REPLACE F02 WITH G08.F03  
                                            IF LEN(ALLTRIM(G08.F05))=4
                                                REPLACE F04 WITH F04+G08.F04*IIF(SEEK(G08.F05,'G06'),G06.F05,0)   &&直接人工
                                            ENDIF    
                                            REPLACE F05 WITH F05+G08.F04*IIF(SEEK(G08.F05,'G06'),G06.F08,0)   &&製造費用                                                                     
                                        ENDIF    
                                    *ELSE                                &&丟到別的地方                              
                                    *ENDIF    
                                    SELECT G08
                                    SKIP
          ENDDO
          SELECT G08
          USE
          SELECT G06
          USE
        G09='G09'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        &&其他委外入庫費用明細檔
        IF !USED('&G09')
            SELE 0
            USE (G09) ALIA G09
         ELSE
            SELE G09
         ENDIF
         SET ORDER TO 1                      
          SELECT G09    &&加入其他委外入庫費用
          GO TOP
          DO WHILE !EOF()
             SELECT E14
             SEEK G09.F01+G09.F10
             IF FOUND()
                 SELECT G13
                 SEEK G09.F05
                 IF FOUND()
                     REPLACE F06 WITH F06+G09.F09                     
                 ENDIF
             ELSE
                 SELECT G11
                 SEEK G09.F05
                 IF !FOUND()
                    APPEND BLANK
                    REPLACE F02 WITH G09.F05
                 ENDIF
                 REPLACE F24 WITH F24+G09.F09    
                 IF F03+F04+F06+F33=0
                    REPLACE F21 WITH 0
                    REPLACE F31 WITH 0   
                 ELSE
                    IF F23=0 AND F24=0 AND F26=0
                       REPLACE F21 WITH ROUND((F23+F24+F26+F35)/(F03+F04+F06+F33),6)
                    ELSE
                       REPLACE F21 WITH ROUND((F23+F24+F26)/(F03+F04+F06),6)
                    ENDIF   
                    &&REPLACE F31 WITH F21*F11                    
                 ENDIF
             ENDIF    
             SELECT A23
             REPLACE F10 WITH F10+G09.F09
             SELECT G09
             SKIP    
          ENDDO
          SELECT G09
          USE
          SELECT E14
          USE
         THISFORM.LBL21.CAPTION='計算成本計算優先順序'
         CREATE CURSOR QIRY;
         (F01 C(43),F05 N(3))
         INDEX ON F01 TAG QIRY1
         SET ORDER TO 1
         SELECT G10
        GO TOP
        DO WHILE !EOF()      &&計算佇列之COUNTER
              T=0
              H=F01
              DO WHILE F01=H
                     T=T+1
                     SKIP
              ENDDO
              SELECT QIRY
              APPEND BLANK
              REPLACE F01 WITH H
              REPLACE F05 WITH T
              THISFORM.LBL41.CAPTION='所有加工品排入'+H
              THISFORM.LBL31.CAPTION='COUNTER'+ALLTRIM(STR(T)) 
              SELECT G10    
         ENDDO
         THISFORM.LBL21.CAPTION='計算直接材料耗用'           
        SELECT G10
        GO TOP
        DO WHILE !EOF()
               SELECT G13
               SEEK G10.F01
               IF FOUND()
                    REPLACE F03 WITH F03+(G10.F03+G10.F04)*G10.F05           
                    THISFORM.LBL41.CAPTION=G13.F01                                 
               ENDIF
              THISFORM.LBL31.CAPTION=G10.F01+'     '+G10.F02        
               SELECT G10
               REPLACE F06 WITH (F03+F04)*F05
               SKIP               
        ENDDO                
         THISFORM.LBL21.CAPTION='開始依佇列計算材料成本' 
        SELECT G10
        SET ORDER TO 2
        SELECT F02 FROM G10  WHERE F02 NOT IN (SELE F01 FROM QIRY) INTO CURSOR G10_1 GROUP BY F02
        IF _TALLY>0
            SELECT G10_1
           GO TOP
           DO WHILE !EOF()
                  SELECT G11
                  SEEK G10_1.F02
                  IF FOUND()
                      THISFORM.LBL31.CAPTION=F02
                     IF F03+F04+F06+F33<>0  &&計算期末單價 
                        IF F23=0 AND F24=0 AND F26=0 
                           REPLACE F21 WITH ROUND((F23+F24+F26+F35)/(F03+F04+F06+F33),6)
                        ELSE
                           REPLACE F21 WITH ROUND((F23+F24+F26)/(F03+F04+F06),6)
                        ENDIF   
                       THISFORM.LBL41.CAPTION=TRANSFORM(F21,'@R,999999999.999999')
                     ELSE
                       REPLACE F21 WITH 0
                       REPLACE F31 WITH 0
                       THISFORM.LBL41.CAPTION=TRANSFORM(F21,'@R,999999999.999999')  
                     ENDIF  
                  ENDIF
                  SELECT G10
                  SET ORDER TO 2
                  SEEK G10_1.F02
                  IF FOUND()
                      DO WHILE F02=G10_1.F02 AND !EOF()
                             SELECT QIRY
                             SEEK G10.F01
                             IF FOUND()
                                 REPLACE F05 WITH F05-1  
                            ENDIF
                            SELECT G10
                            SKIP
                      ENDDO     
                  ENDIF
                  SELECT G10_1
                  SKIP
           ENDDO     
      ENDIF      
      DO WHILE .T.
             SELECT F01 FROM QIRY WHERE F05=0 INTO CURSOR QIRY_1
             IF _TALLY>0
                 SELECT QIRY_1
                 GO TOP
                 DO WHILE !EOF()
                                 SELECT G10
                                 SET ORDER TO 1
                                 SEEK QIRY_1.F01
                                 IF FOUND()
                                     DO WHILE F01=QIRY_1.F01
                                            SELECT G13
                                            SEEK G10.F01
                                            IF FOUND()
                                                REPLACE F03 WITH F03+IIF(SEEK(G10.F02,'G11'),G11.F21*(G10.F03+G10.F04),0)                                 
                                            ENDIF
                                            SELECT G10
                                            SKIP
                                     ENDDO
                                ENDIF
                                SELECT G13
                                SEEK QIRY_1.F01
                                IF FOUND()
                                    SELECT G11
                                    SEEK QIRY_1.F01
                                     IF FOUND()
                                          IF F06<>0
                                              REPLACE F26 WITH G13.F03+G13.F04+G13.F05+G13.F06                                          
                                              REPLACE F16 WITH F26/F06
                                          ENDIF   
                                          IF F03+F04+F06+F33<>0
                                              THISFORM.LBL31.CAPTION=F02
                                              IF F23=0 AND F24=0 AND F26=0
                                                 REPLACE F21 WITH ROUND((F23+F24+F26+F35)/(F03+F04+F06+F33),6)
                                              ELSE
                                                 REPLACE F21 WITH ROUND((F23+F24+F26)/(F03+F04+F06),6)
                                              ENDIF  
                                              &&REPLACE F31 WITH F11*F21
                                              REPLACE F17 WITH IIF(F07<>0,F21,0)
                                              REPLACE F27 WITH IIF(F07<>0,F07*F17,0)
                                              REPLACE F19 WITH IIF(F09<>0,F21,0)
                                              REPLACE F29 WITH IIF(F09<>0,F09*F19,0)
                                              REPLACE F20 WITH IIF(F10<>0,F21,0)
                                              REPLACE F30 WITH IIF(F10<>0,F10*F20,0)
                                              THISFORM.LBL41.CAPTION=TRANSFORM(F21,'@R 999999999.999999')                                              
                                              SELECT G15
                                              REPLACE F05 WITH G11.F21 FOR F01=G11.F02 AND (F03<>'A' AND F03<>'B' AND F03<>'C' AND F03<>'D' AND F03<>'F' AND F03<>'M' AND F03<>'R' AND F03<>'1' AND !EMPTY(F03)) AND F04<>0
                                              REPLACE F05 WITH G11.F21 FOR F05=0 AND (F03<>'A' AND F03<>'B' AND F03<>'C' AND F03<>'D' AND F03<>'F')
                                                                           
                                          ENDIF    
                                      ENDIF
                                ENDIF                                
                                SELECT G10
                                SET ORDER TO 2
                                SEEK QIRY_1.F01
                                IF FOUND()
                                    DO WHILE F02=QIRY_1.F01
                                           SELECT QIRY
                                           SEEK G10.F01
                                           IF FOUND()
                                               REPLACE F05 WITH F05-1  
                                           ENDIF
                                           SELECT G10
                                           SKIP
                                   ENDDO
                               ENDIF
                               SELECT QIRY
                               SEEK QIRY_1.F01
                               IF FOUND()
                                    DELE
                               ENDIF
                               SELECT QIRY_1
                               SKIP
                     ENDDO           
             ELSE
                 EXIT
             ENDIF    
      ENDDO
      SELECT G10
      SET ORDER TO 2
      GO TOP
      DO WHILE !EOF()
            IF EMPTY(F02)
                 DELETE
             ELSE   
                 SELECT G11
                 SEEK G10.F02
                 IF FOUND()
                     SELECT G10
                     REPLACE F05 WITH G11.F21
                     REPLACE F06 WITH F05*(F03+F04)
                 ENDIF
             ENDIF   
             SELECT G10
             SKIP
      ENDDO        
      SELECT G13
      USE
      SELECT G11  &&最後修正
      THISFORM.LBL21.CAPTION='最後修正' 
      GO TOP
      DO WHILE !EOF()
         REPLACE F11 WITH F03+F04-F05+F06-F07+F09-F10
         IF F03<=0 AND F04=0 AND F06=0 AND F33=0 AND F11<0 AND F21=0 AND F13<>0
           
           REPLACE F21 WITH F13                     
         ENDIF      
            REPLACE F31 WITH F11 * F21   &&計算期末小計
            THISFORM.LBL31.CAPTION=F02        
         SKIP
      ENDDO
      DELETE FROM G11 WHERE F02 NOT IN(SELE F01 FROM G15)
      USE
      SELECT G10
      USE
         SELE G15         
         S=0
          GO TOP
          H=F01
          SEEK H
          DO WHILE !EOF()   
                 K=0
                 IF FOUND()            
                     DO WHILE F01=H
                            REPLACE  F06 WITH F04+K  
                            K=F06
                            SKIP
                           THISFORM.LBL41.CAPTION='G15 計算累計量'+'  '+ALLTRIM(STR(S))
                           THISFORM.LBL31.CAPTION=F01+' '+ALLTRIM(STR(F06))                       
                     ENDDO
                 ENDIF
                 H=F01
          ENDDO          
         USE                  
         *****
         SELECT A23
         REPLACE F09 WITH .T.
         =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+' 成本計算作業已成功',0+64,'提示訊息視窗')
         THISFORM.CMND3.SETFOCUS           
       ENDIF  
      ENDPROC  
  ****    
      PROCEDURE CMND2.CLICK        
         IF MESSAGEBOX('是否確定要執行成本反結轉作業',4+32+256,'請確認') = 6 
              IF IIF(SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23'),A23.F09=.F.,.F.)
                   =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔尚未成本結轉,不得進行反結轉作業!',0+48,'提示訊息視窗')
                   THISFORM.MTH_LIST.SETFOCUS
                    RETURN             
             ENDIF    
             IF SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23')=.F.
                 =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔尚未建立,不得進行成本反結轉作業!',0+48,'提示訊息視窗')
                 THISFORM.MTH_LIST.SETFOCUS
                 RETURN             
             ENDIF  
             SELECT  A23
             SEEK DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))
             IF FOUND()
                 REPLACE  F09 WITH .F.
             ENDIF            
             =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+' 成本反結轉作業已成功',0+64,'提示訊息視窗')
             THISFORM.CMND3.SETFOCUS  
         ENDIF  
      ENDPROC           
  ****    
      PROCEDURE CMND3.CLICK        
           SET DECIMALS TO  6
           IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'G08'
           IF FOUND()
              DELETE 
           ENDIF        
           CLOSE TABLE ALL
           THISFORM.RELEASE 
           RETURN  
      ENDPROC      
ENDDEFINE            
*****