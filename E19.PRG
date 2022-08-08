****程式名稱 E19重算MRP
CLOSE ALL
CLEAR
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'E19'
IF !USED('C05')
   SELE 0
   USE C05
ELSE
  SELE 0
ENDIF
SET ORDER TO C055
SET FILTER TO !EMPTY(F02)
IF !USED('D07')
   SELE 0
   USE D07
ELSE
  SELE 0
ENDIF
SET ORDER TO D075     
SET FILTER TO !EMPTY(F02) &&AND F09='1'
IF !USED('E06')
   SELE 0
   USE E06
ELSE
  SELE 0
ENDIF
SET FILTER TO !EMPTY(F02)  &&AND F09='1'
SET ORDER TO E064     
SET FILTER TO !EMPTY(F02)
IF !USED('E16')
   SELE 0
   USE E16
ELSE
  SELE 0
ENDIF
SET ORDER TO E165 
IF !USED('E08')
   SELE 0
   USE E08
ELSE
   SELE E08
ENDIF
SET ORDER TO E082      
*****
E19FORM=CREATEOBJECT("TKE19")
E19FORM.SHOW  
DEFINE CLASS TKE19 AS FORM
  CAPTION='E19.重算MRP'
*!*	  AUTOCENTER=.T.
  CONTROLBOX=.F.
  BORDERSTYLE=1  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  CONTROLCOUNT=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*791
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKE19' 
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*131,;
      TOP=INT_015*340,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<C.選擇',;
      TOOLTIPTEXT='選擇此作業畫面!快速鍵->ALT+A'
      NAME='CMND1'        
  PROCEDURE  INIT       
      v_scrx=SYSMETRIC(1)
      v_scry=SYSMETRIC(2)
      DO CASE
           CASE  v_scrx=640 AND v_scry=480 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=800 AND v_scry=600 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=1024 AND v_scry=768
  	              THISFORM.PICTURE='BMPS\XP1024768.JPG' 
           OTHERWISE
     	              THISFORM.PICTURE='BMPS\XP1024768.JPG' 
      ENDCASE  
 ENDPROC
 PROCEDURE ACTIVATE
       THISFORM.CMND1.SETFOCUS
       THISFORM.CMND1.CLICK	
 ENDPROC
 PROCEDURE CMND1.CLICK  
      TKE19_TEMP=CREATEOBJECT("TKE19_TEMP")  
      TKE19_TEMP.SHOW    
      THISFORM.RELEASE      
      IF !USED('A05')
           SELE 0
           USE A05 
      ELSE
           SELE A05
      ENDIF   
      SET ORDER TO 1    
      SEEK sys_oper+'E19'
      IF FOUND()
         DELETE 
      ENDIF   
      UNLOCK ALL                  
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
ENDPROC                  
ENDDEFINE
*****      
DEFINE CLASS TKE19_TEMP AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='E19.重算MRP'
  TOP=INT_015*80
  LEFT=INT_015*200     
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*400
  WIDTH=INT_015*400
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.    
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKE19_TEMP'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='進貨計劃'
  ADD OBJECT LBL11 AS LABEL WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*40,;
      autosize=.t.,;
      CAPTION='/'+ALLTRIM(STR(RECCOUNT('C05')))
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION='出貨計劃'     
  ADD OBJECT LBL21 AS LABEL WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION='/'+ALLTRIM(STR(RECCOUNT('D07')))
  ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION='生產計劃'     
  ADD OBJECT LBL31 AS LABEL WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION='/'+ALLTRIM(STR(RECCOUNT('E06')))      
  ADD OBJECT LBL4 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*220,;
      AUTOSIZE=.T.,;
      CAPTION='領用計劃'     
  ADD OBJECT LBL41 AS LABEL WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*220,;
      AUTOSIZE=.T.,;
      CAPTION='/'+ALLTRIM(STR(RECCOUNT('E16')))            
  ADD OBJECT LBLZ AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*280,;
      AUTOSIZE=.T.,;
      CAPTION='MRP檔'
  ADD OBJECT TXT1 AS TEXTBOX WITH;
      LEFT=INT_015*75,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*120,;     
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*75,;
      TOP=INT_015*95,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*120,;
      NAME='TXT2'          
  ADD OBJECT TXT3 AS TEXTBOX WITH;
      LEFT=INT_015*75,;
      TOP=INT_015*155,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*120,;     
      NAME='TXT3'    
  ADD OBJECT TXT4 AS TEXTBOX WITH;
      LEFT=INT_015*75,;
      TOP=INT_015*215,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*120,;
      NAME='TXT4'                
  ADD OBJECT TXTZ AS TEXTBOX WITH;
      LEFT=INT_015*75,;
      TOP=INT_015*275,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*120,;
      NAME='TXTZ'                
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*120,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認後開始執行重算MRP',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<C.離開',;
      TOOLTIPTEXT='離開執行,並離開此查詢尋畫面!快速鍵->ALT+C'
      NAME='CMND2'
  PROCEDURE INIT 
      THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
      THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*11,'COMMANDBUTTON')  
      THISFORM.SETALL('MOUSEPOINTER',99,'COMMANDBUTTON')  
      THISFORM.SETALL('MOUSEICON','BMPS\harrow.cur','COMMANDBUTTON')
      THISFORM.CMND1.SETFOCUS           
  ENDPROC
  PROCEDURE CMND1.CLICK
      THISFORM.CMND1.ENABLED=.F.
      THISFORM.CMND2.ENABLED=.F.
      CREATE CURSOR E08TMP;
      (F01 D(8),F02 C(43),F05 N(13,3),F08 N(13,3),F11 N(13,3),F14 N(13,3),F17 N(13,3),F20 N(13,3),F21 N(13,3))
      INDE ON DTOS(F01)+F02 TAG E08TMP1
      INDE ON F02+DTOS(F01) TAG E08TMP2         
      THISFORM.TXT1.VALUE=0
      SELECT  C05
      GO TOP
      DO WHILE !EOF()
         IF !EMPTY(F15)
             SELECT  E08TMP                           &&寫入MRP紀錄檔             
             SET ORDER TO E08TMP2
             SEEK C05.F02+DTOS(MRP_PRD(IIF(EMPTY(C05.F10),C05.F03,C05.F10)))
             IF FOUND()
                 REPLACE  F14 WITH F14+C05.F04
             ELSE
                 APPEND BLANK
                 REPLACE  F01 WITH MRP_PRD(IIF(EMPTY(C05.F10),C05.F03,C05.F10))
                 REPLACE  F02 WITH C05.F02
                 REPLACE  F14 WITH F14+C05.F04
              ENDIF     
              FLG=C05.F02
              DO MRP1_PRC        &&MRP即時計算程序        
              THISFORM.TXT1.VALUE=THISFORM.TXT1.VALUE+1                          
          ENDIF    
          SELECT  C05  
          SKIP
      ENDDO     
      THISFORM.TXT2.VALUE=0 
      SELECT  D07
      GO TOP
      DO WHILE !EOF()
            SELECT E08TMP                           &&寫入MRP紀錄檔
            SET ORDER TO E08TMP2
            SEEK D07.F02+DTOS(MRP_PRD(D07.F03))  
            IF FOUND()
                REPLACE  F05 WITH F05+D07.F04
            ELSE
                APPEND BLANK
                REPLACE  F01 WITH MRP_PRD(D07.F03)
                REPLACE  F02 WITH D07.F02
                REPLACE  F05 WITH F05+D07.F04
            ENDIF     
            FLG=D07.F02
            DO MRP1_PRC        &&MRP即時計算程序   
            THISFORM.TXT2.VALUE=THISFORM.TXT2.VALUE+1
            SELECT  D07
            SKIP                   
      ENDDO   
      THISFORM.TXT3.VALUE=0 
      SELECT  E06
      GO TOP
      DO WHILE !EOF()
             SELECT  E08TMP                           &&寫入MRP紀錄檔
             SET ORDER TO E08TMP2
             SEEK E06.F02+DTOS(MRP_PRD(E06.F03))  
             IF FOUND()
                 REPLACE  F08 WITH F08+E06.F04
             ELSE
                 APPEND BLANK
                 REPLACE  F01 WITH MRP_PRD(E06.F03)
                 REPLACE  F02 WITH E06.F02
                 REPLACE  F08 WITH F08+E06.F04
             ENDIF     
             FLG=E06.F02
             DO MRP1_PRC        &&MRP即時計算程序   
             THISFORM.TXT3.VALUE=THISFORM.TXT3.VALUE+1
             SELECT  E06
             SKIP                   
       ENDDO            
       THISFORM.TXT4.VALUE=0 
       SELECT  E16
       GO TOP
       DO WHILE !EOF()
              SELECT  E08TMP                           &&寫入MRP紀錄檔
              SET ORDER TO E08TMP2
              SEEK E16.F02+DTOS(MRP_PRD(IIF(EMPTY(E16.F10),E16.F03,E16.F10)))  
              IF FOUND()
                  REPLACE  F11 WITH F11+E16.F04
              ELSE
                  APPEND BLANK
                  REPLACE  F01 WITH MRP_PRD(IIF(EMPTY(E16.F10),E16.F03,E16.F10))
                  REPLACE  F02 WITH E16.F02
                  REPLACE  F11 WITH F11+E16.F04
              ENDIF     
              FLG=E16.F02
              DO MRP1_PRC        &&MRP即時計算程序   
              THISFORM.TXT4.VALUE=THISFORM.TXT3.VALUE+1
              SELECT  E16
              SKIP                   
       ENDDO                     
       THISFORM.TXTZ.VALUE=0
       SELECT  E08
       DELETE  ALL
       SELE E08TMP
       GO TOP
       DO WHILE !EOF()
              SELE E08
              APPEND BLANK
              REPLACE  F01 WITH E08TMP.F01
              REPLACE  F02 WITH E08TMP.F02
              REPLACE  F05 WITH E08TMP.F05
              REPLACE  F08 WITH E08TMP.F08
              REPLACE  F11 WITH E08TMP.F11
              REPLACE  F14 WITH E08TMP.F14
              REPLACE  F17 WITH E08TMP.F17
              REPLACE  F20 WITH E08TMP.F20
              THISFORM.TXTZ.VALUE=THISFORM.TXTZ.VALUE+1
              SELECT  E08TMP
              SKIP
       ENDDO   
       =MESSAGEBOX('重算MRP作業已成功',0+64,'提示訊息視窗')  
      THISFORM.CMND1.ENABLED=.T.
      THISFORM.CMND2.ENABLED=.T.    
      THISFORM.CMND2.SETFOCUS 
      ENDPROC   
 PROCEDURE CMND2.CLICK        
            IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'E19'
           IF FOUND()
              DELETE 
           ENDIF        
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
 ENDPROC      
**
 PROC CMND1.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND1.TOP=INT_015*331
             THISFORM.CMND1.LEFT=INT_015*121
  ENDPROC     
  PROC CMND1.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND1.TOP=INT_015*330
             THISFORM.CMND1.LEFT=INT_015*120                
  ENDPROC 
 PROC CMND2.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND2.TOP=INT_015*331
             THISFORM.CMND2.LEFT=INT_015*211
  ENDPROC     
  PROC CMND2.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND2.TOP=INT_015*330
             THISFORM.CMND2.LEFT=INT_015*210
 ENDPROC        
ENDDEFINE            
************自訂整理MRP暫存檔程序
PROC MRP1_PRC
           MRP1=0
           SELE E08tmp
           SEEK FLG
           IF FOUND()
               DO WHILE F02=FLG
                  REPL F20 WITH F05+F08-F11-F14-F17
                  REPL F20 WITH F20+MRP1
                  MRP1=F20
                  SKIP
               ENDDO   
           ENDIF   
          FLG='0' 
ENDPROC
*****