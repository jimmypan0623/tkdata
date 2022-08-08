***程式名稱:K99.重算K17
CLOSE ALL
CLEAR 
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1
IF !USED('A02')	&&權限設定檔
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'K99'      
IF !USED('C01')
   SELE 0
   USE C01 INDE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO C011 
CSTR=LEFT(C01.F01,1) &&客戶編號的第一碼
IF !USED('D01')
   SELE 0
   USE D01 INDE D01
ELSE
   SELE D01
ENDIF
SET ORDER TO D011
DSTR=LEFT(D01.F01,1) &&廠商編號的第一碼
IF !USED('E02')
   SELE 0
   USE E02 INDE E02
ELSE
   SELE E02
ENDIF
SET ORDER TO E021
ESTR=LEFT(E02.F01,1) &&外包商編號的第一碼
CREATE CURSOR SPLY;
(F01 C(5),F04 C(8))
INDE ON F01 TAG SPLY1
INDE ON F04 TAG SLPY2
APPEND FROM C01
APPEND FROM D01
APPEND FROM E02
SELECT SPLY
SET ORDER TO SPLY1
SELECT C01
USE
SELECT D01
USE
SELECT E02
USE
*****
K99FORM=CREATEOBJECT("TKK99")
K99FORM.SHOW  
DEFINE CLASS TKK99 AS FORM
  CAPTION='K99.重算K17'
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
  NAME='TKK99' 
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
      TKK99_TEMP=CREATEOBJECT("TKK99_TEMP")  
      TKK99_TEMP.SHOW    
      THISFORM.RELEASE      
      IF !USED('A05')
           SELE 0
           USE A05 
      ELSE
           SELE A05
      ENDIF   
      SET ORDER TO 1    
      SEEK sys_oper+'K99'
      IF FOUND()
         DELETE 
      ENDIF   
      UNLOCK ALL                  
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
ENDPROC                  
ENDDEFINE
***
DEFINE CLASS TKK99_TEMP AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='K99.重算K17'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*100
  WIDTH=INT_015*290
  TOP=INT_015*200
  LEFT=INT_015*225  
  MOVABLE=.F.
  MAXBUTTON=.F.
  MINBUTTON=.F.  
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKK99_TEMP'
  ADD OBJECT LBLF_MTH AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*15,;
      WIDTH=INT_015*85,;
      CAPTION='選擇重算年月'  
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      WIDTH=INT_015*70,;
      LEFT=INT_015*90,;
      TOP=INT_015*9
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*45,;
      WIDTH=INT_015*85,;
      CAPTION='起始對象編號'
  ADD OBJECT LBLF011 AS LABEL WITH;
      LEFT=INT_015*145,;
      TOP=INT_015*45,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT TXTF01 AS TEXTBOX WITH;
      LEFT=INT_015*90,;
      TOP=INT_015*39,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      MAXLENGTH=5,;
      NAME='TXTF01'     
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*75,;
      WIDTH=INT_015*85,;
      CAPTION='截止對象編號'
  ADD OBJECT LBLF021 AS LABEL WITH;
      LEFT=INT_015*145,;
      TOP=INT_015*75,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT TXTF02 AS TEXTBOX WITH;
      LEFT=INT_015*90,;
      TOP=INT_015*69,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      MAXLENGTH=5,;
      NAME='TXTF02'           
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*230,;
      TOP=INT_015*20,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所輸入之內容->ALT+Y',;
      NAME='CMND1'
 ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*230,;
      TOP=INT_015*55,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消此畫面!快速鍵->ALT+C'
      NAME='CMND2'  
 PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
         THISFORM.SETALL('FONTSIZE',INT_015*10,'COMMANDBUTTON')       
 ENDPROC            
*********  
 PROCEDURE TXTF01.INTERACTIVECHANGE  
     THISFORM.LBLF011.CAPTION=IIF(SEEK(THIS.VALUE,'SPLY'),SPLY.F04,'')
 ENDPROC    
**************  
 PROCEDURE TXTF02.INTERACTIVECHANGE  
     THISFORM.LBLF021.CAPTION=IIF(SEEK(THIS.VALUE,'SPLY'),SPLY.F04,'')
 ENDPROC  
********確認鍵******
PROCEDURE CMND1.CLICK 
     IF EMPTY(THISFORM.TXTF01.VALUE)     
          =MESSAGEBOX('起始對象編號不得空白!',0+48,'提示訊息視窗') 
          THISFORM.TXTF01.SETFOCUS
          RETURN
     ENDIF    
     IF LEN(ALLTRIM(THISFORM.TXTF01.VALUE)) <> 5     
          =MESSAGEBOX('起始對象編號需為五碼!',0+48,'提示訊息視窗') 
          THISFORM.TXTF01.SETFOCUS
          RETURN
     ENDIF        
     IF EMPTY(THISFORM.TXTF02.VALUE)     
          =MESSAGEBOX('截止對象編號不得空白!',0+48,'提示訊息視窗') 
          THISFORM.TXTF02.SETFOCUS
          RETURN
     ENDIF        
     IF LEN(ALLTRIM(THISFORM.TXTF02.VALUE)) <> 5     
          =MESSAGEBOX('截止對象編號需為五碼!',0+48,'提示訊息視窗') 
          THISFORM.TXTF02.SETFOCUS
          RETURN
     ENDIF           
     K24='K24'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)   &&發票明細檔表頭
     K241='K241'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
     K242='K242'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
     IF !USED('&K24')
	 SELE 0
	 USE (K24) ALIA K24
     ELSE
	 SELE K24
     ENDIF
     SET ORDER TO K241      
     K25='K25'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)	&&發票明細檔表身   
     IF !USED('&K25')
          SELE 0
          USE (K25) ALIA K25
     ELSE
          SELE K25
     ENDIF
     SET ORDER TO 4   
     WAIT WINDOW "資料重算中,請稍後!" AT 0,150 NOWAIT NOCLEAR
     SELECT * FROM K25 WHERE F03>=THISFORM.TXTF01.VALUE AND F03<=THISFORM.TXTF02.VALUE ORDER BY F03 INTO CURSOR SPLY_TEMP
     SELECT SPLY_TEMP
     GO TOP
     DO WHILE !EOF()
            TTL1 = 0
            TTL2 = 0
            TTL3 = 0
            TTL4 = 0
            TTL5 = 0
            TTL6 = 0
            TTL7 = 0
            TTL8 = 0
            SELECT * FROM K25 WHERE ALLTRIM(K25.F03) = ALLTRIM(SPLY_TEMP.F03) ORDER BY K25.F01,K25.F02 INTO TABLE K25_TEMP
            SELECT K25_TEMP
            GO TOP
            DO WHILE !EOF()
                   IF K25_TEMP.F16 <> '2'  AND K25_TEMP.F16 <> '4' 
                      TTL1 = TTL1 + K25_TEMP.F08
                      TTL2 = TTL2 + K25_TEMP.F10
                      TTL3 = TTL3 + K25_TEMP.F12
                   ENDIF    
                   IF K25_TEMP.F16<>'2' AND K25_TEMP.F16 <> '4'  AND K25_TEMP.F09='1'
                      TTL4 = TTL4 + 1
                   ENDIF 
                   IF K25_TEMP.F16<>'2' AND K25_TEMP.F16 <> '4'  AND K25_TEMP.F09='2'
                      TTL5 = TTL5 + 1
                   ENDIF 
                   IF K25_TEMP.F16<>'2' AND K25_TEMP.F16 <> '4'  AND K25_TEMP.F09='3'
                      TTL6 = TTL6 + 1
                   ENDIF                    
                   IF K25_TEMP.F16 = '2' 
                      TTL7 = TTL7 + 1
                   ENDIF                   
                   IF K25_TEMP.F16 = '4' 
                      TTL8 = TTL8 + 1
                   ENDIF                                                                                                                                   
                   SELECT  K25_TEMP
                   SKIP
            ENDDO
            SELECT  K25_TEMP
            USE
            SELECT K24
            SET ORDER TO K242
            SEEK  SPLY_TEMP.F03
            IF FOUND()
                REPLACE K24.F03 WITH TTL1	&&銷售金額 
                REPLACE K24.F04 WITH TTL2	&&營業稅額 
                REPLACE K24.F05 WITH TTL3	&&發票總額 
                REPLACE K24.F06 WITH TTL4	&&應稅張數 
                REPLACE K24.F07 WITH TTL5	&&零稅張數  
                REPLACE K24.F08 WITH TTL6	&&免稅張數 
                REPLACE K24.F09 WITH TTL7	&&作廢張數 
                REPLACE K24.F10 WITH TTL8	&&空白張數
            ELSE
                SELECT K24
                APPEND BLANK 
                REPLACE K24.F02 WITH SPLY_TEMP.F03
                REPLACE K24.F03 WITH TTL1	
                REPLACE K24.F04 WITH TTL2	
                REPLACE K24.F05 WITH TTL3	
                REPLACE K24.F06 WITH TTL4	
                REPLACE K24.F07 WITH TTL5	
                REPLACE K24.F08 WITH TTL6
                REPLACE K24.F09 WITH TTL7	
                REPLACE K24.F10 WITH TTL8
           ENDIF
           SELECT SPLY_TEMP
           SKIP
     ENDDO
     SELECT * FROM K25 WHERE F16 = '2' INTO CURSOR K25A
     IF RECCOUNT('K25A') <> 0
         SELECT K24
         SEEK '作廢 '
         IF FOUND()	
              REPLACE K24.F09 WITH RECCOUNT('K25A')
         ELSE
              SELECT K24
              APPEND BLANK 
              REPLACE K24.F02 WITH '作廢 '
              REPLACE K24.F09 WITH RECCOUNT('K25A')                 
         ENDIF
     ENDIF     
     SELECT SPLY_TEMP
     USE
     SELECT K25A
     USE
     SELECT K24
     USE
     SELECT K25
     USE
    WAIT CLEAR     
     =MESSAGEBOX('重算成功!',0+64,'提示訊息視窗')
     THISFORM.CMND2.SETFOCUS
 ENDPROC   
************
 PROC CMND1.MOUSEENTER      
     LPARAMETERS nButton, nShift, nXCoord, nYCoord 
            THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
            THISFORM.CMND1.TOP=INT_015*21
            THISFORM.CMND1.LEFT=INT_015*231
 ENDPROC     
 PROC CMND1.MOUSELEAVE      
     LPARAMETERS nButton, nShift, nXCoord, nYCoord 
            THISFORM.CMND1.FORECOLOR=RGB(0,0,0)  
            THISFORM.CMND1.TOP=INT_015*20
            THISFORM.CMND1.LEFT=INT_015*230                  
 ENDPROC  
********離開鍵***********
 PROCEDURE CMND2.CLICK     
     CLOSE TABLE ALL
     THISFORM.RELEASE 
     RETURN  
 ENDPROC      
 **********
 PROC CMND2.MOUSEENTER      
     LPARAMETERS nButton, nShift, nXCoord, nYCoord 
            THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
            THISFORM.CMND2.TOP=INT_015*56
            THISFORM.CMND2.LEFT=INT_015*231
 ENDPROC     
 PROC CMND2.MOUSELEAVE      
     LPARAMETERS nButton, nShift, nXCoord, nYCoord 
            THISFORM.CMND2.FORECOLOR=RGB(0,0,0)  
            THISFORM.CMND2.TOP=INT_015*55
            THISFORM.CMND2.LEFT=INT_015*230                  
 ENDPROC   
ENDDEFINE       
