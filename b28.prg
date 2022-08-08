**B28.庫存結轉
CLOSE ALL
CLEAR
IF !USED('A23')
   SELE 0
   USE A23
ELSE
   SELE A23
ENDIF
SET ORDER TO A231
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'B28'
*****
B28FORM=CREATEOBJECT("TKB28")
B28FORM.SHOW  
DEFINE CLASS TKB28 AS FORM
  CAPTION='B28.庫存結轉'
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
  NAME='TKB28' 
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
      TKB28_TEMP=CREATEOBJECT("TKB28_TEMP")  
      TKB28_TEMP.SHOW    
      
      IF !USED('A05')
           SELE 0
           USE A05 
      ELSE
           SELE A05
      ENDIF   
      SET ORDER TO 1    
      SEEK sys_oper+'B28'
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
DEFINE CLASS TKB28_TEMP AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='B28.庫存結轉'
  TOP=INT_015*180
  LEFT=INT_015*250  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*180
  WIDTH=INT_015*300
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.  
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKB28_TEMP'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*25,;
      TOP=INT_015*80,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*215,;
      TOP=INT_015*34,;
      WIDTH=INT_015*70            
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*120,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*11,;
      CAPTION='\<Y.執行結轉',;
      TOOLTIPTEXT='確認執行結轉',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*110,;
      TOP=INT_015*120,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*95,;
      FONTSIZE=INT_015*11,;
      CAPTION='\<V.執行反結轉',;
      TOOLTIPTEXT='確認執行反結轉',;
      NAME='CMND2'      
  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*220,;
      TOP=INT_015*120,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*11,;
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      NAME='CMND3'
  PROCEDURE INIT 
      THISFORM.SETALL('FONTSIZE',INT_015*11,'LABEL')
      THISFORM.SETALL('HEIGHT',INT_015*25,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*11,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*11,'COMMANDBUTTON')  
      THISFORM.SETALL('MOUSEPOINTER',99,'COMMANDBUTTON')  
      THISFORM.SETALL('MOUSEICON','BMPS\harrow.cur','COMMANDBUTTON')     
       THISFORM.LBL1.CAPTION='請輸入欲庫存結轉作業之年月'
      THISFORM.MTH_LIST.SETFOCUS           
  ENDPROC
  ***
  PROC MTH_LIST.INIT        
       SELECT A23
       GO TOP
       WITH THIS
            K=0
            S=0
            DO WHILE !EOF()                
                   K=K+1
                  .ADDITEM(LEFT(DTOS(A23.F01),4) + SUBSTR(DTOC(A23.F01),3,3))
                   IF LEFT(DTOS(DATE()),6) = LEFT(DTOS(A23.F01),6)
                       S=K
                   ENDIF       
                   SKIP
            ENDDO   
            .VALUE=S
       ENDWITH   
  ENDPROC 
  ****  
  PROCEDURE CMND1.CLICK
    THISFORM.LBL2.CAPTION='' 
    IF MESSAGEBOX('是否確定要執行結轉作業',4+32+256,'請確認') = 6	
       InputDate=LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
       NextMonth=VAL(RIGHT(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE),2))+1
       IF NextMonth>12
           NextDate1=ALLTRIM(STR((VAL(LEFT(InputDate,4))+1)))+'/01'
       ELSE
           NextDate1=LEFT(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE),5)+PADL(ALLTRIM(STR(NextMonth)),2,'0')
       ENDIF   
       NextDate=LEFT(DTOS(CTOD(ALLTRIM(NextDate1)+'/01')),6)
       IF IIF(SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23'),A23.F07=.T.,.F.)
           =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔已結轉,不得進行結轉作業!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF    
       IF IIF(SEEK(DTOS(CTOD(ALLTRIM(NextDate1)+'/01')),'A23'),A23.F07=.T.,.F.)
           =MESSAGEBOX(ALLTRIM(NextDate1)+'之月份檔已結轉,不得從'+ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'進行結轉作業!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF             
       ****
       IF SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23')=.F.
           =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔未建立,請洽系統管理人員!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF
       IF FILE('B25'+ALLTRIM(InputDate)+'.DBF')=.F.  
           =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之B25.部門庫存月份檔已遺失,請洽系統管理人員!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN  
       ENDIF       
       **
       IF SEEK(DTOS(CTOD(ALLTRIM(NextDate1)+'/01')),'A23')=.F.
           =MESSAGEBOX(ALLTRIM(NextDate1)+'之欲轉入月份檔未建立,請洽系統管理人員!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF
       IF FILE('B25'+ALLTRIM(NextDate)+'.DBF')=.F.  
           =MESSAGEBOX(ALLTRIM(NextDate1)+'之欲轉入B25.部門庫存月份檔已遺失,請洽系統管理人員!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN  
       ENDIF  
       ***
       B39='B39'+InputDate
       B391='B391'+InputDate    
       IF !USED('&B39')
            SELECT  0
            USE (B39) ALIA B39
       ELSE
            SELECT  B39
       ENDIF
       SET ORDER TO B391  
       GO TOP
       IF !EMPTY(B39.F01)
            =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'尚有單據未過帳,請至B39.未過帳單據查詢中查看!',0+48,'提示訊息視窗')
            THISFORM.MTH_LIST.SETFOCUS
            SELECT B39
            USE              
            RETURN 
       ENDIF       
       C08='C08'+InputDate
       C081='C081'+InputDate    
       IF !USED('&C08')
            SELECT  0
            USE (C08) ALIA C08
       ELSE
            SELECT  C08
       ENDIF
       SET ORDER TO 1  
       SELECT F01 FROM C08 WHERE EMPTY(F10) NOWAIT
       IF _TALLY>0
            =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'還有訂單取消單未確認,請營業部相關人員至C08訂單取單中確認或將之刪除!',0+48,'提示訊息視窗')
            THISFORM.MTH_LIST.SETFOCUS
            SELECT C08
            USE              
            RETURN 
       ENDIF              
       *** 
      THISFORM.MTH_LIST.ENABLED=.F.       
      THISFORM.CMND1.ENABLED=.F.
      THISFORM.CMND2.ENABLED=.F.
      THISFORM.CMND3.ENABLED=.F.
       B1='B25'+InputDate    &&部門別欲結轉月份報表
       IF !USED('&B1')
            SELECT  0
            USE (B1) ALIA B1
       ELSE
            SELECT  B1
       ENDIF
       SET ORDER TO 1                   
       B2='B25'+NextDate    &&部門別欲轉入月份報表
       IF !USED('&B2')
           SELECT  0
           USE (B2) ALIA B2
       ELSE
           SELECT  B2
       ENDIF
       SET ORDER TO 1
       CNT1=0 
       THISFORM.LBL2.CAPTION='轉入之資料筆數：'+ALLTRIM(STR(CNT1))+' 筆'                  
       SELECT  B1
       GO TOP        
       DO WHILE !EOF()
              SELECT  B2
              SEEK B1.F01+B1.F02  
              IF FOUND()
                  REPLACE  F03 WITH B1.F15
                  REPLACE  F15 WITH F03+F04-F05-F06+F07+F08-F09+F10-F11+F13-F14-F17+F18-F19-F20+F21
              ELSE
                  IF B1.F15<>0
                      APPEND BLANK
                      REPLACE  F01 WITH B1.F01
                      REPLACE  F02 WITH B1.F02
                      REPLACE  F03 WITH B1.F15
                      REPLACE  F15 WITH B1.F15
                  ENDIF   
              ENDIF
              CNT1=CNT1+1
              THISFORM.LBL2.CAPTION='轉入之資料筆數：'+ALLTRIM(STR(CNT1))+' 筆'    
              SELECT  B1
              SKIP      
       ENDDO
       SELECT  A23
       SEEK DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))
       IF FOUND()
            REPLACE  F07 WITH .T.
       ENDIF  
       SELECT B39
       USE
       SELECT B1
       USE
       SELECT B2
       USE 
       =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+' 結轉作業已成功',0+64,'提示訊息視窗')
       THISFORM.MTH_LIST.ENABLED=.T.    
       THISFORM.CMND1.ENABLED=.T.
       THISFORM.CMND2.ENABLED=.T.
       THISFORM.CMND3.ENABLED=.T.
       THISFORM.CMND3.SETFOCUS
     ENDIF  
 ENDPROC
*********
  PROCEDURE CMND2.CLICK
    THISFORM.LBL2.CAPTION='' 
    IF MESSAGEBOX('是否確定要執行反結轉作業',4+32+256,'請確認') = 6 
       IF IIF(SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23'),A23.F07=.F.,.F.)
           =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔尚未結轉,不得進行反結轉作業!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF    
       IF SEEK(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),'A23')=.F.
           =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'之月份檔尚未建立,不得進行反結轉作業!',0+48,'提示訊息視窗')
           THISFORM.MTH_LIST.SETFOCUS
           RETURN             
       ENDIF  
       SELECT  A23
       SEEK DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))
       IF FOUND()
            REPLACE  F07 WITH .F.
       ENDIF            
       =MESSAGEBOX(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+' 反結轉作業已成功',0+64,'提示訊息視窗')
       THISFORM.MTH_LIST.ENABLED=.T.   
       THISFORM.CMND1.ENABLED=.T.
       THISFORM.CMND2.ENABLED=.T.
       THISFORM.CMND3.ENABLED=.T.
       THISFORM.CMND3.SETFOCUS           
    ENDIF   
 ENDPROC         
 *****      
 PROCEDURE CMND3.CLICK        
            IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'B28'
           IF FOUND()
              DELETE 
           ENDIF        
      CLOSE TABLE ALL
      THISFORM.RELEASE            
 ENDPROC   
      *****
 PROC CMND1.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND1.TOP=INT_015*121
             THISFORM.CMND1.LEFT=INT_015*16
  ENDPROC     
  PROC CMND1.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND1.TOP=INT_015*120
             THISFORM.CMND1.LEFT=INT_015*15                
  ENDPROC 
 PROC CMND2.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND2.TOP=INT_015*121
             THISFORM.CMND2.LEFT=INT_015*111
  ENDPROC     
  PROC CMND2.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND2.TOP=INT_015*120
             THISFORM.CMND2.LEFT=INT_015*110
 ENDPROC            
 PROC CMND3.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND3.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND3.TOP=INT_015*121
             THISFORM.CMND3.LEFT=INT_015*221
  ENDPROC     
  PROC CMND3.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND3.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND3.TOP=INT_015*120
             THISFORM.CMND3.LEFT=INT_015*220
 ENDPROC        
ENDDEFINE            