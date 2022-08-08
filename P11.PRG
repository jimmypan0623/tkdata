******P11.外購交期回覆跟催單
CLOSE ALL
CLEAR 

*!*	RUN  /N Regsvr32 /s MSMAPI32.OCX
AREA1='ITT'
FLG='0'
FCH=''
CHK=''
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'P11'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('A26')
   SELE 0
   USE A26
ELSE 
  SELE A26
ENDIF
SET ORDER TO 1
SEEK  'INT_MIS'
DIR_PATH=A26.F06

IF !USED('D01')
   SELE 0
   USE D01
ELSE 
   SELE D01
ENDIF
SET ORDER TO 1 
IF !USED('P03')
   SELE 0
   USE P03
ELSE 
   SELE P03
ENDIF
SET ORDER TO 1      
IF !USED('P04')
   SELE 0
   USE P04
ELSE 
   SELE P04
ENDIF
SET ORDER TO 1  
P11FORM=CREATEOBJECT("TKP11")
P11FORM.SHOW       
DEFINE CLASS TKP11 AS FORM
  AUTOCENTER=.T.
  CAPTION='P11.外購交期回覆跟催單'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*270
  WIDTH=INT_015*300
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKP11'
   ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*70,;
      TOP=INT_015*30,;
      AUTOSIZE=.T.,;
      CAPTION='廠商編號'
   ADD OBJECT LBL31 AS LABEL WITH;
      LEFT=INT_015*205,;
      TOP=INT_015*30,;
      AUTOSIZE=.T.,;
      CAPTION=''   
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*90,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入下單起始日'
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*150,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入下單終止日'   
  ADD OBJECT TXT3 AS TXT3 WITH;
      LEFT=INT_015*140,;
      TOP=INT_015*25,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*60,;
      FONTSIZE=INT_015*12,; 
      MAXLENGTH=5,;
      NAME='TXT3'           
  ADD OBJECT TXT1 AS TEXTBOX WITH;
      LEFT=INT_015*140,;
      TOP=INT_015*85,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      MAXLENGTH=8,;
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*140,;
      TOP=INT_015*145,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      MAXLENGTH=8,;
      NAME='TXT2'           
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*40,;
      TOP=INT_015*210,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所查詢條件',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*210,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      NAME='CMND2'
      PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
      ENDPROC
      PROCEDURE CMND1.CLICK
         GHK1=THISFORM.TXT1.VALUE
         GHK2=THISFORM.TXT2.VALUE
         GHK3=THISFORM.LBL31.CAPTION
         SELE D07.F01,D07.F02,D07.F03,D07.F04,D07.F07,D07.F13 FROM D07;
         WHERE D07.F09<>'1' AND RIGHT(D07.F01,8)>=THISFORM.TXT1.VALUE AND  RIGHT(D07.F01,8)<=THISFORM.TXT2.VALUE;
         AND D07.F05=ALLTRIM(THISFORM.TXT3.VALUE)  AND D07.F01=ANY(SELE F02 FROM P03 WHERE !EMPTY(F07)) ORDER BY D07.F02,D07.F03 INTO CURSOR ITT_1 NOWAIT
         IF _TALLY>0  &&AND D07.F01=ANY(SELE F02 FROM P03 WHERE !EMPTY(F07))
             CREATE CURSOR ITT;
             (F01 C(10),F02 C(43),F03 D(8),F04 N(13),F07 C(20),F11 N(13),F12 N(4),F13 C(14))
             INDE ON F02+DTOS(F03) TAG ITT1
             SET ORDER TO 1             
             SELE ITT_1
             GO TOP
             DO WHILE !EOF()
                SELE ITT
                APPEND BLANK
                REPLACE  F01 WITH ITT_1.F01
                REPLACE  F02 WITH ITT_1.F02               
                REPLACE  F03 WITH ITT_1.F03
                REPLACE  F04 WITH ITT_1.F04
                REPLACE  F07 WITH ITT_1.F07
                REPLACE F13 WITH ITT_1.F13
                SELE ITT_1
                SKIP   
             ENDDO
             SELE ITT_1
             USE   
             SELE ITT
            GO TOP                         &&計算累計量
            H=F02  
            T=0          
            SEEK H
            DO WHILE !EOF()   
               K=0          
               T=T+1            
*               IF FOUND()                     
                  DO WHILE F02=H
                     K=K+F04      
                     REPL F12 WITH T                                                                                                                                                                              
                     SKIP
                  ENDDO
                  SKIP -1
                  REPL F11 WITH K
                  SKIP 
*               ENDIF
                  H=F02
             ENDDO              
             GO TOP
             ITT_SEEK=CREATEOBJECT("ITT_SEEK")  
             ITT_SEEK.SHOW                           
         ELSE
             =MESSAGEBOX('無此範圍資料!',0+48,'提示訊息視窗')
             RETURN    
         ENDIF
         THISFORM.CMND2.CLICK
      ENDPROC
      PROCEDURE CMND2.CLICK        
           IF !USED('A05')
              SELECT 0
              USE A05
           ELSE
              SELECT A05
           ENDIF   
           SET ORDER TO 1
           SEEK sys_oper+'P11'
           IF FOUND()
              DELETE 
           ENDIF      
           CLOSE TABLE ALL
           THISFORM.RELEASE 
           RETURN  
      ENDPROC      
ENDDEFINE          

**********************************************************************************************
DEFINE CLASS ITT_SEEK AS FORM
*!*	  AUTOCENTER=.T.  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*800
  FONTSIZE=INT_015*12
  CONTROLBOX=.F.
  BORDERSTYLE=1
  SHOWTIPS=.T.
  SHOWWINDOW=1
  MOVABLE=.F.
  WINDOWTYPE=1
  NAME='ITT_SEEK'
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T. 
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;            
      TOP=0,;
      LEFT=INT_015*450               
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=0,;
      ENABLED=.T.,;
      LEFT=INT_015*300    
  ADD OBJECT TRN_FILE AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*520,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<T.轉檔'    
  ADD OBJECT GRID1 AS GRID1 WITH;   
      AUTOCENTER=.T.,;
      TOP=INT_015*30,;
      WIDTH=INT_015*800,;      
      HEIGHT=INT_015*500,;  
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*18,;     
      DELETEMARK=.F.,;
      READONLY=.T.,;
      RECORDSOURCETYPE=1,;
      COLUMNCOUNT=8,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7',;
      COLUMN8.NAME='COLUMN8'
      PROCEDURE INIT 
         THISFORM.CAPTION='外購交期回覆跟催單  廠商:'+GHK3+'  從:'+GHK1+'  到:'+GHK2         
         THISFORM.GRID1.READONLY=.T. 
         THISFORM.GRID1.RECORDSOURCE='ITT'         
         THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限               
         THISFORM.TRN_FILE.ENABLED=IIF(A02.F08='*',.T.,.F.)   &&判斷有無轉檔的權限
         THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F02'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*261
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='希望交期'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.F03'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*65                                                         	                                       
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='採購單號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F01'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='採購數量'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F04'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*70               
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='累計數量'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.F11'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*70     
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='採購用途'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='ITT.F07'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*149 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='備註說明'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(ITT.F01+ITT.F02,'P04'),P04.F13,'')"
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80            
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='運貨方式'        
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="ITT.F13"
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80                                             
               THISFORM.GRID1.SETFOCUS                 
         JK='料品編號'             
         THISFORM.GRID1.SETFOCUS        
      ENDPROC
      PROC KEY_LIST.INIT 
           WITH THIS            
               .ADDITEM('依  料號  +希望交期排列')
               .VALUE=1           
           ENDWITH      
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS XCOLINDEX     
         FCH=THISFORM.ACTIVECONTROL.NAME  
         CHK=''         
         THISFORM.REFRESH
      ENDPROC       
      PROCEDURE TRN_FILE.CLICK
         CREATE CURSOR EXLY;
         (NO C(4),PART_NO C(30),QTY N(9),UNIT_PRICE N(12,2),DTE C(8),PO_NO C(20),FREIGHT C(10),USER_ID C(3),CONSIGNEE C(3),REMARK C(40),CUST C(8))
         INDE ON PART_NO+DTE+PO_NO TAG EXLY_1
         SET ORDER TO 1
         SELE ITT.F02,P04.F15,ITT.F04,P04.F05,ITT.F03,ITT.F01,P03.F04,ITT.F13,P04.F13,ITT.F07 INTO CURSOR EXLY2;
         FROM ITT,P04,P03 WHERE P04.F01+P04.F02=ITT.F01+ITT.F02 AND P03.F02=ITT.F01 
         SELE EXLY2
         GO TOP       
         DO WHILE !EOF()
            SELE EXLY
            APPEND BLANK            
            REPLACE PART_NO WITH IIF(ALLTRIM(EXLY2.F15)='J' .OR. EMPTY(EXLY2.F15),'','('+ALLTRIM(EXLY2.F15)+')')+EXLY2.F02
            REPLACE QTY WITH EXLY2.F04_A
            REPLACE UNIT_PRICE WITH ROUND(EXLY2.F05*1000,IIF(IIF(SEEK(EXLY2.F01+EXLY2.F02,'P04'),P04.F04,'')=INT_011,0,2))
            REPLACE DTE WITH DTOS(EXLY2.F03)
            REPLACE PO_NO WITH LEFT(EXLY2.F01,2)+'-'+RIGHT(EXLY2.F01,8)
*!*	            REPLACE FREIGHT WITH IIF(ALLTRIM(EXLY2.F13)='空運','AIR',IIF(ALLTRIM(EXLY2.F13)='DHL','DHL',IIF(ALLTRIM(EXLY2.F13)='FEDEX','FEDEX',IIF(ALLTRIM(EXLY2.F13)='直取','W/H',''))))
            REPLACE FREIGHT WITH ICASE(ALLTRIM(EXLY2.F13_A) = '海運','',ALLTRIM(EXLY2.F13_A) = '空運','AIR',ALLTRIM(EXLY2.F13_A) = '快遞','FEDEX',ALLTRIM(EXLY2.F13_A))
            REPLACE USER_ID WITH LEFT(EXLY2.F04_B,3)
            REPLACE CONSIGNEE WITH IIF(LEFT(PO_NO,2)='DZ','TWJ','')
            REPLACE REMARK WITH ALLTRIM(EXLY2.F13_B)
            REPLACE CUST WITH SUBSTR(EXLY2.F07,12,8)
            SELE EXLY2
            SKIP
          ENDDO
          SELE EXLY2
          USE
          SELE EXLY
          GO TOP
          K=0
          DO WHILE !EOF()
             K=K+1
             REPLACE NO WITH ALLTRIM(STR(K))
             SKIP
          ENDDO   
          FILENAME='TK'+ALLTRIM(GHK2)
          GCDELIMFILE=PUTFILE('儲存路徑',GHK1+'-'+GHK2,'XLS')
          IF EMPTY(GCDELIMFILE)  && ESC PRESSED
             SELE EXLY
             USE 
             SELE ITT
             THISFORM.GRID1.SETFOCUS
             RETURN
          ELSE
              ON ERROR DO FILE_IMPACT
              IF _TALLY >0
           	   COPY TO (GCDELIMFILE) TYPE XL5    
                   =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗') 
	       ENDIF        
              SELE EXLY
              USE 
              SELE ITT              
*!*	              THISFORM.GRID1.SETFOCUS              
          ENDIF
****
           THISFORM.GRID1.SETFOCUS  
           RETURN
      ENDPROC 
   ENDDEFINE   
***************************************************特殊輸入欄位設定-----------廠商
DEFINE CLASS TXT3 AS TEXTBOX
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0 
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
       ELSE
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02    
       ENDIF          
       IF _TALLY>0
          VDR_SEEK=CREATEOBJECT("VDR_SEEK")  
          VDR_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'        
       ENDIF   
    ELSE
       IF LEN(ALLTRIM(THIS.VALUE))=4
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE F02=ALLTRIM(THIS.VALUE)
          DO CASE
             CASE _TALLY=1
                  THIS.VALUE=VDR.F01
                  THIS.REFRESH
             CASE _TALLY>1
                  VDR_SEEK=CREATEOBJECT("VDR_SEEK")  
                  VDR_SEEK.SHOW    
                  THIS.VALUE=FLG
                  THIS.REFRESH
                  FLG='0'        
          ENDCASE      
       ENDIF    
    ENDIF      
       THISFORM.LBL31.CAPTION=IIF(SEEK(THIS.VALUE,'D01'),D01.F04,'')   
  ENDPROC
ENDDEFINE    
*************************************************************列印
PROCEDURE PNT_PRC  
     REPORT FORM ALLTRIM(INT_116)+'P11' TO PRINT PROMPT PREVIEW     
ENDPROC       
