close all
clear 
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
SEEK sys_oper+'C49'
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
IF !USED('C00')
   SELE 0
   USE C00
ELSE 
  SELE C00
ENDIF
SET ORDER TO 1
IF !USED('C01')
   SELE 0
   USE C01
ELSE 
  SELE C01
ENDIF
SET ORDER TO 1
IF !USED('C04')
   SELE 0
   USE C04
ELSE 
  SELE C04
ENDIF
SET ORDER TO 1
C49form=createobject("tkC49")
C49form.show       
define class tkC49 as form
*!*	  autocenter=.t.
  caption='C49.客戶訂單接單搜尋範圍'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*320
  width=INT_015*420
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkC49'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入日期範圍'    
  ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*250,;
      TOP=INT_015*40,;
      FONTSIZE=INT_015*30,;
      AUTOSIZE=.T.,;
      CAPTION='~'               
  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT TXT1 AS TEXTBOX WITH;
      LEFT=INT_015*130,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      FONTSIZE=INT_015*12,;    
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*300,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      NAME='TXT2'          
 ADD OBJECT O_TPE AS OPTIONGROUP WITH;
      LEFT=INT_015*30,;
      TOP=INT_015*130,;
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
      NAME='O_TPE'                   
ADD OBJECT D_TPE AS OPTIONGROUP WITH;
      LEFT=INT_015*140,;
      TOP=INT_015*140,;
      HEIGHT=INT_015*40,;    
      WIDTH=INT_015*60,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='明細',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='彙總',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;          
      NAME='D_TPE'                        
ADD OBJECT OUTPUT AS OPTIONGROUP WITH;
      LEFT=INT_015*250,;
      TOP=INT_015*140,;
      HEIGHT=INT_015*40,;    
      WIDTH=INT_015*120,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='轉EXCEL檔',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='列印報表',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;          
      NAME='OUTPUT'                         
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*250,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所查詢條件',;
      NAME='CMND1'
  add object cmnd2 as commandbutton with;
      left=INT_015*190,;
      top=INT_015*250,;
      height=INT_015*40,;
      width=INT_015*80,;
      FONTSIZE=INT_015*12,;
      caption='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.O_TPE.VALUE=1
         THISFORM.OUTPUT.VALUE=1         
         THISFORM.TXT1.VALUE=DATE()
         THISFORM.TXT2.VALUE=DATE()
      ENDPROC
**********************************開始搜尋指定範圍之C04內容     
      PROCEDURE D_TPE.INTERACTIVECHANGE
         IF THIS.VALUE=2
            THISFORM.OUTPUT.VALUE=1 
            IF THISFORM.TXT2.VALUE-THISFORM.TXT1.VALUE>31
               =MESSAGEBOX('彙總的日期範圍天數不得大於31',0+48,'提示訊息視窗')
               THISFORM.TXT1.SETFOCUS
               RETURN
            ENDIF
         ENDIF
      ENDPROC    
       PROCEDURE OUTPUT.INTERACTIVECHANGE
         IF THISFORM.D_TPE.VALUE=2 AND THIS.VALUE=2
            =MESSAGEBOX('彙總資料無法轉入報表',0+48,'提示訊息視窗')
            THIS.VALUE=1
            THIS.REFRESH
         ENDIF
      ENDPROC         
      PROCEDURE CMND1.CLICK      
         IF THISFORM.D_TPE.VALUE=2 AND THISFORM.TXT2.VALUE-THISFORM.TXT1.VALUE>31
            =MESSAGEBOX('彙總的日期範圍天數不得大於31',0+48,'提示訊息視窗')
            THISFORM.TXT1.SETFOCUS
            RETURN
         ENDIF
         IF THISFORM.D_TPE.VALUE=1
             SELECT C04.F02 客戶編號,C01.F05 客戶簡稱,C04.F01 接單日期,C04.F04 料品編號,IIF(B01.F97='Y','製品','部品') 類別,C04.F03 訂單號碼,C04.F15-C04.F21 計價數量,C04.F06 幣別,C04.F07 銷售單價,C04.F08 小計,C00.F02 匯率,ROUND(C04.F08*C00.F02,INT_069) 本幣別小計,IIF(SEEK(C01.F18,'A01'),A01.F03,'') 業務A,;
             IIF(SEEK(C01.F33,'A01'),A01.F03,'') 業務B,IIF(SEEK(C01.F23,'A01'),A01.F03,'')  業務助理 FROM C04,C00,C01,B01 WHERE C04.F03=ANY(SELE F02 FROM C03 WHERE F15='1') AND C04.F15-C04.F21>0 AND C01.F01=C04.F02 AND C00.F01=C04.F06;
             AND C04.F01>=THISFORM.TXT1.VALUE AND C04.F01<=THISFORM.TXT2.VALUE AND C04.F04=B01.F01 AND B01.F97$(ICASE(THISFORM.O_TPE.VALUE=1,'N',THISFORM.O_TPE.VALUE=2,'Y','YN'));
              INTO CURSOR BTX ORDER BY C04.F02,C04.F01
             SELECT BTX
             IF _TALLY>0
                GHK1=DTOC(THISFORM.TXT1.VALUE)
                GHK2=DTOC(THISFORM.TXT2.VALUE)      
                FILE_NAME='接單日報表'+ICASE(THISFORM.O_TPE.VALUE=1,'部品',THISFORM.O_TPE.VALUE=2,'製品','全部')++IIF(THISFORM.D_TPE.VALUE=1,'明細','彙總')                 
                IF THISFORM.OUTPUT.VALUE=1
	               GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	               IF EMPTY(GCDELIMFILE)  && ESC PRESSED
  	                  RETURN
	               ELSE
	                 ON ERROR DO FILE_IMPACT
	                 COPY TO (GCDELIMFILE) TYPE XL5   
	                 =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
  	               ENDIF
  	            ELSE
                   REPORT FORM ALLTRIM(INT_116)+'C49' TO PRINT PROMPT PREVIEW  	           
  	            ENDIF             
             ELSE
                =MESSAGEBOX('無此範圍資料 !',0+48,'提示訊息視窗')
                THISFORM.TXT1.SETFOCUS
                RETURN
             ENDIF
         ELSE
             CREATE TABLE BTT (F01 C(5),F02 C(8))
             H=THISFORM.TXT1.VALUE
             
             FOR I=0 TO THISFORM.TXT2.VALUE-THISFORM.TXT1.VALUE
                 L='D'+DTOS(H)
                 ALTER TABLE BTT ADD COLUMN &L N(12,2)
                 H=H+1
             ENDFOR
             ALTER TABLE BTT ADD COLUMN 總計 N(16,2)
             ALTER TABLE BTT ADD COLUMN 業務人員 C(8)
             ALTER TABLE BTT ADD COLUMN 業務助理 C(8)
             INDEX ON F01 TAG BTT
             SET ORDER TO 1
             SELECT C04.F02,C01.F05,C04.F01,SUM(C04.F08*C00.F02) GROUP BY C04.F02,C04.F01 FROM C04,C00,C01 WHERE C04.F03=ANY(SELE F02 FROM C03 WHERE F15='1') AND C04.F01>=THISFORM.TXT1.VALUE AND C04.F01<=THISFORM.TXT2.VALUE AND C00.F01=C04.F06 AND;
             C01.F01=C04.F02 AND C04.F04=ANY(SELE F01 FROM B01 WHERE F97$(ICASE(THISFORM.O_TPE.VALUE=1,'N',THISFORM.O_TPE.VALUE=2,'Y','YN'))) INTO CURSOR BTX ORDER BY C04.F02,C04.F01 NOWAIT
             IF _TALLY>0
                SELECT BTX
                GO TOP
                DO WHILE !EOF()
                   SELECT BTT
                   SEEK BTX.F02
                   IF !FOUND()
                       APPEND BLANK                 
                       REPLACE F01 WITH BTX.F02
                       REPLACE F02 WITH BTX.F05
                   ENDIF        
                   M=FIELD(BTX.F01-THISFORM.TXT1.VALUE+3)
                   REPLACE &M WITH BTX.SUM_EXP_4
                   SELECT BTX
                   SKIP
                ENDDO
                SELECT BTT
                APPEND BLANK
                REPLACE F01 WITH '總計'                
                H=THISFORM.TXT1.VALUE                
                FOR I=0 TO THISFORM.TXT2.VALUE-THISFORM.TXT1.VALUE
                    L='D'+DTOS(H)
                    CALCULATE SUM(&L) TO KK
                    GO BOTT
                    REPLACE &L WITH KK
                    H=H+1
                ENDFOR                       
                SELECT BTT
                GO TOP
                DO WHILE !EOF()
                    TTL=0
                    H=THISFORM.TXT1.VALUE                
                    FOR I=0 TO THISFORM.TXT2.VALUE-THISFORM.TXT1.VALUE
                        L='D'+DTOS(H)                    
                        TTL=TTL+&L
                        H=H+1
                    ENDFOR              
                    REPLACE 總計 WITH TTL    
                    REPLACE 業務人員 WITH IIF(SEEK(IIF(SEEK(F01,'C01'),C01.F33,SPACE(5)),'A01'),A01.F03,SPACE(8))
                    REPLACE 業務助理 WITH IIF(SEEK(IIF(SEEK(F01,'C01'),C01.F23,SPACE(5)),'A01'),A01.F03,SPACE(8))
                   SELECT BTT
                   SKIP
                ENDDO
                   FILE_NAME='接單日報表'+ICASE(THISFORM.O_TPE.VALUE=1,'部品',THISFORM.O_TPE.VALUE=2,'製品','全部')+IIF(THISFORM.D_TPE.VALUE=1,'明細','彙總')
                   
	               GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	               IF EMPTY(GCDELIMFILE)  && ESC PRESSED
  	                  RETURN
	               ELSE
	                 ON ERROR DO FILE_IMPACT
	                 COPY TO (GCDELIMFILE) TYPE XL5   
	                 =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
  	               ENDIF
  	            
             ELSE
                =MESSAGEBOX('無此範圍資料 !',0+48,'提示訊息視窗')
                THISFORM.TXT1.SETFOCUS
                RETURN             
             ENDIF
         ENDIF    
         THISFORM.CMND2.CLICK
      ENDPROC
      procedure cmnd2.click        
           IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'C49'
           IF FOUND()
              DELETE 
           ENDIF        
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          

