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
SEEK sys_oper+'P49'
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
IF !USED('D00')
   SELE 0
   USE D00
ELSE 
  SELE D00
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
IF !USED('P05')
   SELE 0
   USE P05
ELSE 
  SELE P05
ENDIF
SET ORDER TO 1
IF !USED('P04')
   SELE 0
   USE P04
ELSE 
  SELE P04
ENDIF
SET ORDER TO 1
P49form=createobject("tkP49")
P49form.show       
define class tkP49 as form
  autocenter=.t.
  caption='P49.外購下單日期搜尋範圍'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*320
  width=INT_015*420
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkP49'
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
  ADD OBJECT OUTPUT AS OPTIONGROUP WITH;
      LEFT=INT_015*150,;
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
         THISFORM.OUTPUT.VALUE=1         
         THISFORM.TXT1.VALUE=DATE()
         THISFORM.TXT2.VALUE=DATE()
      ENDPROC
**********************************開始搜尋指定範圍之P04內容         IIF(SEEK(IIF(SEEK(IIF(SEEK(LEFT(P04.F11,10),'P05'),P05.F04,''),'C01'),C01.F33,''),'A01'),A01.F03,'') 業務人員
      PROCEDURE CMND1.CLICK      
         SELECT P03.F03 廠商編號,D01.F04 廠商簡稱,P03.F01 下單日期,P03.F14 希望交期,P04.F02 料品編號,P04.F01 採購單號,P04.F03-P04.F16 採購數量,P04.F04 幣別,P04.F05 採購單價,P04.F06 小計,D00.F02 匯率,ROUND(P04.F06*D00.F02,INT_068) 本幣別小計,P04.F22 已回覆數量,RIGHT(P04.F11,9) 需求對象,IIF(SEEK(IIF(SEEK(IIF(SEEK(LEFT(P04.F11,10),'P05'),P05.F04,''),'C01'),C01.F33,''),'A01'),A01.F03,'') 業務擔當;         
         FROM P04,P03,D00,D01 WHERE P03.F02=P04.F01 AND P04.F03-P04.F16>0 AND D01.F01=P03.F03 AND D00.F01=P04.F04 AND P03.F01>=THISFORM.TXT1.VALUE AND P03.F01<=THISFORM.TXT2.VALUE INTO CURSOR BTX ORDER BY P03.F03,P03.F01,P03.F02,P04.F02 NOWAIT   
             SELECT BTX
             IF _TALLY>0
                GHK1=DTCX(THISFORM.TXT1.VALUE)
                GHK2=DTCX(THISFORM.TXT2.VALUE)      
                FILE_NAME='外購下單日報表'            
                IF THISFORM.OUTPUT.VALUE=1
                   IF !USED('D10')
                      SELECT 0
                      USE D10
                   ELSE
                      SELECT D10
                   ENDIF
                   SET ORDER TO 1     
                   CREATE CURSOR BTXMN (廠商編號 C(5),廠商簡稱 C(8),下單日期 D(8),希望交期 D(8),料品編號 C(43),採購單號 C(10),採購數量 N(12),幣別 C(4),採購單價 N(12,6),小計 N(12,2),匯率 N(9,4),本幣別小計 N(12,2),已回覆數量 N(14),需求對象 C(8),業務擔當 C(8),最後進貨日 D(8),已進數量 N(12))
                   SELECT  BTX 
                   GO TOP
                   DO WHILE !EOF()
                      SELECT BTXMN
                      APPEND BLANK
                      REPLACE 廠商編號 WITH BTX.廠商編號
                      REPLACE 廠商簡稱 WITH BTX.廠商簡稱
                      REPLACE 下單日期 WITH BTX.下單日期
                      REPLACE 希望交期 WITH BTX.希望交期
                      REPLACE 料品編號 WITH BTX.料品編號
                      REPLACE 採購單號 WITH BTX.採購單號
                      REPLACE 採購數量 WITH BTX.採購數量
                      REPLACE 幣別 WITH BTX.幣別
                      REPLACE 採購單價 WITH BTX.採購單價
                      REPLACE 小計 WITH BTX.小計
                      REPLACE 匯率 WITH BTX.匯率
                      REPLACE 本幣別小計 WITH BTX.本幣別小計
                      REPLACE 已回覆數量 WITH BTX.已回覆數量
                      REPLACE 需求對象 WITH BTX.需求對象
                      REPLACE 業務擔當 WITH BTX.業務擔當
                      SELECT D10
                      SEEK BTX.採購單號+BTX.料品編號
                      IF FOUND()
                         DO WHILE F01+F02=BTX.採購單號+BTX.料品編號
                            SELECT BTXMN
                            REPLACE 最後進貨日 WITH D10.F03
                            REPLACE 已進數量 WITH 已進數量+D10.F05
                            SELECT D10
                            SKIP
                         ENDDO   
                      ENDIF
                      SELECT BTX
                      SKIP
                   ENDDO
                   SELECT D10
                   USE
                   SELECT BTXMN
                   
	                  GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	                  IF EMPTY(GCDELIMFILE)  && ESC PRESSED
  	                     RETURN
	                  ELSE
	                    ON ERROR DO FILE_IMPACT
	                    COPY TO (GCDELIMFILE) TYPE XL5   
	                    =MESSAGEBOX('儲存成功',0+48,'提示訊息視窗')
  	                  ENDIF
  	                  SELECT BTXMN
  	                  USE
  	            ELSE
                   REPORT FORM ALLTRIM(INT_116)+'P49' TO PRINT PROMPT PREVIEW  	           
  	            ENDIF             
             ELSE
                =MESSAGEBOX('無此範圍資料 !',0+48,'提示訊息視窗')
                THISFORM.TXT1.SETFOCUS
                RETURN
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
           SEEK sys_oper+'P49'
           IF FOUND()
              DELETE 
           ENDIF   
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          

