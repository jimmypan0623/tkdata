***********************************************B12呆滯天數範圍選定
close all
clear 
IF !USED('A23')
   SELECT 0
   USE A23
ELSE 
   SELECT A23
ENDIF
SET ORDER TO 1      
SEEK LEFT(DTOS(DATE()),6)+'01'
 SKIP -1 
IF F07<>.T.
  =MESSAGEBOX('上月庫存尚未結轉至本月故無法使用此功能!',0+48,'提示訊息視窗')
  IF !USED('A05')
     SELECT 0
     USE A05
  ELSE
     SELECT A05
  ENDIF
  SET ORDER TO 1
  SEEK sys_oper+'B12'
  IF FOUND()
     DELETE 
  ENDIF
  CLOSE ALL         
  RETURN
ENDIF
FLG='0'
FCH=''
CHK=''
CHK_STR=''
R=0
FTR_STR='1'
QTD=0
AREA1='BAA'

IF ! USED('A01')
   SELE 0
   USE A01 INDE A01
ELSE
   SELE A01
ENDIF      
SET ORDER TO A01
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'B12'
IF !USED('C20')
   SELECT 0
   USE C20
ELSE
   SELECT C20
ENDIF
SET ORDER TO 1      
IF !USED('C01')
   SELECT 0
   USE C01
ELSE
   SELECT C01
ENDIF
SET ORDER TO C011      
IF !USED('D01')
   SELECT 0
   USE D01
ELSE
   SELECT D01
ENDIF
SET ORDER TO D011      
IF !USED('D10')   &&進貨記錄檔
   SELECT 0
   USE D10
ELSE
   SELECT D10
ENDIF
SET ORDER TO D104      
IF !USED('B01')
   SELE 0
   USE B01 INDE B01
ELSE
   SELE B01
ENDIF
SET ORDER TO 1   

IF !USED('A14')
   SELE 0
   USE A14 INDE A14
ELSE
   SELE A14
ENDIF
SET FILTER TO F04='*'
SET ORDER TO A141
IF !USED('B1A')
   SELECT 0
   USE B1A
ELSE 
   SELECT B1A
ENDIF
SET ORDER TO B1A1      
if !used('B11')
   SELE 0
   USE B11
ELSE
   SELE B11
ENDIF
SET FILTER TO F01<>INT_190
SET ORDER TO B111      
DELETE FROM B1A WHERE F01+F03+F08 NOT IN(SELE F01+F03+F08 FROM B11)   
SELECT F08 FROM B11 GROUP BY F08 INTO CURSOR C01_2 NOWAIT ORDER BY F08
CREATE CURSOR C01_1 (F01 C(5),F04 C(40),F05 C(8),F18 C(5),F23 C(5),F33 C(5))
INDEX ON F01 TAG C01_11
SET ORDER TO 1
SELECT C01_2
GO TOP
DO WHILE !EOF()
   SELECT C01_1
   APPEND BLANK
   REPLACE F01 WITH C01_2.F08
   REPLACE F04 WITH IIF(SEEK(F01,'C01'),C01.F04,'')
   REPLACE F05 WITH IIF(SEEK(F01,'C01'),C01.F05,'')   
   REPLACE F18 WITH IIF(SEEK(F01,'C01'),C01.F18,'')      
   REPLACE F23 WITH IIF(SEEK(F01,'C01'),C01.F23,'')     
   REPLACE F33 WITH IIF(SEEK(F01,'C01'),C01.F33,'')                
   SELECT C01_2
   SKIP
ENDDO

SELECT C01_2
USE
SELECT B11
SET FILTER TO F04<>0
SET RELATION TO F08 INTO C01_1 
SET RELATION TO F01+F03+F08 INTO B1A ADDI
B12form=createobject("tkB12")
B12form.show 


DEFINE CLASS TKB12 AS FORM
  AUTOCENTER=.T.
  CAPTION='請輸入條件範圍'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*390
  WIDTH=INT_015*300
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKB12'
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*10,;
      TOP=INT_015*35,;
      NAME='LABEL2',;
      CAPTION='輸入部門編號'  
  ADD OBJECT LABEL21 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*140,;
      TOP=INT_015*35,;
      NAME='LABEL21',;
      AUTOSIZE=.T.     
  ADD OBJECT LABELA AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*10,;
      TOP=INT_015*75,;
      NAME='LABELA',;
      CAPTION='輸入呆滯天數'  
  ADD OBJECT LABELB AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*80,;
      TOP=INT_015*115,;
      NAME='LABELB',;
      CAPTION='或是'            
  ADD OBJECT LABELC AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*10,;
      TOP=INT_015*155,;
      NAME='LABELC',;
      CAPTION='三個月存貨週轉率低於'                    
  ADD OBJECT LABEL1 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*132,;
      TOP=INT_015*75,;
      NAME='LABEL1',;
      CAPTION='天'
  ADD OBJECT LABEL3 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*150,;
      TOP=INT_015*75,;
      NAME='LABEL3',;
      CAPTION='以上(含)'      
  ADD OBJECT TEXT2 AS TEXT2 WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*30,;
      LEFT=INT_015*90,;
      HEIGHT=INT_015*25,; 
      WIDTH=INT_015*40,;
      MAXLENGTH=5,;
      NAME='TEXT2'      
  ADD OBJECT TEXT1 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*70,;
      LEFT=INT_015*90,;
      HEIGHT=INT_015*25,; 
      WIDTH=INT_015*40,;
      INPUTMASK='9999',;
      ALIGNMENT=3,;
      NAME='TEXT1'
  ADD OBJECT TEXT3 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*150,;
      LEFT=INT_015*140,;
      HEIGHT=INT_015*25,;        
      WIDTH=INT_015*80,;      
      INPUTMASK='9999.999999',;
      ALIGNMENT=3,;            
      NAME='TEXT3'                      
  ADD OBJECT CHK3 AS CHECKBOX WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*50,;
      TOP=INT_015*200,;
      FONTSIZE=INT_015*9,;
      NAME='CHK3',;
      CAPTION='包含零數'      
  ADD OBJECT CHK4 AS CHECKBOX WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*130,;
      TOP=INT_015*200,;
      FONTSIZE=INT_015*9,;
      NAME='CHK4',;
      CAPTION='包含不可出庫存'    
      
  ADD OBJECT CHK5 AS CHECKBOX WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*50,;
      TOP=INT_015*250,;
      FONTSIZE=INT_015*9,;
      NAME='CHK5',;
      CAPTION='或是只擷取'        
  ADD OBJECT TEXT4 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*248,;
      LEFT=INT_015*122,;
      HEIGHT=INT_015*25,;        
      WIDTH=INT_015*43,;    
      MAXLENGTH=4,;                  
      NAME='TEXT4'              
  ADD OBJECT LABEL4 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*167,;
      TOP=INT_015*250,;
      NAME='LABEL4',;
      CAPTION='年度至今殘餘呆料'                 
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      AUTOSIZE=.T.,;
      LEFT=INT_015*40,;    
      HEIGHT=INT_015*25,;   
      TOP=INT_015*314,;  
      WIDTH=INT_015*80,;
      CAPTION='\<Y.執行',;
      TOOLTIPTEXT='確認所輸入的範圍鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      AUTOSIZE=.T.,;
      LEFT=INT_015*120,;
      TOP=INT_015*314,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*80,;
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND2'
    PROCEDURE INIT 
          THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
           THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')      
           THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')
           THISFORM.TEXT2.READONLY=.F.
           THISFORM.TEXT2.VALUE='I000'
           THISFORM.LABEL21.CAPTION=IIF(SEEK('I000 ','A14'),A14.F02,'')
           THISFORM.TEXT1.VALUE=7
           THISFORM.TEXT3.VALUE=0.5   
           THISFORM.CHK3.VALUE=1        
           THISFORM.CHK4.VALUE=1
           THISFORM.CHK5.VALUE=0
           THISFORM.TEXT4.VALUE=LEFT(DTOS(DATE()),4)
           THISFORM.TEXT4.ENABLED=.F.
           THISFORM.TEXT4.READONLY=.T.
    ENDPROC
    PROCEDURE CMND1.CLICK
       IF !SEEK(THISFORM.TEXT2.VALUE,'A14')
          =MESSAGEBOX('無此部門編號!',0+48,'提示訊息視窗')
          THISFORM.TEXT2.SETFOCUS
          RETURN
       ENDIF
       IF THISFORM.CHK5.VALUE=1 AND VAL(ALLTRIM(THISFORM.TEXT4.VALUE))>0
          CHK_STR=THISFORM.TEXT4.VALUE
       ELSE
          CHK_STR=''   
       ENDIF
       IF FTR_STR='1'
         QTD=TTL(THISFORM.TEXT2.VALUE)
       ENDIF   

       CREATE CURSOR INVRT;
       (F01 C(5),F02 C(43),FIN N(16,4),FQY N(16,4),FND N(16,4),RTE N(11,6))
       INDEX ON F01+F02 TAG INVRT
       B25='B25'+LEFT(DTOS(DATE()),6)   &&本月庫存月報表檔
       IF !USED('&B25')
           SELE 0
           USE (B25) ALIA B25 INDE (B25) 
       ELSE
           SELE B25
       ENDIF
       SET ORDER TO 1     
       GO TOP
       DO WHILE !EOF()
          SELECT INVRT
          APPEND BLANK
          REPLACE FIN  WITH B25.F03
          REPLACE F01 WITH B25.F01
          REPLACE F02 WITH B25.F02
          REPLACE FQY WITH B25.F05+B25.F06+B25.F09+B25.F11+B25.F14+B25.F17+B25.F19+B25.F20
          REPLACE FND WITH B25.F15
          SELECT B25
          SKIP
       ENDDO
       SELECT B25
       USE
       SELECT A23
       SET ORDER TO 1
       SEEK LEFT(DTOS(DATE()),6)+'01'
       SKIP -1 
       IF !BOF()  
          B25='B25'+LEFT(DTOS(F01),6)
          IF !USED('&B25')
              SELE 0
              USE (B25) ALIA B25 INDE (B25) 
          ELSE
              SELE B25
          ENDIF
          SET ORDER TO 1      
          SELECT INVRT
          GO TOP
          DO WHILE !EOF()
             SELECT B25
             SEEK INVRT.F01+INVRT.F02
             IF FOUND()
                 SELECT INVRT
                 REPLACE FIN  WITH B25.F03
                 REPLACE FQY WITH FQY+(B25.F05+B25.F06+B25.F09+B25.F11+B25.F14+B25.F17+B25.F19+B25.F20)
                 REPLACE FND WITH FND+B25.F15        
             ELSE
                SELECT INVRT
                REPLACE FIN WITH 0             
             ENDIF     
             SELECT INVRT
             SKIP
          ENDDO
          SELECT B25
          USE
          SELECT A23
          SKIP -1
          IF BOF()
             USE
          ELSE
             B25='B25'+LEFT(DTOS(F01),6)
             IF !USED('&B25')
                SELE 0
                USE (B25) ALIA B25 INDE (B25) 
             ELSE
                SELE B25
             ENDIF              
             SET ORDER TO 1      
             SELECT INVRT
             GO TOP
             DO WHILE !EOF()
                 SELECT B25
                 SEEK INVRT.F01+INVRT.F02
                 IF FOUND()
                     SELECT INVRT
                     REPLACE FIN  WITH B25.F03
                     REPLACE FQY WITH FQY+(B25.F05+B25.F06+B25.F09+B25.F11+B25.F14+B25.F17+B25.F19+B25.F20)
                     REPLACE FND WITH FND+B25.F15    
                 ELSE
                     SELECT INVRT
                     REPLACE FIN WITH 0                                      
                 ENDIF     
                 SELECT INVRT
                 SKIP
             ENDDO
             SELECT B25
             USE      
             SELECT A23
             USE      
          ENDIF    
      ELSE
          SELECT A23
          USE
      ENDIF
      GHI=THISFORM.TEXT2.VALUE
      
      SELECT INVRT
      REPLACE RTE WITH FQY/IIF(FIN+FND<>0,(FIN+FND)/4,1)  ALL              
          SELE B11.F01,;
          B11.F03,;
          B11.F04,;
          B11.F05,;
          A14.F02,;
          B01.F02,;
          B01.F37,;
          INVRT.RTE;
          FROM B11,A14,B01,INVRT WHERE B11.F01=GHI AND B11.F04<>0 AND A14.F01=B11.F01 AND B01.F01=B11.F03 AND ((DATE()-B11.F05>=THISFORM.TEXT1.VALUE;
          ) OR INVRT.RTE<THISFORM.TEXT3.VALUE) AND INVRT.F01+INVRT.F02=B11.F01+B11.F03  ORDER BY B11.F03 NOWAIT  
          IF _TALLY>0
              FLG='1'

                     SELECT B11.F03 料號,B11.F04 庫存數量,B11.F05 最後異動日,B11.F06 用途,DATE()-B11.F05 呆滯天數,INVRT.RTE 存貨週轉率,B11.F08 客戶編號,IIF(SEEK(B11.F08,'C01_1'),C01_1.F05,SPACE(8)) 客戶簡稱,IIF(SEEK(IIF(SEEK(B11.F08,'C01_1','C01_11'),C01_1.F18,SPACE(5)),'A01'),A01.F03,SPACE(8)) 業務A,;
                     IIF(SEEK(IIF(SEEK(B11.F08,'C01_1','C01_11'),C01_1.F33,SPACE(5)),'A01'),A01.F03,SPACE(8)) 業務B,IIF(SEEK(IIF(SEEK(B11.F08,'C01_1','C01_11'),C01_1.F23,SPACE(5)),'A01'),A01.F03,SPACE(8)) 業務助理,;
                     IIF(SEEK(B11.F01+B11.F03+B11.F08,'B1A'),B1A.F10,SPACE(100)) 備註說明,IIF(SEEK(B11.F01+B11.F03+B11.F08,'B1A'),B1A.F17,SPACE(17)) 安全資料 FROM B11,INVRT ;
                      WHERE B11.F01=GHI AND ((DATE()-B11.F05>=THISFORM.TEXT1.VALUE ) OR INVRT.RTE<THISFORM.TEXT3.VALUE)  AND INVRT.F01+INVRT.F02=B11.F01+B11.F03 INTO CURSOR B11A ORDER BY B11.F01,B11.F03 NOWAIT                  

                    SELECT B11A
                    IF INT_028=.F.
                       COPY TO B11KB
                       SELECT B11A 
                       USE 
                       SELECT 0 
                       USE B11KB ALIAS B11A
                       IF !USED('P05')                          
                           SELECT 0
                           USE P05
                       ELSE 
                          SELECT P05    
                       ENDIF    
                       SET ORDER TO P051

                       IF !USED('C03')
                          SELECT 0
                          USE C03
                       ELSE
                          SELECT C03
                       ENDIF
                       SET ORDER TO C031      
                       SELECT B11A
                       GO TOP
                       DO WHILE !EOF()
                          SELECT P05
                          SEEK LEFT(B11A.用途,10)
                          IF FOUND()
                             SELECT B11A
                             REPLACE 客戶編號 WITH P05.F04
                          ELSE
                             SELECT C03
                             SEEK LEFT(B11A.用途,10)    
                             IF FOUND()
                                SELECT B11A
                                REPLACE 客戶編號 WITH C03.F03
                             ELSE
                                SEEK 'CA'+LEFT(B11A.用途,8)
                                IF FOUND()
                                   SELECT B11A
                                   REPLACE 客戶編號 WITH C03.F03
                                ENDIF   
                             ENDIF                            
                          ENDIF   
                          SELECT B11A
                          REPLACE 客戶簡稱 WITH IIF(SEEK(客戶編號,'C01'),C01.F04,'     ')
                          REPLACE 業務A WITH IIF(SEEK(IIF(SEEK(客戶編號,'C01'),C01.F18,'    '),'A01'),A01.F03,'     ')
                          REPLACE 業務B WITH IIF(SEEK(IIF(SEEK(客戶編號,'C01'),C01.F33,'     '),'A01'),A01.F03,'     ')
                          REPLACE 業務助理  WITH IIF(SEEK(IIF(SEEK(客戶編號,'C01'),C01.F23,'     '),'A01'),A01.F03,'     ')
                          SKIP
                       ENDDO
                       SELECT P05
                       USE
                       SELECT C03
                       USE
                    ENDIF   
                    SELECT B11A 
                     
                    
                    DO CASE 
                       CASE  THISFORM.CHK4.VALUE=1 AND THISFORM.CHK3.VALUE=1
                     
                            FTR_STR='1'
                            COPY TO BAA
                       CASE THISFORM.CHK4.VALUE<>1 AND THISFORM.CHK3.VALUE=1
                            COPY TO BAA FOR AT('不可出',備註說明)=0
                       CASE  THISFORM.CHK4.VALUE<>1 AND THISFORM.CHK3.VALUE<>1   
                             SELECT B11A.* FROM B11A INTO TABLE BAA WHERE AT('不可出',備註說明)=0 AND B11A.庫存數量>=VAL(IIF(SEEK(B11A.料號,'C20'),C20.F03,'')) ORDER BY 料號 NOWAIT
                             SELECT BAA
                             USE
                       CASE  THISFORM.CHK4.VALUE=1 AND THISFORM.CHK3.VALUE<>1      
                             FTR_STR='1'                      
                             SELECT B11A.* FROM B11A INTO TABLE BAA WHERE B11A.庫存數量>=VAL(IIF(SEEK(B11A.料號,'C20'),C20.F03,'')) ORDER BY 料號 NOWAIT
                             SELECT BAA 
                             USE
                    
                    ENDCASE
                    SELECT 0
                    USE BAA EXCLUSIVE
                    INDEX ON 料號+客戶編號 TAG BAA1
                    INDEX ON 客戶編號+料號 TAG BAA2
                    INDEX ON 業務B+料號 TAG BAA3
                    INDEX ON 業務助理+料號 TAG BAA4   
                    SELECT INVRT
                    USE                 
                    SELECT B11A 
                    USE
                   
                    SELECT B11
                    SET ORDER TO B112
                    SELE F01,F02,F03,F04 FROM D07 UNION ALL SELE F01,F02,F03,F04*(-1) FROM C05 WHERE F02= ANY(SELE 料號 FROM BAA) INTO CURSOR TKE081             
                    SELE F01,F02,F03,F04 FROM E06 UNION ALL SELE F01,F02,F03,F04*(-1) FROM E16 WHERE F02= ANY(SELE 料號 FROM BAA) INTO CURSOR TKE082
                    CREATE CURSOR TKE08 (F01 C(12),F02 C(43),F03 D(8),F04 N(13,4))
                    INDEX ON F02+DTOS(F03)+F01 TAG TKE08
                    SET ORDER TO 1
                    SELE TKE081
                    GO TOP
                    DO WHILE !EOF()
                       SELE TKE08
                       APPEND BLANK
                       REPL F01 WITH TKE081.F01
                       REPL F02 WITH TKE081.F02
                       REPL F03 WITH TKE081.F03
                       REPL F04 WITH TKE081.F04
                       SELE TKE081
                       SKIP
                    ENDDO
                    SELE TKE081
                    USE
                    SELE TKE082
                    GO TOP
                    DO WHILE !EOF()
                       SELE TKE08
                       APPEND BLANK
                       REPL F01 WITH TKE082.F01
                       REPL F02 WITH TKE082.F02
                       REPL F03 WITH TKE082.F03
                       REPL F04 WITH TKE082.F04
                       SELE TKE082
                       SKIP
                    ENDDO             
                    SELE TKE082
                    USE                      
                    SELECT BAA
                    SET ORDER TO 1
                    SET RELATION TO 料號 INTO D10 ADDI
                    GO TOP
                    IF !EMPTY(CHK_STR)
                       DO WHILE !EOF()
                          IF LEFT(DTOS(IIF(SEEK(BAA.料號,'D10'),D10.F03,CTOD(''))),4)<>CHK_STR
                             DELETE
                          ENDIF   
                          SELECT BAA
                          SKIP
                       ENDDO
                    ENDIF
                    GO TOP
                    ITT_SEEK=createobject("ITT_SEEK")  
                    ITT_SEEK.SHOW                                                                           
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
                 SEEK sys_oper+'B12'
                 IF FOUND()
                    DELETE 
                 ENDIF      
                 FLG='1'   
                 CLOSE TABLES ALL
                THISFORM.RELEASE            
                RETURN       
       ENDPROC                 
       PROCEDURE CHK4.INTERACTIVECHANGE
          IF THIS.VALUE=1
             FTR_STR='1'
          ELSE
             FTR_STR=''
          ENDIF          
       ENDPROC
       PROCEDURE CHK5.INTERACTIVECHANGE
          IF THIS.VALUE=1
             THISFORM.TEXT4.ENABLED=.T.
             THISFORM.TEXT4.READONLY=.F.
*!*	             THISFORM.CHK3.VALUE=1
*!*	             THISFORM.CHK4.VALUE=1
*!*	             THISFORM.CHK3.ENABLED=.F.
*!*	             THISFORM.CHK4.ENABLED=.F.
*!*	             FTR_STR='1'
          ELSE
             THISFORM.TEXT4.ENABLED=.F.
             THISFORM.TEXT4.READONLY=.T.
*!*	             THISFORM.CHK3.ENABLED=.T.
*!*	             THISFORM.CHK4.ENABLED=.T.
          ENDIF       
ENDDEFINE  
***************************************************特殊輸入欄位設定-----------查詢部門
DEFINE CLASS TEXT2 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9  
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) AND F04='*' 
       ELSE
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE F04='*' 
       ENDIF   
       IF _TALLY>0
          DEPART_SEEK=createobject("DEPART_SEEK")  
          DEPART_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'
       ENDIF   
    ENDIF   
       THISFORM.LABEL21.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')         
  ENDPROC
ENDDEFINE
*************************************************************
**********************************************************************************************
define class ITT_SEEK as form
*  autocenter=.t.  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*800
  FONTSIZE=INT_015*12
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='ITT_SEEK'
  ADD OBJECT LBL_REC AS LBL_REC WITH;
      AUTOCENTER =.T.,;
      CAPTION='存貨週轉率',;
      TOP=INT_015*3
  ADD OBJECT REC_COUNT AS REC_COUNT  WITH;
      TOP=INT_015*3
  ADD OBJECT LBL_SVE AS LABEL WITH;
      TOP=INT_015*20,;
      LEFT=INT_015*650,;
      FONTSIZE=INT_015*9,;
      AUTOSIZE=.T.,;
      CAPTION='安全資料'          
  ADD OBJECT LBL_SVE1 AS LABEL WITH;
      TOP=INT_015*20,;
      LEFT=INT_015*703,;
      FONTSIZE=INT_015*9,;
      AUTOSIZE=.T.
  ADD OBJECT TTL1 AS LABEL WITH;
      TOP=INT_015*40,;
      LEFT=INT_015*640,;
      FONTSIZE=INT_015*9,;
      CAPTION='不可出庫存',;
      AUTOSIZE=.T.      
  ADD OBJECT TTL2 AS LABEL WITH;
      TOP=INT_015*40,;
      LEFT=INT_015*703,;
      FONTSIZE=INT_015*9,;
      AUTOSIZE=.T.      
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      WIDTH=INT_015*168,;  
      AUTOCENTER=.T. 
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;            
      TOP=0,;
      LEFT=INT_015*435              
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=0,;
      ENABLED=.T.,;
      LEFT=INT_015*174    
  ADD OBJECT LBLF02 AS LABEL WITH;
      TOP=INT_015*38,;
      LEFT=INT_015*10,;
      HEIGHT=INT_015*23,;      
      WIDTH=INT_015*30,;
      FONTSIZE=INT_015*9,;
      CAPTION='備 註'     
  ADD OBJECT TEXT2 AS TEXTBOX WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*40,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*550,;
      MAXLENGTH=100,;
      FONTSIZE=INT_015*9      
  ADD OBJECT INV_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*453,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*70,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<V.庫存明細',;
      NAME='INV_BOTT'            
  ADD OBJECT IODT_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*387,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      CAPTION='\<Z.預期異動',;
      FONTSIZE=INT_015*9,;
      ENABLED=.T.,;
      VISIBLE=.T.                  
  ADD OBJECT REPLY_REMARK AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*523,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<E.修改'                 
  ADD OBJECT TRN_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*566,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='EXCEL'                 
  ADD OBJECT YES AS COMMANDBUTTON WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*660,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<Y.確定'
  ADD OBJECT NO AS COMMANDBUTTON WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*702,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<N.取消'       
  ADD OBJECT GRID1 AS GRID1 WITH;          
      AUTOCENTER=.T.,;
      TOP=INT_015*61,;  
      HEIGHT=INT_015*465,;     
      WIDTH=INT_015*789,;      
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*18,;          
      DELETEMARK=.F.,;
      READONLY=.T.,;
      RECORDSOURCETYPE=1,;
      COLUMNCOUNT=10,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7',;
      COLUMN8.NAME='COLUMN8',;
      COLUMN9.NAME='COLUMN9',;   
      COLUMN10.NAME='COLUMN10'               
      PROCEDURE INIT 
         THISFORM.CAPTION='B12.物料檢討明細    '+IIF(SEEK(GHI,'A14'),A14.F02,'')+'  '+CHK_STR+IIF(!EMPTY(CHK_STR),'年度至今殘餘','')

         THISFORM.TTL2.CAPTION=TRANSFORM(QTD,'@R 9,999,999,999.999')
         THISFORM.TTL1.VISIBLE=IIF(A02.F08='*',.T.,.F.) AND FTR_STR='1'
         THISFORM.TTL2.VISIBLE=IIF(A02.F08='*',.T.,.F.) AND FTR_STR='1'
         THISFORM.GRID1.READONLY=.T. 
         THISFORM.TEXT2.READONLY=.T.
         THISFORM.YES.VISIBLE=.F.
         THISFORM.NO.VISIBLE=.F.         
         THISFORM.GRID1.RECORDSOURCE='BAA'         
         THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.         
	      THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.            
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
         THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(EMPTY(BAA.備註說明),RGB(255,255,0),'')","COLUMN")
         THISFORM.GRID1.COLUMN5.DYNAMICFORECOLOR="IIF(AT('不可出',BAA.備註說明)>0,RGB(255,0,0),'')"
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='BAA.料號'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='庫存數量'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*80
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='BAA.庫存數量'         
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶編號'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*50
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='BAA.客戶編號'
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶簡稱'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='BAA.客戶簡稱'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65
         THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='備註說明'
         THISFORM.GRID1.COLUMN5.CONTROLSOURCE='BAA.備註說明'
         THISFORM.GRID1.COLUMN5.WIDTH=INT_015*300                            
         THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='業務人員'
         THISFORM.GRID1.COLUMN6.CONTROLSOURCE='BAA.業務b'
         THISFORM.GRID1.COLUMN6.WIDTH=INT_015*55
         THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='業務助理'
         THISFORM.GRID1.COLUMN7.CONTROLSOURCE="BAA.業務助理"       
         THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
         THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最後異動日'
         THISFORM.GRID1.COLUMN8.CONTROLSOURCE="BAA.最後異動日"
         THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60         
         THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='呆滯天數'
         THISFORM.GRID1.COLUMN9.CONTROLSOURCE="BAA.呆滯天數"
         THISFORM.GRID1.COLUMN9.WIDTH=INT_015*50         
         THISFORM.GRID1.COLUMN10.HEADER1.CAPTION='最後進貨日'          
         THISFORM.GRID1.COLUMN10.CONTROLSOURCE="D10.F03"
         THISFORM.GRID1.COLUMN10.WIDTH=INT_015*60         
         JK='料品編號'             
         THISFORM.GRID1.SETFOCUS        
      ENDPROC

  PROC KEY_LIST.INIT 
       WITH THIS 
           .ADDITEM('依料品編號+客戶編號排列')
           .ADDITEM('依客戶編號+料品編號排列')
           .ADDITEM('依業務人員+料品編號排列')                      
           .ADDITEM('依業務助理+料品編號排列')                                 
       ENDWITH      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE BAA
       DO CASE
          CASE THIS.DISPLAYVALUE='依料品編號+客戶編號排列'       
               SET ORDER TO 1 
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='BAA.料號'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='庫存數量'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='BAA.庫存數量'         
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='BAA.客戶編號'
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='BAA.客戶簡稱'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='備註說明'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='BAA.備註說明'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*300                            
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='業務人員'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='BAA.業務b'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='業務助理'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="BAA.業務助理"       
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最後異動日'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="BAA.最後異動日"
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60         
               THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='呆滯天數'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE="BAA.呆滯天數"
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80               
               THISFORM.GRID1.COLUMN6.DYNAMICFORECOLOR="RGB(0,0,0)"               
               THISFORM.GRID1.COLUMN5.DYNAMICFORECOLOR="IIF(AT('不可出',BAA.備註說明)>0,RGB(255,0,0),'')"               
          CASE THIS.DISPLAYVALUE='依客戶編號+料品編號排列'        
               SET ORDER TO 2          
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='BAA.客戶編號'
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='BAA.客戶簡稱'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*65               
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='BAA.料號'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='庫存數量'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='BAA.庫存數量'                        
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='備註說明'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='BAA.備註說明'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*300                            
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='業務人員'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='BAA.業務b'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='業務助理'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="BAA.業務助理"       
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最後異動日'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="BAA.最後異動日"
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60         
               THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='呆滯天數'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE="BAA.呆滯天數"
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*60               
               THISFORM.GRID1.COLUMN6.DYNAMICFORECOLOR="RGB(0,0,0)"               
               THISFORM.GRID1.COLUMN5.DYNAMICFORECOLOR="IIF(AT('不可出',BAA.備註說明)>0,RGB(255,0,0),'')"               
          CASE THIS.DISPLAYVALUE='依業務人員+料品編號排列'
               SET ORDER TO 3            
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='業務人員'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='BAA.業務b'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*55               
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='BAA.料號'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='庫存數量'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='BAA.庫存數量'                        
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='BAA.客戶編號'
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='BAA.客戶簡稱'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65                         
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='備註說明'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='BAA.備註說明'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*300                                 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='業務助理'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="BAA.業務助理"       
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最後異動日'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="BAA.最後異動日"
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60         
               THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='呆滯天數'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE="BAA.呆滯天數"
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*60               
               THISFORM.GRID1.COLUMN5.DYNAMICFORECOLOR="RGB(0,0,0)"               
               THISFORM.GRID1.COLUMN6.DYNAMICFORECOLOR="IIF(AT('不可出',BAA.備註說明)>0,RGB(255,0,0),'')"               
          CASE THIS.DISPLAYVALUE='依業務助理+料品編號排列'
               SET ORDER TO 4
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='業務助理'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE="BAA.業務助理"       
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*55               
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='BAA.料號'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='庫存數量'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='BAA.庫存數量'                        
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='BAA.客戶編號'
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='BAA.客戶簡稱'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65                         
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='備註說明'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='BAA.備註說明'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*300                                 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='業務人員'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='BAA.業務b'
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最後異動日'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="BAA.最後異動日"
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60         
               THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='呆滯天數'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE="BAA.呆滯天數"
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*60          
               THISFORM.GRID1.COLUMN5.DYNAMICFORECOLOR="RGB(0,0,0)"               
               THISFORM.GRID1.COLUMN6.DYNAMICFORECOLOR="IIF(AT('不可出',BAA.備註說明)>0,RGB(255,0,0),'')"               
       ENDCASE 
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC                  
      PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS XColIndex     
       THISFORM.REC_COUNT.CAPTION=TRANSFORM(BAA.存貨週轉率,'@R 999999.999999')
       THISFORM.TEXT2.VALUE=BAA.備註說明       
       THISFORM.LBL_SVE1.CAPTION=BAA.安全資料
         FCH=THISFORM.ACTIVECONTROL.NAME  
         CHK=''         
         THISFORM.REFRESH
      ENDPROC       
  PROC REPLY_REMARK.CLICK
       THISFORM.IODT_BOTT.ENABLED=.F.
       THISFORM.INV_BOTT.ENABLED=.F.
	   THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
	   THISFORM.ORPGROUP.ESC_BOTT.ENABLED=.F.
	   THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
	   THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
	   THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
	   THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
	   THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F. 
	   THISFORM.TRN_BOTT.ENABLED=.F.
	   THISFORM.KEY_LIST.ENABLED=.F.
	   THISFORM.GRID1.ENABLED=.F.
       THIS.ENABLED=.F.
       THISFORM.YES.VISIBLE=.T.
       THISFORM.NO.VISIBLE=.T.      
       THISFORM.TEXT2.READONLY=.F.
       THISFORM.TEXT2.SETFOCUS 
  ENDPROC    
  PROCEDURE YES.CLICK
     SELECT B1A
     SET ORDER TO 1
     SEEK GHI+BAA.料號+IIF(INT_028,BAA.客戶編號,SPACE(5))
     IF !FOUND()
         APPEND BLANK 
         REPLACE F01 WITH GHI
         REPLACE F03 WITH BAA.料號
         REPLACE F08 WITH IIF(INT_028,BAA.客戶編號,SPACE(5))
     ENDIF
     REPLACE F10 WITH THISFORM.TEXT2.VALUE    
     REPLACE F17 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')  
     IF FTR_STR='1'
        DO CASE      
           CASE AT('不可出',BAA.備註說明)=0 AND AT('不可出',THISFORM.TEXT2.VALUE)>0               
             THISFORM.TTL2.CAPTION=TRANSFORM((QTD+IIF(SEEK(BAA.料號,'B01'),B01.F37,0)*BAA.庫存數量),'@R 9,999,999,999.999')
           CASE AT('不可出',BAA.備註說明)>0 AND AT('不可出',THISFORM.TEXT2.VALUE)=0        
              
             THISFORM.TTL2.CAPTION=TRANSFORM((QTD-IIF(SEEK(BAA.料號,'B01'),B01.F37,0)*BAA.庫存數量),'@R 9,999,999,999.999')      
        ENDCASE
     ENDIF   
     SELE BAA   
     REPLACE 備註說明 WITH THISFORM.TEXT2.VALUE  
     REPLACE 安全資料 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'') 
 
     
     THISFORM.NO.CLICK
  ENDPROC
  PROCEDURE NO.CLICK
     THISFORM.TEXT2.READONLY=.T.  
     THISFORM.YES.VISIBLE=.F.
       THISFORM.IODT_BOTT.ENABLED=.T.
       THISFORM.INV_BOTT.ENABLED=.T.      
	   THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.T.
	   THISFORM.ORPGROUP.ESC_BOTT.ENABLED=.T.
	   THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
	   THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
	   THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
	   THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
	   THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T. 
	   THISFORM.KEY_LIST.ENABLED=.T.
	   THISFORM.TRN_BOTT.ENABLED=.T.
	   THISFORM.GRID1.ENABLED=.T. 
     THISFORM.REPLY_REMARK.ENABLED=.T.     
     THIS.Visible=.F.

     THISFORM.GRID1.SETFOCUS     
  ENDPROC      
  **********查看庫存明細
  PROC INV_BOTT.CLICK
         THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
         THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
         THISFORM.INV_BOTT.FORECOLOR=RGB(128,128,128)  
         THISFORM.IODT_BOTT.FORECOLOR=RGB(128,128,128)                         
         THISFORM.REPLY_REMARK.FORECOLOR=RGB(128,128,128)                                  
         SELE B11.F01,B11.F03,A14.F02,B11.F04,B11.F05,B11.F08 FROM B11,A14 WHERE B11.F01<>INT_190 AND B11.F03=BAA.料號 AND A14.F01=B11.F01  INTO CURSOR INV ORDER BY B11.F01
         INV_SEEK=CREATEOBJECT("INV_SEEK")  
         INV_SEEK.SHOW                  
         THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
         THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
         THISFORM.INV_BOTT.FORECOLOR=RGB(0,0,0)                  
         THISFORM.IODT_BOTT.FORECOLOR=RGB(0,0,0) 
         THISFORM.REPLY_REMARK.FORECOLOR=RGB(0,0,0)                                                                    
         AREA1='BAA'         
         SELECT BAA
         THISFORM.GRID1.SETFOCUS
  ENDPROC                 
*********************************查看進出明細
  PROC IODT_BOTT.CLICK

            THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
            THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
            THISFORM.INV_BOTT.FORECOLOR=RGB(128,128,128)  
            THISFORM.IODT_BOTT.FORECOLOR=RGB(128,128,128) 
            THISFORM.REPLY_REMARK.FORECOLOR=RGB(128,128,128)                                              
            LK=BAA.料號
            LV=DATE()
            SELE D07.F01,D07.F03,D07.F04,D07.F05,D07.F07,D07.F09 FROM D07 WHERE D07.F02=LK   INTO CURSOR IODT1 ORDER BY D07.F03 NOWAIT&&
            SELE E06.F01,E06.F03,E06.F04,E06.F05,E06.F07,E06.F09 FROM E06 WHERE E06.F02=LK   INTO CURSOR IODT3  ORDER BY E06.F03 NOWAIT&&            
            SELE C05.F01,C05.F03,C05.F04,C05.F06 FROM C05 WHERE C05.F02=LK  INTO CURSOR IODT2  ORDER BY C05.F03 NOWAIT&&
            SELE E16.F01,E16.F03,E16.F04,E16.F06 FROM E16 WHERE E16.F02=LK  INTO CURSOR IODT4 ORDER BY E16.F03 NOWAIT&&         
            CREATE CURSOR IODT(F01 C(12),F03 D(8),F04 N(13,3),F08 N(13,3),F05 C(50))
            INDE ON DTOS(F03)+F01 TAG IODT
            SET ORDER TO 1
            SELE IODT1
            GO TOP
            DO WHILE !EOF()
               SELE IODT
               APPEND BLANK
               REPL F01 WITH IODT1.F01
               REPL F03 WITH IODT1.F03
               REPL F04 WITH IODT1.F04
               REPL F05 WITH '採:'+IODT1.F05+PADL(IIF(SEEK(IODT1.F05,'D01'),D01.F04,''),8,' ')+IIF(IODT1.F09='1','  ','><')+IIF(!EMPTY(IODT1.F07),'請:','')+IODT1.F07
               SELE IODT1
               SKIP
            ENDDO   
            USE
            SELE IODT3
            GO TOP
            DO WHILE !EOF()
               SELE IODT
               APPEND BLANK
               REPL F01 WITH IODT3.F01
               REPL F03 WITH IODT3.F03
               REPL F04 WITH IODT3.F04
               REPL F05 WITH '產:'+IODT3.F05+PADL(IIF(SEEK(IODT3.F05,'A14'),A14.F02,''),8,' ')+IIF(IODT3.F09='1','  ','><')+'出:'+IODT3.F07+IIF(SEEK(IODT3.F07,'C01'),C01.F05,'')
               SELE IODT3
               SKIP
            ENDDO   
            USE
            SELE IODT2
            GO TOP
            DO WHILE !EOF()
               SELE IODT
               APPEND BLANK
               REPL F01 WITH IODT2.F01
               REPL F03 WITH IODT2.F03
               REPL F04 WITH IODT2.F04*(-1)
               REPL F05 WITH '出:'+IODT2.F06+IIF(SEEK(IODT2.F06,'C01'),C01.F05,'')
               SELE IODT2
               SKIP
            ENDDO   
            USE
            SELE IODT4
            GO TOP
            DO WHILE !EOF()
               SELE IODT
               APPEND BLANK
               REPL F01 WITH IODT4.F01
               REPL F03 WITH IODT4.F03
               REPL F04 WITH IODT4.F04*(-1)
               REPL F05 WITH '領:'+IODT4.F06+IIF(SEEK(IODT4.F06,'A14'),A14.F02,'')           
               SELE IODT4
               SKIP
            ENDDO       
            USE                               
                ORD_NO=''
                STD_NO=BAA.料號
                DTE_NO=DATE()
               IODT_SEEK=CREATEOBJECT("IODT_SEEK")  
               IODT_SEEK.SHOW           
         THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
         THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
         THISFORM.INV_BOTT.FORECOLOR=RGB(0,0,0)                  
         THISFORM.IODT_BOTT.FORECOLOR=RGB(0,0,0)         
         THISFORM.REPLY_REMARK.FORECOLOR=RGB(0,0,0)                                                   
         AREA1='BAA'
         SELECT BAA
        THISFORM.GRID1.SETFOCUS
    ENDPROC 
**********  轉EXCLE檔
    PROC TRN_BOTT.CLICK
         IF A02.F09='*'
            SELECT BAA.*,IIF(SEEK(BAA.料號,'D10'),D10.F03,CTOD('')) 最後進貨日,B01.F37 平均成本 FROM BAA,B01 ORDER BY BAA.料號 WHERE B01.F01=BAA.料號 NOWAIT
         ELSE
            SELECT BAA.*,IIF(SEEK(BAA.料號,'D10'),D10.F03,CTOD('')) 最後進貨日 FROM BAA ORDER BY BAA.料號 NOWAIT
         ENDIF   
         IF _TALLY>0          
               FILE_NAME='B12'+ALLTRIM(IIF(SEEK(GHI,'A14'),A14.F02,''))+DTOS(DATE())              
            GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	        IF EMPTY(GCDELIMFILE)  && ESC PRESSED
  	           RETURN
	        ELSE
	           ON ERROR DO FILE_IMPACT
	           COPY TO (GCDELIMFILE) TYPE XL5   
	           =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
  	        ENDIF  	  
         ENDIF
    ENDPROC
enddefine     
*!*	**********************************************  計算不可出庫存金額
FUNC TTL
     PARA BKK_NO
     qt_od=0
     sele B11
     SET ORDER TO 1
     SEEK BKK_NO
     IF FOUND()
        DO WHILE B11.F01=BKK_NO      
           IF AT('不可出',B1A.F10)>0
              IF EMPTY(CHK_STR)
                 qt_od=qt_od+B11.F04*IIF(SEEK(B11.F03,'B01'),B01.F37,0)
              ELSE
                 IF LEFT(DTOS(IIF(SEEK(B11.F03,'D10'),D10.F03,CTOD(''))),4)=CHK_STR
                    qt_od=qt_od+B11.F04*IIF(SEEK(B11.F03,'B01'),B01.F37,0)
                 ENDIF   
              ENDIF
           ENDIF              
           SELECT B11
           SKIP
         ENDDO
     ELSE
       qt_od=0
     ENDIF
    
     RETURN qt_od
