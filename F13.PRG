close all
clear 
AREA1='EDRQ'
FLG='0'
FCH=''
CHK=''
UKEF=''
CHK_STR=''
CKS=SPACE(10)
OUS=SPACE(5)
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'F13'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('C01')
     SELECT 0
     USE C01
ELSE
     SELECT C01
ENDIF
SET ORDER TO 1          
IF !USED('A14')
     SELECT 0
     USE A14
ELSE
     SELECT A14
ENDIF
SET ORDER TO 1          
SET FILTER TO  F05='*'

IF !USED('B01')
   SELE 0
   USE B01
ELSE 
  SELE B01
ENDIF
SET ORDER TO 1
IF !USED('F01')
   SELE 0
   USE F01
ELSE 
  SELE F01
ENDIF
SET ORDER TO F013
SET FILTER TO  F05=.T.
CREATE CURSOR EDRQ;
(F01 C(43),F02 C(3),F03 N(9),F04 D(8))
INDEX ON F01+F02+DTOS(F04) TAG EDRQ1
SET ORDER TO 1      
SET RELATION TO F01 INTO B01
IF messagebox('是否直接轉入客戶訂單內容',4+32+256,'請確認') = 6	
                     IF !USED('C03')
                          SELECT 0
                          USE C03
                      ELSE
                          SELECT C03
                      ENDIF
                      SET ORDER TO 1        
                     IF !USED('C04')
                         SELECT 0
                         USE C04
                     ELSE
                         SELECT C04
                     ENDIF
                     SET ORDER TO C041            

     F131form=createobject("tkF131")
      F131form.show                        

 ENDIF
 SELECT EDRQ
 GO TOP
F13form=createobject("tkF13")
F13form.show       
*********************************************************************
define class TKF13 as form  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  autocenter=.t.
  caption='F13模擬毛需求'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*420
  WIDTH=INT_015*470
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='TKF13'
  ADD OBJECT SHAPE1 AS SHAPE1 WITH;
      LEFT=INT_015*3,;
      WIDTH=INT_015*464,;
      HEIGHT=INT_015*90
  ADD OBJECT SHAPE2 AS SHAPE2 WITH;    
      LEFT=INT_015*3,;
      TOP=INT_015*92,;
      WIDTH=INT_015*464,;
      HEIGHT=INT_015*290
  ADD OBJECT SHAPE3 AS SHAPE3 WITH;
      LEFT=INT_015*3,;
      TOP=INT_015*384,;
      WIDTH=INT_015*464,;
      HEIGHT=INT_015*35    
  add object cmdgroup as cmdgroup with;      
      TOP=INT_015*349,;
      LEFT=INT_015*60
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      TOP=INT_015*385,;
      LEFT=INT_015*60
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      VISIBLE=.F.    
  ADD OBJECT LBLF01 AS LABEL WITH;
      TOP=INT_015*8,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='製品編號'
  ADD OBJECT LBLF02 AS LABEL WITH;
      TOP=INT_015*8,;
      LEFT=INT_015*326,;
      WIDTH=INT_015*50,;
      CAPTION='版        次'      
  ADD OBJECT LBLF11 AS LABEL WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='品名規格'      
  ADD OBJECT LBLF011 AS LABEL WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*56,;
      AUTOSIZE=.T.                    
  ADD OBJECT LBLF03 AS LABEL WITH;
      TOP=INT_015*58,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='需求數量'
  ADD OBJECT LBLF04 AS LABEL WITH;
      TOP=INT_015*58,;
      LEFT=INT_015*155,;
      WIDTH=INT_015*50,;
      CAPTION='需求日期'        
  ADD OBJECT TXTF01 AS TXTF01 WITH;
      TOP=INT_015*4,;
      LEFT=INT_015*56,;
      WIDTH=INT_015*73,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*261,;
      MAXLENGTH=43,;
      NAME='TXTF01'
  ADD OBJECT TXTF02 AS TEXTBOX WITH;
      TOP=INT_015*4,;
      LEFT=INT_015*376,;
      WIDTH=INT_015*73,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*30,;
      MAXLENGTH=3,;
      NAME='TXTF02'      
  ADD OBJECT TXTF03 AS TEXTBOX WITH;
      TOP=INT_015*54,;
      LEFT=INT_015*56,;
      WIDTH=INT_015*73,;
      HEIGHT=INT_015*25,;
      INPUTMASK='99999999999',;
      ALIGNMENT=1,;
      NAME='TXTF03'      
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      TOP=INT_015*54,;
      LEFT=INT_015*206,;
      WIDTH=INT_015*70,;
      HEIGHT=INT_015*25,;
      NAME='TXTF04'            
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*50,;
      LEFT=INT_015*300                
  add object GRID1 as GRID1 with;        
      AUTOCENTER=.T.,;
      LEFT=INT_015*10,;
      TOP=INT_015*96,;
      WIDTH=INT_015*450,;      
      HEIGHT=INT_015*255,;      
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*19,;      
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=4,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3' ,;
      COLUMN4.NAME='COLUMN4'     
  add object cmnd2 as commandbutton with;
      LEFT=INT_015*191,;
      TOP=INT_015*389,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<T.展料',;
      TOOLTIPTEXT='以此表單展開材料需求!快速鍵->ALT+T'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX') 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')          
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='EDRQ'
         THISFORM.GRID1.CHILDORDER='EDRQ1'         
         THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
         THISFORM.SHRGROUP.VISIBLE=.F.
         THISFORM.SHRGROUP.ENABLED=.F.         
         thisform.setall('readonly',.t.,'textbox')
         THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')         
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='製品編號'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*261
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='EDRQ.F01'
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='版次'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*30
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='EDRQ.F02'         
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='需求數量'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='EDRQ.F03'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*75
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='需求日期'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='EDRQ.F04'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*70   
         THISFORM.GRID1.SETFOCUS       
         FCH='GRID1'
         IF THISFORM.GRID1.ACTIVEROW=0
            THISFORM.CMDGROUP.ENABLED=.F.
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
            THISFORM.CMND2.ENABLED=.F.
            THISFORM.LBLF011.CAPTION=''
         ELSE
            THISFORM.CMDGROUP.ENABLED=.T.         
             THISFORM.CMND2.ENABLED=.T.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.
         ENDIF                     
         JK='製品編號'             
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
         LPARAMETERS XColIndex     
         THISFORM.TXTF01.VALUE=EDRQ.F01
         THISFORM.TXTF02.VALUE=EDRQ.F02
         THISFORM.TXTF03.VALUE=EDRQ.F03
         THISFORM.TXTF04.VALUE=EDRQ.F04
         THISFORM.LBLF011.CAPTION=B01.F02         
         FCH=THISFORM.ACTIVECONTROL.NAME 
         CHK=''
         IF THIS.ACTIVEROW=0
            THISFORM.CMDGROUP.ENABLED=.F.
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
            THISFORM.CMND2.ENABLED=.F.
         ELSE
            THISFORM.CMDGROUP.ENABLED=.T.         
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.T.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.         
            THISFORM.CMND2.ENABLED=.T.            
         ENDIF            
      ENDPROC 
      PROC ORPGROUP.NEW_BOTT.CLICK
           THISFORM.CMND2.ENABLED=.F.
           THISFORM.SetAll('READONLY',.F.,'TEXTBOX')      
           THISFORM.GRID1.ENABLED=.F.               
           THISFORM.TXTF01.READONLY=.F.
           THISFORM.TXTF01.VALUE=''
           THISFORM.TXTF02.VALUE=''
           THISFORM.TXTF03.VALUE=0
           THISFORM.TXTF04.VALUE=DATE()
           THISFORM.TXTF01.SETFOCUS
      ENDPROC
      PROC ORPGROUP.EDIT_BOTT.CLICK
           THISFORM.CMND2.ENABLED=.F.      
           THISFORM.SetAll('READONLY',.F.,'TEXTBOX')                                
           THISFORM.GRID1.ENABLED=.F.               
           THISFORM.TXTF01.READONLY=.F.
           THISFORM.TXTF04.SETFOCUS
      ENDPROC        
      PROC ORPGROUP.DEL_BOTT.CLICK
           IF messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6	            
               SELECT EDRQ
               DELETE 
               SKIP -1
               IF BOF()
                   GO TOP
               ENDIF
              THISFORM.GRID1.SETFOCUS
              IF THISFORM.GRID1.ACTIVEROW=0
                 THISFORM.CMDGROUP.ENABLED=.F.
                 THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
                 THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.   
                 THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.T.            
                 THISFORM.SETALL('VALUE','','TEXTBOX')
                 THISFORM.LBLF011.CAPTION=''
              ELSE      
                 THISFORM.CMDGROUP.ENABLED=.T.

                 THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.
              ENDIF           
           ENDIF   
      ENDPROC
      PROC SHRGROUP.SHURE1_BOTT.CLICK
                 IF EMPTY(THISFORM.TXTF01.VALUE) .OR.  !SEEK(THISFORM.TXTF01.VALUE+THISFORM.TXTF02.VALUE,'F01')
                      =MESSAGEBOX('請輸入正確料號!',0+48,'提示訊息視窗')
                      THISFORM.TXTF01.SETFOCUS
                      RETURN
                 ENDIF
                 IF SEEK(THISFORM.TXTF01.VALUE+THISFORM.TXTF02.VALUE+DTOS(THISFORM.TXTF04.VALUE),'EDRQ')
                     =MESSAGEBOX('已有同樣的料號同樣的需求日期建立在表單中!',0+48,'提示訊息視窗')
                     THISFORM.TXTF04.SETFOCUS
                     RETURN
                 ENDIF
                 IF THISFORM.TXTF03.VALUE=0
                      =MESSAGEBOX('需求數量不得為 0',0+48,'提示訊息視窗')
                      THISFORM.TXTF03.SETFOCUS
                      RETURN
                 ENDIF
                 IF THISFORM.TXTF04.VALUE<DATE()
                      =MESSAGEBOX('此需求日已過期模擬需求無意義!',0+48,'提示訊息視窗')
                      THISFORM.TXTF04.SETFOCUS
                      RETURN
                 ENDIF
                 SELECT EDRQ
                 APPEND BLANK
                 REPLACE F01 WITH THISFORM.TXTF01.VALUE
                 REPLACE F02 WITH THISFORM.TXTF02.VALUE                 
                 REPLACE F03 WITH THISFORM.TXTF03.VALUE
                 REPLACE F04 WITH THISFORM.TXTF04.VALUE                 
                 SEEK THISFORM.TXTF01.VALUE+THISFORM.TXTF02.VALUE+DTOS(THISFORM.TXTF04.VALUE)
                 THISFORM.REFRESH
                THISFORM.ORPGROUP.NEW_BOTT.CLICK    
      ENDPROC         
      PROC SHRGROUP.SHURE2_BOTT.CLICK
                 IF EMPTY(THISFORM.TXTF01.VALUE) .OR.  !SEEK(THISFORM.TXTF01.VALUE+THISFORM.TXTF02.VALUE,'F01')
                      =MESSAGEBOX('請輸入正確料號!',0+48,'提示訊息視窗')
                      THISFORM.TXTF01.SETFOCUS
                      RETURN
                 ENDIF
                 IF THISFORM.TXTF03.VALUE=0
                      =MESSAGEBOX('需求數量不得為 0',0+48,'提示訊息視窗')
                      THISFORM.TXTF03.SETFOCUS
                      RETURN
                 ENDIF
                 IF THISFORM.TXTF04.VALUE<DATE()
                      =MESSAGEBOX('此需求日已過期模擬需求無意義!',0+48,'提示訊息視窗')
                      THISFORM.TXTF04.SETFOCUS
                      RETURN
                 ENDIF      
                 SELECT EDRQ
                 REPLACE F01 WITH THISFORM.TXTF01.VALUE
                 REPLACE F02 WITH THISFORM.TXTF02.VALUE                 
                 REPLACE F03 WITH THISFORM.TXTF03.VALUE
                 REPLACE F04 WITH THISFORM.TXTF04.VALUE
                THISFORM.SHRGROUP.ABANDON_BOTT.CLICK    
      ENDPROC
      PROC SHRGROUP.ABANDON_BOTT.CLICK
           THISFORM.SHRGROUP.ENABLED=.F.
           THISFORM.SHRGROUP.VISIBLE=.F.      
           THISFORM.ORPGROUP.ENABLED=.T.
           THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
           THISFORM.CMDGROUP.ENABLED=.T.
           THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
           THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
           THISFORM.SETALL('READONLY',.T.,'TEXTBOX') 
           THISFORM.TXTF01.READONLY=.T.
           THISFORM.GRID1.ENABLED=.T.
           THISFORM.GRID1.SETFOCUS     
           THISFORM.CMND2.ENABLED=.T.
           UNLOCK ALL      
      ENDPROC
      PROC SHRGROUP.FINISH_BOTT.CLICK
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
      ENDPROC
      procedure cmnd2.click
           SELECT f01 FROM EDRQ WHERE EMPTY(F02) NOWAIT
           IF _TALLY>0
               =MESSAGEBOX(F01+'尚未建立BOM','提示訊息視窗')
               RETURN
           ENDIF    
          IF messagebox('是否將耗損率%一併列入計算?',4+32+256,'請確認') = 6	            
               CHK_STR='Y'
            ELSE
               CHK_STR='N'   
           ENDIF             
           IF !USED('E00')
              SELECT 0
              USE E00
           ELSE
              SELECT E00
           ENDIF
           SET ORDER TO E001      
             IF !USED('B11')
                SELE 0
                USE B11
             ELSE
               SELE B11
             ENDIF      
             SET ORDER TO B112 
             SELECT F01,F03,SUM(F04) FROM B11  GROUP BY F01,F03 ORDER BY F01,F03 INTO CURSOR B11_2 NOWAIT
             CREATE CURSOR B11_1 (F01 C(5),F03 C(43),F04 N(14,4))
             INDEX ON F03+F01 TAG B11_12
             SET ORDER TO B11_12
             SELECT B11_2
             GO TOP
             DO WHILE !EOF()
                SELECT B11_1
                APPEND BLANK
                REPLACE F01 WITH B11_2.F01
                REPLACE F03 WITH B11_2.F03
                REPLACE F04 WITH B11_2.SUM_F04
                SELECT B11_2
                SKIP
             ENDDO
             SELECT B11_2
             USE
             IF !USED('A09')
                SELE 0
                USE A09
             ELSE
                SELE A09
            ENDIF
            SET ORDER TO 1             
             A24='A24'+LEFT(DTOS(DATE()),6)   &&刪除單據暫存檔
             IF !USED('&A24')
                SELE 0
                USE (A24) ALIA A24
             ELSE
                SELE A24
             ENDIF
             SET ORDER TO 1                   
             B39='B39'+LEFT(DTOS(DATE()),6)   &&未過帳單據查詢檔
             IF !USED('&B39')
                SELE 0
                USE (B39) ALIA B39
             ELSE
                SELE B39
             ENDIF
             SET ORDER TO 1     
             B06='B06'+LEFT(DTOS(DATE()),6)   &&未過帳單據查詢檔
             IF !USED('&B06')
                SELE 0
                USE (B06) ALIA B06
             ELSE
                SELE B06
             ENDIF
             SET ORDER TO 1                               
             SELECT F07.F04 STCK_NO,EDRQ.F04-F07.F06 PRS_DTE,ROUND(F07.F07*EDRQ.F03*IIF(CHK_STR='Y',1+F07.F09/100,1),IIF(E00.F05='3',0,4)) SHD_QTY;
             ,F07.F10 TPE FROM F07,EDRQ,E00 WHERE F07.F01+F07.F02=EDRQ.F01+EDRQ.F02 AND EMPTY(F07.F05)  AND F07.F10<>'2' AND E00.F01=F07.F03 INTO CURSOR ITT_2 NOWAIT
             SELECT STCK_NO,PRS_DTE,SUM(SHD_QTY),TPE FROM ITT_2 GROUP BY STCK_NO,TPE,PRS_DTE  INTO CURSOR ITT_1 NOWAIT      
         IF _TALLY>0                         
             CREATE CURSOR ITT;
             (F01 C(12),F02 C(43),F03 D(8),F04 N(14,3),PRS_DTE D(8),TPE C(1))
             INDE ON F02+DTOS(F03) TAG ITT1
             INDEX ON DTOS(F03)+F02 TAG ITT2
             SET RELA TO F02 INTO B01               
             SET ORDER TO 1             
             SELE ITT_1
             GO TOP
             DO WHILE !EOF()
                SELE ITT
                APPEND BLANK
                REPL F02 WITH ITT_1.STCK_NO  
                REPL F03 WITH ITT_1.PRS_DTE+IIF(SEEK(F02,'B01'),B01.F28,0)    &&需求日
                REPL F04 WITH ITT_1.SUM_SHD_QTY
                REPLACE PRS_DTE WITH ITT_1.PRS_DTE &&下單日
                REPLACE TPE WITH ITT_1.TPE
                SELE ITT_1
                SKIP   
             ENDDO

             SELECT ITT_2
             USE
             SELE ITT_1
             USE   
             SELE F01,F02,F03,F04 FROM D07 UNION SELE F01,F02,F03,F04*(-1) FROM C05 WHERE F02= ANY(SELE F02 FROM ITT) INTO CURSOR TKE081             
             SELE F01,F02,F03,F04 FROM E06 UNION SELE F01,F02,F03,F04*(-1) FROM E16 WHERE F02= ANY(SELE F02 FROM ITT)  AND LEFT(F01,8)<>SUBSTR(CKS,3,8) INTO CURSOR TKE082
             CREATE CURSOR TKE08 (F01 C(12),F02 C(43),F03 D(8),F04 N(13,4))
             INDEX ON F02+DTOS(F03)+F01 TAG TKE081
             SET ORDER TO 1
             SELE TKE081   &&
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
             SELECT ITT
             GO TOP
             DO WHILE !EOF()
                    SELECT TKE08
                    APPEND BLANK
                    REPLACE F01 WITH ''
                    REPLACE F02 WITH ITT.F02
                    REPLACE F03 WITH ITT.F03
                    REPLACE F04 WITH ITT.F04*(-1)
                    SELECT ITT
                    SKIP
             ENDDO
             AREA1='ITT'
             SELE ITT
             GO TOP
             ITT_SEEK=createobject("ITT_SEEK")  
             ITT_SEEK.SHOW   
         ELSE
            IF !USED('A05')
               SELECT 0
               USE A05
            ELSE
               SELECT A05
            ENDIF
            SET ORDER TO 1
            SEEK sys_oper+'F13'
            IF FOUND()
               DELETE 
            ENDIF         
            CLOSE ALL
            =MESSAGEBOX('無此資料 !',0+48,'提示訊息視窗')
         ENDIF
            IF !USED('A05')
               SELECT 0
               USE A05
            ELSE
               SELECT A05
            ENDIF
            SET ORDER TO 1
            SEEK sys_oper+'F13'
            IF FOUND()
               DELETE 
            ENDIF         
          CLOSE ALL
          THISFORM.RELEASE 
      endproc    
enddefine            
*************************************************************************************************
************************************************特殊輸入欄位的設定----料號
DEFINE CLASS TXTF01 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=20
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM F01 INTO CURSOR BMH ORDER BY F01 WHERE  F05 AND LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
       ELSE
          SELECT F01,F02 FROM F01 INTO CURSOR BMH ORDER BY F01  WHERE F05=.T.
       ENDIF        
       IF _TALLY>0
          BMH_SEEK=createobject("BMH_SEEK")  
          BMH_SEEK.SHOW    
          THIS.VALUE=FLG
          THISFORM.TXTF02.VALUE=UKEF
          THIS.REFRESH
          FLG='0'     
          UKEF=''
       ENDIF          
    ENDIF  
       THISFORM.LBLF011.CAPTION=IIF(SEEK(THIS.VALUE,'B01'),B01.F02,'')      
  ENDPROC

ENDDEFINE
**********************************************************************************************
define class ITT_SEEK as form
*  autocenter=.t.  
  MOVABLE=.F.
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*789
  FONTSIZE=INT_015*12
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  CAPTION='F13.模擬毛需求--BOM展開'
  name='ITT_SEEK'
*!*	  ADD OBJECT LBL_REC AS LBL_REC WITH;
*!*	      AUTOCENTER =.T.,;
*!*	      CAPTION='在庫數量',;
*!*	      TOP=INT_015*3
*!*	  ADD OBJECT REC_COUNT AS REC_COUNT  WITH;
*!*	      TOP=INT_015*3
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      WIDTH=INT_015*168,;  
      AUTOCENTER=.T. 
  add object ORPgroup as ORPGROUP with;            
      TOP=0,;
      LEFT=INT_015*494               
  add object cmdgroup as cmdgroup with;
      TOP=0,;
      ENABLED=.T.,;
      LEFT=INT_015*174    
  ADD OBJECT TRN_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*388,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<B.轉移轉單'          
  ADD OBJECT PHN_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*455,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<A.轉請購單'          
  ADD OBJECT XLS_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*519,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<L.轉EXCEL'                
    ADD OBJECT  INV_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*710,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<V.庫存明細'        
  add object GRID1 as GRID1 with;          
      AUTOCENTER=.T.,;
      TOP=INT_015*30,;
      WIDTH=INT_015*789,;      
      HEIGHT=INT_015*500,;  
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*18,;          
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=8,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7',;
      COLUMN8.NAME='COLUMN8'
      procedure init 
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='ITT'         
         THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
         THISFORM.PHN_BOTT.ENABLED=IIF(A02.F09='*',.T.,.F.)
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
         THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(ITT.PRS_DTE<=DATE()  AND ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02)<0,RGB(255,255,0),'')","COLUMN")                                    
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F02'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*200
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='需求日期'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F03'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                   
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='需求數量'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F04'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85
         THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應請購日'
         THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.PRS_DTE'
         THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65
         THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='在庫數量'
         THISFORM.GRID1.COLUMN6.CONTROLSOURCE="IIF(SEEK(ITT.F02+B01.F07,'B11_1'),B11_1.F04,0)"
         THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80          
         THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='最後預期結餘'
         THISFORM.GRID1.COLUMN7.CONTROLSOURCE="ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02)"       
         THISFORM.GRID1.COLUMN7.WIDTH=INT_015*90         
         THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='型態'
         THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(ITT.TPE='1','自備','客供')"
         THISFORM.GRID1.COLUMN8.WIDTH=INT_015*35         
         JK='料品編號'             
         THISFORM.GRID1.SETFOCUS        
      ENDPROC
  PROC ORPGROUP.DEL_BOTT.CLICK
       IF MESSAGEBOX('是否確定要刪除此筆資料',4+32+256,'請確認') = 6	          
          SELE ITT
          DELE
          SKIP -1
          IF BOF()
             GO TOP
          ENDIF   
       ENDIF   
  ENDPROC    
  PROC KEY_LIST.INIT 
       with this 
           .additem('依料品編號+需求日期排列')           
           .additem('依需求日期+料品編號排列')                      
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE ITT
       DO CASE
          CASE THIS.DISPLAYVALUE='依料品編號+需求日期排列'         
                     SET ORDER TO 1 
                     THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
                     THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F02'
                     THISFORM.GRID1.COLUMN1.WIDTH=INT_015*200
                     THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
                     THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
                     THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
                     THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='需求日期'
                     THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F03'
                     THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                   
                     THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='需求數量'
                     THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F04'
                     THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85
                     THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應請購日'
                     THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.PRS_DTE'
                     THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65
                     JK='料品編號'             
                     THISFORM.GRID1.SETFOCUS        

          CASE THIS.DISPLAYVALUE='依需求日期+料品編號排列'
                    SET ORDER TO 2
                     THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='需求日期'
                     THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F03'
                     THISFORM.GRID1.COLUMN1.WIDTH=INT_015*65                                       
                     THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
                     THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.F02'
                     THISFORM.GRID1.COLUMN2.WIDTH=INT_015*200
                     THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='品名規格'
                     THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149
                     THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B01.F02'
                     THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='需求數量'
                     THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F04'
                     THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85
                     THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應請購日'
                     THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.PRS_DTE'
                     THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65
                     THISFORM.GRID1.SETFOCUS        

       ENDCASE 
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC                  
      PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS XColIndex     
*       THISFORM.REC_COUNT.CAPTION=TRANSFORM(IIF(SEEK(ITT.F02+B01.F07,'B11_1'),B11_1.F04,0),'@R 99999999999.9999')
         FCH=THISFORM.ACTIVECONTROL.NAME  
         CHK=''         
         THISFORM.REFRESH
      ENDPROC       

      PROC TRN_BOTT.CLICK    &&轉移轉單
           PRT_CHOICE=CREATEOBJECT("PRT_CHOICE")  
           PRT_CHOICE.SHOW          
           SELE ITT
           DO CASE 
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依料品編號+需求日期排列'
                   SET ORDER TO 1
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依需求日期+料品編號排列'
                   SET ORDER TO 2                             
          ENDCASE             
          THIS.ENABLED=.F.
          THISFORM.GRID1.SETFOCUS         
      ENDPROC

      PROC PHN_BOTT.CLICK        &&轉請購單
           ANS_CHOICE=CREATEOBJECT("ANS_CHOICE")  
           ANS_CHOICE.SHOW          
           SELE ITT
           DO CASE 
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依料品編號+需求日期排列'
                   SET ORDER TO 1
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依需求日期+料品編號排列'
                   SET ORDER TO 2                             
          ENDCASE             
          THIS.ENABLED=.F.
          THISFORM.GRID1.SETFOCUS         
      ENDPROC            
       PROC XLS_BOTT.CLICK        &&轉EXCEL
    	    SELECT ITT.F02 料品編號,B01.F02 品名規格,ITT.F03 需求日期,ITT.F04 需求數量,ITT.PRS_DTE 應請購日,IIF(SEEK(ITT.F02+B01.F07,'B11_1'),B11_1.F04,0) 在庫數量,ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02) 預期結餘,IIF(ITT.TPE='1','自備','客供') 供料方式;
	         	   FROM ITT,B01 WHERE B01.F01=ITT.F02 ORDER BY ITT.F01 NOWAIT
	         	      FILENAME='模擬毛需求'+DTOS(DATE())
		              GCDELIMFILE=PUTFILE('儲存路徑',FILENAME,'XLS')
		              IF !EMPTY(GCDELIMFILE) 
		                    ON ERROR DO FILE_IMPACT
		                    COPY TO (GCDELIMFILE) TYPE XL5     
		                    =MESSAGEBOX('儲存成功',0+48,'提示訊息視窗') 
		              ENDIF   		              
          THISFORM.GRID1.SETFOCUS         
      ENDPROC                 
      PROC INV_BOTT.CLICK
           IF !USED('B1A')
              SELECT 0
              USE B1A
           ELSE
              SELECT B1A
           ENDIF
         SET ORDER TO 1
  	     THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
         THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(128,128,128),'COMMANDBUTTON')
         THISFORM.TRN_BOTT.FORECOLOR=RGB(128,128,128)
         THISFORM.PHN_BOTT.FORECOLOR=RGB(128,128,128)
         THISFORM.INV_BOTT.FORECOLOR=RGB(128,128,128)
         SELE B11.F01,B11.F03,A14.F02,B11.F04,B11.F05,B11.F06,B11.F08 FROM B11,A14 WHERE B11.F01<>INT_190 AND B11.F03=ITT.F02 AND A14.F01=B11.F01  INTO CURSOR INV ORDER BY B11.F01
         INV_SEEK=CREATEOBJECT("INV_SEEK")  
         INV_SEEK.SHOW                  
         SELECT B1A
         USE 
         AREA1='ITT'
         THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
         THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON') 
         THISFORM.TRN_BOTT.FORECOLOR=RGB(0,0,0)   
         THISFORM.PHN_BOTT.FORECOLOR=RGB(0,0,0)
         THISFORM.INV_BOTT.FORECOLOR=RGB(0,0,0)
         THISFORM.GRID1.SETFOCUS
      ENDPROC         
enddefine     
********************************************************
DEFINE CLASS ANS_CHOICE AS FORM
  AUTOCENTER=.T.
  CAPTION='請輸入需求備料客戶編號'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*86
  WIDTH=INT_015*250
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='ANS_CHOICE'
  ADD OBJECT LABEL1 AS LABEL WITH;
      LEFT=INT_015*3,;
      TOP=INT_015*23,;
      AUTOSIZE=.T.,;
      FONTSIZE=INT_015*9,;
      CAPTION='客戶編號'         
  ADD OBJECT LABEL11 AS LABEL WITH;
      LEFT=INT_015*122,;
      TOP=INT_015*23,;
      FONTSIZE=INT_015*9,;
      AUTOSIZE=.T.,;
      CAPTION=''    
  ADD OBJECT TEXT1 AS TEXTBOX WITH;
      LEFT=INT_015*80,;
      TOP=INT_015*20,;
      WIDTH=INT_015*40,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      MAXLENGTH=5,;       
      NAME='TEXT1'       
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;     
      TOP=INT_015*11,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;      
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*51,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND2'      
  *****    
  PROCEDURE INIT
       THISFORM.TEXT1.VALUE=OUS
       IF !EMPTY(ALLTRIM(THISFORM.TEXT1.VALUE))
            THISFORM.LABEL11.CAPTION=IIF(SEEK(THISFORM.TEXT1.VALUE,'C01'),C01.F05,'')       
       ENDIF
  ENDPROC
  PROCEDURE TEXT1.INTERACTIVECHANGE
         IF AT('?',ALLTRIM(THIS.VALUE))<>0
             IF AT('?',ALLTRIM(THIS.VALUE))>1             
                   SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
             ELSE
                   SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 
             ENDIF   
             IF _TALLY>0
                 VDR_SEEK=CREATEOBJECT("VDR_SEEK")  
                 VDR_SEEK.SHOW    
                 THIS.VALUE=FLG
                 THIS.REFRESH
                 FLG='0'
             ENDIF   
         ENDIF   
         THISFORM.LABEL11.CAPTION=IIF(SEEK(THIS.VALUE,'C01'),C01.F05,'')
  ENDPROC
  PROCEDURE CMND1.CLICK 
                IF !USED('C20')
                   SELE 0
                   USE C20
                ELSE
                   SELE C20
                ENDIF
                SET ORDER TO 1   	    

                IF !USED('P05')
                   SELE 0
                   USE P05
                ELSE
                   SELE P05
                ENDIF
                SET ORDER TO 1
                IF !USED('P06')
                   SELE 0
                   USE P06
                ELSE
                   SELE P06
                ENDIF
                SET ORDER TO 1              
                FOR I=1 TO 2
                    HD='P001 '+SUBSTR(DTOS(DATE()),3,4)
                    HD1='P001 '
                    HD2='PA'
                    HD3='材料請購單'
                    DO ODR_CRT    &&執行主程式中之單據新增   
                    IF SEEK(PO_NO,'P05')=.T.      &&如果單號重複
                       SELECT MAX(F01) FROM P05 INTO ARRAY GB NOWAIT
                       HD='P001 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                       DO ODR_RPT   &&執行主程式單據編號重複程序               
                    ENDIF                      
                      SELE P05
                      APPEND BLANK
                      REPLACE F01 WITH PO_NO
                      REPLACE F02 WITH DATE()
                      REPLACE F03 WITH '2'
                      REPLACE F04 WITH  THISFORM.TEXT1.VALUE
                      REPLACE F05 WITH IIF(SEEK(THISFORM.TEXT1.VALUE,'C01'),C01.F05,'')
                      REPLACE F06 WITH '廠內耗用'
                      REPLACE F07 WITH SYS_OPER
                      REPLACE F09 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                      REPLACE F10 WITH CKS
                      REPLACE F11 WITH '1'                
                      SELE ITT
                      SET ORDER TO 1
                      GO TOP
                      FI_NO=ITT.F02
                      DO WHILE !EOF()
                             IF ITT.TPE='1'  AND  IIF(SEEK(ITT.F02,'B01'),B01.F97,'')<>'Y'   AND IIF(I=1,LEFT(B01.F02,3)='JST',LEFT(B01.F02,3)<>'JST')   AND  IIF(I=1,LEFT(B01.F02,3)='JST',LEFT(B01.F02,3)<>'JST')   &&若不為客供物料才可以請購                                                                   
                                 SELE P06
                                 APPEND BLANK
                                 REPLACE F01 WITH PO_NO
                                 REPLACE F02 WITH ITT.F02             
                                 REPLACE F05 WITH ITT.F03
                                 REPLACE F06 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                                 REPLACE F07 WITH '廠內耗用'     
                                 SELE ITT                       
                                 DO WHILE F02=FI_NO 
                                        SELE P06              
                                        REPLACE F03 WITH F03+(ITT.F04)                                           
                                        SELE ITT   
                                        SKIP
                                  ENDDO
                            ELSE
                                 SELECT ITT
                                 SKIP
                            ENDIF     
                           FI_NO=F02                    
                      ENDDO                   	 	                      
                      DELE FROM P06 WHERE F01=PO_NO AND F03<=0
                      SELE P06                                
                      SEEK PO_NO
                      IF FOUND()
                          DO WHILE F01=PO_NO
                                 SELE B01
                                 SEEK P06.F02
                                 IF FOUND() 
                                     SELE C20                        &&計算包裝基量
                                     SEEK P06.F02
                                     IF FOUND()
                                         SELE P06                            
                                         IF C20.F15>0
                                             REPLACE  F04 WITH CEILING(F03/C20.F15)*C20.F15
                                         ELSE
                                             REPLACE F04 WITH CEILING(F03)
                                         ENDIF     
                                     ELSE
                                         SELE P06
                                         REPLACE  F04 WITH CEILING(F03)
                                     ENDIF
                                 ENDIF
                                 SELE P06
                                 SKIP
                          ENDDO
                          =MESSAGEBOX(IIF(I=1,'JST部品','非JST部品')+'已轉入請購單單號'+PO_NO,0+48,'請至P03請購單再次確認')                                                                                 
                     ELSE
                          SELE P05
                          DELE
                          IF  SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(DATE()),3,4)
                               SELE A24
                               SEEK 'P001 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                               IF !FOUND()
                                   APPEND BLANK
                                   RLOCK()
                                   REPLACE  F01 WITH 'P001 '
                                   REPLACE  F02 WITH SUBSTR(DTOS(DATE()),3,4)
                                   REPLACE  F05 WITH 'PA'+F02
                                   REPLACE  F06 WITH '請購單'
                                   REPLACE  F08 WITH PO_NO
                                   UNLOCK
                               ENDIF    
                          ENDIF            
*                          =MESSAGEBOX('請購數量或預期結餘量已足,無須再請購!',0+48,'提示訊息視窗')           
                     ENDIF     
*********************************************************************************
                ENDFOR                
                SELE P05
                USE 
                SELE P06
                USE                                      
                SELE C20
                USE     
     THISFORM.RELEASE 
 ENDPROC
 ***** 
 PROCEDURE CMND2.CLICK

      THISFORM.RELEASE
 ENDPROC      
ENDDEFINE      
**************************************************************************************************8
DEFINE CLASS PRT_CHOICE AS FORM
              AUTOCENTER=.T.
              CAPTION='請輸入領料部門編號'
              FONTSIZE=INT_015*9
              HEIGHT=INT_015*86
              WIDTH=INT_015*250
               FONTSIZE=INT_015*9
              CONTROLBOX=.F.
              BORDERSTYLE=2
              SHOWTIPS=.T.
              SHOWWINDOW=1
              WINDOWTYPE=1
               NAME='PRT_CHOICE'
              ADD OBJECT LABEL1 AS LABEL WITH;
                       LEFT=INT_015*3,;
                       TOP=INT_015*23,;
                       AUTOSIZE=.T.,;
                       FONTSIZE=INT_015*9,;
                        CAPTION='部門編號'         
             ADD OBJECT LABEL11 AS LABEL WITH;
                       LEFT=INT_015*122,;
                      TOP=INT_015*23,;
                       FONTSIZE=INT_015*9,;
                       AUTOSIZE=.T.,;
                       CAPTION=''    
             ADD OBJECT TEXT1 AS TEXTBOX WITH;
                      LEFT=INT_015*80,;
                     TOP=INT_015*20,;
                     WIDTH=INT_015*40,;
                     HEIGHT=INT_015*25,;
                     FONTSIZE=INT_015*9,;
                     MAXLENGTH=5,;       
                     NAME='TEXT1'       
            ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
                     LEFT=INT_015*180,;     
                       TOP=INT_015*11,;  
                        HEIGHT=INT_015*25,;    
                    WIDTH=INT_015*50,;
                         FONTSIZE=INT_015*9,;      
                     CAPTION='\<Y.確定',;
                  TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
                      NAME='CMND1'
               ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
                        LEFT=INT_015*180,;
                       TOP=INT_015*51,;
                          HEIGHT=INT_015*25,;    
                      WIDTH=INT_015*50,;
                            FONTSIZE=INT_015*9,;
                     CAPTION='\<C.取消',;
                          TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
                       NAME='CMND2'      
  *****    
            PROCEDURE TEXT1.INTERACTIVECHANGE
                   IF AT('?',ALLTRIM(THIS.VALUE))<>0
                       IF AT('?',ALLTRIM(THIS.VALUE))>1             
                            SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) AND F05='*'
                       ELSE
                            SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE  F05='*'
                       ENDIF   
                       IF _TALLY>0
                          DEPART_SEEK=CREATEOBJECT("DEPART_SEEK")  
                          DEPART_SEEK.SHOW    
                          THIS.VALUE=FLG
                          THIS.REFRESH
                           FLG='0'
                       ENDIF   
                   ENDIF               
                  THISFORM.LABEL11.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')
         ENDPROC
         PROCEDURE CMND1.CLICK 
            IF SEEK(THISFORM.TEXT1.VALUE,'A14')=.F.
                =MESSAGEBOX('無此部門編號!',0+48,'提示訊息視窗')
                THISFORM.TEXT1.SETFOCUS
                RETURN
            ENDIF    
           IF !USED('C20')
              SELE 0
              USE C20
          ELSE
              SELE C20
          ENDIF
          SET ORDER TO 1        
                    HD='B005 '+SUBSTR(DTOS(DATE()),3,4)
                    HD1='B005 '
                    HD2='BE'
                    HD3='移轉單'
                    DO ODR_CRT    &&執行主程式中之單據新增   
                    IF SEEK(PO_NO,'B06')=.T.      &&如果單號重複
                       SELECT MAX(F01) FROM B06 INTO ARRAY GB NOWAIT
                       HD='B005 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                       DO ODR_RPT   &&執行主程式單據編號重複程序               
                    ENDIF                                
             SELE ITT             
             SET ORDER TO 1
             GO TOP             
             QTY=IIF(SEEK(ITT.F02+B01.F07,'B11_1'), B11_1.F04,0)
             RK=0
             TQTY=0
             UKEF=ITT.F02             
             RDY=0
             DO WHILE !EOF()
                   DO WHILE F02=UKEF                    
                          QTY=QTY+ITT.F04*(-1)  &&在庫數量開始減需求量
                          IF QTY>ITT.F04*(-1) AND  ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02)-ITT.F04>=0          &&如果在庫數量大於0而且預期結餘量大於等0者方可轉入移轉單       
                              RK=RK+MIN((ITT.F04+QTY),ITT.F04)*(-1)
                              RDY=RDY+1   &&判斷是否為第一筆
                              IF RK<0 
                                 IF RDY>1   &&若不為同料號之第一筆
                                     SELECT B06
                                     REPLACE F04 WITH F04+MIN(RK*(-1),ITT.F04)
                                  ELSE                                    
                                     SELE B06
                                     APPEND BLANK
                                     REPLACE F01 WITH PO_NO
                                     REPLACE F02 WITH RIGHT(DTOC(DATE()),2)
                                     REPLACE F03 WITH ITT.F02                                  
                                     REPLACE F04 WITH MIN(RK*(-1),ITT.F04)
                                     REPLACE F05 WITH B01.F07
                                     REPLACE F10 WITH THISFORM.TEXT1.VALUE                                 
                                     REPLACE F15 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                                  
                                  ENDIF
                                  SELE ITT
                                  TQTY=TQTY+1
                              ENDIF
                              SELE ITT
                              REPL F04 WITH F04+MIN(ITT.F04+QTY,ITT.F04)*(-1)
                              IF F04<=0 
                                  DELE
                              ENDIF                                                  
                              SKIP                       
                         ELSE       
                              SELE ITT
                              IF F04<=0
                                  DELE
                              ENDIF
                              SKIP                                                
                         ENDIF         
                    ENDDO  
                    RDY=0
                    RK=0                               
                    QTY=IIF(SEEK(ITT.F02+B01.F07,'B11_1','B11_12'), B11_1.F04,0)
                    UKEF=ITT.F02                 
              ENDDO     
              IF TQTY>0                              &&轉入資料後再處理                        
                   SELE F03,SUM(F04) FROM B06 WHERE F03=ANY(SELE F01 FROM B01 WHERE F39='B') AND F01=PO_NO GROUP BY F03 ORDER BY F08 INTO CURSOR TMP3
                   SELE TMP3               &&計算整包整卷領料時此張單據還差多少量補足
                   GO TOP
                   DO WHILE !EOF()
                          SELE C20
                          SEEK TMP3.F03
                          IF FOUND()
                              TTL=CEILING(TMP3.SUM_F04/F15)*F15 
                               IF TTL>0
                                    SELE B06
                                    SEEK PO_NO+C20.F01
                                    IF FOUND()
                                        REPLACE F04 WITH TTL                                                
                                   ENDIF   
                               ENDIF
                          ENDIF  
                          SELE TMP3
                          SKIP
                    ENDDO                
                    =MESSAGEBOX(ALLTRIM(STR(TQTY))+'筆資料已轉入B06移轉單'+PO_NO,0+48,'提示訊息視窗')
              ELSE                                    &&若連一筆都沒有轉入單號必須重用
       	            IF INT_099=.T.
                        SELE A24         
                        SEEK 'B005 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                        IF !FOUND()
                            APPEND BLANK
                            LOCK()
                            REPLACE  F01 WITH 'B005 '
                            REPLACE  F02 WITH SUBSTR(DTOS(DATE()),3,4)
          	            REPLACE  F05 WITH 'BE'+f02
                            REPLACE  F06 WITH '移轉單'
                            REPLACE  F08 WITH PO_NO
                            UNLOCK
                        ENDIF    
                    ENDIF                  
                    =MESSAGEBOX('在庫數量不足無法轉入移轉單!請洽採購催料或儘速請購!',0+48,'提示訊息視窗')                
              ENDIF   
          UNLOCK ALL          
          SELE C20
          USE
          THISFORM.CMND2.CLICK
         ENDPROC
         PROCEDURE CMND2.CLICK
             THISFORM.RELEASE
        ENDPROC      
ENDDEFINE               
         
*************************************************************列印
PROCEDURE PNT_PRC  
    DO CASE 
       CASE KEY()='F02+DTOS(F03)'            
             SELECT ITT.F02,ITT.F03,ITT.F04,ITT.PRS_DTE,ITT.TPE,ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02) OSQTY,;                    
             B01.F02 FROM ITT,B01 WHERE B01.F01=ITT.F02 ORDER BY ITT.F02,ITT.F03 NOWAIT
              IF _TALLY>0
                  REPORT FORM ALLTRIM(INT_116)+'F13A' TO PRINTER PROMPT PREVIEW    
              ENDIF    
       CASE KEY()='DTOS(F03)+F02'            
             SELECT ITT.F02,ITT.F03,ITT.F04,ITT.PRS_DTE,ITT.TPE,ALTT(ITT.F02,'1')+SPL_QTY(ITT.F02) OSQTY,;                    
             B01.F02 FROM ITT,B01 WHERE B01.F01=ITT.F02 ORDER BY ITT.F03,ITT.F02 NOWAIT
             IF _TALLY>0       
                  REPORT FORM ALLTRIM(INT_116)+'F13B' TO PRINTER PROMPT PREVIEW                
              ENDIF     
    ENDCASE            
    SELECT ITT
ENDPROC       
**************************************************************************
FUNC SPL_QTY
     PARA BKK_NO
     QT_OD=0
     SELE TKE08
     SEEK BKK_NO
     IF FOUND()
        DO WHILE .NOT. EOF() .AND. F02 =BKK_NO        
           QT_OD=QT_OD+F04
           SKIP
        ENDDO
     ELSE
       QT_OD=0
     ENDIF   
     RETURN QT_OD
**********************************************************製品料號搜尋畫面
define class BMH_SEEK as form
     autocenter=.t.
     caption='製品料號一覽表'
     fontsize=INT_015*9
     HEIGHT=INT_015*280
    WIDTH=INT_015*314
    CONTROLBOX=.F.
    BORDERSTYLE=2
    SHOWTIPS=.T.
    showwindow=1
    windowtype=1
    name='BMH_SEEK'
   add object GRDB01 as GRID with;          
      AUTOCENTER=.T.,;
      top=0,;
      WIDTH=INT_015*312,;      
      HEIGHT=INT_015*250,;  
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*19,;
      ALLOWCELLSELECTION=.F.,;
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=2,;            
      RECORDSOURCE='BMH',;      
      NAME='GRDB01',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2'      
  add object cmnd1 as commandbutton with;
      autosize=.t.,;
      left=INT_015*20,;    
      HEIGHT=INT_015*25,;   
      top=INT_015*252,;  
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*9,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      caption='\<S.選取',;
      TOOLTIPTEXT='選擇游標所指的紀錄!快速鍵->ALT+S'
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      autosize=.t.,;
      left=INT_015*130,;
      top=INT_015*252,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*9,;     
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;         
      caption='\<C.取消',;
      TOOLTIPTEXT='取消此搜尋畫面!快速鍵->ALT+C'
      name='cmnd2'
      procedure init 
         THISFORM.GRDB01.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRDB01.SETALL('FONTSIZE',INT_015*9,'HEADER')      
         THISFORM.GRDB01.COLUMN1.HEADER1.CAPTION='料      號'
         THISFORM.GRDB01.COLUMN1.WIDTH=INT_015*261
         THISFORM.GRDB01.COLUMN1.CONTROLSOURCE='BMH.F01'
         THISFORM.GRDB01.COLUMN2.HEADER1.CAPTION='版次'
         THISFORM.GRDB01.COLUMN2.WIDTH=INT_015*50
         THISFORM.GRDB01.COLUMN2.CONTROLSOURCE='BMH.F02'
         THISFORM.GRDB01.SETFOCUS
         IF THISFORM.GRDB01.ACTIVEROW<>0
              THISFORM.GRDB01.AFTERROWCOLCHANGE 
         ENDIF       
      ENDPROC
      PROC GRDB01.AFTERROWCOLCHANGE
       LPARAMETERS XColIndex
       FLG=BMH.F01
       UKEF=BMH.F02
      ENDPROC 
      procedure cmnd1.click                               
          SELE BMH
          USE     
          THISFORM.RELEASE 
      endproc
      procedure cmnd2.click
          FLG=''   
          SELE BMH
          USE               
          thisform.release
      endproc
 
enddefine        

    define class tkF131 as form
               caption='輸入欲展開BOM之訂單號碼'
               fontsize=INT_015*9
               height=INT_015*270
               width=INT_015*320
               CONTROLBOX=.F.
               BORDERSTYLE=2
               AUTOCENTER=.T.
               SHOWTIPS=.T.
               showwindow=1
               windowtype=1
               name='tkB35'
               ADD OBJECT LBL1 AS LABEL WITH;
                        LEFT=INT_015*5,;
                        TOP=INT_015*40,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION='請輸入訂單號碼'
               ADD OBJECT LBL2 AS LABEL WITH;
                        LEFT=INT_015*5,;
                        TOP=INT_015*90,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION='需求客戶'                        
               ADD OBJECT LBL21 AS LABEL WITH;
                        LEFT=INT_015*105,;
                        TOP=INT_015*90,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION=''                                            
              ADD OBJECT TXT1 AS TEXTBOX WITH;
                       LEFT=INT_015*130,;
                       TOP=INT_015*35,;
                       HEIGHT=INT_015*30,;
                      WIDTH=INT_015*120,;
                       FONTSIZE=INT_015*12,;
                       MAXLENGTH=10,;
                       NAME='TXT1'    
             ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
                      LEFT=INT_015*20,;
                     TOP=INT_015*210,;
                     HEIGHT=INT_015*40,;
                     WIDTH=INT_015*80,;
                      FONTSIZE=INT_015*12,;
                     CAPTION='\<Y.確認',;
                     TOOLTIPTEXT='確認所查詢條件',;
                     NAME='CMND1'                     
             add object cmnd2 as commandbutton with;
                    left=INT_015*110,;
                    top=INT_015*210,;
                    height=INT_015*40,;
                    width=INT_015*80,;
                    FONTSIZE=INT_015*12,;
                    caption='\<X.離開',;
                   TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
                   name='cmnd2'                     
             PROCEDURE TXT1.INTERACTIVECHANGE
                  IF LEN(ALLTRIM(THIS.VALUE))=10
                      THISFORM.LBL21.CAPTION=IIF(SEEK(THIS.VALUE,'C04'),C04.F02,'')
                      THISFORM.LBL21.CAPTION=THISFORM.LBL21.CAPTION+IIF(SEEK(THISFORM.LBL21.CAPTION,'C01'),C01.F05,'')
                  ENDIF
             ENDPROC     
             PROCEDURE CMND1.CLICK
                     SELECT C03
                     SEEK THISFORM.TXT1.VALUE
                     IF FOUND()
                         SELECT C04
                         SEEK THISFORM.TXT1.VALUE
                         IF FOUND()
                              DO WHILE C04.F03=THISFORM.TXT1.VALUE
                                     IF C04.F10>0 AND IIF(SEEK(C04.F04,'B01'),B01.F97,'')='Y'
                                         SELECT EDRQ
                                         APPEND BLANK
                                         REPLACE F01 WITH C04.F04
                                         REPLACE F02 WITH IIF(SEEK(C04.F04,'F01'),F01.F02,'')
                                         REPLACE F03 WITH C04.F10
                                         REPLACE F04 WITH C03.F05
                                      ENDIF   
                                     SELECT C04
                                     SKIP
                              ENDDO
                         ENDIF
                         OUS=C03.F03
                         CKS=C03.F02
                      ELSE
                          =MESSAGEBOX('無此訂單編號!',0+48,'提示訊息視窗')
                          RETURN
                          THISFORM.TXT1.SETFOCUS
                     ENDIF        
                     THISFORM.CMND2.CLICK
             ENDPROC
             procedure cmnd2.click        
                   SELECT C03
                   USE 
                   SELECT C04
                   USE      
                   THISFORM.RELEASE                  
              endproc      
        ENDDEFINE
*******************************************************************************
