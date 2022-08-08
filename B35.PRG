***製令備料單
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
SEEK sys_oper+'B35'
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
IF !USED('B11')
     SELE 0
     USE B11
ELSE
     SELE B11
ENDIF      
SET ORDER TO B111
SELECT F01,F03,SUM(F04) FROM B11 GROUP BY F01,F03 ORDER BY F01,F03 INTO CURSOR B11_2 NOWAIT
CREATE CURSOR B11_1 (F01 C(5),F03 C(43),F04 N(14,4))
INDEX ON F01+F03 TAG B11_11
INDEX ON F03+F01 TAG B11_12
SET ORDER TO 1
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
IF !USED('E11')
   SELE 0
   USE E11
ELSE
   SELE E11
ENDIF
SET ORDER TO E111   
IF !USED('E16')
   SELE 0
   USE E16
ELSE
   SELE E16
ENDIF
SET ORDER TO E166
B35form=createobject("tkB35")
B35form.show       
define class tkB35 as form
*!*	  autocenter=.t.
  caption='B35.製令備料單搜尋範圍'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*270
  width=INT_015*320
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkB35'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入需求部門編號'
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*120,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入上線截止日期'      
  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT TXT1 AS TXT1 WITH;
      LEFT=INT_015*150,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*60,;
      FONTSIZE=INT_015*12,;
      MAXLENGTH=5,;
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*150,;
      TOP=INT_015*115,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      NAME='TXT2'          
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
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.TXT1.READONLY=.F.
         THISFORM.TXT2.VALUE=DATE()
      ENDPROC
      PROCEDURE CMND1.CLICK      
**********************************先檢查領用計劃之開單未耗量之總數是否與B11之數量相符合             
             GHK=THISFORM.TXT2.VALUE
             GHI=THISFORM.TXT1.VALUE 
             SELECT E16 
             REPLACE F05 WITH 0 FOR F06=GHI
             SELECT B11_1.F01,B11_1.F03,B11_1.F04 FROM B11_1 WHERE B11_1.F01=GHI AND B11_1.F03=ANY(SELE F02 FROM E16 WHERE F06=GHI) INTO CURSOR BTX ORDER BY B11_1.F03
             SELECT BTX
             GO TOP
             DO WHILE !EOF()
                    X=IIF(BTX.F04<0,0,BTX.F04)
                   SELECT E16
                   SEEK GHI+BTX.F03
                   IF FOUND()
                        DO WHILE E16.F06+E16.F02=GHI+BTX.F03
                               IF X>=F04
                                   REPLACE F05 WITH F04
                               ELSE
                                  REPLACE F05 WITH X    
                               ENDIF
                               X=X+F05*(-1)
                               IF X<=0
                                    EXIT
                                ENDIF    
                               SELECT E16
                               SKIP
                        ENDDO
                   ENDIF
                   SELECT BTX
                   SKIP
             ENDDO
*************************************                          
              SELECT E16         &&將E11之領料量轉成跟E16之開單未耗量一樣
              SEEK GHI
              IF FOUND()
                  DO WHILE F06=GHI
                         SELECT E11
                         SEEK E16.F01+E16.F02
                         IF FOUND()
                             REPLACE F13 WITH 0
                             REPLACE F08 WITH E16.F05
                         ENDIF
                        SELECT E16
                        SKIP
                  ENDDO
              ENDIF             
************************************************************                   
         SELE E16.F01,E16.F02,E16.F03,E16.F04-E16.F05,E16.F08,E16.F12,E16.F13;
              FROM E16 WHERE E16.F06=THISFORM.TXT1.VALUE AND E16.F04-E16.F05>0;
              AND E16.F03<=THISFORM.TXT2.VALUE AND E16.F02 =ANY(SELE F01 FROM B01 WHERE F97='N') ORDER BY E16.F02,E16.F03 INTO CURSOR ITT_1
         IF _TALLY>0           

             IF !USED('B01')
                SELE 0
                USE B01
             ELSE
                SELE B01
             ENDIF
             SET ORDER TO 1
              SELECT E16
              SET ORDER TO E164
             SELECT B11_1
             SET ORDER TO B11_12
             IF !USED('E08')
                SELE 0
                USE E08
             ELSE
               SELE E08
             ENDIF
             SET ORDER TO E083         
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
             B08='B08'+LEFT(DTOS(DATE()),6)
             IF !USED('&B08')
                  SELE 0
                  USE (B08) ALIA B08
             ELSE
                  SELE B08
             ENDIF
             SET ORDER TO 1             
             CREATE CURSOR ITT;
             (STCK_NO C(43),MO_NO C(12),QRY_DTE D(8),SHD_QTY N(13,4),RDY_QTY N(13,4),SPLY_TYPE C(1))
             INDE ON STCK_NO+DTOS(QRY_DTE) TAG ITT1
             INDE ON MO_NO+STCK_NO TAG ITT2
             INDE ON DTOS(QRY_DTE)+STCK_NO TAG ITT3
             SET ORDER TO 1             
             SET RELA TO STCK_NO INTO B01 
             SELE ITT_1
             GO TOP
             DO WHILE !EOF()
                SELE ITT
                APPEND BLANK
                REPL STCK_NO WITH ITT_1.F02
                REPL MO_NO WITH ITT_1.F01
                REPL QRY_DTE WITH ITT_1.F03
                REPL SHD_QTY WITH ITT_1.EXP_4
                REPL RDY_QTY WITH ITT_1.F08
                REPLACE SPLY_TYPE WITH ITT_1.F13
                SELE ITT_1
                SKIP   
             ENDDO
             SELE ITT_1
             USE   
             SELE F01,F02,F03,F04 FROM D07 UNION SELE F01,F02,F03,F04*(-1) FROM C05 WHERE F02= ANY(SELE STCK_NO FROM ITT) INTO CURSOR TKE081             
             SELE F01,F02,F03,F04 FROM E06 UNION SELE F01,F02,F03,F04*(-1) FROM E16 WHERE F02= ANY(SELE STCK_NO FROM ITT) INTO CURSOR TKE082
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
             SELE ITT
             GO TOP
             ITT_SEEK=createobject("ITT_SEEK")  
             ITT_SEEK.SHOW   
         ELSE
            CLOSE ALL
            =MESSAGEBOX('無此生產計劃資料 !',0+48,'提示訊息視窗')
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
           SEEK sys_oper+'B35'
           IF FOUND()
              DELETE 
           ENDIF        
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          
***************************************************特殊輸入欄位設定-----------領料部門
DEFINE CLASS TXT1 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9  
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) AND F04='*' AND F05='*' 
       ELSE
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE F04='*' AND F05='*' 
       ENDIF   
       IF _TALLY>0
          DEPART_SEEK=createobject("DEPART_SEEK")  
          DEPART_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'
       ENDIF   
    ENDIF   
       THISFORM.LBLF11.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')         
  ENDPROC
ENDDEFINE  
**********************************************************************************************
define class ITT_SEEK as form
*  autocenter=.t.  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*789
  FONTSIZE=INT_015*12
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='ITT_SEEK'
  ADD OBJECT LBL_REC AS LBL_REC WITH;
      AUTOCENTER =.T.,;
      CAPTION='在庫數量',;
      TOP=INT_015*3
  ADD OBJECT REC_COUNT AS REC_COUNT  WITH;
      TOP=INT_015*3
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      WIDTH=INT_015*168,;  
      AUTOCENTER=.T. 
  add object ORPgroup as ORPGROUP with;            
      TOP=0,;
      LEFT=INT_015*424               
  add object cmdgroup as cmdgroup with;
      TOP=0,;
      ENABLED=.T.,;
      LEFT=INT_015*174    
  ADD OBJECT TRN_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*388,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*60,;
      FONTSIZE=INT_015*9,;
      CAPTION='轉領料單'    
  ADD OBJECT PHN_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*450,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*60,;
      FONTSIZE=INT_015*9,;
      CAPTION='轉請購單'          
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
         thisform.caption='B35.製令備料單   領料部門: '+IIF(SEEK(GHI,'A14'),A14.F02,'')+'  上線截止日:'+DTOC(GHK)       
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='ITT'         
         THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
*         THISFORM.PHN_BOTT.ENABLED=.F.
*         THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
*         THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(ITT.F03-DATE()<0,RGB(255,255,0),'')","COLUMN")                                    
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.STCK_NO'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='需求日期'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.QRY_DTE'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='製令號碼'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.MO_NO'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85                            
         THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應領數量'
         THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.SHD_QTY'
         THISFORM.GRID1.COLUMN5.WIDTH=INT_015*90
         THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預期結餘'
         THISFORM.GRID1.COLUMN6.CONTROLSOURCE="ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO)"       
         THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100 
         THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期最後結餘'
         THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1'))"
         THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100         
         THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='供料方式'
         THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(ITT.SPLY_TYPE='3','客戶供料','廠內發料')"
         THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80          
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
           .additem('依製令號碼+料品編號排列')
           .additem('依需求日期+料品編號排列')                      
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE ITT
       DO CASE
          CASE THIS.DISPLAYVALUE='依料品編號+需求日期排列'         
               SET ORDER TO 1 
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.STCK_NO'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='需求日期'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.QRY_DTE'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='製令號碼'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.MO_NO'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*85                                  
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應領數量'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.SHD_QTY'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*90
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預期結餘'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE="ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO)"       
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期最後結餘'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1'))"
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100                         
               THISFORM.GRID1.SETFOCUS        
          CASE THIS.DISPLAYVALUE='依製令號碼+料品編號排列'          
               SET ORDER TO 2          
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='製令號碼'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.MO_NO'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*85                         
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.STCK_NO'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B01.F02'         
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='需求日期'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.QRY_DTE'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應領數量'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.SHD_QTY'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*90
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預期結餘'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE="ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,MO_NO)"       
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期最後結餘'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1'))"
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100                         
               THISFORM.GRID1.SETFOCUS                   
          CASE THIS.DISPLAYVALUE='依需求日期+料品編號排列'
               SET ORDER TO 3            
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='需求日期'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.QRY_DTE'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*65                            
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.STCK_NO'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B01.F02'
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='製令號碼'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.MO_NO'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85                   
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='應領數量'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.SHD_QTY'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*90
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預期結餘'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE="ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO)"       
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100 
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期最後結餘'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1'))"
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100                         
               THISFORM.GRID1.SETFOCUS                       

       ENDCASE 
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC                  
      PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS XColIndex     
       THISFORM.REC_COUNT.CAPTION=TRANSFORM(IIF(SEEK(ITT.STCK_NO+B01.F07,'B11_1'),B11_1.F04,0),'@R 99999999999.9999')
         FCH=THISFORM.ACTIVECONTROL.NAME  
         CHK=''         
         THISFORM.REFRESH
      ENDPROC       
      PROC TRN_BOTT.CLICK
             SELECT F13 FROM B08 WHERE F13<>'2'  AND F04=GHI NOWAIT
             IF _TALLY>0
                  =MESSAGEBOX('此領料部門尚有B08製令領料單未過帳單據,請全數過帳或刪除後再做此項作業',0+48,'提示訊息視窗')
                  THIS.Enabled=.F. 
                  THISFORM.GRID1.SETFOCUS
*                  RETURN
             ELSE
                  IF !USED('C20')
                        SELE 0
                        USE C20
                  ELSE
                         SELE C20
                  ENDIF
                  SET ORDER TO 1        
                  SELE E16
                  SELE RECNO() FROM E16 WHERE E16.F01+E16.F02=ANY(SELE MO_NO+STCK_NO FROM ITT) INTO ARRAY BK1
                  JJ1=''
                 IF _TALLY>0
                      FOR I=1 TO ALEN(BK1)
                              JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','                               
                      ENDFOR
                      KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
                      RLOCK(KK1,'E16')
                      LL1=RLOCK(KK1,'E16')
                 ELSE
                      LL1=.T.   
                 ENDIF        
                 SELE E11
                 SELE RECNO() FROM E11 WHERE E11.F01+E11.F04=ANY(SELE MO_NO+STCK_NO FROM ITT) INTO ARRAY BK2
                 JJ2=''
                 IF _TALLY>0
                     FOR I=1 TO ALEN(BK2)
                             JJ2=JJ2+ALLTRIM(STR(BK2(I)))+','                               
                     ENDFOR
                     KK2=STUFF(JJ2,RAT(',',JJ2,1),1,'')              
                     RLOCK(KK2,'E11')
                     LL2=RLOCK(KK2,'E11')
                 ELSE
                     LL2=.T.   
                 ENDIF             
                 IF LL1 .AND. LL2         
                     SELE F03,F04 FROM B11_1 WHERE F01=GHI AND F03= ANY(SELE STCK_NO FROM ITT) INTO CURSOR TMP1
                     SELE F02,SUM(F05) FROM E16 WHERE F06=GHI AND F02=ANY(SELE STCK_NO FROM ITT) GROUP BY F02,F03 INTO CURSOR TMP2   
                     CREATE CURSOR TMP;
                     (F01 C(43),F02 N(14,4))
                      INDE ON F01 TAG TMP
                      SET ORDER TO 1                 
                      SELE TMP1
                      GO TOP 
                      DO WHILE !EOF()
                             SELE TMP
                             APPEND BLANK
                             REPLACE  F01 WITH TMP1.F03
                             REPLACE  F02 WITH TMP1.F04
                             SELE TMP1
                             SKIP
                      ENDDO  
                      SELE TMP1
                      USE                
                      SELE TMP2
                      GO TOP
                      DO WHILE !EOF()
                             SELE TMP
                             SET ORDER TO 1
                             SEEK TMP2.F02
                             IF FOUND()
                                  REPLACE F02 WITH F02+TMP2.SUM_F05*(-1)
                             ELSE
                                  APPEND BLANK
                                  REPLACE F01 WITH TMP2.F02
                                  REPLACE F02 WITH TMP2.SUM_F05*(-1)   
                              ENDIF
                              SELE TMP2
                              SKIP   
                       ENDDO
                       SELE TMP2
                       SELE TMP
                       SET ORDER TO 1                                   
                       H='B007 '+SUBSTR(DTOS(DATE()),3,4)
                       SELE A24   
                       SET ORDER TO 1
                       SEEK H           
                       IF FOUND() 
                           PO_NO=F08
                           RLOCK()
                           DELE 
                           UNLOCK
                       ELSE
                            SELE A09
                            SEEK H            
                            IF .NOT.FOUND()                                                  &&如果為本月第一張訂單則重頭開始編號
			         APPEND BLANK                                                
			         REPLACE  f01 WITH 'B007 '
			         REPLACE  f02 WITH  SUBSTR(DTOS(DATE()),3,4)
			         REPLACE  F05 WITH 'BP'+f02
			         REPLACE  F06 WITH '製令領料單'
			         REPLACE  F08 WITH ALLTRIM(F05)+'0000'
		            ENDIF
		            LOCK()
  		            HX=VAL(RIGHT(F08,4))+1
		            PO_NO=LEFT(F08,6)+PADL(ALLTRIM(STR(HX)),4,'0')
		            REPLACE  F08 WITH PO_NO              &&記錄本單號碼,本月份後續的單號碼將以此往下編
		            UNLOCK
                       ENDIF                 
                       IF SEEK(PO_NO,'B08')=.T.                &&如果編號重複
                            SELECT A24
                            SET ORDER TO 1
                            SEEK 'B007 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                            IF FOUND()
                                DO WHILE F01+F02+F08='B007 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                                       RLOCK()
                                       DELETE
                                       UNLOCK
                                       SKIP
                                ENDDO
                            ENDIF
                            SELECT A09
                              SEEK 'B007 '+SUBSTR(DTOS(DATE()),3,4)
                              IF FOUND()		                       
                                   LOCK()		                        
                                   HX=VAL(RIGHT(F08,4))+1
                                   PO_NO=LEFT(F08,6)+PADL(ALLTRIM(STR(HX)),4,'0')
                                   REPLACE  F08 WITH PO_NO                                             &&記錄本訂單號碼,本月份後續的訂單號碼將以此往下編                                       		                        		                        
                                   UNLOCK
                              ENDIF   		                                                                    
                       ENDIF
                       SELE ITT             
                       SET ORDER TO 1
                       GO TOP             
                       RK=IIF(SEEK(ITT.STCK_NO,'TMP'),TMP.F02,0)             
                       QTY=IIF(SEEK(ITT.STCK_NO+B01.F07,'B11_1'),B11_1.F04,0)+RK
                       TQTY=0
                        UKEF=ITT.STCK_NO             
                        DO WHILE !EOF()
                               DO WHILE STCK_NO=UKEF                    
                                      QTY=QTY+ITT.SHD_QTY*(-1)
                                      IF QTY>ITT.SHD_QTY*(-1) AND  ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO)>=0          &&如果在庫數量大於0而且預期結餘量大於等0者方可轉入領料單
                                           SELE E16
                                           SEEK ITT.MO_NO+ITT.STCK_NO
                                           IF FOUND()
                                                REPLACE F05 WITH F05+MIN((ITT.SHD_QTY+QTY),ITT.SHD_QTY)
*                                              REPLACE F05 WITH F05+ITT.SHD_QTY
                                           ENDIF
                                           SELE E11
                                           SEEK ITT.MO_NO+ITT.STCK_NO    
                                           IF FOUND()
                                               REPLACE F13 WITH F13+MIN((ITT.SHD_QTY+QTY),ITT.SHD_QTY)
*                                             REPLACE F13 WITH F13+ITT.SHD_QTY
                                           ENDIF          
                                           RK=RK+MIN((ITT.SHD_QTY+QTY),ITT.SHD_QTY)*(-1)
                                           IF RK<0 
                                               SELE B08
                                               APPEND BLANK
                                               REPLACE F01 WITH PO_NO
                                               REPLACE F02 WITH RIGHT(DTOC(DATE()),2)
                                               REPLACE F04 WITH GHI
                                               REPLACE F05 WITH ITT.MO_NO
                                               REPLACE F06 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                                               REPLACE F07 WITH B01.F07
                                               REPLACE F08 WITH ITT.STCK_NO
                                               REPLACE F10 WITH MIN(RK*(-1),ITT.SHD_QTY)
                                               REPLACE F12 WITH MIN(RK*(-1),ITT.SHD_QTY)
*                                             RK=RK+MIN((ITT.SHD_QTY+QTY),ITT.SHD_QTY)+F10*(-1)
                                               SELE ITT
                                               TQTY=TQTY+1
                                           ENDIF
                                           SELE ITT
                                           REPLACE  SHD_QTY WITH SHD_QTY+MIN(ITT.SHD_QTY+QTY,ITT.SHD_QTY)*(-1)
                                           IF SHD_QTY=0 .OR. SHD_QTY-RDY_QTY<=0
                                                DELE
                                           ENDIF                                                  
                                           SKIP                       
                                     ELSE       
                                            SELE ITT
                                            IF SHD_QTY<=RDY_QTY
                                                DELE
                                            ENDIF
                                             SKIP                                                
                                      ENDIF         
                                ENDDO  
                                RK=IIF(SEEK(ITT.STCK_NO,'TMP'),TMP.F02,0)                                 
                                QTY=IIF(SEEK(ITT.STCK_NO+B01.F07,'B11_1'),B11_1.F04,0)+RK
                                UKEF=ITT.STCK_NO                 
                       ENDDO     
                       IF TQTY>0                              &&轉入資料後再處理
                            SELE B08                                        &&計算整包整卷領料時此張單據還差多少量補足
                            SELE F08,SUM(F12) FROM B08 WHERE F08=ANY(SELE F01 FROM B01 WHERE F39='B') AND F01=PO_NO GROUP BY F08 ORDER BY F08 INTO CURSOR TMP3
                            SELE TMP3
                            GO TOP
                            DO WHILE !EOF()
                                   SELE C20
                                   SEEK TMP3.F08
                                   IF FOUND()
                                       TTL=CEILING(TMP3.SUM_F12/F15)*F15-TMP3.SUM_F12
                                       IF TTL>0
                                           SELE B08
                                           APPEND BLANK
                                           REPLACE F01 WITH PO_NO
                                           REPLACE F02 WITH RIGHT(DTOC(DATE()),2)
                                           REPLACE F04 WITH GHI
                                           REPLACE F06 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                                           REPLACE F07 WITH IIF(SEEK(TMP3.F08,'B01'),B01.F07,GHI)
                                           REPLACE F08 WITH TMP3.F08
                                           REPLACE F12 WITH TTL                                                
                                       ENDIF
                                   ENDIF  
                                   SELE TMP3
                                   SKIP
                            ENDDO                
                             H=''
                             SELECT B08
                             SEEK PO_NO                                         &&重算累計數量
                             H=B08.F01+B08.F08
                             SEEK H
                             DO WHILE F01=PO_NO   
                                    K=0
                                    IF FOUND()            
                                        DO WHILE F01+F08=H
                                               K=K+F12
                                               SKIP                                      
                                        ENDDO
                                        SKIP -1
                                        REPLACE F14 WITH K
                                        SKIP
                                    ENDIF
                                    H=F01+F08
                              ENDDO                      
                             =MESSAGEBOX(ALLTRIM(STR(TQTY))+'筆資料已轉入B08製令領料單'+PO_NO,0+48,'提示訊息視窗')
                            THISFORM.GRID1.SETFOCUS             
                       ELSE                                    &&若連一筆都沒有轉入單號必須重用
       	                     IF INT_099=.T.
                                  SELE A24         
                                  SET ORDER TO 1
                                  SEEK 'B007 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                                  IF !FOUND()
                                      APPEND BLANK
                                      LOCK()
                                      REPLACE  F01 WITH 'B007 '
                                      REPLACE  F02 WITH SUBSTR(DTOS(DATE()),3,4)
          	                      REPLACE  F05 WITH 'BP'+f02
                                      REPLACE  F06 WITH '製令領料單'
                                      REPLACE  F08 WITH PO_NO
                                      UNLOCK
                                  ENDIF    
                                  SELE ITT
                                  DO CASE 
                                         CASE THISFORM.KEY_LIST.DISPLAYVALUE='依料品編號+需求日期排列'
                                                    SET ORDER TO 1
                                          CASE THISFORM.KEY_LIST.DISPLAYVALUE='依製令號碼+料品編號排列'
                                                     SET ORDER TO 2
                                          CASE THISFORM.KEY_LIST.DISPLAYVALUE='依需求日期+料品編號排列'
                                                      SET ORDER TO 3                            
                                  ENDCASE   
                          ENDIF                  
                                =MESSAGEBOX('庫存數量不足無法轉入領料單!請先請洽採購催料或儘速請購!',0+48,'提示訊息視窗')      
                                SELECT ITT
                                THISFORM.GRID1.SETFOCUS
                                THIS.Enabled=.F.          
                      ENDIF   
              ELSE
                    DO FILE_IMPACT 
                          UNLOCK ALL          
                          RETURN
               ENDIF   
               RELEASE ALL LIKE BK*
               UNLOCK ALL          
               SELE TMP
               USE
               SELE C20
               USE
               THIS.ENABLED=.F.
               THISFORM.GRID1.SETFOCUS         
                IF THISFORM.GRID1.ACTIVEROW<>0
*                     THISFORM.PHN_BOTT.ENABLED=.T.
                ELSE   
                      THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
                       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
                       THISFORM.PHN_BOTT.ENABLED=.F.            
                       THISFORM.CMDGROUP.ENABLED=.F.
                 ENDIF
            ENDIF      
      ENDPROC
      PROC PHN_BOTT.CLICK
           ANS_CHOICE=CREATEOBJECT("ANS_CHOICE")  
           ANS_CHOICE.SHOW          
           SELE ITT
           DO CASE 
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依料品編號+需求日期排列'
                   SET ORDER TO 1
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依製令號碼+料品編號排列'
                   SET ORDER TO 2
              CASE THISFORM.KEY_LIST.DISPLAYVALUE='依需求日期+料品編號排列'
                   SET ORDER TO 3                             
          ENDCASE             
          THIS.ENABLED=.F.
          THISFORM.GRID1.SETFOCUS         
      ENDPROC      
enddefine     
********************************************************
DEFINE CLASS ANS_CHOICE AS FORM
  AUTOCENTER=.T.
  CAPTION='請選擇轉請購單之依據'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*115
  WIDTH=INT_015*210
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='ANS_CHOICE'
  ADD OBJECT DATA_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*28,;
      TOP=INT_015*20,;
      AUTOSIZE=.T.,;
      FONTSIZE=INT_015*12,;
      BUTTONCOUNT=3,;
      OPTION1.CAPTION='生產應領量',;
      OPTION1.TOP=0,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.HEIGHT=INT_015*17,;
      OPTION2.CAPTION='預期結餘量',;
      OPTION2.TOP=20,;
      OPTION2.AUTOSIZE=.T.,;
      OPTION2.FONTSIZE=INT_015*9,;  
      OPTION2.HEIGHT=INT_015*17,;          
      OPTION3.CAPTION='最後結餘量',;
      OPTION3.TOP=40,;
      OPTION3.AUTOSIZE=.T.,;
      OPTION3.FONTSIZE=INT_015*9,;  
      OPTION3.HEIGHT=INT_015*17,;          
      NAME='DATA_OPTION'             
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*130,;     
      TOP=INT_015*11,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;      
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*130,;
      TOP=INT_015*51,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND2'      
  *****    
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
                        H='P001 '+SUBSTR(DTOS(DATE()),3,4)
                        SELE A24   
                        SET ORDER TO 1
                        SEEK H           
                        IF FOUND() 
                            PO_NO=F08
                            RLOCK()
                            DELE 
                            UNLOCK
                        ELSE
                            SELE A09
                            SEEK H            
                            IF  !FOUND()                                                  &&如果為本月第一張訂單則重頭開始編號
	    		         APPEND BLANK                                                
		    	         REPLACE  f01 WITH 'P001 '
			         REPLACE  f02 WITH  SUBSTR(DTOS(DATE()),3,4)
			         REPLACE  F05 WITH 'PA'+f02
    			         REPLACE  F06 WITH '請購單'
	    		         REPLACE  F08 WITH ALLTRIM(F05)+'0000'
		            ENDIF
		            LOCK()
    		            HX=VAL(RIGHT(F08,4))+1
	    	            PO_NO=LEFT(F08,6)+PADL(ALLTRIM(STR(HX)),4,'0')
		            REPLACE  F08 WITH PO_NO              &&記錄本單號碼,本月份後續的單號碼將以此往下編
		            UNLOCK
                        ENDIF    
                        IF SEEK(PO_NO,'P05')=.T.      &&如果編號重複
                             SELE A24
                             SEEK 'P001 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                             IF FOUND()
                                  DO WHILE  F01+F02+F08='P001 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                                         RLOCK()
                                         DELE
                                         UNLOCK
                                         SKIP
                                   ENDDO
                              ENDIF                                   
                              SELE A09
                              SEEK 'P001 '+SUBSTR(DTOS(DATE()),3,4)
                              IF FOUND()		                       
                                   LOCK()		                        
                                   HX=VAL(RIGHT(F08,4))+1
                                   PO_NO=LEFT(F08,6)+PADL(ALLTRIM(STR(HX)),4,'0')
                                   REPLACE  F08 WITH PO_NO                                             &&記錄本訂單號碼,本月份後續的訂單號碼將以此往下編                                       		                        		                        
                                   UNLOCK
                              ENDIF   		                                        
                        ENDIF                                              
                        SELE ITT
                        SET ORDER TO 1
                        GO TOP
                        FI_NO=ITT.STCK_NO
                        DO WHILE !EOF()  
                               IF IIF(SEEK(ITT.STCK_NO,'B01'),B01.F97,'')<>'Y'   AND IIF(I=1,LEFT(B01.F02,3)='JST',LEFT(B01.F02,3)<>'JST')  AND IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1'))<0              
                                    SELE P06
                                    APPEND BLANK
                                    REPLACE F01 WITH PO_NO
                                    REPLACE F02 WITH ITT.STCK_NO             
                                    REPLACE F05 WITH ITT.QRY_DTE
                                    REPLACE F06 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                                    REPLACE F07 WITH '廠內耗用'     
                                    SELE ITT                       
                                    DO WHILE STCK_NO=FI_NO 
                                           IF SHD_QTY-RDY_QTY>0 .AND. ITT.SPLY_TYPE<>'3'        &&如果預期結餘小於 0且不為客供品
                                                PO_FLAG=0   &&判斷是否有真的請購之變數
                                                SELE P06    
                                                DO CASE          
                                                      CASE THISFORM.DATA_OPTION.VALUE=1                                   &&依照應領數量請購            
                                                                  REPLACE F03 WITH F03+(ITT.SHD_QTY-ITT.RDY_QTY)
                                                                  PO_FLAG=(ITT.SHD_QTY-ITT.RDY_QTY)
                                                      CASE THISFORM.DATA_OPTION.VALUE=2                                &&依照預期結餘量不足者請購                                                       
                                                                  REPLACE F03 WITH (ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO))*(-1)
                                                                  PO_FLAG=P06.F03
                                                      CASE THISFORM.DATA_OPTION.VALUE=3                                &&依照最後預期結餘量不足者請購
                                                                 REPLACE F03 WITH (IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1')))*(-1)         
                                                                 PO_FLAG= P06.F03
                                                ENDCASE                  
                                                SELE E16
                                                SEEK ITT.MO_NO+ITT.STCK_NO
                                                IF FOUND()
                                                    REPLACE F08 WITH F08+(ITT.SHD_QTY-ITT.RDY_QTY)
                                                    IF PO_FLAG>0   &&若真請購才將請購單號回寫回領用計劃
                                                        REPLACE F14 WITH PO_NO
                                                    ENDIF     
                                                ENDIF
                                                SELE E11
                                                SEEK ITT.MO_NO+ITT.STCK_NO
                                                IF FOUND()
                                                    REPLACE F19 WITH F19 +(ITT.SHD_QTY-ITT.RDY_QTY)
                                                ENDIF                                                
                                                SELE ITT
                                                REPLACE RDY_QTY WITH RDY_QTY+P06.F03
                                           ENDIF
                                           SELE ITT   
                                           SKIP
                                    ENDDO
                                    FI_NO=STCK_NO            
                                ELSE
                                     SELECT ITT
                                     SKIP     
                               ENDIF               
                       ENDDO                   	 	                      
                       SELECT  P06                
                       DELE FOR F03<=0
                       SEEK PO_NO
                       IF FOUND()
                           DO WHILE F01=PO_NO
                                  SELE B01
                                  SEEK P06.F02
                                  IF FOUND() 
                                       SELE C20
                                       SEEK P06.F02
                                       IF FOUND()
                                           SELE P06
                                           REPLACE  F04 WITH CEILING(F03/C20.F15)*C20.F15
                                       ELSE
                                           SELE P06
                                           REPLACE  F04 WITH F03   
                                       ENDIF
                                  ENDIF
                                  SELE P06
                                  SKIP
                           ENDDO
                           SELE P05
                           APPEND BLANK
                           REPLACE F01 WITH PO_NO
                           REPLACE F02 WITH DATE()
                           REPLACE F03 WITH '2'
                           REPLACE F04 WITH  GHI
                           REPLACE F05 WITH IIF(SEEK(GHI,'A14'),A14.F02,'')
                           REPLACE F06 WITH '廠內耗用'
                           REPLACE F07 WITH SYS_OPER
                           REPLACE F09 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                           REPLACE F11 WITH '1'                                   
                           =MESSAGEBOX(IIF(I=1,'JST部品','非JST部品')+'已轉入請購單單號'+PO_NO,0+48,'請至P03請購單再次確認')                                                                                 
                       ELSE
                           IF  SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(DATE()),3,4)
                                SELE A24
                                SET ORDER TO 1
                                SEEK 'P001 '+SUBSTR(DTOS(DATE()),3,4)+PO_NO
                                IF !FOUND()
                                    APPEND BLANK
                                     LOCK()
                                     REPLACE  F01 WITH 'P001 '
                                     REPLACE  F02 WITH SUBSTR(DTOS(DATE()),3,4)
                                     REPLACE  F05 WITH 'PA'+F02
                                     REPLACE  F06 WITH '請購單'
                                     REPLACE  F08 WITH PO_NO
                                     UNLOCK
                                 ENDIF    
                           ENDIF            
                           =MESSAGEBOX(IIF(I=1,'JST部品','非JST部品')+'請購數量或預期結餘量已足,無須再請購!',0+48,'提示訊息視窗')                                
                       ENDIF     
                ENDFOR       
                SELE P05
                USE 
                SELE P06
                USE                                      
                SELE C20
                USE     
     THIS.Enabled=.F.                
     THISFORM.RELEASE 
 ENDPROC
 ***** 
 PROCEDURE CMND2.CLICK
            IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'B35'
           IF FOUND()
              DELETE 
           ENDIF        
      THISFORM.RELEASE
 ENDPROC      
ENDDEFINE      



*************************************************************列印
PROCEDURE PNT_PRC  
    DO CASE 
       CASE KEY()='STCK_NO+DTOS(QRY_DTE)'            
                SELECT ITT.*,ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO) OSQTY,;
                 IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1')) BSQTY FROM ITT ORDER BY ITT.STCK_NO,ITT.QRY_DTE NOWAIT
                IF _TALLY>0
                       REPORT FORM ALLTRIM(INT_116)+'B35A' PREVIEW                                
                ENDIF
       CASE KEY()='MO_NO+STCK_NO'            
                SELECT ITT.*,ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO) OSQTY,;
                 IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1')) BSQTY FROM ITT ORDER BY ITT.MO_NO,ITT.STCK_NO NOWAIT       
                 IF _TALLY>0
                     REPORT FORM ALLTRIM(INT_116)+'B35B' PREVIEW                
                 ENDIF    
       CASE KEY()='DTOS(QRY_DTE)+STCK_NO'   
                SELECT ITT.*,ALTT(ITT.STCK_NO,'1')+MRP_QTY(ITT.STCK_NO,ITT.QRY_DTE,ITT.MO_NO) OSQTY,;
                 IIF(SEEK(ITT.STCK_NO,'E08'),ALTT(ITT.STCK_NO,'1')+E08.F20,ALTT(ITT.STCK_NO,'1')) BSQTY FROM ITT ORDER BY ITT.QRY_DTE,ITT.STCK_NO NOWAIT              
                  IF _TALLY>0         
                      REPORT FORM ALLTRIM(INT_116)+'B35C' PREVIEW     
                  ENDIF    
    ENDCASE            
    SELECT ITT
ENDPROC       
**************************************************************************
FUNC MRP_QTY
     PARA BKK_NO,DTE_SE,MN_P
*     PRIV BKK_NO,QT_OD
*     BKK_NO=C05.F02
     QT_OD=0
     SELE TKE08
     SEEK BKK_NO
     IF FOUND()
        DO WHILE .NOT. EOF() .AND. F02+DTOS(F03)+F01<=BKK_NO+DTOS(DTE_SE)+MN_P
           QT_OD=QT_OD+F04
           SKIP
        ENDDO
     ELSE
       QT_OD=0
     ENDIF   
     RETURN QT_OD