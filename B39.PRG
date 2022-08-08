**程式名稱:未過帳庫存單據
close all
clear 
AREA1='B39'
FLG='0'
FCH=''
CHK=''
IF !USED('A01')
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
SEEK sys_oper+'B39'   
IF !USED('A03')
   SELE 0
   USE A03
ELSE
   SELE A03
ENDIF
SET ORDER TO 1         
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1 
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET ORDER TO 1   
IF !USED('C01')
   SELE 0
   USE C01
ELSE 
   SELE C01
ENDIF
SET ORDER TO C011                  
IF !USED('D01')
   SELE 0
   USE D01
ELSE 
   SELE D01
ENDIF
SET ORDER TO D011                  
IF !USED('E02')
     SELECT 0
     USE E02
ELSE
     SELECT E02
ENDIF
SET ORDER TO E021          
B39='B39'+LEFT(DTOS(DATE()),6)   &&未過帳單劇檔
IF !USED('&B39')
   SELE 0
   USE (B39) ALIA B39 INDE (B39)
ELSE
   SELE B39
ENDIF
SET ORDER TO 1                
GO TOP
SET RELATION  TO F01 INTO A03
SET RELATION  TO F07 INTO A14 ADDI 
SET RELATION  TO F04 INTO C01 ADDI 
SET RELATION  TO F04 INTO D01 ADDI
SET RELATION TO F04 INTO E02 ADDI
B39form=createobject("tkB39")
B39form.show  
define class tkB39 as form
  caption='B39.未過帳單據查詢'
  controlbox=.F.
  BORDERSTYLE=3  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  controlcount=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  showtips=.t.
  showwindow=1
  WIDTH=INT_015*789
  windowtype=1
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      caption=str(reccount()),;
      TOP=INT_015*5
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      autocenter=.t.      
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=9,;            
      RECORDSOURCE='B39',; 
      HEIGHT=INT_015*500,;     
      WIDTH=INT_015*780,;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',; 
      COLUMN4.NAME='COLUMN4',; 
      COLUMN5.NAME='COLUMN5',;             
      COLUMN6.NAME='COLUMN6',;       
      COLUMN7.NAME='COLUMN7',;       
      COLUMN8.NAME='COLUMN8',;
      COLUMN9.NAME='COLUMN9'
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;      
      TOP=INT_015*0,;
      LEFT=INT_015*350              
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=INT_015*0,;
      LEFT=INT_015*250,;
      ENABLED=.T.,;
      VISIBLE=.T.                     
  ADD OBJECT RESCH AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*430,;
      WIDTH=INT_015*70,;
      HEIGHT=INT_015*25,;
      ENABLED=.T.,;
      VISIBLE=.T.,;
      FONTSIZE=INT_015*9,;
      NAME='RESCH',;
      CAPTION='重新搜尋'
  PROCEDURE INIT   
       THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
       THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*40 &&程式編號
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*90 &&單據名稱
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80 &&單據編號 
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*25 &&日期 
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80 &&廠商/客戶
       THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80 &&轉出部門名稱
       THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80 &&轉入部門名稱
       THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80 &&訂單/製令
       THISFORM.GRID1.COLUMN9.WIDTH=INT_015*70 &&補貨日
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='程式編號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='單據名稱'
	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='單據編號'
	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='日期'	   	     	   	   
       THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='客戶/廠商'
	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='轉出部門'	   	   
	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='轉入部門'
  	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='訂單/製令編號'  	   
  	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='補貨日期'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B39.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A03.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B39.F02'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B39.F03'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='IIF(C01.F05<>SPACE(8),C01.F05,IIF(D01.F04<>SPACE(8),D01.F04,E02.F04))'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B39.F06'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'       
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='B39.F08'
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B39.F09'
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS
       JK='程式編號'
  ENDPROC     
  PROC MTH_LIST.INTERACTIVECHANGE
       CHK_STR=IIF(SEEK(DTOS(CTOD((THIS.DISPLAYVALUE+'/01'))),'A23','A231'),LEFT(DTOC(A23.F01),5),'')
       THISFORM.GRID1.RECORDSOURCE=''
       SELE B39
       USE
       B39='B39'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&B39')
          SELE 0
          USE (B39) ALIA B39
       ELSE
          SELE B39
       ENDIF
       SET ORDER TO 1        
       GO TOP
       SET RELA TO F01 INTO A03
       SET RELA TO F07 INTO A14 ADDI 
       SET RELA TO F04 INTO C01 ADDI 
       SET RELA TO F04 INTO D01 ADDI
       THISFORM.GRID1.RECORDSOURCE='B39'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B39.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A03.F02'       
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B39.F02'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B39.F03'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='IIF(C01.F05=SPACE(8),D01.F04,C01.F05)'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B39.F06'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='B39.F08'
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B39.F09'
	   thisform.grid1.setfocus        
	   THISFORM.REFRESH
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
       ENDIF          
  ENDPROC      
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROC RESCH.CLICK
       WAIT WINDOW "重新抓取記錄中,請稍後!" AT 0,150 NOWAIT NOCLEAR
       SELE B39
       DELE ALL  
         B02='B02'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&進貨單
         IF !USED('&B02') 
            SELE 0
            USE (B02) ALIA B02
         ELSE
            SELE B02
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05,F07,F09 FROM B02 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR B02_2 
         IF _TALLY>0
            SELE B02_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B02'
               REPL F02 WITH B02_2.F01
               REPL F03 WITH B02_2.F02
               REPL F04 WITH B02_2.F07
               REPL F07 WITH B02_2.F05
               REPL F08 WITH B02_2.F09
               SELE B02_2
               SKIP
            ENDDO    
         ENDIF  
         SELE B02_2
         USE
         SELE B02 
         USE               
*********************************************
         B03='B03'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&進貨退出單
         IF !USED('&B03') 
            SELE 0
            USE (B03) ALIA B03
         ELSE
            SELE B03
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05,F07,F09,F14 FROM B03 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR B03_2
         IF _TALLY>0
            SELE B03_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B03'
               REPL F02 WITH B03_2.F01
               REPL F03 WITH B03_2.F02
               REPL F04 WITH B03_2.F07
               REPL F05 WITH B03_2.F05
               REPL F06 WITH IIF(SEEK(F05,'A14'),A14.F02,'')
               REPL F08 WITH B03_2.F09
               REPL F09 WITH B03_2.F14
               SELE B03_2
               SKIP               
            ENDDO
         ENDIF   
         SELE B03_2
         USE
         SELE B03
         USE
***********************************************         
         B04='B04'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&出貨單
         IF !USED('&B04') 
            SELE 0
            USE (B04) ALIA B04
         ELSE
            SELE B04
         ENDIF
         SET ORDER TO 1      
         SELE F01,F02,F05,F06,F07 FROM B04 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR B04_2
         IF _TALLY>0
            SELE B04_2
            GO TOP              
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B04'
               REPL F02 WITH B04_2.F01
               REPL F03 WITH B04_2.F02
               REPL F04 WITH B04_2.F06
               REPL F05 WITH B04_2.F05        
               REPL F06 WITH IIF(SEEK(B04_2.F05,'A14'),A14.F02,'')
               REPL F08 WITH B04_2.F07
               SELE B04_2
               SKIP
            ENDDO
         ENDIF 
         SELE B04_2
         USE
         SELE B04
         USE
***************************************        
         B05='B05'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&出貨退回單
         IF !USED('&B05') 
            SELE 0
            USE (B05) ALIA B05
         ELSE
            SELE B05
         ENDIF
         SET ORDER TO 1             
         SELE F01,F02,F05,F10,F11,F14 FROM B05 WHERE EMPTY(F12) GROUP BY F01 ORDER BY F01 INTO CURSOR B05_2
         IF _TALLY>0
            SELE B05_2
            GO TOP                
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B05'
               REPL F02 WITH B05_2.F01
               REPL F03 WITH B05_2.F02
               REPL F04 WITH B05_2.F10
               REPL F07 WITH B05_2.F05
               REPL F08 WITH B05_2.F11
               REPL F09 WITH B05_2.F14               
               SELE B05_2               
               SKIP
            ENDDO
         ENDIF   
         SELE B05_2
         USE
         SELE B05
         USE
*********************************************                 
         B06='B06'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&移轉單
         IF !USED('&B06') 
            SELE 0
            USE (B06) ALIA B06
         ELSE
            SELE B06
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05,F07,F08,F10,F15 FROM B06 WHERE EMPTY(F07) GROUP BY F01 ORDER BY F01 INTO CURSOR B06_2
         IF _TALLY>0
            SELE B06_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B06'
               REPL F02 WITH B06_2.F01
               REPL F03 WITH B06_2.F02
               REPL F05 WITH B06_2.F05
               REPL F06 WITH IIF(SEEK(B06_2.F05,'A14'),A14.F02,'')
               REPL F07 WITH B06_2.F10               
               SELE B06_2
               SKIP
            ENDDO
         ENDIF                        
         SELE B06_2
         USE
         SELE B06
         USE
******************************************
         B07='B07'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&委外入庫單
         IF !USED('&B07') 
            SELE 0
            USE (B07) ALIA B07
         ELSE
            SELE B07
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05,F07,F09 FROM B07 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR B07_2 
         IF _TALLY>0
            SELE B07_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B07'
               REPL F02 WITH B07_2.F01
               REPL F03 WITH B07_2.F02
               REPL F04 WITH B07_2.F07
               REPL F07 WITH B07_2.F05
               REPL F08 WITH B07_2.F09
               SELE B07_2
               SKIP
            ENDDO    
         ENDIF  
         SELE B07_2
         USE
         SELE B07 
         USE                        
******************************************
         B08='B08'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&製令領料單
         IF !USED('&B08') 
            SELE 0
            USE (B08) ALIA B08
         ELSE
            SELE B08
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F07,F04 FROM B08 WHERE EMPTY(F13) GROUP BY F01 ORDER BY F01 INTO CURSOR B08_2 
         IF _TALLY>0
            SELE B08_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B08'
               REPL F02 WITH B08_2.F01
               REPL F03 WITH B08_2.F02
               REPL F05 WITH B08_2.F07
               REPL F06 WITH IIF(SEEK(B08_2.F07,'A14'),A14.F02,'')
               REPL F07 WITH B08_2.F04
               SELE B08_2
               SKIP
            ENDDO    
         ENDIF  
         SELE B08_2
         USE
         SELE B08 
         USE                                 
*****************************************  
         B09='B09'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&盤差調整單
         IF !USED('&B09') 
            SELE 0
            USE (B09) ALIA B09
         ELSE
            SELE B09
         ENDIF
         SET ORDER TO 1                           
         SELE F01,F02,F05 FROM B09 WHERE EMPTY(F11) GROUP BY F01 ORDER BY F01 INTO CURSOR B09_2
         IF _TALLY>0
            SELE B09_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B09'
               REPL F02 WITH B09_2.F01
               REPL F03 WITH B09_2.F02
               REPL F05 WITH B09_2.F05
               REPL F06 WITH IIF(SEEK(B09_2.F05,'A14'),A14.F02,'')               
               SELE B09_2
               SKIP
            ENDDO
         ENDIF   
         SELE B09_2
         USE
         SELE B09
         USE
*****************************************  
         B10='B10'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&報廢單
         IF !USED('&B10') 
            SELE 0
            USE (B10) ALIA B10
         ELSE
            SELE B10
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05 FROM B10 WHERE EMPTY(F09) GROUP BY F01 ORDER BY F01 INTO CURSOR B10_2
         IF _TALLY>0
            SELE B10_2
            GO TOP    
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B10'
               REPL F02 WITH B10_2.F01
               REPL F03 WITH B10_2.F02
               REPL F05 WITH B10_2.F05
               REPL F06 WITH IIF(SEEK(B10_2.F05,'A14'),A14.F02,'')
               SELE B10_2
               SKIP
            ENDDO                        
         ENDIF   
         SELE B10_2
         USE
         SELE B10
         USE
*****************************************  
         B16='B16'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&製令用料耗損單
         IF !USED('&B16') 
            SELE 0
            USE (B16) ALIA B16
         ELSE
            SELE B16
         ENDIF
         SET ORDER TO 1     
         SELE F01,F02,F07,F04 FROM B16 WHERE EMPTY(F13) GROUP BY F01 ORDER BY F01 INTO CURSOR B16_2
         IF _TALLY>0
            SELE B16_2 
            GO TOP            
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B33'
               REPL F02 WITH B16_2.F01
               REPL F03 WITH B16_2.F02
               REPL F05 WITH B16_2.F07
               REPL F06 WITH IIF(SEEK(B16_2.F07,'A14'),A14.F02,'')
               REPL F07 WITH B16_2.F04               
               SELE B16_2
               SKIP
            ENDDO            
         ENDIF                                  
         SELE B16_2
         USE
         SELE B16
         USE
*****************************************  
         B12='B12'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&料品使用對象調撥單
         IF !USED('&B12') 
            SELE 0
            USE (B12) ALIA B12
         ELSE
            SELE B12
         ENDIF
         SET ORDER TO 1     
         SELE F01,F02,F05,F08 FROM B12 WHERE EMPTY(F09) GROUP BY F01 ORDER BY F01 INTO CURSOR B12_2 NOWAIT
         IF _TALLY>0
            SELE B12_2 
            GO TOP            
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B36'
               REPL F02 WITH B12_2.F01
               REPL F03 WITH B12_2.F02
               REPL F05 WITH B12_2.F05
               REPL F06 WITH IIF(SEEK(B12_2.F05,'A14'),A14.F02,'')
               REPLACE F04 WITH B12_2.F08
               SELE B12_2
               SKIP
            ENDDO            
         ENDIF                                  
         SELE B12_2
         USE
         SELE B12
         USE         
**********************************************************************************                  
         B40='B40'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&額外生產耗料單
         IF !USED('&B40') 
            SELE 0
            USE (B40) ALIA B40
         ELSE
            SELE B40
         ENDIF
         SET ORDER TO 1          
         SELE F01,F02,F05 FROM B40 WHERE EMPTY(F11) GROUP BY F01 ORDER BY F01 INTO CURSOR B40_1
         IF _TALLY>0
            SELE B40_1
            GO TOP    
            DO WHILE !EOF()
                   SELE B39
                   APPEND BLANK
                   REPL F01 WITH 'B40'
                   REPL F02 WITH B40_1.F01
                   REPL F03 WITH B40_1.F02
                   REPL F05 WITH B40_1.F05
                   REPL F06 WITH IIF(SEEK(B40_1.F05,'A14'),A14.F02,'')
                   SELE B40_1
                   SKIP
            ENDDO                        
         ENDIF   
         SELE B40_1
         USE
         SELE B40
         USE         
******************************************         
         B55='B55'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&完工入庫單
         IF !USED('&B55') 
            SELE 0
            USE (B55) ALIA B55
         ELSE
            SELE B55
         ENDIF
         SET ORDER TO 1                   
         SELE F01,F02,F07,F04 FROM B55 WHERE EMPTY(F13) GROUP BY F01 ORDER BY F01 INTO CURSOR B55_2
         IF _TALLY>0
            SELE B55_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B55'
               REPL F02 WITH B55_2.F01
               REPL F03 WITH B55_2.F02
               REPL F05 WITH B55_2.F07
               REPL F06 WITH IIF(SEEK(B55_2.F07,'A14'),A14.F02,'')
               REPL F07 WITH B55_2.F04               
               SELE B55_2
               SKIP
            ENDDO
         ENDIF   
         SELE B55_2
         USE
         SELE B55
         USE
******************************************
         B56='B56'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&製品退出/折讓單
         IF !USED('&B56') 
            SELE 0
            USE (B56) ALIA B56
         ELSE
            SELE B56
         ENDIF
         SET ORDER TO 1                   
         SELE F01,F02,F05,F07,F14,F15,F19 FROM B56 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR B56_2
         IF _TALLY>0
            SELE B56_2
            GO TOP         
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPLACE F01 WITH 'B56'
               REPLACE F02 WITH B56_2.F01
               REPLACE F03 WITH B56_2.F02
               REPLACE F04 WITH B56_2.F07
               REPLACE F05 WITH B56_2.F05
               REPLACE F06 WITH IIF(SEEK(B56_2.F05,'A14'),A14.F02,'')
               REPLACE F07 WITH B56_2.F14     
               REPLACE F08 WITH B56_2.F15
               REPLACE F09 WITH B56_2.F19
               SELE B56_2
               SKIP
            ENDDO
         ENDIF   
         SELE B56_2
         USE
         SELE B56
         USE         
******************************************         
         B61='B61'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&借入單
         IF !USED('&B61') 
            SELE 0
            USE (B61) ALIA B61
         ELSE
            SELE B61
         ENDIF
         SET ORDER TO 1                            
         SELE F01,F02,F05,F10 FROM B61 WHERE EMPTY(F12) GROUP BY F01 ORDER BY F01 INTO CURSOR B61_2
         IF _TALLY>0
            SELE B61_2 
            GO TOP                            
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B61'
               REPL F02 WITH B61_2.F01
               REPL F03 WITH B61_2.F02
               REPL F04 WITH B61_2.F10
               REPL F05 WITH B61_2.F05
               REPL F06 WITH IIF(SEEK(B61_2.F05,'A14'),A14.F02,'')               
               SELE B61_2
               SKIP
            ENDDO
         ENDIF   
         SELE B61_2
         USE
         SELE B61
         USE
******************************************         
         B62='B62'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&借入還出單
         IF !USED('&B62') 
            SELE 0
            USE (B62) ALIA B62
         ELSE
            SELE B62
         ENDIF
         SET ORDER TO 1         
         SELE F01,F02,F05,F10 FROM B62 WHERE EMPTY(F12) GROUP BY F01 ORDER BY F01 INTO CURSOR B62_2   
         IF _TALLY>0
            SELE B62_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B62'
               REPL F02 WITH B62_2.F01
               REPL F03 WITH B62_2.F02
               REPL F04 WITH B62_2.F10
               REPL F05 WITH B62_2.F05
               REPL F06 WITH IIF(SEEK(B62_2.F05,'A14'),A14.F02,'')
               SELE B62_2
               SKIP
            ENDDO
         ENDIF                             
         SELE B62_2
         USE
         SELE B62
         USE     
******************************************         
         B63='B63'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&借出單
         IF !USED('&B63') 
            SELE 0
            USE (B63) ALIA B63
         ELSE
            SELE B63
         ENDIF
         SET ORDER TO 1      
         SELE F01,F02,F05,F10 FROM B63 WHERE EMPTY(F12) GROUP BY F01 ORDER BY F01 INTO CURSOR B63_2
         IF _TALLY>0
            SELE B63_2
            GO TOP                                        
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B63'
               REPL F02 WITH B63_2.F01
               REPL F03 WITH B63_2.F02
               REPL F04 WITH B63_2.F10
               REPL F05 WITH B63_2.F05
               REPL F06 WITH IIF(SEEK(B63_2.F05,'A14'),A14.F02,'')               
               SELE B63_2
               SKIP
            ENDDO 
         ENDIF   
         SELE B63_2
         USE
         SELE B63
         USE
******************************************         
         B64='B64'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&借出還入單
         IF !USED('&B64') 
            SELE 0
            USE (B64) ALIA B64
         ELSE
            SELE B64
         ENDIF
         SET ORDER TO 1         
         SELE F01,F02,F05,F10 FROM B64 WHERE EMPTY(F12) GROUP BY F01 ORDER BY F01 INTO CURSOR B64_2
         IF _TALLY>0
            SELE B64_2
            GO TOP  
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'B64'
               REPL F02 WITH B64_2.F01
               REPL F03 WITH B64_2.F02
               REPL F04 WITH B64_2.F10
               REPL F05 WITH B64_2.F05
               REPL F06 WITH IIF(SEEK(B64_2.F05,'A14'),A14.F02,'')               
               SELE B64_2
               SKIP
            ENDDO
         ENDIF                       
         SELE B64_2
         USE
         SELE B64
         USE               
******************************************         
         C38='C38'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&樣品領用單
         IF !USED('&C38') 
            SELE 0
            USE (C38) ALIA C38
         ELSE
            SELE C38
         ENDIF
         SET ORDER TO 1                       
         SELE F01,F02,F05,F10 FROM C38 WHERE EMPTY(F07) GROUP BY F01 ORDER BY F01 INTO CURSOR C38_2                       
         IF _TALLY>0
            SELE C38_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'C38'
               REPL F02 WITH C38_2.F01
               REPL F03 WITH C38_2.F02
               REPL F05 WITH C38_2.F05
               REPL F06 WITH IIF(SEEK(C38_2.F05,'A14'),A14.F02,'')
               REPL F04 WITH C38_2.F10               
               SELE C38_2
               SKIP
            ENDDO
         ENDIF
         SELE C38_2
         USE
         SELE C38
         USE
********************************************
         C40='C40'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&客供品進貨單
         IF !USED('&C40') 
            SELE 0
            USE (C40) ALIA C40
         ELSE
            SELE C40
         ENDIF
         SET ORDER TO 1                       
         SELE F01,F02,F06,F05 FROM C40 WHERE EMPTY(F09) GROUP BY F01 ORDER BY F01 INTO CURSOR C40_2                       
         IF _TALLY>0
            SELE C40_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'C40'
               REPL F02 WITH C40_2.F01
               REPL F03 WITH C40_2.F02
               REPL F07 WITH C40_2.F06
               REPL F04 WITH C40_2.F05              
               SELE C40_2
               SKIP
            ENDDO
         ENDIF
         SELE C40_2
         USE
         SELE C40
         USE
***********************************************
         C41='C41'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&客供品退貨單
         IF !USED('&C41') 
            SELE 0
            USE (C41) ALIA C41
         ELSE
            SELE C41
         ENDIF
         SET ORDER TO 1                       
         SELE F01,F02,F05,F06 FROM C41 WHERE EMPTY(F09) GROUP BY F01 ORDER BY F01 INTO CURSOR C41_2                       
         IF _TALLY>0
            SELE C41_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'C41'
               REPL F02 WITH C41_2.F01
               REPL F03 WITH C41_2.F02
               REPL F05 WITH C41_2.F05
               REPL F06 WITH IIF(SEEK(C41_2.F05,'A14'),A14.F02,'')
               REPL F04 WITH C41_2.F06               
               SELE C41_2
               SKIP
            ENDDO
         ENDIF
         SELE C41_2
         USE
         SELE C41
         USE                  
******************************************         
         D39='D39'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&樣品進貨單
         IF !USED('&D39') 
            SELE 0
            USE (D39) ALIA D39
         ELSE
            SELE D39
         ENDIF
         SET ORDER TO 1                            
         SELE F01,F02,F05,F10 FROM D39 WHERE EMPTY(F07) GROUP BY F01 ORDER BY F01 INTO CURSOR D39_2
         IF _TALLY>0
            SELE D39_2
            GO TOP                  
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'D39'
               REPL F02 WITH D39_2.F01
               REPL F03 WITH D39_2.F02
               REPL F04 WITH D39_2.F05
               REPL F07 WITH D39_2.F10
               SELE D39_2
               SKIP
            ENDDO
         ENDIF
         SELE D39_2
         USE
         SELE D39
         USE      
******************************************         
         H02='H02'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&檢驗單
         IF !USED('&H02') 
            SELE 0
            USE (H02) ALIA H02
         ELSE
            SELE H02
         ENDIF
         SET ORDER TO 1               
         SELE F01,F02,F07,F17 FROM H02 WHERE EMPTY(F13) GROUP BY F01 ORDER BY F01 INTO CURSOR H02_2                               
         IF _TALLY>0
            SELE H02_2
            GO TOP
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'H06'
               REPL F02 WITH H02_2.F01
               REPL F03 WITH H02_2.F02
               REPL F04 WITH H02_2.F07
               REPL F07 WITH H02_2.F17                      
               SELE H02_2
               SKIP
            ENDDO
         ENDIF
         SELE H02_2
         USE
         SELE H02
         USE
**************************************************         
         H04='H04'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&驗退單
         IF !USED('&H04') 
            SELE 0
            USE (H04) ALIA H04
         ELSE
            SELE H04
         ENDIF
         SET ORDER TO 1       
         SELE F01,F02,F05,F15 FROM H04 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR H04_2
         IF _TALLY>0         
            SELE H04_2
            GO TOP                                                
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'H05'
               REPL F02 WITH H04_2.F01
               REPL F03 WITH H04_2.F02
               REPL F04 WITH H04_2.F05
               REPL F09 WITH H04_2.F15               
               SELE H04_2
               SKIP
            ENDDO
         ENDIF 
         SELE H04_2
         USE
         SELE H04
         USE  
**************************************************         
         P12='P12'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&外購進貨單
         IF !USED('&P12') 
            SELE 0
            USE (P12) ALIA P12
         ELSE
            SELE P12
         ENDIF
         SET ORDER TO 1                                
         SELE F01,F02,F07,F09 FROM P12 WHERE EMPTY(F10) GROUP BY F01 ORDER BY F01 INTO CURSOR P12_2
         IF _TALLY>0        
            SELE P12_2
            GO TOP                      
            DO WHILE !EOF()
               SELE B39
               APPEND BLANK
               REPL F01 WITH 'P12'
               REPL F02 WITH P12_2.F01
               REPL F03 WITH P12_2.F02
               REPL F04 WITH P12_2.F07            
               REPL F08 WITH P12_2.F09                              
               SELE P12_2
               SKIP
            ENDDO
         ENDIF   
         SELE P12_2
         USE
         SELE P12
         USE
         ***************
         WAIT CLEAR         
         SELE B39                           
         GO TOP         
         THISFORM.REFRESH
         THISFORM.GRID1.SETFOCUS
   ENDPROC
enddefine        
***********************************************列印程序
PROCEDURE PNT_PRC    

ENDPROC

