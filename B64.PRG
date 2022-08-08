***程式名稱:借出還入單
close all
clear 
FLG='0'
FCH=''
CHK=''
CHK_STR=''
R=0
QTY=0
TQTY=0
HD1='B024 '
HD2='BX'
HD3='借出還入單'
AREA1='B64_1'
AREA2='B64'
IF !USED('A09')
   SELE 0
   USE A09
ELSE
   SELE A09
ENDIF
SET ORDER TO 1
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1         
A24='A24'+LEFT(DTOS(DATE()),6)     &&單號暫存檔
IF !USED('&A24')
   SELE 0
   USE (A24) ALIA A24
ELSE
   SELE A24
ENDIF
SET ORDER TO 1      
B25='B25'+LEFT(DTOS(DATE()),6)    &&部門別月報表
IF !USED('&B25')
   SELE 0
   USE (B25) ALIA B25
ELSE
   SELE B25
ENDIF
SET ORDER TO 1      
B26='B26'+LEFT(DTOS(DATE()),6)   &&庫存異動記錄
IF !USED('&B26')
   SELE 0
   USE (B26) ALIA B26
ELSE
   SELE B26
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
SET ORDER TO 1      
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
SEEK sys_oper+'B64'
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET FILTER TO F04='*'
SET ORDER TO 1     
IF !USED('B01')
   SELE 0
   USE B01 INDE B01
ELSE
   SELE B01
ENDIF
SET ORDER TO 1     
IF !USED('B11')
   SELE 0
   USE B11 
ELSE
   SELE B11
ENDIF
SET ORDER TO B111 
IF !USED('C01')
   SELE 0
   USE C01 INDE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO C011
CREATE CURSOR C01_1 (F01 C(5),F02 C(4),F04 C(40))
INDEX ON F01 TAG C01_1
SET ORDER TO 1
APPEND FROM C01
IF !USED('D07')     &&進貨計畫
   SELE 0   
   USE D07
ELSE
   SELE D07
ENDIF
SET ORDER TO D074
B64='B64'+LEFT(DTOS(DATE()),6)
B641='B641'+LEFT(DTOS(DATE()),6)
IF !USED('&B64')
   SELE 0
   USE (B64) ALIA B64
ELSE
   SELE B64
ENDIF
SET ORDER TO 1      
SET RELA TO F03 INTO B01
SET RELATION TO F06 INTO C01_1 ADDI
CREATE CURSOR B64_1;
(F01 C(10),F02 C(2),F05 C(5),F10 C(5),F12 C(1),F13 C(17),F07 C(20),F90 C(17))
INDE ON F01 TAG B64_11
INDE ON F10+F01 TAG B64_12
INDE ON F05+F02 TAG B64_13
SELE F01,F02,F05,F10,F12,F13,F07,F90 FROM B64 GROUP BY F01 ORDER BY F01 INTO CURSOR B64_2
SELE B64_2
GO TOP
DO WHILE !EOF()
   SELE B64_1
   APPEND BLANK
   REPL F01 WITH B64_2.F01
   REPL F02 WITH B64_2.F02
   REPL F05 WITH B64_2.F05
   REPL F10 WITH B64_2.F10
   REPL F12 WITH B64_2.F12
   REPL F13 WITH B64_2.F13
   REPL F07 WITH B64_2.F07  
   REPL F90 WITH B64_2.F90 
   SELE B64_2
   SKIP   
ENDDO
SELE B64_2
USE   
SELE B64_1
SET ORDER TO 1
SET RELA TO F10 INTO C01
SET RELA TO F05 INTO A14 ADDI
B64form=createobject("tkB64")
B64form.show  
define class tkB64 as form
  caption='B64.借出還入單'
*!*	  autocenter=.t.
  controlbox=.F.
  BORDERSTYLE=1  
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
  NAME='TKB64'
  add object shape1 as shape1 with;
      autocenter=.t.
  ADD OBJECT LBL_REC AS LBL_REC WITH;
      AUTOCENTER =.T.
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      caption=str(reccount())        
  add object shape2 as shape2 with;
      autocenter=.t.
  add object shape3 as shape3 with;
      autocenter=.t.
  ADD OBJECT PAGEFRAME1 AS PAGEFRAME WITH;
      LEFT=INT_015*236,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*200,;
      FONTSIZE=INT_015*9,;
      PAGECOUNT=2,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.T.,;
      PAGE1.NAME='PAGE1',;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='表頭資料',;
      PAGE2.NAME='PAGE2',;
      PAGE2.VISIBLE=.T.,;
      PAGE2.CAPTION='表身詳細資料內容'   
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      autocenter=.t.,;
      WIDTH=INT_015*130
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      autocenter=.t.
  ADD OBJECT grid1 AS grid1 WITH;
      columncount=5,;            
      RECORDSOURCE='B64_1',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;      
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5'
  ADD OBJECT grid2 AS grid2 WITH;
      columncount=4,;            
      RECORDSOURCE='B64',; 
      LINKMASTER='B64_1',;
      RELATIONALEXPR='F01',;
      childorder=B641,;  
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4'      
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.              
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*173,;
      LEFT=INT_015*558          
  ADD OBJECT ANS_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*15,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;      
      FONTSIZE=INT_015*9,;
      CAPTION='\<A.過帳',;
      NAME='ANS_BOTT'                  
  PROCEDURE PAGEFRAME1.PAGE1.INIT     
     THIS.ADDOBJECT('LBLF01','LABEL')
     WITH THIS.LBLF01
          .VISIBLE=.T.
          .LEFT=INT_015*14
          .TOP=INT_015*13
          .WIDTH=INT_015*50
          .CAPTION='還入單號'          
     ENDWITH
     THIS.ADDOBJECT('LBLF10','LABEL')
     WITH THIS.LBLF10
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*38
         .WIDTH=INT_015*50
         .CAPTION='客戶編號'
     ENDWITH
     THIS.ADDOBJECT('LBLF101','LABEL')
     WITH THIS.LBLF101
         .VISIBLE=.T.
         .LEFT=INT_015*115
         .TOP=INT_015*38
         .AUTOSIZE=.T.
     ENDWITH     
     THIS.ADDOBJECT('LBLF05','LABEL')     
     WITH THIS.LBLF05
          .VISIBLE=.T.
          .LEFT=INT_015*14
          .TOP=INT_015*63
          .WIDTH=INT_015*50  
          .CAPTION='收料部門'
     ENDWITH            
     THIS.ADDOBJECT('LBLF051','LABEL')     
     WITH THIS.LBLF051
          .VISIBLE=.T.
          .LEFT=INT_015*115
          .TOP=INT_015*63
          .AUTOSIZE=.T.
     ENDWITH                      
     THIS.ADDOBJECT('LBLF07','LABEL')     
     WITH THIS.LBLF07
          .VISIBLE=.T.
          .LEFT=INT_015*14
          .TOP=INT_015*88
          .WIDTH=INT_015*50  
          .CAPTION='備註說明'
     ENDWITH                 
     THIS.ADDOBJECT('LBLF02','LABEL')
     WITH THIS.LBLF02
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*13
         .WIDTH=INT_015*50
         .CAPTION='還入日期'
     ENDWITH    
     THIS.ADDOBJECT('LBLF12','LABEL')
     WITH THIS.LBLF12
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*38
         .WIDTH=INT_015*50
         .CAPTION='過  帳  碼'
     ENDWITH    
     THIS.ADDOBJECT('LBLF121','LABEL')
     WITH THIS.LBLF121
         .VISIBLE=.T.
         .LEFT=INT_015*419
         .TOP=INT_015*38
         .AUTOSIZE=.T.
     ENDWITH    
     THIS.ADDOBJECT('LBLF13','LABEL')
     WITH THIS.LBLF13
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*63
         .WIDTH=INT_015*50
         .CAPTION='安全資料'
     ENDWITH    
     THIS.ADDOBJECT('LBLF131','LABEL')
     WITH THIS.LBLF131
         .VISIBLE=.T.
         .LEFT=INT_015*419
         .TOP=INT_015*63
         .AUTOSIZE=.T.
     ENDWITH    
     THIS.ADDOBJECT('LBLF90','LABEL')
     WITH THIS.LBLF90
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*88
         .WIDTH=INT_015*50
         .CAPTION='確認人員'
     ENDWITH    
     THIS.ADDOBJECT('LBLF901','LABEL')
     WITH THIS.LBLF901
         .VISIBLE=.T.
         .LEFT=INT_015*419
         .TOP=INT_015*88
         .AUTOSIZE=.T.
     ENDWITH        
     THIS.ADDOBJECT('TXTF01','TEXTBOX')          
     WITH THIS.TXTF01
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*9
          .WIDTH=INT_015*81
          .HEIGHT=INT_015*25
          .MAXLENGTH=10
     ENDWITH        
     THIS.ADDOBJECT('TXTF10','TXTF10')          
     WITH THIS.TXTF10
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*34
          .WIDTH=INT_015*49
          .HEIGHT=INT_015*25
          .MAXLENGTH=5
     ENDWITH             
     THIS.ADDOBJECT('TXTF05','TXTF05')          
     WITH THIS.TXTF05
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*59
          .WIDTH=INT_015*49
          .HEIGHT=INT_015*25
          .MAXLENGTH=5
     ENDWITH        
     THIS.ADDOBJECT('TXTF07','TEXTBOX')          
     WITH THIS.TXTF07
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*84
          .WIDTH=INT_015*300
          .HEIGHT=INT_015*25
          .MAXLENGTH=100
     ENDWITH             
     THIS.ADDOBJECT('TXTF02','TEXTBOX')          
     WITH THIS.TXTF02
          .VISIBLE=.T.
          .LEFT=INT_015*418
          .TOP=INT_015*9
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25
          .MAXLENGTH=2
     ENDWITH        
  ENDPROC   
  PROCEDURE PAGEFRAME1.PAGE2.INIT
	THIS.ADDOBJECT('LBLF03','LABEL')
	WITH THIS.LBLF03  
	     .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*18
	     .WIDTH=INT_015*50
         .CAPTION='料品編號'
    ENDWITH  
    THIS.ADDOBJECT('LBLF031','LABEL')
    WITH THIS.LBLF031
         .VISIBLE=.T.
         .LEFT=INT_015*215
	     .TOP=INT_015*18
	     .autosize=.t.
	ENDWITH     
	THIS.ADDOBJECT('LBLF06','LABEL')
	WITH THIS.LBLF06
	     .VISIBLE=INT_028
         .LEFT=INT_015*14
         .TOP=INT_015*43
	     .WIDTH=INT_015*50
         .CAPTION='原使用者'
    ENDWITH  
    THIS.ADDOBJECT('LBLF061','LABEL')
    WITH THIS.LBLF061
         .VISIBLE=INT_028
         .LEFT=INT_015*115
	     .TOP=INT_015*43
	     .autosize=.t.
	ENDWITH     					
    THIS.ADDOBJECT('LBLF08','Label')
    WITH THIS.LBLF08
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*68
         .WIDTH=INT_015*50
         .CAPTION='借出單號'         
    ENDWITH     

    THIS.ADDOBJECT('LBLF04','LABEL')
    WITH THIS.LBLF04
         .VISIBLE=.T.           
         .LEFT=INT_015*14
         .TOP=INT_015*118
         .WIDTH=INT_015*50
         .CAPTION='還入數量'
    ENDWITH             
    THIS.ADDOBJECT('LBLF13','LABEL')
    WITH THIS.LBLF13
        .VISIBLE=.T.           
        .LEFT=INT_015*328
        .TOP=INT_015*143
        .WIDTH=INT_015*50
        .CAPTION='安全資料'         
    ENDWITH    
    THIS.ADDOBJECT('LBLF131','LABEL')
    WITH THIS.LBLF131
        .VISIBLE=.T.           
        .LEFT=INT_015*380
        .TOP=INT_015*143
        .AUTOSIZE=.T.
    ENDWITH    
	THIS.ADDOBJECT('TXTF03','TXTF03')
	WITH THIS.TXTF03      
	     .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*14
	     .WIDTH=INT_015*261
	     .HEIGHT=INT_015*25
	     .MAXLENGTH=43
	     .NAME='TXTF03'  
	ENDWITH     
	THIS.ADDOBJECT('TXTF06','TXTF06')
	WITH THIS.TXTF06      
	     .VISIBLE=INT_028
	     .LEFT=INT_015*65
	     .TOP=INT_015*39
	     .WIDTH=INT_015*45
	     .HEIGHT=INT_015*25
	     .MAXLENGTH=5
	     .NAME='TXTF06'  
	ENDWITH     					
    THIS.ADDOBJECT('TXTF08','TXTF08')
    WITH THIS.TXTF08
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*64
         .WIDTH=INT_015*80
         .HEIGHT=INT_015*25
         .MAXLENGTH=10
         .NAME='TXTF08'
    ENDWITH         
	THIS.ADDOBJECT('TXTF04','TEXTBOX')
	WITH THIS.TXTF04
	     .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*114
	     .WIDTH=INT_015*100
	     .HEIGHT=INT_015*25
	     .INPUTMASK='999999999999'
	     .NAME='TXTF04'	    
	ENDWITH         
  ENDPROC       
  PROCEDURE PAGEFRAME1.PAGE1.activate
     R=0
          SELE B64_1  
     IF THISFORM.SHRGROUP.VISIBLE=.F.
        THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(B64_1.F12<>'2',RGB(255,0,0),'')","COLUMN")         
        THISFORM.GRID1.ENABLED=.T.   
        THISFORM.GRID1.SETFOCUS 
        THISFORM.KEY_LIST.ENABLED=.T.
        THISFORM.MTH_LIST.ENABLED=.T.
     ELSE
       THISFORM.KEY_LIST.ENABLED=.F.   
       THISFORM.MTH_LIST.ENABLED=.F.   
     ENDIF          
    IF THISFORM.GRID1.ACTIVEROW=0 .AND. thisform.SHRGROUP.ENABLED=.F.
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
        THISFORM.ANS_BOTT.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE1.SETALL('VALUE','','TEXTBOX')        
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=''        
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF121.CAPTION=''   
        THISFORM.PAGEFRAME1.PAGE1.LBLF131.CAPTION=''  
        THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=''                     
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.      
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  .AND. B64_1.F12<>'2' &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   .AND. B64_1.F12<>'2' &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)    &&判斷有無列印的權限  
        THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.
     ENDIF   
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.);
     .AND. !IIF(SEEK(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),'A23','A231'),A23.F07,.F.)
  ENDPROC
  PROCEDURE PAGEFRAME1.PAGE2.activate
       SELE B64 
     IF THISFORM.SHRGROUP.ENABLED=.F.
        THISFORM.GRID2.READONLY=.T.       
        THISFORM.GRID2.SETFOCUS   
     ENDIF 
     IF THISFORM.GRID2.ACTIVEROW=0 .and. thisform.SHRGROUP.ENABLED=.F.
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
        THISFORM.ANS_BOTT.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE2.SETALL('VALUE','','TEXTBOX')
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. B64.F12<>'2'  &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. B64.F12<>'2' &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)  &&&&判斷有無列印的權限  
        THISFORM.ANS_BOTT.ENABLED=IIF(B64.F12='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
     ENDIF   
     R=RECNO('&AREA1')
     THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
     THISFORM.GRID1.ENABLED=.F.     
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)   .AND. B64_1.F12<>'2' &&判斷有無新增的權限
     THISFORM.KEY_LIST.ENABLED=.F.     
     THISFORM.MTH_LIST.ENABLED=.F.     
  ENDPROC   
  
  PROC INIT       
       thisform.setall('height',INT_015*25,'textbox')
       thisform.setall('height',INT_015*17,'label')
       thisform.setall('FONTSIZE',INT_015*9,'textbox')
       thisform.setall('FONTSIZE',INT_015*9,'label')       
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
       THISFORM.SHRGROUP.VISIBLE=.F.
       THISFORM.SHRGROUP.ENABLED=.F.
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'HEADER')        
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID2.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)  
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='還入單號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='客戶編號'
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶簡稱'
       THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='收料部門'
       THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='部門名稱'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B64_1.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B64_1.F10'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C01.F05'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B64_1.F05'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*81
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*49
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*49
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*58
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(B64_1.F12<>'2',RGB(255,0,0),'')","COLUMN")         
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'HEADER')        
       THISFORM.grid2.COLUMN1.HEADER1.CAPTION='料品編號'
       THISFORM.grid2.COLUMN2.HEADER1.CAPTION='品名規格'
	   THISFORM.grid2.COLUMN3.HEADER1.CAPTION='借出單號'
	   THISFORM.grid2.COLUMN4.HEADER1.CAPTION='還入數量'
	   THISFORM.GRID2.LINKMASTER='B64_1'
       THISFORM.GRID2.RELATIONALEXPR='F01'
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='B64.F03'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='B64.F08'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE='B64.F04'
	   THISFORM.grid2.COLUMN1.WIDTH=INT_015*149
	   THISFORM.GRID2.COLUMN2.WIDTH=INT_015*100
       THISFORM.grid2.COLUMN3.WIDTH=INT_015*80
       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*100
       THISFORM.grid1.READONLY=.T.      
       THISFORM.grid2.READONLY=.T.            
       THISFORM.grid1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.ANS_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. B64_1.F12<>'2'  &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. B64_1.F12<>'2' &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限       
          THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.)  .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
       ENDIF          
       THISFORM.KEY_LIST.DISPLAYVALUE='依還入單號排列'      
       JK='退回單號'
       SELE B64_1
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=B64_1.F01
  ENDPROC               
  PROC KEY_LIST.INIT 
       with this 
           .additem('依還入單號排列')
           .additem('依客戶編號排列')
           .additem('依收料部門排列')
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE B64_1 
       DO CASE
          CASE THIS.DISPLAYVALUE='依還入單號排列'
               set order to 1 
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='還入單號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B64_1.F01'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*81
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B64_1.F10'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*49
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C01.F04'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='借料部門'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B64_1.F05'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*49               
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='部門名稱'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*58                                
          CASE THIS.DISPLAYVALUE='依客戶編號排列'
               set order to 2
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B64_1.F10'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*49
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C01.F04'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='還入單號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B64_1.F01'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*81                                                   
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='借料部門'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B64_1.F05'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*49  
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='部門名稱'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*58                 
          CASE THIS.DISPLAYVALUE='依收料部門排列'
               set order to 3
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='借料部門'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B64_1.F05'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*49
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='部門名稱'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*58
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='還入單號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B64_1.F01'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*81               
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶編號'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B64_1.F10'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*49                  
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='客戶簡稱'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C01.F04'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60                                                             
       ENDCASE 
       THISFORM.grid1.SETFOCUS
  ENDPROC      
  PROC MTH_LIST.INTERACTIVECHANGE
       THISFORM.GRID1.RECORDSOURCE=''
       THISFORM.GRID2.RECORDSOURCE=''
       THISFORM.GRID2.LINKMASTER=''
       THISFORM.GRID2.RELATIONALEXPR=''
       THISFORM.GRID2.childorder=''
       SELE A24
       USE
       A24='A24'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&A24')
          SELE 0
          USE (A24) ALIA A24
       ELSE
          SELE A24
       ENDIF
       SET ORDER TO 1             
       SELE B25
       USE
       B25='B25'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&B25')
          SELE 0
          USE (B25) ALIA B25
       ELSE
          SELE B25
       ENDIF
       SET ORDER TO 1                    
       SELE B26
       USE
       B26='B26'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&B26')
          SELE 0
          USE (B26) ALIA B26
       ELSE
          SELE B26
       ENDIF
       SET ORDER TO 1                           
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

       SELE A23
       SET ORDER TO 1
       SEEK DTOS(CTOD(THIS.DISPLAYVALUE+'/01'))
       SELE B64
       SET RELATION OFF INTO C01_1
       set rela off into b01
       SET RELA OFF INTO B64_1
       USE
       SELE B64_1
       set rela off into C01
       set rela off into a14
       USE
       B64='B64'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       B641='B641'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&B64')
          SELE 0
          USE (B64) ALIA B64
       ELSE
          SELE B64
       ENDIF
       SET ORDER TO 1      
       SET RELA TO F03 INTO B01
       SET RELATION TO F06 INTO C01_1 ADDITIVE 
       CREATE CURSOR B64_1;
       (F01 C(10),F02 C(2),F05 C(5),F10 C(5),F12 C(1),F13 C(17),F07 C(20),F90 C(17))
       INDE ON F01 TAG B64_11
       INDE ON F10+F01 TAG B64_12
       INDE ON F05+F02 TAG B64_13
       SELE F01,F02,F05,F10,F12,F13,F07,F90 FROM B64 GROUP BY F01 ORDER BY F01 INTO CURSOR B64_2

       SELE B64_2
       GO TOP
       DO WHILE !EOF()
          SELE B64_1
          APPEND BLANK
          REPL F01 WITH B64_2.F01
          REPL F02 WITH B64_2.F02
          REPL F05 WITH B64_2.F05
          REPL F10 WITH B64_2.F10
          REPL F12 WITH B64_2.F12
          REPL F13 WITH B64_2.F13
          REPL F07 WITH B64_2.F07  
          REPL F90 WITH B64_2.F90 
          SELE B64_2
          SKIP   
       ENDDO
       SELE B64_2
       USE   
       SELE B64_1
       SET ORDER TO 1
       SET RELA TO F10 INTO C01
       SET RELA TO F05 INTO A14 ADDI
       THISFORM.GRID1.RECORDSOURCE='B64_1'
       THISFORM.GRID2.RECORDSOURCE='B64'
	   THISFORM.GRID2.LINKMASTER='B64_1'
       THISFORM.GRID2.RELATIONALEXPR='F01'      
       THISFORM.GRID2.CHILDORDER=B641  
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='B64.F03'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='B64.F08'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE='B64.F04'            
       THISFORM.KEY_LIST.INTERACTIVECHANGE
	   THISFORM.PAGEFRAME1.ACTIVEPAGE=1
	   THISFORM.REFRESH
	   thisform.grid1.setfocus        
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.ANS_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. B64_1.F12<>'2'  &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. B64_1.F12<>'2' &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
          THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.)  .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED               
       ENDIF          
*!*	       THISFORM.KEY_LIST.DISPLAYVALUE='依出貨退回單號排列'      
*!*	       JK='訂單編號'
*!*	       SELE B64_1
*!*	       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=B64_1.F01	   
*       THISFORM.REFRESH       
  ENDPROC
  PROC grid1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=B64_1.F01
       THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=B64_1.F10
       THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=B64_1.F05
       THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=C01.F04
       THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=A14.F02
       THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE=B64_1.F07       
       THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=B64_1.F02
       THISFORM.PAGEFRAME1.PAGE1.LBLF121.CAPTION=B64_1.F12       
       THISFORM.PAGEFRAME1.PAGE1.LBLF131.CAPTION=B64_1.F13
       THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=B64_1.F90
       THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. B64_1.F12<>'2'  &&判斷有無修改的權限
       THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. B64_1.F12<>'2' &&判斷有無刪除的權限
       THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED         
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION  
       IF THIS.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
       ENDIF   
       CHK=''     
*       thisform.refresh
  ENDPROC     
  PROC grid2.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=B64.F03
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=B01.F02
       THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=B64.F04
       THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE=B64.F08 
       THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=B64.F06       
       THISFORM.PAGEFRAME1.PAGE2.LBLF131.CAPTION=F13       
       THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=C01_1.F04            
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       if thisform.pageframe1.activepage=1
          thisform. CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ENDIF   
*       thisform.refresh
  ENDPROC 
  PROCEDURE ORPGROUP.NEW_BOTT.CLICK   
    THISFORM.KEY_LIST.ENABLED=.F. 
    THISFORM.MTH_LIST.ENABLED=.F. 
    THISFORM.ANS_BOTT.ENABLED=.F.
    IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1       
       THISFORM.GRID2.RECORDSOURCE=''
       THISFORM.GRID2.CHILDORDER=''
       THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')     &&處理表頭新增
       THISFORM.GRID1.ENABLED=.F.   
       THISFORM.GRID2.ENABLED=.F.   
          HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
          DO ODR_CRT 
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=PO_NO
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.REFRESH
          thisform.PAGEFRAME1.PAGE1.txtf01.readonly=.t.
          THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE1.TXTF05.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=.F.        
        THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=RIGHT(DTOS(DATE()),2)
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF121.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF131.CAPTION=''                
        THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION='' 
     ELSE
        THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.              &&處理表身新增
        THISFORM.GRID2.ENABLED=.F.   
        THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX')     
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''  
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.READONLY=.F.        
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=''          
        THISFORM.PAGEFRAME1.PAGE2.TXTF08.READONLY=.F.  
        THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE=''     
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=0
        THISFORM.PAGEFRAME1.PAGE2.lblF031.caption=''                  
        THISFORM.PAGEFRAME1.PAGE2.LBLF131.CAPTION=''               
     ENDIF   

  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
     THISFORM.KEY_LIST.ENABLED=.F. 
     THISFORM.MTH_LIST.ENABLED=.F. 
     THISFORM.ANS_BOTT.ENABLED=.F.
     THISFORM.GRID1.ENABLED=.F.   
     THISFORM.GRID2.ENABLED=.F.   
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
        SELE B64_1
        PO_NO=B64_1.F01
        SELE RECNO() FROM B64 WHERE B64.F01=PO_NO AND F12<>'2' INTO ARRAY BK1
        JJ1=''                
        IF _TALLY>0                   
           FOR I= 1 TO ALEN(BK1)
               JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
           ENDFOR  
           KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')
           RLOCK(KK1,'B64')
           LL1=RLOCK(KK1,'B64')
        ELSE 
           =MESSAGEBOX('此單據已由別的工作站過帳中無法修改!',0+64,'提示訊息視窗')
           SELE B64_1           
           UNLOCK ALL
           REPL F12 WITH '2'
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   
        ENDIF   
        IF LL1 =.T.
           THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')     &&處理表頭修改
           THISFORM.PAGEFRAME1.PAGE1.TXTF01.READONLY=.T.
           THISFORM.PAGEFRAME1.PAGE1.TXTF05.READONLY=.F.
           THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
           THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
        ELSE 
           DO file_impact
           UNLOCK ALL
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
        endif   
     ELSE                   
        thisform.pageframe1.page1.enabled=.f.      
        SELE B64                                                      &&處理表身修改
        SELE RECNO() FROM D07 WHERE F01=B64.F08 AND F02=B64.F03 INTO ARRAY BK1
        JJ1=''                
        IF _TALLY>0                   
           FOR I= 1 TO ALEN(BK1)
               JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
           ENDFOR  
           KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')
           RLOCK(KK1,'D07')
           LL1=RLOCK(KK1,'D07')
        ELSE
           LL1=.T.   
        ENDIF         
        IF RLOCK('B64') .AND. LL1   
           THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX')         
           IF INT_028
              THISFORM.PAGEFRAME1.PAGE2.TXTF06.READONLY=.F.                           
              THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
           ELSE   
              THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS
           ENDIF   
        ELSE  
           DO file_impact 
           unlock all
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           
       endif        
     ENDIF   
     RELEASE ALL LIKE BK*
  endproc    
  PROC ORPGROUP.DEL_BOTT.CLICK                    &&刪除資料程序
       IF messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6	        
          if ALIA()='B64_1'           &&刪除整張訂單
             SELE B64_1             
                PO_NO=B64_1.F01
                SELE RECNO() FROM B64 WHERE B64.F01=PO_NO  AND B64.F12<>'2' INTO ARRAY BK1              
                JJ1=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
                   ENDFOR    
                   KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
                   RLOCK(KK1,'B64')
                   LL1=RLOCK(KK1,'B64')
                ELSE
                   LL1=.T.   
                ENDIF    
                SELE RECNO() FROM D07 WHERE F01= ANY (SELE F08 FROM B64 WHERE F01=PO_NO) INTO ARRAY BK2
                JJ2=''                
                IF _TALLY>0                   
                    FOR I= 1 TO ALEN(BK2)
                        JJ2=JJ2+ALLTRIM(STR(BK2(I)))+','
                    ENDFOR  
                    KK2=STUFF(JJ2,RAT(',',JJ2,1),1,'')
                    RLOCK(KK2,'D07')
                    LL2=RLOCK(KK2,'D07')
                ELSE
                   LL2=.T.   
                ENDIF                                  
                IF LL1 .AND. LL2=.T.
                   IF LEN(ALLTRIM(JJ1))>0 
                      SELE B64                   
                      FOR I= 1 TO ALEN(BK1)
                          GO BK1(I) 
                          SELE D07
                          SET ORDER TO D074                 
                          SEEK B64.F08+B64.F03
                            IF FOUND()                               
                               REPL F06 WITH F06+B64.F04*(-1)                          
                            ENDIF                                  
                            SELE B64
                       ENDFOR                                   
       
                   ELSE
                      =MESSAGEBOX('此張單據已由別的工作站過帳中,無法刪除!',0+64,'提示訊息視窗')
                      SELE B64_1
                      REPL F12 WITH '2'
                      UNLOCK ALL
                      RETURN   
                   ENDIF    
                   DELE FROM B64 WHERE B64.F01=PO_NO
                   SELE B39                     &&刪除未過帳單據紀錄
                   SEEK A02.F03+B64_1.F01                   
                   IF FOUND()
                      DELE
                   ENDIF   
                   SELE B64_1
                   DELE
                   UNLOCK ALL
                 IF  SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                     HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)+PO_NO
                     DO ODR_RJT
                  ENDIF  
                ELSE
                   DO file_impact                  
                ENDIF             
             THISFORM.GRID1.SETFOCUS
             IF THISFORM.GRID1.ACTIVEROW=0
                THISFORM.CMDGROUP.ENABLED=.F.
                THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
                THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.               
                THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
                THISFORM.ANS_BOTT.ENABLED=.F.
             ELSE      
                THISFORM.CMDGROUP.ENABLED=.T.
                THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  .AND. B64_1.F12<>'2' &&判斷有無修改的權限
                THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   .AND. B64_1.F12<>'2' &&判斷有無刪除的權限
                THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限
                THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED                
             ENDIF     
          ELSE                              &&刪除表身單筆資料
              SELE B64                  
              SELE RECNO() FROM D07 WHERE F01=B64.F08 AND F02=B64.F03 INTO ARRAY BK1
              JJ1=''                
                IF _TALLY>0                   
                    FOR I= 1 TO ALEN(BK1)
                        JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
                    ENDFOR  
                    KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')
                    RLOCK(KK1,'D07')
                    LL1=RLOCK(KK1,'D07')
                ELSE
                   LL1=.T.   
                ENDIF                                                  
                IF RLOCK('B64') .AND. LL1=.T.
                   IF LEN(ALLTRIM(JJ1))>0
                      SELE D07
                      SET ORDER TO D074
                      SEEK B64.F08+B64.F03
                      IF FOUND()
                         REPL F06 WITH F06+B64.F04*(-1)                     
                      ENDIF 
                   ENDIF    
                   SELE B64
                   DELE
                   SKIP -1
                   IF BOF()
                      GO TOP
                   ENDIF   
                ELSE
                  DO file_impact                  
                ENDIF     
              THISFORM.GRID2.SETFOCUS
              IF THISFORM.GRID2.ACTIVEROW=0         &&如果單身被刪空
                 SELE B39                           &&刪除未過帳單據紀錄       
                 SEEK A02.F03+B64_1.F01
                 IF FOUND()
                    DELE
                 ENDIF                    
                 SELE B64_1                        
                 PO_NO=B64_1.F01                    
                 DELE                               &&連表頭也要刪除
                 IF SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                     HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)+PO_NO
                     DO ODR_RJT
                 ENDIF
                 SELE B64_1
                 SKIP -1
                 IF BOF()
                    GO TOP
                 ENDIF   
                 THISFORM.PAGEFRAME1.ACTIVEPAGE=1 
                 THISFORM.REFRESH          
              ENDIF
          ENDIF
       ENDIF
       RELEASE ALL LIKE BK*   
  ENDPROC  
  PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK               &&新增確認
    IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
       THISFORM.GRID2.RECORDSOURCE='B64'
       THISFORM.GRID2.CHILDORDER=B641
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='B64.F03'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='B64.F08'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE='B64.F04'       
       SELE B64_1 
       COD=SYS(21)
       SET ORDER TO 1
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE,'B64')=.T.  &&如果單號重覆
          SELE MAX(F01) FROM B64 INTO ARRAY GB NOWAIT
          HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)+THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE   
          DO ODR_RPT 
	        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=PO_NO
	        THISFORM.PAGEFRAME1.PAGE1.TXTF01.REFRESH          
       ENDIF 
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE,'C01')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE)
          =MESSAGEBOX('客戶基本主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE,'A14')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE)
          =MESSAGEBOX('工作部門主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF       
       IF EMPTY(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/'+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE))          
          =MESSAGEBOX('日期輸入錯誤!',0+64,'提示訊息視窗') 
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
          SET ORDER TO &COD          
          RETURN
       ENDIF   
       SELE RECNO() FROM D07 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE AND F04>F06 INTO ARRAY RC1
       IF _TALLY=0
          =MESSAGEBOX('未欠此客戶任何物料!',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
          SET ORDER TO &COD
          RETURN
       ENDIF   
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE,'B64')=.T.
         =MESSAGEBOX('資料重複!',0+48,'提示訊息視窗')
         THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
         RETURN
       ENDIF         
       SELE B64_1          
       APPEND BLANK
       REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
       REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
       REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE
       REPLACE F10 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE
       REPLACE F13 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')

       SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       SET ORDER TO &COD
       SELE B39
       APPEND BLANK
       REPL F01 WITH A02.F03
       REPL F02 WITH B64_1.F01
       REPL F03 WITH B64_1.F02
       REPL F04 WITH B64_1.F10
       REPL F05 WITH B64_1.F05
       REPL F06 WITH IIF(SEEK(B64_1.F05,'A14'),A14.F02,'')
       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.  
       THISFORM.PAGEFRAME1.ACTIVEPAGE=2
       THISFORM.ORPGROUP.NEW_BOTT.CLICK
       THISFORM.REFRESH
    ELSE                           &&表身新增確認
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'B01')=.F.
          =MESSAGEBOX('物料基本主檔中無此料號!',0+48,'提示訊息視窗')          
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF   
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07')=.F.          
          =MESSAGEBOX('進貨計劃無此筆資料!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF   
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE>IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07'),D07.F04-D07.F06,0)
          =MESSAGEBOX('還入數量不得大於'+ALLTRIM(STR(D07.F04-D07.F06)),0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS()
          RETURN
       ENDIF   
       IF INT_028 AND !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE) AND !SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE,'C01_1')
           =MESSAGEBOX('無此客戶編號,否則使用對象請保持空白!',0+48,'提示訊息視窗')
           THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
           RETURN
       ENDIF                                                      
       SELE B64
       APPEND BLANK
       LOCK()     
       REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
       REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
       REPLACE F04 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE
       REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
       REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE
       REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE              
       REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE       
       REPLACE F10 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE
       REPLACE F13 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
       UNLOCK ALL         
       seek thisform.pageframe1.page1.txtf01.value+thisform.pageframe1.page2.txtf03.value+THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE
       thisform.grid2.enabled=.t.
       thisform.grid2.setfocus         
       SELE D07                 
       SET ORDER TO D074                                                          &&寫入進貨計劃
       SEEK B64.F08+B64.F03
       IF FOUND()          
          REPL F06 WITH F06+B64.F04              
       ENDIF                                                       
       THISFORM.ORPGROUP.NEW_BOTT.CLICK      
     ENDIF  
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK                          &&修改確認
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE,'A14')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE)
          =MESSAGEBOX('工作部門主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF            
       IF EMPTY(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/'+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE))          
          =MESSAGEBOX('日期輸入錯誤!',0+64,'提示訊息視窗') 
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
          RETURN
        ENDIF                   

           SELE B64
           SEEK B64_1.F01
           IF FOUND()
              DO WHILE F01=B64_1.F01
                 REPL F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
                 REPL F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
                 REPL F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE                                
                 REPL F13 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')  
                 SKIP
              ENDDO    
           ENDIF        
        SELE B64_1
           REPL F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
           REPL F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
           REPL F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE                                                
           REPL F13 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
           UNLOCK ALL    
     ELSE                                                        &&表身修改確認     
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'B01')=.F.
          =MESSAGEBOX('物料基本主檔中無此料號!',0+48,'提示訊息視窗')          
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF   
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07')=.F.          
          =MESSAGEBOX('進貨計劃無此筆資料!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF   
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE>IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07'),D07.F04-D07.F06+B64.F04,0)
          =MESSAGEBOX('還入數量不得大於'+ALLTRIM(STR(D07.F04-D07.F06+B64.F04)),0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS()
          RETURN
       ENDIF          
       IF INT_028 AND !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE) AND !SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE,'C01_1')
           =MESSAGEBOX('無此客戶編號,否則使用對象請保持空白!',0+48,'提示訊息視窗')
           THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
           RETURN
       ENDIF                                                      
       SELE D07
       SET ORDER TO D074
       SEEK B64.F08+B64.F03
       IF FOUND()
          REPLACE F06 WITH F06+B64.F04*(-1)
       ENDIF
       SELE B64                    
       REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
       REPLACE F04 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE
       REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE       
       REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE
       REPLACE F13 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                      
       SELE D07
       SEEK B64.F08+B64.F03
       IF FOUND()
          REPL F06 WITH F06+B64.F04
       ENDIF   
    ENDIF   
    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK    
  ENDPROC   
  procedure SHRGROUP.ABANDON_BOTT.CLICK
      THISFORM.SHRGROUP.ENABLED=.F.
      THISFORM.SHRGROUP.VISIBLE=.F.      
      THISFORM.ORPGROUP.ENABLED=.T.
      THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.CMDGROUP.ENABLED=.T.
      THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE1.TXTF05.READONLY=.T.
      THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=.T.
      THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.T.   
      THISFORM.PAGEFRAME1.PAGE2.TXTF04.READONLY=.T.
      THISFORM.PAGEFRAME1.PAGE2.TXTF06.READONLY=.T.      
      THISFORM.PAGEFRAME1.PAGE2.TXTF08.READONLY=.T.
      THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.
      THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1    
         THISFORM.grid1.ENABLED=.T.
         THISFORM.grid2.ENABLED=.T.        
         THISFORM.GRID1.SETFOCUS
         THISFORM.KEY_LIST.ENABLED=.T.
         THISFORM.MTH_LIST.ENABLED=.T.      
      ELSE
         THISFORM.grid1.ENABLED=.F.
         THISFORM.grid2.ENABLED=.T.        
         THISFORM.GRID2.SETFOCUS
         IF THISFORM.GRID2.ACTIVEROW=0 
            HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)+PO_NO
            DO ODR_RJT
            SELE B39
            SEEK A02.F03+B64_1.F01
            SET ORDER TO 1
            IF FOUND()
               DELE            
            ENDIF   
            SELE B64_1
            DELE
            THISFORM.PAGEFRAME1.ACTIVEPAGE=1
            THISFORM.REFRESH  
         ELSE
            SELE B39
            SET ORDER TO 1
            SEEK A02.F03+B64_1.F01 
            IF FOUND()
               REPL F08 WITH B64.F08
            ENDIF   
         ENDIF   
      ENDIF       
      UNLOCK ALL
      THISFORM.ANS_BOTT.ENABLED=IIF(B64_1.F12='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED              
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
         IF THISFORM.GRID2.RECORDSOURCE<>'B64'
            THISFORM.GRID2.RECORDSOURCE='B64'
            THISFORM.GRID2.CHILDORDER=B641
  	        THISFORM.grid2.COLUMN1.CONTROLSOURCE='B64.F03'
	        THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'
	        THISFORM.grid2.COLUMN3.CONTROLSOURCE='B64.F08'
	        THISFORM.grid2.COLUMN4.CONTROLSOURCE='B64.F04'         
	     ENDIF   
            HD=HD1+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)+PO_NO
            DO ODR_RJT
         SELE B64_1
         DO CASE 
            CASE THISFORM.KEY_LIST.DISPLAYVALUE='依還入單號排列'
                 SET ORDER TO 1
            CASE THISFORM.KEY_LIST.DISPLAYVALUE='依客戶編號排列'
                 SET ORDER TO 2
            CASE THISFORM.KEY_LIST.DISPLAYVALUE='依收料部門排列'
                 SET ORDER TO 3
         ENDCASE   
       ENDIF  
       THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC     
  PROC ANS_BOTT.CLICK     &&過帳程序  
       &&D07(進貨計劃)->F04,F06
       &&B11(庫存明細檔)->庫存數量(F04),最後使用日期(F08)
       &&B39(未過帳單據)->整筆刪除
       if messagebox('是否確認資料無誤可以過帳?',4+32+256,'請確認') = 6	        
    
          SELE RECNO() FROM B64 WHERE B64.F01=B64_1.F01  AND F12<>'2' INTO ARRAY BK1              
               JJ1=''                
               IF _TALLY>0   
                  FOR I= 1 TO ALEN(BK1)
                      JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
                  ENDFOR    
                  KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
                  RLOCK(KK1,'B64')
                  LL1=RLOCK(KK1,'B64')
               ELSE
                  LL1=.T.   
               ENDIF                   
               SELE RECNO() FROM D07 WHERE F01= ANY (SELE F08 FROM B64 WHERE F01=B64_1.F01) INTO ARRAY BK2
               JJ2=''                
               IF _TALLY>0                   
                  FOR I= 1 TO ALEN(BK2)
                      JJ2=JJ2+ALLTRIM(STR(BK2(I)))+','
                   ENDFOR  
                   KK2=STUFF(JJ2,RAT(',',JJ2,1),1,'')
                   RLOCK(KK2,'D07')
                   LL2=RLOCK(KK2,'D07')
               ELSE
                  LL2=.T.   
               ENDIF                  
               IF LL1 .AND. LL2 =.T. 
                  IF LEN(ALLTRIM(JJ1))>0 
                     SELE B64
                     FOR I= 1 TO ALEN(BK1)                      
                         GO BK1(I) 
                         SELE B01                                            &&物料基本主檔
                         SEEK B64.F03
                         IF FOUND()
                            IF !EMPTY(F98)                                   &&如果不是虛擬料號才做庫存運算
                               SELE B11                                            &&部門庫存明細檔 
                               SEEK B64.F05+B64.F03+IIF(INT_028,B64.F06,SPACE(5))                         &&還入收料部門
                               IF FOUND()
                                  REPL F04 WITH F04+(B64.F04)
                                  IF F04=0
                                     DELE
                                  ELSE   
                                     IF F05<CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                        REPL F05 WITH CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                     ENDIF
                                  ENDIF   
                               ELSE
                                  APPEND BLANK
                                  REPL F01 WITH B64.F05
                                  REPL F03 WITH B64.F03
                                  REPL F04 WITH (B64.F04)                                                                   
                                  REPL F05 WITH CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                  IF INT_028
                                     REPLACE F08 WITH B64.F06
                                  ENDIF   
                               ENDIF   
                               SEEK LEFT(INT_190,5)+B64.F03+IIF(INT_028,B64.F06,SPACE(5))                          &&借出入倉還料
                               IF FOUND()
                                  REPL F04 WITH F04+(B64.F04)*(-1)
                                  IF F04=0
                                     DELE
                                  ELSE   
                                     IF F05<CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                        REPL F05 WITH CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                     ENDIF
                                  ENDIF   
                               ELSE
                                  APPEND BLANK
                                  REPL F01 WITH LEFT(INT_190,5)
                                  REPL F03 WITH B64.F03
                                  REPL F04 WITH (B64.F04)*(-1)                                                                   
                                  REPL F05 WITH CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/'+B64.F02))
                                  IF INT_028
                                     REPLACE F08 WITH B64.F06
                                  ENDIF   
                               ENDIF                                  
                               SELE B25                                             &&分部門庫存月報表 
                               SEEK B64.F05+B64.F03                                 &&還入收料部門
                               IF FOUND()
                                  REPL F21 WITH F21+B64.F04
                                  REPL F15 WITH F15+B64.F04
                               ELSE
                                  APPEND BLANK
                                  REPL F01 WITH B64.F05
                                  REPL F02 WITH B64.F03
                                  REPL F21 WITH B64.F04
                                  REPL F15 WITH B64.F04                                           
                               ENDIF   
                               SEEK LEFT(INT_190,5)+B64.F03                                 &&借出入倉還貨
                               IF FOUND()
                                  REPL F19 WITH F19+B64.F04
                                  REPL F15 WITH F15+(B64.F04)*(-1)
                               ELSE
                                  APPEND BLANK
                                  REPL F01 WITH LEFT(INT_190,5)
                                  REPL F02 WITH B64.F03
                                  REPL F19 WITH B64.F04
                                  REPL F15 WITH B64.F04*(-1)                                           
                               ENDIF                                  
                               SELE B26                                            &&庫存異動紀錄
                               APPEND BLANK                                           &&還入收料部門
                               REPL F01 WITH B64.F03
                               REPL F02 WITH B64.F05
                               REPL F03 WITH B64.F02
                               REPL F04 WITH (B64.F04)
                               REPL F06 WITH 'X'
                               REPL F07 WITH B64.F01
                               REPL F08 WITH '還入 '+B64.F08+' '+ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION)
                               APPEND BLANK                                           &&借出入倉還料
                               REPL F01 WITH B64.F03
                               REPL F02 WITH LEFT(INT_190,5)
                               REPL F03 WITH B64.F02
                               REPL F04 WITH (B64.F04)*(-1)
                               REPL F06 WITH 'X'
                               REPL F07 WITH B64.F01
                               REPL F08 WITH '還入 '+B64.F08+' '+ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION)                               
                            ENDIF
                         ENDIF   
                             SELE D07                                            &&進貨計劃
                             SET ORDER TO D074                 
                             SEEK B64.F08+B64.F03
                             IF FOUND()                   
                                SELE D07                                                            
                                REPL F04 WITH F04+B64.F04*(-1) 
                                REPL F06 WITH F06+B64.F04*(-1)                                                             
                                REPL F08 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                             
                                IF F04=0
                                   DELE
                                ENDIF   
                             ENDIF   
                             SELE B64
                             REPL F12 WITH '2'
                             REPL F90 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                            
                     ENDFOR   
                     SELE B39                    &&未過帳庫存單據查詢檔
                     SEEK 'B64'+B64_1.F01
                     IF FOUND()
                        DELE
                     ENDIF                                             
                     SELE B64_1
                     REPL F12 WITH '2'
                     REPL F90 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                  ELSE
                     =MESSAGEBOX('此單據已由別的工作站過帳完成',0+64,'提示訊息視窗')
                     SELE B64_1
                     REPL F12 WITH '2'
                  ENDIF           
                  THIS.ENABLED=.F.
                  THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
                  THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.             
                  UNLOCK ALL
                  IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
                     THISFORM.GRID1.SETFOCUS
                  ELSE
                      THISFORM.GRID2.SETFOCUS
                      THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
                  ENDIF
                  THISFORM.REFRESH
               ELSE
                  DO file_impact
                  unlock all
                 return
                ENDIF   
           unlock all
           RELEASE ALL LIKE BK*
       ENDIF    
  ENDPROC
enddefine    

***********************************************列印程序
PROCEDURE PNT_PRC  

     LK=B64_1.F01     
     HK=LEFT(DTOC(A23.F01),5)
     SELECT B64.F01,;
            B64.F02,;
            B64.F03,;
            B64.F04,;
            B64.F05,;
            B64.F07,;
            B64.F08,;
            B64.F10,; 
            C01.F04,;           
            B01.F02,;
            A14.F02;
     FROM B64,C01,B01,A14 WHERE B64.F01=LK AND C01.F01=B64.F10 AND B01.F01=B64.F03 AND A14.F01=B64.F05 ORDER BY B64.F03 NOWAIT     
     IF _tally >0
*       REPORT FORM B64 PREVIEW
         report form ALLTRIM(INT_116)+'B64' to print prompt 
     ENDIF  
     CLEAR
ENDPROC
************************************************特殊輸入欄位的設定----料號
DEFINE CLASS TXTF03 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       DO CASE
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OPT; 
               ORDER BY F02,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE; 
               AND LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND (D07.F04-D07.F06)>0
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OPT; 
               ORDER BY F02,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE; 
               AND LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND (D07.F04-D07.F06)>0 AND F01=THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE              
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OPT ORDER BY F02,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE;
               AND (D07.F04-D07.F06)>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OPT ORDER BY F02,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE;
               AND (D07.F04-D07.F06)>0 AND F01=THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE              
       ENDCASE 
       IF _TALLY>0  
          OPT_SEEK=createobject("OPT_SEEK")  
          OPT_SEEK.SHOW    
          IF LEN(ALLTRIM(FLG))>1
             THIS.VALUE=FLG
             THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE=CHK_STR
             THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=QTY
             THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=LEFT(FTR_STR,5)        
             THIS.REFRESH          
             FTR_STR=''
             FLG='0'           
             CHK_STR=''
             QTY=0    
          ENDIF                                        
       ENDIF    
    ELSE
       IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE)
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THIS.VALUE,'D07'),D07.F04-D07.F06,0)       
          THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE+THIS.VALUE,'D07'),LEFT(D07.F13,5),'')      
       ENDIF          
    ENDIF  
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=IIF(SEEK(THIS.VALUE,'B01'),B01.F02,'')     
       THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE,'C01_1'),C01_1.F04,'')                                          
  ENDPROC

ENDDEFINE
***************************************************特殊欄位-----------------借出單號
DEFINE CLASS TXTF08 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=10
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       DO CASE
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F02,F03,F04,F06,F13 FROM D07 INTO CURSOR OQT; 
               ORDER BY F01,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE; 
               AND LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND (D07.F04-D07.F06)>0
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F02,F03,F04,F06,F13 FROM D07 INTO CURSOR OQT; 
               ORDER BY F01,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE; 
               AND LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND (D07.F04-D07.F06)>0 AND F02=THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE              
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OQT ORDER BY F01,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE;
               AND (D07.F04-D07.F06)>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F02,F01,F03,F04,F06,F13 FROM D07 INTO CURSOR OQT ORDER BY F01,F03 WHERE F05=THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE;
               AND (D07.F04-D07.F06)>0 AND F02=THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE              
       ENDCASE 
       IF _TALLY>0  
          OQT_SEEK=createobject("OQT_SEEK")  
          OQT_SEEK.SHOW    
          IF LEN(ALLTRIM(FLG))>1
             THIS.VALUE=FLG
             THIS.REFRESH                    
             THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=CHK_STR
             THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=QTY
             THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=LEFT(FTR_STR,5)
             THIS.REFRESH
             FLG='0'         
             FTR_STR=''  
             CHK_STR=''
             QTY=0    
          ENDIF   
       ENDIF    
    ELSE 
       IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE)
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=IIF(SEEK(THIS.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07'),D07.F04-D07.F06,0)        
          THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=IIF(SEEK(THIS.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'D07'),LEFT(D07.F13,5),'')         
       ENDIF   
    ENDIF  
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=IIF(SEEK(THISform.pageframe1.page2.txtf03.VALUE,'B01'),B01.F02,'')      
       THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=IIF(SEEK(THISform.pageframe1.page2.txtf06.VALUE,'C01_1'),C01_1.F04,'')             
  ENDPROC     
ENDDEFINE
***************************************************特殊欄位------------------收料部門
DEFINE CLASS TXTF05 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5 
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) AND LEN(ALLTRIM(F01))=4 AND F04='*'
       ELSE
          SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEN(ALLTRIM(F01))=4 AND F04='*'
       ENDIF   
       IF _TALLY>0
          DEPART_SEEK=createobject("DEPART_SEEK")  
          DEPART_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'
       ENDIF   
    ENDIF   
       THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')  
  ENDPROC
ENDDEFINE  
***************************************************特殊輸入欄位設定-----------客戶
DEFINE CLASS TXTF10 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0 
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          if empty(a02.f12)
             SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) and f18=sys_oper
          else
             SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
          endif
       ELSE
          SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 
       ENDIF
       IF _TALLY>0          
          VDR_SEEK=createobject("VDR_SEEK")  
          VDR_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'        
       ENDIF   
    ELSE
       IF LEN(ALLTRIM(THIS.VALUE))=4
          SELECT F02,F01,F04 FROM C01 INTO CURSOR VDR ORDER BY F02 WHERE F02=ALLTRIM(THIS.VALUE)
          DO CASE
             CASE _TALLY=1
                  THIS.VALUE=VDR.F01
                  THIS.REFRESH
             CASE _TALLY>1
                  VDR_SEEK=createobject("VDR_SEEK")  
                  VDR_SEEK.SHOW    
                  THIS.VALUE=FLG
                  THIS.REFRESH
                  FLG='0'        
          ENDCASE      
       ENDIF    
    ENDIF      
       THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=IIF(SEEK(THIS.VALUE,'C01'),C01.F04,'')  
  ENDPROC
ENDDEFINE
*********************************************************************進貨計劃----依料號排列搜尋
define class OPT_SEEK as form
  autocenter=.t.
  caption='進貨計劃'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*320
  WIDTH=INT_015*522
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='OPT_SEEK'
  add object cmdgroup as cmdgroup with;      
      TOP=INT_015*290,;
      LEFT=INT_015*10
  add object GRID1 as GRID with;          
      AUTOCENTER=.T.,;
      ALLOWCELLSELECTION=.F.,;
      WIDTH=INT_015*522,;      
      HEIGHT=INT_015*293,;      
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*19,;
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=4,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;      
      COLUMN4.NAME='COLUMN4'            
  add object cmnd1 as commandbutton with;
      LEFT=INT_015*183,;
      TOP=INT_015*295,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<S.選取',;
      TOOLTIPTEXT='選取此筆資料!快速鍵->ALT+S'
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      LEFT=INT_015*225,;
      TOP=INT_015*295,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<C.放棄',;
      TOOLTIPTEXT='取消此作業畫面!快速鍵->ALT+C'
      name='cmnd2'
      procedure init 
         AREA1='OPT' 
        SELE OPT
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')         
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='OPT'
         THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.         
         thisform.setall('readonly',.t.,'textbox')
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料        號'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*261
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='OPT.F02'
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='借出單號'
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='OPT.F01'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*83
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預計歸還'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='OPT.F03'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*74
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='還入數量'         
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='OPT.F04-OPT.F06'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*74         
         THISFORM.GRID1.SETFOCUS
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
           LPARAMETERS XColIndex
           FCH='GRID1'
      ENDPROC 
      procedure cmnd1.click            
           FLG=OPT.F02         
           CHK_STR=OPT.F01
           QTY=OPT.F04-OPT.F06      
           FTR_STR=OPT.F13
          AREA1='B64_1'
          THISFORM.RELEASE                                            
      endproc
      procedure cmnd2.click
          FLG='0'
          CHK_STR=''
          QTY=0
          FTR_STR=''
          AREA1='B64_1'
         thisform.release
      endproc    
enddefine            
***************************************進貨計劃-------依借出單號排列搜尋
define class OQT_SEEK as form
  autocenter=.t.
  caption='進貨計劃'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*320
  WIDTH=INT_015*522
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='OQT_SEEK'
  add object cmdgroup as cmdgroup with;      
      TOP=INT_015*290,;
      LEFT=INT_015*10
  add object GRID1 as GRID with;          
      AUTOCENTER=.T.,;
      ALLOWCELLSELECTION=.F.,;
      WIDTH=INT_015*522,;      
      HEIGHT=INT_015*293,;      
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*19,;
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=4,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;      
      COLUMN4.NAME='COLUMN4'            
  add object cmnd1 as commandbutton with;
      LEFT=INT_015*183,;
      TOP=INT_015*295,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<S.選取',;
      TOOLTIPTEXT='選取此筆資料!快速鍵->ALT+S'
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      LEFT=INT_015*225,;
      TOP=INT_015*295,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<C.放棄',;
      TOOLTIPTEXT='取消此作業畫面!快速鍵->ALT+C'
      name='cmnd2'
      procedure init 
         AREA1='OQT' 
         SELE OQT
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')        
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='OQT'
         THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.         
         thisform.setall('readonly',.t.,'textbox')
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='借出單號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='OQT.F01'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*83
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*261
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='OQT.F02'         
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預計歸還'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='OQT.F03'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*74
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='還入數量'         
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='OQT.F04-OQT.F06'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*74         
         THISFORM.GRID1.SETFOCUS
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
           LPARAMETERS XColIndex
           FCH='GRID1'
      ENDPROC 
      procedure cmnd1.click            
           FLG=OQT.F01         
           CHK_STR=OQT.F02
           QTY=OQT.F04-OQT.F06      
           FTR_STR=OQT.F13
          AREA1='B64_1'
          THISFORM.RELEASE                                            
      endproc
      procedure cmnd2.click
          FLG='0'
          CHK_STR=''
          FTR_STR=''
          QTY=0
          AREA1='B64_1'
         thisform.release
      endproc    
enddefine            
***************************************************特殊輸入欄位設定-----------原使用客戶
DEFINE CLASS TXTF06 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9       
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0 
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F02,F01,F04 FROM C01_1 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
       ELSE
          SELECT F02,F01,F04 FROM C01_1 INTO CURSOR VDR ORDER BY F02 
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
          SELECT F02,F01,F04 FROM C01_1 INTO CURSOR VDR ORDER BY F02 WHERE F02=ALLTRIM(THIS.VALUE)
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
       THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=IIF(SEEK(THIS.VALUE,'C01_1'),C01_1.F04,'')  

  ENDPROC
ENDDEFINE



