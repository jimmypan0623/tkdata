***程式名稱:驗後退貨單
close all
clear 
FLG='0'
FCH=''
CHK=''
CHK_STR=''
R=0
QTY=0
TQTY=0
AREA1='H04_1'
AREA2='H04'
IF ! USED('A01')                 &&操作人員基本檔
   SELE 0
   USE A01 INDE A01
ELSE
   SELE A01
ENDIF      
SET ORDER TO A01
IF !USED('A02')                  &&權限設定檔
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'H05'
IF !USED('A14')                   &&部門主檔
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET FILTER TO F04='*'
SET ORDER TO 1     
IF !USED('A09')                   &&單據編號檔
   SELE 0
   USE A09
ELSE
   SELE A09
ENDIF
SET ORDER TO 1
IF !USED('A23')                  &&工作月份檔  
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
IF !USED('B01')                  &&物料基本主檔
   SELE 0
   USE B01 INDE B01
ELSE
   SELE B01
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
IF !USED('D01')              &&廠商基本主檔
   SELE 0
   USE D01 INDE D01         
ELSE
   SELE D01
ENDIF
SET ORDER TO 1
IF !USED('D03')              &&採購訂單表頭
   SELE 0
   USE D03
ELSE
   SELE D03
ENDIF
SET ORDER TO 1               
IF !USED('P03')              &&採購訂單表頭
   SELE 0
   USE P03
ELSE
   SELE P03
ENDIF
SET ORDER TO 1               
IF !USED('D04')               &&採購訂單表身
   SELE 0
   USE D04
ELSE
  SELE D04
ENDIF
SET ORDER TO 1                  
IF !USED('P04')               &&採購訂單表身
   SELE 0
   USE P04
ELSE
  SELE P04
ENDIF
SET ORDER TO 1                  
IF !USED('D07')               &&進貨計劃
    SELE 0
    USE D07
ELSE
    SELE D07
ENDIF
SET ORDER TO D074
IF !USED('E08')                      &&MRP數據
   SELE 0
   USE E08
ELSE
   SELE E08
ENDIF
SET ORDER TO E082                 
IF !USED('H01')                       &&待驗資料檔  
   SELE 0
   USE H01
ELSE
   SELE H01
ENDIF
SET ORDER TO H014      
H04='H04'+LEFT(DTOS(DATE()),6)
H041='H041'+LEFT(DTOS(DATE()),6)
IF !USED('&H04')
   SELE 0
   USE (H04) ALIA H04
ELSE
   SELE H04
ENDIF
SET ORDER TO 1      
SET RELA TO F03 INTO B01
CREATE CURSOR H04_1;
(F01 C(10),F02 C(2),F05 C(5),F10 C(1),F11 C(17),F15 D(8),F90 C(17))
INDE ON F01 TAG H04_11
INDE ON F05+F01 TAG H04_12
SELE F01,F02,F05,F10,F11,F15,F90 FROM H04 GROUP BY F01 ORDER BY F01 INTO CURSOR H04_2
SELE H04_2
GO TOP
DO WHILE !EOF()
   SELE H04_1
   APPEND BLANK
   REPL F01 WITH H04_2.F01
   REPL F02 WITH H04_2.F02
   REPL F05 WITH H04_2.F05
   REPL F10 WITH H04_2.F10
   REPL F11 WITH H04_2.F11
   REPL F15 WITH H04_2.F15
   REPL F90 WITH H04_2.F90
   SELE H04_2
   SKIP   
ENDDO
SELE H04_2
USE   
SELE H04_1
SET ORDER TO 1
SET RELA TO F05 INTO D01
H05form=createobject("tkH05")
H05form.show  
define class tkH05 as form
  caption='H05.驗後退貨單'
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
  NAME='TKH04'
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
      columncount=3,;            
      RECORDSOURCE='H04_1',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3'    
  ADD OBJECT grid2 AS grid2 WITH;
      columncount=5,;            
      RECORDSOURCE='H04',; 
      LINKMASTER='H04_1',;
      RELATIONALEXPR='F01',;
      childorder=H041,;  
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.              
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.              
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*150,;
      LEFT=INT_015*600          
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
          .CAPTION='驗退單號'          
     ENDWITH
     THIS.ADDOBJECT('LBLF05','LABEL')
     WITH THIS.LBLF05
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*38
         .WIDTH=INT_015*50
         .CAPTION='廠商編號'
     ENDWITH
     THIS.ADDOBJECT('LBLF051','LABEL')
     WITH THIS.LBLF051
         .VISIBLE=.T.
         .LEFT=INT_015*115
         .TOP=INT_015*38
         .AUTOSIZE=.T.
     ENDWITH     
     THIS.ADDOBJECT('LBLF15','LABEL')
     WITH THIS.LBLF15
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*63
         .WIDTH=INT_015*50
         .CAPTION='補貨日期'
     ENDWITH          
     THIS.ADDOBJECT('LBLF02','LABEL')
     WITH THIS.LBLF02
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*13
         .WIDTH=INT_015*50
         .CAPTION='驗退日期'
     ENDWITH    
     THIS.ADDOBJECT('LBLF10','LABEL')
     WITH THIS.LBLF10
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*38
         .WIDTH=INT_015*50
         .CAPTION='過  帳  碼'
     ENDWITH    
     THIS.ADDOBJECT('LBLF101','LABEL')
     WITH THIS.LBLF101
         .VISIBLE=.T.
         .LEFT=INT_015*419
         .TOP=INT_015*38
         .AUTOSIZE=.T.
     ENDWITH    
     THIS.ADDOBJECT('LBLF11','LABEL')
     WITH THIS.LBLF11
         .VISIBLE=.T.
         .LEFT=INT_015*368
         .TOP=INT_015*63
         .WIDTH=INT_015*50
         .CAPTION='安全資料'
     ENDWITH    
     THIS.ADDOBJECT('LBLF111','LABEL')
     WITH THIS.LBLF111
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
     THIS.ADDOBJECT('TXTF05','TXTF05')          
     WITH THIS.TXTF05
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*34
          .WIDTH=INT_015*49
          .HEIGHT=INT_015*25
          .MAXLENGTH=5
     ENDWITH            
     THIS.ADDOBJECT('TXTF15','TEXTBOX') 
     WITH THIS.TXTF15
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*59
          .WIDTH=INT_015*70
          .HEIGHT=INT_015*25
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
         .LEFT=INT_015*327
             .TOP=INT_015*18
             .autosize=.t.
        ENDWITH     
    THIS.ADDOBJECT('LBLF09','LABEL')
    WITH THIS.LBLF09
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*43
         .WIDTH=INT_015*50
         .CAPTION='進貨單號'
    ENDWITH     
    THIS.ADDOBJECT('LBLF08','LABEL')
    WITH THIS.LBLF08
         .VISIBLE=.T.
         .LEFT=INT_015*165
         .TOP=INT_015*43
         .AUTOSIZE=.T.
    ENDWITH         
    THIS.ADDOBJECT('LBLF06','Label')
    WITH THIS.LBLF06
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*68
         .WIDTH=INT_015*50
         .CAPTION='退貨數量'         
    ENDWITH     
    THIS.ADDOBJECT('LBLF14','LABEL')    
    WITH THIS.LBLF14
         .VISIBLE=.T.           
         .LEFT=INT_015*14
         .TOP=INT_015*143
         .WIDTH=INT_015*50
         .CAPTION='採購用途'
    ENDWITH     
    THIS.ADDOBJECT('LBLF141','LABEL')
    WITH THIS.LBLF141
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*143
         .AUTOSIZE=.T.
    ENDWITH             

    THIS.ADDOBJECT('LBLF11','LABEL')
    WITH THIS.LBLF11
        .VISIBLE=.T.           
        .LEFT=INT_015*328
        .TOP=INT_015*143
        .WIDTH=INT_015*50
        .CAPTION='安全資料'         
    ENDWITH    
    THIS.ADDOBJECT('LBLF111','LABEL')
    WITH THIS.LBLF111
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
        THIS.ADDOBJECT('TXTF09','TXTF09')
        WITH THIS.TXTF09  
             .VISIBLE=.T.
             .LEFT=INT_015*65
             .TOP=INT_015*39
             .WIDTH=INT_015*90
             .HEIGHT=INT_015*25
             .MAXLENGTH=10
             .NAME='TXTF09'  
        ENDWITH             
        THIS.ADDOBJECT('TXTF06','TEXTBOX')
        WITH THIS.TXTF06
             .VISIBLE=.T.
             .LEFT=INT_015*65
             .TOP=INT_015*64
             .WIDTH=INT_015*100
             .HEIGHT=INT_015*25
             .INPUTMASK='99999999.99'
             .NAME='TXTF06'         
        ENDWITH             
  ENDPROC       
  PROCEDURE PAGEFRAME1.PAGE1.activate
     R=0
          SELE H04_1  
     IF THISFORM.SHRGROUP.VISIBLE=.F.
        THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(H04_1.F10<>'2',RGB(255,0,0),'')","COLUMN")         
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
        THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF111.CAPTION=''        
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.      
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  .AND. H04_1.F10<>'2' &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   .AND. H04_1.F10<>'2' &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)    &&判斷有無列印的權限  
        THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.
     ENDIF   
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.);
     .AND. !IIF(SEEK(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),'A23','A231'),A23.F07,.F.)
  ENDPROC
  PROCEDURE PAGEFRAME1.PAGE2.activate
       SELE H04
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
        THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=''
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. H04.F13<>'2'  &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. H04.F13<>'2' &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)  &&&&判斷有無列印的權限  
        THISFORM.ANS_BOTT.ENABLED=IIF(H04.F13='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED 
     ENDIF   
     R=RECNO('&AREA1')
     THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
     THISFORM.GRID1.ENABLED=.F.     
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)   .AND. H04_1.F10<>'2' &&判斷有無新增的權限
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
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID2.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'HEADER')         
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='檢驗單號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='廠商編號'
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='廠商簡稱'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='H04_1.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='H04_1.F05'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='D01.F04'
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*81
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*49
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(H04_1.F10<>'2',RGB(255,0,0),'')","COLUMN")         
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'HEADER')       
       THISFORM.grid2.COLUMN1.HEADER1.CAPTION='料品編號'
       THISFORM.grid2.COLUMN2.HEADER1.CAPTION='進貨單號'
       THISFORM.grid2.COLUMN3.HEADER1.CAPTION='採購單號'
       THISFORM.grid2.COLUMN4.HEADER1.CAPTION='驗退數量'
       THISFORM.grid2.COLUMN5.HEADER1.CAPTION='採購用途'       
       THISFORM.GRID2.LINKMASTER='H04_1'
       THISFORM.GRID2.RELATIONALEXPR='F01'
       THISFORM.grid2.COLUMN1.CONTROLSOURCE='H04.F03'
       THISFORM.grid2.COLUMN2.CONTROLSOURCE='H04.F09'
       THISFORM.grid2.COLUMN3.CONTROLSOURCE='H04.F08'
       THISFORM.grid2.COLUMN4.CONTROLSOURCE='H04.F06'
       THISFORM.grid2.COLUMN5.CONTROLSOURCE='H04.F14'       
       THISFORM.grid2.COLUMN1.WIDTH=INT_015*149
       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*100
       THISFORM.grid2.COLUMN3.WIDTH=INT_015*100
       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*100
       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*100
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
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. H04_1.F10<>'2'  &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. H04_1.F10<>'2' &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限       
          THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.)  .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
       ENDIF          
       THISFORM.KEY_LIST.DISPLAYVALUE='依驗退單號排列'      
       JK='驗退單號'
       SELE H04_1
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=H04_1.F01
  ENDPROC               
  PROC KEY_LIST.INIT 
       with this 
           .additem('依驗退單號排列')
           .additem('依廠商編號排列')
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE H04_1 
       DO CASE
          CASE THIS.DISPLAYVALUE='依驗退單號排列'
               set order to 1 
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='驗退單號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='H04_1.F01'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*81
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='廠商編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='H04_1.F05'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*49
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='廠商簡稱'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='D01.F04'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60                       
          CASE THIS.DISPLAYVALUE='依廠商編號排列'
               set order to 2
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='廠商編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='H04_1.F05'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*49
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='廠商簡稱'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='D01.F04'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='驗退單號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='H04_1.F01'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*81                                                          
       ENDCASE 

        *SELE &AREA1
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
       SELE H04
       set rela off into b01
       SET RELA OFF INTO H04_1
       USE
       SELE H04_1
       set rela off into D01
       USE
       H04='H04'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       H041='H041'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&H04')
          SELE 0
          USE (H04) ALIA H04
       ELSE
          SELE H04
       ENDIF
       SET ORDER TO 1      
       SET RELA TO F03 INTO B01
       CREATE CURSOR H04_1;       
       (F01 C(10),F02 C(2),F05 C(5),F10 C(1),F11 C(17),F15 D(8),F90 C(17))
       INDE ON F01 TAG H04_11
       INDE ON F05+F01 TAG H04_12
       SELE F01,F02,F05,F10,F11,F15,F90 FROM H04 GROUP BY F01 ORDER BY F01 INTO CURSOR H04_2
       SELE H04_2
       GO TOP
       DO WHILE !EOF()
          SELE H04_1
          APPEND BLANK
          REPL F01 WITH H04_2.F01
          REPL F02 WITH H04_2.F02
          REPL F05 WITH H04_2.F05
          REPL F10 WITH H04_2.F10
          REPL F11 WITH H04_2.F11
          REPL F15 WITH H04_2.F15
          REPL F90 WITH H04_2.F90
          SELE H04_2
          SKIP   
       ENDDO
       SELE H04_2
       USE   
       SELE H04_1
       SET ORDER TO 1
       SET RELA TO F05 INTO D01
       THISFORM.GRID1.RECORDSOURCE='H04_1'
       THISFORM.GRID2.RECORDSOURCE='H04'
       THISFORM.GRID2.LINKMASTER='H04_1'
       THISFORM.GRID2.RELATIONALEXPR='F01'      
       THISFORM.GRID2.CHILDORDER=H041  
       THISFORM.grid2.COLUMN1.CONTROLSOURCE='H04.F03'
       THISFORM.grid2.COLUMN2.CONTROLSOURCE='H04.F09'
       THISFORM.grid2.COLUMN3.CONTROLSOURCE='H04.F08'
       THISFORM.grid2.COLUMN4.CONTROLSOURCE='H04.F06'
       THISFORM.grid2.COLUMN5.CONTROLSOURCE='H04.F14'       
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
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. H04_1.F10<>'2'  &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. H04_1.F10<>'2' &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
          THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.)  .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED               
       ENDIF          
  ENDPROC
  PROC grid1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=H04_1.F01
       THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=H04_1.F05
       THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=D01.F03
       THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=H04_1.F10
       THISFORM.PAGEFRAME1.PAGE1.LBLF111.CAPTION=H04_1.F11   
       THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=H04_1.F90    
       THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=H04_1.F02
       THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE=H04_1.F15
       THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) .AND. H04_1.F10<>'2'  &&判斷有無修改的權限
       THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)  .AND. H04_1.F10<>'2' &&判斷有無刪除的權限
       THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED        
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION  
       IF THIS.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
       ENDIF   
       CHK=''     
  ENDPROC     
  PROC grid2.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=H04.F03
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=B01.F02
       THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=H04.F06               
       THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=H04.F09               
       THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION=H04.F11      
       THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=H04.F08
       THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=H04.F14
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       if thisform.pageframe1.activepage=1
          thisform. CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ENDIF   
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
          H='H003 '+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
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
                REPL f01 WITH 'H003 '
                REPL f02 WITH  SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                REPL F05 WITH 'HC'+f02
                REPL F06 WITH '驗退單'
                REPL F08 WITH ALLTRIM(F05)+'0000'
              ENDIF
              LOCK()
              PO_NO=RIGHT(F08,4)
              HX=VAL(PO_NO)+1
              PO_NO='0000'
              PO_NO=LEFT(F08,6)+STUFF(PO_NO,5-INT(LOG(HX)/LOG(10)+1),INT(LOG(HX)/LOG(10)+1),RIGHT(STR(HX),INT(LOG(HX)/LOG(10)+1)))
              REPL F08 WITH PO_NO                                             &&記錄本訂單號碼,本月份後續的訂單號碼將以此往下編
              UNLOCK
          ENDIF       
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=PO_NO
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.REFRESH
          thisform.PAGEFRAME1.PAGE1.txtf01.readonly=.t.
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
          THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.READONLY=.F.                  
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=''
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=RIGHT(DTOS(DATE()),2)
          THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
          THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE=CTOD('')
          THISFORM.PAGEFRAME1.PAGE1.LBLF101.CAPTION=''
          THISFORM.PAGEFRAME1.PAGE1.LBLF111.CAPTION=''   
          THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=''             
     ELSE
        THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.              &&處理表身新增
        THISFORM.GRID2.ENABLED=.F.   
        THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX')     
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''          
        THISFORM.PAGEFRAME1.PAGE2.lblF031.caption=''            
        THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=''              
        THISFORM.PAGEFRAME1.PAGE2.TXTF09.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=0
        THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION=''    
        THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=''                   
     ENDIF   

  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
     THISFORM.KEY_LIST.ENABLED=.F. 
     THISFORM.MTH_LIST.ENABLED=.F. 
     THISFORM.ANS_BOTT.ENABLED=.F.
     THISFORM.GRID1.ENABLED=.F.   
     THISFORM.GRID2.ENABLED=.F.   
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
        SELE H04_1
        PO_NO=H04_1.F01
        SELE RECNO() FROM H04 WHERE H04.F01=PO_NO AND F13<>'2' INTO ARRAY BK1
        JJ1=''                
        IF _TALLY>0                   
           FOR I= 1 TO ALEN(BK1)
               JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
           ENDFOR  
           KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')
           RLOCK(KK1,'H04')
           LL1=RLOCK(KK1,'H04')
        ELSE 
           =MESSAGEBOX('此單據已由別的工作站過帳中無法修改!',0+64,'提示訊息視窗')
           SELE H04_1           
           UNLOCK ALL
           REPL F13 WITH '2'
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   
        ENDIF   
        IF LL1 =.T.
           THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')     &&處理表頭修改
           THISFORM.PAGEFRAME1.PAGE1.TXTF01.READONLY=.T.
           THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
           THISFORM.PAGEFRAME1.PAGE1.TXTF15.SETFOCUS
        ELSE 
           DO file_impact
           UNLOCK ALL
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
        endif   
     ELSE                      
        THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.            
        SELE H04                                                      &&處理表身修改
        rlock()
        IF RLOCK() 
           THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX')            
           THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
        ELSE  
           DO file_impact 
           unlock all
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           
       endif        
     ENDIF   
  endproc    
  PROC ORPGROUP.DEL_BOTT.CLICK                    &&刪除資料程序
       IF messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6            
          if ALIA()='H04_1'           &&刪除整張訂單
             SELE H04_1             
                PO_NO=H04_1.F01
                SELE RECNO() FROM H04 WHERE H04.F01=PO_NO  AND H04.F13<>'2' INTO ARRAY BK1              
                JJ1=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
                   ENDFOR    
                   KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
                   RLOCK(KK1,'H04')
                   LL1=RLOCK(KK1,'H04')
                ELSE
                   LL1=.T.   
                ENDIF                   
                IF LL1=.T.
                   IF LEN(ALLTRIM(JJ1))>0 
                      SELE H04
                      FOR I= 1 TO ALEN(BK1)
                          GO BK1(I) 
                          SELE H01
                          SET ORDER TO H014                   
                          SEEK H04.F03+H04.F09+H04.F08
                            IF FOUND()
                               REPLACE F22 WITH F22+H04.F06*(-1)                            
                            ENDIF                                  
                            SELE H04
                       ENDFOR                      
                   ELSE
                      =MESSAGEBOX('此張單據已由別的工作站過帳中,無法刪除!',0+64,'提示訊息視窗')
                      SELE H04_1
                      REPL F13 WITH '2'
                      UNLOCK ALL
                      RETURN   
                   ENDIF    
                   DELE FROM H04 WHERE H04.F01=PO_NO     &&表身整個清除
                   SELE B39                     &&刪除未過帳單據紀錄
                   SEEK A02.F03+H04_1.F01                   
                   IF FOUND()
                      DELE
                   ENDIF   
                   SELE H04_1
                   DELE
                   SKIP -1
                   IF BOF()
                      GO TOP
                   ENDIF   
                   UNLOCK ALL
                 IF  SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                      SELE A24
                      SET ORDER TO 1
                      APPEND BLANK
                      LOCK()
                      REPL F01 WITH 'H003 '
                      REPL F02 WITH SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                              REPL F05 WITH 'HC'+f02
                      REPL F06 WITH '驗後退貨單'
                      REPL F08 WITH PO_NO
                      UNLOCK
                  ENDIF  

                ELSE
                   DO file_impact   
                   RETURN               
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
                THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  .AND. H04_1.F10<>'2' &&判斷有無修改的權限
                THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   .AND. H04_1.F10<>'2' &&判斷有無刪除的權限
                THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限
                THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED                
             ENDIF     
          ELSE        
              SELE H04                                                      &&刪除單筆資料及回復期待驗狀態     
              RLOCK() 
              IF RLOCK()                    
                 SELE H01
                 SET ORDER TO H014
                 SEEK H04.F03+H04.F09+H04.F08
                 IF FOUND()
                    REPLACE F22 WITH F22+H04.F06*(-1)                             
                  ENDIF                                              
                  SELE H04
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
                 SEEK A02.F03+H04_1.F01
                 IF FOUND()
                    DELE
                 ENDIF                    
                 SELE H04_1                        
                 PO_NO=H04_1.F01                    
                 DELE                               &&連表頭也要刪除
                 IF SUBSTR(PO_NO,3,4)=SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                      SELE A24
                      SET ORDER TO 1
                      APPEND BLANK
                      LOCK()
                      REPL F01 WITH 'H003 '
                      REPL F02 WITH SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                      REPL F05 WITH 'HC'+f02
                      REPL F06 WITH '驗退單'
                      REPL F08 WITH PO_NO
                 ENDIF
                 SELE H04_1
                 SKIP -1
                 IF BOF()
                    GO TOP
                 ENDIF   
                 UNLOCK ALL                                 
                 THISFORM.PAGEFRAME1.ACTIVEPAGE=1 
                 THISFORM.REFRESH          
              ENDIF
          ENDIF
       ENDIF   
       RELEASE ALL LIKE BK*
  ENDPROC  
  PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK               &&新增確認
    IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
       THISFORM.GRID2.RECORDSOURCE='H04'
       THISFORM.GRID2.CHILDORDER=H041
       THISFORM.grid2.COLUMN1.CONTROLSOURCE='H04.F03'
       THISFORM.grid2.COLUMN2.CONTROLSOURCE='H04.F09'
       THISFORM.grid2.COLUMN3.CONTROLSOURCE='H04.F08'
       THISFORM.grid2.COLUMN4.CONTROLSOURCE='H04.F06'
       THISFORM.grid2.COLUMN5.CONTROLSOURCE='H04.F14'              
       SELE H04_1 
       COD=SYS(21)
       SET ORDER TO 1
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE,'H04_1')=.T.
          SELE A24
          SEEK 'H003 '+SUBSTR(DTOS(DATE()),3,4)
          IF FOUND()
             DO WHILE F08=PO_NO
                RLOCK()
                DELE
                UNLOCK
                SKIP
             ENDDO   
             SELE A09
             SEEK 'H003 '+SUBSTR(DTOS(DATE()),3,4)
             IF FOUND()
                RLOCK()
                PO_NO=RIGHT(F08,4)
                HX=VAL(PO_NO)+1
                PO_NO='0000'                
                PO_NO=LEFT(F08,6)+STUFF(PO_NO,5-INT(LOG(HX)/LOG(10)+1),INT(LOG(HX)/LOG(10)+1),RIGHT(STR(HX),INT(LOG(HX)/LOG(10)+1)))
                REPL F08 WITH PO_NO                                             &&記錄本訂單號碼,本月份後續的訂單號碼將以此往下編
                UNLOCK
                THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=PO_NO
                THISFORM.PAGEFRAME1.PAGE1.TXTF01.REFRESH
             ENDIF                              
          ENDIF   
       ENDIF
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE,'D01')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE)
          =MESSAGEBOX('廠商基本主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF
       SELE RECNO() FROM H01 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE AND F07-F22>0 INTO ARRAY RC1
       IF _TALLY=0                 
          =MESSAGEBOX('此廠商已無進貨單可退!',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS                          
          SET ORDER TO &COD
          RETURN                                     
       ENDIF   
       IF EMPTY(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/'+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE))          
          =MESSAGEBOX('驗退日期輸入錯誤!',0+64,'提示訊息視窗') 
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
          SET ORDER TO &COD          
          RETURN
       ENDIF   
       IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE)
          if messagebox('無補貨日期是否將此張單據內容計入取消量?',4+32+256,'請確認') <> 6	        
             THISFORM.PAGEFRAME1.PAGE1.TXTF15.SETFOCUS
             SET ORDER TO &COD          
             RETURN
          ENDIF 
       ENDIF          
       SELE H04_1          
       APPEND BLANK
       LOCK()
       IF RLOCK()
          REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
          REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
          REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
          REPLACE F11 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
          REPLACE F15 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE
       ENDIF
       SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       SET ORDER TO &COD
       SELE B39
       APPEND BLANK
       REPL F01 WITH A02.F03
       REPL F02 WITH H04_1.F01
       REPL F03 WITH H04_1.F02
       REPL F04 WITH H04_1.F05
       REPL F09 WITH H04_1.F15
       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.  
       THISFORM.PAGEFRAME1.ACTIVEPAGE=2
       THISFORM.ORPGROUP.NEW_BOTT.CLICK
       THISFORM.REFRESH
    ELSE                                                             &&表身新增確認
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'B01')=.F.
          =MESSAGEBOX('物料基本主檔中無此料號!',0+48,'提示訊息視窗')          
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF         
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE,'H01')=.F.          
          =MESSAGEBOX('待驗檔中無此筆資料!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF   
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE;
          +THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION),'H04')=.T. 
          =MESSAGEBOX('此筆資料在此張驗退單重複!',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''
          THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=''
          THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=''
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN            
       ENDIF                         
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE>;
          IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION,'H01'),;
          H01.F07-H01.F22,0)
          =MESSAGEBOX('驗退數量不得大於待驗數量';
          +ALLTRIM(STR(IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION,'H01'),;
          H01.F07-H01.F22,0))),0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
          RETURN
       ENDIF
       SELE H04
       APPEND BLANK
       LOCK()
       IF RLOCK()
          REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
          REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
          REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
          REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
          REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
          REPLACE F07 WITH IIF(SEEK(F03+F09+F08,'H01'),H01.F09,0)          
          REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION
          REPLACE F09 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE
          REPLACE F11 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')          
          REPLACE F12 WITH IIF(SEEK(F03+F09+F08,'H01'),H01.F16,0)          
          REPLACE F13 WITH IIF(EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE),'N','Y')
          REPLACE F14 WITH THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION
          REPLACE F15 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE
          REPLACE F16 WITH IIF(SEEK(F03+F09+F08,'H01'),H01.F13,0)          
      ENDIF
      UNLOCK ALL         
      seek thisform.pageframe1.page1.txtf01.value+thisform.pageframe1.page2.txtf03.value+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION)
      thisform.grid2.enabled=.t.
      thisform.grid2.setfocus  
      SELE H01                 
      SET ORDER TO H014                                                         &&寫入待驗檔中
      SEEK H04.F03+H04.F09+H04.F08
      IF FOUND()      
         REPL F22 WITH F22+H04.F06
      ENDIF               
      THISFORM.ORPGROUP.NEW_BOTT.CLICK      
    ENDIF  
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK                          &&修改確認
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
       IF EMPTY(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/'+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE))          
          =MESSAGEBOX('進貨日期輸入錯誤!',0+64,'提示訊息視窗') 
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
          RETURN
        ENDIF           

           SELE H04
           SEEK H04_1.F01
           IF FOUND()
              DO WHILE F01=H04_1.F01
                 REPL F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
                 REPL F15 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE
                 REPL F11 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')    
                 REPL F13 WITH IIF(EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE),'N','Y')                             
                 SKIP
              ENDDO    
           ENDIF        
        SELE H04_1
           REPL F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
           REPL F15 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF15.VALUE
           REPL F11 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
           UNLOCK ALL    
     ELSE                                                        &&表身修改確認     
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'B01')=.F.
          =MESSAGEBOX('物料基本主檔中無此料號!',0+48,'提示訊息視窗')          
          THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
          RETURN
       ENDIF                 
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE>;
          IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION,'H01'),;
          H01.F07-H01.F22,0)+H04.F06
          =MESSAGEBOX('驗退數量不得大於待退數量';
          +ALLTRIM(STR(IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE+THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION,'H01'),;
          H01.F07-H01.F22,0)+H04.F06)),0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF06.SETFOCUS
          RETURN
       ENDIF                 
       SELE H01                             &&回復待驗狀態
       SEEK H04.F03+H04.F09+H04.F08
       IF FOUND()          
          REPLACE F22 WITH F22+H04.F06*(-1)
       ENDIF   
       SELE H04                    
       REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
       REPLACE F11 WITH DTOC(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')             
       SELE H01                          &&寫回待驗檔中
       SEEK H04.F03+H04.F09+H04.F08
       IF FOUND()   
          REPLACE F22 WITH F22+H04.F06
       ENDIF                                    
       UNLOCK ALL
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
      THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.T.   
      THISFORM.PAGEFRAME1.PAGE2.TXTF09.READONLY=.T.    
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
            SELE A24         
            SET ORDER TO 1
            APPEND BLANK
            LOCK()
            REPL F01 WITH 'H003 '
            REPL F02 WITH SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                    REPL F05 WITH 'HC'+f02
            REPL F06 WITH '驗後退貨單'
            REPL F08 WITH PO_NO
            SELE B39
            SEEK A02.F03+H04_1.F01
            SET ORDER TO 1
            IF FOUND()
               DELE            
            ENDIF   
            SELE H04_1
            DELE
            THISFORM.PAGEFRAME1.ACTIVEPAGE=1
            THISFORM.REFRESH  
         ELSE
            SELE B39
            SET ORDER TO 1
            SEEK A02.F03+H04_1.F01 
            IF FOUND()
               REPL F08 WITH H04.F08            
            ENDIF   
         ENDIF   
      ENDIF       
      UNLOCK ALL
      THISFORM.ANS_BOTT.ENABLED=IIF(H04_1.F10='2',.F.,.T.) .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED          
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
         IF THISFORM.GRID2.RECORDSOURCE<>'H04'
            THISFORM.GRID2.RECORDSOURCE='H04'
            THISFORM.GRID2.CHILDORDER=H041
            THISFORM.grid2.COLUMN1.CONTROLSOURCE='H04.F03'
            THISFORM.grid2.COLUMN2.CONTROLSOURCE='H04.F09'
            THISFORM.grid2.COLUMN3.CONTROLSOURCE='H04.F08'
            THISFORM.grid2.COLUMN4.CONTROLSOURCE='H04.F06'
            THISFORM.grid2.COLUMN5.CONTROLSOURCE='H04.F14'                
             ENDIF   
             IF INT_099=.T.
            SELE A24         
            SET ORDER TO 1
            APPEND BLANK
            LOCK()
            REPL F01 WITH 'H003 '
            REPL F02 WITH SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4)
                    REPL F05 WITH 'HC'+f02
            REPL F06 WITH '驗退單'
            REPL F08 WITH PO_NO
            UNLOCK
         ENDIF   
         SELE H04_1
         DO CASE 
            CASE THISFORM.KEY_LIST.DISPLAYVALUE='依驗退單號排列'
                 SET ORDER TO 1
            CASE THISFORM.KEY_LIST.DISPLAYVALUE='依廠商邊號排列'
                 SET ORDER TO 2
         ENDCASE   
       ENDIF  
       THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC     
  PROC ANS_BOTT.CLICK     &&過帳程序
    if messagebox('是否確認資料無誤可以過帳?',4+32+256,'請確認') = 6	        
       SELE RECNO() FROM H04 WHERE H04.F01=H04_1.F01  AND F10<>'2' INTO ARRAY BK1              
            JJ1=''                
            IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                   JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'H04')
               LL1=RLOCK(KK1,'H04')
            ELSE
               LL1=.T.   
            ENDIF                   
            SELE RECNO() FROM D07 WHERE F01+F02= ANY (SELE F08+F03 FROM H04 WHERE F01=H04_1.F01) INTO ARRAY BK2
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
            SELE RECNO() FROM D04 WHERE F01+F02=ANY (SELE F08+F03 FROM H04 WHERE F01=H04_1.F01) INTO ARRAY BK3
            JJ3=''                
            IF _TALLY>0   
               FOR I= 1 TO ALEN(BK3)
                   JJ3=JJ3+ALLTRIM(STR(BK3(I)))+','
               ENDFOR    
               KK3=STUFF(JJ3,RAT(',',JJ3,1),1,'')              
               RLOCK(KK3,'D04')
               LL3=RLOCK(KK3,'D04')
            ELSE
               LL3=.T.   
            ENDIF                   
            SELE RECNO() FROM P04 WHERE F01+F02=ANY (SELE F08+F03 FROM H04 WHERE F01=H04_1.F01) INTO ARRAY BK4
            JJ4=''                
            IF _TALLY>0   
               FOR I= 1 TO ALEN(BK4)
                   JJ4=JJ4+ALLTRIM(STR(BK4(I)))+','
               ENDFOR    
               KK4=STUFF(JJ4,RAT(',',JJ4,1),1,'')              
               RLOCK(KK4,'P04')
               LL4=RLOCK(KK4,'P04')
            ELSE
               LL4=.T.   
            ENDIF                               
            SELE RECNO() FROM D03 WHERE F02=ANY (SELE F08 FROM H04 WHERE F01=H04_1.F01) INTO ARRAY BK5
            JJ5=''
            IF _TALLY>0   
               FOR I= 1 TO ALEN(BK5)
                   JJ5=JJ5+ALLTRIM(STR(BK5(I)))+','
               ENDFOR    
               KK5=STUFF(JJ5,RAT(',',JJ5,1),1,'')              
               RLOCK(KK5,'D03')
               LL5=RLOCK(KK5,'D03')
            ELSE
               LL5=.T.   
            ENDIF                 
            SELE RECNO() FROM P03 WHERE F02=ANY (SELE F08 FROM H04 WHERE F01=H04_1.F01) INTO ARRAY BK6
            JJ6=''
            IF _TALLY>0   
               FOR I= 1 TO ALEN(BK6)
                   JJ6=JJ6+ALLTRIM(STR(BK6(I)))+','
               ENDFOR    
               KK6=STUFF(JJ6,RAT(',',JJ6,1),1,'')              
               RLOCK(KK6,'P03')
               LL6=RLOCK(KK6,'P03')
            ELSE
               LL6=.T.   
            ENDIF                             

            IF LL1 .AND. LL2 .AND. LL3 .AND. LL4 .AND. LL5 .AND. LL6=.T. 
               IF LEN(ALLTRIM(JJ1))>0 
                  SELE H04
                  FOR I= 1 TO ALEN(BK1)                      
                      GO BK1(I)                          
                      IF H04.F13='Y'                                      &&如果需要補貨
                         SELE D07                                            &&寫入進貨計劃
                         SET ORDER TO D074                 
                         APPEND BLANK
                         LOCK()
                         IF RLOCK()
                            REPLACE F01 WITH H04.F08
                            REPLACE F02 WITH H04.F03
                            REPLACE F03 WITH H04.F15
                            REPLACE F04 WITH H04.F06
                            REPLACE F05 WITH H04.F05
                            REPLACE F07 WITH H04.F14
                            REPLACE F08 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')        
                            REPLACE F09 WITH '1'
                            REPLACE F11 WITH IIF(SEEK(H04.F03+H04.F09+H04.F08,'H01'),H01.F10,CTOD(''))
                            IF LEFT(H04.F09,2)='BA'
                               REPLACE F12 WITH IIF(SEEK(H04.F08+H04.F03,'D04'),D04.F20,0)
                            ELSE
                               REPLACE F12 WITH IIF(SEEK(H04.F08+H04.F03,'P04'),P04.F20,0)                            
                            ENDIF   
                         ENDIF  
                         UNLOCK      
                         SELE E08                           &&寫入MRP紀錄檔
                         SEEK H04.F03+DTOS(MRP_PRD(H04.F15))
                         IF FOUND()
                            REPL F05 WITH F05+H04.F06
                         ELSE
                            APPEND BLANK
                            REPL F01 WITH MRP_PRD(H04.F15)
                            REPL F02 WITH H04.F03
                            REPL F05 WITH F05+H04.F06
                         ENDIF     
                         FLG=H04.F03
                         DO MRP_PRC        &&MRP即時計算程序                        
                      ENDIF                                                             
                      SELE H01                                            &&待驗檔
                      SET ORDER TO H014                 
                      SEEK H04.F03+H04.F09+H04.F08
                      IF FOUND()
                         REPL F07 WITH F07+H04.F06*(-1)
                         REPL F22 WITH F22+H04.F06*(-1)
                         IF F06+F07=0
                            DELE
                         ENDIF                       
                      ENDIF
                      SELE D04           &&訂單表身
                      SEEK H04.F08+H04.F03
                      IF FOUND()
                         IF H04.F13<>'Y'                     
                            REPL F09 WITH F09+(H04.F06)*(-1)
                            REPL F16 WITH F16+H04.F06
                         ENDIF                                   
                         REPL F18 WITH F18+(H04.F06)*(-1)
                      ENDIF   
                      SELE D03          &&訂單表頭
                      SEEK H04.F08
                      IF FOUND()
                         IF F08<>.T. 
                            REPL F08 WITH .T.
                         ENDIF   
                      ENDIF                        
                      SELE P04           &&外購訂單表身
                      SEEK H04.F08+H04.F03
                      IF FOUND()
                         IF H04.F13<>'Y'                     
                            REPL F09 WITH F09+(H04.F06)*(-1)
                            REPL F16 WITH F16+H04.F06
                         ENDIF                                   
                         REPL F18 WITH F18+(H04.F06)*(-1)
                      ENDIF   
                      SELE P03          &&外購訂單表頭
                      SEEK H04.F08
                      IF FOUND()
                         IF F08<>.T. 
                            REPL F08 WITH .T.
                         ENDIF   
                      ENDIF                                                                                         
                      SELE H04
                      REPL F10 WITH '2'
                      REPL F90 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                  ENDFOR   
                  SELE B39                    &&未過帳庫存單據查詢檔
                  SEEK A02.F03+H04_1.F01
                  IF FOUND()
                     DELE
                  ENDIF                                             
                  SELE H04_1
                  REPL F10 WITH '2'
                  REPL F90 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
               ELSE
                     =MESSAGEBOX('此單據已由別的工作站過帳完成',0+64,'提示訊息視窗')
                     SELE H04_1
                     REPL F10 WITH '2'
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
       ENDIF   
       unlock all
       RELEASE ALL LIKE BK*    
  ENDPROC
enddefine    

***********************************************列印程序
PROCEDURE PNT_PRC  
     LK=H04_1.F01     
     HK=LEFT(DTOC(A23.F01),5)
     SELECT H04.F01,;
            H04.F02,;
            H04.F03,;
            H04.F05,;
            H04.F06,;
            H04.F07,;            
            H04.F08,;                    
            H04.F09,;
            H04.F13,;
            H04.F14,;
            H04.F15,;            
            D01.F03,;
            D01.F05,;
            D01.F08,;
            D01.F09,;
            B01.F02;
     FROM H04,D01,B01 WHERE H04.F01=LK AND D01.F01=H04.F05 AND B01.F01=H04.F03 ;
     ORDER BY H04.F03 NOWAIT
     IF _tally >0
*        REPORT FORM H05 PREVIEW
         report form ALLTRIM(INT_116)+'H05' to print prompt        
     ENDIF   
ENDPROC
************************************************特殊輸入欄位的設定----料號
DEFINE CLASS TXTF03 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       DO CASE
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE))                   
               SELECT F01,F03,F11,F10,F07,F12,F22 FROM H01 INTO CURSOR OPT;
               ORDER BY F01,F10,F03 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE; 
               AND LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND H01.F07-H01.F22>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE))                   
               SELECT F01,F03,F11,F10,F07.F12,F22 FROM H01 INTO CURSOR OPT;
               ORDER BY F01,F10,F03 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE; 
               AND LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND H01.F07-H01.F22>0 AND F03=THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE))                   
               SELECT F01,F03,F10,F11,F07,F12,F22 FROM H01 INTO CURSOR OPT ORDER BY F01,F10,F03 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE;
               AND H01.F07-H01.F22>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE))                   
               SELECT F01,F03,F11,F10,F07,F12,F22 FROM H01 INTO CURSOR OPT ORDER BY F01,F10,F03 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE;
               AND H01.F07-H01.F22>0 AND F03=THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE 
       ENDCASE 
       IF _TALLY>0  
          OPT_SEEK=createobject("OPT_SEEK")  
          OPT_SEEK.SHOW    
          IF LEN(ALLTRIM(FLG))>1
             THIS.VALUE=FLG
             THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=CHK_STR
             THIS.REFRESH          
             THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=QTY
             THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=CHK_STR
             THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=FTR_STR
             THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=RK
             FLG='0'           
             CHK_STR=''
             QTY=0    
             FTR_STR=''
             RK=''
          ENDIF
       ENDIF   
    ELSE
*!*	        IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE)
*!*	           THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=IIF(SEEK(THIS.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE,'H01'),H01.F07-H01.F22,0)
*!*	           THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=IIF(SEEK(THIS.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE,'H01'),H01.F11,'')                  
*!*	           THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=IIF(SEEK(THIS.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE,'H01'),H01.F12,'')           
*!*	        ENDIF
    ENDIF         
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=IIF(SEEK(THIS.VALUE,'B01'),B01.F02,'')      
  ENDPROC    
ENDDEFINE
***************************************************特殊欄位-----------------進貨單號
DEFINE CLASS TXTF09 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=10
  FONTSIZE=INT_015*9  
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       DO CASE
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F03,F10,F07,F12,F12,F22 FROM H01 INTO CURSOR OQT;
               ORDER BY F03,F10,F11,F01 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE; 
               AND LEFT(F03,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND H01.F07-H01.F22>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))>1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F03,F10,F11,F07,F12,F22 FROM H01 INTO CURSOR OQT;
               ORDER BY F03,F10,F01 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE; 
               AND LEFT(F03,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1);    
               AND H01.F07-H01.F22>0 AND F01=THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F03,F10,F11,F07,F12,F22 FROM H01 INTO CURSOR OQT ORDER BY F03,F10,F01 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE;
               AND H01.f07-H01.F22>0 
          CASE AT('?',ALLTRIM(THIS.VALUE))=1 .AND. !EMPTY(ALLTRIM(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE))                   
               SELECT F01,F03,F10,F11,F07,F12,F22 FROM H01 INTO CURSOR OQT ORDER BY F03,F10,F01 WHERE F04=THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE;
               AND H01.F07-H01.F22>0 AND F01=THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE 
       ENDCASE 
       IF _TALLY>0  
          OQT_SEEK=createobject("OQT_SEEK")  
          OQT_SEEK.SHOW    
          IF LEN(ALLTRIM(FLG))>1 
             THIS.VALUE=FLG
             THIS.REFRESH                    
             THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=CHK_STR
             THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=QTY
             THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=FTR_STR
             THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=RK
             FLG='0'           
             CHK_STR=''
             QTY=0 
             FTR_STR=''  
          ENDIF   
       ENDIF    
    ELSE
*!*	       IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE)
*!*	          THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THIS.VALUE,'H01'),H01.F07-H01.F22,0)
*!*	          THISFORM.PAGEFRAME1.PAGE2.LBLF141.CAPTION=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THIS.VALUE,'H01'),H01.F12,'')          
*!*	          THISFORM.PAGEFRAME1.PAGE2.LBLF08.CAPTION=IIF(SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE+THIS.VALUE,'H01'),H01.F11,'')                 
*!*	       ENDIF
    ENDIF  
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=IIF(SEEK(THISform.pageframe1.page2.txtf03.VALUE,'B01'),B01.F02,'')      
  ENDPROC     
ENDDEFINE
***************************************************特殊欄位------------------收料部門
DEFINE CLASS TXTF17 AS TEXTBOX
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
       THISFORM.PAGEFRAME1.PAGE1.LBLF171.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')  
  ENDPROC
ENDDEFINE  
***************************************************特殊輸入欄位設定-----------廠商
DEFINE CLASS TXTF05 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9  
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0 
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          if empty(a02.f12)
             SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) and f18=sys_oper
          else
             SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
          endif
       ELSE
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 
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
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE F02=ALLTRIM(THIS.VALUE)
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
       THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=IIF(SEEK(THIS.VALUE,'D01'),D01.F03,'')  
  ENDPROC
ENDDEFINE

*********************************************************************待檢驗資料----依料號排列搜尋
define class OPT_SEEK as form
  autocenter=.t.
  caption='待驗資料'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*320
  WIDTH=INT_015*410
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
      WIDTH=INT_015*410,;      
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
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='OPT.F01'
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='進貨單號'
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='OPT.F03'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*83
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='驗退日期'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='OPT.F10'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*74
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='待退數量'         
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='OPT.F07-OPT.F22'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*74         
         THISFORM.GRID1.SETFOCUS
         IF THISFORM.GRID1.ACTIVEROW<>0
              THISFORM.GRID1.AFTERROWCOLCHANGE
         ENDIF         
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
           LPARAMETERS XColIndex
           FCH='GRID1'
           FLG=OPT.F01         
           CHK_STR=OPT.F03
           QTY=OPT.F07-OPT.F22
           FTR_STR=OPT.F11
           RK=OPT.F12
      ENDPROC 
      procedure cmnd1.click            
          AREA1='H04_1'
          SELE OPT
          USE
          THISFORM.RELEASE                                            
      endproc
      procedure cmnd2.click
          FLG='0'
          CHK_STR=''
          QTY=0
          AREA1='H04_1'
          SELE OPT
          USE
         thisform.release
      endproc    
enddefine            
***************************************待檢驗資料-------依進貨單號排列搜尋
define class OQT_SEEK as form
  autocenter=.t.
  caption='待驗資料'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*320
  WIDTH=INT_015*410
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
      WIDTH=INT_015*410,;      
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
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='進貨單號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='OQT.F03'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*83
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='OQT.F01'         
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='驗退日期'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='OQT.F10'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*74
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='待退數量'         
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='OQT.F07-OQT.F22'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*74         
         THISFORM.GRID1.SETFOCUS
         IF THISFORM.GRID1.ACTIVEROW<>0
              THISFORM.GRID1.AFTERROWCOLCHANGE
         ENDIF         
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
           LPARAMETERS XColIndex
           FCH='GRID1'
           FLG=OQT.F03         
           CHK_STR=OQT.F01
           QTY=OQT.F07-OQT.F22
           FTR_STR=OQT.F11
           RK=OQT.F12
      ENDPROC 
      procedure cmnd1.click            
          AREA1='H04_1'
          SELE OQT
          USE
          THISFORM.RELEASE                                            
      endproc
      procedure cmnd2.click
          FLG='0'
          CHK_STR=''
          QTY=0
          AREA1='H04_1'
          SELE OQT
          USE
         thisform.release
      endproc    
enddefine            
