close all
clear 
FLG='0'
FCH=''
CHK=''
R=0
AREA1='C01'
AREA2='C1600000'
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
SEEK sys_oper+'C13'

IF !USED('C01')
   SELE 0
   USE C01 INDE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO C011      
if !used('C1600000')
   SELE 0
   USE C1600000
ELSE
   SELE C1600000
ENDIF
SET ORDER TO C1600000      
*SET RELA TO F01 INTO C01
C13form=createobject("tkC13")
C13form.show  
define class tkC13 as form
  caption='出貨計劃分析'
  autocenter=.t.
  controlbox=.F.
  BORDERSTYLE=3  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  controlcount=57
  fontsize=9
  height=550
  showtips=.t.
  showwindow=1
  width=789
  windowtype=1
  NAME='TKC13'
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
      LEFT=236,;
      WIDTH=540,;
      HEIGHT=200,;
      PAGECOUNT=2,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.T.,;
      PAGE1.NAME='PAGE1',;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='依客戶編號搜尋紀錄',;
      PAGE2.NAME='PAGE2',;
      PAGE2.VISIBLE=.T.,;
      PAGE2.CAPTION='詳細資料內容'   
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.
  ADD OBJECT grid1 AS grid1 WITH;
      columncount=3,;            
      RECORDSOURCE='C01',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3'      
  ADD OBJECT grid2 AS grid2 WITH;
      columncount=6,;            
      RECORDSOURCE='C1600000',;
      LINKMASTER='C01',;
      RELATIONALEXPR='F01',;
      CHILDORDER='C1600000',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      column5.backcolor=rgb(255,128,64),;
      COLUMN6.NAME='COLUMN6'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=415,;
      LEFT=10,;
      ENABLED=.T.,;
      VISIBLE=.T.                    
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      TOP=455,;
      LEFT=10
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=173,;
      LEFT=558          
  ADD OBJECT ANS_BOTT AS COMMANDBUTTON WITH;
      TOP=490,;
      LEFT=15,;
      HEIGHT=25,;
      AUTOSIZE=.T.,;
      CAPTION='\<A.資料分析',;
      NAME='ANS_BOTT'            
  PROCEDURE PAGEFRAME1.PAGE1.INIT     

     THIS.ADDOBJECT('LBLF01','LABEL')
     WITH THIS.LBLF01
          .VISIBLE=.T.
          .LEFT=14
          .TOP=27
          .WIDTH=50
          .CAPTION='客戶編號'          
     ENDWITH
     THIS.ADDOBJECT('LBLF011','LABEL')     
     WITH THIS.LBLF011
          .VISIBLE=.T.
          .LEFT=214
          .TOP=27
          .AUTOSIZE=.T.  
     ENDWITH            
     THIS.ADDOBJECT('TXTF01','TEXTBOX')     
     WITH THIS.TXTF01
          .VISIBLE=.T.
          .LEFT=64
          .TOP=23
          .WIDTH=149
          .HEIGHT=25
          .MAXLENGTH=20
     ENDWITH        

  ENDPROC   
  PROCEDURE PAGEFRAME1.PAGE2.INIT
	THIS.ADDOBJECT('LBLF02','LABEL')
	WITH THIS.LBLF02 
	     .VISIBLE=.T. 
         .LEFT=14
         .TOP=13
	     .WIDTH=50
         .CAPTION='出貨月份'
    ENDWITH  
    THIS.ADDOBJECT('LBLF03','LABEL')
    WITH THIS.LBLF03
         .VISIBLE=.T.
	     .LEFT=328
	     .TOP=13
	     .WIDTH=50
	     .CAPTION='測試時間'
	ENDWITH     
    THIS.ADDOBJECT('LBLF04','LABEL')
    WITH THIS.LBLF04
         .VISIBLE=.T.    
         .LEFT=14
         .TOP=39
         .WIDTH=50
         .CAPTION='室內溫度'       
    ENDWITH     
    THIS.ADDOBJECT('LBLF05','LABEL')      
    WITH THIS.LBLF05
         .VISIBLE=.T.       
         .LEFT=328
         .TOP=39
         .WIDTH=50
         .CAPTION='測試壓力'
    ENDWITH     
    THIS.ADDOBJECT('LBLF06','LABEL')
    WITH THIS.LBLF06
         .VISIBLE=.T.           
         .LEFT=14
         .TOP=64
         .WIDTH=50
         .CAPTION='動作秒數'
    ENDWITH     
    THIS.ADDOBJECT('LBLF07','LABEL')
    WITH THIS.LBLF07
         .VISIBLE=.T.           
         .LEFT=14
         .TOP=89
         .WIDTH=50
         .CAPTION='備註說明'
    ENDWITH     
    THIS.ADDOBJECT('LBLF08','LABEL')
    WITH THIS.LBLF08
         .VISIBLE=.T.           
         .LEFT=14
         .TOP=139
         .WIDTH=50
         .CAPTION='安全資料'
    ENDWITH     
    THIS.ADDOBJECT('LBLF081','LABEL')
    WITH THIS.LBLF081
         .VISIBLE=.T.           
         .LEFT=65
         .TOP=139
         .AUTOSIZE=.T.          
    ENDWITH     
	THIS.ADDOBJECT('TXTF02','TEXTBOX')
	WITH THIS.TXTF02 
	     .VISIBLE=.T.     
	     .LEFT=65
	     .TOP=10
	     .WIDTH=70
	     .HEIGHT=25
	     .MAXLENGTH=8
	     .NAME='TXTF02'  
	ENDWITH     
	THIS.ADDOBJECT('TXTF03','TEXTBOX')
	WITH THIS.TXTF03
	     .VISIBLE=.T.
	     .LEFT=379
	     .TOP=10
	     .WIDTH=60
	     .HEIGHT=25
	     .MAXLENGTH=8
	     .NAME='TXTF03'	    
	ENDWITH     
    THIS.ADDOBJECT('TXTF04','TEXTBOX')
    WITH THIS.TXTF04
         .VISIBLE=.T.           
         .LEFT=65
         .TOP=35
         .WIDTH=60
         .INPUTMASK='99999999.99'
         .NAME='TXTF04'
    ENDWITH     
    THIS.ADDOBJECT('TXTF05','TEXTBOX')
    WITH THIS.TXTF05
         .VISIBLE=.T.           
         .LEFT=379
         .TOP=35
         .WIDTH=60
         .HEIGHT=25
         .INPUTMASK='99999999.99'
         .NAME='TXTF05'                            
    ENDWITH     
    THIS.ADDOBJECT('TXTF06','TEXTBOX')
    WITH THIS.TXTF06     
         .VISIBLE=.T.           
         .LEFT=65
         .TOP=60
         .WIDTH=60
         .INPUTMASK='99999999.99'
         .NAME='TXTF06'
    ENDWITH     
    THIS.ADDOBJECT('TXTF07','TEXTBOX')
    WITH THIS.TXTF07
         .VISIBLE=.T.           
         .LEFT=65
         .TOP=85
         .WIDTH=149
         .MAXLENGTH=20
         .NAME='TXTF07'                 
    ENDWITH     
  ENDPROC       
  PROCEDURE PAGEFRAME1.PAGE1.activate
     R=0     
     SELE C01     
     SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE     
     THISFORM.GRID1.ENABLED=.T.          
     THISFORM.GRID1.SETFOCUS     
     IF THISFORM.GRID1.ACTIVEROW=0
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=''
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
     ENDIF   
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F. 
     THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
     THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
     THISFORM.KEY_LIST.ENABLED=.T.
  ENDPROC
  PROCEDURE PAGEFRAME1.PAGE2.activate
     SELE C1600000
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
        THISFORM.PAGEFRAME1.PAGE2.SETALL('VALUE','','TEXTBOX')
        THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=0
        THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE=0
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=0
        THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.LBLF081.CAPTION=''
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限  
     ENDIF   
     R=RECNO('&AREA1')
     THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
     THISFORM.GRID1.ENABLED=.F.     
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)   &&判斷有無新增的權限
     THISFORM.KEY_LIST.ENABLED=.F.     
  ENDPROC   

  PROC INIT              
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限 
       THISFORM.ANS_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.)   &&判斷有無資料分析的權限  
       thisform.setall('height',25,'textbox')
       thisform.setall('height',17,'lanbel')
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
       THISFORM.SHRGROUP.VISIBLE=.F.
       THISFORM.SHRGROUP.ENABLED=.F.
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='客戶編號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='索引減碼'
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶簡稱'       
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C01.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C01.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C01.F05'
       THISFORM.grid2.COLUMN1.HEADER1.CAPTION='出貨月份'
       THISFORM.grid2.COLUMN2.HEADER1.CAPTION='測試時間'
	   THISFORM.grid2.COLUMN3.HEADER1.CAPTION='室內溫度'	   
	   THISFORM.grid2.COLUMN4.HEADER1.CAPTION='測試壓力'
	   THISFORM.grid2.COLUMN5.HEADER1.CAPTION='動作秒數'
	   THISFORM.grid2.COLUMN6.HEADER1.CAPTION='備註說明'       	   
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='C1600000.F08'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='C1600000.F03'
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='C1600000.F04'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE='C1600000.F05'
	   THISFORM.grid2.COLUMN5.CONTROLSOURCE='C1600000.F06'
	   THISFORM.grid2.COLUMN6.CONTROLSOURCE='C1600000.F07'
	   THISFORM.grid2.COLUMN1.WIDTH=70
	   THISFORM.GRID2.COLUMN2.WIDTH=60
       THISFORM.grid2.COLUMN3.WIDTH=60
       THISFORM.GRID2.COLUMN4.WIDTH=60
       THISFORM.GRID2.COLUMN5.WIDTH=60
       THISFORM.GRID2.COLUMN6.WIDTH=149
       THISFORM.grid1.READONLY=.T.       
       THISFORM.grid1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.      
       ENDIF          
       THISFORM.KEY_LIST.DISPLAYVALUE='依料品編號排列'       
       JK='料品編號'
       SELE C01
       GO TOP
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=C01.F01
       THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=C01.F02
  ENDPROC           
  PROC KEY_LIST.INIT 
       with this            
           .additem('依料品編號排列')
       endwith      
  ENDPROC 

  PROC grid1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.page1.TXTF01.VALUE=C01.F01
       THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=C01.F02
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION  
       IF THIS.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
       ENDIF   
       CHK=''     
       thisform.refresh
  ENDPROC     
  PROC grid2.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE=F08
       THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=F03
       THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=F04
       THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE=F05
       THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=F06
       THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE=F07
*       THISFORM.PAGEFRAME1.PAGE2.LBLF081.CAPTION=F08
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE      
  ENDPROC 
  PROCEDURE ORPGROUP.NEW_BOTT.CLICK    
     THISFORM.SETALL('READONLY',.F.,'TEXTBOX')       
     THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
     THISFORM.GRID2.ENABLED=.F. 
     THISFORM.ANS_BOTT.ENABLED=.F.
     THISFORM.PAGEFRAME1.PAGE2.TXTF08.SETFOCUS     
     THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE=DATE()
     THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=TIME()
     THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=0
     THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE=0
     THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=0
     THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE=''
     THISFORM.PAGEFRAME1.PAGE2.LBLF081.CAPTION=''
  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK  
     THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
     THISFORM.GRID2.ENABLED=.F.
     THISFORM.ANS_BOTT.ENABLED=.F.
     THISFORM.PAGEFRAME1.PAGE2.TXTF02.SETFOCUS     
     LOCK()
     IF !RLOCK()
        DO file_impact
        return
     endif
  endproc      
  PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK
    IF EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE)
       =MESSAGEBOX('日期不得空白',0+64,'提示訊息視窗')
       THISFORM.PAGEFRAME1.PAGE2.TXTF02.SETFOCUS
       RETURN
    ENDIF   
    SELE C1600000
    APPEND BLANK
    LOCK()
    IF RLOCK()
       REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE
       REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
       REPLACE F04 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE
       REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE
       REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
       REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE
       REPLACE F08 WITH dtoc(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
    ENDIF
    UNLOCK
    SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+DTOS(THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE)+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
    THISFORM.grid2.ENABLED=.T.    
    THISFORM.grid2.SETFOCUS        
    THISFORM.ORPGROUP.NEW_BOTT.CLICK            
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK
    SELE C1600000
    IF RLOCK()
       REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE
       REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
       REPLACE F04 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE
       REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE
       REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
       REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE
       REPLACE F08 WITH DTOC(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
    ELSE
      DO file_impact
      return   
    ENDIF   
    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK    
  ENDPROC   
  procedure SHRGROUP.ABANDON_BOTT.CLICK
      THISFORM.ANS_BOTT.ENABLED=.T.  
      THISFORM.SHRGROUP.ENABLED=.F.
      THISFORM.SHRGROUP.VISIBLE=.F.      
      THISFORM.ORPGROUP.ENABLED=.T.
      THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.CMDGROUP.ENABLED=.T.
      THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')  
      SELE C1600000
      THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.
      THISFORM.grid2.ENABLED=.T.
      THISFORM.grid2.SETFOCUS
      UNLOCK ALL
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK
       SELE C1600000
       THISFORM.SHRGROUP.ABANDON_BOTT.CLICK       
  ENDPROC     
  PROCEDURE ANS_BOTT.CLICK
     LK=''
     ANS_RANGE=createobject("ANS_RANGE")  
     ANS_RANGE.SHOW    
     IF FLG='0'
         =MESSAGEBOX('無此日期範圍資料',0+48,'警示')
     ENDIF   
     FLG='0'
     
  ENDPROC
ENDDEFINE    
**************************************************刪除資料前檢查是否仍有關聯資料在其他資料表格中
procedure dele_shure 
  SELE C1600000   
     LOCK()
     IF RLOCK()

        DELE
        unlock
        SKIP -1
        IF BOF()
           GO TOP
        ENDIF   
     ELSE
       DO file_impact    
     ENDIF      
  
ENDPROC     
***********************************************列印程序
PROCEDURE PNT_PRC  
     LK=C01.F01   
     ANS_RANGE=createobject("ANS_RANGE")  
     ANS_RANGE.SHOW    
     IF FLG='0'
         =MESSAGEBOX('無此日期範圍資料',0+48,'警示')
     ENDIF   
     FLG='0'
ENDPROC

***********************************************分析資料日期範圍選定
define class ANS_RANGE as form
  autocenter=.t.
  caption='請輸入資料日期範圍'
  fontsize=9
  height=190
  width=375
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='ANS_RANGE'
  add object label1 as label with;
      autosize=.t.,;
      height=25,;
      left=30,;
      top=75,;
      name='label1',;
      CAPTION='起始日期'      
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=25,;
      LEFT=190,;
      TOP=75,;
      NAME='LABEL2',;
      CAPTION='截止日期'          
  add object text1 as textbox with;
      autosize=.t.,;
      top=70,;
      width=70,;
      left=80,;
      height=25,;         
      name='text1'
  add object text2 as textbox with;
      autosize=.t.,;
      top=70,;
      width=70,;
      left=240,;
      height=25,;         
      name='text2'    
  add object cmnd1 as commandbutton with;
      autosize=.t.,;
      left=60,;    
      height=25,;   
      top=144,;  
      width=80,;
      caption='\<Y.執行',;
      TOOLTIPTEXT='確認所輸入的範圍鍵值!快速鍵->ALT+Y',;
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      autosize=.t.,;
      left=200,;
      top=144,;
      height=25,;
      width=80,;
      caption='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      name='cmnd2'
  procedure init 
      THISFORM.TEXT1.VALUE=DATE()
      THISFORM.TEXT2.VALUE=DATE()
  ENDPROC    
   procedure cmnd1.click 
      IF LK=C01.F01
         SELECT C1600000.F01,;
         C1600000.F08,;
         C1600000.F03,;
         C1600000.F04,;          
         C1600000.F05,;
         C1600000.F06,;
         C1600000.F07,;                    
         C01.F02;
         FROM C1600000,C01;     
         WHERE C1600000.F01=LK  AND C01.F01=C1600000.F01 AND C1600000.F08>=THISFORM.TEXT1.VALUE;
         AND  C1600000.F08<=THISFORM.TEXT2.VALUE INTO CURSOR C1600000PNT ORDER BY C1600000.F08,C1600000.F03
         IF _tally >0
            FLG='1'
            REPORT FORM C1600000 PREVIEW
         ENDIF  
      ELSE
         SELECT F09,F08,AVG(F04),AVG(F05),AVG(F06),MIN(F04),MIN(F05),MIN(F06),MAX(F04),MAX(F05),MAX(F06),;
         KMP(F08,'f04'),KMP(F08,'f05'),KMP(F08,'f06');
         GROUP BY F08 INTO CURSOR ANSPT1 FROM C1600000;
         WHERE F09=C01.F01 AND F08>=THISFORM.TEXT1.VALUE AND F08<=THISFORM.TEXT2.VALUE order by C1600000.f08
         IF _TALLY>0
            SELECT F01,f08,F06 INTO CURSOR STDCLC FROM C1600000 WHERE  F09=C01.F01 AND F08>=THISFORM.TEXT1.VALUE AND F08<=THISFORM.TEXT2.VALUE
            LK=ALLTRIM(DTOC(THISFORM.TEXT1.VALUE))+'～'+ALLTRIM(DTOC(THISFORM.TEXT2.VALUE))                  
            SMPL_QTY=createobject("SMPL_QTY")  
            SMPL_QTY.SHOW    

         ENDIF            
      ENDIF   
         THISFORM.RELEASE                         
  endproc
  procedure cmnd2.click
      FLG='1'   
      thisform.release
  endproc      
enddefine            
********************************************************
define class SMPL_QTY as form
  autocenter=.t.
  caption='請輸入以下資料'
  fontsize=9
  height=190
  width=375
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='SMPL_QTY'
  add object label1 as label with;
      autosize=.t.,;
      height=25,;
      left=20,;
      top=25,;
      name='label1',;
      CAPTION='每組樣本量'      
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=25,;
      LEFT=20,;
      TOP=50,;
      NAME='LABEL2',;
      CAPTION='動作秒數平均數與標準差是否為已知'
  ADD OBJECT LABEL3 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=25,;
      LEFT=20,;
      TOP=75,;
      NAME='LABEL3',;
      CAPTION='已知動作秒數平均數'              
  ADD OBJECT LABEL4 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=25,;
      LEFT=20,;
      TOP=100,;
      NAME='LABEL4',;
      CAPTION='已知動作秒數標準差'    
  add object text1 as textbox with;
      autosize=.t.,;
      top=20,;
      width=70,;
      left=80,;
      height=25,;         
      name='text1'
  add OBJECT opgyn as optiongroup with;
      top=45,;
      width=72,; 
      height=25,;
      LEFT=220,;
      buttoncount=2,;
      VALUE=2
  add object text2 as textbox with;
      autosize=.t.,;
      top=70,;
      width=70,;
      left=130,;
      height=25,;
      inputmask='99999999.99',;         
      name='text2'    
    add object text3 as textbox with;
      autosize=.t.,;
      top=95,;
      width=70,;
      left=130,;
      height=25,;         
      inputmask='99999999.99',;
      name='text3'        
  add object cmnd1 as commandbutton with;
      autosize=.t.,;
      left=60,;    
      height=25,;   
      top=144,;  
      width=80,;
      caption='\<Y.執行',;
      TOOLTIPTEXT='確認已輸入正確的資料!快速鍵->ALT+Y',;
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      autosize=.t.,;
      left=200,;
      top=144,;
      height=25,;
      width=80,;
      caption='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      name='cmnd2'
  procedure init 
      thisFORM.OPGYN.option1.LEFT=5
      thisFORM.OPGYN.option1.WIDTH=48
      thisFORM.OPGYN.option2.LEFT=38
      thisFORM.OPGYN.option2.WIDTH=48
      THISFORM.OPGYN.OPTION2.TOP=5
      thisFORM.OPGYN.option1.caption='是'
      thisFORM.OPGYN.option2.caption='否'
      THISFORM.LABEL3.VISIBLE=.F.
      THISFORM.LABEL4.VISIBLE=.F.     
      THISFORM.TEXT2.VISIBLE=.F.
      THISFORM.TEXT3.VISIBLE=.F.
      THISFORM.TEXT1.VALUE=0
      THISFORM.TEXT2.VALUE=0
      THISFORM.TEXT3.VALUE=0
  ENDPROC    
  PROCEDURE OPGYN.CLICK
     DO CASE
        CASE THIS.VALUE=1       
             THISFORM.LABEL3.VISIBLE=.T.
             THISFORM.LABEL4.VISIBLE=.T.     
             THISFORM.TEXT2.VISIBLE=.T.
             THISFORM.TEXT3.VISIBLE=.T.
             THISFORM.TEXT2.SETFOCUS
        CASE THIS.VALUE=2
             THISFORM.LABEL3.VISIBLE=.F.
             THISFORM.LABEL4.VISIBLE=.F.     
             THISFORM.TEXT2.VISIBLE=.F.
             THISFORM.TEXT3.VISIBLE=.F.
   
     ENDCASE
  ENDPROC
   procedure cmnd1.click 
      IF THISFORM.TEXT1.VALUE<2
         =MESSAGEBOX('每組樣本量不得小於2',0+64,'提示訊息視窗')
         THISFORM.TEXT1.SETFOCUS
         RETURN
      ENDIF   
      IF THISFORM.TEXT1.VALUE>25
         =MESSAGEBOX('每組樣本量大於25者本系統無法計算',0+64,'提示訊息視窗')
         THISFORM.TEXT1.SETFOCUS
         RETURN
      ENDIF   
      IF THISFORM.OPGYN.VALUE=1 .AND. THISFORM.TEXT2.VALUE=0
         =MESSAGEBOX('已知平均數不得等於0',0+64,'提示訊息視窗')
         THISFORM.TEXT2.SETFOCUS
         RETURN
      ENDIF   
      FLG='1'   
      RK=THISFORM.TEXT1.VALUE
      CHK_STR=THISFORM.TEXT2.VALUE
      FTR_STR=THISFORM.TEXT3.VALUE
      DO ANSC13        
      THISFORM.RELEASE                         
  endproc
  procedure cmnd2.click
      FLG='1' 
      thisform.release
  endproc      
enddefine            
**********
FUNCTION KMP
PARAMETERS  mdate,mfield  
H=mfield
  SELECT F08,&H FROM C1600000 INTO CURSOR CLC_STD WHere F08=mdate
  CALCULATE STD(&H) TO TMP2
  RETURN TMP2
************
