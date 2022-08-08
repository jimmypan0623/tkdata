***程式名稱:單階材料表
close all
clear 
FLG='0'
FCH=''
CHK=''
CHK_STR=''
FTR_STR=''
UKEF=''
RTG_NB=0
R=0
QTY=0
TQTY=0
AREA1='F01'
AREA2='F02'     
RTE_TABLE='E05'
IF ! USED('A14')
   SELE 0
   USE A14 
ELSE
   SELE A14
ENDIF      
SET FILTER TO F05='*'   &&部門別有生產者才做為加工部門
SET ORDER TO 1
GO TOP
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
SEEK sys_oper+'F01'    
IF !USED('B01')
   SELE 0
   USE B01 
ELSE
   SELE B01
ENDIF
SET ORDER TO 1           
SELE F01,F02 FROM B01 WHERE F97='Y' INTO CURSOR B01_2 ORDER BY F01
CREATE CURSOR B01_1;
(F01 C(43),F02 C(40))
INDE ON F01 TAG B01_11
SET ORDER TO B01_11
SELE B01_2
GO TOP
DO WHILE !EOF()
   SELE B01_1
   APPEND BLANK
   REPL F01 WITH B01_2.F01
   REPL F02 WITH B01_2.F02
   SELE B01_2
   SKIP
ENDDO 
SELE B01_2
USE
IF !USED('F06')
    SELE 0 
    USE F06
ELSE
    SELE F06
ENDIF
SET ORDER TO F061        
IF !USED('E00')
   SELE 0
   USE E00
ELSE
   SELE E00
ENDIF
SET ORDER TO E001 
IF !USED('E05')
   SELE 0
   USE E05
ELSE
   SELE E05
ENDIF
SET ORDER TO E051      
IF !USED('F03')
   SELE 0
   USE F03
ELSE
   SELE F03
ENDIF
SET ORDER TO F031 
IF !USED('F07')
   SELE 0
   USE F07
ELSE
   SELE F07
ENDIF
SET ORDER TO F071      
IF !USED('F01')
   SELE 0
   USE F01
ELSE 
   SELE F01
ENDIF
SET ORDER TO F011     
SET RELA TO F01 INTO B01_1  
SET RELA TO F07 INTO A14 ADDI
IF !USED('F02')
   SELE 0
   USE F02
ELSE 
   SELE F02
ENDIF
SET ORDER TO F021
SET RELA TO F04 INTO B01
SET RELA TO F03 INTO E00 ADDI
F01form=createobject("tkF01")
F01form.show  
define class tkF01 as form
  caption='F01.單階材料表'
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
  NAME='TKF01'
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
  ADD OBJECT PASS_CND AS LABEL WITH;
      LEFT=INT_015*60,;
      TOP=INT_015*525,;
      AUTOSIZE=.T.            
  ADD OBJECT PAGEFRAME1 AS PAGEFRAME WITH;
      LEFT=INT_015*236,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*200,;
      FONTSIZE=INT_015*9,;
      PAGECOUNT=2,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.T.,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;                     
      PAGE1.NAME='PAGE1',;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='表頭資料',;
      PAGE2.NAME='PAGE2',;
      PAGE2.VISIBLE=.T.,;
      PAGE2.CAPTION='表身詳細資料內容'   
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      autocenter=.t.
  ADD OBJECT grid1 AS grid1 WITH;
      columncount=3,;            
      RECORDSOURCE='F01',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3'
  ADD OBJECT grid2 AS grid2 WITH;
      columncount=6,;            
      RECORDSOURCE='F02',; 
      LINKMASTER='F01',;
      RELATIONALEXPR='F01+F02',;
      childorder='F021',;  
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.                             
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.               
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*173,;
      LEFT=INT_015*558          
  ADD OBJECT RTG_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*57,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*40,;      
      CAPTION='\<T.製程',;
      NAME='RTG_BOTT'                       
  ADD OBJECT ISB_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*98,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*65,;      
      CAPTION='\<B.代用料號',;
      NAME='ISB_BOTT'                         
  ADD OBJECT VRS_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*162,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*63,;      
      CAPTION='\<V.整組複製',;
      NAME='VRS_BOTT'                             
  ADD OBJECT ANS_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*15,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*40,;      
      CAPTION='\<A.確認',;
      NAME='ANS_BOTT'                             
  PROCEDURE PAGEFRAME1.PAGE1.INIT 
        THIS.ADDOBJECT('PASS_CND','LABEL')
        WITH THIS.PASS_CND
                   .LEFT=INT_015*14
                  .TOP=INT_015*144
    	          .AUTOSIZE=.T.
    	          .CAPTION='GHDFHT'  
       ENDWITH    
     THIS.ADDOBJECT('LBLF01','LABEL')
     WITH THIS.LBLF01
          .VISIBLE=.T.
          .LEFT=INT_015*14
          .TOP=INT_015*7
          .WIDTH=INT_015*50
          .CAPTION='組成品號'          
     ENDWITH
     THIS.ADDOBJECT('LBLF011','LABEL')
     WITH THIS.LBLF011
          .VISIBLE=.T.
          .LEFT=INT_015*326
          .TOP=INT_015*7          
          .AUTOSIZE=.T.         
     ENDWITH     
     THIS.ADDOBJECT('LBLF02','LABEL')
     WITH THIS.LBLF02
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*32
         .WIDTH=INT_015*50
         .CAPTION='版        本'
     ENDWITH
     THIS.ADDOBJECT('LBLF03','LABEL')
     WITH THIS.LBLF03
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*57
         .WIDTH=INT_015*50
         .CAPTION='圖        號'
     ENDWITH     
     THIS.ADDOBJECT('LBLF06','LABEL')
     WITH THIS.LBLF06
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*82
         .WIDTH=INT_015*50
         .CAPTION='工作說明'
     ENDWITH          
     THIS.ADDOBJECT('LBLF07','LABEL')
     WITH THIS.LBLF07
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*107
         .WIDTH=INT_015*50
         .CAPTION='完工部門'
     ENDWITH               
     THIS.ADDOBJECT('LBLF071','LABEL')
     WITH THIS.LBLF071
         .VISIBLE=.T.
         .LEFT=INT_015*120
         .TOP=INT_015*107
         .AUTOSIZE=.T.
     ENDWITH                    
     THIS.ADDOBJECT('LBLE05','LABEL')
     WITH THIS.LBLE05
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*132
         .WIDTH=INT_015*370
         .CAPTION='生產途程'
     ENDWITH                              
      THIS.ADDOBJECT('RTGBOX','EDITBOX')&&文書編輯EDITBOX顯示生產途程
      WITH THIS.RTGBOX
                  .LEFT=INT_015*64
                 .TOP=INT_015*128
                 .WIDTH=INT_015*470
                 .HEIGHT=INT_015*48
                 .FONTSIZE=INT_015*8
                 .VISIBLE=.T.
                 .READONLY=.T.
                 .MAXLENGTH=500
                 .NAME='RTGBOX'           
      ENDWITH         
     THIS.ADDOBJECT('LBLF05','LABEL')
     WITH THIS.LBLF05
         .VISIBLE=.T.
         .LEFT=INT_015*328
         .TOP=INT_015*32
         .WIDTH=INT_015*370
         .CAPTION='是否確認'
     ENDWITH               
     THIS.ADDOBJECT('LBLF051','LABEL')
     WITH THIS.LBLF051
         .VISIBLE=.T.
         .LEFT=INT_015*379
         .TOP=INT_015*32
         .CAPTION=''
     ENDWITH                    
     THIS.ADDOBJECT('LBLF04','LABEL')
     WITH THIS.LBLF04
         .VISIBLE=.T.
         .LEFT=INT_015*328
         .TOP=INT_015*57
         .WIDTH=INT_015*370
         .CAPTION='安全資料'
     ENDWITH                    
     THIS.ADDOBJECT('LBLF041','LABEL')
     WITH THIS.LBLF041
         .VISIBLE=.T.
         .LEFT=INT_015*379
         .TOP=INT_015*57
         .AUTOSIZE=.T.
         .CAPTION=''
     ENDWITH                
     THIS.ADDOBJECT('LBLF90','LABEL')
     WITH THIS.LBLF90
         .VISIBLE=.T.
         .LEFT=INT_015*328
         .TOP=INT_015*109
         .WIDTH=INT_015*370
         .CAPTION='確認人員'
     ENDWITH                    
     THIS.ADDOBJECT('LBLF901','LABEL')
     WITH THIS.LBLF901
         .VISIBLE=.T.
         .LEFT=INT_015*379
         .TOP=INT_015*109
         .AUTOSIZE=.T.
         .CAPTION=''
     ENDWITH          
     THIS.ADDOBJECT('TXTF01','TXTF01')          
     WITH THIS.TXTF01
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*3
          .WIDTH=INT_015*261
          .HEIGHT=INT_015*25
          .MAXLENGTH=43
     ENDWITH        
     THIS.ADDOBJECT('TXTF02','TEXTBOX')          
     WITH THIS.TXTF02
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*27
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=3
     ENDWITH        
     THIS.ADDOBJECT('TXTF03','TEXTBOX')          
     WITH THIS.TXTF03
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*53
          .WIDTH=INT_015*149
          .HEIGHT=INT_015*25
          .MAXLENGTH=20
     ENDWITH                  
     THIS.ADDOBJECT('TXTF06','TEXTBOX')          
     WITH THIS.TXTF06
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*78
          .WIDTH=INT_015*335
          .HEIGHT=INT_015*25
          .MAXLENGTH=43
     ENDWITH     
     THIS.ADDOBJECT('TXTF07','TXTDPT')          
     WITH THIS.TXTF07
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*103
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=5          
     ENDWITH                       
  ENDPROC   
  PROCEDURE PAGEFRAME1.PAGE2.INIT
	THIS.ADDOBJECT('LBLF04','LABEL')
	WITH THIS.LBLF04
	     .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*7
	     .WIDTH=INT_015*50
         .CAPTION='材料品號'
    ENDWITH  
	THIS.ADDOBJECT('LBLF041','LABEL')
	WITH THIS.LBLF041
	     .VISIBLE=.T.
         .LEFT=INT_015*327
         .TOP=INT_015*7
	     .AUTOSIZE=.T.
         .CAPTION=''
    ENDWITH      
    THIS.ADDOBJECT('LBLF05','LABEL')
    WITH THIS.LBLF05
         .VISIBLE=.T.
         .LEFT=INT_015*14
	     .TOP=INT_015*32
	     .WIDTH=INT_015*50
	     .CAPTION='版        本'
	ENDWITH     
    THIS.ADDOBJECT('LBLF07','LABEL')
    WITH THIS.LBLF07
         .VISIBLE=.T.
	     .LEFT=INT_015*14
	     .TOP=INT_015*57
	     .WIDTH=INT_015*50
	     .CAPTION='使用數量'
	ENDWITH     
	THIS.ADDOBJECT('LBLF071','LABEL')
    WITH THIS.LBLF071
         .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*57
	     .AUTOSIZE=.T.
	     .CAPTION='1/'
	ENDWITH     
    THIS.ADDOBJECT('LBLF12','LABEL')
    WITH THIS.LBLF12
         .VISIBLE=.T.
	     .LEFT=INT_015*174
	     .TOP=INT_015*57
	     .WIDTH=INT_015*50
	     .CAPTION='耗損比率'
	ENDWITH     	
    THIS.ADDOBJECT('LBLF121','LABEL')
    WITH THIS.LBLF121
         .VISIBLE=.T.
	     .LEFT=INT_015*268
	     .TOP=INT_015*57
	     .AUTOSIZE=.T.
	     .CAPTION='%'
	ENDWITH     		
    THIS.ADDOBJECT('LBLF09','LABEL')
    WITH THIS.LBLF09
         .VISIBLE=.T.    
         .LEFT=INT_015*14
         .TOP=INT_015*82
         .WIDTH=INT_015*50
         .CAPTION='廠        牌'       
    ENDWITH        
    THIS.ADDOBJECT('LBLF11','Label')
    WITH THIS.LBLF11
         .VISIBLE=.T.
         .LEFT=INT_015*328
         .TOP=INT_015*57
         .WIDTH=INT_015*50
         .CAPTION='代用料數'         
    ENDWITH     
    THIS.ADDOBJECT('LBLF111','Label')
    WITH THIS.LBLF111
         .VISIBLE=.T.
         .LEFT=INT_015*379
         .TOP=INT_015*57
         .AUTOSIZE=.T.
         .CAPTION=''         
    ENDWITH         
    THIS.ADDOBJECT('LBLF08','LABEL')
    WITH THIS.LBLF08
         .VISIBLE=.T.    
         .LEFT=INT_015*14
         .TOP=INT_015*107
         .WIDTH=INT_015*50
         .CAPTION='備註說明'       
    ENDWITH    
    THIS.ADDOBJECT('LBLF06','LABEL')
    WITH THIS.LBLF06
         .VISIBLE=.T.    
         .LEFT=INT_015*14
         .TOP=INT_015*132
         .WIDTH=INT_015*50
         .CAPTION='供料方式'       
    ENDWITH        
    THIS.ADDOBJECT('LBLF061','LABEL')
    WITH THIS.LBLF061
         .VISIBLE=.T.    
         .LEFT=INT_015*84
         .TOP=INT_015*132
         .AUTOSIZE=.T.
         .CAPTION=''       
    ENDWITH            
    THIS.ADDOBJECT('LBLF03','LABEL')
    WITH THIS.LBLF03
         .VISIBLE=.T.
	     .LEFT=INT_015*328
	     .TOP=INT_015*32
	     .WIDTH=INT_015*50
	     .CAPTION='投入製程'
	ENDWITH     		    
    THIS.ADDOBJECT('LBLF031','LABEL')
    WITH THIS.LBLF031
         .VISIBLE=.T.
	     .LEFT=INT_015*420
	     .TOP=INT_015*32
	     .AUTOSIZE=.T.
	     .CAPTION=''
	ENDWITH     				
    THIS.ADDOBJECT('LBLF10','LABEL')
    WITH THIS.LBLF10
        .VISIBLE=.T.           
        .LEFT=INT_015*328
        .TOP=INT_015*156
        .WIDTH=INT_015*50
        .CAPTION='安全資料'         
    ENDWITH    
    THIS.ADDOBJECT('LBLF101','LABEL')
    WITH THIS.LBLF101
        .VISIBLE=.T.           
        .LEFT=INT_015*380
        .TOP=INT_015*156
        .AUTOSIZE=.T.
    ENDWITH    
	THIS.ADDOBJECT('TXTF04','TXTF04')
	WITH THIS.TXTF04      
	     .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*3
	     .WIDTH=INT_015*261
	     .HEIGHT=INT_015*25
	     .MAXLENGTH=43
	     .NAME='TXTF04'  
	ENDWITH     
	THIS.ADDOBJECT('TXTF05','TEXTBOX')
	WITH THIS.TXTF05
	     .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*28
	     .WIDTH=INT_015*30
	     .HEIGHT=INT_015*25
	     .MAXLENGTH=3
	     .NAME='TXTF05'	    
	ENDWITH     
    THIS.ADDOBJECT('TXTF07','TEXTBOX')
    WITH THIS.TXTF07
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*53
         .WIDTH=INT_015*80
         .HEIGHT=INT_015*25
         .INPUTMASK='99999999.9999'
         .ALIGNMENT=1
         .NAME='TXTF07'
    ENDWITH     
    THIS.ADDOBJECT('TXTF12','TEXTBOX')
    WITH THIS.TXTF12
         .VISIBLE=.T.           
         .LEFT=INT_015*225
         .TOP=INT_015*53
         .WIDTH=INT_015*42
         .HEIGHT=INT_015*25
         .INPUTMASK='99.99'
         .ALIGNMENT=1
         .NAME='TXTF12'
    ENDWITH         
    THIS.ADDOBJECT('TXTF09','TEXTBOX')
    WITH THIS.TXTF09
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*78
         .WIDTH=INT_015*224
         .HEIGHT=INT_015*25
         .MAXLENGTH=30
         .NAME='TXTF09'
    ENDWITH             
    THIS.ADDOBJECT('TXTF08','TEXTBOX')
    WITH THIS.TXTF08
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*103
         .WIDTH=INT_015*223         
         .HEIGHT=INT_015*25
         .MAXLENGTH=30
         .NAME='TXTF08'                                     
    ENDWITH                 
    THIS.ADDOBJECT('SPLY_TYPE','SPLY_TYPE')    
    THIS.ADDOBJECT('TXTF06','TEXTBOX')
    WITH THIS.TXTF06
         .VISIBLE=.T.           
         .LEFT=INT_015*65
         .TOP=INT_015*128
         .WIDTH=INT_015*15
         .HEIGHT=INT_015*25
         .MAXLENGTH=1 
         .NAME='TXTF06'
    ENDWITH                 
    THIS.ADDOBJECT('TXTF03','TXTF03')
    WITH THIS.TXTF03
         .VISIBLE=.T.           
         .LEFT=INT_015*379
         .TOP=INT_015*28
         .WIDTH=INT_015*40
         .HEIGHT=INT_015*25
         .MAXLENGTH=4
         .NAME='TXTF03'
    ENDWITH             
  ENDPROC       
  PROCEDURE PAGEFRAME1.PAGE1.activate
     R=0
     SELE F01
     IF THISFORM.SHRGROUP.VISIBLE=.F.
        THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(F01.F05=.F.,RGB(255,0,0),'')","COLUMN")              
        THISFORM.GRID1.ENABLED=.T.   
        THISFORM.GRID1.SETFOCUS   
        THISFORM.KEY_LIST.ENABLED=.T.
        THISFORM.ISB_BOTT.ENABLED=.F.
     ELSE
       THISFORM.KEY_LIST.ENABLED=.F.   
     ENDIF                
    IF THISFORM.GRID1.ACTIVEROW=0 .AND. thisform.SHRGROUP.ENABLED=.F.
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
        THISFORM.ANS_BOTT.ENABLED=.F.
        THISFORM.RTG_BOTT.ENABLED=.F.
        THISFORM.VRS_BOTT.ENABLED=.F.        
        THISFORM.PAGEFRAME1.PAGE1.SETALL('VALUE','','TEXTBOX')
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF041.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF071.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.      
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)    &&判斷有無列印的權限 
        THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
        THISFORM.RTG_BOTT.ENABLED=.T.        
        THISFORM.VRS_BOTT.ENABLED=.T.                
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.
     ENDIF   

  ENDPROC
  PROCEDURE PAGEFRAME1.PAGE2.activate
     SELE F02
     IF THISFORM.SHRGROUP.ENABLED=.F.
        THISFORM.RTG_BOTT.ENABLED=.F.
        THISFORM.VRS_BOTT.ENABLED=.F.                
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
        THISFORM.ISB_BOTT.ENABLED=.F.        
        THISFORM.PAGEFRAME1.PAGE2.SETALL('VALUE','','TEXTBOX')
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.LBLF041.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE2.LBLF101.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION='0'     
        THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=''   
        THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=''
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)   &&判斷有無修改的權限
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)  &&&&判斷有無列印的權限 
        THISFORM.ISB_BOTT.ENABLED=!THISFORM.SHRGROUP.ENABLED         
        THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED        &&判斷有無確認的權限
     ENDIF   
     R=RECNO('&AREA1')
     THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
     THISFORM.GRID1.ENABLED=.F.     
     THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)    &&判斷有無新增的權限
     THISFORM.KEY_LIST.ENABLED=.F.     
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
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(F01.F05=.F.,RGB(255,0,0),'')","COLUMN")                
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'HEADER')         
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='組成品號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='版本'    
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='圖        號'          
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='F01.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='F01.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='F01.F03'
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*30
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149       
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'HEADER')       
       THISFORM.grid2.COLUMN1.HEADER1.CAPTION='材料編號'       
       THISFORM.grid2.COLUMN2.HEADER1.CAPTION='品名規格'              
       THISFORM.grid2.COLUMN3.HEADER1.CAPTION='版本'
	   THISFORM.grid2.COLUMN4.HEADER1.CAPTION='使用數量'	  
	   THISFORM.grid2.COLUMN5.HEADER1.CAPTION='使用廠牌'
	   THISFORM.grid2.COLUMN6.HEADER1.CAPTION='投入製程'
	   THISFORM.GRID2.LINKMASTER='F01'
       THISFORM.GRID2.RELATIONALEXPR='F01+F02'
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='F02.F04'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'	   
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='F02.F05'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE="IIF(IIF(SEEK(F02.F04,'B01'),B01.F39,'')='C','1/'+ALLTRIM(STR(F02.F07)),TRANSFORM(F02.F07,'@R 9999999.999'))"
	   THISFORM.grid2.COLUMN5.CONTROLSOURCE='F02.F09'
  	   THISFORM.grid2.COLUMN6.CONTROLSOURCE='E00.F04'
	   THISFORM.grid2.COLUMN1.WIDTH=INT_015*149
       THISFORM.grid2.COLUMN2.WIDTH=INT_015*149
       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*30
       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*70
       THISFORM.GRID2.COLUMN4.ALIGNMENT=2
       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*80
       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*40
       THISFORM.grid1.READONLY=.T.      
       THISFORM.grid2.READONLY=.T.            
       THISFORM.grid1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.ANS_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限 
          THISFORM.ANS_BOTT.ENABLED=!F01.F05  .AND. IIF(A02.F11='*',.T.,.F.) .AND. !THISFORM.SHRGROUP.ENABLED
       ENDIF          
       THISFORM.KEY_LIST.DISPLAYVALUE='依組成品號排列'      
       JK='組成品號'
       SELE F01
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=F01.F01       
  ENDPROC               
  PROC KEY_LIST.INIT 
       with this 
           .additem('依組成品號排列')      
           .additem('依圖    號排列')                 
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE F01
       DO CASE
          CASE THIS.DISPLAYVALUE='依組成品號排列'
               set order to F011
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='組成品號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='F01.F01'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='版本'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='F01.F02'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*30
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='圖        號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='F01.F03'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149
               THISFORM.PAGEFRAME1.PAGE1.LBLF01.TOP=INT_015*13
               THISFORM.PAGEFRAME1.PAGE1.LBLF011.TOP=INT_015*13
               THISFORM.PAGEFRAME1.PAGE1.TXTF01.TOP=INT_015*9
               THISFORM.PAGEFRAME1.PAGE1.LBLF03.TOP=INT_015*63
               THISFORM.PAGEFRAME1.PAGE1.TXTF03.TOP=INT_015*59
          CASE THIS.DISPLAYVALUE='依圖    號排列'
               set order to F012
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='圖        號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='F01.F03'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='版本'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='F01.F02'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*30
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='組成品號'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='F01.F01'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149        
               THISFORM.PAGEFRAME1.PAGE1.LBLF03.TOP=INT_015*13
               THISFORM.PAGEFRAME1.PAGE1.TXTF03.TOP=INT_015*9                
               THISFORM.PAGEFRAME1.PAGE1.LBLF01.TOP=INT_015*63
               THISFORM.PAGEFRAME1.PAGE1.LBLF011.TOP=INT_015*63
               THISFORM.PAGEFRAME1.PAGE1.TXTF01.TOP=INT_015*59                                                                                                           
       ENDCASE 
       THISFORM.grid1.SETFOCUS
  ENDPROC        
  PROC grid1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex       
       THISFORM.PASS_CND.BACKCOLOR=IIF(F01.F05=.T.,RGB(0,255,0),RGB(255,0,0))            
       THISFORM.PASS_CND.CAPTION=IIF(F01.F05=.T.,'    已確認    ','    未確認    ')       
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=F01.F01
       THISFORM.PAGEFRAME1.PAGE1.lblf011.caption=B01_1.F02
       THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=F01.F02
       THISFORM.PAGEFRAME1.PAGE1.TXTF03.VALUE=F01.F03
       THISFORM.PAGEFRAME1.PAGE1.TXTF06.VALUE=F01.F06
       THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE=F01.F07       
       THISFORM.PAGEFRAME1.PAGE1.lblf041.caption=F01.F04
       THISFORM.PAGEFRAME1.PAGE1.lblf071.caption=A14.F02           
       THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=IIF(F01.F05=.T.,'是','否')
       THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=F01.F90
       THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE=ROUTE(F01.F01,F01.F02)       
       THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED        
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION  
       UKEF=THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE
       RTG_NB=OCCURS('》',THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE)              
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
       THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=F02.F04
       THISFORM.PAGEFRAME1.PAGE2.LBLF041.CAPTION=B01.F02
       THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=E00.F04
       THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE=F02.F05
       THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=F02.F06
       THISFORM.PAGEFRAME1.PAGE2.LBLF061.CAPTION=IIF(F02.F06='1','廠內發料',IIF(F02.F06='2','廠商自備','客戶提供'))
       THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE=F02.F07
       THISFORM.PAGEFRAME1.PAGE2.TXTF07.LEFT=INT_015*(65+IIF(B01.F39='C',20,0))
       THISFORM.PAGEFRAME1.PAGE2.TXTF07.INPUTMASK=IIF(B01.F39='C','99999999','999999.9999')
       THISFORM.PAGEFRAME1.PAGE2.TXTF07.WIDTH=INT_015*IIF(B01.F39='C',65,80)
       THISFORM.PAGEFRAME1.PAGE2.TXTF12.VISIBLE=IIF(B01.F39='B',.T.,.F.)       
       THISFORM.PAGEFRAME1.PAGE2.LBLF12.VISIBLE=IIF(B01.F39='B',.T.,.F.)              
       THISFORM.PAGEFRAME1.PAGE2.LBLF121.VISIBLE=IIF(B01.F39='B',.T.,.F.)                     
       THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE=F02.F12       
       THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=F02.F03
       THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=F02.F09
       THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE=F02.F08     
       THISFORM.PAGEFRAME1.PAGE2.LBLF101.CAPTION=F02.F10       
       THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION=ALLTRIM(STR(F02.F11))
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
       if thisform.pageframe1.activepage=1
          thisform.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ENDIF   
*       thisform.refresh
  ENDPROC 
  PROCEDURE ORPGROUP.NEW_BOTT.CLICK    
    THISFORM.KEY_LIST.ENABLED=.F. 
    THISFORM.ANS_BOTT.ENABLED=.F.
    THISFORM.ISB_BOTT.ENABLED=.F.
    THISFORM.RTG_BOTT.ENABLED=.F.
    THISFORM.VRS_BOTT.ENABLED=.F.            
    IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1       
       THISFORM.GRID2.RECORDSOURCE=''
       THISFORM.GRID2.CHILDORDER=''
       THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')     &&處理表頭新增
       THISFORM.GRID1.ENABLED=.F.   
       THISFORM.GRID2.ENABLED=.F.   
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.SETFOCUS
        THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF03.VALUE=''     
        THISFORM.PAGEFRAME1.PAGE1.TXTF06.VALUE=''   
        THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE=''           
        THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF041.CAPTION=''        
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF051.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.TXTF07.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE1.LBLF071.CAPTION=''
        THISFORM.PAGEFRAME1.PAGE1.LBLF901.CAPTION=''
     ELSE
        THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.              &&處理表身新增
        THISFORM.GRID2.ENABLED=.F.   
        THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX')     
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS        
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE=''         
        THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=''          
        THISFORM.PAGEFRAME1.PAGE2.LBLF041.CAPTION=''                            
        THISFORM.PAGEFRAME1.PAGE2.TXTF04.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE=0        
        THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE=0                
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.F.
        THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION='0'
        THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VISIBLE=.T.                        
        THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VALUE=1       
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.VISIBLE=.F.
        THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE='1'
        THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE=''
        THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE=''              
        THISFORM.PAGEFRAME1.PAGE2.LBLF101.CAPTION='' 
        THISFORM.PAGEFRAME1.PAGE2.LBLF111.CAPTION='0'                       
     ENDIF   

  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
     THISFORM.KEY_LIST.ENABLED=.F. 
     THISFORM.ANS_BOTT.ENABLED=.F.     
     THISFORM.RTG_BOTT.ENABLED=.F.               
     THISFORM.VRS_BOTT.ENABLED=.F.             
     THISFORM.ISB_BOTT.ENABLED=.F.                    
     THISFORM.GRID1.ENABLED=.F.   
     THISFORM.GRID2.ENABLED=.F.        
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
        SELE F01
        LOCK()
        IF !RLOCK()
           DO file_impact
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
*           return
        ELSE
           THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')     &&處理表頭修改
           THISFORM.PAGEFRAME1.PAGE1.TXTF02.READONLY=.T.
           THISFORM.PAGEFRAME1.PAGE1.TXTF07.READONLY=.F.
           THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.
           THISFORM.PAGEFRAME1.PAGE1.TXTF03.SETFOCUS
        endif        
     ELSE                                                      &&處理表身修改            
        SELE F01
        RLOCK()
        IF !RLOCK()
           DO file_impact
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK            
        ENDIF
        SELE F02
        RLOCK()
        IF RLOCK() 
           THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.F.,'TEXTBOX') 
           THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
           THISFORM.PAGEFRAME1.PAGE2.TXTF05.READONLY=.T.
           THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VISIBLE=.T.  
           THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VALUE=VAL(F02.F06)
           THISFORM.PAGEFRAME1.PAGE2.TXTF06.VISIBLE=.F.           
           THISFORM.PAGEFRAME1.PAGE2.TXTF07.SETFOCUS
           THISFORM.PAGEFRAME1.PAGE2.TXTF09.READONLY=!EMPTY(F05)
           THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.F.
        ELSE  
           DO file_impact 
           unlock all
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           
       endif             
     ENDIF  
  endproc    
  PROC ORPGROUP.DEL_BOTT.CLICK                    &&刪除資料程序
       IF messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6
          if thisform.pageframe1.activepage=1          &&刪除整組BOM
             SELE F01             
                PO_NO=F01.F01+F01.F02               
                SELE RECNO() FROM F07 WHERE F07.F01+F07.F02=PO_NO  INTO ARRAY BK4
                JJ4=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK4)
                       JJ4=JJ4+ALLTRIM(STR(BK4(I)))+','
                   ENDFOR    
                   KK4=STUFF(JJ4,RAT(',',JJ4,1),1,'')              
                   RLOCK(KK4,'F07')
                   LL4=RLOCK(KK4,'F07')
                ELSE
                   LL4=.T.   
                ENDIF                                                    
                SELE RECNO() FROM F06 WHERE F06.F01+F06.F02=PO_NO  INTO ARRAY BK3             
                JJ3=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK3)
                       JJ3=JJ3+ALLTRIM(STR(BK3(I)))+','
                   ENDFOR    
                   KK3=STUFF(JJ3,RAT(',',JJ3,1),1,'')              
                   RLOCK(KK3,'F06')
                   LL3=RLOCK(KK3,'F06')
                ELSE
                   LL3=.T.   
                ENDIF                                    
                SELE RECNO() FROM F02 WHERE F02.F01+F02.F02=PO_NO  INTO ARRAY BK1              
                JJ1=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
                   ENDFOR    
                   KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
                   RLOCK(KK1,'F02')
                   LL1=RLOCK(KK1,'F02')
                ELSE
                   LL1=.T.   
                ENDIF                   
                SELE RECNO() FROM E05 WHERE E05.F01+E05.F02=PO_NO  INTO ARRAY BK2            
                JJ2=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK2)
                       JJ2=JJ2+ALLTRIM(STR(BK2(I)))+','
                   ENDFOR    
                   KK2=STUFF(JJ2,RAT(',',JJ2,1),1,'')              
                   RLOCK(KK2,'E05')
                   LL2=RLOCK(KK2,'E05')
                ELSE
                   LL2=.T.   
                ENDIF                      
                SELE RECNO() FROM F03 WHERE F03.F01+F03.F02=PO_NO  INTO ARRAY BK5            
                JJ5=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK5)
                       JJ5=JJ5+ALLTRIM(STR(BK5(I)))+','
                   ENDFOR    
                   KK5=STUFF(JJ5,RAT(',',JJ5,1),1,'')              
                   RLOCK(KK5,'F03')
                   LL5=RLOCK(KK5,'F03')
                ELSE
                   LL5=.T.   
                ENDIF                                                   
                IF LL1 .AND. LL2 .AND. LL3 .AND. LL4 .AND. LL5 .AND. RLOCK('F01')
                   SELE F06
                   DELE FROM F06 WHERE F01+F02=PO_NO     
                   SELE E05
                   DELE FROM E05 WHERE F01+F02=PO_NO                                                
                   SELE F02    
                   DELE FROM F02 WHERE F01+F02=PO_NO             
                   IF F01.F05=.T.                         &&如果此組BOM已確認過
                      SELE F03
                      DELE FROM F03 WHERE F01+F02=PO_NO
                      SELE F07
                      DELE FROM F07 WHERE F01+F02=PO_NO
                      FLG='1'
                      THISFORM.ANS_BOTT.CLICK                    &&BOM重算       
                      FLG='0'
                      DELE FROM F03 WHERE F01+F02=PO_NO
                      DELE FROM F07 WHERE F01+F02=PO_NO               
                   ENDIF   
                   SELE F01
                   DELE
                   SKIP -1
                   IF BOF()
                      GO TOP
                   ENDIF   
                   UNLOCK ALL
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
                THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)    &&判斷有無刪除的權限
                THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限
                THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED                           
             ENDIF     
          ELSE                                                                &&刪除單筆資料  
              SELE F02
              PO_NO=F01+F02+F03+F04+F05
                SELE RECNO() FROM F06 WHERE F01+F02+F03+F04+F05=PO_NO  INTO ARRAY BK3             
                JJ3=''                
                IF _TALLY>0   
                   FOR I= 1 TO ALEN(BK3)
                       JJ3=JJ3+ALLTRIM(STR(BK3(I)))+','
                   ENDFOR    
                   KK3=STUFF(JJ3,RAT(',',JJ3,1),1,'')              
                   RLOCK(KK3,'F06')
                   LL3=RLOCK(KK3,'F06')
                ELSE
                   LL3=.T.   
                ENDIF                        
              IF LL3 .AND. RLOCK('F02')                                          
                 SELE F06
                 DELE FROM F06 WHERE F01+F02+F03+F04+F05=F02.F01+F02.F02+F02.F03+F02.F04+F02.F05      
                 SELE F02                                                                   
                 DELE
                 SKIP -1
                 IF BOF()
                    GO TOP
                 ENDIF   
                 IF F01.F05=.T.                     &&如果已確認過        
                    SELE F03
                    DELE FROM F03 WHERE F03.F01+F03.F02=F01.F01+F01.F02
                    SELE F07
                    DELE FROM F07 WHERE F07.F01+F07.F02=F01.F01+F01.F02     
                    FLG='1'            
                    THISFORM.ANS_BOTT.CLICK                    &&BOM重算
                    FLG='0'
                 ENDIF                    
              ELSE
                 DO file_impact                  
                 RETURN
              ENDIF                   
              THISFORM.GRID2.SETFOCUS
              IF THISFORM.GRID2.ACTIVEROW=0         &&如果單身被刪空           
                 SELE E05
                 DELE FROM E05 WHERE E05.F01+E05.F02=F01.F01+F01.F02
                 IF F01.F05=.T.                     &&如果已確認過        
                    SELE F03
                    DELE FROM F03 WHERE F03.F01+F03.F02=F01.F01+F01.F02
                    SELE F07
                    DELE FROM F07 WHERE F07.F01+F07.F02=F01.F01+F01.F02      
                    PO_NO=F01.F01+F01.F02
                    FLG='1'
                    THISFORM.ANS_BOTT.CLICK                    &&BOM重算       
                    FLG='0'
                    DELE FROM F03 WHERE F01+F02=PO_NO
                    DELE FROM F07 WHERE F01+F02=PO_NO                          
                 ENDIF                                     
                 
                 SELE F01                        
                 DELE                               &&連表頭也要刪除
                 SKIP -1
                 IF BOF()
                    GO TOP
                 ENDIF   
                 THISFORM.PAGEFRAME1.ACTIVEPAGE=1 
                 THISFORM.REFRESH          
*                 THISFORM.ANS_BOTT.ENABLED=IIF(A02.F11='*',.T.,.F.)                 
                 SELE F02   
              ENDIF
          ENDIF
       ENDIF
       CHK_STR=''                    
       RELEASE ALL LIKE BK*   
  ENDPROC  
  PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK               &&新增確認
    IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1                               &&表頭新增確認
       THISFORM.GRID2.RECORDSOURCE='F02'
       THISFORM.GRID2.CHILDORDER='F021'
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='F02.F04'
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'
	   THISFORM.grid2.COLUMN3.CONTROLSOURCE='F02.F05'
	   THISFORM.grid2.COLUMN4.CONTROLSOURCE="IIF(IIF(SEEK(F02.F04,'B01'),B01.F39,'')='C','1/'+ALLTRIM(STR(F02.F07)),TRANSFORM(F02.F07,'@R 9999999.999'))"
  	   THISFORM.grid2.COLUMN5.CONTROLSOURCE='F02.F09'
   	   THISFORM.grid2.COLUMN6.CONTROLSOURCE='E00.F04'       
       SELE F01 
       COD=SYS(21)
       SET ORDER TO F011
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE,'F01')=.T.
          =MESSAGEBOX('此版本組成品號已建立資料!,請直接去表身詳細資料內容新增',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.SETFOCUS
          RETURN  
       ENDIF
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE,'B01_1')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE)
          =MESSAGEBOX('物料基本主檔中無此編號請先建立或是此料號在基本主檔中請改設成展開!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF01.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF
       IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE)
          =MESSAGEBOX('版本不得空白!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF       
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE,'A14')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE)
          =MESSAGEBOX('工作部門主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF07.SETFOCUS
          SET ORDER TO &COD
          RETURN            
       ENDIF              
       SELE F01          
       APPEND BLANK
       LOCK()
       IF RLOCK()
          REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE          
          REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE   
          REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF03.VALUE   
          REPLACE F04 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
          REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF06.VALUE       
          REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE
       ELSE
          DO file_impact                             
       ENDIF
*       UNLOCK
       SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE      
       SET ORDER TO &COD
       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.
       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.  
       if messagebox('是否建立此組成品製程?',4+32+0,'請確認') = 6	 
          SELE E05   
          THISFORM.RTG_BOTT.CLICK
          SELE F01
       endif                
       THISFORM.PAGEFRAME1.ACTIVEPAGE=2
       THISFORM.ORPGROUP.NEW_BOTT.CLICK
       THISFORM.REFRESH
    ELSE         &&表身新增確認    
       IF SEEK(THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE,'B01')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE)
          =MESSAGEBOX('物料基本主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS
          RETURN        
       ENDIF    
       IF seek(thisform.pageframe1.page1.txtf01.value+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF04.value+THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE,'F02')=.T.
          =MESSAGEBOX('此材料編號已建立!',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF04.SETFOCUS
          RETURN
       ENDIF
        IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE)
            IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'E05','E052')=.F.
                 =MESSAGEBOX('此加工品無此製程編號!',0+48,'提示訊息視窗')
                 THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
                 RETURN
             ENDIF     
       ENDIF             
       IF LEN(ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE))=4 .AND. THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VALUE=2
          =MESSAGEBOX('廠內加工之供料方式不得為廠商自備!',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.SETFOCUS
          RETURN
       ENDIF
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE<=0
          =MESSAGEBOX('用料數量不得小於等於零',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF07.SETFOCUS
          RETURN
       ENDIF   
       IF THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE<0
          =MESSAGEBOX('耗損比率不得小於零',0+48,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE2.TXTF12.SETFOCUS
          RETURN
       ENDIF          
         IF F01.F05=.T.
            SELE F03 
            DELE FROM F03 WHERE F03.F01+F03.F02=F01.F01+F01.F02
            SELE F07
            DELE FROM F07 WHERE F07.F01+F07.F01=F01.F01+F01.F02     
            SELE F01
            REPL F05 WITH .F. 
            THISFORM.PASS_CND.BACKCOLOR=RGB(255,0,0)
            THISFORM.PASS_CND.CAPTION='    未確認    '
         ENDIF                
       SELE F02
       APPEND BLANK       
       rlock()
          REPLACE F01 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
          REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
          REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
          REPLACE F04 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF04.VALUE
          REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE
          REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
          REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE
          REPLACE F12 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE          
          REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE
          REPLACE F09 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE
          REPLACE F10 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
      UNLOCK ALL         
      thisform.grid2.enabled=.t.
      thisform.grid2.setfocus  
      if messagebox('是否建立代用料?',4+32+256,'請確認') = 6	 
         SELE F06    
         THISFORM.ISB_BOTT.CLICK
         SELE F02  
      endif         
      seek thisform.pageframe1.page1.txtf01.value+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF04.value+THISFORM.PAGEFRAME1.PAGE2.TXTF05.VALUE                     
      thisform.grid2.enabled=.t.
      thisform.grid2.setfocus         
*!*	      SELE F01
*!*	      REPL F05 WITH .F.
      THISFORM.ORPGROUP.NEW_BOTT.CLICK      
    ENDIF  
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK                          &&修改確認
     IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1                            &&表頭修改確認
        IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE)
           =MESSAGEBOX('版本不得空白!',0+48,'提示訊息視窗')
           THISFORM.PAGEFRAME1.PAGE1.TXTF02.SETFOCUS
           RETURN
        ENDIF
       IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE,'A14')=.F. .OR. EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE)
          =MESSAGEBOX('工作部門主檔中無此編號請先建立!',0+64,'提示訊息視窗')
          THISFORM.PAGEFRAME1.PAGE1.TXTF07.SETFOCUS
          RETURN            
       ENDIF                      
       IF THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE<>F01.F07   &&若完工部門修改
            IF messagebox('每一製程之加工部門是否與最後完工部門相同?',4+32+0,'請確認') = 6
                IF LEN(ALLTRIM(F01.F07))=5 .AND. LEN(ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE))=4  &&若是原由尾外加工變廠內自製
                    SELE F02
                    SEEK F01.F01+F01.F02
                    IF FOUND()
                        DO WHILE F01+F02=F01.F01+F01.F02
                               IF F06='2'
                                   =MESSAGEBOX(F04+'之供料方式請先修改為不為廠商自備!',0+48,'提示訊息視窗')
                                   SELE F01
                                  THISFORM.PAGEFRAME1.PAGE1.TXTF07.SETFOCUS
                                   RETURN
                              ENDIF               
                              SKIP
                        ENDDO 
                    ENDIF
                 ENDIF            
                 SELECT E05                 
                 REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE FOR E05.F01+E05.F02=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
            ELSE
                SELECT E05
                SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE+PADL(ALLTRIM(STR(RTG_NB)),3,'0')
                IF FOUND()
                    IF LEN(ALLTRIM(F01.F07))=5 .AND. LEN(ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE))=4  &&若是原由尾外加工變廠內自製                
                        SELECT F02
                        SEEK F01.F01+F02.F02
                        IF FOUND()
                            DO WHILE F01+F02=F01.F01+F01.F02
                                   IF F06='2' AND F03=E05.F07
                                       =MESSAGEBOX(F04+'之供料方式請先修改為不為廠商自備!',0+48,'提示訊息視窗')            
                                       RETURN                       
                                   ENDIF
                                   SKIP
                            ENDDO
                        ENDIF           
                    ENDIF    
                    SELECT E05                    
                    REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE
                ENDIF
            ENDIF
       ENDIF
        SELE F01
        REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
        REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF03.VALUE
        REPLACE F04 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                   
        REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF06.VALUE    
        REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE
     ELSE                                                            &&表身修改確認
        IF THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE<=0
           =MESSAGEBOX('用料數量不得小於等於零',0+48,'提示訊息視窗')
           THISFORM.PAGEFRAME1.PAGE2.TXTF07.SETFOCUS
           RETURN
        ENDIF        
        IF !EMPTY(THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE)
            IF SEEK(THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE+THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE,'E05','E052')=.F.
                 =MESSAGEBOX('此加工品無此製程編號!',0+48,'提示訊息視窗')
                 THISFORM.PAGEFRAME1.PAGE2.TXTF03.SETFOCUS
                 RETURN
             ENDIF     
       ENDIF             
        IF THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE<0
           =MESSAGEBOX(' 耗損比率不得小於零',0+48,'提示訊息視窗')
           THISFORM.PAGEFRAME1.PAGE2.TXTF12.SETFOCUS
           RETURN
        ENDIF                
        IF THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VALUE=2
             IF LEN(ALLTRIM(IIF(SEEK(F02.F01+F02.F02+F02.F03,'E05','E052'),E05.F08,'')))=4
                 =MESSAGEBOX('廠內加工之供料方式不得為廠商自備!',0+48,'提示訊息視窗')
                 THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.SETFOCUS
                 RETURN
             ENDIF   
        ENDIF        
        SELE F06 
        SEEK F02.F01+F02.F02+F02.F03+F02.F04+F02.F05                
        IF FOUND()
             DO WHILE F01+F02+F04+F05=F02.F01+F02.F02+F02.F04+F02.F05 
                    REPL F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE 
                    SKIP
             ENDDO                         
        ENDIF    
        SELE F02
        REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF03.VALUE
        REPLACE F06 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE
        REPLACE F07 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF07.VALUE
        REPLACE F12 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE        
        REPLACE F08 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF08.VALUE 
        REPLACE F09 WITH THISFORM.PAGEFRAME1.PAGE2.TXTF09.VALUE 
        REPLACE F10 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
         IF F01.F05=.T.
            SELE F03 
            DELE FROM F03 WHERE F03.F01+F03.F02=F01.F01+F01.F02
            SELE F07
            DELE FROM F07 WHERE F07.F01+F07.F01=F01.F01+F01.F02     
            SELE F01
            REPL F05 WITH .F. 
            THISFORM.PASS_CND.BACKCOLOR=RGB(255,0,0)
            THISFORM.PASS_CND.CAPTION='    未確認    '
         ENDIF                 
     ENDIF           
 *      ELSE
*         DO file_impact
*          return   
*       ENDIF   
     UNLOCK ALL
     THISFORM.SHRGROUP.ABANDON_BOTT.CLICK    
  ENDPROC   
  procedure SHRGROUP.ABANDON_BOTT.CLICK
      THISFORM.SHRGROUP.ENABLED=.F.
      THISFORM.SHRGROUP.VISIBLE=.F.      
      THISFORM.ORPGROUP.ENABLED=.T.
      THISFORM.ANS_BOTT.ENABLED=.T.      
      THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.CMDGROUP.ENABLED=.T.
      THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE2.SETALL('READONLY',.T.,'TEXTBOX')  
      THISFORM.PAGEFRAME1.PAGE1.TXTF01.READONLY=.T.  
      THISFORM.PAGEFRAME1.PAGE1.TXTF07.READONLY=.T.
      THISFORM.PAGEFRAME1.PAGE2.TXTF04.READONLY=.T. 
      THISFORM.PAGEFRAME1.PAGE2.TXTF03.READONLY=.T. 
      THISFORM.PAGEFRAME1.PAGE2.TXTF06.VISIBLE=.T.
      THISFORM.PAGEFRAME1.PAGE2.SPLY_TYPE.VISIBLE=.F.
      THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.      
      THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.      
      THISFORM.grid2.ENABLED=.T.                         
      THISFORM.grid1.ENABLED=.T.           
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1  
         THISFORM.RTG_BOTT.ENABLED=.T.                           
         THISFORM.VRS_BOTT.ENABLED=.T.                 
         THISFORM.GRID1.SETFOCUS
         THISFORM.KEY_LIST.ENABLED=.T.    
      ELSE      
         THISFORM.GRID1.ENABLED=.F.
         THISFORM.GRID2.SETFOCUS      
         THISFORM.ISB_BOTT.ENABLED=.T.                                 
         IF THISFORM.GRID2.ACTIVEROW=0
             SELECT E05
             DELETE FROM E05 WHERE F01+F02=F01.F01+F01.F02
            SELE F01
            DELE
            THISFORM.PAGEFRAME1.ACTIVEPAGE=1
            THISFORM.REFRESH                          
         ENDIF   
      ENDIF      
      UNLOCK ALL
      *THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED                   
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
         IF THISFORM.GRID2.RECORDSOURCE<>'F02'
            THISFORM.GRID2.RECORDSOURCE='F02'
            THISFORM.GRID2.LINKMASTER='F01'  
            THISFORM.GRID2.CHILDORDER='F021'                               
            THISFORM.GRID2.RELATIONALEXPR='F01+F02'
   	        THISFORM.grid2.COLUMN1.CONTROLSOURCE='F02.F04'
	        THISFORM.grid2.COLUMN2.CONTROLSOURCE='B01.F02'	   
	        THISFORM.grid2.COLUMN3.CONTROLSOURCE='F02.F05'
	        THISFORM.grid2.COLUMN4.CONTROLSOURCE="IIF(IIF(SEEK(F02.F04,'B01'),B01.F39,'')='C','1/'+ALLTRIM(STR(F02.F07)),TRANSFORM(F02.F07,'@R 9999999.999'))"
	        THISFORM.grid2.COLUMN5.CONTROLSOURCE='F02.F09'
  	        THISFORM.grid2.COLUMN6.CONTROLSOURCE='E00.F04'   
*  	        THISFORM.GRID2.REFRESH      
	     ENDIF   
         SELE F01
         IF THISFORM.KEY_LIST.DISPLAYVALUE='依組成品號排列'
            SET ORDER TO F011
         ELSE
            SET ORDER TO F012
         ENDIF   	     
      ENDIF  
      THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC   
  PROCEDURE RTG_BOTT.CLICK                   &&製程按鈕
         SELE E05.F07,E00.F04 FROM E05,E00 WHERE E05.F01+E05.F02=F01.F01+F01.F02 AND E00.F01=E05.F07 ORDER BY E05.F06 INTO CURSOR E05_1  
         IF _TALLY>0
            SELECT F01,F04 FROM E00 WHERE F01 NOT IN (SELECT F07 FROM E05 WHERE E05.F01+E05.F02=F01.F01+F01.F02) ORDER BY F01 INTO CURSOR E00_1
         ELSE
            SELECT F01,F04 FROM E00  ORDER BY F01 INTO CURSOR E00_1
         ENDIF   
         ROUTE_SEEK=createobject("ROUTE_SEEK")  
         ROUTE_SEEK.SHOW                
             FLG='0'                    
             AREA1='F01'
             SELE F01       
  ENDPROC  
  PROCEDURE ISB_BOTT.CLICK
         SUB_SEEK=createobject("SUB_SEEK")  
         SUB_SEEK.SHOW                
             FLG='0'                    
             AREA1='F01'
             SELE F02          
  ENDPROC
  PROCEDURE VRS_BOTT.CLICK            &&版本複製
            FTR_STR=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE 
            FLG=''
            VRS_COPY=createobject("VRS_COPY")  
            VRS_COPY.SHOW   
            IF !EMPTY(FLG)
               SELE * FROM F01 WHERE F01+F02=FTR_STR INTO CURSOR VRS1 
               SELE F01            
               APPEND BLANK
               REPLACE  F01 WITH LEFT(FLG,43)
               REPLACE  F02 WITH RIGHT(FLG,3)
               REPLACE  F03 WITH VRS1.F03
               REPLACE F04 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
               REPLACE  F06 WITH VRS1.F06
               REPLACE F07 WITH VRS1.F07
               SELE VRS1 
               USE
               SELE * FROM E05 WHERE F01+F02=FTR_STR INTO CURSOR VRS2
               IF _TALLY>0
                  SELE VRS2
                  GO TOP
                  DO WHILE !EOF()
                         SELE E05
                         APPEND BLANK
                         REPLACE  F01 WITH LEFT(FLG,43)
                         REPLACE  F02 WITH RIGHT(FLG,3)
                         REPLACE  F04 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                         REPLACE F05 WITH VRS2.F05
                         REPLACE  F06 WITH VRS2.F06
                         REPLACE  F07 WITH VRS2.F07
                         REPLACE  F08 WITH VRS2.F08
                         REPLACE  F09 WITH VRS2.F09
                         REPLACE  F10 WITH VRS2.F10
                         REPLACE  F11 WITH VRS2.F11
                         REPLACE  F12 WITH VRS2.F12
                         REPLACE  F13 WITH VRS2.F13
                         REPLACE  F99 WITH VRS2.F99
                         SELE VRS2
                         SKIP
                   ENDDO  
                ENDIF   
                SELE VRS2
                USE
                SELE * FROM F06 WHERE F01+F02=FTR_STR INTO CURSOR VRS3 
                IF _TALLY>0
                   SELE VRS3
                   GO TOP 
                   DO WHILE !EOF()
                          SELE F06 
                          APPEND BLANK
                          REPLACE  F01 WITH LEFT(FLG,43)
                          REPLACE  F02 WITH RIGHT(FLG,3)
                          REPLACE  F03 WITH VRS3.F03
                          REPLACE  F04 WITH VRS3.F04
                          REPLACE  F05 WITH VRS3.F05
                          REPLACE  F06 WITH VRS3.F06
                          REPLACE  F07 WITH VRS3.F07
                          REPLACE  F08 WITH VRS3.F08
                          SELE VRS3
                          SKIP                   
                   ENDDO
                ENDIF
                SELE * FROM F02 WHERE F01+F02=FTR_STR INTO CURSOR VRS4 ORDER BY F04,F05
                IF _TALLY>0
                   SELE VRS4
                   GO TOP
                   DO WHILE !EOF()
                      SELE F02
                      APPEND BLANK
                      REPLACE  F01 WITH LEFT(FLG,43)
                      REPLACE  F02 WITH RIGHT(FLG,3)
                      REPLACE  F03 WITH VRS4.F03
                      REPLACE  F04 WITH VRS4.F04
                      REPLACE  F05 WITH VRS4.F05
                      REPLACE  F06 WITH VRS4.F06
                      REPLACE  F07 WITH VRS4.F07
                      REPLACE  F08 WITH VRS4.F08
                      REPLACE  F09 WITH VRS4.F09
                      REPLACE  F10 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                      REPLACE  F11 WITH VRS4.F11
                      REPLACE F12 WITH VRS4.F12
                      SELE VRS4
                      SKIP
                   ENDDO
                ENDIF
                SELE VRS4
                USE     
                SELE F01
                THISFORM.GRID1.SETFOCUS      
               THISFORM.PAGEFRAME1.ACTIVEPAGE=2                
            ELSE
               SELE F01
               THISFORM.GRID1.SETFOCUS                
            ENDIF             
            FTR_STR=''
            FLG=''              
  ENDPROC
  PROCEDURE ANS_BOTT.CLICK        &&BOM 確認程序
         IF MESSAGEBOX('是否確認資料無誤?',4+32+256,'請確認') = 6	 
            IF !USED('C20')
               SELECT 0
               USE C20
            ELSE
               SELECT C20
            ENDIF
            SET ORDER TO 1          
            SET FILTER TO VAL(C20.F03)>0
              S=0
              SELECT F02
              SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
              DO WHILE F01+F02=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE  &&開始搜尋每一筆之製程欄位是否為空白
                      IF EMPTY(F03) 
                           S=S+1  &&若有則加到變數 S
                           EXIT   &&馬上離開此迴圈
                      ENDIF     
                     SKIP
              ENDDO
              IF S>0   &&如果其中有一筆材料沒打上投入製程編號
                  =MESSAGEBOX('料號:'+ALLTRIM(F02.F04)+'之投入製程尚未建立!',0+48,'提示訊息視窗')
                  THISFORM.PAGEFRAME1.ACTIVEPAGE=2
                  SELE F02
                  THISFORM.GRID2.SETFOCUS
                  RETURN    &&回去原來的畫面且不做BOM確認動作   
             ENDIF     
            SELE E05           &&製程檔
            SET ORDER TO E052  &&料號+版本+製程編號
            PRW=''
            SELECT F02  &&材料表身檔
            SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE       &&從材料表身檔中的每筆料耗所投入製程去尋找E05是否沒有對應到
            DO WHILE F01+F02=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE+THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE    
                   SELECT E05
                   SEEK F02.F01+F02.F02+F02.F03
                   IF !FOUND()     
                        PRW=F02.F03          &&若未對應到的話將該製程編號記起來到 PRW 變數中        
                   ENDIF
                   SELECT F02
                   SKIP
            ENDDO        
            IF !EMPTY(PRW)  &&若沒對應到
                  =MESSAGEBOX('製程:'+PRW+'    '+IIF(SEEK(PRW,'E00'),E00.F04,'')+'可能已被刪除,請於表頭建立!',0+48,'提示訊息視窗') 
                 THISFORM.PAGEFRAME1.ACTIVEPAGE=1 &&先回表頭再說
                  SELE F01
                  THISFORM.GRID1.SETFOCUS
               RETURN                                            
            ENDIF            
            IF FLG<>'1'     &&如果此程序不是承接刪除程序才做以下判斷
                SELECT E05
                SEEK F01.F01+F01.F02
                IF !FOUND() 
                    =MESSAGEBOX('請先建立製程再做確認!',0+48,'修改後再做確認')
                    THISFORM.PAGEFRAME1.ACTIVEPAGE=1
                     SELE F01
                     THISFORM.GRID1.SETFOCUS
                    RETURN
                ELSE
                    REPLACE  F99 WITH 0 FOR F01+F02=F01.F01+F01.F02
                    SELE F02
                    SEEK F01.F01+F01.F02
                    DO WHILE F01+F02=F01.F01+F01.F02
                           SELE E05
                           SEEK F02.F01+F02.F02+F02.F03
                           IF FOUND()
                               DO CASE  &&製程型態
                                      CASE F05='1' &&一次或不一定  
                                                 IF F99=0
                                                     REPLACE F99 WITH 1
                                                ENDIF
                                      CASE F05='2' &&看使用幾項就是幾次
                                                  REPLACE  F99 WITH F99+1 
                                      CASE F05='3'  &&使用數量即是製程加工次數
                                            REPLACE  F99 WITH F99+F02.F07
                                      CASE F05='4'  &&批次生產
                                           IF F99=0
                                              REPLACE F99 WITH IIF(SEEK(F02.F01,'C20'),VAL(C20.F03),1)
                                           ENDIF   
                               ENDCASE                  
                           ENDIF   
                          SELE F02
                          SKIP
                    ENDDO       
                    SELE E05
                    SEEK F01.F01+F01.F02
                    IF FOUND()
                       DO WHILE F01+F02=F01.F01+F01.F02
                          IF F99=0
                             IF F05='4'
                                REPLACE F99 WITH VAL(IIF(SEEK(F01.F01,'C20'),C20.F03,'1')) 
                             ELSE 
                                REPLACE F99 WITH 1  
                             ENDIF
                          ENDIF 
                          SKIP
                       ENDDO
                    ENDIF
               ENDIF               
            ENDIF   
            SELE E05
            SET ORDER TO E051   &&料號+版本+製程順序
            CREATE CURSOR BM4;
            (F01 C(43),F02 C(3),F03 C(50),F04 C(43),F05 C(3),F06 N(14,4),F07 N(14,4),F08 C(4))          
            CREATE CURSOR BM3; &&父階料號+父階版本+子階料號+子階版本
            (F01 C(43),F02 C(3),F04 C(43),F05 C(3),F07 N(8,4),F08 N(8,4),F09 N(2),F10 N(4),F11 N(4),F90 N(2))        &&記錄搜詢軌跡檔
            INDE ON F04+F05+F01+F02 TAG BM31
            SET ORDER TO BM31
            CREATE CURSOR BM2; &&子階料號+子階版本+父階料號+父階版本+階次
            (F01 C(43),F02 C(3),F04 C(43),F05 C(3),F08 C(2)) &&祖宗八代紀錄檔
            INDE ON F01+F02+F08 TAG BM21 DESC  &&子階料號+版本降冪排列
*!*	            INDE ON F04+F05 TAG BM22
            CREATE CURSOR BM1;
            (F01 C(43),F02 C(3),F03 C(4),F04 C(43),F05 C(3),F06 N(4),F07 N(8,4),F08 N(2),F09 N(5,2),F10 C(1),F90 N(2))
            SET RELA TO F03 INTO E00

               PT=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE  &&成品料號
               OP=THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE  &&成品料號版本

            SELE BM2  &&祖宗檔
            APPEND BLANK   &&建立初始值
            REPL F01 WITH PT
            REPL F02 WITH OP
            REPL F04 WITH PT
            REPL F05 WITH OP
            REPL F08 WITH '00'
            PRIORITY=0       &&統計階層變數
            FATHER=PT        &&父階料號
            EGG=OP           &&父階版本 
            SON=''
            BIRD=''
            FLP=0
            TQTY=1
            SELE F02
            set order to F022   &&往回找爸爸
            SEEK FATHER+EGG
            IF FOUND()   &&如果找到
               DO WHILE FLP>=0        &&開始執行回圈
                  PRIORITY=PRIORITY+1  
                  IF PRIORITY>INT_181
                     =MESSAGEBOX('此組BOM可能有父階子階錯亂的問題!請重新檢查修正後再做確認!',0+48,'提示訊息視窗')
                     RETURN
                     EXIT
                  ENDIF   
                  SELE BM2    &&記錄來源
                  APPEND BLANK      
                  REPL F01 WITH PT       &&要找的完成品號寫進來
                  REPL F02 WITH OP
                  REPL F04 WITH F02.F01   &&將祖父寫進來 
                  REPL F05 WITH F02.F02
                  REPL F08 WITH PADL(PRIORITY,2,'0')   
                  FATHER=F04  &&寫完之後阿公變爸爸
                  EGG=F05
                  SELE BM3    &&並紀錄在軌跡檔
                  APPEND BLANK
                  REPL F01 WITH F02.F04  &&子
                  REPL F02 WITH F02.F05
                  REPL F04 WITH F02.F01  &&父
                  REPL F05 WITH F02.F02
                  REPL F09 WITH PRIORITY         
                  SELE F02
                  SEEK FATHER+EGG &&再回頭去找祖先
                  IF FOUND()
*!*	                     SON=F02.F01   &&找到就將角色對調                     
*!*	                     BIRD=F02.F02  
                     FATHER=F02.F04
                     EGG=F02.F05
                  ELSE  
                     SEEK BM3.F01+BM3.F02+BM3.F04+BM3.F05  &&由軌跡檔去找F02(逆)
                     IF FOUND()
                        SKIP
                        IF F04+F05=BM3.F01+BM3.F02                         
                           PRIORITY=PRIORITY-1  
                           SELECT BM3
                           DELETE                               
                        ELSE
                           FLP=PRIORITY
                           DO WHILE .T.
                              FLP=FLP-1 
                              IF FLP<0
                                 EXIT
                              ENDIF   
                              SELE BM3   
                              SON=BM3.F01
                              BIRD=BM3.F02
                              DELE  
                              SEEK SON+BIRD
                              IF FOUND()     
                                 PRIORITY=BM3.F09-1   
                                 SON=BM3.F01
                                 BIRD=BM3.F02                        
                                 FATHER=BM3.F04
                                 EGG=BM3.F05
                                 DELE                                    
                                 SELE F02
                                 SEEK SON+BIRD+FATHER+EGG
                                 SKIP
                                 IF F04+F05=SON+BIRD
                                    EXIT            
                                 ENDIF
                              ENDIF        
                           ENDDO                  
                        ENDIF                         
                     ENDIF
                  ENDIF   
               ENDDO
            ENDIF   
            DELE FROM F07 WHERE F01+F02=ANY(SELE F04+F05 FROM BM2)
            DELE FROM F03 WHERE F01+F02=ANY(SELE F04+F05 FROM BM2)  &&樹狀材料表
            SELE BM2   
               *****************************************************************************************************開始尋找子孫
               SET ORDER TO BM21
               GO TOP
               DO WHILE !EOF()
                  PT=F04
                  OP=F05
                  QTY=IIF(SEEK(PT,'B01'),B01.F28,1)
                  SELE BM1   
                  APPEND BLANK 
                  REPL F01 WITH PT  &&設定初始值
                  REPL F02 WITH OP
                  REPL F04 WITH PT
                  REPL F05 WITH OP
                  REPL F06 WITH QTY  &&lead time
                  REPL F07 WITH 1
                  REPL F08 WITH 0
                  PRIORITY=0
                  FATHER=PT
                  EGG=OP
                  SON=''
                  BIRD=''
                  FLP=0
                  TQTY=1
                  LN=1
                  IF !EMPTY(EGG)  &&如果版本空白就不用往下找了
                     SELE F02
                     set order to F021                     
                     SEEK FATHER+EGG
                     IF FOUND()
                        DO WHILE FLP>=0    
                           PRIORITY=PRIORITY+1
                           QTY=QTY+B01.F28+B01.F31
                           IF PRIORITY>INT_181  &&階層數已經大於系統設定最大階層數
                              =MESSAGEBOX('此組BOM可能有父階子階錯亂的問題!請重新檢查修正後再做確認!',0+48,'提示訊息視窗')
                              RETURN
                              EXIT
                           ENDIF   
                           SELE BM1
                           APPEND BLANK
                           REPLACE  F01 WITH PT
                           REPLACE  F02 WITH OP
                           REPLACE  F03 WITH F02.F03
                           REPLACE  F04 WITH F02.F04
                           REPLACE  F05 WITH F02.F05
                           REPLACE F06 WITH QTY
                           REPLACE F07 WITH TQTY*F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE F08 WITH PRIORITY        
                           REPLACE F09 WITH F02.F12
                           REPLACE F10 WITH F02.F06
                           REPLACE F90 WITH LN
                           FATHER=F04
                           EGG=F05
                           TQTY=F07             
                           SELE BM4
                           APPEND BLANK
                           REPLACE F01 WITH BM1.F01
                           REPLACE F02 WITH BM1.F02
                           REPLACE F03 WITH PADL(BM1.F08,2,'0')+PADL(ALLTRIM(STR(BM1.F90)),2,'0')+SPACE(6*BM1.F08)+IIF(BM1.F08<>0,'╠═','')                       
                           REPLACE F04 WITH BM1.F04
                           REPLACE F05 WITH BM1.F05
                           REPLACE F06 WITH F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE F07 WITH BM1.F07
                           REPLACE F08 WITH BM1.F03
                           SELE BM3                       
                           APPEND BLANK                                 
                           REPLACE  F01 WITH F02.F01
                           REPLACE  F02 WITH F02.F02
                           REPLACE  F04 WITH F02.F04
                           REPLACE  F05 WITH F02.F05
                           REPLACE  F07 WITH F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE  F08 WITH BM1.F07
                           REPLACE  F09 WITH PRIORITY                   
                           REPLACE  F10 WITH B01.F28+B01.F31         
                           REPLACE  F11 WITH QTY
                           REPLACE F90 WITH LN
                           SELE F02
                           SEEK FATHER+EGG
                           IF FOUND()
                              FATHER=F02.F04
                              EGG=F02.F05
                              LN=1
                           ELSE      &&找不到下一階 
                              SEEK BM3.F01+BM3.F02+BM3.F04+BM3.F05  &&回上一階
                              IF FOUND()
                                 SKIP      &&往下移動一筆
                                 IF F01+F02=BM3.F01+BM3.F02  &&如果該筆父階也相同則參數要回到原階時的狀況(表示有本房兄弟旁系)
                                    PRIORITY=PRIORITY-1      &&往上退一階層
                                    TQTY=BM3.F08/BM3.F07     &&找到父親用量
                                    QTY=BM3.F11-BM3.F10      &&LEAD TIME
                                    LN=LN+1
                                    SELECT BM3               &&將軌跡刪除
                                    DELETE 
                                 ELSE                        &&如果無本房旁系
                                    FLP=PRIORITY
                                    DO WHILE .T.             
                                       FLP=FLP-1 
                                       IF FLP<0
                                          EXIT
                                       ENDIF   
                                       SELE BM3   
                                       SON=BM3.F01
                                       BIRD=BM3.F02
                                       DELETE   &&找到原軌跡紀錄就刪除
                                       IF !EMPTY(BIRD)  &&如果是不是空蛋(版次空白)
                                          SEEK SON+BIRD &&再從子階開始找
                                          IF FOUND()     
                                             PRIORITY=BM3.F09-1   
                                             TQTY=BM3.F08/BM3.F07
                                             QTY=BM3.F11-BM3.F10
                                             SON=BM3.F01
                                             BIRD=BM3.F02                        
                                             FATHER=BM3.F04
                                             EGG=BM3.F05
                                             LN=BM3.F90
                                             DELETE                                     
                                             SELE F02
                                             SEEK SON+BIRD+FATHER+EGG
                                             SKIP
                                             IF F01+F02=SON+BIRD
                                                LN=LN+1
                                                EXIT            
                                             ENDIF
                                          ENDIF   
                                       ENDIF               &&子階版次是否空白判斷完
                                    ENDDO                  
                                 ENDIF                         
                              ENDIF
                            ENDIF   
                        ENDDO
                     ENDIF        
                  ENDIF     &&版本空白判斷忠
               SELE BM2
               SKIP
            ENDDO   
            SELE BM2
            USE 
            SELE BM3
            USE            
            SELE BM1
            GO TOP
            KM=''
            DO WHILE !EOF()
                  SELE F07
                  APPEND BLANK
                  REPL F01 WITH BM1.F01
                  REPL F02 WITH BM1.F02
                  REPL F03 WITH BM1.F03
                  REPL F04 WITH BM1.F04
                  REPL F05 WITH BM1.F05
                  REPL F06 WITH BM1.F06
                  REPL F07 WITH BM1.F07
                  REPL F08 WITH PADL(ALLTRIM(STR(BM1.F08)),2,'0')
                  IF BM1.F09>0
                      REPLACE F09 WITH BM1.F09
                  ENDIF    
                  REPLACE F10 WITH BM1.F10
               SELE BM1
               SKIP
           ENDDO   
           SELE BM4
           DO WHILE RECNO()>1
              IF LEFT(F03,4)='0000'
                 SKIP -1
              ENDIF   
              L=AT('╠═',F03)   
              B=RECNO()
              SKIP -1
              IF AT('╠═',F03)>L  &&如果前筆的位置較後
                 DO WHILE AT('╠═',F03)>L 
                    H=STUFF(F03,L,2,'║')           
        	        REPLACE  F03 WITH H
        	        SKIP -1
    	         ENDDO     	     
              ENDIF
              GO B-1
           ENDDO
           SELE BM4
           
           GO BOTT
           IF LEFT(F03,4)<>'0000'
              L=AT('╠═',F03)   
              H=STUFF(F03,L,2,'╚')    
              REPL F03 WITH H      
           ENDIF       
              DO WHILE RECNO()>1
                 KM=LEFT(F03,4)
                 SKIP -1
                 IF LEFT(F03,4)>KM
                    L=AT('╠═',F03)   
                    H=STUFF(F03,L,2,'╚')    
                    REPL F03 WITH H             
                 ENDIF   
              ENDDO
              SELE BM4
              GO TOP
              DO WHILE !EOF()
                 L=AT('╠═',F03)   
                 SKIP
                 IF SUBSTR(F03,L,2)=SPACE(2)
                    SKIP -1
                    H=STUFF(F03,L,2,'╚')    
                    REPL F03 WITH H  
                    SKIP                 
                 ENDIF   
              ENDDO  
              SELE BM4
 *****************************             
              GO BOTTOM   &&刪掉多餘的豎線
              DO WHILE !BOF()
                 TMP_STK_NO=F01
                 L=ATC('║',F03)
                 DO WHILE F01=TMP_STK_NO AND !BOF()                
                     IF L>0    
                       IF ATC('║',F03)=L
                          H=STUFF(F03,L,2,SPACE(2))           
        	              REPLACE  F03 WITH H   
        	           ELSE
        	              IF ATC('╠═',F03)=L
        	                H=STUFF(F03,L,2,'╚')    
                            REPLACE  F03 WITH H   
                            L=0  &&自宮
                          ENDIF             	                   	                 	        
                       ENDIF
                     ENDIF
                     SKIP -1
                    
                 ENDDO   
                
              ENDDO
******************************************              
              GO TOP
              DO WHILE !EOF()
                 SELECT F03
                 APPEND BLANK
                 REPLACE F01 WITH BM4.F01
                 REPLACE F02 WITH BM4.F02
                 REPLACE F03 WITH BM4.F03
                 REPLACE F04 WITH BM4.F04
                 REPLACE F05 WITH BM4.F05
                 REPLACE F06 WITH BM4.F06
                 REPLACE F07 WITH BM4.F07
                 REPLACE F08 WITH BM4.F08
                 SELE BM4
                 SKIP
              ENDDO
              SELE BM4
              USE
              SELE BM1
              USE

              SELE F01
              REPLACE F90 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')                            
              REPLACE F05 WITH .T.
              THISFORM.PASS_CND.BACKCOLOR=IIF(F01.F05=.T.,RGB(0,255,0),RGB(255,0,0))            
              THISFORM.PASS_CND.CAPTION=IIF(F01.F05=.T.,'    已確認    ','    未確認    ')      
              SELECT C20
              USE              
              IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1
                  SELE F01 
                 THISFORM.GRID1.SETFOCUS
              ELSE
                 SELE F02 
                 THISFORM.GRID2.SETFOCUS
              ENDIF     
              THISFORM.ANS_BOTT.ENABLED=.F.
           ENDIF     
  ENDPROC  
enddefine    
***********************************************列印程序
PROCEDURE PNT_PRC  
     IF !USED('A13')
         SELE 0
         USE A13
     ELSE
        SELE A13
     ENDIF   
     SET ORDER TO 1
     SEEK 'F01'            
     LK1=F01.F01   &&完成品號
     LK2=F01.F02   &&版次
     LK3=F01.F03   &&圖號
     LK4=F01.F05   &&是否已確認
     LK5=F01.F06   &&工作說明
     LK6=F01.F07   &&預設完工部門
      SELE   F02.F03 RTO_NO,;  &&製程編號
            F02.F04  MTRL_NO,;        &&材料編號	            
            F02.F05  VRS_NO,;        && 版本                      
            F02.F07  UNIT_QTU,;         &&用量  
            F02.F08   RMRK,;        &&備註說明
            F02.F06   SPLY_TYPE,;        &&供料方式
            F02.F09   BRAND1,;        &&原料廠牌
            F02.F11  SUB_NO,;        &&代用料數
            F02.F12  LSE_RTE,;        &&耗損率
            B01.F02  MTRL_NAME,;    &&材料品名  
            F06.F06   BRAND2,;   &&代用料廠牌
            F06.F07 SUB_MTRL,; &&代用料號
            F06.F08 SUB_VRS;  &&代用料版本
            FROM F02,B01,F06 WHERE F02.F01+F02.F02=LK1+LK2 AND B01.F01=F02.F04 AND F06.F01+F06.F02+F06.F04+F06.F05=LK1+LK2+F02.F04+F02.F05; 
UNION;
     SELE   F02.F03 RTO_NO,;  &&製程編號
            F02.F04  MTRL_NO,;        &&材料編號	            
            F02.F05  VRS_NO,;        && 版本                      
            F02.F07  UNIT_QTU,;         &&用量
            F02.F08   RMRK,;        &&備註說明
            F02.F06   SPLY_TYPE,;        &&供料方式
            F02.F09   BRAND1,;        &&原料廠牌
            F02.F11  SUB_NO,;        &&代用料數
            F02.F12  LSE_RTE,;        &&耗損率
            B01.F02  MTRL_NAME,;    &&材料品名  
            SPACE(30) BRAND2,;
            SPACE(43) SUB_MTRL,;
            SPACE(3) SUB_VRS;            
            FROM F02,B01 WHERE F02.F01+F02.F02=LK1+LK2 AND B01.F01=F02.F04 AND F02.F04+F02.F05 NOT IN(SELE  F04+F05 FROM F06 WHERE F01+F02=LK1+LK2);
            INTO CURSOR BILL1
            SELECT * FROM BILL1 ORDER BY MTRL_NO,SUB_MTRL NOWAIT
     IF _tally >0
        REPORT FORM ALLTRIM(INT_116)+'F01' PREVIEW        
*         report form F02TEST to print prompt 
     ENDIF  
ENDPROC
***************************************************特殊欄位------------------完工部門
DEFINE CLASS TXTDPT AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5 
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
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
       THISFORM.PAGEFRAME1.PAGE1.LBLF071.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')  
  ENDPROC
ENDDEFINE  
************************************************特殊輸入欄位的設定----代用料號
DEFINE CLASS TXTF07 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02,F97 FROM B01 INTO CURSOR MTR ORDER BY F01 WHERE  LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
       ELSE
          SELECT F01,F02,F97 FROM B01 INTO CURSOR MTR ORDER BY F01  
       ENDIF        
       IF _TALLY>0
          MAN_SEEK=createobject("MTR_SEEK")  
          MAN_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'     
       ENDIF    
    ENDIF  
    THISFORM.LBLF071.CAPTION=IIF(SEEK(THIS.VALUE,'B01'),B01.F02,'')   
    THISFORM.TXTF08.READONLY=IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F97,'')='Y',.F.,.T.)   
  ENDPROC
ENDDEFINE
************************************************特殊輸入欄位的設定----製程輸入
DEFINE CLASS TXTF03 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=4
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT E05.F07,E00.F04 FROM E05,E00 INTO CURSOR RTE ORDER BY E05.F06 WHERE  LEFT(E05.F07,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1) AND E05.F01+E05.F02=F01.F01+F01.F02 AND E00.F01=E05.F07
       ELSE
          SELECT E05.F07,E00.F04 FROM E05,E00 INTO CURSOR RTE ORDER BY E05.F06 WHERE E05.F01+E05.F02=F01.F01+F01.F02 AND E00.F01=E05.F07
       ENDIF        
       IF _TALLY>0
          RTE_SEEK=createobject("RTE_SEEK")  
          RTE_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'     
       ENDIF    
    ENDIF  
    THISFORM.PAGEFRAME1.PAGE2.LBLF031.CAPTION=IIF(SEEK(THIS.VALUE,'E00'),E00.F04,'')   
  ENDPROC
ENDDEFINE
************************************************特殊輸入欄位的設定----表身料號
DEFINE CLASS TXTF04 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02,F97 FROM B01 INTO CURSOR MTR ORDER BY F01 WHERE  LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
       ELSE
          SELECT F01,F02,F97 FROM B01 INTO CURSOR MTR ORDER BY F01  
       ENDIF        
       IF _TALLY>0
          MAN_SEEK=createobject("MTR_SEEK")  
          MAN_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'     
       ENDIF    
    ENDIF  
    THISFORM.PAGEFRAME1.PAGE2.LBLF041.CAPTION=IIF(SEEK(THIS.VALUE,'B01'),B01.F02,'')   
    THISFORM.PAGEFRAME1.PAGE2.TXTF05.READONLY=IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F97,'')='Y',.F.,.T.)   
    THISFORM.PAGEFRAME1.PAGE2.TXTF09.READONLY=IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F97,'')='Y',.T.,.F.)      
    
    THISFORM.PAGEFRAME1.PAGE2.LBLF12.VISIBLE= IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='B',.T.,.F.)      
    THISFORM.PAGEFRAME1.PAGE2.LBLF121.VISIBLE= IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='B',.T.,.F.)          
    THISFORM.PAGEFRAME1.PAGE2.TXTF12.VISIBLE= IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='B',.T.,.F.)              
    THISFORM.PAGEFRAME1.PAGE2.TXTF12.VALUE= IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='B',INT_182,0)                  
    THISFORM.PAGEFRAME1.PAGE2.TXTF07.LEFT=INT_015*IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='C',85,65)      
    THISFORM.PAGEFRAME1.PAGE2.TXTF07.WIDTH= INT_015*IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='C',65,80)  
    THISFORM.PAGEFRAME1.PAGE2.TXTF07.INPUTMASK= IIF(IIF(SEEK(THIS.VALUE,'B01'),B01.F39,'')='C','9999999','999999.9999')  
  ENDPROC
ENDDEFINE

***************************************************特殊輸入欄位設定-----------表頭料號
DEFINE CLASS TXTF01 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
       ELSE
          SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01    
       ENDIF        
       IF _TALLY>0
          MAN_SEEK=createobject("MTR_SEEK")  
          MAN_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'     
       ENDIF    
    ENDIF  
    THISFORM.PAGEFRAME1.PAGE1.LBLF011.CAPTION=IIF(SEEK(THIS.VALUE,'B01_1'),B01_1.F02,'')      
  ENDPROC
ENDDEFINE
*********************************************************************代用料號搜尋畫面
define class SUB_SEEK as form
  autocenter=.t.
  caption='代用料號'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*370
  WIDTH=INT_015*666
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='SUB_SEEK'
  ADD OBJECT SHAPE1 AS SHAPE1 WITH;
      LEFT=INT_015*3,;
      WIDTH=INT_015*661,;
      HEIGHT=INT_015*90
  ADD OBJECT SHAPE2 AS SHAPE2 WITH;    
      LEFT=INT_015*3,;
      TOP=INT_015*92,;
      WIDTH=INT_015*661,;
      HEIGHT=INT_015*240
  ADD OBJECT SHAPE3 AS SHAPE3 WITH;
      LEFT=INT_015*3,;
      TOP=INT_015*334,;
      WIDTH=INT_015*661,;
      HEIGHT=INT_015*35    
  add object cmdgroup as cmdgroup with;      
      TOP=INT_015*299,;
      LEFT=INT_015*122
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      TOP=INT_015*335,;
      LEFT=INT_015*122
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      VISIBLE=.F.    
  ADD OBJECT LBLF07 AS LABEL WITH;
      TOP=INT_015*8,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='代用料號'
  ADD OBJECT LBLF071 AS LABEL WITH;
      TOP=INT_015*8,;
      LEFT=INT_015*317,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT LBLF08 AS LABEL WITH;
      TOP=INT_015*33,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='版        本'
  ADD OBJECT LBLF06 AS LABEL WITH;
      TOP=INT_015*58,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*50,;
      CAPTION='廠        牌'      
  ADD OBJECT TXTF07 AS TXTF07 WITH;
      TOP=INT_015*4,;
      LEFT=INT_015*56,;
      WIDTH=INT_015*261,;
      HEIGHT=INT_015*25,;
      MAXLENGTH=43,;
      NAME='TXTF07'
  ADD OBJECT TXTF08 AS TEXTBOX WITH;
      TOP=INT_015*29,;
      LEFT=INT_015*56,;
      WIDTH=INT_015*30,;
      HEIGHT=INT_015*25,;
      MAXLENGTH=3,;
      NAME='TXTF08'      
  ADD OBJECT TXTF06 AS TEXTBOX WITH;
      TOP=INT_015*54,;
      LEFT=INT_015*56,;
      WIDTH=INT_015*224,;
      HEIGHT=INT_015*25,;
      MAXLENGTH=30,;
      NAME='TXTF06'            
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*55,;
      LEFT=INT_015*561               
  add object GRID1 as GRID with;          
      AUTOCENTER=.T.,;
      LEFT=INT_015*10,;
      TOP=INT_015*96,;
      WIDTH=INT_015*649,;      
      HEIGHT=INT_015*205,;      
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
  add object cmnd2 as commandbutton with;
      LEFT=INT_015*253,;
      TOP=INT_015*339,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      caption='\<X.離開',;
      TOOLTIPTEXT='取消此作業畫面!快速鍵->ALT+X'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX') 
         THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')          
         AREA1='F06' 
        SELE F06
        SET RELA TO F07 INTO B01        
        SET FILTER TO F01+F02+F03+F04+F05=F02.F01+F02.F02+F02.F03+F02.F04+F02.F05
         SEEK F02.F01+F02.F02+F02.F03+F02.F04+F02.F05
         THISFORM.GRID1.READONLY=.T. 
         thisform.grid1.recordsource='F06'
         THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.ESC_BOTT.VISIBLE=.F.
         THISFORM.SHRGROUP.VISIBLE=.F.
         THISFORM.SHRGROUP.ENABLED=.F.                  
         thisform.setall('readonly',.t.,'textbox')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')         
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='代用料號'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*261
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='F06.F07'         
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'                           
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*170	        
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='版本'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='F06.F08'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*30
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='使用廠牌'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='F06.F06'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*224        
         THISFORM.GRID1.SETFOCUS
         FCH='GRID1'
         IF THISFORM.GRID1.ACTIVEROW=0
            THISFORM.CMDGROUP.ENABLED=.F.
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
            THISFORM.LBLF071.CAPTION=''
         ELSE
            THISFORM.CMDGROUP.ENABLED=.T.         
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.T.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.
         ENDIF                     
         THISFORM.CMND2.ENABLED=.T.                     
         THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.T.
      ENDPROC
      PROC GRID1.AFTERROWCOLCHANGE
         LPARAMETERS XColIndex     
         THISFORM.TXTF07.VALUE=F06.F07
         THISFORM.TXTF08.VALUE=F06.F08
         THISFORM.TXTF06.VALUE=F06.F06
         THISFORM.LBLF071.CAPTION=B01.F02
         FCH=THISFORM.ACTIVECONTROL.NAME 
         IF THIS.ACTIVEROW=0
            THISFORM.CMDGROUP.ENABLED=.F.
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
         ELSE
            THISFORM.CMDGROUP.ENABLED=.T.         
            THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.T.
            THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.         
         ENDIF                     
         THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.T.                     
      ENDPROC 
      procedure cmnd2.click
         SELE F06
         SET RELA OFF INTO B01
         SET FILTER TO
         thisform.release         
      endproc    
      PROC ORPGROUP.NEW_BOTT.CLICK
           THISFORM.CMND2.ENABLED=.F.
           SELE F06
           THISFORM.TXTF07.READONLY=.F.
           THISFORM.TXTF08.READONLY=.F.
           THISFORM.TXTF06.READONLY=.F.           
           THISFORM.TXTF07.VALUE=''
           THISFORM.TXTF08.VALUE=''
           THISFORM.TXTF06.VALUE=''           
           THISFORM.LBLF071.CAPTION=''           
           THISFORM.TXTF07.SETFOCUS

      ENDPROC
      PROC ORPGROUP.EDIT_BOTT.CLICK
           THISFORM.CMND2.ENABLED=.F.      
           SELE F06
           RLOCK()
           IF !RLOCK()
              DO file_impact
              THISFORM.GRID1.SETFOCUS
              THISFORM.SHRGROUP.ABANDON_BOTT.CLICK 
           ELSE
              THISFORM.TXTF07.READONLY=.F.
              THISFORM.TXTF08.READONLY=IIF(B01.F97='Y',.F.,.T.)
              THISFORM.TXTF07.SETFOCUS
           ENDIF                 
      ENDPROC        
      PROC ORPGROUP.DEL_BOTT.CLICK
           IF messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6
              SELE F02 
              REPL F11 WITH F11+(-1)	            
              SELE F06  
              RLOCK()           
              IF RLOCK()                
                 DELE
                 UNLOCK             
                 IF BOF()
                    GO TOP
                 ELSE
                   SKIP -1   
                 ENDIF   
              ELSE
                 DO file_impact                             
              ENDIF         
              THISFORM.GRID1.SETFOCUS
              IF THISFORM.GRID1.ACTIVEROW=0
                 THISFORM.CMDGROUP.ENABLED=.F.
                 THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
                 THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.   
                 THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.T.            
                 THISFORM.SETALL('VALUE','','TEXTBOX')
              ELSE      
                 THISFORM.CMDGROUP.ENABLED=.T.
                 THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.T.
                 THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.T.
              ENDIF           
           ENDIF   
           thisform.CMND2.enabled=.T.
      ENDPROC
      PROC SHRGROUP.SHURE1_BOTT.CLICK
           IF SEEK(THISFORM.TXTF07.VALUE,'B01')=.F.
              =MESSAGEBOX('無此基本料號',0+48,'提示訊息視窗')
              THISFORM.TXTF07.SETFOCUS
              RETURN
           ENDIF
           IF THISFORM.TXTF07.VALUE=F02.F04
              =MESSAGEBOX('代用料號不得等於原料號',0+48,'提示訊息視窗')
              THISFORM.TXTF07.SETFOCUS           
              RETURN
           ENDIF
           IF SEEK(F02.F01+F02.F02+F02.F03+F02.F04+F02.F05+THISFORM.TXTF07.VALUE+THISFORM.TXTF08.VALUE,'F06')=.T.
              =MESSAGEBOX('此版本料號重複建立!',0+64,'提示訊息視窗')
              THISFORM.TXTF07.SETFOCUS
              RETURN              
           ENDIF
           IF EMPTY(THISFORM.TXTF08.VALUE) AND IIF(SEEK(THISFORM.TXTF07.VALUE,'B01'),B01.F97,'')='Y'
              =MESSAGEBOX('版本不得空白',0+48,'提示訊息視窗')
              THISFORM.TXTF08.SETFOCUS
              RETURN
           ENDIF   
           SELE F02
           REPL F11 WITH F11+1           
           SELE F06
           APPEND BLANK
           RLOCK()
           REPL F01 WITH F02.F01
           REPL F02 WITH F02.F02
           REPL F03 WITH F02.F03
           REPL F04 WITH F02.F04
           REPL F05 WITH F02.F05
           REPL F07 WITH THISFORM.TXTF07.VALUE
           REPL F08 WITH THISFORM.TXTF08.VALUE     
           REPL F06 WITH THISFORM.TXTF06.VALUE                      
           UNLOCK
           SEEK F02.F01+F02.F02+F02.F03+F02.F04+F02.F05+THISFORM.TXTF07.VALUE+THISFORM.TXTF08.VALUE           
           THISFORM.GRID1.ENABLED=.T.           
           THISFORM.GRID1.SETFOCUS                      
          THISFORM.ORPGROUP.NEW_BOTT.CLICK    
      ENDPROC         
      PROC SHRGROUP.SHURE2_BOTT.CLICK
           IF EMPTY(THISFORM.TXTF08.VALUE) AND IIF(SEEK(THISFORM.TXTF07.VALUE,'B01'),B01.F97,'')='Y'
              =MESSAGEBOX('版本不得空白',0+48,'提示訊息視窗')
              THISFORM.TXTF08.SETFOCUS
              RETURN
           ENDIF         
           SELE F06
           REPL F07 WITH THISFORM.TXTF07.VALUE
           REPL F08 WITH THISFORM.TXTF08.VALUE
           REPL F06 WITH THISFORM.TXTF06.VALUE
*           SEEK F02.F01+F02.F02+F02.F03+F02.F04+F02.F05+THISFORM.TXTF07.VALUE+THISFORM.TXTF08.VALUE
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
           THISFORM.TXTF07.READONLY=.T.
           THISFORM.CMND2.ENABLED=.T.
           THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.T.                                                                     
           THISFORM.GRID1.ENABLED=.T.
           THISFORM.GRID1.SETFOCUS
           THISFORM.REFRESH
           UNLOCK ALL      
      ENDPROC
      PROC SHRGROUP.FINISH_BOTT.CLICK
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
      ENDPROC
 
enddefine            
*********************************************************************建立製程
DEFINE CLASS ROUTE_SEEK AS FORM
  autocenter=.t.
  FONTSIZE=INT_015*9
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='route_SEEK'
       TOP=INT_015*17
       LEFT=INT_015*34
       HEIGHT=INT_015*282
       WIDTH=INT_015*430
       CAPTION='製程編輯'+'    '+F01.F01+'  '+F01.F02
       DRAGIT=.F.                                &&---紀錄可拖曳的項目       
       NODRAGIT=.F.                              &&---紀錄不可拖曳的項目
       ADD OBJECT LABEL1 AS LABEL WITH;
           LEFT=INT_015*20,;
           HEIGHT=INT_015*25,;
           AUTOSIZE=.T.,;
           CAPTION='未加入之生產製程' 
       ADD OBJECT LABEL2 AS LABEL WITH;
           LEFT=INT_015*270,;
           HEIGHT=INT_015*25,;
           AUTOSIZE=.T.,;
           CAPTION='已排入之生產製程'        
       ADD OBJECT LIST1 AS LISTBOX WITH;                &&來源表列
           ROWSOURCETYPE=1,;
           VALUE=0,;
           DRAGICON='BMPS\DRAGMOVE.CUR',;
           HEIGHT=INT_015*188,;
           LEFT=INT_015*11,;
           TOP=INT_015*28,;
           WIDTH=INT_015*156,;
           ROWHEIGHT=INT_015*18,;  
           FONTSIZE=INT_015*9,;      
           TOOLTIPTEXT='未排入製程項目',;
           NAME='LIST1'
        ADD OBJECT LIST2 AS LISTBOX WITH;               &&目的表列
            ROWSOURCETYPE=1,;
            VALUE=0,;
            DRAGICON='DRAGMOVE.CUR',;
            HEIGHT=INT_015*185,;
            LEFT=INT_015*263,;
            TOP=INT_015*26,;
            WIDTH=INT_015*162,;
            ROWHEIGHT=INT_015*18,;  
            FONTSIZE=INT_015*9,;        
            TOOLTIPTEXT='已排入製程項目',;             
            NAME='LIST2'
         ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
             TOP=INT_015*39,;
             LEFT=INT_015*195,;
             HEIGHT=INT_015*34,;
             WIDTH=INT_015*38,;
             CAPTION='',;
             TOOLTIPTEXT='單筆資料匯入',;
             PICTURE='BMPS\FORWARDONE.BMP',;
             NAME='CMND1'
         ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
             TOP=INT_015*74,;
             LEFT=INT_015*195,;
             HEIGHT=INT_015*34,;
             WIDTH=INT_015*38,;
             CAPTION='',;
             TOOLTIPTEXT='全部資料匯入',;             
             PICTURE='BMPS\FORWARDALL.BMP',;             
             NAME='CMND2'
         ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
             TOP=INT_015*135,;
             LEFT=INT_015*195,;
             HEIGHT=INT_015*34,;
             WIDTH=INT_015*38,;
             CAPTION='',;
             TOOLTIPTEXT='單筆資料匯出',;             
             PICTURE='BMPS\BACKONE.BMP',;             
             NAME='CMND3'
         ADD OBJECT CMND4 AS COMMANDBUTTON WITH;
             TOP=INT_015*170,;
             LEFT=INT_015*195,;
             HEIGHT=INT_015*34,;
             WIDTH=INT_015*38,;
             CAPTION='',;
             TOOLTIPTEXT='全部資料匯出',;                          
             PICTURE='BMPS\BACKALL.BMP',;                          
             NAME='CMND4'             
         ADD OBJECT CMND5 AS COMMANDBUTTON WITH;
             TOP=INT_015*238,;
             LEFT=INT_015*284,;
             HEIGHT=INT_015*30,;
             WIDTH=INT_015*99,;
             CAPTION='\<X.離開',;
             NAME='CMND5'    
         ADD OBJECT CMND6 AS COMMANDBUTTON WITH;
             TOP=INT_015*238,;
             LEFT=INT_015*138,;
             HEIGHT=INT_015*30,;
             WIDTH=INT_015*99,;
             CAPTION='\<S.確認',;
             NAME='CMND6'                 
          PROCEDURE INIT
             THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')
             THISFORM.SETALL('FONTSIZE',INT_015*9,'LISTBOX')
             THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')
          ENDPROC   
          PROCEDURE LIST1.INIT        
                    with this
                         SELE E00_1
                         GO TOP
                         DO WHILE !EOF()                           
                            .ADDITEM(E00_1.F01+'  '+E00_1.F04)              
                             SKIP
                         ENDDO   
                         SELE E00_1
                         USE
                         .VALUE=1
                     endwith   
          ENDPROC                
          PROCEDURE LIST2.INIT
                    WITH THIS
                         SELE E05_1
                         GO TOP
                         DO WHILE !EOF()
                           .ADDITEM(E05_1.F07+'  '+E05_1.F04)
                           SKIP
                         ENDDO    
                         SELE E05_1
                         USE
                         .VALUE=1
                    ENDWITH
          ENDPROC
          PROCEDURE NODRAG                             &&當拖曳物件進入按鈕物件時出現禁止圖示
             PARAMETER nState
             DO CASE
                CASE nState=0                &&Enter
                     THISFORM.LIST1.DRAGICON=THIS.NODRAGIT
                     THISFORM.LIST2.DRAGICON=THIS.NODRAGIT
                CASE nState=1
                     THISFORM.LIST1.DRAGICON=THIS.DRAGIT
                     THISFORM.LIST2.DRAGICON=THIS.DRAGIT
              ENDCASE
           ENDPROC
           PROCEDURE DRAGOVER                          &&拖曳的物件進入表格
              LPARAMETERS oSource,nXCoord,nYCoord,nState
              THIS.NODRAG(nState)
           ENDPROC

           PROCEDURE Load                              &&指定拖曳時的游標圖案
              THIS.DRAGIT='BMPS\DRAGMOVE.CUR'
              THIS.NODRAGIT='BMPS\NO.CUR'
           ENDPROC
           PROCEDURE LIST1.DblClick                    &&在來源表列按兩次滑鼠左鍵
              THISFORM.LIST2.ADDITEM(THIS.LIST(THIS.Listindex))
              THIS.RemoveItem(THIS.ListIndex)
           ENDPROC
           PROCEDURE LIST1.MouseMove                   &&在來源表列中拖曳
              LPARAMETER nButton,nShift,nXCoord,nYCoord
              ITEM=THIS.VALUE   
              IF THIS.ListCount>0 AND nButton=1              &&尚有選項&MouseLeft
                 IF THIS.Selected(ITEM)                      &&已選取
                    THIS.DRAGICON=THISFORM.DRAGIT
                    THIS.DRAG(1)
                 ENDIF
              ENDIF      
           ENDPROC   
           PROCEDURE LIST1.DRAGOVER                    &&當有拖曳中的物件經過LIST1的上方
              LPARAMETER oSource,nXCoord,nYCoord,nState
              IF oSource.NAME='LIST2'
                 DO CASE
                    CASE nState=0                      &&Enter
                         THISFORM.LIST2.DRAGICON=THISFORM.DRAGIT
                    CASE nState=1                      &&Leave
                         THISFORM.LIST2.DRAGICON=THISFORM.NODRAGIT
                  ENDCASE
               ENDIF                                          
            ENDPROC
            PROCEDURE LIST1.DragDrop                   &&當有拖曳的物件放在LIST1物件中
               LPARAMETER oSource,nXCoord,nYCoord,nState                              
               IF UPPER(oSource.NAME)='LIST2'
                  THISFORM.LIST2.DblClick
               ENDIF
            ENDPROC
            PROCEDURE LIST2.DblClick                   &&在目的表列連按兩次滑鼠左鍵
               THISFORM.LIST1.Additem(THIS.List(This.ListIndex))
               THIS.RemoveItem(THIS.ListIndex)
            ENDPROC
            PROCEDURE LIST2.DragDrop                   &&當有拖曳的物件放在LIST2物件中
               LPARAMETERS oSource,nXCoord,nYCoord
               IF UPPER(oSource.NAME)='LIST1'
                  THISFORM.LIST1.DblClick
               ENDIF
            ENDPROC
            PROCEDURE LIST2.DragOver                   &&當有拖曳中的物件經過LIST2的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               IF oSource.NAME='LIST1'
                  DO CASE 
                     CASE nState=0             &&Enter
                          THISFORM.LIST1.DragIcon=THISFORM.DRAGIT
                     CASE nState=1             &&Leave
                          THISFORM.LIST2.Dragicon=THISFORM.NODRAGIT
                  ENDCASE
               ENDIF                
            ENDPROC
            PROCEDURE LIST2.MouseMove                  &&在目的表列中拖曳
               LPARAMETERS nButton,nShift,nXCoord,nYCoord
               ITEM=THIS.VALUE
               IF THIS.ListCount>0 AND nButton=1      &&尚有選項&MouseLeft
                  IF THIS.Selected(ITEM)              &&已選取   
                     THIS.DragIcon=THISFORM.DRAGIT
                     THIS.Drag(1)
                  ENDIF
               ENDIF
            ENDPROC
            PROCEDURE CMND1.CLICK                     &&>單筆匯入按鈕
               ITEM=THISFORM.LIST1.VALUE
               IF THISFORM.LIST1.SELECTED(ITEM)
                  THISFORM.LIST2.AddItem(THISFORM.LIST1.List(ITEM))
                  THISFORM.LIST1.RemoveItem(THISFORM.LIST1.ListIndex)
               ENDIF            
            ENDPROC
            PROCEDURE CMND1.Dragover                  &&當有拖曳中的物件經過CMND1上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC
            PROCEDURE CMND2.CLICK                     &&>>全部筆數匯入按鈕
               FOR I=1 TO THISFORM.LIST1.ListCount
                   THISFORM.LIST2.AddItem(THISFORM.LIST1.List(1))
                   THISFORM.LIST1.RemoveItem(1)
               ENDFOR
            ENDPROC
            PROCEDURE CMND2.Dragover                  &&當有拖曳中的物件經過CMND2上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC
            PROCEDURE CMND3.CLICK                     &&<單筆匯出按扭
               ITEM=THISFORM.LIST2.VALUE
               IF THISFORM.LIST2.Selected(ITEM)
                  THISFORM.LIST1.AddItem(THISFORM.LIST2.List(ITEM))
                  THISFORM.LIST2.RemoveItem(THISFORM.LIST2.ListIndex)
               ENDIF
            ENDPROC
            PROCEDURE CMND3.Dragover                  &&當有拖曳中的物件經過CMND3上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC           
            PROCEDURE CMND4.CLICK                     &&<<全部筆數匯出按鈕
               FOR I=1 TO THISFORM.LIST2.ListCount
                   THISFORM.LIST1.AddItem(THISFORM.LIST2.List(1))
                   THISFORM.LIST2.RemoveItem(1)
               ENDFOR
            ENDPROC
            PROCEDURE CMND4.DragOver                  &&有拖曳中的物件經過CMND4的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC
            PROCEDURE CMND5.CLICK                     &&離開按鈕
               RELEASE THISFORM
               CLEAR EVENT
            ENDPROC
            PROCEDURE CMND5.Dragover                  &&當有拖曳中的物件經過CMND5的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC        
            PROCEDURE CMND6.CLICK                    &&確認按鈕
               DELE FROM E05 WHERE E05.F01+E05.F02=F01.F01+F01.F02
               SELE E05
               FOR I=1 TO THISFORM.LIST2.ListCount
                   APPEND BLANK
                   REPLACE  F01 WITH F01.F01
                   REPLACE  F02 WITH F01.F02
                   REPLACE  F04 WITH DTCX(DATE())+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
                   REPLACE  F06 WITH PADL(ALLTRIM(STR(I)),3,'0')
                   REPLACE  F07 WITH THISFORM.LIST2.List(I)
                   REPLACE  F08 WITH F01.F07
                   REPLACE  F05 WITH IIF(SEEK(F07,'E00'),E00.F05,'1')
                   REPLACE  F09 WITH IIF(SEEK(F07,'E00'),E00.F02,'1')
                   REPLACE  F10 WITH IIF(SEEK(F07,'E00'),E00.F06,'1')                   
                   REPL F99 WITH 0
               ENDFOR                              
               IF F01.F05=.T.
                  SELE F03 
                  DELE FROM F03 WHERE F03.F01+F03.F02=F01.F01+F01.F02
                  SELE F07
                  DELE FROM F07 WHERE F07.F01+F07.F01=F01.F01+F01.F02     
                  SELE F01
                  REPL F05 WITH .F. 
*!*	                  THISFORM.PASS_CND.BACKCOLOR=RGB(255,0,0)
*!*	                  THISFORM.PASS_CND.CAPTION='    未確認    '
               ENDIF                                       
               RELEASE THISFORM
               CLEAR EVENT
            ENDPROC
            PROCEDURE CMND6.Dragover                &&當有拖曳中的物件經過CMND6的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC              
            PROCEDURE LABEL1.Dragover                &&當有拖曳中的物件經過LABEL1的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC                          
            PROCEDURE LABEL2.Dragover                &&當有拖曳中的物件經過LABEL2的上方
               LPARAMETERS oSource,nXCoord,nYCoord,nState
               THISFORM.NoDrag(nState)
            ENDPROC                                      
ENDDEFINE
***************************************************供料方式
DEFINE CLASS SPLY_TYPE AS COMBOBOX
  LEFT=INT_015*65
  TOP=INT_015*128
  HEIGHT=INT_015*25
  WIDTH=INT_015*80
  FONTSIZE=INT_015*9
  BOUNDCOLUMN=1
  COLUMNCOUNT=1
  ROWSOURCETYPE=1
  VALUE=1
  STYLE=2
  NAME='SPLY_TYPE'
  PROCEDURE INIT
     WITH THIS 
          .ADDITEM('廠內發料')
          .ADDITEM('廠商自備')          
          .ADDITEM('客戶提供')          
     ENDWITH   
  ENDPROC
  PROCEDURE INTERACTIVECHANGE
     THISFORM.PAGEFRAME1.PAGE2.TXTF06.VALUE=ALLTRIM(STR(THIS.VALUE))
  ENDPROC
ENDDEFINE  
**********************************************&&複製後之料號及版本
define class VRS_COPY as form
  autocenter=.t.
  caption=FTR_STR+'BOM複製'
  fontsize=INT_015*9
  HEIGHT=INT_015*190
  WIDTH=INT_015*400
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='VRS_COPY'
  add object label1 as label with;
      autosize=.t.,;
      HEIGHT=INT_015*25,;
      left=INT_015*30,;
      top=INT_015*45,;
      FONTSIZE=INT_015*9,;
      name='label1'  
  add object label2 as label with;
      autosize=.t.,;
      HEIGHT=INT_015*25,;
      left=INT_015*30,;
      top=INT_015*75,;
      FONTSIZE=INT_015*9,;
      name='label2'
      add object text1 as textbox with;
          autosize=.t.,;
          top=INT_015*40,;
          WIDTH=INT_015*261,;
          left=INT_015*110,;
          HEIGHT=INT_015*25,; 
          FONTSIZE=INT_015*9,;
          MAXLENGTH=43,;
          name='text1'
      add object text2 as textbox with;
          autosize=.t.,;
          top=INT_015*70,;
          WIDTH=INT_015*40,;
          left=INT_015*110,;
          HEIGHT=INT_015*25,; 
          FONTSIZE=INT_015*9,;
          MAXLENGTH=3,;
          name='text2'          
      add object cmnd1 as commandbutton with;
          autosize=.t.,;
          left=INT_015*20,;    
          HEIGHT=INT_015*25,;   
          top=INT_015*144,;  
          WIDTH=INT_015*80,;
          FONTSIZE=INT_015*9,;          
          caption='\<Y.執行',;
          TOOLTIPTEXT='確認所輸入的版本!快速鍵->ALT+Y',;
          name='cmnd1'
      add object cmnd2 as commandbutton with;
          autosize=.t.,;
          left=INT_015*120,;
          top=INT_015*144,;
          HEIGHT=INT_015*25,;
          WIDTH=INT_015*80,;
          FONTSIZE=INT_015*9,;
          caption='\<C.取消',;
          TOOLTIPTEXT='取消複製作業!快速鍵->ALT+C',;
          name='cmnd2'
      procedure init 
          THISFORM.label1.caption='請輸入料號'                     
          THISFORM.label2.caption='請輸入版本'               
          THISFORM.TEXT1.VALUE=LEFT(FTR_STR,43)
          THISFORM.TEXT2.VALUE=''
      ENDPROC    
      PROCEDURE TEXT1.INTERACTIVECHANGE
             IF AT('?',ALLTRIM(THIS.VALUE))<>0
                  IF AT('?',ALLTRIM(THIS.VALUE))>1             
                      SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
                  ELSE
                      SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01    
                  ENDIF        
                  IF _TALLY>0
                      MAN_SEEK=createobject("MTR_SEEK")  
                      MAN_SEEK.SHOW    
                      THIS.VALUE=FLG
                      THIS.REFRESH
                      FLG=''     
                   ENDIF    
              ENDIF              
      ENDPROC
      procedure cmnd1.click
             IF EMPTY(THISFORM.TEXT1.VALUE)
                 =MESSAGEBOX('料號不得空白',0+48,'提示訊息視窗')
                 THISFORM.TEXT1.SETFOCUS
                 RETURN
              ENDIF   
             IF EMPTY(THISFORM.TEXT2.VALUE)
                 =MESSAGEBOX('版本不得空白',0+48,'提示訊息視窗')
                 THISFORM.TEXT2.SETFOCUS
                 RETURN
              ENDIF                       
             IF SEEK(THISFORM.TEXT1.VALUE,'B01_1')=.F.
                =MESSAGEBOX('此料號不得建立BOM',0+48,'提示訊息視窗')
                THISFORM.TEXT1.SETFOCUS
                RETURN
             ENDIF                     
             SELE F01
             IF SEEK(THISFORM.TEXT1.VALUE+THISFORM.TEXT2.VALUE,'F01')=.T.
                =MESSAGEBOX('此料號之版本已建立',0+48,'提示訊息視窗')
                THISFORM.TEXT1.SETFOCUS
                RETURN
             ENDIF               
             IF EMPTY(THISFORM.TEXT1.VALUE)
                 =MESSAGEBOX('料號不得空白',0+48,'提示訊息視窗')
                 THISFORM.TEXT1.SETFOCUS
                 RETURN
              ENDIF   
             IF EMPTY(THISFORM.TEXT2.VALUE)
                 =MESSAGEBOX('版本不得空白',0+48,'提示訊息視窗')
                 THISFORM.TEXT2.SETFOCUS
                 RETURN
              ENDIF                 
             FLG=THISFORM.TEXT1.VALUE+THISFORM.TEXT2.VALUE                
          THISFORM.RELEASE                
      endproc
      procedure cmnd2.click
          FLG=''
          thisform.release
      endproc      
enddefine            
****************************************************
DEFINE CLASS TEXT1 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=43
  FONTSIZE=INT_015*9
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)    
       ELSE
          SELECT F01,F02 FROM B01_1 INTO CURSOR MTR ORDER BY F01    
       ENDIF        
       IF _TALLY>0
          MAN_SEEK=createobject("MTR_SEEK")  
          MAN_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'     
       ENDIF    
    ENDIF  
  ENDPROC
ENDDEFINE