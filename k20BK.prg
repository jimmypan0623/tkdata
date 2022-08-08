close all
clear 
FLG='0'
FCH=''
CHK=''
CHK_STR=''
FTR_STR=''
UKEF=''
R=0
QTY=0
TQTY=0
AREA1='K37'
AREA2='K39'     
RTE_TABLE='K38'
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'K20'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('A23')
   SELE 0
   USE A23
ELSE 
  SELE A23
ENDIF
SET FILTER TO MOD(VAL(SUBSTR(DTOS(F01),5,2)),2)<>0
SET ORDER TO 1
GO BOTT
K20form=createobject("tkK20")
K20form.show       
define class tkK20 as form
  autocenter=.t.
  caption='K20.產生電子計算機發票'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*220
  width=INT_015*320
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkK20'
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*160,;
      TOP=INT_015*35,;
      WIDTH=INT_015*80,;      
      HEIGHT=INT_015*30,;    
      FONTSIZE=INT_015*12  
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入發票月份'    

  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  PROC MTH_LIST.INIT    
       WITH THIS
            SELE A23
            GO TOP
            K=0
            S=0
            DO WHILE !EOF()                
               K=K+1
               IF INT_012=1
                   SHDT=PADL(ALLTRIM(STR(VAL(LEFT(DTOS(A23.F01),4))-1911)),4,'0')+'/'+SUBSTR(DTOS(A23.F01),5,2)
                   
               ELSE
                   SHDT=LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)
               ENDIF    
               .ADDITEM(SHDT) &&               
               IF LEFT(DTOS(DATE()),6)=LEFT(DTOS(A23.F01),6)
                   S=K
               ELSE
                   
                   S=K-1
               ENDIF       
               SKIP
            ENDDO   

            .VALUE=S
       ENDWITH   
  ENDPROC       
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*150,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所查詢條件',;
      NAME='CMND1'
  add object cmnd2 as commandbutton with;
      left=INT_015*190,;
      top=INT_015*150,;
      height=INT_015*40,;
      width=INT_015*80,;
      FONTSIZE=INT_015*12,;
      caption='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.MTH_LIST.SETFOCUS
      ENDPROC
**********************************開始搜尋指定範圍之發票       
      PROCEDURE CMND1.CLICK      
      
          K37='K37'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          K38='K38'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          K39='K39'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          K391='K391'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
        IF FILE('&K37'+'.DBF')=.F.
             =MESSAGEBOX('尚未產生月份檔!',0+48,'提示訊息視窗')
             THISFORM.CMND2.CLICK
        ELSE 
          IF !USED('&K37')
             SELE 0
             USE (K37) ALIA K37
          ELSE
             SELE K37
          ENDIF
          SET ORDER TO 1
          IF !USED('&K38')
             SELE 0
             USE (K38) ALIA K38
          ELSE
             SELE K38
          ENDIF
          SET ORDER TO 1
          IF !USED('&K39')
             SELE 0
             USE (K39) ALIA K39
          ELSE
             SELE K39
          ENDIF
          SET ORDER TO K391
          IF RECCOUNT()=0 AND IIF(A02.F05='*',.T.,.F.)
             IF MESSAGEBOX('尚未產生發票號碼是否現在產生?',4+32+256,'請確認') = 6
                MTH_NO=RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                K2Jform=createobject("tkK2J")
                K2Jform.show              
                HKS=THISFORM.MTH_LIST.DISPLAYVALUE
                K2Kform=createobject("tkK2K")
                K2Kform.show                                              
             ELSE
                THISFORM.CMND2.CLICK
             ENDIF   
          ELSE
             HKS=THISFORM.MTH_LIST.DISPLAYVALUE
             K2Kform=createobject("tkK2K")
             K2Kform.show                                           
          ENDIF    
             
        ENDIF        
          THISFORM.CMND2.CLICK
      ENDPROC
      procedure cmnd2.click        
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          

********
    define class tkK2J as form
               caption='輸入欲產生之發票號碼'
               fontsize=INT_015*9
               height=INT_015*270
               width=INT_015*320
               CONTROLBOX=.F.
               BORDERSTYLE=2
               AUTOCENTER=.T.
               SHOWTIPS=.T.
               showwindow=1
               windowtype=1
               name='tkK2J'
               ADD OBJECT LBL1 AS LABEL WITH;
                        LEFT=INT_015*5,;
                        TOP=INT_015*40,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION='請輸入字軌'
               ADD OBJECT LBL2 AS LABEL WITH;
                        LEFT=INT_015*5,;
                        TOP=INT_015*90,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION='請輸入起始號碼'                        
               ADD OBJECT LBL3 AS LABEL WITH;
                        LEFT=INT_015*5,;
                        TOP=INT_015*140,;
                        FONTSIZE=INT_015*12,;
                        AUTOSIZE=.T.,;
                        CAPTION='請輸入截止號碼'                                            
              ADD OBJECT TXT1 AS TEXTBOX WITH;
                       LEFT=INT_015*130,;
                       TOP=INT_015*35,;
                       HEIGHT=INT_015*30,;
                       WIDTH=INT_015*40,;
                       FONTSIZE=INT_015*12,;
                       MAXLENGTH=2,;
                       NAME='TXT1'    
              ADD OBJECT TXT2 AS TEXTBOX WITH;
                       LEFT=INT_015*130,;
                       TOP=INT_015*85,;
                       HEIGHT=INT_015*30,;
                       WIDTH=INT_015*100,;
                       FONTSIZE=INT_015*12,;
                       MAXLENGTH=8,;
                       NAME='TXT2'    
             ADD OBJECT TXT3 AS TEXTBOX WITH;
                       LEFT=INT_015*130,;
                       TOP=INT_015*135,;
                       HEIGHT=INT_015*30,;
                       WIDTH=INT_015*100,;
                       FONTSIZE=INT_015*12,;
                       MAXLENGTH=8,;
                       NAME='TXT3'    
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
             PROCEDURE CMND1.CLICK
                   IF ASC(LEFT(THISFORM.TXT1.VALUE,1))<65 OR ASC(LEFT(THISFORM.TXT1.VALUE,1))>90
                      =MESSAGEBOX('字軌錯誤!',0+48,'提示訊息視窗')
                      THISFORM.TXT1.SETFOCUS
                      RETURN
                   ENDIF
                   IF ASC(RIGHT(THISFORM.TXT1.VALUE,1))<65 OR ASC(RIGHT(THISFORM.TXT1.VALUE,1))>90
                      =MESSAGEBOX('字軌錯誤!',0+48,'提示訊息視窗')
                      THISFORM.TXT1.SETFOCUS
                      RETURN
                   ENDIF
                   IF (VAL(THISFORM.TXT3.VALUE)-VAL(THISFORM.TXT2.VALUE)+1)%50<>0
                      =MESSAGEBOX('起訖號碼有誤!',0+48,'提示訊息視窗')
                      THISFORM.TXT2.SETFOCUS
                      RETURN
                   ENDIF
                   IF THISFORM.TXT3.VALUE<THISFORM.TXT2.VALUE
                       =MESSAGEBOX('截止號碼小於起始號碼',0+48,'提示訊息視窗')
                       THISFORM.TXT2.SETFOCUS
                       RETURN
                     ENDIF             
                     IF !(RIGHT(THISFORM.TXT2.VALUE,2)='00' OR RIGHT(THISFORM.TXT2.VALUE,2)='50')
                       =MESSAGEBOX('起始號碼有誤',0+48,'提示訊息視窗')
                       THISFORM.TXT3.SETFOCUS
                       RETURN
                     ENDIF                     
                     IF !(RIGHT(THISFORM.TXT3.VALUE,2)='49' OR RIGHT(THISFORM.TXT3.VALUE,2)='99')
                       =MESSAGEBOX('截止號碼有誤',0+48,'提示訊息視窗')
                       THISFORM.TXT3.SETFOCUS
                       RETURN
                     ENDIF
                     
                     SCNT=1
                     STD_NO=THISFORM.TXT2.VALUE
                     END_NO=THISFORM.TXT3.VALUE
                     DO WHILE STD_NO<=THISFORM.TXT3.VALUE
                        SELECT K37
                        APPEND BLANK
                        REPLACE F01 WITH PADL(ALLTRIM(STR(SCNT)),4,'0')
                        REPLACE F02 WITH '31'
                        REPLACE F03 WITH THISFORM.TXT1.VALUE+STD_NO
                        REPLACE F06 WITH F03
                        FOR I=1 TO 50
                            SELECT K39
                            APPEND BLANK
                            REPLACE F01 WITH K37.F01
                            REPLACE F02 WITH THISFORM.TXT1.VALUE+STD_NO
                            STD_NO=PADL(ALLTRIM(STR(VAL(STD_NO)+1)),8,'0')
                        ENDFOR
                        SELECT K37
                        REPLACE F04 WITH THISFORM.TXT1.VALUE+PADL(ALLTRIM(STR(VAL(STD_NO)-1)),8,'0')
                        REPLACE F07 WITH 50
                        SCNT=SCNT+1
                     ENDDO
                        CT=INT((SCNT-1)/2)
                        SELECT K37 
                        REPLACE F08 WITH MTH_NO FOR VAL(F01)<=CT
                        REPLACE F08 WITH PADL(ALLTRIM(STR(VAL(MTH_NO)+1)),2,'0') FOR VAL(F01)>CT
                     THISFORM.CMND2.CLICK
             ENDPROC
             procedure cmnd2.click        

                   THISFORM.RELEASE                  
              endproc      
        ENDDEFINE
********************************************************************
define class tkK2K as form
  caption='K20.電子計算機發票號碼內容檔'
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
  NAME='TKK2K'
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
      RECORDSOURCE='K37',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3'
  ADD OBJECT grid2 AS grid2 WITH;
      columncount=2,;            
      RECORDSOURCE='K39',; 
      LINKMASTER='K37',;
      RELATIONALEXPR='F01',;
      childorder='K391',;  
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2'
     
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.                             
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.               
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*173,;
      LEFT=INT_015*558          
  ADD OBJECT EXL_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*7,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*64,;      
      CAPTION='\<L.轉申報檔',;
      NAME='EXL_BOTT'                        
  ADD OBJECT EPT_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*71,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*91,;      
      CAPTION='\<Q.列印空白發票',;
      NAME='EPT_BOTT'                              
  ADD OBJECT RTG_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*490,;
      LEFT=INT_015*164,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      WIDTH=INT_015*64,;      
      CAPTION='\<T.編輯人員',;
      NAME='RTG_BOTT'                       
          
  PROCEDURE PAGEFRAME1.PAGE1.INIT 
     THIS.ADDOBJECT('LBLF01','LABEL')
     WITH THIS.LBLF01
          .VISIBLE=.T.
          .LEFT=INT_015*14
          .TOP=INT_015*7
          .WIDTH=INT_015*50
          .CAPTION='冊        號'          
     ENDWITH
     THIS.ADDOBJECT('LBLF02','LABEL')
     WITH THIS.LBLF02
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*32
         .WIDTH=INT_015*50
         .CAPTION='類        別'
     ENDWITH
     THIS.ADDOBJECT('LBLF021','LABEL')
     WITH THIS.LBLF021
         .VISIBLE=.T.
         .LEFT=INT_015*105
         .TOP=INT_015*32
         .WIDTH=INT_015*50
         .CAPTION=''
     ENDWITH
     THIS.ADDOBJECT('LBLF08','LABEL')
     WITH THIS.LBLF08
         .VISIBLE=.T.
         .LEFT=INT_015*220
         .TOP=INT_015*32
         .WIDTH=INT_015*50
         .CAPTION='歸屬月份'
     ENDWITH
     THIS.ADDOBJECT('LBLF03','LABEL')
     WITH THIS.LBLF03
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*57
         .WIDTH=INT_015*50
         .CAPTION='起始號碼'
     ENDWITH     
     THIS.ADDOBJECT('LBLF04','LABEL')
     WITH THIS.LBLF04
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*82
         .WIDTH=INT_015*50
         .CAPTION='截止號碼'
     ENDWITH          
     THIS.ADDOBJECT('LBLF05','LABEL')
     WITH THIS.LBLF05
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*107
         .WIDTH=INT_015*50
         .CAPTION='最後異動'
     ENDWITH               
     THIS.ADDOBJECT('LBLF06','LABEL')
     WITH THIS.LBLF06
         .VISIBLE=.T.
         .LEFT=INT_015*214
         .TOP=INT_015*107
         .WIDTH=INT_015*50
         .CAPTION='待用號碼'
     ENDWITH                    
     THIS.ADDOBJECT('LBLE05','LABEL')
     WITH THIS.LBLE05
         .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*132
         .WIDTH=INT_015*50
         .CAPTION='開立人員'
     ENDWITH                              
     THIS.ADDOBJECT('LBLE07','LABEL')
     WITH THIS.LBLE07
         .VISIBLE=.T.
         .LEFT=INT_015*214
         .TOP=INT_015*132
         .WIDTH=INT_015*50
         .CAPTION='剩餘張數'
     ENDWITH                                   
      THIS.ADDOBJECT('RTGBOX','EDITBOX')&&文書編輯發票開立人員
      WITH THIS.RTGBOX
                  .LEFT=INT_015*64
                 .TOP=INT_015*128
                 .WIDTH=INT_015*100
                 .HEIGHT=INT_015*48
                 .FONTSIZE=INT_015*8
                 .VISIBLE=.T.
                 .READONLY=.T.
                 .MAXLENGTH=500
                 .NAME='RTGBOX'           
      ENDWITH         
     THIS.ADDOBJECT('TXTF01','TEXTBOX')          
     WITH THIS.TXTF01
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*3
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25
          .MAXLENGTH=4
     ENDWITH        

     THIS.ADDOBJECT('TXTF02','TEXTBOX')          
     WITH THIS.TXTF02
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*27
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2
     ENDWITH        
     THIS.ADDOBJECT('TXTF08','TEXTBOX')          
     WITH THIS.TXTF08
          .VISIBLE=.T.
          .LEFT=INT_015*270
          .TOP=INT_015*27
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2
     ENDWITH        
     THIS.ADDOBJECT('INV_TYPE','INV_TYPE')     
     THIS.ADDOBJECT('TXTF03','TEXTBOX')          
     WITH THIS.TXTF03
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*53
          .WIDTH=INT_015*80
          .HEIGHT=INT_015*25
          .MAXLENGTH=20
     ENDWITH                  
     THIS.ADDOBJECT('TXTF04','TEXTBOX')          
     WITH THIS.TXTF04
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*78
          .WIDTH=INT_015*80
          .HEIGHT=INT_015*25
          .MAXLENGTH=10
     ENDWITH     
     THIS.ADDOBJECT('TXTF05','TEXTBOX')          
     WITH THIS.TXTF05
          .VISIBLE=.T.
          .LEFT=INT_015*64
          .TOP=INT_015*103
          .WIDTH=INT_015*80
          .HEIGHT=INT_015*25
          
     ENDWITH     
     THIS.ADDOBJECT('TXTF06','TEXTBOX')          
     WITH THIS.TXTF06
          .VISIBLE=.T.
          .LEFT=INT_015*264
          .TOP=INT_015*103
          .WIDTH=INT_015*80
          .HEIGHT=INT_015*25
          .MAXLENGTH=10
     ENDWITH         
     THIS.ADDOBJECT('TXTF07','TEXTBOX')          
     WITH THIS.TXTF07
          .VISIBLE=.T.
          .LEFT=INT_015*264
          .TOP=INT_015*128
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .INPUTMASK='99'
     ENDWITH         
     
  ENDPROC   
  PROCEDURE PAGEFRAME1.PAGE2.INIT
	THIS.ADDOBJECT('LBLF02','LABEL')
	WITH THIS.LBLF02
	     .VISIBLE=.T.
         .LEFT=INT_015*14
         .TOP=INT_015*7
	     .WIDTH=INT_015*50
         .CAPTION='發票號碼'
    ENDWITH  
	THIS.ADDOBJECT('TXTF02','TEXTBOX')
	WITH THIS.TXTF02  
	     .VISIBLE=.T.
	     .LEFT=INT_015*65
	     .TOP=INT_015*3
	     .WIDTH=INT_015*100
	     .HEIGHT=INT_015*25
	     .MAXLENGTH=10
	     .NAME='TXTF02'  
	ENDWITH     
  ENDPROC       
  PROCEDURE PAGEFRAME1.PAGE1.activate
    R=0
    SELE K37
     IF THISFORM.SHRGROUP.VISIBLE=.F.
        THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR",'',"COLUMN")              
        THISFORM.GRID1.ENABLED=.T.   
        THISFORM.GRID1.SETFOCUS   
        THISFORM.KEY_LIST.ENABLED=.T.
     ELSE
        THISFORM.KEY_LIST.ENABLED=.F.   
     ENDIF                
    IF THISFORM.GRID1.ACTIVEROW=0 .AND. thisform.SHRGROUP.ENABLED=.F. 
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
        THISFORM.PAGEFRAME1.PAGE1.LBLF021.CAPTION=''
           
        THISFORM.PAGEFRAME1.PAGE1.SETALL('VALUE','','TEXTBOX')
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=''
     
        THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE=''
       
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.      
        THISFORM.RTG_BOTT.ENABLED=.F.
        THISFORM.EXL_BOTT.ENABLED=.F.
        THISFORM.EPT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
     ELSE
        THISFORM.PAGEFRAME1.PAGE1.LBLF021.CAPTION=IIF(K37.F02='32','二聯式','三聯式')
        THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND K37.F07=50 &&判斷有無修改的權限   
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.) &&判斷有無列印全部發票的權限 
        THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=IIF(A02.F07='*',.T.,.F.) &&判斷有無列印全部發票的權限     
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
        THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T. 
        THISFORM.RTG_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.) AND K37.F07>0
        THISFORM.EXL_BOTT.ENABLED=IIF(A02.F09='*',.T.,.F.) AND MONTH(DATE())>MONTH(CTOD(HKS+'/01'))
        THISFORM.EPT_BOTT.ENABLED=IIF(A02.F10='*',.T.,.F.) AND MONTH(DATE())>MONTH(CTOD(HKS+'/01'))
     ENDIF   

  ENDPROC
  PROCEDURE PAGEFRAME1.PAGE2.activate
     SELE K39
     THISFORM.GRID2.SETFOCUS
     IF THISFORM.GRID2.ACTIVEROW=0 
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.                          
        THISFORM.PAGEFRAME1.PAGE2.SETALL('VALUE','','TEXTBOX')               
     ELSE
        THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
        THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.
        THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.                      
     ENDIF  
        THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
        THISFORM.RTG_BOTT.ENABLED=.F. 
        THISFORM.EXL_BOTT.ENABLED=.F. 
        THISFORM.EPT_BOTT.ENABLED=.F. 
        THISFORM.SHRGROUP.ENABLED=.F.
        THISFORM.SHRGROUP.VISIBLE=.F.
     R=RECNO('&AREA1')
     THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
     THISFORM.GRID1.ENABLED=.F.     
     THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
     THISFORM.KEY_LIST.ENABLED=.F.     
     THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.
  ENDPROC     

  PROC INIT       
       THISFORM.Caption=THISFORM.Caption+HKS
       thisform.setall('height',INT_015*25,'textbox')
*       thisform.setall('height',INT_015*17,'label')
       thisform.setall('FONTSIZE',INT_015*9,'textbox')
       thisform.setall('FONTSIZE',INT_015*9,'label')       
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')              
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID2.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)               
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID1.setall('FONTSIZE',INT_015*9,'HEADER')         
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='冊號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='類別'            
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='聯式'  
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='K37.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='K37.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE="IIF(K37.F02='31','三聯式','二聯式')"
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*50
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*30
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*50
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.GRID2.setall('FONTSIZE',INT_015*9,'HEADER')       
       THISFORM.grid2.COLUMN1.HEADER1.CAPTION='發票號碼'       
       THISFORM.grid2.COLUMN2.HEADER1.CAPTION='開立日期'       
                 
	   THISFORM.GRID2.LINKMASTER='K37'
       THISFORM.GRID2.RELATIONALEXPR='F01+F02'
	   THISFORM.grid2.COLUMN1.CONTROLSOURCE='K39.F02'	
	   THISFORM.grid2.COLUMN1.WIDTH=INT_015*100      
	   THISFORM.grid2.COLUMN2.CONTROLSOURCE='K39.F03'	
	   THISFORM.grid2.COLUMN2.WIDTH=INT_015*75
       THISFORM.GRID2.SETALL("DYNAMICBACKCOLOR","IIF(K39.F04='+',RGB(255,0,0),'')","COLUMN")	 
       THISFORM.grid1.READONLY=.T.      
       THISFORM.grid2.READONLY=.T.            
       THISFORM.grid1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.         
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  AND K37.F07=50&&判斷有無修改的權限    
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.) &&判斷有無列印全部發票的權限     
          THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=IIF(A02.F07='*',.T.,.F.) &&判斷有無列印全部發票的權限       
       ENDIF          
        THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.        
        THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F. 
        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
        THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.

        THISFORM.SHRGROUP.ENABLED=.F.
        THISFORM.SHRGROUP.VISIBLE=.F.
        THISFORM.KEY_LIST.ENABLED=.T.
        THISFORM.KEY_LIST.VALUE=1   
        JK='冊號'
        SELE K37
        THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=K37.F01       
  ENDPROC               
  PROC KEY_LIST.INIT 
       with this 
           .additem('依冊號排列')                     
       endwith      
  ENDPROC 
  PROC grid1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex           
       THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE=K37.F01
       THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=K37.F02
       THISFORM.PAGEFRAME1.PAGE1.TXTF03.VALUE=K37.F03
       THISFORM.PAGEFRAME1.PAGE1.TXTF04.VALUE=K37.F04
       THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=K37.F05       
       THISFORM.PAGEFRAME1.PAGE1.TXTF06.VALUE=K37.F06     
       THISFORM.PAGEFRAME1.PAGE1.TXTF07.VALUE=K37.F07
       THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=K37.F08
       THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE=INV_MAN(K37.F01)   
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION  
       UKEF=THISFORM.PAGEFRAME1.PAGE1.RTGBOX.VALUE              
       IF THIS.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
          THISFORM.PAGEFRAME1.PAGE1.LBLF021.CAPTION=''
       ELSE
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND K37.F07=50 &&判斷有無修改的權限    
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
          THISFORM.PAGEFRAME1.PAGE1.LBLF021.CAPTION=IIF(K37.F02='32','二聯式','三聯式')
       ENDIF   
       CHK=''     
*       thisform.refresh
  ENDPROC     
  PROC grid2.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.PAGEFRAME1.PAGE2.TXTF02.VALUE=K39.F02
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
       if thisform.pageframe1.activepage=1
          thisform.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ENDIF   
*       thisform.refresh
  ENDPROC 
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
     THISFORM.KEY_LIST.ENABLED=.F. 
     THISFORM.RTG_BOTT.ENABLED=.F.       
     THISFORM.EXL_BOTT.ENABLED=.F.  
     THISFORM.EPT_BOTT.ENABLED=.F.                             
     THISFORM.GRID1.ENABLED=.F.   
     THISFORM.GRID2.ENABLED=.F.        
     THISFORM.PAGEFRAME1.PAGE1.INV_TYPE.ENABLED=.T.
     THISFORM.PAGEFRAME1.PAGE1.INV_TYPE.VISIBLE=.T.
     THISFORM.PAGEFRAME1.PAGE1.INV_TYPE.VALUE=IIF(K37.F02='32',2,1)
        SELECT K37
        LOCK()
        IF !RLOCK()
           DO file_impact
           THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
*           return
        ELSE
           THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')     &&處理表頭修改     
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.READONLY=.F.                 
           THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.           
        endif                            
  endproc      
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK                          &&修改確認
    IF !(VAL(THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE)-VAL(RIGHT(HKS,2))=0 OR VAL(THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE)-VAL(RIGHT(HKS,2))=1)
       =MESSAGEBOX('歸屬月份錯誤',0+48,'提示訊息')
       thisform.pageframe1.page1.txtf08.setfocus
       return
    ENDIF
    SELECT K38
    SEEK THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
    IF FOUND()
       DO WHILE F01=THISFORM.PAGEFRAME1.PAGE1.TXTF01.VALUE
          REPLACE F03 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE          
          SKIP
       ENDDO   
    ENDIF
    SELECT K37
    REPLACE F02 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE
    REPLACE F08 WITH PADL(ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE),2,'0')
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
      THISFORM.PAGEFRAME1.PAGE1.INV_TYPE.ENABLED=.F.
      THISFORM.PAGEFRAME1.PAGE1.INV_TYPE.VISIBLE=.F.             
      THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.      
      THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.      
      THISFORM.grid2.ENABLED=.T.                         
      THISFORM.grid1.ENABLED=.T.           
     
         THISFORM.RTG_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.) AND K37.F07>0                          
         THISFORM.EXL_BOTT.ENABLED=IIF(A02.F09='*',.T.,.F.) AND MONTH(DATE())>MONTH(CTOD(HKS+'/01'))   
         THISFORM.EPT_BOTT.ENABLED=IIF(A02.F10='*',.T.,.F.) AND MONTH(DATE())>MONTH(CTOD(HKS+'/01'))    
         THISFORM.GRID1.SETFOCUS
         THISFORM.KEY_LIST.ENABLED=.T.    

      UNLOCK ALL
      *THISFORM.ANS_BOTT.ENABLED=!F01.F05 .AND. IIF(A02.F11='*',.T.,.F.)  .AND. !THISFORM.SHRGROUP.ENABLED                   
   ENDPROC
     PROCEDURE RTG_BOTT.CLICK                   &&編輯人員
         SELE K38.F02,A01.F03 FROM K38,A01 WHERE A01.F01=K38.F02 AND K38.F01=K37.F01 ORDER BY K38.F02 INTO CURSOR K38_1  
         IF _TALLY>0
            SELECT F01,F03 FROM A01 WHERE F01 NOT IN (SELECT F02 FROM K38 WHERE K38.F01=K37.F01) AND F04='S000 ' ORDER BY F01 INTO CURSOR A01_1 
         ELSE
            SELECT F01,F03 FROM A01  ORDER BY F01 INTO CURSOR A01_1 WHERE F04='S000 '
         ENDIF   
         INVMAN_SEEK=createobject("INVMAN_SEEK")  
         INVMAN_SEEK.SHOW                
             FLG='0'                    
             AREA1='K37'
             SELE K37     
    ENDPROC  
    PROCEDURE EXL_BOTT.CLICK    &&轉申報檔
           IF RIGHT(HKS,2)='11'
              JST=PADL(ALLTRIM(STR(YEAR(CTOD(HKS+'/01'))+1-1911)),3,'0')+'01'
                     
          ELSE
             JST=SUBSTR(HKS,2,3)+PADL(ALLTRIM(STR(VAL(RIGHT(HKS,2))+2)),2,'0')
             
          ENDIF    
          TXMTH1=LEFT(DTOS(CTOD(HKS+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(HKS,2)))),2,'0')
          TXMTH2=LEFT(DTOS(CTOD(HKS+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(HKS,2))+1)),2,'0')
          B061='B06'+TXMTH1
          B062='B06'+TXMTH2
          C131='C13'+TXMTH1
          C132='C13'+TXMTH2
          K251='K25'+TXMTH1
          K252='K25'+TXMTH2
          **********************開檔案
          IF !USED('C01')
              SELECT 0
              USE C01
          ELSE
              SELECT C01
          ENDIF
          SET ORDER TO 1        
          IF !USED('&B06')
             SELECT 0
             USE &B061 ALIAS B061
          ELSE
             SELECT B061             
          ENDIF   
          SET ORDER TO 2          
          IF !USED('&B062')
             SELECT 0
             USE &B062 ALIAS B062
          ELSE
             SELECT B062             
          ENDIF   
          SET ORDER TO 2                       
          IF !USED('&C131')
             SELECT 0
             USE &C131 ALIAS C131
          ELSE
             SELECT C131             
          ENDIF   
          SET ORDER TO 3          
          IF !USED('&C132')
             SELECT 0
             USE &C132 ALIAS C132
          ELSE
             SELECT C132             
          ENDIF   
          SET ORDER TO 3             
          IF !USED('&K251')
             SELECT 0
             USE &K251 ALIAS K251
          ELSE
             SELECT K251             
          ENDIF             
          SET ORDER TO 1
          
          IF !USED('&K252')
             SELECT 0
             USE &K252 ALIAS K252
          ELSE
             SELECT K252             
          ENDIF             
          SET ORDER TO 1
          CREATE CURSOR SALE;
          (F01 C(5),F02 C(8),F03 C(5),F04 C(1),F05 C(4),F06 C(10),F07 C(8),F08 C(12),F09 C(5),F10 C(12),F11 C(4),F12 C(4),F13 C(10),F14 C(15),F15 C(10),F16 C(1),F17 C(20),F18 C(1),F99 C(2))
          INDEX ON F06 TAG SALE
          SET ORDER TO 1
          SELECT K39
          GO TOP
          DO WHILE !EOF()
             SELECT SALE
             APPEND BLANK
                           
                    REPLACE F01 WITH JST
                    REPLACE F02 WITH 'TOKUTSU'
                    
                    REPLACE F04 WITH '1'
                    REPLACE F05 WITH  PADL(K39.F01,4,'0')
                    REPLACE F06 WITH K39.F02
                    REPLACE F99 WITH IIF(SEEK(F06,'K251'),K251.F01,IIF(SEEK(F06,'K252'),K252.F01,' '))   
                    REPLACE F03 WITH IIF(SEEK(F06,'K251'),SUBSTR(K251.F19,2,3)+RIGHT(K251.F19,2),IIF(SEEK(F06,'K252'),SUBSTR(K252.F19,2,3)+RIGHT(K252.F19,2),' '))    
                    REPLACE F07 WITH IIF(SEEK(F06,'K251'),ICASE(K251.F16='2','!',K251.F16='4','-',K251.F04),IIF(SEEK(F06,'K252'),ICASE(K252.F16='2','!',K252.F16='4','-',K252.F04),'+')) 
                       REPLACE F08 WITH PADL(ALLTRIM(STR(IIF(SEEK(F06,'K251'),IIF(F99='31',K251.F08,K251.F12),IIF(SEEK(F06,'K252'),IIF(F99='31',K252.F08,K252.F12),0)))),12,'0')
                       REPLACE F09 WITH IIF(SEEK(F06,'K251'),IIF(F99='31',ICASE(K251.F09='1','5',K251.F09='2','04',K251.F09='3','*',''),ICASE(K251.F09='1','5',K251.F09='2','01',K251.F09='3','*','')),IIF(SEEK(F06,'K252'),IIF(F99='31',ICASE(K252.F09='1','5',K252.F09='2','04',K252.F09='3','*',''),ICASE(K252.F09='1','5',K252.F09='2','01',K252.F09='3','*','')),' '))
                       REPLACE F10 WITH PADL(ALLTRIM(STR(IIF(SEEK(F06,'K251'),K251.F10,IIF(SEEK(F06,'K252'),K252.F10,0)))),12,'0')
                       REPLACE F11 WITH IIF(SEEK(F06,'K251'),RIGHT(K251.F19,2)+K251.F02,IIF(SEEK(F06,'K252'),RIGHT(K252.F19,2)+K252.F02,' ')) 
                       REPLACE F13 WITH IIF(SEEK(F06,'K251'),K251.F03,IIF(SEEK(F06,'K252'),K252.F03,' '))
                    IF F99='32'
                       REPLACE F17 WITH IIF(SEEK(F06,'K251'),K251.F06,IIF(SEEK(F06,'K252'),K252.F06,' ')) 
   
                       REPLACE F18 WITH IIF(SEEK(F06,'K251'),K251.F05,IIF(SEEK(F06,'K252'),K252.F05,' ')) 
                    ELSE
                       IF F09='04'
                          REPLACE F18 WITH '1'
                       ENDIF   
                    ENDIF                    
                    REPLACE F16 WITH 'N'
                    SELECT K39
                    REPLACE F04 WITH ''
                    IF LEFT(SALE.F07,1)='+'
                       REPLACE F03 WITH CTOD('')
                       
                    ELSE
                       REPLACE F03 WITH CTOD(PADL(LEFT(SALE.F03,3),4,'0')+'/'+LEFT(SALE.F11,2)+'/'+RIGHT(SALE.F11,2))
                    ENDIF
             SKIP
          ENDDO
          SELECT SALE
          DO WHILE !BOF()             
          
             IF LEFT(F07,1)='+' AND (RIGHT(F06,2)='99' OR  RIGHT(F06,2)='49')
                REPLACE F07 WITH ' '
                DO WHILE .T.
                   SKIP -1
                   IF LEFT(F07,1)<>'+'
                      EXIT
                   ELSE
                      REPLACE F05 WITH PADL(ALLTRIM(STR(CEILING(RECNO()/50))),4,'0')
                      IF RIGHT(F06,2)='00' OR RIGHT(F06,2)='50'
                         REPLACE F07 WITH '*'
                      ELSE   
                         REPLACE F07 WITH SPACE(1)   
                      ENDIF   
                   ENDIF
                ENDDO                
                SKIP
                REPLACE F07 WITH '*'           
             ENDIF
             SKIP -1
          ENDDO          
          
          SELECT F01 送件年月,F02 用戶代號,PADL(ALLTRIM(STR(VAL(LEFT(TXMTH2,4))-1911)),3,'0')+RIGHT(TXMTH2,2) 所屬年月,F04 發票聯式,F05 第冊數,F06 發票號碼,IIF(LEFT(F07,1)='!','+'+SPACE(7),F07) 統一編號,F08 銷售額,F09 稅別,F10 稅額,F11 月日,F12 營業稅科目類別代號,F13 客戶代號,F14 公斤數,F18 通關註記,F15+F16 出售註記及列印註記 FROM SALE NOWAIT
          IF _TALLY>0
             FILE_NAME='銷項發票明細'+JST
             GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
             IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                 COPY TO (GCDELIMFILE) TYPE XL5
                 =MESSAGEBOX('儲存成功',0+48,'提示訊息視窗')
             ENDIF
          ENDIF                                                  
          SELECT F01 送件年月,F02 用戶代號,F06+SPACE(1) 發票號碼_序,F03+RIGHT(F11,2) 開立日期,PADR(IIF(SEEK(ALLTRIM(F13),'C01'),C01.F04,' '),12, ' ') 買受人,PADR(ALLTRIM(F07),8,SPACE(1)) 統一編號,;
          '電子連接器等'+SPACE(4) 貨物勞務,PADL(ALLTRIM(STR(TTL2(F06,F03))),12,'0') 數量,SUBSTR(F09,2,1) 外銷方式,SPACE(12) 備註,IIF(F99='31','發票'+SPACE(8),IIF(F18='1','X7'+SPACE(10),SPACE(12))) 非經海關證明文件名稱,;
          IIF(F99='31',F06+SPACE(12),IIF(F18='1',F17,SPACE(22))) 非經海關證明文件號碼,IIF(F18='2',PADL('0',12,'0'),F08) 非經海關證明文件金額,IIF(F99='32' AND F18='2','G3'+SPACE(2),SPACE(4)) 經海關出口報單類別,;
          IIF(F99='32' AND F18='2',F17,SPACE(20)) 經海關出口報單號碼,IIF(F99='32' AND F18='2',F08,PADL('0',12,'0')) 經海關出口報單金額,;
          F03+RIGHT(F11,2) 輸出或結匯日期,IIF(IIF(F99='31','1',F18)='1','Y',SPACE(1)) 外匯證明文件,F04 發票聯式,IIF(F99='31','1',F18) 通關註記;
          FROM SALE WHERE LEFT(F09,1)='0' AND LEFT(F07,1)<>'!'
          IF _TALLY>0
             FILE_NAME='零稅率清單'+JST
             GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
             IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XL5
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
             ENDIF      
          ENDIF               
          UPDATE K39 SET F04='+' WHERE F02=ANY(SELE F06 FROM SALE WHERE LEFT(F07,1)='+')
          SELECT SALE
          USE
          SELECT C01
          USE
          SELECT K251
          USE 
          SELECT K252
          USE
          SELECT C131
          USE
          SELECT C132
          USE
          SELECT B061
          USE 
          SELECT B062
          USE 
          SELECT K37
          THISFORM.GRID2.SETALL("DYNAMICBACKCOLOR","IIF(K39.F04='+',RGB(255,0,0),'')","COLUMN")
          THISFORM.GRID1.SETFOCUS
    ENDPROC
     PROCEDURE EPT_BOTT.CLICK    &&列印空白發票
       SELECT K39      
       SELECT F02 FROM K39 WHERE F04='+' INTO CURSOR K39_1 ORDER BY F02 NOWAIT 
       IF _TALLY>0
          SELECT K39_1
          GO TOP
          DO WHILE !EOF()
             SELECT K39.F02 FROM K39 WHERE K39.F02=K39_1.F02
             IF _TALLY>0
                REPORT FORM ALLTRIM(INT_116)+'INV_C' TO PRINT 
             ENDIF   
             SELECT K39_1
             SKIP  
          ENDDO    
          SELECT K39_1
          USE
       ELSE
          =MESSAGEBOX('無空白發票可列印!或是請先做轉申報檔後再列印',0+48,'提示訊息視窗')
       ENDIF
       SELECT K37
       THISFORM.GRID1.SETFOCUS
     ENDPROC
  ENDDEFINE        

*********************************************************************發票種類
DEFINE CLASS INV_TYPE AS OPTIONGROUP
  HEIGHT=INT_015*25
  TOP=INT_015*27
  LEFT=INT_015*64
  HEIGHT=INT_015*25
  WIDTH=INT_015*120
  FONTSIZE=INT_015*9
  BUTTONCOUNT=2
  VALUE=1
  NAME='INV_TYPE'
  OPTION1.CAPTION='三聯式'
  OPTION1.TOP=INT_015*5
  OPTION1.WIDTH=INT_015*52
  OPTION1.LEFT=INT_015*5
  OPTION1.NAME='OPTION1'
  OPTION2.CAPTION='二聯式'
  OPTION2.TOP=INT_015*5
  OPTION2.WIDTH=INT_015*52
  OPTION2.LEFT=INT_015*59
  OPTION2.NAME='OPTION2'
  PROC INTERACTIVECHANGE
       THISFORM.PAGEFRAME1.PAGE1.TXTF02.VALUE=IIF(THIS.Value=1,'31','32')
  ENDPROC     

ENDDEFINE  
*****
*****************************開立發票人員顯示自訂函數
FUNC INV_MAN
     PARA BKK_NO
     qt_od=SPACE(0)
     SELECT  &RTE_TABLE
     SET ORDER TO 1
     
     SEEK BKK_NO
     IF FOUND()
          DO WHILE F01=BKK_NO
                 qt_od=qt_od+PADR(ALLTRIM(F02),6,SPACE(1))+IIF(SEEK(F02,'A01'),A01.F03,'')
                SKIP
          ENDDO
          **qt_od=LEFT(qt_od,LEN(qt_od)-2)       
     ELSE
          qt_od=SPACE(0)
     ENDIF
     SELE &AREA1
     RETURN qt_od       
**********************************************************       編輯人員
DEFINE CLASS INVMAN_SEEK AS FORM
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
       CAPTION='人員編輯'+'    '+K37.F01+'冊'
       DRAGIT=.F.                                &&---紀錄可拖曳的項目       
       NODRAGIT=.F.                              &&---紀錄不可拖曳的項目
       ADD OBJECT LABEL1 AS LABEL WITH;
           LEFT=INT_015*20,;
           HEIGHT=INT_015*25,;
           AUTOSIZE=.T.,;
           CAPTION='未加入之人員' 
       ADD OBJECT LABEL2 AS LABEL WITH;
           LEFT=INT_015*270,;
           HEIGHT=INT_015*25,;
           AUTOSIZE=.T.,;
           CAPTION='已加入之人員'        
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
           TOOLTIPTEXT='未加入人員列表',;
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
            TOOLTIPTEXT='已加入人員列表',;             
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
             THISFORM.Caption='人員編輯'+'    '+K37.F01+'冊'
             THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')
             THISFORM.SETALL('FONTSIZE',INT_015*9,'LISTBOX')
             THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')
          ENDPROC   
          PROCEDURE LIST1.INIT        
                    with this
                         SELE A01_1
                         GO TOP
                         DO WHILE !EOF()                           
                            .ADDITEM(A01_1.F01+'  '+A01_1.F03)              
                             SKIP
                         ENDDO   
                         SELE A01_1
                         USE
                         .VALUE=1
                     endwith   
          ENDPROC                
          PROCEDURE LIST2.INIT
                    WITH THIS
                         SELE K38_1
                         GO TOP
                         DO WHILE !EOF()
                           .ADDITEM(K38_1.F02+'  '+K38_1.F03)
                           SKIP
                         ENDDO    
                         SELE K38_1
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
               DELE FROM K38 WHERE K38.F01=K37.F01
               SELE K38
               FOR I=1 TO THISFORM.LIST2.ListCount
                   APPEND BLANK
                   REPLACE  F01 WITH K37.F01
                   REPLACE  F03 WITH K37.F02
                  
                   REPLACE  F02 WITH THISFORM.LIST2.List(I)
                 
               ENDFOR                              
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
*********************************
*******************************************   計算該單出貨總數量
FUNC TTL2
     PARA BKK_NO,M_TH
     qt_od=0
     IF MOD(VAL(M_TH),2)=0
        SELECT C132
     ELSE
        SELECT C131
     ENDIF   
     SEEK BKK_NO
     IF FOUND()
         DO WHILE F17=BKK_NO                                      
            qt_od=qt_od+(f06)                                                    
            SKIP
         ENDDO
     ELSE
        IF MOD(VAL(M_TH),2)=0
           SELECT B062
        ELSE
           SELECT B061
        ENDIF   
        SEEK BKK_NO     
        IF FOUND()
           DO WHILE F20=BKK_NO                           
              qt_od=qt_od+(f04)                                                    
              SKIP
           ENDDO           
        ELSE
          qt_od=0
        ENDIF   
     ENDIF
     SELECT SALE
     RETURN qt_od
**********************************************列印發票
PROCEDURE PNT_PRC  
    S=0
    HK=''
 
     INV_RANGE=CREATEOBJECT("INV_RANGE")  
     INV_RANGE.SHOW    
     IF S=0
        =MESSAGEBOX('無此範圍資料',0+48,'警示')
     ENDIF   
     
     CLEAR     
ENDPROC
*******************************************************列印範圍選定
DEFINE CLASS INV_RANGE AS FORM
  AUTOCENTER=.T.
  CAPTION='請輸入列印資料範圍'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*150
  WIDTH=INT_015*500
  FONTSIZE=INT_015*9
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.    
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='ANS_RANGE'
  ADD OBJECT LABEL1 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*25,;
      TOP=INT_015*40,;
      NAME='LABEL1'
  ADD OBJECT TEXT1 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*34,;
      LEFT=INT_015*78,;
      HEIGHT=INT_015*25,; 
      NAME='TEXT1'
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*25,;
      TOP=INT_015*90,;
      NAME='LABEL2'      
  ADD OBJECT TEXT2 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      TOP=INT_015*84,;
      LEFT=INT_015*78,;
      HEIGHT=INT_015*25,;         
      NAME='TEXT2'    
 ADD OBJECT PYP_TPE AS OPTIONGROUP WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*50,;
      HEIGHT=INT_015*40,;    
      WIDTH=INT_015*120,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='列印出貨單發票',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='列印送貨單發票',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;          
      NAME='PYP_TPE'                                          
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*400,;     
      TOP=INT_015*40,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*70,;
      FONTSIZE=INT_015*10,;
      CAPTION='\<P.直接列印',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'

  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*400,;
      TOP=INT_015*80,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*70,;
      FONTSIZE=INT_015*10,;
      CAPTION='\<C.取   消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND3'    
  PROCEDURE INIT 
      THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')      
      THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')        
      THISFORM.LABEL1.CAPTION='起始冊號'
      THISFORM.LABEL2.CAPTION='截止冊號'
      THISFORM.TEXT1.WIDTH=INT_015*40
      THISFORM.TEXT2.WIDTH=INT_015*40 
      THISFORM.TEXT1.MAXLENGTH=4
      THISFORM.TEXT2.MAXLENGTH=4
      THISFORM.TEXT1.VALUE=K37.F01
      THISFORM.TEXT2.VALUE=K37.F01
      THISFORM.PYP_TPE.VALUE=1
  ENDPROC    
  PROCEDURE PYP_TPE.INTERACTIVECHANGE
    IF THIS.VALUE=2
       FTR_STR='1'
    ELSE
       FTR_STR='2'   
    ENDIF   
  ENDPROC
  PROCEDURE CMND1.CLICK      
    DIMENSION TXMTH(2)
          TXMTH(1)=LEFT(DTOS(CTOD(HKS+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(HKS,2)))),2,'0')
          TXMTH(2)=LEFT(DTOS(CTOD(HKS+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(HKS,2))+1)),2,'0')             
 
*!*	          B061='B06'+TXMTH(1)
*!*	          B062='B06'+TXMTH(2)
*!*	          B041='B04'+TXMTH(1)
*!*	          B042='B04'+TXMTH(2)

          **********************開檔案
          IF !USED('B01')
              SELECT 0
              USE B01
          ELSE
              SELECT B01
          ENDIF
          SET ORDER TO 1        
          IF !USED('C01')
              SELECT 0
              USE C01
          ELSE
              SELECT C01
          ENDIF
          SET ORDER TO 1        
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
          SET ORDER TO 1             
          IF FTR_STR='1'
             FOR K=1 TO 2
                B06='B06'+TXMTH(K)
                IF !USED('&B06')
                   SELECT 0
                   USE &B06 ALIAS B06
                ELSE
                   SELECT B06            
                ENDIF   
                SET ORDER TO 1      
                SELE F01,F02,F06,F20,F22,F23 FROM B06 WHERE F07='2' AND F20= ANY(SELE F02 FROM K39 WHERE F01>=THISFORM.TEXT1.VALUE AND F01<=THISFORM.TEXT2.VALUE) GROUP BY F01 ORDER BY F20 INTO CURSOR B06_2 
                IF _TALLY>0
                   CREATE CURSOR B06_1;
                   (F01 C(10),F02 C(2),F06 C(5),F20 C(10),F22 C(2),F23 C(1))
                   SELECT B06_2
                   GO TOP
                   DO WHILE !EOF()
                      SELECT B06_1
                      APPEND BLANK
                      REPLACE F01 WITH B06_2.F01
                      REPLACE F02 WITH B06_2.F02
                      REPLACE F06 WITH B06_2.F06
                      REPLACE F20 WITH B06_2.F20
                      REPLACE F22 WITH B06_2.F22
                      REPLACE F23 WITH B06_2.F23
                      SELECT B06_2
                      SKIP
                   ENDDO
                   SELECT B06_2
                   USE
                   SELECT B06_1
                   GO TOP
                   DO WHILE !EOF()
                         HK=ALLTRIM(STR(VAL(LEFT(TXMTH(K),4))-1911))+RIGHT(TXMTH(K),2)    
                      LK=B06_1.F01                                                  
                      SELE B06
                      CALC SUM(ROUND(F04*F14*F17,INT_069)) FOR F01=LK TO TMP1
                      
                      SELECT B06.F01,;
                      B06.F02,;
                      B06.F03,;
                      B06.F04,;
                      B06.F05,;
                      B06.F06,;
                      B06.F08,;
                      B06.F09,;
                      B06.F10,;
                      B06.F11,;
                      B06.F14,;
                      B06.F17,;
                      B06.F20,;
                      C01.F04,;
                      C01.F07,;
                      C01.F12,;
                      C01.F13,;
                      C01.F32,;
                      C01.F41,;
                      C01.F42,;
                      C03.F14,;
                      C04.F17,;             
                      B01.F02,;
                      B01.F99;
                      FROM B06,C01,C04,C03,B01 WHERE B06.F01=LK AND C01.F01=B06.F06 AND C04.F03+C04.F04=B06.F09+B06.F03 AND C03.F02=B06.F09 AND B01.F01=B06.F03 ORDER BY B06.F03 INTO CURSOR INV_DB NOWAIT
                      IF _TALLY>0
                         IF _tally<6
                            FLG='0'
                         ELSE
                            FLG='1'
                            SELECT * FROM INV_DB GROUP BY F01 NOWAIT
                         ENDIF
                       	    *REPORT FORM ALLTRIM(INT_116)+'INV_B' PREVIEW               
                            REPORT FORM ALLTRIM(INT_116)+'INV_B' TO PRINT 
                      ENDIF
                      SELECT B06_1
                      SKIP
                      S=S+1
                   ENDDO                
                ENDIF                                                       
                SELECT B06
                USE    
             ENDFOR
             
          ELSE
             FOR J=1 TO 2
                B04='B04'+TXMTH(J)
                IF !USED('&B04')
                   SELECT 0
                   USE &B04 ALIAS B04
                ELSE
                   SELECT B04           
                ENDIF   
                SET ORDER TO 1          
                SELE F01,F02,F06,F20,F22,F23 FROM B04 WHERE F10='2' AND  F20= ANY(SELE F02 FROM K39 WHERE F01>=THISFORM.TEXT1.VALUE AND F01<=THISFORM.TEXT2.VALUE) GROUP BY F01 ORDER BY F20 INTO CURSOR B04_2                                                        
                IF _TALLY>0
                   CREATE CURSOR B04_1;
                   (F01 C(10),F02 C(2),F06 C(5),F20 C(10),F22 C(2),F23 C(1))
                   SELECT B04_2
                   GO TOP
                   DO WHILE !EOF()
                      SELECT B04_1
                      APPEND BLANK
                      REPLACE F01 WITH B04_2.F01
                      REPLACE F02 WITH B04_2.F02
                      REPLACE F06 WITH B04_2.F06
                      REPLACE F20 WITH B04_2.F20
                      REPLACE F22 WITH B04_2.F22
                      REPLACE F23 WITH B04_2.F23
                      SELECT B04_2
                      SKIP
                   ENDDO
                   SELECT B04_2
                   USE
                   SELECT B04_1
                   GO TOP
                   DO WHILE !EOF()
                      HK=ALLTRIM(STR(VAL(LEFT(TXMTH(J),4))-1911))+RIGHT(TXMTH(J),2)        
                      LK=B04_1.F01       
                      SELE B04
                      CALC SUM(F17) FOR F01=LK TO TMP1
                                      
                      SELECT B04.F01,;
                      B04.F02,;
                      B04.F03,;
                      B04.F04,;
                      B04.F05,;
                      B04.F06,;
                      B04.F07,; 
                      B04.F09,;
                      B04.F12,;  
                      B04.F13,;
                      B04.F14,;
                      B04.F15,;
                      B04.F16,;
                      B04.F17,;
                      B04.F20,;
                      B04.F21,;
                      B04.F23,;
                      B04.F24,;
                      C01.F04,;
                      C01.F08,;
                      C01.F10,;
                      C01.F12,;
                      C01.F13,;
                      C01.F24,;
                      C01.F41,;
                      C01.F42,;
                      C03.F06,;
                      C03.F14,;
                      C04.F17,;
                      B01.F02,;
                      B01.F04,;            
                      B01.F99;
                      FROM B04,C01,C03,C04,B01 WHERE B04.F01=LK AND C01.F01=B04.F06 AND C03.F02=B04.F07 AND C04.F03=B04.F07 AND C04.F04=B04.F03;
                      AND B01.F01=B04.F03 ORDER BY B04.F03 INTO CURSOR INV_DB  NOWAIT     
                      IF _TALLY>0
                         IF _tally<6
                            FLG='0'
                         ELSE
                            FLG='1'
                            SELECT * FROM INV_DB GROUP BY F01 NOWAIT
                         ENDIF
                       	    REPORT FORM ALLTRIM(INT_116)+'INV_A' PREVIEW               
                            *REPORT FORM ALLTRIM(INT_116)+'INV_A' TO PRINT 
                      ENDIF
                      SELECT B04_1
                      SKIP
                      S=S+1
                   ENDDO        
                   
                ENDIF                
                SELECT B04
                USE
             ENDFOR
          ENDIF   

           SELECT B01
           USE
           SELECT C01
           USE
           SELECT C03
           USE
           SELECT C04
           USE                                                               
     THISFORM.RELEASE    
  ENDPROC

  ****
  PROCEDURE CMND3.CLICK
      S=1          
      THISFORM.RELEASE
  ENDPROC        
ENDDEFINE  



