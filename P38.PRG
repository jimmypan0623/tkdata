**程式名稱:P38.採購交期異動表
set excl off
close all
clear 
AREA1='P38'
FLG='0'
FCH=''
CHK=''
IF !USED('A01')
   SELE 0
   USE A01
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
SEEK sys_oper+'P38'      
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1 
IF !USED ('B01')
   SELE 0
   USE B01 
ELSE 
   SELE B01
ENDIF
SET ORDER TO 1              

IF !USED('P38')
   SELE 0
   USE P38
ELSE
   SELE P38
ENDIF
SET ORDER TO P382                
SET FILTER TO !EMPTY(F05)
GO TOP
P38form=createobject("tkP38")
P38form.show  
define class tkP38 as form
  caption='P38.外購交期異動紀錄'
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
  ADD OBJECT LBL_REC AS LBL_REC WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*630,;
      WIDTH=INT_015*60.,;
      CAPTION='安全資料'
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      LEFT=INT_015*680,;
      TOP=INT_015*5
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=9,;            
      RECORDSOURCE='P38',; 
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
      LEFT=INT_015*400              
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=INT_015*0,;
      LEFT=INT_015*230,;
      ENABLED=.T.,;
      VISIBLE=.T.                 
 ADD OBJECT EXCEL_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*459,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<T.轉EXCEL'                                           
  PROCEDURE INIT
       SELE P38
       GO TOP
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限   
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)  
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')   
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF((DATE()-CTOD(LEFT(P38.F08,8)))<4 .AND. (DATE()-CTOD(LEFT(P38.F08,8)))=>0,RGB(0,255,255),'')","COLUMN")
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
       THISFORM.GRID1.COLUMN6.WIDTH=INT_015*85
       THISFORM.GRID1.COLUMN7.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60
       THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='採購單號'	          
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='原交期'
	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='原數量'	   	   
  	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='新交期' 
  	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='新數量'   	   
  	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='採購用途'   	     	   
  	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最初交期'   
  	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='異動單號'      	     	      
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='P38.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='P38.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='P38.F03'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='P38.F04'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='P38.F05'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='P38.F06'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='P38.F07'
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='P38.F09'
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='P38.F10'
       THISFORM.GRID1.READONLY=.T.  
       THISFORM.GRID1.COLUMN9.READONLY=.F.     
       JK='採購單號'
       THISFORM.GRID1.SETFOCUS
  ENDPROC          
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex       
      * THISFORM.GRID1.TOOLTIPTEXT=IIF((DATE()-CTOD(LEFT(P38.F08,8)))<4 .AND. (DATE()-CTOD(LEFT(P38.F08,8)))>0,LEFT(P38.F08,8)+'異動','')
       THISFORM.REC_COUNT.CAPTION=P38.F08
       THISFORM.GRID1.TOOLTIPTEXT=ALLTRIM(LEFT(P38.F08,8))+'異動'
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROC KEY_LIST.INIT
       with this 
           .additem('依採購單號+料號+  原交期排列')
           .additem('依料號+採購單號+最初交期排列')
           .additem('依採購用途+料號+最初交期排列')
           .additem('依異動單號+料號+  原交期排列')
       endwith     
  ENDPROC
  PROC KEY_LIST.INTERACTIVECHANGE
       SELE P38
       DO CASE
          CASE THIS.DISPLAYVALUE='依採購單號+料號+  原交期排列'
               SET ORDER TO P381
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='採購單號'	          
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
       	       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='原交期'
        	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='原數量'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='新交期' 
        	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='新數量'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='採購用途'   	     	   
        	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='最初交期'   
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='異動單號'	     	      
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='P38.F01'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='P38.F02'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='P38.F03'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='P38.F04'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='P38.F05'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='P38.F06'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='P38.F07'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='P38.F09'    
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='P38.F10'                          
          CASE THIS.DISPLAYVALUE='依料號+採購單號+最初交期排列'
               SET ORDER TO P385
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60               
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80               
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料        號'	                 	          
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='採購單號'
       	       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='最初交期'
        	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='原交期'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='原數量' 
        	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='新交期'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='新數量'   	     	   
        	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='採購用途'   	  
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='異動單號'	     	              	      	      
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='P38.F02'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='P38.F01'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='P38.F09'                                             
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='P38.F03'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='P38.F04'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='P38.F05'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='P38.F06'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='P38.F07'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='P38.F10'               
          CASE THIS.DISPLAYVALUE='依採購用途+料號+最初交期排列'
               SET ORDER TO P383
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149               
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60                              
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80               
        	   THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='採購用途'   	     	                     
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 	          
       	       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='最初交期'
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='採購單號'       	       
        	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='原交期'	   	   
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='原數量' 
        	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='新交期'   	   
  	           THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='新數量'   	
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='異動單號'	     	        	                	   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='P38.F07'                 	           
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='P38.F02'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='P38.F09'                                                            
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='P38.F01'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='P38.F03'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='P38.F04'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='P38.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='P38.F06'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='P38.F10'               
          CASE THIS.DISPLAYVALUE='依異動單號+料號+  原交期排列'
               SET ORDER TO P384
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80               
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*60                              
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*85
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*149
        	   THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='異動單號'   	     	                                    
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	       
        	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='原交期'	   	                            	          
       	       THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='最初交期'
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='採購單號'       	       
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='原數量' 
        	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='新交期'   	   
  	           THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='新數量'   	     	   
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='採購用途'   	     	                       	           
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='P38.F10'                                                     
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='P38.F02'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='P38.F03'               
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='P38.F09'                                                            
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='P38.F01'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='P38.F04'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='P38.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='P38.F06'    
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='P38.F07'                 	                          
       ENDCASE  
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION       
       THISFORM.GRID1.SETFOCUS             
  ENDPROC
    PROCEDURE EXCEL_BOTT.CLICK
       FLG='0'
       EXCEL_CHK=CREATEOBJECT("EXCEL_CHK")  
       EXCEL_CHK.SHOW    
       IF FLG='0'
            =MESSAGEBOX('無此資料範圍',0+48,'提示訊息視窗')
        ELSE
            FLG='0'
        ENDIF    
       THISFORM.GRID1.SETFOCUS
  ENDPROC 
enddefine        
***********************************************列印程序
PROCEDURE PNT_PRC     
          
*              REPORT FORM P38 PREVIEW
*              report form P38 to print prompt        
          
ENDPROC

****************************************
***********************************************轉EXCEL日期範圍選定
define class EXCEL_CHK as form
  autocenter=.t.
  caption='請輸入列印資料範圍'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*190
  WIDTH=INT_015*480
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='EXCEL_CHK'
  add object label1 as label with;
      autosize=.t.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*30,;
      TOP=INT_015*75,;
      name='label1'
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*250,;
      TOP=INT_015*75,;
      NAME='LABEL2'
  add object text1 as textbox with;
      autosize=.t.,;
      TOP=INT_015*70,;
      LEFT=INT_015*80,;
      HEIGHT=INT_015*25,; 
      name='text1'
  add object text2 as textbox with;
      autosize=.t.,;
      TOP=INT_015*70,;
      LEFT=INT_015*300,;
      HEIGHT=INT_015*25,;         
      name='text2'    
  add object cmnd1 as commandbutton with;
      autosize=.t.,;
      LEFT=INT_015*100,;    
      HEIGHT=INT_015*25,;   
      TOP=INT_015*144,;  
      WIDTH=INT_015*80,;
      caption='\<Y.執行',;
      TOOLTIPTEXT='確認所輸入的範圍鍵值!快速鍵->ALT+Y',;
      name='cmnd1'
  add object cmnd2 as commandbutton with;
      autosize=.t.,;
      LEFT=INT_015*300,;
      TOP=INT_015*144,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*80,;
      caption='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      name='cmnd2'  
 ******************************
 procedure init 
      THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')      
      THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON')            

              THISFORM.LABEL1.CAPTION='起始日期'
              THISFORM.LABEL2.CAPTION='截止日期'
              THISFORM.TEXT1.WIDTH=INT_015*70
              THISFORM.TEXT2.WIDTH=INT_015*70            
              THISFORM.TEXT1.VALUE=DATE()
              THISFORM.TEXT2.VALUE=DATE()              


  ENDPROC
******************************
procedure cmnd1.click 
   	  FILE_NAME='P38外購交期異動紀錄'
          SELECT P05.F05 需求對象,C01.F18 業務員編號,A01.F03 姓名,P38.F01 採購單號,P38.F02 料號,P38.F03 原交期,P38.F04 原數量,P38.F05 新交期,P38.F06 新數量,LEFT(P38.F07,10) 請購單號,;
          P38.F09 最初交期,CTOD(LEFT(P38.F08,8)) 異動日期,P38.F10 異動單號 FROM P38,P05,C01,A01;
          WHERE CTOD(LEFT(P38.F08,8))>= THISFORM.TEXT1.VALUE AND CTOD(LEFT(P38.F08,8))<= THISFORM.TEXT2.VALUE AND !EMPTY(P38.F05) AND !EMPTY(P38.F10) AND P05.F01=LEFT(P38.F07,10) AND;
          C01.F01=P05.F04 and A01.F01=C01.F18 ORDER BY P05.F04,P38.F01,P38.F02,P38.F04
	 GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	 IF EMPTY(GCDELIMFILE)  && ESC PRESSED
             RETURN
          ELSE 		  
              ON ERROR DO FILE_IMPACT          
              IF _tally >0              
                   COPY TO (GCDELIMFILE) TYPE XL5   
		   =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗') 

                 FLG='1'
              ENDIF
        ENDIF        
       THISFORM.RELEASE                         
  endproc           
******************************      
  procedure cmnd2.click
      FLG='1'   
      thisform.release
  endproc      
enddefine        

