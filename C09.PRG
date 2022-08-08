**程式名稱:C09客戶訂單取消紀錄
set excl off
close all
clear 
AREA1='C09'
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
SEEK sys_oper+'C09'      
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1             
IF !USED('C01')
   SELE 0
   USE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO C011
IF !USED('C09')
   SELE 0
   USE C09
ELSE
   SELE C09
ENDIF
SET ORDER TO C091
SET RELATION TO F06 INTO C01
GO TOP
C09form=createobject("tkC09")
C09form.show  
define class tkC09 as form
  caption='C09.客戶訂單取消紀錄'
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
      RECORDSOURCE='C09',; 
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
       SELE C09
       GO TOP
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限   
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)  
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')   
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
       THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
       THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
       THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
       THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120             
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='訂單號碼'	          
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='取消日期'
	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='取消單號'	   	   
  	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='取消數量' 
  	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='客戶編號'   	   
  	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='客戶簡稱'   
  	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='客戶單號'     	   	     	      	     	      
  	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'     	   	     	      	     	        	   
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C09.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F03'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C09.F04'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F05'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F06'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C01.F05'
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F07'       
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'       
       THISFORM.GRID1.READONLY=.T.  
       THISFORM.GRID1.COLUMN9.READONLY=.F.     
       JK='訂單編號'
       THISFORM.GRID1.SETFOCUS
  ENDPROC          
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex       
       THISFORM.REC_COUNT.CAPTION=C09.F09
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROC KEY_LIST.INIT
       with this 
           .additem('依訂單號碼+料品編號排列')
           .additem('依料品編號+取消日期排列')
           .additem('依客戶編號+取消日期排列')
           .additem('依客戶單號+料品編號排列')
           .additem('依取消單號+料品編號排列')           
       endwith     
  ENDPROC
  PROC KEY_LIST.INTERACTIVECHANGE
       SELE C09
       DO CASE
          CASE THIS.DISPLAYVALUE='依訂單號碼+料品編號排列'
               SET ORDER TO C091
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120              
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='訂單號碼'	          
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='取消日期'
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='取消單號'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='取消數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='客戶編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='客戶簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='客戶單號'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'     	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F01'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C09.F02'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F03'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C09.F04'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F05'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F06'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C01.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F07'       
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'       
          CASE THIS.DISPLAYVALUE='依料品編號+取消日期排列'
               SET ORDER TO C092
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*65               
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120         
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料        號'	                 
	           THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='取消日期'                   
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='訂單號碼'	          
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='取消單號'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='取消數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='客戶編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='客戶簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='客戶單號'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'     
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F02'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C09.F03'        	   	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F01'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C09.F04'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F05'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F06'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C01.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F07'       
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'       
          CASE THIS.DISPLAYVALUE='依客戶編號+取消日期排列'
               SET ORDER TO C093
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*40  
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*65               
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                                         
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120        
  	           THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='客戶編號'   	
  	           THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='客戶簡稱'     	              
	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='取消日期'                     	                          
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='料        號'	                 
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='訂單號碼'	          
	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='取消單號'	   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='取消數量' 
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='客戶單號'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F06'    
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C01.F05'               
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F03'        	   	   	     	      	     	        	                      	     
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C09.F02'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F01'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F04'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C09.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F07'       
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'             
          CASE THIS.DISPLAYVALUE='依客戶單號+料品編號排列'
               SET ORDER TO C094
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80    
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149                                 
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*40  
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65                
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65                                         
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120          
         	   THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='客戶單號' 
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                          	       	   	     	      	     	                     
  	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='客戶編號' 
  	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='客戶簡稱'     	             	   
	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='取消日期'                     	                          
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='訂單號碼'	          
	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='取消單號'	   	   
  	           THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='取消數量' 
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F07'      
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C09.F02'                        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F06'    
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C01.F05'               
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F03'        	   	   	     	      	     	        	                      	     
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F01'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C09.F04'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F05'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'             
          CASE THIS.DISPLAYVALUE='依取消單號+料品編號排列'
               SET ORDER TO C095
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80   
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149                           
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120         
	           THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='取消單號'	   
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 	           	                      
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='訂單號碼'	          
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='取消日期'
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='取消數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='客戶編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='客戶簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='客戶單號'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='備註說明'     
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='C09.F04'    
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='C09.F02'                   	   	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='C09.F01'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='C09.F03'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='C09.F05'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='C09.F06'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='C01.F05'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='C09.F07'       
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='C09.F08'                                  	                          
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
          
*              REPORT FORM C09 PREVIEW
*              report form C09 to print prompt        
          
ENDPROC

****************************************
***********************************************轉EXCEL日期範圍選定
define class EXCEL_CHK as form
  autocenter=.t.
  caption='請輸入轉檔資料範圍'
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
   	  FILE_NAME='C09客戶訂單取消紀錄'
          SELECT C09.F01 訂單編號,C09.F02 料品編號,C09.F03 取消日期,C09.F04 取消單號,C09.F05 取消數量,C09.F06 客戶編號,C01.F05 客戶簡稱,C09.F07 客戶單號,C09.F08 備註說明,C09.F09 安全資料 ;
          FROM C09,C01;
          WHERE C09.F03>= THISFORM.TEXT1.VALUE AND C09.F03<= THISFORM.TEXT2.VALUE AND C01.F01=C09.F06 ORDER BY C09.F01,C09.F02 NOWAIT
          
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

