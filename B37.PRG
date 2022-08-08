**程式名稱:B37料品調撥紀錄
set excl off
close all
clear 
AREA1='B13'
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
SEEK sys_oper+'B37'                  
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
IF !USED('B13')
   SELE 0
   USE B13
ELSE
   SELE B13
ENDIF
SET ORDER TO B131
SET RELATION TO F05 INTO A14
GO TOP
B37form=createobject("tkB37")
B37form.show  
define class tkB37 as form
  caption='B37.料品調撥紀錄'
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
      RECORDSOURCE='B13',; 
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
       SELE B13
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
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='調撥單號'	          
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='調撥日期'
	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='撥出對象'	   	   
  	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥數量' 
  	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='部門編號'   	   
  	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='部門簡稱'   
  	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='撥入對象'     	   	     	      	     	      
  	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'     	   	     	      	     	        	   
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B13.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B13.F03'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F02'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F04'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B13.F05'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'       
       THISFORM.GRID1.READONLY=.T.  
       THISFORM.GRID1.COLUMN9.READONLY=.F.     
       JK='調撥單號'
       THISFORM.GRID1.SETFOCUS
  ENDPROC          
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex       
       THISFORM.REC_COUNT.CAPTION=B13.F10
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROC KEY_LIST.INIT
       with this 
           .additem('依調撥單號+料品編號排列')
           .additem('依料品編號+調撥日期排列')
           .additem('依部門編號+調撥日期排列')
           .additem('依撥入對象+料品編號排列')
           .additem('依撥出對象+料品編號排列')           
       endwith     
  ENDPROC
  PROC KEY_LIST.INTERACTIVECHANGE
       SELE B13
       DO CASE
          CASE THIS.DISPLAYVALUE='依調撥單號+料品編號排列'
               SET ORDER TO B131
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120              
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='調撥單號'	          
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='調撥日期'
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='撥出對象'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='部門編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='部門簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='撥入對象'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'     	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B13.F01'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B13.F03'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F02'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F04'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B13.F05'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'       
          CASE THIS.DISPLAYVALUE='依料品編號+調撥日期排列'
               SET ORDER TO B132
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
	           THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='調撥日期'                   
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='調撥單號'	          
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='撥出對象'	   	   
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='部門編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='部門簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='撥入對象'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'     
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B13.F03'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B13.F02'        	   	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F01'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F04'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B13.F05'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'       
          CASE THIS.DISPLAYVALUE='依部門編號+調撥日期排列'
               SET ORDER TO B133
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*40  
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*65               
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                                         
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120        
  	           THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='部門編號'   	
  	           THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='部門簡稱'     	              
	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='調撥日期'                     	                          
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='料        號'	                 
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥單號'	          
	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='撥出對象'	   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='調撥數量' 
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='撥入對象'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B13.F05'    
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A14.F02'               
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F02'        	   	   	     	      	     	        	                      	     
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B13.F03'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F01'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='B13.F04'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'             
          CASE THIS.DISPLAYVALUE='依撥入對象+料品編號排列'
               SET ORDER TO B134
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80    
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149                                 
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*40  
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65                
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*65                                         
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120          
         	   THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='撥入對象' 
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                          	       	   	     	      	     	                     
  	           THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='部門編號' 
  	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='部門簡稱'     	             	   
	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥日期'                     	                          
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='調撥單號'	          
	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='撥出對象'	   	   
  	           THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='調撥數量' 
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'   
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B13.F03'                        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F05'    
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='A14.F02'               
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F02'        	   	   	     	      	     	        	                      	     
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B13.F01'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='B13.F04'
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'             
          CASE THIS.DISPLAYVALUE='依撥出對象+料品編號排列'
               SET ORDER TO B135
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80   
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149                           
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*40
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*65
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80       
               THISFORM.GRID1.COLUMN9.WIDTH=INT_015*120         
	           THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='撥出對象'	   
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 	           	                      
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='調撥單號'	          
	           THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='調撥日期'
  	           THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='調撥數量' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='部門編號'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='部門簡稱'   
         	   THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='撥入對象'     	   	     	      	     	      
        	   THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='調撥用途'     
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE="IIF(SEEK(B13.F07,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B13.F03'                   	   	   	     	      	     	        	   
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='B13.F01'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='B13.F02'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='B13.F04'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='B13.F05'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='A14.F02'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE="IIF(SEEK(B13.F08,'C01'),C01.F05,'')"
               THISFORM.GRID1.COLUMN9.CONTROLSOURCE='B13.F09'                                  	                          
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
          
*              REPORT FORM B13 PREVIEW
*              report form B13 to print prompt        
          
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
   	  FILE_NAME='B37料品調撥紀錄'
          SELECT B13.F01 調撥單號,B13.F03 料品編號,B13.F02 調撥日期,B13.F07 撥出對象,IIF(SEEK(B13.F07,'C01'),C01.F05,SPACE(5)) 對象簡稱1,;
          B13.F04 調撥數量,B13.F05 部門編號,A14.F02 部門簡稱,B13.F08 撥入對象,IIF(SEEK(B13.F08,'C01'),C01.F04,SPACE(5)) 對象簡稱2,B13.F09 調撥用途,B13.F10 安全資料;
          FROM B13,A14;
          WHERE B13.F02>= THISFORM.TEXT1.VALUE AND B13.F02<= THISFORM.TEXT2.VALUE AND A14.F01=B13.F05 ORDER BY B13.F01,B13.F03 NOWAIT
          
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

