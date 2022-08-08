**程式名稱:E07.檢視MRP數據
set excl off
close all
clear 
AREA1='E08'
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
SEEK sys_oper+'E07'      
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1 
IF !USED('P05')
   SELE 0
   USE P05
ELSE
   SELE P05
ENDIF
SET ORDER TO P051
IF !USED('C01')
   SELE 0
   USE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO C011
IF !USED('C03')
   SELE 0
   USE C03
ELSE
   SELE C03
ENDIF
SET ORDER TO C031
IF !USED('D01')
   SELE 0
   USE D01
ELSE
   SELE D01
ENDIF
SET ORDER TO D011
IF !USED('A14')
   SELE 0
   USE A14
ELSE
   SELE A14
ENDIF
SET ORDER TO A141                  
IF !USED('B11')
   SELE 0
   USE B11
ELSE
   SELE B11
ENDIF
SET ORDER TO B112    
       
IF !USED('E08')
   SELE 0
   USE E08
ELSE
   SELE E08
ENDIF
SET ORDER TO E081        
DELE FOR F05+F08+F11+F14+F17=0
GO TOP
******************
E07form=createobject("tkE07")
E07form.show  
define class tkE07 as form
  caption='E07.檢視MRP數據'+space(40)+'計算週期：'+iif(int_005='1','日',iif(int_005='2','週','月'))
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
      AUTOCENTER =.T.,;
      CAPTION='目前存量',;
      TOP=INT_015*5
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      TOP=INT_015*5      
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.,;
      WIDTH=INT_015*130
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=7,;            
      RECORDSOURCE='E08',; 
      HEIGHT=INT_015*500,;     
      WIDTH=INT_015*780,;     
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',; 
      COLUMN4.NAME='COLUMN4',; 
      COLUMN5.NAME='COLUMN5',;             
      COLUMN6.NAME='COLUMN6',;       
      COLUMN7.NAME='COLUMN7'
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;      
      TOP=INT_015*0,;
      LEFT=INT_015*400              
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=INT_015*0,;
      LEFT=INT_015*140,;
      ENABLED=.T.,;
      VISIBLE=.T. 
  ADD OBJECT IODT_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*360,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      CAPTION='\<Z.進出明細',;
      FONTSIZE=INT_015*9,;
      ENABLED=.T.,;
      VISIBLE=.T.
 ADD OBJECT EXCEL_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*459,;
      HEIGHT=INT_015*25,;      
      WIDTH=INT_015*65,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<T.轉EXCEL'                                     
**********************                      
  PROCEDURE INIT
       SELE E08
       GO TOP
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限     
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')   
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*60
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*200
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*100
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*100
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*100
       THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100
       THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='週        期'	          
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預計入貨'
	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='預計產出'	   	   
  	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='預計領用' 
  	   THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預計出貨'   	   
  	   THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期結餘'   	     	  
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='E08.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='E08.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='E08.F05'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='E08.F08'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='E08.F11'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='E08.F14'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE="E08.F20+ALTT(E08.F02,'1')"
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(E08.F20+ALTT(E08.F02,'1')<0 AND E08.F11+E08.F14+E08.F17>0,RGB(255,0,0),IIF(E08.F20+ALTT(E08.F02,'1')>0 AND E08.F05+E08.F08>0,RGB(255,255,0),''))","COLUMN")                                 
       THISFORM.GRID1.READONLY=.T.       
       JK='週        期'
       THISFORM.GRID1.SETFOCUS
  ENDPROC
***********************            
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex       
       THISFORM.REC_COUNT.CAPTION=TRANSFORM(ALTT(E08.F02,'1'),'@R 999999999999.999')
       FCH=THISFORM.ACTIVECONTROL.NAME       
       CHK=''
       thisform.refresh
  ENDPROC
***********************         
  PROC KEY_LIST.INIT
       with this 
           .additem('依日期+料號排列')
           .additem('依料號+日期排列')
       endwith     
  ENDPROC
*******************  
  PROC KEY_LIST.INTERACTIVECHANGE
       SELE E08
       DO CASE
          CASE THIS.DISPLAYVALUE='依日期+料號排列'
               SET ORDER TO E081
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*200
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='週        期'	          
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料        號'	                 
        	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預計入貨'
        	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='預計產出'	   	   
        	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='預計領用' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預計出貨'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期結餘'   	     
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='E08.F01'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='E08.F02'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='E08.F05'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='E08.F08'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='E08.F11'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='E08.F14'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="E08.F20+ALTT(E08.F02,'1')"  	           	      
          CASE THIS.DISPLAYVALUE='依料號+日期排列'
               SET ORDER TO E082
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*200
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*100            
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料        號'	                 
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='週        期'	                         
        	   THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預計入貨'
        	   THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='預計產出'	   	   
        	   THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='預計領用' 
  	           THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預計出貨'   	   
  	           THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='預期結餘'   	     	      
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='E08.F02'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='E08.F01'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='E08.F05'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='E08.F08'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='E08.F11'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='E08.F14'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE="E08.F20+ALTT(E08.F02,'1')"               
       ENDCASE  
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION       
       THISFORM.GRID1.SETFOCUS      
  ENDPROC
  PROC IODT_BOTT.CLICK
            LK=E08.F02
            LV=E08.F01
            DO DETAL
            VHJ=E08.F20+ALTT(E08.F02,'1')    
            K=0                       
            SELE INSCH
            GO BOTT
            DO WHILE !BOF()
               REPL F08 WITH VHJ-K
               VHJ=VHJ-K
               K=F04
               SKIP -1
            ENDDO   
            
               INSCH_SEEK=CREATEOBJECT("INSCH_SEEK")  
               INSCH_SEEK.SHOW           

            SELE E08
           THISFORM.GRID1.SETFOCUS  
  ENDPROC
  PROCEDURE EXCEL_BOTT.CLICK
      SELE E08
       EXCEL_CHK=CREATEOBJECT("EXCEL_CHK")  
       EXCEL_CHK.SHOW    
       THISFORM.GRID1.SETFOCUS
  ENDPROC 
enddefine  
***********************************************列印範圍選定
define class ANS_RANGE as form
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
  name='ANS_RANGE'
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
      DO CASE
         CASE KEY()='DTOS(F01)+F02'    
              THISFORM.LABEL1.CAPTION='起始日期'
              THISFORM.LABEL2.CAPTION='截止日期'
              THISFORM.TEXT1.WIDTH=INT_015*70
              THISFORM.TEXT2.WIDTH=INT_015*70            
              THISFORM.TEXT1.VALUE=E08.F01
              THISFORM.TEXT2.VALUE=E08.F01              
         CASE KEY()='F02+DTOS(F01)'
              THISFORM.LABEL1.CAPTION='起始料號'
              THISFORM.LABEL2.CAPTION='截止料號'
              THISFORM.TEXT1.WIDTH=INT_015*150
              THISFORM.TEXT2.WIDTH=INT_015*150
              THISFORM.TEXT1.MAXLENGTH=20              
              THISFORM.TEXT2.MAXLENGTH=20              
              THISFORM.TEXT1.VALUE=E08.F02
              THISFORM.TEXT2.VALUE=E08.F02                                   
      ENDCASE   
  ENDPROC
******************************
procedure cmnd1.click 
  DO CASE
      CASE KEY()='DTOS(F01)+F02'    
         SELECT E08.*,ALTT(E08.F02,'1') FROM E08 WHERE;
 		 E08.F01>= THISFORM.TEXT1.VALUE AND E08.F01<= THISFORM.TEXT2.VALUE ;
 		  ORDER BY E08.F01,E08.F02 NOWAIT 
           IF _tally >0              
              REPORT FORM ALLTRIM(INT_116)+'E07_A' PREVIEW 
*!*	               report form E08_A to print prompt        
              FLG='1'
           ENDIF  
      CASE KEY()='F02+DTOS(F01)' 
            SELECT E08.*,ALTT(E08.F02,'1') FROM E08 WHERE  ;
 		    E08.F02>= THISFORM.TEXT1.VALUE  AND E08.F02<= THISFORM.TEXT2.VALUE ;
 		    ORDER BY E08.F02,E08.F01 NOWAIT
            IF _tally >0
              REPORT FORM ALLTRIM(INT_116)+'E07_B' PREVIEW
*!*	               report form E08_B to print prompt        
              FLG='1'
           ENDIF                                 
       ENDCASE 
       THISFORM.RELEASE                         
  endproc           
******************************      
  procedure cmnd2.click
      FLG='1'   
      thisform.release
  endproc      
enddefine        
***********************************************列印程序
PROCEDURE PNT_PRC  
   FLG='0'
   SELE E08
   ANS_RANGE=createobject("ANS_RANGE")  
   ANS_RANGE.SHOW    
   IF FLG<>'1'
       =MESSAGEBOX('無此範圍資料',0+48,'警示')
   ENDIF   
   FLG='0'     
ENDPROC
***************************************週期內料況明細
DEFINE CLASS INSCH_SEEK AS FORM
  AUTOCENTER=.T.
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*320
  WIDTH=INT_015*725
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='INSCH_SEEK'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;      
      TOP=INT_015*290,;
      LEFT=INT_015*0
  ADD OBJECT GRID1 AS GRID WITH;          
      AUTOCENTER=.T.,;
      WIDTH=INT_015*725,;      
      HEIGHT=INT_015*293,;     
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*19,; 
      DELETEMARK=.F.,;
      READONLY=.T.,;
      RECORDSOURCETYPE=1,;
      COLUMNCOUNT=8,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;      
      COLUMN6.NAME='COLUMN6',;      
      COLUMN7.NAME='COLUMN7',;  
      COLUMN8.NAME='COLUMN8'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*295,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='取消此作業畫面!快速鍵->ALT+C'
      NAME='CMND2'
      PROCEDURE INIT 
         THISFORM.CAPTION=E08.F02+SPACE(5)+DTCX(E08.F01)
         AREA1='INSCH'  
         SELE INSCH
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
         THISFORM.GRID1.READONLY=.T. 
         THISFORM.GRID1.RECORDSOURCE='INSCH'
         THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.         
         THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
         THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(INSCH.F99<>'1',RGB(255,255,0),'')","COLUMN")                                   
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='單據號碼'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='INSCH.F01'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*83
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='異動日期'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='INSCH.F03'         
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預期異動數量'  
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*110
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='INSCH.F04'   
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='預期結餘'  
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*110
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='INSCH.F08'  
         THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='客戶簡稱'  
         THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
         THISFORM.GRID1.COLUMN5.CONTROLSOURCE="IIF(SEEK(INSCH.F05,'C01'),C01.F05,'')"
         THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='業務人員'  
         THISFORM.GRID1.COLUMN6.WIDTH=INT_015*70
         THISFORM.GRID1.COLUMN6.CONTROLSOURCE="IIF(SEEK(INSCH.F06,'A01'),A01.F03,'')"
         THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='確認'  
         THISFORM.GRID1.COLUMN7.WIDTH=INT_015*20
         THISFORM.GRID1.COLUMN7.CONTROLSOURCE="IIF(INSCH.F99='1','是','否')"                   
         THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='備註說明'  
         THISFORM.GRID1.COLUMN8.WIDTH=INT_015*261
         THISFORM.GRID1.COLUMN8.CONTROLSOURCE='INSCH.F07'            
         THISFORM.GRID1.SETFOCUS 
      ENDPROC 
      PROC GRID1.AFTERROWCOLCHANGE
           LPARAMETERS XCOLINDEX
      ENDPROC 
      PROCEDURE CMND2.CLICK
          SELE INSCH
          USE
          AREA1='E08'
         THISFORM.RELEASE
      ENDPROC    
ENDDEFINE   
********************************************************
DEFINE CLASS EXCEL_CHK AS FORM
  AUTOCENTER=.T.
  CAPTION='請選擇資料轉EXCEL之選項'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*150
  WIDTH=INT_015*350
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='EXCEL_CHK '
ADD OBJECT DATA_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*67,;
      TOP=INT_015*10,;
      AUTOSIZE=.T.,;
      FONTSIZE=INT_015*12,;
      BUTTONCOUNT=3,;
      OPTION1.CAPTION='預期結餘大於 0',;
      OPTION1.TOP=0,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='預期結餘小於 0',;
      OPTION2.TOP=INT_015*20,;
      OPTION2.AUTOSIZE=.T.,;
      OPTION3.CAPTION='全部資料',;
      OPTION3.TOP=INT_015*40,;
      OPTION3.AUTOSIZE=.T.,;      
      NAME='DATA_OPTION'       
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*193,;     
      TOP=INT_015*11,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*193,;
      TOP=INT_015*51,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND2'      
  ADD OBJECT LBL1 AS LABEL WITH;
       LEFT=INT_015*5,;
       TOP=INT_015*110,;
       FONTSIZE=INT_015*9,;
       AUTOSIZE=.T.,;
       CAPTION=''      
  ADD OBJECT LBL2 AS LABEL WITH;
       LEFT=INT_015*5,;
       TOP=INT_015*135,;
       FONTSIZE=INT_015*9,;
       AUTOSIZE=.T.,;
       CAPTION=''             
  ADD OBJECT LBL3 AS LABEL WITH;
       LEFT=INT_015*90,;
       TOP=INT_015*82,;
       FONTSIZE=INT_015*9,;
       AUTOSIZE=.T.,;
       CAPTION='到'                    
  ADD OBJECT TXT1 AS TEXTBOX WITH;
       LEFT=INT_015*5,;
       TOP=INT_015*80,;
       HEIGHT=INT_015*25,;
       WIDTH=INT_015*70,;
       FONTSIZE=INT_015*9,;
       VALUE=DATE()       
  ADD OBJECT TXT2 AS TEXTBOX WITH;
       LEFT=INT_015*120,;
       TOP=INT_015*80,;
       HEIGHT=INT_015*25,;
       WIDTH=INT_015*70,;
       FONTSIZE=INT_015*9,;
       VALUE=DATE()              
  *****    
  PROCEDURE CMND1.CLICK   
	DO CASE 
       CASE THISFORM.DATA_OPTION.VALUE=1 &&有預期結餘者
            	        FILE_NAME='E07 檢視MRP數據(有預期結餘)'
            	        SELECT E08.F02 STCK_NO,E08.F01 PRD,E08.F05 FINC,E08.F08 FPNC,E08.F11 FRNC,E08.F14 FONC,E08.F20+ALTT(E08.F02,'1') FSNC;
            	        FROM E08 WHERE E08.F05>0 AND E08.F20+ALTT(E08.F02,'1')>0 AND F01>=THISFORM.TXT1.VALUE AND F01<=THISFORM.TXT2.VALUE ORDER BY E08.F02,E08.F01 INTO CURSOR E08_A NOWAIT
	    CASE THISFORM.DATA_OPTION.VALUE=2	&&預期結餘為負
	    		 FILE_NAME='E07 檢視MRP數據(預期結餘為負)'
            	        SELECT E08.F02 STCK_NO,E08.F01 PRD,E08.F05 FINC,E08.F08 FPNC,E08.F11 FRNC,E08.F14 FONC,E08.F20+ALTT(E08.F02,'1') FSNC;
            	        FROM E08 WHERE  E08.F20+ALTT(E08.F02,'1')<0 AND F01>=THISFORM.TXT1.VALUE AND F01<=THISFORM.TXT2.VALUE ORDER BY E08.F02,E08.F01 INTO CURSOR E08_A NOWAIT
	    CASE THISFORM.DATA_OPTION.VALUE=3	&&全部
	    		 FILE_NAME='E07 檢視MRP數據(全部資料)'
            	        SELECT E08.F02 STCK_NO,E08.F01 PRD,E08.F05 FINC,E08.F08 FPNC,E08.F11 FRNC,E08.F14 FONC,E08.F20+ALTT(E08.F02,'1') FSNC;
            	        FROM E08 WHERE F01>=THISFORM.TXT1.VALUE AND F01<=THISFORM.TXT2.VALUE ORDER BY E08.F02,E08.F01 INTO CURSOR E08_A NOWAIT
	 ENDCASE 	
         IF messagebox('是否將進出明細一併轉入?',4+32+256,'請確認') = 6 	 
              FLG='1'
         ELSE
              FLG='0'
         ENDIF
	 GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	 IF EMPTY(GCDELIMFILE)  && ESC PRESSED
             RETURN
          ELSE
              ON ERROR DO FILE_IMPACT
              IF _TALLY >0
                   CREATE CURSOR E08_B;
                   (STCK_NO C(43),PRD D(8),FINC N(16,2),FPNC N(16,2),FRNC N(16,2),FONC N(16,2),FSNC N(16,2))              
                   CREATE CURSOR E08_C;
                   (STCK_NO C(43),F01 C(12),F03 D(8),F04 N(13,3),F05 C(5),F06 C(5),F08 N(13,3),F07 C(50),F99 C(1))                   
                   SELECT E08_A
                   GO TOP
                   DO WHILE !EOF()
                      IF FLG='1'
                              LK=E08_A.STCK_NO
                              LV=E08_A.PRD
                              DO DETAL
                              VHJ=E08_A.FSNC   
                              K=0                       
                              SELE INSCH
                              GO BOTT
                              DO WHILE !BOF()
                                     REPLACE  F08 WITH VHJ-K
                                     VHJ=VHJ-K
                                     K=F04
                                     SKIP -1
                              ENDDO  
                              GO TOP
                              DO WHILE !EOF()                                     
                                     SELECT E08_C
                                     APPEND BLANK           
                                     REPLACE STCK_NO WITH E08_A.STCK_NO 
                                   REPLACE F01 WITH INSCH.F01
                                   REPLACE F03 WITH INSCH.F03
                                   REPLACE F04 WITH INSCH.F04
                                   REPLACE F05 WITH INSCH.F05
                                   REPLACE F06 WITH INSCH.F06
                                   REPLACE F07 WITH INSCH.F07
                                   REPLACE F08 WITH INSCH.F08            
                                   REPLACE F99 WITH INSCH.F99        
                                   THISFORM.LBL2.CAPTION=F05+'  '+F06+'  '                                                                                                                                                    
                                   THISFORM.LBL1.CAPTION=F01+'  '+ALLTRIM(STR(F04))+'  '+F07                                                                                                                 
                                     SELECT INSCH
                                     SKIP
                                ENDDO      
                             SELECT INSCH
                             USE                                           
                      ELSE
                          SELECT E08_B
                          APPEND BLANK
                          REPLACE STCK_NO WITH E08_A.STCK_NO
                          REPLACE PRD WITH  E08_A.PRD
                          REPLACE FINC WITH E08_A.FINC
                          REPLACE FPNC WITH E08_A.FPNC
                          REPLACE FRNC WITH E08_A.FRNC
                          REPLACE FONC WITH E08_A.FONC
                          REPLACE FSNC WITH E08_A.FSNC
                          THISFORM.LBL1.CAPTION=STCK_NO+'  '+DTCX(PRD)+'  '+ALLTRIM(STR(FSNC))                          
                      ENDIF                                                                                             
                      SELECT E08_A
                      SKIP
                   ENDDO
                   SELECT E08_A
                   USE
                   IF FLG='1'
                        SELECT STCK_NO 料號,F01 單據號碼,F03 異動日期,F04 異動數量,F08 預期結餘,F05 客戶編號,IIF(SEEK(F05,'C01'),C01.F05,'') 客戶簡稱,IIF(SEEK(F06,'A01'),A01.F03,'') 業務人員,IIF(F99='1','是','否') 是否確認,F07 備註說明 FROM E08_C  NOWAIT ORDER BY STCK_NO,F03
                   ELSE
                        SELECT STCK_NO 料號,PRD 週期,FINC 當期進貨,FPNC  當期產出,FRNC 當期耗用,FONC 當期出貨,FSNC 當期結餘;
                        FROM E08_B   NOWAIT                                     
                   ENDIF     
                   COPY TO (GCDELIMFILE) TYPE XL5
		   =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗') 
	       ENDIF                             
          ENDIF	
          SELECT E08
     THISFORM.RELEASE                         
  ENDPROC
 ***** 
  PROCEDURE CMND2.CLICK
      THISFORM.RELEASE
  ENDPROC      
ENDDEFINE
**********************************************計算明細
PROCEDURE DETAL
                          SELE D07.F01,D07.F03,D07.F04,D07.F05,D07.F07,D07.F09,D07.F13 FROM D07 WHERE D07.F02=LK AND MRP_PRD(D07.F03)=LV  INTO CURSOR INSCH1 ORDER BY D07.F03 NOWAIT&&
                          SELE E06.F01,E06.F03,E06.F04,E06.F05,E06.F07,E06.F09 FROM E06 WHERE E06.F02=LK AND MRP_PRD(E06.F03)=LV  INTO CURSOR INSCH3  ORDER BY E06.F03 NOWAIT&&            
                          SELE C05.F01,IIF(EMPTY(C05.F10),C05.F03,C05.F10) F03,C05.F04,C05.F06,C05.F10,C05.F11 FROM C05 WHERE C05.F02=LK AND MRP_PRD(IIF(EMPTY(C05.F10),C05.F03,C05.F10))=LV INTO CURSOR INSCH2 ORDER BY C05.F03 NOWAIT&&
                          SELE E16.F01,IIF(EMPTY(E16.F10),E16.F03,E16.F10) F03,E16.F04,E16.F06,E16.F10 FROM E16 WHERE E16.F02=LK AND MRP_PRD(IIF(EMPTY(E16.F10),E16.F03,E16.F10))=LV INTO CURSOR INSCH4 ORDER BY E16.F03 NOWAIT&&         
                          CREATE CURSOR INSCH(F01 C(12),F03 D(8),F04 N(13,3),F08 N(13,3),F05 C(5),F06 C(5),F07 C(50),F99 C(1))
                          INDE ON F03 TAG INSCH
                          SET ORDER TO 1
                          SELE INSCH1
                          GO TOP
                          DO WHILE !EOF()
                                 SELE INSCH
                                 APPEND BLANK
                                 REPL F01 WITH INSCH1.F01
                                 REPL F03 WITH INSCH1.F03
                                 REPL F04 WITH INSCH1.F04
                                 REPLACE F05 WITH IIF(SEEK(LEFT(INSCH1.F07,10),'P05'),P05.F04,'')
                                 REPLACE F06 WITH IIF(SEEK(F05,'C01'),C01.F33,'')
                                 REPL F07 WITH '採:'+INSCH1.F13+IIF(SEEK(INSCH1.F05,'D01'),D01.F04,'')+IIF(!EMPTY(INSCH1.F07),'請:','')+INSCH1.F07
                                 REPLACE F99 WITH INSCH1.F09
                                 SELE INSCH1
                                 SKIP
                          ENDDO   
            USE
                          SELE INSCH3
                          GO TOP
                          DO WHILE !EOF()
                                 SELE INSCH
                                 APPEND BLANK
                                 REPL F01 WITH INSCH3.F01
                                 REPL F03 WITH INSCH3.F03
                                 REPL F04 WITH INSCH3.F04
                                 REPLACE F05 WITH IIF(SEEK('CA'+LEFT(F01,8),'C03'),C03.F03,'')
                                 REPLACE F06 WITH IIF(SEEK(F05,'C01'),C01.F33,'')                                 
                                 REPL F07 WITH '產:'+INSCH3.F05+IIF(SEEK(INSCH3.F05,'A14'),A14.F02,'')
                                 REPLACE F99 WITH INSCH3.F09                                 
                                 SELE INSCH3
                                 SKIP
                         ENDDO   
            USE
                         SELE INSCH2
                         GO TOP
                         DO WHILE !EOF()
                                SELE INSCH
                                APPEND BLANK
                                REPL F01 WITH INSCH2.F01
                                REPL F03 WITH INSCH2.F03
                                REPL F04 WITH INSCH2.F04*(-1)
                                 REPLACE F05 WITH IIF(SEEK(LEFT(F01,10),'C03'),C03.F03,'')
                                 REPLACE F06 WITH IIF(SEEK(F05,'C01'),C01.F33,'')                                             
                                REPL F07 WITH INSCH2.F11
                                REPLACE F99 WITH IIF(!EMPTY(INSCH2.F10),'1','')
                                SELE INSCH2
                                SKIP
                         ENDDO   
            USE
                         SELE INSCH4
                         GO TOP
                         DO WHILE !EOF()
                                SELE INSCH
                                APPEND BLANK
                                REPL F01 WITH INSCH4.F01
                                REPL F03 WITH INSCH4.F03
                                REPL F04 WITH INSCH4.F04*(-1)
                                 REPLACE F05 WITH IIF(SEEK('CA'+LEFT(F01,8),'C03'),C03.F03,'')
                                 REPLACE F06 WITH IIF(SEEK(F05,'C01'),C01.F33,'')                                             
                                REPL F07 WITH '領:'+INSCH4.F06+IIF(SEEK(INSCH4.F06,'A14'),A14.F02,'')   
                                REPLACE F99 WITH IIF(!EMPTY(INSCH4.F10),'1','')                                        
                                SELE INSCH4
                                SKIP
                         ENDDO                           
            USE
                       RETURN            
