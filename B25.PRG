**程式名稱:B25.部門別月報表
SET EXCL OFF
CLOSE ALL
CLEAR 
AREA1='B25'
FLG='0'
FCH=''
CHK=''
INDEX_KEY=''
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
SEEK sys_oper+'B25'

SEEK SYS_OPER+'B25'      
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
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET ORDER TO 1        
SET FILTER TO F04='*'
GO TOP
G11='G11'+LEFT(DTOS(DATE()),6)   &&月報表檔
IF !USED('&G11')
   SELE 0
   USE (G11) ALIA G11 INDE (G11)
ELSE
   SELE G11
ENDIF
SET ORDER TO 1  
B25='B25'+LEFT(DTOS(DATE()),6)   &&月報表檔
IF !USED('&B25')
   SELE 0
   USE (B25) ALIA B25 INDE (B25)
ELSE
   SELE B25
ENDIF
SET ORDER TO 1  
SET FILTER TO F01=A14.F01
GO TOP
SET RELATION  TO F01 INTO A14 
SET RELATION  TO F02 INTO B01 ADDI
B25FORM=CREATEOBJECT("TKB25")
B25FORM.SHOW  
DEFINE CLASS TKB25 AS FORM
  CAPTION='B25.部門別庫存月報表'
  CONTROLBOX=.F.
  BORDERSTYLE=3  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  CONTROLCOUNT=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WIDTH=INT_015*789
  WINDOWTYPE=1
  ADD OBJECT SHAPE1 AS SHAPE1 WITH;
      AUTOCENTER=.T.           
  ADD OBJECT SHAPE2 AS SHAPE2 WITH;
      AUTOCENTER=.T.
  ADD OBJECT SHAPE3 AS SHAPE3 WITH;
      AUTOCENTER=.T.
*!*	  ADD OBJECT LBL_REC AS LBL_REC WITH;
*!*	       TOP=INT_015*300
*!*	  ADD OBJECT REC_COUNT AS REC_COUNT  WITH;
*!*	      CAPTION=STR(RECCOUNT()),;
*!*	      TOP=INT_015*300  
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.,;
      WIDTH=INT_015*130
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      AUTOCENTER=.T.      
  ADD OBJECT GRID1 AS GRID1 WITH;
      COLUMNCOUNT=2,;            
      RECORDSOURCE='B25',;     
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.              
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.
*!*	  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
*!*	      TOP=INT_015*475,;
*!*	      LEFT=INT_015*410              
  ADD OBJECT ANS_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*460,;
      LEFT=INT_015*74,;      
      FONTSIZE=INT_015*9,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*65,;
      CAPTION='\<A.庫存分析'
*!*	  ADD OBJECT LOOK_BOTT AS COMMANDBUTTON WITH;
*!*	      TOP=INT_015*420,;             
*!*	      LEFT=INT_015*183,;
*!*	      WIDTH=INT_015*40,;
*!*	      HEIGHT=INT_015*25,;
*!*	      FONTSIZE=INT_015*9,;
*!*	      CAPTION='\<F.搜尋'      
   ADD OBJECT TXTF01 AS TEXTBOX WITH; 
   	  VISIBLE=.F.,;
      LEFT=INT_015*295,;
      TOP=INT_015*450,;
      WIDTH=INT_015*50,;
      NAME='TXTF01'         
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*27,;
      WIDTH=INT_015*50,;
      CAPTION='料品編號'  
  ADD OBJECT LBLF021 AS LABEL WITH;
      LEFT=INT_015*557,;
      TOP=INT_015*27,;
      AUTOSIZE=.T.,;
      CAPTION='' 
  ADD OBJECT TXTF02 AS TEXTBOX WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*23,;
      WIDTH=INT_015*261,;
      NAME='TXTF02'      
  ADD OBJECT LBLF03 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*52,;
      WIDTH=INT_015*50,;
      CAPTION='期初數量'
  ADD OBJECT TXTF03 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*48,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF03' 
  ADD OBJECT LBLF04 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*77,;
      WIDTH=INT_015*50,;
      CAPTION='進貨數量'
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*73,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF04' 
  ADD OBJECT LBLF05 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*77,;
      WIDTH=INT_015*50,;
      CAPTION='進貨退出'
  ADD OBJECT TXTF05 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*73,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF05'  
  ADD OBJECT LBLF06 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*102,;
      WIDTH=INT_015*50,;
      CAPTION='出貨數量'
  ADD OBJECT TXTF06 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*98,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF06' 
  ADD OBJECT LBLF07 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*102,;
      WIDTH=INT_015*50,;
      CAPTION='出貨退回'
  ADD OBJECT TXTF07 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*98,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF07' 
  ADD OBJECT LBLF08 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*127,;
      WIDTH=INT_015*50,;
      CAPTION='轉入數量'
  ADD OBJECT TXTF08 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*123,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF08' 
  ADD OBJECT LBLF09 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*127,;
      WIDTH=INT_015*50,;
      CAPTION='轉出數量'
  ADD OBJECT TXTF09 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*123,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF09'   
  ADD OBJECT LBLF10 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*152,;
      WIDTH=INT_015*50,;
      CAPTION='生產數量'
  ADD OBJECT TXTF10 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*148,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF10' 
  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*152,;
      WIDTH=INT_015*50,;
      CAPTION='產耗數量'
  ADD OBJECT TXTF11 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*148,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF11'     
  ADD OBJECT LBLF18 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*177,;
      WIDTH=INT_015*50,;
      CAPTION='借入數量'
  ADD OBJECT TXTF18 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*173,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF18' 
  ADD OBJECT LBLF19 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*177,;
      WIDTH=INT_015*50,;
      CAPTION='還出數量'
  ADD OBJECT TXTF19 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*173,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF19'   
  ADD OBJECT LBLF20 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*202,;
      WIDTH=INT_015*50,;
      CAPTION='借出數量'
  ADD OBJECT TXTF20 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*198,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF20' 
  ADD OBJECT LBLF21 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*202,;
      WIDTH=INT_015*50,;
      CAPTION='還入數量'
  ADD OBJECT TXTF21 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*198,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF21'  
  ADD OBJECT LBLF14 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*227,;
      WIDTH=INT_015*50,;
      CAPTION='損耗數量'
  ADD OBJECT TXTF14 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*223,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF14' 
  ADD OBJECT LBLF17 AS LABEL WITH;
      LEFT=INT_015*504,;
      TOP=INT_015*227,;
      WIDTH=INT_015*50,;
      CAPTION='他耗數量'
  ADD OBJECT TXTF17 AS TEXTBOX WITH;
      LEFT=INT_015*555,;
      TOP=INT_015*223,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF17' 
  ADD OBJECT LBLF13 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*252,;
      WIDTH=INT_015*50,;
      CAPTION='盤差數量'
  ADD OBJECT TXTF13 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*248,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF13'
  ADD OBJECT LBLF15 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*277,;
      WIDTH=INT_015*50,;
      CAPTION='期末數量'
  ADD OBJECT TXTF15 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*273,;
      WIDTH=INT_015*80,;
      MAXLENGTH=40,;
      NAME='TXTF15'  
**************                             
  PROCEDURE INIT  

       SEEK A14.F01
       CHK_STR=IIF(SEEK(DTOS(CTOD((THISFORM.MTH_LIST.DISPLAYVALUE+'/01'))),'A23','A231'),LEFT(DTOC(A23.F01),5),'')
*!*	       THISFORM.CMDGROUP.SEEK_BOTT.VISIBLE=.F.
*!*	       THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限   
       THISFORM.ANS_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.)           &&判斷有無列印成本的權限
*!*	       THISFORM.ANS_BOTT.VISIBLE=IIF(A02.F08='*',.T.,.F.)           &&判斷有無列印成本的權限  
       THISFORM.SETALL('HEIGHT',INT_015*25,'TEXTBOX')
       THISFORM.SETALL('HEIGHT',INT_015*17,'LABEL')
       THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
       THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')       
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(B25.F15<0,RGB(255,0,0),'')","COLUMN")       
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER') 
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*130
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B25.F02'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
       THISFORM.GRID1.READONLY=.T.
       THISFORM.GRID1.SETFOCUS
        IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.ANS_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限   
       	  THISFORM.ANS_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.)           &&判斷有無列印成本的權限
       ENDIF          
       THISFORM.KEY_LIST.INTERACTIVECHANGE
       JK='料品編號'    
        CHK=THISFORM.TXTF01.VALUE       
       THISFORM.REFRESH
  ENDPROC 
*************  
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS NCOLINDEX 
       SELE B25
       SET ORDER TO 1
       THISFORM.TXTF01.VALUE=B25.F01
       THISFORM.TXTF02.VALUE=B25.F02
       THISFORM.LBLF021.CAPTION=B01.F02
       THISFORM.TXTF03.VALUE=B25.F03
       THISFORM.TXTF04.VALUE=B25.F04
       THISFORM.TXTF05.VALUE=B25.F05
       THISFORM.TXTF06.VALUE=B25.F06
       THISFORM.TXTF07.VALUE=B25.F07
       THISFORM.TXTF08.VALUE=B25.F08
       THISFORM.TXTF09.VALUE=B25.F09
       THISFORM.TXTF10.VALUE=B25.F10
       THISFORM.TXTF11.VALUE=B25.F11
       THISFORM.TXTF13.VALUE=B25.F13
       THISFORM.TXTF14.VALUE=B25.F14
       THISFORM.TXTF15.VALUE=B25.F15
       THISFORM.TXTF17.VALUE=B25.F17
       THISFORM.TXTF18.VALUE=B25.F18
       THISFORM.TXTF19.VALUE=B25.F19
       THISFORM.TXTF20.VALUE=B25.F20
       THISFORM.TXTF21.VALUE=B25.F21
       THISFORM.GRID1.TOOLTIPTEXT=IIF(B25.F15<0,'期末數量小於零','')       
       FCH=THISFORM.ACTIVECONTROL.NAME
       JK=THIS.COLUMN1.HEADER1.CAPTION 
       IF THIS.ACTIVEROW=0
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.
       ENDIF  
       CHK=THISFORM.TXTF01.VALUE
*!*	       THISFORM.REFRESH
  ENDPROC 
***********            
  PROC KEY_LIST.INIT 
	     WITH THIS 
           SELE A14     
           GO TOP
           DO WHILE !EOF() 
              .ADDITEM(A14.F01+SPACE(2)+A14.F02)
              SKIP
           ENDDO
           .VALUE=1
       ENDWITH 
  ENDPROC
*************   
  PROC KEY_LIST.INTERACTIVECHANGE        
	INDEX_KEY=LEFT(THISFORM.KEY_LIST.DISPLAYVALUE,5) 
       SELE B25
	   SET FILTER TO F01=INDEX_KEY
       GO TOP
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC 
************    
  PROC MTH_LIST.INTERACTIVECHANGE
       CHK_STR=IIF(SEEK(DTOS(CTOD((THIS.DISPLAYVALUE+'/01'))),'A23','A231'),LEFT(DTOC(A23.F01),5),'')
       THISFORM.GRID1.RECORDSOURCE=''
       SELE B25
       SET RELATION OFF INTO B01
       SET RELATION OFF INTO A14
       USE
       SELECT G11
       USE       
       G11='G11'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
        IF !USED('&G11')
   			SELE 0
   			USE (G11) ALIA G11 INDE (G11)
		ELSE
   			SELE G11
		ENDIF
		SET ORDER TO 1         
       B25='B25'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
        IF !USED('&B25')
   			SELE 0
   			USE (B25) ALIA B25 INDE (B25)
		ELSE
   			SELE B25
		ENDIF
		SET ORDER TO 1  
       SET RELATION  TO F01 INTO A14 
       SET RELATION  TO F02 INTO B01 ADDI 
       THISFORM.GRID1.RECORDSOURCE='B25'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='B25.F02'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='B01.F02'
	   THISFORM.KEY_LIST.INTERACTIVECHANGE
	   THISFORM.GRID1.SETFOCUS        
	   THISFORM.REFRESH
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.ANS_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
          THISFORM.ANS_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.)           &&判斷有無列印成本的權限
       ENDIF       
      * FLG='0'
  ENDPROC  
*************             
  PROC ANS_BOTT.CLICK
       LK=LEFT(THISFORM.KEY_LIST.DISPLAYVALUE,5)
       DEP=ALLTRIM(RIGHT(THISFORM.KEY_LIST.DISPLAYVALUE,8))       
       DEPART=LK+SPACE(2)+DEP       
     IF messagebox('是否只算到結帳日之庫存',4+32+256,'請確認') = 6	     
         B26='B26'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)  &&庫存異動檔
         IF !USED('&B26')
              SELE 0
              USE (B26) ALIA B26 INDE (B26)
        ELSE
            SELE B26
        ENDIF
        SET ORDER TO 1                 
           ANS_CHOICE=CREATEOBJECT("ANS_CHOICE")  
           ANS_CHOICE.SHOW                  

     ELSE 
       SELE B25.F01,B25.F02,((B25.F05+B25.F06+B25.F09+B25.F11+B25.F14+B25.F17+B25.F19+B25.F20)/((B25.F03+B25.F15)/2)),B25.F15,G11.F21 FROM B25,G11 WHERE B25.F01=LK AND B25.F15<>0 AND B25.F02=G11.F02 ORDER BY B25.F02 NOWAIT INTO CURSOR B25I000
       IF _TALLY>0
           FILENAME=LEFT(CHK_STR,2)+'年'+RIGHT(CHK_STR,2)+'月'+DEP+'庫存分析月報表'+'(B25)'
           COPY TO  &FILENAME TYPE XL5 
     
           REPORT FORM ALLTRIM(INT_116)+'B25B' PREVIEW
       ELSE
            =MESSAGEBOX('本月之加權平均成本尚未計算!請先執行G08再列印本報表!',0+48,'提示訊息視窗')            
       ENDIF
     ENDIF         
       THISFORM.GRID1.SETFOCUS
  ENDPROC 
ENDDEFINE  
***********************************************列印程序
PROCEDURE PNT_PRC     
          LK=A14.F01
           SELECT * FROM B25 ORDER BY F02 WHERE F01=LK AND (F03<>0 OR F04<>0;
            OR F05<>0 OR F06<>0 OR F07<>0 OR F08<>0 OR F09<>0 OR F10<>0 OR F11<>0;
            OR F12<>0 OR F13<>0 OR F14<>0 OR F15<>0  OR F17<>0 OR F18<>0;
            OR F19<>0 OR F20<>0 OR F21<>0 )NOWAIT
           IF _TALLY >0
              DEPART=A14.F01+SPACE(2)+A14.F02
              REPORT FORM ALLTRIM(INT_116)+'B25' PREVIEW
*              REPORT FORM B25 TO PRINT PROMPT        
           ENDIF                                                             
ENDPROC   
****************  
********************************************************
DEFINE CLASS ANS_CHOICE AS FORM
  AUTOCENTER=.T.
  CAPTION='請輸入結帳日'
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*86
  WIDTH=INT_015*250
  FONTSIZE=INT_015*9
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='ANS_CHOICE'
  ADD OBJECT LABEL1 AS LABEL WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*23,;
      AUTOSIZE=.T.,;
      FONTSIZE=INT_015*9,;
      CAPTION='結帳日期'         
  ADD OBJECT LABEL11 AS LABEL WITH;
      LEFT=INT_015*122,;
      TOP=INT_015*23,;
      FONTSIZE=INT_015*9,;
      AUTOSIZE=.T.,;
      CAPTION=''    
  ADD OBJECT TEXT1 AS TEXTBOX WITH;
      LEFT=INT_015*80,;
      TOP=INT_015*20,;
      WIDTH=INT_015*25,;
      HEIGHT=INT_015*25,;
      FONTSIZE=INT_015*9,;
      MAXLENGTH=2,;       
      NAME='TEXT1'       
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;     
      TOP=INT_015*11,;  
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;      
      CAPTION='\<Y.確定',;
      TOOLTIPTEXT='確認所選擇的鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*51,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<C.取消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND2'      
  *****    
  PROCEDURE INIT
       THISFORM.TEXT1.VALUE=RIGHT(DTOS(DATE()),2)
  ENDPROC
  PROCEDURE CMND1.CLICK        
        CREATE CURSOR JIEZ (F01 C(5),F02 C(43),F03 N(14,4),F04 N(14,4),F05 N(14,4),F06 N(14,4),F07 N(14,4),;
        F08 N(14,4),F09 N(14,4),F10 N(14,4),F11 N(14,4),F13 N(14,4),F14 N(14,4),F17 N(14,4),F18 N(14,4),F19 N(14,4),;
        F20 N(14,4),F21 N(14,4),F15 N(14,4),F23 N(16,6),F24 N(16,6),F16 N(16,6),F12 N(16,6))
        INDEX ON F01+F02 TAG JIEZ
        SET ORDER TO 1
        SELECT B25
        SEEK LK
        DO WHILE F01=LK
                SELECT JIEZ
                APPEND BLANK
                REPLACE F01 WITH B25.F01
                REPLACE F02 WITH B25.F02                
                SELECT B26
                SEEK B25.F01+B25.F02
               IF FOUND()
                   DO WHILE F02+F01=B25.F01+B25.F02 AND F03<=THISFORM.TEXT1.VALUE
                          SELECT JIEZ
                          DO CASE
                                CASE B26.F06='A'  OR B26.F06='F' &&進貨
                                           REPLACE F04 WITH F04+B26.F04
                                CASE B26.F06='B'  OR B26.F06='J' &&進貨退出
                                           REPLACE F05 WITH F05+B26.F04*(-1)
                                CASE B26.F06='C'   &&出貨 
                                           REPLACE F06 WITH F06+B26.F04*(-1)                                
                                CASE B26.F06='D'   &&出貨退回
                                           REPLACE F07 WITH F07+B26.F04                                
                                CASE B26.F06='E' OR B26.F06='P'    &&移轉
                                          IF B26.F04>0
                                              REPLACE F08 WITH F08+B26.F04
                                          ELSE
                                               REPLACE F09 WITH F09+B26.F04*(-1)
                                          ENDIF    
                                CASE B26.F06='M'         &&產出或產耗
                                          IF B26.F04>0
                                              REPLACE F10 WITH F10+B26.F04
                                          ELSE
                                             REPLACE F11 WITH F11+B26.F04*(-1)
                                          ENDIF       
                                CASE B26.F06='T'         &&產出或產耗
                                          IF B26.F04>0
                                              REPLACE F10 WITH F10+B26.F04
                                          ELSE
                                             REPLACE F11 WITH F11+B26.F04*(-1)
                                          ENDIF                                                                                                                                                                                                                                    
                                CASE B26.F06='H'          &&盤差數量                     
                                          REPLACE F13 WITH F13+B26.F04 
                                CASE B26.F06='K'  OR B26.F06='Q'      &&損耗數量
                                           REPLACE F14 WITH F14+B26.F04*(-1)
                                CASE B26.F06='I' OR B26.F06='Z'          &&其他耗用
                                           REPLACE F17 WITH F17+B26.F04*(-1)                                
                                CASE B26.F06='U'                                    &&借入                             
                                           REPLACE F18 WITH F18+B26.F04
                                CASE B26.F06='V'                                   &&借入還出
                                           REPLACE F19 WITH F19+B26.F04*(-1)                                                                                                           
                                CASE B26.F06='W'                                  &&借出
                                           REPLACE F20 WITH F20+B26.F04*(-1)
                                CASE B26.F06='X'                                   &&借出還入
                                           REPLACE F21 WITH F21+B26.F04                                                                         
                          ENDCASE
                          SELECT B26
                          SKIP
                   ENDDO       
               ENDIF
                    SELECT JIEZ
                    REPLACE F15 WITH B25.F03+F04-F05-F06+F07+F08-F09+F10-F11+F13-F14-F17+F18-F19-F20+F21
                    REPLACE F23 WITH IIF(SEEK(B25.F02,'G11'),G11.F13,0) &&期初單價                          
                    REPLACE F24 WITH IIF(SEEK(B25.F02,'G11'),G11.F14,0) &&進貨單價                                              
                    REPLACE F16 WITH IIF(SEEK(B25.F02,'G11'),G11.F21,0) &&平均單價
                    REPLACE F12 WITH ROUND(F15*F16,2)                           &&期末小計           
               SELECT B25
               SKIP
        ENDDO
           SELECT B26
           USE            
           SELECT A23
           SEEK DTOS(CTOD(CHK_STR+'/01'))
           SKIP -1
           IF !BOF()
                B26='B26'+LEFT(DTOS(A23.F01),6)
                IF !USED('&B26')
                     SELE 0
                     USE (B26) ALIA B26 INDE (B26)
                ELSE
                    SELE B26
                 ENDIF
                SET ORDER TO 1
                SET FILTER TO F03>THISFORM.TEXT1.VALUE AND F02=LK
                GO TOP
                DO WHILE !EOF()
                          SELECT JIEZ
                          SEEK B26.F02+B26.F01
                          IF FOUND()
                               DO CASE
                                      CASE B26.F06='A'  OR B26.F06='F' &&進貨
                                                REPLACE F04 WITH F04+B26.F04
                                      CASE B26.F06='B'  OR B26.F06='J' &&進貨退出
                                                 REPLACE F05 WITH F05+B26.F04*(-1)
                                      CASE B26.F06='C'   &&出貨 
                                                 REPLACE F06 WITH F06+B26.F04*(-1)                                
                                      CASE B26.F06='D'   &&出貨退回
                                                 REPLACE F07 WITH F07+B26.F04                                
                                      CASE B26.F06='E' OR B26.F06='P'    &&移轉
                                                 IF B26.F04>0
                                                     REPLACE F08 WITH F08+B26.F04
                                                ELSE
                                                     REPLACE F09 WITH F09+B26.F04*(-1)
                                                ENDIF    
                                      CASE B26.F06='M'         &&產出或產耗
                                                 IF B26.F04>0
                                                      REPLACE F10 WITH F10+B26.F04
                                                 ELSE
                                                      REPLACE F11 WITH F11+B26.F04*(-1)
                                                 ENDIF       
                                      CASE B26.F06='T'         &&產出或產耗
                                                 IF B26.F04>0
                                                     REPLACE F10 WITH F10+B26.F04
                                                 ELSE
                                                     REPLACE F11 WITH F11+B26.F04*(-1)
                                                 ENDIF                                                                                                                                                                                                                                    
                                      CASE B26.F06='H'          &&盤差數量                     
                                                 REPLACE F13 WITH F13+B26.F04 
                                      CASE B26.F06='K'  OR B26.F06='Q'      &&損耗數量
                                                 REPLACE F14 WITH F14+B26.F04*(-1)
                                      CASE B26.F06='I' OR B26.F06='Z'          &&其他耗用
                                                 REPLACE F17 WITH F17+B26.F04*(-1)                                
                                      CASE B26.F06='U'                                    &&借入                             
                                                 REPLACE F18 WITH F18+B26.F04
                                      CASE B26.F06='V'                                   &&借入還出
                                                 REPLACE F19 WITH F19+B26.F04*(-1)                                                                                                           
                                      CASE B26.F06='W'                                  &&借出
                                                 REPLACE F20 WITH F20+B26.F04*(-1)
                                      CASE B26.F06='X'                                   &&借出還入
                                                 REPLACE F21 WITH F21+B26.F04                                                                         
                                ENDCASE                
                        ENDIF        
                       SELECT B26
                       SKIP
                ENDDO
           ENDIF      
           SELECT B26
           USE
           SELECT JIEZ
           REPLACE F03 WITH F15-F04+F05+F06-F07-F08+F09-F10+F11-F13+F14+F17-F18+F19+F20-F21 ALL
           DELETE  FROM JIEZ  WHERE  (F03=0 AND F04=0;
            AND F05=0 AND F06=0 AND F07=0 AND F08=0 AND F09=0 AND F10=0 AND F11=0;
            AND F13=0 AND F14=0 AND F15=0  AND F17=0 AND F18=0;
            AND F19=0 AND F20=0 AND F21=0 )                               
           FILENAME=LEFT(CHK_STR,2)+'年'+RIGHT(CHK_STR,2)+'月'+DEP+'結帳至'+THISFORM.TEXT1.VALUE+'日庫存表'           
           SELECT JIEZ
           REPORT FORM ALLTRIM(INT_116)+'B25A' PREVIEW       
            SELECT F01 庫存部門,F02 料品編號,F03 期初數量,F04 進貨數量,F05 進貨退出,F06 出貨數量,F07 出貨退回,F08 轉入數量,;
            F09 轉出數量,F10 生產數量,F11 生產耗用,F13 盤差數量,F14 損耗數量,F17 其他耗用,F18 借入數量,F19 借入還出,F20 借出數量,;
            F21 借出還入,F15 期末數量,F23 期初單價,F24 進貨單價,F16 平均單價,F12 期末小計;
            FROM JIEZ  ORDER BY F02 NOWAIT      
           COPY TO &FILENAME TYPE XL5
           SELECT JIEZ
           USE
         THISFORM.CMND2.CLICK
  ENDPROC
  PROCEDURE CMND2.CLICK 
         THISFORM.RELEASE
  ENDPROC    
ENDDEFINE  