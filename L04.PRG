**程式名稱:每月攤提名細表
close all
clear 
AREA1='L03'
FLG='0'
FCH=''
CHK=''
CHK_STR=''
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
SEEK sys_oper+'L04'   
IF !USED('A03')
   SELE 0
   USE A03
ELSE
   SELE A03
ENDIF
SET ORDER TO 1         
IF !USED('A23')
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1 
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET ORDER TO 1   
IF !USED('L00')
   SELE 0
   USE L00
ELSE 
   SELE L00
ENDIF
SET ORDER TO 1
IF !USED('L01')
   SELE 0
   USE L01
ELSE 
   SELE L01
ENDIF
SET ORDER TO 1            
IF !USED('L02')
   SELE 0
   USE L02
ELSE 
   SELE L02
ENDIF
SET ORDER TO 1 
IF !USED('S50')
   SELE 0
   USE S50
ELSE 
   SELE S50
ENDIF
SET ORDER TO 1               
L03='L03'+LEFT(DTOS(DATE()),6)   &&每月攤提明細檔
       L031='L031'+LEFT(DTOS(DATE()),6) 
       L032='L032'+LEFT(DTOS(DATE()),6) 
       L033='L033'+LEFT(DTOS(DATE()),6) 
       L034='L034'+LEFT(DTOS(DATE()),6) 
IF !USED('&L03')
   SELE 0
   USE (L03) ALIA L03 INDE (L03)
ELSE
   SELE L03
ENDIF
SET ORDER TO L031                
GO TOP
SET RELATION  TO F01 INTO L01
SET RELATION TO F04 INTO L00 ADDI
SET RELATION  TO F02 INTO A14 ADDI 
SET RELATION  TO F03 INTO S50 ADDI 
L04form=createobject("tkL04")
L04form.show  
define class tkL04 as form
  caption='L04.每月攤提名細表'
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
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      caption=str(reccount()),;
      TOP=INT_015*5
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.,;
      WIDTH=INT_015*120
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      autocenter=.t.      
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=13,;            
      RECORDSOURCE='L03',; 
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
      COLUMN9.NAME='COLUMN9',;
      COLUMN10.NAME='COLUMN10',;
      COLUMN11.NAME='COLUMN11',;            
      COLUMN12.NAME='COLUMN12',;
      COLUMN13.NAME='COLUMN13'        
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;      
      TOP=INT_015*0,;
      LEFT=INT_015*420              
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=INT_015*0,;
      LEFT=INT_015*220,;
      ENABLED=.T.,;
      VISIBLE=.T.                     
  ADD OBJECT RESCH AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*440,;
      WIDTH=INT_015*75,;
      HEIGHT=INT_015*25,;
      ENABLED=.T.,;
      VISIBLE=.T.,;
      FONTSIZE=INT_015*9,;
      NAME='RESCH',;
      CAPTION='\<A.重新計算'
  PROCEDURE INIT   
       	THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       	THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       	THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       	THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
       	THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.T.
       	THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       	THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
       	THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       	THISFORM.GRID1.COLUMN1.WIDTH=INT_015*120
       	THISFORM.GRID1.COLUMN2.WIDTH=INT_015*120          
       	THISFORM.GRID1.COLUMN3.WIDTH=INT_015*40
        THISFORM.GRID1.COLUMN4.WIDTH=INT_015*55    
        THISFORM.GRID1.COLUMN5.WIDTH=INT_015*40
        THISFORM.GRID1.COLUMN6.WIDTH=INT_015*65                                                             
        THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55  
       	THISFORM.GRID1.COLUMN8.WIDTH=INT_015*80 &&每月折舊
       	THISFORM.GRID1.COLUMN9.WIDTH=INT_015*80 &&累計折舊
       	THISFORM.GRID1.COLUMN10.WIDTH=INT_015*80 &&目前殘值
       	THISFORM.GRID1.COLUMN11.WIDTH=INT_015*80 &&預留殘值              
       	THISFORM.GRID1.COLUMN12.WIDTH=INT_015*80 &&取得金額
       	THISFORM.GRID1.COLUMN13.WIDTH=INT_015*80 &&改良金額       
       	THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='財產編號'
       	THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='規格說明'
	THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='部門編號'
	THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='部門名稱'	   	     	   	   
       	THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='人員編號'
	THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='人員姓名'	   	   
	THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='類別區分'
  	THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='每月折舊攤提'  	   
  	THISFORM.GRID1.COLUMN9.HEADER1.CAPTION='累計折舊攤提'
  	THISFORM.GRID1.COLUMN10.HEADER1.CAPTION='目前殘值'
  	THISFORM.GRID1.COLUMN11.HEADER1.CAPTION='預留殘值'  	     	   
        THISFORM.GRID1.COLUMN12.HEADER1.CAPTION='取得金額'  	     	     	   
  	THISFORM.GRID1.COLUMN13.HEADER1.CAPTION='改良金額'  	     	     	     	   
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L03.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='L01.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L03.F02'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='L03.F03'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S50.F02'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='L00.F04'       
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='L03.F07+L03.F11'
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='L03.F08'
       THISFORM.GRID1.COLUMN10.CONTROLSOURCE='L03.F09'       
       THISFORM.GRID1.COLUMN11.CONTROLSOURCE='L03.F10'              
       THISFORM.GRID1.COLUMN12.CONTROLSOURCE='L01.F06'       
       THISFORM.GRID1.COLUMN13.CONTROLSOURCE='L01.F07'        
       CHK_STR=THISFORM.MTH_LIST.DISPLAYVALUE       
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS
       JK='財產編號'
  ENDPROC     
  PROC KEY_LIST.INIT 
       WITH THIS 
           .ADDITEM('依財產編號排列')
           .ADDITEM('依部門編號排列')
           .ADDITEM('依保管人員排列')
           .ADDITEM('依類別區分排列')           
       ENDWITH      
  ENDPROC 
  PROCEDURE KEY_LIST.INTERACTIVECHANGE
            SELECT L03
            DO CASE
               CASE THIS.VALUE=1
                    SET ORDER TO L031
                    THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='財產編號'
                    THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='規格說明'
                    THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='部門編號'
                    THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='部門名稱'         
                    THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='人員編號'
                    THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='人員姓名'                               
                    THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='類別區分'                                                   
    	            THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L03.F01'
                    THISFORM.GRID1.COLUMN2.CONTROLSOURCE='L01.F02'     
    	            THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L03.F02'
                    THISFORM.GRID1.COLUMN4.CONTROLSOURCE='A14.F02'        
    	            THISFORM.GRID1.COLUMN5.CONTROLSOURCE='L03.F03'
                    THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S50.F02'                                
                    THISFORM.GRID1.COLUMN7.CONTROLSOURCE='L00.F04'                                                    
                    THISFORM.GRID1.COLUMN1.WIDTH=INT_015*120
                    THISFORM.GRID1.COLUMN2.WIDTH=INT_015*120          
                    THISFORM.GRID1.COLUMN3.WIDTH=INT_015*40
                    THISFORM.GRID1.COLUMN4.WIDTH=INT_015*55    
                    THISFORM.GRID1.COLUMN5.WIDTH=INT_015*40
                    THISFORM.GRID1.COLUMN6.WIDTH=INT_015*65                                                             
                    THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55                                                                                 
               CASE THIS.VALUE=2
                    SET ORDER TO L032
                    THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='部門編號'
                    THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='部門名稱'
                    THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='財產編號'
                    THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='規格說明'       
                    THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='人員編號'
                    THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='人員姓名'                                    
                    THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='類別區分'                                                        
    	            THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L03.F02'
                    THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A14.F02'   
    	            THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L03.F01'
                    THISFORM.GRID1.COLUMN4.CONTROLSOURCE='L01.F02'       
    	            THISFORM.GRID1.COLUMN5.CONTROLSOURCE='L03.F03'
                    THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S50.F02'                                   
                    THISFORM.GRID1.COLUMN7.CONTROLSOURCE='L00.F04'                                                       
                    THISFORM.GRID1.COLUMN1.WIDTH=INT_015*40
                    THISFORM.GRID1.COLUMN2.WIDTH=INT_015*55    
                    THISFORM.GRID1.COLUMN3.WIDTH=INT_015*120
                    THISFORM.GRID1.COLUMN4.WIDTH=INT_015*120     
                    THISFORM.GRID1.COLUMN5.WIDTH=INT_015*40
                    THISFORM.GRID1.COLUMN6.WIDTH=INT_015*65                                                            
                    THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55
               CASE THIS.VALUE=3
                    SET ORDER TO L033
                    THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='人員編號'
                    THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='人員姓名'
                    THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='財產編號'
                    THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='規格說明'     
                    THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='部門編號'
                    THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='部門名稱'  
                    THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='類別區分'                                   
    	            THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L03.F03'
                    THISFORM.GRID1.COLUMN2.CONTROLSOURCE='S50.F02'    
    	            THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L03.F01'
                    THISFORM.GRID1.COLUMN4.CONTROLSOURCE='L01.F02'    
    	            THISFORM.GRID1.COLUMN5.CONTROLSOURCE='L03.F02'
                    THISFORM.GRID1.COLUMN6.CONTROLSOURCE='A14.F02'                                     
                    THISFORM.GRID1.COLUMN7.CONTROLSOURCE='L00.F04'                                                         
                    THISFORM.GRID1.COLUMN1.WIDTH=INT_015*50
                    THISFORM.GRID1.COLUMN2.WIDTH=INT_015*80        
                    THISFORM.GRID1.COLUMN3.WIDTH=INT_015*120
                    THISFORM.GRID1.COLUMN4.WIDTH=INT_015*120    
                    THISFORM.GRID1.COLUMN5.WIDTH=INT_015*40
                    THISFORM.GRID1.COLUMN6.WIDTH=INT_015*55                                                          
                    THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55                                                                              
               CASE THIS.VALUE=4  
                    SET ORDER TO L034
                    THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='類別區分'
                    THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='財產編號'
                    THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='規格說明'                    
                    THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='部門編號'                                        
                    THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='部門名稱'                                                            
                    THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='人員編號'                                                                                
                    THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='人員姓名'                                                                                                                        
    	            THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L00.F04'
                    THISFORM.GRID1.COLUMN2.CONTROLSOURCE='L03.F01'  
                    THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L01.F02'                                                            
                    THISFORM.GRID1.COLUMN4.CONTROLSOURCE='L03.F02'                                                                                
                    THISFORM.GRID1.COLUMN5.CONTROLSOURCE='A14.F02'         
                    THISFORM.GRID1.COLUMN6.CONTROLSOURCE='L03.F03'        
                    THISFORM.GRID1.COLUMN7.CONTROLSOURCE='S50.F02'   
                    THISFORM.GRID1.COLUMN1.WIDTH=INT_015*55
                    THISFORM.GRID1.COLUMN2.WIDTH=INT_015*120                               
                    THISFORM.GRID1.COLUMN3.WIDTH=INT_015*120                                                   
                    THISFORM.GRID1.COLUMN4.WIDTH=INT_015*40                                                                       
                    THISFORM.GRID1.COLUMN5.WIDTH=INT_015*55   
                    THISFORM.GRID1.COLUMN6.WIDTH=INT_015*50                                                                       
                    THISFORM.GRID1.COLUMN7.WIDTH=INT_015*55                                                                                                               
            ENDCASE                    
           JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION            
           THISFORM.GRID1.REFRESH
           THISFORM.GRID1.SETFOCUS                 
                           
  ENDPROC  
  PROC MTH_LIST.INTERACTIVECHANGE
       THISFORM.GRID1.RECORDSOURCE=''
       SELE L03
       SET RELATION OFF INTO L01
       SET RELATION OFF INTO S50
       SET RELATION OFF INTO A14
       SET RELATION OFF INTO L00
       USE
       L03='L03'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       L031='L031'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       L032='L032'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       L033='L033'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       L034='L034'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&L03')
          SELE 0
          USE (L03) ALIA L03
       ELSE
          SELE L03
       ENDIF
       SET ORDER TO L031        
       SET RELATION  TO F01 INTO L01
       SET RELATION TO F04 INTO L00 ADDI
       SET RELATION  TO F02 INTO A14 ADDI 
       SET RELATION  TO F03 INTO S50 ADDI 
       THISFORM.GRID1.RECORDSOURCE='L03'
       GO TOP       
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L03.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='L01.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='L03.F02'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE='L03.F03'
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S50.F02'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE='L00.F04'       
       THISFORM.GRID1.COLUMN8.CONTROLSOURCE='L03.F07+L03.F11'
       THISFORM.GRID1.COLUMN9.CONTROLSOURCE='L03.F08'
       THISFORM.GRID1.COLUMN10.CONTROLSOURCE='L03.F09'       
       THISFORM.GRID1.COLUMN11.CONTROLSOURCE='L03.F10'              
       THISFORM.GRID1.COLUMN12.CONTROLSOURCE='L01.F06'       
       THISFORM.GRID1.COLUMN13.CONTROLSOURCE='L01.F07'  
       CHK_STR=THIS.DISPLAYVALUE            
       THISFORM.KEY_LIST.INTERACTIVECHANGE           
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
       ENDIF          
  ENDPROC      
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROC RESCH.CLICK
       WAIT WINDOW "重新計算攤提資料中,請稍後!" AT 0,150 NOWAIT NOCLEAR
       SELECT L03
       SET RELATION OFF INTO L01
       SET RELATION OFF INTO S50
       SET RELATION OFF INTO L00
       SET RELATION OFF INTO A14
       DTLT=YEAR(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01'))*12+VAL(RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2))
       DELETE ALL
       SELECT L01
       GO TOP 
       DO WHILE !EOF()
          IF L01.F09='Y' AND CTOD(L01.F13+'/01')<CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01') AND L01.F32<L01.F26+L01.F06+L01.F07
             SELECT L03
             APPEND BLANK
             REPLACE F01 WITH L01.F01  &&財產編號
             REPLACE F02 WITH L01.F04  &&部門編號
             REPLACE F03 WITH L01.F05
             REPLACE F04 WITH L01.F27
             REPLACE F05 WITH L01.F11
             REPLACE F06 WITH L01.F12 
             REPLACE F07 WITH L01.F17  
                A=0
                B=0
                C=0
                D=0
                X=0
                Y=''
                Z=''             
             SELECT L02
             SEEK L03.F01             
             IF FOUND()
                DO WHILE F01=L03.F01
                       IF CTOD(F07+'/01')<=CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')                                    
                           X=DTLT-YEAR(CTOD(F07+'/01'))*12-VAL(RIGHT(F07,2))
                           REPLACE F05 WITH F03-F04*X  &&目前改良殘值
                           REPLACE F09 WITH THISFORM.MTH_LIST.DISPLAYVALUE
                           A=A+F04  
                           B=B+F06
                           C=C+F04*X    
                           D=D+F05                  
                       ENDIF    
                       SELECT L02
                       SKIP
                ENDDO
             ENDIF 
             SELECT L03
             REPLACE F08 WITH L01.F17*(DTLT-YEAR(CTOD(L01.F13+'/01'))*12-VAL(RIGHT(L01.F13,2)))+C
             REPLACE F09 WITH L01.F06-L01.F17*(DTLT-YEAR(CTOD(L01.F13+'/01'))*12-VAL(RIGHT(L01.F13,2)))+D
             REPLACE F10 WITH L01.F22+B
             REPLACE F11 WITH A            
             SELECT L01
             REPLACE F19 WITH THISFORM.MTH_LIST.DISPLAYVALUE
             REPLACE F20 WITH DTLT-YEAR(CTOD(F13+'/01'))*12-VAL(RIGHT(F13,2))
             REPLACE F21 WITH F06-F17*F20
             REPLACE F29 WITH D
             REPLACE F25 WITH L03.F09
             REPLACE F32 WITH L03.F08
          ENDIF    
          SELECT L01          
          SKIP
       ENDDO
         ***************
         WAIT CLEAR         
         SELE L03                
         SET RELATION  TO F01 INTO L01
         SET RELATION TO F04 INTO L00 ADDI
         SET RELATION  TO F02 INTO A14 ADDI 
         SET RELATION  TO F03 INTO S50 ADDI 
         GO TOP         
         THISFORM.REFRESH
         THISFORM.GRID1.SETFOCUS
   ENDPROC
enddefine        
***********************************************列印程序
PROCEDURE PNT_PRC     
    SELECT L03
    DO CASE 
       CASE KEY()='F01'            
            SELECT L03.F01,;
                   L01.F02,;
                   L03.F04,;
                   L00.F04,;
                   L03.F02,;
                   A14.F02,;
                   L03.F03,;
                   S50.F02,;
                   L03.F07,;
                   L03.F11,;
                   L03.F08,;
                   L03.F09,;
                   L03.F10,;
                   L01.F06,;
                   L01.F07,;
                   L01.F13,;
                   L01.F14;
                   FROM L03,L01,L00,A14,S50 WHERE L01.F01=L03.F01 AND A14.F01=L03.F02 AND S50.F01=L03.F03;
                   AND L00.F01=L03.F04 ORDER BY L03.F01 NOWAIT 
            IF _tally>0
                REPORT FORM ALLTRIM(INT_116)+'L04A' TO PRINT PROMPT PREVIEW            
            ENDIF
       CASE KEY()='F02'
            SELECT L03.F01,;
                   L01.F02,;
                   L03.F04,;
                   L00.F04,;
                   L03.F02,;
                   A14.F02,;
                   L03.F03,;
                   S50.F02,;
                   L03.F07,;
                   L03.F11,;
                   L03.F08,;
                   L03.F09,;
                   L03.F10,;
                   L01.F06,;
                   L01.F07,;
                   L01.F13,;
                   L01.F14;
                   FROM L03,L01,L00,A14,S50 WHERE L01.F01=L03.F01 AND A14.F01=L03.F02 AND S50.F01=L03.F03;
                   AND L00.F01=L03.F04 ORDER BY L03.F02,L03.F01 NOWAIT 
            IF _tally>0
                REPORT FORM ALLTRIM(INT_116)+'L04B' TO PRINT PROMPT PREVIEW            
            ENDIF       
       CASE KEY()='F03'
            SELECT L03.F01,;
                   L01.F02,;
                   L03.F04,;
                   L00.F04,;
                   L03.F02,;
                   A14.F02,;
                   L03.F03,;
                   S50.F02,;
                   L03.F07,;
                   L03.F11,;
                   L03.F08,;
                   L03.F09,;
                   L03.F10,;
                   L01.F06,;
                   L01.F07,;
                   L01.F13,;
                   L01.F14;
                   FROM L03,L01,L00,A14,S50 WHERE L01.F01=L03.F01 AND A14.F01=L03.F02 AND S50.F01=L03.F03;
                   AND L00.F01=L03.F04 ORDER BY L03.F03,L03.F01 NOWAIT 
            IF _tally>0
                REPORT FORM ALLTRIM(INT_116)+'L04C' TO PRINT PROMPT PREVIEW            
            ENDIF              
       CASE KEY()='F04'
            SELECT L03.F01,;
                   L01.F02,;
                   L03.F04,;
                   L00.F04,;
                   L03.F02,;
                   A14.F02,;
                   L03.F03,;
                   S50.F02,;
                   L03.F07,;
                   L03.F11,;
                   L03.F08,;
                   L03.F09,;
                   L03.F10,;
                   L01.F06,;
                   L01.F07,;
                   L01.F13,;
                   L01.F14;
                   FROM L03,L01,L00,A14,S50 WHERE L01.F01=L03.F01 AND A14.F01=L03.F02 AND S50.F01=L03.F03;
                   AND L00.F01=L03.F04 ORDER BY L03.F04,L03.F01 NOWAIT 
            IF _tally>0
                REPORT FORM ALLTRIM(INT_116)+'L04D' TO PRINT PROMPT PREVIEW            
            ENDIF                  
    ENDCASE
                SELECT L03.F01 AS 財產編號,;
                   L01.F02 AS 規格說明,;
                   L03.F04 AS 類別代號,;
                   L00.F04 AS 類別區分,;
                   L03.F02 AS 保管部門編號,;
                   A14.F02 AS 保管部門簡稱,;
                   L03.F03 AS 保管人編號,;
                   S50.F02 AS 保管人名稱,;
                   L03.F07+L03.F11 AS 攤提金額,;
                   L03.F08 AS 累計攤提金額,;
                   L03.F09 AS 目前殘值,;
                   L03.F10 AS 預留殘值,;
                   L01.F06 AS 取得金額,;
                   L01.F07 AS 改良金額,;
                   L01.F13 AS 起始月份,;
                   L01.F14 AS 截止月份;
                   FROM L03,L01,L00,A14,S50 WHERE L01.F01=L03.F01 AND A14.F01=L03.F02 AND S50.F01=L03.F03;
                   AND L00.F01=L03.F04 ORDER BY L03.F01 NOWAIT
                   FILENAME="固定資產攤提明細表"+LEFT(DTOS(CTOD(CHK_STR+'/01')),6)
            IF _tally>0
                COPY TO &FILENAME TYPE XL5
            ENDIF                      
ENDPROC

