**程式名稱:S21.刷卡資料紀錄檔
SET EXCL OFF
CLOSE ALL
CLEAR 
PRG_NO='S21'
AREA1='S21'
FLG='0'
FCH=''
CHK=''
HK=''
IF !USED('A01')	&&操作人員資料設定
   SELECT  0
   USE A01
ELSE 
   SELECT  A01
ENDIF
SET ORDER TO A01    
IF !USED('A02')	&&權限設定檔
   SELECT  0
   USE A02 INDE A02
ELSE
   SELECT  A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'S21'   
IF !USED('A23')	&&月份檔
    SELE 0
    USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO A231   
IF !USED('S15')	&&年度假日設定檔
   SELE 0
   USE S15 
ELSE
   SELE S15
ENDIF
SET ORDER TO S151
S21='S21'+LEFT(DTOS(DATE()),6)
S211='S211'+LEFT(DTOS(DATE()),6)
S212='S212'+LEFT(DTOS(DATE()),6)
IF FILE('&S21'+'.DBF')=.T.
    IF !USED('&S21')	&&刷卡資料
         SELE 0
         USE (S21) ALIA S21
    ELSE
         SELE S21
    ENDIF
ELSE
    CREATE CURSOR S21 (F01 C(4),F02 C(8),F03 C(6),F04 D(8),F05 C(1),F06 C(5))
    INDE ON F01 TAG  S211 ASCENDING  
    INDE ON F02 TAG  S212 ASCENDING  
ENDIF 
SELECT F01,F02,F03 FROM S21 WHERE !EMPTY(F01) GROUP BY F01 ORDER BY F01 INTO CURSOR S21TMP
INDE ON F03 TAG S21TMP1 ASCENDING 
SET ORDER TO S21TMP1 
SELECT S21
GO TOP
DO WHILE !EOF()
   IF EMPTY(S21.F01) = .T.
	  SELECT S21TMP
	  SEEK S21.F03
	  IF FOUND()
	     REPLACE S21.F01 WITH S21TMP.F01
	     REPLACE S21.F02 WITH S21TMP.F02
	  ENDIF
   ENDIF
   SELECT S21
   SKIP
ENDDO
SELECT S21TMP
USE
SELECT S21
IF A02.F10<>'*'
    SET FILTER TO ALLTRIM(S21.F01)=IIF(SEEK(sys_oper,'A01'),ALLTRIM(A01.F12),.F.)
ENDIF
SELECT S21
SET ORDER TO S211   
GO TOP
*********
S21FORM=CREATEOBJECT("TKS21")
S21FORM.SHOW  
DEFINE CLASS TKS21 AS FORM
  CAPTION='S21.刷卡資料紀錄檔'
*!*	  AUTOCENTER=.T.
  CONTROLBOX=.F.
  BORDERSTYLE=1  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  CONTROLCOUNT=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WIDTH=INT_015*791
  WINDOWTYPE=1
 ADD OBJECT LBL_REC AS LBL_REC WITH;
     TOP=INT_015*8,;
      LEFT=INT_015*630,;
      AUTOSIZE=.T.,;
      BACKCOLOR=RGB(255,255,0),;
      CAPTION='資料筆數：'  
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      WIDTH=INT_015*125,;
      AUTOCENTER=.T.
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
     WIDTH=INT_015*65,;
      AUTOCENTER=.T.             
  ADD OBJECT GRID1 AS GRID1 WITH;
      COLUMNCOUNT=7,;            
      RECORDSOURCE='S21',; 
      HEIGHT=INT_015*500,;     
      WIDTH=INT_015*780,;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',; 
      COLUMN4.NAME='COLUMN4',; 
      COLUMN5.NAME='COLUMN5',;             
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7'
   ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      TOP=INT_015*0,;
      LEFT=INT_015*205,;
      ENABLED=.T.,;
      VISIBLE=.T.            
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;      
      TOP=INT_015*0,;
      LEFT=INT_015*405        
  ADD OBJECT TR_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*446,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*88,;      
      FONTSIZE=INT_015*9,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<T.資料轉至S16',;
      NAME='TR_BOTT'                                                  
  PROCEDURE INIT
       HK=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)
       THISFORM.LBL_REC.VISIBLE=IIF(A02.F10='*',.T.,.F.)
       THISFORM.LBL_REC.CAPTION='資料筆數：'+ALLTRIM(TRANSFORM(RECCOUNT('S21'),'@R 999,999,999,999'))       
       THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限  
       THISFORM.TR_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.) AND IIF(A02.F10='*',.T.,.F.)
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)   
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(DOW(S21.F04)=7 OR DOW(S21.F04)=1 OR INDEXSEEK(DTOS(S21.F04),.F.,'S15','S151')=.T.,RGB(0,255,255),'')","COLUMN")            
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')   
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*55
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*55
       THISFORM.GRID1.COLUMN3.WIDTH=INT_015*55
       THISFORM.GRID1.COLUMN4.WIDTH=INT_015*70
       THISFORM.GRID1.COLUMN5.WIDTH=INT_015*55
       THISFORM.GRID1.COLUMN6.WIDTH=INT_015*55
       THISFORM.GRID1.COLUMN7.WIDTH=INT_015*120
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='人員編號'
       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='人員姓名'
       THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='卡片號碼'	   
       THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='進出日期'
       THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='進出時段'		
       THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='進出時間'	
       THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='假日判斷'   
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='S21.F01'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='S21.F02'
       THISFORM.GRID1.COLUMN3.CONTROLSOURCE='S21.F03'
       THISFORM.GRID1.COLUMN4.CONTROLSOURCE='S21.F04'
       THISFORM.GRID1.COLUMN5.CONTROLSOURCE="IIF(S21.F05='1','上午','下午')"
       THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S21.F06'
       THISFORM.GRID1.COLUMN7.CONTROLSOURCE="DAY_WEEK(DOW(S21.F04))+IIF(INDEXSEEK(DTOS(S21.F04),.T.,'S15','S151'),SPACE(5)+ALLTRIM(S15.F02),'')"
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS    
       IF THISFORM.GRID1.ACTIVEROW=0 
          THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
	      THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
	      THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
	      THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.    
	      THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.	        
	      THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.   
	      THISFORM.TR_BOTT.ENABLED=.F.
       ELSE
          THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
	      THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
	      THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
	      THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.    
	      THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.	        
	      THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限  
	      THISFORM.TR_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.) AND IIF(A02.F10='*',.T.,.F.)  
       ENDIF       
       JK='人員編號'
  ENDPROC    
  PROC KEY_LIST.INIT 
       WITH THIS 
           .ADDITEM('依人員編號排列')
           .ADDITEM('依人員姓名排列')      
       ENDWITH      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE S21
       DO CASE
          CASE THIS.DISPLAYVALUE='依人員編號排列'          
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='人員編號'
		       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='人員姓名'
		       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='S21.F01'
		       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='S21.F02'	        
               SET ORDER TO S211
          CASE THIS.DISPLAYVALUE='依人員姓名排列'          
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='人員姓名'
		       THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='人員編號'		     
		       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='S21.F02'
		       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='S21.F01'	     	        
               SET ORDER TO S212   
       ENDCASE 
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC             
  ***********
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
               ENDIF       
               SKIP
            ENDDO   
            .VALUE=S
       ENDWITH   
  ENDPROC       
***********
  PROC MTH_LIST.INTERACTIVECHANGE
    HK=ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)
    THISFORM.GRID1.RECORDSOURCE=''         
    SELECT S21
    USE
	S21='S21'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
	S211='S211'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
	S212='S212'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
	IF FILE('&S21'+'.DBF')=.T.
	    IF !USED('&S21')	&&刷卡資料
	         SELE 0
	         USE (S21) ALIA S21
	    ELSE
	         SELE S21
	    ENDIF
	ELSE
	    CREATE CURSOR S21 (F01 C(4),F02 C(8),F03 C(6),F04 D(8),F05 C(1),F06 C(5))
	    INDE ON F01 TAG  S211 ASCENDING  
	    INDE ON F02 TAG  S212 ASCENDING  
	ENDIF
	SELECT F01,F02,F03 FROM S21 WHERE !EMPTY(F01) GROUP BY F01 ORDER BY F01 INTO CURSOR S21TMP
	INDE ON F03 TAG S21TMP1 ASCENDING 
	SET ORDER TO S21TMP1 
	SELECT S21
	GO TOP
	DO WHILE !EOF()
	   IF EMPTY(S21.F01) = .T.
		  SELECT S21TMP
		  SEEK S21.F03
		  IF FOUND()
		     REPLACE S21.F01 WITH S21TMP.F01
		     REPLACE S21.F02 WITH S21TMP.F02
		  ENDIF
	   ENDIF
	   SELECT S21
	   SKIP
	ENDDO
	SELECT S21TMP
	USE
	SELECT S21	
	IF A02.F10<>'*'
	    SET FILTER TO ALLTRIM(S21.F01)=IIF(SEEK(sys_oper,'A01'),ALLTRIM(A01.F12),.F.)
	ENDIF	   
	SELECT S21
	SET ORDER TO S211 	
	GO TOP        
	THISFORM.LBL_REC.VISIBLE=IIF(A02.F10='*',.T.,.F.)
	THISFORM.LBL_REC.CAPTION='資料筆數：'+ALLTRIM(TRANSFORM(RECCOUNT('S21'),'@R 999,999,999,999'))   
	THISFORM.GRID1.RECORDSOURCE='S21'
        THISFORM.GRID1.COLUMN1.CONTROLSOURCE='S21.F01'
        THISFORM.GRID1.COLUMN2.CONTROLSOURCE='S21.F02'
        THISFORM.GRID1.COLUMN3.CONTROLSOURCE='S21.F03'
        THISFORM.GRID1.COLUMN4.CONTROLSOURCE='S21.F04'
        THISFORM.GRID1.COLUMN5.CONTROLSOURCE="IIF(S21.F05='1','上午','下午')"
        THISFORM.GRID1.COLUMN6.CONTROLSOURCE='S21.F06'
        THISFORM.GRID1.COLUMN7.CONTROLSOURCE="DAY_WEEK(DOW(S21.F04))+IIF(INDEXSEEK(DTOS(S21.F04),.T.,'S15','S151'),SPACE(5)+ALLTRIM(S15.F02),'')"        
        THISFORM.GRID1.SETFOCUS
         IF  THISFORM.GRID1.ACTIVEROW=0 
              THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
	      THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
	      THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
	      THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.    
	      THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.	        
	      THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.   
	      THISFORM.TR_BOTT.ENABLED=.F.
         ELSE
              THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
	      THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
	      THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
	      THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.    
	      THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.T.	        
	      THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限
	      THISFORM.TR_BOTT.ENABLED=IIF(A02.F08='*',.T.,.F.) AND IIF(A02.F10='*',.T.,.F.)    
        ENDIF
*!*	        THISFORM.REFRESH
 ENDPROC      
*******        
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS NCOLINDEX
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       THISFORM.REFRESH
  ENDPROC       
  ***資料轉至S16
 PROCEDURE  TR_BOTT.CLICK  
     IF MESSAGEBOX('是否確定要將' + HK + '之資料轉入S16.上下班紀錄檔中,且轉入後將覆蓋已存在之資料!',4+32+256,'請選擇') = 6
	    WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR	
	    IF !USED('S01')	&&薪資資料設定
	       SELECT 0
	       USE S01 
	    ELSE
	       SELECT S01
	    ENDIF
	    SET ORDER TO S011	    
	    IF !USED('S00')	&&薪資參數設定
	       SELECT 0
	       USE S00 
	    ELSE
	       SELECT S00
	    ENDIF
	    SET ORDER TO S005
		S00_F63 = ALLTRIM(IIF(SEEK('Y','S00'),S00.F63,'F'))			&&加班時間不滿時則以30分鐘計算
		S00_F64 = ALLTRIM(IIF(SEEK('Y','S00'),S00.F64,'F'))			&&加班時間若滿時則以60分鐘計算
	    S00_F65 = IIF(INDEXSEEK('Y',.T.,'S00','S005'),S00.F65,'2')	&&加班資料是否依S19加班申請(修改)作業為主(1:是 2:否)
        IF !USED('S05')  &&人事薪資基本檔
            SELECT 0
            USE S05
        ELSE
            SELECT S05
        ENDIF
        SET ORDER TO S051
		IF !USED('S06')   &&薪資項目及扣款項目
		   SELE 0
		   USE S06 
		ELSE
		   SELE S06
		ENDIF
		SET ORDER TO S061        
	    IF !USED('S17')	&&請假作業檔
	       SELECT 0
	       USE S17
	    ELSE 
	       SELECT S17
	    ENDIF
	    SET ORDER TO S171
	    IF !USED('S18')	&&請假作業檔
	       SELECT 0
	       USE S18
	    ELSE 
	       SELECT S18
	    ENDIF
	    SET ORDER TO S181
	    IF !USED('S19')	&&請假作業檔
	       SELECT 0
	       USE S19
	    ELSE 
	       SELECT S19
	    ENDIF
	    SET ORDER TO S191
        S16 = 'S16' + LEFT(DTOS(CTOD(ALLTRIM(HK) + '/01')),6)
        S161 = 'S161' + LEFT(DTOS(CTOD(ALLTRIM(HK) + '/01')),6)
        S162 = 'S162' + LEFT(DTOS(CTOD(ALLTRIM(HK) + '/01')),6)
        IF !USED('&S16')	&&上下班紀錄檔
           SELECT 0
           USE (S16) ALIA S16
        ELSE
           SELECT S16
        ENDIF
        SELECT S16
        SET ORDER TO S161
        *****先計算健保費有補助65歲以上之眷口(主要判斷是否已滿65歲)
        SELECT S05
        GO TOP
        DO WHILE !EOF()
		   ***眷口人數大於不扣款之眷口數則為0,否則為 ROUND(投保金額*(全民健康保險費率)*(被保險人及眷屬月負擔金額比率),0)*(眷口人數+本人)
		   PEOP = IIF(S05.F16 >= IIF(INDEXSEEK('B',.T.,'S01','S011'),S01.F06,0),S01.F06 - 1,S05.F16)
		   PEOP1 = PEOP
		   SINPO = ROUND(S05.F23 * (IIF(INDEXSEEK('B',.T.,'S01','S011'),S01.F02,0) / 100) * (IIF(INDEXSEEK('B',.T.,'S01','S011'),S01.F03,0) / 100),0)
*!*			   SINPO = SINPO1 * (PEOP + 1)
		   IF PEOP > 0
			  SELECT S01.F12,S01.F13,S01.F14,S01.F15,S01.F16,S01.F18,0000 AS PE FROM S01 WHERE S01.F01 ='Z' AND S01.F12 = S05.F01 INTO CURSOR S01_A READWRITE 
			  SELECT S01_A
			  INDEX ON S01_A.F18 TAG S01_A1
			  INDEX ON DTOS(S01_A.F16) + S01_A.F14 TAG S01_A2
			  SET ORDER TO S01_A1
			  GO TOP
			  ***先判斷是否為殘障
			  DO WHILE !EOF() AND PEOP1 <> 0
			     DO CASE
			        CASE S01_A.F18 = '0'	&&重度殘障 
			          	 REPLACE S01_A.PE WITH SINPO - ROUND(SINPO * (S01.F09 / 100),0)
			        CASE S01_A.F18 = '1'	&&中度殘障 
			          	 REPLACE S01_A.PE WITH SINPO - CEILING(ROUND(SINPO * (S01.F10 / 100),0))
			        CASE S01_A.F18 = '2'	&&輕度殘障 
			             REPLACE S01_A.PE WITH SINPO - CEILING(ROUND(SINPO * (S01.F11 / 100),0))
			        OTHERWISE
			             REPLACE S01_A.PE WITH SINPO			             
			     ENDCASE
			     PEOP1 = PEOP1 - 1
			     SELECT S01_A
			     SKIP
			  ENDDO
			  **在判斷設籍地是否有補助65歲(含)以上老人,且設籍須滿一年(含)INT((CTOD(HK + '/' + ALLTRIM(STR(DAY_NUM(RIGHT(ALLTRIM(HK),2))))) - S01_A.F16) / 365) >= 65 &&非殘障+設籍一年+有補助+大於等於65歲以上
			  DATE1 = CTOD(HK + '/' + ALLTRIM(STR(DAY_NUM(RIGHT(ALLTRIM(HK),2)))))
			  PEOP1 = PEOP
			  SELECT S01_A
			  SET ORDER TO S01_A2
			  GO TOP
			  DO WHILE !EOF() AND PEOP1 <> 0
			     IF S01_A.F15 = 0 AND IIF(SEEK('Y' + S01_A.F14,'S01'),S01.F15,0) = 1 AND (YEAR(DATE1) - YEAR(S01_A.F16)) >= 65 AND MONTH(DATE1) >= MONTH(S01_A.F16) AND DAY(DATE1) >= DAY(S01_A.F16)
			        REPLACE S01_A.PE WITH IIF(IIF(SEEK('B','S01'),S01.F08,0) >= SINPO,0,SINPO - S01.F08)
			     ENDIF
			     PEOP1 = PEOP1 - 1
			     SELECT S01_A
			     SKIP
			  ENDDO
			  ***相加每一眷屬可扣繳之保費
			  PEOP1 = PEOP
			  SELECT S01_A
			  GO TOP
			  DO WHILE !EOF() AND PEOP1 <> 0
			     SINPO = SINPO + S01_A.PE
			     PEOP1 = PEOP1 - 1
			     SELECT S01_A
			     SKIP
			  ENDDO
			  SELECT S01_A
			  USE
		   ENDIF
		   SELECT S06
		   SET ORDER TO S061	&& F01+F06+F07
		   SEEK S05.F01+'K'+'健保費'  &&健保費
		   IF FOUND()
		      S06_F08 = S06.F08
		      IF SINPO <> S06.F08
			     REPLACE F08 WITH SINPO
			     REPLACE F99 WITH DTOC(DATE()) +' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')	
		   	     IF EMPTY(S06.F08)
			        DELETE 
			     ENDIF   
			  ENDIF	 
		   ELSE
		      S06_F08 = 0
		      SELECT S06
		      APPEND BLANK 
		      REPLACE F01 WITH S05.F01
		      REPLACE F02 WITH S05.F02
		      REPLACE F03 WITH S05.F05
		      REPLACE F04 WITH S05.F06
		      REPLACE F05 WITH '2'
		      REPLACE F06 WITH 'K'
		      REPLACE F07 WITH '健保費'
		      REPLACE F08 WITH INT(SINPO)
		      REPLACE F99 WITH DTOC(DATE()) +' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')		
		       IF EMPTY(S06.F08)
		          DELETE 
		       ENDIF
		   ENDIF
		   **若金額不相同則加入S07A.調薪修改作業中
		   IF SINPO <> 0 AND SINPO <> S06_F08
		      IF !USED('S07')	 &&S07 調薪修改作業
		   	     SELECT  0
		   	  	 USE S07 
		      ELSE
		         SELECT S07
			  ENDIF
			  SELECT S07
			  APPEND BLANK 
			  REPLACE S07.F01 WITH S06.F01
			  REPLACE S07.F02 WITH S06.F02
			  REPLACE S07.F03 WITH S06.F03
			  REPLACE S07.F04 WITH S06.F04
			  REPLACE S07.F05 WITH DATE()
			  REPLACE S07.F06 WITH '2'
			  REPLACE S07.F07 WITH '2'
			  REPLACE S07.F08 WITH '健保費調整'				 
			  REPLACE S07.F09 WITH '健保費'
			  REPLACE S07.F11 WITH S06.F08 - SINPO
			  REPLACE S07.F12 WITH S06_F08
			  REPLACE S07.F13 WITH S06.F08
			  REPLACE S07.F14 WITH 'Y'
			  REPLACE S07.F15 WITH '1'
			  REPLACE S07.F99 WITH DTOC(DATE()) +' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
			  SELECT S06
			  REPLACE S06.F09 WITH S07.F05
			  REPLACE S06.F10 WITH S07.F08
			  SELECT S07
			  USE
		   ENDIF		   
		   SELECT S05
		   SKIP	   
		ENDDO  
		SELECT S01
		USE   
		*****先整理S21上下班資料
		SELECT S21.F01,S21.F02,S05.F05,S05.F06,S05.F29,S05.F30 FROM S21,S05 WHERE S21.F01 = S05.F01 AND !EMPTY(S21.F01) ORDER BY S21.F01 GROUP BY S21.F01 INTO CURSOR S21A	&&整理刷卡資料之人員清單
		SELECT S21A
		GO TOP
		DO WHILE !EOF()
		   S00_F52 = IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F52,'')    &&休息時段起
		   S00_F53 = IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F53,'')    &&休息時段止
		   S00_F66 = IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F66,'')	   &&上班刷卡時段止
		   S00_F67 = IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F67,'')    &&下班刷卡時段起
		   IF IIF(INDEXSEEK(S21A.F01,.T.,'S16','&S161'),S16.F19 <> 'Y',.T.) = .T.	&&S16.上下班紀錄尚未確認
		      IF S21A.F30 = '1'		&&判斷該人員是否需要刷卡
		      	 SELECT S16
		      	 DELETE FROM S16 WHERE S16.F02 = S21A.F01 AND EMPTY(S16.F31)	&&刪除是否曠職或離職(1:曠職,2:離職,'':無)
		         FOR I = 1 TO DAY_NUM(HK)
			         ***先帶入該人員之基本資料
					 SELECT S16
					 SET ORDER TO S161
					 SEEK S21A.F01 + PADL(ALLTRIM(STR(I)),2,'0')
					 IF !FOUND()
						SELECT S16
						APPEND BLANK
						REPLACE S16.F01 WITH PADL(ALLTRIM(STR(I)),2,'0')
					 ENDIF
			         REPLACE S16.F02 WITH S21A.F01
			         REPLACE S16.F03 WITH S21A.F02
			         REPLACE S16.F04 WITH S21A.F05
			         REPLACE S16.F05 WITH S21A.F06
		             REPLACE S16.F06 WITH IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F50,'')
		             REPLACE S16.F07 WITH IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F51,'')
		             REPLACE S16.F08 WITH IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F54,'')
		             REPLACE S16.F09 WITH IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F55,'')
		             REPLACE S16.F10 WITH IIF(INDEXSEEK('I' + S21A.F29,.T.,'S00','S005'),S00.F56,'')
			         REPLACE S16.F18 WITH DOW(CTOD(ALLTRIM(HK + '/' + S16.F01)))
			         REPLACE S16.F19 WITH 'N'
			         REPLACE S16.F32 WITH '1'
			         REPLACE S16.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
					 ***再帶入該人員之刷卡資料
			         SELECT S21.* FROM S21 WHERE S21.F01 = S21A.F01 AND DTOS(S21.F04) = DTOS(CTOD(HK + '/' + PADL(ALLTRIM(STR(I)),2,'0'))) ORDER BY S21.F01,S21.F05,S21.F06 INTO CURSOR S21B
			      	 SELECT S21B
			      	 GO TOP
			      	 IF _TALLY > 0
			         	SELECT S16
				        REPLACE S16.F11 WITH S21B.F06		&&第一筆資料(上班)
				        SELECT S21B
				        GO BOTTOM
				        SELECT S16
				        IF ALLTRIM(S21B.F06) <> ALLTRIM(S16.F11)
				       	   REPLACE S16.F12 WITH S21B.F06	&&最後一筆資料(下班)
				        ENDIF
				        ***判斷第一筆資料是否在 班別:下班時段 之後
				        IF !EMPTY(S16.F11) AND S16.F11 > S16.F07
				            REPLACE S16.F12 WITH S16.F11
				        	REPLACE S16.F11 WITH ''				        	
				        ENDIF
				        ***判斷最後一筆資料是否在 班別:上班時段 之前
				        IF !EMPTY(S16.F12) AND S16.F12 < S16.F06
				            REPLACE S16.F11 WITH S16.F12
				        	REPLACE S16.F12 WITH ''
				        ENDIF
			      	 ENDIF
			      	 SELECT S21B
			      	 USE
		     	 ENDFOR
		      ENDIF
		   ENDIF   
		   SELECT S21A
		   SKIP
		ENDDO
		SELECT S21A
		USE
		*****加入需刷卡但卻無刷卡資料之人員
		SELECT S05.F01,S05.F02,S05.F05,S05.F06,S05.F29 FROM S05 WHERE S05.F30 = '1' AND F28 <> '1' AND F28 <> '3' AND F28 <> '7' AND F28 <> '8' ORDER BY S05.F01 INTO CURSOR S05_TEMP
		SELECT S05_TEMP
		GO TOP
		DO WHILE !EOF()
		   SELECT S16
		   SEEK S05_TEMP.F01
		   IF !FOUND()
		   	   FOR I=1 TO DAY_NUM(HK)
				   SELECT S16
				   APPEND BLANK
			       REPLACE S16.F01 WITH PADL(ALLTRIM(STR(I)),2,'0')
				   REPLACE S16.F02 WITH S05_TEMP.F01
				   REPLACE S16.F03 WITH S05_TEMP.F02
				   REPLACE S16.F04 WITH S05_TEMP.F05
				   REPLACE S16.F05 WITH S05_TEMP.F06
				   REPLACE S16.F06 WITH IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F50,'')
				   REPLACE S16.F07 WITH IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F51,'')
				   REPLACE S16.F08 WITH IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F54,'')
				   REPLACE S16.F09 WITH IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F55,'')
				   REPLACE S16.F10 WITH IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F56,'')
				   REPLACE S16.F18 WITH DOW(CTOD(ALLTRIM(HK + '/' + S16.F01)))
				   REPLACE S16.F19 WITH 'N'
				   REPLACE S16.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
		   	   ENDFOR
		   ENDIF
		   SELECT S05_TEMP
		   SKIP
		ENDDO
		SELECT S05_TEMP
		USE		
		*****帶入各項單據
		SELECT F02 FROM S16 ORDER BY F02 GROUP BY F02 INTO CURSOR S16_1
		*****S17.上下班修改作業
		SELECT S17.* FROM S17,S16_1 WHERE S16_1.F02 = S17.F03 AND LEFT(DTOS(S17.F09),6) = LEFT(DTOS(CTOD(HK + '/01')),6) AND S17.F13 = 'Y' AND LEFT(ALLTRIM(S17.F02),2) = 'SA' ORDER BY S17.F02 INTO CURSOR S17B
		SELECT S17B
		GO TOP
		DO WHILE !EOF()
		   FOR I = 0 TO S17B.F10 - S17B.F09
		       D = VAL(RIGHT(DTOC(S17B.F09),2)) + I
		       SELECT S16
		       SEEK S17B.F03 + PADL(ALLTRIM(STR(D)),2,'0')
		       IF FOUND()
				  DO CASE
					 CASE !EMPTY(S17B.F11) AND !EMPTY(S17B.F12)		&&變更上下班  
					      REPLACE S16.F21 WITH S16.F11	&&變更前上班時間
						  REPLACE S16.F22 WITH S16.F12	&&變更前下班時間
			              REPLACE S16.F11 WITH S17B.F11
		                  REPLACE S16.F12 WITH S17B.F12					 
		             CASE !EMPTY(S17B.F11) AND EMPTY(S17B.F12)		&&變更上班
		                  REPLACE S16.F21 WITH S16.F11
		                  REPLACE S16.F11 WITH S17B.F11
		             CASE EMPTY(S17B.F11) AND !EMPTY(S17B.F12)		&&變更下班
		                  REPLACE S16.F22 WITH S16.F12
		                  REPLACE S16.F12 WITH S17B.F12
		          ENDCASE
		       	  REPLACE S16.F20 WITH S17B.F02			&&S17.上下班修改補登作業單號
		       ENDIF
		   ENDFOR
		   SELECT S17B
		   SKIP
		ENDDO
		SELECT S17B
		USE
		SELECT S17
		USE 
		*****S19.加班申請(修改)作業
		IF S00_F65 = '1'	&&加班資料是否依S19加班申請(修改)作業為主(1:是 2:否)
		   SELECT S19.* FROM S19,S16_1 WHERE S16_1.F02 = S19.F03 AND  LEFT(DTOS(CTOD(HK + '/01')),6) = LEFT(DTOS(S19.F09),6) AND S19.F13='Y' ORDER BY S19.F02 INTO CURSOR S19B
		   SELECT S19B
		   GO TOP    
		   DO WHILE !EOF()
			  FOR I = 0 TO S19B.F10 - S19B.F09
			      D = VAL(RIGHT(DTOC(S19B.F09),2)) + I
			      SELECT S16
			      SEEK S19B.F03 + PADL(ALLTRIM(STR(D)),2,'0')
			      IF FOUND()
		             DO CASE
		                CASE !EMPTY(S19B.F11) AND !EMPTY(S19B.F12)
		                     REPLACE S16.F27 WITH S16.F13	&&變更前加班上班時間
		                     REPLACE S16.F28 WITH S16.F14	&&變更前加班下班時間
		                     REPLACE S16.F13 WITH S19B.F11
		                     REPLACE S16.F14 WITH S19B.F12
		                CASE !EMPTY(S19B.F11) AND EMPTY(S19B.F12)
		                	 REPLACE S16.F27 WITH S16.F13
		                     REPLACE S16.F13 WITH S19B.F11
		                CASE EMPTY(S19B.F11) AND !EMPTY(S19B.F12)
		                     REPLACE S16.F28 WITH S16.F14
		                   	 REPLACE S16.F14 WITH S19B.F12
		             ENDCASE
			         REPLACE S16.F26 WITH S19B.F02
			         ****修正平常加班與假日加班之加班上班時間
			         IF S16.F18 <> 1 AND S16.F18 <> 7	&&平常日
			         	IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 <> '1',.T.) = .T.	&&平常日且非假日
			         	   IF S16.F13 < S16.F10
		                      REPLACE S16.F13 WITH S16.F10
		             	   ENDIF
			         	ENDIF
			         ENDIF
			         IF S16.F18 = 1 OR S16.F18 = 7		&&假日
			         	IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 = '2',.F.) = .T.	&&假日補班
			         	   IF S16.F13 < S16.F10
		                      REPLACE S16.F13 WITH S16.F10
		             	   ENDIF
			         	ENDIF
			         ENDIF
			      ENDIF
			  ENDFOR
			  SELECT S19B
			  SKIP
		   ENDDO
		   SELECT S19B
		   USE
		ENDIF
		SELECT S19
		USE
		*****整理加班時間
		SELECT S16
		SET ORDER TO S161
		GO TOP
		DO WHILE !EOF()
		   IF S00_F65 <> '1'	&&加班資料不依S19加班申請(修改)作業為主
		   	  IF S16.F18 <> 1 AND S16.F18 <> 7		&&平常日
		   	  	 IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 <> '1',.T.) = .T.	&&平常日且非假日
		   	  	 	IF S16.F12 > S16.F10 AND !EMPTY(S16.F12)
		   	  	 	   REPLACE S16.F14 WITH S16.F12	&&加班下班時間
		   	  	 	   REPLACE S16.F13 WITH S16.F10	&&加班上班時間
		   	  	 	   REPLACE S16.F12 WITH S16.F10 &&下班時間
		   	  	 	ENDIF
		   	  	 ENDIF
		   	  ENDIF
			  IF S16.F18 = 1 OR S16.F18 = 7			&&假日
			     IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 = '2',.F.) = .T.		&&假日補班
		   	  	 	IF S16.F12 > S16.F10 AND !EMPTY(S16.F12)
		   	  	 	   REPLACE S16.F14 WITH S16.F12	&&加班下班時間
		   	  	 	   REPLACE S16.F13 WITH S16.F10	&&加班上班時間
		   	  	 	   REPLACE S16.F12 WITH S16.F10 &&下班時間
		   	  	 	ENDIF
			     ENDIF
			  ENDIF
			  IF S16.F18 = 1 OR S16.F18 = 7			&&假日
			     IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 <> '2',.T.) = .T.	&&國定假日
			   	    REPLACE S16.F13 WITH S16.F11	&&加班上班時間
			   	    REPLACE S16.F14 WITH S16.F12	&&加班下班時間
			   	    REPLACE S16.F11 WITH ''			&&上班時間
			   	    REPLACE S16.F12 WITH '' 		&&下班時間
			   	 ENDIF
			  ENDIF
		   ELSE					&&假日有加班上下班時間
		   	  IF S16.F18 = 1 OR S16.F18 = 7	OR IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 = '1',.F.)		&&假日
		   	     IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 <> '2',.T.) = .T.	&&國定假日
*!*			   	        IF EMPTY(S16.F26)
		   	     	   REPLACE S16.F11 WITH ''
		   	     	   REPLACE S16.F12 WITH ''
*!*			   	     	ENDIF   
		   	     ENDIF
		   	  ENDIF		   	  
		   ENDIF
		   SELECT S16
		   SKIP
		ENDDO
		*****S18.請假作業  
		SELECT S18.* FROM S18,S16_1 WHERE S16_1.F02 = S18.F03 AND  LEFT(DTOS(CTOD(HK + '/01')),6) >= LEFT(DTOS(S18.F10),6);
		       AND LEFT(DTOS(CTOD(HK + '/01')),6) <= LEFT(DTOS(S18.F11),6) AND S18.F18 = 'Y' AND LEFT(ALLTRIM(S18.F02),2) = 'SB' ORDER BY S18.F02 INTO CURSOR S18B
		SELECT S18B
		GO TOP
		DO WHILE !EOF()
		   StarDate = S18B.F10	&&請假起始日期
		   DO WHILE LEFT(DTOS(StarDate),6) <> LEFT(DTOS(CTOD(HK + '/01')),6)
			  IF (MONTH(StarDate) + 1) < 13
				 StarDate = CTOD(ALLTRIM(STR(YEAR(StarDate))) + '/' + PADL(ALLTRIM(STR(MONTH(StarDate) + 1)),2,'0') + '/01')
			  ELSE
				 StarDate = CTOD(ALLTRIM(STR(YEAR(StarDate) + 1)) + '/01/01')
			  ENDIF
		   ENDDO
		   IF LEFT(DTOS(S18B.F11),6) <> LEFT(DTOS(CTOD(HK + '/01')),6)
*!*				  EndDate = CTOD(LEFT(DTOC(StarDate),6) + ALLTRIM(STR(DAY_NUM(ALLTRIM(LEFT(DTOC(S18B.F11),5))))))
			  EndDate = CTOD(LEFT(DTOC(StarDate),6) + ALLTRIM(STR(DAY_NUM(ALLTRIM(LEFT(DTOC(StarDate),5))))))
		   ELSE
			  EndDate = S18B.F11 &&請假截止日期
		   ENDIF 
		   FOR I = 0 TO DAY(EndDate) - DAY(StarDate)
			   SELECT S16
			   SEEK S18B.F03 + PADL(ALLTRIM(STR(I + DAY(StarDate))),2,'0')
			   IF FOUND()
				  REPLACE S16.F15 WITH 0
				  REPLACE S16.F16 WITH 0
				  REPLACE S16.F23 WITH S18B.F02
				  REPLACE S16.F24 WITH IIF(DAY(StarDate) = (I + DAY(StarDate)),S18B.F12,ALLTRIM(IIF(INDEXSEEK(S18B.F03,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F50,''),'')))
				  REPLACE S16.F25 WITH IIF((I + DAY(StarDate)) = DAY(EndDate),S18B.F13,ALLTRIM(IIF(INDEXSEEK(S18B.F03,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F51,''),'')))
				  REPLACE S16.F29 WITH S16.F11		&&請假前之上班時間
				  REPLACE S16.F30 WITH S16.F12		&&請假前之下班時間
		          *****調整上下班時間
		          DO CASE
		             CASE S16.F24 = S16.F06 AND S16.F25 = S16.F07	&&請整天
		                  REPLACE S16.F11 WITH ''
		                  REPLACE S16.F12 WITH ''
		             CASE S16.F24 <> S16.F06 AND S16.F25 = S16.F07
		                  REPLACE S16.F12 WITH S16.F24
		             CASE S16.F24 = S16.F06 AND S16.F25 <> S16.F07
		                  REPLACE S16.F11 WITH S16.F25
		          ENDCASE
				  S00_F52 = ALLTRIM(IIF(INDEXSEEK(S16.F02,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F52,''),''))	&&班別:休息時段起
				  S00_F53 = ALLTRIM(IIF(INDEXSEEK(S16.F02,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F53,''),''))	&&班別:休息時段止
		          IF !EMPTY(S16.F11) AND !EMPTY(S16.F12) AND !EMPTY(S00_F52) AND !EMPTY(S00_F53)
		             IF BETWEEN(S16.F11,S00_F52,S00_F53) AND S16.F12 > S00_F53
		                REPLACE S16.F11 WITH S00_F53
		             ENDIF
		             IF BETWEEN(S16.F12,S00_F52,S00_F53) AND S16.F11 < S00_F52
		                REPLACE S16.F12 WITH S00_F52
		             ENDIF
		          ENDIF
		          *****整理非假日補班之請假
				  IF S16.F18 = 1 OR S16.F18 = 7			&&假日
				     IF IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.T.,'S15','S151'),S15.F03 <> '2',.T.) = .T.		&&假日補班
			   	  	 	IF !EMPTY(S16.F23)
			   	  	 	   REPLACE S16.F23 WITH ''
			   	  	 	   REPLACE S16.F24 WITH ''
			   	  	 	   REPLACE S16.F25 WITH ''
			   	  	 	ENDIF
				     ENDIF
				  ENDIF
		       ENDIF
		   ENDFOR
		   SELECT S18B
		   SKIP
		ENDDO
		SELECT S18B
		USE
		SELECT S18
		USE
		SELECT S16_1
		USE
		*****計算遲到、早退、加班時數
		SELECT S16
		SET ORDER TO S161
		GO TOP
		DO WHILE !EOF()
		   S00_F52 = ALLTRIM(IIF(INDEXSEEK(S16.F02,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F52,''),''))	&&班別:休息時段起
		   S00_F53 = ALLTRIM(IIF(INDEXSEEK(S16.F02,.T.,'S05','S051'),IIF(INDEXSEEK('I' + S05.F29,.T.,'S00','S005'),S00.F53,''),''))	&&班別:休息時段止
		   S00_F52_H = VAL(LEFT(ALLTRIM(S00_F52),2))		&&班別:休息時段起(時)
		   S00_F52_M = VAL(RIGHT(ALLTRIM(S00_F52),2))		&&班別:休息時段起(分)
		   S00_F53_H = VAL(LEFT(ALLTRIM(S00_F53),2))		&&班別:休息時段止(時)
		   S00_F53_M = VAL(RIGHT(ALLTRIM(S00_F53),2))		&&班別:休息時段止(分)
		   S16_F07_H = VAL(LEFT(ALLTRIM(S16.F07),2))		&&班別:下班時段(時)
		   S16_F07_M = VAL(RIGHT(ALLTRIM(S16.F07),2))   	&&班別:下班時段(分)
		   S16_F08_H = VAL(LEFT(ALLTRIM(S16.F08),2))		&&班別:上班遲到緩衝時間(時)
		   S16_F08_M = VAL(RIGHT(ALLTRIM(S16.F08),2))   	&&班別:上班遲到緩衝時間(分)
		   S16_F09_H = VAL(LEFT(ALLTRIM(S16.F09),2))		&&班別:下班早退緩衝時間(時)
		   S16_F09_M = VAL(RIGHT(ALLTRIM(S16.F09),2))   	&&班別:下班早退緩衝時間(分)
		   S16_F10_H = VAL(LEFT(ALLTRIM(S16.F10),2))		&&班別:加班上班緩衝時間(時)
		   S16_F10_M = VAL(RIGHT(ALLTRIM(S16.F10),2))   	&&班別:加班上班緩衝時間(分)
		   S16_F11_H = VAL(LEFT(ALLTRIM(S16.F11),2))		&&上班時間(時)
		   S16_F11_M = VAL(RIGHT(ALLTRIM(S16.F11),2))		&&上班時間(分)
		   S16_F12_H = VAL(LEFT(ALLTRIM(S16.F12),2))		&&下班時間(時)
		   S16_F12_M = VAL(RIGHT(ALLTRIM(S16.F12),2))		&&下班時間(分)
		   S16_F13_H = VAL(LEFT(ALLTRIM(S16.F13),2))		&&加班上班時間(時)
		   S16_F13_M = VAL(RIGHT(ALLTRIM(S16.F13),2))		&&加班上班時間(分)
		   S16_F14_H = VAL(LEFT(ALLTRIM(S16.F14),2)) + IIF(LEFT(ALLTRIM(S16.F14),2) = '00',24,0)	&&加班下班時間(時)
		   S16_F14_M = VAL(RIGHT(ALLTRIM(S16.F14),2))		&&加班下班時間(分)
		   *****計算遲到分鐘數
		   IF S16.F11 >= S16.F08 AND !EMPTY(S16.F11) AND !EMPTY(S16.F08)
		      DO CASE
		         CASE S16_F11_M >= S16_F08_M
		              REPLACE S16.F15 WITH ((S16_F11_H - S16_F08_H) * 60) + (S16_F11_M - S16_F08_M)
		         CASE S16_F12_M < S16_F08_M
		              REPLACE S16.F15 WITH ((S16_F11_H - S16_F08_H) * 60) - (S16_F08_M - S16_F11_M)
		      ENDCASE
		   ENDIF
		   **判斷是否有請假
		   IF !EMPTY(S16.F23) AND S16.F11 >= S16.F25 AND !EMPTY(S16.F11)
		   	   REPLACE S16.F15 WITH 0
		   ENDIF
		   **無上班時間則遲到為零
		   IF EMPTY(S16.F11)
		      REPLACE S16.F15 WITH 0
		   ENDIF
		   *****計算早退分鐘數
		   IF S16.F12 <= S16.F09 AND !EMPTY(S16.F12) AND !EMPTY(S16.F09)
		      DO CASE
		         CASE S16_F12_M <= S16_F09_M
		              REPLACE S16.F16 WITH ((S16_F09_H - S16_F12_H) * 60) + (S16_F09_M - S16_F12_M)
		         CASE S16_F12_M > S16_F09_M
		              REPLACE S16.F16 WITH ((S16_F09_H - S16_F12_H) * 60) - (S16_F12_M - S16_F09_M)
		      ENDCASE
		   ENDIF
		   **判斷是否有請假
		   IF !EMPTY(S16.F23) AND S16.F12 <= S16.F25 AND !EMPTY(S16.F12)
		   	  REPLACE S16.F16 WITH 0
		   ENDIF
		   **無下班時間則早退為零
		   IF EMPTY(S16.F12)
		      REPLACE S16.F16 WITH 0
		   ENDIF		   
		   *****計算加班小時數
		   S16_F171=0
		   IF S16.F13 < S16.F14
		      IF S16.F18 = 7 OR S16.F18 = 1 OR IIF(INDEXSEEK(DTOS(CTOD(HK + '/' + S16.F01)),.F.,'S15','S151'),S15.F03 = '1',.F.)	&&假日加班
		         IF S16.F14 <= S16.F10
		   			DO CASE
		           	   CASE S16_F13_M <= S16_F14_H
		                    S16_F171 = ((S16_F14_H - S16_F13_H) * 60) + (S16_F14_M - S16_F13_M)
		           	   CASE S16_F13_M > S16_F14_H
		                    S16_F171 = ((S16_F14_H - S16_F13_H) * 60) - (S16_F13_M - S16_F14_M)
		            ENDCASE
		         ELSE
		   			DO CASE 
		           	   CASE S16_F13_M <= S16_F07_H
		                    S16_F171 = ((S16_F07_H - S16_F13_H) * 60) + (S16_F07_M - S16_F13_M)
		           	   CASE S16_F13_M > S16_F07_H
		                    S16_F171 = ((S16_F07_H - S16_F13_H) * 60) - (S16_F13_M - S16_F07_M)
		            ENDCASE
			        DO CASE 
			           CASE S16_F10_M <= S16_F14_H
			                S16_F171 = S16_F171 + ((S16_F14_H - S16_F10_H) * 60) + (S16_F14_M - S16_F10_M)
			           CASE S16_F10_M > S16_F14_H
			                S16_F171 = S16_F171 + ((S16_F14_H - S16_F10_H) * 60) - (S16_F10_M - S16_F14_M)
			        ENDCASE
		         ENDIF
		         IF S16.F13 < S00_F52 AND S16.F14 > S00_F53
		            S16_F171 = S16_F171 - ((S00_F53_H - S00_F52_H) * 60) + (S00_F53_M - S00_F52_M)
		         ENDIF
		      ELSE		&&一般上班之加班
		         DO CASE
		            CASE S16.F13 >= S16.F10
		                 DO CASE 
		                    CASE S16_F13_M <= S16_F14_H
		                         S16_F171 = ((S16_F14_H - S16_F13_H) * 60) + (S16_F14_M - S16_F13_M)
		                    CASE S16_F13_M > S16_F14_H
		                       	 S16_F171 = ((S16_F14_H - S16_F13_H) * 60) - (S16_F13_M - S16_F14_M)
		                 ENDCASE
		            CASE S16.F13 < S16.F10
		                 DO CASE
		                    CASE S16_F10_M <= S16_F14_H
		                         S16_F171 = ((S16_F14_H - S16_F10_H) * 60) + (S16_F14_M - S16_F10_M)
		                    CASE S16_F10_M > S16_F14_H
		                       	 S16_F171 = ((S16_F14_H - S16_F10_H) * 60) - (S16_F10_M - S16_F14_M)
		                 ENDCASE
		        ENDCASE
		      ENDIF
		      *****判斷加班計算範圍
			  IF MOD(S16_F171,60) = 0
				 REPLACE S16.F17 WITH ROUND(S16_F171 / 60,1)
			  ELSE
				 DO CASE
			        CASE S00_F63 <> 'F' AND S00_F64 <> 'F'
			           	 IF MOD(S16_F171,60) >= VAL(S00_F63) AND MOD(S16_F171,60) < VAL(S00_F64)
			                REPLACE S16.F17 WITH INT(S16_F171 / 60) + 0.5
			             ELSE
			                IF MOD(S16_F171,60) >= VAL(S00_F64) AND MOD(S16_F171,60) < 60
			             	   REPLACE S16.F17 WITH INT(S16_F171 / 60) + 1
			                ELSE
			             	   REPLACE S16.F17 WITH INT(S16_F171 / 60)
			                ENDIF
			             ENDIF
			        CASE S00_F63 <> 'F' AND S00_F64 = 'F'
			             IF MOD(S16_F171,60) >= VAL(S00_F63) AND MOD(S16_F171,60) < 60
			                REPLACE S16.F17 WITH INT(S16_F171 / 60) + 0.5
			             ELSE
			                REPLACE S16.F17 WITH INT(S16_F171 / 60)
			             ENDIF
			        CASE S00_F63 = 'F' AND S00_F64 <> 'F'
			             IF MOD(S16_F171,60) >= VAL(S00_F64) AND MOD(S16_F171,60) < 60
			                REPLACE S16.F17 WITH INT(S16_F171 / 60) + 1
			             ELSE
			                REPLACE S16.F17 WITH INT(S16_F171 / 60)
			             ENDIF
			        CASE S00_F63 = 'F' AND S00_F64 = 'F'
			             REPLACE S16.F17 WITH ROUND(S16_F171 / 60,1)
			     ENDCASE
		      ENDIF
		   ENDIF
		   ***判斷加班小時數是否為零
		   IF S16.F17 = 0
		      REPLACE S16.F13 WITH ''
		      REPLACE S16.F14 WITH ''
		   ENDIF
		   ***無加班上下班時間加班小時數也為零
		   IF EMPTY(S16.F13) AND EMPTY(S16.F14)
		      REPLACE S16.F17 WITH 0
		   ENDIF
		   SELECT S16
		   SKIP
		ENDDO
		******
        SELECT S00
        USE
        SELECT S05
        USE
        SELECT S06
        USE
        SELECT S16
        USE
        WAIT CLEAR
        =MESSAGEBOX('刷卡資料已成功轉至S16.上下班紀錄作業!',0+64,'提示視窗')
     ENDIF
     THISFORM.GRID1.SETFOCUS
 ENDPROC   
ENDDEFINE      
***********************************************列印程序
PROCEDURE PNT_PRC  
   SELECT S21
   ANS_RANGE=CREATEOBJECT("ANS_RANGE")  
   ANS_RANGE.SHOW    
ENDPROC 
***********************************************列印範圍選定
DEFINE CLASS ANS_RANGE AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='請輸入列印資料範圍'
  TOP=INT_015*200
  LEFT=INT_015*250 
  HEIGHT=INT_015*120
  WIDTH=INT_015*280
  FONTSIZE=INT_015*9
  MOVABLE=.F.
  MAXBUTTON=.F.
  MINBUTTON=.F.   
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='ANS_RANGE'
  ADD OBJECT LABEL1 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*20,;
      TOP=INT_015*36,;
      CAPTION='起始員工編號',;
      NAME='LABEL1'
  ADD OBJECT TEXT1 AS TEXTBOX WITH;
      TOP=INT_015*30,;
      LEFT=INT_015*100,;
      HEIGHT=INT_015*25,; 
      WIDTH=INT_015*70,; 
      MAXLENGTH=4,;
      NAME='TEXT1'      
  ADD OBJECT LABEL2 AS LABEL WITH;
      AUTOSIZE=.T.,;
      HEIGHT=INT_015*25,;
      LEFT=INT_015*20,;
      TOP=INT_015*66,;
      CAPTION='截止員工編號',;
      NAME='LABEL2'
  ADD OBJECT TEXT2 AS TEXTBOX WITH;
      TOP=INT_015*60,;
      LEFT=INT_015*100,;
      HEIGHT=INT_015*25,;   
      WIDTH=INT_015*70,;     
      MAXLENGTH=4,;   
      NAME='TEXT2'    
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*190,;    
      HEIGHT=INT_015*25,; 
      WIDTH=INT_015*70,;    
      TOP=INT_015*14,;  
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<P.列         印',;
      TOOLTIPTEXT='列印所輸入的範圍鍵值!快速鍵->ALT+Y',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*190,;    
      HEIGHT=INT_015*25,; 
      WIDTH=INT_015*70,;    
      TOP=INT_015*45,;  
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<E.轉EXCEL',;
      TOOLTIPTEXT='將資料轉至EXCEL格式!快速鍵->ALT+Y',;
      NAME='CMND2'      
  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*190,;
      TOP=INT_015*76,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*70,; 
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<C.取        消',;
      TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
      NAME='CMND3'
  PROCEDURE INIT 
      THISFORM.SETALL('FONTSIZE',INT_015*10,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*10,'LABEL')      
      THISFORM.SETALL('FONTSIZE',INT_015*10,'COMMANDBUTTON')
      THISFORM.TEXT1.READONLY=IIF(A02.F10='*',.F.,.T.)  
      THISFORM.TEXT2.READONLY=IIF(A02.F10='*',.F.,.T.)  
      DO CASE  
             CASE KEY()='F01'   
                        THISFORM.LABEL1.CAPTION='起始員工編號' 
                        THISFORM.LABEL2.CAPTION='截止員工編號' 
                        THISFORM.TEXT1.MAXLENGTH=4
                        THISFORM.TEXT2.MAXLENGTH=4
                        THISFORM.TEXT1.VALUE=S21.F01
                        THISFORM.TEXT2.VALUE=S21.F01
             CASE KEY()='F02' 
                        THISFORM.LABEL1.CAPTION='起始員工姓名' 
                        THISFORM.LABEL2.CAPTION='截止員工姓名'        
                        THISFORM.TEXT1.MAXLENGTH=8
                        THISFORM.TEXT2.MAXLENGTH=8                        
                        THISFORM.TEXT1.VALUE=S21.F02
                        THISFORM.TEXT2.VALUE=S21.F02                              
      ENDCASE     
  ENDPROC   
  *** 
  PROCEDURE CMND1.CLICK 
      IF  EMPTY(THISFORM.TEXT1.VALUE)
           =MESSAGEBOX('起始'+IIF(KEY()='F01','編號','姓名')+'不得為空白!',0+64,'提示訊息視窗')
           THISFORM.TEXT1.SETFOCUS
  	   RETURN 
      ENDIF
      IF  EMPTY(THISFORM.TEXT2.VALUE)
           =MESSAGEBOX('截止'+IIF(KEY()='F01','編號','姓名')+'不得為空白!',0+64,'提示訊息視窗')
           THISFORM.TEXT2.SETFOCUS
  	   RETURN 
      ENDIF      
      DO CASE  
             CASE KEY()='F01' 
                        SELECT  * FROM S21 WHERE F01>=THISFORM.TEXT1.VALUE AND F01<=THISFORM.TEXT2.VALUE ORDER BY F01,F04,F05,F06 NOWAIT                        
		        IF  _TALLY >0
		             REPORT FORM ALLTRIM(INT_116)+'S21_1'  TO PRINTER PROMPT  PREVIEW  
		        ELSE
		             =MESSAGEBOX('查無此範圍資料',0+48,'警示')
		             THISFORM.TEXT1.SETFOCUS
		             RETURN     
		        ENDIF                          
             CASE KEY()='F02'   
                        SELECT  * FROM S21 WHERE F02>=THISFORM.TEXT1.VALUE AND F02<=THISFORM.TEXT2.VALUE  ORDER BY F02,F04,F05,F06 NOWAIT		   
                        IF  _TALLY >0
		             REPORT FORM ALLTRIM(INT_116)+'S21_2' TO PRINTER PROMPT  PREVIEW
		        ELSE
		             =MESSAGEBOX('查無此範圍資料',0+48,'警示')
		             THISFORM.TEXT1.SETFOCUS
		             RETURN     
		        ENDIF                               
      ENDCASE         
           THISFORM.CMND3.CLICK
  ENDPROC  
  ****
  PROCEDURE CMND2.CLICK 
      IF  EMPTY(THISFORM.TEXT1.VALUE)
           =MESSAGEBOX('起始'+IIF(KEY()='F01','編號','姓名')+'不得為空白!',0+64,'提示訊息視窗')
           THISFORM.TEXT1.SETFOCUS
  	   RETURN 
      ENDIF
      IF  EMPTY(THISFORM.TEXT2.VALUE)
           =MESSAGEBOX('截止'+IIF(KEY()='F01','編號','姓名')+'不得為空白!',0+64,'提示訊息視窗')
           THISFORM.TEXT2.SETFOCUS
  	   RETURN 
      ENDIF         
      FILENAME='S21 刷卡出勤資料'+LEFT(DTOS(CTOD(HK+'/01')),6)
      DO CASE  
             CASE KEY()='F01' 
                        SELECT F01 員工編號,F02 員工姓名,F04 出勤日期,IIF(F05='1','上午','下午') 進出時段,F06 進出時間,F03 卡片號碼;
                                       FROM S21 WHERE F01>=THISFORM.TEXT1.VALUE AND F01<=THISFORM.TEXT2.VALUE ORDER BY F01,F04,F05,F06 NOWAIT
		        IF  _TALLY >0
		             GCDELIMFILE=PUTFILE('儲存路徑',FILENAME,'XLS')
			     IF !EMPTY(GCDELIMFILE) 
			          ON ERROR DO FILE_IMPACT
			          COPY TO (GCDELIMFILE) TYPE XL5     
			          =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗') 
			     ELSE
		                 THISFORM.TEXT1.SETFOCUS
		                 RETURN				     
			     ENDIF  
		        ELSE
		             =MESSAGEBOX('查無此範圍資料',0+48,'警示')
		             THISFORM.TEXT1.SETFOCUS
		             RETURN     
		        ENDIF                          
             CASE KEY()='F02'   
                        SELECT F02 員工姓名,F01 員工編號,F04 出勤日期,IIF(F05='1','上午','下午') 進出時段,F06 進出時間,F03 卡片號碼;
                                       FROM S21 WHERE F02>=THISFORM.TEXT1.VALUE AND F02<=THISFORM.TEXT2.VALUE ORDER BY F02,F04,F05,F06 NOWAIT
                        IF  _TALLY >0
		             GCDELIMFILE=PUTFILE('儲存路徑',FILENAME,'XLS')
			     IF !EMPTY(GCDELIMFILE) 
			          ON ERROR DO FILE_IMPACT
			          COPY TO (GCDELIMFILE) TYPE XL5     
			          =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')  
			     ELSE
		                 THISFORM.TEXT1.SETFOCUS
		                 RETURN			     
			     ENDIF  		
		       ELSE
		             =MESSAGEBOX('查無此範圍資料',0+48,'警示')
		             THISFORM.TEXT1.SETFOCUS
		             RETURN     
		       ENDIF                               
      ENDCASE         
           THISFORM.CMND3.CLICK
  ENDPROC           
  ****
  PROCEDURE CMND3.CLICK
      THISFORM.RELEASE
  ENDPROC      
  PROC CMND1.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND1.TOP=INT_015*15
             THISFORM.CMND1.LEFT=INT_015*191
  ENDPROC     
  PROC CMND1.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND1.TOP=INT_015*14
             THISFORM.CMND1.LEFT=INT_015*190                  
  ENDPROC     
  PROC CMND2.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND2.TOP=INT_015*46
             THISFORM.CMND2.LEFT=INT_015*191
  ENDPROC     
  PROC CMND2.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND2.TOP=INT_015*45
             THISFORM.CMND2.LEFT=INT_015*190                  
  ENDPROC     
  PROC CMND3.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND3.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND3.TOP=INT_015*77
             THISFORM.CMND3.LEFT=INT_015*191
  ENDPROC     
  PROC CMND3.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND3.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND3.TOP=INT_015*76
             THISFORM.CMND3.LEFT=INT_015*190                  
  ENDPROC   
ENDDEFINE   
*****假日判斷
FUNCTION DAY_WEEK
     PARA JUK4
     JO4=''
     DO CASE
        CASE JUK4=2
             JO4='星期一'     
        CASE JUK4=3
             JO4='星期二'   
        CASE JUK4=4
             JO4='星期三'   
        CASE JUK4=5
             JO4='星期四'   
        CASE JUK4=6
             JO4='星期五'   
        CASE JUK4=7
             JO4='星期六'   
        CASE JUK4=1
             JO4='星期日' 
     ENDCASE
     RETURN JO4               
       