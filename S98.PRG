**程式名稱:S98.出勤資料轉入薪資作業
SET EXCLUSIVE OFF
CLOSE ALL
CLEAR 
PRG_NO='S98'
IF !USED('A01')  &&操作人員資料設定
   SELECT 0
   USE A01
ELSE 
   SELECT A01
ENDIF
SET ORDER TO A01
IF !USED('A02')
   SELECT 0
   USE A02 INDE A02
ELSE
   SELECT A02
ENDIF
SET ORDER TO A021
SEEK SYS_OPER + 'S98'
*****
S98FORM=CREATEOBJECT("TKS98")
S98FORM.SHOW
DEFINE CLASS TKS98 AS FORM
  CAPTION='S98.出勤資料轉入薪資作業'
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
  WIDTH=INT_015*791
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKS20'
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*131,;
      TOP=INT_015*340,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<C.選擇',;
      TOOLTIPTEXT='選擇此作業畫面!快速鍵->ALT+A'
      NAME='CMND1'
  PROCEDURE INIT
      v_scrx=SYSMETRIC(1)
      v_scry=SYSMETRIC(2)
      DO CASE
         CASE v_scrx=640 .and. v_scry=480 
              THISFORM.PICTURE='BMPS\XP800600.JPG'
         CASE v_scrx=800 .and. v_scry=600 
              THISFORM.PICTURE='BMPS\XP800600.JPG'
         CASE v_scrx=1024 .and. v_scry=768
  	          THISFORM.PICTURE='BMPS\XP1024768.JPG'
         OTHERWISE
     	      THISFORM.PICTURE='BMPS\XP1024768.JPG'
      ENDCASE
 ENDPROC
 PROCEDURE ACTIVATE
      THISFORM.CMND1.SETFOCUS
      THISFORM.CMND1.CLICK
 ENDPROC
 PROCEDURE CMND1.CLICK
      TKS98_TEMP=CREATEOBJECT("TKS98_TEMP")
      TKS98_TEMP.SHOW
      THISFORM.RELEASE
      IF !USED('A04')
         SELECT 0
         USE A04 INDE A04
      ELSE
         SELECT A04
      ENDIF
      SET ORDER TO A041
      SEEK SYS_OPER
      IF FOUND()
         REPLACE A04.F02 WITH ''
         REPLACE A04.F03 WITH ''
         REPLACE A04.F04 WITH ''
      ENDIF
      UNLOCK ALL
      CLOSE TABLE ALL
      THISFORM.RELEASE
      RETURN
ENDPROC
ENDDEFINE
*******
DEFINE CLASS TKS98_TEMP AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='S98.出勤資料轉入薪資作業'
  FONTSIZE=INT_015*9
  TOP=INT_015*200
  LEFT=INT_015*215
  HEIGHT=INT_015*95
  WIDTH=INT_015*340
  MOVABLE=.F.
  MAXBUTTON=.F.
  MINBUTTON=.F.
  CONTROLBOX=.F.
  BORDERSTYLE=1
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKS98_TEMP'
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*45,;
      TOP=INT_015*15,;
      BACKCOLOR=RGB(255,255,0),;
      AUTOSIZE=.T.,;
      CAPTION='請輸入欲轉入出勤資料之西元年月'
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*60,;
      WIDTH=INT_015*75,;
      CAPTION='出勤西元年月'
  ADD OBJECT TXTF02 AS TEXTBOX WITH;
      LEFT=INT_015*90,;
      TOP=INT_015*54,;
      MAXLENGTH=7,;
      INPUTMASK='9999/99',;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*60,;
      NAME='TXTF02'
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*225,;
      TOP=INT_015*55,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;
      CAPTION='\<Y.執行',;
      TOOLTIPTEXT='執行所選取之內容->ALT+Y',;
      NAME='CMND1'
 ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*280,;
      TOP=INT_015*55,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='離開此作業!快速鍵->ALT+X'
      NAME='CMND2'
********************
 PROCEDURE INIT
      THISFORM.SETALL('FONTSIZE',INT_015*10,'LABEL')
      THISFORM.SETALL('FONTSIZE',INT_015*10,'TEXTBOX')
      THISFORM.SETALL('FONTSIZE',INT_015*10,'COMMANDBUTTON')
      THISFORM.LBLF01.FONTSIZE = INT_015*12
      THISFORM.TXTF02.READONLY = .F.
      THISFORM.TXTF02.VALUE = ALLTRIM(STR(YEAR(DATE()))) + '/' + PADL(ALLTRIM(STR(MONTH(DATE()))),2,'0')
      THISFORM.TXTF02.SETFOCUS
 ENDPROC   
********確認鍵******
PROCEDURE CMND1.CLICK
      IF LEN(ALLTRIM(THISFORM.TXTF02.VALUE)) = 1 
	     =MESSAGEBOX('出勤西元年月不得為空白,請重新輸入!',0+64,'提示視窗')
	     THISFORM.TXTF02.SETFOCUS
	     RETURN
      ENDIF
      IF LEN(ALLTRIM((THISFORM.TXTF02.VALUE))) <> 7 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2)) <= 0 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2)) >= 13
         =MESSAGEBOX('出勤西元年月格式不正確,請重新輸入',0+64,'提示視窗')
         THISFORM.TXTF02.SETFOCUS
         RETURN
      ENDIF
      S10 = 'S10' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
      IF FILE(ALLTRIM(S10) + '.DBF') = .F.
         =MESSAGEBOX(THISFORM.TXTF02.VALUE + '之S10.薪資計算作業資料表尚未開檔,請開檔後在執行此作業',0+64,'提示視窗')
         THISFORM.TXTF02.SETFOCUS
         RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S0A.DBF') = .T.	&&年度缺勤紀錄檔
         IF !USED(ALLTRIM(INT_999) + '\S0A')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S0A' SHARED
            SELECT * FROM S0A WHERE !EMPTY(S0A.F10) AND LEFT(DTOS(S0A.F08),6) = LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6) ORDER BY S0A.F01,S0A.F05 INTO CURSOR S0AA
            SELECT S0A
            USE
         ELSE
            =MESSAGEBOX('德通系統資料之 S0A.DBF 錯誤請查明後在執行此作業',0+64,'提示視窗')
            THISFORM.TXTF02.SETFOCUS
            RETURN
         ENDIF
      ELSE
         =MESSAGEBOX('德通系統資料之 S0A.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
         THISFORM.TXTF02.SETFOCUS
         RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S14.DBF') = .T.	&&年度出勤紀錄檔
         IF !USED(ALLTRIM(INT_999) + '\S14')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S14'  SHARED
            SET ORDER TO S141
            SELECT * FROM S14 WHERE S14.F08 = CTOD(THISFORM.TXTF02.VALUE + '/01') ORDER BY S14.F01 INTO CURSOR S14_TEMP
            SELECT S14
            USE
         ELSE
            =MESSAGEBOX('德通資料之 S14.DBF 錯誤請查明後在執行此作業',0+64,'提示視窗')
            THISFORM.TXTF02.SETFOCUS
            RETURN
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S14.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      ******
      IF FILE(ALLTRIM(INT_999) + '\S0B.DBF') = .T.	&&年假記錄檔 
         IF !USED(ALLTRIM(INT_999) + '\S0B')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S0B' SHARED
            SELECT * FROM S0B ORDER BY F01 INTO CURSOR S0B_TK
            SELECT S0B
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S0B.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S00.DBF') = .T.	&&薪資參數設定
         IF !USED(ALLTRIM(INT_999) + '\S00')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S00' SHARED
            SELECT * FROM S00 ORDER BY F01 INTO CURSOR S00_TK
            SELECT S00
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S00.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S01.DBF') = .T.	&&薪資資料設定
         IF !USED(ALLTRIM(INT_999) + '\S00')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S01' SHARED
            SELECT * FROM S01 ORDER BY F01 INTO CURSOR S01_TK
            SELECT S01
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S01.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S02.DBF') = .T.	&&所得稅表
         IF !USED(ALLTRIM(INT_999) + '\S0B')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S02' SHARED
            SELECT * FROM S02 ORDER BY F01,F02 INTO CURSOR S02_TK
            SELECT S02
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S02.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S03.DBF') = .T.	&&健保費表
         IF !USED(ALLTRIM(INT_999) + '\S03')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S03' SHARED
            SELECT * FROM S03 ORDER BY F01 INTO CURSOR S03_TK
            SELECT S03
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S03.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S04.DBF') = .T.	&&勞保費表
         IF !USED(ALLTRIM(INT_999) + '\S04')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S04' SHARED
            SELECT * FROM S04 ORDER BY F01 INTO CURSOR S04_TK
            SELECT S04
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S04.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S05.DBF') = .T.	&&人事薪資基本檔
         IF !USED(ALLTRIM(INT_999) + '\S05')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S05' SHARED
            SELECT * FROM S05 ORDER BY F01 INTO CURSOR S05_TK
            SELECT S05
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S05.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S06.DBF') = .T.	&&薪資項目及扣款項目
         IF !USED(ALLTRIM(INT_999) + '\S06')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S06' SHARED
            SELECT * FROM S06 WHERE F05 = '2' AND F06 = 'J' OR F06 = 'K' OR F06 = 'O' ORDER BY F01 INTO CURSOR S06_TK
            SELECT S06
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S06.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      IF FILE(ALLTRIM(INT_999) + '\S07.DBF') = .T.	&&調薪修改作業
         IF !USED(ALLTRIM(INT_999) + '\S07')
            SELECT  0
            USE ALLTRIM(INT_999) + '\S07' SHARED
            SELECT * FROM S07 WHERE S07.F15 <> '2' ORDER BY F01 INTO CURSOR S07_TK
            SELECT S07
            USE
         ENDIF
      ELSE
          =MESSAGEBOX('德通資料之 S07.DBF 路徑錯誤請查明後在執行此作業',0+64,'提示視窗')
          THISFORM.TXTF02.SETFOCUS
          RETURN
      ENDIF
      ****
      IF MESSAGEBOX('此轉出作業會把 S10.薪資計算作業中尚未確認之資料全部刪除後重新轉入,是否確定要執行',4+32+256,'請確認')=6
         IF !USED('S0A')		&&年度缺勤天數累計
             SELECT  0
             USE S0A
         ELSE
             SELECT  S0A
         ENDIF
         SET ORDER TO S0A1
         IF !USED('S00')		&&薪資參數設定
            SELECT 0
            USE S00
         ELSE
            SELECT S00
         ENDIF
	     SET ORDER TO S001	&& F01+F05
	     ***帶入德通系統中S00.薪資參數設定有改變之資料
	     SELECT S00_TK
	     GO TOP
	     DO WHILE !EOF()
	     	IF S00_TK.F01 = 'B'		&&加班費設定
	     	   SELECT S00
	     	   SET ORDER TO S001	&& F01+F05
	     	   SEEK 'B' + S00_TK.F05
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'B'
	     	      REPLACE S00.F05 WITH S00_TK.F05
	     	   ENDIF
     	   	   REPLACE S00.F06 WITH S00_TK.F06
     	   	   REPLACE S00.F07 WITH S00_TK.F07
     	   	   REPLACE S00.F08 WITH S00_TK.F08
     	   	   REPLACE S00.F09 WITH S00_TK.F09
     	   	   REPLACE S00.F10 WITH S00_TK.F10
     	   	   REPLACE S00.F46 WITH S00_TK.F46
     	   	   REPLACE S00.F47 WITH S00_TK.F47
     	   	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
	     	IF S00_TK.F01 = 'C'		&&遲到設定
	     	   SELECT S00
	     	   SEEK 'C'
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'C'
	     	   ENDIF
	     	   REPLACE S00.F11 WITH S00_TK.F11
	     	   REPLACE S00.F12 WITH S00_TK.F12
	     	   REPLACE S00.F13 WITH S00_TK.F13
	     	   REPLACE S00.F14 WITH S00_TK.F14
	     	   REPLACE S00.F15 WITH S00_TK.F15
	     	   REPLACE S00.F16 WITH S00_TK.F16
	     	   REPLACE S00.F17 WITH S00_TK.F17
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
	     	IF S00_TK.F01 = 'D'		&&早退設定
	     	   SELECT S00
	     	   SEEK 'D'
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'D'
	     	   ENDIF
	     	   REPLACE S00.F11 WITH S00_TK.F11
	     	   REPLACE S00.F12 WITH S00_TK.F12
	     	   REPLACE S00.F13 WITH S00_TK.F13
	     	   REPLACE S00.F14 WITH S00_TK.F14
	     	   REPLACE S00.F15 WITH S00_TK.F15
	     	   REPLACE S00.F16 WITH S00_TK.F16
	     	   REPLACE S00.F17 WITH S00_TK.F17
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
	     	IF S00_TK.F01 = 'E'		&&年假設定
	     	   SELECT S00
	     	   SET ORDER TO S002	&&F01+ALLTRIM(STR(F18))+ALLTRIM(STR(F19))
	     	   SEEK 'E' + ALLTRIM(STR(S00_TK.F18)) + ALLTRIM(STR(S00_TK.F19))
	     	   IF !FOUND()
	     	   	  SELECT S00
	     	   	  APPEND BLANK 
	     	   	  REPLACE S00.F18 WITH S00_TK.F18
	     	   	  REPLACE S00.F19 WITH S00_TK.F19
	     	   ENDIF
	     	   REPLACE S00.F20 WITH S00_TK.F20
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF	   
	     	IF S00_TK.F01 = 'F'		&&年假限制設定
	     	   SELECT S00
	     	   SEEK 'F'
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'F'
	     	   ENDIF
	     	   REPLACE S00.F21 WITH S00_TK.F21
	     	   REPLACE S00.F22 WITH S00_TK.F22
	     	   REPLACE S00.F23 WITH S00_TK.F23
	     	   REPLACE S00.F24 WITH S00_TK.F24
	     	   REPLACE S00.F25 WITH S00_TK.F25
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF	    
	     	IF S00_TK.F01 = 'G'		&&請假扣款設定
	     	   SELECT S00
	     	   SET ORDER TO S003	&& F01+F26
	     	   SEEK 'G' + S00_TK.F26
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK
	     	      REPLACE S00.F01 WITH 'G'
	     	      REPLACE S00.F26 WITH S00_TK.F26
	     	   ENDIF
	     	   REPLACE S00.F26 WITH S00_TK.F26
	     	   REPLACE S00.F27 WITH S00_TK.F27
	     	   REPLACE S00.F28 WITH S00_TK.F28
	     	   REPLACE S00.F29 WITH S00_TK.F29
	     	   REPLACE S00.F30 WITH S00_TK.F30
	     	   REPLACE S00.F31 WITH S00_TK.F31
	     	   REPLACE S00.F32 WITH S00_TK.F32
	     	   REPLACE S00.F33 WITH S00_TK.F33
	     	   REPLACE S00.F34 WITH S00_TK.F34
	     	   REPLACE S00.F35 WITH S00_TK.F35
	     	   REPLACE S00.F36 WITH S00_TK.F36
	     	   REPLACE S00.F37 WITH S00_TK.F37
	     	   REPLACE S00.F38 WITH S00_TK.F38
	     	   REPLACE S00.F58 WITH S00_TK.F58
	     	   REPLACE S00.F62 WITH S00_TK.F62
	     	   REPLACE S00.F68 WITH S00_TK.F68
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
	     	IF S00_TK.F01 = 'I'		&&班別設定
	     	   SELECT S00
	     	   SET ORDER TO S005	&& F01+F48
	     	   SEEK 'I' + S00_TK.F48
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'I'
	     	      REPLACE S00.F48 WITH S00_TK.F48
	     	   ENDIF	     	 
	     	   REPLACE S00.F49 WITH S00_TK.F49
	     	   REPLACE S00.F50 WITH S00_TK.F50
	     	   REPLACE S00.F51 WITH S00_TK.F51
	     	   REPLACE S00.F52 WITH S00_TK.F52
	     	   REPLACE S00.F53 WITH S00_TK.F53
	     	   REPLACE S00.F54 WITH S00_TK.F54
	     	   REPLACE S00.F55 WITH S00_TK.F55
	     	   REPLACE S00.F56 WITH S00_TK.F56
	     	   REPLACE S00.F57 WITH S00_TK.F57
	     	   REPLACE S00.F66 WITH S00_TK.F66
	     	   REPLACE S00.F67 WITH S00_TK.F67
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
	     	IF S00_TK.F01 = 'Y'		&&加班參數設定
	     	   SELECT S00
	     	   SEEK 'Y'
	     	   IF !FOUND()
	     	      SELECT S00
	     	      APPEND BLANK 
	     	      REPLACE S00.F01 WITH 'Y'
	     	   ENDIF
	     	   REPLACE S00.F63 WITH S00_TK.F63
	     	   REPLACE S00.F64 WITH S00_TK.F64
	     	   REPLACE S00.F65 WITH S00_TK.F65
	     	   REPLACE S00.F99 WITH S00_TK.F99
	     	ENDIF
*!*		     	IF S00_TK.F01 = 'H'		&&懲處扣款設定
*!*		     	ENDIF	     	
*!*		     	IF S00_TK.F01 = 'J'		&&其他所得發放設定
*!*		     	ENDIF     	
*!*		     	IF S00_TK.F01 = 'K'		&&年終獎金發放設定
*!*		     	ENDIF
*!*		     	IF S00_TK.F01 = 'L'		&&未休年假獎金發放設定
*!*		     	ENDIF	     	
*!*		     	IF S00_TK.F01 = 'Z'		&&基本參數設定
*!*		     	ENDIF
	     	SELECT S00_TK
	     	SKIP
	     ENDDO
	     SELECT S00_TK
	     USE
	     ***
         IF !USED('S01')		&&薪資資料設定
            SELECT 0
            USE S01
         ELSE
            SELECT S01
         ENDIF
         SET ORDER TO S011
	     ***帶入德通系統中S01.薪資資料設定有改變之資料
	     SELECT S01_TK
	     GO TOP
	     DO WHILE !EOF()
	     	SELECT S01
	     	SEEK S01_TK.F01 + S01_TK.F07
	     	IF !FOUND()
	     	   SELECT S01
	     	   APPEND BLANK
	     	   REPLACE S01.F01 WITH S01_TK.F01
	     	ENDIF
	     	REPLACE S01.F02 WITH S01_TK.F02
	     	REPLACE S01.F03 WITH S01_TK.F03
	     	REPLACE S01.F04 WITH S01_TK.F04
	     	REPLACE S01.F05 WITH S01_TK.F05
	     	REPLACE S01.F06 WITH S01_TK.F06
	     	REPLACE S01.F07 WITH S01_TK.F07
	     	REPLACE S01.F08 WITH S01_TK.F08
	     	REPLACE S01.F09 WITH S01_TK.F09
	     	REPLACE S01.F10 WITH S01_TK.F10
	     	REPLACE S01.F11 WITH S01_TK.F11
	     	REPLACE S01.F12 WITH S01_TK.F12
	     	REPLACE S01.F13 WITH S01_TK.F13
	     	REPLACE S01.F14 WITH S01_TK.F14
	     	REPLACE S01.F15 WITH S01_TK.F15
	     	REPLACE S01.F16 WITH S01_TK.F16
	     	REPLACE S01.F17 WITH S01_TK.F17
	     	REPLACE S01.F18 WITH S01_TK.F18
	     	REPLACE S01.F99 WITH S01_TK.F99
	     	SELECT S01_TK
	     	SKIP
	     ENDDO
	     SELECT S01_TK
	     USE
	     ***遲到設定參數
	     SELECT S00
	     SET ORDER TO S001
	     SEEK 'C'
	     IF FOUND()
	        LATE_F11 = S00.F11 	&&遲到分鐘數
		    LATE_F12 = S00.F12 	&&是否影響全勤
		    LATE_F13 = S00.F13 	&&是否有超過規定分鐘數需另扣金額
		    LATE_F14 = S00.F14 	&&超過分鐘數
		    LATE_F15 = S00.F15 	&&扣款底數
		    LATE_F16 = S00.F16 	&&扣款比例
		    LATE_F17 = S00.F17 	&&扣款金額
	     ELSE
		    LATE_F11 = 0
		    LATE_F12 = ''
		    LATE_F13 = ''
		    LATE_F14 = 0
		    LATE_F15 = ''
		    LATE_F16 = 0
		    LATE_F17 = 0
         ENDIF
         ***早退設定參數
         SELECT S00
         SEEK 'D'
         IF FOUND()
		    LEAVE_F11 = S00.F11 &&早退分鐘數
		    LEAVE_F12 = S00.F12 &&是否影響全勤
		    LEAVE_F13 = S00.F13 &&是否有超過規定分鐘數需另扣金額
		    LEAVE_F14 = S00.F14 &&超過分鐘數
		    LEAVE_F15 = S00.F15 &&扣款底數
		    LEAVE_F16 = S00.F16 &&扣款比例
		    LEAVE_F17 = S00.F17 &&扣款金額
	     ELSE
		    LEAVE_F11 = 0
		    LEAVE_F12 = ''
		    LEAVE_F13 = ''
		    LEAVE_F14 = 0
		    LEAVE_F15 = ''
		    LEAVE_F16 = 0
		    LEAVE_F17 = 0
         ENDIF
         IF !USED('S0B')  		&&年假記錄檔
            SELECT 0
            USE S0B
         ELSE
            SELECT S0B
         ENDIF
         SET ORDER TO S0B1   
         IF !USED('S02')  		&&所得稅表
            SELECT 0
            USE S02
         ELSE
            SELECT S02
         ENDIF     
         SET ORDER TO S021
         IF !USED('S03')  		&&健保費表
            SELECT 0
            USE S03
         ELSE
            SELECT S03
         ENDIF
         SET ORDER TO S031
         IF !USED('S04')  		&&勞保費表
            SELECT 0
            USE S04
         ELSE
            SELECT S04
         ENDIF
         SET ORDER TO S041
         IF !USED('S05')  		&&人事薪資基本檔
            SELECT 0
            USE S05 
         ELSE
            SELECT S05
         ENDIF
         SET ORDER TO S051
	     IF !USED('S06')		&&薪資項目及扣款項目
	        SELECT 0
	        USE S06 
		 ELSE
		    SELECT S06
		 ENDIF
		 SET ORDER TO S061
		 IF !USED('S07')		&&調薪修改作業
		   SELECT 0
		   USE S07
		 ELSE
		   SELECT S07
		 ENDIF
		 SET ORDER TO S071
		 IF !USED('S08')		&&懲處作業
		   SELECT 0
		   USE S08
		 ELSE
		   SELECT S08
		 ENDIF
		 SET ORDER TO S081
		 IF !USED('S09')		&&其他所得發放作業
		   SELECT 0
		   USE S09
		 ELSE
		   SELECT S09
		 ENDIF
		 SET ORDER TO S091
	     S101 ='S101' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
	     S102 = 'S102' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
	     S103 = 'S103' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
	     S104 = 'S104' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
	     S105 = 'S105' + LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6)
	     IF !USED('&S10')	&&薪資計算作業
	        SELECT 0
	        USE (S10) ALIAS S10
	     ELSE
	        SELECT S10
	     ENDIF
	     SET ORDER TO S101
	     SELECT S10
	     DELETE FROM S10 WHERE S10.F98 = 'N'   &&刪除未確認之人員
         *****
         WAIT WINDOW "資料匯入中,請稍後!" AT 0,150 NOWAIT NOCLEAR
         ***判斷所得稅表資料是否有變動
         IF FILE(ALLTRIM(INT_999) + '\S02.DBF') = .T.
	        SELECT S02_TK
	        GO TOP
	        DO WHILE !EOF()
	           SELECT S02
	           SEEK S02_TK.F01 + S02_TK.F02
	           IF !FOUND()
	              SELECT S02
	         	  APPEND BLANK
	           ENDIF
	           REPLACE S02.F01 WITH S02_TK.F01
	           REPLACE S02.F02 WITH S02_TK.F02
	           REPLACE S02.F03 WITH S02_TK.F03
	           REPLACE S02.F04 WITH S02_TK.F04
	           REPLACE S02.F05 WITH S02_TK.F05
	           REPLACE S02.F06 WITH S02_TK.F06
	           REPLACE S02.F07 WITH S02_TK.F07
	           REPLACE S02.F08 WITH S02_TK.F08
	           REPLACE S02.F09 WITH S02_TK.F09
	           REPLACE S02.F10 WITH S02_TK.F10
	           REPLACE S02.F11 WITH S02_TK.F11
	           REPLACE S02.F12 WITH S02_TK.F12
	           REPLACE S02.F13 WITH S02_TK.F13
	           REPLACE S02.F14 WITH S02_TK.F14
	           REPLACE S02.F99 WITH S02_TK.F99
	           SELECT S02_TK
	           SKIP
	        ENDDO
	        SELECT S02_TK
	        USE
	     ENDIF
	     SELECT S02
	     USE
         ***判斷勞保費表資料是否有變動
         IF FILE(ALLTRIM(INT_999) + '\S03.DBF') = .T.
            SELECT S03_TK
            GO TOP
            DO WHILE !EOF()
               SELECT S03
               SEEK S03_TK.F01
               IF !FOUND()
                  SELECT S03
                  APPEND BLANK 
               ENDIF
               REPLACE S03.F01 WITH S03_TK.F01
               REPLACE S03.F02 WITH S03_TK.F02
               REPLACE S03.F03 WITH S03_TK.F03
               REPLACE S03.F04 WITH S03_TK.F04
               REPLACE S03.F05 WITH S03_TK.F05
               REPLACE S03.F06 WITH S03_TK.F06
               REPLACE S03.F07 WITH S03_TK.F07
               REPLACE S03.F08 WITH S03_TK.F08
               REPLACE S03.F09 WITH S03_TK.F09
               REPLACE S03.F10 WITH S03_TK.F10
               REPLACE S03.F99 WITH S03_TK.F99
               SELECT S03_TK
               SKIP
            ENDDO
            SELECT S03_TK
            USE
         ENDIF
         SELECT S03
         USE
         ***判斷健保費表資料是否有變動
         IF FILE(ALLTRIM(INT_999) + '\S04.DBF') = .T.
            SELECT S04_TK
            GO TOP
            DO WHILE !EOF()
               SELECT S04
               SEEK S04_TK.F01
               IF !FOUND()
                  SELECT S04
                  APPEND BLANK
               ENDIF
               REPLACE S04.F01 WITH S04_TK.F01
               REPLACE S04.F02 WITH S04_TK.F02
               REPLACE S04.F03 WITH S04_TK.F03
               REPLACE S04.F04 WITH S04_TK.F04
               REPLACE S04.F05 WITH S04_TK.F05
               REPLACE S04.F06 WITH S04_TK.F06
               REPLACE S04.F99 WITH S04_TK.F99
               SELECT S04_TK
               SKIP
            ENDDO
            SELECT S04_TK
            USE
         ENDIF
         SELECT S04
         USE
         ***先判斷員工之勞保與健保投保資料是否有變動
         IF FILE(ALLTRIM(INT_999) + '\S05.DBF') = .T.
           SELECT S05_TK
           GO TOP
           DO WHILE !EOF()
              SELECT S05
              SET ORDER TO S051
              SEEK S05_TK.F01
              IF FOUND()
                 REPLACE S05.F04 WITH S05_TK.F04
                 REPLACE S05.F07 WITH S05_TK.F07
                 REPLACE S05.F08 WITH S05_TK.F08
                 REPLACE S05.F09 WITH S05_TK.F09
                 REPLACE S05.F13 WITH S05_TK.F13
                 REPLACE S05.F14 WITH S05_TK.F14
                 REPLACE S05.F16 WITH S05_TK.F16
                 REPLACE S05.F17 WITH S05_TK.F17
                 REPLACE S05.F18 WITH S05_TK.F18
                 REPLACE S05.F19 WITH S05_TK.F19
                 REPLACE S05.F20 WITH S05_TK.F20
                 REPLACE S05.F21 WITH S05_TK.F21
                 REPLACE S05.F22 WITH S05_TK.F22
                 REPLACE S05.F23 WITH S05_TK.F23
                 REPLACE S05.F24 WITH S05_TK.F24
                 REPLACE S05.F30 WITH S05_TK.F30
                 REPLACE S05.F31 WITH S05_TK.F31
                 REPLACE S05.F32 WITH S05_TK.F32
                 REPLACE S05.F33 WITH S05_TK.F33
              ELSE
                 SELECT S05
                 APPEND BLANK 
                 REPLACE S05.F01 WITH S05_TK.F01
                 REPLACE S05.F02 WITH S05_TK.F02
                 REPLACE S05.F03 WITH S05_TK.F03
                 REPLACE S05.F04 WITH S05_TK.F04
                 REPLACE S05.F05 WITH S05_TK.F05
                 REPLACE S05.F06 WITH S05_TK.F06
                 REPLACE S05.F07 WITH S05_TK.F07
                 REPLACE S05.F08 WITH S05_TK.F08
                 REPLACE S05.F09 WITH S05_TK.F09
                 REPLACE S05.F10 WITH S05_TK.F10
                 REPLACE S05.F11 WITH S05_TK.F11
                 REPLACE S05.F12 WITH S05_TK.F12
                 REPLACE S05.F13 WITH S05_TK.F13
                 REPLACE S05.F14 WITH S05_TK.F14
                 REPLACE S05.F15 WITH S05_TK.F15
                 REPLACE S05.F16 WITH S05_TK.F16
                 REPLACE S05.F17 WITH S05_TK.F17
                 REPLACE S05.F18 WITH S05_TK.F18
                 REPLACE S05.F19 WITH S05_TK.F19
                 REPLACE S05.F20 WITH S05_TK.F20
                 REPLACE S05.F21 WITH S05_TK.F21
                 REPLACE S05.F22 WITH S05_TK.F22
                 REPLACE S05.F23 WITH S05_TK.F23
                 REPLACE S05.F24 WITH S05_TK.F24
                 REPLACE S05.F25 WITH S05_TK.F25
                 REPLACE S05.F26 WITH S05_TK.F26
                 REPLACE S05.F27 WITH S05_TK.F27
                 REPLACE S05.F28 WITH S05_TK.F28
                 REPLACE S05.F29 WITH S05_TK.F29
                 REPLACE S05.F30 WITH S05_TK.F30
                 REPLACE S05.F31 WITH S05_TK.F31
                 REPLACE S05.F32 WITH S05_TK.F32
                 REPLACE S05.F33 WITH S05_TK.F33
                 REPLACE S05.F99 WITH S05_TK.F99
              ENDIF
              SELECT S05_TK
              SKIP
           ENDDO
           SELECT S05_TK
           USE
         ENDIF
         ***帶入德通系統之S0B.年假記錄檔
         IF FILE(ALLTRIM(INT_999) + '\S0B.DBF') = .T.
            SELECT S0B_TK
            GO TOP
            DO WHILE !EOF()
               SELECT S0B
               SEEK S0B_TK.F01 + S0B_TK.F03
               IF FOUND()
                  REPLACE S0B.F04 WITH S0B_TK.F04
                  REPLACE S0B.F05 WITH S0B_TK.F05
               ELSE
                  SELECT S0B
                  APPEND BLANK
                  REPLACE S0B.F01 WITH S0B_TK.F01
                  REPLACE S0B.F02 WITH S0B_TK.F02
                  REPLACE S0B.F03 WITH S0B_TK.F03
                  REPLACE S0B.F04 WITH S0B_TK.F04
                  REPLACE S0B.F05 WITH S0B_TK.F05
               ENDIF
               SELECT S0B_TK
               SKIP
            ENDDO
            SELECT S0B_TK
            USE
         ENDIF
         ***帶入德通系統之S06.薪資項目及扣款項目
         IF FILE(ALLTRIM(INT_999) + '\S06.DBF') = .T.
            SELECT S06_TK
            GO TOP
            DO WHILE !EOF()
               SELECT S06
               SET ORDER TO S061   && F01+F06+F07
               SEEK S06_TK.F01 + S06_TK.F06 + S06_TK.F07
               IF FOUND()
                  IF S06.F06 = 'J' AND S06.F08 <> S06_TK.F08
                     REPLACE S06.F08 WITH S06_TK.F08
                     REPLACE S06.F99 WITH S06_TK.F99
                  ENDIF
                  IF S06.F06 = 'K' AND S06.F08 <> S06_TK.F08
                     REPLACE S06.F08 WITH S06_TK.F08
                     REPLACE S06.F99 WITH S06_TK.F99
                  ENDIF
                  IF S06.F06 = 'O' AND S06.F08 <> S06_TK.F08
                     REPLACE S06.F08 WITH S06_TK.F08
                     REPLACE S06.F99 WITH S06_TK.F99
                  ENDIF
               ELSE
                  SELECT S06
                  APPEND BLANK
                  REPLACE S06.F01 WITH S06_TK.F01
                  REPLACE S06.F02 WITH S06_TK.F02
                  REPLACE S06.F03 WITH S06_TK.F03
                  REPLACE S06.F04 WITH S06_TK.F04
                  REPLACE S06.F05 WITH S06_TK.F05
                  REPLACE S06.F06 WITH S06_TK.F06
                  REPLACE S06.F07 WITH S06_TK.F07
                  REPLACE S06.F08 WITH S06_TK.F08
                  REPLACE S06.F09 WITH S06_TK.F09
                  REPLACE S06.F10 WITH S06_TK.F10
                  REPLACE S06.F99 WITH S06_TK.F99
               ENDIF
               SELECT S06_TK
               SKIP
            ENDDO
            SELECT S06_TK
            USE
         ENDIF
         ***帶入德通系統之S07.調薪修改作業資料
         IF FILE(ALLTRIM(INT_999) + '\S07.DBF') = .T.
           SELECT S07_TK
           GO TOP
           DO WHILE !EOF()
              SELECT S07
              SET ORDER TO S071   && F01+DTOS(F05)+F06+F08+F09+F15
              SEEK S07_TK.F01 + DTOS(S07_TK.F05) + S07_TK.F06 + S07_TK.F08 + S07_TK.F09 + S07_TK.F15
              IF !FOUND()
                 SELECT S07
                 APPEND BLANK
                 REPLACE S07.F01 WITH S07_TK.F01
                 REPLACE S07.F02 WITH S07_TK.F02
                 REPLACE S07.F03 WITH S07_TK.F03
                 REPLACE S07.F04 WITH S07_TK.F04
                 REPLACE S07.F05 WITH S07_TK.F05
                 REPLACE S07.F06 WITH S07_TK.F06
                 REPLACE S07.F07 WITH S07_TK.F07
                 REPLACE S07.F08 WITH S07_TK.F08
                 REPLACE S07.F09 WITH S07_TK.F09
                 REPLACE S07.F10 WITH S07_TK.F10
                 REPLACE S07.F11 WITH S07_TK.F11
                 REPLACE S07.F12 WITH S07_TK.F12
                 REPLACE S07.F13 WITH S07_TK.F13
                 REPLACE S07.F14 WITH S07_TK.F14
                 REPLACE S07.F15 WITH S07_TK.F15
                 REPLACE S07.F99 WITH S07_TK.F99
              ENDIF
              SELECT S07_TK
              SKIP
           ENDDO
           SELECT S07_TK
           USE
         ENDIF
         *****
         SELECT S14_TEMP
         GO TOP
         DO WHILE !EOF()
            SELECT S05
            SET ORDER TO S051
            SEEK S14_TEMP.F01
            IF FOUND()
               DO CASE 
                  CASE S05.F15 = 'M'   &&發薪總類為月薪
                       MONTH_MONEY = IIF(INDEXSEEK(S05.F01 + 'A',.T.,'S06','S061'),S06.F08,0)   	 					&&底薪
                       DAY_MONEY = IIF(IIF(INDEXSEEK('Z',.T.,'S00','S001'),S00.F02,0) <> 0,MONTH_MONEY / S00.F02,0)    	&&日薪
                       TIME_MONEY = IIF(IIF(INDEXSEEK('Z',.T.,'S00','S001'),S00.F03,0) <> 0,DAY_MONEY / S00.F03,0)     	&&時薪
                  CASE S05.F15 = 'D'   &&發薪總類為日薪
                       MONTH_MONEY = 0
                       DAY_MONEY = IIF(INDEXSEEK(S05.F01 + 'B',.T.,'S06','S061'),S06.F08,0)
                       TIME_MONEY = IIF(IIF(INDEXSEEK('Z',.T.,'S00','S001'),S00.F03,0) <> 0,DAY_MONEY / S00.F03,0)
                  CASE S05.F15 = 'T'   &&發薪總類為時薪
                       MONTH_MONEY = 0
                       DAY_MONEY = 0
                       TIME_MONEY = IIF(INDEXSEEK(S05.F01 + 'C',.T.,'S06','S061'),S06.F08,0)
               ENDCASE
               *****A.基本考勤
               S00Z_F02 = IIF(INDEXSEEK('Z',.T.,'S00','S001'),S00.F02,MONTH(CTOD(ALLTRIM(THISFORM.TXTF02.VALUE) + '/01'))) 	&&每月工作日數
               SELECT S10
               SET ORDER TO S101
               SEEK S14_TEMP.F01 + 'A' && F01+F97+F12
               IF !FOUND()
                  SELECT S10
                  APPEND BLANK
                  REPLACE S10.F01 WITH S05.F01
                  REPLACE S10.F02 WITH S05.F02
                  REPLACE S10.F03 WITH S05.F05
                  REPLACE S10.F04 WITH S05.F06
   	 	          IF (RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2) = '02' AND S14_TEMP.F07 < 28) OR (RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2) <> '02' AND S14_TEMP.F07 < S00Z_F02)
   	 		         REPLACE S10.F05 WITH S14_TEMP.F07
   	 	          ELSE
   	 		         REPLACE S10.F05 WITH S00Z_F02
   	 	          ENDIF
   	 	          S10_F05 = S10.F05
                  REPLACE S10.F06 WITH S14_TEMP.F05
                  REPLACE S10.F07 WITH S14_TEMP.F06
		          REPLACE S10.F97 WITH 'A'
		          REPLACE S10.F98 WITH 'N'
		          REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
		          ****判斷全勤扣款
		          IF S10.F06 >= LATE_F11 && AND LATE_F11<>0 && 判斷是否達到遲到標準
		             IF LATE_F12 = 'Y'  	          		&&是否影響全勤(有無全勤)
		                REPLACE S10.F08 WITH 'Y'
		             ELSE
		                IF S10.F07 >= LEAVE_F11 			&& 判斷是否達到早退標準
		                   REPLACE S10.F08 WITH IIF(LEAVE_F12 = 'Y','Y','N')    &&是否影響全勤(有無全勤)
		                ELSE
		                   REPLACE S10.F08 WITH 'N'
		                ENDIF
		             ENDIF
		          ELSE 		&&若沒達到遲到標準則在判斷是否達到早退標準
		             IF S10.F07 >= LEAVE_F11 && AND LEAVE_F11<>0				&& 判斷是否達到早退標準
		                REPLACE S10.F08 WITH IIF(LEAVE_F12 = 'Y','Y','N')       &&是否影響全勤(有無全勤)
		             ELSE
		                REPLACE S10.F08 WITH 'N'
		             ENDIF
		          ENDIF
		          ****判斷額外扣款(遲到)
		          IF S10.F06 >= LATE_F14 AND LATE_F13 = 'Y'  &&超過等於額外扣款分鐘數 AND 是否需要額外扣款
		             DO CASE
		                CASE LATE_F15 = 'T'  		&&額外扣款底數為時薪
		                     REPLACE S10.F09 WITH ROUND(TIME_MONEY * (LATE_F16/100),0)
		                CASE LATE_F15 = 'D'  		&&額外扣款底數為日薪
		                     REPLACE S10.F09 WITH ROUND(DAY_MONEY * (LATE_F16/100),0)
		                CASE LATE_F15 = 'B'  		&&額外扣款底數為底薪
		                     REPLACE S10.F09 WITH ROUND(MONTH_MONEY * (LATE_F16/100),0)
		                CASE EMPTY(LATE_F15)        &&額外扣款為金額
		                     REPLACE S10.F09 WITH LATE_F17
		             ENDCASE
		          ENDIF
		          ****判斷額外扣款(早退)
		          IF S10.F07 >= LEAVE_F14 AND LEAVE_F13 = 'Y'  &&超過等於額外扣款分鐘數 AND 是否需要額外扣款
		             DO CASE
		                CASE LEAVE_F15 = 'T'  		&&額外扣款底數為時薪
		                     REPLACE S10.F10 WITH ROUND(TIME_MONEY * (LATE_F16/100),0)
		                CASE LEAVE_F15 = 'D'  		&&額外扣款底數為日薪
		                     REPLACE S10.F10 WITH ROUND(DAY_MONEY * (LATE_F16/100),0)
		                CASE LEAVE_F15 = 'B'  		&&額外扣款底數為底薪
		                     REPLACE S10.F10 WITH ROUND(MONTH_MONEY * (LATE_F16/100),0)
		                CASE EMPTY(LEAVE_F15)       &&額外扣款為金額
		                     REPLACE S10.F10 WITH LEAVE_F17
		             ENDCASE
		          ENDIF
               ENDIF
               *** 
               SELECT S10
               SET ORDER TO S101
               SEEK S14_TEMP.F01 + 'A' && F01+F97+F12
               IF FOUND() AND S10.F98 <> 'Y'
                  S10_F08A = S10.F08 
                  **********************F.加班明細
                  ********平日加班
                  IF EMPTY(S14_TEMP.F09) = .F.
                     SELECT S10
                     APPEND BLANK 
                     REPLACE S10.F01 WITH S05.F01
                     REPLACE S10.F02 WITH S05.F02
                     REPLACE S10.F03 WITH S05.F05
                     REPLACE S10.F04 WITH S05.F06
                     REPLACE S10.F33 WITH '1'
                     REPLACE S10.F34 WITH IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F07,'')
                     REPLACE S10.F35 WITH S14_TEMP.F09
                     REPLACE S10.F36 WITH IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F46,'')
                     REPLACE S10.F37 WITH IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F47,0)
                     REPLACE S10.F44 WITH S14_TEMP.F12
                     DO CASE
		                CASE S10.F34 = 'H'  &&時薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F09,0) * TIME_MONEY),0)		&&補助金額+(加班數值*時薪)
		                CASE S10.F34 = 'D'  &&日薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F09,0) * DAY_MONEY),0)		&&補助金額+(加班數值*日薪)
		                CASE S10.F34 = 'B'  &&底薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F09,0) * MONTH_MONEY),0)	&&補助金額+(加班數值*月薪)
		                OTHERWISE        	&&給予金額
		                     REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P1',.T.,'S00','S001'),S00.F10,0)),0)					&&補助金額+(加班數值*給予金額)
		             ENDCASE
		             REPLACE S10.F39 WITH  S10.F38 - (S10.F37 * S10.F44)	&&加班小計=加班總所得(包括津貼金額)-津貼金額
		             REPLACE S10.F97 WITH 'F'
		             REPLACE S10.F98 WITH 'N'
		             REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
                  ENDIF
                  ********假日加班
                  IF EMPTY(S14_TEMP.F10) = .F.
                     SELECT S10
                     APPEND BLANK
                     REPLACE S10.F01 WITH S05.F01
                     REPLACE S10.F02 WITH S05.F02
                     REPLACE S10.F03 WITH S05.F05
                     REPLACE S10.F04 WITH S05.F06
                     REPLACE S10.F33 WITH '2'
                     REPLACE S10.F34 WITH IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F07,'')
                     REPLACE S10.F35 WITH S14_TEMP.F10
                     REPLACE S10.F36 WITH IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F46,'')
                     REPLACE S10.F37 WITH IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F47,0)
                     REPLACE S10.F44 WITH S14_TEMP.F13
                     DO CASE
		                CASE S10.F34 = 'H'  &&時薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F09,0) * TIME_MONEY),0)		&&補助金額+(加班數值*時薪)
		                CASE S10.F34 = 'D'  &&日薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F09,0) * DAY_MONEY),0)		&&補助金額+(加班數值*日薪)
		                CASE S10.F34 = 'B'  &&底薪
		       	             REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F09,0) * MONTH_MONEY),0)	&&補助金額+(加班數值*月薪)
		                OTHERWISE        	&&給予金額
		                     REPLACE S10.F38 WITH ROUND((S10.F37 * S10.F44) + (S10.F35 * IIF(INDEXSEEK('B' + 'P2',.T.,'S00','S001'),S00.F10,0)),0)					&&補助金額+(加班數值*給予金額)
		             ENDCASE
		             REPLACE S10.F39 WITH  S10.F38 - (S10.F37 * S10.F44)	&&加班小計=加班總所得(包括津貼金額)-津貼金額
		             REPLACE S10.F97 WITH 'F'
		             REPLACE S10.F98 WITH 'N'
		             REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
                  ENDIF
                  ******************C.薪資所得
                  SELECT S06.* FROM S06 WHERE S06.F05='1' AND ALLTRIM(S06.F01) = ALLTRIM(S05.F01) ORDER BY S06.F06 INTO CURSOR S06_TEMP
		          SELECT S06_TEMP
		          GO TOP
		          DO WHILE !EOF()
                     SELECT S10
                     APPEND BLANK 
                     REPLACE S10.F01 WITH S05.F01
                     REPLACE S10.F02 WITH S05.F02
                     REPLACE S10.F03 WITH S05.F05
                     REPLACE S10.F04 WITH S05.F06
    			     REPLACE S10.F18 WITH S06_TEMP.F06
    			     REPLACE S10.F19 WITH S06_TEMP.F07
    			     REPLACE S10.F20 WITH IIF(S10_F08A = 'Y' AND S06_TEMP.F06 = 'F',0,S06_TEMP.F08)
    			     DO CASE
    			        CASE S10_F08A ='Y' AND S06_TEMP.F06 = 'F'
    			             REPLACE S10.F20 WITH 0 
    			        CASE S10_F08A <> 'Y' AND S06_TEMP.F06 = 'F'
    			             IF S10_F05 < S00Z_F02
    			                REPLACE S10.F20 WITH ROUND((S06_TEMP.F08/S00Z_F02) * S10_F05,0)
    			             ELSE
    			                REPLACE S10.F20 WITH S06_TEMP.F08
    			             ENDIF
    			        CASE S06_TEMP.F06 = 'A'
	                         IF S10_F05 < S00Z_F02
	                            DO CASE
	                               CASE S05.F15 = 'M'
	                                    REPLACE S10.F20 WITH DAY_MONEY * S10_F05
	                               CASE S05.F15 = 'D'
	                                    REPLACE S10.F20 WITH DAY_MONEY * S10_F05
	                               CASE S05.F15 = 'T'
	                                    REPLACE S10.F20 WITH TIME_MONEY * S10_F05
	                            ENDCASE
	                         ELSE
	                            REPLACE S10.F20 WITH S06_TEMP.F08
	                         ENDIF
    			        CASE S06_TEMP.F06 = 'E' OR S06_TEMP.F06 = 'D'
		                     IF S10_F05 < S00Z_F02
		                        REPLACE S10.F20 WITH ROUND((S06_TEMP.F08/S00Z_F02) * S10_F05,0)
		                     ELSE
		                        REPLACE S10.F20 WITH S06_TEMP.F08
		                     ENDIF
		                OTHERWISE
		                     REPLACE S10.F20 WITH S06_TEMP.F08
		             ENDCASE
    			     REPLACE S10.F21 WITH S06_TEMP.F08
		             REPLACE S10.F97 WITH 'C'
		             REPLACE S10.F98 WITH 'N'
		             REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
		             SELECT S06_TEMP
		             SKIP
		          ENDDO
		          SELECT S06_TEMP
		          USE
		          *****帶入此人員之其他所得發放(此發放需為合併該月份發放之所得)
		          SELECT S09.* FROM S09 WHERE ALLTRIM(S09.F01) = ALLTRIM(S05.F01) AND S09.F07 <> '1' ;
			             AND LEFT(DTOS(S09.F08),6) = LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE+'/01')),6) AND !EMPTY(S09.F08);
			             AND S09.F15 = 'Y' AND (S09.F16 = '1' OR S09.F16 = '3') ORDER BY S09.F05 INTO CURSOR S09_TEMP
		          SELECT S09_TEMP
		          GO TOP
		          DO WHILE !EOF()
                     SELECT S10
                     APPEND BLANK
                     REPLACE S10.F01 WITH S05.F01
                     REPLACE S10.F02 WITH S05.F02
                     REPLACE S10.F03 WITH S05.F05
                     REPLACE S10.F04 WITH S05.F06
			         REPLACE S10.F18 WITH 'I'
    			     REPLACE S10.F19 WITH S09_TEMP.F06
			         REPLACE S10.F20 WITH S09_TEMP.F13
			         REPLACE S10.F22 WITH S09_TEMP.F05
	                 REPLACE S10.F97 WITH 'C'
	                 REPLACE S10.F98 WITH 'N'
	                 REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	                 ********帶入扣稅金額
	                 IF (S09_TEMP.F12) <> 0		&&稅率不為零則有扣稅金額
                        SELECT S10
                        APPEND BLANK
                        REPLACE S10.F01 WITH S05.F01
                        REPLACE S10.F02 WITH S05.F02
                        REPLACE S10.F03 WITH S05.F05
                        REPLACE S10.F04 WITH S05.F06
			            REPLACE S10.F24 WITH 'N'
			            REPLACE S10.F25 WITH ALLTRIM(S09_TEMP.F06) + '扣稅額'
		                REPLACE S10.F26 WITH S09_TEMP.F13 - S09_TEMP.F17
			            REPLACE S10.F97 WITH 'D'
			            REPLACE S10.F98 WITH 'N'
			            REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	                 ENDIF
	                 SELECT S09_TEMP
	                 SKIP
		          ENDDO
		          SELECT S09_TEMP
		          USE
		          *****帶入此人員之特殊其他所得發放(此發放需為合併該月份發放之所得)
		          SELECT S12.* FROM S12 WHERE ALLTRIM(S12.F01) = ALLTRIM(S05.F01) AND S12.F08 <> '1' ;
			             AND LEFT(DTOS(S12.F09),6) = LEFT(DTOS(CTOD(THISFORM.TXTF02.VALUE + '/01')),6) AND !EMPTY(S12.F09);
			             AND S12.F19 = 'Y' AND (S12.F20 = '1' OR S12.F20 = '3') ORDER BY S12.F05 INTO CURSOR S12_TEMP
		          SELECT S12_TEMP
		          GO TOP
		          DO WHILE !EOF()	
                     SELECT S10
                     APPEND BLANK 
                     REPLACE S10.F01 WITH S05.F01
                     REPLACE S10.F02 WITH S05.F02
                     REPLACE S10.F03 WITH S05.F05
                     REPLACE S10.F04 WITH S05.F06
			         REPLACE S10.F18 WITH 'I'
    			     REPLACE S10.F19 WITH IIF(S12_TEMP.F06 = '1','年終獎金發放','未休年假獎金發放')
			         REPLACE S10.F20 WITH S12_TEMP.F16
			         REPLACE S10.F22 WITH S12_TEMP.F05
	                 REPLACE S10.F97 WITH 'C'
	                 REPLACE S10.F98 WITH 'N'
	                 REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	                 ********帶入扣稅金額
	                 IF (S12_TEMP.F15) <> 0		&&稅率不為零則有扣稅金額
                        SELECT S10
                        APPEND BLANK
                        REPLACE S10.F01 WITH S05.F01
                        REPLACE S10.F02 WITH S05.F02
                        REPLACE S10.F03 WITH S05.F05
                        REPLACE S10.F04 WITH S05.F06
			            REPLACE S10.F24 WITH 'N'
			            REPLACE S10.F25 WITH IIF(S12_TEMP.F06 = '1','年終獎金發放','未休年假獎金發放') + '扣稅額'
		                REPLACE S10.F26 WITH S12_TEMP.F16 - S12_TEMP.F17
			            REPLACE S10.F97 WITH 'D'
			            REPLACE S10.F98 WITH 'N'
			            REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	                 ENDIF
	                 SELECT S12_TEMP
	                 SKIP
		          ENDDO
		          SELECT S12_TEMP
		          USE
		          ******************D.薪資扣款
				  SELECT S06.* FROM S06 WHERE S06.F05 = '2' AND ALLTRIM(S06.F01) = ALLTRIM(S05.F01) ORDER BY S06.F06 INTO CURSOR S06_TEMP
				  SELECT S06_TEMP
				  GO TOP
		      	  DO WHILE !EOF()
			          SELECT S10
			          APPEND BLANK
                      REPLACE S10.F01 WITH S05.F01
                      REPLACE S10.F02 WITH S05.F02
                      REPLACE S10.F03 WITH S05.F05
                      REPLACE S10.F04 WITH S05.F06
                      REPLACE S10.F24 WITH S06_TEMP.F06
    			      REPLACE S10.F25 WITH S06_TEMP.F07
    			      IF S06_TEMP.F06 = 'J'  &&勞保費需照天數計算
				         REPLACE S10.F26 WITH ROUND(S05.F18 * ((IIF(INDEXSEEK('C',.T.,'S01','S011'),S01.F02,0) + IIF(INDEXSEEK('C',.T.,'S01','S011'),S01.F06,0))/100) * (IIF(INDEXSEEK('C',.T.,'S01','S011'),S01.F03,0)/100)/30 * IIF(S10_F05 > 30,30,S10_F05),0)
    			      ELSE
    			         REPLACE S10.F26 WITH S06_TEMP.F08
    			      ENDIF
			          REPLACE S10.F23 WITH S06_TEMP.F08
			          REPLACE S10.F97 WITH 'D'
			          REPLACE S10.F98 WITH 'N'
			          REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
                      SELECT S06_TEMP
                      SKIP
                  ENDDO
                  SELECT S06_TEMP
                  USE
                  ******************B.缺勤項目
                  SELECT S0AA.* FROM S0AA WHERE S0AA.F01 = S05.F01 ORDER BY S0AA.F05 INTO CURSOR S0A_TEMP
                  ***先刪除已存在S0A之資料
    			  SELECT S0A_TEMP
    			  GO TOP
    		  	  DO WHILE !EOF()
    		  	     SELECT S0A
*!*	    		  	     SET ORDER TO S0A3   && f01+f05+LEFT(DTOS(F08),6)+f09+f10
*!*	    		  	     SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
					 SET ORDER TO S0A1 && F01+F05+DTOS(F08)+F09+F10
					 SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10
    		  	     IF FOUND()
*!*	    		  	        IF S0A_TEMP.F05 = 'A1'
*!*	    		  	           SELECT S05
*!*					           REPLACE S05.F12 WITH S05.F12-S0A_TEMP.F07
*!*	    		  	        ENDIF
    		  	        SELECT S0A
    		  	        DELETE
    		  	     ELSE
    		  	        SELECT S0A
*!*	    		  	        SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+'SB'+SPACE(8)
    		  	        SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + 'SB' + SPACE(8)
    		  	        IF FOUND()
*!*	    		  	           IF S0A_TEMP.F05 = 'A1'
*!*	    		  	              SELECT S05
*!*					              REPLACE S05.F12 WITH S05.F12-S0A_TEMP.F07
*!*	    		  	           ENDIF
    		  	           REPLACE S0A.F07 WITH S0A.F07 - S0A_TEMP.F07
    		  	           IF S0A.F07 = 0
    		  	              SELECT S0A
    		  	              DELETE
    		  	           ENDIF
    		  	        ENDIF
    		  	     ENDIF
    		  	     SELECT S0A_TEMP
    		  	     SKIP
    		  	  ENDDO
    		  	  ******
    			  SELECT S0A_TEMP
    			  GO TOP
    		  	  DO WHILE !EOF()
    		  	     SELECT S10
    		  	     SET ORDER TO S101	&& F01+F97+F12+F42
    		  	     SEEK S05.F01 + 'B' + S0A_TEMP.F05 + S0A_TEMP.F10
    		  	     IF !FOUND()
    			        SELECT S10
    			        SEEK S05.F01 + 'B' + S0A_TEMP.F05
    			        IF FOUND()
    			           SELECT S10
			               REPLACE S10.F14 WITH S10.F14 + S0A_TEMP.F07
			               REPLACE S10.F42 WITH 'SB'
    			        ELSE
    			           SELECT S10
    			           APPEND BLANK
    			           REPLACE S10.F01 WITH S05.F01
                           REPLACE S10.F02 WITH S05.F02
                           REPLACE S10.F03 WITH S05.F05
                           REPLACE S10.F04 WITH S05.F06
                           REPLACE S10.F12 WITH S0A_TEMP.F05
			               REPLACE S10.F13 WITH S0A_TEMP.F06
			               REPLACE S10.F14 WITH S0A_TEMP.F07
			               REPLACE S10.F15 WITH S0A_TEMP.F09
			               REPLACE S10.F42 WITH S0A_TEMP.F10
    			           REPLACE S10.F97 WITH 'B'
    			           REPLACE S10.F98 WITH 'N'
    			        ENDIF
    			     ENDIF
    			     REPLACE S10.F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
    			     SELECT S0A
*!*	    			     SET ORDER TO S0A3   && f01+f05+LEFT(DTOS(F08),6)+f09+f10
*!*	    			     SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
    			     SET ORDER TO S0A1 && F01+F05+DTOS(F08)+F09+F100
    			     SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10
    			     IF !FOUND()
    			        SELECT S0A
*!*	    			        SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
				        SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10
    			        IF FOUND()
    			           SELECT S0A
    			           REPLACE S0A.F07 WITH S0A.F07 + S0A_TEMP.F07
    			           REPLACE S0A.F10 WITH 'SB'
    			        ELSE
    			           SELECT S0A
			               APPEND BLANK 
			               REPLACE S0A.F01 WITH S0A_TEMP.F01
			               REPLACE S0A.F02 WITH S0A_TEMP.F02
			               REPLACE S0A.F03 WITH S0A_TEMP.F03
			               REPLACE S0A.F04 WITH S0A_TEMP.F04
			               REPLACE S0A.F05 WITH S0A_TEMP.F05
			               REPLACE S0A.F06 WITH S0A_TEMP.F06
			               REPLACE S0A.F07 WITH S0A_TEMP.F07
			               REPLACE S0A.F08 WITH S0A_TEMP.F08
			               REPLACE S0A.F09 WITH S0A_TEMP.F09
			               REPLACE S0A.F10 WITH S0A_TEMP.F10
			               REPLACE S0A.F99 WITH S0A_TEMP.F99
    			        ENDIF
*!*	    			        IF S0A_TEMP.F05 = 'A1'
*!*					           SELECT S05
*!*					           REPLACE S05.F12 WITH S05.F12 + S0A_TEMP.F07
*!*	    			        ENDIF
    			     ENDIF
    			     SELECT S0A_TEMP
    			     SKIP
    			  ENDDO
    			  SELECT S0A_TEMP
    			  USE
    			  ***B.計算扣款
    			  SELECT S10.F12 FROM S10 WHERE S10.F01 = S05.F01 AND !EMPTY(S10.F12) ORDER BY S10.F12 INTO CURSOR S10BB
    			  SELECT S10BB
    			  GO TOP
    			  DO WHILE !EOF()
			         SELECT S10
			         SET ORDER TO S101   && F01+F97+F12
			         SEEK S05.F01 + 'B' + S10BB.F12
			         IF FOUND()
			            SELECT S00
			            SET ORDER TO S003
			            SEEK 'G' + S10.F12
			            IF FOUND()
			               DO CASE 		&&判斷扣款比例或金額
				              CASE S00.F30 = 'H'		&&底數為時薪
				                   REPLACE S10.F16 WITH ROUND(TIME_MONEY * (S00.F31/100) * S10.F14,0)
				              CASE S00.F30 = 'D'		&&底數為日薪
			   	                   REPLACE S10.F16 WITH ROUND(DAY_MONEY * (S00.F31/100) * S10.F14,0)
				              CASE S00.F30 = 'B'		&&底數為底薪
				     	           REPLACE S10.F16 WITH ROUND(MONTH_MONEY * (S00.F31/100) * S10.F14,0)
				              CASE EMPTY(S00.F30)	
				   	               REPLACE S10.F16 WITH ROUND(S00.F32 * S10.F14,0)	
		                   ENDCASE  
				           ****是否扣全勤
			               IF S00.F33 = 'Y'   &&判斷是否扣全勤
			                  NON_LEAVE_DAY = IIF(INDEXSEEK(S05.F01 + 'F',.T.,'S06','S061'),S06.F08,0) &&全勤獎金
			                  IF S10.F14 > S00.F34          &&不扣全勤緩衝數
				                 DO CASE
					                CASE S10.F14 = S00.F35	&&全勤扣款方式一,扣當月所得之全勤金額比例
			         	                 REPLACE S10.F17 WITH IIF(NON_LEAVE_DAY <> 0,ROUND(NON_LEAVE_DAY * (S00.F36/100),0),0)
			         	            CASE S10.F14 = S00.F37	&&全勤扣款方式二,扣當月所得之全勤金額比例
			         	                 REPLACE S10.F17 WITH IIF(NON_LEAVE_DAY <> 0,ROUND(NON_LEAVE_DAY * (S00.F38/100),0),0)
			         	            OTHERWISE
			         	                 REPLACE S10.F17 WITH NON_LEAVE_DAY
			         	         ENDCASE
			         	      ELSE
			         	         REPLACE S10.F17 WITH 0
			                  ENDIF
			               ELSE
			                  REPLACE S10.F17 WITH 0
			               ENDIF
    			        ENDIF
    			        ***若此缺勤項目需扣全勤則扣除
			            S10_F17 = S10.F17
			            SELECT S10
			            SET ORDER TO S102	&& F01+F97+F18+F19+DTOS(F22)
			            SEEK S05.F01 + 'C' + 'F'
			            IF FOUND()
			               REPLACE S10.F20 WITH IIF((S10.F20 - S10_F17) <= 0,0,S10.F20 - S10_F17)
			            ENDIF
			            ***若有扣除全勤則F08是否有扣除全勤要設定成Y
			            IF S10.F20 = 0
			               SELECT S10
			               SEEK S05.F01 + 'A'
			               IF FOUND()
			                  REPLACE S10.F08 WITH 'Y'
			               ENDIF
    			        ENDIF
    			     ENDIF
    			     SELECT S10BB
    			     SKIP
    			  ENDDO
    			  SELECT S10BB
    			  USE
			      ********E.懲處項目
				  SELECT S08.* FROM S08 WHERE S08.F15 = 'Y' AND S08.F12 <> 0 AND ALLTRIM(S08.F01) = ALLTRIM(S05.F01) ORDER BY S08.F05 INTO CURSOR S08_TEMP
				  SELECT S08_TEMP
				  GO TOP
				  DO WHILE !EOF()
				     SELECT S10
				     APPEND BLANK
	                 REPLACE S10.F01 WITH S05.F01
	                 REPLACE S10.F02 WITH S05.F02
	                 REPLACE S10.F03 WITH S05.F05
	                 REPLACE S10.F04 WITH S05.F06
			         REPLACE S10.F27 WITH S08_TEMP.F05
			         REPLACE S10.F28 WITH S08_TEMP.F06
			         REPLACE S10.F29 WITH S08_TEMP.F07
			         DO CASE 
			            CASE S08_TEMP.F08 = 'H'  &&時薪
			          		 REPLACE S10.F30 WITH ROUND(TIME_MONEY * (S08_TEMP.F09/100),0)
			            CASE S08_TEMP.F08 = 'D'  &&日薪
			                 REPLACE S10.F30 WITH ROUND(DAY_MONEY * (S08_TEMP.F09/100),0)
			            CASE S08_TEMP.F08 = 'B'  &&底薪
			                 REPLACE S10.F30 WITH ROUND(MONTH_MONEY * (S08_TEMP.F09/100),0)
			            CASE S08_TEMP.F08 = 'T'  &&實領薪資
			                 REPLACE S10.F30 WITH 0
			            CASE EMPTY(S08_TEMP.F08)  &&扣款金額
						     REPLACE S10.F30 WITH S08_TEMP.F10
			         ENDCASE
			         REPLACE S10.F31 WITH S08_TEMP.F13
			         REPLACE S10.F40 WITH S08_TEMP.F08
			         REPLACE S10.F41 WITH S08_TEMP.F09
    			     REPLACE S10.F97 WITH 'E'
    			     REPLACE S10.F98 WITH 'N'
    			     REPLACE S10.F99 WITH DTOC(DATE()) + ' ' + IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
                     SELECT S08_TEMP
                     SKIP
			      ENDDO
			      SELECT S08_TEMP
			      USE
			   ********
               ENDIF
            ENDIF 
            SELECT S14_TEMP
            SKIP
         ENDDO
         ******************************************
         *****帶入無需刷卡且S10未確認之人員出勤資料
         SELECT F01,F02 FROM S10 WHERE F98 <> 'Y' GROUP BY F01 ORDER BY F01 INTO CURSOR S10ABC
         SELECT S05.F01,S05.F02 FROM S05 WHERE S10ABC.F01 = S05.F01 OR S05.F30='2' INTO CURSOR S05ABC
         SELECT S10ABC
         USE
         SELECT S05ABC
         GO TOP
         DO WHILE !EOF()
            SELECT S05
            SET ORDER TO S051
            SEEK S05ABC.F01
            IF FOUND() AND S05.F30 = '2'
               SELECT S0AA.* FROM S0AA WHERE  S0AA.F01=S05.F01 ORDER BY S0AA.F05 INTO CURSOR S0A_TEMP
               ***先刪除已存在S0A之資料
    		   SELECT S0A_TEMP
    		   GO TOP
    		   DO WHILE !EOF()
    		      SELECT S0A
*!*	    		      SET ORDER TO S0A3   && F01+F05+LEFT(DTOS(F08),6)+F09+F10
*!*	    		      SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
    		      SET ORDER TO S0A1   && F01+F05+DTOS(F08)+F09+F10
    		      SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10  
    		      IF FOUND()
*!*	    		         IF S0A_TEMP.F05 = 'A1'
*!*	    		            SELECT S05
*!*				            REPLACE S05.F12 WITH S05.F12 - S0A_TEMP.F07
*!*	    		 	     ENDIF
    		  	     SELECT S0A
    		  	     DELETE
    		  	  ELSE
    		  	     SELECT S0A
*!*	    		  	     SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+'SB'+SPACE(8)
    		  	     SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + 'SB' + SPACE(8)
    		  	     IF FOUND() 
*!*	    		  	        IF S0A_TEMP.F05 = 'A1'
*!*	    		  	           SELECT S05
*!*					           REPLACE S05.F12 WITH S05.F12-S0A_TEMP.F07
*!*	    		  	        ENDIF
    		  	        REPLACE S0A.F07 WITH S0A.F07 - S0A_TEMP.F07
    		  	        IF S0A.F07 = 0
    		  	           SELECT S0A
    		  	           DELETE
    		  	        ENDIF
    		  	     ENDIF
    		  	  ENDIF
    		  	  SELECT S0A_TEMP
    		  	  SKIP
    		   ENDDO
    		   ****** 再加入S0A
    		   SELECT S0A_TEMP
    		   GO TOP
    		   DO WHILE !EOF()
    			  SELECT S0A
*!*	    			  SET ORDER TO S0A3   && F01+F05+LEFT(DTOS(F08),6)+F09+F10
*!*	    			  SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
    			  SET ORDER TO S0A1   && F01+F05+DTOS(F08)+F09+F10
    			  SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10
    			  IF !FOUND()
    			     SELECT S0A
*!*	    			     SEEK S0A_TEMP.F01+S0A_TEMP.F05+LEFT(DTOS(S0A_TEMP.F08),6)+S0A_TEMP.F09+S0A_TEMP.F10
			         SEEK S0A_TEMP.F01 + S0A_TEMP.F05 + DTOS(S0A_TEMP.F08) + S0A_TEMP.F09 + S0A_TEMP.F10
    			     IF FOUND()
    			        SELECT S0A
    			        REPLACE S0A.F07 WITH S0A.F07 + S0A_TEMP.F07
    			        REPLACE S0A.F10 WITH 'SB'
    			     ELSE
    			        SELECT S0A
    			        APPEND BLANK
    			        REPLACE S0A.F01 WITH S0A_TEMP.F01
    			        REPLACE S0A.F02 WITH S0A_TEMP.F02
    			        REPLACE S0A.F03 WITH S0A_TEMP.F03
    			        REPLACE S0A.F04 WITH S0A_TEMP.F04
    			        REPLACE S0A.F05 WITH S0A_TEMP.F05
    			        REPLACE S0A.F06 WITH S0A_TEMP.F06
    			        REPLACE S0A.F07 WITH S0A_TEMP.F07
    			        REPLACE S0A.F08 WITH S0A_TEMP.F08
    			        REPLACE S0A.F09 WITH S0A_TEMP.F09
    			        REPLACE S0A.F10 WITH S0A_TEMP.F10
    			        REPLACE S0A.F99 WITH S0A_TEMP.F99
    			     ENDIF
*!*	    			     IF S0A_TEMP.F05 = 'A1'
*!*				            SELECT S05
*!*					        REPLACE S05.F12 WITH S05.F12+S0A_TEMP.F07
*!*	    			     ENDIF
    			  ENDIF
    			  SELECT S0A_TEMP
    			  SKIP
    		   ENDDO
    		   SELECT S0A_TEMP
    		   USE
            ENDIF
            SELECT S05ABC
            SKIP
         ENDDO
         SELECT S05ABC
         USE
         ******
         WAIT CLEAR
         =MESSAGEBOX('出勤資料已成功轉入S10.薪資計算作業中,請至該作業中查詢!',0+64,'提示視窗')
	     THISFORM.TXTF02.VALUE = ''
      ENDIF
      THISFORM.CMND2.SETFOCUS
ENDPROC
********離開鍵***********
 PROCEDURE CMND2.CLICK
      CLOSE TABLE ALL
      THISFORM.RELEASE
      RETURN
 ENDPROC
 ***
 PROCEDURE CMND1.MOUSEENTER
      LPARAMETERS nButton, nShift, nXCoord, nYCoord
        THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
        THISFORM.CMND1.TOP=INT_015*56
        THISFORM.CMND1.LEFT=INT_015*226
  ENDPROC
  PROCEDURE CMND1.MOUSELEAVE
      LPARAMETERS nButton, nShift, nXCoord, nYCoord
        THISFORM.CMND1.FORECOLOR=RGB(0,0,0)
        THISFORM.CMND1.TOP=INT_015*55
        THISFORM.CMND1.LEFT=INT_015*225
  ENDPROC
 PROCEDURE CMND2.MOUSEENTER
      LPARAMETERS nButton, nShift, nXCoord, nYCoord
        THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
        THISFORM.CMND2.TOP=INT_015*56
        THISFORM.CMND2.LEFT=INT_015*281
  ENDPROC
  PROCEDURE CMND2.MOUSELEAVE
      LPARAMETERS nButton, nShift, nXCoord, nYCoord
        THISFORM.CMND2.FORECOLOR=RGB(0,0,0)
        THISFORM.CMND2.TOP=INT_015*55
        THISFORM.CMND2.LEFT=INT_015*280
 ENDPROC
ENDDEFINE