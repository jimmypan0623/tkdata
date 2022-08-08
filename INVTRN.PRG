****程式名稱 INVCRT發票轉帳作業
CLEAR ALL
SET TALK OFF
SET EXCL OFF
SET DELE ON
SET CENT OFF
SET SAFE OFF
SET DECIMALS TO 6
SET REPROCESS TO 1
_SCREEN.CAPTION=''
_SCREEN.CONTROLBOX=.F.
ZOOM WIND SCREEN MIN
CLEAR ALL
CLOS DATA ALL
CLOS TABL ALL
SET TALK OFF
SET EXCL OFF
SET DELE ON
SET CENT OFF
SET SAFE OFF
SET DECIMALS TO 6
SET ENGINEBEHAVIOR 70 
SET REPORTBEHAVIOR 80
SET REPROCESS TO 1
_SCREEN.CAPTION=''
_SCREEN.CONTROLBOX=.F.
ZOOM WINDOW  SCREEN MAX
HIDE WINDOW SCREEN
DEACTIVATE WIND SCREEN
SET SYSMENU OFF
SET STATUS BAR OFF
SET MULTILOCKS ON
DECLARE _FPRESET IN msvcrt20.dll 
CAPSLOCK(.T.)

SET DATE TO YMD
************************
SELE 0
USE A26
X_CHK = .F.
DO WHIL ! EOF()
    X_CHECK = TRIM(F08)
    X_OK = &X_CHECK
    REPL F03 WITH IIF(X_OK,'  ','※')
    IF ! X_OK
      X_CHK = .T.
    ENDI
    X_NAME = TRIM(F01)
    DO CASE
      CASE F04 = 'C'
        X_VAL = LEFT(F06,MAX(F05,1))
      CASE F04 = 'N'
        X_VAL = VAL(F06)
      CASE F04 = 'L'
        X_VAL = ( F06='Y' )
    ENDC 
    PUBLIC &X_NAME
    &X_NAME = X_VAL
    IF X_NAME = 'INT_000'
    	&X_NAME = ALLTRIM(&X_NAME)
    ENDI
    SKIP
ENDD
*
DO CASE
    CASE X_CHK            && FJMIS 進入系統時有錯誤
	       	ERROR("系統參數錯誤，請以ＦＪＤＩＡＧ重新設定後執行作業 ! ")
	    CLOS DATA
    	QUIT
    CASE X_CHK            && FJDIAG 進入系統時有錯誤
	        ERROR("系統參數錯誤，請重新設定才可執行ＦＪＭＩＳ ! ")
	    CLOS DATA
 	    RETU
    OTHE
ENDC

SELE A26
USE
*****************************
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

IF !USED('B01')
    SELECT 0
    USE B01
ELSE
   SELECT B01
ENDIF
SET ORDER TO 1
IF !USED('A01')
    SELECT 0
    USE A01
ELSE
    SELECT A01
ENDIF
SET ORDER TO 1   
PUBLIC UKEF
UKEF=30000
PUBLIC MSGPATH
MSGPATH='\\192.168.10.213\tkdata\INVDGT\'
&&MSGPATH='\\192.168.10.199\Mis\FIGBK\'
PUBLIC DTES[2]

DO FORM INVTRN

PROCEDURE INVHEAD
****處理上傳電子發票表頭檔
                           RSTB=PADR(ALLTRIM(STR((RAND()*10000))),4,'0') &&隨機碼
                           STRTOFILE('H'+;  
                           PADR(LEFT(K25TMP.F15,10),36,' ')+;
                           LEFT(DTOS(DATE()),4)+'-'+SUBSTR(DTOS(DATE()),5,2)+'-'+RIGHT(DTOS(DATE()),2)+SPACE(1)+TIME()+; 
                           K25TMP.F07+; 
                           LEFT(DTOS(DTES[I]),4)+'-'+SUBSTR(DTOS(DTES[I]),5,2)+'-'+K25TMP.F02;  
                           +TIME()+;                     
                           PADR(INT_107,10,' ')+;
                           PADR(CHRTRAN(INT_000,' ',''),60,' ')+;
                           PADR(CHRTRAN(INT_109,' ',''),100,' ')+;
                           SPACE(12)+;
                           SPACE(15)+;
                           SPACE(15)+;
                           SPACE(40)+;
                           PADR(IIF(!EMPTY(K25TMP.F04),K25TMP.F04,'0000000000'),10,' ')+;
                           PADR(IIF(!EMPTY(ALLTRIM(K25TMP.F04)),IIF(SEEK(K25TMP.F03,'C01'),C01.F41,''),'0000'),60,' ')+;
                           SPACE(100)+;
                           SPACE(12)+;
                           SPACE(15)+;
                           SPACE(15)+;
                           SPACE(40)+;
                           PADR(K25TMP.F03+SPACE(1)+LEFT(K25TMP.F15,10)+SPACE(1)+IIF(K25TMP.F21=INT_011,'','單價幣別:'+IIF(K25TMP.F21='US  ','USD ',K25TMP.F21)),200,' ')+;
                           K25TMP.F05+;
                           '07'+;
                           '0'+;
                           SPACE(6)+;
                           SPACE(64)+;
                           SPACE(64)+;
                           'Y'+;
                           SPACE(7)+;
                           PADR(IIF(EMPTY(ALLTRIM(K25TMP.F04)),RSTB,''),4,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='1',IIF(EMPTY(K25TMP.F04),K25TMP.F12,K25TMP.F08),0),'@R 999999999999')),12,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='3',K25TMP.F08,0),'@R 999999999999')),12,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='2',K25TMP.F08,0),'@R 999999999999')),12,' ')+;
                           K25TMP.F09+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='1',INT_002/100,0),'@R 99.99')),5,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(EMPTY(K25TMP.F04),0,K25TMP.F10),'@R 999999999999')),12,' ')+;
                           PADR(LTRIM(TRANSFORM(K25TMP.F12,'@R 999999999999')),12,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F21=INT_011,0,K25TMP.F23),'@R 999999999.99')),12,' ')+;
                           PADR(LTRIM(TRANSFORM(IIF(K25TMP.F21=INT_011,1,K25TMP.F22),'@R 9999999.9999')),12,' ')+;
                           PADR(IIF(K25TMP.F21=INT_011,'   ',IIF(ALLTRIM(K25TMP.F21)='US','USD',K25TMP.F21)),3,' ')+;
                           IIF(K25TMP.F01='31' AND K25TMP.F09='2','1',' ')+IIF(SEEK(K25TMP.F03,'C01'),C01.F47,' ')+CHR(13) + CHR(10),FILE_NAME2,1)
                           
                           IF INT_160
                              ****處理列印電子發票證明聯表頭檔
                              STRTOFILE('H'+;
                              PADR(LEFT(K25TMP.F15,10),36,' ')+;
                              LEFT(DTOS(DATE()),4)+'-'+SUBSTR(DTOS(DATE()),5,2)+'-'+RIGHT(DTOS(DATE()),2)+SPACE(1)+TIME()+;
                              K25TMP.F07+;
                              LEFT(DTOS(DTES[I]),4)+'-'+SUBSTR(DTOS(DTES[I]),5,2)+'-'+K25TMP.F02;
                              +TIME()+;
                              PADR(LEFT(K25TMP.F15,10),20,' ')+;
                              PADR(INT_107,10,' ')+;
                              PADR(CHRTRAN(INT_000,' ',''),60,' ')+;
                              PADR(CHRTRAN(INT_109,' ',''),100,' ')+;
                              SPACE(15)+;
                              PADR(CHRTRAN(INT_000,' ',''),20,' ')+;
                              PADR(IIF(!EMPTY(K25TMP.F04),K25TMP.F04,' '),10,' ')+;
                              PADR(IIF(SEEK(K25TMP.F03,'C01'),C01.F41,' '),60,' ')+;
                              SPACE(100)+;
                              SPACE(96)+;
                              '07'+;
                              PADR(IIF(EMPTY(ALLTRIM(K25TMP.F04)),RSTB,''),4,' ')+;
                              PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='1',IIF(EMPTY(K25TMP.F04),K25TMP.F12,K25TMP.F08),0),'@R 999999999999')),12,' ')+;
                              PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='3',K25TMP.F08,0),'@R 999999999999')),12,' ')+;
                              PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='2',K25TMP.F08,0),'@R 999999999999')),12,' ')+;
                              K25TMP.F09+;
                              PADR(LTRIM(TRANSFORM(IIF(K25TMP.F09='1',INT_002/100,0),'@R 99.99')),5,' ')+;
                              PADR(LTRIM(TRANSFORM(IIF(EMPTY(K25TMP.F04),0,K25TMP.F10),'@R 999999999999')),12,' ')+;
                              PADR(LTRIM(TRANSFORM(K25TMP.F12,'@R 999999999999')),12,' ')+;
                              SPACE(1)+;
                              '3'+SPACE(32)+;
                              SPACE(32)+;
                              PADR('USERQQ',40,' ')+CHR(13) + CHR(10),FILE_NAME3,1)    
                           ENDIF
                           
ENDPROC                           
