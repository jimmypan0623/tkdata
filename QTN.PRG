****程式名稱 C02報價紀錄日期查詢轉EXCEL
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
HIDE WINDOW SCREEN
DEACTIVATE WIND SCREEN
SET SYSMENU OFF
SET STATUS BAR OFF
SET MULTILOCKS ON
CAPSLOCK(.T.)
PUBLIC PR
IF !USED('A26')
     SELECT 0
     USE A26
ELSE
     SELECT A26
ENDIF     
SET ORDER TO 1
 SEEK 'INT_012'
 IF FOUND()
      IF ALLTRIM(F06)='1'
          SET DATE TO TAIWAN
      ELSE
          SET DATE TO YMD
      ENDIF        
ENDIF     
 
DO FORM QTN
********產地別
     FUNC ORG_PLACE
     PARA JUK
     JO='       '
     DO CASE
        CASE JUK='B '
             JO='比利時'
        CASE JUK='C '
             JO='美國'
        CASE JUK='F '
             JO='法國'
        CASE JUK='G '
             JO='上海'
        CASE JUK='J '
             JO='日本'      
        CASE JUK='Q '
             JO='韓國'
        CASE JUK='W '
             JO='印尼'
        CASE JUK='Z '
             JO='日本'
        CASE JUK='D '
             JO='馬來西亞'
        CASE JUK='E '
             JO='歐洲'
        OTHERWISE
             JO='        '             
     ENDCASE                                               
    RETURN JO