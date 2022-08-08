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
CREATE CURSOR B09K (F01 C(10),F02 C(2),F03 C(35),F04 N(14,4),F05 C(5),F07 C(100))
CREATE CURSOR MOFIX (F01 C(12),F02 N(8))
INDEX ON F01 TAG MOFIX
SET ORDER TO 1
     
     
   


                            IF !USED('E03')
                                  SELECT 0
                                  USE E03
                            ELSE
                                  SELECT E03
                            ENDIF
                            SET ORDER TO E031
                            IF !USED('E05')
                                 SELECT 0
                                 USE E05
                            ELSE
                                 SELECT E05
                            ENDIF
                            SET ORDER TO E052
                            IF !USED('E06')
                                 SELECT 0
                                 USE E06
                            ELSE
                                 SELECT E06
                            ENDIF
                            SET ORDER TO E061                            
                            IF !USED('E07')
                                 SELECT 0
                                 USE E07
                            ELSE
                                 SELECT E07
                            ENDIF
                            SET ORDER TO E071
                            IF !USED('E11')
                                 SELECT 0
                                 USE E11
                            ELSE
                                 SELECT E11
                            ENDIF
                            SET ORDER TO E111    
                            IF !USED('E16')
                                 SELECT 0
                                 USE E16
                            ELSE
                                 SELECT E16
                            ENDIF
                            SET ORDER TO E161      
IF !USED('A09')
   SELECT 0
   USE A09
ELSE 
   SELECT A09
ENDIF
SET ORDER TO 1                             
IF !USED('A23')
   SELECT 0
   USE A23
ELSE 
   SELECT A23
ENDIF
SET ORDER TO 1    
A24='A24'+LEFT(DTOS(DATE()),6)
IF !USED('&A24')
   SELE 0
   USE (A24) ALIA A24
ELSE
   SELE A24
ENDIF
SET ORDER TO 1                
B09='B09'+LEFT(DTOS(DATE()),6)
IF !USED('&B09')
   SELE 0
   USE (B09) ALIA B09
ELSE
   SELE B09
ENDIF
SET ORDER TO 1                                                            
DO FORM MOFIX
