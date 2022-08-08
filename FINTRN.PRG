CLEAR ALL
CLOSE ALL
SET DATE TO TAIWAN
USE E15 EXCLUSIVE ALIAS TRN
ZAP
SET DEFAULT TO Z:\TKDATA
YRM=DTOS(CTOD('2005/08/01'))
IF !USED('A23')
   SELECT 0
   USE A23
ELSE    
   SELECT A23
ENDIF   
SET ORDER TO A231
SEEK YRM
IF FOUND()
   DO WHILE !EOF()
      H='B07'+LEFT(DTOS(A23.F01),6)
      T='B55'+LEFT(DTOS(A23.F01),6)
      IF !USED('&H')
         SELECT 0
         USE &H
      ELSE
         SELECT &H
      ENDIF
      SET ORDER TO 1
      IF !USED('&T')
         SELECT 0
         USE &T
      ELSE
         SELECT &T
      ENDIF
      SET ORDER TO 1
      IF !USED('TRN')
         SELECT 0
         USE TRN  
      ELSE
         SELECT TRN
      ENDIF
      SET ORDER TO 1
*      DELETE ALL
      SELECT &H
      GO TOP
      DO WHILE !EOF()
         IF &H..F10='2'
            SELECT TRN
            SEEK &H..F09+&H..F11
            IF !FOUND()
                APPEND BLANK
                REPLACE F01 WITH &H..F09
                REPLACE F03 WITH &H..F03
                REPLACE F04 WITH &H..F08
                REPLACE F05 WITH &H..F11
                REPLACE F06 WITH &H..F14
            ENDIF    
            IF F02<CTOD(LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'/'+&H..F02)            
                REPLACE F02 WITH CTOD(LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'/'+&H..F02)                            
            ENDIF   
            REPLACE F07 WITH F07+&H..F04           
         ENDIF   
         SELECT &H
         SKIP
      ENDDO       
      SELECT &T
      GO TOP
      DO WHILE !EOF()
         IF &T..F13='2'
            SELECT TRN
            SEEK &T..F05+SPACE(4)
            IF !FOUND()
                APPEND BLANK
                REPLACE F01 WITH &T..F05
                REPLACE F03 WITH &T..F08
                REPLACE F04 WITH &T..F11
                REPLACE F06 WITH &T..F07                          
            ENDIF
            IF F02<CTOD(LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'/'+&T..F02)            
                REPLACE F02 WITH CTOD(LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'/'+&T..F02)            
            ENDIF    
            REPLACE F07 WITH F07+&T..F12         
         ENDIF   
         SELECT &T
         SKIP
      ENDDO            
      SELECT A23
      SKIP
   ENDDO   
ENDIF
SELECT TRN
BROW
CLOSE ALL
SET DEFAULT TO D:\TKMIS