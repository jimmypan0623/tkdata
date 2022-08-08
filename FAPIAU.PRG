CLEAR ALL
CLOSE ALL
INT_117='V0093'
INT_118='V0298'
SET DEFAULT TO Z:\TKDATA
SET DELETED ON 
SET EXCLUSIVE OFF
SELECT 0
USE K25200808 ALIAS K25
SET ORDER TO 1
SELECT 0
USE K2F
SET ORDER TO 1
SELECT K2F.F01,K2F.F02,K25.F01,K25.F03,K25.F09 FROM K2F,K25 WHERE K25.F07=K2F.F01 INTO CURSOR K2FBK ORDER BY K2F.F01 NOWAIT
SELECT 0
USE C13200808 ALIA C13
SET ORDER TO 4
SELECT 0
USE D19200808 ALIA D19
SET ORDER TO 4
SELECT 0
USE P19200808 ALIAS P19
SET ORDER TO 5
SELECT 0
USE E20200808 ALIAS E20
SET ORDER TO 4
SELECT K25
GO TOP
DO WHILE !EOF()
   IF LEFT(F01,1)='3'
      SELECT C13
      SEEK LEFT(K25.F15,10)
      IF FOUND()
         DO WHILE F02=LEFT(K25.F15,10)
            REPLACE F17 WITH K25.F07
            IF LEFT(F04,4)<>'µ|ÃB'
               REPLACE F16 WITH ICASE(K25.F09='2','00',K25.F09='3','09',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) <> 0,'03',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) = 0,'02','06')
            ENDIF
            SKIP   
         ENDDO
      ENDIF
   ELSE
     DEP=''
      DO CASE   
         CASE LEFT(K25.F15,2)$'BABB' AND (K25.F03<>INT_117 AND K25.F03<>INT_118)
              DEP='D19'
         CASE LEFT(K25.F15,2)$'BMBR'   
              DEP='E20'
         CASE LEFT(K25.F15,2)$'BBPB' OR (K25.F03=INT_117 OR K25.F03=INT_118)
              DEP='P19'
         OTHERWISE
              DEP=''     
      ENDCASE
      IF !EMPTY(DEP)
          SELECT &DEP
          IF DEP='P19' AND (K25.F03=INT_117 OR K25.F03=INT_118)
             SET ORDER TO 5
                 IF DEP='P19' AND (K25.F03=INT_117 OR K25.F03=INT_118)
                    SET ORDER TO 5
                    SEEK '³Æµù:'+SUBSTR(K25.F15,3,6)
                    IF FOUND()
                       DO WHILE LEFT(F17,11)='³Æµù:'+SUBSTR(K25.F15,3,6)
                          REPLACE F20 WITH K25.F07
                          IF LEFT(F04,4)<>'µ|ÃB'
                             REPLACE F19 WITH ICASE(K25.F09='2','00',K25.F09='3','09',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) <> 0,'03',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) = 0,'02','06')
                          ENDIF   
                          SKIP
                       ENDDO
                    ENDIF
                 ENDIF   
          ELSE
                    SET ORDER TO 4    
                    SEEK LEFT(K25.F15,10)
                    IF FOUND()
                       DO WHILE F02=LEFT(K25.F15,10)
                          REPLACE F20 WITH K25.F07
                          IF LEFT(F04,4)<>'µ|ÃB'
                             REPLACE F19 WITH ICASE(K25.F09='2','00',K25.F09='3','09',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) <> 0,'03',K25.F09='1' AND MOD(VAL(RIGHT(K25.F01,1)),2) = 0,'02','06')                    
                          ENDIF   
                          SKIP
                       ENDDO
                    ENDIF
          ENDIF   
      ENDIF
   ENDIF
   SELECT K25
   SKIP
ENDDO

SELECT K2FBK
GO TOP
DO WHILE !EOF()
   IF LEFT(F01_B,1)='3'
      SELECT C13
      SEEK LEFT(K2FBK.F02,10)
      IF FOUND()
         DO WHILE F02=LEFT(K2FBK.F02,10)
            REPLACE F17 WITH K2FBK.F01_A
            IF LEFT(F04,4)<>'µ|ÃB'
               REPLACE F16 WITH ICASE(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) = 0,'02','06')
            ENDIF
            SKIP   
         ENDDO
      ENDIF
   ELSE
     DEP=''
      DO CASE   
         CASE LEFT(K2FBK.F02,2)$'BABB' AND (K2FBK.F03<>INT_117 AND K2FBK.F03<>INT_118)
              DEP='D19'
         CASE LEFT(K2FBK.F02,2)$'BMBR'   
              DEP='E20'
         CASE LEFT(K2FBK.F02,2)$'BBPB' OR (K2FBK.F03=INT_117 OR K2FBK.F03=INT_118)
              DEP='P19'
         OTHERWISE
              DEP=''     
      ENDCASE
      IF !EMPTY(DEP)
          SELECT &DEP
          IF DEP='P19' AND (K2FBK.F03=INT_117 OR K2FBK.F03=INT_118)
             SET ORDER TO 5
             SEEK '³Æµù:'+SUBSTR(K2FBK.F02,3,6)
             IF FOUND()
                DO WHILE LEFT(F17,11)=SUBSTR(K2FBK.F02,3,6)
                   REPLACE F20 WITH K2FBK.F01_A
                   IF LEFT(F04,4)<>'µ|ÃB'                   
                      REPLACE F19 WITH ICASE(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) = 0,'02','06')
                   ENDIF   
                   SKIP
                ENDDO
             ENDIF
          ELSE
             SET ORDER TO 4
             SEEK LEFT(K2FBK.F02,10)
             IF FOUND()
                DO WHILE F02=LEFT(K2FBK.F02,10)
                   REPLACE F20 WITH K2FBK.F01_A
                   IF LEFT(F04,4)<>'µ|ÃB'
                      REPLACE F19 WITH ICASE(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' AND MOD(VAL(RIGHT(K2FBK.F01_B,1)),2) = 0,'02','06')                   
                   ENDIF
                   SKIP
                ENDDO
             ENDIF   
          ENDIF   
      ENDIF
   ENDIF
   SELECT K2FBK
   SKIP
ENDDO
CLOSE ALL
CLEAR ALL
SET DEFAULT TO D:\TKMIS
