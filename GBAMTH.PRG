CLOSE ALL
CLEAR ALL
SET DELETED ON
SET EXCLUSIVE OFF
YRS='2022'
SELECT 0
USE gbamth
GO TOP
DO WHILE !EOF()
   FOR I=2 TO 25
       REPLACE (FIELD(I)) WITH 0
   ENDFOR
   SKIP
ENDDO
DO WHILE !BOF()
   FOR I=1 TO 7
       MTH1='B04'+YRS+PADL(ALLTRIM(STR(I)),2,'0')
       MTH2='C10'+YRS+PADL(ALLTRIM(STR(I)),2,'0')
       SELECT 0
       USE \\192.168.10.213\TKDATA\&MTH1 ALIAS B04      
       **USE &MTH1 ALIAS B04      
       SET ORDER TO 1
       SELECT 0
       USE \\192.168.10.213\TKDATA\&MTH2 ALIAS C10
       **USE &MTH2 ALIAS C10
       SET ORDER TO 3
    
       SEEK GBAMTH.F01
       IF FOUND()
          K=0
          DO WHILE F03=GBAMTH.F01
             IF F02='C8152' OR F02='C4017' OR F02='C4032' OR F02='C8520' OR F02='C2958'
                K=K+C10.F08
                SELECT GBAMTH                 
                REPLACE (FIELD(I*2)) WITH K
                SELECT B04
                SEEK C10.F04+C10.F03
                IF FOUND()
                   SELECT GBAMTH
                   REPLACE (FIELD(I*2+1)) WITH B04.F15
                ENDIF
             ENDIF
            SELECT C10
            SKIP
          ENDDO
       ENDIF
       SELECT B04
       USE
       SELECT C10
       USE
   ENDFOR 
  
   SELECT GBAMTH
   SKIP -1
ENDDO
CLOSE ALL


