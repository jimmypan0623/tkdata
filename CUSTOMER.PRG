CLOSE ALL
CLEAR ALL
SELECT 0
USE CUSTOMER ALIAS CST
SET DEFAULT TO \\192.168.10.200\MIS\MIS\TKDATA
SELECT 0
USE B01
SET ORDER TO 1
SELECT 0
USE C01
SET ORDER TO 1
SELECT CST
GO TOP
DO WHILE !EOF()
       CSNAME=ALLTRIM(IIF(SEEK(CST.F01,'C01'),C01.F05,''))
       FOR I =1 TO 5
                H='C10'+'2003'+PADL(ALLTRIM(STR(I)),2,'0')                
                K='G11'+'2003'+PADL(ALLTRIM(STR(I)),2,'0')                
               SELECT 0
               USE &H ALIAS C10
                SELECT 0
               USE &K ALIAS G11
*!*	               SELECT C10.F01 出貨日期,C10.F03 料品編號,C10.F04 出貨單號,C10.F05 訂單編號,C10.F07 銷售單價,C10.F08 出貨數量,G11.F21 成本單價,C10.F07-G11.F21 毛利,IIF(B01.F07='I001 ','線材','部品') 出貨類別;
*!*	               FROM C10,G11,B01 WHERE C10.F02=CST.F01 AND G11.F02=C10.F03 AND B01.F01=C10.F03  ORDER BY C10.F03,C10.F01 NOWAIT
               SELECT C10.F01 出貨日期,C10.F03 料品編號,C10.F04 出貨單號,C10.F05 訂單編號,C10.F07 銷售單價,C10.F08 出貨數量,B01.F37 成本單價,C10.F07-B01.F37 毛利,IIF(B01.F07='I001 ','線材','部品') 出貨類別;
               FROM C10,B01 WHERE C10.F02=CST.F01  AND B01.F01=C10.F03  ORDER BY C10.F03,C10.F01 NOWAIT
               IF _TALLY>0
                   FNAME=CST.F01+CSNAME+'2003'+PADL(ALLTRIM(STR(I)),2,'0')          
                   COPY TO C:\TOKUTSU\&FNAME TYPE XL5     
               ENDIF    
               SELECT C10
               USE 
               SELECT G11
               USE 
       ENDFOR
       SELECT CST
       SKIP
ENDDO       
CLOSE ALL
CLEAR ALL
SET DEFAULT TO C:\TKMIS
