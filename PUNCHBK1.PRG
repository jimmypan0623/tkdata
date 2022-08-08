CLEAR ALL
CLOSE ALL
SET EXCLUSIVE OFF
SET DELETED ON
DTS='202205'
S16='S16'+DTS
 IF !USED('&S16')
   SELE 0
   USE (S16) ALIA S16
ELSE
   SELE S16
ENDIF

SET ORDER TO 1



 

SELECT F01,F02,F05,F04,F03,F11,F12,F24,F25,(EMPTY(F11) AND EMPTY(F24)) OR (EMPTY(F11) AND F24>"11:59") A1,!EMPTY(F24) AND F24 <="09:00" C1,F25>="17:30" D1 ,;
(EMPTY(F12) AND F25<'12:00') B1 FROM S16 WHERE  BETWEEN(F18,2,6) AND DTS+F01<=DTOS(DATE()) AND DTS+F01 NOT IN(SELECT DTOS(F01) FROM S15 WHERE F03='1') ORDER BY F02,F01;
INTO CURSOR PUNCH1 NOWAIT
CREATE CURSOR PUNCH2 (F02 C(4),F05 C(8),F03 C(8),W01 C(8),W02 C(8),W03 C(8),W04 C(8),W05 C(8),W06 C(8),W07 C(8),W08 C(8),W09 C(8),W10 C(8),W11 C(8),W12 C(8),;
W13 C(8),W14 C(8),W15 C(8),W16 C(8),W17 C(8),W18 C(8),W19 C(8),W20 C(8),W21 C(8),W22 C(8),W23 C(8),W24 C(8),W25 C(8),W26 C(8),W27 C(8),W28 C(8),W29 C(8),W30 C(8),;
W31 C(8))
STAFFNO='我是某個值'

SELECT PUNCH1
GO TOP 
DO WHILE !EOF()
      SELECT PUNCH2
    IF PUNCH1.F02<>STAFFNO
     
       APPEND BLANK
       REPLACE F02 WITH PUNCH1.F02
        REPLACE F05 WITH PUNCH1.F05 
       REPLACE F03 WITH PUNCH1.F03      
       STAFFNO=F02   
    ENDIF
  
        F_NO=VAL(PUNCH1.F01)
       
        REPLACE (FIELD(F_NO+3)) WITH PRSNT_STATE(PUNCH1.A1,PUNCH1.B1,PUNCH1.C1,PUNCH1.D1)
        TMPSTATE=(FIELD(F_NO+3))
        IF DTOS(DATE())>=DTS+PUNCH1.F01 AND &TMPSTATE="下午缺勤"  
           REPLACE (FIELD(F_NO+3)) WITH '尚未下班'
        ENDIF
       SELECT PUNCH1
 
    SKIP
ENDDO
DIMENSION TbHead(31)   &&將日期填在陣列準備寫在表頭
SELECT PUNCH2
GO TOP
FOR i=1 TO 31
     TMPSTATE=(FIELD(i+3))
     IF &TMPSTATE=SPACE(8) 
         TbHead(i)=''         
     ELSE    &&如果該欄非空白表示當日上班
         TbHead(i)=PADL(ALLTRIM(STR(i)),2,'0')  &&此時才填入日期於表頭
     ENDIF
ENDFOR

 FILE_NAME="D:\git\S28"+dts+".json"
 **先寫表頭
  STRTOFILE('{'+'"'+'DPTNAME'+'"'+':'+'"'+'部門名稱'+'"'+',';
 +'"'+'STAFFNAME'+'"'+':'+'"'+'員工姓名'+'"'+','+'"'+'W01'+'"'+':'+'"'+TbHead(1)+'"'+',';
 +'"'+'W02'+'"'+':'+'"'+TbHead(2)+'"'+','+'"'+'W03'+'"'+':'+'"'+TbHead(3)+'"'+',';
 +'"'+'W04'+'"'+':'+'"'+TbHead(4)+'"'+','+'"'+'W05'+'"'+':'+'"'+TbHead(5)+'"'+',';
 +'"'+'W06'+'"'+':'+'"'+TbHead(6)+'"'+','+'"'+'W07'+'"'+':'+'"'+TbHead(7)+'"'+',';
 +'"'+'W08'+'"'+':'+'"'+TbHead(8)+'"'+','+'"'+'W09'+'"'+':'+'"'+TbHead(9)+'"'+',';
 +'"'+'W10'+'"'+':'+'"'+TbHead(10)+'"'+','+'"'+'W11'+'"'+':'+'"'+TbHead(11)+'"'+',';
 +'"'+'W12'+'"'+':'+'"'+TbHead(12)+'"'+','+'"'+'W13'+'"'+':'+'"'+TbHead(13)+'"'+',';
 +'"'+'W14'+'"'+':'+'"'+TbHead(14)+'"'+','+'"'+'W15'+'"'+':'+'"'+TbHead(15)+'"'+',';
 +'"'+'W16'+'"'+':'+'"'+TbHead(16)+'"'+','+'"'+'W17'+'"'+':'+'"'+TbHead(17)+'"'+',';
 +'"'+'W18'+'"'+':'+'"'+TbHead(18)+'"'+','+'"'+'W19'+'"'+':'+'"'+TbHead(19)+'"'+',';
 +'"'+'W20'+'"'+':'+'"'+TbHead(20)+'"'+','+'"'+'W21'+'"'+':'+'"'+TbHead(21)+'"'+',';
 +'"'+'W22'+'"'+':'+'"'+TbHead(22)+'"'+','+'"'+'W23'+'"'+':'+'"'+TbHead(23)+'"'+',';
 +'"'+'W24'+'"'+':'+'"'+TbHead(24)+'"'+','+'"'+'W25'+'"'+':'+'"'+TbHead(25)+'"'+',';
 +'"'+'W26'+'"'+':'+'"'+TbHead(26)+'"'+','+'"'+'W27'+'"'+':'+'"'+TbHead(27)+'"'+',';
 +'"'+'W28'+'"'+':'+'"'+TbHead(28)+'"'+','+'"'+'W29'+'"'+':'+'"'+TbHead(29)+'"'+',';
 +'"'+'W30'+'"'+':'+'"'+TbHead(30)+'"'+','+'"'+'W31'+'"'+':'+'"'+TbHead(31)+'"'+'}';
 +CHR(13) + CHR(10),FILE_NAME,1)
&&再寫表身
SELECT PUNCH2
 
DO WHILE !EOF()
  
 STRTOFILE('{'+'"'+'DPTNAME'+'"'+':'+'"'+ALLTRIM(F05)+'"'+',';
 +'"'+'STAFFNAME'+'"'+':'+'"'+ALLTRIM(F03)+'"'+','+'"'+'W01'+'"'+':'+'"'+ALLTRIM(W01)+'"'+',';
 +'"'+'W02'+'"'+':'+'"'+ALLTRIM(W02)+'"'+','+'"'+'W03'+'"'+':'+'"'+ALLTRIM(W03)+'"'+',';
 +'"'+'W04'+'"'+':'+'"'+ALLTRIM(W04)+'"'+','+'"'+'W05'+'"'+':'+'"'+ALLTRIM(W05)+'"'+',';
 +'"'+'W06'+'"'+':'+'"'+ALLTRIM(W06)+'"'+','+'"'+'W07'+'"'+':'+'"'+ALLTRIM(W07)+'"'+',';
 +'"'+'W08'+'"'+':'+'"'+ALLTRIM(W08)+'"'+','+'"'+'W09'+'"'+':'+'"'+ALLTRIM(W09)+'"'+',';
 +'"'+'W10'+'"'+':'+'"'+ALLTRIM(W10)+'"'+','+'"'+'W11'+'"'+':'+'"'+ALLTRIM(W11)+'"'+',';
 +'"'+'W12'+'"'+':'+'"'+ALLTRIM(W12)+'"'+','+'"'+'W13'+'"'+':'+'"'+ALLTRIM(W13)+'"'+',';
 +'"'+'W14'+'"'+':'+'"'+ALLTRIM(W14)+'"'+','+'"'+'W15'+'"'+':'+'"'+ALLTRIM(W15)+'"'+',';
 +'"'+'W16'+'"'+':'+'"'+ALLTRIM(W16)+'"'+','+'"'+'W17'+'"'+':'+'"'+ALLTRIM(W17)+'"'+',';
 +'"'+'W18'+'"'+':'+'"'+ALLTRIM(W18)+'"'+','+'"'+'W19'+'"'+':'+'"'+ALLTRIM(W19)+'"'+',';
 +'"'+'W20'+'"'+':'+'"'+ALLTRIM(W20)+'"'+','+'"'+'W21'+'"'+':'+'"'+ALLTRIM(W21)+'"'+',';
 +'"'+'W22'+'"'+':'+'"'+ALLTRIM(W22)+'"'+','+'"'+'W23'+'"'+':'+'"'+ALLTRIM(W23)+'"'+',';
 +'"'+'W24'+'"'+':'+'"'+ALLTRIM(W24)+'"'+','+'"'+'W25'+'"'+':'+'"'+ALLTRIM(W25)+'"'+',';
 +'"'+'W26'+'"'+':'+'"'+ALLTRIM(W26)+'"'+','+'"'+'W27'+'"'+':'+'"'+ALLTRIM(W27)+'"'+',';
 +'"'+'W28'+'"'+':'+'"'+ALLTRIM(W28)+'"'+','+'"'+'W29'+'"'+':'+'"'+ALLTRIM(W29)+'"'+',';
 +'"'+'W30'+'"'+':'+'"'+ALLTRIM(W30)+'"'+','+'"'+'W31'+'"'+':'+'"'+ALLTRIM(W31)+'"'+'}';
 +CHR(13) + CHR(10),FILE_NAME,1)

   SELECT PUNCH2
   SKIP
ENDDO
 

*********************************

FUNCTION PRSNT_STATE
   PARAMETERS A,B,C,D
   COND='全日到勤'
   IF (A AND B)=.T.
      COND='全天缺勤'
   ELSE
     IF A=.T.
        COND='上午缺勤'
     ENDIF
     IF B=.T.
        COND='下午缺勤'
     ENDIF
   ENDIF
   IF ( C AND D)=.T.
      COND='全天請假'
   ELSE
      IF C=.T.
        COND='上午請假'
      ENDIF
      IF D=.T.
        COND='下午請假'
      ENDIF
   ENDIF
   RETURN COND
