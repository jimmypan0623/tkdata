**���ԧB�Ʀr���⤤����B
CLEAR 
CLOSE ALL

FLG=.F.

A=12474040203
?A
H=INT(LOG10(A))
OTMN=''
FOR I=H TO 0 STEP -1
    BK=INT(A/10^I)
    OTMN=OTMN+IIF(RIGHT(OTMN,2)='�s' AND NBR(BK)='�s','',NBR(BK))
    IF I%4=0  
       IF RIGHT(OTMN,2)='�s'
          OTMN=LEFT(OTMN,LEN(OTMN)-2)
          IF RIGHT(OTMN,2)$'�a�լB'
             FLG=.F.
          ELSE
             FLG=.T.   
          ENDIF 
       ENDIF  
    ELSE
       FLG=.F.          
    ENDIF
    OTMN=OTMN+IIF((NBR(BK)='�s' AND I%4<>0) OR (FLG AND I%4=0),'',K1(I))   
    A=A-BK*10^I
    ?A
ENDFOR

?OTMN+'����'
FUNCTION K1
   PARAMETERS BT
   RV=''
   DO CASE     
      CASE BT=1
           RV='�B'
      CASE BT=2
           RV='��'
      CASE BT=3
           RV='�d'
      CASE BT=4
           RV='�U'
      CASE BT=5
           RV='�B'
      CASE BT=6
           RV='��'
      CASE BT=7
           RV='�a'
      CASE BT=8
           RV='��'
      CASE BT=9
           RV='�B'          
      CASE BT=10
           RV='��'  
      CASE BT=11
           RV='�a'    
      CASE BT=12
           RV='��'        
       CASE BT=13
           RV='�B'  
       CASE BT=14
           RV='��'        
       CASE BT=15
           RV='�a'                                               
   ENDCASE
   RETURN RV
  
  
 FUNCTION NBR
     PARAMETERS GS
     RS=''
  DO CASE   
      CASE GS=0
           RS='�s'
      CASE GS=1
           RS='��'
      CASE GS=2
           RS='�L'
      CASE GS=3
           RS='��'
      CASE GS=4
           RS='�v'
      CASE GS=5
           RS='��'
      CASE GS=6
           RS='��'
      CASE GS=7
           RS='�m'
      CASE GS=8
           RS='��'
      CASE GS=9
           RS='�h'     
  ENDCASE
  RETURN RS         
           
     