**┰计传衡いゅ肂
CLEAR 
CLOSE ALL

FLG=.F.

A=12474040203
?A
H=INT(LOG10(A))
OTMN=''
FOR I=H TO 0 STEP -1
    BK=INT(A/10^I)
    OTMN=OTMN+IIF(RIGHT(OTMN,2)='箂' AND NBR(BK)='箂','',NBR(BK))
    IF I%4=0  
       IF RIGHT(OTMN,2)='箂'
          OTMN=LEFT(OTMN,LEN(OTMN)-2)
          IF RIGHT(OTMN,2)$'ㄕ珺'
             FLG=.F.
          ELSE
             FLG=.T.   
          ENDIF 
       ENDIF  
    ELSE
       FLG=.F.          
    ENDIF
    OTMN=OTMN+IIF((NBR(BK)='箂' AND I%4<>0) OR (FLG AND I%4=0),'',K1(I))   
    A=A-BK*10^I
    ?A
ENDFOR

?OTMN+'じ俱'
FUNCTION K1
   PARAMETERS BT
   RV=''
   DO CASE     
      CASE BT=1
           RV='珺'
      CASE BT=2
           RV='ㄕ'
      CASE BT=3
           RV=''
      CASE BT=4
           RV='窾'
      CASE BT=5
           RV='珺'
      CASE BT=6
           RV='ㄕ'
      CASE BT=7
           RV=''
      CASE BT=8
           RV='货'
      CASE BT=9
           RV='珺'          
      CASE BT=10
           RV='ㄕ'  
      CASE BT=11
           RV=''    
      CASE BT=12
           RV=''        
       CASE BT=13
           RV='珺'  
       CASE BT=14
           RV='ㄕ'        
       CASE BT=15
           RV=''                                               
   ENDCASE
   RETURN RV
  
  
 FUNCTION NBR
     PARAMETERS GS
     RS=''
  DO CASE   
      CASE GS=0
           RS='箂'
      CASE GS=1
           RS='滁'
      CASE GS=2
           RS='禠'
      CASE GS=3
           RS='把'
      CASE GS=4
           RS='竩'
      CASE GS=5
           RS='ヮ'
      CASE GS=6
           RS='嘲'
      CASE GS=7
           RS='琺'
      CASE GS=8
           RS=''
      CASE GS=9
           RS='╤'     
  ENDCASE
  RETURN RS         
           
     