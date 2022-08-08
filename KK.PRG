   OMT1=MONTH(DATE())+1   &&計算本月最後一天為幾號
          IF OMT1>12
              OMT1=OMT1%12
              OMT2=VAL(LEFT(DTOS(DATE())),4))+1
           ELSE
              OMT2=VAL(LEFT(DTOS(DATE()),4))  
          ENDIF
          OMT3=RIGHT(DTOS(CTOD(ALLTRIM(STR(OMT2))+'/'+PADL(ALLTRIM(STR(OMT1)),2,'0')+'/01')-1),2)
          ?OMT3
 
 