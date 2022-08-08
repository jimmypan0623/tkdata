CLEAR ALL
CLOSE ALL

? chk_unit('28579479')

Function chk_unit()
   Lparameters p_unit
   LOCAL W_CHECK,W_CNT, W_CNT1,ii
   IF EMPTY(p_unit)
      RETURN .F.
   endif 
   If ' '$p_unit
      RETURN .F.
   ENDIF
   IF LEN(ALLTRIM(p_unit))<>8
      RETURN .F.
   ENDIF   
   Store 0 To W_CNT, W_CNT1
   ii = 1
   FOR ii=1 TO 8
       W_CHECK = Str(Val(Substr(p_unit,ii,1))*Val(Substr('12121241',ii,1)),2)
       W_CNT = W_CNT+Val(Substr(W_CHECK,1,1))+Val(Substr(W_CHECK,2,1))
   NEXT ii
   If Val(Substr(p_unit,7,1))>=7
      W_CNT = W_CNT-Val(Substr(p_unit,7,1))*4+Int(Val(Substr(p_unit,7,1))*4/10)
      W_CNT1 = W_CNT1-Val(Substr(p_unit,7,1))*4+Mod(Val(Substr(p_unit,7,1))*4,10)
      If !Mod(W_CNT,10)#0 .And. !Mod(W_CNT1,10)#0
         RETURN .F.
      Endif
   Else
      If Mod(W_CNT,10)#0
         RETURN .F.
      Endif
   ENDIF
   RETURN .T.
Endfunc

