
SELECT A02.F03,A03.F02 FROM A02,A01 WHERE A02.F03=sys_oper AND a03.f01=a02.f03 INTO CURSOR cbx ORDER BY a02.f03
bt_cnt=_tally
DEFINE CHOICE_BTT AS COMMANDGROUP
      buttoncount=bt_cnt
      PROCEDURE init
           WITH THIS
                FOR i=1 TO bt_cnt
                        
                ENDFOR
            ENDWITH        
          
   S=S+IIF(bt_cnt>32,35,50)*INT_015
   IF S>IIF(bt_cnt>32,399,364)*INT_015
      S=0
      IK=IK+198*INT_015
   ENDIF   
  SKIP
ENDDO   
      endproc

ENDDEFINE