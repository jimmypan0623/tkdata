INT_015=1.28
DEFINE WINDOW WIN FROM 0,0 TO 800,600 ;
      FONT "ARIAL",12;
      STYLE "TU";
      TITLE "測試";
      CLOSE ;
      FLOAT;
      GROW
      ACTIVATE WINDOW WIN
      SYS_OPER='00002'
      
SELECT A02.F03,A03.F02 FROM A02,A03 WHERE LEFT(A02.F03,1)='A' AND A02.F01=sys_oper AND A03.F01=A02.F03 INTO TABLE AA ORDER BY A02.F03
SELECT A02.F03,A03.F02 FROM A02,A03 WHERE LEFT(A02.F03,1)='B' AND A02.F01=sys_oper AND A03.F01=A02.F03 INTO TABLE BB ORDER BY A02.F03
SELECT A02.F03,A03.F02 FROM A02,A03 WHERE LEFT(A02.F03,1)='C' AND A02.F01=sys_oper AND A03.F01=A02.F03 INTO TABLE CC ORDER BY A02.F03
SELECT A02.F03,A03.F02 FROM A02,A03 WHERE LEFT(A02.F03,1)='D' AND A02.F01=sys_oper AND A03.F01=A02.F03 INTO TABLE DD ORDER BY A02.F03

*SET STATUS BAR ON
DEFINE WINDOW mywindow AT 5,5 SIZE 27,71


SET SYSMENU TO 

DEFINE MENU mymenu IN WINDOW WIN;
     FONT"新細明體",12;
     STYLE  'BI'
DEFINE PAD mypadA OF mymenu PROMPT "\<A. 系統管理"
DEFINE PAD mypadB OF mymenu PROMPT "\<B. 庫存管理"
DEFINE PAD mypadC OF mymenu PROMPT "\<C. 銷貨管理"
DEFINE PAD mypadD OF mymenu PROMPT "\<D. 採購管理"
DEFINE PAD mypadX OF mymenu PROMPT "\<X. 離        開"

ON  SELECTION PAD mypadA OF mymenu ACTIVATE POPUP AA1

ON  SELECTION PAD mypadB OF mymenu ACTIVATE POPUP BB1

ON  SELECTION PAD mypadC OF mymenu ACTIVATE POPUP CC1

ON  SELECTION PAD mypadD OF mymenu ACTIVATE POPUP DD1

ON  SELECTION PAD mypadX OF mymenu  DO leave

 DEFINE POPUP AA1 FONT "新細明體",12 STYLE 'BN' FROM 1,0 PROMPT FIELD AA.F03+'  '+AA.F02 SCROLL  

 ON SELECTION POPUP AA1 DO browit  WITH PROMPT()

DEFINE POPUP BB1  FONT "新細明體",12 STYLE 'BN'  FROM 1,8  PROMPT FIELD BB.F03+'  '+BB.F02 SCROLL  

ON SELECTION POPUP BB1 DO  browit WITH PROMPT()

DEFINE POPUP CC1   FONT "新細明體",12 STYLE 'BN'  FROM 1,16  PROMPT FIELD CC.F03+'  '+CC.F02 SCROLL 

ON SELECTION POPUP CC1 DO browit WITH PROMPT()

DEFINE POPUP DD1   FONT "新細明體",12 STYLE 'BN'  FROM 1,24  PROMPT FIELD DD.F03+'  '+DD.F02 SCROLL 

ON SELECTION POPUP DD1 DO browit WITH PROMPT()



*!*	DEFINE POPUP DD1  FONT "新細明體",12 STYLE 'BN'  FROM 1,14  PROMPT FIELD F03+'   '+F02 SCROLL
*!*	ON SELECTION POPUP DD1  DO browit WITH PROMPT()


*!*	CLOSE DATABASES 
mPad=""
	DO WHILE mPad==""
               IF !USED('AA')
                   SELECT 0
                   USE AA
                ELSE
                    SELECT AA
                ENDIF
                IF !USED('BB')
                    SELECT 0
                    USE BB
                 ELSE
                    SELECT  BB
                ENDIF
                IF !USED('CC')
                     SELECT 0
                     USE CC
                 ELSE
                      SELECT CC
                 ENDIF
                 IF !USED('DD')
                      SELECT 0
                      USE DD
                 ELSE
                      SELECT DD
                 ENDIF                                 

	       ACTIVATE MENU mymenu
               mPad=PAD()
	ENDDO
RELEASE MENUS mymenu EXTEND
SET SYSMENU TO DEFAULT
PROCEDURE leave
        DEACTIVATE MENU mymenu
        
RETURN



PROCEDURE browit
PARAMETERS mprompt
OPX=LEFT(mprompt,3)

       DO  (OPX)     

       SELECT 0
       USE AA
       SELECT 0
       USE BB
       SELECT 0
       USE CC
       SELECT 0
       USE DD
 RETURN 
      