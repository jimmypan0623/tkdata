close all
clear 
FLG='0'
TIMCNT=0
AREA1='A02'
AREA2=''
CHK=''
PJMform=createobject("tkPJM")
PJMform.show  
define class tkPJM as form
  i=0 
  caption='  主  畫  面  選  單'
  autocenter=.t.
  controlbox=.F.
  BORDERSTYLE=3  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  controlcount=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  showtips=.t.
  showwindow=1
  WIDTH=INT_015*789
  windowtype=1
  NAME='TKPJM'

  ADD OBJECT TIMER1 AS TIMER WITH;
      INTERVAL=200
  ADD OBJECT CMDA AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*20,;
      CAPTION='\<A.系統管理'
  ADD OBJECT CMDB AS COMMANDBUTTON WITH;      
      LEFT=INT_015*10,;
      TOP=INT_015*80,;
      CAPTION='\<B.庫存管理'
  ADD OBJECT CMDC AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*140,;
      CAPTION='\<C.業務管理'            
  ADD OBJECT CMDD AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*200,;
      CAPTION='\<D.採購管理'
  ADD OBJECT CMDP AS COMMANDBUTTON WITH;
      LEFT=INT_015*558,;
      TOP=INT_015*20,;
      CAPTION='\<P.外購管理'            
  ADD OBJECT CMDE AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*260,;
      CAPTION='\<E.生產管理'      
  ADD OBJECT CMDF AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*320,;
      CAPTION='\<F.料表管理' 
 ADD OBJECT CMDG AS COMMANDBUTTON WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*380,;
      CAPTION='\<G.成本計算'                    
  ADD OBJECT CMDH AS COMMANDBUTTON WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*20,;
      CAPTION='\<H.品質管理'                        
  ADD OBJECT CMDK AS COMMANDBUTTON WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*80,;
      CAPTION='\<K.財務管理'                  
  ADD OBJECT CMDS AS COMMANDBUTTON WITH;
      LEFT=INT_015*200,;
      TOP=INT_015*140,;
      CAPTION='\<S.人事薪資'            
  ADD OBJECT CMDQ AS COMMANDBUTTON WITH;
      LEFT=INT_015*558,;
      TOP=INT_015*80,;
      CAPTION='\<Q.統計品管'                   
  ADD OBJECT CMDQT AS COMMANDBUTTON WITH;
      LEFT=INT_015*558,;
      TOP=INT_015*140,;
      CAPTION='\<R.重新登錄'
*!*	  ADD OBJECT CMDBK1 AS COMMANDBUTTON WITH;
*!*	      LEFT=INT_015*558,;
*!*	      TOP=INT_015*200,;
*!*	      CAPTION='資料重整'                 
*!*	  ADD OBJECT CMDBK2 AS COMMANDBUTTON WITH;
*!*	      LEFT=INT_015*558,;
*!*	      TOP=INT_015*260,;
*!*	      CAPTION='產生檔案'            
*!*	  ADD OBJECT CMDBK3 AS COMMANDBUTTON WITH;
*!*	      LEFT=INT_015*558,;
*!*	      TOP=INT_015*320,;
*!*	      CAPTION='參數設定'              
*!*	  ADD OBJECT CMDBK4 AS COMMANDBUTTON WITH;
*!*	      LEFT=INT_015*558,;
*!*	      TOP=INT_015*380,;
*!*	      CAPTION='單據編號暫存檔'                        
*!*	  ADD OBJECT CMDBK5 AS COMMANDBUTTON WITH;
*!*	      LEFT=INT_015*558,;
*!*	      TOP=INT_015*440,;
*!*	      CAPTION='訂單表身檔'                  
  PROC INIT       
       thisform.setall('FONTSIZE',INT_015*14,'COMMANDBUTTON')
       thisform.setall('AUTOSIZE',.T.,'COMMANDBUTTON')
  ENDPROC     
  PROC CMDA.CLICK
       BK='A'
       DO CHOICE       
  ENDPROC
  PROC CMDB.CLICK
       BK='B'
       DO CHOICE
  ENDPROC
  PROC CMDC.CLICK
       BK='C'
       DO CHOICE
  ENDPROC
  PROC CMDD.CLICK
       BK='D'
       DO CHOICE
  ENDPROC  
  PROC CMDE.CLICK
       BK='E'
       DO CHOICE
  ENDPROC   
  PROC CMDF.CLICK
       BK='F'
       DO CHOICE
  ENDPROC
  PROC CMDG.CLICK
       BK='G'
       DO CHOICE
  ENDPROC     
  PROC CMDH.CLICK
       BK='H'
       DO CHOICE
  ENDPROC   
  PROC CMDK.CLICK
       BK='K'
       DO CHOICE
  ENDPROC     
  PROC CMDP.CLICK
       BK='P'
       DO CHOICE
  ENDPROC        
  PROC CMDQ.CLICK
       BK='Q'
       DO CHOICE
  ENDPROC  
  PROC CMDS.CLICK
       BK='S'
       DO CHOICE
  ENDPROC            
*!*	  PROC CMDBK1.CLICK
*!*	       DO FORM INDEX
*!*	  ENDPROC
*!*	  PROC CMDBK2.CLICK
*!*	       DO FORM JIMMY
*!*	  ENDPROC  
*!*	  PROC CMDBK3.CLICK
*!*	       DO FORM A26
*!*	  ENDPROC  
*!*	  PROC CMDBK4.CLICK
*!*	      DO BK4    
*!*	  ENDPROC    
*!*	  PROC CMDBK5.CLICK
*!*	      DO BK5    
*!*	  ENDPROC      
  PROC CMDQT.CLICK
       IF !USED('A04')
          SELE 0
          USE A04
       ELSE 
          SELE MODI 
       Endif 
       SET ORDER TO 1
       SEEK sys_oper
       if found()
          replace f01 with 'zzzzz'
          REPLACE F04 WITH ''
          REPLACE F06 WITH ''
       endif   
       close table all       
       THISFORM.RELEASE
  ENDPROC   
*!*	  PROC TIMER1.TIMER       
*!*	       THISFORM.i=THISFORM.i+1 
*!*	       IF mod(THISFORM.i,20)=0 and TIMCNT=0
*!*	          ?CHR(7)
*!*	          ??CHR(7)
*!*	          ??CHR(7)
*!*	          =MESSAGEBOX('您有單據未簽核',0+64,'請馬上處理')
*!*	          DO Q04
*!*	       ENDIF   
*!*	  ENDPROC     
ENDDEFINE       
PROCEDURE CHOICE
   SELECT a02.F03,a03.f02 FROM A02,a03 WHERE A02.F01=sys_oper and LEFT(A02.F03,1)=BK AND A03.F01=A02.F03 INTO CURSOR PRGS ORDER BY A02.F03
*   SELECT a02.F03,a03.f02 FROM A02,a03  WHERE A02.F01=sys_oper and LEFT(A02.F03,1)=BK and A03.F01=A02.F03 INTO CURSOR PRGS ORDER BY A02.F03   
       PRG_SEEK=createobject("PRG_SEEK")  
       PRG_SEEK.SHOW    
      
       FLG='0'
ENDPROC       
