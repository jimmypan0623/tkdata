q00form=createobject("tkq00")
q00form.show  
define class tkq00 as form
  caption='管制圖'+B01.F01+B01.F02
  autocenter=.t.
  BORDERSTYLE=3  
  CONTROLBOX=.F.
  MOVABLE=.F.
  controlcount=57
  fontsize=9
  height=550
  showtips=.t.
  showwindow=1
  width=789
  windowtype=1
  NAME='TKq00'
  ADD OBJECT LBLGD AS LABEL WITH;
      LEFT=88,;
      TOP=510,;
      HEIGHT=17,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT LBLHD1 AS LABEL WITH;
      LEFT=350,;
      TOP=5,;
      HEIGHT=17,;
      AUTOSIZE=.T.,;
      FONTBOLD=.T.,;
      CAPTION='不良率管制圖'    
  ADD OBJECT LBLXUB AS LABEL WITH;
      HEIGHT=17,;
      width=75,;
      alignment=1
  ADD OBJECT LBLXCB AS LABEL WITH;
      HEIGHT=17,;
      width=75,;
      alignment=1,;      
      AUTOSIZE=.T.
  ADD OBJECT LBLXLB AS LABEL WITH;
      HEIGHT=17,;
      width=75,;
      alignment=1

  proc init
       CALC MAX(EXP_6) TO TMP4
       CALC MIN(EXP_6) TO TMP5
       IF (FTR_STR+3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100>TMP4
          TMP4=(FTR_STR+3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100
       ENDIF
       IF (FTR_STR-3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100<TMP5
          TMP5=(FTR_STR-3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100
       ENDIF
       RANGE1=400/(TMP4-TMP5)
       UXBAR=RANGE1*(((FTR_STR+3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100)-TMP5)
       LXBAR=RANGE1*(((FTR_STR-3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100)-TMP5)
       CXBAR=RANGE1*(FTR_STR*100-TMP5)
       thisform.line(80,505,800,505)
       thisform.line(80,22,800,22)
       thisform.line(80,22,80,505)
       THISFORM.LINE(80,490-CXBAR,800,490-CXBAR)
       THISFORM.FORECOLOR=255
       THISFORM.LINE(80,490-UXBAR,800,490-UXBAR)
       THISFORM.LINE(80,490-LXBAR,800,490-LXBAR)    

       SELE ANSPT2
       IF RECCOUNT()<=35 
          K=1
       ELSE
          K=INT(RECCOUNT()/35)
          IF MOD(RECCOUNT(),35)>0
             K=K+1
          ENDIF   
       ENDIF   
       M=1/2^(K-1)
       GO TOP
       DO WHILE !EOF()
          THISFORM.FORECOLOR=0

          THISFORM.LINE(M*RECNO()*20+80,505,M*RECNO()*20+80,22)
          IF MOD(RECNO(),2^(K-1))=0  
             OH=ALLTRIM(STR(RECNO()))
             IF LEN(OH)=1
                OH='0'+OH
             ENDIF           
             THISFORM.LBLGD.CAPTION=THISFORM.LBLGD.CAPTION+'  '+OH                     
          ENDIF   

          THISFORM.FORECOLOR=65522
          THISFORM.CIRCLE(2,M*RECNO()*20+80,490-RANGE1*(EXP_6-TMP5))

          IF RECNO()>1
             THISFORM.FORECOLOR=65280
             THISFORM.LINE(RKK,FKK,M*RECNO()*20+80,490-RANGE1*(EXP_6-TMP5))

          ENDIF    
          RKK=M*RECNO()*20+80
          FKK=490-RANGE1*(EXP_6-TMP5)

          SKIP
       ENDDO   
       THISFORM.LBLXUB.CAPTION=TRANSFORM((FTR_STR+3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100,'99999999.99')
       THISFORM.LBLXUB.TOP=490-UXBAR-5
       THISFORM.LBLXCB.CAPTION=TRANSFORM(FTR_STR*100,'99999999.99')
       THISFORM.LBLXCB.TOP=490-CXBAR-5
       THISFORM.LBLXLB.CAPTION=TRANSFORM((FTR_STR-3*SQRT((FTR_STR-FTR_STR^2)/RECCOUNT()))*100,'99999999.99')
       THISFORM.LBLXLB.TOP=490-LXBAR-5

  endproc
  PROC click
       thisform.release
  ENDPROC     
enddefine