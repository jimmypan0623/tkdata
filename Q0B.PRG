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
      CAPTION='平均數管制圖'
  ADD OBJECT LBLHD2 AS LABEL WITH;
      LEFT=350,;
      TOP=258,;
      HEIGHT=17,;
      AUTOSIZE=.T.,;
      FONTBOLD=.T.,;
      CAPTION='標準差管制圖'            
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
  ADD OBJECT LBLRUB AS LABEL WITH;
      HEIGHT=17,;
      width=75,;
      alignment=1      
  ADD OBJECT LBLRCB AS LABEL WITH;
      HEIGHT=17,;
      alignment=1,;      
      width=75
  ADD OBJECT LBLRLB AS LABEL WITH;
      HEIGHT=17,;
      alignment=1,;      
      width=75
  proc init
       CALC MAX(AVG_F06) TO TMP4
       CALC MIN(AVG_F06) TO TMP5
       CALC MAX(EXP_14) TO TMP6
       CALC MIN(EXP_14) TO TMP7       
       IF TMP('AVG_F06')+SAM1.A1*TMP('EXP_14')>TMP4
          TMP4=TMP('AVG_F06')+SAM1.A1*TMP('EXP_14')
       ENDIF
       IF TMP('AVG_F06')-SAM1.A1*TMP('EXP_14')<TMP5
          TMP5=TMP('AVG_F06')-SAM1.A1*TMP('EXP_14')   
       ENDIF
       IF TMP('EXP_14')*SAM1.B4>TMP6
          TMP6=TMP('EXP_14')*SAM1.B4
       ENDIF    
       IF TMP('EXP_14')*SAM1.B3<TMP7
          TMP7=TMP('EXP_14')*SAM1.B3
       ENDIF     
       RANGE1=200/(TMP4-TMP5)
       RANGE2=200/(TMP6-TMP7)
*       RANGE1=200/TMP4
*       RANGE2=200/TMP6
       UXBAR=RANGE1*((TMP('AVG_F06')+SAM1.A1*TMP('EXP_14'))-TMP5)
       LXBAR=RANGE1*((TMP('AVG_F06')-SAM1.A1*TMP('EXP_14'))-TMP5)
       CXBAR=RANGE1*(TMP('AVG_F06')-TMP5)
       URBAR=RANGE2*((TMP('EXP_14')*SAM1.B4)-TMP7)
       LRBAR=RANGE2*((TMP('EXP_14')*SAM1.B3)-TMP7)
       CRBAR=RANGE2*(TMP('EXP_14')-TMP7)
       thisform.line(80,505,800,505)
       thisform.line(80,252,800,252)
       thisform.line(80,22,800,22)
       thisform.line(80,275,800,275)
       thisform.line(80,22,80,252)
       thisform.line(80,275,80,505)
*!*	       thisform.line(390,0,390,570)       
       THISFORM.LINE(80,237-CXBAR,800,237-CXBAR)
       THISFORM.LINE(80,490-CRBAR,800,490-CRBAR)
       THISFORM.FORECOLOR=255
       THISFORM.LINE(80,237-UXBAR,800,237-UXBAR)
       THISFORM.LINE(80,237-LXBAR,800,237-LXBAR)    
       THISFORM.LINE(80,490-URBAR,800,490-URBAR)
       THISFORM.LINE(80,490-LRBAR,800,490-LRBAR)
       SELE ANSPT1
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
          THISFORM.LINE(M*RECNO()*20+80,252,M*RECNO()*20+80,22)
          THISFORM.LINE(M*RECNO()*20+80,505,M*RECNO()*20+80,275)
          IF MOD(RECNO(),2^(K-1))=0  
             OH=ALLTRIM(STR(RECNO()))
             IF LEN(OH)=1
                OH='0'+OH
             ENDIF           
             THISFORM.LBLGD.CAPTION=THISFORM.LBLGD.CAPTION+'  '+OH                     
          ENDIF   

          THISFORM.FORECOLOR=65522
          THISFORM.CIRCLE(2,M*RECNO()*20+80,237-RANGE1*(ANSPT1.AVG_F06-TMP5))
          THISFORM.CIRCLE(2,M*RECNO()*20+80,490-RANGE2*((EXP_14)-TMP7))          
          IF RECNO()>1
             THISFORM.FORECOLOR=65280
             THISFORM.LINE(RKK,FKK,M*RECNO()*20+80,237-RANGE1*(ANSPT1.AVG_F06-TMP5))
             THISFORM.LINE(RKK,GKK,M*RECNO()*20+80,490-RANGE2*((EXP_14)-TMP7))
          ENDIF    
          RKK=M*RECNO()*20+80
          FKK=237-RANGE1*(ANSPT1.AVG_F06-TMP5)
          GKK=490-RANGE2*((EXP_14)-TMP7)
          SKIP
       ENDDO   
       THISFORM.LBLXUB.CAPTION=TRANSFORM(TMP('AVG_F06')+SAM1.A1*TMP('EXP_14'),'99999999.99')
       THISFORM.LBLXUB.TOP=237-UXBAR-5
       THISFORM.LBLXCB.CAPTION=TRANSFORM(TMP('AVG_F06'),'99999999.99')
       THISFORM.LBLXCB.TOP=237-CXBAR-5
       THISFORM.LBLXLB.CAPTION=TRANSFORM(TMP('AVG_F06')-SAM1.A1*TMP('EXP_14'),'99999999.99')
       THISFORM.LBLXLB.TOP=237-LXBAR-5
       THISFORM.LBLRUB.CAPTION=TRANSFORM(TMP('EXP_14')*SAM1.B4,'99999999.99')
       THISFORM.LBLRUB.TOP=490-URBAR-5
       THISFORM.LBLRCB.CAPTION=TRANSFORM(TMP('EXP_14'),'99999999.99')
       THISFORM.LBLRCB.TOP=490-CRBAR-5
       THISFORM.LBLRLB.CAPTION=TRANSFORM(TMP('EXP_14')*SAM1.B3,'99999999.99')
       THISFORM.LBLRLB.TOP=490-LRBAR-5       
  endproc
  PROC click
       thisform.release
  ENDPROC     
enddefine