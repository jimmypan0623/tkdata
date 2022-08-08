close all
clear 
FLG='0'
FCH=''
AREA1='L00'
CHK=''
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE 
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'L00'   
IF !USED('A01')
   SELE 0
   USE A01 INDE A01
ELSE 
   SELE A01
ENDIF
SET ORDER TO A01
IF !USED('L00')
   SELE 0
   USE L01 INDE L01
ELSE
   SELE L01
ENDIF
SET ORDER TO L014
IF !USED('L00')
   SELE 0
   USE L00 INDE L00
ELSE
   SELE L00
ENDIF
SET ORDER TO 1  
GO TOP
L00form=createobject("tkL00")
L00form.show  
define class tkL00 as form
  caption='L00.固定資產類別設定'
*!*	  autocenter=.t.
  controlbox=.F.
  BORDERSTYLE=1
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
  NAME='TKL00'
  add object shape1 as shape1 with;
      autocenter=.t.
  add object shape2 as shape2 with;
      autocenter=.t.
  add object shape3 as shape3 with;
      autocenter=.t.
  ADD OBJECT LBL_REC AS LBL_REC WITH;
      AUTOCENTER =.T.
  ADD OBJECT REC_COUNT AS REC_COUNT  with;
      caption=str(reccount())      
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*27,;
      WIDTH=INT_015*50,;
      CAPTION='類別編號'
  ADD OBJECT LBLF04 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*52,;
      WIDTH=INT_015*50,;
      CAPTION='類別說明'
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*77,;
      WIDTH=INT_015*50,;
      CAPTION='預留殘值'            
  ADD OBJECT LBLF021 AS LABEL WITH;
      LEFT=INT_015*354,;
      TOP=INT_015*77,;
      WIDTH=INT_015*50,;
      CAPTION='%'                
  ADD OBJECT LBLF05 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*102,;
      WIDTH=INT_015*50,;
      CAPTION='攤提期限'                 
  ADD OBJECT LBLF051 AS LABEL WITH;
      LEFT=INT_015*337,;
      TOP=INT_015*102,;
      WIDTH=INT_015*20,;
      CAPTION='年'                
  ADD OBJECT LBLF06 AS LABEL WITH;
      LEFT=INT_015*389,;
      TOP=INT_015*102,;
      WIDTH=INT_015*20,;
      CAPTION='月'                      
  ADD OBJECT LBLF03 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*152,;
      WIDTH=INT_015*50,;
      CAPTION='安全資料'
  ADD OBJECT LBLF031 AS LABEL WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*152,;
      AUTOSIZE=.T.    
  
  ADD OBJECT TXTF01 AS TEXTBOX WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*23,;
      WIDTH=INT_015*49,;
      MAXLENGTH=4,;
      NAME='TXTF01'  
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      LEFT=INT_015*295,;
      TOP=INT_015*48,;
      WIDTH=INT_015*74,;
      MAXLENGTH=20,;
      NAME='TXTF04'                 
  ADD OBJECT TXTF02 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*73,;
      WIDTH=INT_015*60,;
      HEIGHT=INT_015*25,;
      INPUTMASK='99.9999',;
      NAME='TXTF02'                         
  ADD OBJECT TXTF05 AS TEXTBOX WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*98,;
      WIDTH=INT_015*40,;
      HEIGHT=INT_015*25,;
      INPUTMASK='999',;
      NAME='TXTF05'                       
  ADD OBJECT TXTF06 AS TEXTBOX WITH;
      LEFT=INT_015*355,;
      TOP=INT_015*98,;
      WIDTH=INT_015*30,;
      HEIGHT=INT_015*25,;
      INPUTMASK='99',;
      NAME='TXTF06'                       
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=2,;            
      RECORDSOURCE='L00',;      
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2'
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.              
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*455,;
      LEFT=INT_015*558              
  PROC INIT
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)   &&判斷有無新增的權限
       THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  &&判斷有無修改的權限
       THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
       THISFORM.ORPGROUP.PNT_BOTT.VISIBLE=.F.
       THISFORM.SHRGROUP.VISIBLE=.F.
       THISFORM.SHRGROUP.ENABLED=.F.
       thisform.setall('height',INT_015*25,'textbox')
       thisform.setall('FONTSIZE',INT_015*9,'label')
       thisform.setall('FONTSIZE',INT_015*9,'textbox')
       thisform.setall('height',INT_015*17,'label')       
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
       thisform.setall('FONTSIZE',INT_015*9,'COLUMN')
       thisform.setall('FONTSIZE',INT_015*9,'HEADER')       
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*49
       THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='類別編號'
	   THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='類別說明'
	   THISFORM.GRID1.COLUMN1.CONTROLSOURCE='L00.F01'
	   THISFORM.GRID1.COLUMN2.CONTROLSOURCE='L00.F04'
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS
       JK='幣別代號'
  ENDPROC     
  PROC KEY_LIST.INIT 
       with this 
           .additem('依類別代號排列')
       endwith      
  ENDPROC 
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.TXTF01.VALUE=F01
       THISFORM.TXTF02.VALUE=F02
       THISFORM.TXTF04.VALUE=F04
       THISFORM.TXTF05.VALUE=F05
       THISFORM.TXTF06.VALUE=F06              
       THISFORM.LBLF031.CAPTION=F03    
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROCEDURE ORPGROUP.NEW_BOTT.CLICK    
       THISFORM.SETALL('READONLY',.F.,'TEXTBOX')    
       THISFORM.TXTF01.VALUE=''
       THISFORM.TXTF02.VALUE=0
       THISFORM.TXTF05.VALUE=0
       THISFORM.TXTF06.VALUE=0              
       THISFORM.TXTF04.VALUE=''
       THISFORM.LBLF031.CAPTION=''
       THISFORM.TXTF01.SETFOCUS()
       THISFORM.GRID1.ENABLED=.F.
       FLG='1'
  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
     THISFORM.TXTF01.READONLY=.T.
     THISFORM.GRID1.ENABLED=.F.
     THISFORM.TXTF04.SETFOCUS
     LOCK()
     IF !RLOCK()
        DO file_impact
        return
     endif
  endproc 
  PROCEDURE ORPGROUP.DEL_BOTT.CLICK
     if messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6	   
        IF SEEK(L00.F01,'L01')=.T.   
           =MESSAGEBOX('此類別已於L01資產管理基本檔中使用,不得刪除',0+48,'提示訊息視窗')
           RETURN
        ENDIF
        SELE L00
        LOCK()
        IF RLOCK()        
           SELE L00
           DELE
           unlock
           SKIP -1
           IF BOF()
              GO TOP
           ENDIF   
           THISFORM.GRID1.SETFOCUS
           IF THISFORM.GRID1.ACTIVEROW=0
              THISFORM.SETALL('VALUE','','TEXTBOX')
              THISFORM.CMDGROUP.ENABLED=.F.
              THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
              THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
              THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
           ELSE
              THISFORM.CMDGROUP.ENABLED=.T.
              THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)  &&判斷有無修改的權限
              THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
              THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限              
           ENDIF
        ELSE
          DO file_impact    
        ENDIF
     ENDIF        
  ENDPROC
  PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK
    if empty(THISFORM.TXTf01.VALUE)
       =messagebox('類別編號不得空白!',0+48,'提示視窗')
       thisform.txtf01.setfocus
       return
    endif   
    SELE L00
    SET ORDER TO 1
    IF SEEK(THISFORM.TXTF01.VALUE,'L00')=.T.
       =MESSAGEBOX('此類別代號已經建立!',0+48,'請注意')
       THISFORM.TXTF01.VALUE=''
       THISFORM.TXTF01.SETFOCUS  
       RETURN
    ENDIF
    APPEND BLANK
    LOCK()
    IF RLOCK()
       REPLACE F01 WITH THISFORM.TXTF01.VALUE
       REPLACE F02 WITH THISFORM.TXTF02.VALUE
       REPLACE F04 WITH THISFORM.TXTF04.VALUE
       REPLACE F05 WITH THISFORM.TXTF05.VALUE
       REPLACE F06 WITH THISFORM.TXTF06.VALUE              
       REPLACE F03 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
    ENDIF
    UNLOCK
    THISFORM.REC_COUNT.CAPTION=STR(RECCOUNT())
    THISFORM.GRID1.ENABLED=.T.
    THISFORM.REFRESH
    THISFORM.GRID1.SETFOCUS    
    THISFORM.ORPGROUP.NEW_BOTT.CLICK            
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK    
    SELE L00
    IF RLOCK()
       REPLACE F02 WITH THISFORM.TXTF02.VALUE
       REPLACE F04 WITH THISFORM.TXTF04.VALUE    
       REPLACE F05 WITH THISFORM.TXTF05.VALUE
       REPLACE F06 WITH THISFORM.TXTF06.VALUE       
       REPLACE F03 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')          
    ELSE
      DO file_impact
      return   
    ENDIF   
    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK    
  ENDPROC   
  procedure SHRGROUP.ABANDON_BOTT.CLICK
      THISFORM.SHRGROUP.ENABLED=.F.
      THISFORM.SHRGROUP.VISIBLE=.F.      
      THISFORM.ORPGROUP.ENABLED=.T.
      THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.CMDGROUP.ENABLED=.T.
      THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.KEY_LIST.ENABLED=.T.      
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')   
      UNLOCK ALL
      SELE L00
      THISFORM.GRID1.ENABLED=.T. 
      THISFORM.GRID1.SETFOCUS   
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK   
       THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC     
enddefine     
**************************************************刪除資料前檢查是否仍有關聯資料在其他資料表格中
procedure dele_shure 

        
     SELE L00

     LOCK()
     IF RLOCK()        
        SELE L00
        DELE
        unlock
        SKIP -1
        IF BOF()
           GO TOP
        ENDIF   
     ELSE
       DO file_impact    
     ENDIF      
ENDPROC     
************************************************
