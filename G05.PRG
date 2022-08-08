close all
clear 
SET DECIMALS TO 9
FLG='0'
FCH=''
AREA1='G06'
CHK=''
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE 
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'G05'   
IF !USED('A01')
   SELE 0
   USE A01 INDE A01
ELSE 
   SELE A01
ENDIF
SET ORDER TO A01
IF !USED('A23')
   SELE 0
   USE A23
ELSE 
   SELE A23
ENDIF
SET ORDER TO 1
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
   SELE A14
ENDIF
SET ORDER TO A141
SET FILTER TO F04='*' AND F05='*'  
G06='G06'+LEFT(DTOS(DATE()),6)   &&月報表檔
IF !USED('&G06')
   SELE 0
   USE (G06) ALIA G06 
ELSE
   SELE G06
ENDIF
SET ORDER TO 1  
SET RELATION TO F02 INTO A14
GO TOP
G05form=createobject("tkG05")
G05form.show  
define class tkG05 as form
  caption='G05.直接人工及製造費用'
*!*	  autocenter=.t.
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
  NAME='TKG05'
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
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*27,;
      WIDTH=INT_015*50,;
      CAPTION='部門編號'
  ADD OBJECT LBLF021 AS LABEL WITH;
      LEFT=INT_015*341,;
      TOP=INT_015*27,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT LBLF03 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*52,;
      WIDTH=INT_015*50,;
      CAPTION='直接人工'
  ADD OBJECT LBLF04 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*77,;
      WIDTH=INT_015*50,;
      CAPTION='相關費用'      
  ADD OBJECT LBLF10 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*102,;
      WIDTH=INT_015*50,;
      CAPTION='委外工資'            
  ADD OBJECT LBLF101 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*102,;
      AUTOSIZE=.T.,;
      CAPTION=''                  
  ADD OBJECT LBLF12 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*127,;
      WIDTH=INT_015*50,;
      CAPTION='間接材料'            
  ADD OBJECT LBLF121 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*127,;
      AUTOSIZE=.T.,;
      CAPTION=''                              
  ADD OBJECT LBLF07 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*152,;
      WIDTH=INT_015*50,;
      CAPTION='有效工點'            
  ADD OBJECT LBLF071 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*152,;
      AUTOSIZE=.T.,;
      CAPTION=''                        
  ADD OBJECT LBLF05 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*177,;
      WIDTH=INT_015*50,;
      CAPTION='每點人工'            
  ADD OBJECT LBLF051 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*177,;
      AUTOSIZE=.T.,;
      CAPTION=''                              
  ADD OBJECT LBLF08 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*202,;
      WIDTH=INT_015*50,;
      CAPTION='每點製費'            
  ADD OBJECT LBLF081 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*202,;
      AUTOSIZE=.T.,;
      CAPTION=''                                    
  ADD OBJECT LBLF11 AS LABEL WITH;      
      LEFT=INT_015*244,;
      TOP=INT_015*227,;
      WIDTH=INT_015*50,;
      CAPTION='每點委外'            
  ADD OBJECT LBLF111 AS LABEL WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*227,;
      AUTOSIZE=.T.,;
      CAPTION=''                                          
  ADD OBJECT LBLF06 AS LABEL WITH;
      LEFT=INT_015*244,;
      TOP=INT_015*302,;
      WIDTH=INT_015*50,;
      CAPTION='安全資料'
  ADD OBJECT LBLF061 AS LABEL WITH;
      LEFT=INT_015*295,;
      TOP=INT_015*302,;
      AUTOSIZE=.T.      
  ADD OBJECT TXTF02 AS TEXTBOX WITH;      
      LEFT=INT_015*295,;
      TOP=INT_015*23,;
      WIDTH=INT_015*40,;
      MAXLENGTH=5,;
      NAME='TXTF02'  
  ADD OBJECT TXTF03 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      LEFT=INT_015*295,;
      TOP=INT_015*48,;
      WIDTH=INT_015*100,;
      INPUTMASK='99999999.99',;
      NAME='TXTF03'                 
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      AUTOSIZE=.T.,;
      LEFT=INT_015*295,;
      TOP=INT_015*73,;
      WIDTH=INT_015*100,;
      INPUTMASK='99999999.99',;
      NAME='TXTF04'                       

  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.,;
      WIDTH=INT_015*130
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
            AUTOCENTER=.T.            
  ADD OBJECT GRID1 AS GRID1 WITH;
      columncount=2,;            
      RECORDSOURCE='G06',;      
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
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='部門代號'
	   THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='部門名稱'
	   THISFORM.GRID1.COLUMN1.CONTROLSOURCE='G06.F02'
	   THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A14.F02'
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID1.SETFOCUS
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.TXTF02.VALUE=''
          THISFORM.TXTF03.VALUE=0
          THISFORM.TXTF04.VALUE=0
          THISFORM.LBLF021.CAPTION=''
          THISFORM.LBLF061.CAPTION=''
          THISFORM.LBLF101.CAPTION=''          
          THISFORM.LBLF121.CAPTION=''                              
          THISFORM.LBLF071.CAPTION=''                    
          THISFORM.LBLF051.CAPTION=''                              
          THISFORM.LBLF081.CAPTION=''                                        
          THISFORM.LBLF111.CAPTION=''                                                  
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)   &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
       ENDIF                           
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.);
      .AND. !IIF(SEEK(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),'A23','A231'),A23.F09,.F.)                 
       JK='部門編號'
  ENDPROC     

  PROC KEY_LIST.INIT 
       with this 
           .additem('依部門編號排列')
       endwith      
  ENDPROC
  PROC MTH_LIST.INTERACTIVECHANGE
       THISFORM.GRID1.RECORDSOURCE=''
       SELE G06
       SET RELATION OFF INTO A14
       USE
       G06='G06'+LEFT(DTOS(CTOD(THIS.DISPLAYVALUE+'/01')),6)
       IF !USED('&G06')
          SELE 0
          USE (G06) ALIA G06
       ELSE
          SELE G06
       ENDIF
       SET ORDER TO 1                 
       GO TOP
       SET RELA TO F02 INTO A14 
       THISFORM.GRID1.RECORDSOURCE='G06'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='G06.F02'
       THISFORM.GRID1.COLUMN2.CONTROLSOURCE='A14.F02'           
	   THISFORM.REFRESH
	   thisform.grid1.setfocus        
       IF THISFORM.GRID1.ACTIVEROW=0
          THISFORM.CMDGROUP.ENABLED=.F.
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
          THISFORM.TXTF02.VALUE=''
          THISFORM.TXTF03.VALUE=0
          THISFORM.TXTF04.VALUE=0
          THISFORM.LBLF021.CAPTION=''
          THISFORM.LBLF061.CAPTION=''          
          THISFORM.LBLF101.CAPTION=''          
          THISFORM.LBLF121.CAPTION=''                    
          THISFORM.LBLF071.CAPTION=''    
          THISFORM.LBLF051.CAPTION=''                        
          THISFORM.LBLF081.CAPTION=''                    
          THISFORM.LBLF111.CAPTION=''                                                                                                      
       ELSE
          THISFORM.CMDGROUP.ENABLED=.T.      
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)   &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限    
       ENDIF                    
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.);
      .AND. !IIF(SEEK(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),'A23','A231'),A23.F09,.F.)                        
  ENDPROC     
  PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS nColIndex
       THISFORM.TXTF02.VALUE=F02
       THISFORM.LBLF021.CAPTION=A14.F02       
       THISFORM.TXTF03.VALUE=F03
       THISFORM.TXTF04.VALUE=F04
       THISFORM.LBLF061.CAPTION=F06  
       THISFORM.LBLF101.CAPTION=TRANSFORM(F10,'@R 999999999.99')
       THISFORM.LBLF121.CAPTION=TRANSFORM(F12,'@R 99999999.999999')       
       THISFORM.LBLF071.CAPTION=TRANSFORM(F07,'@R 999999999.99')       
       THISFORM.LBLF051.CAPTION=TRANSFORM(F05,'@R 99999.999999999')              
       THISFORM.LBLF081.CAPTION=TRANSFORM(F08,'@R 99999.999999999')                     
       THISFORM.LBLF111.CAPTION=TRANSFORM(F11,'@R 99999.999999999')                            
       FCH=THISFORM.ACTIVECONTROL.NAME
       CHK=''
       thisform.refresh
  ENDPROC       
  PROCEDURE ORPGROUP.NEW_BOTT.CLICK    
       THISFORM.SETALL('READONLY',.F.,'TEXTBOX')    
       THISFORM.MTH_LIST.ENABLED=.F.
       THISFORM.TXTF02.VALUE=''
       THISFORM.TXTF03.VALUE=0
       THISFORM.TXTF04.VALUE=0
       THISFORM.LBLF021.CAPTION=''       
       THISFORM.LBLF061.CAPTION=''
       THISFORM.LBLF101.CAPTION=''       
       THISFORM.LBLF121.CAPTION=''              
       THISFORM.LBLF071.CAPTION=''              
       THISFORM.LBLF051.CAPTION=''
       THISFORM.LBLF081.CAPTION=''       
       THISFORM.LBLF111.CAPTION=''              
       THISFORM.TXTF02.SETFOCUS()
       THISFORM.GRID1.ENABLED=.F.
       FLG='1'
  ENDPROC  
  PROCEDURE ORPGROUP.EDIT_BOTT.CLICK
      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')    
       THISFORM.MTH_LIST.ENABLED=.F.      
     THISFORM.GRID1.ENABLED=.F.
     THISFORM.TXTF02.READONLY=.T.
     THISFORM.TXTF03.SETFOCUS
     SELECT G06
     LOCK()
     IF !RLOCK()
        DO file_impact
        return
     endif
  endproc 
  PROCEDURE ORPGROUP.DEL_BOTT.CLICK
     if messagebox('是否確定要刪除此筆資料',4+32+256,'請確認') = 6	      
        SELE G06
        LOCK()
        IF RLOCK()        
           SELECT G06
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
              THISFORM.LBLF021.CAPTION=''         
              THISFORM.LBLF061.CAPTION=''     
              THISFORM.LBLF101.CAPTION=''              
              THISFORM.LBLF121.CAPTION=''                            
              THISFORM.LBLF071.CAPTION=''                                          
              THISFORM.LBLF051.CAPTION=''                                                        
              THISFORM.LBLF081.CAPTION=''                                                                      
              THISFORM.LBLF111.CAPTION=''                                                                                    
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
    if empty(THISFORM.TXTf02.VALUE)
       =messagebox('部門編號不得空白!',0+64,'提示視窗')
       thisform.txtf02.setfocus
       return
    endif   
    SELE G06
    SET ORDER TO 1
    IF SEEK(THISFORM.TXTF02.VALUE,'G06')=.T.
       =MESSAGEBOX('此部門本月之工資與費用已經建立!',0+64,'請注意')
       THISFORM.TXTF02.VALUE=''
       THISFORM.TXTF02.SETFOCUS  
       RETURN
    ENDIF
    APPEND BLANK
    LOCK()
    IF RLOCK()
       REPLACE F02 WITH THISFORM.TXTF02.VALUE
       REPLACE F03 WITH THISFORM.TXTF03.VALUE
       REPLACE F04 WITH THISFORM.TXTF04.VALUE
       REPLACE F06 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
    ENDIF    
    UNLOCK
    SEEK THISFORM.TXTF02.VALUE
    THISFORM.REC_COUNT.CAPTION=STR(RECCOUNT())
    THISFORM.GRID1.ENABLED=.T.
    THISFORM.REFRESH
    THISFORM.GRID1.SETFOCUS    
    THISFORM.ORPGROUP.NEW_BOTT.CLICK            
  ENDPROC
  PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK    
    SELE G06
    IF RLOCK()
       REPLACE F02 WITH THISFORM.TXTF02.VALUE
       REPLACE F03 WITH THISFORM.TXTF03.VALUE    
       REPLACE F04 WITH THISFORM.TXTF04.VALUE
       REPLACE F06 WITH IIF(INT_012=1,PADL(ALLTRIM(STR(YEAR(DATE())-1911)),3,' ')+RIGHT(DTOC(DATE()),6),dtoc(DATE())+' ')+IIF(SEEK(sys_oper,'A01'),A01.F03,'')          
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
          THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)   &&判斷有無修改的權限
          THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.)   &&判斷有無刪除的權限
          THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)   &&判斷有無列印的權限  
       THISFORM.MTH_LIST.ENABLED=.T.          
      UNLOCK ALL
      SELE G06
      THISFORM.GRID1.ENABLED=.T.
      THISFORM.GRID1.SETFOCUS      
  ENDPROC
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK   
       THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC     
   PROCEDURE TXTF02.INTERACTIVECHANGE
        IF AT('?',ALLTRIM(THIS.VALUE))<>0
            IF AT('?',ALLTRIM(THIS.VALUE))>1             
               SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE LEFT(F01,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)  AND F04='*'  AND F05='*' AND LEN(ALLTRIM(F01))=4
            ELSE
               SELECT F01,F02 FROM A14 INTO CURSOR DEPART ORDER BY F01 WHERE  F04='*'  AND F05='*' AND LEN(ALLTRIM(F01))=4
            ENDIF   
            IF _TALLY>0
                DEPART_SEEK=createobject("DEPART_SEEK")  
                DEPART_SEEK.SHOW    
                THIS.VALUE=FLG
                THIS.REFRESH
                FLG='0'
            ENDIF   
        ENDIF   
       THISFORM.LBLF021.CAPTION=IIF(SEEK(THIS.VALUE,'A14'),A14.F02,'')            
   ENDPROC
enddefine    
**************************************************刪除資料前檢查是否仍有關聯資料在其他資料表格中
procedure dele_shure 

        
     SELE G06

     LOCK()
     IF RLOCK()        
        SELE G06
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
