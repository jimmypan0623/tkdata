close all
clear 
FLG='0'
FCH=''
CHK=''
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'C96'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('A23')
   SELE 0
   USE A23
ELSE 
  SELE A23
ENDIF
SET ORDER TO 1
IF !USED('C01')
   SELE 0
   USE C01
ELSE 
  SELE C01
ENDIF
SET ORDER TO 1
IF !USED('B01')
   SELE 0
   USE B01
ELSE 
  SELE B01
ENDIF
SET ORDER TO 1
C96form=createobject("tkC96")
C96form.show       
define class tkC96 as form
*!*	  autocenter=.t.
  caption='C96.客戶別銷售-毛利(率%)分析'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*320
  width=INT_015*420
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkC96'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入日期範圍'    
  ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*250,;
      TOP=INT_015*40,;
      FONTSIZE=INT_015*30,;
      AUTOSIZE=.T.,;
      CAPTION='~'               
  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT TXT1 AS TEXTBOX WITH;
      LEFT=INT_015*130,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      FONTSIZE=INT_015*12,;    
      INPUTMASK='####/##',;
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*300,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*90,;
      INPUTMASK='####/##',;      
      NAME='TXT2'          
 ADD OBJECT O_TPE AS OPTIONGROUP WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*130,;
      HEIGHT=INT_015*60,;    
      WIDTH=INT_015*60,;      
      BUTTONCOUNT=3,;
      OPTION1.CAPTION='部品',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='製品',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*4,;
      OPTION2.TOP=INT_015*23,;
      OPTION2.AUTOSIZE=.T.,;      
      OPTION3.CAPTION='全部',;
      OPTION3.FONTSIZE=INT_015*9,;
      OPTION3.LEFT=INT_015*4,;
      OPTION3.TOP=INT_015*43,;
      OPTION3.AUTOSIZE=.T.,;      
      NAME='O_TPE'                   
      
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*250,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所查詢條件',;
      NAME='CMND1'
  add object cmnd2 as commandbutton with;
      left=INT_015*190,;
      top=INT_015*250,;
      height=INT_015*40,;
      width=INT_015*80,;
      FONTSIZE=INT_015*12,;
      caption='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.O_TPE.VALUE=1
         THISFORM.TXT1.VALUE='' &&LEFT(DTOS(DATE()),4)+'/'+SUBSTR(DTOS(DATE()),5,2) 
         THISFORM.TXT2.VALUE='' &&LEFT(DTOS(DATE()),4)+'/'+SUBSTR(DTOS(DATE()),5,2) 
         THISFORM.TXT1.SETFOCUS
      ENDPROC
**********************************開始搜尋指定範圍之內容         
      PROCEDURE CMND1.CLICK      
        IF EMPTY(CTOD(ALLTRIM(THISFORM.TXT1.VALUE)+'/01')) 
           MESSAGEBOX('月份輸入錯誤',0+48,'提示訊息視窗')
           THISFORM.TXT1.SETFOCUS
           RETURN           
        ENDIF
        IF SEEK(DTOS(CTOD(ALLTRIM(THISFORM.TXT1.VALUE)+'/01')),'A23')=.F.
           MESSAGEBOX('無此月份資料',0+48,'提示訊息視窗')
           THISFORM.TXT1.SETFOCUS
           RETURN                   
        ENDIF
        IF EMPTY(CTOD(ALLTRIM(THISFORM.TXT2.VALUE)+'/01'))
           MESSAGEBOX('月份輸入錯誤',0+48,'提示訊息視窗')
           THISFORM.TXT2.SETFOCUS
           RETURN           
        ENDIF           
        IF SEEK(DTOS(CTOD(ALLTRIM(THISFORM.TXT2.VALUE)+'/01')),'A23')=.F.
           MESSAGEBOX('無此月份資料',0+48,'提示訊息視窗')
           THISFORM.TXT2.SETFOCUS
           RETURN                   
        ENDIF        
        IF THISFORM.TXT1.VALUE>THISFORM.TXT2.VALUE
           MESSAGEBOX('起始月份不得大於迄止月份',0+48,'提示訊息視窗')
           THISFORM.TXT2.SETFOCUS        
        ENDIF
        DO CASE 
           CASE THISFORM.O_TPE.VALUE=1
                M='N'
           CASE THISFORM.O_TPE.VALUE=2
                M='Y'
           CASE THISFORM.O_TPE.VALUE=3
                M='YN'          
         ENDCASE       
             CREATE TABLE CMP1 (客戶編號 C(5),客戶簡稱 C(8),業務A C(8),業務B C(8),業助 C(8))

            INDEX ON 客戶編號 TAG CMP1
            SET ORDER TO 1

        SELECT A23
        SEEK DTOS(CTOD(ALLTRIM(THISFORM.TXT1.VALUE)+'/01'))
        IF FOUND() 
           I=0
           DO WHILE F01 <=CTOD(ALLTRIM(THISFORM.TXT2.VALUE)+'/01')
              I=I+1
              H=LEFT(DTOS(A23.F01),6)
              J='售'+H
              K='毛'+H
              L='率'+H
              SELECT CMP1
              ALTER TABLE CMP1 ADD COLUMN &J N(16,2)
              ALTER TABLE CMP1 ADD COLUMN &K N(16,2)              
              ALTER TABLE CMP1 ADD COLUMN &L N(7,3)              
              R='C10'+H
              S='G11'+H
              SELECT 0
              USE &S ALIAS GAA
              SET ORDER TO 1
              SELECT 0
              USE &R ALIAS CAK
              SET ORDER TO 2
              SELECT CAK
              GO TOP
              DO WHILE !EOF()
                 IF IIF(SEEK(CAK.F03,'B01'),B01.F97,'')$M
                    SELECT CMP1
                    SET ORDER TO 1                    
                    SEEK CAK.F02
                    IF !FOUND()
                       APPEND BLANK
                       REPLACE 客戶編號 WITH CAK.F02
                       REPLACE 客戶簡稱 WITH CAK.F09
                       REPLACE 業務A WITH CAK.F11
                       REPLACE 業務B WITH CAK.F13
                       REPLACE 業助 WITH CAK.F15
                    ENDIF   
                    F1=FIELD(3+3*I)
                    REPLACE &F1 WITH &F1+CAK.F07*CAK.F08
                    F2=FIELD(3+3*I+1)                 
                    SELECT GAA
                    SEEK CAK.F03
                    IF FOUND()       
                       SELECT CMP1
                       REPLACE &F2 WITH &F2+(GAA.F21*CAK.F08) 
                    ENDIF             
                 ENDIF
                 SELECT CAK
                 SKIP
              ENDDO
              SELECT CAK
              USE
              SELECT GAA
              USE
              SELECT CMP1
              REPLACE &F2 WITH &F1-&F2 ALL  
              F3=FIELD(3+3*I+2)
              GO TOP
              DO WHILE !EOF()
                 IF &F1<>0
                    REPLACE &F3 WITH &F2/&F1*100 
                 ENDIF
                 SELECT CMP1
                 SKIP
              ENDDO                                       
              SELECT A23
              SKIP
           ENDDO
          SELECT CMP1



      


               FILE_NAME='客戶別銷售毛利分析'+ICASE(THISFORM.O_TPE.VALUE=1,'部品',THISFORM.O_TPE.VALUE=2,'製品','全部')
	           GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
	           IF EMPTY(GCDELIMFILE)  && ESC PRESSED
  	              RETURN
	           ELSE
	             ON ERROR DO FILE_IMPACT
	             COPY TO (GCDELIMFILE) TYPE XL5   
	             =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
  	          ENDIF


         THISFORM.CMND2.CLICK           
        ENDIF

        
      ENDPROC
      procedure cmnd2.click        
           IF !USED('A05')
              SELECT 0
              USE A05
            ELSE
              SELECT A05
            ENDIF
            SET ORDER TO 1
            SEEK sys_oper+'C96'
            IF FOUND()
              DELETE 
            ENDIF          
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          
******************************

