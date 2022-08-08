set excl off
close all
clear 
AREA1='ITT'
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
SEEK sys_oper+'D10'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('D01')
   SELE 0
   USE D01
ELSE 
   SELE D01
ENDIF
SET ORDER TO 1      
D10form=createobject("tkD10")
D10form.show       
define class tkD10 as form
   AUTOCENTER=.T.
  caption='D10.交貨提示跟催單'
  fontsize=INT_015*9
  height=INT_015*270
  width=INT_015*300
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkD10'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入廠商編號'
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*120,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入基準天數'      
  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT TXT1 AS TXT1 WITH;
      LEFT=INT_015*120,;
      TOP=INT_015*35,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*50,;
      MAXLENGTH=5,;
      NAME='TXT1'    
  ADD OBJECT TXT2 AS TEXTBOX WITH;
      LEFT=INT_015*120,;
      TOP=INT_015*115,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*50,;
      NAME='TXT2'          
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*210,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.確認',;
      TOOLTIPTEXT='確認所查詢條件',;
      NAME='CMND1'
  add object cmnd2 as commandbutton with;
      left=INT_015*110,;
      top=INT_015*210,;
      height=INT_015*40,;
      width=INT_015*80,;
      FONTSIZE=INT_015*12,;
      caption='\<X.離開',;
      TOOLTIPTEXT='離開此查詢尋畫面!快速鍵->ALT+X'
      name='cmnd2'
      procedure init 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.TXT1.READONLY=.F.
         THISFORM.TXT2.VALUE=0
      ENDPROC
      PROCEDURE CMND1.CLICK
         SELE D07.F01,D07.F02,D07.F03,D07.F04,D07.F05,D07.F06,B01.F02 FROM D07,B01;
         WHERE IIF(EMPTY(THISFORM.TXT1.VALUE),D07.F05>SPACE(5),D07.F05=THISFORM.TXT1.VALUE) AND  D07.F03<=DATE()+THISFORM.TXT2.VALUE;
         AND LEFT(D07.F01,1)<>'B' AND B01.F01=D07.F02  ORDER BY D07.F05,D07.F02,D07.F03 INTO CURSOR ITT_1
         IF _TALLY>0
             GHK=THISFORM.TXT2.VALUE
             CREATE CURSOR ITT;
             (F01 C(10),F02_A C(43),F02_B C(40),F03 D(8),F04 N(11,2),F05 C(5))
             INDE ON F05+DTOS(F03) TAG ITT1
             INDE ON F01+F02_A TAG ITT2
             INDE ON F02_A+DTOS(F03) TAG ITT3
             SET ORDER TO 1             
             SET RELA TO F05 INTO D01             
             SELE ITT_1
             GO TOP
             DO WHILE !EOF()
                SELE ITT
                APPEND BLANK
                REPL F01 WITH ITT_1.F01
                REPL F02_A WITH ITT_1.F02_A
                REPL F02_B WITH ITT_1.F02_B
                REPL F03 WITH ITT_1.F03
                REPL F04 WITH ITT_1.F04
                REPL F05 WITH ITT_1.F05
                SELE ITT_1
                SKIP   
             ENDDO
             SELE ITT_1
             USE   
             SELE ITT
             GO TOP
             ITT_SEEK=createobject("ITT_SEEK")  
             ITT_SEEK.SHOW                           
         ENDIF
         THISFORM.CMND2.CLICK
      ENDPROC
      procedure cmnd2.click        
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          
***************************************************特殊輸入欄位設定-----------廠商
DEFINE CLASS TXT1 AS TEXTBOX
  READONLY=.T.
  MAXLENGTH=5
  FONTSIZE=INT_015*9  
  PROCEDURE INTERACTIVECHANGE
    IF AT('?',ALLTRIM(THIS.VALUE))<>0 
       IF AT('?',ALLTRIM(THIS.VALUE))>1             
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE LEFT(F02,AT('?',THIS.VALUE)-1)=LEFT(THIS.VALUE,AT('?',THIS.VALUE)-1)
       ELSE
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02    
       ENDIF          
       IF _TALLY>0
          VDR_SEEK=createobject("VDR_SEEK")  
          VDR_SEEK.SHOW    
          THIS.VALUE=FLG
          THIS.REFRESH
          FLG='0'        
       ENDIF   
    ELSE
       IF LEN(ALLTRIM(THIS.VALUE))=4
          SELECT F02,F01,F04 FROM D01 INTO CURSOR VDR ORDER BY F02 WHERE F02=ALLTRIM(THIS.VALUE)
          DO CASE
             CASE _TALLY=1
                  THIS.VALUE=VDR.F01
                  THIS.REFRESH
             CASE _TALLY>1
                  VDR_SEEK=createobject("VDR_SEEK")  
                  VDR_SEEK.SHOW    
                  THIS.VALUE=FLG
                  THIS.REFRESH
                  FLG='0'        
          ENDCASE      
       ENDIF
    
    ENDIF      
       THISFORM.LBLF11.CAPTION=IIF(SEEK(THIS.VALUE,'D01'),D01.F04,'')         
  ENDPROC
ENDDEFINE  
**********************************************************************************************
define class ITT_SEEK as form
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.  
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*789
  FONTSIZE=INT_015*12
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='ITT_SEEK'
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T. 
  add object ORPgroup as ORPGROUP with;            
      TOP=0,;
      LEFT=INT_015*400                  
  add object cmdgroup as cmdgroup with;
      TOP=0,;
      ENABLED=.T.,;
      LEFT=INT_015*300      
  add object GRID1 as GRID1 with;          
      AUTOCENTER=.T.,;
      TOP=INT_015*30,;
      WIDTH=INT_015*789,;      
      HEIGHT=INT_015*500,;  
      ROWHEIGHT=INT_015*18,;
      HEADERHEIGHT=INT_015*18,;          
      deletemark=.f.,;
      READONLY=.T.,;
      recordsourcetype=1,;
      columncount=8,;            
      NAME='GRID1',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;      
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7',;
      COLUMN8.NAME='COLUMN8'      
      PROCEDURE INIT 
         THISFORM.CAPTION='交貨提示跟催單  基準天數:'+ALLTRIM(STR(GHK))+'  天'         
         THISFORM.GRID1.READONLY=.T. 
         THISFORM.GRID1.RECORDSOURCE='ITT'         
         THISFORM.ORPGROUP.NEW_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
         THISFORM.ORPGROUP.DEL_BOTT.VISIBLE=.F.
         THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
         THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
         THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(ITT.F03-DATE()<0,RGB(255,255,0),'')","COLUMN")                                    
         THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='廠商編號'
         THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F05'
         THISFORM.GRID1.COLUMN1.WIDTH=INT_015*50
         THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='廠商簡稱'
         THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
         THISFORM.GRID1.COLUMN2.CONTROLSOURCE='D01.F04'
         THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預進貨日'
         THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F03'
         THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                   
         THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='料品編號'
         THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F02_A'
         THISFORM.GRID1.COLUMN4.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='品名規格'
         THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.F02_B'
         THISFORM.GRID1.COLUMN5.WIDTH=INT_015*149
         THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='採購單號'
         THISFORM.GRID1.COLUMN6.CONTROLSOURCE='ITT.F01'
         THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
         THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='進貨數量'
         THISFORM.GRID1.COLUMN7.CONTROLSOURCE='ITT.F04'
         THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100               
         THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='剩餘天數'
         THISFORM.GRID1.COLUMN8.CONTROLSOURCE='ITT.F03-DATE()'
         THISFORM.GRID1.COLUMN8.WIDTH=INT_015*100    
         JK='廠商編號'             
         THISFORM.GRID1.SETFOCUS        
      ENDPROC
  PROC KEY_LIST.INIT 
       with this 
           .additem('依廠商編號+預進貨日排列')
           .additem('依採購單號+  料號  排列')
           .additem('依  料號  +預進貨日排列')           
       endwith      
  ENDPROC 
  PROC KEY_LIST.INTERACTIVECHANGE                 
       SELE ITT
       DO CASE
          CASE THIS.DISPLAYVALUE='依廠商編號+預進貨日排列'         
               SET ORDER TO 1 
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='廠商編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F05'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='廠商簡稱'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='D01.F04'
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預進貨日'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F03'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                   
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F02_A'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='ITT.F02_B'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='採購單號'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='ITT.F01'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='進貨數量'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='ITT.F04'
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100               
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='剩餘天數'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='ITT.F03-DATE()'
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*100         
               THISFORM.GRID1.SETFOCUS             
          CASE THIS.DISPLAYVALUE='依採購單號+  料號  排列'          
               SET ORDER TO 2          
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='採購單號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F01'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*80   
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.F02_A'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F02_B'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*149                           
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='廠商編號'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F05'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='廠商簡稱'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='D01.F04'
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='預進貨日'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='ITT.F03'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*65                   
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='進貨數量'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='ITT.F04'
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100               
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='剩餘天數'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='ITT.F03-DATE()'
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*100         
               THISFORM.GRID1.SETFOCUS                   
          CASE THIS.DISPLAYVALUE='依  料號  +預進貨日排列'
               SET ORDER TO 3           
               THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='料品編號'
               THISFORM.GRID1.COLUMN1.CONTROLSOURCE='ITT.F02_A'
               THISFORM.GRID1.COLUMN1.WIDTH=INT_015*149
               THISFORM.GRID1.COLUMN2.HEADER1.CAPTION='品名規格'
               THISFORM.GRID1.COLUMN2.CONTROLSOURCE='ITT.F02_B'
               THISFORM.GRID1.COLUMN2.WIDTH=INT_015*149         
               THISFORM.GRID1.COLUMN3.HEADER1.CAPTION='預進貨日'
               THISFORM.GRID1.COLUMN3.CONTROLSOURCE='ITT.F03'
               THISFORM.GRID1.COLUMN3.WIDTH=INT_015*65                                                         	                                       
               THISFORM.GRID1.COLUMN4.HEADER1.CAPTION='廠商編號'
               THISFORM.GRID1.COLUMN4.CONTROLSOURCE='ITT.F05'
               THISFORM.GRID1.COLUMN4.WIDTH=INT_015*50
               THISFORM.GRID1.COLUMN5.HEADER1.CAPTION='廠商簡稱'
               THISFORM.GRID1.COLUMN5.WIDTH=INT_015*60
               THISFORM.GRID1.COLUMN5.CONTROLSOURCE='D01.F04'         
               THISFORM.GRID1.COLUMN6.HEADER1.CAPTION='採購單號'
               THISFORM.GRID1.COLUMN6.CONTROLSOURCE='ITT.F01'
               THISFORM.GRID1.COLUMN6.WIDTH=INT_015*80
               THISFORM.GRID1.COLUMN7.HEADER1.CAPTION='進貨數量'
               THISFORM.GRID1.COLUMN7.CONTROLSOURCE='ITT.F04'
               THISFORM.GRID1.COLUMN7.WIDTH=INT_015*100               
               THISFORM.GRID1.COLUMN8.HEADER1.CAPTION='剩餘天數'
               THISFORM.GRID1.COLUMN8.CONTROLSOURCE='ITT.F03-DATE()'
               THISFORM.GRID1.COLUMN8.WIDTH=INT_015*100         
               THISFORM.GRID1.SETFOCUS                   

       ENDCASE 
       JK=THISFORM.GRID1.COLUMN1.HEADER1.CAPTION
       THISFORM.GRID1.REFRESH
       THISFORM.GRID1.SETFOCUS
  ENDPROC                  
      PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS XColIndex     
         FCH=THISFORM.ACTIVECONTROL.NAME  
         CHK=''         
         THISFORM.REFRESH
      ENDPROC       
enddefine     
*************************************************************列印
PROCEDURE PNT_PRC  
    DO CASE 
       CASE KEY()='F05+DTOS(F03)'            
            REPORT FORM ALLTRIM(INT_116)+'D10A' PREVIEW    
       CASE KEY()='F01+F02_A'            
            REPORT FORM ALLTRIM(INT_116)+'D10B' PREVIEW                
       CASE KEY()='F02_A+DTOS(F03)'            
            REPORT FORM ALLTRIM(INT_116)+'D10C' PREVIEW     
    ENDCASE            
ENDPROC       