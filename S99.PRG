**程式名稱:S99.薪資資料移除作業
SET EXCL OFF
CLOSE ALL
CLEAR 
PRG_NO='S99'
IF !USED('A02')		&&權限設定檔
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'S20'    
IF !USED('A23')		&&月份檔
   SELE 0
   USE A23 INDE A23
ELSE
   SELE A23
ENDIF
SELECT A23
GO TOP
FIRST_DATE=LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2) &&系統月份檔第一筆開帳月份
GO BOTTOM
LAST_DATE=LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)  &&系統月份檔最後一筆開帳月份
*******
S99FORM=CREATEOBJECT("TKS99")
S99FORM.SHOW  
DEFINE CLASS TKS99 AS FORM
  CAPTION='S99.薪資資料移除作業'
*!*	  AUTOCENTER=.T.
  CONTROLBOX=.F.
  BORDERSTYLE=1  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  CONTROLCOUNT=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*550
  WIDTH=INT_015*791
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKS99'
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*131,;
      TOP=INT_015*340,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      CAPTION='\<C.選擇',;
      TOOLTIPTEXT='選擇此作業畫面!快速鍵->ALT+A'
      NAME='CMND1'        
  PROCEDURE  INIT       
      v_scrx=SYSMETRIC(1)
      v_scry=SYSMETRIC(2)
      DO CASE
           CASE  v_scrx=640 .and. v_scry=480 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=800 .and. v_scry=600 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=1024 .and. v_scry=768
  	              THISFORM.PICTURE='BMPS\XP1024768.JPG' 
           OTHERWISE
     	              THISFORM.PICTURE='BMPS\XP1024768.JPG' 
      ENDCASE  
 ENDPROC
 PROCEDURE ACTIVATE
       THISFORM.CMND1.SETFOCUS
       THISFORM.CMND1.CLICK	
 ENDPROC
 PROCEDURE CMND1.CLICK  
      TKS99_TEMP=CREATEOBJECT("TKS99_TEMP")  
      TKS99_TEMP.SHOW    
      THISFORM.RELEASE      
      IF !USED('A04')
           SELE 0
           USE A04 INDE A04
      ELSE
           SELE A04
      ENDIF   
      SET ORDER TO 1    
      SEEK SYS_OPER
      IF FOUND()
          REPL F02 WITH ''
          REPL F03 WITH ''
          REPL F04 WITH ''
      ENDIF   
      UNLOCK ALL                  
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
ENDPROC                  
ENDDEFINE
*******
DEFINE CLASS TKS99_TEMP AS FORM
*!*	  AUTOCENTER=.T.
  CAPTION='S99.薪資資料移除作業'
  FONTSIZE=INT_015*9
  TOP=INT_015*100
  LEFT=INT_015*215  
  HEIGHT=INT_015*320
  WIDTH=INT_015*385
  MOVABLE=.F.
  MAXBUTTON=.F.
  MINBUTTON=.F.  
  CONTROLBOX=.F.
  BORDERSTYLE=1
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKS99_TEMP'
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*10,;
      TOP=INT_015*20,;
      AUTOSIZE=.T.,;
      BACKCOLOR=RGB(255,255,0),;
      CAPTION='請選擇欲移除資料之作業項目(注意!此作業將會刪除選取之資料)'  
  ADD OBJECT CHK_S05 AS CHECKBOX WITH;  
      LEFT=INT_015*20,;
      TOP=INT_015*60,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S05.人事薪資基本檔'   
  ADD OBJECT CHK_S06 AS CHECKBOX WITH;  
      LEFT=INT_015*200,;
      TOP=INT_015*60,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S06.離職人員薪資基本檔'         
  ADD OBJECT CHK_S07 AS CHECKBOX WITH;  
      LEFT=INT_015*20,;
      TOP=INT_015*85,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S07.調薪修改作業'   
  ADD OBJECT CHK_S08 AS CHECKBOX WITH;  
      LEFT=INT_015*200,;
      TOP=INT_015*85,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S08.懲處作業'            
  ADD OBJECT CHK_S09 AS CHECKBOX WITH;  
      LEFT=INT_015*20,;
      TOP=INT_015*110,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S09.其他所得發放作業'   
  ADD OBJECT CHK_S10 AS CHECKBOX WITH;  
      LEFT=INT_015*200,;
      TOP=INT_015*110,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S10.薪資計算作業'          
  ADD OBJECT CHK_S11 AS CHECKBOX WITH;  
      LEFT=INT_015*20,;
      TOP=INT_015*135,;
      FONTSIZE=INT_015*10,;
      AUTOSIZE=.T.,;
      CAPTION='S11.薪資月報表'   
*!*	  ADD OBJECT CHK_S13 AS CHECKBOX WITH;  
*!*	      LEFT=INT_015*200,;
*!*	      TOP=INT_015*135,;
*!*	      FONTSIZE=INT_015*10,;
*!*	      AUTOSIZE=.T.,;
*!*	      CAPTION='S13.缺勤報表查詢'           
      *******
  ADD OBJECT LBLF02 AS LABEL WITH;
      LEFT=INT_015*90,;
      TOP=INT_015*200,;
      AUTOSIZE=.T.,;
      BACKCOLOR=RGB(255,255,0),;
      CAPTION='請選擇欲移除資料之西元日期範圍'     
  ADD OBJECT LBLF03 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*240,;
     WIDTH=INT_015*75,;
      CAPTION='起始西元年月'           
  ADD OBJECT TXTF03 AS TEXTBOX WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*234,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*55,;
      MAXLENGTH=7,;
      INPUTMASK='9999/99'
      NAME='TXTF03'              
  ADD OBJECT LBLF04 AS LABEL WITH;
      LEFT=INT_015*220,;
      TOP=INT_015*240,;
     WIDTH=INT_015*75,;
      CAPTION='截止西元年月'           
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      LEFT=INT_015*300,;
      TOP=INT_015*234,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*55,;
      MAXLENGTH=7,;
      INPUTMASK='9999/99'
      NAME='TXTF04'               
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*105,;
      TOP=INT_015*275,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<Y.執行',;
      TOOLTIPTEXT='執行所選取之內容->ALT+Y',;
      NAME='CMND1'
 ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*250,;
      TOP=INT_015*275,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;
      FONTSIZE=INT_015*10,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='離開此作業!快速鍵->ALT+X'
      NAME='CMND2'
********************
 PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*10,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*10,'CHECKBOX')
         THISFORM.SETALL('VALUE',0,'CHECKBOX') 
         THISFORM.SETALL('FONTSIZE',INT_015*10,'COMMANDBUTTON')  
         THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
         THISFORM.TXTF03.VALUE=FIRST_DATE
         THISFORM.TXTF04.VALUE=LAST_DATE
         THISFORM.CMND2.SETFOCUS
 ENDPROC   
 ***
 PROCEDURE CHK_S05.INTERACTIVECHANGE
    DO CASE
           CASE (THIS.VALUE=1 OR THISFORM.CHK_S06.VALUE=1) AND; 
                      THISFORM.CHK_S07.VALUE=0 AND THISFORM.CHK_S08.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')   
                      THISFORM.CMND2.SETFOCUS       
           CASE (THIS.VALUE=0 OR THISFORM.CHK_S06.VALUE=0) AND; 
                      THISFORM.CHK_S07.VALUE=0 AND THISFORM.CHK_S08.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=1 OR THISFORM.CHK_S06.VALUE=1 OR;
                      THISFORM.CHK_S07.VALUE=1 OR THISFORM.CHK_S08.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')  
                      THISFORM.TXTF03.SETFOCUS      
           CASE THIS.VALUE=0 OR THISFORM.CHK_S06.VALUE=0 OR;
                      THISFORM.CHK_S07.VALUE=1 OR THISFORM.CHK_S08.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')      
                      THISFORM.TXTF03.SETFOCUS                                    
    ENDCASE   
 ENDPROC     
 ***
 PROCEDURE CHK_S06.INTERACTIVECHANGE
    DO CASE
           CASE (THIS.VALUE=1 OR THISFORM.CHK_S05.VALUE=1) AND; 
                      THISFORM.CHK_S07.VALUE=0 AND THISFORM.CHK_S08.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')          
                      THISFORM.CMND2.SETFOCUS
           CASE (THIS.VALUE=0 OR THISFORM.CHK_S05.VALUE=0) AND; 
                      THISFORM.CHK_S07.VALUE=0 AND THISFORM.CHK_S08.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=1 OR THISFORM.CHK_S05.VALUE=1 OR;
                      THISFORM.CHK_S07.VALUE=1 OR THISFORM.CHK_S08.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')     
                      THISFORM.TXTF03.SETFOCUS   
           CASE THIS.VALUE=0 OR THISFORM.CHK_S05.VALUE=0 OR;
                      THISFORM.CHK_S07.VALUE=1 OR THISFORM.CHK_S08.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')  
                      THISFORM.TXTF03.SETFOCUS                                        
    ENDCASE   
 ENDPROC  
 ***
 PROCEDURE CHK_S07.INTERACTIVECHANGE
    DO CASE
           CASE THIS.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
                      THISFORM.TXTF03.SETFOCUS
           CASE THIS.VALUE=0 AND THISFORM.CHK_S08.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=0 OR THISFORM.CHK_S08.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')    
                      THISFORM.TXTF03.SETFOCUS                  
    ENDCASE   
 ENDPROC    
 ***
 PROCEDURE CHK_S08.INTERACTIVECHANGE
    DO CASE
           CASE THIS.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
                      THISFORM.TXTF03.SETFOCUS
           CASE THIS.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND;
                      THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=0 OR THISFORM.CHK_S07.VALUE=1 OR;
                      THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX') 
                      THISFORM.TXTF03.SETFOCUS                     
    ENDCASE   
 ENDPROC   
 ***
 PROCEDURE CHK_S09.INTERACTIVECHANGE
    DO CASE
           CASE THIS.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
                      THISFORM.TXTF03.SETFOCUS
           CASE THIS.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND;
                      THISFORM.CHK_S08.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=0 OR THISFORM.CHK_S07.VALUE=1 OR;
                      THISFORM.CHK_S08.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX') 
                      THISFORM.TXTF03.SETFOCUS                     
    ENDCASE   
 ENDPROC   
 ***
 PROCEDURE CHK_S10.INTERACTIVECHANGE
    DO CASE
           CASE THIS.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
                      THISFORM.TXTF03.SETFOCUS
           CASE THIS.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND;
                      THISFORM.CHK_S08.VALUE=0 AND THISFORM.CHK_S09.VALUE=0 AND;
                      THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=0 OR THISFORM.CHK_S07.VALUE=1 OR;
                      THISFORM.CHK_S08.VALUE=1 OR THISFORM.CHK_S09.VALUE=1 OR;
                      THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')  
                      THISFORM.TXTF03.SETFOCUS                    
    ENDCASE   
 ENDPROC    
 ***
 PROCEDURE CHK_S11.INTERACTIVECHANGE
    DO CASE
           CASE THIS.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
                      THISFORM.TXTF03.SETFOCUS
           CASE THIS.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND;
                      THISFORM.CHK_S08.VALUE=0 AND THISFORM.CHK_S09.VALUE=0 AND;
                      THISFORM.CHK_S10.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0
                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
                      THISFORM.CMND2.SETFOCUS
           CASE THIS.VALUE=0 OR THISFORM.CHK_S07.VALUE=1 OR;
                      THISFORM.CHK_S08.VALUE=1 OR THISFORM.CHK_S09.VALUE=1 OR;
                      THISFORM.CHK_S10.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')        
                      THISFORM.TXTF03.SETFOCUS              
    ENDCASE   
 ENDPROC    
*!*	 ***
*!*	 PROCEDURE CHK_S13.INTERACTIVECHANGE
*!*	    DO CASE
*!*	           CASE THIS.VALUE=1
*!*	                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
*!*	                      THISFORM.TXTF03.SETFOCUS
*!*	           CASE THIS.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND;
*!*	                      THISFORM.CHK_S08.VALUE=0 AND THISFORM.CHK_S09.VALUE=0 AND;
*!*	                      THISFORM.CHK_S10.VALUE=0 AND THISFORM.CHK_S11.VALUE=0
*!*	                      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')
*!*	                      THISFORM.CMND2.SETFOCUS
*!*	           CASE THIS.VALUE=0 OR THISFORM.CHK_S07.VALUE=1 OR;
*!*	                      THISFORM.CHK_S08.VALUE=1 OR THISFORM.CHK_S09.VALUE=1 OR;
*!*	                      THISFORM.CHK_S10.VALUE=1 OR THISFORM.CHK_S11.VALUE=1
*!*	                      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')      
*!*	                      THISFORM.TXTF03.SETFOCUS                
*!*	    ENDCASE   
*!*	 ENDPROC     
********確認鍵******
PROCEDURE CMND1.CLICK  
      IF  THISFORM.CHK_S05.VALUE=0 AND THISFORM.CHK_S06.VALUE=0 AND THISFORM.CHK_S07.VALUE=0 AND ;
           THISFORM.CHK_S08.VALUE=0 AND THISFORM.CHK_S09.VALUE=0 AND THISFORM.CHK_S10.VALUE=0 AND ;
           THISFORM.CHK_S11.VALUE=0 &&AND THISFORM.CHK_S13.VALUE=0 
           =MESSAGEBOX('最少須選取一項作業項目,請重新輸入!',0+64,'提示視窗')
    	   THISFORM.CMND2.SETFOCUS
    	   RETURN
      ENDIF   
      IF THISFORM.CHK_S07.VALUE=1 OR THISFORM.CHK_S08.VALUE=1 OR;
          THISFORM.CHK_S09.VALUE=1 OR THISFORM.CHK_S10.VALUE=1 OR;
          THISFORM.CHK_S11.VALUE=1 &&OR THISFORM.CHK_S13.VALUE=1
	  IF  LEN(ALLTRIM(THISFORM.TXTF03.VALUE))=1 
	       =MESSAGEBOX('起始西元年月不得為空白,請重新輸入!',0+64,'提示視窗')
	       THISFORM.TXTF03.SETFOCUS
	        RETURN
	  ENDIF       
	  IF  LEN(ALLTRIM(THISFORM.TXTF03.VALUE))<>7  OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF03.VALUE),2))<=0 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF03.VALUE),2))>=13
	        =MESSAGEBOX('起始西元年月格式不正確,請重新輸入!',0+64,'提示視窗')
	       THISFORM.TXTF03.SETFOCUS
	        RETURN
	  ENDIF             
	  IF  LEN(ALLTRIM(THISFORM.TXTF04.VALUE))=1 
	        =MESSAGEBOX('截止西元年月不得為空白,請重新輸入!',0+64,'提示視窗')
	    	THISFORM.TXTF04.SETFOCUS
	    	RETURN
	  ENDIF       
	  IF  LEN(ALLTRIM(THISFORM.TXTF04.VALUE))<>7  OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF04.VALUE),2))<=0 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF04.VALUE),2))>=13
	        =MESSAGEBOX('截止西元年月格式不正確,請重新輸入!',0+64,'提示視窗')
	 	THISFORM.TXTF04.SETFOCUS
	    	RETURN
	  ENDIF   
	  IF CTOD(THISFORM.TXTF03.VALUE+'/01')>CTOD(THISFORM.TXTF04.VALUE+'/01')
	        =MESSAGEBOX('起始西元年月不得大於截止西元年月,請重新輸入!',0+64,'薪資系統日期範圍：'+FIRST_DATE+' ～ '+LAST_DATE)
	       THISFORM.TXTF03.SETFOCUS
	        RETURN	  
	  ENDIF	
	  IF CTOD(THISFORM.TXTF04.VALUE+'/01')<CTOD(THISFORM.TXTF03.VALUE+'/01')
	        =MESSAGEBOX('截止西元年月不得小於起始西元年月,請重新輸入!',0+64,'薪資系統日期範圍：'+FIRST_DATE+' ～ '+LAST_DATE)
	       THISFORM.TXTF04.SETFOCUS
	        RETURN	  
	  ENDIF		 
	  IF INDEXSEEK(DTOS(CTOD(THISFORM.TXTF03.VALUE+'/01')),.F.,'A23','A231')=.F.
	        =MESSAGEBOX('起始西元年月在薪資系統中並不存在,請重新輸入!',0+64,'薪資系統日期範圍：'+FIRST_DATE+' ～ '+LAST_DATE)
	       THISFORM.TXTF03.SETFOCUS
	        RETURN	  
	  ENDIF	  
	  IF INDEXSEEK(DTOS(CTOD(THISFORM.TXTF04.VALUE+'/01')),.F.,'A23','A231')=.F.
	        =MESSAGEBOX('截止西元年月在薪資系統中並不存在,請重新輸入!',0+64,'薪資系統日期範圍：'+FIRST_DATE+' ～ '+LAST_DATE)
	       THISFORM.TXTF04.SETFOCUS
	        RETURN	  
	  ENDIF		   	  
      ENDIF          
      ***
  IF MESSAGEBOX('是否確定要移除所選擇的作業項目',4+32+256,'請確認')=6 
      IF THISFORM.CHK_S05.VALUE=1
	   IF !USED('S06')   &&薪資項目及扣款項目
		SELECT  0
		USE S06 
	   ELSE
	        SELECT  S06
	   ENDIF	  
	   IF !USED('S05')  &&人事薪資基本檔
	       SELECT  0
	       USE S05 
	   ELSE
	       SELECT  S05
	   ENDIF
           SELECT  RECNO() FROM S05 WHERE F28='A' OR F28='B' OR F28='C' OR F28='2' OR F28='4' OR  F28='9' INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S05')
               LL1=RLOCK(KK1,'S05')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S05
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I)   
                       DELETE FROM S06 WHERE ALLTRIM(S06.F01)=ALLTRIM(S05.F01)
                       SELECT S05                      
                       DELETE 	
                ENDFOR      
           ENDIF  
           SELECT S06
           USE
           SELECT S05
           USE            
           RELEASE ALL LIKE BK*   
           =MESSAGEBOX('S05.人事薪資基本檔已刪除成功!',0+64,'提示視窗')
           THISFORM.CMND2.SETFOCUS
      ENDIF
      ***
      IF THISFORM.CHK_S06.VALUE=1
	   IF !USED('S06')   &&薪資項目及扣款項目
		SELECT  0
		USE S06 
	   ELSE
	        SELECT  S06
	   ENDIF	  
	   IF !USED('S05')  &&人事薪資基本檔
	       SELECT  0
	       USE S05 
	   ELSE
	       SELECT  S05
	   ENDIF
           SELECT  RECNO() FROM S05 WHERE F28='1'  OR F28='3'  OR F28='7' OR F28='8'  INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S05')
               LL1=RLOCK(KK1,'S05')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S05
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I) 
                       DELETE FROM S06 WHERE ALLTRIM(S06.F01)=ALLTRIM(S05.F01)
                       SELECT S05                              
                       DELETE 	
                ENDFOR      
           ENDIF 
           SELECT S06
           USE
           SELECT S05
           USE      
           RELEASE ALL LIKE BK*   
           =MESSAGEBOX('S06.離職人員薪資基本檔已刪除成功!',0+64,'提示視窗')
           THISFORM.CMND2.SETFOCUS
      ENDIF
      ***
      IF THISFORM.CHK_S07.VALUE=1
	   IF !USED('S07')  &&調薪修改作業
	       SELECT  0
	       USE S07 
	   ELSE
	       SELECT  S07
	   ENDIF
           SELECT  RECNO() FROM S07 WHERE F05>=CTOD(THISFORM.TXTF03.VALUE+'/01')  AND  F05<=CTOD(THISFORM.TXTF04.VALUE+'/01')   INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S07')
               LL1=RLOCK(KK1,'S07')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S07
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I) 
                       DELETE 	
                ENDFOR      
           ENDIF 
           SELECT S07
           USE      
           RELEASE ALL LIKE BK*    
           =MESSAGEBOX('S07.調薪修改作業已刪除成功!',0+64,'提示視窗')     
           THISFORM.CMND2.SETFOCUS
      ENDIF
      ***
      IF THISFORM.CHK_S08.VALUE=1
	   IF !USED('S08')  &&懲處作業資料
	       SELECT  0
	       USE S08 
	   ELSE
	       SELECT  S08
	   ENDIF
           SELECT  RECNO() FROM S08 WHERE F05>=CTOD(THISFORM.TXTF03.VALUE+'/01')  AND  F05<=CTOD(THISFORM.TXTF04.VALUE+'/01')   INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S08')
               LL1=RLOCK(KK1,'S08')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S08
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I) 
                       DELETE 	
                ENDFOR      
           ENDIF 
           SELECT S08
           USE      
           RELEASE ALL LIKE BK*    
           =MESSAGEBOX('S08.懲處作業已刪除成功!',0+64,'提示視窗')  
           THISFORM.CMND2.SETFOCUS      
      ENDIF
      ***
      IF THISFORM.CHK_S09.VALUE=1
	   IF !USED('S09')  &&其他所得發放作業
	       SELECT  0
	       USE S09 
	   ELSE
	       SELECT  S09
	   ENDIF
           SELECT  RECNO() FROM S09 WHERE F05>=CTOD(THISFORM.TXTF03.VALUE+'/01')  AND  F05<=CTOD(THISFORM.TXTF04.VALUE+'/01')   INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S09')
               LL1=RLOCK(KK1,'S09')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S09
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I) 
                       DELETE 	
                ENDFOR      
           ENDIF 
           SELECT S09
           USE      
           RELEASE ALL LIKE BK*   
           =MESSAGEBOX('S09.其他所得發放作業已刪除成功!',0+64,'提示視窗')   
           THISFORM.CMND2.SETFOCUS          
      ENDIF
      ***
      IF THISFORM.CHK_S10.VALUE=1
      	   S1=LEFT(DTOS(CTOD(THISFORM.TXTF03.VALUE+'/01')),6)	&&起始西元年月
	   H1='S10'+ALLTRIM(S1)
	   F1=ALLTRIM(H1)+'.DBF'
	   S2=LEFT(DTOS(CTOD(THISFORM.TXTF04.VALUE+'/01')),6)	&&截止西元年月
	   H2='S10'+ALLTRIM(S2)
	   F2=ALLTRIM(H2)+'.DBF'
	   DO WHILE  FILE('&F1')=.T.   
	    	  IF !USED('&F1')
	   	      SELECT  0
	              USE &F1 
	          ELSE
	              SELECT  &F1
	          ENDIF     
	    	  SELECT &H1
		  DELETE ALL
		  USE
		  IF  H1=H2
	   	       EXIT 
	          ENDIF  
	          IF  VAL(RIGHT(S1,2))+1=13
	   	        S1=ALLTRIM(STR(VAL(LEFT(S1,4))+1)+'01')
	   	        H1='S10'+ALLTRIM(S1)
	                F1=ALLTRIM(H1)+'.DBF'  
	          ELSE
	                S1=ALLTRIM(STR(VAL(S1)+1))
	                H1='S10'+ALLTRIM(S1)
	                F1=ALLTRIM(H1)+'.DBF'  
	          ENDIF    
	    ENDDO  
	    =MESSAGEBOX('S10.薪資計算作業已刪除成功!',0+64,'提示視窗')   
	    THISFORM.CMND2.SETFOCUS 
      ENDIF
      ***
      IF THISFORM.CHK_S11.VALUE=1
	   IF !USED('S11')  &&薪資月報表
	       SELECT  0
	       USE S11 
	   ELSE
	       SELECT  S11
	   ENDIF
           SELECT  RECNO() FROM S11 WHERE F01>=CTOD(THISFORM.TXTF03.VALUE+'/01')  AND  F01<=CTOD(THISFORM.TXTF04.VALUE+'/01')   INTO ARRAY BK1             
           JJ1=''                
           IF _TALLY>0   
               FOR I= 1 TO ALEN(BK1)
                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
               ENDFOR    
               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
               RLOCK(KK1,'S11')
               LL1=RLOCK(KK1,'S11')
           ELSE
               LL1=.T.   
           ENDIF                   
           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
               SELECT S11
                FOR I= 1 TO ALEN(BK1)                      
                       GO BK1(I) 
                       DELETE 	
                ENDFOR      
           ENDIF 
           SELECT S11
           USE      
           RELEASE ALL LIKE BK*   
           =MESSAGEBOX('S11.薪資月報表已刪除成功!',0+64,'提示視窗')  
           THISFORM.CMND2.SETFOCUS       
      ENDIF
      ***
*!*	      IF THISFORM.CHK_S13.VALUE=1
*!*		   IF !USED('S0A')  &&年度缺勤天數累計
*!*		       SELECT  0
*!*		       USE S0A
*!*		   ELSE
*!*		       SELECT  S0A
*!*		   ENDIF
*!*	           SELECT  RECNO() FROM S0A WHERE F08>=CTOD(THISFORM.TXTF03.VALUE+'/01')  AND  F08<=CTOD(THISFORM.TXTF04.VALUE+'/01')   INTO ARRAY BK1             
*!*	           JJ1=''                
*!*	           IF _TALLY>0   
*!*	               FOR I= 1 TO ALEN(BK1)
*!*	                       JJ1=JJ1+ALLTRIM(STR(BK1(I)))+','
*!*	               ENDFOR    
*!*	               KK1=STUFF(JJ1,RAT(',',JJ1,1),1,'')              
*!*	               RLOCK(KK1,'S0A')
*!*	               LL1=RLOCK(KK1,'S0A')
*!*	           ELSE
*!*	               LL1=.T.   
*!*	           ENDIF                   
*!*	           IF LL1=.T. AND LEN(ALLTRIM(JJ1))>0  	 	    
*!*	               SELECT S0A
*!*	                FOR I= 1 TO ALEN(BK1)                      
*!*	                       GO BK1(I) 
*!*	                       DELETE 	
*!*	                ENDFOR      
*!*	           ENDIF 
*!*	           SELECT S0A
*!*	           USE      
*!*	           RELEASE ALL LIKE BK*      
*!*	           =MESSAGEBOX('S13.缺勤報表查詢已刪除成功!',0+64,'提示視窗')  
*!*	           THISFORM.CMND2.SETFOCUS     
*!*	      ENDIF
   ELSE
       THISFORM.CMND2.SETFOCUS  
   ENDIF                                                  
ENDPROC       
********離開鍵***********
 PROCEDURE CMND2.CLICK     
     CLOSE TABLE ALL
     THISFORM.RELEASE 
     RETURN  
 ENDPROC     
 ***
 PROC CMND1.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND1.TOP=INT_015*276
             THISFORM.CMND1.LEFT=INT_015*106
  ENDPROC     
  PROC CMND1.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND1.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND1.TOP=INT_015*275
             THISFORM.CMND1.LEFT=INT_015*105                
  ENDPROC 
 PROC CMND2.MOUSEENTER      
       LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND2.TOP=INT_015*276
             THISFORM.CMND2.LEFT=INT_015*251
  ENDPROC     
  PROC CMND2.MOUSELEAVE      
        LPARAMETERS nButton, nShift, nXCoord, nYCoord 
             THISFORM.CMND2.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND2.TOP=INT_015*275
             THISFORM.CMND2.LEFT=INT_015*250
 ENDPROC                                              
ENDDEFINE     