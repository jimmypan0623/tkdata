********A30.資料庫處理作業
CLOSE ALL
CLEAR 
IF !USED('A02')
   SELECT  0
   USE A02 INDE A02
ELSE
   SELECT  A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'A30'
IF !USED('A20')
    SELECT  0
    USE A20 INDE A20
ELSE 
    SELECT  A20
ENDIF
SET ORDER TO A20   
IF !USED('A21')
    SELECT  0
    USE A21 INDE A21
ELSE 
    SELECT  A21
ENDIF
SET ORDER TO A211
A30FORM=CREATEOBJECT("TKA30")
A30FORM.SHOW  
DEFINE CLASS TKA30 AS FORM
  CAPTION='A30.資料庫處理作業'
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
  NAME='TKA30' 
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*131,;
      TOP=INT_015*340,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*40,;
      FONTSIZE=INT_015*9,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      CAPTION='\<C.選擇',;
      TOOLTIPTEXT='選擇此作業畫面!快速鍵->ALT+A'
      NAME='CMND1'        
  PROCEDURE  INIT       
      v_scrx=SYSMETRIC(1)
      v_scry=SYSMETRIC(2)
      DO CASE
           CASE  v_scrx=640 AND v_scry=480 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=800 AND v_scry=600 
                      THISFORM.PICTURE='BMPS\XP800600.JPG' 
           CASE  v_scrx=1024 AND v_scry=768
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
      TKA30_TEMP=CREATEOBJECT("TKA30_TEMP")  
      TKA30_TEMP.SHOW    
      THISFORM.RELEASE      
      IF !USED('A05')
           SELE 0
           USE A05
      ELSE
           SELE A05
      ENDIF   
      SET ORDER TO 1    
      SEEK SYS_OPER+'A30'
      IF FOUND()
         DELETE 
      ENDIF   
      UNLOCK ALL                  
      CLOSE TABLE ALL
      THISFORM.RELEASE 
      RETURN  
ENDPROC                  
ENDDEFINE
*****
***********
DEFINE CLASS TKA30_TEMP AS FORM
  CAPTION='A30.資料庫處理作業'
*!*	  AUTOCENTER=.T.
  TOP=INT_015*190
  LEFT=INT_015*130
  CONTROLBOX=.F.
  BORDERSTYLE=1  
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.
  CONTROLCOUNT=57
  FONTSIZE=INT_015*9
  HEIGHT=INT_015*190
  WIDTH=INT_015*520
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKA30_TEMP'    
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      VISIBLE=.F.            
  ADD OBJECT PAGEFRAME1 AS PAGEFRAME WITH;
      TOP=INT_015*5,;
      LEFT=INT_015*5,;
      WIDTH=INT_015*510,;
      HEIGHT=INT_015*175,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;        
      PAGECOUNT=5,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.T.,;
      PAGE1.NAME='PAGE1',;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='資料庫重整作業',;            
      PAGE2.NAME='PAGE2',;
      PAGE2.VISIBLE=.T.,;
      PAGE2.CAPTION='資料庫建立作業',;      
      PAGE3.NAME='PAGE3',;
      PAGE3.VISIBLE=.T.,;
      PAGE3.CAPTION='資料庫備份作業',;        
      PAGE4.NAME='PAGE4',;
      PAGE4.VISIBLE=.T.,;
      PAGE4.CAPTION='資料庫匯入作業',;   
      PAGE5.NAME='PAGE5',;
      PAGE5.VISIBLE=.T.,;
      PAGE5.CAPTION='資料庫刪除作業'                   
  ****
 ADD OBJECT OPTIONGROUP_CHOICE AS OPTIONGROUP WITH;
     TOP=INT_015*60,;
     LEFT=INT_015*25,;  
     ALIGNMENT=0,;
     HEIGHT=INT_015*75,;
    WIDTH=INT_015*450,;
     BUTTONCOUNT=2,;
     VALUE=1,;
     OPTION1.CAPTION='資料重整',;
     OPTION1.FONTSIZE=INT_015*11,;  
     OPTION1.TOP=INT_015*7,;
     OPTION1.AUTOSIZE=.T.,;
     OPTION1.LEFT=INT_015*5,;
     OPTION1.NAME='OPTION1',;
     OPTION2.CAPTION='索引重整',;
     OPTION2.FONTSIZE=INT_015*11,; 	   
     OPTION2.TOP=INT_015*50,;
     OPTION2.AUTOSIZE=.T.,;
     OPTION2.LEFT=INT_015*5,;
     OPTION2.NAME='OPTION2'    	         
  ADD OBJECT LBLF01 AS LABEL WITH;
      LEFT=INT_015*135,;
      TOP=INT_015*91,;
      AUTOSIZE=.T.,;
      CAPTION='資料重整之西元年月範圍'              
  ADD OBJECT LBLF011 AS LABEL WITH;
      LEFT=INT_015*392,;
      TOP=INT_015*91,;
      AUTOSIZE=.T.,;
      CAPTION='～'                 
  ADD OBJECT TXTF01 AS TEXTBOX WITH;
      LEFT=INT_015*330,;
      TOP=INT_015*87,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*55,;
      MAXLENGTH=7,;
      INPUTMASK='9999/99'
      NAME='TXTF01'      
  ADD OBJECT TXTF02 AS TEXTBOX WITH;
      LEFT=INT_015*415,;
      TOP=INT_015*87,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*55,;
      MAXLENGTH=7,;
      INPUTMASK='9999/99'
      NAME='TXTF02'   
*****      
  ADD OBJECT LBLF04 AS LABEL WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*30,;
      TOP=INT_015*55,;
      AUTOSIZE=.T.,;
      CAPTION='請選擇資料備份之目的路徑'           
  ADD OBJECT TXTF04 AS TEXTBOX WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*40,;
      TOP=INT_015*80,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*350,;
      NAME='TXTF04'       
  ADD OBJECT CMND5 AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*393,;
      TOP=INT_015*80,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*50,;   
      CAPTION='\<S.選取',;
      TOOLTIPTEXT='選取路徑!快速鍵->ALT+S'
      NAME='CMND5'                   
  *******      
  ADD OBJECT LBLF031 AS LABEL WITH;
      LEFT=INT_015*30,;
      TOP=INT_015*142,;
      AUTOSIZE=.T.,;
      CAPTION=''     
  ADD OBJECT LBLF032 AS LABEL WITH;
      LEFT=INT_015*30,;
      TOP=INT_015*165,;
      AUTOSIZE=.T.,;
      CAPTION=''         
  ****        
  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*360,;
      TOP=INT_015*145,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;   
      CAPTION='\<Y.執行',;
      TOOLTIPTEXT='執行所選取之作業!快速鍵->ALT+X'
      NAME='CMND3'                                                               
  ADD OBJECT CMND4 AS COMMANDBUTTON WITH;
      LEFT=INT_015*435,;
      TOP=INT_015*145,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;   
      CAPTION='\<X.離開',;
      TOOLTIPTEXT='取消此作業畫面!快速鍵->ALT+X'
      NAME='CMND4'        
   ********   此按鈕不顯示
  ADD OBJECT CMND_TEMP AS COMMANDBUTTON WITH;
      VISIBLE=.F.,;
      LEFT=INT_015*10,;
      TOP=INT_015*305,;
      HEIGHT=INT_015*30,;
      WIDTH=INT_015*65,;   
      CAPTION='\<T.重整',;
      TOOLTIPTEXT='重整畫面!快速鍵->ALT+T'
      NAME='CMND_TEMP '       
  *******       
 PROCEDURE INIT     
       THISFORM.PAGEFRAME1.PAGE1.FONTSIZE=INT_015*10
       THISFORM.PAGEFRAME1.PAGE2.FONTSIZE=INT_015*10  
       THISFORM.PAGEFRAME1.PAGE3.FONTSIZE=INT_015*10
       THISFORM.PAGEFRAME1.PAGE4.FONTSIZE=INT_015*10 
       THISFORM.PAGEFRAME1.PAGE5.FONTSIZE=INT_015*10   
       THISFORM.PAGEFRAME1.PAGE1.ENABLED=IIF(A02.F08='*',.T.,.F.)  
       THISFORM.PAGEFRAME1.PAGE2.ENABLED=IIF(A02.F09='*',.T.,.F.)  
       THISFORM.PAGEFRAME1.PAGE3.ENABLED=IIF(A02.F10='*',.T.,.F.)  
       THISFORM.PAGEFRAME1.PAGE4.ENABLED=IIF(A02.F11='*',.T.,.F.)  
       THISFORM.PAGEFRAME1.PAGE5.ENABLED=IIF(A02.F12='*',.T.,.F.)  
       DO CASE
             CASE THISFORM.PAGEFRAME1.PAGE1.ENABLED=.T.
		        THISFORM.PAGEFRAME1.PAGE1.SETFOCUS   
		        THISFORM.PAGEFRAME1.PAGE1.ACTIVATE   		                   
             CASE THISFORM.PAGEFRAME1.PAGE2.ENABLED=.T.
		        THISFORM.PAGEFRAME1.PAGE2.SETFOCUS  
		        THISFORM.PAGEFRAME1.PAGE2.ACTIVATE  		                          
             CASE THISFORM.PAGEFRAME1.PAGE3.ENABLED=.T.
		        THISFORM.PAGEFRAME1.PAGE3.SETFOCUS   
		        THISFORM.PAGEFRAME1.PAGE3.ACTIVATE    		                         
             CASE THISFORM.PAGEFRAME1.PAGE4.ENABLED=.T.  
		        THISFORM.PAGEFRAME1.PAGE4.SETFOCUS
		        THISFORM.PAGEFRAME1.PAGE4.ACTIVATE                   
             CASE THISFORM.PAGEFRAME1.PAGE5.ENABLED=.T.
		        THISFORM.PAGEFRAME1.PAGE5.SETFOCUS   
		        THISFORM.PAGEFRAME1.PAGE5.ACTIVATE  		        
       ENDCASE          
       THISFORM.SETALL('FONTSIZE',INT_015*11,'LABEL')
       THISFORM.SETALL('FONTSIZE',INT_015*10,'TEXTBOX')
       THISFORM.SETALL('HEIGHT',INT_015*25,'TEXTBOX')
       THISFORM.SETALL('FONTSIZE',INT_015*11,'COMMANDBUTTON')  
       THISFORM.SETALL('MOUSEPOINTER',99,'COMMANDBUTTON')  
       THISFORM.SETALL('MOUSEICON','BMPS\harrow.cur','COMMANDBUTTON')                                
       THISFORM.TXTF01.VALUE=''
       THISFORM.TXTF02.VALUE=''
 ENDPROC  
 ******
 PROCEDURE PAGEFRAME1.PAGE1.ACTIVATE
       THISFORM.PAGEFRAME1.PAGE1.FONTITALIC=.T.
       THISFORM.PAGEFRAME1.PAGE1.FORECOLOR=RGB(0,0,255)
       THISFORM.PAGEFRAME1.PAGE2.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE2.FORECOLOR=RGB(0,0,0)  
       THISFORM.PAGEFRAME1.PAGE3.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE3.FORECOLOR=RGB(0,0,0)      
       THISFORM.PAGEFRAME1.PAGE4.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE4.FORECOLOR=RGB(0,0,0)    
       THISFORM.PAGEFRAME1.PAGE5.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE5.FORECOLOR=RGB(0,0,0)  
       THISFORM.SETALL('VISIBLE',.T.,'LABEL')
       THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX')     
       THISFORM.SETALL('VALUE','','TEXTBOX')   
       THISFORM.SETALL('READONLY',.F.,'TEXTBOX')              
       THISFORM.LBLF01.LEFT=INT_015*135
       THISFORM.LBLF01.TOP=INT_015*91
       THISFORM.LBLF011.LEFT=INT_015*392
       THISFORM.LBLF011.TOP=INT_015*91       
       THISFORM.TXTF01.LEFT=INT_015*330
       THISFORM.TXTF01.TOP=INT_015*87        
       THISFORM.TXTF02.LEFT=INT_015*415
       THISFORM.TXTF02.TOP=INT_015*87           
       THISFORM.LBLF04.VISIBLE=.F.    
       THISFORM.TXTF04.VISIBLE=.F.  
       THISFORM.CMND5.VISIBLE=.F.  
       THISFORM.OPTIONGROUP_CHOICE.VISIBLE=.T.
       THISFORM.OPTIONGROUP_CHOICE.OPTION1.CAPTION='資料重整'        
       THISFORM.OPTIONGROUP_CHOICE.OPTION2.CAPTION='索引重整'  
       THISFORM.OPTIONGROUP_CHOICE.VALUE=1
       THISFORM.LBLF01.CAPTION='資料重整之西元年月範圍' 
       THISFORM.TXTF01.SETFOCUS   
 ENDPROC 
 ***
 PROCEDURE PAGEFRAME1.PAGE2.ACTIVATE
       THISFORM.PAGEFRAME1.PAGE1.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE1.FORECOLOR=RGB(0,0,0)
       THISFORM.PAGEFRAME1.PAGE2.FONTITALIC=.T.
       THISFORM.PAGEFRAME1.PAGE2.FORECOLOR=RGB(0,0,255)  
       THISFORM.PAGEFRAME1.PAGE3.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE3.FORECOLOR=RGB(0,0,0)      
       THISFORM.PAGEFRAME1.PAGE4.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE4.FORECOLOR=RGB(0,0,0)    
       THISFORM.PAGEFRAME1.PAGE5.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE5.FORECOLOR=RGB(0,0,0)  
       THISFORM.SETALL('VISIBLE',.F.,'LABEL')
       THISFORM.SETALL('VISIBLE',.F.,'TEXTBOX')
       THISFORM.SETALL('VALUE','','TEXTBOX')   
       THISFORM.SETALL('READONLY',.F.,'TEXTBOX')    
       THISFORM.LBLF01.LEFT=INT_015*135
       THISFORM.LBLF01.TOP=INT_015*91
       THISFORM.LBLF011.LEFT=INT_015*392
       THISFORM.LBLF011.TOP=INT_015*91       
       THISFORM.TXTF01.LEFT=INT_015*330
       THISFORM.TXTF01.TOP=INT_015*87        
       THISFORM.TXTF02.LEFT=INT_015*415
       THISFORM.TXTF02.TOP=INT_015*87            
       THISFORM.LBLF04.VISIBLE=.F.    
       THISFORM.TXTF04.VISIBLE=.F.  
       THISFORM.CMND5.VISIBLE=.F. 
       THISFORM.OPTIONGROUP_CHOICE.VISIBLE=.T.  
       THISFORM.OPTIONGROUP_CHOICE.OPTION1.CAPTION='建立基本檔'        
       THISFORM.OPTIONGROUP_CHOICE.OPTION2.CAPTION='建立月份檔'  
       THISFORM.OPTIONGROUP_CHOICE.VALUE=1       
       THISFORM.LBLF01.CAPTION='建立月份檔之西元年月範圍'      
 ENDPROC 
 PROCEDURE PAGEFRAME1.PAGE3.ACTIVATE
       THISFORM.PAGEFRAME1.PAGE1.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE1.FORECOLOR=RGB(0,0,0)
       THISFORM.PAGEFRAME1.PAGE2.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE2.FORECOLOR=RGB(0,0,0)  
       THISFORM.PAGEFRAME1.PAGE3.FONTITALIC=.T.
       THISFORM.PAGEFRAME1.PAGE3.FORECOLOR=RGB(0,0,255)      
       THISFORM.PAGEFRAME1.PAGE4.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE4.FORECOLOR=RGB(0,0,0) 
       THISFORM.PAGEFRAME1.PAGE5.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE5.FORECOLOR=RGB(0,0,0)   
       THISFORM.SETALL('VISIBLE',.T.,'LABEL')
       THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX')             
       THISFORM.SETALL('VALUE','','TEXTBOX')          
       THISFORM.OPTIONGROUP_CHOICE.VISIBLE=.F.  
       THISFORM.LBLF01.LEFT=INT_015*30
       THISFORM.LBLF01.TOP=INT_015*116
       THISFORM.LBLF011.LEFT=INT_015*260
       THISFORM.LBLF011.TOP=INT_015*116       
       THISFORM.TXTF01.LEFT=INT_015*200
       THISFORM.TXTF01.TOP=INT_015*110        
       THISFORM.TXTF02.LEFT=INT_015*280
       THISFORM.TXTF02.TOP=INT_015*110          
       THISFORM.LBLF04.VISIBLE=.T.    
       THISFORM.TXTF04.VISIBLE=.T.  
       THISFORM.CMND5.VISIBLE=.T. 
       THISFORM.TXTF04.READONLY=.T. 
       THISFORM.LBLF01.CAPTION='資料備份之西元年月範圍'        
       THISFORM.LBLF04.CAPTION='請選擇資料備份之目的路徑'    
       THISFORM.CMND5.SETFOCUS
 ENDPROC 
 PROCEDURE PAGEFRAME1.PAGE4.ACTIVATE
       THISFORM.PAGEFRAME1.PAGE1.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE1.FORECOLOR=RGB(0,0,0)
       THISFORM.PAGEFRAME1.PAGE2.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE2.FORECOLOR=RGB(0,0,0)  
       THISFORM.PAGEFRAME1.PAGE3.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE3.FORECOLOR=RGB(0,0,0)      
       THISFORM.PAGEFRAME1.PAGE4.FONTITALIC=.T.
       THISFORM.PAGEFRAME1.PAGE4.FORECOLOR=RGB(0,0,255)  
       THISFORM.PAGEFRAME1.PAGE5.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE5.FORECOLOR=RGB(0,0,0)      
       THISFORM.SETALL('VISIBLE',.T.,'LABEL')
       THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX')                  
       THISFORM.SETALL('VALUE','','TEXTBOX')             
       THISFORM.OPTIONGROUP_CHOICE.VISIBLE=.F.  
       THISFORM.LBLF01.LEFT=INT_015*30
       THISFORM.LBLF01.TOP=INT_015*116
       THISFORM.LBLF011.LEFT=INT_015*260
       THISFORM.LBLF011.TOP=INT_015*116       
       THISFORM.TXTF01.LEFT=INT_015*200
       THISFORM.TXTF01.TOP=INT_015*110        
       THISFORM.TXTF02.LEFT=INT_015*280
       THISFORM.TXTF02.TOP=INT_015*110           
       THISFORM.LBLF04.VISIBLE=.T.    
       THISFORM.TXTF04.VISIBLE=.T.  
       THISFORM.CMND5.VISIBLE=.T. 
       THISFORM.TXTF04.READONLY=.T. 
       THISFORM.LBLF01.CAPTION='匯入資料之西元年月範圍'         
       THISFORM.LBLF04.CAPTION='請選擇匯入資料之來源路徑'   
       THISFORM.CMND5.SETFOCUS 
 ENDPROC    
 PROCEDURE PAGEFRAME1.PAGE5.ACTIVATE
       THISFORM.PAGEFRAME1.PAGE1.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE1.FORECOLOR=RGB(0,0,0)
       THISFORM.PAGEFRAME1.PAGE2.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE2.FORECOLOR=RGB(0,0,0)  
       THISFORM.PAGEFRAME1.PAGE3.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE3.FORECOLOR=RGB(0,0,0)      
       THISFORM.PAGEFRAME1.PAGE4.FONTITALIC=.F.
       THISFORM.PAGEFRAME1.PAGE4.FORECOLOR=RGB(0,0,0)  
       THISFORM.PAGEFRAME1.PAGE5.FONTITALIC=.T.
       THISFORM.PAGEFRAME1.PAGE5.FORECOLOR=RGB(0,0,255)              
       THISFORM.SETALL('VISIBLE',.F.,'LABEL')
       THISFORM.SETALL('VISIBLE',.F.,'TEXTBOX')
       THISFORM.SETALL('VALUE','','TEXTBOX') 
       THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
       THISFORM.LBLF01.LEFT=INT_015*135
       THISFORM.LBLF01.TOP=INT_015*91
       THISFORM.LBLF011.LEFT=INT_015*392
       THISFORM.LBLF011.TOP=INT_015*91       
       THISFORM.TXTF01.LEFT=INT_015*330
       THISFORM.TXTF01.TOP=INT_015*87        
       THISFORM.TXTF02.LEFT=INT_015*415
       THISFORM.TXTF02.TOP=INT_015*87               
       THISFORM.LBLF04.VISIBLE=.F.    
       THISFORM.TXTF04.VISIBLE=.F.  
       THISFORM.CMND5.VISIBLE=.F. 
       THISFORM.OPTIONGROUP_CHOICE.VISIBLE=.T.  
       THISFORM.OPTIONGROUP_CHOICE.OPTION1.CAPTION='刪除基本檔'        
       THISFORM.OPTIONGROUP_CHOICE.OPTION2.CAPTION='刪除月份檔'  
       THISFORM.OPTIONGROUP_CHOICE.VALUE=1       
       THISFORM.LBLF01.CAPTION='刪除月份檔之西元年月範圍'      
 ENDPROC   
****
PROCEDURE  OPTIONGROUP_CHOICE.INTERACTIVECHANGE
      DO CASE
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=1
		       THISFORM.SETALL('VISIBLE',.T.,'LABEL')
		       THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX') 
		       THISFORM.LBLF04.VISIBLE=.F.  
		       THISFORM.TXTF04.VISIBLE=.F.      
		       THISFORM.CMND5.VISIBLE=.F.            
                       IF THIS.VALUE=1
            	           THISFORM.LBLF01.CAPTION='資料重整之西元年月範圍'    
            	       ELSE
            	           THISFORM.LBLF01.CAPTION='索引重整之西元年月範圍'  
            	       ENDIF
            	       THISFORM.TXTF01.SETFOCUS
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=2
           	       IF THIS.VALUE=1
			    THISFORM.SETALL('VISIBLE',.F.,'LABEL')
			    THISFORM.SETALL('VISIBLE',.F.,'TEXTBOX')
			    THISFORM.CMND4.SETFOCUS
           	       ELSE
			    THISFORM.SETALL('VISIBLE',.T.,'LABEL')
			    THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX')
			    THISFORM.LBLF04.VISIBLE=.F. 
			    THISFORM.TXTF04.VISIBLE=.F.  
			    THISFORM.CMND5.VISIBLE=.F. 
			    THISFORM.TXTF01.SETFOCUS
		       ENDIF     
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=5
           	       IF THIS.VALUE=1
			    THISFORM.SETALL('VISIBLE',.F.,'LABEL')
			    THISFORM.SETALL('VISIBLE',.F.,'TEXTBOX')
			    THISFORM.CMND4.SETFOCUS
           	       ELSE
			    THISFORM.SETALL('VISIBLE',.T.,'LABEL')
			    THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX')
			    THISFORM.LBLF04.VISIBLE=.F. 
			    THISFORM.TXTF04.VISIBLE=.F.  
			    THISFORM.CMND5.VISIBLE=.F. 
			    THISFORM.TXTF01.SETFOCUS
		       ENDIF               
      ENDCASE        
 ENDPROC   
  ************
  PROCEDURE CMND3.CLICK
       IF THISFORM.PAGEFRAME1.ACTIVEPAGE=1 OR THISFORM.PAGEFRAME1.ACTIVEPAGE=3 OR THISFORM.PAGEFRAME1.ACTIVEPAGE=4;
          OR (THISFORM.PAGEFRAME1.ACTIVEPAGE=2  AND THISFORM.OPTIONGROUP_CHOICE.VALUE=2) OR (THISFORM.PAGEFRAME1.ACTIVEPAGE=5  AND THISFORM.OPTIONGROUP_CHOICE.VALUE=2)
	  IF  LEN(ALLTRIM(THISFORM.TXTF01.VALUE))=1 
		=MESSAGEBOX('起始日期不得為空白,請重新輸入!',0+64,'提示訊息')
		THISFORM.CMND_TEMP.CLICK
		THISFORM.TXTF01.SETFOCUS
		RETURN
	   ENDIF             
	   IF LEN(ALLTRIM((THISFORM.TXTF01.VALUE)))<>7 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF01.VALUE),2))<=0 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF01.VALUE),2))>=13
	       =MESSAGEBOX('起始日期不正確',0+64,'提示訊息')
	       THISFORM.CMND_TEMP.CLICK
	       THISFORM.TXTF01.SETFOCUS
	       RETURN
	   ENDIF
	   IF  LEN(ALLTRIM(THISFORM.TXTF02.VALUE))=1 
	        =MESSAGEBOX('截止日期不得為空白,請重新輸入!',0+64,'提示訊息')
		THISFORM.CMND_TEMP.CLICK
		THISFORM.TXTF02.SETFOCUS
		RETURN
	   ENDIF         
	   IF LEN(ALLTRIM((THISFORM.TXTF02.VALUE)))<>7 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2))<=0 OR VAL(RIGHT(ALLTRIM(THISFORM.TXTF02.VALUE),2))>=13
	       =MESSAGEBOX('截止日期不正確',0+64,'提示訊息')
	       THISFORM.CMND_TEMP.CLICK
	       THISFORM.TXTF02.SETFOCUS
	       RETURN
	   ENDIF 	 
	   IF CTOD(THISFORM.TXTF01.VALUE+'/01')>CTOD(THISFORM.TXTF02.VALUE+'/01')
	       =MESSAGEBOX('起始日期不得大於截止日期,請重新輸入!',0+64,'提示訊息')
	       THISFORM.CMND_TEMP.CLICK
	       THISFORM.TXTF01.SETFOCUS
	       RETURN	  
	   ENDIF	
	   IF CTOD(THISFORM.TXTF02.VALUE+'/01')<CTOD(THISFORM.TXTF01.VALUE+'/01')
	       =MESSAGEBOX('截止日期不得小於起始日期,請重新輸入!',0+64,'提示訊息')
	       THISFORM.CMND_TEMP.CLICK
	       THISFORM.TXTF02.SETFOCUS
	       RETURN	  
	   ENDIF	
      ENDIF
      THISFORM.SETALL('ENABLED',.F.,'COMMANDBUTTON')
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX')    
      THISFORM.OPTIONGROUP_CHOICE.ENABLED=.F.      
      THISFORM.LBLF031.VISIBLE=.T.
      THISFORM.LBLF032.VISIBLE=.T.
      ***匯出所選取之作業項目
      CREATE CURSOR CK_TEMP (F01 C(3))
      SELECT F01 FROM A20  ORDER BY F01 INTO CURSOR A20A
      SELECT A20A
      GO TOP
      DO WHILE !EOF()
             SELECT CK_TEMP
             APPEND BLANK 
             REPLACE CK_TEMP.F01 WITH A20A.F01
             SELECT A20A
             SKIP
      ENDDO
      SELECT A20A
      USE                                                       
   ***
    DO CASE
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=1	&&資料庫重整作業
		       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE3.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE4.ENABLED=.F. 
		       THISFORM.PAGEFRAME1.PAGE5.ENABLED=.F. 
		       IF THISFORM.OPTIONGROUP_CHOICE.VALUE=1
		            IF  MESSAGEBOX('您是否確定要執行資料重整作業?',4+48+256,'請確認') = 6 
			         ***重整基本檔,暫存檔,系統檔
			         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='0' OR A20.F02='9' OR A20.F02 ='z') ORDER BY A20.F01 INTO CURSOR A20_TEMP
			         SELECT A20_TEMP
			         GO TOP
			         DO WHILE !EOF()
			                TableName=ALLTRIM(A20_TEMP.F01)	          
		               	        IF FILE(TableName+'.DBF')=.T.  
		               	            IF !USED('&TableName') 
			                        SELECT  0
			                        USE &TableName  EXCLUSIVE 
			                        PACK
			                        F3=TableName+'.FXP'
			                        F5=TableName+'.BAK'
            						F7=TableName+'.ERR'
			                        IF FILE('&F5')=.T.
			                            DELETE FILE  &F3
			                            DELETE FILE  &F5
			                            DELETE FILE  &F7
			                        ENDIF
			                        THISFORM.LBLF031.CAPTION='重整基本檔：'+TableName+' '+ALLTRIM(A20_TEMP.F07)
			                        SELECT  &TableName
			                        USE
			                     ENDIF    
			                ENDIF     
			                SELECT A20_TEMP
			                SKIP
			         ENDDO
			         SELECT A20_TEMP
			         USE
			         ***重整月份檔
			         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='1' OR A20.F02='3') ORDER BY A20.F01 INTO CURSOR A20_TEMP
			         IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			              DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			         ELSE
			             MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			             DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			         ENDIF  			         
			         DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		                 DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)			         
	                                SELECT A20_TEMP
	                                GO TOP
	                                DO WHILE !EOF()
	                                      TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
		               	              IF FILE(TableName+'.DBF')=.T. 
		               	                  IF !USED('&TableName') 
			                               SELECT  0
			                               USE &TableName  EXCLUSIVE 
			                               PACK
				                       F3=TableName+'.FXP'
				                       F5=TableName+'.BAK'
						       F7=TableName+'.ERR'
				                       IF FILE('&F5')=.T.
				                           DELETE FILE  &F3
				                           DELETE FILE  &F5
				                           DELETE FILE  &F7
				                       ENDIF			                               
			                               THISFORM.LBLF031.CAPTION='重整月份檔：'+TableName+' '+ALLTRIM(A20_TEMP.F07)
			                               SELECT  &TableName
			                               USE
			                          ENDIF    
			                      ENDIF    
	                                      SELECT A20_TEMP
	                                      SKIP
	                                ENDDO
	                                IF (VAL(RIGHT(DateFile,2))+1)>12
			                     DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                ELSE
			                     MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                     DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                ENDIF
			          ENDDO
			          SELECT A20_TEMP
			          USE        			         
			         ***
			         =MESSAGEBOX('資料重整已完成!',0+64,'提示訊息')
			         THISFORM.CMND_TEMP.CLICK
			    ELSE
			         THISFORM.CMND_TEMP.CLICK     
			    ENDIF 
			    **************    
		       ELSE	&&重整索引
		            IF  MESSAGEBOX('您是否確定要執行索引重整作業?',4+48+256,'請確認') = 6 
			         ***重整索引之基本檔,暫存檔,系統檔
			         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='0' OR A20.F02='9' OR A20.F02='z') ORDER BY A20.F01 INTO CURSOR A20_TEMP  
			         SELECT A20_TEMP
			         GO TOP
			         DO WHILE !EOF()
			               	TableName=ALLTRIM(A20_TEMP.F01)	&&資料檔名
			                 SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
					 SELECT A21_TEMP
					 IF _TALLY>0
					    IF FILE(TableName+'.DBF')=.T.
			               	        IF !USED('&TableName') 
						     DELETE FILE TableName+'.CDX'
			                             SELECT  0
			                             USE &TableName  EXCLUSIVE
			                             SELECT A21_TEMP 
			                             GO TOP
						     DO WHILE !EOF()
			                                   THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+'  '+ALLTRIM(A20_TEMP.F07) 
			                                   THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
			                                   IndexName=ALLTRIM(A21_TEMP.F02)	&&索引名稱
			                                   IndexKey=ALLTRIM(A21_TEMP.F03)	&&索引內容
			                                   IF A21_TEMP.F04='1'		&&順向排序
			                                       SELECT &TableName
			                                       INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
			                                   ELSE					&&逆向排序
			                                       SELECT &TableName
			                                       INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
			                                   ENDIF
			                                   SELECT A21_TEMP
			                                   SKIP   
			                              ENDDO  
			                              SELECT &TableName
			                              USE
			                        ENDIF	      
			                    ENDIF
				        ENDIF 
				        SELECT A21_TEMP
				        USE
				        SELECT A20_TEMP
				        SKIP   
			         ENDDO
			         SELECT A20_TEMP
			         USE
			         ***重整索引之月份檔
			         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='1' OR A20.F02='3') ORDER BY A20.F01 INTO CURSOR A20_TEMP
			         IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			              DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			         ELSE
			             MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			             DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			         ENDIF  			         
			         DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		                 DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)			         
	                                SELECT A20_TEMP
	                                GO TOP
	                                DO WHILE !EOF()
	                                      TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                      SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
					      SELECT A21_TEMP
					      IF _TALLY>0	                                      
		               	                 IF FILE(TableName+'.DBF')=.T. 
			               	             IF !USED('&TableName') 
						         DELETE FILE TableName+'.CDX'
			                                 SELECT  0
			                                 USE &TableName  EXCLUSIVE
			                                 SELECT A21_TEMP 
			                                 GO TOP
						         DO WHILE !EOF()	
						                THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
			                                        THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
			                                         IndexName=ALLTRIM(A21_TEMP.F02) + ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6)) 	&&索引名稱
			                                         IndexKey=ALLTRIM(A21_TEMP.F03)		&&索引內容
			                                         IF A21_TEMP.F04='1'	&&順向排序
			                                              SELECT &TableName
			                                              INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
			                                         ELSE					&&逆向排序
			                                              SELECT &TableName
			                                              INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
			                                         ENDIF
			                                         SELECT A21_TEMP
			                                         SKIP           	                 
						         ENDDO
						         SELECT &TableName
						         USE
						     ENDIF
		               	                 ENDIF
		               	              ENDIF
		               	              SELECT A21_TEMP
		               	              USE
		               	              SELECT A20_TEMP
		               	              SKIP  
		               	         ENDDO       
	                                 IF (VAL(RIGHT(DateFile,2))+1)>12
			                     DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                 ELSE
			                     MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                     DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                 ENDIF
			         ENDDO
			         SELECT A20_TEMP
			         USE			         
			          =MESSAGEBOX('索引重整已完成!',0+64,'提示訊息')
			         THISFORM.CMND_TEMP.CLICK
			    ELSE
			         THISFORM.CMND_TEMP.CLICK		         
		            ENDIF
		       ENDIF   
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=2	&&資料庫建立作業
		         THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.  
		         THISFORM.PAGEFRAME1.PAGE3.ENABLED=.F.  
		         THISFORM.PAGEFRAME1.PAGE4.ENABLED=.F. 
		         THISFORM.PAGEFRAME1.PAGE5.ENABLED=.F. 
			     IF MESSAGEBOX('您是否確定要執行資料庫建立作業?',4+48+256,'請確認') = 6 							                        
                    IF THISFORM.OPTIONGROUP_CHOICE.VALUE=1	&&建立基本檔
                       SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='0' OR A20.F02='9' ORDER BY A20.F01 INTO CURSOR A20_TEMP
        		       SELECT A20_TEMP
		               GO TOP
		               DO WHILE !EOF()
		                  TableName=ALLTRIM(A20_TEMP.F01)	&&資料檔名
		                  IF FILE(TableName+'.DBF')=.F.
		                     SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                 ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                 SELECT A22_TEMP
			                 GO TOP
			                 IF _TALLY>0
			                    FieldName=ALLTRIM(A22_TEMP.F01)	&&欄位名稱
			                    TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
			                    CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
			                    SELECT A22_TEMP
			                    SKIP
			                    DO WHILE !EOF()		&&建立欄位
			                       FieldName=ALLTRIM(A22_TEMP.F01)
			                       TypeName=ALLTRIM(A22_TEMP.F02)
			                       ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
			                       SELECT A22_TEMP
			                       SKIP
			                    ENDDO 
			                    SELECT A22_TEMP
			                    USE
			                    SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
			                    SELECT A21_TEMP
			                    GO TOP
			                    DO WHILE !EOF()		&&建立索引
			                       THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+'  '+ALLTRIM(A20_TEMP.F07) 
			                       THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
			                       IndexName=ALLTRIM(A21_TEMP.F02)	&&索引名稱
			                       IndexKey=ALLTRIM(A21_TEMP.F03)	        &&索引內容
			                       IF A21_TEMP.F04='1'		&&順向排序
			                          SELECT &TableName
			                          INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
			                       ELSE					&&逆向排序
			                          SELECT &TableName
			                          INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
			                       ENDIF   
			                       SELECT A21_TEMP
			                       SKIP
			                    ENDDO
			                    SELECT &TableName
			                    USE
			                    F5=TableName+'.BAK'
			                    IF FILE('&F5')=.T.
			                       DELETE FILE  &F5
			                    ENDIF
			                    SELECT A21_TEMP
			                    USE 
			                 ENDIF
		                  ENDIF  
		                  SELECT A20_TEMP
		                  SKIP
                       ENDDO
		               SELECT A20_TEMP
		               USE       
		               		                IF !USED('C00')
		                    SELECT 0
		                    USE C00
		                ELSE
		                   SELECT C00
		                ENDIF   
		                SET ORDER TO 1
		                SEEK INT_011
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH INT_011
		                   REPLACE F02 WITH 1
		                   REPLACE F03 WITH DTCX(DATE())+' SET UP'
		                ENDIF
		                USE		                
		                IF !USED('D00')
		                   SELECT 0
		                   USE D00
		                ELSE
		                   SELECT D00
		                ENDIF   
		                SET ORDER TO 1
		                SEEK INT_011
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH INT_011
		                   REPLACE F02 WITH 1
		                   REPLACE F03 WITH DTCX(DATE())+'  SET UP'
		                ENDIF          
		                USE		                		                 
		                IF !USED('A09')
		                   SELECT 0
		                   USE A09
		                ELSE
		                   SELECT A09
		                ENDIF
		                SET ORDER TO 1
		                SEEK 'C000 '
		                IF !FOUND()
		                   APPEND BLANK
		                   REPLACE F01 WITH 'C000'
		                   REPLACE F06 WITH '客戶基本主檔編碼'
		                   REPLACE F08 WITH 'C0000' 
		                ENDIF
		                SEEK 'D000 '
		                IF !FOUND()                    
		                   APPEND BLANK
		                   REPLACE F01 WITH 'D000'
		                   REPLACE F06 WITH '廠商基本主檔編碼'
		                   REPLACE F08 WITH 'V0000'            
		                ENDIF
		                SEEK 'E000 '
		                IF !FOUND()         
		                   APPEND BLANK
		                   REPLACE F01 WITH 'E000'
		                   REPLACE F06 WITH '協力廠商基本主檔編碼'
		                   REPLACE F08 WITH 'W0000'                  		                   
		                ENDIF  
		                USE 
		               =MESSAGEBOX('基本檔已建立完畢!',0+64,'提示訊息視窗')
		               THISFORM.CMND_TEMP.CLICK
		                ****
                    ELSE		&&建立月份檔
			           IF FILE('A23.DBF')<>.T.
                          =MESSAGEBOX('請先建立基本檔後在建立月份檔!',0+64,'提示訊息視窗') 
                          THISFORM.CMND_TEMP.CLICK    
      			       ELSE     
      			          IF !USED('A23')
			                 SELECT  0
			                 USE A23 INDE A23
			              ELSE 
			                 SELECT  A23
			              ENDIF
			              SET ORDER TO A231
			              GO TOP
			              IF CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01')<A23.F01 AND !EMPTY(A23.F01)	
                             =MESSAGEBOX('系統月份檔之起始日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故不得小於此日期,請重新輸入!',0+64,'提示訊息視窗') 
                             THISFORM.CMND_TEMP.CLICK
				             THISFORM.TXTF01.SETFOCUS
				             RETURN	 			                   
			              ENDIF
			              ***			             
			              SELECT A23
			              GO BOTTOM 
			              IF MONTH(A23.F01) = 12
			                 IF YEAR(A23.F01) + 1 <> YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))
			                    =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次年1月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                    THISFORM.CMND_TEMP.CLICK
				                THISFORM.TXTF01.SETFOCUS
				                RETURN	
				             ELSE
				                IF MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01')) <> 1
			                       =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次年1月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                       THISFORM.CMND_TEMP.CLICK
				                   THISFORM.TXTF01.SETFOCUS
				                   RETURN					                   
				                ENDIF
			                 ENDIF 			                 
			              ELSE
			                 IF YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))=YEAR(A23.F01) AND !EMPTY(A23.F01)
			                    IF (MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))-MONTH(A23.F01))<>1
			                        =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                        THISFORM.CMND_TEMP.CLICK
				                    THISFORM.TXTF01.SETFOCUS
				                    RETURN	
			                    ENDIF
			                 ELSE
			                    IF YEAR(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))=(YEAR(A23.F01)+1) AND !EMPTY(A23.F01)
			                       IF (MONTH(CTOD(ALLTRIM(THISFORM.TXTF01.VALUE)+'/01'))<>1 AND MONTH(A23.F01)<>12)=.F.
			                          =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                          THISFORM.CMND_TEMP.CLICK
				                      THISFORM.TXTF01.SETFOCUS
				                      RETURN	
			                       ENDIF		                    
			                    ELSE
			                       IF !EMPTY(A23.F01)
			                          =MESSAGEBOX('系統月份檔之最後建立日期為'+LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)+'故建立日期需從次月開始,請重新輸入!',0+64,'提示訊息視窗') 
			                          THISFORM.CMND_TEMP.CLICK
				                      THISFORM.TXTF01.SETFOCUS
				                      RETURN
				                   ENDIF    
			                    ENDIF
			                 ENDIF   
			              ENDIF	             
			              *****
                          SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='1'  ORDER BY A20.F01 INTO CURSOR A20_TEMP		      
			              IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			                 DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			              ELSE
			                 MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			                 DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			              ENDIF   			      
			              DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		                  DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)
			                 SELECT A23
			                 SEEK DTOS(CTOD(DateFile+'/01'))
			                 IF !FOUND()
			                    APPEND BLANK
			                    REPLACE  F01 WITH CTOD(DateFile+'/01')
			                 ENDIF		
			                 ***	            
			                 SELECT A20_TEMP
			                 GO TOP
			                 DO WHILE !EOF()
			                    TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                    IF FILE(TableName+'.DBF')=.F.
			                       SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                       ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                       SELECT A22_TEMP
				                   GO TOP
				                   IF _TALLY>0
				                      FieldName=ALLTRIM(A22_TEMP.F01)    &&欄位名稱
				                      TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
				                      CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
				                      SELECT A22_TEMP
				                      SKIP
				                      DO WHILE !EOF()		&&建立欄位
				                         FieldName=ALLTRIM(A22_TEMP.F01)
				                         TypeName=ALLTRIM(A22_TEMP.F02)
				                         ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
				                         SELECT A22_TEMP
				                         SKIP
				                      ENDDO 
				                      SELECT A22_TEMP
				                      USE
				                      ***
				                      SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
				                      SELECT A21_TEMP
				                      GO TOP
				                      DO WHILE !EOF()		&&建立索引
				                         THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
				                         THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
				                         IndexName=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))		&&索引名稱
				                         IndexKey=ALLTRIM(A21_TEMP.F03)	       &&索引內容
				                         IF A21_TEMP.F04='1'		&&順向排序
				                            SELECT &TableName
				                            INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
				                         ELSE						&&逆向排序
				                            SELECT &TableName
				                            INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
				                         ENDIF   
				                         SELECT A21_TEMP
				                         SKIP
				                      ENDDO
				                      ***帶入各項暫存檔至該月份檔
							          IF  LEFT(TableName,3) = 'C13'  AND FILE('C1C.DBF') = .T.	&&應收帳款對帳單檔
							              IF !USED('C1C') 
								             SELE 0
								             USE C1C
								          ELSE
								             SELE C1C
								          ENDIF
							              SELECT &TableName
							              APPEND FROM C1C
							              SELECT C1C
							              DELETE ALL
							              USE
							          ENDIF
							          IF LEFT(TableName,3) = 'D19'  AND FILE('D1I.DBF') = .T.	&&應付帳款對帳單檔
							             IF !USED('D1I') 
								           SELE 0
								           USE D1I
								         ELSE
								           SELE D1I
								         ENDIF
							             SELECT &TableName
							             APPEND FROM D1I
							             SELECT D1I
							             DELETE ALL
							             USE
							          ENDIF	
							          IF LEFT(TableName,3) = 'P19'  AND FILE('P1I.DBF') = .T.	&&外購應付帳款對帳單檔
							             IF !USED('P1I') 
								            SELE 0
								            USE P1I
								         ELSE
								            SELE P1I
								         ENDIF
							             SELECT &TableName
							             APPEND FROM P1I
							             SELECT P1I
							             DELETE ALL
							             USE
							          ENDIF
							          IF LEFT(TableName,3) = 'K24'  AND FILE('K2D.DBF') = .T. 	&&發票表頭檔
							             IF !USED('K2D')
								            SELE 0
								            USE K2D
								         ELSE
								            SELE K2D
								         ENDIF
							             SELECT &TableName
							             APPEND FROM K2D
							             SELECT K2D
							             DELETE ALL
							             USE
							          ENDIF
							          IF LEFT(TableName,3) = 'K25' AND FILE('K2E.DBF') = .T.	&&發票明細檔
							             IF !USED('K2E') 
								            SELE 0
								            USE K2E
								         ELSE
								            SELE K2E
								         ENDIF
							             SELECT &TableName
							             APPEND FROM K2E
							             SELECT K2E
							             DELETE ALL
							             USE
							          ENDIF			
							          	       							       							       						       
				                      ***
				                      SELECT &TableName
				                      USE
					                  F5=TableName+'.BAK'
					                  IF FILE('&F5')=.T.
					                     DELETE FILE  &F5
					                  ENDIF				                              
				                      SELECT A21_TEMP
				                      USE 
				                   ENDIF
			                    ENDIF  
			                    SELECT A20_TEMP
			                    SKIP
	                         ENDDO
	                         IF (VAL(RIGHT(DateFile,2))+1)>12
			                    DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                 ELSE
			                    MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                    DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                 ENDIF  
			              ENDDO
			              SELECT A20_TEMP
			              USE 
***********************************************建立奇月檔
                         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='3'  ORDER BY A20.F01 INTO CURSOR A20_TEMP		      
			             IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			                 DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			             ELSE
			                 MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			                 DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			             ENDIF   	
			             DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
			             DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)     			                
		                    SELECT A20_TEMP
			                GO TOP
			                DO WHILE !EOF()
			                   IF VAL(ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6)))%2<>0 
			                      TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                   
			                      IF FILE(TableName+'.DBF')=.F.
			                         SELECT  A22.F01,A22.F02,A22.F03,A22.F04,A22.F05 FROM A22 WHERE ALLTRIM(A22.F05)=ALLTRIM(A20_TEMP.F01);
			                         ORDER BY A22.F05,A22.F01 INTO CURSOR A22_TEMP
			                         SELECT A22_TEMP
				                     GO TOP
				                     IF _TALLY>0
				                        FieldName=ALLTRIM(A22_TEMP.F01)    &&欄位名稱
				                        TypeName=ALLTRIM(A22_TEMP.F02)    &&欄位型態
				                        CREATE TABLE &TableName (&FieldName &TypeName(A22_TEMP.F03,A22_TEMP.F04)) 
				                        SELECT A22_TEMP
				                        SKIP
				                        DO WHILE !EOF()		&&建立欄位
				                           FieldName=ALLTRIM(A22_TEMP.F01)
				                           TypeName=ALLTRIM(A22_TEMP.F02)
				                           ALTER TABLE  &TableName ADD &FieldName  &TypeName(A22_TEMP.F03,A22_TEMP.F04)
				                           SELECT A22_TEMP
				                           SKIP
				                        ENDDO 
				                        SELECT A22_TEMP
				                        USE
				                        ***
				                        SELECT A21.F02,A21.F03,A21.F04 FROM A21 WHERE ALLTRIM(A21.F01)=ALLTRIM(A20_TEMP.F01) ORDER BY A21.F01 INTO CURSOR A21_TEMP
				                        SELECT A21_TEMP
				                        GO TOP
				                        DO WHILE !EOF()		&&建立索引
				                           THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
				                           THISFORM.LBLF032.CAPTION=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A21_TEMP.F03)+' '+IIF(A21_TEMP.F04='1','順向排序','逆向排序')
				                           IndexName=ALLTRIM(A21_TEMP.F02)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))		&&索引名稱
				                           IndexKey=ALLTRIM(A21_TEMP.F03)	       &&索引內容
				                           IF A21_TEMP.F04='1'		&&順向排序
				                              SELECT &TableName
				                              INDEX ON  &IndexKey TAG &IndexName  ASCENDING 
				                           ELSE						&&逆向排序
				                              SELECT &TableName
				                              INDEX ON  &IndexKey TAG &IndexName  DESCENDING 
				                           ENDIF   
				                           SELECT A21_TEMP
				                           SKIP
				                        ENDDO			  				                        
			                            SELECT &TableName
				                        USE
					                    F5=TableName+'.BAK'
					                    IF FILE('&F5')=.T.
					                       DELETE FILE  &F5
					                    ENDIF				                              
				                        SELECT A21_TEMP
				                        USE 				                        
				                     ENDIF  
				                  ENDIF
				               ENDIF   	 
			                   SELECT A20_TEMP
			                   SKIP
	                        ENDDO	                            	
	                        IF (VAL(RIGHT(DateFile,2))+1)>12
			                   DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                ELSE
			                   MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                   DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                ENDIF  
			             ENDDO
			             SELECT A20_TEMP
			             USE                   			                	                
			               
***************************************************			              
			              SELECT A23
			              USE     				              		              
			              =MESSAGEBOX('月份檔已建立完畢!',0+64,'提示訊息視窗') 
			              THISFORM.CMND_TEMP.CLICK  				      
			            ENDIF
			       **********
                     ENDIF       
                 ELSE
                     THISFORM.CMND_TEMP.CLICK
                 ENDIF    
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=3	&&資料庫備份作業
		       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE4.ENABLED=.F. 
		       THISFORM.PAGEFRAME1.PAGE5.ENABLED=.F. 	
		       IF MESSAGEBOX('您是否確定要執行資料庫備份作業?',4+48+256,'請確認') = 6 		
		            IF EMPTY(ALLTRIM(THISFORM.TXTF04.VALUE))
			         =MESSAGEBOX('路徑位址不得為空白,請重新輸入!',0+64,'提示訊息')
			         THISFORM.CMND_TEMP.CLICK
		                 THISFORM.CMND5.SETFOCUS
		                 RETURN
			     ENDIF  			
			     **備份基本檔,暫存檔,系統檔
			     SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='0' OR A20.F02='9' OR A20.F02='z')  ORDER BY A20.F01 INTO CURSOR A20_TEMP
			     SELECT A20_TEMP
			     GO TOP
			     DO WHILE !EOF()
			            TableName=ALLTRIM(A20_TEMP.F01)	&&資料檔名
			            IF FILE(TableName+'.DBF')=.T. 
			                IF !USED('&TableName')
			                    SELECT 0
			                    USE &TableName  
			                ELSE
			                    SELECT &TableName
			                ENDIF 
			                THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+'  '+ALLTRIM(A20_TEMP.F07) 
			                COPY TO ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF'
			                SELECT &TableName
			                USE
			            ENDIF
			            SELECT A20_TEMP
			            SKIP
			     ENDDO
			     SELECT A20_TEMP
			     USE		     
			     ***備份月份檔
			     SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='1'  ORDER BY A20.F01 INTO CURSOR A20_TEMP
			     IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			          DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			     ELSE
			          MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			          DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			     ENDIF   			      
			     DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		             DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)
			            SELECT A20_TEMP
			            GO TOP
			            DO WHILE !EOF()
			                  TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
			                  IF FILE(TableName+'.DBF')=.T.
     					      IF !USED('&TableName')
			                          SELECT 0
			                          USE &TableName  
			                      ELSE
			                          SELECT &TableName
			                      ENDIF 
			                      THISFORM.LBLF031.CAPTION=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))+'  '+ALLTRIM(A20_TEMP.F07) 
			                      COPY TO ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF'
			                      SELECT &TableName
			                      USE
			                  ENDIF 
			                  SELECT A20_TEMP
			                  SKIP
			            ENDDO     
	                            IF (VAL(RIGHT(DateFile,2))+1)>12
			                 DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			            ELSE
			                 MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                 DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			            ENDIF 				     
			     ENDDO
			     SELECT A20_TEMP
			     USE
		              =MESSAGEBOX('資料已備份成功!',0+64,'提示訊息視窗') 
			     THISFORM.CMND_TEMP.CLICK  				     
		       ELSE
		             THISFORM.CMND_TEMP.CLICK
		       ENDIF                       
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=4	&&資料庫匯入作業
		       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE3.ENABLED=.F. 
		       THISFORM.PAGEFRAME1.PAGE5.ENABLED=.F.   
		       IF MESSAGEBOX('您是否確定要執行資料庫匯入作業?',4+48+256,'請確認') = 6
		          IF EMPTY(ALLTRIM(THISFORM.TXTF04.VALUE))
			         =MESSAGEBOX('路徑位址不得為空白,請重新輸入!',0+64,'提示訊息')
			         THISFORM.CMND_TEMP.CLICK
		             THISFORM.CMND5.SETFOCUS
		             RETURN
			      ENDIF 
			     ****匯入系統檔A01.個人編號資料,A02.個人授權記錄,A14.分部門設定資料檔,A26.系統變數設定檔
			     IF FILE('A01.DBF') = .T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+'A01.DBF') = .T.
			         IF !USED('A01')
			              SELECT  0
			              USE A01
			         ELSE 
			              SELECT  A01
			         ENDIF	
			         THISFORM.LBLF031.CAPTION='匯入系統檔：A01.個人編號資料'
			         APPEND FROM  ALLTRIM(THISFORM.TXTF04.VALUE)+'A01.DBF'
			     ENDIF
			     IF FILE('A02.DBF') = .T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+'A02.DBF') = .T.
			         IF !USED('A02')
			              SELECT  0
			              USE A02
			         ELSE 
			              SELECT  A02
			         ENDIF	
			         THISFORM.LBLF031.CAPTION='匯入系統檔：A02.個人授權記錄'
			         APPEND FROM  ALLTRIM(THISFORM.TXTF04.VALUE)+'A02.DBF'		         
			     ENDIF
			     IF FILE('A14.DBF') = .T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+'A14.DBF') = .T.
			         IF !USED('A14')
			              SELECT  0
			              USE A14
			         ELSE 
			              SELECT  A14
			         ENDIF	
			         THISFORM.LBLF031.CAPTION='匯入系統檔：A14.分部門設定資料檔'
			         APPEND FROM  ALLTRIM(THISFORM.TXTF04.VALUE)+'A14.DBF'
			     ENDIF			     			     
			     IF FILE('A26.DBF') = .T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+'A26.DBF') = .T.
			         IF !USED(ALLTRIM(THISFORM.TXTF04.VALUE)+'A26.DBF')
			            SELECT 0
			            USE ALLTRIM(THISFORM.TXTF04.VALUE)+'A26.DBF'
			            SELECT A26.* FROM A26 ORDER BY A26.F01 INTO CURSOR A26_TEMP
			            SELECT A26
			            USE  
			            IF !USED('A26')
			               SELECT  0
			               USE A26
			            ELSE 
			               SELECT  A26
			            ENDIF
			            SELECT A26
			            SET ORDER TO 1
			            THISFORM.LBLF031.CAPTION='匯入系統檔：A26.系統參數設定'
			            SELECT A26_TEMP
			            GO TOP
			            DO WHILE !EOF()
			               SELECT A26
			               SEEK A26_TEMP.F01
			               IF FOUND()
			                  REPLACE A26.F06 WITH A26_TEMP.F06
			               ELSE
			                  SELECT A26
			                  APPEND BLANK
			                  REPLACE A26.F01 WITH A26_TEMP.F01
			                  REPLACE A26.F02 WITH A26_TEMP.F02
			                  REPLACE A26.F03 WITH A26_TEMP.F03
			                  REPLACE A26.F04 WITH A26_TEMP.F04
			                  REPLACE A26.F05 WITH A26_TEMP.F05
			                  REPLACE A26.F06 WITH A26_TEMP.F06
			                  REPLACE A26.F07 WITH A26_TEMP.F07
			                  REPLACE A26.F08 WITH A26_TEMP.F08
			                  REPLACE A26.F09 WITH A26_TEMP.F09
			                  REPLACE A26.F10 WITH A26_TEMP.F10
			               ENDIF
			               SELECT A26_TEMP
			               SKIP
			            ENDDO
			            SELECT A26_TEMP
			            USE
			         ENDIF   
			     ENDIF			     
			     ***A23.系統月份工作檔
			     IF FILE('A23.DBF') = .T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+'A23.DBF') = .T.
			         IF !USED(ALLTRIM(THISFORM.TXTF04.VALUE)+'A23.DBF')
			              SELECT  0
			              USE ALLTRIM(THISFORM.TXTF04.VALUE)+'A23'
			              SELECT * FROM A23 ORDER BY F01 INTO CURSOR A23_TEMP
			              SELECT A23
			              USE
			         ELSE
			              CREATE CURSOR A23_TEMP;
  				     (F01 D(8),F02 C(1),F03 C(1),F04 N(1),F05 N(1),F06 N(1),F07 L(1),F08 L(1),F09 L(1),F10 N(1),F11 L(1),F12 L(1),F13 L(1),F14 L(1))
			         ENDIF				         
			         IF !USED('A23')
			            SELECT  0
			            USE A23
			         ELSE 
			            SELECT  A23
			         ENDIF	
			         SET ORDER TO A231 && DTOS(F01)
			         THISFORM.LBLF031.CAPTION='匯入基本檔：A23.系統月份工作檔'
			         SELECT A23_TEMP
			         GO TOP
			         DO WHILE !EOF()
			            SELECT A23
			            SEEK DTOS(A23_TEMP.F01)
			            IF FOUND()
			               REPLACE A23.F02 WITH A23_TEMP.F02
			               REPLACE A23.F03 WITH A23_TEMP.F03
			               REPLACE A23.F04 WITH A23_TEMP.F04
			               REPLACE A23.F05 WITH A23_TEMP.F05
			               REPLACE A23.F06 WITH A23_TEMP.F06
			               REPLACE A23.F07 WITH A23_TEMP.F07
			               REPLACE A23.F08 WITH A23_TEMP.F08
			               REPLACE A23.F09 WITH A23_TEMP.F09
			               REPLACE A23.F10 WITH A23_TEMP.F10
			               REPLACE A23.F11 WITH A23_TEMP.F11
			               REPLACE A23.F12 WITH A23_TEMP.F12
			               REPLACE A23.F13 WITH A23_TEMP.F13
			               REPLACE A23.F14 WITH A23_TEMP.F14
			            ENDIF
			            SELECT A23_TEMP
			            SKIP
			         ENDDO
			         SELECT A23_TEMP
			         USE
			         SELECT A23
			         USE
			     ENDIF			     
			     ****匯入基本檔    		               
			     SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='0' ORDER BY A20.F01 INTO CURSOR A20_TEMP 
			     SELECT A20_TEMP
			     GO TOP
			     DO WHILE !EOF()
			            TableName=ALLTRIM(A20_TEMP.F01)	
			            IF FILE(TableName+'.DBF')=.T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF')=.T.	            
		               	         IF !USED('&TableName')
			                     SELECT  0
			                     USE &TableName 
			                 ELSE 
			                      SELECT  &TableName
			                 ENDIF		                 
			                 IF TableName<>'A23'
			                      SELECT &TableName
			                      DELETE FROM &TableName
			                      THISFORM.LBLF031.CAPTION='匯入基本檔：'+TableName+' '+ALLTRIM(A20_TEMP.F07)  
			                      APPEND FROM  ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF'   
			                 ENDIF
			                 *****
			                 SELECT  &TableName
			                 USE
			            ENDIF    
			            SELECT A20_TEMP
			            SKIP
		             ENDDO 
		             SELECT A20_TEMP
		             USE
		             ****匯入月份檔  
			     SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='1' ORDER BY A20.F01 INTO CURSOR A20_TEMP
			     IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			          DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			     ELSE
			          MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			          DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			     ENDIF  
			     DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		             DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)			           
			            SELECT A20_TEMP
			            GO TOP
			            DO WHILE !EOF()
			                   TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名			
				            IF FILE(TableName+'.DBF')=.T. AND FILE(ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF')=.T.	            
			               	         IF !USED('&TableName')
				                      SELECT  0
				                      USE &TableName 
				                 ELSE 
				                      SELECT  &TableName
				                 ENDIF
			                         SELECT &TableName
			                         DELETE FROM &TableName				                 
				                 THISFORM.LBLF031.CAPTION='匯入月份檔：'+TableName+' '+ALLTRIM(A20_TEMP.F07)				                 
				                 APPEND FROM  ALLTRIM(THISFORM.TXTF04.VALUE)+TableName+'.DBF'
				                 SELECT  &TableName
				                 USE
				             ENDIF
				             SELECT A20_TEMP
				             SKIP
			             ENDDO       			
		                     IF (VAL(RIGHT(DateFile,2))+1)>12
		                         DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
		                     ELSE
		                         MonthNo=MONTH(CTOD(DateFile+'/01'))+1
		                         DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
		                     ENDIF  
		              ENDDO 		        
		             ***       
		              SELECT A20_TEMP
		              USE
		              =MESSAGEBOX('資料庫匯入作業已成功!',0+64,'提示訊息視窗')
		              THISFORM.CMND_TEMP.CLICK		              
		       ELSE  
		              THISFORM.CMND_TEMP.CLICK
		       ENDIF
            CASE THISFORM.PAGEFRAME1.ACTIVEPAGE=5	&&資料庫刪除作業
		       THISFORM.PAGEFRAME1.PAGE1.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE2.ENABLED=.F.  
		       THISFORM.PAGEFRAME1.PAGE3.ENABLED=.F. 
		       THISFORM.PAGEFRAME1.PAGE4.ENABLED=.F.                     
                        IF THISFORM.OPTIONGROUP_CHOICE.VALUE=1	&&刪除基本檔
			     IF MESSAGEBOX('您是否確定要刪除基本檔?',4+48+256,'請確認') = 6 
			         SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND A20.F02='0' ORDER BY A20.F01 INTO CURSOR A20_TEMP
			         SELECT A20_TEMP
			         GO TOP
			         DO WHILE !EOF()
			                TableName=ALLTRIM(A20_TEMP.F01)
			                F1=TableName+'.DBF'
			                F2=TableName+'.CDX'
			                F3=TableName+'.FXP'
			                F4=TableName+'.FPT'
			                F5=TableName+'.BAK'
			                F6=TableName+'.TBK'
			                F7=TableName+'.ERR'
			                IF FILE('&F1')=.T. AND !USED('&F1')
			                    THISFORM.LBLF031.CAPTION=TableName+'  '+ALLTRIM(A20_TEMP.F07)
			                     DELETE FILE  &F1
			                     DELETE FILE  &F2
			                     DELETE FILE  &F3
			                     DELETE FILE  &F4
			                     DELETE FILE  &F5
			                     DELETE FILE  &F6
			                     DELETE FILE  &F7
			                ELSE
			                     THISFORM.LBLF031.CAPTION=TableName+'不存在系統中'
			                ENDIF
			                SELECT A20_TEMP    
			                SKIP
			          ENDDO
			          SELECT A20_TEMP
			          USE
			          =MESSAGEBOX('基本檔已刪除完畢!',0+64,'提示訊息視窗')
			          THISFORM.CMND_TEMP.CLICK
			    ELSE
			          THISFORM.CMND_TEMP.CLICK
			    ENDIF                           
                        ELSE		&&刪除月份檔
			    IF MESSAGEBOX('您是否確定要刪除月份檔?',4+48+256,'請確認') = 6 	
			          IF FILE('A23.DBF')=.T.
			              IF !USED('A23')
			                  SELECT  0
			                  USE A23 INDE A23
			              ELSE 
			                   SELECT  A23
			              ENDIF
			               SET ORDER TO A231 
			           ENDIF 
			           SELECT A20.F01,A20.F07 FROM A20,CK_TEMP WHERE CK_TEMP.F01=A20.F01 AND (A20.F02='1' OR A20.F02='3')  ORDER BY A20.F01 INTO CURSOR A20_TEMP
			           IF (VAL(RIGHT(THISFORM.TXTF02.VALUE,2))+1)>12
			               DateEnd=ALLTRIM(STR(YEAR(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1))+'/01'
			           ELSE
			               MonthNo=MONTH(CTOD(THISFORM.TXTF02.VALUE+'/01'))+1
			               DateEnd=LEFT(THISFORM.TXTF02.VALUE,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			           ENDIF   			      
			           DateFile=ALLTRIM(THISFORM.TXTF01.VALUE)
		                   DO WHILE ALLTRIM(DateEnd)<>ALLTRIM(DateFile)			           
			                  IF FILE('A23.DBF')=.T.
			                       SELECT A23
			                       SEEK DTOS(CTOD(DateFile+'/01'))
			                       IF FOUND()
			                           DELETE 
			                       ENDIF				           	  
			                  ENDIF
			           	  SELECT A20_TEMP
			           	  GO TOP
			                  DO WHILE !EOF()
			                        TableName=ALLTRIM(A20_TEMP.F01)+ALLTRIM(LEFT(DTOS(CTOD(DateFile+'/01')),6))	&&資料檔名
				                F1=TableName+'.DBF'
				                F2=TableName+'.CDX'
				                F3=TableName+'.FXP'
				                F4=TableName+'.FPT'
				                F5=TableName+'.BAK'
				                F6=TableName+'.TBK'
				                F7=TableName+'.ERR'
				                IF FILE('&F1')=.T. AND !USED('&F1')
				                    THISFORM.LBLF031.CAPTION=TableName+'  '+ALLTRIM(A20_TEMP.F07)
				                     DELETE FILE  &F1
				                     DELETE FILE  &F2
				                     DELETE FILE  &F3
				                     DELETE FILE  &F4
				                     DELETE FILE  &F5
				                     DELETE FILE  &F6
				                     DELETE FILE  &F7
				                ELSE
				                     THISFORM.LBLF031.CAPTION=TableName+'不存在系統中'
				                ENDIF
				                SELECT A20_TEMP   
			                        SKIP
			                  ENDDO
			          	  IF (VAL(RIGHT(DateFile,2))+1)>12
			                       DateFile=ALLTRIM(STR(YEAR(CTOD(DateFile+'/01'))+1))+'/01'
			                  ELSE
			                       MonthNo=MONTH(CTOD(DateFile+'/01'))+1
			                       DateFile=LEFT(DateFile,5)+IIF(MonthNo<10,'0'+ALLTRIM(STR(MonthNo)),ALLTRIM(STR(MonthNo)))			             
			                  ENDIF  
			          ENDDO     
			          SELECT A20_TEMP
			          USE
			          IF FILE('A23.DBF')=.T.
			              SELECT A23
			              USE
			          ENDIF			          
			          =MESSAGEBOX('月份檔已刪除完畢!',0+64,'提示訊息視窗')
			          THISFORM.CMND_TEMP.CLICK
			    ELSE
			          THISFORM.CMND_TEMP.CLICK
			    ENDIF  
			    ***********                        
                        ENDIF                
      ENDCASE
      THISFORM.CMND_TEMP.CLICK
  ENDPROC   
  ***    
  PROCEDURE CMND4.CLICK
      THISFORM.RELEASE   
  ENDPROC  
  ****
  PROCEDURE CMND5.CLICK
      IF THISFORM.PAGEFRAME1.ACTIVEPAGE=3	
          THISFORM.TXTF04.VALUE=ALLTRIM(GETDIR(CURDIR(),"請選擇備份資料之目的路徑","選擇路徑",64))
      ELSE
          THISFORM.TXTF04.VALUE=ALLTRIM(GETDIR(CURDIR(),"請選擇匯入資料之來源路徑","選擇路徑",16384))
      ENDIF
  ENDPROC    
******
  PROCEDURE CMND_TEMP.CLICK
      THISFORM.PAGEFRAME1.PAGE1.ENABLED=IIF(A02.F08='*',.T.,.F.)  
      THISFORM.PAGEFRAME1.PAGE2.ENABLED=IIF(A02.F09='*',.T.,.F.)  
      THISFORM.PAGEFRAME1.PAGE3.ENABLED=IIF(A02.F10='*',.T.,.F.)  
      THISFORM.PAGEFRAME1.PAGE4.ENABLED=IIF(A02.F11='*',.T.,.F.)  
      THISFORM.PAGEFRAME1.PAGE5.ENABLED=IIF(A02.F12='*',.T.,.F.)  
      THISFORM.OPTIONGROUP_CHOICE.ENABLED=.T.
      THISFORM.SETALL('READONLY',.F.,'TEXTBOX')
      THISFORM.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.TXTF04.READONLY=.T. 
      THISFORM.TXTF01.VALUE=''  
      THISFORM.TXTF02.VALUE=''  
      THISFORM.LBLF031.CAPTION='' 
      THISFORM.LBLF032.CAPTION=''  
      THISFORM.CMND4.SETFOCUS
  ENDPROC   
  ********     
 PROC CMND3.MOUSEENTER      
       LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND3.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND3.TOP=INT_015*146
             THISFORM.CMND3.LEFT=INT_015*361
  ENDPROC     
  PROC CMND3.MOUSELEAVE      
        LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND3.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND3.TOP=INT_015*145
             THISFORM.CMND3.LEFT=INT_015*360                
  ENDPROC 
 PROC CMND4.MOUSEENTER      
       LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND4.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND4.TOP=INT_015*146
             THISFORM.CMND4.LEFT=INT_015*436
  ENDPROC     
  PROC CMND4.MOUSELEAVE      
        LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND4.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND4.TOP=INT_015*145
             THISFORM.CMND4.LEFT=INT_015*435
 ENDPROC   
 PROC CMND5.MOUSEENTER      
       LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND5.FORECOLOR=RGB(255,0,0)
             THISFORM.CMND5.TOP=INT_015*81
             THISFORM.CMND5.LEFT=INT_015*394
  ENDPROC     
  PROC CMND5.MOUSELEAVE      
        LPARAMETERS NBUTTON, NSHIFT, NXCOORD, NYCOORD 
             THISFORM.CMND5.FORECOLOR=RGB(0,0,0)  
             THISFORM.CMND5.TOP=INT_015*80
             THISFORM.CMND5.LEFT=INT_015*393
 ENDPROC    
ENDDEFINE    
