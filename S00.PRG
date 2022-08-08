**程式名稱:S00.薪資參數設定
CLOSE ALL
CLEAR 
AREA1='S00_1'
AREA2='S00'
FLG='0'
FCH=''
CHK=''
R=0
CHK_NULL=''         &&判別年假設定內容是否為空
IF !USED('A01')	&&操作人員資料設定
   SELE 0
   USE A01
ELSE 
   SELE A01
ENDIF
SET ORDER TO A01      
IF !USED('A02')	&&權限設定檔
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'S00'     
IF !USED('A26')    &&系統參數設定
   SELE 0
   USE A26 INDE A26
ELSE
   SELE A26
ENDIF
SET ORDER TO A26         
IF !USED('S01')	&&薪資資料設定
   SELE 0
   USE S01 
ELSE
   SELE S01
ENDIF
SET ORDER TO S011
IF !USED('S00')	&&薪資參數設定
   SELE 0
   USE S00 
ELSE
   SELE S00
ENDIF
SET ORDER TO S001
*****
CREATE CURSOR S00_1 (F01 C(1),F02 C(30)) 
INDE ON F01 TAG S00_11
SELECT S00_1
APPEND BLANK
REPLACE F01 WITH 'B'
REPLACE F02 WITH '加班費設定'
APPEND BLANK
REPLACE F01 WITH 'C'
REPLACE F02 WITH '遲到設定'
APPEND BLANK
REPLACE F01 WITH 'D'
REPLACE F02 WITH '早退設定'
APPEND BLANK
REPLACE F01 WITH 'E'
REPLACE F02 WITH '年假範圍設定'
APPEND BLANK
REPLACE F01 WITH 'F'
REPLACE F02 WITH '年假限制設定'
APPEND BLANK
REPLACE F01 WITH 'G'
REPLACE F02 WITH '請假扣款設定'
APPEND BLANK
REPLACE F01 WITH 'H'
REPLACE F02 WITH '懲處扣款設定'
APPEND BLANK
REPLACE F01 WITH 'I'
REPLACE F02 WITH '班別設定'
APPEND BLANK
REPLACE F01 WITH 'J'
REPLACE F02 WITH '其他所得發放設定'
APPEND BLANK
REPLACE F01 WITH 'K'
REPLACE F02 WITH '年終獎金發放設定'
APPEND BLANK
REPLACE F01 WITH 'L'
REPLACE F02 WITH '未休年假獎金發放設定'
APPEND BLANK
REPLACE F01 WITH 'X'
REPLACE F02 WITH '公司資料設定'
APPEND BLANK
REPLACE F01 WITH 'Y'
REPLACE F02 WITH '加班參數設定'
APPEND BLANK
REPLACE F01 WITH 'Z'
REPLACE F02 WITH '基本參數設定'
SELECT S00_1
SET ORDER TO S00_11
GO TOP
********
S00FORM=CREATEOBJECT("TKS00")
S00FORM.SHOW  
DEFINE CLASS TKS00 AS FORM
  CAPTION='S00.薪資參數設定'
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
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WIDTH=INT_015*791
  WINDOWTYPE=1
  NAME='TKS00'
  ADD OBJECT SHAPE1 AS SHAPE1 WITH;
      AUTOCENTER=.T.
  ADD OBJECT SHAPE2 AS SHAPE2 WITH;
      AUTOCENTER=.F.
  ADD OBJECT SHAPE3 AS SHAPE3 WITH;
      AUTOCENTER=.T.  
  ADD OBJECT KEY_LIST AS KEY_LIST WITH;
      AUTOCENTER=.T.
  ADD OBJECT GRID1 AS GRID1 WITH;
      TOP=INT_015*10,;
      HEIGHT=INT_015*405,;
      COLUMNCOUNT=1,;            
      RECORDSOURCE='S00_1',;      
      COLUMN1.NAME='COLUMN1'
  ADD OBJECT GRID2 AS GRID2 WITH; 
      TOP=INT_015*283,;
      HEIGHT=INT_015*235,;  
      COLUMNCOUNT=12,;              
      RECORDSOURCE='S00',;
      LINKMASTER='S00_1',;
      RELATIONALEXPR='F01',;
      CHILDORDER='S001',;
      COLUMN1.NAME='COLUMN1',;
      COLUMN2.NAME='COLUMN2',;
      COLUMN3.NAME='COLUMN3',;
      COLUMN4.NAME='COLUMN4',;
      COLUMN5.NAME='COLUMN5',;
      COLUMN6.NAME='COLUMN6',;
      COLUMN7.NAME='COLUMN7',;
      COLUMN8.NAME='COLUMN8',;
      COLUMN9.NAME='COLUMN9',;
      COLUMN10.NAME='COLUMN10',;
      COLUMN11.NAME='COLUMN11',;
      COLUMN12.NAME='COLUMN12'       
  ADD OBJECT CMDGROUP AS CMDGROUP WITH;
      AUTOSIZE=.T.              
  ADD OBJECT ORPGROUP AS ORPGROUP WITH;    
      AUTOSIZE=.T.
  ADD OBJECT ED_BOTT AS COMMANDBUTTON WITH;
      TOP=INT_015*460,;
      LEFT=INT_015*56,;
      HEIGHT=INT_015*25,;
      WIDTH=INT_015*42,;      
      FONTSIZE=INT_015*9,;
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;     
      CAPTION='\<E.修改',;
      NAME='ED_BOTT'              
   ***     
ADD OBJECT PAGEFRAME1 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.T.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='加班費設定'    
ADD OBJECT PAGEFRAME2 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='遲到設定'
ADD OBJECT PAGEFRAME3 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='早退設定'   
ADD OBJECT PAGEFRAME4 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='年假範圍設定'    
ADD OBJECT PAGEFRAME5 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='年假限制設定'    
ADD OBJECT PAGEFRAME6 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='請假扣款設定'    
ADD OBJECT PAGEFRAME7 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='懲處扣款設定'                                                 
ADD OBJECT PAGEFRAME8 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='班別設定'         
ADD OBJECT PAGEFRAME9 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='其他所得發放設定'              
ADD OBJECT PAGEFRAME10 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='年終獎金發放設定'    
ADD OBJECT PAGEFRAME11 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='未休年假獎金發放設定'                
ADD OBJECT PAGEFRAME12 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='公司資料設定'         
ADD OBJECT PAGEFRAME13 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,;
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='加班參數設定'           
ADD OBJECT PAGEFRAME14 AS PAGEFRAME WITH;
      LEFT=INT_015*237,;
      WIDTH=INT_015*540,;
      HEIGHT=INT_015*279,;
      FONTSIZE=INT_015*9,;    
      MOUSEPOINTER=99,; 
      MOUSEICON='BMPS\harrow.cur',;               
      PAGECOUNT=1,;
      TABSTRETCH=1,;
      ACTIVEPAGE=1,;
      VISIBLE=.F.,;
      PAGE1.VISIBLE=.T.,;
      PAGE1.CAPTION='基本參數設定'                 
  ADD OBJECT SHRGROUP AS SHRGROUP WITH;
      TOP=INT_015*200,;
      LEFT=INT_015*600                                  
  *************************設定PAGEFRAME1***************************  加班費設定-B
PROCEDURE PAGEFRAME1.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF05','LABEL') 
     WITH THIS.LBLF05
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .WIDTH=INT_015*50
          .CAPTION='加班編號'          
     ENDWITH
  THIS.ADDOBJECT('TXTF05','TEXTBOX')      
     WITH THIS.TXTF05
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2         
     ENDWITH
  THIS.ADDOBJECT('LBLF06','LABEL') 
     WITH THIS.LBLF06
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*55
          .WIDTH=INT_015*50
          .CAPTION='加班時段'          
     ENDWITH
 THIS.ADDOBJECT('OPTIONGROUP06','OPTIONGROUP')
    WITH THIS.OPTIONGROUP06
           .VISIBLE=.T.
   	   .TOP=INT_015*49
   	   .LEFT=INT_015*71    
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*170
	   .BUTTONCOUNT=2
	   .VALUE=1
	   .ENABLED=.F.
	   .NAME='OPTIONGROUP06'
	   .OPTION1.CAPTION='平日加班'
	   .OPTION1.FONTSIZE=INT_015*9      
	   .OPTION1.TOP=INT_015*5
	   .OPTION1.AUTOSIZE=.T.
	   .OPTION1.LEFT=INT_015*5 
	   .OPTION1.NAME='OPTION1'
	   .OPTION2.CAPTION='假日加班'
	   .OPTION2.FONTSIZE=INT_015*9    
	   .OPTION2.TOP=INT_015*5
	   .OPTION2.AUTOSIZE=.T.
	   .OPTION2.LEFT=INT_015*90
	   .OPTION2.NAME='OPTION2'   	   
     ENDWITH          
  THIS.ADDOBJECT('LBLF07','LABEL') 
     WITH THIS.LBLF07
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*80
          .WIDTH=INT_015*50
          .CAPTION='單        位'          
     ENDWITH      
 THIS.ADDOBJECT('OPTIONGROUP07','OPTIONGROUP')
     WITH THIS.OPTIONGROUP07
       .VISIBLE=.T.
   	   .TOP=INT_015*74
   	   .LEFT=INT_015*71   
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*85
	   .BUTTONCOUNT=2
	   .VALUE=1
	   .ENABLED=.F.
	   .NAME='OPTIONGROUP07'
	   .OPTION1.CAPTION='時'
	   .OPTION1.FONTSIZE=INT_015*9      
	   .OPTION1.TOP=INT_015*5
	   .OPTION1.AUTOSIZE=.T.
	   .OPTION1.LEFT=INT_015*5 
	   .OPTION1.NAME='OPTION1'
	   .OPTION2.CAPTION='日'
	   .OPTION2.FONTSIZE=INT_015*9    
	   .OPTION2.TOP=INT_015*5
	   .OPTION2.AUTOSIZE=.T.
	   .OPTION2.LEFT=INT_015*50
	   .OPTION2.NAME='OPTION2'   	       
     ENDWITH
  THIS.ADDOBJECT('LBLF08','LABEL') 
     WITH THIS.LBLF08
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*107
          .AUTOSIZE=.T.
          .CAPTION='計算方式'          
     ENDWITH  
 THIS.ADDOBJECT('OPTIONGROUP08','OPTIONGROUP08')
     WITH THIS.OPTIONGROUP08
   	   .TOP=INT_015*99
   	   .LEFT=INT_015*71    
     ENDWITH  
 THIS.ADDOBJECT('TXTF08','TEXTBOX')      
     WITH THIS.TXTF08
          .VISIBLE=.T.
          .LEFT=INT_015*142
          .TOP=INT_015*101
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=4                
     ENDWITH                              
THIS.ADDOBJECT('COMBOBOX08','COMBOBOX')
    WITH THIS.COMBOBOX08
           .VISIBLE=.F.
	   .TOP=INT_015*101
	   .LEFT=INT_015*142
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*50
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX08'	   
   ENDWITH   
  THIS.ADDOBJECT('LBLF09','LABEL') 
     WITH THIS.LBLF09
          .VISIBLE=.T.
          .LEFT=INT_015*194
          .TOP=INT_015*107
          .AUTOSIZE=.T.
          .CAPTION=' 計算費率'          
     ENDWITH     
  THIS.ADDOBJECT('TXTF09','TEXTBOX')      
     WITH THIS.TXTF09
          .VISIBLE=.T.
          .LEFT=INT_015*247
          .TOP=INT_015*101
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=6        
	  .INPUTMASK='999.99'            
     ENDWITH    
  THIS.ADDOBJECT('TXTF10','TEXTBOX')      
     WITH THIS.TXTF10
          .VISIBLE=.T.
          .LEFT=INT_015*418
          .TOP=INT_015*101
          .WIDTH=INT_015*60
          .HEIGHT=INT_015*25
          .MAXLENGTH=7        
	  .INPUTMASK='9999999'            
     ENDWITH         
  THIS.ADDOBJECT('LBLF46','LABEL') 
     WITH THIS.LBLF46
          .VISIBLE=.T.
          .LEFT=INT_015*250
          .TOP=INT_015*30
          .WIDTH=INT_015*100
          .CAPTION='每次加班津貼補助'          
     ENDWITH
  THIS.ADDOBJECT('TXTF46','TEXTBOX')      
     WITH THIS.TXTF46
          .VISIBLE=.T.
          .LEFT=INT_015*351
          .TOP=INT_015*24
          .WIDTH=INT_015*140
          .HEIGHT=INT_015*25
          .MAXLENGTH=30         
     ENDWITH      
  THIS.ADDOBJECT('COMBOBOX46','COMBOBOX46') 
     WITH THIS.COMBOBOX46
          .LEFT=INT_015*351
          .TOP=INT_015*24
     ENDWITH         
 THIS.ADDOBJECT('LBLF47','LABEL') 
     WITH THIS.LBLF47
          .VISIBLE=.T.
          .LEFT=INT_015*300
          .TOP=INT_015*55
          .WIDTH=INT_015*50
          .CAPTION='補助金額'          
     ENDWITH
  THIS.ADDOBJECT('TXTF47','TEXTBOX')      
     WITH THIS.TXTF47
          .VISIBLE=.T.
          .LEFT=INT_015*351
          .TOP=INT_015*49
          .WIDTH=INT_015*60
          .HEIGHT=INT_015*25
          .MAXLENGTH=7         
	  .INPUTMASK='9999999'            
     ENDWITH                       
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                       
  ENDPROC
*************************設定PAGEFRAME2***************************        遲到設定-C      
  PROCEDURE PAGEFRAME2.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF11C','LABEL') 
     WITH THIS.LBLF11C
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='每月遲到共計                 分鐘(含)以上計作遲到，是否扣除全勤：'          
     ENDWITH
 THIS.ADDOBJECT('SPINNER11C','SPINNER')      &&遲到分鐘數
    WITH THIS.SPINNER11C
          .VISIBLE=.T.
	  .TOP=INT_015*24
	  .LEFT=INT_015*95
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45  
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=100
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=100
	  .INPUTMASK='999'     
     ENDWITH    
THIS.ADDOBJECT('YESNO_CHOICE12C','YESNO_CHOICE')  &&遲到是否扣全勤
    WITH THIS.YESNO_CHOICE12C
   	 .TOP=INT_015*24
         .LEFT=INT_015*360       
     ENDWITH               
THIS.ADDOBJECT('OPTIONGROUP13C','OPTIONGROUP13C')  &&遲到是否設定額外扣款
    WITH THIS.OPTIONGROUP13C
   	 .TOP=INT_015*49
   	 .LEFT=INT_015*20
     ENDWITH               
THIS.ADDOBJECT('LBLF14C','LABEL') 
     WITH THIS.LBLF14C 
          .VISIBLE=.T.
          .LEFT=INT_015*128
          .TOP=INT_015*58
          .AUTOSIZE=.T.
          .CAPTION='每月遲到超過                  分鐘(含)以上須再另扣' 
     ENDWITH     
 THIS.ADDOBJECT('SPINNER14C','SPINNER')      &&遲到扣款分鐘數
    WITH THIS.SPINNER14C
          .VISIBLE=.T.
	  .TOP=INT_015*51
	  .LEFT=INT_015*204
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45  
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=100
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=100
	  .INPUTMASK='999'     
     ENDWITH        
THIS.ADDOBJECT('OPTIONGROUP15C','OPTIONGROUP15C')  &&遲到超過分鐘數之另扣款項目
    WITH THIS.OPTIONGROUP15C
   	 .TOP=INT_015*78
   	 .LEFT=INT_015*130
     ENDWITH            
  THIS.ADDOBJECT('TXTF15C','TEXTBOX')      &&遲到扣款底數
     WITH THIS.TXTF15C
          .VISIBLE=.T.
          .LEFT=INT_015*204
          .TOP=INT_015*80   
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25      
          .MAXLENGTH=4           
     ENDWITH           
THIS.ADDOBJECT('COMBOBOX15C','COMBOBOX')
    WITH THIS.COMBOBOX15C
           .VISIBLE=.F.
	   .TOP=INT_015*80
	   .LEFT=INT_015*204
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*55
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX15C'	   
   ENDWITH  
THIS.ADDOBJECT('LBLF16C','LABEL') 
     WITH THIS.LBLF16C 
          .VISIBLE=.T.
          .LEFT=INT_015*268
          .TOP=INT_015*87
          .AUTOSIZE=.T.
          .CAPTION='扣款比例                   ％'  
     ENDWITH        
  THIS.ADDOBJECT('TXTF16C','TEXTBOX')      &&遲到扣款比例
     WITH THIS.TXTF16C
          .VISIBLE=.T.
          .LEFT=INT_015*320
          .TOP=INT_015*80
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25      
          .MAXLENGTH=6   
          .INPUTMASK='999.99'         
     ENDWITH       
  THIS.ADDOBJECT('TXTF17C','TEXTBOX')      &&遲到扣款金額
     WITH THIS.TXTF17C
          .VISIBLE=.T.
          .LEFT=INT_015*457
          .TOP=INT_015*80
          .WIDTH=INT_015*53
          .HEIGHT=INT_015*25      
          .MAXLENGTH=7  
          .INPUTMASK='9999999'         
     ENDWITH                   
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH          
  ENDPROC 
*************************設定PAGEFRAME3***************************        早退設定-C      
  PROCEDURE PAGEFRAME3.PAGE1.INIT     
  THIS.ADDOBJECT('LBLF11D','LABEL') 
     WITH THIS.LBLF11D
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='每月早退共計                 分鐘(含)以上計作早退，是否扣除全勤：'           
     ENDWITH
 THIS.ADDOBJECT('SPINNER11D','SPINNER')      &&早退分鐘數
    WITH THIS.SPINNER11D
          .VISIBLE=.T.
	  .TOP=INT_015*24
	  .LEFT=INT_015*95
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45  
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=100
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=100
	  .INPUTMASK='999'     
     ENDWITH   
THIS.ADDOBJECT('YESNO_CHOICE12D','YESNO_CHOICE')  &&早退是否扣全勤
    WITH THIS.YESNO_CHOICE12D
   	 .TOP=INT_015*24
         .LEFT=INT_015*360
     ENDWITH           
THIS.ADDOBJECT('OPTIONGROUP13D','OPTIONGROUP13D')  &&早退是否設定額外扣款
    WITH THIS.OPTIONGROUP13D
   	 .TOP=INT_015*49
   	 .LEFT=INT_015*20
     ENDWITH              
THIS.ADDOBJECT('LBLF14D','LABEL') 
     WITH THIS.LBLF14D 
          .VISIBLE=.T.
          .LEFT=INT_015*128
          .TOP=INT_015*58
          .AUTOSIZE=.T.
          .CAPTION='每月早退超過                  分鐘(含)以上須再另扣' 
     ENDWITH     
 THIS.ADDOBJECT('SPINNER14D','SPINNER')      &&早退扣款分鐘數
    WITH THIS.SPINNER14D
          .VISIBLE=.T.
	  .TOP=INT_015*51
	  .LEFT=INT_015*204
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45  
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=100
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=100
	  .INPUTMASK='999'     
     ENDWITH        
THIS.ADDOBJECT('OPTIONGROUP15D','OPTIONGROUP15D')  &&遲到超過分鐘數之另扣款項目
    WITH THIS.OPTIONGROUP15D
   	 .TOP=INT_015*78
   	 .LEFT=INT_015*130
     ENDWITH         
  THIS.ADDOBJECT('TXTF15D','TEXTBOX')      &&扣款底數
     WITH THIS.TXTF15D
          .VISIBLE=.T.
          .LEFT=INT_015*204
          .TOP=INT_015*80
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25      
          .MAXLENGTH=4          
     ENDWITH          
THIS.ADDOBJECT('COMBOBOX15D','COMBOBOX')
    WITH THIS.COMBOBOX15D
           .VISIBLE=.F.
	   .TOP=INT_015*80
	   .LEFT=INT_015*204
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*55
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX15D'	   
   ENDWITH     
THIS.ADDOBJECT('LBLF16D','LABEL') 
     WITH THIS.LBLF16D 
          .VISIBLE=.T.
          .LEFT=INT_015*268
          .TOP=INT_015*87
          .AUTOSIZE=.T.
          .CAPTION='扣款比例                   ％'  
     ENDWITH     
  THIS.ADDOBJECT('TXTF16D','TEXTBOX')      &&扣款比例
     WITH THIS.TXTF16D
          .VISIBLE=.T.
          .LEFT=INT_015*320
          .TOP=INT_015*80
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25      
          .MAXLENGTH=6   
          .INPUTMASK='999.99'         
     ENDWITH          
  THIS.ADDOBJECT('TXTF17D','TEXTBOX')      &&遲到扣款金額
     WITH THIS.TXTF17D
          .VISIBLE=.T.
          .LEFT=INT_015*457
          .TOP=INT_015*80
          .WIDTH=INT_015*53
          .HEIGHT=INT_015*25      
          .MAXLENGTH=7  
          .INPUTMASK='9999999'         
     ENDWITH                
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                          
  ENDPROC   
  *************************設定PAGEFRAME4***************************  年假範圍設定-E
PROCEDURE PAGEFRAME4.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF18','LABEL') 
     WITH THIS.LBLF18
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='年資範圍設定：任職滿             年，未滿            年者，每年給予年假            天'          
     ENDWITH
  THIS.ADDOBJECT('TXTF18','TEXTBOX')      &&起始年資
     WITH THIS.TXTF18
          .VISIBLE=.T.
          .LEFT=INT_015*143 
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH        
  THIS.ADDOBJECT('TXTF19','TEXTBOX')      &&終止年資
     WITH THIS.TXTF19
          .VISIBLE=.T.
          .LEFT=INT_015*230
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH                     
  THIS.ADDOBJECT('TXTF20','TEXTBOX')     &&可休天數   
     WITH THIS.TXTF20
          .VISIBLE=.T.
          .LEFT=INT_015*373
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH  
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*190
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*190
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH   
  ENDPROC   
  *************************設定PAGEFRAME5***************************  年假限制設定-F
PROCEDURE PAGEFRAME5.PAGE1.INIT  
  THIS.ADDOBJECT('LBLF21','LABEL') 
     WITH THIS.LBLF21
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .WIDTH=INT_015*95
          .CAPTION='最大年資設定'          
     ENDWITH         
  THIS.ADDOBJECT('OPTIONGROUP21','OPTIONGROUP21')
     WITH THIS.OPTIONGROUP21
	   .TOP=INT_015*20
	   .LEFT=INT_015*95
     ENDWITH  
  THIS.ADDOBJECT('TXTF21','TEXTBOX')      &&現制最大年資
     WITH THIS.TXTF21
          .VISIBLE=.T.
          .LEFT=INT_015*228
          .TOP=INT_015*22
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH
  THIS.ADDOBJECT('LBLF22','LABEL') 
     WITH THIS.LBLF22
          .VISIBLE=.T.
          .LEFT=INT_015*263
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='年且每年增加年假             天'          
     ENDWITH          
  THIS.ADDOBJECT('TXTF22','TEXTBOX')     &&每年增加年假之天數
     WITH THIS.TXTF22
          .VISIBLE=.T.
          .LEFT=INT_015*361
          .TOP=INT_015*22
          .WIDTH=INT_015*33
          .HEIGHT=INT_015*25
	  .INPUTMASK='999'          
          .MAXLENGTH=3          
     ENDWITH     
  THIS.ADDOBJECT('LBLF23','LABEL') 
     WITH THIS.LBLF23
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*58
          .WIDTH=INT_015*95
          .CAPTION='年假上限設定'          
     ENDWITH                       
  THIS.ADDOBJECT('OPTIONGROUP23','OPTIONGROUP23')
     WITH THIS.OPTIONGROUP23
	   .TOP=INT_015*49
	   .LEFT=INT_015*95
     ENDWITH 
  THIS.ADDOBJECT('TXTF23','TEXTBOX')      &&年假上限
     WITH THIS.TXTF23
          .VISIBLE=.T.
          .LEFT=INT_015*288
          .TOP=INT_015*51
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=3          
     ENDWITH 
  THIS.ADDOBJECT('LBLF231','LABEL') 
     WITH THIS.LBLF231
          .VISIBLE=.T.
          .LEFT=INT_015*322
          .TOP=INT_015*58
          .AUTOSIZE=.T.
          .CAPTION='天(含)'          
     ENDWITH      
  THIS.ADDOBJECT('LBLF24','LABEL') 
     WITH THIS.LBLF24
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*87
          .WIDTH=INT_015*95 
          .CAPTION='到職未滿一年'        
     ENDWITH    
  THIS.ADDOBJECT('OPTIONGROUP24','OPTIONGROUP24')
     WITH THIS.OPTIONGROUP24
	   .TOP=INT_015*78
	   .LEFT=INT_015*95
     ENDWITH 
  THIS.ADDOBJECT('TXTF24','TEXTBOX')      
     WITH THIS.TXTF24
          .VISIBLE=.T.
          .LEFT=INT_015*168
          .TOP=INT_015*80
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH  
  THIS.ADDOBJECT('LBLF241','LABEL') 
     WITH THIS.LBLF241
          .VISIBLE=.T.
          .LEFT=INT_015*202
          .TOP=INT_015*87
          .WIDTH=INT_015*15
          .CAPTION='天'        
     ENDWITH       
  THIS.ADDOBJECT('TXTF25','TEXTBOX')      
     WITH THIS.TXTF25
          .VISIBLE=.T.
          .LEFT=INT_015*300
          .TOP=INT_015*80
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
	  .INPUTMASK='99'          
          .MAXLENGTH=2          
     ENDWITH 
  THIS.ADDOBJECT('LBLF25','LABEL') 
     WITH THIS.LBLF25
          .VISIBLE=.T.
          .LEFT=INT_015*335
          .TOP=INT_015*87
          .AUTOSIZE=.T. 
          .CAPTION='天乘上到職月數'        
     ENDWITH                                                            
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                      
  ENDPROC  
  *************************設定PAGEFRAME6***************************  請假扣款項目-G  
PROCEDURE PAGEFRAME6.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF26','LABEL') 
     WITH THIS.LBLF26
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .WIDTH=INT_015*50
          .CAPTION='假別編號'          
     ENDWITH
  THIS.ADDOBJECT('TXTF26','TEXTBOX')      
     WITH THIS.TXTF26
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2         
     ENDWITH
  THIS.ADDOBJECT('LBLF58','LABEL') 
     WITH THIS.LBLF58
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*31
          .WIDTH=INT_015*50
          .CAPTION='是否給薪'          
     ENDWITH    
 THIS.ADDOBJECT('OPTIONGROUP58','OPTIONGROUP58')
     WITH THIS.OPTIONGROUP58
     	   .VISIBLE=.T.
   	   .TOP=INT_015*24
   	   .LEFT=INT_015*261          
     ENDWITH    
  THIS.ADDOBJECT('LBLF68','LABEL') 
     WITH THIS.LBLF68
          .VISIBLE=.T.
          .LEFT=INT_015*355
          .TOP=INT_015*31
          .WIDTH=INT_015*75
          .CAPTION='是否計算假日'
     ENDWITH    
 THIS.ADDOBJECT('OPTIONGROUP68','OPTIONGROUP')
     WITH THIS.OPTIONGROUP68
     	   .VISIBLE=.T.
   	       .TOP=INT_015*24
   	       .LEFT=INT_015*435 
		   .HEIGHT=INT_015*25
		   .WIDTH=INT_015*85   	       
		   .BUTTONCOUNT=2
		   .VALUE=1
		   .ENABLED=.F.
		   .NAME='OPTIONGROUP68'
		   .OPTION1.CAPTION='是'
		   .OPTION1.FONTSIZE=INT_015*9      
		   .OPTION1.TOP=INT_015*5
		   .OPTION1.AUTOSIZE=.T.
		   .OPTION1.LEFT=INT_015*5 
		   .OPTION1.NAME='OPTION1'
		   .OPTION2.CAPTION='否'
		   .OPTION2.FONTSIZE=INT_015*9    
		   .OPTION2.TOP=INT_015*5
		   .OPTION2.AUTOSIZE=.T.
		   .OPTION2.LEFT=INT_015*50
		   .OPTION2.NAME='OPTION2'       
     ENDWITH           
  THIS.ADDOBJECT('LBLF27','LABEL') 
     WITH THIS.LBLF27
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*55
          .WIDTH=INT_015*50
          .CAPTION='假別項目'          
     ENDWITH
  THIS.ADDOBJECT('TXTF27','TEXTBOX')      
     WITH THIS.TXTF27
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*49
          .WIDTH=INT_015*130
          .HEIGHT=INT_015*25
          .MAXLENGTH=30          
     ENDWITH   
  THIS.ADDOBJECT('LBLF28','LABEL') 
     WITH THIS.LBLF28
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*55
          .WIDTH=INT_015*50
          .CAPTION='計數單位'          
     ENDWITH    
 THIS.ADDOBJECT('OPTIONGROUP28','OPTIONGROUP28')
     WITH THIS.OPTIONGROUP28
     	   .VISIBLE=.T.
   	   .TOP=INT_015*49
   	   .LEFT=INT_015*261          
     ENDWITH          
  THIS.ADDOBJECT('LBLF62','LABEL') 
     WITH THIS.LBLF62
          .VISIBLE=.T.
          .LEFT=INT_015*400
          .TOP=INT_015*55
          .AUTOSIZE=.T.
          .CAPTION='請假基數                 日'          
     ENDWITH
  THIS.ADDOBJECT('TXTF62','TEXTBOX')      
     WITH THIS.TXTF62
          .VISIBLE=.T.
          .LEFT=INT_015*452
          .TOP=INT_015*49
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25
          .MAXLENGTH=4   
          .INPUTMASK='99.9'        
     ENDWITH       
  THIS.ADDOBJECT('LBLF29','LABEL') 
     WITH THIS.LBLF29
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*83
          .AUTOSIZE=.T.
          .CAPTION='年度限制'          
     ENDWITH           
 THIS.ADDOBJECT('OPTIONGROUP29','OPTIONGROUP29')
     WITH THIS.OPTIONGROUP29
   	   .TOP=INT_015*74
   	   .LEFT=INT_015*71    
     ENDWITH  
  THIS.ADDOBJECT('TXTF29','TEXTBOX')      &&年度限制
     WITH THIS.TXTF29
          .VISIBLE=.T.
          .LEFT=INT_015*318
          .TOP=INT_015*76
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25
          .MAXLENGTH=4        
	  .INPUTMASK='999.9'            
     ENDWITH       
  THIS.ADDOBJECT('LBLF291','LABEL') 
     WITH THIS.LBLF291
          .VISIBLE=.T.
          .LEFT=INT_015*365
          .TOP=INT_015*83
          .AUTOSIZE=.T.
          .CAPTION='天(不含)'          
     ENDWITH        
  THIS.ADDOBJECT('LBLF30','LABEL') 
     WITH THIS.LBLF30
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*111
          .AUTOSIZE=.T.
          .CAPTION='扣款方式'          
     ENDWITH        
 THIS.ADDOBJECT('OPTIONGROUP30','OPTIONGROUP30')
     WITH THIS.OPTIONGROUP30
   	   .TOP=INT_015*103
   	   .LEFT=INT_015*71    
     ENDWITH            
  THIS.ADDOBJECT('TXTF30','TEXTBOX')     &&扣款底數 
     WITH THIS.TXTF30
          .VISIBLE=.T.
          .LEFT=INT_015*208
          .TOP=INT_015*105
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=4            
     ENDWITH         
THIS.ADDOBJECT('COMBOBOX30','COMBOBOX')
    WITH THIS.COMBOBOX30
       .VISIBLE=.F.
	   .TOP=INT_015*105
	   .LEFT=INT_015*208
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*55
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX30'	   
   ENDWITH  	   
  THIS.ADDOBJECT('LBLF31','LABEL') 
     WITH THIS.LBLF31
          .VISIBLE=.T.
          .LEFT=INT_015*268
          .TOP=INT_015*111
          .AUTOSIZE=.T.
          .CAPTION='扣款比例                   ％'          
     ENDWITH   
  THIS.ADDOBJECT('TXTF31','TEXTBOX')      &&扣款比例
     WITH THIS.TXTF31
          .VISIBLE=.T.
          .LEFT=INT_015*318
          .TOP=INT_015*105
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=6        
	  .INPUTMASK='999.99'            
     ENDWITH    
  THIS.ADDOBJECT('TXTF32','TEXTBOX')      &&扣款金額
     WITH THIS.TXTF32
          .VISIBLE=.T.
          .LEFT=INT_015*463
          .TOP=INT_015*105
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7        
	  .INPUTMASK='9999999'            
     ENDWITH      
     ***************
  THIS.ADDOBJECT('LBLF33','LABEL') 
     WITH THIS.LBLF33
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*138
          .WIDTH=INT_015*50
          .CAPTION='影響全勤'          
     ENDWITH     
 THIS.ADDOBJECT('YESNO_CHOICE33','YESNO_CHOICE')
     WITH THIS.YESNO_CHOICE33 
   	   .TOP=INT_015*132
   	   .LEFT=INT_015*71
     ENDWITH  
  THIS.ADDOBJECT('LBLF34','LABEL') 
     WITH THIS.LBLF34
          .VISIBLE=.T.
          .LEFT=INT_015*171
          .TOP=INT_015*139
          .AUTOSIZE=.T.
          .CAPTION='不扣全勤緩衝天數               天(含)'          
     ENDWITH      
  THIS.ADDOBJECT('TXTF34','TEXTBOX')      
     WITH THIS.TXTF34
          .VISIBLE=.T.
          .LEFT=INT_015*270
          .TOP=INT_015*133
          .WIDTH=INT_015*35
          .HEIGHT=INT_015*25
          .MAXLENGTH=3        
	  .INPUTMASK='99.9'            
     ENDWITH                
  THIS.ADDOBJECT('LBLF35','LABEL') 
     WITH THIS.LBLF35
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*164
          .AUTOSIZE=.T.
          .CAPTION='全勤扣款方式一： 缺勤天數             ，扣款比例                    ％ '
     ENDWITH      
  THIS.ADDOBJECT('TXTF35','TEXTBOX')      &&全勤扣款方式一缺勤天數
     WITH THIS.TXTF35
          .VISIBLE=.T.
          .LEFT=INT_015*170
          .TOP=INT_015*158
          .WIDTH=INT_015*35
          .HEIGHT=INT_015*25
          .MAXLENGTH=3        
	  .INPUTMASK='99.9'            
     ENDWITH 
  THIS.ADDOBJECT('TXTF36','TEXTBOX')     &&全勤扣款方式一扣款比例 
     WITH THIS.TXTF36
          .VISIBLE=.T.
          .LEFT=INT_015*270
          .TOP=INT_015*158
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=6         
	  .INPUTMASK='999.99'            
     ENDWITH 
  THIS.ADDOBJECT('LBLF37','LABEL') 
     WITH THIS.LBLF37
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*189
          .AUTOSIZE=.T.
          .CAPTION='全勤扣款方式二： 缺勤天數             ，扣款比例                    ％ '
     ENDWITH      
  THIS.ADDOBJECT('TXTF37','TEXTBOX')       &&全勤扣款方式二缺勤天數
     WITH THIS.TXTF37
          .VISIBLE=.T.
          .LEFT=INT_015*170
          .TOP=INT_015*183
          .WIDTH=INT_015*35
          .HEIGHT=INT_015*25
          .MAXLENGTH=3        
	  .INPUTMASK='99.9'            
     ENDWITH 
  THIS.ADDOBJECT('TXTF38','TEXTBOX')      &&全勤扣款方式二扣款比例
     WITH THIS.TXTF38
          .VISIBLE=.T.
          .LEFT=INT_015*270
          .TOP=INT_015*183
           .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=6         
	  .INPUTMASK='999.99'            
     ENDWITH                      
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                      
  ENDPROC
  *************************設定PAGEFRAME7***************************  懲處扣款設定-G
PROCEDURE PAGEFRAME7.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF39','LABEL') 
     WITH THIS.LBLF39
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .WIDTH=INT_015*50
          .CAPTION='懲處編號'          
     ENDWITH
  THIS.ADDOBJECT('TXTF39','TEXTBOX')      
     WITH THIS.TXTF39
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*24
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2         
     ENDWITH
  THIS.ADDOBJECT('LBLF40','LABEL') 
     WITH THIS.LBLF40
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*55
          .WIDTH=INT_015*50
          .CAPTION='懲處項目'          
     ENDWITH
  THIS.ADDOBJECT('TXTF40','TEXTBOX')      
     WITH THIS.TXTF40
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*49
          .WIDTH=INT_015*140
          .HEIGHT=INT_015*25
          .MAXLENGTH=30          
     ENDWITH   
  THIS.ADDOBJECT('LBLF41','LABEL') 
     WITH THIS.LBLF41
          .VISIBLE=.T.
          .LEFT=INT_015*250
          .TOP=INT_015*55
          .WIDTH=INT_015*80
          .CAPTION=' 連續扣款月數'          
     ENDWITH
  THIS.ADDOBJECT('TXTF41','TEXTBOX')      
     WITH THIS.TXTF41
          .VISIBLE=.T.
          .LEFT=INT_015*330
          .TOP=INT_015*49
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25
          .MAXLENGTH=2
          .INPUTMASK='99'         
     ENDWITH        
  THIS.ADDOBJECT('LBLF42','LABEL') 
     WITH THIS.LBLF42
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*83
          .AUTOSIZE=.T.
          .CAPTION='扣款方式'          
     ENDWITH  
 THIS.ADDOBJECT('OPTIONGROUP42','OPTIONGROUP42')
     WITH THIS.OPTIONGROUP42
   	   .TOP=INT_015*74
   	   .LEFT=INT_015*71    
     ENDWITH  
  THIS.ADDOBJECT('TXTF42','TEXTBOX')      
     WITH THIS.TXTF42
          .VISIBLE=.T.
          .LEFT=INT_015*145
          .TOP=INT_015*76
          .WIDTH=INT_015*75
          .HEIGHT=INT_015*25
          .MAXLENGTH=8              
     ENDWITH                        
THIS.ADDOBJECT('COMBOBOX42','COMBOBOX')
    WITH THIS.COMBOBOX42
           .VISIBLE=.F.
	   .TOP=INT_015*76
	   .LEFT=INT_015*145
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*75
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪,薪資總額'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX42'	   
   ENDWITH   
  THIS.ADDOBJECT('LBLF43','LABEL') 
     WITH THIS.LBLF43
          .VISIBLE=.T.
          .LEFT=INT_015*223
          .TOP=INT_015*83
          .AUTOSIZE=.T.
          .CAPTION=' 扣款比例                   ％'          
     ENDWITH     
  THIS.ADDOBJECT('TXTF43','TEXTBOX')      
     WITH THIS.TXTF43
          .VISIBLE=.T.
          .LEFT=INT_015*277
          .TOP=INT_015*76
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=6        
	  .INPUTMASK='999.99'            
     ENDWITH    
  THIS.ADDOBJECT('TXTF44','TEXTBOX')      
     WITH THIS.TXTF44
          .VISIBLE=.T.
          .LEFT=INT_015*418
          .TOP=INT_015*76
          .WIDTH=INT_015*60
          .HEIGHT=INT_015*25
          .MAXLENGTH=7        
	  .INPUTMASK='9999999'            
     ENDWITH  
    THIS.ADDOBJECT('LBLF45','LABEL') 
     WITH THIS.LBLF45
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*108
          .WIDTH=INT_015*50
          .CAPTION='備註說明'          
     ENDWITH
  THIS.ADDOBJECT('TXTF45','TEXTBOX')      
     WITH THIS.TXTF45
          .VISIBLE=.T.
          .LEFT=INT_015*71
          .TOP=INT_015*103
          .WIDTH=INT_015*320
          .HEIGHT=INT_015*25
          .MAXLENGTH=50    
     ENDWITH                                     
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                      
  ENDPROC  
  *************************設定PAGEFRAME8***************************  班別設定-I
PROCEDURE PAGEFRAME8.PAGE1.INIT 
  THIS.ADDOBJECT('LBLF48','LABEL') 
     WITH THIS.LBLF48
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*33
          .WIDTH=INT_015*80
          .CAPTION='班 別 編 號'          
     ENDWITH        
  THIS.ADDOBJECT('TXTF48','TEXTBOX')     
     WITH THIS.TXTF48
          .VISIBLE=.T.
          .LEFT=INT_015*85
          .TOP=INT_015*26
          .WIDTH=INT_015*30
          .HEIGHT=INT_015*25   
          .MAXLENGTH=2          
     ENDWITH        
  THIS.ADDOBJECT('LBLF49','LABEL') 
     WITH THIS.LBLF49
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*58
          .WIDTH=INT_015*80
          .CAPTION='班 別 說 明'          
     ENDWITH        
  THIS.ADDOBJECT('TXTF49','TEXTBOX')     
     WITH THIS.TXTF49
          .VISIBLE=.T.
          .LEFT=INT_015*85
          .TOP=INT_015*51
          .WIDTH=INT_015*170
          .HEIGHT=INT_015*25   
          .MAXLENGTH=30          
     ENDWITH          
  THIS.ADDOBJECT('LBLF50','LABEL') 
     WITH THIS.LBLF50
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*82
          .AUTOSIZE=.T.
          .CAPTION='上 班 時 段                   ，下 班 時 段'          
     ENDWITH        
  THIS.ADDOBJECT('TXTF50','TXTF50')     
     WITH THIS.TXTF50
          .VISIBLE=.T.
          .FONTSIZE=INT_015*9
          .LEFT=INT_015*85
          .TOP=INT_015*76
     ENDWITH   
  THIS.ADDOBJECT('TXTF51','TXTF51')     
     WITH THIS.TXTF51
          .VISIBLE=.T.
          .FONTSIZE=INT_015*9
          .LEFT=INT_015*210
          .TOP=INT_015*76
     ENDWITH                      
  THIS.ADDOBJECT('LBLF52','LABEL') 
     WITH THIS.LBLF52
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*107
          .AUTOSIZE=.T.
          .CAPTION='休息時段起                  ，休息時段止'          
     ENDWITH      
  THIS.ADDOBJECT('TXTF52','TXTF52')     
     WITH THIS.TXTF52
          .VISIBLE=.T.
          .LEFT=INT_015*85
          .TOP=INT_015*101
	  .FONTSIZE=INT_015*9
     ENDWITH     
  THIS.ADDOBJECT('TXTF53','TXTF53')     
     WITH THIS.TXTF53
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*101
	  .FONTSIZE=INT_015*9  
     ENDWITH     
  THIS.ADDOBJECT('LBLF57','LABEL') 
     WITH THIS.LBLF57
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*132
          .AUTOSIZE=.T.
          .CAPTION='上班總時數'       
     ENDWITH    
  THIS.ADDOBJECT('TXTF57','TEXTBOX')     
     WITH THIS.TXTF57
          .VISIBLE=.T.
          .LEFT=INT_015*85
          .TOP=INT_015*126
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=3 
          .INPUTMASK='99.9'        
     ENDWITH     
  THIS.ADDOBJECT('LBLF66','LABEL') 
     WITH THIS.LBLF66
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*157
          .AUTOSIZE=.T.
          .CAPTION='上班刷卡時段止'       
     ENDWITH                 
  THIS.ADDOBJECT('TXTF66','TEXTBOX')     
     WITH THIS.TXTF66
          .VISIBLE=.T.
          .LEFT=INT_015*105
          .TOP=INT_015*151
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=5     
          .INPUTMASK='99:99'     
     ENDWITH        
  THIS.ADDOBJECT('LBLF67','LABEL') 
     WITH THIS.LBLF67
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*182
          .AUTOSIZE=.T.
          .CAPTION='下班刷卡時段起'       
     ENDWITH       
  THIS.ADDOBJECT('TXTF67','TEXTBOX')     
     WITH THIS.TXTF67
          .VISIBLE=.T.
          .LEFT=INT_015*105
          .TOP=INT_015*176
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=5   
          .INPUTMASK='99:99'       
     ENDWITH        
     ***
  THIS.ADDOBJECT('LBLF54','LABEL') 
     WITH THIS.LBLF54
          .VISIBLE=.T.
          .LEFT=INT_015*300
          .TOP=INT_015*58
          .WIDTH=INT_015*95
          .CAPTION='上班遲到緩衝時間'          
     ENDWITH          
  THIS.ADDOBJECT('TXTF54','TEXTBOX')     
     WITH THIS.TXTF54
          .VISIBLE=.T.
          .LEFT=INT_015*400
          .TOP=INT_015*52
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=5 
          .INPUTMASK='99:99'        
     ENDWITH           
  THIS.ADDOBJECT('LBLF55','LABEL') 
     WITH THIS.LBLF55
          .VISIBLE=.T.
          .LEFT=INT_015*300
          .TOP=INT_015*83
          .WIDTH=INT_015*95
          .CAPTION='下班早退緩衝時間'          
     ENDWITH          
  THIS.ADDOBJECT('TXTF55','TEXTBOX')     
     WITH THIS.TXTF55
          .VISIBLE=.T.
          .LEFT=INT_015*400
          .TOP=INT_015*77
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=5 
          .INPUTMASK='99:99'        
     ENDWITH          
  THIS.ADDOBJECT('LBLF56','LABEL') 
     WITH THIS.LBLF56
          .VISIBLE=.T.
          .LEFT=INT_015*300
          .TOP=INT_015*108
          .WIDTH=INT_015*95
          .CAPTION='加班上班緩衝時間'          
     ENDWITH          
  THIS.ADDOBJECT('TXTF56','TEXTBOX')     
     WITH THIS.TXTF56
          .VISIBLE=.T.
          .LEFT=INT_015*400
          .TOP=INT_015*102
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25   
          .MAXLENGTH=5 
          .INPUTMASK='99:99'        
     ENDWITH                                                  
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                      
  ENDPROC    
*************************設定PAGEFRAME9***************************        其他所得發放設定 -J       
  PROCEDURE PAGEFRAME9.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF59','LABEL') 
     WITH THIS.LBLF59
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='其他所得發放之固定扣稅率                    ％'          
     ENDWITH      
  THIS.ADDOBJECT('TXTF59','TEXTBOX')      
     WITH THIS.TXTF59
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*165
          .TOP=INT_015*24
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
           .INPUTMASK='99.9'  
          .MAXLENGTH=4              
     ENDWITH                                 
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                          
  ENDPROC  
*************************設定PAGEFRAME10***************************        年終獎金發放設定 -K       
  PROCEDURE PAGEFRAME10.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF68','LABEL') 
     WITH THIS.LBLF68
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='年終獎金發放方式'         
     ENDWITH                                        
 THIS.ADDOBJECT('OPTIONGROUP68','OPTIONGROUP681')
    WITH THIS.OPTIONGROUP68
           .VISIBLE=.T.
   	   .TOP=INT_015*49
   	   .LEFT=INT_015*30       
     ENDWITH       
  **** 方式一                              
  THIS.ADDOBJECT('TXTF691','TEXTBOX')	&&基數      
     WITH THIS.TXTF691
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*122
          .TOP=INT_015*52
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25
           .INPUTMASK='999.99'  
          .MAXLENGTH=6             
     ENDWITH               
  THIS.ADDOBJECT('LBLF701','LABEL') 
     WITH THIS.LBLF701
          .VISIBLE=.T.
          .LEFT=INT_015*170
          .TOP=INT_015*60
          .AUTOSIZE=.T.
          .CAPTION='Ｘ 發放底數                   Ｘ（到職月數／12）' 
     ENDWITH                       
 THIS.ADDOBJECT('TXTF701','TEXTBOX')		&&發放底數      
     WITH THIS.TXTF701
          .VISIBLE=.T.
          .LEFT=INT_015*237
          .TOP=INT_015*52
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=4                
     ENDWITH                              
THIS.ADDOBJECT('COMBOBOX701','COMBOBOX')		&&發放底數  
    WITH THIS.COMBOBOX701
           .VISIBLE=.F.
	   .TOP=INT_015*52
	   .LEFT=INT_015*237
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*50
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX701'	   
   ENDWITH                                
 **** 方式二   
  THIS.ADDOBJECT('TXTF692','TEXTBOX')	&&基數      
     WITH THIS.TXTF692
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*122
          .TOP=INT_015*78
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25
           .INPUTMASK='999.99'  
          .MAXLENGTH=6              
     ENDWITH    
  THIS.ADDOBJECT('LBLF702','LABEL') 
     WITH THIS.LBLF702
          .VISIBLE=.T.
          .LEFT=INT_015*170
          .TOP=INT_015*86
          .AUTOSIZE=.T.
          .CAPTION='Ｘ 發放底數' 
     ENDWITH     
 THIS.ADDOBJECT('TXTF702','TEXTBOX')		&&發放底數        
     WITH THIS.TXTF702
          .VISIBLE=.T.
          .LEFT=INT_015*237
          .TOP=INT_015*78
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=4                
     ENDWITH                              
THIS.ADDOBJECT('COMBOBOX702','COMBOBOX')		&&發放底數  
    WITH THIS.COMBOBOX702
           .VISIBLE=.F.
	   .TOP=INT_015*78
	   .LEFT=INT_015*237
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*50
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX702'	   
   ENDWITH    
 **** 方式三         
  THIS.ADDOBJECT('TXTF693','TEXTBOX')	&&基數      
     WITH THIS.TXTF693
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*122
          .TOP=INT_015*104
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25
           .INPUTMASK='999.99'  
          .MAXLENGTH=6              
     ENDWITH   
  THIS.ADDOBJECT('LBLF71','LABEL') 
     WITH THIS.LBLF71
          .VISIBLE=.T.
          .LEFT=INT_015*170
          .TOP=INT_015*112
          .AUTOSIZE=.T.
          .CAPTION='Ｘ 發放金額' 
     ENDWITH         
 THIS.ADDOBJECT('TXTF71','TEXTBOX')		&&發放金額        
     WITH THIS.TXTF71
          .VISIBLE=.T.
          .LEFT=INT_015*237
          .TOP=INT_015*104
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7                
     ENDWITH       
     ****       
  THIS.ADDOBJECT('LBLF72','LABEL') 
     WITH THIS.LBLF72
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*150
          .AUTOSIZE=.T.
          .CAPTION='年終獎金發放扣稅率                    ％'         
     ENDWITH               
  THIS.ADDOBJECT('TXTF72','TEXTBOX')      
     WITH THIS.TXTF72
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*128
          .TOP=INT_015*143
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
           .INPUTMASK='9999999'  
          .MAXLENGTH=7              
     ENDWITH        
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                          
  ENDPROC     
*************************設定PAGEFRAME11***************************        未休年假獎金發放設定 -L       
  PROCEDURE PAGEFRAME11.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF68','LABEL') 
     WITH THIS.LBLF68
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='未休年假獎金發放方式'         
     ENDWITH                   
 THIS.ADDOBJECT('OPTIONGROUP68','OPTIONGROUP682')
    WITH THIS.OPTIONGROUP68
           .VISIBLE=.T.
   	   .TOP=INT_015*49
   	   .LEFT=INT_015*30    
     ENDWITH                                             
  **** 方式一                                                         
 THIS.ADDOBJECT('TXTF701','TEXTBOX') 	&&發放底數     
     WITH THIS.TXTF701
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*52
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=4                
     ENDWITH                              
THIS.ADDOBJECT('COMBOBOX701','COMBOBOX')		&&發放底數
    WITH THIS.COMBOBOX701
           .VISIBLE=.F.
	   .TOP=INT_015*52
	   .LEFT=INT_015*210
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*50
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX701'	   
   ENDWITH                 
  THIS.ADDOBJECT('LBLF69','LABEL') 
     WITH THIS.LBLF69
          .VISIBLE=.T.
          .LEFT=INT_015*265
          .TOP=INT_015*60
          .AUTOSIZE=.T.
          .CAPTION='Ｘ 發放比率' 
     ENDWITH         
  THIS.ADDOBJECT('TXTF69','TEXTBOX')	&&發放比率      
     WITH THIS.TXTF69
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*330
          .TOP=INT_015*52
          .WIDTH=INT_015*45
          .HEIGHT=INT_015*25
           .INPUTMASK='999.99'  
          .MAXLENGTH=6              
     ENDWITH                     
 **** 方式二   
 THIS.ADDOBJECT('TXTF702','TEXTBOX')		&&發放底數      
     WITH THIS.TXTF702
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*78
          .WIDTH=INT_015*50
          .HEIGHT=INT_015*25
          .MAXLENGTH=4                
     ENDWITH                              
THIS.ADDOBJECT('COMBOBOX702','COMBOBOX')		&&發放底數
    WITH THIS.COMBOBOX702
           .VISIBLE=.F.
	   .TOP=INT_015*78
	   .LEFT=INT_015*210
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*50
	   .FONTSIZE=INT_015*9   
	   .BOUNDCOLUMN=1
	   .ROWSOURCE=',時薪,日薪,底薪'
	   .ROWSOURCETYPE=1
	   .STYLE=2    
	   .VALUE=1
	   .NAME='COMBOBOX702'	   
   ENDWITH    
 **** 方式三         
 THIS.ADDOBJECT('TXTF71','TEXTBOX')		&&發放金額      
     WITH THIS.TXTF71
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*104
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7                
     ENDWITH       
     ****       
  THIS.ADDOBJECT('LBLF72','LABEL') 
     WITH THIS.LBLF72
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*150
          .AUTOSIZE=.T.
          .CAPTION='未休年假獎金發放扣稅率                    ％'         
     ENDWITH               
  THIS.ADDOBJECT('TXTF72','TEXTBOX')      
     WITH THIS.TXTF72
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*150
          .TOP=INT_015*143
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
           .INPUTMASK='9999999'  
          .MAXLENGTH=7              
     ENDWITH              
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                          
  ENDPROC  
*************************設定PAGEFRAME12***************************        公司資料設定 -X     
  PROCEDURE PAGEFRAME12.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF_INT_000','LABEL') 
     WITH THIS.LBLF_INT_000
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .WIDTH=INT_015*80
          .CAPTION='公司名稱(中文)'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_000','TEXTBOX')     
     WITH THIS.TXTF_INT_000
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*24
          .WIDTH=INT_015*420
          .HEIGHT=INT_015*25   
          .MAXLENGTH=50    
     ENDWITH         
  THIS.ADDOBJECT('LBLF_INT_076','LABEL') 
     WITH THIS.LBLF_INT_076
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*55
          .WIDTH=INT_015*80
          .CAPTION='公司名稱(英文)'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_076','TEXTBOX')     
     WITH THIS.TXTF_INT_076
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*49
          .WIDTH=INT_015*420
          .HEIGHT=INT_015*25   
          .MAXLENGTH=50    
     ENDWITH        
  THIS.ADDOBJECT('LBLF_INT_071','LABEL') 
     WITH THIS.LBLF_INT_071
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*80
          .WIDTH=INT_015*80
          .CAPTION='公司地址(中文)'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_071','TEXTBOX')     
     WITH THIS.TXTF_INT_071
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*74
          .WIDTH=INT_015*420
          .HEIGHT=INT_015*25   
          .MAXLENGTH=50    
     ENDWITH                
  THIS.ADDOBJECT('LBLF_INT_077','LABEL') 
     WITH THIS.LBLF_INT_077
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*105
          .WIDTH=INT_015*80
          .CAPTION='公司地址(英文)'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_077','TEXTBOX')     
     WITH THIS.TXTF_INT_077
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*99
          .WIDTH=INT_015*420
          .HEIGHT=INT_015*25   
          .MAXLENGTH=70    
     ENDWITH      
  THIS.ADDOBJECT('LBLF_INT_080','LABEL') 
     WITH THIS.LBLF_INT_080
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*130
          .WIDTH=INT_015*80
          .CAPTION='公   司   網   址'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_080','TEXTBOX')     
     WITH THIS.TXTF_INT_080
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*124
          .WIDTH=INT_015*420
          .HEIGHT=INT_015*25   
          .MAXLENGTH=40    
     ENDWITH         
  THIS.ADDOBJECT('LBLF_INT_078','LABEL') 
     WITH THIS.LBLF_INT_078
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*155
          .WIDTH=INT_015*80
          .CAPTION='公   司   電   話'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_078','TEXTBOX')     
     WITH THIS.TXTF_INT_078
          .VISIBLE=.T.
          .LEFT=INT_015*101
          .TOP=INT_015*149
          .WIDTH=INT_015*100
          .HEIGHT=INT_015*25   
          .MAXLENGTH=20
     ENDWITH        
  THIS.ADDOBJECT('LBLF_INT_079','LABEL') 
     WITH THIS.LBLF_INT_079
          .VISIBLE=.T.
          .LEFT=INT_015*210
          .TOP=INT_015*155
          .WIDTH=INT_015*50
          .CAPTION='公司傳真'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_079','TEXTBOX')     
     WITH THIS.TXTF_INT_079
          .VISIBLE=.T.
          .LEFT=INT_015*263
          .TOP=INT_015*149
          .WIDTH=INT_015*100
          .HEIGHT=INT_015*25   
          .MAXLENGTH=20
     ENDWITH                 
  THIS.ADDOBJECT('LBLF_INT_107','LABEL') 
     WITH THIS.LBLF_INT_107
          .VISIBLE=.T.
          .LEFT=INT_015*370
          .TOP=INT_015*155 
          .WIDTH=INT_015*50
          .CAPTION='公司統編'          
     ENDWITH
  THIS.ADDOBJECT('TXTF_INT_107','TEXTBOX')     
     WITH THIS.TXTF_INT_107
          .VISIBLE=.T.
          .LEFT=INT_015*421
          .TOP=INT_015*149
          .WIDTH=INT_015*100
          .HEIGHT=INT_015*25   
          .MAXLENGTH=8 
     ENDWITH    
  ENDPROC   
  *************************設定PAGEFRAME13***************************        加班參數設定 -Y       
  PROCEDURE PAGEFRAME13.PAGE1.INIT   
 THIS.ADDOBJECT('OPTIONGROUP63','OPTIONGROUP63')
     WITH THIS.OPTIONGROUP63
   	   .TOP=INT_015*30
   	   .LEFT=INT_015*20    
     ENDWITH    
  THIS.ADDOBJECT('LBLF63','LABEL') 
     WITH THIS.LBLF63
          .VISIBLE=.T.
          .LEFT=INT_015*143 
          .TOP=INT_015*39
          .AUTOSIZE=.T.
          .CAPTION='加班時間不滿               （含）分鐘，則以30分鐘計算'          
     ENDWITH
  THIS.ADDOBJECT('TXTF63','TEXTBOX')     
     WITH THIS.TXTF63
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*220
          .TOP=INT_015*32
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25   
          .MAXLENGTH=2    
          .INPUTMASK='99'   
     ENDWITH         
 THIS.ADDOBJECT('OPTIONGROUP64','OPTIONGROUP64')
     WITH THIS.OPTIONGROUP64
   	   .TOP=INT_015*59
   	   .LEFT=INT_015*20    
     ENDWITH         
  THIS.ADDOBJECT('LBLF64','LABEL') 
     WITH THIS.LBLF64
          .VISIBLE=.T.
          .LEFT=INT_015*143
          .TOP=INT_015*67
          .AUTOSIZE=.T.
          .CAPTION='加班時間若滿               （含）分鐘，則以60分鐘計算'      
     ENDWITH
  THIS.ADDOBJECT('TXTF64','TEXTBOX')     
     WITH THIS.TXTF64
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*220
          .TOP=INT_015*61
          .WIDTH=INT_015*40
          .HEIGHT=INT_015*25   
          .MAXLENGTH=2
          .INPUTMASK='99'    
     ENDWITH        
  THIS.ADDOBJECT('LBLF65','LABEL') 
     WITH THIS.LBLF65
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*95
          .AUTOSIZE=.T.
          .CAPTION='加班計算是否依S19.加班申請(修改)作業為主'          
     ENDWITH     
 THIS.ADDOBJECT('OPTIONGROUP65','OPTIONGROUP')
     WITH THIS.OPTIONGROUP65
     	   .VISIBLE=.T.
   	   .TOP=INT_015*88
   	   .LEFT=INT_015*260  
	   .HEIGHT=INT_015*25
	   .WIDTH=INT_015*85
	   .BUTTONCOUNT=2
	   .VALUE=1
	   .ENABLED=.F.
	   .NAME='OPTIONGROUP65'
	   .OPTION1.CAPTION='是'
	   .OPTION1.FONTSIZE=INT_015*9      
	   .OPTION1.TOP=INT_015*5
	   .OPTION1.AUTOSIZE=.T.
	   .OPTION1.LEFT=INT_015*5 
	   .OPTION1.NAME='OPTION1'
	   .OPTION2.CAPTION='否'
	   .OPTION2.FONTSIZE=INT_015*9    
	   .OPTION2.TOP=INT_015*5
	   .OPTION2.AUTOSIZE=.T.
	   .OPTION2.LEFT=INT_015*50
	   .OPTION2.NAME='OPTION2'   	       
     ENDWITH     
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH      
  ENDPROC 
*************************設定PAGEFRAME14***************************        基本參數設定 -Z       
  PROCEDURE PAGEFRAME14.PAGE1.INIT   
  THIS.ADDOBJECT('LBLF02','LABEL') 
     WITH THIS.LBLF02
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='每月共計                   天'          
     ENDWITH  
 THIS.ADDOBJECT('SPINNER02','SPINNER')      
    WITH THIS.SPINNER02
          .VISIBLE=.T.
	  .TOP=INT_015*24
	  .LEFT=INT_015*72
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45  
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=31
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=31
	  .INPUTMASK='99'     
     ENDWITH      
  THIS.ADDOBJECT('LBLF03','LABEL') 
     WITH THIS.LBLF03
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*55
          .AUTOSIZE=.T.
          .CAPTION='每天共計                   小時'          
     ENDWITH
 THIS.ADDOBJECT('SPINNER03','SPINNER')      
    WITH THIS.SPINNER03
          .VISIBLE=.T.
	  .TOP=INT_015*49
	  .LEFT=INT_015*72
          .HEIGHT=INT_015*25     
	  .WIDTH=INT_015*45    
	  .KEYBOARDLOWVALUE=0
	  .KEYBOARDHIGHVALUE=24
	  .SPINNERLOWVALUE=0
	  .SPINNERHIGHVALUE=24
	  .INPUTMASK='99'     
     ENDWITH    
  THIS.ADDOBJECT('LBLF04','LABEL') 
     WITH THIS.LBLF04
          .VISIBLE=.T.
          .LEFT=INT_015*200
          .TOP=INT_015*30
          .AUTOSIZE=.T.
          .CAPTION='全勤獎金'          
     ENDWITH
  THIS.ADDOBJECT('TXTF04','TEXTBOX')      
     WITH THIS.TXTF04
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*250
          .TOP=INT_015*24
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7              
     ENDWITH                    
  THIS.ADDOBJECT('LBLF60','LABEL') 
     WITH THIS.LBLF60
          .VISIBLE=.T.
          .LEFT=INT_015*200
          .TOP=INT_015*55
          .AUTOSIZE=.T.
          .CAPTION='伙  食  費'          
     ENDWITH
  THIS.ADDOBJECT('TXTF60','TEXTBOX')      
     WITH THIS.TXTF60
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*250
          .TOP=INT_015*49
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7              
     ENDWITH          
  THIS.ADDOBJECT('LBLF61','LABEL') 
     WITH THIS.LBLF61
          .VISIBLE=.T.
          .LEFT=INT_015*200
          .TOP=INT_015*80
          .AUTOSIZE=.T.
          .CAPTION='交  通  費'          
     ENDWITH
  THIS.ADDOBJECT('TXTF61','TEXTBOX')      
     WITH THIS.TXTF61
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*250
          .TOP=INT_015*74
          .WIDTH=INT_015*55
          .HEIGHT=INT_015*25
          .MAXLENGTH=7              
     ENDWITH      
  THIS.ADDOBJECT('LBLF73','LABEL') 
     WITH THIS.LBLF73
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*120
          .AUTOSIZE=.T.
          .CAPTION='S10.薪資計算作業之報表明細欄位數'          
     ENDWITH         
  THIS.ADDOBJECT('TXTF73','TEXTBOX')      
     WITH THIS.TXTF73
          .VISIBLE=.T.
          .ALIGNMENT=1
          .LEFT=INT_015*220
          .TOP=INT_015*115
          .WIDTH=INT_015*35
          .HEIGHT=INT_015*25
          .MAXLENGTH=2              
     ENDWITH               
 THIS.ADDOBJECT('LBLF99','LABEL') 
     WITH THIS.LBLF99
          .VISIBLE=.T.
          .LEFT=INT_015*20
          .TOP=INT_015*220
          .WIDTH=INT_015*50
          .CAPTION='安全資料'          
     ENDWITH  
  THIS.ADDOBJECT('LBLF991','LABEL') 
     WITH THIS.LBLF991
          .VISIBLE=.T.
          .LEFT=INT_015*73
          .TOP=INT_015*220
          .AUTOSIZE=.T.
          .CAPTION=''          
     ENDWITH                          
  ENDPROC 
**********  
  PROCEDURE INIT 
       THISFORM.PAGEFRAME1.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME2.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME3.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME4.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME5.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME6.PAGE1.FONTSIZE=INT_015*9     
       THISFORM.PAGEFRAME7.PAGE1.FONTSIZE=INT_015*9 
       THISFORM.PAGEFRAME8.PAGE1.FONTSIZE=INT_015*9  
       THISFORM.PAGEFRAME9.PAGE1.FONTSIZE=INT_015*9  
       THISFORM.PAGEFRAME10.PAGE1.FONTSIZE=INT_015*9 
       THISFORM.PAGEFRAME11.PAGE1.FONTSIZE=INT_015*9 
       THISFORM.PAGEFRAME12.PAGE1.FONTSIZE=INT_015*9
       THISFORM.PAGEFRAME13.PAGE1.FONTSIZE=INT_015*9 
       THISFORM.PAGEFRAME14.PAGE1.FONTSIZE=INT_015*9
       THISFORM.KEY_LIST.VISIBLE=.F.
       THISFORM.SHRGROUP.VISIBLE=.F.
       THISFORM.SHRGROUP.ENABLED=.F.       
       THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F.                              
       THISFORM.SETALL('FONTSIZE',INT_015*9,'TEXTBOX')
       THISFORM.SETALL('FONTSIZE',INT_015*9,'LABEL')  
       THISFORM.SETALL('FONTSIZE',INT_015*9,'SPINNER')               
       THISFORM.SETALL('READONLY',.T.,'TEXTBOX')       
       THISFORM.SETALL('READONLY',.T.,'SPINNER')
       THISFORM.GRID1.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID2.BACKCOLOR=IIF(SEEK(sys_oper,'A01'),IIF(A01.F11=0,10009146,A01.F11),10009146)
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID1.SETALL('FONTSIZE',INT_015*9,'HEADER')
       THISFORM.GRID2.SETALL('FONTSIZE',INT_015*9,'COLUMN')
       THISFORM.GRID2.SETALL('FONTSIZE',INT_015*9,'HEADER')      
       THISFORM.GRID1.COLUMN1.WIDTH=INT_015*185                         
       THISFORM.GRID1.COLUMN1.HEADER1.CAPTION='設定項目'
       THISFORM.GRID1.COLUMN1.CONTROLSOURCE='S00_1.F02'    
       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*55
       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*55
       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*30
       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55
       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*55
       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*55
       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*140
       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*55  
       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50   
       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50        
       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='加班編號'
       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='時段說明'
       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='單位'
       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='計算基準'
       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='計算費率'
       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='計算金額'
       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='每次加班津貼補助'
       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='補助金額'  
       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION=''  
       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION=''  
       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION=''  
       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION=''  
       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F05'
       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(S00.F06='1','平日加班','假日加班')"
       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(S00.F07='H','時','日')"
       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F08),'',BASE_CLASS(S00.F08))"
       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F09),'',S00.F09)"
       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F10),'',S00.F10)"
       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F46'
       THISFORM.GRID2.COLUMN8.CONTROLSOURCE="IIF(!EMPTY(S00.F47),S00.F47,'')"
       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F26'
       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F26'
       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F26'
       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F26'
       THISFORM.GRID1.READONLY=.T.       
       THISFORM.GRID2.READONLY=.T. 
       THISFORM.GRID1.SETFOCUS
       THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
       THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
       THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.   
       THISFORM.ED_BOTT.ENABLED=.F.    
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
       THISFORM.ORPGROUP.PNT_BOTT.ENABLED=IIF(A02.F07='*',.T.,.F.)       
       JK=''
  ENDPROC     
  ********
PROC GRID1.AFTERROWCOLCHANGE
       LPARAMETERS NCOLINDEX
       AREA1='S00_1'
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
       THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
       THISFORM.ED_BOTT.ENABLED=.F.
       THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.    
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")      
       DO CASE
             CASE S00_1.F01='B'
                        SELECT S00
                        SET ORDER TO S001   
                       THISFORM.PAGEFRAME1.VISIBLE=.T.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*30
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*140
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*55  
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50   
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50                         
		       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='加班編號'
		       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='時段說明'
		       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='單位'
		       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='計算基準'
		       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='計算費率'
		       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='計算金額'
       		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='每次加班津貼補助'
       		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='補助金額'  
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION=''  
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION=''  
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION=''  
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION=''  
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(S00.F06='1','平日加班','假日加班')"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(S00.F07='H','時','日')"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F08),'',BASE_CLASS(S00.F08))"
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F09),'',S00.F09)"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F10),'',S00.F10)"
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F46'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE="IIF(!EMPTY(S00.F47),S00.F47,'')"
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F26'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F26'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F26'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F26'        
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=S00.F05
         	       THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE=VAL(S00.F06)
         	       THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=IIF(S00.F07='H',1,2)
         	       THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=IIF(!EMPTY(S00.F08),1,2)
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=BASE_CLASS(S00.F08)
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE=IIF(EMPTY(S00.F09),'',S00.F09)
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=IIF(EMPTY(S00.F10),'',S00.F10)
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE=S00.F46
         	       THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE=IIF(EMPTY(S00.F47),'',S00.F47)
		       THISFORM.PAGEFRAME1.PAGE1.LBLF991.CAPTION=S00.F99		                    
             CASE S00_1.F01='C'       
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.T.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.       
                       THISFORM.PAGEFRAME14.VISIBLE=.F.       
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*65  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*90 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='遲到分鐘數' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='扣除全勤' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='額外扣款分鐘數' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='扣款底數' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='扣款比例' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='扣款金額' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F11'
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(S00.F12='Y','是','否')"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F14),'',S00.F14)"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))"
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F16),'',S00.F16)"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F17),'',S00.F17)"
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                                   
		       THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE=S00.F11      
		       THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=IIF(S00.F12='Y',1,2)   
		       THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=IIF(S00.F13<>'Y',1,2)              
		       THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE=IIF(S00.F13<>'Y',0,S00.F14)                   
		       THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=IIF(EMPTY(S00.F15),2,1)                    
		       THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))             
		       THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=IIF(EMPTY(S00.F16),'',S00.F16)              
		       THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=IIF(EMPTY(S00.F17),'',S00.F17)
		       THISFORM.PAGEFRAME2.PAGE1.LBLF991.CAPTION=S00.F99  
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'C'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='C'   
             CASE S00_1.F01='D'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.T.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.    
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*65  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*90 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='遲到分鐘數' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='扣除全勤' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='額外扣款分鐘數' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='扣款底數' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='扣款比例' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='扣款金額' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F11'
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(S00.F12='Y','是','否')"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F14),'',S00.F14)"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))"
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F16),'',S00.F16)"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F17),'',S00.F17)"
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                           
                       THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE=S00.F11  
                       THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=IIF(S00.F12='Y',1,2)
                       THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=IIF(S00.F13<>'Y',1,2)	
                       THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE=IIF(S00.F13<>'Y',0,S00.F14)
                       THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=IIF(EMPTY(S00.F15),2,1)  
                       THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))
                       THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=IIF(EMPTY(S00.F16),'',S00.F16)        
                       THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=IIF(EMPTY(S00.F17),'',S00.F17)  
                       THISFORM.PAGEFRAME3.PAGE1.LBLF991.CAPTION=S00.F99        
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'D'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='D'                         
             CASE S00_1.F01='E'
                        SELECT S00
                        SET ORDER TO S002                    
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.T.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.  
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*70
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*75         
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50   		       
		       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='年資範圍(年)'
		       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='給予年假天數'        
                       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 		         
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="ALLTRIM(ALLTRIM(STR(S00.F18))+'～'+ALLTRIM(STR(S00.F19)))"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F20'      
                       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'    		                                                
         	       THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE=S00.F18
         	       THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE=S00.F19
         	       THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE=S00.F20   
		       THISFORM.PAGEFRAME4.PAGE1.LBLF991.CAPTION=S00.F99                         
             CASE S00_1.F01='F'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.T.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.   
                       THISFORM.PAGEFRAME14.VISIBLE=.F.           
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*180  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*150 
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*140 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='最大年資限制' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='年假上限限制' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='到職未滿一年' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="IIF(EMPTY(S00.F21),'','任職滿'+ALLTRIM(STR(S00.F21))+'年且每年增加年假'+ALLTRIM(STR(S00.F22))+'天')"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(EMPTY(S00.F23),'','年假最高給假天數'+ALLTRIM(STR(S00.F23))+'天(含)')"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F24),'以年假'+ALLTRIM(STR(S00.F25))+'天乘上到職月數','給予年假'+ALLTRIM(STR(S00.F24))+'天')"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                                      
		       THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=IIF(EMPTY(S00.F21),1,2)
		       THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE=IIF(EMPTY(S00.F21),'',S00.F21)
		       THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE=IIF(EMPTY(S00.F21),'',S00.F22)       
		       THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=IIF(EMPTY(S00.F23),1,2)       
		       THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE=IIF(EMPTY(S00.F23),'',S00.F23) 
		       THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=IIF(EMPTY(S00.F25),1,2)       
		       THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE=IIF(EMPTY(S00.F25),S00.F24,'')
		       THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE=IIF(EMPTY(S00.F25),'',S00.F25)   
		       THISFORM.PAGEFRAME5.PAGE1.LBLF991.CAPTION=S00.F99   
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'F'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='F'  		                             		                         	                                               
             CASE S00_1.F01='G'
                        SELECT S00
                        SET ORDER TO S003                  
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.T.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.    
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*90
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*40  
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50     
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*70
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*80
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50  
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='假別編號'
		       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='假別項目'
		       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='單位'      
		       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='請假基數'  
                       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='是否給薪'  	  	       
		       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='扣款底數'     
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='扣款比例(％)'                
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='扣款金額'
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='影響全勤'                     
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='是否計算假日' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION=''   
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION=''          
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F26'  
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F27'    
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(S00.F28='H','小時',IIF(S00.F28='S','次/月','日'))"  
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F62'     
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(S00.F58='1','是','否')"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F30),'',BASE_CLASS(S00.F30))"
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE="IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),'',IIF(EMPTY(S00.F31),'',S00.F31))"   
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE="IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),'',IIF(EMPTY(S00.F32),'',S00.F32))"
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE="IIF(S00.F33='Y','是','否')"  
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE="IIF(S00.F68='1','是','否')"
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'
		       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)     	   
		       THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)    
		       THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)      		       
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE=S00.F26
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE=S00.F27
         	       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=IIF(S00.F28='H',2,IIF(S00.F28='S',3,1))      	       
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE=IIF(EMPTY(S00.F29),'',S00.F29)
         	       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=IIF(EMPTY(S00.F29),1,2)
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF30.VALUE=BASE_CLASS(S00.F30)
         	       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),1,IIF(EMPTY(S00.F30),3,2))
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=IIF(EMPTY(S00.F31),'',S00.F31)
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=IIF(EMPTY(S00.F32),'',S00.F32)     
		       THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=IIF(S00.F33='Y',1,2) 
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=IIF(S00.F33='Y',S00.F34,'')
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=IIF(S00.F33='Y',S00.F35,'')
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=IIF(S00.F33='Y',S00.F36,'')
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=IIF(S00.F33='Y',S00.F37,'')
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=IIF(S00.F33='Y',S00.F38,'')
         	       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=VAL(S00.F58)
         	       THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE=VAL(S00.F68)
         	       THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=S00.F62
		       THISFORM.PAGEFRAME6.PAGE1.LBLF991.CAPTION=S00.F99
		       DO CASE		&&計數單位
		       	      CASE S00.F28='D'	&&日
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='天(不含)'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝天數               天(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤天數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤天數             ，扣款比例                    ％ '		       	      		 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 日'  
		       	      CASE S00.F28='H'	&&小時
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='小時(不含)'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝時數               時(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤時數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤時數             ，扣款比例                    ％ '		       	      		 		       	      
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 小時' 
		       	      CASE S00.F28='S'	&&次
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='次(不含)'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝次數               次(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤次數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤次數             ，扣款比例                    ％ '		       	      		 		       	      
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 次' 
		       ENDCASE	                                                                          
             CASE S00_1.F01='H'
                        SELECT S00
                        SET ORDER TO S004                    
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.T.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.     
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*120
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*70
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50  
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*80
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*260   
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50   	                                                                     
		       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='懲處編號'
		       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='懲處項目'
		       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='扣款底數'     
		       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='扣款比例(％)'                
		       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='扣款金額'  
		       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='連續扣款月數'
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='備註說明'         
                       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
                       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 		           
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F39'  
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F40'    
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F42),'',BASE_CLASS(S00.F42))"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F43),'',S00.F43)"    
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F44),'',S00.F44)"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F41'   
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F45'    
                       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
                       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'   	
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE=S00.F39
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE=S00.F40
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF41.VALUE=S00.F41
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF42.VALUE=BASE_CLASS(S00.F42)
         	       THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=IIF(EMPTY(S00.F43),2,1)
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=IIF(EMPTY(S00.F43),'',S00.F43)
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=IIF(EMPTY(S00.F44),'',S00.F44)
         	       THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE=S00.F45        	             	       
		       THISFORM.PAGEFRAME7.PAGE1.LBLF991.CAPTION=S00.F99	                       	                                           
             CASE S00_1.F01='I'
                        SELECT S00
                        SET ORDER TO S005                  
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.T.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.         
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*140
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*65
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*65  
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*65 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*100 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*100 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*100 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*95 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*95   	       
		       THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='班別編號'
		       THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='班別說明'
		       THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='上班時段'     
		       THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='下班時段'         
		       THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='休息時段起'  
		       THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='休息時段止'        
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='上班總時數'     
                       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='上班遲到緩衝時間' 
                       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='下班早退緩衝時間' 
                       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='加班上班緩衝時間' 
                       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='上班刷卡時段止' 
                       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='下班刷卡時段起' 		                                    
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F48'  
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F49'    
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F50' 
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F51' 
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F52' 
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F53' 
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F57' 
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F54'        
                       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F55'
                       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F56'
                       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F66'
                       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F67'	
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE=S00.F48
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE=S00.F49
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE=S00.F50
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE=S00.F51
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE=S00.F52
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE=S00.F53
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE=S00.F54  	
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE=S00.F55 
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE=S00.F56     
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=S00.F57  
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE=S00.F66
         	       THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE=S00.F67   
		       THISFORM.PAGEFRAME8.PAGE1.LBLF991.CAPTION=S00.F99	 	
             CASE S00_1.F01='J'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.T.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.
                       THISFORM.PAGEFRAME14.VISIBLE=.F.                    
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*80  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='扣稅率（％）' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F59'
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                         
		       THISFORM.PAGEFRAME9.PAGE1.TXTF59.VALUE=S00.F59         
		       THISFORM.PAGEFRAME9.PAGE1.LBLF991.CAPTION=S00.F99 
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'J'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='J' 
             CASE S00_1.F01='K'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.T.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.
                       THISFORM.PAGEFRAME14.VISIBLE=.F.                    
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*100  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='年終獎金計算方式' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="IIF(S00.F68='1','方式一',IIF(S00.F68='2','方式二','方式三'))"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'
		       DO CASE 
		              CASE S00.F68 = '1' 
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = S00.F69
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = ''		                         
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''
		              CASE S00.F68 = '2' 
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = S00.F69
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = ''		                         
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''
		              CASE S00.F68 = '3' 
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = S00.F69
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = S00.F71
		             OTHERWISE 
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''	     
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''	
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''	
		       ENDCASE
		       THISFORM.PAGEFRAME10.PAGE1.TXTF72.VALUE = S00.F72
		       THISFORM.PAGEFRAME10.PAGE1.LBLF991.CAPTION=S00.F99 
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'K'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='K' 	
             CASE S00_1.F01='L'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.T.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.
                       THISFORM.PAGEFRAME14.VISIBLE=.F.                    
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*120  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='未休年假獎金計算方式' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="IIF(S00.F68='1','方式一',IIF(S00.F68='2','方式二','方式三'))"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'
		       DO CASE 
		              CASE S00.F68 = '1' 
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = S00.F69                         
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''
		              CASE S00.F68 = '2' 
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2                         
		                          THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''
		              CASE S00.F68 = '3' 
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''       
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''   		                          
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = S00.F71
		             OTHERWISE 
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''	     
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''	
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''	
		       ENDCASE
		       THISFORM.PAGEFRAME11.PAGE1.TXTF72.VALUE = S00.F72		       
		       THISFORM.PAGEFRAME11.PAGE1.LBLF991.CAPTION=S00.F99 
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'L'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='L' 		       	       
             CASE S00_1.F01='X'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.T.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.          
                       THISFORM.PAGEFRAME14.VISIBLE=.F.
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*50
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                               
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_000.VALUE=IIF(INDEXSEEK('INT_000',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_071.VALUE=IIF(INDEXSEEK('INT_071',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_076.VALUE=IIF(INDEXSEEK('INT_076',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_077.VALUE=IIF(INDEXSEEK('INT_077',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_078.VALUE=IIF(INDEXSEEK('INT_078',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_079.VALUE=IIF(INDEXSEEK('INT_079',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_080.VALUE=IIF(INDEXSEEK('INT_080',.T.,'A26','A26'),ALLTRIM(A26.F06),'')
		       THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_107.VALUE=IIF(INDEXSEEK('INT_107',.T.,'A26','A26'),ALLTRIM(A26.F06),'')         
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.  
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.)	    
             CASE S00_1.F01='Y'		&&加班參數設定-Y
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.T.          
                       THISFORM.PAGEFRAME14.VISIBLE=.F.	
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*230
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*230
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*240
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='加班參數一' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='加班參數二' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='加班計算是否依S19.加班申請(修改)作業為主' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="IIF(ALLTRIM(S00.F63)<>'F','加班時間不滿'+ALLTRIM(S00.F63)+'(含)分鐘,則以30分鐘計算','不設定')"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(ALLTRIM(S00.F64)<>'F','加班時間若滿'+ALLTRIM(S00.F64)+'(含)分鐘,則以60分鐘計算','不設定')"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F65) OR S00.F65='2','否','是')"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'     
                       THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE=IIF(ALLTRIM(S00.F63)='F','',ALLTRIM(S00.F63))
                       THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE=IIF(ALLTRIM(S00.F64)='F','',ALLTRIM(S00.F64))
		       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=IIF(ALLTRIM(S00.F63)<>'F',2,1)     
		       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=IIF(ALLTRIM(S00.F64)<>'F',2,1)                       
                       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.VALUE=IIF(EMPTY(S00.F65) OR S00.F65='2',2,1)                         
                       THISFORM.PAGEFRAME13.PAGE1.LBLF991.CAPTION=S00.F99 
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'Y'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='Y'  		                                              
             CASE S00_1.F01='Z'
                       THISFORM.PAGEFRAME1.VISIBLE=.F.
                       THISFORM.PAGEFRAME2.VISIBLE=.F.
                       THISFORM.PAGEFRAME3.VISIBLE=.F.
                       THISFORM.PAGEFRAME4.VISIBLE=.F.
                       THISFORM.PAGEFRAME5.VISIBLE=.F.
                       THISFORM.PAGEFRAME6.VISIBLE=.F.
                       THISFORM.PAGEFRAME7.VISIBLE=.F.
                       THISFORM.PAGEFRAME8.VISIBLE=.F.
                       THISFORM.PAGEFRAME9.VISIBLE=.F.
                       THISFORM.PAGEFRAME10.VISIBLE=.F.
                       THISFORM.PAGEFRAME11.VISIBLE=.F.                       
                       THISFORM.PAGEFRAME12.VISIBLE=.F.
                       THISFORM.PAGEFRAME13.VISIBLE=.F.
                       THISFORM.PAGEFRAME14.VISIBLE=.T.                    
		       THISFORM.GRID2.COLUMN1.WIDTH=INT_015*75  
		       THISFORM.GRID2.COLUMN2.WIDTH=INT_015*85
		       THISFORM.GRID2.COLUMN3.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN4.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN5.WIDTH=INT_015*55 
		       THISFORM.GRID2.COLUMN6.WIDTH=INT_015*75 
		       THISFORM.GRID2.COLUMN7.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN8.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN9.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN10.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN11.WIDTH=INT_015*50 
		       THISFORM.GRID2.COLUMN12.WIDTH=INT_015*50
	               THISFORM.GRID2.COLUMN1.HEADER1.CAPTION='每月共計(天)' 
	               THISFORM.GRID2.COLUMN2.HEADER1.CAPTION='每天共計(小時)' 
	               THISFORM.GRID2.COLUMN3.HEADER1.CAPTION='全勤獎金' 
	               THISFORM.GRID2.COLUMN4.HEADER1.CAPTION='伙食費' 
	               THISFORM.GRID2.COLUMN5.HEADER1.CAPTION='交通費' 
	               THISFORM.GRID2.COLUMN6.HEADER1.CAPTION='報表欄位數' 
		       THISFORM.GRID2.COLUMN7.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN8.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN9.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN10.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN11.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN12.HEADER1.CAPTION='' 
		       THISFORM.GRID2.COLUMN1.CONTROLSOURCE="IIF(EMPTY(S00.F02),'',S00.F02)"
		       THISFORM.GRID2.COLUMN2.CONTROLSOURCE="IIF(EMPTY(S00.F03),'',S00.F03)"
		       THISFORM.GRID2.COLUMN3.CONTROLSOURCE="IIF(EMPTY(S00.F04),'',S00.F04)"
		       THISFORM.GRID2.COLUMN4.CONTROLSOURCE="IIF(EMPTY(S00.F60),'',S00.F60)"
		       THISFORM.GRID2.COLUMN5.CONTROLSOURCE="IIF(EMPTY(S00.F61),'',S00.F61)"
		       THISFORM.GRID2.COLUMN6.CONTROLSOURCE="IIF(EMPTY(S00.F61),0,S00.F73)"
		       THISFORM.GRID2.COLUMN7.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN8.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN9.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN10.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN11.CONTROLSOURCE='S00.F05'
		       THISFORM.GRID2.COLUMN12.CONTROLSOURCE='S00.F05'                         
		       THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE=S00.F02   
		       THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE=S00.F03   
		       THISFORM.PAGEFRAME14.PAGE1.TXTF04.VALUE=S00.F04       
		       THISFORM.PAGEFRAME14.PAGE1.TXTF60.VALUE=S00.F60 
		       THISFORM.PAGEFRAME14.PAGE1.TXTF61.VALUE=S00.F61       
		       THISFORM.PAGEFRAME14.PAGE1.LBLF991.CAPTION=S00.F99 
		       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.) AND S00.F01<>'Z'   
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='Z'  		                                                  
       ENDCASE   
       FCH='GRID1'
       CHK=''                    
  ENDPROC
 ****   
   PROC GRID1.MOUSEENTER
       LPARAMETERS nButton, nShift, nXCoord, nYCoord
         IF THISFORM.SHRGROUP.VISIBLE=.F. 
             THISFORM.GRID1.ENABLED=.T. 
         ELSE    
             THISFORM.GRID1.ENABLED=.F.                         
         ENDIF    
 ENDPROC  
*********          
PROC GRID2.AFTERROWCOLCHANGE
       LPARAMETERS NCOLINDEX
       AREA2='S00'
       R=RECNO('&AREA1')
       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
       THISFORM.GRID1.ENABLED=.F. 
       IF THISFORM.GRID2.ACTIVEROW=0  
           THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
           THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
           THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
           THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.       
           THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
           THISFORM.ED_BOTT.ENABLED=.F.
           THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
       ELSE
           THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.T.      
           THISFORM.CMDGROUP.UP_BOTT.ENABLED=.T.
           THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.T.
           THISFORM.CMDGROUP.END_BOTT.ENABLED=.T.            
           THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) 
           THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) 
           THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.) 
       ENDIF     
       THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
       DO CASE
             CASE  S00.F01='B'  &&加班費設定-B
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=S00.F05
         	        THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE=VAL(S00.F06)
         	        THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=IIF(S00.F07='H',1,2)
         	        THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=IIF(!EMPTY(S00.F08),1,2)
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=BASE_CLASS(S00.F08)
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE=IIF(EMPTY(S00.F09),'',S00.F09)
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=IIF(EMPTY(S00.F10),'',S00.F10)
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE=S00.F46
         	        THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE=IIF(EMPTY(S00.F47),'',S00.F47)
		        THISFORM.PAGEFRAME1.PAGE1.LBLF991.CAPTION=S00.F99		
		        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.) AND S00.F05<>'P1' 	 AND S00.F05<>'P2'   	        
            CASE  S00.F01='C'  &&遲到設定-C
		        THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE=S00.F11      
		        THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=IIF(S00.F12='Y',1,2)   
		        THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=IIF(S00.F13<>'Y',1,2)              
		        THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE=IIF(S00.F13<>'Y',0,S00.F14)                   
		        THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=IIF(EMPTY(S00.F15),2,1)                    
		        THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))             
		        THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=IIF(EMPTY(S00.F16),'',S00.F16)              
		        THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=IIF(EMPTY(S00.F17),'',S00.F17)
		        THISFORM.PAGEFRAME2.PAGE1.LBLF991.CAPTION=S00.F99  
		       IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		       ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		       ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='C'  			        
            CASE  S00.F01='D'  &&早退設定-D
    			THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE=S00.F11  
                        THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=IIF(S00.F12='Y',1,2)
                        THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=IIF(S00.F13<>'Y',1,2)	
                        THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE=IIF(S00.F13<>'Y',0,S00.F14)
                        THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=IIF(EMPTY(S00.F15),2,1)  
                        THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=IIF(EMPTY(S00.F15),'',BASE_CLASS(S00.F15))
                        THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=IIF(EMPTY(S00.F16),'',S00.F16)        
                        THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=IIF(EMPTY(S00.F17),'',S00.F17)  
                        THISFORM.PAGEFRAME3.PAGE1.LBLF991.CAPTION=S00.F99         
		       IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		       ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		       ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='D'  	                        
            CASE  S00.F01='E'   &&年假範圍設定-E		  
         	        THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE=S00.F18
         	        THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE=S00.F19
         	        THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE=S00.F20   
		        THISFORM.PAGEFRAME4.PAGE1.LBLF991.CAPTION=S00.F99 	
	    CASE   S00.F01='F'   &&年假限制設定-F
		        THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=IIF(EMPTY(S00.F21),1,2)
		        THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE=IIF(EMPTY(S00.F21),'',S00.F21)
		        THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE=IIF(EMPTY(S00.F21),'',S00.F22)       
		        THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=IIF(EMPTY(S00.F23),1,2)       
		        THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE=IIF(EMPTY(S00.F23),'',S00.F23) 
		        THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=IIF(EMPTY(S00.F25),1,2)       
		        THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE=IIF(EMPTY(S00.F25),S00.F24,'')
		        THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE=IIF(EMPTY(S00.F25),'',S00.F25)   
		        THISFORM.PAGEFRAME5.PAGE1.LBLF991.CAPTION=S00.F99   	
		       IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		       ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		       ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='F'  			          	       
            CASE  S00.F01='G'   &&請假扣款項目-G		
		      	THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)     	   
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)    
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		      	THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)          	       
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE=S00.F26
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE=S00.F27
         	        THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=IIF(S00.F28='H',2,IIF(S00.F28='S',3,1))      	       
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE=IIF(EMPTY(S00.F29),'',S00.F29)
         	        THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=IIF(EMPTY(S00.F29),1,2)
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF30.VALUE=BASE_CLASS(S00.F30)
         	        THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),1,IIF(EMPTY(S00.F30),3,2))
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=IIF(EMPTY(S00.F31),'',S00.F31)
		        THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=IIF(EMPTY(S00.F32),'',S00.F32)     
		        THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=IIF(S00.F33='Y',1,2) 
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=IIF(S00.F33='Y',S00.F34,'')
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=IIF(S00.F33='Y',S00.F35,'')
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=IIF(S00.F33='Y',S00.F36,'')
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=IIF(S00.F33='Y',S00.F37,'')
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=IIF(S00.F33='Y',S00.F38,'')
         	        THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=VAL(S00.F58)
         	        THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE=VAL(S00.F68)
         	        THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=S00.F62
		        THISFORM.PAGEFRAME6.PAGE1.LBLF991.CAPTION=S00.F99
		         DO CASE		&&計數單位
		       	       CASE S00.F28='D'	&&日
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='天'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝天數               天(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤天數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤天數             ，扣款比例                    ％ '		       	      		 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 日' 
		       	       CASE S00.F28='H'	&&小時
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='小時'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝時數               時(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤時數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤時數             ，扣款比例                    ％ '		       	      		 		       	      
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 小時' 
		       	       CASE S00.F28='S'	&&次
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='次/月'
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝次數               次(含)' 
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤次數             ，扣款比例                    ％ '
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤次數             ，扣款比例                    ％ '		       	      		 		       	      
		       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 次/月' 
		         ENDCASE		        
		        THISFORM.ORPGROUP.DEL_BOTT.ENABLED=IIF(A02.F06='*',.T.,.F.) AND S00.F26<>'A1' 	 AND S00.F26<>'A2'   	       
            CASE  S00.F01='H'    &&懲處扣款設定-H	  
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE=S00.F39
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE=S00.F40
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF41.VALUE=S00.F41
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF42.VALUE=BASE_CLASS(S00.F42)
         	        THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=IIF(EMPTY(S00.F43),2,1)
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=IIF(EMPTY(S00.F43),'',S00.F43)
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=IIF(EMPTY(S00.F44),'',S00.F44)
         	        THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE=S00.F45        	             	       
		        THISFORM.PAGEFRAME7.PAGE1.LBLF991.CAPTION=S00.F99		
            CASE  S00.F01='I'    &&班別設定-I	  
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE=S00.F48
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE=S00.F49
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE=S00.F50
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE=S00.F51
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE=S00.F52
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE=S00.F53
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE=S00.F54 
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE=S00.F55  
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE=S00.F56   	  
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=S00.F57     
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE=S00.F66   
         	        THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE=S00.F67   
		        THISFORM.PAGEFRAME8.PAGE1.LBLF991.CAPTION=S00.F99	
            CASE  S00.F01='J'    &&其他所得發放設定-J	  
         	        THISFORM.PAGEFRAME9.PAGE1.TXTF59.VALUE=S00.F59		
         	        THISFORM.PAGEFRAME9.PAGE1.LBLF991.CAPTION=S00.F99         
		        IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		        ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		        ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='J'  	
            CASE  S00.F01='K'    &&年終獎金發放設定-K	  
         	        THISFORM.GRID2.SETFOCUS	
		        IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		        ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		        ENDIF         	        	        
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='K'  	            	  
		        DO CASE
		               CASE S00.F68 = '1'
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = S00.F69
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE =  ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE =  ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''
		               CASE S00.F68 = '2'
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = S00.F69
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE =  ''		                         
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''
		               CASE S00.F68 = '3'
		                         THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = S00.F69		                         
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = S00.F71
		        ENDCASE
		        THISFORM.PAGEFRAME10.PAGE1.TXTF72.VALUE = S00.F72
         	        THISFORM.PAGEFRAME10.PAGE1.LBLF991.CAPTION=S00.F99  
         	         THISFORM.GRID2.REFRESH	
            CASE  S00.F01='L'    &&未休年假獎金發放設定-L	  	      
         	        THISFORM.GRID2.SETFOCUS
		        IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		        ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		        ENDIF     		        
		        THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='L'  	
		        DO CASE
		               CASE S00.F68 = '1'
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = S00.F69
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''
		               CASE S00.F68 = '2'
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = BASE_CLASS(S00.F70)
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''
		               CASE S00.F68 = '3'
		                         THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''       
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''     
		                         THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = S00.F71
		        ENDCASE
		        THISFORM.PAGEFRAME11.PAGE1.TXTF72.VALUE = S00.F72
         	        THISFORM.PAGEFRAME11.PAGE1.LBLF991.CAPTION=S00.F99
            CASE  S00.F01='Y'    &&加班參數設定-Y	  
                       THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE=IIF(ALLTRIM(S00.F63)='F','',ALLTRIM(S00.F63))
                       THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE=IIF(ALLTRIM(S00.F64)='F','',ALLTRIM(S00.F64)) 
		       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=IIF(ALLTRIM(S00.F63)<>'F',2,1)     
		       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=IIF(ALLTRIM(S00.F64)<>'F',2,1)                        
                       THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.VALUE=IIF(EMPTY(S00.F65) OR S00.F65='2',2,1)
		       THISFORM.PAGEFRAME13.PAGE1.LBLF991.CAPTION=S00.F99 
		       IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		       ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		       ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='Y'  			        
            CASE  S00.F01='Z'    &&基本參數設定-Z	  
		       THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE=S00.F02   
		       THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE=S00.F03   
		       THISFORM.PAGEFRAME14.PAGE1.TXTF04.VALUE=S00.F04       
		       THISFORM.PAGEFRAME14.PAGE1.TXTF60.VALUE=S00.F60 
		       THISFORM.PAGEFRAME14.PAGE1.TXTF61.VALUE=S00.F61   
		       THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE=S00.F73    
		       THISFORM.PAGEFRAME14.PAGE1.LBLF991.CAPTION=S00.F99 
		       IF  THISFORM.GRID2.ACTIVEROW=0  
		             THISFORM.ORPGROUP.NEW_BOTT.ENABLED=IIF(A02.F04='*',.T.,.F.)
		       ELSE
		            THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
		       ENDIF
		       THISFORM.ED_BOTT.ENABLED=IIF(A02.F05='*',.T.,.F.) AND S00.F01='Z'  		        	
       ENDCASE                    
       FCH='GRID2'
       JK=THIS.COLUMN1.HEADER1.CAPTION
       CHK=''           
  ENDPROC  
*******新增
PROCEDURE ORPGROUP.NEW_BOTT.CLICK 
  R=RECNO('&AREA1')
  THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")  
  THISFORM.GRID1.ENABLED=.F.   
  THISFORM.GRID2.ENABLED=.F. 
  THISFORM.ED_BOTT.ENABLED=.F. 	
  DO CASE
         CASE  S00_1.F01='B'  &&加班費設定-B      
         	    THISFORM.PAGEFRAME1.PAGE1.SETALL('READONLY',.F.,'TEXTBOX') 
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.ENABLED=.T.
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.ENABLED=.T.	
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.ENABLED=.T.    
		    THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VISIBLE=.T.
		    THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VISIBLE=.T.		       	            	         
         	    THISFORM.PAGEFRAME1.PAGE1.TXTF46.VISIBLE=.F.  
         	    THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=.T.          	       
		    THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE=''
		    THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=''	
		    THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE=0		              
		    THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=''				       
		    THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE=''	
		    THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE=0	             
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE=1			           
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=1	
		    THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=1
		    THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=1		       		       
		    THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE=1	   		              
		    THISFORM.PAGEFRAME1.PAGE1.LBLF991.CAPTION=''		
		    THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS	
         CASE  S00_1.F01='C'  &&遲到設定-C 
         	    THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.READONLY=.F.
         	    THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.ENABLED=.T.      
         	    THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.ENABLED=.T.   
		    THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE=0           
		    THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=1	       
		    THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=1	            
		    THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE=0                
		    THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1                    
		    THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=''             
		    THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=0           
		    THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=0 
		    THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.SETFOCUS  
		    THISFORM.PAGEFRAME2.PAGE1.LBLF991.CAPTION=''                
         CASE  S00_1.F01='D'  &&早退設定-D   
         	    THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.READONLY=.F.
         	    THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.ENABLED=.T.   
         	    THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.ENABLED=.T.      
       		    THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE=0
       		    THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=1
       		    THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=1
       		    THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE=0
       		    THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1
       		    THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=''
       		    THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=0
       		    THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=0 
       		    THISFORM.PAGEFRAME3.PAGE1.LBLF991.CAPTION=''       		    
       		    THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.SETFOCUS
         CASE  S00_1.F01='E'   &&年假範圍設定-E       	       
         	    THISFORM.PAGEFRAME4.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
		    THISFORM.PAGEFRAME4.PAGE1.TXTF18.READONLY=.F.    
		    THISFORM.PAGEFRAME4.PAGE1.TXTF19.READONLY=.F.    	
		    THISFORM.PAGEFRAME4.PAGE1.TXTF20.READONLY=.F. 
		    THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE=0
		    THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE=0	
		    THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE=0 
		    THISFORM.PAGEFRAME4.PAGE1.LBLF991.CAPTION=''
		    THISFORM.PAGEFRAME4.PAGE1.TXTF18.SETFOCUS		    
         CASE  S00_1.F01='F'   &&年假限制設定-F	
       		    THISFORM.PAGEFRAME5.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
       		    THISFORM.PAGEFRAME5.PAGE1.TXTF24.READONLY=.F.
       		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.ENABLED=.T.  
       		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.ENABLED=.T.  
       		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.ENABLED=.T.  
		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=1
		    THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE=0
		    THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE=0      
		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=1   
		    THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE=0
		    THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=1       
		    THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE=0
		    THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE=0 
		    THISFORM.PAGEFRAME4.PAGE1.LBLF991.CAPTION=''    
		    THISFORM.PAGEFRAME5.PAGE1.TXTF21.SETFOCUS 		                  	                	    
         CASE  S00_1.F01='G'   &&缺勤扣款項目-G 	           	                   	       
         	    THISFORM.PAGEFRAME6.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')          
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.F.   
		    THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=.T.  
		    THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=.T.
		    THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=.T.            	             	       	       
	            THISFORM.PAGEFRAME6.PAGE1.TXTF26.READONLY=.F.
	            THISFORM.PAGEFRAME6.PAGE1.TXTF27.READONLY=.F.     
	            THISFORM.PAGEFRAME6.PAGE1.TXTF62.READONLY=.F.
	            THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.ENABLED=.T.   
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.ENABLED=.T.	
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.ENABLED=.T.
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.ENABLED=.T.
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.ENABLED=.T.
		    THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.ENABLED=.T.
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=1		       
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=1	 
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=1 
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2	  
		    THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE=2	    		    
		    THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=2
		    THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE=''
		    THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE=''	
	            THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE=''		
		    THISFORM.PAGEFRAME6.PAGE1.TXTF30.VALUE=''
		    THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=''				       
		    THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=''
		    THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=''
		    THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=''
		    THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=''	 
		    THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=''	 
		    THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=''	
		    THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=0 
		    THISFORM.PAGEFRAME6.PAGE1.LBLF991.CAPTION=''			    		       
		    THISFORM.PAGEFRAME6.PAGE1.TXTF26.SETFOCUS
         CASE  S00_1.F01='H'     &&懲處扣款設定-H           	                     	       
         	    THISFORM.PAGEFRAME7.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')   	         	       
		    THISFORM.PAGEFRAME7.PAGE1.TXTF42.VISIBLE=.F.		
		    THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.ENABLED=.T.
		    THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=1			                    
		    THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VISIBLE=.T.
		    THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=1
		    THISFORM.PAGEFRAME7.PAGE1.TXTF44.READONLY=.T.
		    THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE=''
		    THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE=''	
		    THISFORM.PAGEFRAME7.PAGE1.TXTF41.VALUE=0			       
		    THISFORM.PAGEFRAME7.PAGE1.TXTF42.VALUE=''
		    THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=0
		    THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=''	
   		    THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE='' 		       
		    THISFORM.PAGEFRAME7.PAGE1.LBLF991.CAPTION=''
		    THISFORM.PAGEFRAME7.PAGE1.TXTF39.SETFOCUS		    	
         CASE  S00_1.F01='I'   &&班別設定-I        	       
         	    THISFORM.PAGEFRAME8.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')      
		    THISFORM.PAGEFRAME8.PAGE1.TXTF50.READONLY=.F.
		    THISFORM.PAGEFRAME8.PAGE1.TXTF51.READONLY=.F.
		    THISFORM.PAGEFRAME8.PAGE1.TXTF52.READONLY=.F.
		    THISFORM.PAGEFRAME8.PAGE1.TXTF53.READONLY=.F.
		    THISFORM.PAGEFRAME8.PAGE1.TXTF57.READONLY=.T.
		    THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE=''	
		    THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE=''	
		    THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=0
		    THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE=''
		    THISFORM.PAGEFRAME8.PAGE1.LBLF991.CAPTION=''			     
		    THISFORM.PAGEFRAME8.PAGE1.TXTF48.SETFOCUS
         CASE  S00_1.F01='J'   &&其他所得發放設定-J 	      
      		    THISFORM.PAGEFRAME9.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
		    THISFORM.PAGEFRAME9.PAGE1.TXTF59.VALUE=0     
		    THISFORM.PAGEFRAME9.PAGE1.LBLF991.CAPTION=''   
		    THISFORM.PAGEFRAME9.PAGE1.TXTF59.SETFOCUS     	
         CASE  S00_1.F01='K'   &&年終獎金發放設定-K 	      
      		    THISFORM.PAGEFRAME10.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
      		    THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.ENABLED=.T.
      		    THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
      		    THISFORM.PAGEFRAME10.PAGE1.TXTF701.VISIBLE = .F.
      		    THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VISIBLE = .T.
      		    THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
      		    THISFORM.PAGEFRAME10.PAGE1.TXTF691.READONLY = .F.
		    THISFORM.PAGEFRAME10.PAGE1.TXTF701.READONLY = .F.
		    THISFORM.PAGEFRAME10.PAGE1.TXTF72.READONLY = .F.
		    THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = 0
		    THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = ''		    
		    THISFORM.PAGEFRAME10.PAGE1.LBLF991.CAPTION=''   
		    THISFORM.PAGEFRAME10.PAGE1.TXTF691.SETFOCUS   
         CASE  S00_1.F01='L'   &&未休年假獎金發放設定-L 	      
      		    THISFORM.PAGEFRAME11.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
      		    THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.ENABLED=.T.
      		    THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
      		    THISFORM.PAGEFRAME11.PAGE1.TXTF701.VISIBLE = .F.
      		    THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VISIBLE = .T.
      		    THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
      		    THISFORM.PAGEFRAME11.PAGE1.TXTF69.READONLY = .F.
		    THISFORM.PAGEFRAME11.PAGE1.TXTF701.READONLY = .F.
		    THISFORM.PAGEFRAME11.PAGE1.TXTF72.READONLY = .F.
		    THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = 0
		    THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = ''
		    THISFORM.PAGEFRAME11.PAGE1.LBLF991.CAPTION=''   
		    THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.SETFOCUS   		    		    	    
         CASE  S00_1.F01='Y'   &&加班參數設定-Y 	      
      		    THISFORM.PAGEFRAME13.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.ENABLED=.T.
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.ENABLED=.T.
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.ENABLED=.T.      		    
		    THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE=''     
		    THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE='' 
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=1
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=1		
		    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.VALUE=2
		    THISFORM.PAGEFRAME13.PAGE1.LBLF991.CAPTION=''   
		    THISFORM.PAGEFRAME13.PAGE1.SETFOCUS    	    
         CASE  S00_1.F01='Z'   &&基本參數設定-Z 	      
      		    THISFORM.PAGEFRAME14.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
      		    THISFORM.PAGEFRAME14.PAGE1.SETALL('READONLY',.F.,'SPINNER')
		    THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE=0
		    THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE=0   
		    THISFORM.PAGEFRAME14.PAGE1.TXTF04.VALUE=0     
		    THISFORM.PAGEFRAME14.PAGE1.TXTF60.VALUE=0 
		    THISFORM.PAGEFRAME14.PAGE1.TXTF61.VALUE=0  
		    THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE=9 
		    THISFORM.PAGEFRAME14.PAGE1.LBLF991.CAPTION=''   
		    THISFORM.PAGEFRAME14.PAGE1.SPINNER02.SETFOCUS     		      	    	       		       	       
     ENDCASE  	
ENDPROC
********新增確認
PROCEDURE SHRGROUP.SHURE1_BOTT.CLICK
  DO CASE
         CASE S00_1.F01='B' &&加班費設定-B
         	    IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE)     
                        =MESSAGEBOX('加班類別不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
                        RETURN
                    ENDIF
         	    IF LEN(ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE))<2
                         =MESSAGEBOX('加班類別不得少於兩碼!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
                        RETURN
                    ENDIF                       
         	    IF INDEXSEEK('B'+THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE,.F.,'S00','S001')=.T.   
                        =MESSAGEBOX('此加班類別已存在,請查明後再輸入!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE1.TXTF05.SETFOCUS
                        RETURN
                    ENDIF                
                    IF THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=1           
         	         IF THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=1     
                             =MESSAGEBOX('計算底數不得空白!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.SETFOCUS
                             RETURN
                         ENDIF    
         	         IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE)
                              =MESSAGEBOX('計算比例不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME1.PAGE1.TXTF09.SETFOCUS
                              RETURN
                         ENDIF     
                    ELSE
         	         IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE)
                             =MESSAGEBOX('計算金額不得為零!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
                             RETURN
                         ENDIF                        
                    ENDIF                              
         	    IF  THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE<>1 AND EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE)     
                         =MESSAGEBOX('您已選擇津貼補助項目,故補助金額不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME1.PAGE1.TXTF47.SETFOCUS
                         RETURN
                    ENDIF
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
 		        REPLACE F01 WITH 'B'		       
		        REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
		        REPLACE F06 WITH ALLTRIM(STR(THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE))
		        REPLACE F07 WITH IIF(THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=1,'H','D')
		        IF THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=1
		             DO CASE 
		                   CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=2
		                              REPLACE F08 WITH 'H'
		                   CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=3
		                              REPLACE F08 WITH 'D'
		                   CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=4
		                              REPLACE F08 WITH 'B'
		             ENDCASE    
		             REPLACE F09 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE
		        ELSE
		             REPLACE F10 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE
		        ENDIF
		        REPLACE F46 WITH ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE)
		        IF THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE<>1
		             REPLACE F47 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE
		        ENDIF
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF
		    THISFORM.ORPGROUP.NEW_BOTT.CLICK 
         CASE S00_1.F01='C'   &&遲到設定-C
                    IF THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=2
                         IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE)
	                      =MESSAGEBOX('您已選擇設定遲到額外扣款,故每月遲到"超過"分鐘數不得為零!',0+64,'提示視窗')
	                      THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.SETFOCUS
	                      RETURN
	                 ENDIF                        
 	       	         IF THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE<=THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE
                   	      =MESSAGEBOX('您已選擇設定額外扣款,故每月遲到"超過"分鐘數不得小於或等於每月遲到分鐘數!',0+64,'提示視窗')
  		             THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.SETFOCUS
                  	     RETURN
                         ENDIF  
                         IF THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1                         
	                      IF  THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=1
	                            =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款底數不得為空白!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.SETFOCUS
	                            RETURN
	                      ENDIF   
	                      IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE)
	                           =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款比列不得為零!',0+64,'提示視窗')
	                          THISFORM.PAGEFRAME2.PAGE1.TXTF16C.SETFOCUS
	                           RETURN
	                      ENDIF   	                          
	                  ELSE      
	                      IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE)
	                           =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款金額不得為零!',0+64,'提示視窗')
	                           THISFORM.PAGEFRAME2.PAGE1.TXTF17C.SETFOCUS
	                           RETURN
	                      ENDIF  
	                  ENDIF 
                    ENDIF                                   
		    SELECT S00
		    APPEND BLANK
		    LOCK()
		    IF RLOCK()
		        REPLACE F01 WITH  'C'                            
		        REPLACE F11 WITH THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE
		        REPLACE F12 WITH  IIF(THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=1,'Y','N')
		        IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=1
		              REPLACE F13 WITH 'N'	
		              REPLACE F14 WITH 0
		              REPLACE F15 WITH ''
		              REPLACE F16 WITH 0
		              REPLACE F17 WITH 0
		        ELSE
		              REPLACE F13 WITH 'Y'
		              REPLACE F14 WITH THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE
		              IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1 
		                    DO CASE 
			                  CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=2
			                             REPLACE F15 WITH 'H'
			                  CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=3
			                             REPLACE F15 WITH 'D'
			                  CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=4
			                             REPLACE F15 WITH 'B'	                              
			             ENDCASE
		            	     REPLACE F16 WITH THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE	
		            	     REPLACE F17 WITH 0 
		               ELSE
		            	     REPLACE F15 WITH ''
		            	     REPLACE F16 WITH 0
		            	     REPLACE F17 WITH THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE
		               ENDIF                 	  
		        ENDIF
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF     
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK                                 
         CASE S00_1.F01='D'   &&早退設定-D
                    IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=2
                          IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE)
	                       =MESSAGEBOX('您已選擇設定早退額外扣款,故每月早退"超過"分鐘數不得為零!',0+64,'提示視窗')
	                      THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.SETFOCUS
	                      RETURN
	                 ENDIF  
 	       	         IF THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE<=THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE
                             =MESSAGEBOX('您已選擇設定額外扣款,故每月早退"超過"分鐘數不得小於或等於每月早退分鐘數!',0+64,'提示視窗')
  		             THISFORM.PAGEFRAME1.PAGE1.SPINNER14D.SETFOCUS
                  	     RETURN
                         ENDIF  
                         IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1                         
	                       IF  THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=1
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款底數不得為空白!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.SETFOCUS
	                            RETURN
	                       ENDIF   
	                       IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE)
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款比列不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.TXTF16D.SETFOCUS
	                            RETURN
	                       ENDIF   	                          
	                 ELSE      
	                       IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE)
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款金額不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.TXTF17D.SETFOCUS
	                            RETURN
	                       ENDIF  
	                 ENDIF 
                    ENDIF     
		    SELECT S00
		    APPEND BLANK
		    LOCK()
		    IF RLOCK()
		        REPLACE F01 WITH  'D'                            
		        REPLACE F11 WITH THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE
		        REPLACE F12 WITH  IIF(THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=1,'Y','N')
		        IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=1
		              REPLACE F13 WITH 'N'	
		              REPLACE F14 WITH 0
		              REPLACE F15 WITH ''
		              REPLACE F16 WITH 0
		              REPLACE F17 WITH 0
		        ELSE
		              REPLACE F13 WITH 'Y'
		              REPLACE F14 WITH THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE
		              IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1 
		                    DO CASE 
			                  CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=2
			                             REPLACE F15 WITH 'H'
			                  CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=3
			                             REPLACE F15 WITH 'D'
			                  CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=4
			                             REPLACE F15 WITH 'B'	                              
			             ENDCASE
		            	     REPLACE F16 WITH THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE	
		            	     REPLACE F17 WITH 0 
		               ELSE
		            	     REPLACE F15 WITH ''
		            	     REPLACE F16 WITH 0
		            	     REPLACE F17 WITH THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE
		               ENDIF                 	  
		        ENDIF
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF   
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK                                             
         CASE S00_1.F01='E'   &&年假範圍設定-E
         	    IF EMPTY(THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE)     
                        =MESSAGEBOX('起始年資不得為零!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME4.PAGE1.TXTF18.SETFOCUS
                        RETURN
                    ENDIF                       
         	    IF EMPTY(THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE)     
                         =MESSAGEBOX('終止年資不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME4.PAGE1.TXTF19.SETFOCUS
                         RETURN
                    ENDIF    
         	    IF EMPTY(THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE)     
                         =MESSAGEBOX('給予年假天數不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME4.PAGE1.TXTF20.SETFOCUS
                         RETURN
                    ENDIF                       
         	    IF THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE>THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE     
                         =MESSAGEBOX('起始年資不得大於終止年資!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME4.PAGE1.TXTF18.SETFOCUS
                         RETURN
                    ENDIF
         	    IF INDEXSEEK('E'+ALLTRIM(STR(THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE))+ALLTRIM(STR(THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE)),.F.,'S00','S002')=.T.
                        =MESSAGEBOX('此年資範圍已存在,請查明後再輸入!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME4.PAGE1.TXTF18.SETFOCUS
                        RETURN
                    ENDIF  
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
 			 REPLACE F01 WITH 'E'		       
		         REPLACE F18 WITH THISFORM.PAGEFRAME4.PAGE1.TXTF18.VALUE
		         REPLACE F19 WITH THISFORM.PAGEFRAME4.PAGE1.TXTF19.VALUE
		         REPLACE F20 WITH THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE
		         REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF
		    THISFORM.ORPGROUP.NEW_BOTT.CLICK 
         CASE S00_1.F01='F'   &&年假限制設定-F		    
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2                       
         	          IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE)     
                              =MESSAGEBOX('您已選擇設定最大年資,故最大年資不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF21.SETFOCUS
                              RETURN
                          ENDIF                     
                   ENDIF
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=2                       
         	          IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE)     
                              =MESSAGEBOX('您已選擇設定年假上限,故年假上限不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF23.SETFOCUS
                              RETURN
                          ENDIF
                   ENDIF        
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=2                       
         	         IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE)     
                              =MESSAGEBOX('您已選擇到職未滿一年之,故乘上到職月數之年假天數不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF25.SETFOCUS
                              RETURN
                          ENDIF
                   ENDIF        
                   SELECT S00
                   APPEND BLANK
                   LOCK()
                   IF RLOCK()    
                        REPLACE F01 WITH 'F'
		        REPLACE F21 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE,0)
		        REPLACE F22 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE,0)
		        REPLACE F23 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE,0)
		        REPLACE F24 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=1,THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE,0)
		        REPLACE F25 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE,0)				      		        
		   ENDIF  
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           
         CASE S00_1.F01='G'   &&請假扣款項目-G  
         	    IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE)     
                        =MESSAGEBOX('假別編號不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME6.PAGE1.TXTF26.SETFOCUS
                        RETURN
                    ENDIF   
         	    IF LEN(ALLTRIM(THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE))<2
                        =MESSAGEBOX('假別編號不得少於兩碼!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE6.TXTF26.SETFOCUS
                        RETURN
                    ENDIF                         
         	    IF INDEXSEEK('G'+THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE,.F.,'S00','S003')=.T.
                        =MESSAGEBOX('此假別編號已存在,請查明後再輸入!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE6.TXTF26.SETFOCUS
                        RETURN
                    ENDIF                                             
         	    IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE)     
                        =MESSAGEBOX('假別項目不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME6.PAGE1.TXTF27.SETFOCUS
                        RETURN
                    ENDIF    
                    IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=2 AND EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE) 
                          =MESSAGEBOX('限制年度缺勤數不得為零!',0+64,'提示視窗')
      		          THISFORM.PAGEFRAME6.PAGE1.TXTF29.SETFOCUS
                          RETURN
                    ENDIF  
                    IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=2 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
         	          IF THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=1     
                               =MESSAGEBOX('扣款底數不得空白!',0+64,'提示視窗')
      		               THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.SETFOCUS
                               RETURN
                          ENDIF                                 
	                  IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE)     
	                       =MESSAGEBOX('扣款比例不得為零!',0+64,'提示視窗')
	      		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.SETFOCUS
	                       RETURN
	                  ENDIF                       
	            ELSE     
	                  IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=3 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
	                        IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE)     
	                            =MESSAGEBOX('扣款金額不得為零!',0+64,'提示視窗')
	      		            THISFORM.PAGEFRAME6.PAGE1.TXTF32.SETFOCUS
	                            RETURN
	                        ENDIF 
	                  ENDIF         	                
	             ENDIF  
		     IF THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=1 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
		         IF  THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE<>0
		              IF THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE<>0 AND THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE<=THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE    
	                           =MESSAGEBOX('全勤扣款方式一之缺勤數不可小於或等於不扣全勤緩衝數!',0+64,'提示視窗')
	      		           THISFORM.PAGEFRAME6.PAGE1.TXTF35.SETFOCUS
	                           RETURN
	                      ENDIF
		              IF THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE<>0 AND THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE<=THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE    
	                          =MESSAGEBOX('全勤扣款方式二之缺勤數不可小於或等於不扣全勤緩衝數!',0+64,'提示視窗')
	      		          THISFORM.PAGEFRAME6.PAGE1.TXTF37.SETFOCUS
	                           RETURN
	                      ENDIF
	                  ENDIF 
	         	  IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE)     
	         	      IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE)
	                           =MESSAGEBOX('全勤扣款方式一之缺勤數您已輸入資料,故扣款比例不得為零',0+64,'提示視窗')
	      		           THISFORM.PAGEFRAME1.PAGE1.TXTF36.SETFOCUS
	                           RETURN
	                      ENDIF      
	                  ENDIF
	         	  IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE)     
	         	       IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE)
	                            =MESSAGEBOX('全勤扣款方式一之扣款比例您已輸入資料,故缺勤數不得為零',0+64,'提示視窗')
	      		            THISFORM.PAGEFRAME6.PAGE1.TXTF35.SETFOCUS
	                            RETURN
	                        ENDIF      
	                  ENDIF 	                        
	         	  IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE)     
	         	      IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE)
	                           =MESSAGEBOX('全勤扣款方式二之缺勤數您已輸入資料,故扣款比例不得為零',0+64,'提示視窗')
	      		           THISFORM.PAGEFRAME6.PAGE1.TXTF38.SETFOCUS
	                           RETURN
	                      ENDIF      
	                  ENDIF
	         	  IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE)     
	         	      IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE)
	                           =MESSAGEBOX('全勤扣款方式二之扣款比例您已輸入資料,故缺勤數不得為零',0+64,'提示視窗')
	      		           THISFORM.PAGEFRAME6.PAGE1.TXTF37.SETFOCUS
	                           RETURN
	                      ENDIF      
	                  ENDIF 	                     	        
		     ENDIF
                     SELECT S00
		     APPEND BLANK  
		     LOCK()
		     IF RLOCK()   
		         REPLACE F01 WITH 'G'
		         REPLACE F26 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE    
		         REPLACE F27 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE
		         REPLACE F28 WITH IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=1,'D',IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=2,'H','S'))		            
		         REPLACE F29 WITH IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=2,THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE,0)
		         IF THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=2 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
		             DO CASE 
		                   CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=2
		                              REPLACE F30 WITH 'H'
		                   CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=3
		                              REPLACE F30 WITH 'D'
		                   CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=4
		                              REPLACE F30 WITH 'B'
		             ENDCASE 	
		             REPLACE F31 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE
		        ENDIF	            
		        IF THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=3 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
		             REPLACE F32 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE		            
		        ENDIF      
		        IF THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=1 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
		             REPLACE F33 WITH 'Y'
			     REPLACE F34 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE
			     REPLACE F35 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE
			     REPLACE F36 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE
			     REPLACE F37 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE
			     REPLACE F38 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE
			 ELSE
		             REPLACE F33 WITH 'N'			         
			 ENDIF    
			 IF THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=1
			      REPLACE F30 WITH ''
			      REPLACE F31 WITH 0
			      REPLACE F32 WITH 0
			      REPLACE F33 WITH 'N'	
			 ENDIF
			 REPLACE F58 WITH ALLTRIM(STR(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE))
			 REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE))
			 REPLACE F62 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE 
		     REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
		    ENDIF	
		    THISFORM.ORPGROUP.NEW_BOTT.CLICK 
         CASE S00_1.F01='H'   &&懲處扣款設定-H
         	   IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE)     
                       =MESSAGEBOX('懲處編號不得空白!',0+64,'提示視窗')
      		       THISFORM.PAGEFRAME7.PAGE1.TXTF39.SETFOCUS
                       RETURN
                   ENDIF   
         	   IF LEN(ALLTRIM(THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE))<2
                       =MESSAGEBOX('懲處編號不得少於兩碼!',0+64,'提示視窗')
      		      THISFORM.PAGEFRAME7.PAGE1.TXTF39.SETFOCUS
                      RETURN
                   ENDIF                        
         	   IF INDEXSEEK('H'+THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE,.F.,'S00','S004')=.T.
                       =MESSAGEBOX('此懲處編號已存在,請查明後再輸入!',0+64,'提示視窗')
      		       THISFORM.PAGEFRAME7.PAGE1.TXTF39.SETFOCUS
                       RETURN
                   ENDIF                                             
         	   IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE)     
                       =MESSAGEBOX('懲處項目不得空白!',0+64,'提示視窗')
      		       THISFORM.PAGEFRAME7.PAGE1.TXTF40.SETFOCUS
                       RETURN
                   ENDIF    
                   IF  THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=1
                         IF THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=1     
                              =MESSAGEBOX('扣款底數不得空白!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.SETFOCUS
                             RETURN
                         ENDIF     
	                 IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE)     
	                     =MESSAGEBOX('扣款比例不得為零!',0+64,'提示視窗')
	      		     THISFORM.PAGEFRAME7.PAGE1.TXTF43.SETFOCUS
	                     RETURN
	                 ENDIF                       
	           ELSE     
	                 IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE)     
	                     =MESSAGEBOX('扣款金額不得為零!',0+64,'提示視窗')
	      	             THISFORM.PAGEFRAME7.PAGE1.TXTF44.SETFOCUS
	                      RETURN
	                 ENDIF     	                
	           ENDIF  
                   SELECT S00
		   APPEND BLANK  
		   LOCK()
		   IF RLOCK()   
		       REPLACE F01 WITH 'H'
		       REPLACE F39 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF39.VALUE
		       REPLACE F40 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE
		       REPLACE F41 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF41.VALUE		                 
		       IF THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=1
		            DO CASE 
		                  CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=2
		                             REPLACE F42 WITH 'H'
		                  CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=3
		                             REPLACE F42 WITH 'D'
		                  CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=4
		                             REPLACE F42 WITH 'B'
		                  CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=5
		                             REPLACE F42 WITH 'T'		                              
		            ENDCASE  
		            REPLACE F43 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE
		       ELSE
		            REPLACE F44 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE		            
		       ENDIF 	 
   		       REPLACE F45 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE			    
		       REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF
		    THISFORM.ORPGROUP.NEW_BOTT.CLICK 
         CASE S00_1.F01='I'   &&班別設定-I
         	    IF EMPTY(THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE)     
                        =MESSAGEBOX('班別編號不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME8.PAGE1.TXTF48.SETFOCUS
                        RETURN
                    ENDIF  
         	    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE))<>2
                        =MESSAGEBOX('班別編號必須為兩碼!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME8.PAGE1.TXTF48.SETFOCUS
                        RETURN
                    ENDIF                         
	            IF INDEXSEEK('I'+THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE,.F.,'S00','S005')=.T.     
                        =MESSAGEBOX('此班別編號已存在,請重新輸入!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME8.PAGE1.TXTF48.SETFOCUS
                        RETURN
                    ENDIF       
         	    IF EMPTY(THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE)     
                        =MESSAGEBOX('班別說明不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME8.PAGE1.TXTF49.SETFOCUS
                        RETURN
                    ENDIF   
                    ** 
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE,2)))
		       =MESSAGEBOX('上班時段不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF50.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))>59
		        =MESSAGEBOX('上班時段格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF50.SETFOCUS
		        RETURN  
		    ENDIF        
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE,2)))
		       =MESSAGEBOX('下班時段不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF51.SETFOCUS
		        RETURN       
		    ENDIF			     
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))>59
		       =MESSAGEBOX('下班時段格式不正確!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF51.SETFOCUS
		       RETURN
		    ENDIF		    
		    **
		    IF  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))<>1
		         IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))<>5 OR ;
		             LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2)))=1 OR;
		             LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2)))=1 OR;
		             VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))>23 OR;
		             VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))>59
		             =MESSAGEBOX('起始休息時段格式不正確!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF52.SETFOCUS
		              RETURN
		        ENDIF      
		        IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE,2)))
             		     =MESSAGEBOX('起始休息時段已有資料故截止休息時段不得空白!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF53.SETFOCUS
		             RETURN       
		        ENDIF			        
		    ENDIF	
		    **
		    IF  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))<>1	         
		        IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE,2)))
             		     =MESSAGEBOX('截止休息時段已有資料故起始休息時段不得空白!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF52.SETFOCUS
		             RETURN       
		        ENDIF				 
			 IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))<>5 OR ;
			     LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2)))=1 OR;
			     LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2)))=1 OR;
			     VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))>23 OR;
			     VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))>59
			     =MESSAGEBOX('截止休息時段格式不正確!',0+64,'提示視窗')
			    THISFORM.PAGEFRAME8.PAGE1.TXTF53.SETFOCUS
			     RETURN
			 ENDIF    
		    ENDIF	  
		    IF EMPTY(THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE)
		       =MESSAGEBOX('上班總時數不得為零!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF57.SETFOCUS
		        RETURN       
		    ENDIF	
		    **	
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE)) = 5
		        IF LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2)))=1 OR;
		            LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2)))=1 OR;
		           VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2))>23 OR;
		           VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2))>59
		           =MESSAGEBOX('上班刷卡時段止格式不正確!',0+64,'提示視窗')
		           THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		           RETUR	
		        ENDIF   
		    ENDIF    	    	 
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 
		        IF LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2)))=1 OR;
		            LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2)))=1 OR;
		            VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2))>23 OR;
		            VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2))>59
		            =MESSAGEBOX('下班刷卡時段起格式不正確!',0+64,'提示視窗')
		            THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		            RETUR	
		        ENDIF    
		    ENDIF  		    	       		    
		    ***
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE,2)))
		       =MESSAGEBOX('上班遲到緩衝時間可以與上班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))>59
		        =MESSAGEBOX('上班遲到緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  
		    ENDIF   
		    ***
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))<VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
		        =MESSAGEBOX('上班遲到緩衝時間不得小於上班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))<VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
		        =MESSAGEBOX('上班遲到緩衝時間不得小於上班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  		         
		    ENDIF		    
		    **
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE,2)))
		       =MESSAGEBOX('下班早退緩衝時間可以與下班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>59
		        =MESSAGEBOX('下班早退緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  
		    ENDIF  		    
		    ***
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('下班早退緩衝時間不得大於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('下班早退緩衝時間不得大於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  		         
		    ENDIF				       
		    **
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE,2)))
		       =MESSAGEBOX('加班上班緩衝時間可以與下班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))>59
		        =MESSAGEBOX('加班上班緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  
		    ENDIF  	
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))<VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('加班上班緩衝時間不得小於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))<VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('加班上班緩衝時間不得小於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  		         
		    ENDIF			
		    ***
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE < THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE
		       =MESSAGEBOX('上班刷卡時段止需大於上班時段!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		        RETURN       
		    ENDIF			    
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE < THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE
		       =MESSAGEBOX('上班刷卡時段止不可小於上班遲到緩衝時間!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		        RETURN       
		    ENDIF		
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE > THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE
		       =MESSAGEBOX('下班刷卡時段起需小於下班時段!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		        RETURN       
		    ENDIF			    
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE > THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE
		       =MESSAGEBOX('下班刷卡時段起不可大於下班早退緩衝時間!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		        RETURN       
		    ENDIF
		    ***                  
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
		        REPLACE F01 WITH 'I'
		        REPLACE F48 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF48.VALUE
		        REPLACE F49 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE
		        REPLACE F50 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE
		        REPLACE F51 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE
		        REPLACE F52 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE
		        REPLACE F53 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE
		        REPLACE F54 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE
		        REPLACE F55 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE
		        REPLACE F56 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE
		        REPLACE F57 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
		        REPLACE F66 WITH IIF(LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5,THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE,'')
		        REPLACE F67 WITH IIF(LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5,THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE,'') 
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF 
		    THISFORM.ORPGROUP.NEW_BOTT.CLICK 	
         CASE S00_1.F01='J'   &&其他所得發放設定-J	 	
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
		        REPLACE F01 WITH 'J'
		        REPLACE F59 WITH THISFORM.PAGEFRAME9.PAGE1.TXTF59.VALUE         
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF    
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK        
         CASE S00_1.F01='K'   &&年終獎金發放設定-K	 	
                    DO CASE 
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE)
		      			   =MESSAGEBOX('方式一之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF691.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
		      			   =MESSAGEBOX('方式一之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.SETFOCUS
		        		   RETURN       
		    		      ENDIF		  
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE)
		      			   =MESSAGEBOX('方式二之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF692.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 1
		      			   =MESSAGEBOX('方式二之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.SETFOCUS
		        		   RETURN       
		    		      ENDIF		    	
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE)
		      			   =MESSAGEBOX('方式三之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF693.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE)
		      			   =MESSAGEBOX('方式三之發放金額不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF71.SETFOCUS
		        		   RETURN       
		    		      ENDIF	        		      
                    ENDCASE 
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
		        REPLACE F01 WITH 'K'
		        REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE))
		        DO CASE 
		              CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
		              		 REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0
		              CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2         
		              		 REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0	
		               CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3  
		               		  REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE
		               		  REPLACE F70 WITH ''    
		               		  REPLACE F71 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE    	               
		        ENDCASE    
		        REPLACE F72 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF72.VALUE
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF    
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   		
         CASE S00_1.F01='L'   &&未休年假獎金發放設定-L	 	
                    DO CASE 
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1	
                                      IF THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
		      			   =MESSAGEBOX('方式一之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF EMPTY(THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE)
		      			   =MESSAGEBOX('方式一之發放比率不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.TXTF69.SETFOCUS
		        		   RETURN       
		    		      ENDIF			    		        
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2		
                                      IF THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 1
		      			   =MESSAGEBOX('方式二之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.SETFOCUS
		        		   RETURN       
		    		      ENDIF		    	
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3	
                                      IF EMPTY(THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE)
		      			   =MESSAGEBOX('方式三之發放金額不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.TXTF71.SETFOCUS
		        		   RETURN       
		    		      ENDIF	        		      
                    ENDCASE                    
                    SELECT S00
		    APPEND BLANK  
		    LOCK()
		    IF RLOCK()   
		        REPLACE F01 WITH 'L'
		        REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE))
		        DO CASE 
		              CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
		              		  REPLACE F69 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0
		              CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2         
		              		 REPLACE F69 WITH 0
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0	
		               CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3  
		               		  REPLACE F69 WITH 0
		               		  REPLACE F70 WITH ''    
		               		  REPLACE F71 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE    	               
		        ENDCASE    
		        REPLACE F72 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF72.VALUE
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		    ENDIF    
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   		        
         CASE S00_1.F01='Y'   &&加班參數設定-Y	 
		    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=2   
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)>60 
		             =MESSAGEBOX('加班時間不滿之分鐘數不可超過60分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN 
		         ENDIF     
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)=0 
		             =MESSAGEBOX('加班時間不滿之分鐘數不可為零!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN 
		         ENDIF		         
		    ENDIF 
		    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=2  
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)>60    
		             =MESSAGEBOX('加班時間若滿之分鐘數不可超過60分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF64.SETFOCUS
		             RETURN 
		         ENDIF  
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)=0    
		             =MESSAGEBOX('加班時間若滿之分鐘數不可為零!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF64.SETFOCUS
		             RETURN 
		         ENDIF    		           
		    ENDIF                     
                    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=2 AND THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=2
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)>VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)
		             =MESSAGEBOX('加班時間不滿之分鐘數不可大於加班時間若滿之分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN  
		         ENDIF   	
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)=VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)
		             =MESSAGEBOX('加班時間不滿之分鐘數不可等於加班時間若滿之分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN  
		         ENDIF  		          		         	         
		   ENDIF		      		                   
                   SELECT S00
                   APPEND BLANK 
		   IF RLOCK()  
		       REPLACE F01 WITH 'Y'
		       REPLACE F63 WITH IIF(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=1,'F',ALLTRIM(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE))
		       REPLACE F64 WITH IIF(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=1,'F',ALLTRIM(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE))
		       REPLACE F65 WITH ALLTRIM(STR(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.VALUE))
		       REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		   ENDIF 
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           	       
	 CASE S00_1.F01='Z'   &&基本參數設定-Z	     
         	    IF EMPTY(THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE)     
                         =MESSAGEBOX('天數不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME14.PAGE1.SPINNER02.SETFOCUS
                         RETURN
                    ENDIF 
         	    IF EMPTY(THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE)     
                        =MESSAGEBOX('時數不得為零!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME14.PAGE1.SPINNER03.SETFOCUS
                         RETURN
                    ENDIF   
         	    IF THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE <= 0     
                   =MESSAGEBOX('報表欄位數必須大於零!',0+64,'提示視窗')
      		       THISFORM.PAGEFRAME14.PAGE1.TXTF73.SETFOCUS
                   RETURN
                ENDIF                                         
                   SELECT S00
                   APPEND BLANK 
		   IF RLOCK()  
		       REPLACE F01 WITH 'Z'
		       REPLACE F02 WITH THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE
		       REPLACE F03 WITH THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE
		       REPLACE F04 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF04.VALUE
		       REPLACE F60 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF60.VALUE
		       REPLACE F61 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF61.VALUE
		       REPLACE F73 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE
		       REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
		   ENDIF 
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK  	                                                           		       		       		       		       		       		                                                                 		      
  ENDCASE 
  UNLOCK
  THISFORM.REFRESH       
ENDPROC
**********修改
PROCEDURE ED_BOTT.CLICK
   SELECT S00
   R=RECNO('&AREA1')
   THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","IIF(RECNO('&AREA1')=R,RGB(255,255,0),'')","COLUMN")
   THISFORM.GRID1.ENABLED=.F.   
   THISFORM.GRID2.ENABLED=.F. 
   THISFORM.CMDGROUP.TOP_BOTT.ENABLED=.F.      
   THISFORM.CMDGROUP.UP_BOTT.ENABLED=.F.
   THISFORM.CMDGROUP.NEXT_BOTT.ENABLED=.F.
   THISFORM.CMDGROUP.END_BOTT.ENABLED=.F.      
   THISFORM.ED_BOTT.ENABLED=.F.
   THISFORM.ORPGROUP.NEW_BOTT.ENABLED=.F.
   THISFORM.ORPGROUP.EDIT_BOTT.ENABLED=.F.
   THISFORM.ORPGROUP.DEL_BOTT.ENABLED=.F.
   THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.
   THISFORM.ORPGROUP.ESC_BOTT.ENABLED=.F.
   THISFORM.SHRGROUP.ENABLED=.T.
   THISFORM.SHRGROUP.VISIBLE=.T.
   THISFORM.SHRGROUP.SHURE1_BOTT.ENABLED=.F.
   THISFORM.SHRGROUP.SHURE1_BOTT.VISIBLE=.F.
   THISFORM.SHRGROUP.SHURE2_BOTT.ENABLED=.T.
   THISFORM.SHRGROUP.SHURE2_BOTT.VISIBLE=.T.  
   THISFORM.SHRGROUP.ABANDON_BOTT.VISIBLE=.T.
   THISFORM.SHRGROUP.FINISH_BOTT.ENABLED=.F.
   THISFORM.SHRGROUP.FINISH_BOTT.VISIBLE=.F.        
   DO CASE	       
         CASE S00.F01='B'    &&加班費設定-B       
		   THISFORM.PAGEFRAME1.PAGE1.TXTF05.READONLY=.T.         	
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.ENABLED=.T.
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.ENABLED=.T.		       
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.ENABLED=.T.
		   THISFORM.PAGEFRAME1.PAGE1.TXTF08.VISIBLE=IIF(!EMPTY(S00.F08),.F.,.T.)
		   THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VISIBLE=IIF(!EMPTY(S00.F08),.T.,.F.)		       
		   THISFORM.PAGEFRAME1.PAGE1.TXTF46.VISIBLE=.F.	
		   THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VISIBLE=.T.	
		   THISFORM.PAGEFRAME1.PAGE1.TXTF08.READONLY=IIF(!EMPTY(S00.F08),.F.,.T.)
		   THISFORM.PAGEFRAME1.PAGE1.TXTF09.READONLY=IIF(!EMPTY(S00.F08),.F.,.T.)
		   THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=IIF(!EMPTY(S00.F08),.T.,.F.)
		   THISFORM.PAGEFRAME1.PAGE1.TXTF47.READONLY=.F.		       	       	       	       		       	
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE=VAL(S00.F06)
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=IIF(S00.F07='H',1,2)       	
		   THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=IIF(!EMPTY(S00.F08),1,2) 
		   THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=IIF(EMPTY(S00.F08),1,IIF(S00.F08='H',2,IIF(S00.F08='D',3,4)))
		   THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.DISPLAYVALUE=S00.F46		         	       
		   THISFORM.PAGEFRAME1.PAGE1.LBLF991.CAPTION=''		
		   IF !EMPTY(S00.F08)
		       THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.SETFOCUS
		   ELSE
		       THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
		   ENDIF
         CASE S00.F01='C'    &&遲到設定    	       
         	   THISFORM.PAGEFRAME2.PAGE1.SETALL('READONLY',.F.,'SPINNER')
		   THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.ENABLED=.T.			       
		   THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.ENABLED=.T.		       	
		   THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.ENABLED=IIF(S00.F13<>'Y',.F.,.T.)		       	       	       		       
		   THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=IIF(S00.F12='Y',1,2)		       
		   THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=IIF(S00.F13<>'Y',1,2)		       
		   THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.READONLY=IIF(S00.F13<>'Y',.T.,.F.)		       
		   THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=IIF(EMPTY(S00.F15),2,1)		       	       
		   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VISIBLE=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)   		       
		   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.READONLY=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)   		       
		   THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.F.,.T.),.F.)   		      	
		   THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=IIF(EMPTY(S00.F15),1,IIF(S00.F15='H',2,IIF(S00.F15='D',3,4)))	       	
		   THISFORM.PAGEFRAME2.PAGE1.TXTF16C.READONLY=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)   
		   THISFORM.PAGEFRAME2.PAGE1.TXTF17C.READONLY=IIF(S00.F13='Y',IIF(EMPTY(S00.F15),.F.,.T.),.T.)     	       	       	       	       	       
		   THISFORM.PAGEFRAME2.PAGE1.LBLF991.CAPTION=''		
		   THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.SETFOCUS		       
         CASE S00.F01='D'    &&早退設定
         	   THISFORM.PAGEFRAME3.PAGE1.SETALL('READONLY',.F.,'SPINNER')
         	   THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.ENABLED=.T.
         	   THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.ENABLED=.T.
         	   THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.ENABLED=IIF(S00.F13<>'Y',.F.,.T.)	
         	   THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=IIF(S00.F12='Y',1,2)
         	   THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=IIF(S00.F13<>'Y',1,2)
         	   THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.READONLY=IIF(S00.F13<>'Y',.T.,.F.)
         	   THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=IIF(EMPTY(S00.F15),2,1)
         	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VISIBLE=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)
         	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.READONLY=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)
         	   THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.F.,.T.),.F.)
         	   THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=IIF(EMPTY(S00.F15),1,IIF(S00.F15='H',2,IIF(S00.F15='D',3,4)))
         	   THISFORM.PAGEFRAME3.PAGE1.TXTF16D.READONLY=IIF(S00.F13<>'N',IIF(EMPTY(S00.F15),.T.,.F.),.T.)
         	   THISFORM.PAGEFRAME3.PAGE1.TXTF17D.READONLY=IIF(S00.F13='Y',IIF(EMPTY(S00.F15),.F.,.T.),.T.)
		   THISFORM.PAGEFRAME3.PAGE1.LBLF991.CAPTION=''		
		   THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.SETFOCUS	         	       
         CASE S00.F01='E'    &&年假範圍設定-E        	       
         	   THISFORM.PAGEFRAME4.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')    
         	   THISFORM.PAGEFRAME4.PAGE1.TXTF20.READONLY=.F.           	       	       	       	       
		   THISFORM.PAGEFRAME4.PAGE1.LBLF991.CAPTION=''	
		   THISFORM.PAGEFRAME4.PAGE1.TXTF20.SETFOCUS  
         CASE S00.F01='F'    &&年假限制設定-F
		   THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.ENABLED=.T.	
		   THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.ENABLED=.T.		
		   THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.ENABLED=.T.	         
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF21.READONLY=IIF(EMPTY(S00.F21),.T.,.F.)
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF22.READONLY=IIF(EMPTY(S00.F21),.T.,.F.)
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF23.READONLY=IIF(EMPTY(S00.F23),.T.,.F.)
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF24.READONLY=IIF(EMPTY(S00.F25),.F.,.T.)
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF25.READONLY=IIF(EMPTY(S00.F25),.T.,.F.)
     		   THISFORM.PAGEFRAME5.PAGE1.LBLF991.CAPTION=''
     		   THISFORM.PAGEFRAME5.PAGE1.TXTF21.SETFOCUS         
         CASE S00.F01='G'   &&請假扣款項目-G        	       
         	   THISFORM.PAGEFRAME6.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)
		   THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),.F.,IIF(EMPTY(S00.F31),.F.,.T.)) AND IIF(S00.F58<>'1',.T.,.F.)		        	   		   
		   THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=IIF(S00.F58<>'1', IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),.T.,IIF(EMPTY(S00.F31),.T.,.F.)),.F.)		       
		   THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)    
		   THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=IIF(S00.F58<>'1',.T.,.F.) 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=IIF(S00.F58<>'1',.T.,.F.)               	   
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.ENABLED=.T.	
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.ENABLED=.T.
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.ENABLED=IIF(S00.F58<>'1',.T.,.F.)	
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.ENABLED=.T.
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.ENABLED=.T.
		   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.ENABLED=IIF(S00.F58<>'1',.T.,.F.)	 
		   THISFORM.PAGEFRAME6.PAGE1.TXTF27.READONLY=IIF(S00.F26<>'A1',.F.,.T.)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=IIF(S00.F58<>'1', IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),.T.,IIF(EMPTY(S00.F31),.T.,.F.)),.F.)		       
		   THISFORM.PAGEFRAME6.PAGE1.TXTF31.READONLY=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),.T.,IIF(EMPTY(S00.F31),.T.,.F.))
		   THISFORM.PAGEFRAME6.PAGE1.TXTF32.READONLY=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),.T.,IIF(EMPTY(S00.F32),.T.,.F.))		       
		   THISFORM.PAGEFRAME6.PAGE1.TXTF34.READONLY=IIF(S00.F33='Y',.F.,.T.)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF35.READONLY=IIF(S00.F33='Y',.F.,.T.)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF36.READONLY=IIF(S00.F33='Y',.F.,.T.)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF37.READONLY=IIF(S00.F33='Y',.F.,.T.)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF38.READONLY=IIF(S00.F33='Y',.F.,.T.)	
		   THISFORM.PAGEFRAME6.PAGE1.TXTF62.READONLY=IIF(S00.F28='S',.T.,.F.)    		       		       		       
		   THISFORM.PAGEFRAME6.PAGE1.TXTF26.VALUE=S00.F26
		   THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE=S00.F27
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=IIF(S00.F28='D',1,IIF(S00.F28='H',2,3))     
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=IIF(EMPTY(S00.F29),1,2)
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),1,IIF(EMPTY(S00.F31),3,2))
		   THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=IIF(EMPTY(S00.F30),1,IIF(S00.F30='H',2,IIF(S00.F30='D',3,4)))
		   THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),'',IIF(EMPTY(S00.F31),'',S00.F31))
		   THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=IIF(EMPTY(S00.F31) AND EMPTY(S00.F32),'',IIF(EMPTY(S00.F32),'',S00.F32))
		   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=IIF(S00.F33='Y',1,2)
		   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=IIF(S00.F33='Y',S00.F34,'')
		   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=IIF(S00.F33='Y',S00.F35,'')
		   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=IIF(S00.F33='Y',S00.F36,'')
		   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=IIF(S00.F33='Y',S00.F37,'')
		   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=IIF(S00.F33='Y',S00.F38,'')
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=VAL(S00.F58)
		   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE=VAL(S00.F68)
		   THISFORM.PAGEFRAME6.PAGE1.LBLF991.CAPTION=''
		   THISFORM.PAGEFRAME6.PAGE1.TXTF27.SETFOCUS  
         CASE S00.F01='H'    &&懲處扣款設定-H           	       
         	   THISFORM.PAGEFRAME7.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
		   THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.ENABLED=.T.	
		   THISFORM.PAGEFRAME7.PAGE1.TXTF39.READONLY=.T.
		   THISFORM.PAGEFRAME7.PAGE1.TXTF42.READONLY=IIF(EMPTY(S00.F42),.T.,.F.)
		   THISFORM.PAGEFRAME7.PAGE1.TXTF43.READONLY=IIF(EMPTY(S00.F43),.T.,.F.)
		   THISFORM.PAGEFRAME7.PAGE1.TXTF44.READONLY=IIF(EMPTY(S00.F44),.T.,.F.)
		   THISFORM.PAGEFRAME7.PAGE1.TXTF42.VISIBLE=IIF(EMPTY(S00.F42),.T.,.F.)
		   THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VISIBLE=IIF(EMPTY(S00.F42),.F.,.T.)
		   THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=IIF(EMPTY(S00.F42),2,1)   
		   THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=IIF(EMPTY(S00.F42),1,IIF(S00.F42='H',2,IIF(S00.F42='D',3,IIF(S00.F42='B',4,5))))
		   THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=IIF(EMPTY(S00.F43),'',S00.F43)
		   THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=IIF(EMPTY(S00.F44),'',S00.F44)
		   THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE=S00.F45
		   THISFORM.PAGEFRAME7.PAGE1.LBLF991.CAPTION=''
		   THISFORM.PAGEFRAME7.PAGE1.TXTF40.SETFOCUS	
         CASE S00.F01='I'  &&班別設定-I        	       
         	   THISFORM.PAGEFRAME8.PAGE1.SETALL('READONLY',.F.,'TEXTBOX') 
         	   THISFORM.PAGEFRAME8.PAGE1.SETALL('READONLY',.F.,'SPINNER')     	          	       
 		   THISFORM.PAGEFRAME8.PAGE1.TXTF48.READONLY=.T.       	       	       	       	       
		   THISFORM.PAGEFRAME8.PAGE1.TXTF50.READONLY=.F.
		   THISFORM.PAGEFRAME8.PAGE1.TXTF51.READONLY=.F.		
		   THISFORM.PAGEFRAME8.PAGE1.TXTF52.READONLY=.F.
		   THISFORM.PAGEFRAME8.PAGE1.TXTF53.READONLY=.F.
		   THISFORM.PAGEFRAME8.PAGE1.TXTF57.READONLY=.T.   
		   THISFORM.PAGEFRAME8.PAGE1.LBLF991.CAPTION=''	
 		   THISFORM.PAGEFRAME8.PAGE1.TXTF49.SETFOCUS   	
         CASE S00.F01='J'   &&其他所得發放設定-J
         	   THISFORM.PAGEFRAME9.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
		   THISFORM.PAGEFRAME9.PAGE1.LBLF991.CAPTION=''
		   THISFORM.PAGEFRAME9.PAGE1.TXTF59.SETFOCUS	 		  
         CASE  S00_1.F01='K'   &&年終獎金發放設定-K 	      
      		    THISFORM.PAGEFRAME10.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
      		    THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.ENABLED=.T.
      		    DO CASE 
      		          CASE S00.F68 = '1'
      		                     THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF701.VISIBLE = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VISIBLE = .T.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF691.READONLY = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF701.READONLY = .F.
      		                     DO CASE 
      		                            CASE S00.F70 = 'H'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 2
      		                            CASE S00.F70 = 'D'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 3
      		                            CASE S00.F70 = 'B'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 4 
      		                     ENDCASE
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF691.SETFOCUS	
      		          CASE S00.F68 = '2'
      		                     THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF702.VISIBLE = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VISIBLE = .T.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF692.READONLY = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF702.READONLY = .F.
      		                     DO CASE 
      		                            CASE S00.F70 = 'H'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 2
      		                            CASE S00.F70 = 'D'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 3
      		                            CASE S00.F70 = 'B'
      		                                       THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 4 
      		                     ENDCASE
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF692.SETFOCUS	
      		          CASE S00.F68 = '3'
      		                     THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3     
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF693.READONLY = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF71.READONLY = .F.
      		                     THISFORM.PAGEFRAME10.PAGE1.TXTF693.SETFOCUS
      		    ENDCASE 
		    THISFORM.PAGEFRAME10.PAGE1.TXTF72.READONLY = .F.
		    THISFORM.PAGEFRAME10.PAGE1.LBLF991.CAPTION=''   	          		    
         CASE  S00_1.F01='L'   &&未休年假獎金發放設定-L 	      
      		    THISFORM.PAGEFRAME11.PAGE1.SETALL('READONLY',.T.,'TEXTBOX')
      		    THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.ENABLED=.T.
      		    DO CASE 
      		          CASE S00.F68 = '1'
      		                     THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF701.VISIBLE = .F.
      		                     THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VISIBLE = .T.
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF69.READONLY = .F.
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF701.READONLY = .F.
      		                     DO CASE 
      		                            CASE S00.F70 = 'H'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 2
      		                            CASE S00.F70 = 'D'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 3
      		                            CASE S00.F70 = 'B'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 4 
      		                     ENDCASE
      		                     THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.SETFOCUS	
      		          CASE S00.F68 = '2'
      		                     THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF702.VISIBLE = .F.
      		                     THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VISIBLE = .T.
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF702.READONLY = .F.
      		                     DO CASE 
      		                            CASE S00.F70 = 'H'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 2
      		                            CASE S00.F70 = 'D'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 3
      		                            CASE S00.F70 = 'B'
      		                                       THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 4 
      		                     ENDCASE
      		                     THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.SETFOCUS	
      		          CASE S00.F68 = '3'
      		                     THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3     
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF71.READONLY = .F.
      		                     THISFORM.PAGEFRAME11.PAGE1.TXTF71.SETFOCUS
      		    ENDCASE       		    
		    THISFORM.PAGEFRAME11.PAGE1.TXTF72.READONLY = .F.
		    THISFORM.PAGEFRAME11.PAGE1.LBLF991.CAPTION=''   	    	       
         CASE S00_1.F01='X'  &&公司資料設定        	       
         	   THISFORM.PAGEFRAME12.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
		   THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_000.SETFOCUS	
	 CASE S00.F01='Y'  &&加班參數設定-Y
	    	    THISFORM.PAGEFRAME13.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
	    	    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.ENABLED=.T.
	    	    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.ENABLED=.T.
	    	    THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.ENABLED=.T.
	    	     THISFORM.PAGEFRAME13.PAGE1.SETFOCUS
         CASE S00.F01='Z'   &&基本參數設定-Z          	       
         	   THISFORM.PAGEFRAME14.PAGE1.SETALL('READONLY',.F.,'TEXTBOX')
         	   THISFORM.PAGEFRAME14.PAGE1.SETALL('READONLY',.F.,'SPINNER')
		   THISFORM.PAGEFRAME14.PAGE1.LBLF991.CAPTION=''
		   THISFORM.PAGEFRAME14.PAGE1.SPINNER02.SETFOCUS			       	       		       
    ENDCASE   
    IF S00_1.F01<>'X'  
	LOCK()
	IF !RLOCK()
	    DO FILE_IMPACT
	    RETURN
	ENDIF  
    ENDIF
ENDPROC
*******修改確認
PROCEDURE SHRGROUP.SHURE2_BOTT.CLICK
  DO CASE           	  
         CASE S00.F01='B'        &&加班費設定-B                     
                    IF THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=1           
         	         IF THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=1     
                             =MESSAGEBOX('計算底數不得空白!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.SETFOCUS
                             RETURN
                         ENDIF    
         	         IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE)
                             =MESSAGEBOX('計算比例不得為零!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME1.PAGE1.TXTF09.SETFOCUS
                             RETURN
                         ENDIF     
                    ELSE
         	         IF EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE)
                             =MESSAGEBOX('計算金額不得為零!',0+64,'提示視窗')
      		             THISFORM.PAGEFRAME1.PAGE1.TXTF10.SETFOCUS
                              RETURN
                          ENDIF                        
                    ENDIF 
         	    IF  THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE<>1 AND EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE)     
                         =MESSAGEBOX('您已選擇津貼補助項目,故補助金額不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME1.PAGE1.TXTF47.SETFOCUS
                         RETURN
                    ENDIF
         	    IF THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE=1 AND !EMPTY(THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE)     
                        =MESSAGEBOX('您選擇津貼補助項目為空白,故不得有補助金額!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE1.TXTF47.SETFOCUS
                        RETURN
                    ENDIF   
                    SELECT S00
                    IF RLOCK()                                                          
		        REPLACE F05 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF05.VALUE
		        REPLACE F06 WITH ALLTRIM(STR(THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.VALUE))
		        REPLACE F07 WITH IIF(THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.VALUE=1,'H','D')
		        IF THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.VALUE=1
	            	    DO CASE 
		                  CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=2
		                             REPLACE F08 WITH 'H'
		                  CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=3
		                             REPLACE F08 WITH 'D'
		                  CASE THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=4
		                             REPLACE F08 WITH 'B'	                              
		             ENDCASE
			     REPLACE F09 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE
		             REPLACE F10 WITH 0	
		        ELSE
		             REPLACE F08 WITH ''	
		             REPLACE F09 WITH 0	
		             REPLACE F10 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE
		        ENDIF
		        REPLACE F46 WITH ALLTRIM(THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE)
		        IF THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VALUE<>1
		             REPLACE F47 WITH THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE
		        ELSE
		             REPLACE F47 WITH 0	     
		        ENDIF
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	           ELSE
	                DO FILE_IMPACT
	                RETURN 		        
		   ENDIF     
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK       
         CASE S00.F01='C'   &&遲到設定-C
                   IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=2
                         IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE)
	                       =MESSAGEBOX('您已選擇設定遲到額外扣款,故每月遲到"超過"分鐘數不得為零!',0+64,'提示視窗')
	                      THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.SETFOCUS
	                      RETURN
	                 ENDIF                        
 	       	         IF THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE<=THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE
                              =MESSAGEBOX('您已選擇設定額外扣款,故每月遲到"超過"分鐘數不得小於或等於每月遲到分鐘數!',0+64,'提示視窗')
  		             THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.SETFOCUS
                  	     RETURN
                         ENDIF  
                         IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1                         
	                       IF  THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=1
	                            =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款底數不得為空白!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.SETFOCUS
	                            RETURN
	                       ENDIF   
	                       IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE)
	                            =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款比列不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME2.PAGE1.TXTF16C.SETFOCUS
	                            RETURN
	                       ENDIF   	                          
	                 ELSE      
	                       IF  EMPTY(THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE)
	                             =MESSAGEBOX('您已選擇設定遲到額外扣款,故另扣項目之扣款金額不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME2.PAGE1.TXTF17C.SETFOCUS
	                            RETURN
	                       ENDIF  
	                 ENDIF 
                    ENDIF         
		   SELECT S00
		   IF RLOCK()                              
		       REPLACE F11 WITH THISFORM.PAGEFRAME2.PAGE1.SPINNER11C.VALUE
		       REPLACE F12 WITH  IIF(THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.VALUE=1,'Y','N')
		       IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.VALUE=1
		            REPLACE F13 WITH 'N'	
		            REPLACE F14 WITH 0
		            REPLACE F15 WITH ''
		            REPLACE F16 WITH 0
		            REPLACE F17 WITH 0
		       ELSE
		            REPLACE F13 WITH 'Y'
		            REPLACE F14 WITH THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE
		            IF  THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1 
		               	 DO CASE 
			               CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=2
			                          REPLACE F15 WITH 'H'
			               CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=3
			                          REPLACE F15 WITH 'D'
			               CASE THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=4
			                          REPLACE F15 WITH 'B'	                              
			         ENDCASE
		            	 REPLACE F16 WITH THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE	
		            	 REPLACE F17 WITH 0 
		            ELSE
		             	 REPLACE F15 WITH ''
		            	 REPLACE F16 WITH 0
		            	 REPLACE F17 WITH THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE
		             ENDIF                 	  
		       ENDIF
		       REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	           ELSE
	                DO FILE_IMPACT
	                RETURN 		       
		   ENDIF    
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK                           
         CASE S00.F01='D'   &&早退設定-D   
                   IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=2
                         IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE)
	                      =MESSAGEBOX('您已選擇設定早退額外扣款,故每月早退"超過"分鐘數不得為零!',0+64,'提示視窗')
	                      THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.SETFOCUS
	                      RETURN
	                 ENDIF  
 	       	         IF THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE<=THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE
                              =MESSAGEBOX('您已選擇設定額外扣款,故每月早退"超過"分鐘數不得小於或等於每月早退分鐘數!',0+64,'提示視窗')
  		             THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.SETFOCUS
                  	     RETURN
                         ENDIF  
                         IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1                         
	                       IF  THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=1
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款底數不得為空白!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.SETFOCUS
	                            RETURN
	                       ENDIF   
	                       IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE)
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款比列不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.TXTF16D.SETFOCUS
	                            RETURN
	                       ENDIF   	                          
	                 ELSE      
	                       IF  EMPTY(THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE)
	                            =MESSAGEBOX('您已選擇設定早退額外扣款,故另扣項目之扣款金額不得為零!',0+64,'提示視窗')
	                            THISFORM.PAGEFRAME3.PAGE1.TXTF17D.SETFOCUS
	                            RETURN
	                       ENDIF  
	                 ENDIF 
                   ENDIF                                                
		   SELECT S00
		   IF RLOCK()                              
		       REPLACE F11 WITH THISFORM.PAGEFRAME3.PAGE1.SPINNER11D.VALUE
		       REPLACE F12 WITH  IIF(THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.VALUE=1,'Y','N')
		       IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.VALUE=1
		            REPLACE F13 WITH 'N'	
		            REPLACE F14 WITH 0
		            REPLACE F15 WITH ''
		            REPLACE F16 WITH 0
		            REPLACE F17 WITH 0
		       ELSE
		            REPLACE F13 WITH 'Y'
		            REPLACE F14 WITH THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE
		            IF  THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1 
		               	 DO CASE 
			               CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=2
			                          REPLACE F15 WITH 'H'
			               CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=3
			                          REPLACE F15 WITH 'D'
			               CASE THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=4
			                          REPLACE F15 WITH 'B'	                              
			         ENDCASE
		            	 REPLACE F16 WITH THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE	
		            	 REPLACE F17 WITH 0 
		            ELSE
		             	 REPLACE F15 WITH ''
		            	 REPLACE F16 WITH 0
		            	 REPLACE F17 WITH THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE
		             ENDIF                 	  
		       ENDIF
		       REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	           ELSE
	                DO FILE_IMPACT
	                RETURN 		       
		   ENDIF    
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK        
         CASE S00.F01='E'     &&年假範圍設定-E
         	    IF EMPTY(THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE) AND CHK_NULL<>'1'   
                        =MESSAGEBOX('給予年假天數不得為零!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME4.PAGE1.TXTF20.SETFOCUS
                        RETURN
                    ENDIF     
                    SELECT S00
                    IF RLOCK()    
                    	REPLACE F20 WITH THISFORM.PAGEFRAME4.PAGE1.TXTF20.VALUE
			REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	            ELSE
	                DO FILE_IMPACT
	                RETURN 			
                    ENDIF
                    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
         CASE S00.F01='F'     &&年假限制設定-F
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2                       
         	          IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE)     
                              =MESSAGEBOX('您已選擇設定最大年資,故最大年資不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF21.SETFOCUS
                              RETURN
                          ENDIF                     
                   ENDIF
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=2                       
         	          IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE)     
                              =MESSAGEBOX('您已選擇設定年假上限,故年假上限不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF23.SETFOCUS
                              RETURN
                          ENDIF
                   ENDIF        
    		   IF  THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=2                       
         	         IF EMPTY(THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE)     
                              =MESSAGEBOX('您已選擇到職未滿一年之,故乘上到職月數之年假天數不得為零!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME5.PAGE1.TXTF25.SETFOCUS
                              RETURN
                          ENDIF
                   ENDIF        
                   SELECT S00
                   IF RLOCK()     
		        REPLACE F21 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE,0)
		        REPLACE F22 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE,0)
		        REPLACE F23 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE,0)
		        REPLACE F24 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=1,THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE,0)
		        REPLACE F25 WITH IIF(THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.VALUE=2,THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE,0)				      
	           ELSE
	                DO FILE_IMPACT
	                RETURN 		        
		   ENDIF  
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK      
         CASE S00.F01='G'           &&請假扣款項目-G
         	    IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE)     
                         =MESSAGEBOX('假別項目不得空白!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME6.PAGE1.TXTF27.SETFOCUS
                         RETURN
                    ENDIF    
                    IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=2 AND EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE) 
                          =MESSAGEBOX('限制年度缺勤數不得為零!',0+64,'提示視窗')
      		          THISFORM.PAGEFRAME6.PAGE1.TXTF29.SETFOCUS
                           RETURN
                    ENDIF                         
                    IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=2 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
         	          IF THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=1     
                              =MESSAGEBOX('扣款底數不得空白!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.SETFOCUS
                              RETURN
                          ENDIF                                 
	                  IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE)     
	                       =MESSAGEBOX('扣款比例不得為零!',0+64,'提示視窗')
	      		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.SETFOCUS
	                       RETURN
	                  ENDIF                       
	             ELSE     
	                  IF  THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=3 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
	         	         IF EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE)     
	                              =MESSAGEBOX('扣款金額不得為零!',0+64,'提示視窗')
	      		              THISFORM.PAGEFRAME6.PAGE1.TXTF32.SETFOCUS
	                              RETURN
	                         ENDIF 
	                   ENDIF         	                
	             ENDIF
		     IF THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=1 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
		          IF  THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE<>0  
		                IF  THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE<>0 AND THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE<=THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE     
	                             =MESSAGEBOX('全勤扣款方式一之缺勤數不可小於或等於不扣全勤緩衝數!',0+64,'提示視窗')
	      		             THISFORM.PAGEFRAME6.PAGE1.TXTF35.SETFOCUS
	                             RETURN
	                        ENDIF
		                IF THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE<>0 AND THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE<=THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE    
	                            =MESSAGEBOX('全勤扣款方式二之缺勤數不可小於或等於不扣全勤緩衝數!',0+64,'提示視窗')
	      		            THISFORM.PAGEFRAME6.PAGE1.TXTF37.SETFOCUS
	                            RETURN
	                       ENDIF
	                  ENDIF		        
	         	  IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE)     
	         	       IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE)
	                             =MESSAGEBOX('全勤扣款方式一之缺勤數您已輸入資料,故扣款比例不得為零',0+64,'提示視窗')
	      		             THISFORM.PAGEFRAME6.PAGE1.TXTF36.SETFOCUS
	                             RETURN
	                        ENDIF      
	                   ENDIF
	         	   IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE)     
	         	        IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE)
	                             =MESSAGEBOX('全勤扣款方式一之扣款比例您已輸入資料,故缺勤數不得為零',0+64,'提示視窗')
	      		             THISFORM.PAGEFRAME6.PAGE1.TXTF35.SETFOCUS
	                             RETURN
	                       ENDIF      
	                   ENDIF 	                        
	         	   IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE)     
	         	       IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE)
	                             =MESSAGEBOX('全勤扣款方式二之缺勤數您已輸入資料,故扣款比例不得為零',0+64,'提示視窗')
	      		             THISFORM.PAGEFRAME6.PAGE1.TXTF38.SETFOCUS
	                             RETURN
	                       ENDIF      
	                   ENDIF
	         	   IF !EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE)     
	         	       IF  EMPTY(THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE)
	                             =MESSAGEBOX('全勤扣款方式二之扣款比例您已輸入資料,故缺勤數不得為零',0+64,'提示視窗')
	      		             THISFORM.PAGEFRAME6.PAGE1.TXTF37.SETFOCUS
	                             RETURN
	                       ENDIF      
	                   ENDIF 	                     	        
		     ENDIF   		        
		     SELECT S00
		     IF RLOCK()
		        REPLACE F27 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF27.VALUE
		        REPLACE F28 WITH IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=1,'D',IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.VALUE=2,'H','S'))		            
		        REPLACE F29 WITH IIF(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.VALUE=2,THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE,0)
		        IF THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=1
		           REPLACE F30 WITH ''
			       REPLACE F31 WITH 0
			       REPLACE F32 WITH 0	
		           REPLACE F33 WITH 'N'	
			       REPLACE F34 WITH  0
			       REPLACE F35 WITH  0
			       REPLACE F36 WITH  0
			       REPLACE F37 WITH  0
			       REPLACE F38 WITH  0				       
		        ELSE
		              DO CASE
		                 CASE THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=1 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
			                  REPLACE F30 WITH ''
			                  REPLACE F31 WITH 0
			         	      REPLACE F32 WITH 0			            
		                 CASE THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=2 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
				               DO CASE 
			                      CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=2
			                           REPLACE F30 WITH 'H'
                                  CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=3
	                                   REPLACE F30 WITH 'D'
	                              CASE THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=4
	                                   REPLACE F30 WITH 'B'	                              
	                           ENDCASE	                      	    	            
		     			       REPLACE F31 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE
		           		       REPLACE F32 WITH 0
		                 CASE THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=3 AND THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE=2
			                  REPLACE F30 WITH ''
			                  REPLACE F31 WITH 0
			                  REPLACE F32 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE		            
		              ENDCASE  
		              IF THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=1 
		                 REPLACE F33 WITH  'Y'
			             REPLACE F34 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE
			             REPLACE F35 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE
			             REPLACE F36 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE
			             REPLACE F37 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE
			             REPLACE F38 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE
			          ELSE
		                 REPLACE F33 WITH  'N'	
			             REPLACE F34 WITH  0
			             REPLACE F35 WITH  0
			             REPLACE F36 WITH  0
			             REPLACE F37 WITH  0
			             REPLACE F38 WITH  0		            	 		         
			          ENDIF 		              
		          ENDIF    

  		        REPLACE F58 WITH ALLTRIM(STR(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.VALUE))  
			    REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.VALUE)) 
			    REPLACE F62 WITH THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE   
  	            REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'') 

	         ELSE   
	                DO FILE_IMPACT
	                RETURN 		          	
		     ENDIF     
		     THISFORM.SHRGROUP.ABANDON_BOTT.CLICK      
         CASE S00.F01='H'            &&懲處扣款設定-H
         	    IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE)     
                        =MESSAGEBOX('懲處項目不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME7.PAGE1.TXTF40.SETFOCUS
                        RETURN
                    ENDIF    
                    IF  THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=1
                          IF THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=1     
                               =MESSAGEBOX('扣款底數不得空白!',0+64,'提示視窗')
      		              THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.SETFOCUS
                              RETURN
                          ENDIF     
	                  IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE)     
	                      =MESSAGEBOX('扣款比例不得為零!',0+64,'提示視窗')
	      		      THISFORM.PAGEFRAME7.PAGE1.TXTF43.SETFOCUS
	                      RETURN
	                  ENDIF                       
	             ELSE     
	         	  IF EMPTY(THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE)     
	                       =MESSAGEBOX('扣款金額不得為零!',0+64,'提示視窗')
	      		       THISFORM.PAGEFRAME7.PAGE1.TXTF44.SETFOCUS
	                       RETURN
	                  ENDIF     	                
	             ENDIF
                     SELECT S00
		     IF RLOCK()   
		          REPLACE F40 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF40.VALUE
		          REPLACE F41 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF41.VALUE
		          IF THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.VALUE=1
		               DO CASE 
		                      CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=2
		                                 REPLACE F42 WITH 'H'
		                      CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=3
		                                 REPLACE F42 WITH 'D'
		                      CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=4
		                                 REPLACE F42 WITH 'B'
		                      CASE THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VALUE=5
		                                 REPLACE F42 WITH 'T'		                              
		               ENDCASE		            
		               REPLACE F43 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE
		               REPLACE F44 WITH 0
		          ELSE
		               REPLACE F42 WITH ''
		               REPLACE F43 WITH 0
		               REPLACE F44 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE
		          ENDIF 	     
		          REPLACE F45 WITH THISFORM.PAGEFRAME7.PAGE1.TXTF45.VALUE	            
		          REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'') 
		    ELSE
		          DO FILE_IMPACT
		          RETURN   
		    ENDIF   		        
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK      
        	    THISFORM.PAGEFRAME1.PAGE1.ACTIVATE    
         CASE S00.F01='I'   &&班別設定-I
	            IF EMPTY(THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE)     
                         =MESSAGEBOX('班別說明不得空白!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME8.PAGE1.TXTF49.SETFOCUS
                         RETURN
                    ENDIF             
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))=1
		       =MESSAGEBOX('上班時段不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF50.SETFOCUS
		       RETURN
		    ENDIF           
                    ** 
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE,2)))
		       =MESSAGEBOX('上班時段不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF50.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))>59
		        =MESSAGEBOX('上班時段格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF50.SETFOCUS
		        RETURN  
		    ENDIF        
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE,2)))
		       =MESSAGEBOX('下班時段不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF51.SETFOCUS
		        RETURN       
		    ENDIF			     
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))>59
		       =MESSAGEBOX('下班時段格式不正確!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF51.SETFOCUS
		       RETURN
		    ENDIF		    
		    **
		    IF  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))<>1
		         IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))<>5 OR ;
		             LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2)))=1 OR;
		             LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2)))=1 OR;
		             VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))>23 OR;
		             VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))>59
		             =MESSAGEBOX('起始休息時段格式不正確!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF52.SETFOCUS
		              RETURN
		        ENDIF      
		        IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE,2)))
             		     =MESSAGEBOX('起始休息時段已有資料故截止休息時段不得空白!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF53.SETFOCUS
		             RETURN       
		        ENDIF			        
		    ENDIF	
		    **
		    IF  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))<>1	         
		        IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE,2)))
             		     =MESSAGEBOX('截止休息時段已有資料故起始休息時段不得空白!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME8.PAGE1.TXTF52.SETFOCUS
		             RETURN       
		        ENDIF				 
			 IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))<>5 OR ;
			     LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2)))=1 OR;
			     LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2)))=1 OR;
			     VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))>23 OR;
			     VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))>59
			     =MESSAGEBOX('截止休息時段格式不正確!',0+64,'提示視窗')
			    THISFORM.PAGEFRAME8.PAGE1.TXTF53.SETFOCUS
			     RETURN
			 ENDIF    
		    ENDIF	  
		    IF EMPTY(THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE)
		       =MESSAGEBOX('上班總時數不得為零!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF57.SETFOCUS
		        RETURN       
		    ENDIF			    
		    **	
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE)) = 5
		        IF LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2)))=1 OR;
		            LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2)))=1 OR;
		           VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2))>23 OR;
		           VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE),2))>59
		           =MESSAGEBOX('上班刷卡時段止格式不正確!',0+64,'提示視窗')
		           THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		           RETUR	
		        ENDIF   
		    ENDIF    	    	 
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 
		        IF LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2)))=1 OR;
		            LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2)))=1 OR;
		            VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2))>23 OR;
		            VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE),2))>59
		            =MESSAGEBOX('下班刷卡時段起格式不正確!',0+64,'提示視窗')
		            THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		            RETUR	
		        ENDIF    
		    ENDIF  		    
		    ***
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE,2)))
		       =MESSAGEBOX('上班遲到緩衝時間可以與上班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))>59
		        =MESSAGEBOX('上班遲到緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  
		    ENDIF   
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))<VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
		        =MESSAGEBOX('上班遲到緩衝時間不得小於上班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE),2))<VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
		        =MESSAGEBOX('上班遲到緩衝時間不得小於上班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF54.SETFOCUS
		        RETURN  		         
		    ENDIF		    
		    **
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE,2)))
		       =MESSAGEBOX('下班早退緩衝時間可以與下班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>59
		        =MESSAGEBOX('下班早退緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  
		    ENDIF  
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('下班早退緩衝時間不得大於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE),2))>VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('下班早退緩衝時間不得大於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF55.SETFOCUS
		        RETURN  		         
		    ENDIF				       
		    **
		    IF EMPTY(ALLTRIM(LEFT(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE,2))) AND EMPTY(ALLTRIM(RIGHT(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE,2)))
		       =MESSAGEBOX('加班上班緩衝時間可以與下班時段相同但不得空白!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN       
		    ENDIF		          
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE))<>5 OR ;
		        LEN(ALLTRIM(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2)))=1 OR;
		        LEN(ALLTRIM(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2)))=1 OR;
		        VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))>23 OR;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))>59
		        =MESSAGEBOX('加班上班緩衝時間格式不正確!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  
		    ENDIF  	
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))<VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('加班上班緩衝時間不得小於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  		         
		    ENDIF
		    IF VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)) AND;
		        VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE),2))<VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))
		        =MESSAGEBOX('加班上班緩衝時間不得小於下班時段!',0+64,'提示視窗')
		        THISFORM.PAGEFRAME8.PAGE1.TXTF56.SETFOCUS
		        RETURN  		         
		    ENDIF	
		    ***		
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE < THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE
		       =MESSAGEBOX('上班刷卡時段止需大於上班時段!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		        RETURN       
		    ENDIF			    
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE < THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE
		       =MESSAGEBOX('上班刷卡時段止不可小於上班遲到緩衝時間!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF66.SETFOCUS
		        RETURN       
		    ENDIF		
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE > THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE
		       =MESSAGEBOX('下班刷卡時段起需小於下班時段!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		        RETURN       
		    ENDIF			    
		    IF LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5 AND LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE))=5 AND THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE > THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE
		       =MESSAGEBOX('下班刷卡時段起不可大於下班早退緩衝時間!',0+64,'提示視窗')
		       THISFORM.PAGEFRAME8.PAGE1.TXTF67.SETFOCUS
		        RETURN       
		    ENDIF		        	    		    	    
		    ***       		                                                                    	                          
	            SELECT S00
	            IF RLOCK()
	                 REPLACE F49 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF49.VALUE
	                 REPLACE F50 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE
	                 REPLACE F51 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE
	                 REPLACE F52 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE
	                 REPLACE F53 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE
	                 REPLACE F54 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE
	                 REPLACE F55 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE
	                 REPLACE F56 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE
	                 REPLACE F57 WITH THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
	                 REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(sys_oper,'A01'),A01.F03,'')
		         REPLACE F66 WITH IIF(LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE))=5,THISFORM.PAGEFRAME8.PAGE1.TXTF66.VALUE,'')
		         REPLACE F67 WITH IIF(LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE))=5,THISFORM.PAGEFRAME8.PAGE1.TXTF67.VALUE,'')                 	                 
	            ELSE
	                 DO FILE_IMPACT
	                 RETURN   
	            ENDIF     	                          
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   
        CASE S00_1.F01='J'  &&其他所得發放設定		        
		    SELECT S00
	            IF RLOCK()
		        REPLACE F59 WITH THISFORM.PAGEFRAME9.PAGE1.TXTF59.VALUE            
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	            ELSE
	                DO FILE_IMPACT
	                RETURN   		        
	            ENDIF 		                                                       		
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK       
         CASE S00_1.F01='K'   &&年終獎金發放設定-K	 	
                    DO CASE 
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE)
		      			   =MESSAGEBOX('方式一之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF691.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
		      			   =MESSAGEBOX('方式一之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.SETFOCUS
		        		   RETURN       
		    		      ENDIF		  
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE)
		      			   =MESSAGEBOX('方式二之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF692.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 1
		      			   =MESSAGEBOX('方式二之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.SETFOCUS
		        		   RETURN       
		    		      ENDIF		    	
                           CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE)
		      			   =MESSAGEBOX('方式三之基數不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF693.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF EMPTY(THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE)
		      			   =MESSAGEBOX('方式三之發放金額不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME10.PAGE1.TXTF71.SETFOCUS
		        		   RETURN       
		    		      ENDIF	        		      
                    ENDCASE 
                    SELECT S00
		    LOCK()
		    IF RLOCK()   
		        REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE))
		        DO CASE 
		              CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 1
		              		 REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0
		              CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 2         
		              		 REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0	
		               CASE THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.VALUE = 3  
		               		  REPLACE F69 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE
		               		  REPLACE F70 WITH ''    
		               		  REPLACE F71 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE    	               
		        ENDCASE    
		        REPLACE F72 WITH THISFORM.PAGEFRAME10.PAGE1.TXTF72.VALUE
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	            ELSE
	                 DO FILE_IMPACT
	                 RETURN   		        
		    ENDIF     
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   		
         CASE S00_1.F01='L'   &&未休年假獎金發放設定-L	 	
                    DO CASE 
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1	
                                      IF THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
		      			   =MESSAGEBOX('方式一之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.SETFOCUS
		        		   RETURN       
		    		      ENDIF		
                                      IF EMPTY(THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE)
		      			   =MESSAGEBOX('方式一之發放比率不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.TXTF69.SETFOCUS
		        		   RETURN       
		    		      ENDIF			    		        
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2		
                                      IF THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 1
		      			   =MESSAGEBOX('方式二之發放底數不得為空白!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.SETFOCUS
		        		   RETURN       
		    		      ENDIF		    	
                           CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3	
                                      IF EMPTY(THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE)
		      			   =MESSAGEBOX('方式三之發放金額不得為零!',0+64,'提示視窗')
		       			   THISFORM.PAGEFRAME11.PAGE1.TXTF71.SETFOCUS
		        		   RETURN       
		    		      ENDIF
                    ENDCASE                    
                    SELECT S00
		    LOCK()
		    IF RLOCK()   
		        REPLACE F68 WITH ALLTRIM(STR(THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE))
		        DO CASE 
		              CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 1
		              		  REPLACE F69 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0
		              CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 2         
		              		 REPLACE F69 WITH 0
		              		 DO CASE 
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 2
		                   			  REPLACE F70 WITH 'H'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 3
		                   			  REPLACE F70 WITH 'D'
		                               CASE THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 4
		                   			  REPLACE F70 WITH 'B'
		                          ENDCASE
		                          REPLACE F71 WITH 0	
		               CASE THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.VALUE = 3  
		               		  REPLACE F69 WITH 0
		               		  REPLACE F70 WITH ''    
		               		  REPLACE F71 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE    	               
		        ENDCASE    
		        REPLACE F72 WITH THISFORM.PAGEFRAME11.PAGE1.TXTF72.VALUE
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	            ELSE
	                 DO FILE_IMPACT
	                 RETURN   		        
		    ENDIF    
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   			               
        CASE S00_1.F01='X'  &&公司資料設定
         	   IF EMPTY(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_000.VALUE)     
                        =MESSAGEBOX('公司名稱(中文)不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE9.TXTF_INT_000.SETFOCUS
                        RETURN
                   ENDIF
         	   IF EMPTY(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_071.VALUE)     
                        =MESSAGEBOX('公司地址(中文)不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE9.TXTF_INT_071.SETFOCUS
                        RETURN
                   ENDIF  
         	   IF EMPTY(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_107.VALUE)     
                        =MESSAGEBOX('公司統編不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE9.TXTF_INT_107.SETFOCUS
                        RETURN
                   ENDIF 
         	   IF LEN(ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_107.VALUE))<8
                        =MESSAGEBOX('公司統編不得不得小於九碼!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE9.TXTF_INT_107.SETFOCUS
                        RETURN
                   ENDIF                                                                     
         	   IF EMPTY(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_078.VALUE)     
                        =MESSAGEBOX('公司電話不得空白!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME1.PAGE9.TXTF_INT_078.SETFOCUS
                        RETURN
                   ENDIF                       
                   SELECT A26
                   SET ORDER TO A26   && F01
		   SEEK 'INT_000'
		   IF FOUND()
		        REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_000.VALUE)
		   ENDIF 
		   SEEK 'INT_071'
		   IF FOUND()
		        REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_071.VALUE)
		   ENDIF 
		   SEEK 'INT_076'
		   IF FOUND()
		        REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_076.VALUE)
		   ENDIF 
		   SEEK 'INT_077'
		   IF FOUND()
		       REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_077.VALUE)
		   ENDIF 
		   SEEK 'INT_078'
		   IF FOUND()
		        REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_078.VALUE)
		   ENDIF 
		   SEEK 'INT_079'
		   IF FOUND()
		        REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_079.VALUE)
		   ENDIF 
		   SEEK 'INT_080'
		   IF FOUND()
		       REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_080.VALUE)
		   ENDIF 
		   SEEK 'INT_107'
		   IF FOUND()
		       REPLACE F06 WITH ALLTRIM(THISFORM.PAGEFRAME12.PAGE1.TXTF_INT_107.VALUE)
		   ENDIF 		        
		   THISFORM.SHRGROUP.ABANDON_BOTT.CLICK  
         CASE S00_1.F01='Y'   &&加班參數設定-Y	
		    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=2   
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)>60 
		             =MESSAGEBOX('加班時間不滿之分鐘數不可超過60分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN 
		         ENDIF     
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)=0 
		             =MESSAGEBOX('加班時間不滿之分鐘數不可為零!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN 
		         ENDIF		         
		    ENDIF 
		    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=2  
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)>60    
		             =MESSAGEBOX('加班時間若滿之分鐘數不可超過60分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF64.SETFOCUS
		             RETURN 
		         ENDIF  
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)=0    
		             =MESSAGEBOX('加班時間若滿之分鐘數不可為零!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF64.SETFOCUS
		             RETURN 
		         ENDIF    		           
		    ENDIF  	                       
                    IF THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=2 AND THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=2	         
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)>VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)
		             =MESSAGEBOX('加班時間不滿之分鐘數不可大於加班時間若滿之分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN  
		         ENDIF   	
		         IF VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE)=VAL(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE)
		             =MESSAGEBOX('加班時間不滿之分鐘數不可等於加班時間若滿之分鐘!',0+64,'提示視窗')
		             THISFORM.PAGEFRAME13.PAGE1.TXTF63.SETFOCUS
		             RETURN  
		         ENDIF  		          		         	         
		    ENDIF 
		    SELECT S00
	            IF RLOCK()
		        REPLACE F63 WITH IIF(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.VALUE=1,'F',ALLTRIM(THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE))
		        REPLACE F64 WITH IIF(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.VALUE=1,'F',ALLTRIM(THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE))
		        REPLACE F65 WITH ALLTRIM(STR(THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.VALUE))
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'') 
	            ELSE
	                DO FILE_IMPACT
	                RETURN   		        
	            ENDIF 			    
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           	
         CASE S00.F01='Z'   &&基本參數設定-Z
         	    IF EMPTY(THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE)     
                         =MESSAGEBOX('天數不得為零!',0+64,'提示視窗')
      		         THISFORM.PAGEFRAME14.PAGE1.SPINNER02.SETFOCUS
                         RETURN
                    ENDIF 
         	    IF EMPTY(THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE)     
                        =MESSAGEBOX('時數不得為零!',0+64,'提示視窗')
      		        THISFORM.PAGEFRAME14.PAGE1.SPINNER03.SETFOCUS
                         RETURN
                    ENDIF   
         	    IF THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE <= 0    
                   =MESSAGEBOX('報表欄位數必須大於零!',0+64,'提示視窗')
      		       THISFORM.PAGEFRAME14.PAGE1.TXTF73.SETFOCUS
                    RETURN
                ENDIF                              
		    SELECT S00
	            IF RLOCK()
		        REPLACE F02 WITH THISFORM.PAGEFRAME14.PAGE1.SPINNER02.VALUE
		        REPLACE F03 WITH THISFORM.PAGEFRAME14.PAGE1.SPINNER03.VALUE
		        REPLACE F04 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF04.VALUE
		        REPLACE F60 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF60.VALUE
		        REPLACE F61 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF61.VALUE
		        REPLACE F73 WITH THISFORM.PAGEFRAME14.PAGE1.TXTF73.VALUE		        
		        REPLACE F99 WITH DTOC(DATE())+' '+IIF(SEEK(SYS_OPER,'A01'),A01.F03,'')
	            ELSE
	                DO FILE_IMPACT
	                RETURN   		        
	            ENDIF 		                                                       		
		    THISFORM.SHRGROUP.ABANDON_BOTT.CLICK           	        		       	         	             	      
  ENDCASE 
ENDPROC
*******刪除
PROCEDURE ORPGROUP.DEL_BOTT.CLICK  
IF  MESSAGEBOX('是否確定要刪除此筆資料',4+32+256,'請確認')=6 
     DO CASE
           CASE  S00.F01='B'   &&加班費設定-B
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF  
           CASE  S00.F01='C'   &&遲到設定-C
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF  	
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 		       
           CASE  S00.F01='D'   &&早退設定-D
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 		
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 		              	       
           CASE  S00.F01='E'  &&年假範圍設定-E
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 	
           CASE  S00.F01='F'  &&年假限制設定-F
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 		
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 		              
           CASE  S00.F01='G'    &&缺勤扣款項目-G
         	       SELECT S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 	
           CASE  S00.F01='H'   &&懲處扣款設定-H
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 	
           CASE  S00.F01='I'    &&班別設定-I
                       IF  MESSAGEBOX('若S05.人事薪資基本檔中已存在此班別,則刪除後將無法判別該班別,請問是否仍要刪除?',4+32+256,'請確認')=6 
	         	    SELECT S00
	         	    LOCK()
	         	    IF RLOCK()
	         	        DELETE 
	         	        UNLOCK
	         	        SKIP -1
	         	        IF BOF()
	         	            GO TOP
	         	            ENDIF  
			    ELSE
			            DO FILE_IMPACT
			    ENDIF 	
		      ENDIF	
           CASE  S00.F01='J'   &&其他所得發放設定-J
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 	
           CASE  S00.F01='K'   &&年終獎金發放設定-K
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 			    
           CASE  S00.F01='L'   &&未休年假獎金發放設定-L
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 			          	      
           CASE  S00.F01='Y'   &&加班參數設定-Y
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 			      
           CASE  S00.F01='Z'   &&基本參數設定-Z
         	       SELECT  S00
         	       LOCK()
         	       IF RLOCK()
         	           DELETE 
         	           UNLOCK
         	           SKIP -1
         	           IF BOF()
         	               GO TOP
         	           ENDIF  
		       ELSE
		           DO FILE_IMPACT
		       ENDIF 
		       THISFORM.GRID1.ENABLED=.T. 
		       THISFORM.GRID1.SETALL("DYNAMICBACKCOLOR","","COLUMN")  
		       THISFORM.GRID1.SETFOCUS 		         	              	           	       	       	         		       
     ENDCASE 
     THISFORM.SHRGROUP.ABANDON_BOTT.CLICK   
   ENDIF
ENDPROC
*******
PROCEDURE SHRGROUP.ABANDON_BOTT.CLICK
      THISFORM.KEY_LIST.VISIBLE=.F. 
      THISFORM.SHRGROUP.ENABLED=.F.
      THISFORM.SHRGROUP.VISIBLE=.F.      
      THISFORM.ORPGROUP.ENABLED=.T.
      THISFORM.ORPGROUP.EDIT_BOTT.VISIBLE=.F.
      THISFORM.ORPGROUP.PNT_BOTT.ENABLED=.F.   &&判斷有無列印的權限     
      THISFORM.ORPGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.ORPGROUP.ESC_BOTT.ENABLED=.T.
      THISFORM.CMDGROUP.ENABLED=.T.
      THISFORM.CMDGROUP.SETALL('ENABLED',.T.,'COMMANDBUTTON')
      THISFORM.CMDGROUP.SETALL('FORECOLOR',RGB(0,0,0),'COMMANDBUTTON')
      THISFORM.CMDGROUP.SEEK_BOTT.ENABLED=.F. 
      THISFORM.SETALL('READONLY',.T.,'TEXTBOX') 
      THISFORM.SETALL('VISIBLE',.T.,'TEXTBOX') 
      THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP06.ENABLED=.F.            
      THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP07.ENABLED=.F. 
      THISFORM.PAGEFRAME1.PAGE1.OPTIONGROUP08.ENABLED=.F.      
      THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VISIBLE=.F.
      THISFORM.PAGEFRAME1.PAGE1.COMBOBOX46.VISIBLE=.F.         
      THISFORM.PAGEFRAME2.SETALL('READONLY',.T.,'SPINNER') 
      THISFORM.PAGEFRAME2.PAGE1.YESNO_CHOICE12C.ENABLED=.F.	
      THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP13C.ENABLED=.F.	 
      THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.ENABLED=.F.	
      THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=.F.  	     
      THISFORM.PAGEFRAME3.SETALL('READONLY',.T.,'SPINNER') 
      THISFORM.PAGEFRAME3.PAGE1.YESNO_CHOICE12D.ENABLED=.F. 
      THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP13D.ENABLED=.F.    
      THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.ENABLED=.F.
      THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=.F.  
      THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP21.ENABLED=.F.
      THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP23.ENABLED=.F.    
      THISFORM.PAGEFRAME5.PAGE1.OPTIONGROUP24.ENABLED=.F.  
      THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP28.ENABLED=.F.
      THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP29.ENABLED=.F.            
      THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.ENABLED=.F.    
      THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP58.ENABLED=.F.
      THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP68.ENABLED=.F.
      THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.ENABLED=.F.                
      THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.F.  
      THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VISIBLE=.F.      
      THISFORM.PAGEFRAME7.PAGE1.OPTIONGROUP42.ENABLED=.F.   
      THISFORM.PAGEFRAME8.PAGE1.TXTF50.READONLY=.T.
      THISFORM.PAGEFRAME8.PAGE1.TXTF51.READONLY=.T.  
      THISFORM.PAGEFRAME8.PAGE1.TXTF52.READONLY=.T.
      THISFORM.PAGEFRAME8.PAGE1.TXTF53.READONLY=.T.                         
      THISFORM.PAGEFRAME8.SETALL('READONLY',.T.,'SPINNER')
      THISFORM.PAGEFRAME10.PAGE1.SETALL('VISIBLE',.T.,'TEXTBOX') 
      THISFORM.PAGEFRAME10.PAGE1.SETALL('VISIBLE',.F.,'COMBOBOX')
      THISFORM.PAGEFRAME10.PAGE1.OPTIONGROUP68.ENABLED=.F.      
      THISFORM.PAGEFRAME11.PAGE1.SETALL('VISIBLE',.T.,'TEXTBOX') 
      THISFORM.PAGEFRAME11.PAGE1.SETALL('VISIBLE',.F.,'COMBOBOX')
      THISFORM.PAGEFRAME11.PAGE1.OPTIONGROUP68.ENABLED=.F.
      THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP63.ENABLED=.F.
      THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP64.ENABLED=.F. 
      THISFORM.PAGEFRAME13.PAGE1.OPTIONGROUP65.ENABLED=.F.
      THISFORM.PAGEFRAME14.SETALL('READONLY',.T.,'SPINNER') 
      THISFORM.GRID1.ENABLED=.T.
      THISFORM.GRID2.ENABLED=.T.
      THISFORM.GRID2.SETFOCUS  
       IF THISFORM.GRID2.ACTIVEROW=0 
           THISFORM.GRID1.SETFOCUS
           THISFORM.GRID1.REFRESH
       ELSE
           THISFORM.GRID2.SETFOCUS
           THISFORM.GRID2.REFRESH
       ENDIF
      UNLOCK ALL
  ENDPROC   
*********   
  PROCEDURE SHRGROUP.FINISH_BOTT.CLICK
      THISFORM.SHRGROUP.ABANDON_BOTT.CLICK
  ENDPROC
ENDDEFINE     
********* 列印     
 PROCEDURE PNT_PRC  
 
 ENDPROC
****************加班計算比例設定
DEFINE CLASS OPTIONGROUP08 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*420
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP08'
   OPTION1.CAPTION='計算基準'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='計算金額'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*280
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=''
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.VISIBLE=.F.
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.READONLY=.F.
           THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VISIBLE=.T.
      	   THISFORM.PAGEFRAME1.PAGE1.TXTF09.READONLY=.F. 
      	   THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=.T. 
	   THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE=0
	   THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=''	
	   THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VALUE=1	      
      ELSE     	         	    
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.VALUE=''
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.VISIBLE=.T.
           THISFORM.PAGEFRAME1.PAGE1.TXTF08.READONLY=.T.           
           THISFORM.PAGEFRAME1.PAGE1.COMBOBOX08.VISIBLE=.F.      
      	   THISFORM.PAGEFRAME1.PAGE1.TXTF09.READONLY=.T.     
      	   THISFORM.PAGEFRAME1.PAGE1.TXTF10.READONLY=.F. 
	   THISFORM.PAGEFRAME1.PAGE1.TXTF09.VALUE=''
	   THISFORM.PAGEFRAME1.PAGE1.TXTF10.VALUE=0	   
      ENDIF    
  ENDPROC    
ENDDEFINE  
**************************************加班津貼補助項目
DEFINE CLASS COMBOBOX46 AS COMBOBOX 
       VISIBLE=.F.
       HEIGHT=INT_015*25
       WIDTH=INT_015*140
       FONTSIZE=INT_015*9   
       BOUNDCOLUMN=1
       ROWSOURCETYPE=1
       STYLE=0   
       VALUE=1
       NAME='COMBOBOX46'
 PROCEDURE INIT    
       THIS.ADDITEM('')
       SELECT F07 FROM S01 WHERE F01='F' ORDER BY F07 INTO CURSOR S01_TEMP
       SELECT S01_TEMP
       GO TOP
       DO WHILE  !EOF()
              THIS.ADDITEM(ALLTRIM(S01_TEMP.F07))
      	      SELECT  S01_TEMP
      	      SKIP
       ENDDO
       SELECT S01_TEMP
       USE   
  ENDPROC  
   PROCEDURE INTERACTIVECHANGE
       IF  THIS.VALUE=1
            THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE=''
            THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE=''
            THISFORM.PAGEFRAME1.PAGE1.TXTF47.READONLY=.T.            
       ELSE
            THISFORM.PAGEFRAME1.PAGE1.TXTF46.VALUE=THIS.DISPLAYVALUE
            THISFORM.PAGEFRAME1.PAGE1.TXTF47.VALUE=0          
            THISFORM.PAGEFRAME1.PAGE1.TXTF47.READONLY=.F.              
       ENDIF	
       THISFORM.PAGEFRAME1.PAGE1.TXTF46.REFRESH     
       THISFORM.PAGEFRAME1.PAGE1.TXTF47.REFRESH          
  ENDPROC       
ENDDEFINE
******************************************每月遲到超過分鐘數是否設定另外加扣款
DEFINE CLASS OPTIONGROUP13C AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*500 
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP13C'
   OPTION1.CAPTION='不設定'
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='設定'
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*65
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF  THIS.VALUE=1
           THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.ENABLED=.F.
	   THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.READONLY=.T.      
	   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.READONLY=.T.     
	   THISFORM.PAGEFRAME2.PAGE1.TXTF16C.READONLY=.T.   
	   THISFORM.PAGEFRAME2.PAGE1.TXTF17C.READONLY=.T.   
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VISIBLE=.T. 
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=.F.
	   THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE=0 	   	
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=''  
           THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=''  
           THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=''              
      ELSE
           THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.ENABLED=.T.      
	   THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.READONLY=.F.      
	   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.READONLY=.F.     
	   THISFORM.PAGEFRAME2.PAGE1.TXTF16C.READONLY=.F.   
	   THISFORM.PAGEFRAME2.PAGE1.TXTF17C.READONLY=.T.   
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VISIBLE=.F. 
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=.T. 
           THISFORM.PAGEFRAME2.PAGE1.OPTIONGROUP15C.VALUE=1            
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=1         
           THISFORM.PAGEFRAME2.PAGE1.SPINNER14C.VALUE=0  	   
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=''  
           THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=0
           THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=''      
      ENDIF       
  ENDPROC    
ENDDEFINE  
*****************************************遲到超過分鐘數之另扣款項目
DEFINE CLASS OPTIONGROUP15C AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*390
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.
   NAME='OPTIONGROUP15C'
   OPTION1.CAPTION='扣款底數'
   OPTION1.FONTSIZE=INT_015*9      
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5 
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='扣款金額'
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*260
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF  THIS.VALUE=1   
	   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.READONLY=.F.      
	   THISFORM.PAGEFRAME2.PAGE1.TXTF16C.READONLY=.F.   
	   THISFORM.PAGEFRAME2.PAGE1.TXTF17C.READONLY=.T.  
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VISIBLE=.F. 	         
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=.T.         
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VALUE=1         
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=''  
           THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=0
           THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=''            
      ELSE
	   THISFORM.PAGEFRAME2.PAGE1.TXTF15C.READONLY=.T.      
	   THISFORM.PAGEFRAME2.PAGE1.TXTF16C.READONLY=.T.   
	   THISFORM.PAGEFRAME2.PAGE1.TXTF17C.READONLY=.F.  
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VISIBLE=.T. 	         
           THISFORM.PAGEFRAME2.PAGE1.COMBOBOX15C.VISIBLE=.F.                
           THISFORM.PAGEFRAME2.PAGE1.TXTF15C.VALUE=''  
           THISFORM.PAGEFRAME2.PAGE1.TXTF16C.VALUE=''
           THISFORM.PAGEFRAME2.PAGE1.TXTF17C.VALUE=0    
      ENDIF       
  ENDPROC        
ENDDEFINE   
****************************************每月早退超過分鐘數是否設定另外加扣款
DEFINE CLASS OPTIONGROUP13D AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*500
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP13D'
   OPTION1.CAPTION='不設定'
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='設定'
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*65
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF  THIS.VALUE=1    
           THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.ENABLED=.F.
	   THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.READONLY=.T.      
	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.READONLY=.T.     
	   THISFORM.PAGEFRAME3.PAGE1.TXTF16D.READONLY=.T.   
	   THISFORM.PAGEFRAME3.PAGE1.TXTF17D.READONLY=.T.   
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VISIBLE=.T. 
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=.F.
	   THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE=0 	   	
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=''  
           THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=''  
           THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=''              
      ELSE
           THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.ENABLED=.T.      
	   THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.READONLY=.F.      
	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.READONLY=.F.     
	   THISFORM.PAGEFRAME3.PAGE1.TXTF16D.READONLY=.F.   
	   THISFORM.PAGEFRAME3.PAGE1.TXTF17D.READONLY=.T.   
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VISIBLE=.F. 
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=.T. 
           THISFORM.PAGEFRAME3.PAGE1.OPTIONGROUP15D.VALUE=1            
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=1         
           THISFORM.PAGEFRAME3.PAGE1.SPINNER14D.VALUE=0  	   
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=''  
           THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=0
           THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=''      
      ENDIF  
  ENDPROC    
ENDDEFINE  
*************************************早退分鐘數之另扣款項目
DEFINE CLASS OPTIONGROUP15D AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*390
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.
   NAME='OPTIONGROUP15D'
   OPTION1.CAPTION='扣款底數'
   OPTION1.FONTSIZE=INT_015*9      
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5 
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='扣款金額'
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*260
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF  THIS.VALUE=1   
	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.READONLY=.F.      
	   THISFORM.PAGEFRAME3.PAGE1.TXTF16D.READONLY=.F.   
	   THISFORM.PAGEFRAME3.PAGE1.TXTF17D.READONLY=.T.  
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VISIBLE=.F. 	         
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=.T.         
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VALUE=1         
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=''  
           THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=0
           THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=''            
      ELSE
	   THISFORM.PAGEFRAME3.PAGE1.TXTF15D.READONLY=.T.      
	   THISFORM.PAGEFRAME3.PAGE1.TXTF16D.READONLY=.T.   
	   THISFORM.PAGEFRAME3.PAGE1.TXTF17D.READONLY=.F.  
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VISIBLE=.T. 	         
           THISFORM.PAGEFRAME3.PAGE1.COMBOBOX15D.VISIBLE=.F.                
           THISFORM.PAGEFRAME3.PAGE1.TXTF15D.VALUE=''  
           THISFORM.PAGEFRAME3.PAGE1.TXTF16D.VALUE=''
           THISFORM.PAGEFRAME3.PAGE1.TXTF17D.VALUE=0    
      ENDIF       
  ENDPROC      
ENDDEFINE   
****************年假最大年資設定
DEFINE CLASS  OPTIONGROUP21 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*340
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP21'
   OPTION1.CAPTION='不設定'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='任職滿'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*78
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF21.READONLY=.T. 
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF22.READONLY=.T.
	   THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE=''      	   
      ELSE
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF21.READONLY=.F.
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF22.READONLY=.F.
	   THISFORM.PAGEFRAME5.PAGE1.TXTF21.VALUE=0 
	   THISFORM.PAGEFRAME5.PAGE1.TXTF22.VALUE=0     	         
      ENDIF    
  ENDPROC   
ENDDEFINE 
****************年假最多天數設定
DEFINE CLASS  OPTIONGROUP23 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*340
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP23'
   OPTION1.CAPTION='不設定'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='年假最高給假天數'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*78
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF23.READONLY=.T. 
	   THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE=''       	   
      ELSE
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF23.READONLY=.F.  
	   THISFORM.PAGEFRAME5.PAGE1.TXTF23.VALUE=0      	       
      ENDIF    
  ENDPROC    
ENDDEFINE  
****************未滿一年之設定
DEFINE CLASS  OPTIONGROUP24 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*340
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP24'
   OPTION1.CAPTION='給予年假'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='以年假'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*150
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF24.READONLY=.F. 
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF25.READONLY=.T.
	   THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE=0 
	   THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE='' 	         	          	   
      ELSE
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF24.READONLY=.T.     
      	   THISFORM.PAGEFRAME5.PAGE1.TXTF25.READONLY=.F.       	
	   THISFORM.PAGEFRAME5.PAGE1.TXTF24.VALUE='' 
	   THISFORM.PAGEFRAME5.PAGE1.TXTF25.VALUE=0       	       
      ENDIF    
  ENDPROC    
ENDDEFINE   
****************************請假扣款設定-是否給薪
DEFINE CLASS OPTIONGROUP58 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*25
   WIDTH=INT_015*85
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP58'
   OPTION1.CAPTION='是'
   OPTION1.TOP=INT_015*5
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='否'
   OPTION2.TOP=INT_015*5
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*50
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
     THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VALUE=1     
     THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VALUE=2  
      IF THIS.VALUE=1
      	   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.ENABLED=.F.     
      	   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.ENABLED=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=.F.    
      	   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=.F.        	   
      	   THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.F.   
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=.F.   
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=.F.    
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=.F.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=.F.      	   	   
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=.T.      	
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF31.READONLY=.T.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF32.READONLY=.T.  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.READONLY=.T. 
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.READONLY=.T. 
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.READONLY=.T. 
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.READONLY=.T. 
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.READONLY=.T.  	    	      	   
      ELSE
      	   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.ENABLED=.T.     
      	   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.ENABLED=.T.   
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.OPTIONGROUP30.VISIBLE=.T.    
      	   THISFORM.PAGEFRAME6.PAGE1.YESNO_CHOICE33.VISIBLE=.T.        	   
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF30.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF31.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF31.VISIBLE=.T.   
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF32.VISIBLE=.T.    
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF33.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF34.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF35.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.LBLF37.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VISIBLE=.T.
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VISIBLE=.T.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF30.VALUE=''  	
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=''   
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=''  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=''  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=''  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=''  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=''  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=''         	   
      ENDIF    
  ENDPROC    
ENDDEFINE  
****************請假扣款設定-計數單位
DEFINE CLASS  OPTIONGROUP28 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*25
   WIDTH=INT_015*132
   BUTTONCOUNT=3
   VALUE=1
   ENABLED=.F.
   NAME='OPTIONGROUP28'
   OPTION1.CAPTION='日'
   OPTION1.FONTSIZE=INT_015*9      
   OPTION1.TOP=INT_015*5
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5 
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='時'
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.TOP=INT_015*5
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*42
   OPTION2.NAME='OPTION2'   	 
   OPTION3.CAPTION='次/月'
   OPTION3.FONTSIZE=INT_015*9    
   OPTION3.TOP=INT_015*5
   OPTION3.AUTOSIZE=.T.
   OPTION3.LEFT=INT_015*83
   OPTION3.NAME='OPTION3' 	
   PROCEDURE INTERACTIVECHANGE
  	DO CASE		&&計數單位
       	      CASE THIS.VALUE=1	&&日
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='天(不含)'
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝天數               天(含)' 
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤天數             ，扣款比例                    ％ '
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤天數             ，扣款比例                    ％ '		       	      		 
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 日' 
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.READONLY=.F.
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=1
       	      CASE THIS.VALUE=2	&&小時
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='小時(不含)' 
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝時數               時(含)' 
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤時數             ，扣款比例                    ％ '
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤時數             ，扣款比例                    ％ '		
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 小時'        
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.READONLY=.F.	     
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=0 		 		       	      
       	      CASE THIS.VALUE=3	&&次
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF291.CAPTION='次/月(不含)'
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF34.CAPTION='不扣全勤緩衝次數               次(含)' 
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF35.CAPTION='全勤扣款方式一： 缺勤次數             ，扣款比例                    ％ '
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF37.CAPTION='全勤扣款方式二： 缺勤次數             ，扣款比例                    ％ '		
       	      		 THISFORM.PAGEFRAME6.PAGE1.LBLF62.CAPTION='請假基數                 次/月'     
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.READONLY=.T.   	     
       	      		 THISFORM.PAGEFRAME6.PAGE1.TXTF62.VALUE=1 		 		       	      
       ENDCASE  
  ENDPROC    
ENDDEFINE     
*************************年度可休限制(請假缺勤扣款)
DEFINE CLASS OPTIONGROUP29 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*455
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP29'
   OPTION1.CAPTION='不限制'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='限制年度缺勤數最多不得超過'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*70
   OPTION2.NAME='OPTION2'  
   PROCEDURE INTERACTIVECHANGE
      DO CASE 
            CASE THIS.VALUE=1       	   
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF29.READONLY=.T.    
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE=''                
            CASE THIS.VALUE=2      	        	     
		       THISFORM.PAGEFRAME6.PAGE1.TXTF29.READONLY=.F. 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF29.VALUE=0           	         	    
      ENDCASE   
  ENDPROC    
ENDDEFINE 
*************************扣款設定(請假扣款)
DEFINE CLASS OPTIONGROUP30 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*455
   BUTTONCOUNT=3
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP30'
   OPTION1.CAPTION='不扣款'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='扣款底數'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*70
   OPTION2.NAME='OPTION2' 
   OPTION3.CAPTION='扣款金額'
   OPTION3.FONTSIZE=INT_015*9 	   
   OPTION3.TOP=INT_015*7
   OPTION3.AUTOSIZE=.T.
   OPTION3.LEFT=INT_015*325
   OPTION3.NAME='OPTION3'   
   PROCEDURE INTERACTIVECHANGE
      THISFORM.PAGEFRAME6.PAGE1.TXTF30.VALUE=''   
      DO CASE 
            CASE THIS.VALUE=1       	   
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.T.     
	      	       THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.F.      
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=.T.  	      	            
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF31.READONLY=.T.     
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF32.READONLY=.T. 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=''
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=''	                  
            CASE THIS.VALUE=2     	   
		       THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.F.     
		       THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.T.     
		       THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VALUE=1     
		       THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=.F.    		       	     	     
		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.READONLY=.F. 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.READONLY=.T. 
		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=0
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=''		      
            CASE THIS.VALUE=3       	   
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF30.VISIBLE=.T.     
	      	       THISFORM.PAGEFRAME6.PAGE1.COMBOBOX30.VISIBLE=.F.   
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF30.READONLY=.T.             
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF31.READONLY=.T.     
	      	       THISFORM.PAGEFRAME6.PAGE1.TXTF32.READONLY=.F.       	       
		       THISFORM.PAGEFRAME6.PAGE1.TXTF31.VALUE=''
		       THISFORM.PAGEFRAME6.PAGE1.TXTF32.VALUE=0	      	         	    
      ENDCASE   
  ENDPROC    
ENDDEFINE 
****************扣款設定(懲處扣款)
DEFINE CLASS OPTIONGROUP42 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*420
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.	   
   NAME='OPTIONGROUP42'
   OPTION1.CAPTION='扣款底數'
   OPTION1.FONTSIZE=INT_015*9  
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='扣款金額'
   OPTION2.FONTSIZE=INT_015*9 	   
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*280
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.VALUE=''
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.VISIBLE=.F.
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.READONLY=.F.
           THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VISIBLE=.T.
      	   THISFORM.PAGEFRAME7.PAGE1.TXTF43.READONLY=.F. 
      	   THISFORM.PAGEFRAME7.PAGE1.TXTF44.READONLY=.T. 
	   THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=0
	   THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=''		      
      ELSE     	         	    
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.VALUE=''
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.VISIBLE=.T.
           THISFORM.PAGEFRAME7.PAGE1.TXTF42.READONLY=.T.           
           THISFORM.PAGEFRAME7.PAGE1.COMBOBOX42.VISIBLE=.F.      
      	   THISFORM.PAGEFRAME7.PAGE1.TXTF43.READONLY=.T.     
      	   THISFORM.PAGEFRAME7.PAGE1.TXTF44.READONLY=.F. 
	   THISFORM.PAGEFRAME7.PAGE1.TXTF43.VALUE=''
	   THISFORM.PAGEFRAME7.PAGE1.TXTF44.VALUE=0	   
      ENDIF    
  ENDPROC    
ENDDEFINE 
******班別設定-上班時段
DEFINE CLASS TXTF50 AS TEXTBOX
        VISIBLE=.T.
        READONLY=.T.
        HEIGHT=INT_015*25     
	WIDTH=INT_015*45  
	MAXLENGTH=5
	INPUTMASK='99:99' 
   PROCEDURE INTERACTIVECHANGE	
        THISFORM.PAGEFRAME8.PAGE1.TXTF54.VALUE=THIS.VALUE
        F50_H=VAL(LEFT(ALLTRIM(THIS.VALUE),2))
        F50_M=VAL(RIGHT(ALLTRIM(THIS.VALUE),2))
        F51_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))+IIF(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)='00',24,0)
        F51_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))  
        F52_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))
        F52_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))   
        F53_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))
        F53_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))                       
        DO CASE
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=1 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=1
                         IF F50_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)-(F50_M-F51_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)+(F51_M-F50_M))/60,1)
                         ENDIF
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=5 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=5
                         IF F50_M>=F52_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)-(F50_M-F52_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)+(F52_M-F50_M))/60,1)
                         ENDIF
                         IF F53_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)-(F53_M-F51_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)+(F51_M-F53_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ENDIF                      
         ENDCASE
  ENDPROC  	
ENDDEFINE 	
******班別設定-下班時段
DEFINE CLASS TXTF51 AS TEXTBOX
        VISIBLE=.T.
        READONLY=.T.
        HEIGHT=INT_015*25     
	WIDTH=INT_015*45  
	MAXLENGTH=5
	INPUTMASK='99:99' 
   PROCEDURE INTERACTIVECHANGE	
        THISFORM.PAGEFRAME8.PAGE1.TXTF55.VALUE=THIS.VALUE
        THISFORM.PAGEFRAME8.PAGE1.TXTF56.VALUE=THIS.VALUE
        F50_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F50_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F51_H=VAL(LEFT(ALLTRIM(THIS.VALUE),2))+IIF(LEFT(ALLTRIM(THIS.VALUE),2)='00',24,0)
        F51_M=VAL(RIGHT(ALLTRIM(THIS.VALUE),2))  
        F52_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))
        F52_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))   
        F53_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))
        F53_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))                       
        DO CASE
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=1 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=1
                         IF F50_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)-(F50_M-F51_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)+(F51_M-F50_M))/60,1)
                         ENDIF
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=5 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=5
                         IF F50_M>=F52_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)-(F50_M-F52_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)+(F52_M-F50_M))/60,1)
                         ENDIF
                         IF F53_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)-(F53_M-F51_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)+(F51_M-F53_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ENDIF                      
         ENDCASE        
  ENDPROC       
ENDDEFINE 	
******班別設定-休息時段起
DEFINE CLASS TXTF52 AS TEXTBOX
        VISIBLE=.T.
        READONLY=.T.
        HEIGHT=INT_015*25     
	WIDTH=INT_015*45  
	MAXLENGTH=5
	INPUTMASK='99:99' 
   PROCEDURE INTERACTIVECHANGE	
        F50_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F50_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F51_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))+IIF(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)='00',24,0)
        F51_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))  
        F52_H=VAL(LEFT(ALLTRIM(THIS.VALUE),2))
        F52_M=VAL(RIGHT(ALLTRIM(THIS.VALUE),2))   
        F53_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))
        F53_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE),2))                       
        DO CASE
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=1 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=1
                         IF F50_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)-(F50_M-F51_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)+(F51_M-F50_M))/60,1)
                         ENDIF
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=5 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=5
                         IF F50_M>=F52_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)-(F50_M-F52_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)+(F52_M-F50_M))/60,1)
                         ENDIF
                         IF F53_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)-(F53_M-F51_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)+(F51_M-F53_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ENDIF                      
         ENDCASE    
  ENDPROC       
ENDDEFINE 	
******班別設定-休息時段止
DEFINE CLASS TXTF53 AS TEXTBOX
        VISIBLE=.T.
        READONLY=.T.
        HEIGHT=INT_015*25     
	WIDTH=INT_015*45  
	MAXLENGTH=5
	INPUTMASK='99:99' 
   PROCEDURE INTERACTIVECHANGE	
        F50_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F50_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF50.VALUE),2))
        F51_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))+IIF(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2)='00',24,0)
        F51_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF51.VALUE),2))  
        F52_H=VAL(LEFT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))
        F52_M=VAL(RIGHT(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE),2))   
        F53_H=VAL(LEFT(ALLTRIM(THIS.VALUE),2))
        F53_M=VAL(RIGHT(ALLTRIM(THIS.VALUE),2))                       
        DO CASE
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=1 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=1
                         IF F50_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)-(F50_M-F51_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F50_H)*60)+(F51_M-F50_M))/60,1)
                         ENDIF
              CASE LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF52.VALUE))=5 AND  LEN(ALLTRIM(THISFORM.PAGEFRAME8.PAGE1.TXTF53.VALUE))=5
                         IF F50_M>=F52_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)-(F50_M-F52_M))/60,1)
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F52_H-F50_H)*60)+(F52_M-F50_M))/60,1)
                         ENDIF
                         IF F53_M>=F51_M
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)-(F53_M-F51_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ELSE
                             THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE=ROUND((((F51_H-F53_H)*60)+(F51_M-F53_M))/60,1)+THISFORM.PAGEFRAME8.PAGE1.TXTF57.VALUE
                         ENDIF                      
         ENDCASE    
  ENDPROC       
ENDDEFINE 	
***********************************************是否影響全勤(遲到早退,缺勤扣款)
DEFINE CLASS YESNO_CHOICE AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*25
   WIDTH=INT_015*85
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='YESNO_CHOICE'
   OPTION1.CAPTION='是'
   OPTION1.TOP=INT_015*5
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='否'
   OPTION2.TOP=INT_015*5
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*50
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
      IF THIS.VALUE=1
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.READONLY=.F.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.READONLY=.F.      	
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.READONLY=.F.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.READONLY=.F.  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.READONLY=.F. 
	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=0
	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=0
	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=0
	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=0
	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=0	   	   	   	   
      ELSE
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.READONLY=.T.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.READONLY=.T.      	
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.READONLY=.T.      
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.READONLY=.T.  
      	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.READONLY=.T.   
	   THISFORM.PAGEFRAME6.PAGE1.TXTF34.VALUE=''
	   THISFORM.PAGEFRAME6.PAGE1.TXTF35.VALUE=''
	   THISFORM.PAGEFRAME6.PAGE1.TXTF36.VALUE=''
	   THISFORM.PAGEFRAME6.PAGE1.TXTF37.VALUE=''
	   THISFORM.PAGEFRAME6.PAGE1.TXTF38.VALUE=''      	     
      ENDIF    
  ENDPROC    
ENDDEFINE  
***********************************************是否設定加班時間計算方式一
DEFINE CLASS OPTIONGROUP63 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*400
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP63'
   OPTION1.CAPTION='不設定'
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='設定'
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*80
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
       THISFORM.PAGEFRAME13.PAGE1.TXTF63.READONLY=IIF(THIS.VALUE=1,.T.,.F.) 
       THISFORM.PAGEFRAME13.PAGE1.TXTF63.VALUE=IIF(THIS.VALUE=1,'','0')     	      
  ENDPROC    
ENDDEFINE  
***********************************************是否設定加班時間計算方式二
DEFINE CLASS OPTIONGROUP64 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*29
   WIDTH=INT_015*400
   BUTTONCOUNT=2
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP64'
   OPTION1.CAPTION='不設定'
   OPTION1.TOP=INT_015*7
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5
   OPTION1.FONTSIZE=INT_015*9    
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='設定'
   OPTION2.TOP=INT_015*7
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*80
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.NAME='OPTION2'
   PROCEDURE INTERACTIVECHANGE
       THISFORM.PAGEFRAME13.PAGE1.TXTF64.READONLY=IIF(THIS.VALUE=1,.T.,.F.) 
       THISFORM.PAGEFRAME13.PAGE1.TXTF64.VALUE=IIF(THIS.VALUE=1,'','0')     
  ENDPROC    
ENDDEFINE 
*****************年終獎金發放設定
DEFINE CLASS OPTIONGROUP681 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*85
   WIDTH=INT_015*375
   BUTTONCOUNT=3
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP681'
   OPTION1.CAPTION='方式一：基數'
   OPTION1.FONTSIZE=INT_015*9      
   OPTION1.TOP=INT_015*10
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5 
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='方式二：基數'
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.TOP=INT_015*36
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*5
   OPTION2.NAME='OPTION2'   	   
   OPTION3.CAPTION='方式三：基數'
   OPTION3.FONTSIZE=INT_015*9    
   OPTION3.TOP=INT_015*62
   OPTION3.AUTOSIZE=.T.
   OPTION3.LEFT=INT_015*5
   OPTION3.NAME='OPTION3'  	 
   PROCEDURE INTERACTIVECHANGE
      DO CASE 
             CASE THIS.Value = 1
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VISIBLE = .F.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VISIBLE = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VISIBLE = .T.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VISIBLE = .F. 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.READONLY = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.READONLY = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.READONLY = .T.     
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = 0        
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''           
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = '' 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''  
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 1
             CASE THIS.Value = 2
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VISIBLE = .T.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VISIBLE = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VISIBLE = .F.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VISIBLE = .T. 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.READONLY = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.READONLY = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.READONLY = .T.   
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''        
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = 0           
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = '' 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = ''  
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 1                        
             CASE THIS.Value = 3
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VISIBLE = .T.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VISIBLE = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VISIBLE = .T.
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VISIBLE = .F. 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.READONLY = .T.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.READONLY = .F.
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.READONLY = .F.        
                        THISFORM.PAGEFRAME10.PAGE1.TXTF691.VALUE = ''        
                        THISFORM.PAGEFRAME10.PAGE1.TXTF692.VALUE = ''           
                        THISFORM.PAGEFRAME10.PAGE1.TXTF693.VALUE = 0
                        THISFORM.PAGEFRAME10.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME10.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME10.PAGE1.TXTF71.VALUE = 0  
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME10.PAGE1.COMBOBOX702.VALUE = 1
      ENDCASE
  ENDPROC    
ENDDEFINE 
*****************未休年假獎金發放設定
DEFINE CLASS OPTIONGROUP682 AS OPTIONGROUP
   VISIBLE=.T.
   HEIGHT=INT_015*85
   WIDTH=INT_015*410
   BUTTONCOUNT=3
   VALUE=1
   ENABLED=.F.   
   NAME='OPTIONGROUP682'
   OPTION1.CAPTION='方式一：未休天數 Ｘ 發放底數'
   OPTION1.FONTSIZE=INT_015*9      
   OPTION1.TOP=INT_015*10
   OPTION1.AUTOSIZE=.T.
   OPTION1.LEFT=INT_015*5 
   OPTION1.NAME='OPTION1'
   OPTION2.CAPTION='方式二：未休天數 Ｘ 發放底數'
   OPTION2.FONTSIZE=INT_015*9    
   OPTION2.TOP=INT_015*36
   OPTION2.AUTOSIZE=.T.
   OPTION2.LEFT=INT_015*5
   OPTION2.NAME='OPTION2'   	   
   OPTION3.CAPTION='方式三：未休天數 Ｘ 發放金額'
   OPTION3.FONTSIZE=INT_015*9    
   OPTION3.TOP=INT_015*62
   OPTION3.AUTOSIZE=.T.
   OPTION3.LEFT=INT_015*5
   OPTION3.NAME='OPTION3'  	 
   PROCEDURE INTERACTIVECHANGE
      DO CASE 
             CASE THIS.Value = 1
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VISIBLE = .F.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VISIBLE = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VISIBLE = .T.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VISIBLE = .F. 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.READONLY = .F.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.READONLY = .F.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.READONLY = .T.     
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = 0        
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''  
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 1
             CASE THIS.Value = 2
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VISIBLE = .T.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VISIBLE = .F.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VISIBLE = .F.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VISIBLE = .T. 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.READONLY = .F.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.READONLY = .T.     
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''        
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = ''  
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 1                   
             CASE THIS.Value = 3
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VISIBLE = .T.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VISIBLE = .F.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VISIBLE = .T.
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VISIBLE = .F. 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.READONLY = .T.
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.READONLY = .F.     
                        THISFORM.PAGEFRAME11.PAGE1.TXTF69.VALUE = ''        
                        THISFORM.PAGEFRAME11.PAGE1.TXTF701.VALUE = '' 
                        THISFORM.PAGEFRAME11.PAGE1.TXTF702.VALUE = ''
                        THISFORM.PAGEFRAME11.PAGE1.TXTF71.VALUE = 0
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX701.VALUE = 1
                        THISFORM.PAGEFRAME11.PAGE1.COMBOBOX702.VALUE = 1    
      ENDCASE
  ENDPROC    
ENDDEFINE 




