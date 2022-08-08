****程式名稱 K18發票明細異常紀錄
CLOSE ALL
CLEAR
FLG=''
SET DECIMALS TO 9
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK SYS_OPER+'K18'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1
IF !USED('A14')
   SELE 0
   USE A14
ELSE 
  SELE A14
ENDIF
SET ORDER TO 1
IF !USED('A23')
   SELE 0
   USE A23
ELSE 
  SELE A23
ENDIF
SET ORDER TO 1
IF !USED('B01')
   SELE 0
   USE B01
ELSE
   SELE B01
ENDIF
SET ORDER TO 1   
IF !USED('C01')
   SELE 0
   USE C01
ELSE
   SELE C01
ENDIF
SET ORDER TO 1
IF !USED('D01')
   SELE 0
   USE D01
ELSE
   SELE D01
ENDIF
SET ORDER TO 1      
IF !USED('E02')
   SELE 0
   USE E02
ELSE
   SELE E02
ENDIF
SET ORDER TO 1 
 
*********
TKK18FORM=CREATEOBJECT("TKK18")
TKK18FORM.SHOW       
DEFINE CLASS TKK18 AS FORM
  CAPTION='K18.發票明細'
*!*	  AUTOCENTER=.T.  
  FONTSIZE=INT_015*9
  TOP=INT_015*90
  LEFT=INT_015*150    
  HEIGHT=INT_015*400
  WIDTH=INT_015*450
  CONTROLBOX=.F.
  BORDERSTYLE=2
  MAXBUTTON=.F.
  MINBUTTON=.F.
  MOVABLE=.F.
  CLOSABLE=.F.  
  SHOWTIPS=.T.
  SHOWWINDOW=1
  WINDOWTYPE=1
  NAME='TKK18'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*15,;
      TOP=INT_015*250,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入發票月份'
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='發票號碼'     
  ADD OBJECT LBL21 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''
  ADD OBJECT LBL3 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION='抓取明細'
  ADD OBJECT LBL31 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*100,;
      AUTOSIZE=.T.,;
      CAPTION=''      
  ADD OBJECT LBL4 AS LABEL WITH;
      LEFT=INT_015*20,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION='寫入筆數'            
  ADD OBJECT LBL41 AS LABEL WITH;
      LEFT=INT_015*85,;
      TOP=INT_015*160,;
      AUTOSIZE=.T.,;
      CAPTION=''                
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*160,;
      TOP=INT_015*245,;
      WIDTH=INT_015*80,;      
      HEIGHT=INT_015*30,;    
      FONTSIZE=INT_015*12
  ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
      LEFT=INT_015*50,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*120,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<Y.計算發票明細',;
      TOOLTIPTEXT='開始抓取進銷項單據資料',;
      NAME='CMND1'
  ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
      LEFT=INT_015*180,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*120,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<T.轉出文字檔',;
      TOOLTIPTEXT='開始匯出進銷項單據資料',;
      NAME='CMND2'      
  ADD OBJECT CMND3 AS COMMANDBUTTON WITH;
      LEFT=INT_015*330,;
      TOP=INT_015*330,;
      HEIGHT=INT_015*40,;
      WIDTH=INT_015*80,;
      FONTSIZE=INT_015*12,;
      CAPTION='\<C.離開',;
      TOOLTIPTEXT='離開,並離開此畫面!快速鍵->ALT+C'
      NAME='CMND3'
 ********     
      PROCEDURE INIT 
         THISFORM.SETALL('FONTSIZE',INT_015*12,'LABEL')
         THISFORM.SETALL('FONTSIZE',INT_015*12,'TEXTBOX')
         THISFORM.MTH_LIST.SETFOCUS
      ENDPROC
  ***
  PROC MTH_LIST.INIT    
       WITH THIS
            SELE A23
            GO TOP
            K=0
            S=0
            DO WHILE !EOF()                
               K=K+1
               IF INT_012=1
                   SHDT=PADL(ALLTRIM(STR(VAL(LEFT(DTOS(A23.F01),4))-1911)),4,'0')+'/'+SUBSTR(DTOS(A23.F01),5,2)
               ELSE
                   SHDT=LEFT(DTOS(A23.F01),4)+'/'+SUBSTR(DTOS(A23.F01),5,2)
               ENDIF    
               .ADDITEM(SHDT) &&               
               IF LEFT(DTOS(DATE()),6)=LEFT(DTOS(A23.F01),6)
                   S=K
               ENDIF       
               SKIP
            ENDDO   
            .VALUE=S
       ENDWITH   
  ENDPROC 
  ****        
 **********    
      PROCEDURE CMND1.CLICK
      IF MESSAGEBOX('是否確定要執行發票明細核對作業',4+32+256,'請確認') = 6 
         SELECT A23
         SEEK DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01'))
         IF !FOUND()
            =MESSAGEBOX('無此月份資料!',0+48,'提示訊息視窗')
            THISFORM.MTH_LIST.SETFOCUS
            RETURN
         ENDIF

         IF !USED('C01')
              SELE 0
              USE C01
         ELSE
              SELE C01
         ENDIF
          SET ORDER TO 1            
         IF !USED('D01')
              SELE 0
              USE D01
         ELSE
              SELE D01
         ENDIF
          SET ORDER TO 1          
         IF !USED('E02')
              SELE 0
              USE E02
         ELSE
              SELE E02
         ENDIF
          SET ORDER TO 1         
        B06='B06'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B06')
            SELE 0
            USE (B06) ALIA B06
         ELSE
            SELE B06
         ENDIF          
         SET ORDER TO 1                    
        D19='D19'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&D19')
            SELE 0
            USE (D19) ALIA D19
         ELSE
            SELE D19
         ENDIF
         SET ORDER TO 4            
        P19='P19'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&P19')
            SELE 0
            USE (P19) ALIA P19
         ELSE
            SELE P19
         ENDIF
         SET ORDER TO 5                     
        E20='E20'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&E20')
            SELE 0
            USE (E20) ALIA E20
         ELSE
            SELE E20
         ENDIF
         SET ORDER TO 4            
        C13='C13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&C13')
            SELE 0
            USE (C13) ALIA C13
         ELSE
            SELE C13
         ENDIF
         SET ORDER TO 4            
         K25='K25'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)
        IF !USED('&K25')
            SELE 0
            USE (K25) ALIA K25
         ELSE
            SELE K25
         ENDIF
         
         SET ORDER TO 1          
         SET FILTER TO F16<>'2'
         IF !USED('K2F')
            SELECT 0
            USE K2F
         ELSE
            SELECT K2F
         ENDIF
         SET ORDER TO 1      
         SELECT K2F.F01,K2F.F02,K25.F01,K25.F03,K25.F09 FROM K2F,K25 WHERE K25.F07=K2F.F01 INTO CURSOR K2FBK ORDER BY K2F.F01 NOWAIT         
         CHK=LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),4)+'/'+SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),5,2)
         
********先核對重算發票與明細
             SELECT K25
             Go Top
    Do While !Eof()
        THISFORM.LBL21.CAPTION=F07
    	If Left(F01,1)='3'
    		Select C13
    		Seek Left(K25.F15,10)
    		If Found()
    			Do While F02=Left(K25.F15,10)        
                    THISFORM.LBL31.CAPTION=F02+SPACE(2)+F05                        
    				Replace F17 With K25.F07
    				If Left(F04,4)<>'稅額'                    
    					Replace F16 With Icase(K25.F09='2','00',K25.F09='3','09',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) <> 0,'03',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) = 0,'02','06')                        
    				Endif
    				Skip
    			ENDDO
    	    ENDIF
    		
    	Else
    		DEP=''
    		Do Case
    		Case Left(K25.F15,2)$'BABB' And (K25.F03<>INT_117 And K25.F03<>INT_118)
    			DEP='D19'
    		Case Left(K25.F15,2)$'BMBR'
    			DEP='E20'
    		Case Left(K25.F15,2)='PB' Or (Left(K25.F15,2)='BB' AND (K25.F03=INT_117 Or K25.F03=INT_118))
    			DEP='P19'
    		Otherwise
    			DEP=''
    		Endcase
    		If !Empty(DEP)
    			Select &DEP
    			If DEP='P19' And (K25.F03=INT_117 Or K25.F03=INT_118)
    				Set Order To 5
    				If DEP='P19' And (K25.F03=INT_117 Or K25.F03=INT_118)
    					Set Order To 5
    					Seek '備註:'+Substr(K25.F15,3,6)+IIF(SUBSTR(K25.F15,11,4)='運費','進貨費用明細轉入',SPACE(1))
    					If Found()
    						Do While Left(F17,11)='備註:'+Substr(K25.F15,3,6)+IIF(SUBSTR(K25.F15,11,4)='運費','進貨費用明細轉入',SPACE(1))
                               THISFORM.LBL31.CAPTION=F02+SPACE(2)+F05                                                    
    							Replace F20 With K25.F07
    							If Left(F04,4)<>'稅額'
    								Replace F19 With Icase(K25.F09='2','00',K25.F09='3','09',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) <> 0,'03',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) = 0,'02','06')
    							Endif
    							Skip
    						Enddo
    					Endif
    				Endif
    			Else
    				Set Order To 4
    				Seek Left(K25.F15,10)
    				If Found()
    					Do While F02=Left(K25.F15,10)
                            THISFORM.LBL31.CAPTION=F02+SPACE(2)+F05                                                
    						Replace F20 With K25.F07
    						If Left(F04,4)<>'稅額'
    							Replace F19 With Icase(K25.F09='2','00',K25.F09='3','09',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) <> 0,'03',K25.F09='1' And Mod(Val(Right(K25.F01,1)),2) = 0,'02','06')
    						Endif
    						Skip
    					Enddo
    				Endif
    			Endif
    		Endif
    	Endif
    	Select K25
    	Skip
    Enddo

    Select K2FBK
    Go Top
    Do While !Eof()
    	If Left(F01_B,1)='3'
    		Select C13
    		Seek Left(K2FBK.F02,10)
    		If Found()
    			Do While F02=Left(K2FBK.F02,10)
    				Replace F17 With K2FBK.F01_A
    				If Left(F04,4)<>'稅額'
    					Replace F16 With Icase(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) = 0,'02','06')
    				Endif
    				Skip
    			ENDDO
    		Endif
    	Else
    		DEP=''
    		Do Case
    		Case Left(K2FBK.F02,2)$'BABB' And (K2FBK.F03<>INT_117 And K2FBK.F03<>INT_118)
    			DEP='D19'
    		Case Left(K2FBK.F02,2)$'BMBR'
    			DEP='E20'
            Case Left(K2FBK.F02,2)='PB' Or ( LEFT(K2FBK.F02,2)='BB' AND (K25.F03=INT_117 Or K25.F03=INT_118))
    			DEP='P19'
    		Otherwise
    			DEP=''
    		Endcase
    		If !Empty(DEP)
    			Select &DEP
    			If DEP='P19' And (K2FBK.F03=INT_117 Or K2FBK.F03=INT_118)
    				Set Order To 5
    				Seek '備註:'+Substr(K2FBK.F02,3,6)+IIF(SUBSTR(K25.F15,11,4)='運費','進貨費用明細轉入',SPACE(1))
    				If Found()
    					Do While Left(F17,11)=Substr(K2FBK.F02,3,6)+IIF(SUBSTR(K25.F15,11,4)='運費','進貨費用明細轉入',SPACE(1))
    						Replace F20 With K2FBK.F01_A
    						If Left(F04,4)<>'稅額'
    							Replace F19 With Icase(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) = 0,'02','06')
    						Endif
    						Skip
    					Enddo
    				Endif
    			Else
    				Set Order To 4
    				Seek Left(K2FBK.F02,10)
    				If Found()
    					Do While F02=Left(K2FBK.F02,10)
    						Replace F20 With K2FBK.F01_A
    						If Left(F04,4)<>'稅額'
    							Replace F19 With Icase(K2FBK.F09='2','00',K2FBK.F09='3','09',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) <> 0,'03',K2FBK.F09='1' And Mod(Val(Right(K2FBK.F01_B,1)),2) = 0,'02','06')
    						Endif
    						Skip
    					Enddo
    				Endif
    			Endif
    		Endif
    	Endif
    	Select K2FBK
    	Skip
    Enddo
    CREATE CURSOR C12ERR (類別 C(2),發票號碼 C(10),對象編號 C(5),對象簡稱 C(8),憑證單號 C(10),發票總額 N(14,2),明細總額 N(14,2),差異金額 N(14,2))
    INDEX ON 類別+發票號碼 TAG C12ERR1
    SET ORDER TO 1
    CREATE CURSOR D19ERR (類別 C(2),發票號碼 C(10),對象編號 C(5),對象簡稱 C(8),憑證單號 C(10),發票總額 N(14,2),明細總額 N(14,2),差異金額 N(14,2))  
    INDEX ON 類別+發票號碼 TAG D19ERR1  
    SET ORDER TO 1  
    CREATE CURSOR P19ERR (類別 C(2),發票號碼 C(10),對象編號 C(5),對象簡稱 C(8),憑證單號 C(10),發票總額 N(14,2),明細總額 N(14,2),差異金額 N(14,2))    
    INDEX ON 類別+發票號碼 TAG P19ERR1    
    SET ORDER TO 1  
    CREATE CURSOR E20ERR (類別 C(2),發票號碼 C(10),對象編號 C(5),對象簡稱 C(8),憑證單號 C(10),發票總額 N(14,2),明細總額 N(14,2),差異金額 N(14,2))      
    INDEX ON 類別+發票號碼 TAG E20ERR1    
    SET ORDER TO 1  
    SELECT K25
    GO TOP
    DO WHILE !EOF()
       TCT=0
       IF LEFT(K25.F01,1)='3'
          SELECT C13
          SET ORDER TO 3
          SEEK K25.F07
          IF FOUND()
             DO WHILE C13.F17=K25.F07
                TCT=TCT+C13.F08
                SKIP
             ENDDO   
             IF K25.F12<>ROUND(TCT,INT_068)
                THISFORM.LBL41.CAPTION=K25.F07
                SELECT C12ERR
                APPEND BLANK
                REPLACE 類別 WITH K25.F01
                REPLACE 發票號碼 WITH K25.F07
                REPLACE 對象編號 WITH K25.F03
                REPLACE 對象簡稱 WITH IIF(SEEK(K25.F03,'C01'),C01.F05,'')
                REPLACE 憑證單號 WITH LEFT(K25.F15,10)
                REPLACE 發票總額 WITH K25.F12
                REPLACE 明細總額 WITH TCT              
                REPLACE 差異金額 WITH K25.F12-TCT
             ENDIF   
          ELSE
             SELECT B06
             SET ORDER TO 2
             SEEK K25.F07
             IF FOUND()
                DO WHILE B06.F20=K25.F07
                   TCT=TCT+ROUND(B06.F04*B06.F14*B06.F17,INT_069)
                   SKIP
                ENDDO
                IF K25.F12<>TCT
                   SELECT C12ERR
                   APPEND BLANK
                   REPLACE 類別 WITH K25.F01
                   REPLACE 發票號碼 WITH K25.F03
                   REPLACE 對象編號 WITH K25.F03
                   REPLACE 對象簡稱 WITH IIF(SEEK(K25.F03,'C01'),C01.F05,'')
                   REPLACE 憑證單號 WITH LEFT(K25.F15,10)
                   REPLACE 發票總額 WITH K25.F12
                   REPLACE 明細總額 WITH TCT              
                   REPLACE 差異金額 WITH K25.F12-TCT                  
                ENDIF   
             ELSE
                THISFORM.LBL41.CAPTION=K25.F07
                SELECT C12ERR
                APPEND BLANK
                REPLACE 類別 WITH K25.F01
                REPLACE 發票號碼 WITH K25.F07
                REPLACE 對象編號 WITH K25.F03
                REPLACE 對象簡稱 WITH IIF(SEEK(K25.F03,'C01'),C01.F05,'')
                REPLACE 憑證單號 WITH LEFT(K25.F15,10)              
                REPLACE 發票總額 WITH K25.F12
                REPLACE 差異金額 WITH K25.F12
             ENDIF   
          ENDIF
       ELSE
            DEP=''
            Do Case
            Case Left(K25.F15,2)$'BABB' And (K25.F03<>INT_117 And K25.F03<>INT_118)
                  DEP='D19'
            Case Left(K25.F15,2)$'BMBR'
                DEP='E20'
            Case Left(K25.F15,2)='PB' Or ( LEFT(K25.F15,2)='BB' AND (K25.F03=INT_117 Or K25.F03=INT_118))
                DEP='P19'
            Otherwise
                DEP=''
            ENDCASE
            IF !EMPTY(DEP)            
               SELECT &DEP
               SET ORDER TO 3
               SEEK K25.F07
               IF FOUND()
                  DO WHILE F20=K25.F07
                     TCT=TCT+F08
                     SKIP
                  ENDDO
                  IF K25.F12<>TCT
                     THISFORM.LBL41.CAPTION=K25.F07
                     DO CASE
                        CASE DEP='P19'
                             SELECT P19ERR
                        CASE DEP='D19'
                             SELECT D19ERR
                        CASE DEP='E20'
                             SELECT E20ERR
                     ENDCASE   
                     APPEND BLANK
                     REPLACE 類別 WITH K25.F01
                     REPLACE 發票號碼 WITH K25.F07
                     REPLACE 對象編號 WITH K25.F03
                     REPLACE 對象簡稱 WITH IIF(DEP='E20',IIF(SEEK(K25.F03,'E02'),E02.F04,'') ,IIF(SEEK(K25.F03,'D01'),D01.F04,''))
                     REPLACE 發票總額 WITH K25.F12
                     REPLACE 憑證單號 WITH LEFT(K25.F15,10)                 
                     REPLACE 明細總額 WITH TCT              
                     REPLACE 差異金額 WITH K25.F12-TCT
                  ENDIF                 
               ELSE
                  THISFORM.LBL41.CAPTION=K25.F07
                  DO CASE
                     CASE DEP='P19'
                          SELECT P19ERR
                     CASE DEP='D19'
                          SELECT D19ERR
                     CASE DEP='E20'
                          SELECT E20ERR
                  ENDCASE                 
                  APPEND BLANK
                  REPLACE 類別 WITH K25.F01
                  REPLACE 發票號碼 WITH K25.F07
                  REPLACE 對象編號 WITH K25.F03
                  REPLACE 對象簡稱 WITH IIF(DEP='E20',IIF(SEEK(K25.F03,'E02'),E02.F04,'') ,IIF(SEEK(K25.F03,'D01'),D01.F04,''))
                  REPLACE 憑證單號 WITH LEFT(K25.F15,10)                 
                  REPLACE 發票總額 WITH K25.F12
                  REPLACE 差異金額 WITH K25.F12
               ENDIF
            ENDIF
         ENDIF   
         SELECT K25
         SKIP
      ENDDO
      SELECT C12ERR
      IF RECCOUNT()<>0
          FILE_NAME='C12ERR'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
          IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XLS
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF        
      ENDIF    
      SELECT D19ERR
      IF RECCOUNT()<>0
          FILE_NAME='D19ERR'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
          IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XLS 
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF             
      ENDIF    
      SELECT P19ERR
      IF RECCOUNT()<>0
          FILE_NAME='P19ERR'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
          IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XLS  
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF             
      ENDIF    
      SELECT E20ERR
      IF RECCOUNT()<>0
          FILE_NAME='E20ERR'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
          IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XLS   
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF                                     
      ENDIF              
  ENDIF

         THISFORM.CMND3.CLICK
      ENDPROC  
  ****    
      PROCEDURE CMND2.CLICK        
         IF !USED('E00')
            SELECT 0
            USE E00
         ELSE
            SELECT E00
         ENDIF
         SET ORDER TO 1       
         IF !USED('B01')
            SELECT 0
            USE B01
         ELSE
            SELECT B01
         ENDIF
         SET ORDER TO 1 
         IF !USED('C01')
              SELE 0
              USE C01
         ELSE
              SELE C01
         ENDIF
          SET ORDER TO 1            
         IF !USED('D01')
              SELE 0
              USE D01
         ELSE
              SELE D01
         ENDIF
          SET ORDER TO 1          
         IF !USED('E02')
              SELE 0
              USE E02
         ELSE
              SELE E02
         ENDIF
          SET ORDER TO 1   

        IF !USED('P03')
           SELECT 0
           USE P03
        ELSE
           SELECT P03
        ENDIF
        SET ORDER TO 1
        B02='B02'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)   
        IF !USED('&B02')
            SELECT 0
            USE (B02) ALIAS B02
        ELSE
            SELECT B02
        ENDIF        
        SET ORDER TO 1
        B04='B04'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)   
        IF !USED('&B04')
            SELECT 0
            USE (B04) ALIAS B04
        ELSE
            SELECT B04
        ENDIF        
        SET ORDER TO 1       
        B05='B05'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)   
        IF !USED('&B05')
            SELECT 0
            USE (B05) ALIAS B05
        ELSE
            SELECT B05
        ENDIF        
        SET ORDER TO 1               
        B06='B06'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B06')
            SELE 0
            USE (B06) ALIA B06
         ELSE
            SELE B06
         ENDIF          
         SET ORDER TO 1                    
        B07='B07'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)  
        IF !USED('&B07')
            SELECT 0
            USE (B07) ALIAS B07
        ELSE
            SELECT B07
        ENDIF        
        SET ORDER TO 1             
        B08='B08'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)  
        IF !USED('&B08')
            SELECT 0
            USE (B08) ALIAS B08
        ELSE
            SELECT B08
        ENDIF        
        SET ORDER TO 1                     
        B09='B09'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B09')
            SELE 0
            USE (B09) ALIA B09
         ELSE
            SELE B09
         ENDIF          
         SET ORDER TO 1                                                                              
        B10='B10'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B10')
            SELE 0
            USE (B10) ALIA B10
         ELSE
            SELE B10
         ENDIF          
         SET ORDER TO 1                                                                                    
        B40='B40'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B40')
            SELE 0
            USE (B40) ALIA B40
         ELSE
            SELE B40
         ENDIF          
         SET ORDER TO 1                                                                      
        B61='B61'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B61')
            SELE 0
            USE (B61) ALIA B61
         ELSE
            SELE B61
         ENDIF          
         SET ORDER TO 1                                                            
        B62='B62'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B62')
            SELE 0
            USE (B62) ALIA B62
         ELSE
            SELE B62
         ENDIF          
         SET ORDER TO 1                                                  
        B63='B63'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B63')
            SELE 0
            USE (B63) ALIA B63
         ELSE
            SELE B63
         ENDIF          
         SET ORDER TO 1                                        
        B64='B64'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B64')
            SELE 0
            USE (B64) ALIA B64
         ELSE
            SELE B64
         ENDIF          
         SET ORDER TO 1                              
        C38='C38'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&C38')
            SELE 0
            USE (C38) ALIA C38
         ELSE
            SELE C38
         ENDIF          
         SET ORDER TO 1                    
        D39='D39'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&D39')
            SELE 0
            USE (D39) ALIA D39
         ELSE
            SELE D39
         ENDIF          
         SET ORDER TO 1        

        B55='B55'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B55')
            SELE 0
            USE (B55) ALIA B55
         ELSE
            SELE B55
         ENDIF          
         SET ORDER TO 1
        B56='B56'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B56')
            SELE 0
            USE (B56) ALIA B56
         ELSE
            SELE B56
         ENDIF          
         SET ORDER TO 1     
        B16='B16'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&B16')
            SELE 0
            USE (B16) ALIA B16
         ELSE
            SELE B16
         ENDIF          
         SET ORDER TO 1              
        E13='E13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&E13')
            SELE 0
            USE (E13) ALIA E13
         ELSE
            SELE E13
         ENDIF          
         SET FILTER TO F08=.T.
         SET ORDER TO 2                              
         
        E14='E14'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&E14')
            SELE 0
            USE (E14) ALIA E14
         ELSE
            SELE E14
         ENDIF          
         SET ORDER TO 2       
            

        D19='D19'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&D19')
            SELE 0
            USE (D19) ALIA D19
         ELSE
            SELE D19
         ENDIF
         SET ORDER TO 4            
        P19='P19'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&P19')
            SELE 0
            USE (P19) ALIA P19
         ELSE
            SELE P19
         ENDIF
         SET ORDER TO 5                     
        E20='E20'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&E20')
            SELE 0
            USE (E20) ALIA E20
         ELSE
            SELE E20
         ENDIF
         SET ORDER TO 4            
         G07='G07'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G07')
            SELE 0
            USE (G07) ALIA G07
         ELSE
            SELE G07
         ENDIF
         SET ORDER TO 1                    
         G09='G09'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&G09')
            SELE 0
            USE (G09) ALIA G09
         ELSE
            SELE G09
         ENDIF
         SET ORDER TO 1                  
         P17='P17'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&P17')
            SELE 0
            USE (P17) ALIA P17
         ELSE
            SELE P17
         ENDIF
         SET ORDER TO 1                               
        C13='C13'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&C13')
            SELE 0
            USE (C13) ALIA C13
         ELSE
            SELE C13
         ENDIF
         SET ORDER TO 4       
        K25='K25'+LEFT(DTOS(CTOD(ALLTRIM(THISFORM.MTH_LIST.DISPLAYVALUE)+'/01')),6)        
        IF !USED('&K25')
            SELE 0
            USE (K25) ALIA K25
         ELSE
            SELE K25
         ENDIF
         SET ORDER TO 4                     
          OMT1=VAL(RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2))+1   &&計算本月最後一天為幾號
          IF OMT1>12
              OMT1=OMT1%12
              OMT2=VAL(LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),4))+1
           ELSE
              OMT2=VAL(LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),4))  
          ENDIF
          OMT3=RIGHT(DTOS(CTOD(ALLTRIM(STR(OMT2))+'/'+PADL(ALLTRIM(STR(OMT1)),2,'0')+'/01')-1),2)

          UKEF=THISFORM.MTH_LIST.DISPLAYVALUE
          TRN_RANGE=CREATEOBJECT("TRN_RANGE")  
          TRN_RANGE.SHOW    

          
          THISFORM.CMND3.CLICK
      ENDPROC      
  ****    
      PROCEDURE CMND3.CLICK        
           IF !USED('A05')
              SELECT 0
              USE A05
           ELSE
              SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'K18'
           IF FOUND()
              DELETE 
           ENDIF      
           CLOSE TABLE ALL
           THISFORM.RELEASE 
           RETURN  
      ENDPROC      
ENDDEFINE            
*****
***********************************************轉檔選項
DEFINE CLASS TRN_RANGE AS FORM
*!*	       AUTOCENTER=.T.
   CAPTION='請選擇轉檔範圍'
   TOP=INT_015*180
   LEFT=INT_015*220          
   FONTSIZE=INT_015*9
   HEIGHT=INT_015*170
   WIDTH=INT_015*355
   FONTSIZE=INT_015*9
   MOVABLE=.F.
   MAXBUTTON=.F.
   MINBUTTON=.F.      
   CONTROLBOX=.F.
   BORDERSTYLE=2
   SHOWTIPS=.T.
   SHOWWINDOW=1
   WINDOWTYPE=1
   NAME='TRN_RANGE'
  ADD OBJECT SIDE_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*60,;
      TOP=INT_015*40,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*185,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='內帳',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='外帳',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*100,;
      OPTION2.TOP=INT_015*3,;
      OPTION2.AUTOSIZE=.T.,;      
      NAME='SIDE_OPTION' 
  ADD OBJECT INOUT_OPTION AS OPTIONGROUP WITH;
      LEFT=INT_015*60,;
      TOP=INT_015*95,;
      HEIGHT=INT_015*25,;    
      WIDTH=INT_015*185,;      
      BUTTONCOUNT=3,;
      OPTION1.CAPTION='進項',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='銷項',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*64,;
      OPTION2.TOP=INT_015*3,;
      OPTION2.AUTOSIZE=.T.,;      
      OPTION3.CAPTION='全部',;
      OPTION3.FONTSIZE=INT_015*9,;
      OPTION3.LEFT=INT_015*124,;
      OPTION3.TOP=INT_015*3,;
      OPTION3.AUTOSIZE=.T.,;      
      NAME='INOUT_OPTION'
   ADD OBJECT CMND1 AS COMMANDBUTTON WITH;
       LEFT=INT_015*60,;    
       HEIGHT=INT_015*25,;   
       TOP=INT_015*140,;  
       WIDTH=INT_015*50,;
       CAPTION='\<Y.執  行',;
       TOOLTIPTEXT='確認所輸入的範圍鍵值!快速鍵->ALT+Y',;
       NAME='CMND1'
   ADD OBJECT CMND2 AS COMMANDBUTTON WITH;
       LEFT=INT_015*200,;
       TOP=INT_015*140,;
       HEIGHT=INT_015*25,;
       WIDTH=INT_015*50,;
       CAPTION='\<C.取  消',;
       TOOLTIPTEXT='取消作業!快速鍵->ALT+C',;
       NAME='CMND2'   
   PROCEDURE INIT 
      
      THISFORM.SETALL('FONTSIZE',INT_015*9,'COMMANDBUTTON') 
      THISFORM.SETALL('MOUSEPOINTER',99,'COMMANDBUTTON')
      THISFORM.SETALL('MOUSEICON','BMPS\harrow.CUR','COMMANDBUTTON')          
      THISFORM.SIDE_OPTION.VALUE=1               
      THISFORM.INOUT_OPTION.VALUE=1                
          
   ENDPROC             
   
   PROCEDURE CMND1.CLICK
           CREATE CURSOR TRNAC;   &&進銷累積檔
         (F01 C(4),F02 C(15),F03 C(8),F04 C(2),F05 C(3),F06 C(3),F07 C(10),F08 C(10),F09 C(2),F10 C(2),F11 C(11),F12 C(3),F13 C(8),F14 C(20),F15 C(15),;
         F16 C(15),F17 C(1),F18 C(10),F19 C(20),F20 C(1),F21 C(4),F22 C(15),F23 C(15),F24 C(1),F25 C(10),F26 C(10),F27 C(12),F28 C(20),F29 C(15),;         
         F30 C(15),F31 C(15),F32 C(4),F33 C(20),F34 C(1),F35 C(12),F36 C(15),F37 C(13),F38 C(15),F39 C(13),F40 C(15),F41 C(30),F42 C(90),F43 C(1),F44 C(10),;
         F45 C(250),F46 C(100),F47 C(100),F48 C(100),F49 C(100),F50 C(100),F51 C(15),F52 C(15),F53 C(15),F54 C(15),F55 C(15),F56 C(8),F57 C(8),;
         F58 C(100),F59 C(100),F60 C(100),F61 C(100),F62 C(100),F63 C(15),F64 C(15),F65 C(15),F66 C(15),F67 C(15),F68 C(8),F69 C(8),F70 C(1))
        
   
     IF THISFORM.INOUT_OPTION.VALUE=2  OR THISFORM.INOUT_OPTION.VALUE=3  &&如果只轉銷項或全部
        SELECT K25
        SET ORDER TO 1
         SELECT * FROM C13 WHERE F05=ANY(SELE F01 FROM B01) AND !(F08=0 AND F06>0) INTO CURSOR C13T1 ORDER BY F02,F05 NOWAIT
         H=0         
         SELECT C13T1
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=F02
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH C13T1.F02
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(C13T1.F01),4))-1911))+RIGHT(DTOS(C13T1.F01),4),8,' ')
            REPLACE F04 WITH IIF(LEFT(C13T1.F02,2)='BC','20',IIF(C13T1.F18='03','22','21'))
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號或由海外倉出貨   
               IF C13T1.F08>0
                  SELECT B04
                  SEEK C13T1.F02+C13T1.F05+C13T1.F04
                  IF FOUND() AND LEN(ALLTRIM(B04.F05))=5 AND LEFT(B04.F05,1)='S'
                     SELECT TRNAC
                     REPLACE F05 WITH IIF(SEEK(B04.F05,'A14'),IIF(EMPTY(A14.F06),'A00',A14.F06),'A00')
                  ELSE       
                    SELECT TRNAC                             
                    REPLACE F05 WITH 'A00'          
                  ENDIF        
               ELSE
                  SELECT B05
                  SEEK C13T1.F02+C13T1.F05+C13T1.F04
                  IF FOUND() AND LEN(ALLTRIM(B05.F05))=5 AND LEFT(B05.F05,1)='S'
                     SELECT TRNAC
                     REPLACE F05 WITH IIF(SEEK(B05.F05,'A14'),IIF(EMPTY(A14.F06),'A00',A14.F06),'A00')
                  ELSE       
                    SELECT TRNAC                             
                    REPLACE F05 WITH 'A00'          
                  ENDIF            
               ENDIF  
            ENDIF                              
            REPLACE F07 WITH C13T1.F03
*!*	            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳
*!*	               REPLACE F08 WITH 'NNNNNNNNNN' 
*!*	            ELSE
*!*	               REPLACE F08 WITH 'NNNNNYNNNN'
*!*	            ENDIF   
            REPLACE F09 WITH '00'
            REPLACE F10 WITH '00'            
            REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'C01'),C01.F33,'')
            REPLACE F12 WITH IIF(C13T1.F13=INT_011,SPACE(3),IIF(C13T1.F13='US ','USD',C13T1.F13))
            REPLACE F13 WITH IIF(C13T1.F13=INT_011,'1'+SPACE(7),PADL(ALLTRIM(TRANSFORM(C13T1.F14,'@R ###.####')),8,'0'))
            REPLACE F15 WITH C13T1.F04
            REPLACE F17 WITH RIGHT(C13T1.F16,1)
            REPLACE F18 WITH C13T1.F17

            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(C13T1.F79,'B01'),B01.F99,'')
            IF THISFORM.SIDE_OPTION.VALUE=2 AND ALLTRIM(F33)='00005962'
               REPLACE F20 WITH 'N'
               REPLACE F21 WITH IIF(LEFT(C13T1.F02,2)='BC','202',IIF(C13T1.F18='03','22','21'))
            ELSE
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH IIF(LEFT(C13T1.F02,2)='BC','20',IIF(C13T1.F18='03','22','21'))
            ENDIF            
            
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(C13T1.F06),'@R ############')),12,'0')
*!*	            IF C13T1.F13 <>INT_011
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(ROUND(C13T1.F07,6),'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(ROUND(ABS(C13T1.F06*C13T1.F07),2),'@R ##########.##')),13,'0')                        
*!*	            ENDIF
            REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(C13T1.F07*C13T1.F14,'@R #######.######')),15,'0')
            REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(ABS(IIF(C13T1.F16='02',ROUND(C13T1.F08/(1+INT_002/100),INT_069),C13T1.F08)),'@R #############')),13,'0')                         
            REPLACE F43 WITH IIF(C13T1.F16='02','Y','N')
            REPLACE F44 WITH C13T1.F03
            REPLACE F70 WITH ICASE(IIF(SEEK(C13T1.F17,'K25'),K25.F09+K25.F01,'')='232','1',IIF(SEEK(C13T1.F17,'K25'),K25.F09+K25.F01,'')='231','4','')
            SELECT C13T1
            SKIP
            IF F02<>K
               H=0
            ENDIF
         ENDDO          
         SELECT * FROM C13 WHERE F05=ANY(SELE F01 FROM B01) AND F08=0 AND F06>0 INTO CURSOR C13T2 ORDER BY F02,F05 NOWAIT
         H=0         
         SELECT C13T2
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=F02
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH C13T2.F02+'A'
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(C13T2.F01),4))-1911))+RIGHT(DTOS(C13T2.F01),4),8,' ')
            REPLACE F04 WITH '70'
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
               REPLACE F05 WITH 'A00'               
            ENDIF                              
            REPLACE F07 WITH C13T2.F03       
            REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'C01'),C01.F33,'')            
            REPLACE F15 WITH C13T2.F04
            REPLACE F17 WITH RIGHT(C13T2.F16,1) 
*!*	            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳 
*!*	               REPLACE F08 WITH 'NNNNNYNNNN'
*!*	            ELSE
*!*	               REPLACE F08 WITH 'NNNNNNNNNN'
*!*	            ENDIF   
            REPLACE F20 WITH 'Y'
            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(C13T2.F79,'B01'),B01.F99,'')
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(C13T2.F06),'@R ############')),12,'0')
            SELECT C13T2
            SKIP
            IF F02<>K
               H=0
            ENDIF
         ENDDO        
     ENDIF    
        IF THISFORM.INOUT_OPTION.VALUE=1 OR THISFORM.INOUT_OPTION.VALUE=3  &&只轉進項或全部        
         SELECT * FROM D19 WHERE F05 =ANY(SELECT F01 FROM B01) UNION ALL SELECT * FROM P19 WHERE F05=ANY (SELECT F01 FROM B01 WHERE LEFT(F98,1)='*');
         AND LEFT(F04,2)<>'FE' ORDER BY F02,F05 INTO CURSOR D19T NOWAIT
         H=0         
         SELECT D19T
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=IIF(F03=INT_117,LEFT(F02,6)+SUBSTR(F17,6,6),F02)
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH IIF(D19T.F03=INT_117,LEFT(D19T.F02,6)+SUBSTR(D19T.F17,6,6),D19T.F02)
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(D19T.F01),4))-1911))+RIGHT(DTOS(D19T.F01),4),8,' ')
            REPLACE F04 WITH IIF(LEFT(D19T.F02,2)$'BAPB','10',IIF(D19T.F21='03','12','11'))
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
               REPLACE F05 WITH 'A00'               
            ENDIF                              
            REPLACE F07 WITH D19T.F03
            REPLACE F09 WITH '00'
            REPLACE F10 WITH '00'             
            REPLACE F12 WITH IIF(D19T.F13=INT_011,SPACE(3),IIF(D19T.F13='US ','USD',D19T.F13))
            REPLACE F13 WITH IIF(D19T.F13=INT_011,'1'+SPACE(7),PADL(ALLTRIM(TRANSFORM(D19T.F14,'@R ###.####')),8,'0'))
            REPLACE F16 WITH D19T.F04
            REPLACE F17 WITH RIGHT(D19T.F19,1)
            REPLACE F18 WITH D19T.F20
            REPLACE F20 WITH 'Y'
            REPLACE F21 WITH IIF(LEFT(D19T.F02,2)$'BAPB','10',IIF(D19T.F21='03','12','11'))
            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(D19T.F05,'B01'),B01.F99,'')
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(D19T.F06),'@R #########.##')),12,'0')
*!*	            IF D19T.F13<>INT_011
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(ROUND(D19T.F07,6),'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(ROUND(ABS(D19T.F06*D19T.F07),2),'@R ##########.##')),13,'0')                        
*!*	            ENDIF            
            REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(D19T.F07*D19T.F14,'@R #######.######')),15,'0')
            REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(ABS(D19T.F08),'@R #############')),13,'0')             
            REPLACE F43 WITH IIF(D19T.F19='02','Y','N')
            REPLACE F44 WITH D19T.F03
            SELECT D19T
            SKIP
            IF F03<>INT_117
               IF F02<>K
                  H=0
               ENDIF
            ELSE
               IF LEFT(D19T.F02,6)+SUBSTR(D19T.F17,6,6)<>K
                  H=0
               ENDIF
            ENDIF   
         ENDDO                       
         SELECT * FROM E20 WHERE EMPTY(F15) AND F05 =ANY(SELECT F01 FROM B01) AND F04 NOT IN(SELECT F01 FROM E14 WHERE F08<>.T.);  
         AND LEFT(F04,2)<>'FE' ORDER BY F02,F05 INTO CURSOR E20T1 NOWAIT   &&無耗料者且整組入貨視為採購進貨
         H=0         
         SELECT E20T1
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=F02
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH E20T1.F02
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(E20T1.F01),4))-1911))+RIGHT(DTOS(E20T1.F01),4),8,' ')
            REPLACE F04 WITH IIF(LEFT(E20T1.F02,2)='BM','10',IIF(E20T1.F21='03','12','11'))
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
               REPLACE F05 WITH 'A00'               
            ENDIF                              
            REPLACE F07 WITH E20T1.F03
            REPLACE F09 WITH '00'
            REPLACE F10 WITH '00'             
            REPLACE F12 WITH IIF(E20T1.F13=INT_011,SPACE(3),IIF(E20T1.F13='US ','USD',E20T1.F13))
            REPLACE F13 WITH IIF(E20T1.F13=INT_011,'1'+SPACE(7),PADL(ALLTRIM(TRANSFORM(E20T1.F14,'@R ###.####')),8,'0'))
            REPLACE F16 WITH E20T1.F04
            REPLACE F17 WITH RIGHT(E20T1.F19,1)
            REPLACE F18 WITH E20T1.F20
            REPLACE F20 WITH 'Y'
            REPLACE F21 WITH IIF(LEFT(E20T1.F02,2)='BM','10',IIF(E20T1.F21='03','12','11'))
            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(E20T1.F05,'B01'),B01.F99,'')
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(IIF(EMPTY(E20T1.F15),ABS(E20T1.F06),0),'@R ##########.##')),12,'0')

               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(ROUND(E20T1.F07,6),'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(ROUND(ABS(E20T1.F06*E20T1.F07),2),'@R ##########.##')),13,'0')                        
  
            REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(E20T1.F07*E20T1.F14,'@R #######.######')),15,'0')
            REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(ABS(E20T1.F08),'@R #############')),13,'0')             
            REPLACE F43 WITH IIF(E20T1.F19='02','Y','N')
            REPLACE F44 WITH E20T1.F03
            SELECT E20T1
            SKIP
            IF F02<>K
               H=0
            ENDIF
         ENDDO            
         SELECT * FROM E20 WHERE F05 =ANY(SELECT F01 FROM B01) AND F04 = ANY (SELECT F01 FROM E14 WHERE F08<>.T.);  
         AND LEFT(F04,2)<>'FE' ORDER BY F02,F05 INTO CURSOR E20T2 NOWAIT   &&有耗料者視為進貨費用
         H=0         
         SELECT E20T2
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=F02
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH E20T2.F02+'A'
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(E20T2.F01),4))-1911))+RIGHT(DTOS(E20T2.F01),4),8,' ')
            REPLACE F04 WITH IIF(LEFT(E20T2.F02,2)='BM','10',IIF(E20T2.F21='03','12','11'))
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
               REPLACE F05 WITH 'A00'               
            ENDIF                              
            REPLACE F07 WITH E20T2.F03
            REPLACE F09 WITH '00'
            REPLACE F10 WITH '00'             
            REPLACE F12 WITH IIF(E20T2.F13=INT_011,SPACE(3),IIF(E20T2.F13='US ','USD',E20T2.F13))
            REPLACE F13 WITH IIF(E20T2.F13=INT_011,'1'+SPACE(7),PADL(ALLTRIM(TRANSFORM(E20T2.F14,'@R ###.####')),8,'0'))
            REPLACE F16 WITH E20T2.F04
            REPLACE F17 WITH RIGHT(E20T2.F19,1)
            REPLACE F18 WITH E20T2.F20
            
            IF THISFORM.SIDE_OPTION.VALUE=1
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH IIF(LEFT(E20T2.F02,2)='BM','10',IIF(E20T2.F21='03','12','11'))
               
            ELSE
               REPLACE F21 WITH IIF(LEFT(E20T2.F02,2)='BM','102',IIF(E20T2.F21='03','12','11'))
               REPLACE F20 WITH 'N'
            ENDIF   
            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(E20T2.F05,'B01'),B01.F99,'')
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL('0',12,'0')

               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(ROUND(E20T2.F07,6),'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(ROUND(ABS(E20T2.F06*E20T2.F07),2),'@R ##########.##')),13,'0')                        
      
            REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(E20T2.F07*E20T2.F14,'@R #######.######')),15,'0')
            REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(ABS(E20T2.F08),'@R #############')),13,'0')             
            REPLACE F43 WITH IIF(E20T2.F19='02','Y','N')
            REPLACE F44 WITH E20T2.F03
            SELECT E20T2
            SKIP
            IF F02<>K
               H=0
            ENDIF
         ENDDO                     
         SELECT * FROM E20 WHERE !EMPTY(F15) AND F05 =ANY(SELECT F01 FROM B01) AND F04 NOT IN(SELECT F01 FROM E14 WHERE F08<>.T.);  
         AND LEFT(F04,2)<>'FE' ORDER BY F02,F05 INTO CURSOR E20T3 NOWAIT   &&無耗料者且非整組入貨視為進貨費用
         H=0         
         SELECT E20T3
         GO TOP 
         DO WHILE !EOF()
            H=H+1
            K=F02
            SELECT TRNAC
            APPEND BLANK
            REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
            REPLACE F02 WITH E20T3.F02+'D'
            REPLACE F03 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(E20T3.F01),4))-1911))+RIGHT(DTOS(E20T3.F01),4),8,' ')
            REPLACE F04 WITH IIF(LEFT(E20T3.F02,2)='BM','10',IIF(E20T3.F21='03','12','11'))
            IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
               REPLACE F05 WITH 'A00'               
            ENDIF                              
            REPLACE F07 WITH E20T3.F03
            REPLACE F09 WITH '00'
            REPLACE F10 WITH '00'             
            REPLACE F12 WITH IIF(E20T3.F13=INT_011,SPACE(3),IIF(E20T3.F13='US ','USD',E20T3.F13))
            REPLACE F13 WITH IIF(E20T3.F13=INT_011,'1'+SPACE(7),PADL(ALLTRIM(TRANSFORM(E20T3.F14,'@R ###.####')),8,'0'))
            REPLACE F16 WITH E20T3.F04
            REPLACE F17 WITH RIGHT(E20T3.F19,1)
            REPLACE F18 WITH E20T3.F20
            REPLACE F20 WITH 'Y'
            REPLACE F21 WITH IIF(LEFT(E20T3.F02,2)='BM','10',IIF(E20T3.F21='03','12','11'))              
            REPLACE F24 WITH 'N'
            REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
            REPLACE F33 WITH IIF(SEEK(E20T3.F05,'B01'),B01.F99,'')
            REPLACE F34 WITH '0'
            REPLACE F35 WITH PADL('0',12,'0')

               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(ROUND(E20T3.F07,6),'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(ROUND(ABS(E20T3.F06*E20T3.F07),2),'@R ##########.##')),13,'0')                        
       
            REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(E20T3.F07*E20T3.F14,'@R #######.######')),15,'0')
            REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(ABS(E20T3.F08),'@R #############')),13,'0')             
            REPLACE F43 WITH IIF(E20T3.F19='02','Y','N')
            REPLACE F44 WITH E20T3.F03
            SELECT E20T3
            SKIP
            IF F02<>K
               H=0
            ENDIF
         ENDDO                     

        
         SELECT G07           &&內購其他進貨費用
            H=0
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH G07.F01+'F'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+IIF(SEEK(G07.F01,'B02'),B02.F02,OMT3),8,' ')  
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH IIF(SEEK(G07.F01,'B02'),B02.F07,'VZ001')
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'             
               REPLACE F13 WITH '1'+SPACE(7)
               REPLACE F16 WITH G07.F10
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(G07.F05,'B01'),B01.F99,'')
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL('0',12,'0')
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(G07.F09,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(G07.F09,'@R #############')),13,'0')       
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(G07.F09,'@R #######.######')),15,'0') 
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(G07.F09,'@R #############')),13,'0')        
               REPLACE F43 WITH 'N'
               REPLACE F44 WITH F07              
               SELECT G07
               SKIP
               IF F01<>K
                  H=0
               ENDIF   
            ENDDO
        
         SELECT  G09                 &&委外進貨費用        
            H=0
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH G09.F01+'F'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+IIF(SEEK(G09.F01,'B07'),B07.F02,OMT3),8,' ')  
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH IIF(SEEK(G09.F01,'B07'),B07.F07,'VZ001')
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'       
               REPLACE F13 WITH '1'+SPACE(7)      
               REPLACE F16 WITH IIF(SEEK(G09.F01,'B07'),B07.F15,' ')
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(G09.F05,'B01'),B01.F99,'')
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL('0',12,'0')
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(G09.F09,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(G09.F09,'@R #############')),13,'0')      
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(G09.F09,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(G09.F09,'@R #############')),13,'0')                        
               REPLACE F43 WITH 'N'
               REPLACE F44 WITH '0018'              
               SELECT G09
               SKIP
               IF F01<>K
                  H=0
               ENDIF   
            ENDDO
          SELECT  P17               &&外購進貨費用     
          SELECT * FROM P17 INTO CURSOR P17T ORDER BY F01,F05 NOWAIT   
          IF _TALLY>0
            H=0
            SELECT P17T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH 'PB'+SUBSTR(DTOS(CTOD((UKEF+'/01'))),3,4)+P17T.F01+'F'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+OMT3,8,' ')  
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH IIF(SEEK(P17T.F11,'P03'),P03.F03,'VZ001')
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'             
               REPLACE F13 WITH '1'+SPACE(7)
               REPLACE F16 WITH P17T.F11
               REPLACE F17 WITH '3'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(P17T.F05,'B01'),B01.F99,'')
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL('0',12,'0')
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(P17T.F09,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(P17T.F09,'@R #############')),13,'0')      
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(P17T.F09,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(P17T.F09,'@R #############')),13,'0')                       
               REPLACE F43 WITH 'N'
               REPLACE F44 WITH F07              
               SELECT P17T
               SKIP
               IF F01<>K
                  H=0
               ENDIF   
            ENDDO            
         ENDIF   
         SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ')  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13='*' ) INTO CURSOR B06T1 ORDER BY F01,F03 NOWAIT &&由非成本部門轉入成本部門視為樣入
         IF _TALLY>0
            H=0
            SELECT B06T1
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B06T1.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B06T1.F02,8,' ')              
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH 'VZ001'
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'
               REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'D01'),D01.F39,'')
               REPLACE F13 WITH '1'+SPACE(7)
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B06T1.F79,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B06T1.F04),'@R ########.###')),12,'0') 
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(0,'@R ##########.##')),13,'0')                                
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #############')),13,'0')             
               REPLACE F43 WITH 'N'
*!*	               REPLACE F44 WITH F07                  
               SELECT B06T1
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF
         SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*')  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ') INTO CURSOR B06T2 ORDER BY F01,F03 NOWAIT  &&由成本部門轉入非成本部門視為樣出   
         IF _TALLY>0
            H=0
            SELECT B06T2
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B06T2.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B06T2.F02,8,' ')               
               REPLACE F04 WITH '70'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH IIF(EMPTY(ALLTRIM(B06T2.F06)),'C6666',B06T2.F06)
               REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'C01'),C01.F33,'')
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B06T2.F79,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B06T2.F04),'@R ########.###')),12,'0')                    
               SELECT B06T2
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO         
         ENDIF      
         IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳
            SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*' AND LEN(ALLTRIM(F06))<>3)  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13='*' AND LEN(ALLTRIM(F06))=3 ) INTO CURSOR B06T3 ORDER BY F01,F03 NOWAIT  &&由倉庫轉入海外客戶HUB倉改成撥出入單 
            IF _TALLY>0
               FOR I=1 TO 2
                   H=0
                   SELECT B06T3
                   GO TOP
                   DO WHILE !EOF()
                      H=H+1
                      K=F01
                      SELECT TRNAC
                      APPEND BLANK  &&先做撥出再做撥入
                      REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
                      REPLACE F02 WITH B06T3.F01+IIF(I=1,'A','B')
                      REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B06T3.F02,8,' ')               
                      REPLACE F04 WITH IIF(I=1,'01','00')                                                     
                      REPLACE F05 WITH IIF(I=1,'A00',IIF(SEEK(B06T3.F10,'A14'),A14.F06,'A00'))                                                                                                   
                      REPLACE F20 WITH 'Y'
                      REPLACE F24 WITH 'N'
                      REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
                      REPLACE F33 WITH IIF(SEEK(B06T3.F79,'B01'),B01.F99,'')            
                      REPLACE F34 WITH '0'
                      REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B06T3.F04),'@R ########.###')),12,'0')      
                      SELECT B06T3
                      SKIP
                      IF F01<>K
                         H=0
                      ENDIF
                   ENDDO         
               ENDFOR    
            ENDIF      
            SELECT * FROM B06 WHERE F04>0 AND F07='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*' AND LEN(ALLTRIM(F06))=3)  AND F10 =ANY(SELECT F01 FROM A14 WHERE F13='*' AND LEN(ALLTRIM(F06))<>3 ) INTO CURSOR B06T4 ORDER BY F01,F03 NOWAIT  &&由海外客戶HUB倉轉入倉庫亦改成撥出入單 
            IF _TALLY>0
               FOR I=1 TO 2
                   H=0
                   SELECT B06T4
                   GO TOP
                   DO WHILE !EOF()
                      H=H+1
                      K=F01
                      SELECT TRNAC
                      APPEND BLANK  &&先做撥出
                      REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
                      REPLACE F02 WITH B06T4.F01+IIF(I=1,'A','B')
                      REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B06T4.F02,8,' ')               
                      REPLACE F04 WITH IIF(I=1,'01','00')                                                     
                      REPLACE F05 WITH IIF(I=1,IIF(SEEK(B06T4.F05,'A14'),A14.F06,'A00'),'A00')                                                                                                   
                      REPLACE F20 WITH 'Y'
                      REPLACE F24 WITH 'N'
                      REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
                      REPLACE F33 WITH IIF(SEEK(B06T4.F03,'B01'),B01.F99,'')            
                      REPLACE F34 WITH '0'
                      REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B06T4.F04),'@R ########.###')),12,'0')                            
                      SELECT B06T4
                      SKIP
                      IF F01<>K
                         H=0
                      ENDIF
                   ENDDO
               ENDFOR             
            ENDIF                  
            
         ENDIF
         SELECT B07   &&有耗料者視為入庫單
         SELECT * FROM B07 WHERE F10='2' AND F13='1' AND F09 =ANY(SELECT F01 FROM E14 WHERE F08<>.T.) INTO CURSOR B07T1 ORDER BY F01,F03 NOWAIT
         IF _TALLY>0
            H=0
            SELECT B07T1
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B07T1.F01+'B'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B07T1.F02,8,' ')             
               REPLACE F04 WITH '40'               
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B07T1.F03,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B07T1.F04),'@R ########.###')),12,'0')                    
               SELECT B07T1
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF                           
         SELECT B07   &&無耗料且單價為零者者視為免費進貨
         SELECT * FROM B07 WHERE F10='2' AND F13='1' AND F09 NOT IN (SELECT F01 FROM E14 WHERE F08<>.T.) AND F17=0 INTO CURSOR B07T2 ORDER BY F01,F03 NOWAIT
         IF _TALLY>0
            H=0
            SELECT B07T2
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B07T2.F01+'C'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B07T2.F02,8,' ')             
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH B07T2.F07
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'
               REPLACE F13 WITH '1'+SPACE(7)
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B07T2.F03,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B07T2.F04),'@R ########.###')),12,'0')      
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(0,'@R ##########.##')),13,'0')                                
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #############')),13,'0')             
               REPLACE F43 WITH 'N'
*!*	               REPLACE F44 WITH F07                                               
               SELECT B07T2
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF                                    
         SELECT * FROM B08 WHERE F12>0 AND F13='2'  AND F07 =ANY (SELECT F01 FROM A14 WHERE F13=SPACE(1) AND F01<>'I00C ') AND F04 =ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B08T ORDER BY F01,F08 NOWAIT &&從非成本部門領料視為樣入
         IF _TALLY>0
            H=0
            SELECT B08T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B08T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B08T.F02,8,' ')              
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH 'VZ001'
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'
               REPLACE F13 WITH '1'+SPACE(7)
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B08T.F08,'B01'),B01.F99,'')           
               REPLACE F34 WITH '0' 
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B08T.F12),'@R ########.###')),12,'0')              
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(0,'@R ##########.##')),13,'0')                                
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #############')),13,'0')             
               REPLACE F43 WITH 'N'
*!*	               REPLACE F44 WITH F07                        
               SELECT B08T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO         
         ENDIF        
         SELECT * FROM B09 WHERE F04>0 AND F11='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B09T1 ORDER BY F01,F03 NOWAIT  &&盤盈
         IF _TALLY>0
            H=0
            SELECT B09T1
            GO TOP 
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B09T1.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B09T1.F02,8,' ')
               REPLACE F04 WITH '80'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                                                       
                  IF LEN(ALLTRIM(B09T1.F05))=5 AND LEFT(B09T1.F05,1)='S'                     
                     REPLACE F05 WITH IIF(SEEK(B09T1.F05,'A14'),IIF(EMPTY(A14.F06),'A00',A14.F06),'A00')
                  ELSE                                                   
                     REPLACE F05 WITH 'A00'          
                  ENDIF        
               ENDIF                                 
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B09T1.F03,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B09T1.F04),'@R #######.####')),12,'0')     
               SELECT B09T1
               SKIP
               IF F01<>K
                  H=0
               ENDIF                   
            ENDDO
         ENDIF   
         SELECT * FROM B09 WHERE F04<0 AND F11='2' AND F05 =ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B09T2 ORDER BY F01,F03 NOWAIT &&盤虧
         IF _TALLY>0
            H=0
            SELECT B09T2
            GO TOP 
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B09T2.F01+'A'
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B09T2.F02,8,' ')
               REPLACE F04 WITH '81'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                                                       
                  IF LEN(ALLTRIM(B09T2.F05))=5 AND LEFT(B09T2.F05,1)='S'                     
                     REPLACE F05 WITH IIF(SEEK(B09T2.F05,'A14'),IIF(EMPTY(A14.F06),'A00',A14.F06),'A00')
                  ELSE                                                   
                     REPLACE F05 WITH 'A00'          
                  ENDIF        
               ENDIF                                              
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B09T2.F03,'B01'),B01.F99,'')    
               REPLACE F34 WITH '0'        
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B09T2.F04),'@R ########.###')),12,'0')     
               SELECT B09T2
               SKIP
               IF F01<>K
                  H=0
               ENDIF                   
            ENDDO         
         ENDIF    
         SELECT B10          &&報廢
         SELECT * FROM B10 WHERE F09='2' AND F05 = ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B10T ORDER BY F01,F03 NOWAIT
         IF _TALLY>0
            H=0
            SELECT B10T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B10T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B10T.F02,8,' ')              
               REPLACE F04 WITH '96'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B10T.F03,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B10T.F04),'@R ########.###')),12,'0')                    
               SELECT B10T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF
         SELECT B40  &&領用
         SELECT * FROM B40 WHERE F07='N' AND F11='2' AND F05 = ANY (SELECT F01 FROM A14 WHERE F13='*') ORDER BY F01,F03 INTO CURSOR B40T NOWAIT
         IF _TALLY>0
            H=0
            SELECT B40T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B40T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B40T.F02,8,' ')   
               REPLACE F04 WITH '50'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')       
               REPLACE F33 WITH IIF(SEEK(B40T.F03,'B01'),B01.F99,'')          
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B40T.F04),'@R ########.###')),12,'0')                    
               SELECT B40T
               SKIP
               IF F01<>K
                  H=0
               ENDIF               
            ENDDO
         ENDIF
*************************************借出入單先暫時不管         
*!*	         H=0
*!*	         SELECT B61
*!*	         GO TOP
*!*	         DO WHILE !EOF()
*!*	           H=H+1
*!*	           K=F01
*!*	           SELECT TRNAC
*!*	           APPEND BLANK
*!*	           REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
*!*	           REPLACE F02 WITH B61.F01
*!*	           REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B61.F02,8,' ') 
*!*	           REPLACE F04 WITH '58'
*!*	           REPLACE F07 WITH B61.F10
*!*	           REPLACE F20 WITH 'Y'
*!*	           REPLACE F24 WITH 'N'
*!*	           REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')       
*!*	           REPLACE F33 WITH IIF(SEEK(B61.F03,'B01'),B01.F99,'')          
*!*	           REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B61.F04),'@R ########.###')),12,'0')                    
*!*	               SELECT B61
*!*	               SKIP
*!*	               IF F01<>K
*!*	                  H=0
*!*	               ENDIF                
*!*	         ENDDO
*!*	         H=0
*!*	         SELECT B62
*!*	         GO TOP
*!*	         DO WHILE !EOF()
*!*	           H=H+1
*!*	           K=F01
*!*	           SELECT TRNAC
*!*	           APPEND BLANK
*!*	           REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
*!*	           REPLACE F02 WITH B62.F01
*!*	           REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B62.F02,8,' ')   
*!*	           REPLACE F04 WITH '59'
*!*	           REPLACE F07 WITH B62.F10
*!*	           REPLACE F20 WITH 'Y'
*!*	           REPLACE F24 WITH 'N'
*!*	           REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')       
*!*	           REPLACE F33 WITH IIF(SEEK(B62.F03,'B01'),B01.F99,'')          
*!*	           REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B62.F04),'@R ########.###')),12,'0')                    
*!*	               SELECT B62
*!*	               SKIP
*!*	               IF F01<>K
*!*	                  H=0
*!*	               ENDIF                
*!*	         ENDDO         
*!*	         H=0
*!*	         SELECT B63
*!*	         GO TOP
*!*	         DO WHILE !EOF()
*!*	           H=H+1
*!*	           K=F01
*!*	           SELECT TRNAC
*!*	           APPEND BLANK
*!*	           REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
*!*	           REPLACE F02 WITH B63.F01
*!*	           REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B63.F02,8,' ')   
*!*	           REPLACE F04 WITH '55'
*!*	           REPLACE F07 WITH B63.F10
*!*	           REPLACE F20 WITH 'Y'
*!*	           REPLACE F24 WITH 'N'
*!*	           REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')       
*!*	           REPLACE F33 WITH IIF(SEEK(B63.F03,'B01'),B01.F99,'')          
*!*	           REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B63.F04),'@R ########.###')),12,'0')                    
*!*	               SELECT B63
*!*	               SKIP
*!*	               IF F01<>K
*!*	                  H=0
*!*	               ENDIF                
*!*	         ENDDO                  
*!*	         H=0
*!*	         SELECT B64
*!*	         GO TOP
*!*	         DO WHILE !EOF()
*!*	           H=H+1
*!*	           K=F01
*!*	           SELECT TRNAC
*!*	           APPEND BLANK
*!*	           REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
*!*	           REPLACE F02 WITH B64.F01
*!*	           REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B64.F02,8,' ')  
*!*	           REPLACE F04 WITH '56'
*!*	           REPLACE F07 WITH B64.F10
*!*	           REPLACE F20 WITH 'Y'
*!*	           REPLACE F24 WITH 'N'
*!*	           REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')       
*!*	           REPLACE F33 WITH IIF(SEEK(B64.F03,'B01'),B01.F99,'')          
*!*	           REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B64.F04),'@R ########.###')),12,'0')                    
*!*	               SELECT B64
*!*	               SKIP
*!*	               IF F01<>K
*!*	                  H=0
*!*	               ENDIF                               
*!*	         ENDDO             
************************************         
         SELECT C38  &&樣出
         SELECT * FROM C38 WHERE F07='2' AND F05 = ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR C38T ORDER BY F01,F79 NOWAIT
         IF _TALLY>0
            H=0
            SELECT C38T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH C38T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+C38T.F02,8,' ')               
               REPLACE F04 WITH '70'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                                 
               REPLACE F07 WITH C38T.F10
               REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'C01'),C01.F33,'')
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(C38T.F79,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(C38T.F04),'@R ########.###')),12,'0')                    
               SELECT C38T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF              
         SELECT D39    &&樣入改成進貨
         SELECT * FROM D39 WHERE F07='2' AND F10 = ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR D39T ORDER BY F01,F03 NOWAIT
         IF _TALLY>0
            H=0
            SELECT D39T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH D39T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+D39T.F02,8,' ')              
               REPLACE F04 WITH '10'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                  
               REPLACE F07 WITH D39T.F05
               REPLACE F09 WITH '00'
               REPLACE F10 WITH '00'
               REPLACE F11 WITH IIF(SEEK(ALLTRIM(F07),'D01'),D01.F39,'')
               REPLACE F13 WITH '1'+SPACE(7)               
               REPLACE F17 WITH '6'
               REPLACE F20 WITH 'Y'
               REPLACE F21 WITH '10'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(D39T.F03,'B01'),B01.F99,'')            
               REPLACE F34 WITH '0'
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(D39T.F04),'@R ########.###')),12,'0')            
               REPLACE F38 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F39 WITH PADL(ALLTRIM(TRANSFORM(0,'@R ##########.##')),13,'0')                                
               REPLACE F36 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #######.######')),15,'0')
               REPLACE F37 WITH PADL(ALLTRIM(TRANSFORM(0,'@R #############')),13,'0')             
               REPLACE F43 WITH 'N'
*!*	               REPLACE F44 WITH F07                                       
               SELECT D39T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF              
         SELECT B55
         SELECT * FROM B55 WHERE  F13='2'  INTO CURSOR B55T ORDER BY F01,F08 NOWAIT
         IF _TALLY>0
            H=0
            SELECT B55T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B55T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B55T.F02,8,' ')         
               REPLACE F04 WITH '40'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                  
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B55T.F08,'B01'),B01.F99,'')          
               REPLACE F34 WITH '0'  
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B55T.F12),'@R ########.###')),12,'0')                    
               SELECT B55T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF              
         SELECT B56
         SELECT * FROM B56 WHERE F10='2' AND F13='1' AND EMPTY(ALLTRIM(F07)) INTO CURSOR B56T ORDER BY F01,F03 NOWAIT  &&只抓廠內自製單號
         IF _TALLY>0
            H=0
            SELECT B56T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B56T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B56T.F02,8,' ')              
               REPLACE F04 WITH '41'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                   REPLACE F05 WITH 'A00'               
               ENDIF                  
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B56T.F03,'B01'),B01.F99,'')         
               REPLACE F34 WITH '0'   
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B56T.F04),'@R ########.###')),12,'0')                    
               SELECT B56T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF       
         SELECT B16
         SELECT * FROM B16 WHERE F13='2' AND F07 = ANY (SELECT F01 FROM A14 WHERE F13='*') INTO CURSOR B16T ORDER BY F01,F08 NOWAIT
         IF _TALLY>0
            H=0
            SELECT B16T
            GO TOP
            DO WHILE !EOF()
               H=H+1
               K=F01
               SELECT TRNAC
               APPEND BLANK
               REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
               REPLACE F02 WITH B16T.F01
               REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+B16T.F02,8,' ')   
               REPLACE F19 WITH IIF(SEEK(B16T.F09,'B01'),B01.F99,'')                  
               REPLACE F04 WITH '30'
               IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                  REPLACE F05 WITH 'A00'               
               ENDIF                  
               REPLACE F20 WITH 'Y'
               REPLACE F24 WITH 'N'
               REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
               REPLACE F33 WITH IIF(SEEK(B16T.F08,'B01'),B01.F99,'')         
               REPLACE F34 WITH '0'   
               REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(B16T.F12),'@R ########.###')),12,'0')           
               REPLACE F40 WITH PADL(ALLTRIM(TRANSFORM(0,'@R ###############')),15,'0')        
               SELECT B16T
               SKIP
               IF F01<>K
                  H=0
               ENDIF
            ENDDO
         ENDIF                 
         SELECT B55
         SET ORDER TO 2
         SELECT B07
         SET ORDER TO 2         
         SELECT F01,SUM(F05/F99) FROM E13 WHERE F08=.T. GROUP BY F01 INTO CURSOR E13T ORDER BY F01 NOWAIT        
         SELECT F01,F02,F03,SUM(F05),F07 FROM E14 WHERE F08<>.T. GROUP BY F01,F02 INTO CURSOR E14T ORDER BY F01,F02 NOWAIT            
         IF _TALLY>0                     
             SELECT E14T.*,E13T.SUM_EXP_2 FROM E14T,E13T WHERE E13T.f01=E14T.f01 INTO CURSOR E14T1 ORDER BY E14T.F01,E14T.F02 NOWAIT
             IF _TALLY>0
                 H=0
                GO TOP
                DO WHILE !EOF()
                   H=H+1
                   K=E14T1.F01
                   SELECT TRNAC
                   APPEND BLANK
                   REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
                   REPLACE F02 WITH SUBSTR(DTOS(CTOD((UKEF+'/01'))),3,2)+CHR(IIF(VAL(RIGHT(UKEF,2))<10,48,55)+VAL(RIGHT(UKEF,2)))+E14T1.F01
                   REPLACE F03 WITH PADR(ALLTRIM(F01)+RIGHT(UKEF,2)+E14T1.F03,8,' ')                                
                   REPLACE F19 WITH IIF(SEEK(IIF(SEEK(E14T1.F01,'B55'),B55.F08,IIF(SEEK(E14T1.F01,'B07'),B07.F03,'')),'B01'),B01.F99,'')                                
                   REPLACE F04 WITH '30'
                   IF THISFORM.SIDE_OPTION.VALUE=2  &&如果是外帳增設一個虛擬部門編號                                       
                      REPLACE F05 WITH 'A00'               
                   ENDIF   
                   REPLACE F20 WITH 'Y'
                   REPLACE F24 WITH 'N'
                   REPLACE F32 WITH PADL(ALLTRIM(STR(H)),4,'0')
                   REPLACE F33 WITH IIF(SEEK(E14T1.F02,'B01'),B01.F99,'')       
                   REPLACE F34 WITH '0'     
                   REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(E14T1.SUM_F05),'@R ########.###')),12,'0')                 
                   REPLACE F40 WITH PADL(ALLTRIM(TRANSFORM(E14T1.SUM_EXP_2,'@R ###############')),15,'0')                     
                   SELECT E14T1
                   SKIP
                   IF F01<>K
                      H=0
                   ENDIF
                ENDDO
            ENDIF    
         ENDIF             
      ENDIF  &&進銷項轉檔判斷完畢   
          SELECT TRNAC

          FILE_NAME=IIF(THISFORM.SIDE_OPTION.VALUE=1,IIF(INT_107='20938737','A_TOKUTSU','A_'+INT_107),'A_'+INT_107)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'013')  && A_TOKUTSU.013 進銷累積檔
          IF !EMPTY(GCDELIMFILE)  
                COPY TO (GCDELIMFILE) TYPE  SDF                
                =MESSAGEBOX('儲存成功',0+48,'提示訊息視窗')
          ENDIF              

*!*	          CREATE CURSOR ACTT;
*!*	          (F01 C(15),F02 C(8),F03 C(2),F04 C(10),F05 C(11),F06 C(8),F07 C(15),F08 C(1),F09 C(15),F10 C(15),F11 C(15),F12 C(15),F13 C(15),F14 C(15),;
*!*	          F15 C(1),F16 C(10),F17 C(10),F18 C(4),F19 C(15),F20 C(1),F21 C(15),F22 C(4),F23 C(30),F25 C(1),F26 C(20),F27 C(12),F28 C(10),F29 C(10),F30 C(1),F31 C(12),;
*!*	          F32 C(15),F33 C(3),F34 C(8),F35 C(15),F36 C(15),F37 C(15))
*!*	          SELECT F01,F02,F03,SUM(F08),F13,F14,F19,F20,F21 FROM E20 WHERE F05 =ANY(SELECT F01 FROM B01) AND F04 NOT IN (SELECT F01 FROM E14) INTO CURSOR E20T GROUP BY F02 ORDER BY F02 NOWAIT
*!*	          SELECT E20T
*!*	          GO TOP
*!*	          DO WHILE !EOF()
*!*	             SELECT ACTT   &&帳款檔
*!*	                APPEND BLANK 
*!*	                REPLACE F01 WITH E20T.F02
*!*	                REPLACE F02 WITH PADR(ALLTRIM(STR(VAL(LEFT(DTOS(E20T.F01),4))-1911))+RIGHT(DTOS(E20T.F01),4),8,' ')
*!*	                REPLACE F03 WITH IIF(LEFT(E20T.F02,2)='BM','10',IIF(E20T.F21='03','12','11'))
*!*	                REPLACE F04 WITH E20T.F03
*!*	                REPLACE F08 WITH RIGHT(E20T.F19,1)
*!*	                IF F03<>'12'
*!*	                   REPLACE F09 WITH PADL(ALLTRIM(TRANSFORM(E20T.SUM_F08,'@R ############.##')),15,'0')
*!*	                   REPLACE F10 WITH PADL(ALLTRIM(TRANSFORM(ROUND(E20T.SUM_F08*(1+IIF(F08='3',INT_002/100,0)),0),'@R ############.##')),15,'0')
*!*	                ELSE
*!*	                   REPLACE F12 WITH PADL(ALLTRIM(TRANSFORM(ABS(ABS(E20T.SUM_F08)),'@R ############.##')),15,'0')
*!*	                ENDIF                       
*!*	                REPLACE F15 WITH 'N'
*!*	                REPLACE F16 WITH E20T.F20
*!*	                REPLACE F17 WITH E20T.F20
*!*	                REPLACE F25 WITH '2'
*!*	                REPLACE F26 WITH DTOS(DATE())
*!*	                REPLACE F32 WITH E20T.F02
*!*	                IF ALLTRIM(INT_011)<>ALLTRIM(E20T.F13)
*!*	                   REPLACE F33 WITH IIF(ALLTRIM(E20T.F13)='US','USD',E20T.F13)           
*!*	                   REPLACE F34 WITH PADL(ALLTRIM(TRANSFORM(E20T.F14,'@R ###.####')),8,'0')       
*!*	                   REPLACE F35 WITH PADL(ALLTRIM(TRANSFORM(ABS(ROUND(E20T.SUM_F08/E20T.F14,2)),'@R ############.##')),15,'0')
*!*	                ENDIF
*!*	             SELECT E20T
*!*	             SKIP
*!*	          ENDDO
*!*	          SELECT ACTT
*!*	          FILE_NAME='A_TOKUTSU'  &&+LEFT(DTOS(CTOD(UKEF+'/01')),6)
*!*	          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'025')  && A_TOKUTSU.025
*!*	          IF !EMPTY(GCDELIMFILE)  
*!*	                COPY TO (GCDELIMFILE) TYPE SDF
*!*	                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
*!*	          ENDIF                
          CREATE CURSOR CVCD;  &&客戶供應商檔
          (F01 C(10),F02 C(1),F03 C(12),F04 C(60),F05 C(12),F06 C(4),F07 C(8),F08 C(9),F09 C(5),F10 C(60),F11 C(60),F12 C(60),F13 C(20),F14 C(20),F15 C(20),;
          F16 C(20),F17 C(20),F18 C(20),F19 C(20),F20 C(30),F21 C(12),F22 C(12),F23 C(30),F25 C(18),F26 C(1),F27 C(5),F28 C(15),F29 C(10),F30 C(11),;
          F31 C(11),F32 C(8),F33 C(8),F34 C(15),F35 C(15),F36 C(15),F37 C(15),F38 C(1),F39 C(1),F40 C(30),F41 C(30),F42 C(1),F43 C(2),F44 C(2),F45 C(2),F46 C(2),;
          F47 C(2),F48 C(2),F49 C(2),F50 C(2),F51 C(2),F52 C(2),F53 C(2),F54 C(2),F55 C(5),F56 C(5),F57 C(20),F58 C(10),F59 C(10))
          SELECT C01
          GO TOP
          DO WHILE !EOF()
             SELECT CVCD
             APPEND BLANK
             REPLACE F01 WITH C01.F01
             REPLACE F02 WITH '2'
             REPLACE F03 WITH ALLTRIM(STUFF(C01.F05,AT('?',C01.F05),1,IIF(AT('?',C01.F05)>0,'',LEFT(C01.F05,1))))  &&STUFF(C01.F05,AT('?',C01.F05),1,IIF(AT('?',C01.F05)>0,'',LEFT(C01.F05,1)))
             REPLACE F04 WITH ALLTRIM(STUFF(C01.F04,AT('?',C01.F04),1,IIF(AT('?',C01.F04)>0,'',LEFT(C01.F04,1))))
*!*	             REPLACE F05 WITH C01.F21
             REPLACE F07 WITH C01.F10
             REPLACE F10 WITH ALLTRIM(STUFF(C01.F06,AT('?',C01.F06),1,IIF(AT('?',C01.F06)>0,'',LEFT(C01.F06,1))))
             REPLACE F11 WITH ALLTRIM(STUFF(C01.F06,AT('?',C01.F06),1,IIF(AT('?',C01.F06)>0,'',LEFT(C01.F06,1))))
             REPLACE F12 WITH ALLTRIM(STUFF(C01.F07,AT('?',C01.F07),1,IIF(AT('?',C01.F07)>0,'',LEFT(C01.F07,1))))
             REPLACE F13 WITH C01.F13
             REPLACE F14 WITH C01.F13
             REPLACE F15 WITH C01.F13
             REPLACE F16 WITH C01.F14
             REPLACE F21 WITH ALLTRIM(ALLTRIM(STUFF(C01.F11,AT('?',C01.F11),1,IIF(AT('?',C01.F11)>0,'',LEFT(C01.F11,1)))))
*!*	             REPLACE F22 WITH ALLTRIM(ALLTRIM(STUFF(C01.F12,AT('?',C01.F12),1,IIF(AT('?',C01.F12)>0,'',LEFT(C01.F12,1)))))
*!*	             REPLACE F26 WITH SPACE(1)
*!*	             REPLACE F27 WITH SPACE(5)
             REPLACE F29 WITH SPACE(10)
             REPLACE F30 WITH ALLTRIM(C01.F33)
             REPLACE F31 WITH ALLTRIM(C01.F23)
             REPLACE F38 WITH IIF(C01.F29='31','3','2')
             REPLACE F39 WITH ICASE(C01.F15='0','1',C01.F15='3','3','2')
             REPLACE F42 WITH '2'
             REPLACE F43 WITH IIF(C01.F15='2',' 2',' 1')
             REPLACE F44 WITH C01.F17
             REPLACE F59 WITH C01.F01
             SELECT C01 
             SKIP
          ENDDO
          SELECT D01
          GO TOP
          DO WHILE !EOF()
             SELECT CVCD
             APPEND BLANK
             REPLACE F01 WITH D01.F01
             REPLACE F02 WITH '1'
             REPLACE F03 WITH ALLTRIM(ALLTRIM(STUFF(D01.F04,AT('?',D01.F04),1,IIF(AT('?',D01.F04)>0,'',LEFT(D01.F04,1)))))
             REPLACE F04 WITH ALLTRIM(ALLTRIM(STUFF(D01.F03,AT('?',D01.F03),1,IIF(AT('?',D01.F03)>0,'',LEFT(D01.F03,1)))))
             REPLACE F07 WITH D01.F06
             REPLACE F10 WITH ALLTRIM(ALLTRIM(STUFF(D01.F05,AT('?',D01.F05),1,IIF(AT('?',D01.F05)>0,'',LEFT(D01.F05,1)))))
             REPLACE F11 WITH ALLTRIM(ALLTRIM(STUFF(D01.F21,AT('?',D01.F21),1,IIF(AT('?',D01.F21)>0,'',LEFT(D01.F21,1)))))
             REPLACE F12 WITH ALLTRIM(ALLTRIM(STUFF(D01.F05,AT('?',D01.F05),1,IIF(AT('?',D01.F05)>0,'',LEFT(D01.F05,1)))))
             REPLACE F13 WITH D01.F09
             REPLACE F14 WITH D01.F09
             REPLACE F15 WITH D01.F09
             REPLACE F16 WITH D01.F10
             REPLACE F21 WITH ALLTRIM(ALLTRIM(STUFF(D01.F07,AT('?',D01.F07),1,IIF(AT('?',D01.F07)>0,'',LEFT(D01.F07,1)))))
             REPLACE F22 WITH ALLTRIM(ALLTRIM(STUFF(D01.F08,AT('?',D01.F08),1,IIF(AT('?',D01.F08)>0,'',LEFT(D01.F08,1)))))
             REPLACE F26 WITH D01.F12
             REPLACE F31 WITH D01.F39
             REPLACE F38 WITH '3'
             REPLACE F39 WITH '2'
             REPLACE F42 WITH '2'
             REPLACE F43 WITH ' 1'
             REPLACE F44 WITH D01.F15
             REPLACE F59 WITH D01.F01
             SELECT D01
             SKIP
          ENDDO
          SELECT E02
          GO TOP
          DO WHILE !EOF()
             SELECT CVCD
             APPEND BLANK
             REPLACE F01 WITH E02.F01
             REPLACE F02 WITH '1'
             REPLACE F03 WITH ALLTRIM(ALLTRIM(STUFF(E02.F04,AT('?',E02.F04),1,IIF(AT('?',E02.F04)>0,'',LEFT(E02.F04,1)))))
             REPLACE F04 WITH ALLTRIM(ALLTRIM(STUFF(E02.F03,AT('?',E02.F03),1,IIF(AT('?',E02.F03)>0,'',LEFT(E02.F03,1)))))
             REPLACE F07 WITH E02.F13
             REPLACE F10 WITH ALLTRIM(ALLTRIM(STUFF(E02.F05,AT('?',E02.F05),1,IIF(AT('?',E02.F05)>0,'',LEFT(E02.F05,1)))))
             REPLACE F11 WITH ALLTRIM(ALLTRIM(STUFF(E02.F05,AT('?',E02.F05),1,IIF(AT('?',E02.F05)>0,'',LEFT(E02.F05,1)))))
             REPLACE F12 WITH ALLTRIM(ALLTRIM(STUFF(E02.F05,AT('?',E02.F05),1,IIF(AT('?',E02.F05)>0,'',LEFT(E02.F05,1)))))
             REPLACE F13 WITH E02.F06
             REPLACE F14 WITH E02.F06
             REPLACE F15 WITH E02.F06
             REPLACE F16 WITH E02.F07
             REPLACE F22 WITH ALLTRIM(ALLTRIM(STUFF(E02.F08,AT('?',E02.F08),1,IIF(AT('?',E02.F08)>0,'',LEFT(E02.F08,1)))))
             REPLACE F31 WITH E02.F39
             REPLACE F38 WITH '3'
             REPLACE F39 WITH '2'
             REPLACE F42 WITH '2'
             REPLACE F43 WITH '1'
             REPLACE F44 WITH E02.F11
             REPLACE F59 WITH E02.F01
             SELECT E02
             SKIP
          ENDDO
          SELECT CVCD
          FILE_NAME=IIF(THISFORM.SIDE_OPTION.VALUE=1,IIF(INT_107='20938737','TOKUTSU',INT_107),INT_107)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'011')    && TOKUTSU.011
          IF !EMPTY(GCDELIMFILE)               
                COPY TO (GCDELIMFILE) TYPE SDF
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF                                    
          CREATE CURSOR INVT;    &&庫存檔
          (F01 C(4),F02 C(20),F03 C(1),F04 C(15),F05 C(30),F06 C(60),F07 C(1),F08 C(15),F09 C(3),F10 C(12),F11 C(15),F12 C(1),F13 C(15),F14 C(1),F15 C(15),;
          F16 C(1),F17 C(15),F18 C(1),F19 C(8),F20 C(8),F21 C(15),F22 C(1),F23 C(3),F24 C(15),F25 C(8),F26 C(15),F27 C(1),F28 C(3),F29 C(15),F30 C(6),F31 C(12),;
          F32 C(20),F33 C(30),F34 C(10),F35 C(15),F36 C(15),F37 C(15),F38 C(15),F39 C(15),F40 C(15),F41 C(15),F42 C(15),F43 C(15),F44 C(15),F45 C(15),F46 C(15),;
          F47 C(15),F48 C(30),F51 C(30))
          SELECT * FROM B01 WHERE !EMPTY(F98) AND (F99=ANY(SELE F33 FROM TRNAC) OR F99=ANY (SELECT F28 FROM TRNAC))  INTO CURSOR B01T ORDER BY F01 NOWAIT &&
          SELE B01T
          GO TOP
          DO WHILE !EOF()
             SELECT INVT
             APPEND BLANK
             REPLACE F01 WITH PADR(ALLTRIM(STR(VAL(LEFT(UKEF,4)))),4,' ')
             REPLACE F02 WITH B01T.F99
             REPLACE F03 WITH '0'
             REPLACE F04 WITH PADL(ALLTRIM(TRANSFORM(1,'@R ###############')),15,'0')       
             REPLACE F05 WITH B01T.F01
             REPLACE F06 WITH B01T.F01
             REPLACE F07 WITH IIF(B01T.F97='Y','4',ICASE(B01T.F98='* *','3',B01T.F98=' **','4',B01T.F98='*  ' ,'1',B01T.F98='A  ','2','2'))
             REPLACE F09 WITH ICASE(ALLTRIM(B01T.F04)$'PCSpcs','021',ALLTRIM(B01T.F04)='M','022','PCS')
             REPLACE F12 WITH 'N'
             REPLACE F14 WITH 'N'
             REPLACE F16 WITH 'N'
             REPLACE F18 WITH 'N'
             REPLACE F22 WITH 'N'
             REPLACE F30 WITH IIF(B01T.F97='Y','加工品','部品')
             REPLACE F34 WITH IIF(B01T.F97<>'Y' AND LEFT(B01T.F02,3)='JST',INT_117,'')
             REPLACE F48 WITH STRCONV(B01T.F29,1)
             REPLACE F51 WITH B01T.F02
             SELECT B01T
             SKIP
          ENDDO
          SELECT INVT
          FILE_NAME=IIF(THISFORM.SIDE_OPTION.VALUE=1,IIF(INT_107='20938737','A_TOKUTSU','A_'+INT_107),'A_'+INT_107)
          GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'007')
          IF !EMPTY(GCDELIMFILE)  && A_TOKUTSU.007
                COPY TO (GCDELIMFILE) TYPE SDF   
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
          ENDIF                          
            
 
        THISFORM.CMND2.CLICK                  
   ENDPROC  
   PROCEDURE CMND2.CLICK
        FLG='1'   
        THISFORM.RELEASE                   
   ENDPROC                      
ENDDEFINE