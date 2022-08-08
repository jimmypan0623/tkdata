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
SEEK sys_oper+'K19'
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
SET FILTER TO MOD(VAL(SUBSTR(DTOS(F01),5,2)),2)<>0
SET ORDER TO 1
GO BOTT
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
K19form=createobject("tkK19")
K19form.show       
define class tkK19 as form
*!*	  autocenter=.t.
  caption='K19.發票月份轉檔選擇'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*320
  width=INT_015*420
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkK19'
  ADD OBJECT MTH_LIST AS MTH_LIST WITH;
      LEFT=INT_015*160,;
      TOP=INT_015*35,;
      WIDTH=INT_015*80,;      
      HEIGHT=INT_015*30,;    
      FONTSIZE=INT_015*12  
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION='請輸入申報月份'    

  ADD OBJECT LBLF11 AS LABEL WITH;
      LEFT=INT_015*210,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''      
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
               ELSE
                   S=K-1
               ENDIF       
               SKIP
            ENDDO   
           
            .VALUE=S
       ENDWITH   
  ENDPROC       
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
         THISFORM.MTH_LIST.SETFOCUS
      ENDPROC
**********************************開始搜尋指定範圍之發票       
      PROCEDURE CMND1.CLICK      
          IF RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)='01'
             TXMTH1=ALLTRIM(STR(YEAR(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01'))-1))+'11'
             TXMTH2=ALLTRIM(STR(YEAR(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01'))-1))+'12'             
          ELSE
             TXMTH1=LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2))-2)),2,'0')
             TXMTH2=LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),4)+PADL(ALLTRIM(STR(VAL(RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2))-1)),2,'0')
          ENDIF    
          B061='B06'+TXMTH1
          B062='B06'+TXMTH2
          C131='C13'+TXMTH1
          C132='C13'+TXMTH2
          K251='K25'+TXMTH1
          K252='K25'+TXMTH2
          ************************************先作銷項          
          IF !USED('&B06')
             SELECT 0
             USE &B061 ALIAS B061
          ELSE
             SELECT B061             
          ENDIF   
          SET ORDER TO 2          
          IF !USED('&B062')
             SELECT 0
             USE &B062 ALIAS B062
          ELSE
             SELECT B062             
          ENDIF   
          SET ORDER TO 2                       
          IF !USED('&C131')
             SELECT 0
             USE &C131 ALIAS C131
          ELSE
             SELECT C131             
          ENDIF   
          SET ORDER TO 3          
          IF !USED('&C132')
             SELECT 0
             USE &C132 ALIAS C132
          ELSE
             SELECT C132             
          ENDIF   
          SET ORDER TO 3             
          IF !USED('&K251')
             SELECT 0
             USE &K251 ALIAS K251
          ELSE
             SELECT K251             
          ENDIF   
          
          SET ORDER TO 1
          IF !USED('&K252')
             SELECT 0
             USE &K252 ALIAS K252
          ELSE
             SELECT K252             
          ENDIF   
          
          SET ORDER TO 1          
          SELECT * FROM K251 UNION (SELE * FROM K252 )  INTO CURSOR K25 NOWAIT
          CALCULATE MAX(LEFT(F07,2)) FOR F01='31' AND LEFT(F07,2)<>'BC' AND SUBSTR(F07,3,4)<>SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4) AND ISUPPER(LEFT(F07,1)) AND ISUPPER(SUBSTR(F07,2,1)) TO CT1
          CALCULATE MAX(LEFT(F07,2)) FOR F01='32' AND LEFT(F07,2)<>'BC' AND SUBSTR(F07,3,4)<>SUBSTR(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),3,4) AND ISUPPER(LEFT(F07,1)) AND ISUPPER(SUBSTR(F07,2,1)) TO CT2
          CREATE CURSOR SALE1;
          (F01 C(5),F02 C(8),F03 C(5),F04 C(1),F05 C(4),F06 C(10),F07 C(8),F08 C(12),F09 C(5),F10 C(12),F11 C(4),F12 C(4),F13 C(10),F14 C(15),F15 C(10),F16 C(1),F17 C(20),F18 C(1))
          INDEX ON F06 TAG SALE1
          SET ORDER TO 1
          SELECT * FROM K25 WHERE LEFT(F01,2)='31' AND LEN(ALLTRIM(F07))=10 AND LEFT(F07,2)=CT1 AND F08>=0  ORDER BY F07 INTO CURSOR K25T1 NOWAIT
          IF _TALLY>0
*!*	             YMC=PADL(ALLTRIM(STR(MOD(VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,6,2)),2)+VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,6,2))+1)),2,'0')
*!*	             IF YMC='13'
*!*	                YC=PADL(ALLTRIM(STR(VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,1,4))+1)),3,'0')
*!*	                YMC='01'
*!*	             ELSE
*!*	                YC=SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)
*!*	             ENDIF             
             SELECT K25T1
             GO TOP
             ST_KEY=''
             IF RIGHT(F07,2)='00' OR RIGHT(F07,2)='50'
                ST_KEY=F07                
             ELSE
                IF VAL(RIGHT(F07,2))<50
                   ST_KEY=CT1+LEFT(ST_KEY,6)+'00'
                ELSE
                   ST_KEY=CT1+LEFT(ST_KEY,6)+'50'                
                ENDIF                 
                SELECT SALE1

                FOR I=VAL(RIGHT(ST_KEY,8)) TO VAL(RIGHT(K25T1.F07,8))-1
                    APPEND BLANK
                    REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                    REPLACE F02 WITH 'TOKUTSU'
                    REPLACE F03 WITH SUBSTR(K25T1.F19,2,3)+RIGHT(K25T1.F19,2)
                    REPLACE F04 WITH '3'
                    REPLACE F05 WITH '0001'
                    REPLACE F06 WITH ST_KEY
                    REPLACE F07 WITH '+'
                    REPLACE F16 WITH 'N'                    
                    ST_KEY=CT1+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')
                ENDFOR    
             ENDIF                    
             SELECT K25T1
             DO WHILE !EOF()
                IF ST_KEY<>K25T1.F07
                   SELECT SALE1 
                   IF VAL(RIGHT(K25T1.F07,8))-VAL(RIGHT(ST_KEY,8))<100
                      TB=VAL(RIGHT(K25T1.F07,8))-VAL(RIGHT(ST_KEY,8))
                   ELSE                      
                      IF VAL(RIGHT(ST_KEY,2))>50
                         TB=100-VAL(RIGHT(ST_KEY,2))
                      ELSE
                         TB=50-VAL(RIGHT(ST_KEY,2))                         
                      ENDIF
                   ENDIF                      
                      FOR I=1 TO TB
                          APPEND BLANK
                          REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                          REPLACE F02 WITH 'TOKUTSU'
                          REPLACE F03 WITH SUBSTR(K25T1.F19,2,3)+RIGHT(K25T1.F19,2)
                          REPLACE F04 WITH '3'
                          REPLACE F06 WITH ST_KEY
                          REPLACE F07 WITH '+'
                          REPLACE F16 WITH 'N'               
                          ST_KEY=CT1+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')                
                      ENDFOR                    
                ENDIF
                   SELECT SALE1
                   APPEND BLANK
                   REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                   REPLACE F02 WITH 'TOKUTSU'
                   REPLACE F03 WITH SUBSTR(K25T1.F19,2,3)+RIGHT(K25T1.F19,2)
                   REPLACE F04 WITH '3'
                   REPLACE F06 WITH K25T1.F07
                   REPLACE F07 WITH ICASE(K25T1.F16='2','!',K25T1.F16='4','-',K25T1.F04)
                   REPLACE F08 WITH PADL(ALLTRIM(TRANSFORM(K25T1.F08,'@R ############')),12,'0')
                   REPLACE F09 WITH ICASE(K25T1.F09='1','5',K25T1.F09='2','04',K25T1.F09='3','*','')
                   REPLACE F10 WITH PADL(ALLTRIM(TRANSFORM(K25T1.F10,'@R ############')),12,'0')
                   REPLACE F11 WITH RIGHT(K25T1.F19,2)+K25T1.F02
                   REPLACE F13 WITH K25T1.F03
                   REPLACE F16 WITH 'N'
                   IF F09='04'
                      REPLACE F18 WITH '1'
                   ENDIF   
                SELECT K25T1
                ST_KEY=CT1+PADL(ALLTRIM(STR(VAL(RIGHT(F07,8))+1)),8,'0')                
                SKIP
             ENDDO

          ENDIF          
          SELECT SALE1
          SET ORDER TO 1
          GO BOTT
          IF RIGHT(F06,2)<>'99' OR RIGHT(F06,2)<>'49'
             IF VAL(RIGHT(F06,2))>50
                CTT=99-VAL(RIGHT(F06,2))
             ELSE    
                CTT=49-VAL(RIGHT(F06,2))
             ENDIF
             ST_KEY=CT1+PADL(ALLTRIM(STR(VAL(RIGHT(F06,8))+1)),8,'0')                             
             FOR I=1 TO CTT
                 APPEND BLANK
                 REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                 REPLACE F02 WITH 'TOKUTSU'
                 REPLACE F03 WITH SUBSTR(K25T1.F19,2,3)+RIGHT(K25T1.F19,2)
                 REPLACE F04 WITH '3'
                 REPLACE F06 WITH ST_KEY
                 REPLACE F07 WITH '+'
                 REPLACE F16 WITH 'N'               
                 ST_KEY=CT1+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')                             
             ENDFOR
          ENDIF
          GO BOTT
          IF MOD(RECCOUNT(),50)<>0
             =MESSAGEBOX('銷項三聯式發票號碼有重複請再次檢查K17',0+48,'提示訊息視窗')
             RETURN
          ENDIF
          DO WHILE !BOF()             
            REPLACE F05 WITH PADL(ALLTRIM(STR(CEILING(RECNO()/50))),4,'0')          
             IF LEFT(F07,1)='+' AND (RIGHT(F06,2)='99' OR  RIGHT(F06,2)='49')
                REPLACE F07 WITH ' '
                DO WHILE .T.
                   SKIP -1
                   IF LEFT(F07,1)<>'+'
                      EXIT
                   ELSE
                      REPLACE F05 WITH PADL(ALLTRIM(STR(CEILING(RECNO()/50))),4,'0')
                      IF RIGHT(F06,2)='00' OR RIGHT(F06,2)='50'
                         REPLACE F07 WITH '*'
                      ELSE      
                         REPLACE F07 WITH   ' '
                      ENDIF      
                   ENDIF
                ENDDO                
                SKIP
                REPLACE F07 WITH '*'           
             ENDIF
             SKIP -1
          ENDDO
          
          CREATE CURSOR SALE2;
          (F01 C(5),F02 C(8),F03 C(5),F04 C(1),F05 C(4),F06 C(10),F07 C(8),F08 C(12),F09 C(5),F10 C(12),F11 C(4),F12 C(4),F13 C(10),F14 C(15),F15 C(10),F16 C(1),F17 C(22),F18 C(1))
          INDEX ON F06 TAG SALE2
          SET ORDER TO 1
          SELECT * FROM K25 WHERE LEFT(F01,2)='32' AND LEN(ALLTRIM(F07))=10 AND LEFT(F07,2)=CT2 AND F08>=0  ORDER BY F07 INTO CURSOR K25T2 NOWAIT
          IF _TALLY>0
*!*	             YMC=PADL(ALLTRIM(STR(MOD(VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,6,2)),2)+VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,6,2))+1)),2,'0')
*!*	             IF YMC='13'
*!*	                YC=PADL(ALLTRIM(STR(VAL(SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,1,4))+1)),3,'0')
*!*	                YMC='01'
*!*	             ELSE
*!*	                YC=SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)
*!*	             ENDIF             
             SELECT K25T2
             GO TOP
             ST_KEY=''
             IF RIGHT(F07,2)='00' OR RIGHT(F07,2)='50'
                ST_KEY=F07                
             ELSE
                IF VAL(RIGHT(F07,2))<50
                   ST_KEY=CT2+LEFT(ST_KEY,6)+'00'
                ELSE
                   ST_KEY=CT2+LEFT(ST_KEY,6)+'50'                
                ENDIF                 
                SELECT SALE2

                FOR I=VAL(RIGHT(ST_KEY,8)) TO VAL(RIGHT(K25T2.F07,8))-1
                    APPEND BLANK
                    REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                    REPLACE F02 WITH 'TOKUTSU'
                    REPLACE F03 WITH SUBSTR(K25T2.F19,2,3)+RIGHT(K25T2.F19,2)
                    REPLACE F04 WITH '3'
                    REPLACE F05 WITH '0001'
                    REPLACE F06 WITH ST_KEY
                    REPLACE F07 WITH '+'
                    REPLACE F16 WITH 'N'                                                           
                    ST_KEY=CT2+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')
                ENDFOR    
             ENDIF                    
             SELECT K25T2
             DO WHILE !EOF()
                IF ST_KEY<>K25T2.F07
                   SELECT SALE2 
                   IF VAL(RIGHT(K25T2.F07,8))-VAL(RIGHT(ST_KEY,8))<100
                      TB=VAL(RIGHT(K25T2.F07,8))-VAL(RIGHT(ST_KEY,8))
                   ELSE                      
                      IF VAL(RIGHT(ST_KEY,2))>50
                         TB=100-VAL(RIGHT(ST_KEY,2))
                      ELSE
                         TB=50-VAL(RIGHT(ST_KEY,2))                         
                      ENDIF
                   ENDIF                      
                      FOR I=1 TO TB
                          APPEND BLANK
                          REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                          REPLACE F02 WITH 'TOKUTSU'
                          REPLACE F03 WITH SUBSTR(K25T2.F19,2,3)+RIGHT(K25T2.F19,2)
                          REPLACE F04 WITH '2'
                          REPLACE F06 WITH ST_KEY
                          REPLACE F07 WITH '+'
                          REPLACE F16 WITH 'N'            
                          ST_KEY=CT2+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')                
                      ENDFOR                    
                ENDIF
                   SELECT SALE2
                   APPEND BLANK
                   REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                   REPLACE F02 WITH 'TOKUTSU'
                   REPLACE F03 WITH SUBSTR(K25T2.F19,2,3)+RIGHT(K25T2.F19,2)
                   REPLACE F04 WITH '2'
                   REPLACE F06 WITH K25T2.F07
                   REPLACE F07 WITH ICASE(K25T2.F16='2','!',K25T2.F16='4','-',SPACE(8))
                   REPLACE F08 WITH PADL(ALLTRIM(TRANSFORM(K25T2.F12,'@R ############')),12,'0')
                   REPLACE F09 WITH ICASE(K25T2.F09='1','5',K25T2.F09='2','01',K25T2.F09='3','*','')
                   REPLACE F10 WITH PADL(ALLTRIM(TRANSFORM(K25T2.F10,'@R ############')),12,'0')
                   REPLACE F11 WITH RIGHT(K25T2.F19,2)+K25T2.F02
                   REPLACE F13 WITH K25T2.F03
                   REPLACE F16 WITH 'N'
                   REPLACE F17 WITH K25T2.F06
                   REPLACE F18 WITH K25T2.F05
                SELECT K25T2
                ST_KEY=CT2+PADL(ALLTRIM(STR(VAL(RIGHT(F07,8))+1)),8,'0')                
                SKIP
             ENDDO

          ENDIF          
          SELECT SALE2
          SET ORDER TO 1
          GO BOTT
          IF RIGHT(F06,2)<>'99' OR RIGHT(F06,2)<>'49'
             IF VAL(RIGHT(F06,2))>50
                CTT=99-VAL(RIGHT(F06,2))
             ELSE    
                CTT=49-VAL(RIGHT(F06,2))
             ENDIF
             ST_KEY=CT2+PADL(ALLTRIM(STR(VAL(RIGHT(F06,8))+1)),8,'0')                             
             FOR I=1 TO CTT
                 APPEND BLANK
                 REPLACE F01 WITH SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                 REPLACE F02 WITH 'TOKUTSU'
                 REPLACE F03 WITH SUBSTR(K25T2.F19,2,3)+RIGHT(K25T2.F19,2)
                 REPLACE F04 WITH '2'
                 REPLACE F06 WITH ST_KEY
                 REPLACE F07 WITH '+'
                 REPLACE F16 WITH 'N'            
                 ST_KEY=CT2+PADL(ALLTRIM(STR(VAL(RIGHT(ST_KEY,8))+1)),8,'0')                             
             ENDFOR
          ENDIF
          GO BOTT
          IF MOD(RECCOUNT(),50)<>0
             =MESSAGEBOX('銷項二聯式發票號碼有重複請再次檢查K17',0+48,'提示訊息視窗')
             RETURN
          ENDIF
          DO WHILE !BOF()             
             REPLACE F05 WITH PADL(ALLTRIM(STR(CEILING(RECNO()/50))),4,'0')          
             IF LEFT(F07,1)='+' AND (RIGHT(F06,2)='99' OR  RIGHT(F06,2)='49')
                REPLACE F07 WITH ' '
                DO WHILE .T.
                   SKIP -1
                   IF LEFT(F07,1)<>'+'
                      EXIT
                   ELSE
                      REPLACE F05 WITH PADL(ALLTRIM(STR(CEILING(RECNO()/50))),4,'0')
                      IF RIGHT(F06,2)='00' OR RIGHT(F06,2)='50'
                         REPLACE F07 WITH '*'
                      ELSE   
                         REPLACE F07 WITH SPACE(1)   
                      ENDIF   
                   ENDIF
                ENDDO                
                SKIP
                REPLACE F07 WITH '*'           
             ENDIF
             SKIP -1
          ENDDO          
          SELECT * FROM SALE1 UNION ALL SELECT * FROM SALE2 INTO CURSOR SALE NOWAIT
          SELECT F01 送件年月,F02 用戶代號,PADL(ALLTRIM(STR(VAL(LEFT(TXMTH2,4))-1911)),3,'0')+RIGHT(TXMTH2,2) 所屬年月,F04 發票聯式,F05 第冊數,F06 發票號碼,IIF(LEFT(F07,1)='!','+'+SPACE(7),F07) 統一編號,F08 銷售額,F09 稅別,F10 稅額,F11 月日,F12 營業稅科目類別代號,F13 客戶代號,F14 公斤數,F18 通關註記,0 免稅銷售額,F15+F16 出售註記及列印註記 FROM SALE NOWAIT
          IF _TALLY>0
             FILE_NAME='銷項發票明細'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
             GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
             IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                 COPY TO (GCDELIMFILE) TYPE XL5
                 =MESSAGEBOX('儲存成功',0+48,'提示訊息視窗')
             ENDIF
          ENDIF                                                  
          SELECT F01 送件年月,F02 用戶代號,F06+SPACE(1) 發票號碼_序,F03+RIGHT(F11,2) 開立日期,PADR(IIF(SEEK(ALLTRIM(F13),'C01'),C01.F04,' '),12, ' ') 買受人,PADR(ALLTRIM(F07),8,SPACE(1)) 統一編號,;
          '電子連接器等'+SPACE(4) 貨物勞務,PADL(ALLTRIM(STR(TTL2(F06,F03))),12,'0') 數量,SUBSTR(F09,2,1) 外銷方式,SPACE(12) 備註,IIF(LEFT(F06,2)=CT1,'發票'+SPACE(8),IIF(F18='1','X7'+SPACE(10),SPACE(12))) 非經海關證明文件名稱,;
          IIF(LEFT(F06,2)=CT1,F06+SPACE(12),IIF(F18='1',F17,SPACE(22))) 非經海關證明文件號碼,IIF(F18='2',PADL('0',12,'0'),F08) 非經海關證明文件金額,IIF(LEFT(F06,2)<>CT1 AND F18='2','G3'+SPACE(2),SPACE(4)) 經海關出口報單類別,IIF(LEFT(F06,2)<>CT1 AND F18='2',F17,SPACE(20)) 經海關出口報單號碼,IIF(LEFT(F06,2)<>CT1 AND F18='2',F08,PADL('0',12,'0')) 經海關出口報單金額,;
          F03+RIGHT(F11,2) 輸出或結匯日期,IIF(IIF(LEFT(F06,2)=CT1,'1',F18)='1','Y',SPACE(1)) 外匯證明文件,F04 發票聯式,IIF(LEFT(F06,2)=CT1,'1',F18) 通關註記 FROM SALE WHERE LEFT(F09,1)='0' AND LEFT(F07,1)<>'!'
          IF _TALLY>0
             FILE_NAME='零稅率清單'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
             GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
             IF !EMPTY(GCDELIMFILE)  && ESC PRESSED
                COPY TO (GCDELIMFILE) TYPE XL5
                =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
             ENDIF      
          ENDIF                
          
          ******************************************************再做進項       
                   
          SELECT K25             
          CREATE CURSOR RECP1;
          (F01 C(5),F02 C(8),F03 C(3),F04 C(2),F05 C(2),F06 C(1),F07 C(4),F08 C(1),F09 C(2),F10 C(1),F11 C(12),F12 C(5),F13 C(12),F14 C(7),F15 C(8),F16 C(10),F17 C(15),F18 C(1),F19 C(12),F20 C(22),F21 C(14))
          INDEX ON F21 TAG RECP1
          SET ORDER TO 1
          SELECT * FROM K25 WHERE LEFT(F01,2)='21' AND LEN(ALLTRIM(F07))=10 AND LEFT(F07,2)=CT1 AND F08>=0  ORDER BY F07 INTO CURSOR K25T1 NOWAIT          
          IF _TALLY>0
             SELECT K25T1
             GO TOP
             STS='1'
             SQR='0'
             DO WHILE !EOF()
                SELECT RECP1
                APPEND BLANK
                REPLACE F01 WITH  SUBSTR(THISFORM.MTH_LIST.DISPLAYVALUE,2,3)+RIGHT(THISFORM.MTH_LIST.DISPLAYVALUE,2)
                REPLACE F02 WITH 'TOKUTSU'
                REPLACE F03 WITH SUBSTR(K25T1.F19,2,3)
                REPLACE F04 WITH IIF(MOD(VAL(RIGHT(K25T1.F19,2)),2)<>0,RIGHT(K25T1.F19,2),PADL(ALLTRIM(STR(VAL(RIGHT(K25T1.F19,2))-1)),2,'0'))
                REPLACE F05 WITH IIF(MOD(VAL(RIGHT(K25T1.F19,2)),2)=0,RIGHT(K25T1.F19,2),PADL(ALLTRIM(STR(VAL(RIGHT(K25T1.F19,2))+1)),2,'0'))
                REPLACE F06 WITH '1'
                REPLACE F07 WITH PADL(STS,4,'0')
                REPLACE F08 WITH 'Y'
                REPLACE F09 WITH PADL(SQR,2,'0')
                REPLACE F10 WITH '1'
                REPLACE F11 WITH PADL(ALLTRIM(STR(K25T1.F08)),12,'0')
                REPLACE F12 WITH ALLTRIM(STR(INT_002))
                REPLACE F13 WITH PADL(ALLTRIM(STR(K25T1.F10)),12,'0')
                REPLACE F14 WITH SUBSTR(K25T1.F19,2,3)+RIGHT(K25T1.F19,2)+K25T1.F02
                REPLACE F15 WITH K25T1.F04
                REPLACE F16 WITH K25T1.F03
                REPLACE F18 WITH 'N'
                REPLACE F21 WITH K25T1.F07
                IF SQR='99'
                   SQR='00'
                   STS=ALLTRIM(STR(VAL(STS)+1))
                ELSE    
                    SQR=ALLTRIM(STR(VAL(SQR)+1))
                ENDIF
                SELECT K25T1
                SKIP
             ENDDO
             SELECT F01 送件年月,F02 用戶代號,F03 所屬年度,F04 所屬月份_起,F05 所屬月份_迄,;
             F06 憑證類別,F07 第冊數,F08 扣抵否,F09 序號,F10 類別,F11 憑證金額,F12 稅別,F13 稅額,F14 日期,F15 開立行號,F16 供應商代號,F17 公斤數,;
             F18 列印註記,F19 徵收費,F20 報單號碼,F21 發票號碼 FROM RECP1 NOWAIT
             IF _TALLY>0
                FILE_NAME='進項發票明細'+LEFT(DTOS(CTOD(THISFORM.MTH_LIST.DISPLAYVALUE+'/01')),6)
                GCDELIMFILE=PUTFILE('儲存路徑',FILE_NAME,'XLS')
                IF !EMPTY(GCDELIMFILE)  
                   COPY TO (GCDELIMFILE) TYPE XL5
                   =MESSAGEBOX('儲存成功',0+64,'提示訊息視窗')
                ENDIF      
             ENDIF   
          ENDIF                                                       
         THISFORM.CMND2.CLICK
      ENDPROC
      procedure cmnd2.click        
           IF !USED('A05')
               SELECT 0
               USE A05
           ELSE
               SELECT A05
           ENDIF
           SET ORDER TO 1
           SEEK sys_oper+'K19'
           IF FOUND()
              DELETE 
           ENDIF        
           close table all
           THISFORM.RELEASE 
           return  
      endproc      
enddefine          

********
*******************************************   計算該單出貨總數量
FUNC TTL2
     PARA BKK_NO,M_TH
     qt_od=0
     IF MOD(VAL(RIGHT(M_TH,2)),2)=0
        SELECT C132
     ELSE
        SELECT C131
     ENDIF   
     SEEK BKK_NO
     IF FOUND()
         DO WHILE F17=BKK_NO                                      
            qt_od=qt_od+(f06)                                                    
            SKIP
         ENDDO
     ELSE
        IF MOD(VAL(RIGHT(M_TH,2)),2)=0
           SELECT B062
        ELSE
           SELECT B061
        ENDIF   
        SEEK BKK_NO     
        IF FOUND()
           DO WHILE F20=BKK_NO                           
              qt_od=qt_od+(f04)                                                    
              SKIP
           ENDDO           
        ELSE
          qt_od=0
        ENDIF   
     ENDIF
     SELECT SALE
     RETURN qt_od
