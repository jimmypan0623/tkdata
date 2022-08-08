close all
clear 
FLG='0'
IF !USED('A02')
   SELE 0
   USE A02 INDE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021      
SEEK sys_oper+'F08'
IF !USED('A01')
   SELE 0
   USE A01
ELSE 
  SELE A01
ENDIF
SET ORDER TO 1


F08form=createobject("tkF08")
F08form.show       
define class tkF08 as form
*!*	  autocenter=.t.
  caption='F08.ＢＯＭ總檢查'
  AUTOCENTER=.T.
  fontsize=INT_015*9
  height=INT_015*320
  width=INT_015*420
  CONTROLBOX=.F.
  BORDERSTYLE=2
  SHOWTIPS=.T.
  showwindow=1
  windowtype=1
  name='tkF08'
  ADD OBJECT LBL1 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*40,;
      AUTOSIZE=.T.,;
      CAPTION=''    
  ADD OBJECT LBL2 AS LABEL WITH;
      LEFT=INT_015*5,;
      TOP=INT_015*80,;
      FONTSIZE=INT_015*30,;
      AUTOSIZE=.T.,;
      CAPTION=''               
ADD OBJECT OUTPUT AS OPTIONGROUP WITH;
      LEFT=INT_015*100,;
      TOP=INT_015*140,;
      HEIGHT=INT_015*40,;    
      AUTOSIZE=.T.,;      
      BUTTONCOUNT=2,;
      OPTION1.CAPTION='未確認',;
      OPTION1.FONTSIZE=INT_015*9,;
      OPTION1.LEFT=INT_015*4,;
      OPTION1.TOP=INT_015*3,;
      OPTION1.AUTOSIZE=.T.,;
      OPTION2.CAPTION='已確認',;
      OPTION2.FONTSIZE=INT_015*9,;
      OPTION2.LEFT=INT_015*70,;
      OPTION2.TOP=INT_015*3,;
      OPTION2.AUTOSIZE=.T.,;          
      NAME='OUTPUT'                         
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
      

         THISFORM.OUTPUT.VALUE=1         

      ENDPROC
      PROCEDURE CMND1.CLICK
         IF !USED('C20')
            SELECT 0
            USE C20
         ELSE
            SELECT C20
         ENDIF
         SET ORDER TO 1
         SET FILTER TO VAL(F03)>0      
         IF !USED('B01')
            SELECT 0
            USE B01
         ELSE
            SELECT B01
         ENDIF      
         SET ORDER TO 1
         SELE F01,F02 FROM B01 WHERE F97='Y' INTO CURSOR B01_2 ORDER BY F01
         CREATE CURSOR B01_1;
         (F01 C(43),F02 C(40))
         INDE ON F01 TAG B01_11
         SET ORDER TO B01_11
         SELE B01_2
         GO TOP
         DO WHILE !EOF()
            SELE B01_1
            APPEND BLANK
            REPL F01 WITH B01_2.F01
            REPL F02 WITH B01_2.F02
            SELE B01_2
            SKIP
         ENDDO 
         SELE B01_2
         USE         
         IF !USED('E00')
            SELECT 0
            USE E00
         ELSE
            SELECT E00
         ENDIF      
         SET ORDER TO 1
         IF !USED('E05')
            SELECT 0 
            USE E05
         ELSE
            SELECT E05
         ENDIF     
         SET ORDER TO E052
         IF !USED('F02')
            SELECT 0
            USE F02
         ELSE
            SELECT F02
         ENDIF
         SET RELATION TO F04 INTO B01
         SET ORDER TO F021      
         IF !USED('F01')
            SELECT 0
            USE F01
         ELSE
            SELECT F01
         ENDIF
         SET ORDER TO F011    
         SET RELA TO F01 INTO B01_1  
         IF THISFORM.OUTPUT.VALUE=1
            CREATE CURSOR BM2; &&子階料號+子階版本+父階料號+父階版本+階次
            (F01 C(43),F02 C(3),F04 C(43),F05 C(3),F08 C(2)) &&祖宗八代紀錄檔
            INDE ON F01+F02+F08 TAG BM21 DESC  &&子階料號+版本降冪排列
         ELSE
             DELE FROM F07 
             DELE FROM F03 
         ENDIF      
         CREATE CURSOR BM4;
         (F01 C(43),F02 C(3),F03 C(50),F04 C(43),F05 C(3),F06 N(14,4),F07 N(14,4),F08 C(4))          
         CREATE CURSOR BM3;
         (F01 C(43),F02 C(3),F04 C(43),F05 C(3),F07 N(8,4),F08 N(8,4),F09 N(2),F10 N(4),F11 N(4),F90 N(2))        &&記錄搜詢軌跡檔
         INDE ON F04+F05+F01+F02 TAG BM31
         SET ORDER TO BM31          
         CREATE CURSOR BM1;
         (F01 C(43),F02 C(3),F03 C(4),F04 C(43),F05 C(3),F06 N(4),F07 N(8,4),F08 N(2),F09 N(5,2),F10 C(1),F90 N(2))
         SET RELA TO F03 INTO E00                                  
         SELECT 0
         SELECT F01
         GO TOP
         DO WHILE !EOF()   
            IF F01.F05=IIF(THISFORM.OUTPUT.VALUE=2,.T.,.F.)
               IF THISFORM.OUTPUT.VALUE=1
                  PT=F01.F01
                  OP=F01.F02
                  S=0
                  SELECT F02
                  SEEK F01.F01+F01.F02
                  DO WHILE F01+F02=F01.F01+F01.F02 &&開始搜尋每一筆之製程欄位是否為空白
                     IF EMPTY(F03) 
                        S=S+1  &&若有則加到變數 S
                        EXIT   &&馬上離開此迴圈
                     ENDIF     
                     SKIP
                  ENDDO
                  IF S>0   &&如果其中有一筆材料沒打上投入製程編號
                     =MESSAGEBOX('料號:'+ALLTRIM(F02.F04)+'之投入製程尚未建立!',0+48,'提示訊息視窗')
                     RETURN    &&回去原來的畫面且不做BOM確認動作   
                  ENDIF     
                  SELE E05           &&製程檔
                  SET ORDER TO E052  &&料號+版本+製程編號
                  PRW=''
                  SELECT F02  &&材料表身檔
                  SEEK F01.F01+F01.F02       &&從材料表身檔中的每比料耗所投入製程去尋找E05是否沒有對應到
                  DO WHILE F01+F02=F01.F01+F01.F02
                     SELECT E05
                     SEEK F02.F01+F02.F02+F02.F03
                     IF !FOUND()     
                        PRW=F02.F03          &&有的話將該製程編號記起來到 PRW 變數中        
                     ENDIF
                     SELECT F02
                     SKIP
                  ENDDO        
                  IF !EMPTY(PRW)  &&若沒對應到
                     =MESSAGEBOX(F01.F01+'之製程:'+PRW+'    '+IIF(SEEK(PRW,'E00'),E00.F04,'')+'可能已被刪除,請於表頭建立!',0+48,'提示訊息視窗') 
                     RETURN                                            
                  ENDIF            
                  SELECT E05
                  SEEK F01.F01+F01.F02
                  IF !FOUND() 
                     =MESSAGEBOX(F01.F01+'請先建立製程再做確認!',0+48,'修改後再做確認')
                     SELE F01
                     RETURN
                  ELSE
                     REPLACE  F99 WITH 0 FOR F01+F02=F01.F01+F01.F02
                     SELE F02
                     SEEK F01.F01+F01.F02
                     DO WHILE F01+F02=F01.F01+F01.F02
                        SELE E05
                        SEEK F02.F01+F02.F02+F02.F03
                        IF FOUND()
                           DO CASE  &&製程型態
                              CASE F05='1' &&一次或不一定  
                                   IF F99=0
                                      REPLACE F99 WITH 1
                                   ENDIF
                              CASE F05='2' &&看使用幾項就是幾次
                                   REPLACE  F99 WITH F99+1 
                              CASE F05='3'  &&使用數量即是製程加工次數
                                   REPLACE  F99 WITH F99+F02.F07
                              CASE F05='4'  &&批次生產
                                   REPLACE F99 WITH IIF(SEEK(F01.F01,'C20'),VAL(C20.F03),1)
                           ENDCASE                  
                        ENDIF   
                        SELE F02
                        SKIP
                     ENDDO       
                     SELE E05
                     REPLACE  F99 WITH 1 FOR F01+F02=F01.F01+F01.F02 AND F99=0          
                  ENDIF               
                  SELE BM2  &&祖宗檔
                  APPEND BLANK   &&建立初始值
                  REPL F01 WITH PT
                  REPL F02 WITH OP
                  REPL F04 WITH PT
                  REPL F05 WITH OP
                  REPL F08 WITH '00'
                  PRIORITY=0       &&統計階層變數
                  FATHER=PT        &&父階料號
                  EGG=OP           &&父階版本 
                  SON=''
                  BIRD=''
                  FLP=0
                  TQTY=1
                  SELE F02
                  set order to F022   &&往回找爸爸
                  SEEK FATHER+EGG
                  IF FOUND()   &&如果找到
                     DO WHILE FLP>=0        &&開始執行回圈
                        PRIORITY=PRIORITY+1  
                        IF PRIORITY>INT_181
                           =MESSAGEBOX('此組BOM可能有父階子階錯亂的問題!請重新檢查修正後再做確認!',0+48,'提示訊息視窗')
                           FLG='1'
                           EXIT
                        ENDIF   
                        SELE BM2    &&記錄來源
                        APPEND BLANK      
                        REPL F01 WITH PT       &&要找的完成品號寫進來
                        REPL F02 WITH OP
                        REPL F04 WITH F02.F01   &&將祖父寫進來 
                        REPL F05 WITH F02.F02
                        REPL F08 WITH PADL(PRIORITY,2,'0')   
                        FATHER=F04  &&寫完之後阿公變爸爸
                        EGG=F05
                        SELE BM3    &&並紀錄在軌跡檔
                        APPEND BLANK
                        REPL F01 WITH F02.F04  &&子
                        REPL F02 WITH F02.F05
                        REPL F04 WITH F02.F01  &&父
                        REPL F05 WITH F02.F02
                        REPL F09 WITH PRIORITY         
                        SELE F02
                        SEEK FATHER+EGG &&再回頭去找祖先
                        IF FOUND()
                           FATHER=F02.F04
                           EGG=F02.F05
                        ELSE  
                           SEEK BM3.F01+BM3.F02+BM3.F04+BM3.F05  &&由軌跡檔去找F02(逆)
                           IF FOUND()
                              SKIP
                              IF F04+F05=BM3.F01+BM3.F02                         
                                 PRIORITY=PRIORITY-1  
                                 SELECT BM3
                                 DELETE                               
                              ELSE
                                 FLP=PRIORITY
                                 DO WHILE .T.
                                    FLP=FLP-1 
                                    IF FLP<0
                                       EXIT
                                    ENDIF   
                                    SELE BM3   
                                    SON=BM3.F01
                                    BIRD=BM3.F02
                                    DELE  
                                    SEEK SON+BIRD
                                    IF FOUND()     
                                       PRIORITY=BM3.F09-1   
                                       SON=BM3.F01
                                       BIRD=BM3.F02                        
                                       FATHER=BM3.F04
                                       EGG=BM3.F05
                                       DELE                                    
                                       SELE F02
                                       SEEK SON+BIRD+FATHER+EGG
                                       SKIP
                                       IF F04+F05=SON+BIRD
                                          EXIT            
                                       ENDIF
                                    ENDIF        
                                 ENDDO                  
                              ENDIF                         
                           ENDIF
                        ENDIF   
                     ENDDO
                     IF FLG='1'
                        SELECT BM2
                        USE
                        RETURN                  
                     ENDIF
                  ENDIF                   
                  DELE FROM F07 WHERE F01+F02=ANY(SELE F04+F05 FROM BM2)
                  DELE FROM F03 WHERE F01+F02=ANY(SELE F04+F05 FROM BM2)  &&樹狀材料表                 
                  SELECT BM2
                 
                 SET ORDER TO BM21
                 GO TOP
                 DO WHILE !EOF()
                    DO EXTRALP WITH BM2.F04,BM2.F05
                       SELECT BM2
                       SKIP               
                 ENDDO
                 SELECT F01
                 REPLACE F05 WITH .T.
                 SELECT BM2
                 ZAP
             ELSE    
                PA1=F01.F01
                PA2=F01.F02
                DO EXTRALP WITH PA1,PA2    
             ENDIF
          ENDIF
          THISFORM.LBL1.CAPTION=F01.F01+F01.F02   
          SELECT F01
          SKIP
        ENDDO
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
           SEEK sys_oper+'F08'
           IF FOUND()
              DELETE 
           ENDIF           
           close table all
           THISFORM.RELEASE 
           return  
      endproc            
ENDDEFINE      
**
*******************************
PROCEDURE EXTRALP 
         PARAMETERS FA1,SN1
                  PT=FA1
                  OP=SN1
                  QTY=IIF(SEEK(PT,'B01'),B01.F28,1)
                  SELE BM1
                  APPEND BLANK 
                  REPL F01 WITH PT  &&設定初始值
                  REPL F02 WITH OP
                  REPL F04 WITH PT
                  REPL F05 WITH OP
                  REPL F06 WITH QTY  &&lead time
                  REPL F07 WITH 1
                  REPL F08 WITH 0
                  PRIORITY=0
                  FATHER=PT
                  EGG=OP
                  SON=''
                  BIRD=''
                  FLP=0
                  TQTY=1
                  LN=1
                  IF !EMPTY(EGG)  &&如果版本空白就不用往下找了
                     SELE F02
                     set order to F021   
                     SEEK FATHER+EGG
                     IF FOUND()
                        DO WHILE FLP>=0    
                           PRIORITY=PRIORITY+1
                           QTY=QTY+B01.F28+B01.F31
                           IF PRIORITY>INT_181
                              =MESSAGEBOX(PT+SPACE(2)+OP+'此組BOM可能有父階子階錯亂的問題!請重新檢查修正後再做確認!',0+48,'提示訊息視窗')
                              FLG='1'
                              SELECT F01
                              REPLACE F05 WITH .F.
                              EXIT
                           ENDIF   
                           SELE BM1
                           APPEND BLANK
                           REPLACE  F01 WITH PT
                           REPLACE  F02 WITH OP
                           REPLACE  F03 WITH F02.F03
                           REPLACE  F04 WITH F02.F04
                           REPLACE  F05 WITH F02.F05
                           REPLACE  F06 WITH QTY
                           REPLACE  F07 WITH TQTY*F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE  F08 WITH PRIORITY
                           REPLACE F09 WITH F02.F12
                           REPLACE F10 WITH F02.F06      
                           REPLACE F90 WITH LN
                           FATHER=BM1.F04
                           EGG=BM1.F05
                           TQTY=BM1.F07              
                           SELE BM4
                           APPEND BLANK
                           REPLACE F01 WITH BM1.F01
                           REPLACE F02 WITH BM1.F02
                           REPLACE F03 WITH PADL(BM1.F08,2,'0')+PADL(ALLTRIM(STR(BM1.F90)),2,'0')+SPACE(6*BM1.F08)+IIF(BM1.F08<>0,'╠═','')                       
                           REPLACE F04 WITH BM1.F04
                           REPLACE F05 WITH BM1.F05
                           REPLACE F06 WITH F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE F07 WITH BM1.F07
                           REPLACE F08 WITH BM1.F03
                           SELE BM3                       
                           APPEND BLANK                                 
                           REPLACE  F01 WITH F02.F01
                           REPLACE  F02 WITH F02.F02
                           REPLACE  F04 WITH F02.F04
                           REPLACE  F05 WITH F02.F05
                           REPLACE  F07 WITH F02.F07^IIF(IIF(SEEK(F04,'B01'),B01.F39,'')='C',-1,1)
                           REPLACE  F08 WITH BM1.F07
                           REPLACE  F09 WITH PRIORITY                   
                           REPLACE  F10 WITH B01.F28+B01.F31         
                           REPLACE  F11 WITH QTY
                           REPLACE F90 WITH LN
                           SELE F02
                           SEEK FATHER+EGG
                           IF FOUND()
                              FATHER=F02.F04
                              EGG=F02.F05
                              LN=1                           
                           ELSE  
                              SEEK BM3.F01+BM3.F02+BM3.F04+BM3.F05
                              IF FOUND()
                                 SKIP      
                                 IF F01+F02=BM3.F01+BM3.F02  &&如果下一筆的父階也一樣則參數要回到原階時的狀況
                                    PRIORITY=PRIORITY-1  
                                    TQTY=BM3.F08/BM3.F07     
                                    QTY=BM3.F11-BM3.F10
                                    LN=LN+1                                 
                                    SELECT BM3               &&將軌跡刪除
                                    DELETE 
                                 ELSE
                                    FLP=PRIORITY
                                    DO WHILE .T.
                                       FLP=FLP-1 
                                       IF FLP<0
                                          EXIT
                                       ENDIF   
                                       SELE BM3   
                                       SON=BM3.F01
                                       BIRD=BM3.F02
                                       DELETE   &&找到原軌跡紀錄就刪除
                                       IF !EMPTY(BIRD)   &&如果版次非空白
                                          SEEK SON+BIRD
                                          IF FOUND()     
                                             PRIORITY=BM3.F09-1   
                                             TQTY=BM3.F08/BM3.F07
                                             QTY=BM3.F11-BM3.F10
                                             SON=BM3.F01
                                             BIRD=BM3.F02                        
                                             FATHER=BM3.F04
                                             EGG=BM3.F05
                                             LN=BM3.F90                                                                                
                                             DELETE                                     
                                             SELE F02
                                             SEEK SON+BIRD+FATHER+EGG
                                             SKIP
                                             IF F01+F02=SON+BIRD
                                                LN=LN+1                                        
                                                EXIT            
                                             ENDIF
                                          ENDIF        
                                       ENDIF   
                                    ENDDO                  
                                 ENDIF                         
                              ENDIF
                           ENDIF   
                        ENDDO
                        IF FLG='1'
                           SELECT BM1
                           ZAP 
                           SELECT BM3
                           ZAP
                           SELECT BM4
                           ZAP
                           FLG='0'
                           RETURN
                        ENDIF
                     ENDIF   
                  ENDIF        


       
            SELE BM1
            GO TOP
            KM=''
            DO WHILE !EOF()
                  SELE F07
                  APPEND BLANK
                  REPL F01 WITH BM1.F01
                  REPL F02 WITH BM1.F02
                  REPL F03 WITH BM1.F03
                  REPL F04 WITH BM1.F04
                  REPL F05 WITH BM1.F05
                  REPL F06 WITH BM1.F06
                  REPL F07 WITH BM1.F07
                  REPL F08 WITH PADL(ALLTRIM(STR(BM1.F08)),2,'0')
                  IF BM1.F09>0
                      REPLACE F09 WITH BM1.F09
                  ENDIF    
                  REPLACE F10 WITH BM1.F10

               SELE BM1
               SKIP
           ENDDO   
 
           SELE BM4
           DO WHILE RECNO()>1
              IF LEFT(F03,4)='0000'
                 SKIP -1
              ENDIF   
              L=AT('╠═',F03)   
              B=RECNO()
              SKIP -1
              IF AT('╠═',F03)>L
                 DO WHILE AT('╠═',F03)>L 
                    H=STUFF(F03,L,2,'║')           
        	        REPL F03 WITH H
        	        SKIP -1
    	         ENDDO     	     
              ENDIF
              GO B-1
           ENDDO
           SELE BM4         
           GO BOTT
******************************************
           IF LEFT(F03,4)<>'0000'
              L=AT('╠═',F03)   
              H=STUFF(F03,L,2,'╚')    
              REPLACE  F03 WITH H      
            ENDIF       
              DO WHILE RECNO()>1
                 KM=LEFT(F03,4)
                 SKIP -1
                 IF LEFT(F03,4)>KM
                    L=AT('╠═',F03)   
                    H=STUFF(F03,L,2,'╚')    
                    REPLACE  F03 WITH H             
                 ENDIF   
              ENDDO
              SELE BM4
              GO TOP
              DO WHILE !EOF()
                 L=AT('╠═',F03)   
                 SKIP
                 IF SUBSTR(F03,L,2)=SPACE(2)
                    SKIP -1
                    H=STUFF(F03,L,2,'╚')    
                    REPL F03 WITH H  
                    SKIP                 
                 ENDIF   
              ENDDO  
              SELE BM4
              GO TOP
              DO WHILE !EOF()
                 SELE F03
                 APPEND BLANK
                 REPLACE F01 WITH BM4.F01
                 REPLACE F02 WITH BM4.F02
                 REPLACE F03 WITH BM4.F03
                 REPLACE F04 WITH BM4.F04
                 REPLACE F05 WITH BM4.F05
                 REPLACE F06 WITH BM4.F06
                 REPLACE F07 WITH BM4.F07                 
                 REPLACE F08 WITH BM4.F08
                 SELE BM4
                 SKIP
              ENDDO
              SELE BM4
              ZAP  
              SELE BM1
              ZAP