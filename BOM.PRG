SET EXCLUSIVE OFF
SET DELETED ON
CLEAR ALL
CLOSE ALL
SELECT 0
USE BOM
SET ORDER TO BOM1
CREATE CURSOR BM1 (父 C(1),子 C(1),用量 N(16,4),階 N(3),數 N(2))
CREATE CURSOR BM3 (父 C(1),子 C(1),用量 N(16,4),總量 N(16,4),階 N(3),數 N(2)) &&軌跡檔
INDEX ON 子+父 TAG BM31        &&可逆推
SET ORDER TO BM31
PT='A'  &&祖宗
PRTY=0  &&階次
FLP=0   &&階次旗標
FATHER=PT  &&設定祖宗變數PT為第一個父親
SON=''
TQTY=1
N=1
SELECT BM1
APPEND BLANK
REPLACE 父 WITH PT
REPLACE 子 WITH PT
REPLACE 階 WITH 0
REPLACE 用量 WITH 1
SELECT BOM
SEEK FATHER
IF FOUND()
   DO WHILE FLP>=0
      PRTY=PRTY+1
      IF PRTY>25    &&設定只能最多25階否則為無窮迴圈跳出
         =MESSAGEBOX('BOM建錯',0+48,'提示訊息視窗')      
         EXIT
      ENDIF   
      SELECT BM1
      APPEND BLANK
      REPLACE 父 WITH PT
      REPLACE 子 WITH BOM.子 
      REPLACE 用量 WITH TQTY*BOM.用量
      REPLACE 階 WITH PRTY
      REPLACE 數 WITH N 
      FATHER=BOM.子 &&將BOM.子階料號丟到變數FATHER準備再往下搜尋子階
      TQTY=BM1.用量
      SELECT BM3
      APPEND BLANK
      REPLACE 父 WITH BOM.父
      REPLACE 子 WITH BOM.子
      REPLACE 用量 WITH BOM.用量
      REPLACE 總量 WITH BM1.用量
      REPLACE 階 WITH PRTY
      REPLACE 數 WITH N 
      SELECT BOM
      SEEK FATHER   &&繼續往子階找下去
      IF FOUND()    &&如果還有再準備往下找
*         SON=父     &&往父階遞迴時的紀錄變數
         FATHER=BOM.子  &&往子階時的搜尋記錄變數
         N=1
      ELSE
         SEEK BM3.父+BM3.子  &&若沒有子階則將剛剛BM3所紀錄的本階資料當鍵值指標指回原來的紀錄
         IF FOUND()          
            SON=BM3.父  &&先把目前指向的父階變數放入SON待會兒遞迴         
            SKIP             &&回到BOM檔原紀錄後往下移動一筆
            IF BOM.父=BM3.父 &&判斷是否有旁系兄弟(若是同個爹即是兄弟)
               PRTY=PRTY-1   &&果然有兄弟此時輩份要往上升即階數減 1            
               TQTY=BM3.總量/BM3.用量  
               N=N+1
               *測試開始
                SELECT BM1
                APPEND BLANK
                REPLACE 父 WITH PT
                REPLACE 子 WITH BOM.父                 
                REPLACE 階 WITH PRTY                               
               *測試結束
               SELECT BM3   &&往上階跳時記錄檔刪掉
               DELETE              
            ELSE             &&如果沒有兄弟的話       
                *測試開始
                   SELECT BM1
                   APPEND BLANK
                   REPLACE 父 WITH PT
                   REPLACE 子 WITH SON                 
                   REPLACE 階 WITH PRTY-1                              
               *測試結束 
               FLP=PRTY      &&先將目前的階數紀錄起來丟到變數FLP
               DO WHILE .T.
                  FLP=FLP-1   &&是不是整串BOM都已經搜尋過一遍了 
                  IF FLP<0    &&是的話跳出迴圈
                     EXIT
                  ENDIF
         
                  SELECT BM3  &&回到軌跡紀錄檔            
                  DELETE      &&先刪掉以免重覆遞迴
                  SEEK SON    
                  IF FOUND()
                     PRTY=BM3.階-1
                     SON=BM3.父
                     FATHER=BM3.子
                     TQTY=BM3.總量/BM3.用量      
                     N=BM3.數     
                     *測試開始
                                
                   SELECT BM1
                   APPEND BLANK
                   REPLACE 父 WITH PT
                   REPLACE 子 WITH SON                 
                   REPLACE 階 WITH PRTY                              
               
                     *測試結束             
                     SELECT BM3                       
                     DELETE
                     SELECT BOM
                     SEEK SON+FATHER
                     IF FOUND()
                        SKIP
                        IF 父=SON
                           N=N+1                                             
                           EXIT
                        ENDIF   
                     ENDIF   
                  ENDIF
               ENDDO   
            ENDIF
         ENDIF   
      ENDIF  
   ENDDO

***************
   
ENDIF
**************

            KM=0

           CREATE CURSOR BM4;
           (料 C(1),樹 C(50),數 N(2))
           SELE BM1
               SET ORDER TO 
           GO TOP
           DO WHILE !EOF()          
              SELE BM4
              APPEND BLANK
              REPL 料 WITH BM1.父
              REPL 樹 WITH PADL(BM1.階,2,'0')+SPACE(2)+SPACE(6*BM1.階)+IIF(BM1.階<>0,'╠═','')+BM1.子
              REPLACE 數 WITH BM1.數
              SELE BM1           
              SKIP
           ENDDO    
           SELE BM4
           DO WHILE RECNO()>1
              IF LEFT(樹,2)='00'
                 SKIP -1
              ENDIF   
              L=AT('╠═',樹)   
              B=RECNO()
              SKIP -1
              IF AT('╠═',樹)>L
                 DO WHILE AT('╠═',樹)>L 
                    H=STUFF(樹,L,2,'║')           
        	        REPL 樹 WITH H
        	        SKIP -1
    	         ENDDO     	     
              ENDIF
              GO B-1
           ENDDO
           SELE BM4
           GO BOTT
           IF LEFT(樹,2)<>'00'
              L=AT('╠═',樹)   
              H=STUFF(樹,L,2,'╚')    
              REPL 樹 WITH H      
              ENDIF       
              DO WHILE RECNO()>1
                 KM=LEFT(樹,2)
                 SKIP -1
                 IF LEFT(樹,2)>KM
                    L=AT('╠═',樹)   
                    H=STUFF(樹 ,L,2,'╚')    
                    REPL 樹  WITH H             
                 ENDIF   
              ENDDO
              SELE BM4
              GO TOP
              DO WHILE !EOF()
                 L=AT('╠═',樹 )   
                 SKIP
                 IF SUBSTR(樹 ,L,2)=SPACE(2)
                    SKIP -1
                    H=STUFF(樹 ,L,2,'╚')    
                    REPL 樹  WITH H  
                    SKIP                 
                 ENDIF   
              ENDDO  
              SELE BM4
              GO TOP
              BROW
SELECT BM1
BROWSE
 
CLOSE ALL
