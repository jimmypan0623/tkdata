CLOSE ALL
CLEAR
DIME  AOU[6,3],IIN[6],GUA[6],GUB[6]
   IF !USED('A07')
       SELECT 0
       USE A07
   ELSE
       SELECT A07
   ENDIF 
   SET ORDER TO 1
   SEEK sys_oper+DTOS(DATE())
   IF FOUND()        
      FOR I= 1 TO 6
          CLM=FIELD(2+I)
          IIN[I]=&CLM &&找出本日登入時卜過的每爻紀錄(只有6,7,8,9四種可能)
      ENDFOR
   ELSE   &&如果沒紀錄則電腦占卜一次
      SELECT A07
      APPEND BLANK  &&並將其存到A07
      REPLACE F01 WITH sys_oper
      REPLACE F02 WITH DATE()
      FOR I=1 TO 6 &&丟六次銅板
         T=0
         FOR J=1 TO 3 &&每次同時丟三枚
            AOU[I,J]=2+MOD(round(12*(RAND()*val(left(time(),2))/24+RAND()*val(substr(time(),4,2))/60+RAND()*val(right(time(),2))/60),0),2) &&和為非2即3
            T=T+AOU[I,J]
         ENDFOR
         IIN[I]=T &&紀錄每爻的和(只有6,7,8,9四種可能)
         CLM=FIELD(2+I)
         REPLACE &CLM WITH T
      ENDFOR
   ENDIF        
   SELECT A07
   USE
   
   
   ***********本卦及變卦
   BS=0
   GE=0
   FOR I=1 TO 3               &&下卦
      BS=BS+MOD(IIN[I],2)*2^(3-I)  &&以 6 或 8 陰爻為 0; 7 或 9 陽爻為1,轉成二進制後再轉成十進制紀錄下卦總和
      GE=GE+ICASE(IIN[I]=6,1*2^(3-I),IIN[I]=9,0, MOD(IIN[I],2)*2^(3-I)) &&紀錄是否為變卦
   ENDFOR
   KS=0
   DK=0
   FOR J=4 TO 6 STEP 1        &&上卦
       KS=KS+MOD(IIN[J],2)*2^(6-J)  &&以 6 或 8 陰爻為 0; 7 或 9 陽爻為1,轉成二進制後再轉成十進制紀錄上卦總和
       DK=DK+ICASE(IIN[J]=6,1*2^(6-J),IIN[J]=9,0, MOD(IIN[J],2)*2^(6-J)) &&紀錄是否為變卦
   ENDFOR
***************   
IF !USED('JAUGAN')
   SELECT 0
   USE JAUGAN
ELSE
   SELECT JAUGAN
ENDIF
SET ORDER TO 1      
IF !USED('ICHIN')
   SELE 0
   USE ICHIN
ELSE
   SELE ICHIN
ENDIF      
SET ORDER TO ICHIN
IF !USED('GUZAI')
   SELE 0
   USE GUAZI
ELSE
   SELE GUAZI
ENDIF
SET ORDER TO 1      
IF !USED('A03')
   SELE 0
   USE A03
ELSE
   SELE A03
ENDIF      
SET ORDER TO A03
IF !USED('A02')
   SELE 0
   USE A02
ELSE
   SELE A02
ENDIF
SET ORDER TO A021 

**


   MENU2=CREATEOBJECT('TKMENU2')
   MENU2.REFRESH
   IK=0
   S=0

   IAO1='▅▅▅'
   IAO2='▅    ▅'
   WK=LIOSO(TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))-1     
   FOR I=1 TO 6 
       WK=WK+1
       JMY='LABL0'+ALLTRIM(STR(I))
       MENU2.ADDOBJECT('&JMY','&JMY')
       MENU2.SETALL('BACKSTYLE',0,'&JMY')
       MENU2.SETALL('TOP',INT_015*(495-15*(I-1)),'&JMY')
       MENU2.SETALL('LEFT',INT_015*23,'&JMY')
       MENU2.SETALL('VISIBLE',.T.,'&JMY')
       MENU2.SETALL('FORECOLOR',IIF(MOD(IIN[I],3)=0,RGB(255,0,0),RGB(255,255,255)),'&JMY')     &&判斷何親所屬之爻為動爻以藍色顯示
       MENU2.SETALL('FONTSIZE',INT_015*8,'&JMY')   
       MENU2.SETALL('AUTOSIZE',.T.,'&JMY')
       MENU2.SETALL('TOOLTIPTEXT',SOLIO(WK%6),'&JMY') &&安六獸       
       MENU2.SETALL('CAPTION',IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(I)),'GUAZI'),GUAZI.F04+'  '+IIF(IIN[I]%2=0,IAO2,IAO1)+' '+GUAZI.F05+HRT(LEFT(GUAZI.F04,2)),'')+GPP(LEFT(GUAZI.F04,2))+SHS(LEFT(GUAZI.F04,2)),'&JMY')     &&安五行及畫爻及安六親
   ENDFOR
   MENU2.SETALL('VISIBLE',.T.,'COMMANDBUTTON')
   MENU2.SETALL('AUTOSIZE',.T.,'COMMANDBUTTON')
   MENU2.SHOW

********
DEFINE CLASS TKMENU2 AS FORM
       CONTROLBOX=.F.
       MOVABLE=.F.
       CLOSABLE=.F.       
       SHOWWINDOW=1
       BORDERSTYLE=1  
       controlcount=57
       FONTSIZE=INT_015*9
       HEIGHT=INT_015*550
       showtips=.t.
       WIDTH=INT_015*794
       WINDOWTYPE=1
       NAME='TKMENU2'                                                          
       add object label1 as label with;
           TOP=INT_015*418,;          
           WIDTH=INT_015*10,;
           AUTOSIZE=.T.,;
           FONTSIZE=INT_015*14,;
           FONTBOLD=.T.,;
           FONTNAME='標楷體',;                      
           WORDWRAP=.T.,;
           BACKSTYLE=0,;
           caption=''  
       add object label2 as label with;
           LEFT=INT_015*122,;
           TOP=INT_015*418,;
           WIDTH=INT_015*668,;
           HEIGHT=INT_015*70,;
           FONTSIZE=INT_015*14,;
           FONTBOLD=.T.,;
           FONTNAME='標楷體',;                      
           WORDWRAP=.T.,;
           BACKSTYLE=0,;
           caption=''                    
       PROC INIT 
                  IF INT_015=1
                      THISFORM.PICTURE='BMPS\XP800600.JPG'     
                  ELSE    
                      THISFORM.PICTURE='BMPS\XP1024768+LONG.jpg' 
                  ENDIF       
                  THISFORM.caption=SPACE(2)+TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1))+DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))+'日';
                  +'('+DG(((10-VAL(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))%10+MOD(MOD(VAL(SYS(1))-2442130,60),12)+1)%12);
                  +DG(((10-VAL(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))%10+MOD(MOD(VAL(SYS(1))-2442130,60),12)+2)%12)+'空)';   
                  +SPACE(10)+IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'JAUGAN'),(JAUGAN.F02+JAUGAN.F03),'')               
*!*	                  +SPACE(10)+GS(IIN[6])+GS(IIN[5])+GS(IIN[4])+GS(IIN[3])+GS(IIN[2])+GS(IIN[1])+SPACE(2);

*!*	                  +IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),ICHIN.F02,'')+IIF(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))<>ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'  之  ';
*!*	                  +IIF(SEEK(ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'ICHIN'),ICHIN.F02,''),'')+IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),SPACE(10)+'卦義為：'+;
*!*	                  ICHIN.F04+SPACE(4)+'大象為：'+ICHIN.F05+SPACE(4)+'作用為：'+ICHIN.F06+SPACE(4)+'代表為：'+ICHIN.F07,'')              
                  ****

                  THISFORM.DRAWWIDTH=INT_015*10
                  Q=0
                  FOR I=6 TO 1 STEP -1        &&計算幾支動爻
                          IF MOD(IIN[I],2)=0
                               IF IIN[I]=6 
                                   AINS(GUA,1,1)
                                   STORE I TO GUA(1)
                                   Q=Q+1                 
                               ELSE
                                   AINS(GUB,1,1)
                                   STORE I TO GUB(1)
                                ENDIF      
                          ELSE
                               IF IIN[I]=9
                                   AINS(GUA,1,1)
                                   STORE I TO GUA(1)
                                   Q=Q+1                
                               ELSE
                                   AINS(GUB,1,1)
                                   STORE I TO GUB(1)
                               ENDIF      
                          ENDIF          
                   ENDFOR    
                   thisform.label1.caption=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),ICHIN.F02,'')    &&顯示本卦卦名
                   DO CASE
                         CASE Q=0                                        &&六爻安靜則顯示本卦卦辭
                                   thisform.label2.caption=ichin.f03
                         CASE Q=1                                        &&一爻動則顯示該爻爻辭 
                                   THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(1))),'GUAZI'),GUAZI.F03,'') 
                         CASE Q=2                                        &&二爻動       
                                    IF IIN(GUA(1))=IIN(GUA(2))                      &&若同為陰爻或同為陽爻則以上爻爻辭顯示 
                                        THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'')       
                                    ELSE     
                                        IF IIN(GUA(1))>IIN(GUA(2))                   &&若為一陰一陽則顯示陰爻爻辭(蓋陽主現在陰主未來) 
                                            THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'')       
                                        ELSE
                                            THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(1))),'GUAZI'),GUAZI.F03,'')       
                                        ENDIF   
                                    ENDIF
                         CASE Q=3                                        &&三爻動則取中間爻之爻辭     
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'') 
                         CASE Q=4                                        &&四爻動取兩安靜爻中之下爻辭 
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUB(1))),'GUAZI'),GUAZI.F03,'')  
                         CASE Q=5                                        &&五爻動則取未動之爻辭
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUB(1))),'GUAZI'),GUAZI.F03,'')  
                         CASE Q=6                                        &&六爻皆動     
                                    DO CASE
                                          CASE ALLTRIM(STR(KS))+ALLTRIM(STR(BS))='00'    &&若是坤卦六爻全動則取用六
                                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+'A','GUAZI'),GUAZI.F03,'')                        
                                          CASE ALLTRIM(STR(KS))+ALLTRIM(STR(BS))='77'    &&若是乾卦六爻全動則取用九                  
                                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+'A','GUAZI'),GUAZI.F03,'')
                                     OTHERWISE                                                                           &&其他卦六爻全動則取彖辭
                                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+'0','GUAZI'),GUAZI.F03,'')
                                    ENDCASE                              
                   ENDCASE
       ENDPROC         
       PROCEDURE ACTIVATE
            DO MENU4
            THISFORM.RELEASE
       ENDPROC       

ENDDEFINE
*******************************
DEFINE CLASS LABL01 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
DEFINE CLASS LABL02 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
DEFINE CLASS LABL03 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
DEFINE CLASS LABL04 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
DEFINE CLASS LABL05 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
DEFINE CLASS LABL06 AS LABEL
       PROC MOUSEENTER      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(255,255,255)
                    THIS.FORECOLOR=RGB(0,0,0)
                ELSE
                    THIS.ForeColor= RGB(0,0,255) 
                ENDIF  
       ENDPROC     
       PROC MOUSELEAVE      
                LPARAMETERS nButton, nShift, nXCoord, nYCoord 
                IF THIS.ForeColor=RGB(0,0,0)
                    THIS.FORECOLOR=RGB(255,255,255)
                ELSE
                    THIS.ForeColor= RGB(255,0,0) 
                ENDIF  
       ENDPROC                    
ENDDEFINE
*******************************       
FUNC LIOSO
     PARA LSA
     POL=0
     DO CASE
        CASE LSA='甲' OR LSA='乙'
             POL=1
        CASE LSA='丙' OR LSA='丁'
             POL=2
        CASE LSA='戊' 
             POL=3
        CASE LSA='己'
             POL=4
        CASE LSA='庚' OR LSA='辛'
             POL=5 
        CASE LSA='壬' OR LSA='癸'
             POL=6
      ENDCASE     
      RETURN POL                  
*************************          
FUNC SOLIO
     PARA ASL
     LOP=''
     DO CASE
        CASE ASL=1
             LOP='青龍'
        CASE ASL=2
             LOP='朱雀'
        CASE ASL=3 
             LOP='勾陳'
        CASE ASL=4
             LOP='螣蛇'
        CASE ASL=5
             LOP='白虎' 
        CASE ASL=0
             LOP='玄武'
      ENDCASE     
      RETURN LOP                  
*******************************桃花爻◆
FUNCTION HRT
      PARAMETERS DZ
      DV=DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))
      DO CASE
            CASE ATC(DV,'申子辰')>0
                       IF DZ='酉'
                           RETURN '◆'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(DV,'巳酉丑')>0
                       IF DZ='午'
                            RETURN '◆'
                        ELSE
                            RETURN SPACE(0)
                         ENDIF       
            CASE ATC(DV,'寅午戌')>0
                       IF DZ='卯'
                            RETURN '◆'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(DV,'亥卯未')>0
                       IF DZ='子'
                           RETURN '◆'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF       
        OTHERWISE 
                RETURN SPACE(0)                  
       ENDCASE                 
*******************************天乙貴人爻★甲戊庚牛羊, 乙己鼠猴鄉, 丙丁豬雞位, 壬癸兔蛇藏, 庚辛逢虎馬, 此是貴人方.
FUNCTION GPP
     PARAMETERS PZ
     DV=TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1))
     DO CASE
            CASE ATC(DV,'甲戊庚')>0     
                       IF ATC(PZ,'丑未')>0
                           RETURN '★'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF    
            CASE ATC(DV,'乙己')>0     
                       IF ATC(PZ,'子申')>0
                           RETURN '★'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                           
            CASE ATC(DV,'丙丁')>0     
                       IF ATC(PZ,'亥酉')>0
                           RETURN '★'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                                               
            CASE ATC(DV,'壬癸')>0     
                       IF ATC(PZ,'卯巳')>0
                           RETURN '★'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                                               
            CASE ATC(DV,'庚辛')>0     
                       IF ATC(PZ,'寅午')>0
                           RETURN '★'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF               
      OTHERWISE 
                RETURN SPACE(0)                                                 
     ENDCASE      
*******************************驛馬●寅午戌日占卦，卦現申爻，申爻即驛馬爻。 申子辰日占卦，卦現寅爻，寅爻即驛馬爻。巳酉丑日占卦，卦現亥爻，亥爻即驛馬爻。亥卯未日占卦，卦現巳爻，巳爻即驛馬爻。
FUNCTION SHS
      PARAMETERS SZ
      SV=DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))
      DO CASE
            CASE ATC(SV,'申子辰')>0
                       IF SZ='寅'
                           RETURN '●'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(SV,'巳酉丑')>0
                       IF SZ='亥'
                            RETURN '●'
                        ELSE
                            RETURN SPACE(0)
                         ENDIF       
            CASE ATC(SV,'寅午戌')>0
                       IF SZ='申'
                            RETURN '●'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(SV,'亥卯未')>0
                       IF SZ='巳'
                           RETURN '●'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF       
        OTHERWISE 
                RETURN SPACE(0)                  
       ENDCASE 
***         

