
DIME  AOU[6,3],IIN[6]
FOR I=1 TO 6
    T=0
    FOR J=1 TO 3
        AOU[I,J]=round(ROUND(RAND()*val(left(time(),2))/24,0);
        +round(RAND()*val(substr(time(),4,2))/60,0)+round(RAND()*val(right(time(),2))/60,0)*rand(),0)
        T=T+AOU[I,J]
    ENDFOR
    IIN[I]=T
ENDFOR     
BS=0
FOR I=1 TO 3
    BS=BS+MOD(IIN[I],2)*2^(3-I)
ENDFOR
KS=0
FOR J=4 TO 6 STEP 1
    KS=KS+MOD(IIN[J],2)*2^(6-J)
ENDFOR
IF !USED('ICHIN')
   SELE 0
   USE ICHIN
ELSE
   SELE ICHIN
ENDIF      
SET ORDER TO ICHIN
MENU2=CREATEOBJECT('TKMENU2')
MENU2.REFRESH

MENU2.SETALL('VISIBLE',.T.,'COMMANDBUTTON')

MENU2.SETALL('AUTOSIZE',.T.,'COMMANDBUTTON')
MENU2.SHOW
********
DEFINE CLASS TKMENU2 AS FORM
       CONTROLBOX=.F.
       MOVABLE=.F.
       CLOSABLE=.F.       
       SHOWWINDOW=1
       BORDERSTYLE=3  
       controlcount=57
       FONTSIZE=INT_015*9
       HEIGHT=INT_015*550
       showtips=.t.
       WIDTH=INT_015*789
       WINDOWTYPE=1
       NAME='TKMENU2'   
       add object label1 as label with;
           TOP=INT_015*410,;          
           WIDTH=INT_015*10,;
           AUTOSIZE=.T.,;
           FONTSIZE=INT_015*16,;
           FONTBOLD=.T.,;
           WORDWRAP=.T.,;
           caption=''  
       add object label2 as label with;
           LEFT=INT_015*110,;
           TOP=INT_015*460,;
           WIDTH=INT_015*690,;
           AUTOSIZE=.T.,;
           FONTSIZE=INT_015*16,;
           FONTBOLD=.T.,;
           WORDWRAP=.T.,;
           caption=''           
       ADD OBJECT CMDQT AS COMMANDBUTTON WITH;
           LEFT=INT_015*680,;
           TOP=INT_015*480,;
           FONTSIZE=INT_015*9,;
           MOUSEPOINTER=99,;
           MOUSEICON='BMPS\PAN_OPEN.cur',;
           CAPTION='\<R.重新登錄'
       PROC INIT
           THISFORM.caption='按鈕操作'+SPACE(10)+ALLTRIM(STR(IIN[1]))+ALLTRIM(STR(IIN[2]))+ALLTRIM(STR(IIN[3]))+ALLTRIM(STR(IIN[4]))+ALLTRIM(STR(IIN[5]))+ALLTRIM(STR(IIN[6]))+GS(BS)+GS(KS)+'<---勢運日本' 
           THISFORM.DRAWWIDTH=INT_015*10
           K=0
           FOR I=6 TO 1 STEP -1
               IF MOD(IIN[I],2)=0
                  IF IIN[I]=0
                     THISFORM.FORECOLOR=255                  
                  ELSE
                     THISFORM.FORECOLOR=0
                  ENDIF      
                  THISFORM.LINE(INT_015*40,INT_015*(K*20+420),INT_015*55,INT_015*(K*20+420))
                  THISFORM.LINE(INT_015*75,INT_015*(K*20+420),INT_015*90,INT_015*(K*20+420))
               ELSE
                  IF IIN[I]=3
                     THISFORM.FORECOLOR=255                  
                  ELSE
                     THISFORM.FORECOLOR=0
                  ENDIF      
                  THISFORM.LINE(INT_015*40,INT_015*(K*20+420),INT_015*90,INT_015*(K*20+420))
               ENDIF   
               K=K+1
           ENDFOR    
           thisform.label1.caption=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),ICHIN.F02,'')
           thisform.label2.caption=ichin.f03           
       ENDPROC   
       PROC CMDQT.CLICK
            IF !USED('A04')
              SELE 0
              USE A04
            ELSE 
              SELE A04
            Endif 
            SET ORDER TO 1
            SEEK sys_oper
            IF FOUND()
               REPLACE F01 WITH 'zzzzz'
               REPLACE F04 WITH ''
               REPLACE F06 WITH ''
            ENDIF   
            CLOSE TABLE ALL       
            RELEASE ALL 
            THISFORM.RELEASE
*            CLEAR EVENTS
*            RETURN  
       ENDPROC            
ENDDEFINE
RETURN
***************

*******
PROC GS
     PARA BK
     OL=''
     DO CASE
        CASE BK=7
             OL='天'      
        CASE BK=6
             OL='澤'      
        CASE BK=5
             OL='火'      
        CASE BK=4
             OL='雷'      
        CASE BK=3
             OL='風'      
        CASE BK=2
             OL='水'      
        CASE BK=1
             OL='山'      
        CASE BK=0
             OL='地'           
      ENDCASE  
      RETURN OL
***********

