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
          IIN[I]=&CLM &&��X����n�J�ɤR�L���C������(�u��6,7,8,9�|�إi��)
      ENDFOR
   ELSE   &&�p�G�S�����h�q���e�R�@��
      SELECT A07
      APPEND BLANK  &&�ñN��s��A07
      REPLACE F01 WITH sys_oper
      REPLACE F02 WITH DATE()
      FOR I=1 TO 6 &&�᤻���ɪO
         T=0
         FOR J=1 TO 3 &&�C���P�ɥ�T�T
            AOU[I,J]=2+MOD(round(12*(RAND()*val(left(time(),2))/24+RAND()*val(substr(time(),4,2))/60+RAND()*val(right(time(),2))/60),0),2) &&�M���D2�Y3
            T=T+AOU[I,J]
         ENDFOR
         IIN[I]=T &&�����C�����M(�u��6,7,8,9�|�إi��)
         CLM=FIELD(2+I)
         REPLACE &CLM WITH T
      ENDFOR
   ENDIF        
   SELECT A07
   USE
   
   
   ***********�������ܨ�
   BS=0
   GE=0
   FOR I=1 TO 3               &&�U��
      BS=BS+MOD(IIN[I],2)*2^(3-I)  &&�H 6 �� 8 ������ 0; 7 �� 9 ������1,�ন�G�i���A�ন�Q�i������U���`�M
      GE=GE+ICASE(IIN[I]=6,1*2^(3-I),IIN[I]=9,0, MOD(IIN[I],2)*2^(3-I)) &&�����O�_���ܨ�
   ENDFOR
   KS=0
   DK=0
   FOR J=4 TO 6 STEP 1        &&�W��
       KS=KS+MOD(IIN[J],2)*2^(6-J)  &&�H 6 �� 8 ������ 0; 7 �� 9 ������1,�ন�G�i���A�ন�Q�i������W���`�M
       DK=DK+ICASE(IIN[J]=6,1*2^(6-J),IIN[J]=9,0, MOD(IIN[J],2)*2^(6-J)) &&�����O�_���ܨ�
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

   IAO1='�f�f�f'
   IAO2='�f    �f'
   WK=LIOSO(TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))-1     
   FOR I=1 TO 6 
       WK=WK+1
       JMY='LABL0'+ALLTRIM(STR(I))
       MENU2.ADDOBJECT('&JMY','&JMY')
       MENU2.SETALL('BACKSTYLE',0,'&JMY')
       MENU2.SETALL('TOP',INT_015*(495-15*(I-1)),'&JMY')
       MENU2.SETALL('LEFT',INT_015*23,'&JMY')
       MENU2.SETALL('VISIBLE',.T.,'&JMY')
       MENU2.SETALL('FORECOLOR',IIF(MOD(IIN[I],3)=0,RGB(255,0,0),RGB(255,255,255)),'&JMY')     &&�P�_��˩��ݤ������ʤ��H�Ŧ����
       MENU2.SETALL('FONTSIZE',INT_015*8,'&JMY')   
       MENU2.SETALL('AUTOSIZE',.T.,'&JMY')
       MENU2.SETALL('TOOLTIPTEXT',SOLIO(WK%6),'&JMY') &&�w���~       
       MENU2.SETALL('CAPTION',IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(I)),'GUAZI'),GUAZI.F04+'  '+IIF(IIN[I]%2=0,IAO2,IAO1)+' '+GUAZI.F05+HRT(LEFT(GUAZI.F04,2)),'')+GPP(LEFT(GUAZI.F04,2))+SHS(LEFT(GUAZI.F04,2)),'&JMY')     &&�w����εe���Φw����
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
           FONTNAME='�з���',;                      
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
           FONTNAME='�з���',;                      
           WORDWRAP=.T.,;
           BACKSTYLE=0,;
           caption=''                    
       PROC INIT 
                  IF INT_015=1
                      THISFORM.PICTURE='BMPS\XP800600.JPG'     
                  ELSE    
                      THISFORM.PICTURE='BMPS\XP1024768+LONG.jpg' 
                  ENDIF       
                  THISFORM.caption=SPACE(2)+TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1))+DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))+'��';
                  +'('+DG(((10-VAL(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))%10+MOD(MOD(VAL(SYS(1))-2442130,60),12)+1)%12);
                  +DG(((10-VAL(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1)))%10+MOD(MOD(VAL(SYS(1))-2442130,60),12)+2)%12)+'��)';   
                  +SPACE(10)+IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'JAUGAN'),(JAUGAN.F02+JAUGAN.F03),'')               
*!*	                  +SPACE(10)+GS(IIN[6])+GS(IIN[5])+GS(IIN[4])+GS(IIN[3])+GS(IIN[2])+GS(IIN[1])+SPACE(2);

*!*	                  +IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),ICHIN.F02,'')+IIF(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))<>ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'  ��  ';
*!*	                  +IIF(SEEK(ALLTRIM(STR(DK))+ALLTRIM(STR(GE)),'ICHIN'),ICHIN.F02,''),'')+IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),SPACE(10)+'���q���G'+;
*!*	                  ICHIN.F04+SPACE(4)+'�j�H���G'+ICHIN.F05+SPACE(4)+'�@�ά��G'+ICHIN.F06+SPACE(4)+'�N���G'+ICHIN.F07,'')              
                  ****

                  THISFORM.DRAWWIDTH=INT_015*10
                  Q=0
                  FOR I=6 TO 1 STEP -1        &&�p��X��ʤ�
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
                   thisform.label1.caption=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS)),'ICHIN'),ICHIN.F02,'')    &&��ܥ������W
                   DO CASE
                         CASE Q=0                                        &&�����w�R�h��ܥ�������
                                   thisform.label2.caption=ichin.f03
                         CASE Q=1                                        &&�@���ʫh��ܸӤ����� 
                                   THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(1))),'GUAZI'),GUAZI.F03,'') 
                         CASE Q=2                                        &&�G����       
                                    IF IIN(GUA(1))=IIN(GUA(2))                      &&�Y�P�������ΦP�������h�H�W��������� 
                                        THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'')       
                                    ELSE     
                                        IF IIN(GUA(1))>IIN(GUA(2))                   &&�Y���@���@���h��ܳ�������(�\���D�{�b���D����) 
                                            THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'')       
                                        ELSE
                                            THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(1))),'GUAZI'),GUAZI.F03,'')       
                                        ENDIF   
                                    ENDIF
                         CASE Q=3                                        &&�T���ʫh��������������     
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUA(2))),'GUAZI'),GUAZI.F03,'') 
                         CASE Q=4                                        &&�|���ʨ���w�R�������U���� 
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUB(1))),'GUAZI'),GUAZI.F03,'')  
                         CASE Q=5                                        &&�����ʫh�����ʤ�����
                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+ALLTRIM(STR(GUB(1))),'GUAZI'),GUAZI.F03,'')  
                         CASE Q=6                                        &&�����Ұ�     
                                    DO CASE
                                          CASE ALLTRIM(STR(KS))+ALLTRIM(STR(BS))='00'    &&�Y�O�[���������ʫh���Τ�
                                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+'A','GUAZI'),GUAZI.F03,'')                        
                                          CASE ALLTRIM(STR(KS))+ALLTRIM(STR(BS))='77'    &&�Y�O�����������ʫh���ΤE                  
                                                    THISFORM.LABEL2.CAPTION=IIF(SEEK(ALLTRIM(STR(KS))+ALLTRIM(STR(BS))+'A','GUAZI'),GUAZI.F03,'')
                                     OTHERWISE                                                                           &&��L���������ʫh��ν��
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
        CASE LSA='��' OR LSA='�A'
             POL=1
        CASE LSA='��' OR LSA='�B'
             POL=2
        CASE LSA='��' 
             POL=3
        CASE LSA='�v'
             POL=4
        CASE LSA='��' OR LSA='��'
             POL=5 
        CASE LSA='��' OR LSA='��'
             POL=6
      ENDCASE     
      RETURN POL                  
*************************          
FUNC SOLIO
     PARA ASL
     LOP=''
     DO CASE
        CASE ASL=1
             LOP='�C�s'
        CASE ASL=2
             LOP='����'
        CASE ASL=3 
             LOP='�ĳ�'
        CASE ASL=4
             LOP='�f�D'
        CASE ASL=5
             LOP='�ժ�' 
        CASE ASL=0
             LOP='�ȪZ'
      ENDCASE     
      RETURN LOP                  
*******************************������
FUNCTION HRT
      PARAMETERS DZ
      DV=DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))
      DO CASE
            CASE ATC(DV,'�Ӥl��')>0
                       IF DZ='��'
                           RETURN '��'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(DV,'�x����')>0
                       IF DZ='��'
                            RETURN '��'
                        ELSE
                            RETURN SPACE(0)
                         ENDIF       
            CASE ATC(DV,'�G�Ȧ�')>0
                       IF DZ='�f'
                            RETURN '��'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(DV,'��f��')>0
                       IF DZ='�l'
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF       
        OTHERWISE 
                RETURN SPACE(0)                  
       ENDCASE                 
*******************************�ѤA�Q�H�����ҥ�������, �A�v���U�m, ���B������, �ЬѨ߳D��, �����{�갨, ���O�Q�H��.
FUNCTION GPP
     PARAMETERS PZ
     DV=TG(RIGHT(STR(MOD(VAL(SYS(1))-2442130,60)),1))
     DO CASE
            CASE ATC(DV,'�ҥ���')>0     
                       IF ATC(PZ,'����')>0
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF    
            CASE ATC(DV,'�A�v')>0     
                       IF ATC(PZ,'�l��')>0
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                           
            CASE ATC(DV,'���B')>0     
                       IF ATC(PZ,'�註')>0
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                                               
            CASE ATC(DV,'�Ь�')>0     
                       IF ATC(PZ,'�f�x')>0
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF                                               
            CASE ATC(DV,'����')>0     
                       IF ATC(PZ,'�G��')>0
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF               
      OTHERWISE 
                RETURN SPACE(0)                                                 
     ENDCASE      
*******************************�氨���G�Ȧ���e���A���{�Ӥ��A�Ӥ��Y�氨���C �Ӥl����e���A���{�G���A�G���Y�氨���C�x������e���A���{����A����Y�氨���C��f����e���A���{�x���A�x���Y�氨���C
FUNCTION SHS
      PARAMETERS SZ
      SV=DG(MOD(MOD(VAL(SYS(1))-2442130,60),12))
      DO CASE
            CASE ATC(SV,'�Ӥl��')>0
                       IF SZ='�G'
                           RETURN '��'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(SV,'�x����')>0
                       IF SZ='��'
                            RETURN '��'
                        ELSE
                            RETURN SPACE(0)
                         ENDIF       
            CASE ATC(SV,'�G�Ȧ�')>0
                       IF SZ='��'
                            RETURN '��'
                       ELSE
                            RETURN SPACE(0)   
                        ENDIF    
            CASE ATC(SV,'��f��')>0
                       IF SZ='�x'
                           RETURN '��'
                       ELSE
                           RETURN SPACE(0)
                       ENDIF       
        OTHERWISE 
                RETURN SPACE(0)                  
       ENDCASE 
***         

