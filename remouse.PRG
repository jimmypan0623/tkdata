DEFINE POPUP popRight SHORTCUT RELATIVE FROM MROW(),MCOL() 
*!*	DEFINE BAR _med_cut OF popRight PROMPT "剪下(<T)" 
DEFINE BAR 7 OF popRight PROMPT "-----------" 
DEFINE BAR _med_copy OF popRight PROMPT "複製(<C)" 
DEFINE BAR _med_paste OF popRight PROMPT "貼上(<P)" 
*!*	DEFINE BAR _med_clear OF popRight PROMPT "刪除(<D)" 
DEFINE BAR _med_slcta OF popRight PROMPT "全選(<S)" 
PUSH KEY CLEAR 
ACTIVATE POPUP popRight 
ON KEY LABEL RIGHTMOUSE  DO remouse