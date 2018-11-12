'-------------------------------------------------------------------------------
'-- VBS-Script-Datei
'-- Neu erstellt am 17.09.2008 07:52:54
'-- Ersteller: 
'-- Kommentar: 
'-------------------------------------------------------------------------------
Option Explicit  'Erzwingt die explizite Deklaration aller Variablen in einem Skript.

dim x1,x2,xref,xbeg,noise,x1_step,x2_step,limit,xbeg_step,count,input_chn_no, x_end, x_end_step
dim input_chn,t_start,t_end

t_start=0.3
t_end=0.2

x1=t_start
x_end=t_end
input_chn="test/VS"


'create channel for result of calculation
if groupindexget("calculations")<>0 then call groupdel(groupindexget("calculations"))

call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/input")

call groupdefaultset(groupindexget("calculations"))
input_chn_no=cno("input")


noise=0.2*(cmax("calculations/input")-cmin("calculations/input"))
msgbox noise


xbeg=x1
x1_step=wf_step(x1,"calculations/input")
x_end_step=wf_step(x_end,"calculations/input")
xbeg_step=x1_step

if t_start < t_end then
          msgbox "vorwärts"
          for count=x1_step to x1_step+20
            limit=limit+chdx(count,input_chn_no)
          next

          limit=limit/20


          while (abs((chdx(x1_step,input_chn_no)-limit)) < noise) and x1_step <= x_end_step
            x1_step=x1_step+1
          wend

          x2_step=x1_step

          while (abs(chdx(x2_step,input_chn_no)-limit) > noise/3) and (x2_step > xbeg_step)
            x2_step=x2_step-1
          wend

          x1=wf_time(x1_step,"calculations/input")
          x2=wf_time(x2_step,"calculations/input")

          msgbox "x1: "&x1&"x2: "&x2

          call chncopy("calculations/input","calculations/reg_1")
          call chncopy("calculations/input","calculations/reg_2")

          call select_timerange("calculations/reg_1",xbeg,x2)
          call select_timerange("calculations/reg_2",x2,x1)

  elseif t_start > t_end then

          for count=x1_step to x1_step-20 step -1
            limit=limit+chdx(count,input_chn_no)
          next

          limit=limit/20


          while (abs((chdx(x1_step,input_chn_no)-limit)) < noise) and x1_step >= x_end_step
            x1_step=x1_step-1
          wend

          x2_step=x1_step

          while (abs(chdx(x2_step,input_chn_no)-limit) > noise/3) and (x2_step < xbeg_step)
            x2_step=x2_step+1
          wend

          x1=wf_time(x1_step,"calculations/input")
          x2=wf_time(x2_step,"calculations/input")

          msgbox "x1: "&x1&"x2: "&x2

          call chncopy("calculations/input","calculations/reg_1")
          call chncopy("calculations/input","calculations/reg_2")

          call select_timerange("calculations/reg_1",x2,xbeg)
          call select_timerange("calculations/reg_2",x1,x2)

  else 
          msgbox "t_start=t_end! End of function."

end if








'calculate regression
call chnregrXYcalc("","calculations/reg_1","","calculations/reg_1","linear","Partition complete area",1000,1)

dim y_cut_1,y_cut_2,slope_1,slope_2, arisepoint

y_cut_1=regrcoeffa
slope_1=regrcoeffb

'calculate regression
call chnregrXYcalc("","calculations/reg_2","","calculations/reg_2","linear","Partition complete area",1000,1)

y_cut_2=regrcoeffa
slope_2=regrcoeffb

arisepoint=(trunc(((y_cut_2-y_cut_1)/(slope_1-slope_2)*100000)+0.5))/100000

msgbox arisepoint

call chnalloc("arisepoint_x",1,,,,groupindexget("calculations"))
chd(1,"calculations/arisepoint_x")=arisepoint