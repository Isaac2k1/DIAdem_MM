'-------------------------------------------------------------------------------
'-- VBS script file
'-- Created on 2007-01-17 22:45:37
'-- Author: Kaan Oenen/Mathias Knaak
'-- Comment: Collection of evaluation routines
'   Last Update: 2008-01-01.00.00.00
'   Version: 1.0.0
'   Reviewed:
'-------------------------------------------------------------------------------
'dim test
'test=chnpropget("test/Travel Drive"," wf_start_offset")
'msgbox test
'call y_value_of_time("y_yalue",0.19,"test/Travel Drive","yVal.",0,0,0,0,0)

'call extrempoint("extrema",0.16,0.6,"test/High current S","First",0,0,0,0,0,0)
'call msgbox(chd(1,"evaluations/extrema"))

'msgbox perform_eval("test/VS","1|2|3/4|5|6|7|8/5|9|10")


'call two_channel_math("ergebnis",0.1,0.5,"test/I-Shunt","-","test/TRVx","","")

option explicit







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'function get_var_val
'reads value of variable from "evaluations"-group


Function get_var_val(var_inp)
      If Left(var_inp,4) = "Var:" Then  ' It is a variable
        get_var_val = Mid(var_inp,5,Len(var_inp))
        get_var_val = CHT(1, "evaluations/" & get_var_val)
        'msgbox "value of " & var_inp & " = " & get_var_val
      Else  'It is a number
        get_var_val = val(var_inp)
      End if
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sub run_eval
' calls different evaluation routines for each evaluation

Sub run_eval(eva_ind,output_name,output_unit,eva_inps,scale_params)

  
 


  Select Case eva_ind
    
      Case 1  'signal_arise_time
       
         if perform_eval(eva_inps(2),eva_inps(5)) then
           Call sign_arise_time(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),eva_inps(4),eva_inps(5),eva_inps(6),eva_inps(7))    
         end if   


      Case 2  'level_cross_time
       
         if perform_eval(eva_inps(2),eva_inps(7)) then
           Call level_cross_time(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),get_var_val(eva_inps(4)),"",eva_inps(6),"","","")    
         end if         

      Case 3  'y_value_of_time
      
        if perform_eval(eva_inps(2),eva_inps(7)) then
          Call y_value_of_time(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),get_var_val(eva_inps(3)),eva_inps(4),"","","")    
        end if         
      
      Case 4  'Extremum: time value
      
        if perform_eval(eva_inps(2),eva_inps(7)) then
           Call t_extrempoint(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),"",eva_inps(6),"","","")    
        end if         
      
      Case 5  'Extremum: y-value
      
        if perform_eval(eva_inps(2),eva_inps(7)) then
          Call y_extrempoint(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),"","",eva_inps(6),"","","")    
        end if         
      
      Case 7  'Slope of Secant
      
        if perform_eval(eva_inps(2),eva_inps(4)) then
          Call slope_of_secant(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),"","","")    
        end if         
     
      Case 8  'Slope of Tangent
      
        if perform_eval(eva_inps(2),eva_inps(8)) then
          Call slope_of_tangent(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),eva_inps(4),eva_inps(5),eva_inps(6),eva_inps(7),eva_inps(8),eva_inps(9),eva_inps(10))    
        end if       
     
      
      Case 9 'Regression Line (Slope)

        if perform_eval(eva_inps(2),eva_inps(5)) then
          call regression_line(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),eva_inps(5),"","","Slope")    
        end if

      Case 10 'Regression Line (Y-Cut)

        if perform_eval(eva_inps(2),eva_inps(5)) then
          call regression_line(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),eva_inps(5),"","","Y_Cut")    
        end if


      Case 11 'Regression Line (RegrPrec)

        if perform_eval(eva_inps(2),eva_inps(5)) then
          call regression_line(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),eva_inps(5),"","","RegrPrec")    
        end if


      Case 12 'y-value of plateau
      
        if perform_eval(eva_inps(2),eva_inps(6)) then
          Call find_plateau(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),get_var_val(eva_inps(4)),eva_inps(5),"","","")    
        end if         
      
      Case 13 ' Average Value
      
        if perform_eval(eva_inps(2),eva_inps(4)) then
          Call average(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),"","","")    
        end if         
      
      Case 14 ' Effective Value
      
        if perform_eval(eva_inps(2),eva_inps(4)) then
          Call rms_value(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),"","","")    
        end if         
      
      Case 15 ' integral Value
      
        if perform_eval(eva_inps(2),eva_inps(6)) then
          Call integral_value(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),eva_inps(5),"","","")    
        end if         
      
      Case 31 ' channel differential
      
        if perform_eval(eva_inps(2),eva_inps(6)) then
          Call channel_differential(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),get_var_val(eva_inps(5)),"")    
        end if      
            
      Case 32 ' channel integral
      
        if perform_eval(eva_inps(2),eva_inps(6)) then
          Call channel_integral(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),get_var_val(eva_inps(5)),"")    
        end if         
      
      Case 34 ' channel smoothing
       
         if perform_eval(eva_inps(2),eva_inps(5)) then
          Call channel_smoothing(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),eva_inps(5))    
         end if     
         
       Case 35 ' remove outlier
       
         if perform_eval(eva_inps(2),eva_inps(4)) then
          Call remove_outlier(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),eva_inps(4))
         end if           
      
      Case 36 ' Two channel Math
       
         if perform_eval(eva_inps(2),eva_inps(5)) then
          Call two_channel_math(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),eva_inps(5))    
         end if           
      
      Case 37 ' One channel Math
       
         if perform_eval(eva_inps(2),eva_inps(5)) then
          Call one_chn_math(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),eva_inps(5),eva_inps(6))    
         end if         


      Case 38 'Fourier Transformation

        if perform_eval(eva_inps(2),eva_inps(7)) then
          call chn_fft(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),get_var_val(eva_inps(5)),eva_inps(6),eva_inps(7))
        end if  

      Case 51 ' Two Variable Math
       
         if perform_eval("",eva_inps(3)) then
          Call two_var_math(output_name,get_var_val(eva_inps(0)),eva_inps(1),get_var_val(eva_inps(2)),eva_inps(3),eva_inps(4),eva_inps(5),eva_inps(6))    
         end if         

      
      case 52 'One Variable Math
      
         if perform_eval("",eva_inps(2)) then
          Call one_var_math(output_name,eva_inps(0),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),eva_inps(4),eva_inps(5))    
         end if         

      
      
      Case 53 ' T-Zero

       Call t_zero(output_name,eva_inps(0))

      Case 54 ' T-Max

       Call t_max(output_name,eva_inps(0))

      case 55 ' cut channel

         if perform_eval(eva_inps(0),eva_inps(3)) then
          Call cut_chn(output_name,eva_inps(0),get_var_val(eva_inps(1)),get_var_val(eva_inps(2)),eva_inps(3))    
         end if         




     Case 61  'document Variable: time
         if perform_eval("",eva_inps(1)) then
          Call docu_time(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 62  'document Variable: y_value
         if perform_eval("",eva_inps(1)) then
          Call docu_y_val(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 63 'document Variable: string
         if perform_eval("",eva_inps(1)) then
          Call docu_string(output_name,eva_inps(0),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 64  'document Variable: number
         if perform_eval("",eva_inps(1)) then
          Call docu_number(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 65  'select Document string
         if perform_eval("",eva_inps(1)) then
          Call sel_doc_string(output_name,eva_inps(0),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         




     Case 66  'runtime Variable: time
         if perform_eval("",eva_inps(1)) then
          Call runtime_time(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 67  'runtime Variable: y_Value
         if perform_eval("",eva_inps(1)) then
          Call runtime_y_val(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 68  'runtime Variable: Integer
         if perform_eval("",eva_inps(1)) then
          Call runtime_integer(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),eva_inps(3))    
         end if         

    Case 70  'Code dependent Integer
         if perform_eval("",eva_inps(2)) then
          Call code_dep_integer(output_name,eva_inps(0),eva_inps(1),eva_inps(2),eva_inps(3),eva_inps(4))    
         end if         
    
    Case 71  'Code dependent Integer
         if perform_eval("",eva_inps(2)) then
          Call code_dep_integer(output_name,eva_inps(0),eva_inps(1),eva_inps(2),eva_inps(3),eva_inps(4))    
         end if  

    Case 72  'Code dependent string
         if perform_eval("",eva_inps(2)) then
          Call code_dep_string(output_name,eva_inps(0),eva_inps(1),eva_inps(2),eva_inps(3),eva_inps(4))    
         end if  

    Case 82 'export to Excel
         if perform_eval("",eva_inps(3)) then 
         call excel_export(output_name,eva_inps(0),eva_inps(1),eva_inps(2))
        end if

    Case 83 'export TDM
         if perform_eval("",eva_inps(2)) then 
         call tdm_export(output_name,eva_inps(0),eva_inps(1))
        end if


  End Select

'write output unit
  'if output_unit <> "" then 
  call chnpropset("evaluations/"&output_name,"unit_string",output_unit)




'if autoscaling in sw-channels is switched off, cut channel to selected y-range

'call msgbox(scale_params(0)&"  "&scale_params(1)&"  "&scale_params(2))

if scale_params(0)="0" then
  call msgbox(scale_params(0)&"  "&scale_params(1)&"  "&scale_params(2))
  dim range_start,range_end
  range_start=scale_params(2)-0.5*scale_params(1)
  range_end=scale_params(2)+0.5*scale_params(1)
  call cut_chn_range(range_start,range_end,"evaluations/"&output_name)
end if


End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Sub TwoChnOperation(output_name,chn1,math_op,chn2)

  Select case math_op
    
    Case "+"  'Addition
      
      'Use DIAdem function for addition
      Call CHNADD(chn1,chn2,"evaluations/"&output_name) '... Y,CALCYCHN,E 
      
      Msgbox "Summation of two channels done!" &Vbcrlf&_
              chn1& " + " &chn2& " = " &output_name
    
    
    Case "-"  'Subtraction
      
      Call CHNSUB(chn1,chn2,"evaluations/"&output_name) '... Y,Y1,E 
      
      Msgbox "Substraction of two channels done!" &Vbcrlf&_
              chn1& " - " &chn2& " = " &output_name

    Case "*"    'Multiplication    
      
      Call CHNMUL(chn1,chn2,"evaluations/"&output_name) '... Y,CALCYCHN,E       

      Msgbox "Multiplication of two channels done!" &Vbcrlf&_
              chn1& " * " &chn2& " = " &output_name
      
    Case "/"    'Divide
      
      Call CHNDIV(chn1,chn2,"evaluations/"&output_name)

  End Select



End Sub


Sub ChnScale(result_chn,input_chn,chn_scal_fct,chn_off)
   chn_scal_fct = Val(chn_scal_fct)
   chn_off = Val(chn_off)
  Call CHNLINSCALE(input_chn,"evaluations/"&result_chn,chn_scal_fct,chn_off) '... Y,E,CHNSCALEFACTOR,CHNSCALEOFFSET 

  'Msgbox input_chn&" is scaled with factor "&chn_scal_fct& " and an Offset of "&chn_off& "has been added"&VbcrLf&_
   '      "Result channel is: "&result_chn 

End Sub

Sub Example3(inp1,inp2,inp3)

  Call CHNOFFSET(inp2,"/"&inp1,inp3,"mean value offset") '... Y,E,CHNOFFSETVALUE,CHNOFFSETMODE 

End Sub

Sub MaxValue(result_var,input_chn)
  Call UIAutoRefreshSet(True)
  Dim max_val, min_val, absolute
  max_val = ChnPropGet(input_chn, "Maximum")
  min_val = ChnPropGet(input_chn, "Minimum")
  msgbox "min_val = " & min_val
  msgbox "max_val = " & max_val
  max_val = val(Replace(max_val, "." , ","))
  min_val = val(Replace(min_val, "." , ","))
  If abs(min_val) > abs(max_val) Then
    absolute = abs(min_val)
  Else
    absolute = abs(max_val) 
  End if
  CHT(1,"evaluations/" & result_var ) = absolute
  Call UIAutoRefreshSet(False)
End Sub

Sub ChnDivide(result_chn,input_chn,chn_scal_fct)
   'msgbox "factor = " & chn_scal_fct
  chn_scal_fct = 1.0/Val(chn_scal_fct)
 
  Call CHNLINSCALE(input_chn,"evaluations/"&result_chn,chn_scal_fct,0) '... Y,E,CHNSCALEFACTOR,CHNSCALEOFFSET 
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Basic evaluations
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub sign_arise_time
'
' Mathias Knaak
' 19.09.2008
' desription: searches for the beginning uo a signal.
' INPUT:(result_var,byval t_start,byval t_end,input_chn,signal_noise,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel in evaluations-Group
't_start, t_end: time for start and end of search in the channel, if t_start > t_end use backward search
'input_chn: Input channel
'signal noise: percentage of signal range to define signal arise point
'mark_plot: mark the result on plot

sub sign_arise_time(result_var,byval t_start,byval t_end,input_chn,signal_noise,mark_plot,when_eva, when_plot,when_report)

dim x1,x2,xref,xbeg,noise,x1_step,x2_step,limit,xbeg_step,count,input_chn_no, x_end, x_end_step

'convert variables to values
t_start=val(t_start)
t_end=val(t_end)

'x1: time so begin search
x1=t_start
'x_end: maximum value to search
x_end=t_end


'create channel group for calculations
if groupindexget("calculations")<>0 then call groupdel(groupindexget("calculations"))

call groupcreate("calculations",groupcount+1)

'copy input channel to calculations group
call chncopy(input_chn,"calculations/input")
'set default group (to use chdx-function)
call groupdefaultset(groupindexget("calculations"))
input_chn_no=cno("input")

'define noise level: percentage of signal_noise times peak-to-peak value of input channel
noise=signal_noise/100*(cmax("calculations/input")-cmin("calculations/input"))


'convert time values into sample numbers
xbeg=x1
x1_step=wf_step(x1,"calculations/input")
x_end_step=wf_step(x_end,"calculations/input")
xbeg_step=x1_step

'begin forward search
if t_start < t_end then
            
            'calculate average value of the first 20 values
          for count=x1_step to x1_step+20
            limit=limit+chdx(count,input_chn_no)
          next
          limit=limit/20

          'begin at position x1_step and move forward until change of value is > than noise level
          while (abs((chdx(x1_step,input_chn_no)-limit)) < noise) and x1_step <= x_end_step
            x1_step=x1_step+1
          wend
          'set x2_step to new x1_step
          x2_step=x1_step
          'move backward from x2_step until signal change is greater than noise/3
          while (abs(chdx(x2_step,input_chn_no)-limit) > noise/3) and (x2_step > xbeg_step)
            x2_step=x2_step-1
          wend
          
          'get time values from sample numbers
          x1=wf_time(x1_step,"calculations/input")
          x2=wf_time(x2_step,"calculations/input")

          'copy input channel to calculate regression lines
          call chncopy("calculations/input","calculations/reg_1")
          call chncopy("calculations/input","calculations/reg_2")

          'cut channels to selected time range for regression lines

          call select_timerange("calculations/reg_1",xbeg,x2)
          call select_timerange("calculations/reg_2",x2,x1)

  elseif t_start > t_end then 'use bachward seach, same funtionality in other respects as desribed above

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

  '        msgbox "x1: "&x1&"x2: "&x2

          call chncopy("calculations/input","calculations/reg_1")
          call chncopy("calculations/input","calculations/reg_2")

          call select_timerange("calculations/reg_1",x2,xbeg)
          call select_timerange("calculations/reg_2",x1,x2)

  else 
          msgbox "t_start=t_end! End of function."

end if

'calculate two regression lines between xbeg until x2 and x2 until x1
'save slope and y_cut


'calculate regression
call chnregrXYcalc("","calculations/reg_1","","calculations/reg_1","linear","Partition complete area",1000,1)

dim y_cut_1,y_cut_2,slope_1,slope_2, arisepoint

y_cut_1=regrcoeffa
slope_1=regrcoeffb

'calculate regression
call chnregrXYcalc("","calculations/reg_2","","calculations/reg_2","linear","Partition complete area",1000,1)

y_cut_2=regrcoeffa
slope_2=regrcoeffb

'calculate intersection point of both regression lines, round value
arisepoint=(trunc(((y_cut_2-y_cut_1)/(slope_1-slope_2)*100000)+0.5))/100000

'msgbox arisepoint

'call chnalloc("arisepoint_x",1,,,,groupindexget("calculations"))
chd(1,"evaluations/"&result_var)=arisepoint

'add marker point
if mark_plot="True" then call add_marker(arisepoint,chd_wf_time(arisepoint,input_chn),"time",input_chn)

'delete calculations group
call groupdel(groupindexget("calculations"))

end sub









''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub level_cross_time
'
' Mathias Knaak
' 31.07.08 
'
'Description: finds the specified crossing-point of an input channel with a given y-Value. 
'
'INPUT:(result_var,t_start,t_end,input_chn,y_level,cross_num, sig_noise,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel in evaluations-Group
't_start, t_end: time for start and end of search in the channel
'input_chn: Input channel
'y_level: y-level to find
'cross_num: number of crossing to find
'all other inputs are not yet implemented

'OUTPUT: channel in "evaluations"-Group with the name of result_var, time of found crossing_point in the first value
'
'
'

sub level_cross_time(result_var,t_start,t_end,input_chn,y_level,cross_num, sig_noise,mark_plot,when_eva, when_plot,when_report)

dim step_val, cross_count    'step_val: variable to count steps in the for-loop
cross_count=0                'cross_count: number of found crossing 

'msgbox t_start


'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,,,,,groupindexget("evaluations"))

t_start=val(t_start)
t_end=val(t_end)

'search crossing point
'begin forward search
if t_start < t_end then

    'look at every value in given time range
      for step_val = wf_step(t_start,input_chn) to wf_step(t_end,input_chn)+1
            'if level of input_channel crosses the given Y-Value...
            if (((chd(step_val, input_chn)<=y_level) and (chd(step_val+1, input_chn)>y_level)) or ((chd(step_val, input_chn)>=y_level) and (chd(step_val+1, input_chn)<y_level))) then
                '...crossing point found
                cross_count=cross_count+1 'increase number of found crossing points
                if cross_count=cross_num then 'if crossing found, write results to result channel
                  CHD(1,"evaluations/"&result_var) =  wf_time(step_val, input_chn)
                  exit for            
                end if 'cross_count=cross_num then
            end if   '(((chd(step_val, input_chn)<=y_level) [...]
      next

end if

'begin reverse search if t_start > t_end

if t_start > t_end then

    'look at every value in given time range
          for step_val = wf_step(t_start,input_chn)+1 to wf_step(t_end,input_chn) step -1
                if (((chd(step_val, input_chn)<=y_level) and (chd(step_val+1, input_chn)>y_level)) or ((chd(step_val, input_chn)>=y_level) and (chd(step_val+1, input_chn)<y_level))) then
                    'crossing point found
                    cross_count=cross_count+1
                    if cross_count=cross_num then
                      CHD(1,"evaluations/"&result_var) =  wf_time(step_val, input_chn)
                      exit for            
                    end if
                end if
          next

end if

if mark_plot="True" then call add_marker(wf_time(step_val, input_chn),y_level,"time",input_chn)



end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub y_value_of_time
'
'Mathias Knaak
'4.08.2008
'
'description: returns the y-value of a function at a given time, select value, 1st. or 2nd. derivation
'
'INPUT:(result_var,input_time,input_chn,y_type,fit_points,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
'input_time: time of level to find
'input_chn: Input data
'fit_points: number of points for approximation
'all other inputs are not yet implemented

sub y_value_of_time(result_var,input_time,input_chn,y_type,fit_points,mark_plot,when_eva, when_plot,when_report)
dim step_val

'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,,,,,groupindexget("evaluations"))

call chncopy (input_chn,"calculations/diff_chn") 

if fit_points>0 then  'use approximation
  dim sample_count, start_step, chn_length
  'calculate first value for approximation
  start_step=trunc(wf_step(input_time,"calculations/diff_chn")-0.5*fit_points+0.5)
  call chnalloc("approx_points",chnlength("calculations/diff_chn"),,,,groupindexget("calculations"))
    sample_count=1
    'write novalues at the beginning of the channel
    for sample_count=1 to start_step-1
       chd(sample_count,"calculations/approx_points")=novalue
    next
    
    'write values to fit curve
    sample_count=1
    for sample_count=1 to fit_points
      chd(start_step,"calculations/approx_points")=chd(start_step,"calculations/diff_chn")
      start_step=start_step+1
    next

    'write novalues at the end of the channel
    chn_length=chnlength("calculations/diff_chn")
    for sample_count=start_step to chn_length
       chd(sample_count,"calculations/approx_points")=novalue
    next
'convert channel to Waveform-Channel
call chnwfpropset("calculations/approx_points","zeit","s",val(chnpropget(input_chn, "wf_start_offset")),val(chnpropget(input_chn, "wf_increment")))

'approximate channel:
'use quadratic-function for approximation (0-, 1st.- and 2nd.-order function) 
Call ApprAnsatzOff             
ApprAnsatzFct(1) ="Yes"
ApprAnsatzFct(2) ="Yes"
ApprAnsatzFct(3) ="Yes"

Call ChnApprXYCalc("","calculations/approx_points","","calculations/diff_chn","Partition complete area",len(input_chn),1) '... XW,Y,E,E,XChnStyle,XNo,XDiv 

'else 
'  call chncopy (input_chn,"calculations/diff_chn") 

end if    'fit_points>0 then



  select case y_type 'select type of result value, derivate channel if neccessary

      case "Slope"
          call Chndifferentiate (,"calculations/diff_chn",,"calculations/diff_chn")
      case "Curv."
          call Chndifferentiate (,"calculations/diff_chn",,"calculations/diff_chn")
          call Chndifferentiate (,"calculations/diff_chn",,"calculations/diff_chn")
  end select

'step_val=wf_step(input_time, input_chn)

'save result_value in result-channel
chd(1,"evaluations/"&result_var)=chd_wf_time(input_time,"calculations/diff_chn") 

call groupdel(groupindexget("calculations"))

if mark_plot="True" then call add_marker(input_time,chd_wf_time(input_time,input_chn),"value",input_chn)


end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub t_extrempoint
'
'find the time of extremum in a channel, type of extremum ist selectable
'


'Mathias Knaak
'5.08.2008
'INPUT:(result_var,t_start,t_end,input_chn,extrema_type,sig_noise,fit_points,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'extrema_type: type of extremum
'all other inputs are not yet implemented

sub t_extrempoint(result_var,t_start,t_end,input_chn,extrema_type,sig_noise,fit_points,mark_plot,when_eva, when_plot,when_report)
dim result
'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

'copy channel for calculations
call chncopy(input_chn,"calculations/extrema")

'reduce channel to selected timerange
call select_timerange("calculations/extrema",t_start,t_end)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,,,,,groupindexget("evaluations"))

select case extrema_type  'select time of extrema

  case "First"  'first extremum
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",1,"Max.Peaks","Time")
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",1,"Min.Peaks","Time")
    result=valmin(chd(1,"calculations/Peak_X_pos_time"),chd(1,"calculations/Peak_X_neg_time"))

  case "First Pos." 'first positive extremum
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",1,"Max.Peaks","Time")
    result=chd(1,"calculations/Peak_X_pos_time")
  
  case "First Neg." 'first negative extremun
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",1,"Min.Peaks","Time")
    result=chd(1,"calculations/Peak_X_neg_time")
  
  case "First Abs." 'maximum of "first pos." and "first neg."
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",10,"Max.Peaks","Time")
    Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",10,"Min.Peaks","Time")
      if  chd(1,"calculations/Peak_Y_pos_time") > chd(1,"calculations/Peak_Y_neg_time") then
            result=chd(1,"calculations/Peak_X_pos_time")
      elseif chd(1,"calculations/Peak_Y_pos_time") = chd(1,"calculations/Peak_Y_neg_time") then
          result=valmin(chd(1,"calculations/Peak_X_pos_time"),chd(1,"calculations/Peak_X_neg_time"))
      else
          result=chd(1,"calculations/Peak_X_neg_time")
      end if
  
  case "Highest Pos." 'highest pos. extremum
       Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_X_pos_amplitude") 
  
  case "Highest Neg."  'highest neg. extremum
       Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_X_neg_amplitude")  
  
  case "Highest Abs." 'maximum of highest pos. and neg.
          Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude")  
          Call ChnPeakFind("","calculations/extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
      if   chd(1,"calculations/Peak_Y_pos_amplitude") > chd(1,"calculations/Peak_Y_neg_amplitude") then
            result=chd(1,"calculations/Peak_X_pos_amplitude")
      elseif chd(1,"calculations/Peak_Y_pos_amplitude") = chd(1,"calculations/Peak_Y_neg_amplitude") then
          result=valmin(chd(1,"calculations/Peak_X_pos_amplitude"),chd(1,"calculations/Peak_X_neg_amplitude"))
      else
          result=chd(1,"calculations/Peak_X_neg_amplitude")
      end if
   
end select

chd(1,"evaluations/"&result_var)=result

if mark_plot="True" then call add_marker(result,chd_wf_time(result,input_chn),"time",input_chn)

call groupdel(groupindexget("calculations"))

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub y_extrempoint
'
'find y-value of extrema in a channel, almost the same function as t_extrempoint; return value y_value, not time
'Mathias Knaak
'5.08.2008
'INPUT:(result_var,t_start,t_end,input_chn,extrema_type,sig_noise,fit_points,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'extrema_type: type of extremum
'all other inputs are not yet implemented

sub y_extrempoint(result_var,t_start,t_end,input_chn,extrema_type,sig_noise,fit_points,mark_plot,when_eva, when_plot,when_report)
dim result,result_x
'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/y_extrema")

call select_timerange("calculations/y_extrema",t_start,t_end)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,,,,,groupindexget("evaluations"))

select case extrema_type

  case "First"
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",1,"Max.Peaks","Time")
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",1,"Min.Peaks","Time")
    if chd(1,"calculations/Peak_X_pos_time")< chd(1,"calculations/Peak_X_neg_time") then
      result=chd(1,"calculations/Peak_Y_pos_time")
      result_x=chd(1,"calculations/Peak_X_pos_time")    
    else    
      result=chd(1,"calculations/Peak_Y_neg_time")
      result_x=chd(1,"calculations/Peak_X_neg_time")
    end if

  case "First Pos."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",1,"Max.Peaks","Time")
    result=chd(1,"calculations/Peak_Y_pos_time")
    result_x=chd(1,"calculations/Peak_X_pos_time")

  case "First Neg."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",1,"Min.Peaks","Time")
    result=chd(1,"calculations/Peak_Y_neg_time")
    result_x=chd(1,"calculations/Peak_X_neg_time")

  case "First Abs."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",10,"Max.Peaks","Time")
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",10,"Min.Peaks","Time")
      if  chd(1,"calculations/Peak_Y_pos_time") > chd(1,"calculations/Peak_Y_neg_time") then
            result=chd(1,"calculations/Peak_Y_pos_time")
            result_x=chd(1,"calculations/Peak_X_pos_time")
      elseif chd(1,"calculations/Peak_Y_pos_time") = chd(1,"calculations/Peak_Y_neg_time") then
          result=chd(1,"calculations/Peak_Y_pos_time")
          result_x=chd(1,"calculations/Peak_X_pos_time")
      else
          result=chd(1,"calculations/Peak_Y_neg_time")
          result_x=chd(1,"calculations/Peak_X_neg_time")
      end if
  
  case "Highest Pos."
       Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_Y_pos_amplitude") 
       result_x=chd(1,"calculations/Peak_X_pos_amplitude") 
  case "Highest Neg."
       Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_Y_neg_amplitude")  
       result_x=chd(1,"calculations/Peak_X_neg_amplitude")  
  case "Highest Abs."
          Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude")  
          Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
      if   chd(1,"calculations/Peak_Y_pos_amplitude") > chd(1,"calculations/Peak_Y_neg_amplitude") then
            result=chd(1,"calculations/Peak_Y_pos_amplitude")
            result_x=chd(1,"calculations/Peak_X_pos_amplitude")
      elseif chd(1,"calculations/Peak_Y_pos_amplitude") = chd(1,"calculations/Peak_Y_neg_amplitude") then
          result=chd(1,"calculations/Peak_Y_pos_amplitude")
          result_x=chd(1,"calculations/Peak_X_pos_amplitude")
      else
          result=chd(1,"calculations/Peak_Y_neg_amplitude")
          result_x=chd(1,"calculations/Peak_X_neg_amplitude")
      end if
   
end select

chd(1,"evaluations/"&result_var)=result

if mark_plot="True" then call add_marker(result_x,result,"value",input_chn)

call groupdel(groupindexget("calculations"))
end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub slope_of_secant
'
'Mathias Knaak
'5.08.2008
'finds the slope of a secant, defined by two time values and an imput channel

'INPUT:(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'all other inputs are not yet implemented


sub slope_of_secant(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,,,,,groupindexget("evaluations"))  
dim result
'calculate differential with delta y / delta x
result= chd_wf_time(t_end,input_chn)-chd_wf_time(t_start,input_chn)
result=result/(t_end-t_start)

'write result
chd(1,"evaluations/"&result_var)=result

if mark_plot="True" then
    call add_marker(t_start,chd_wf_time(t_start,input_chn),"time",input_chn)
    call add_marker(t_end,chd_wf_time(t_end,input_chn),"time","")
    call add_marker_line(t_start,chd_wf_time(t_start,input_chn),t_end,chd_wf_time(t_end,input_chn),input_chn)
end if


end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sub slope_of_tangent(result_var,byval t_start,byval t_end,input_chn,byval ref_time,tangent_type,true_tangent,signal_noise,mark_plot,when_eva, when_plot,when_report)
'
'Mathias Knaak
'22.09.2008
'decription: find the slope of a tangent on an function. the tangent will cross the input_channel at ref_time and touch the function 
' in the time range between t_start and t_end
'Input: 
'result_var: name of result_variable
't_start,t_end: search range
'input_channel
'ref_time: intersection between tangent and input_channel
'tangent_type: "First, First Positiv, First Negativ, Second, Maximum"
'true tangent: a true tangent may not touch the channel at t_end
'signal noise: not used
'mark plot: mark the result on the plot
sub slope_of_tangent(result_var,byval t_start,byval t_end,input_chn,byval ref_time,tangent_type,true_tangent,signal_noise,mark_plot,when_eva, when_plot,when_report)

dim x_start, x_end,x_act,y_ref,y_act,x_ref,x_tan_1,x_tan_2,slope_1,slope_2, ref_slope,chn_number,act_slope
dim result,result_x

slope_1=0
slope_2=0

result=0
result_x=0



'create group for calculations
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/input")

call groupdefaultset(groupindexget("calculations"))

'x_ref: Sample number of reference time
x_ref=wf_step(ref_time,"calculations/input")

'y_ref: y_value of reference time
y_ref=chd(x_ref,"calculations/input")

'x_start,x_end: sample number of begin and end of search range
x_start=wf_step(t_start,"calculations/input")
x_end=wf_step(t_end,"calculations/input")

'ref_slope: slope of secant between x_ref and x_start. 
ref_slope=(chd(x_start,"calculations/input")-y_ref)/(x_start-x_ref)

'begin search for tangents one step behind x_start, x_start may not be point of the tangent
x_start=x_start+1

'a true tangent may not go through x_end
if true_tangent="True" then  x_end=x_end-1

'get channel number for chdx-function
chn_number=cno("input")

'set position of tangent behind the end of search range
x_tan_1=x_end+1
x_tan_2=x_end+1

'search complete range between t_start and t_end
for x_act=x_start to x_end 
  'get y_value at actual position
  y_act=chdx(x_act,chn_number)
  'get slope of sekant between x_ref and actual position (x_act)
  act_slope=(y_act-y_ref)/(x_act-x_ref)
  'save slope as tangent if slope is steeper than reference slope and saved slope 
  if (act_slope > ref_slope) and (act_slope > slope_1 ) then
    slope_1=act_slope
    x_tan_1=x_act
  'save negative slope
  elseif (act_slope < ref_slope) and (abs(act_slope) > abs(slope_2) ) then
    slope_2=act_slope
    x_tan_2=x_act
  end if
next

select case tangent_type 'select type of slope for the result

    case "First"  'first tangent
      if (x_tan_1 < x_tan_2) then 
        result=(chd(x_tan_1,"calculations/input")-y_ref)/(wf_time(x_tan_1,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_1,"calculations/input")
      elseif x_tan_1 > x_tan_2 then 
        result=(chd(x_tan_2,"calculations/input")-y_ref)/(wf_time(x_tan_2,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_2,"calculations/input")
      end if

    case "F-Pos"  'first tangent steeper than reference tangent
        result=(chd(x_tan_1,"calculations/input")-y_ref)/(wf_time(x_tan_1,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_1,"calculations/input")
    case "F-Neg"  'first tangent under the reference tangent
        result=(chd(x_tan_2,"calculations/input")-y_ref)/(wf_time(x_tan_2,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_2,"calculations/input")
    case "Secnd"  'second tangent
      if (x_tan_1 > x_tan_2) and (x_tan_1 < x_end) then 
        result=(chd(x_tan_1,"calculations/input")-y_ref)/(wf_time(x_tan_1,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_1,"calculations/input")
      elseif (x_tan_1 < x_tan_2) and (x_tan_2 < x_end) then 
        result=(chd(x_tan_2,"calculations/input")-y_ref)/(wf_time(x_tan_2,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_2,"calculations/input")
      end if

    case "Maxim"  'tangent with maximum slope
      if slope_1 > abs(slope_2) then 
        result=(chd(x_tan_1,"calculations/input")-y_ref)/(wf_time(x_tan_1,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_1,"calculations/input")
      else
        result=(chd(x_tan_2,"calculations/input")-y_ref)/(wf_time(x_tan_2,"calculations/input")-ref_time)
        result_x=wf_time(x_tan_2,"calculations/input")
      end if

end select

'write result
chd(1,"evaluations/"&result_var)=result

'write marker
if mark_plot="True" then
    call add_marker(ref_time,chd_wf_time(ref_time,input_chn),"time",input_chn)
    call add_marker(result_x,chd_wf_time(result_x,input_chn),"time","")
    call add_marker_line(ref_time,chd_wf_time(ref_time,input_chn),result_x,chd_wf_time(result_x,input_chn),"")
end if
'delete calculations group
call groupdel(groupindexget("calculations"))


end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub regression_line(result_var,t_start,t_end,input_chn,regress_type,mark_plot,when_eva,when_plot,when_report,result_type)
' Mathias Knaak
' 17.09.2008
' description: calculates the regression from an input channel
'Input:
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'regress_type: Linear,Logarithmisch,Exponentiell, Power
'mark_plot: Mark result on plot
'result_type: Slope / Y_Cut / RegrPrec

sub regression_line(result_var,t_start,t_end,input_chn,regress_type,mark_plot,when_eva,when_plot,when_report,result_type)

'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/input")
call select_timerange("calculations/input",t_start,t_end)

dim regr_type 'Type of regression
dim result 'result value

select case regress_type 'select type of regression

  case "Lin"
    regr_type="linear"

  case "Log"
    regr_type="logarith.n"

  case "Exp"
    regr_type="exponential"

  case "Pwr"
    regr_type="power"

end select

'calculate regression
call chnregrXYcalc("","calculations/input","","calculations/regr_curve",regr_type,"Partition complete area",1000,1)

select case result_type

  case "Slope"
    result=RegrcoeffB

  case "Y_Cut"
    result=RegrcoeffA

  case "RegrPrec"
    result=100*RegrPrecision

end select

'copy result to evaluations group
chd(1,"evaluations/"&result_var)=result

'copy result channel to evaluations group
if mark_plot then
  call chncopy("calculations/regr_curve","evaluations/regression curve")
end if

call groupdel(groupindexget("calculations"))

end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub find_plateau
'
'Mathias Knaak
'6.08.2008
'
'Desription: this evaluation finds the longest plateau in a given time range
'return value is mean y-value of the found plateau; additional values give start and end time of the plateau
'
'INPUT:(result_var,t_start,t_end,input_chn,class_num,sig_noise,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'class_num: Number of classes to divide y-range
'all other inputs are not yet implemented
'
'

sub find_plateau(result_var,t_start,t_end,input_chn,class_num,sig_noise,mark_plot,when_eva, when_plot,when_report)

dim int_found_start,int_found_end,int_found_y,int_found_length 'Values of found maximum Plateau
dim limit,range, end_step, count

'width of y-class
range=0.5*((chnpropget(input_chn,"maximum")-chnpropget(input_chn,"minimum"))/class_num)

'only y-values above "limit" are used to find plateau; time periods at minimum y-value will be ignored
limit=0.25*(chnpropget(input_chn,"maximum")-chnpropget(input_chn,"minimum"))+chnpropget(input_chn,"minimum")

dim act_pos,int_act_start,int_act_end, int_act_y  'values of actual plateau

int_act_start=wf_step(t_start,input_chn) 'begin search at t_start
act_pos=int_act_start
int_found_start=act_pos
end_step=wf_step(t_end,input_chn)
int_found_y=0

while act_pos < end_step  'search from t_start until t_end
  act_pos=act_pos+1  
  
  'check every value in the channel, if y_value at act_pos differs from y_value at start position
  if abs(chd(act_pos,input_chn)-chd(int_act_start,input_chn))> range then
    if (act_pos - int_act_start) > int_found_length then 'if Plateau is longer than the old one
      'calculate mean y-value in plateau range
      int_act_y=0
      for count=int_act_start to act_pos-1
        int_act_y=int_act_y+chd(count,input_chn)
      next  'count=int_act_start to act_pos-1
      int_act_y=int_act_y/(act_pos -int_act_start)
      
      if int_act_y > limit then 'if y-Value ist above limit...
        '...save values of found plateau 
        int_found_start=int_act_start
        int_found_end=act_pos-1
        int_found_length=int_found_end - int_found_start
        int_found_y=int_act_y
      end if  'int_act_y > limit then
    end if
  int_act_start=act_pos  'begin next search at actual position
  'int_start_y=chd(act_pos,input_chn)
  end if    ' if abs(chd(act_pos,input_chn)-chd(int_act_start,input_chn))> range then

wend    'act_pos < end_step

'call msgbox(wf_time(int_found_start,input_chn)&" , "&wf_time(int_found_end,input_chn))

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,4,,,,groupindexget("evaluations"))  

chd(1,"evaluations/"&result_var)=int_found_y
'chd(2,"evaluations/"&result_var)=wf_time(int_found_start,input_chn)
'chd(3,"evaluations/"&result_var)=wf_time(int_found_end,input_chn)

if mark_plot="True" then

'call    add_marker(wf_time(int_found_start,input_chn),int_found_y,"value",input_chn)
'call    add_marker(wf_time(int_found_end,input_chn),int_found_y,"value",input_chn)
call add_marker_line(wf_time(int_found_start,input_chn),int_found_y,wf_time(int_found_end,input_chn),int_found_y,input_chn)
end if



end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub average
'
'Mathias Knaak
'7.08.2008
' calculates the average of y-values in defined time range
'
''INPUT:(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'all other inputs are not yet implemented
'
'
sub average(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)

dim start, stop_count ,average_value ,count

average_value=0
'define begin and end of range
start=wf_step(t_start,input_chn)
stop_count=wf_step(t_end,input_chn)

'summarize all values in given range
for count=start to stop_count
  average_value=average_value+chd(count,input_chn)
next
'divide 
average_value=average_value/(1+stop_count-start)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,4,,,,groupindexget("evaluations"))  
'save result
chd(1,"evaluations/"&result_var)=average_value

if mark_plot="True" then

'call    add_marker(t_start,average_value,"value",input_chn)
'call    add_marker(t_end,average_value,"value",input_chn)
call    add_marker_line(t_start,average_value,t_end,average_value,input_chn)
end if


end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub rms_value
'
'Mathias Knaak
'7.08.08
'
' calculates the rms-value of a signal in given time-range
'Der Effektivwert wird mit der Integral-Methode nach Simpson berechnet, Algorithmus wie in rebadas
' 
''INPUT:(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'all other inputs are not yet implemented

sub rms_value(result_var,t_start,t_end,input_chn,mark_plot,when_eva, when_plot,when_report)


dim start, stop_count ,rms_val ,count

rms_val=0
'get step values in the input channel for given time values
start=wf_step(t_start,input_chn)
stop_count=wf_step(t_end,input_chn)-1

for count=start to stop_count 'loop over all values in the given time range
  'Summe über y(n)^2 + y(n)*y(n+1) + y(n+1)^2
  rms_val=rms_val+chd(count,input_chn)*chd(count,input_chn)+chd(count,input_chn)*chd(count+1,input_chn)+chd(count+1,input_chn)*chd(count+1,input_chn)
next

'calculate square root and divide by time range
rms_val=sqrt(rms_val/(1+stop_count-start)/3)

'call msgbox(rms_val)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,4,,,,groupindexget("evaluations"))  

chd(1,"evaluations/"&result_var)=rms_val

if mark_plot="True" then

'call    add_marker(t_start,rms_val,"value",input_chn)
'call    add_marker(t_end,rms_val,"value",input_chn)
call    add_marker_line(t_start,rms_val,t_end,rms_val,input_chn)
end if



end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub integral_value
'
'Mathias Knaak
'21.08.08
'
'calculates the integral value of a channel in a given time range
' 
''INPUT:(result_var,t_start,t_end,input_chn,int_type,int_exp,mark_plot,when_eva, when_plot,when_report)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'all other inputs are not yet implemented

sub integral_value(result_var,t_start,t_end,input_chn,int_type,int_exp,mark_plot,when_eva, when_plot,when_report)


dim start, stop_count ,int_val ,count

int_val=0
'determine step values for the time range
start=wf_step(t_start,input_chn)
stop_count=wf_step(t_end,input_chn)-2

'use "1" as exponent, if no value as input
if not isnumeric(int_exp) then int_exp=1


dim channel

channel=cno(input_chn)
call groupdefaultset(chngroup(input_chn))


select case int_type

 case "Lin."

      for count=start to stop_count 'loop over all values in time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+chdx(count,channel)+(4*chdx(count+1,channel))+chdx(count+2,channel)
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+2)/6*(t_end-t_start)


case "Abs."
      
      for count=start to stop_count 'loop over all values in time range
        'Summe über (y(n) + 4*y(n+1) + y(n+2))^int_exp
        int_val=int_val+(abs(chdx(count,channel))+(4*abs(chdx(count+1,channel)))+abs(chdx(count+2,channel)))^int_exp
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+2)/(6^int_exp)*(t_end-t_start)


case "Quad"
      
      for count=start to stop_count 'loop over all values in time range
        'Summe über (y(n) + 4*y(n+1) + y(n+2))^2
        int_val=int_val+(chdx(count,channel)+(4*chdx(count+1,channel))+chdx(count+2,channel))^2
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+2)/36*(t_end-t_start)



end select




'call msgbox(int_val)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,4,,,,groupindexget("evaluations"))  
'write result
chd(1,"evaluations/"&result_var)=int_val

if mark_plot="True" then

'call    add_marker(t_start,int_val,"value",input_chn)
'call    add_marker(t_end,int_val,"value",input_chn)
call    add_marker_line(t_start,int_val,t_end,int_val,input_chn)
end if

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub channel_integral
'
'Mathias Knaak
'21.08.08
'
'
' 
''INPUT:(result_var,t_start,t_end,input_chn,int_type,int_exp,boundary,when_eva)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'int_type: Type of Integral: Linear, Absolut, Quadratisch
'int_exp: Exponent for Absolut Integral
'boundary: initial value for the integral


sub channel_integral(result_var,t_start,t_end,input_chn,int_type,int_exp,boundary,when_eva)


dim start, stop_count ,int_val ,count,time_step,length,wf_offset

'get time steps of input channel
time_step=chnpropvalget(input_chn,"wf_increment")
wf_offset=chnpropvalget(input_chn,"wf_start_offset")
length=cl(input_chn)
'msgbox length
'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then 
  call chnalloc(result_var,length,,,,groupindexget("evaluations"))  
  call chnwfpropset("evaluations/"&result_var,"x","s",wf_offset,time_step)
else
  call chnrealloc("evaluations/"&result_var,length)
  call chnwfpropset("evaluations/"&result_var,"x","s",wf_offset,time_step)
end if

'boundary condition
int_val=boundary
'get step values for time range
start=wf_step(t_start,input_chn)
stop_count=wf_step(t_end,input_chn)

'for count=start to stop_count-2 'loop over complete time range
'  'Summe über y(n) + 4*y(n+1) + y(n+2)
'  int_val=int_val+(chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn))*(time_step/6)
'  'write every calculated integral value into channel
'  chd(count,"evaluations/"&result_var)=int_val
'next

dim channel

channel=cno(input_chn)
call groupdefaultset(chngroup(input_chn))

select case int_type

 case "Lin."

      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+(chdx(count,channel)+(4*chdx(count+1,channel))+chdx(count+2,channel))*(time_step/6)
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next


case "Abs."
      
      for count=start to stop_count-1 'loop over complete time range
        'calculate integral value
        int_val=int_val+((abs(chdx(count,channel))+(4*abs(chdx(count+1,channel)))+abs(chdx(count+2,channel)))^int_exp)*(time_step/(6^int_exp))
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next


case "Quad"
      
      for count=start to stop_count-1 'loop over complete time range
        'calculate integral value
        int_val=int_val+((chdx(count,channel)+(4*chdx(count+1,channel))+chdx(count+2,channel))^2)*(time_step/36)
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next



end select

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub channel_smoothing
'
'Mathias Knaak
'25.08.08
'
'
' 
''INPUT:(result_var,t_start,t_end,input_chn,fit_type,fit_points,when_eva)
'result_var: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'fit_type: linear, quadratic, cubic
'fit_points: number of points for fitting
'all other inputs are not yet implemented

sub channel_smoothing(result_var,t_start,t_end,input_chn,fit_type,byval fit_points,when_eva)


dim start, stop_count,count,half_points,smooth_value,time_step,wf_offset,length

'get time steps of input channel
time_step=chnpropvalget(input_chn,"wf_increment")
wf_offset=chnpropvalget(input_chn,"wf_start_offset")
length=cl(input_chn)
'msgbox length
'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then 
  call chnalloc(result_var,length,,,,groupindexget("evaluations"))  
  call chnwfpropset("evaluations/"&result_var,"x","s",wf_offset,time_step)
else
  call chnrealloc("evaluations/"&result_var,length)
  call chnwfpropset("evaluations/"&result_var,"x","s",wf_offset,time_step)
end if

'get step values for time range
start=wf_step(t_start,input_chn)
stop_count=wf_step(t_end,input_chn)

half_points=trunc(fit_points/2)
fit_points=(2*half_points)+1
smooth_value=0

'call msgbox (fit_points &"  ,  "& half_points)

'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)
call chncopy(input_chn,"calculations/input")
call chncopy("evaluations/"&result_var,"calculations/result")
call chnrealloc("calculations/result",length)
cl("calculations/result")=length
dim  c_res, position,c_input
c_input=cno("calculations/input")

c_res=cno("calculations/result")

groupdefaultset(groupindexget("calculations"))



select case fit_type

 case "Lin."
      smooth_value=chdx(start,c_input)
      for count=start to stop_count 'loop over complete time range
        
        if count < (start+half_points) then
          chdx(count,c_res)=(smooth_value/(1+(count-start)*2))
          smooth_value=smooth_value+chdx(count+(count-start)+1,c_input)+chdx(count+(count-start)+2,c_input)
          
          'chd(count,"evaluations/"&result_var)=smooth_value/(1+count-start)
        'elseif  (count > (start+half_points)) and (count <(start+fit_points))  then
        '  smooth_value=smooth_value+chdx(count,c_input)
        '  chdx(count-half_points,c_res)=(smooth_value/(1+count-start))
        elseif (count >= (start+half_points)) and count < (stop_count-half_points)  then
          chdx(count,c_res)=smooth_value/(fit_points)
          smooth_value=smooth_value+chdx(count+half_points+1,c_input)-chdx(count-half_points,c_input)
         else
          chdx(count,c_res)=smooth_value/((stop_count-count)*2+1)
          smooth_value=smooth_value-chdx(count-(stop_count-count)+1,c_input)-chdx(count-(stop_count-count),c_input)
          'call msgbox((fit_points-count+stop_count)&" ,"&count&" , "&smooth_value)
          
        end if

      next


case "Quad"
      call msgbox ("Quadratic smoothing is not implemented."&chr(13)&"Linear smoothing will be used.") 
      smooth_value=chdx(start,c_input)
      for count=start to stop_count 'loop over complete time range
        
        if count < (start+half_points) then
          chdx(count,c_res)=(smooth_value/(1+(count-start)*2))
          smooth_value=smooth_value+chdx(count+(count-start)+1,c_input)+chdx(count+(count-start)+2,c_input)
          
          'chd(count,"evaluations/"&result_var)=smooth_value/(1+count-start)
        'elseif  (count > (start+half_points)) and (count <(start+fit_points))  then
        '  smooth_value=smooth_value+chdx(count,c_input)
        '  chdx(count-half_points,c_res)=(smooth_value/(1+count-start))
        elseif (count >= (start+half_points)) and count < (stop_count-half_points)  then
          chdx(count,c_res)=smooth_value/(fit_points)
          smooth_value=smooth_value+chdx(count+half_points+1,c_input)-chdx(count-half_points,c_input)
         else
          chdx(count,c_res)=smooth_value/((stop_count-count)*2+1)
          smooth_value=smooth_value-chdx(count-(stop_count-count)+1,c_input)-chdx(count-(stop_count-count),c_input)
          'call msgbox((fit_points-count+stop_count)&" ,"&count&" , "&smooth_value)
          
        end if

      next


case "Cub."
      call msgbox ("Cubic smoothing is not implemented."&chr(13)&"Linear smoothing will be used.")
      smooth_value=chdx(start,c_input)
      for count=start to stop_count 'loop over complete time range
        
        if count < (start+half_points) then
          chdx(count,c_res)=(smooth_value/(1+(count-start)*2))
          smooth_value=smooth_value+chdx(count+(count-start)+1,c_input)+chdx(count+(count-start)+2,c_input)
          
          'chd(count,"evaluations/"&result_var)=smooth_value/(1+count-start)
        'elseif  (count > (start+half_points)) and (count <(start+fit_points))  then
        '  smooth_value=smooth_value+chdx(count,c_input)
        '  chdx(count-half_points,c_res)=(smooth_value/(1+count-start))
        elseif (count >= (start+half_points)) and count < (stop_count-half_points)  then
          chdx(count,c_res)=smooth_value/(fit_points)
          smooth_value=smooth_value+chdx(count+half_points+1,c_input)-chdx(count-half_points,c_input)
         else
          chdx(count,c_res)=smooth_value/((stop_count-count)*2+1)
          smooth_value=smooth_value-chdx(count-(stop_count-count)+1,c_input)-chdx(count-(stop_count-count),c_input)
          'call msgbox((fit_points-count+stop_count)&" ,"&count&" , "&smooth_value)
          
        end if

      next



end select

call chndel("evaluations/"&result_var)
call select_timerange("calculations/result",t_start,t_end)
call chncopy("calculations/result","evaluations/"&result_var)
call groupdel(groupindexget("calculations"))



end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub remove_outlier
'Mathias Knaak
'22.09.2008
'
sub remove_outlier(result_chn,byval t_start,byval t_end,input_chn,limit,when_eva)

dim name,chn_length

'create group for calculations
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/input")

call groupdefaultset(groupindexget("calculations"))

call select_timerange("calculations/input",t_start,t_end)

chn_length=cl(input_chn)

'group_name="Travel_"&Name
'channelname=group_name&"/Travel"
'T1=group_name&"/Travel_NV"
'T2=channelname
'T4=channelname&"_Diff_Y"

'diff_chn=channelname&"_Diff_Y"
'NV_chn="Travel_NV"

L1=limit

' differentiate inputchannel
Call ChnDifferentiate("","calculations/input","","calculations/diff") 

' calculate absolute Values of differentiated channel
Call chncalculate ("CH(""calculations/diff"")=abs(CH(""calculations/diff""))")

'create new channel 
Call Chnalloc("NV_removed")
'remove distorted values: if derivation is greater than threshold, Values will be replaced with NOVALUES
Call ChnCalculate("Ch(""calculations/NV_removed"") = Ch(""calculations/input"") + CTNV(Ch(""calculations/diff"")>L1) ")    
'generate time channel and rename channel (bug in DIAdem)
name=ChnFromWfXGen("calculations/diff","calculations/diff_time") '... Y,E            
cn(name)="diff_time"
'interpolate NoValues
Call ChnNovHandle("calculations/diff_time","calculations/NV_removed","Interpolate","XY",1,0,0) 
'create Waveformchannel
Call ChnToWfChn("calculations/diff_time","calculations/NV_removed",0)            
'insert one "zero" for original channel length (at differentiation one value was "lost") 
call chnrealloc("calculations/NV_removed",chn_length)
Call ChnAreaInsert0("calculations/NV_removed", 1, 1)                                 

call select_timerange("calculations/NV_removed",t_start,t_end)

call chncopy("calculations/NV_removed","evaluations/"&result_chn)
call groupdel(groupindexget("calculations"))

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''















''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub two_channel_math
'9.9.2008 Mathias Knaak
'
'description: allows math. operations with two channels. ( + - * / )
'if the sample rate is not the same, the result will have the highest of both sample rates.

sub two_channel_math(result_chn,t_start,t_end,chn_one,math_op,chn_two, when_eva)

'dim start, stop_count,count,half_points,smooth_value,time_step,wf_offset
dim time_step_1, time_step_2,wf_offset_1, wf_offset_2
dim name

'get time steps of input channels
time_step_1=chnpropvalget(chn_one,"wf_increment")
time_step_2=chnpropvalget(chn_two,"wf_increment")

wf_offset_1=chnpropvalget(chn_one,"wf_start_offset")
wf_offset_2=chnpropvalget(chn_two,"wf_start_offset")

if cl("evaluations/"&result_chn)<>0 then
  call chndel("evaluations/"&result_chn)
  call chnalloc(result_chn,,,,,groupindexget("evaluations"))
end if


'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

'copy channels to new "calculations"-Group
'if both channels have different sample rates, scale both channels to the maximum of both sample rates


if time_step_1=time_step_2 then 'both channels have the same sample rate, copy both channels
  call chncopy(chn_one,"calculations/chn_one")
  call chncopy(chn_two,"calculations/chn_two")
  call chnwfpropset("evaluations/"&result_chn,"x","s",wf_offset_1,time_step_1)
  'msgbox "identical channels"
elseif  time_step_1<time_step_2 then  ' different sample rates
  'copy channel with higher sample rate
  call chncopy(chn_one,"calculations/chn_one")
  'generate channel with x-values
  name=chnfromwfxgen(chn_one,"calculations/x_values")
  cn(name)="x_values"
  'map channel to higher sample rate
  Call Chnmaplincalc(,chn_two,"calculations/x_values","calculations/chn_two",1,"Const. value",NOVALUE,"analogue")
  'create Waveform-channel
  call chntowfchn("calculations/x_values","calculations/chn_two",1)
  
  call chnwfpropset("calculations/chn_two","x","s",wf_offset_1,time_step_1)
  call chnwfpropset("evaluations/"&result_chn,"x","s",wf_offset_1,time_step_1)
elseif time_step_1>time_step_2 then
  call chncopy(chn_two,"calculations/chn_two")
  name=chnfromwfxgen(chn_two,"calculations/x_values")
  cn(name)="x_values"

  'name=chnfromwfxgen(chn_one,"calculations/x_values_1")
  'cn(name)="x_values_1"
  'msgbox chn_one
  Call Chnmaplincalc(,chn_one,"calculations/x_values","calculations/chn_one",1,"Const. value",NOVALUE,"analogue")
  call chntowfchn("calculations/x_values","calculations/chn_one",1)
  call chnwfpropset("calculations/chn_one","x","s",wf_offset_2,time_step_2)
  call chnwfpropset("evaluations/"&result_chn,"x","s",wf_offset_2,time_step_2)
'  msgbox "fertig mit else"
end if

'length=cl("calculations/chn_one")
'msgbox length
'call chnrealloc("evaluations/"&result_chn,length)

'''''''''''''
Dim count_start,count_end,second_start,count
'read time range
count_start=wf_step(t_start,"calculations/chn_one")
count_end=wf_step(t_end,"calculations/chn_one")
second_start=wf_step(t_start,"calculations/chn_two")

call chnrealloc("evaluations/"&result_chn,count_end)
call chncopy("evaluations/"&result_chn,"calculations/result")
call chnrealloc("calculations/result",count_end)
cl("calculations/result")=count_end

'msgbox count_start&"  "&count_end
'do not refrest date portal after every change of a value
Call UIAutoRefreshSet(False)   

'calculate difference of each valua

dim c_one, c_two, c_res, position
c_one=cno("calculations/chn_one")
c_two=cno("calculations/chn_two")
c_res=cno("calculations/result")

groupdefaultset(groupindexget("calculations"))

select case math_op

    case "+"
  '    msgbox "Anfang der berechnung"
'      For count=0 To count_end-count_start
'        chd(count+count_start,"evaluations/"&result_chn)=chd(count+count_start,"calculations/chn_one")+chd(count+second_start,"calculations/chn_two")
'      Next

      For count=0 To count_end-count_start
         chdx(count+count_start,c_res)=chdx(count+count_start,c_one)+chdx(count+second_start,c_two)
      Next


    case "-"
    
 '     For count=0 To count_end-count_start
 '       chd(count+count_start,"evaluations/"&result_chn)=chd(count+count_start,"calculations/chn_one")-chd(count+second_start,"calculations/chn_two")
 '     Next

      For count=0 To count_end-count_start
         chdx(count+count_start,c_res)=chdx(count+count_start,c_one)-chdx(count+second_start,c_two)
      Next

    case "*"

'      For count=0 To count_end-count_start
'        chd(count+count_start,"evaluations/"&result_chn)=chd(count+count_start,"calculations/chn_one")*chd(count+second_start,"calculations/chn_two")
'      Next

      For count=0 To count_end-count_start
         chdx(count+count_start,c_res)=chdx(count+count_start,c_one)*chdx(count+second_start,c_two)
      Next
      
    case "/"
    
'      For count=0 To count_end-count_start
'        chd(count+count_start,"evaluations/"&result_chn)=chd(count+count_start,"calculations/chn_one")/chd(count+second_start,"calculations/chn_two")
'      Next
      
      Call CHNCALCULATE("Ch(""calculations/chn_two"") = Ch(""calculations/chn_two"")+CTNV(Ch(""calculations/chn_two"")=0)")
      For count=0 To count_end-count_start
         chdx(count+count_start,c_res)=chdx(count+count_start,c_one)/chdx(count+second_start,c_two)
      Next

end select

      For count=1 To count_start-1
         chdx(count,c_res)=Novalue
      Next



Call UIAutoRefreshSet(True) 

'msgbox "abgeschlossen"
call chndel("evaluations/"&result_chn)
call chncopy("calculations/result","evaluations/"&result_chn)
call groupdel(groupindexget("calculations"))

''''''''''''



end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub one_chn_math (result_chn,t_start,t_end,input_chn,math_op,input_var,when_eva,formula_string)
' Mathias Knaak
' 16.09.2008
' description: allows several mathematical operations with one channel
' Input:
' result_chn: Name of result channel
' t_start,t_end: time for calculations
' input_chn: input
' math_op: * , / , + , - , ^ , f(input), constant, absolut
' input_var: variable for calculation
' when eva: used by "perform_eva"
' formula_string: string for calculation with Diadem calculator, use "yVal" as reference for the input value
'                 if string ends with ".txt", load formula from file at autoactpath\formula_files

sub one_chn_math(result_chn,t_start,t_end,input_chn,math_op,byval input_var,when_eva,formula_string)

'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

'copy input channel to new calculations group
call chncopy(input_chn,"calculations/input")

'cut timerange to desired values
call select_timerange("calculations/input",t_start,t_end)

dim factor, count, start_count, end_count,chn_number,formula

'get sample number at begin and end of timerange
start_count=wf_step(t_start,"calculations/input")
end_count=wf_step(t_end,"calculations/input")

input_var=val(input_var)

'select type of math operation
select case math_op

  case "*"  'Multiplication
    'scale values with variable
    call chnlinscale("calculations/input","calculations/input",input_var,0)

  case "/" 'Division
    'prevent dividion by zero
    factor=1000000000
    if input_var<> 0 then factor = 1/input_var
    ' multiply with reciprocal value of input_var
    call chnlinscale("calculations/input","calculations/input",factor,0)

  case "+"  
    'Add input_var
    call chnlinscale("calculations/input","calculations/input",1,input_var)

  case "-"
    'subtract input_var
    call chnlinscale("calculations/input","calculations/input",1,-1*input_var)

  case "^"
    call groupdefaultset(groupindexget("calculations"))
    chn_number=cno("input")
    for count=start_count to end_count
      chdx(count,chn_number)=chdx(count,chn_number)^input_var
    next

  case "Constant"
    'write constant values 
    call groupdefaultset(groupindexget("calculations"))
    chn_number=cno("input")
    
    for count=start_count to end_count
      chdx(count,chn_number)=input_var
    next

  case "Absolut"
    'calculate absolute values
    call groupdefaultset(groupindexget("calculations"))
    chn_number=cno("input")
    
    for count=start_count to end_count
      chdx(count,chn_number)=abs(chdx(count,chn_number))
    next

  case "f()" 'use a free defined formula for the Diadem calculator
    'remove all novalues 
    call chnnovhandle("calculations/input","calculations/input","SetValue","XY",True,True,1)
      
    call groupdefaultset(groupindexget("calculations"))
    
    formula=formula_string
    'if formula string ends with ".txt", it is a file name
    if right(formula,4)=".txt" then
      Dim tfh
      ' open text file, read formula from text file
      tfh = TextFileOpen(autoactpath & "\formula_files\"&formula, tfRead)
      If TextFileError(tfh) = 0 Then
        formula=Textfilereadln(tfh)
      end if
    textfileclose(tfh)
    end if
        
    'replace "yVal" with input channel data
    formula=replace(formula,"yVal","CH(""input"")")
    ' add string for result of calculation
    formula="CH(""input"")="&formula
    
    'calculate formula    
    call chncalculate(formula)
    'remove all Values out of selected time range
    call select_timerange("calculations/input",t_start,t_end)

end select

'copy result channel to the evaluations group
call chncopy("calculations/input","evaluations/"&result_chn)
call groupdel(groupindexget("calculations"))


end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub channel_differential
'
'Mathias Knaak
'11.09.08
'
'
' 
''INPUT:(result_chn,t_start,t_end,input_chn,deriv_type,fit_type,fit_points,when_eva)
'result_chn: name of result channel
't_start, t_end: time for start and end of search in the channel
'input_chn: Input data
'deriv_type: Slope or Curve (1st or 2nd derivation)
'fit_type: type of curve fitting (Lin, Quad. or Cubic)
'fit_points: number of points for curve fitting
'when_eva: When is evaluation to perform

sub channel_differential(result_chn,t_start,t_end,input_chn,deriv_type,fit_type,fit_points,when_eva)

call channel_smoothing(result_chn,t_start,t_end,input_chn,fit_type,fit_points, when_eva)

select case deriv_type

    case "Slope"

      Call ChnDeriveCalc("","evaluations/"&result_chn,"evaluations/"&result_chn)

    case "Curve"

      Call ChnDeriveCalc("","evaluations/"&result_chn,"evaluations/"&result_chn)
      Call ChnDeriveCalc("","evaluations/"&result_chn,"evaluations/"&result_chn)

end select

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub chn_fft
'
' Mathias Knaak
' 12.09.2008
' desription: calculates the FFT of a signal channel
' Input:
' result_chn: name of result channel
' t_start,t_end: Timerange for evaluation
' input_chn
' FFT_Type: type of result: Ampl.-Spectr.: Amplitude
'                           Phase-Spectr.: Phase
'                           FFT $ IFT    : IFFT

sub chn_fft(result_chn,t_start,t_end,input_chn,FFT_type,LP_freq,HP_freq,Trapez_lines,when_eva)

'call msgbox (t_start&" ; "&t_end)

'create channel for result of calculation
if groupindexget("calculations")=0 then call groupcreate("calculations",groupcount+1)

call chncopy(input_chn,"calculations/input")
call select_timerange("calculations/input",t_start,t_end)

call chnnovhandle("calculations/input","calculations/input","SetValue","XY",True,True,0)

'call msgbox (lp_freq&" , "&hp_freq)

'hp_freq=val(hp_freq)
'hp_freq=val(hp_freq)

'Call UIAutoRefreshSet(True)     

if LP_freq<>HP_freq then

  if LP_freq>0 and HP_freq=0 then 'Low-Pass Filter

    Call ChnFiltCalc("","calculations/input","calculations/input","FIR","Bessel","Low pass",4,LP_freq,100,1000,1.2,1000,"Rectangle",1,0) 
    msgbox "Low pass"
  
  elseif LP_freq=0 and HP_freq>0 then 'High-Pass Filter
    
    Call ChnFiltCalc("","calculations/input","calculations/input","FIR","Bessel","High pass",4,HP_freq,100,1000,1.2,1000,"Rectangle",1,0)     
    msgbox "high pass"
  
  elseif LP_freq>0 and HP_freq>LP_freq then 'Band Stop Filter

    Call ChnFiltCalc("","calculations/input","calculations/input","FIR","Bessel","Band Stop",4,HP_freq,LP_freq,HP_freq,1.2,1000,"Rectangle",1,0)     
    msgbox "band stop"
  
  elseif LP_freq>0 and HP_freq<LP_freq then 'Band Pass Filter

   'Call ChnFiltCalc("","calculations/input","calculations/input","IIR","Bessel","Band Pass",4,HP_freq,HP_freq,LP_freq,1.2,200,"Rectangle",1,0)      
   Call ChnFiltCalc("","calculations/input","calculations/input","FIR","Bessel","Band Pass",4,HP_freq,HP_freq,LP_freq,1.2,1000,"Rectangle",1,0)      
   msgbox "band pass"
  
  end if
  
end if

'Call ChnCharacter("calculations/input")
Call ChnCharacterAll()
'msgbox t_start
'call msgbox (t_start&" ; "&t_end)


call groupdefaultset(groupindexget("calculations"))

FFTIntervUser    ="NumberStartOverl"
FFTIntervPara(1) =1
FFTIntervPara(2) =12400
FFTIntervPara(3) =1
FFTIntervOverl   =0
FFTWndFct        ="Rectangle"
FFTWndPara       =10
FFTWndChn        ="[2]/I-Injection"
FFTWndCorrectTyp ="No"
FFTAverageType   ="No"
FFTAmplFirst     ="Amplitude"




select case FFT_Type

      case "Ampl.-Spectr."
      
        FFTAmpl          =1
        FFTAmplType      ="Ampl.Peak"
        FFTCalc          =0
        FFTAmplExt       ="No"
        FFTPhase         =0
        FFTCepstrum      =0

        call Chnfft1("","calculations/input")

        call chncopy("calculations/AmplitudePeak","evaluations/"&result_chn)

      case "Phase-Spectr."

        FFTAmpl          =0
        FFTAmplType      ="Ampl.Peak"
        FFTCalc          =0
        FFTAmplExt       ="No"
        FFTPhase         =1
        FFTCepstrum      =0

        call Chnfft1("","calculations/input")

        call chncopy("calculations/Phase","evaluations/"&result_chn)

      case "FFT $ IFT"

      msgbox "This function is not available!"

end select

call groupdel(groupindexget("calculations"))



end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sub two_var_math
' Mathias Knaak
' 16.09.2008
'
'sub two_var_math(result_var,byval xVar,math_op,byval yvar,when_eva, when_plot,when_report,eval_func)
'result_var: Name of result value
'xVar,yVar: input_values for calculation 
'math_op: select mathematical operator
'when_eva, when_plot,when_report
'eval_func: formula or file for evaluation, use syntax of diadem calculator

sub two_var_math(result_var,byval xVar,math_op,byval yVar,when_eva, when_plot,when_report,eval_func)

dim result,formula

xVar=val(xVar)
yVar=val(yVar)

select case math_op

    case "*"
    
      result=xVar*yVar
    
    case "/"
    
      result=xVar/yVar
    
    case "+"
    
      result=xVar+yVar
    
    case "-"
    
      result=xVar-yVar
    
    case "^"
    
      result=xVar^yVar
    
    case "%"
    
      result=xVar*yVar/100
    
    case "Mid"
    
      result=(xVar+yVar)/2
    
    case "Min"
    
      result=valmin(xVar,yVar)
    
    case "Max"
    
      result=valmax(xVar,yVar)
    
    case "aMin"
    
      result=abs(valmin(xVar,yVar))
    
    case "aMax"
    
      result=abs(valmax(xVar,yVar))
    
    case "MinA"
    
      result=valmin(abs(xVar),abs(yVar))
    
    case "MaxA"
    
      result=valmax(abs(xVar),abs(yVar))
    
    case "f()"

     ' result=17    
     ' msgbox "funktion ohne funktion"
      formula=eval_func
      
    'if formula string ends with ".txt", it is a file name
    if right(formula,4)=".txt" then
      Dim tfh
      ' open text file, read formula from text file
      tfh = TextFileOpen(autoactpath & "\formula_files\"&formula, tfRead)
      If TextFileError(tfh) = 0 Then
        formula=Textfilereadln(tfh)
      end if
    textfileclose(tfh)
    end if
    'call globaldim ("x_value,y_value")
     R1 = xVar
     R2 = yVar
    
    

    'replace "xVal" and "yVal" with input channel data
    formula=replace(formula,"xVal","R1")
    formula=replace(formula,"yVal","R2")
    ' add string for result of calculation
    formula="R3="&formula
    
    'calculate formula    
    call chncalculate(formula)
    
    result=R3



end select

chd(1,"evaluations/"&result_var)=result


end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sub one_var_math
' Mathias Knaak
' 17.09.2008
'
'sub one_var_math(result_var,math_op,byval xVar,when_eva, when_plot,when_report,eval_func)
'result_var: Name of result value
'xVar,yVar: input_values for calculation 
'math_op: select mathematical operator
'when_eva, when_plot,when_report
'eval_func: formula or file for evaluation, use syntax of diadem calculator

sub one_var_math(result_var,math_op,byval xVar,when_eva, when_plot,when_report,eval_func)

dim result,formula

xVar=val(xVar)

select case math_op

    case "Ln"
      if xVar>0 then
        result=ln(xVar)        
      else
        result=Novalue
        call msgbox ("Error:  ln("&xVar&")  is no valid input")
      end if
    case "Exp"

      result=exp(xVar)        

    case "Lg"
    
    if xVar>0 then
      result=lg(xVar)        
    else
      result=Novalue
      call msgbox ("Error:  lg("&xVar&")  is no valid input")
    end if

    case "10^"
      result=10^(xVar)        
    
    case "Sin"
      result=sin(xVar)        
    
    case "aSin"
      if abs(xVar)<=1 then
        result=asin(xVar)        
      else
        result=Novalue
        call msgbox ("Error: asin("&xVar&") is no valid input")
      end if
          
    case "Cos"
      result=cos(xVar)        
    
    case "aCos"
    if abs(xVar)<=1 then
      result=acos(xVar)        
    else
      result=Novalue
      call msgbox ("Error: acos("&xVar&") is no valid input")
    end if
    
    case "Tan"
      result=tan(xVar)        
    
    case "aTan"
      result=atan(xVar)        
    
    case "SqRt"
    if xVar>=0 then
      result=sqrt(xVar)        
    else
      result=Novalue
      call msgbox ("Error: sqrt("&xVar&") is no valid input")
    end if

    case "Abs"
      result=abs(xVar)        
    
    case "-"
      result=-(xVar)        
    
    case "1/x"
      if xVar<>0 then
        result=1/(xVar)            
      else
        result=Novalue
        call msgbox ("Error: Division by zero")
      end if
    
    case "f()"
      formula=eval_func
      
      'if formula string ends with ".txt", it is a file name
      if right(formula,4)=".txt" then
        Dim tfh
        ' open text file, read formula from text file
        tfh = TextFileOpen(autoactpath & "\formula_files\"&formula, tfRead)
        If TextFileError(tfh) = 0 Then
          formula=Textfilereadln(tfh)
        end if
      textfileclose(tfh)
      end if
      'call globaldim ("x_value,y_value")
       R1 = xVar

      'replace "xVal"  with input value data
      formula=replace(formula,"xVal","R1")

      ' add string for result of calculation
      formula="R3="&formula
    
      'calculate formula    
      call chncalculate(formula)
    
      result=R3

end select

chd(1,"evaluations/"&result_var)=result

end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub cut_chn 
'
sub cut_chn(result_chn,input_chn, byval t_start, byval t_end,when_eva)

dim wf_start,wf_inc,count,copy_start, copy_end, copy_count, input_chn_no, res_chn_no

Call UIAutoRefreshSet(False)  

wf_start=t_start+chnpropvalget(input_chn,"wf_start_offset")
wf_inc=chnpropvalget(input_chn,"wf_increment")



copy_start=wf_step(t_start,input_chn)
copy_end=wf_step(t_end,input_chn)

copy_count=copy_end-copy_start

'call chnrealloc("evaluations/"&result_chn, copy_count)

'create channel group for calculations
if groupindexget("calculations")<>0 then call groupdel(groupindexget("calculations"))

call groupcreate("calculations",groupcount+1)

'copy input channel to calculations group
call chncopy(input_chn,"calculations/input")
'call chncopy("evaluations/"&result_chn,"calculations/result")
'set default group (to use chdx-function)
call groupdefaultset(groupindexget("calculations"))
input_chn_no=cno("calculations/input")

call chnalloc("result",copy_count,1,,,groupindexget("calculations"))

call chnwfpropset("calculations/result","x","s",wf_start,wf_inc)
'input_chn_no=cno(input_chn)
res_chn_no=cno("calculations/result")

cl("calculations/result")=copy_count

for count=0 to copy_count-1
  chd(count+1,res_chn_no)=chd(count+copy_start+1,input_chn_no)
next

call chndel("evaluations/"&result_chn)
call chncopy("calculations/result","evaluations/"&result_chn)

call groupdel(groupindexget("calculations"))
Call UIAutoRefreshSet(False)  
end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''












'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub t_zero
'M.Knaak
'gives the start value for the time as return value, return value is "wf_start_offset"

sub t_zero(result_var,input_chn)

chd(1,"evaluations/"&result_var)=Chnpropvalget(input_chn,"wf_start_offset")


end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub t_max
'M.Knaak
'gives the last time value of a channel

sub t_max(result_var,input_chn)

chd(1,"evaluations/"&result_var)=Chnpropvalget(input_chn,"wf_start_offset")+(Chnpropvalget(input_chn,"wf_increment"))*(Chnpropvalget(input_chn,"length")-1)


end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Documentations and Updates
' Functions for Documentation and Updates of Variables


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' docu_time
' 23.09.2008
' Mathias Knaak
' description: input of a time Variable
' INPUT:
' result_var: Name of Variable
' time_value: input value
' when_eva: when is update to perform

sub docu_time(result_var,time_value,when_eva,when_plot,when_report)

  chd(1,"evaluations/"&result_var)=time_value

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' docu_y_val
' 23.09.2008
' Mathias Knaak
' description: input of a y-Variable
' INPUT:
' result_var: Name of Variable
' y_value: input value
' when_eva: when is update to perform

sub docu_y_val(result_var,y_value,when_eva,when_plot,when_report)

  chd(1,"evaluations/"&result_var)=y_value

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' docu_number
' 23.09.2008
' Mathias Knaak
' description: input of a integer Variable
' INPUT:
' result_var: Name of Variable
' integer_value: input value
' when_eva: when is update to perform

sub docu_number(result_var,integer_value,when_eva,when_plot,when_report)

  chd(1,"evaluations/"&result_var)=trunc(integer_value)

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' docu_string
' 23.09.2008
' Mathias Knaak
' description: input of a String
' INPUT:
' result_var: Name of Variable
' text_string: input value
' when_eva: when is update to perform

sub docu_string(result_var,text_string,when_eva,when_plot,when_report)

  cht(1,"evaluations/"&result_var)=text_string

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' runtime_time
' 23.09.2008
' Mathias Knaak
' description: runtime input of a time Variable
' INPUT:
' result_var: Name of Variable
' time_value: input value
' when_eva: when is update to perform

sub runtime_time(result_var,byval time_value,when_eva,when_plot,when_report)

time_value=inputbox("Please enter new value for "&result_var&" !","Runtime Input",time_value)

  chd(1,"evaluations/"&result_var)=val(time_value)

end sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' runtime_y_val
' 23.09.2008
' Mathias Knaak
' description: runtime input of a  y-Variable
' INPUT:
' result_var: Name of Variable
' y_value: input value
' when_eva: when is update to perform

sub runtime_y_val(result_var,byval y_value,when_eva,when_plot,when_report)

y_value=inputbox("Please enter new value for "&result_var&" !","Runtime Input",y_value)

  chd(1,"evaluations/"&result_var)=val(y_value)

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' runtime_integer
' 23.09.2008
' Mathias Knaak
' description: runtime input of a integer Variable
' INPUT:
' result_var: Name of Variable
' integer_value: input value
' when_eva: when is update to perform

sub runtime_integer(result_var,byval integer_value,when_eva,when_plot,when_report)

integer_value=inputbox("Please enter new value for "&result_var&" !","Runtime Input",integer_value)

  chd(1,"evaluations/"&result_var)=val(integer_value)

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sel_doc_string(result_var,file_name,when_eva,when_plot,when_report)
'1.10.2008
'Mathias Knaak
'description: reads data from a list file, user can select entries from the list
'INPUT:
' result_var: Name of Variable
' file_name: name of file with list
' when_eva: when is update to perform


sub sel_doc_string(result_var,byval file_name,when_eva,when_plot,when_report)

Call globaldim("docu_string_text")
dim max, I

'max = MaxOrd(file_name)

'Fill all fields of docu_str_sel with empty values
For I = 1 To 100
  docu_str_sel_(I) = ""
Next

'msgbox file_name

'Fill in the text field
For I = 0 To MaxOrd(file_name)    'In dyn.enum. file the index starts with zero
'Get the values from the dyn. enumeration variable and assign it to docu_str_sel_()
 docu_str_sel_(I+1) = VEnum(file_name,I) 'text_field_ index starts with 1
Next

file_name=replace(file_name,"_","")

'generate headline for dialog
T4="Select "&file_name&" :"

'call dialog
if SUDDlgShow("Dlg_docu_string",AutoActPath & "edit_v1.8.sud","")="IDOk" then
  'write result to result channel
  cht(1,"evaluations/"&result_var)=docu_string_text
end if

end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'code_dep_integer
'M.Knaak
'02.10.2008
'description: read a value from a file, line number is defined by Acqui/Plot/Timecode
'INPUT:
' result_var: Name of Variable
' code type: select type of code to select the line to read data from file
' file_name: name of file with list
' when_eva: when is update to perform


sub code_dep_integer(result_var,code_type,file_name,when_eva, when_plot,when_report)

dim line_index 'number of line in the file to read the value 

select case code_type 'select the type of code

  case "Acqui-Code"
   
    if grouppropexist(2,"acquiCode") then 'if property "acquicode" exists, 
      line_index=grouppropget(2,"acquiCode") 'then set "line_index" to value of this property
    else
      line_index=-1
    end if

  case "Cycle-Code"

    if grouppropexist(2,"PlotCode") then
      line_index=grouppropget(2,"PlotCode")
    else
      line_index=-1
    end if

  case "TimingCode"

    if grouppropexist(2,"TrCode") then
      line_index=grouppropget(2,"TrCode")
    else
      line_index=-1
    end if 

end select


Dim tfh, result
' open text file, read line from text file
tfh = TextFileOpen(autoactpath & "\Document_code_files\"&file_name, tfRead)
If TextFileError(tfh) = 0 Then
call Textfileseek(tfh,val(line_index)) ' set textfilehandle to desired line
 result=Textfilereadln(tfh)
end if
textfileclose(tfh)

chd(1,"evaluations/"&result_var)=val(result)

end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'code_dep_string
'M.Knaak
'02.10.2008
'description: read a value from a file, line number is defined by Acqui/Plot/Timecode
'INPUT:
' result_var: Name of Variable
' code type: select type of code to select the line to read data from file
' file_name: name of file with list
' when_eva: when is update to perform


sub code_dep_string(result_var,code_type,file_name,when_eva, when_plot,when_report)

dim line_index 'number of line in the file to read the value 

select case code_type 'select the type of code

  case "Acqui-Code"
   
    if grouppropexist(2,"acquiCode") then 'if property "acquicode" exists, 
      line_index=grouppropget(2,"acquiCode") 'then set "line_index" to value of this property
    else
      line_index=-1
    end if

  case "Cycle-Code"

    if grouppropexist(2,"PlotCode") then
      line_index=grouppropget(2,"PlotCode")
    else
      line_index=-1
    end if

  case "TimingCode"

    if grouppropexist(2,"TrCode") then
      line_index=grouppropget(2,"TrCode")
    else
      line_index=-1
    end if 

end select


Dim tfh, result
' open text file, read line from text file
tfh = TextFileOpen(autoactpath & "\Document_code_files\"&file_name, tfRead)
If TextFileError(tfh) = 0 Then
call Textfileseek(tfh,val(line_index)) ' set textfilehandle to desired line
 result=Textfilereadln(tfh)
end if
textfileclose(tfh)

cht(1,"evaluations/"&result_var)=result

end sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'excel_export
'M.Knaak
'2.10.2008
'description: exports selected files to excel files
'Input:
'result_var: name of result, saves file name
'config file: use config file to save data without user interaction, search path ist autoactpath\export_config
'save_path: Path to save files; name of file will be generated automatically from group description
'time channel: select one channel to generate a channel with time values (otherwise Excel can save only sample counts)


sub excel_export(result_var,config_file, save_path, time_chn)

dim file_name,name
'call groupdefaultset(2)
'read group description; importrbs writes serial- and shotnumber into description
file_name=grouppropget(2,"description")
'replace "/" with "_"
file_name=replace(file_name,"/","_")
'add ".xls"
file_name="\"&file_name&".xls"

'generate time channel for Export
if time_chn<>"" then 
name=ChnFromWfXGen(time_chn,"/Time")
cn(name)="Time"
end if

'call export to Excel with dialog window so select channels and options
if config_file="" then
  call Excelexport(save_path&file_name,,1,"")
end if

'call export to Excel without dialog window, load settings from config file
if config_file<>"" then
  call Excelexport(save_path&file_name,,0,autoactpath&"export_config\"&config_file)
end if
'delete backslash in filename
file_name=replace(file_name,"\","")
'write filename to result channel
cht(1,"evaluations/"&result_var)=file_name

end sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'tdm_export
'M.Knaak
'2.10.2008
'description: exports selected files to tdm files
'Input:
'result_var: name of result, saves file name
'save_path: Path to save files; name of file will be generated automatically from group description


sub tdm_export(result_var,save_path,save_all)

dim file_name,saved_channels
call groupdefaultset(2)
'read group description; importrbs writes serial- and shotnumber into description
file_name=grouppropget(2,"description")
'replace "/" with "_"
file_name=replace(file_name,"/","_")
'add ".tdm"
file_name="\"&file_name&".tdm"

if save_all="True" then
'saved_channels="'[1]/[2]'-'[1]/["&groupchncount(1)&"]','[2]/[1]'-'[2]/["&groupchncount(2)&"]'"
'msgbox "true"
call datafilesave(save_path&file_name,"TDM")
elseif save_all="False" then
'msgbox "false"
saved_channels="'[2]/[1]'-'[2]/["&groupchncount(2)&"]'"
call datafilesavesel(save_path&file_name,"TDM",saved_channels)
end if 

'msgbox saved_channels
'call datafilesavesel(save_path&file_name,"TDM",saved_channels)

'delete backslash in filename
file_name=replace(file_name,"\","")
'write filename to result channel
'msgbox file_name
cht(1,"evaluations/"&result_var)=file_name

end sub




















'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Help functions
'






'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'function chd_wf_time
' 
' Mathias Knaak
' 07.08.08
'is a function like "CHD" with time values as input
'
'Input: imput_time, input_chn
'Return Value: y-Value of channel at specified time
'

function chd_wf_time(input_time, input_chn)

dim chn_step
chn_step=1+((input_time-Chnpropvalget(input_chn,"wf_start_offset"))/Chnpropvalget(input_chn,"wf_increment"))
chd_wf_time=chd(trunc(chn_step), input_chn)+ Frac(chn_step)*(chd(trunc(chn_step+1), input_chn)-chd(trunc(chn_step), input_chn))



end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'function wf_time
'
' Mathias Knaak
' 07.08.08

'calculates the time corresponding to a specified step value 

function wf_time(value_num, input_chn)

wf_time=Chnpropvalget(input_chn,"wf_start_offset")+ (value_num - 1)*Chnpropvalget(input_chn,"wf_increment")

end function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'function wf_step
'
' Mathias Knaak
' 07.08.08


'calculates the step_no corresponding to a specified value of time

function wf_step(input_time, input_chn)
dim step_no
step_no=1+trunc((val(input_time)-Chnpropvalget(input_chn,"wf_start_offset"))/Chnpropvalget(input_chn,"wf_increment"))
if step_no<1 then step_no=1
wf_step=step_no
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub select_timerange
'Mathias Knaak
'4.08.2008
'fills all values out of selected time range with NOVALUE

sub select_timerange(input_channel, byval time_start, byval time_end)

dim sample_count, start_step, chn_length, t_help
  
  if val(time_end) < val(time_start) then
    msgbox "Wertetausch"
    t_help=time_start
    time_start=time_end
    time_end=t_help
  end if

  start_step=wf_step(time_start,input_channel)
  chn_length=chnlength(input_channel)
  
    for sample_count=1 to start_step-1
       chd(sample_count,input_channel)=novalue
    next

    sample_count=1
    start_step=wf_step(time_end,input_channel)+1
    for sample_count=start_step to chn_length
       chd(sample_count,input_channel)=novalue
    next


end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sub reverse_channel
'Mathias Knaak
'5.08.2008
'reverses a waveform-channel
'
'Input: Channel Name
'

sub reverse_channel(input_chn)
dim new_inc, new_offset

new_offset=Chnpropvalget(input_chn,"wf_start_offset")+CL(input_chn)*Chnpropvalget(input_chn,"wf_increment")
new_inc=-Chnpropvalget(input_chn,"wf_increment")

call chnwfpropset(input_chn,Chnpropvalget(input_chn,"wf_xname"),Chnpropvalget(input_chn,"wf_xunit_string"),new_offset,new_inc)


end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub add_marker
' creates an x/y-pair to mark a position in a report
' "type" selects type of marker: "time" adds a cross, "value" adds a horizontal line
'
sub add_marker(x_value,y_value,mark_type,input_chn)

dim position

select case mark_type

  case "time"

      'create channels for result of calculation
      if cno("evaluations/x_mark")=0 then call chnalloc("x_mark",100,,,,groupindexget("evaluations"))  
      if cno("evaluations/y_mark")=0 then call chnalloc("y_mark",100,,,,groupindexget("evaluations"))  

      'get length of channel
      position=cl("evaluations/x_mark")+1
      'write result
      chd(position,"evaluations/x_mark")=x_value
      chd(position,"evaluations/y_mark")=y_value



  case "value"

    'create channels for result of calculation
    if cno("evaluations/x_mark_value")=0 then call chnalloc("x_mark_value",100,,,,groupindexget("evaluations"))  
    if cno("evaluations/y_mark_value")=0 then call chnalloc("y_mark_value",100,,,,groupindexget("evaluations"))  

    'get length of channel
    position=cl("evaluations/x_mark_value")+1
    'write result
    chd(position,"evaluations/x_mark_value")=x_value
    chd(position,"evaluations/y_mark_value")=y_value

end select

if groupindexget("mark_lines")=0 then 
  call groupcreate("mark_lines")
  call chnalloc("channel_list",100,,DataTypeString,,groupindexget("mark_lines"))
end if

dim count,chn_found
chn_found=false

if input_chn <> "" then 
  
  for count=1 to cl("mark_lines/channel_list")
    if input_chn=cht(count,"mark_lines/channel_list") then chn_found=true
  next

  if not chn_found then cht(cl("mark_lines/channel_list")+1,"mark_lines/channel_list")=input_chn
end if


end sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub remove_marker
' removes all xy-marker
'
'
sub remove_marker()

dim position

'delete marker channels
if cno("evaluations/x_mark")<>0 then call chndel("evaluations/x_mark")
if cno("evaluations/y_mark")<>0 then call chndel("evaluations/y_mark")

if cno("evaluations/x_mark_value")<>0 then call chndel("evaluations/x_mark_value")
if cno("evaluations/y_mark_value")<>0 then call chndel("evaluations/y_mark_value")

if groupindexget("mark_lines")<>0 then call groupdel(groupindexget("mark_lines"))

end sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub add_marker_line
' adds a marker line in the mark_lines group
'call "draw_mark_lines" to draw mark lines in view-Window

sub add_marker_line(x1,y1,x2,y2,input_chn)

if groupindexget("mark_lines")=0 then 
  call groupcreate("mark_lines")
  call chnalloc("channel_list",100,,DataTypeString,,groupindexget("mark_lines"))
end if


dim chn_count,count,chn_found
chn_found=false
'get number of channels for new channel names
chn_count=(groupchncount(groupindexget("mark_lines"))-1)/2+1

'add new channels to the group
call chnalloc("X-Marker_"&chn_count,2,,,,groupindexget("mark_lines"))
call chnalloc("Y-Marker_"&chn_count,2,,,,groupindexget("mark_lines"))

'write values into channels
chd(1,"mark_lines/X-Marker_"&chn_count)=x1
chd(2,"mark_lines/X-Marker_"&chn_count)=x2
chd(1,"mark_lines/Y-Marker_"&chn_count)=y1
chd(2,"mark_lines/Y-Marker_"&chn_count)=y2

if input_chn <> "" then 
  
  for count=1 to cl("mark_lines/channel_list")
    if input_chn=cht(count,"mark_lines/channel_list") then chn_found=true
  next

  if not chn_found then cht(cl("mark_lines/channel_list")+1,"mark_lines/channel_list")=input_chn
end if


end sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sub draw_mark_lines
'adds all mark_lines in "mark_lines" to view-window
'adds all selected curves to View Window

sub draw_mark_lines

View.loadlayout(autoactpath&"view_for_evaluation.TDV")
Call WndShow("VIEW","MAXIMIZE")


if groupindexget("mark_lines")=0 then exit sub

dim marker_count,count
'get number of channels for new channel names
marker_count=groupchncount(groupindexget("mark_lines"))/2

Dim oMySheet: set oMySheet = View.ActiveSheet
Dim oOldArea: Set oOldArea = oMySheet.ActiveArea
dim channel_name

for count=1 to marker_count 
call oOldArea.DisplayObj.Curves.Add("mark_lines/X-Marker_"&count,"mark_lines/Y-Marker_"&count)  
next

for count=1 to cl("mark_lines/channel_list")
  channel_name=cht(count,"mark_lines/channel_list")
  call oOldArea.DisplayObj.Curves.Add("",channel_name)  
next

end sub










'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'function perform_eval(input_chn,eval_code)
'
'
'
'

function perform_eval(input_chn,eval_code)

perform_eval=False



dim perform_code, test_code, plot_code,time_code,test_code_input, time_code_input,plot_code_input,group_index
dim perform_test, perform_plot,perform_time
dim testvalue

perform_plot=false
perform_test=false
perform_time=false

if eval_code="" then exit function


'msgbox eval_code

'Split evaluation code into separate values
perform_code=split(eval_code,"/")
test_code=split(perform_code(0),"|")
plot_code=split(perform_code(1),"|")
time_code=split(perform_code(2),"|")

'read test/time/plot-code from input-channel
if input_chn<>"" then
  group_index=chngroup(input_chn)
else 
 group_index=2
end if


if grouppropexist(group_index,"acquiCode") then
  test_code_input=grouppropget(group_index,"acquiCode")
else
  test_code_input=""
end if

if grouppropexist(group_index,"PlotCode") then
  plot_code_input=grouppropget(group_index,"PlotCode")
else
  plot_code_input=""
end if

if grouppropexist(group_index,"TrCode") then
  time_code_input=grouppropget(group_index,"TrCode")
else
  time_code_input=""
end if


'check, if evaluation code of input_channel is element of eval_code

for each testvalue in test_code
  if testvalue=test_code_input then perform_test=true
next

for each testvalue in plot_code
  if testvalue=plot_code_input then perform_plot=true
next

for each testvalue in time_code
  if testvalue=time_code_input then perform_time=true
next


'return true, if at least one test code in each group is true 

perform_eval= perform_test and perform_plot and perform_time

end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' sub cut_chn_range
' M.Knaak: 11.09.2008
' description: cuts channel range to desired values
' all other values will be filled with novalues
' Input:
' range_start: Range Start 
' range_end:Range End
' input_chn: complete Channelname

sub cut_chn_range(range_start,range_end,input_chn)

call globaldim ("cut_range_start,cut_range_end")
cut_range_start=range_start
cut_range_end=range_end

  Call CHNCALCULATE("Ch("""&input_chn&""") = Ch("""&input_chn&""")+CTNV(Ch("""&input_chn&""")<cut_range_start or Ch("""&input_chn&""")>cut_range_end)")

end sub














