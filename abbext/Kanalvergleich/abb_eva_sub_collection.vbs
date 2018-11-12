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
    
      Case 2  'level_cross_time
       Call level_cross_time(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),get_var_val(eva_inps(4)),"","","","","")    

      Case 3  'y_value_of_time
       Call y_value_of_time(output_name,get_var_val(eva_inps(0)),eva_inps(1),eva_inps(2),get_var_val(eva_inps(3)),"","","","")    

      Case 4  'Extremum: time value
       Call t_extrempoint(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),"","","","","")    

      Case 5  'Extremum: y-value
       Call y_extrempoint(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),"","","","","","")    

      Case 7  'Slope of Secant
       Call slope_of_secant(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),"","","","")    

      Case 12 'y-value of plateau
       Call find_plateau(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),get_var_val(eva_inps(3)),get_var_val(eva_inps(4)),"","","","")    

      Case 13 ' Average Value
       Call average(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),"","","","")    

      Case 14 ' Effective Value
       Call rms_value(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),"","","","")    

      Case 15 ' integral Value
       Call integral_value(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),"","","","")    

      Case 32 ' channel integral
       Call channel_integral(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),get_var_val(eva_inps(5)),"")    
      
      Case 34 ' channel smoothing
       'msgbox eva_inps(5)
       Call channel_smoothing(output_name,get_var_val(eva_inps(0)),get_var_val(eva_inps(1)),eva_inps(2),eva_inps(3),get_var_val(eva_inps(4)),eva_inps(5))    
        
      
      Case 53 ' T-Zero
       Call t_zero(output_name,get_var_val(eva_inps(0)))

      Case 54 ' T-Max
       Call t_max(output_name,get_var_val(eva_inps(0)))


     Case 61
      Call TwoChnOperation(output_name,eva_inps(0),eva_inps(1),eva_inps(2))    
     Case 62
      ' inputs: output name, input channel, scale factor, offset
      Call ChnScale(output_name,eva_inps(0),get_var_val(eva_inps(1)),get_var_val(eva_inps(2)))
      
'     Case 63
'      eval1_input3 = Cint(eval1_input3)
'     Call Example3(eval1_input1,eval1_input2,eval1_input3)
      Case 64
        Call MaxValue(output_name,eva_inps(0))    
      Case 65
        Call ChnDivide(output_name,eva_inps(0),get_var_val(eva_inps(1)))
  End Select


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

'search crossing point
'begin forward search
if t_start < t_end then
'msgbox t_end
    'look at every value in given time range
      for step_val = wf_step(t_start,input_chn) to wf_step(t_end,input_chn)+1
            'if level of input_channel crosses the given Y-Value...
            if (((chd(step_val, input_chn)<=y_level) and (chd(step_val+1, input_chn)>y_level)) or ((chd(step_val, input_chn)>=y_level) and (chd(step_val+1, input_chn)<y_level))) then
                '...crossing point found
                cross_count=cross_count+1 'increase number of found crossing points
                if cross_count=cross_num then 'if crossing if found, write results to result channel
                  CHD(1,"evaluations/"&result_var) =  wf_time(step_val, input_chn)
                  call msgbox ( wf_time(step_val, input_chn))
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
                      call msgbox ( wf_time(step_val, input_chn))
                      exit for            
                    end if
                end if
          next

end if


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
dim result
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
    else    
      result=chd(1,"calculations/Peak_Y_neg_time")
    end if

  case "First Pos."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",1,"Max.Peaks","Time")
    result=chd(1,"calculations/Peak_Y_pos_time")
  
  case "First Neg."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",1,"Min.Peaks","Time")
    result=chd(1,"calculations/Peak_Y_neg_time")
  
  case "First Abs."
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_time","calculations/Peak_Y_pos_time",10,"Max.Peaks","Time")
    Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_time","calculations/Peak_Y_neg_time",10,"Min.Peaks","Time")
      if  chd(1,"calculations/Peak_Y_pos_time") > chd(1,"calculations/Peak_Y_neg_time") then
            result=chd(1,"calculations/Peak_Y_pos_time")
      elseif chd(1,"calculations/Peak_Y_pos_time") = chd(1,"calculations/Peak_Y_neg_time") then
          result=chd(1,"calculations/Peak_Y_pos_time")
      else
          result=chd(1,"calculations/Peak_Y_neg_time")
      end if
  
  case "Highest Pos."
       Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_Y_pos_amplitude") 
  case "Highest Neg."
       Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
       result=chd(1,"calculations/Peak_Y_neg_amplitude")  
  case "Highest Abs."
          Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_pos_amplitude","calculations/Peak_Y_pos_amplitude",10,"Max.Peaks","Amplitude")  
          Call ChnPeakFind("","calculations/y_extrema","calculations/Peak_X_neg_amplitude","calculations/Peak_Y_neg_amplitude",10,"Min.Peaks","Amplitude") 
      if   chd(1,"calculations/Peak_Y_pos_amplitude") > chd(1,"calculations/Peak_Y_neg_amplitude") then
            result=chd(1,"calculations/Peak_Y_pos_amplitude")
      elseif chd(1,"calculations/Peak_Y_pos_amplitude") = chd(1,"calculations/Peak_Y_neg_amplitude") then
          result=chd(1,"calculations/Peak_Y_pos_amplitude")
      else
          result=chd(1,"calculations/Peak_Y_neg_amplitude")
      end if
   
end select

chd(1,"evaluations/"&result_var)=result


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


select case int_type

 case "Lin."

      for count=start to stop_count 'loop over all values in time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn)
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+1)/6*(t_end-t_start)


case "Abs."
      
      for count=start to stop_count 'loop over all values in time range
        'Summe über (y(n) + 4*y(n+1) + y(n+2))^int_exp
        int_val=int_val+(abs(chd(count,input_chn))+(4*abs(chd(count+1,input_chn)))+abs(chd(count+2,input_chn)))^int_exp
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+1)/(6^int_exp)*(t_end-t_start)


case "Quad"
      
      for count=start to stop_count 'loop over all values in time range
        'Summe über (y(n) + 4*y(n+1) + y(n+2))^2
        int_val=int_val+(chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn))^2
      next

' divide by number of steps and multiply with time range
int_val=int_val/(stop_count-start+1)/36*(t_end-t_start)



end select




'call msgbox(int_val)

'create channel for result of calculation
if cno("evaluations/"&result_var)=0 then call chnalloc(result_var,4,,,,groupindexget("evaluations"))  
'write result
chd(1,"evaluations/"&result_var)=int_val


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
'boundary: initial value for the integral
'all other inputs are not yet implemented

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
  'call chnrealloc("evaluations/"&result_var,length)
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


select case int_type

 case "Lin."

      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+(chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn))*(time_step/6)
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next


case "Abs."
      
      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+((abs(chd(count,input_chn))+(4*abs(chd(count+1,input_chn)))+abs(chd(count+2,input_chn)))^int_exp)*(time_step/(6^int_exp))
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next


case "Quad"
      
      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+((chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn))^2)*(time_step/36)
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

sub channel_smoothing(result_var,t_start,t_end,input_chn,fit_type,fit_points,when_eva)


dim start, stop_count,count,half_points,smooth_value,time_step,wf_offset

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
smooth_value=0
select case fit_type

 case "Lin."

      for count=start to stop_count+half_points 'loop over complete time range
        
        if count < (start+half_points) then
          smooth_value=smooth_value+chd(count,input_chn)
          'chd(count,"evaluations/"&result_var)=smooth_value/(1+count-start)
        elseif  (count >= (start+half_points)) and (count <(start+fit_points))  then
          smooth_value=smooth_value+chd(count,input_chn)
          chd(count-half_points,"evaluations/"&result_var)=(smooth_value/(1+count-start))
        elseif (count >= (start+fit_points)) and count <= stop_count  then
          smooth_value=smooth_value+chd(count,input_chn)-chd(count-fit_points,input_chn)
          chd(count-half_points,"evaluations/"&result_var)=smooth_value/(fit_points)
        else
          smooth_value=smooth_value-chd(count-fit_points,input_chn)
          chd(count-half_points,"evaluations/"&result_var)=smooth_value/(fit_points-count+stop_count)
          'call msgbox((fit_points-count+stop_count)&" ,"&count&" , "&smooth_value)
          
        end if

      next


case "Quad"
      
      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+((abs(chd(count,input_chn))+(4*abs(chd(count+1,input_chn)))+abs(chd(count+2,input_chn)))^int_exp)*(time_step/(6^int_exp))
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next


case "Cub."
      
      for count=start to stop_count-1 'loop over complete time range
        'Summe über y(n) + 4*y(n+1) + y(n+2)
        int_val=int_val+((chd(count,input_chn)+(4*chd(count+1,input_chn))+chd(count+2,input_chn))^2)*(time_step/36)
        'write every calculated integral value into channel
        chd(count+1,"evaluations/"&result_var)=int_val
      next



end select

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''











''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
step_no=1+trunc((input_time-Chnpropvalget(input_chn,"wf_start_offset"))/Chnpropvalget(input_chn,"wf_increment"))
if step_no<1 then step_no=1
wf_step=step_no
end function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'sub select_timerange
'Mathias Knaak
'4.08.2008
'fills all values out of selected time range with NOVALUE

sub select_timerange(input_channel,t_start,t_end)

dim sample_count, start_step, chn_length, t_help
  
  if t_end < t_start then
    t_help=t_start
    t_start=t_end
    t_end=t_help
  end if

  start_step=wf_step(t_start,input_channel)
  chn_length=chnlength(input_channel)
  
    for sample_count=1 to start_step-1
       chd(sample_count,input_channel)=novalue
    next

    sample_count=1
    start_step=wf_step(t_end,input_channel)
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
'
'
sub add_marker(x_value,y_value)

dim position

'create channel for result of calculation
if cno("evaluations/x_mark")=0 then call chnalloc(x_mark,100,,,,groupindexget("evaluations"))  
if cno("evaluations/y_mark")=0 then call chnalloc(y_mark,100,,,,groupindexget("evaluations"))  

'get length of channel
position=cl("evaluations/x_mark")+1
'write result
chd(position,"evaluations/x_mark")=x_value
chd(position,"evaluations/y_mark")=y_value

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















end function


























