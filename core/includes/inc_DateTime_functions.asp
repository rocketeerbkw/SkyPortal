<% 
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':: date.time functions
function longDate(tIsoDate)
  dim m, d, y
  y = left(tIsoDate,4)
  m = mid(tIsoDate,5,2)
  d = mid(tIsoDate,7,2)
  longDate = FormatDateTime(DateSerial(y,m,d),1)
end function

function ChkDateFormat(tIsoDate)
	ChkDateFormat =  isdate("" & Mid(tIsoDate, 5,2) & "/" & Mid(tIsoDate, 7,2) & "/" & Mid(tIsoDate, 1,4) & " " & Mid(tIsoDate, 9,2) & ":" & Mid(tIsoDate, 11,2) & ":" & Mid(tIsoDate, 13,2) & "") 
end function

function getDateFormat()
  curDt = dateserial(2006,03,09)
  tFmt = ""
  if strComp(Month("03/09/2006"),"3") = 0 then
    if right(curDt,4) = "2006" then
      tFmt = "mm/dd/yyyy"
	else
      tFmt = "yyyy/mm/dd"
	end if
  else
      tFmt = "dd/mm/yyyy"
  end if
  getDateFormat = tFmt
end function

function getDateDiff(dat1,dat2)
  'dat1 = today
  'dat2 = submit date
  tmpDatDiff = 0
  year1 = left(dat1,4)
  month1 = mid(dat1,5,2)
  day1 = mid(dat1,7,2)
  tmpDat1 = DateSerial(year1,month1,day1)
  
  year2 = left(dat2,4)
  month2 = mid(dat2,5,2)
  day2 = mid(dat2,7,2)
  tmpDat2 = DateSerial(year2,month2,day2)
  
  'tmpDatDiff = DateDiff("d", tmpDat1, tmpDat2)
  
  if year1 = year2 then
    if month1 = month2 then
	  tmpDateDiff = day1 - day2
	else
	  if month1 > month2 then
	    tmpM = month1 - month2
		if tmpM = 1 then
		  tmpDateDiff = (30 - day2) + day1
		else
		  tmpM = (tmpM - 1)*30
		  tmpDateDiff = (30 - day2) + day1 + tmpM
		end if
	  end if
	end if
  else
    if year1 > year2 then
	  tmpY = year1 - year2
	  lastYrDays = ((12 - month2)*30) + (30 - day2)
	  if tmpY = 1 then
		if month1 = 01 then
		  tmpDateDiff = lastYrDays + day1
		else
		  tmpDateDiff = lastYrDays + ((month1 - 1)*30) + day1
		end if
	  else
	    tmpDateDiff = ((tmpY -1)*365) + lastYrDays + ((month1 - 1)*30) + day1
	  end if
	end if
  end if
  getDateDiff = tmpDateDiff
end function

function getDayDiff(iso1,iso2)
  'dat1 = today's datestring
  'dat2 = other date datestring
  getDayDiff = DateDiff("d", chkDate2(iso1), chkDate2(iso2))
end function

function DateToStr2(dtDateTime)
  DateToStr2 = DateToStr(dtDateTime)
end function

function DateToStr(dtDateTime)
	DateToStr = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

function StrToDate(tIsoDate)
  if not isNumeric(tIsoDate) then
    StrToDate = txtError
  else
    if strMCurDateString = tIsoDate then
      tmpDate = chkDate2(tIsoDate) & ChkTime2(tIsoDate)
	else
      'tmpDate = chkDate2(strDateTime) & ChkTime2(strDateTime)
      tmpDate = chkDate2(tIsoDate) & ChkTime2(tIsoDate)
	  tmpDate = DateAdd("h", strMTimeAdjust , tmpDate)
	end if
	  if strTimeType = 12 then
	    tmpDate = FormatDateTime(tmpDate,2) & " " & FormatDateTime(tmpDate,3)
	  else
	    tmpDate = FormatDateTime(tmpDate,2) & " " & FormatDateTime(tmpDate,4)
	  end if
	  StrToDate = tmpDate
  end if
end function

function ChkTime(tIsoDate)
  if tIsoDate <> "" then
    tmpAmPm = ""
    tmpTime = ""
    if strMCurDateString = tIsoDate then
      ChkTime = ChkTime2(tIsoDate)
    else
	  tmpTime = chkDate2(tIsoDate)
	  if strTimeType = 12 then
		if cint(Mid(tIsoDate, 9,2)) > 12 then
			tmpAmPm = "PM"
			tmpTime = tmpTime & " " & _
			(cint(Mid(tIsoDate, 9,2)) -12) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & tmpAmPm
		elseif cint(Mid(tIsoDate, 9,2)) = 12 then
			tmpAmPm = "PM"
			tmpTime = tmpTime & " " & _
			cint(Mid(tIsoDate, 9,2)) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & tmpAmPm
		elseif cint(Mid(tIsoDate, 9,2)) = 0 then
			tmpAmPm = "AM"
			tmpTime = tmpTime & " " & _
			(cint(Mid(tIsoDate, 9,2)) +12) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & tmpAmPm
		else
			tmpAmPm = "AM"
			tmpTime = tmpTime & " " & _
			Mid(tIsoDate, 9,2) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " ' & tmpAmPm
		end if
		tmpTime = DateAdd("h", strMTimeAdjust , tmpTime)
		tmpTime = FormatDateTime(tmpTime,3)
	  else
		tmpTime = tmpTime & " " & _
		Mid(tIsoDate, 9,2) & ":" & _
		Mid(tIsoDate, 11,2) & ":" & _
		Mid(tIsoDate, 13,2) '& " " & tmpAmPm
		'Mid(fTime, 13,2) & " " & tmpAmPm
		'response.Write("tIsoDate: " & tIsoDate)
		'response.Write("<br>tmpTime: " & tmpTime)
		tmpTime = DateAdd("h", strMTimeAdjust , tmpTime)
		tmpTime = FormatDateTime(tmpTime,4)
	  end if
    end if
  end if
	ChkTime = tmpTime ' & tmpAmPm
end function

function ChkTime2(tIsoDate)
  if tIsoDate <> "" then
	if strTimeType = 12 then
		if cint(Mid(tIsoDate, 9,2)) > 12 then
			ChkTime2 = ChkTime2 & " " & _
			(cint(Mid(tIsoDate, 9,2)) -12) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & "PM"
		elseif cint(Mid(tIsoDate, 9,2)) = 12 then
			ChkTime2 = ChkTime2 & " " & _
			cint(Mid(tIsoDate, 9,2)) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & "PM"
		elseif cint(Mid(tIsoDate, 9,2)) = 0 then
			ChkTime2 = ChkTime2 & " " & _
			(cint(Mid(tIsoDate, 9,2)) +12) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & "AM"
		else
			ChkTime2 = ChkTime2 & " " & _
			Mid(tIsoDate, 9,2) & ":" & _
			Mid(tIsoDate, 11,2) & ":" & _
			Mid(tIsoDate, 13,2) & " " & "AM"
		end if
		
	else
		ChkTime2 = ChkTime2 & " " & _
		Mid(tIsoDate, 9,2) & ":" & _
		Mid(tIsoDate, 11,2) & ":" & _
		Mid(tIsoDate, 13,2) 
	end if
  end if
end function

function ChkTime3(fTime)
  if fTime = "" then
	exit function
  end if
  tmpAmPm = ""
  tmpTime = ""
  if strMCurDateString = fTime then
    ChkTime = ChkTime2(fTime)
  else
	tmpTime = chkDate2(fTime)
	if strTimeType = 12 then
		if cint(Mid(fTime, 9,2)) > 12 then
			tmpAmPm = "PM"
			tmpTime = tmpTime & " " & _
			(cint(Mid(fTime, 9,2)) -12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & tmpAmPm
		elseif cint(Mid(fTime, 9,2)) = 12 then
			tmpAmPm = "PM"
			tmpTime = tmpTime & " " & _
			cint(Mid(fTime, 9,2)) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & tmpAmPm
		elseif cint(Mid(fTime, 9,2)) = 0 then
			tmpAmPm = "AM"
			tmpTime = tmpTime & " " & _
			(cint(Mid(fTime, 9,2)) +12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & tmpAmPm
		else
			tmpAmPm = "AM"
			tmpTime = tmpTime & " " & _
			Mid(fTime, 9,2) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & tmpAmPm
		end if
		tmpTime = DateAdd("h", strMTimeAdjust , tmpTime)
		tmpTime = FormatDateTime(tmpTime,3)
		aTime = split(tmpTime,":")
		tmpTime = aTime(0) & ":" & aTime(1) & " " & tmpAmPm
	else
		tmpTime = tmpTime & " " & _
		Mid(fTime, 9,2) & ":" & _
		Mid(fTime, 11,2) & ":" & _
		Mid(fTime, 13,2) 
		tmpTime = DateAdd("h", strMTimeAdjust , tmpTime)
		tmpTime = FormatDateTime(tmpTime,4)
		aTime = split(tmpTime,":")
		tmpTime = aTime(0) & ":" & aTime(1)
	end if
	ChkTime3 = tmpTime ' & tmpAmPm
  end if
end function

function ChkDate(tIsoDate)
  if not isNumeric(tIsoDate) then
    ChkDate = "" & strCurDate
  else
    year1 = left(tIsoDate,4)
    month1 = mid(tIsoDate,5,2)
    day1 = mid(tIsoDate,7,2)
    if strMCurDateString = tIsoDate then
      ChkDate = "" & DateSerial(year1,month1,day1)
    else
      tpDate = DateSerial(year1,month1,day1) & " " & ChkTime2(tIsoDate)
	  tpDate = DateAdd("h", strMTimeAdjust , tpDate)
	  ChkDate = FormatDateTime(tpDate,2)
	  'ChkDate = strTimeType & " | " & tpDate
	end if
  end if
end function

function chkDate2(tIsoDate)
  cyear1 = left(tIsoDate,4)
  cmonth1 = mid(tIsoDate,5,2)
  cday1 = mid(tIsoDate,7,2)
  chkDate2 = DateSerial(cyear1,cmonth1,cday1)
end function

function getDay(dat)
  tmpDay = mid(dat,7,2)
  if left(tmpDay,1) = "0" then
    tmpDay = right(tmpDay,1)
  end if
  getDay = tmpDay
end function

function getMonth(dat)
  tmpMonth = mid(dat,5,2)
  if left(tmpMonth,1) = "0" then
    tmpMonth = right(tmpMonth,1)
  end if
  getMonth = tmpMonth
end function

function getYear(dat)
  getYear = mid(dat,1,4)
end function
':: end date/time functions
%>
