<%
  public include, include_vars
  set include = new cls_include

  class cls_include

    private sub class_initialize()
      set include_vars = server.createobject("scripting.dictionary")
    end sub
    private sub class_deactivate()
      arr_variables.removeall
      set include_vars = nothing
      set include = nothing
    end sub

    public default function include(byval str_path)
      dim str_source
      if str_path <> "" then
        str_source = readfile(str_path)
        if str_source <> "" then
          'processincludes str_source
          convert2code str_source
          formatcode str_source
          if str_source <> "" then
              executeglobal str_source
              include_vars.removeall
          end if
        end if
      end if
    end function
	
	Public function writeSource(s)
        if s <> "" then
          convert2code s
          formatcode s
          if s <> "" then
              executeglobal s
              include_vars.removeall
          end if
        end if
	end Function

    private sub convert2code(str_source)
      dim i, str_temp, arr_temp, int_len
      if str_source <> "" then
        if instr(str_source,"%" & ">") > instr(str_source,"<" & "%") then
          str_temp = replace(str_source,"<" & "%","|%")
          str_temp = replace(str_temp,"%" & ">","|")
          if left(str_temp,1) = "|" then str_temp = right(str_temp,len(str_temp) - 1)
          if right(str_temp,1) = "|" then str_temp = left(str_temp,len(str_temp) - 1)
          arr_temp = split(str_temp,"|")
          int_len = ubound(arr_temp)
          if (int_len + 1) > 0 then
            for i = 0 to int_len
              str_temp = trim(arr_temp(i))
              str_temp = replace(str_temp,vbcrlf & vbcrlf,vbcrlf)
              if left(str_temp,2) = vbcrlf then str_temp = right(str_temp,len(str_temp) - 2)
              if right(str_temp,2) = vbcrlf then str_temp = left(str_temp,len(str_temp) - 2)
              if left(str_temp,1) = "%" then
                str_temp = right(str_temp,len(str_temp) - 1)
                if left(str_temp,1) = "=" then
                  str_temp = right(str_temp,len(str_temp) - 1)
                  str_temp = "response.write " & str_temp
                end if
              else
                if str_temp <> "" then
                  include_vars.add i, str_temp
                  str_temp = "response.write include_vars.item(" & i & ")" 
                end if
              end if
              str_temp = replace(str_temp,chr(34) & chr(34) & " & ","")
              str_temp = replace(str_temp," & " & chr(34) & chr(34),"")
              if right(str_temp,2) <> vbcrlf then str_temp = str_temp
              arr_temp(i) = str_temp
            next
            str_source = join(arr_temp,vbcrlf)
          end if
        else
          if str_source <> "" then
            include_vars.add "var", str_source
            str_source = "response.write include_vars.item(""var"")" 
          end if
        end if
      end if
    end sub

    private sub processincludes(str_source)
      dim int_start, str_path, str_mid, str_temp
      str_source = replace(str_source,"<!-- #","<!--#")
      int_start = instr(str_source,"<!--#include")
      str_mid = lcase(getbetween(str_source,"<!--#include","-->"))
      do until int_start = 0
        str_mid = lcase(getbetween(str_source,"<!--","-->"))
        int_start = instr(str_mid,"#include")
        if int_start >  0 then
          str_temp = lcase(getbetween(str_mid,chr(34),chr(34)))
          str_temp = trim(str_temp)
          str_path = readfile(str_temp)
          str_source = replace(str_source,"<!--" & str_mid & "-->",str_path & vbcrlf)
        end if
        int_start = instr(str_source,"#include")
      loop
    end sub

    private sub formatcode(str_code)
      dim i, arr_temp, int_len
      str_code = replace(str_code,vbcrlf & vbcrlf,vbcrlf)
      if left(str_code,2) = vbcrlf then str_code = right(str_code,len(str_code) - 2)
      str_code = trim(str_code)
      if instr(str_code,vbcrlf) > 0 then
        arr_temp = split(str_code,vbcrlf)
        for i = 0 to ubound(arr_temp)
          arr_temp(i) = ltrim(arr_temp(i))
          if arr_temp(i) <> "" then arr_temp(i) = arr_temp(i) & vbcrlf
        next
        str_code = join(arr_temp,"")
        arr_temp = vbnull
      end if
    end sub

    private function readfile(str_path)
      dim objfso, objfile
      if str_path <> "" then
        if instr(str_path,":") = 0 and instr(str_path,"\\") = 0 then str_path = server.mappath(str_path)
        set objfso = server.createobject("scripting.filesystemobject")
        if objfso.fileexists(str_path) then
		  if strUnicode = "YES" then
            set objfile = objfso.opentextfile(str_path, 1, false, -1)
		  else
            set objfile = objfso.opentextfile(str_path, 1, false)
		  end if
          if err.number = 0 then
            readfile = objfile.readall
            objfile.close
		  else
		    readfile = err.Source & " - " & err.description
          end if
          set objfile = nothing
        end if
        set objfso = nothing
      end if
    end function

    private function getbetween(strdata, strstart, strend)
      dim lngstart, lngend
      lngstart = instr(strdata, strstart) + len(strstart)
      if (lngstart <> 0) then
        lngend = instr(lngstart, strdata, strend)
        if (lngend <> 0) then
          getbetween = mid(strdata, lngstart, lngend - lngstart)
        end if
      end if
    end function

  end class
%>