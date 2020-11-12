<%
if rs.state <> 0 then rs.close

set rs = nothing
set cmd = nothing
conn.close
set conn = nothing

%>