<%@ Language=VBScript %>
<%
option explicit
dim set_var : set_var = trUE
%>
<!--#include file="../inc/some_include.asp"-->
<%

dim cc : cc = "42"

set xyz = new PunkRocker

dim x_y_z_

dim s : s = "hello " & _
   "there " & _
   "you " & _
   "coder " & _
   "person"
REM ok this is fun
x = &h32F
x1 = &o777
y = 4
z = 4.33
r = -5e-10-20
d = #12-13-2014#
d = #12/13/2014 3:16:17 am#
if x <> 32 and y <= 16 then
	call foo(32)
end if
mystr = "hello ""there"" you"
' it's 100%
str = "xyz"%>
<%'QAPI Stuff%>
<script type="application/javascript">
alert("<%= "ok" %>")
</script>
<%
str = str &hour(d)

xxx = &h6

dt1 = #31-Dec-1999 21:26:38#
dt2 = #1999-12-31 21:26:38#
dt3 = #12/31/1999 9:26:38 PM#
dt4 = #31-Dec-1999#
dt5 = #21:26:38#

sub foo(x)
	x = 10 REM ok
	x = 20 REM! ok
	y = VbOK
end sub

if I_LIKE_THIS then ccv = "123"

public class Clown
	private x
	public Zero
	public property let Name(s)
		x = len(s)
	end property
	public property get Name()
		Name = x
	end property

	sub class_Initialize
	end sub

	sub class_terminate
	end sub

	sub foo()

if x or y then
	print x
	print y
elseif z then
	print z
else
	print ""
end if

	end sub

	function bar(none)

select case foo
case "1", "2"
	x =1
case "3"
	x = 4
case else
	x = 0
end select

	end function
end class

%>
