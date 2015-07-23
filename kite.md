Crazy as a Kite: VBScript Information
=====================================

A few VBScript-isms that might help with your code.

### Statements

"call", "class", "const", "dim", "do", "loop", "erase", "execute", "executeglobal", "exit", "for", "each", "next", "function", "if", "then", "else", "elseif", "on", "error", "resume", "goto", "option", "explicit", "private", "public", "property", "let", "get", "set", "redim", "randomize", "rem", "select", "case", "stop", "sub", "while", "wend", "with", "end", "raise"

### Builtin Functions

"abs", "array", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "chr", "cint", "clng", "conversions", "cos", "createobject", "csng", "cstr", "date", "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day", "escape", "eval", "exp", "filter", "formatcurrency", "formatdatetime", "formatnumber", "formatpercent", "getlocale", "getobject", "getref", "hex", "hour", "inputbox", "instr", "instrrev", "int, fix", "isarray", "isdate", "isempty", "isnull", "isnumeric", "isobject", "join", "lbound", "lcase", "left", "len", "loadpicture", "log", "ltrim; rtrim; and trim", "maths", "mid", "minute", "month", "monthname", "msgbox", "now", "oct", "replace", "rgb", "right", "rnd", "round", "scriptengine", "scriptenginebuildversion", "scriptenginemajorversion", "scriptengineminorversion", "second", "setlocale", "sgn", "sin", "space", "split", "sqr", "strcomp", "string", "strreverse", "tan", "time", "timer", "timeserial", "timevalue", "typename", "ubound", "ucase", "unescape", "vartype", "weekday", "weekdayname", "year"

### Keywords

"null", "empty", "nothing", "true", "false"

### Color Constants

"vbblack", "vbred", "vbgreen", "vbyellow", "vbblue", "vbmagenta", "vbcyan", "vbwhite"

### Compare Constants

"vbbinarycompare", “vbtextcompare"

### Date Constants

"vbsunday", "vbmonday", "vbtuesday", "vbwednesday", "vbthursday", "vbfriday", "vbsaturday", "vbusesystemdayofweek", "vbfirstjan1", "vbfirstfourdays", "vbfirstfullweek"

### Date Format Constants

"vbgeneraldate", "vblongdate", "vbshortdate", "vblongtime", "vbshorttime"

### Miscellaneous Constants

“vbobjecterror"

### MsgBox Constants

"vbokonly", "vbokcancel", "vbabortretryignore", "vbyesnocancel", "vbyesno", "vbretrycancel", "vbcritical", "vbquestion", "vbexclamation", "vbinformation", "vbdefaultbutton1", "vbdefaultbutton2", "vbdefaultbutton3", "vbdefaultbutton4", "vbapplicationmodal", "vbsystemmodal", "vbok", "vbancel", "vbabort", "vbretry", "vbignore", "vbyes", "vbno"

### String Constants

"vbcr", "vbcrlf", "vbformfeed", "vblf", "vbnewline", "vbnullchar", "vbnullstring", "vbtab", “vbverticaltab"

### Tristate Constants

"vbusedefault", "vbtrue", "vbfalse"

### VarType Constants

"vbempty", "vbnull", "vbinteger", "vblong", "vbsingle", "vbdouble", "vbcurrency", "vbdate", "vbstring", "vbobject", "vberror", "vbboolean", "vbvariant", "vbdataobject", "vbdecimal", "vbbyte", "vbarray"

### Operators

"^", "*", "/", "\\", "mod", "+", "-", "&", "=", "<", "<=", ">", ">=", "<>", "and", "not", "or", "xor"

### Line Management

* `_` will extend code to the next line.
* `:` allows you to have multiple statements on the same line, like `dim foo : foo = 3`
* ASP Delimiters: `<%`, `%>`

### Characters

* No such thing.
* Use `Chr(13)` or `ChrW(144)`, for example.

### Strings

* Enclosed in `"`
* Use `""` to escape quotes in strings.
* No other known escape characters: need to do `"str" & vbTab & "str2"`.
* Use `& _` for multi-line strings.

### Integers

* `&h33AA` - hex format
* `&o07` - octal format
* `9944` - normal format.
* VBScript does not handle 64-bit integers.

### Numbers

* Pretty typical - `3.0`, `3e3`.

### Currency, Decimal

* No known literal syntax.
* Currency maps well to a specific .Net type and maps to the Money type in SQL Server.

### Dates

* No time zones on dates.
* There is a date and time literal syntax: `#10-11-2014#`, `#12-12-2012 12:12:12 PM#`, and `#12:33:44#` for instance.

### Exceptions

* Use `raise` to raise an exception.
* You can't catch an exception. All you can do is `on error resume next` and then check `if err then`.
* Use `on error goto 0` to return to normal throw on error behavior.
* It will be very tricky to map VBScript error handling code into a C# exception model. A common paradigm that simulates exceptions from other languages is to wrap what would be in the try block in a function:

Example

	' Foo is sort of like a try block
	Function Foo()
		' raise some errors
	End Function

	' Here is our simulated try/catch
	On error resume next
	Foo()
	If err then
		' Do something about it
		Err.clear
	End if
	On error goto 0

### Calling Functions

* VBScript is very stupid about `Sub` vs. `Function`.
* Using the `Call` statement, it doesn't matter if it's a `sub` or `function`: `Call Foo(x, y, z)`
* To return a value from a fuunction, assign the value to the function name. This does not make it return, it just provides the value to return.

Example:

	function foo()
		foo = 3
		x = 2 ' this still executes
	end function

### Comments

* No multi-line comments
* Can use `REM` or `'`

### Nuggets of Wisdom from [@async_io](https://twitter.com/async_io)

#### VBScript's typelessness is useless episode 1:

	"3" - "2" = 1
	"3" / "2" = 1.5
	"3" * "2" = 6
	"3" + "2" = 32

'nuff said.

#### This is valid vbScript:

	i = 42
	if 0 = 1 then dim i

#### True = -1.

What was wrong with `True` being `not(False)` and `False` being `0` like every other language I have no idea.

#### When calling a sub...

...you do not enclose the argument list in parentheses. When calling a function you do. Except that if the sub only accepts one argument, putting it in parentheses works because its "operator precedence" causes the `(single_arg)` to reduce to `single_arg` first. However, when you have 2 arguments, that no longer works. Also, it is not true if you say `Call MySub`, in which case the `()` becomes mandatory.

#### Operator precedence

The operator precedence and language grammar of vb-based languages is obtuse at best, and horrid and dangerous at worst. It is also completely undocumented. I do not attribute this to malice, but poor design carried out over a number of years. I don't envy the team that had to support the creeping horror of vb syntax up until VB.Net. I can't imagine what the BNF spec for vbScript must look like. I am not sure if one could even be written up.

#### Data Structures

The language actively discourages good use of data structures.  There's no good in-built struct-like object, and you can't statically initialize a dictionary (hashtable) easily using something like `{"foo", 42, "bar", 666}`.

#### Set

If you are assigning an object to a variable, you have to say `set`. This is irritating at best. However, it gets better/worse. If the object you're assigning to the variable has a "default" property on it, the assignment won't actually thrown an error. No, it will put the value of the "default" property into the variable. Then when you try to use the variable later, it's not the object you expected.

#### Arrays

What was so wrong with making both of these return -1?

	dim foo()
	dim bar : bar = Array()
	WScript.stdout.writeline(ubound(foo)) ' Microsoft VBScript runtime error (0x9): Subscript out of range
	WScript.stdout.writeline(ubound(bar)) ' -1

#### Types

Notice that

	const foo = 3

causes `typename(foo)` to return `Integer` and yet

	dim bar(foo)

yields the following error:

	Microsoft VBScript compilation error (0x402): Expected integer constant

even though foo is CONST AND AN INTEGER!!!!!

#### Short-ciruit evaluation

List of languages that do not implement short-circuit evaluation of boolean `and`: vbScript

#### Be really specific

You have to say `end _blocktype_`, but the parser throws an error anyway if you say `end function` when you mean `end sub`

#### Switch

VBScript has a switch/case statement, but unlike every single other language's implementation, it does not do fall-throughs from one case to the next.  This might actually not be such a bad design choice since you can comma-separate multiple possibilities, however it does make it confusing for people with a background in real programming languages.

#### And vs &

Why was it a good idea to make `and` be both the boolean and logical and operator?  `&` was taken for string concatenation, which was moronic enough in its own right, but apparently vb programmers were too stupid to understand `&&` and `||` even though they can all write dhtml javascript which uses those exact operators.  Consequently, if you try to use boolean `and`, don't forget to put parenthese around the expression including the `and`, otherwise you will end up with it most likely being used as a boolean operator.

`if (5 and 4) > 0 then` is what you want, not `if 5 and 4 > 0 then` because the second form will evaluate as `if 5 and if 4 > 0` instead of `if 5 bitwise-and 4 > 0`.

#### Divide by zero

You would think that if you knew for a non-zero integer that you were dividing it by zero, you would know for all cases.  You would be wrong:

	dim foo
	foo = 4/0 ' Microsoft VBScript runtime error (0xB): Division by zero
	foo = 0/0 ' Microsoft VBScript runtime error (0x6): Overflow

#### Sub vs Function

`Sub Foo() : Foo = bar : End Sub` does not generate a "compile time" error even though the interpreter knows you CANNOT RETURN A VALUE FROM A SUB...

#### Okay...

Explain this one:

	c = cint(true) ' no error
	cint(true)     ' Microsoft VBScript runtime error (0x1CA): Variable uses an Automation type not supported in VBScript
	true           ' Microsoft VBScript compilation error (0x400): Expected statement

#### Whitespace matters

Yes, it does.

	dim h0F
	s = s &h0F  ' Microsoft VBScript compilation error (0x401): Expected end of statement
	s = s & h0F ' no error
	s = s &h    ' no error
	s = s &h0   ' Microsoft VBScript compilation error (0x401): Expected end of statement
