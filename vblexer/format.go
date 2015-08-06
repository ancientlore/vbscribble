package vblexer

import (
	"strings"
)

func formatConstant(c string) string {
	c = strings.TrimPrefix(c, "vb")
	switch c {
	case "binarycompare":
		return "vbBinaryCompare"
	case "textcompare":
		return "vbTextCompare"
	case "usesystemdayofweek":
		return "vbUseSystemDayOfWeek"
	case "firstjan1":
		return "vbFirstJan1"
	case "firstfourdays":
		return "vbFirstFourDays"
	case "firstfullweek":
		return "vbFirstFullWeek"
	case "generaldate":
		return "vbGeneralDate"
	case "longdate":
		return "vbLongDate"
	case "shortdate":
		return "vbShortDate"
	case "longtime":
		return "vbLongTime"
	case "shorttime":
		return "vbShortTime"
	case "objecterror":
		return "vbObjectError"
	case "okonly":
		return "vbOkOnly"
	case "okcancel":
		return "vbOkCancel"
	case "abortretryignore":
		return "vbAbortRetryIgnore"
	case "yesnocancel":
		return "vbYesNoCancel"
	case "yesno":
		return "vbYesNo"
	case "retrycancel":
		return "vbRetryCancel"
	case "defaultbutton1", "defaultbutton2", "defaultbutton3", "defaultbutton4":
		return "vbDefaultButton" + strings.TrimPrefix(c, "defaultbutton")
	case "applicationmodal":
		return "vbApplicationModal"
	case "systemmodal":
		return "vbSystemModal"
	case "crlf":
		return "vbCrLf"
	case "formfeed":
		return "vbFormFeed"
	case "newline":
		return "vbNewLine"
	case "nullchar":
		return "vbNullChar"
	case "nullstring":
		return "vbNullString"
	case "verticaltab":
		return "vbVerticalTab"
	case "usedefault":
		return "vbUseDefault"
	case "dataobject":
		return "vbDataObject"
	default:
		return "vb" + strings.Title(c)
	}
}

func formatFunction(f string) string {
	switch f {
	case "cbool", "cbyte", "cdate", "cdbl", "cint", "clng", "csng", "cstr":
		return "C" + strings.Title(strings.TrimPrefix(f, "c"))
	case "ccur":
		return "CCur"
	case "createobject":
		return "CreateObject"
	case "dateadd", "datediff", "datepart", "dateserial", "datevalue":
		return "Date" + strings.Title(strings.TrimPrefix(f, "date"))
	case "formatcurrency", "formatnumber", "formatpercent":
		return "Format" + strings.Title(strings.TrimPrefix(f, "format"))
	case "formatdatetime":
		return "FormatDateTime"
	case "getlocale", "getobject", "getref":
		return "Get" + strings.Title(strings.TrimPrefix(f, "get"))
	case "inputbox":
		return "InputBox"
	case "instr":
		return "InStr"
	case "instrrev":
		return "InStrRev"
	case "isarray", "isdate", "isempty", "isnull", "isnumeric", "isobject":
		return "Is" + strings.Title(strings.TrimPrefix(f, "is"))
	case "lbound", "lcase", "ltrim":
		return "L" + strings.Title(strings.TrimPrefix(f, "l"))
	case "loadpicture":
		return "LoadPicture"
	case "rtrim":
		return "RTrim"
	case "maths":
		return "MathS"
	case "monthname":
		return "MonthName"
	case "msgbox":
		return "MsgBox"
	case "scriptengine":
		return "ScriptEngine"
	case "scriptenginebuildversion":
		return "ScriptEngineBuildVersion"
	case "scriptenginemajorversion":
		return "ScriptEngineMajorVersion"
	case "scriptengineminorversion":
		return "ScriptEngineMinorVersion"
	case "setlocale":
		return "SetLocale"
	case "strcomp", "strreverse":
		return "Str" + strings.Title(strings.TrimPrefix(f, "str"))
	case "timeserial", "timevalue":
		return "Time" + strings.Title(strings.TrimPrefix(f, "time"))
	case "typename":
		return "TypeName"
	case "ubound", "ucase":
		return "U" + strings.Title(strings.TrimPrefix(f, "u"))
	case "unescape":
		return "UnEscape"
	case "vartype":
		return "VarType"
	case "weekday":
		return "WeekDay"
	case "weekdayname":
		return "WeekDayName"
	default:
		return strings.Title(f)
	}
}
