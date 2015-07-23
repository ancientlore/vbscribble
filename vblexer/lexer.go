// Package vblexer implements a VBScript lexer.
package vblexer

import (
	"github.com/ancientlore/vbscribble/vbscanner"
	"io"
	"strconv"
	"strings"
	"time"
)

//go:generate stringer -type=TokenType

// TokenType describes the type of token read by the scanner
type TokenType int

// Token types
const (
	EOF                 TokenType = iota // end of file
	STATEMENT                            // language statements (reserved words)
	FUNCTION                             // builtin functions
	KEYWORD                              // keywords - nothing, null, empty
	KEYWORD_BOOL                         // boolean keywords - true and false
	COLOR_CONSTANT                       // constants for colors, like vbRed
	COMPARE_CONSTANT                     // constants to specify binary or text comparison
	DATE_CONSTANT                        // constants relating to dates, like vbSunday
	DATEFORMAT_CONSTANT                  // constants for specifying date formats
	MISC_CONSTANT                        // miscellaenous constants
	MSGBOX_CONSTANT                      // constants for message boxes like MbOk
	STRING_CONSTANT                      // constant strings like vbCrLf
	TRISTATE_CONSTANT                    // constants for tristate use
	VARTYPE_CONSTANT                     // constants for variant types
	IDENTIFIER                           // identifiers
	STRING                               // string literals
	INT                                  // integer literals
	FLOAT                                // float literals
	DATE                                 // date literals
	COMMENT                              // comments
	HTML                                 // HTML fragments in the ASP
	CHAR                                 // random characters like parens
	EOL                                  // end of line
	OP                                   // operator
)

// Lex uses a scanner to read and classify VBScript tokens
type Lex struct {
	s        vbscanner.Scanner
	Filename string
	Line     int
}

// Init prepares the lexer for use.
func (lex *Lex) Init(src io.Reader, fname string, initialMode vbscanner.Mode) {
	lex.s.Init(src, initialMode)
	lex.Filename = fname
	lex.Line = 1
}

// Lex returns the next token in the steam and classifies it. The values returned are
// the token type, the converted value, and the raw value as a string.
func (lex *Lex) Lex() (TokenType, interface{}, string) {
	tok, value := lex.s.Scan()

	// return tok.String(), value
	if tok == vbscanner.EOF {
		return EOF, nil, value
	}

	switch tok {
	case vbscanner.Ident:
		stoken := strings.ToLower(value)
		switch stoken {
		// statements
		case "call", "class", "const", "dim", "do", "loop", "erase", "execute", "executeglobal", "exit", "for", "each", "next", "function", "if", "then", "else", "elseif", "on", "error", "resume", "goto", "option", "explicit", "private", "public", "property", "let", "get", "set", "redim", "randomize", "rem", "select", "case", "stop", "sub", "while", "wend", "with", "end", "raise":
			return STATEMENT, stoken, value

		// functions
		case "abs", "array", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "chr", "cint", "clng", "conversions", "cos", "createobject", "csng", "cstr", "date", "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day", "escape", "eval", "exp", "filter", "formatcurrency", "formatdatetime", "formatnumber", "formatpercent", "getlocale", "getobject", "getref", "hex", "hour", "inputbox", "instr", "instrrev", "int, fix", "isarray", "isdate", "isempty", "isnull", "isnumeric", "isobject", "join", "lbound", "lcase", "left", "len", "loadpicture", "log", "ltrim; rtrim; and trim", "maths", "mid", "minute", "month", "monthname", "msgbox", "now", "oct", "replace", "rgb", "right", "rnd", "round", "scriptengine", "scriptenginebuildversion", "scriptenginemajorversion", "scriptengineminorversion", "second", "setlocale", "sgn", "sin", "space", "split", "sqr", "strcomp", "string", "strreverse", "tan", "time", "timer", "timeserial", "timevalue", "typename", "ubound", "ucase", "unescape", "vartype", "weekday", "weekdayname", "year":
			return FUNCTION, stoken, value

		// Keywords
		case "null", "empty", "nothing":
			return KEYWORD, stoken, value
		case "true", "false":
			b, err := strconv.ParseBool(value)
			if err != nil {
				panic(err)
			}
			return KEYWORD_BOOL, b, value

		// constants
		case "vbblack", "vbred", "vbgreen", "vbyellow", "vbblue", "vbmagenta", "vbcyan", "vbwhite":
			return COLOR_CONSTANT, stoken, value
		case "vbbinarycompare", "vbtextcompare":
			return COMPARE_CONSTANT, stoken, value
		case "vbsunday", "vbmonday", "vbtuesday", "vbwednesday", "vbthursday", "vbfriday", "vbsaturday", "vbusesystemdayofweek", "vbfirstjan1", "vbfirstfourdays", "vbfirstfullweek":
			return DATE_CONSTANT, stoken, value
		case "vbgeneraldate", "vblongdate", "vbshortdate", "vblongtime", "vbshorttime":
			return DATEFORMAT_CONSTANT, stoken, value
		case "vbobjecterror":
			return MISC_CONSTANT, stoken, value
		case "vbokonly", "vbokcancel", "vbabortretryignore", "vbyesnocancel", "vbyesno", "vbretrycancel", "vbcritical", "vbquestion", "vbexclamation", "vbinformation", "vbdefaultbutton1", "vbdefaultbutton2", "vbdefaultbutton3", "vbdefaultbutton4", "vbapplicationmodal", "vbsystemmodal", "vbok", "vbancel", "vbabort", "vbretry", "vbignore", "vbyes", "vbno":
			return MSGBOX_CONSTANT, stoken, value
		case "vbcr", "vbcrlf", "vbformfeed", "vblf", "vbnewline", "vbnullchar", "vbnullstring", "vbtab", "vbverticaltab":
			return STRING_CONSTANT, stoken, value
		case "vbusedefault", "vbtrue", "vbfalse":
			return TRISTATE_CONSTANT, stoken, value
		case "vbempty", "vbnull", "vbinteger", "vblong", "vbsingle", "vbdouble", "vbcurrency", "vbdate", "vbstring", "vbobject", "vberror", "vbboolean", "vbvariant", "vbdataobject", "vbdecimal", "vbbyte", "vbarray":
			return VARTYPE_CONSTANT, stoken, value
		default:
			return IDENTIFIER, value, value
		}

	// Values
	case vbscanner.String:
		return STRING, value, value
	case vbscanner.Integer:
		base := 10
		str := strings.ToLower(value)
		if strings.HasPrefix(str, "&h") {
			base = 16
			str = strings.TrimPrefix(str, "&h")
		} else if strings.HasPrefix(strings.ToLower(str), "&o") {
			base = 8
			str = strings.TrimPrefix(str, "&o")
		}
		i, err := strconv.ParseInt(str, base, 64)
		if err != nil {
			panic(err)
		}
		return INT, i, value
	case vbscanner.Float:
		f, err := strconv.ParseFloat(value, 64)
		if err != nil {
			panic(err)
		}
		return FLOAT, f, value
	case vbscanner.Date:
		formats := []string{
			"02-Jan-2006 15:04:05",  // #31-Dec-1999 21:26:00#
			"02-JAN-2006 15:04:05",  // #31-DEC-1999 21:26:00#
			"2006-01-02 15:04:05",   // #1999-12-31 21:26:00#
			"01/02/2006 3:04:05 PM", // #12/31/1999 9:26:00 PM#
			"01/02/2006 3:04:05 pm", // #12/31/1999 9:26:00 pm#
			"02-Jan-2006",           // #31-Dec-1999#
			"02-JAN-2006",           // #31-DEC-1999#
			"2006-01-02",            // #1999-12-31#
			"01/02/2006",            // #12/31/1999#
			"01-02-2006",            // #12-31-1999#
			"15:04:05",              // #21:26:00#
		}
		loc, _ := time.LoadLocation("America/New_York")
		for _, fmt := range formats {
			t, err := time.ParseInLocation(fmt, value, loc)
			if err == nil {
				return DATE, t, value
			}
		}
		return DATE, value, value
	case vbscanner.Comment:
		return COMMENT, value, value
	case vbscanner.Html:
		return HTML, value, value
	case vbscanner.Char:
		return CHAR, value, value
	case vbscanner.EOL:
		lex.Line++
		return EOL, value, value
	case vbscanner.Op:
		return OP, value, value
	}

	panic("How did we get here?")
	// return "IDENT", value
}