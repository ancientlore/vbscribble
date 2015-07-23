// Package vbscanner implements a VBScript scanner.
package vbscanner

import (
	"bufio"
	"bytes"
	"io"
	"strings"
	"unicode"
)

//go:generate stringer -type=TokenType

// TokenType describes the type of token read by the scanner
type TokenType int

// Token types
const (
	EOF     TokenType = iota // end of file
	EOL                      // end of line - needed for vb
	Ident                    // identifiers
	String                   // string literals
	Integer                  // integer literals
	Float                    // float literals
	Date                     // date literals
	Comment                  // comments
	Html                     // the HTML fragments in the ASP
	Char                     // random characters like parens
	Op                       // operators
)

//go:generate stringer -type=Mode

// Mode describes whether we are scanning HTML or VBS code
type Mode int

// Mode types
const (
	HTML_MODE Mode = iota // when reading the HTML part of ASP
	VBS_MODE              // when reading the VBScript part of ASP
)

// Scanner reads a stream and provides VBS tokens scanned from it
type Scanner struct {
	rdr  *bufio.Reader
	mode Mode
	eof  bool
	buf  bytes.Buffer
}

// Init sets up the scanner with the given reader
func (s *Scanner) Init(src io.Reader, initialMode Mode) {
	s.rdr = bufio.NewReader(src)
	s.mode = initialMode
}

// nextIs reads the next rune and returns true if it matches c. If
// it doesn't match, the read rune is unread.
func (s *Scanner) nextIs(c rune) bool {
	r, _, err := s.rdr.ReadRune()
	if err == io.EOF {
		return false
	} else if err != nil {
		panic(err)
	} else if r == c {
		return true
	}
	err = s.rdr.UnreadRune()
	if err != nil {
		panic(err)
	}
	return false
}

// nextIs reads the next rune and returns true if it matches c.
// It always leaves the run unread.
func (s *Scanner) peek(c rune) bool {
	x := false
	r, _, err := s.rdr.ReadRune()
	if err == io.EOF {
		return false
	} else if err != nil {
		panic(err)
	} else if r == c {
		x = true
	}
	err = s.rdr.UnreadRune()
	if err != nil {
		panic(err)
	}
	return x
}

// scanHtml returns the block of HTML up to the next "<%" or end of stream.
func (s *Scanner) scanHtml() string {
	s.buf.Reset()
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if r == '<' && s.nextIs('%') {
			s.mode = VBS_MODE
			return s.buf.String()
		} else {
			s.buf.WriteRune(r)
		}
	}
}

// temporary
func (s *Scanner) scanCode() string {
	s.buf.Reset()
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if r == '%' && s.nextIs('>') {
			s.mode = HTML_MODE
			return s.buf.String()
		} else {
			s.buf.WriteRune(r)
		}
	}
}

func (s *Scanner) Scan() (TokenType, string) {
	if s.eof {
		return EOF, ""
	} else if s.mode == HTML_MODE {
		return Html, s.scanHtml()
	} else {
		for {
			r, _, err := s.rdr.ReadRune()
			if err == io.EOF {
				s.eof = true
				if s.buf.Len() != 0 {
					panic("Premature EOF")
				}
				return EOF, ""
			} else if err != nil {
				panic(err)
			}

			if r == '%' && s.nextIs('>') {
				s.mode = HTML_MODE
				return Html, s.scanHtml()
			} else {
				if r == 'R' || r == 'r' {
					b, err := s.rdr.Peek(3)
					if err == nil {
						str := strings.ToLower(string(b[0:2]))
						if str == "em" && (b[2] == '\t' || b[2] == ' ') {
							b := make([]byte, 3)
							_, err = s.rdr.Read(b)
							if err != nil {
								panic(err)
							}
							return Comment, s.scanComment()
						}
					}
				}
				if unicode.IsLetter(r) {
					s := s.scanIdent(r)
					switch strings.ToLower(s) {
					case "mod", "and", "not", "or", "xor":
						return Op, s
					default:
						return Ident, s
					}
				} else if unicode.IsNumber(r) {
					return s.scanNumber(r)
				} else if r == '&' && s.isHexOctNum() {
					return s.scanNumber(r)
				} else if unicode.IsSpace(r) {
					if r == '\n' {
						return EOL, "\n"
					}
				} else if r == ':' {
					return EOL, ":"
				} else if r == '"' {
					return String, s.scanString()
				} else if r == '\'' {
					return Comment, s.scanComment()
				} else if r == '#' {
					return Date, s.scanDate()
				} else if r == '^' || r == '*' || r == '/' || r == '+' || r == '-' || r == '&' || r == '=' {
					return Op, string(r)
				} else if r == '<' || r == '>' {
					if r == '<' && s.nextIs('>') {
						return Op, "<>"
					} else if s.nextIs('=') {
						return Op, string(r) + "="
					} else {
						return Op, string(r)
					}
				} else if r == '\\' && s.nextIs('\\') {
					return Op, "\\\\"
				} else {
					return Char, string(r)
				}
			}
		}

	}
}

// isHexOctNum returns true if the upcoming bytes (after the already read &) represent a number
func (s *Scanner) isHexOctNum() bool {
	ch, err := s.rdr.Peek(2)
	if err == nil {
		if ch[0] == 'h' || ch[0] == 'H' || ch[0] == 'o' || ch[0] == 'O' {
			if unicode.IsDigit(rune(ch[1])) {
				return true
			}
		}
	}
	return false
}

// scanIdent scans an identifier
func (s *Scanner) scanIdent(c rune) string {
	s.buf.Reset()
	s.buf.WriteRune(c)
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '.' {
			s.buf.WriteRune(r)
		} else {
			err = s.rdr.UnreadRune()
			if err != nil {
				panic(err)
			}
			return s.buf.String()
		}
	}
}

// scanString returns a string
func (s *Scanner) scanString() string {
	s.buf.Reset()
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if r != '"' {
			if r == '\r' || r == '\n' {
				panic("unterminated string constant")
			}
			s.buf.WriteRune(r)
		} else {
			if s.nextIs('"') {
				s.buf.WriteRune(r)
			} else {
				str := s.buf.String()
				if strings.Contains(str, "%>") {
					panic("String contains %>")
				}
				return str
			}
		}
	}
}

// scanDate returns a date
func (s *Scanner) scanDate() string {
	s.buf.Reset()
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if r != '#' {
			const allowed = "/-: \tAPMJANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDECapmjanfebmaraprmayjunjulaugsepoctnovdec"
			if r == '\r' || r == '\n' {
				panic("unterminated Date constant")
			} else if unicode.IsDigit(r) || strings.ContainsRune(allowed, r) {
				s.buf.WriteRune(r)
			} else {
				panic("Invalid date characters")
			}
		} else {
			str := s.buf.String()
			if strings.Contains(str, "%>") {
				panic("Date contains %>")
			}
			return str
		}
	}
}

// scanInteger returns an integer or float
func (s *Scanner) scanNumber(c rune) (TokenType, string) {
	s.buf.Reset()
	s.buf.WriteRune(c)
	var t TokenType = Integer
	first := true
	hex := false
	oct := false
	signReady := false
	gotSign := false
	gotE := false
	gotDot := false
	for {
		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return t, s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if first && r == 'h' || r == 'H' {
			s.buf.WriteRune(r)
			first = false
			hex = true
		} else if first && r == 'o' || r == 'O' {
			s.buf.WriteRune(r)
			first = false
			oct = true
		} else if unicode.IsDigit(r) {
			s.buf.WriteRune(r)
			signReady = false
		} else if hex && (r == 'A' || r == 'B' || r == 'C' || r == 'D' || r == 'E' || r == 'F' || r == 'a' || r == 'b' || r == 'c' || r == 'd' || r == 'e' || r == 'f') {
			s.buf.WriteRune(r)
			signReady = false
		} else if !hex && !oct && !gotDot && r == '.' {
			t = Float
			s.buf.WriteRune(r)
			signReady = false
			gotDot = true
		} else if signReady && !gotSign && (r == '+' || r == '-') {
			gotSign = true
			s.buf.WriteRune(r)
			signReady = false
		} else if !hex && !oct && !gotE && (r == 'e' || r == 'E') {
			t = Float
			s.buf.WriteRune(r)
			signReady = true
			gotE = true
		} else {
			err = s.rdr.UnreadRune()
			if err != nil {
				panic(err)
			}
			return t, s.buf.String()
		}
	}
}

// scanCommebt returns the comment
func (s *Scanner) scanComment() string {
	s.buf.Reset()
	for {
		// check for ASP terminator in comment
		asp, err := s.rdr.Peek(2)
		if err == nil {
			if string(asp) == "%>" {
				return s.buf.String()
			}
		}

		r, _, err := s.rdr.ReadRune()
		if err == io.EOF {
			s.eof = true
			return s.buf.String()
		} else if err != nil {
			panic(err)
		}

		if r != '\n' {
			if r != '\r' {
				s.buf.WriteRune(r)
			}
		} else {
			err = s.rdr.UnreadRune()
			if err != nil {
				panic(err)
			}
			return s.buf.String()
		}
	}
}
