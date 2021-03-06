package main

import (
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/ancientlore/vbscribble/vblexer"
	"github.com/ancientlore/vbscribble/vbscanner"
)

var (
	respWrite = flag.Bool("rw", false, "Use Response.Write formatting")
)

func main() {
	flag.Parse()

	for _, pattern := range flag.Args() {
		files, err := filepath.Glob(pattern)
		if err != nil {
			log.Fatal(err)
		}
		for _, f := range files {
			fi, err := os.Stat(f)
			if err != nil {
				log.Fatal(err)
			}
			if !fi.IsDir() {
				fmt.Fprintln(os.Stderr, "\n*** ", f, " ***")
				fil, err := os.Open(f)
				if err != nil {
					log.Fatal(err)
				}
				func(fil io.Reader, f string) {
					var lex vblexer.Lex
					defer func() {
						if r := recover(); r != nil {
							log.Print("PARSE ERROR ", f, ":", lex.Line, ": ", r)
						}
					}()
					lex.Init(fil, f, vbscanner.HTML_MODE)
					aft := ""
					tabs := 0
					startLine := true
					paren := false
					prevK := vblexer.EOF
					var prevT interface{}
					needStarter := false
					remTabAfterEOL := false
					noTabs := false
					if *respWrite {
						fmt.Print("<%")
					}
					for k, t, v := lex.Lex(); k != vblexer.EOF; k, t, v = lex.Lex() {
						if needStarter {
							if k != vblexer.FILE_INCLUDE && k != vblexer.VIRTUAL_INCLUDE && k != vblexer.HTML {
								fmt.Print("<%")
							}
							needStarter = false
						}
						if startLine {
							if k == vblexer.STATEMENT {
								if t == "End" {
									pv := v
									k, t, v = lex.Lex()
									if k != vblexer.EOF {
										t = "End " + t.(string)
										v = pv + " " + v
										tabs--
										/*
											if t == "End Select" {
												tabs--
											}
										*/
									}
								}
								switch t {
								case "Else", "Elseif", "Case", "Wend", "Next", "Loop":
									tabs--
								}
							}
							if tabs < 0 {
								tabs = 0
							}
							if prevK != vblexer.HTML && !noTabs {
								fmt.Print(strings.Repeat("\t", tabs))
							}
							noTabs = false
							if remTabAfterEOL {
								remTabAfterEOL = false
								tabs--
							}
							startLine = false
							aft = ""
							paren = false
						} else {
							aft = " "
						}
						if paren {
							paren = false
							aft = ""
						}
						if prevK == vblexer.STATEMENT && prevT == "Then" {
							if k != vblexer.EOL && k != vblexer.HTML {
								tabs--
							}
						}
						switch k {
						case vblexer.EOF:
						case vblexer.STATEMENT:
							fmt.Print(aft)
							switch t {
							case "Elseif":
								fmt.Print("ElseIf")
							case "Redim":
								fmt.Print("ReDim")
							case "Executeglobal":
								fmt.Print("ExecuteGlobal")
							case "Wend":
								fmt.Print("WEnd")
							case "Byref":
								fmt.Print("ByRef")
							case "Byval":
								fmt.Print("ByVal")
							default:
								fmt.Print(t)
							}
							switch t {
							case "If", "Function", "Sub", "Class", "Property", "For", "With", "While", "Case": // "Select"
								if !(prevK == vblexer.STATEMENT && prevT == "Exit") {
									tabs++
								}
							case "Else":
								if !(prevK == vblexer.STATEMENT && prevT == "Case") {
									tabs++
								}
							case "Elseif": // "Do"
								tabs++
							}
						case vblexer.FUNCTION:
							fmt.Print(aft)
							fmt.Print(t)
						case vblexer.KEYWORD, vblexer.KEYWORD_BOOL:
							fmt.Print(aft)
							fmt.Print(t)
						case vblexer.COLOR_CONSTANT, vblexer.COMPARE_CONSTANT, vblexer.DATE_CONSTANT, vblexer.DATEFORMAT_CONSTANT, vblexer.MISC_CONSTANT, vblexer.MSGBOX_CONSTANT, vblexer.STRING_CONSTANT, vblexer.TRISTATE_CONSTANT, vblexer.VARTYPE_CONSTANT:
							fmt.Print(aft)
							fmt.Print(t)
						case vblexer.IDENTIFIER:
							fmt.Print(aft)
							fmt.Print(t)
						case vblexer.STRING:
							fmt.Print(aft)
							fmt.Printf("\"%s\"", strings.Replace(v, "\"", "\"\"", -1))
						case vblexer.INT:
							fmt.Print(aft)
							fmt.Print(v)
						case vblexer.FLOAT:
							fmt.Print(aft)
							fmt.Print(v)
						case vblexer.DATE:
							fmt.Print(aft)
							fmt.Print("#", v, "#")
						case vblexer.COMMENT:
							fmt.Print(aft)
							fmt.Printf("' %s", t)
						case vblexer.HTML:
							if *respWrite {
								lines := strings.Split(strings.Replace(v, "\r", "", -1), "\n")
								for index, line := range lines {
									if index == 0 {
										fmt.Println()
										fmt.Print(aft)
										fmt.Print("Response.Write ")
									} else {
										fmt.Print(strings.Repeat("\t", tabs+1))
										fmt.Print("& vbCrLf & ")
									}
									fmt.Printf("\"%s\"\n", strings.Replace(line, "\"", "\"\"", -1))
								}
							} else {
								if prevK != vblexer.EOF && prevK != vblexer.FILE_INCLUDE && prevK != vblexer.VIRTUAL_INCLUDE && prevK != vblexer.HTML {
									fmt.Print(aft)
									fmt.Print("%>")
								}
								fmt.Print(v)
								needStarter = true
							}
							startLine = true
						case vblexer.FILE_INCLUDE:
							if prevK != vblexer.HTML && prevK != vblexer.FILE_INCLUDE && prevK != vblexer.VIRTUAL_INCLUDE {
								fmt.Print(aft)
								fmt.Print("%>")
							}
							if *respWrite {
								fmt.Print(aft)
								fmt.Print("%>")
							}
							fmt.Printf(`<!--#include file="%s"-->`, v)
							if *respWrite {
								fmt.Print("<%")
								fmt.Print(aft)
							}
							needStarter = true
							startLine = true
						case vblexer.VIRTUAL_INCLUDE:
							if prevK != vblexer.HTML && prevK != vblexer.FILE_INCLUDE && prevK != vblexer.VIRTUAL_INCLUDE {
								fmt.Print(aft)
								fmt.Print("%>")
							}
							if *respWrite {
								fmt.Print(aft)
								fmt.Print("%>")
							}
							fmt.Printf(`<!--#include virtual="%s"-->`, v)
							if *respWrite {
								fmt.Print("<%")
								fmt.Print(aft)
							}
							needStarter = true
							startLine = true
						case vblexer.CHAR:
							if prevK == vblexer.STATEMENT || prevK == vblexer.OP {
								fmt.Print(aft)
							}
							fmt.Print(t)
							if t == "(" {
								paren = true
							}
						case vblexer.EOL:
							if t == ":" {
								fmt.Print(aft)
								fmt.Print(t)
								fmt.Print(" ")
								noTabs = true
							} else {
								fmt.Println()
							}
							startLine = true
						case vblexer.OP:
							fmt.Print(aft)
							fmt.Print(t)
						case vblexer.CONTINUATION:
							fmt.Print(aft)
							fmt.Print(t)
							tabs++
							remTabAfterEOL = true
						default:
							panic("Unexpected token type")
						}
						prevK = k
						prevT = t
					}
					if *respWrite {
						fmt.Println("%>")
					}
				}(fil, f)
				fil.Close()
			}
		}
	}
}
