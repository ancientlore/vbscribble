package main

import (
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/ancientlore/vbscribble/vblexer"
	"github.com/ancientlore/vbscribble/vbscanner"
)

func main() {
	var root string
	var obj bool
	var objNew bool
	flag.StringVar(&root, "root", ".", "Root folder to search")
	flag.BoolVar(&obj, "obj", false, "Show COM objects used in each file")
	flag.BoolVar(&objNew, "new", false, "Show objects created with new in each file")
	flag.Parse()

	filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			fmt.Println(err.Error())
			return err
		}
		if !info.IsDir() && filepath.Ext(info.Name()) == ".asp" {
			messages := make([]string, 0)
			fil, err := os.Open(path)
			if err != nil {
				return err
			}
			defer fil.Close()
			func(fil io.Reader, f string) {
				var lex vblexer.Lex
				defer func() {
					if r := recover(); r != nil {
						fmt.Println("*** ", f, " ***")
						fmt.Println("PARSE ERROR:", lex.Line, ": ", r)
						fmt.Println()
					}
				}()
				lex.Init(fil, f, vbscanner.HTML_MODE)
				creatingObj := false
				newingObj := false
				for k, t, v := lex.Lex(); k != vblexer.EOF; k, t, v = lex.Lex() {
					switch k {
					case vblexer.STATEMENT:
						switch t {
						case "Stop":
							messages = append(messages, fmt.Sprintf("%d: Statement [Stop] should not be used in production code", lex.Line))
						case "Execute", "Executeglobal":
							messages = append(messages, fmt.Sprintf("%d: Statement [%s] is not recommended", lex.Line, t))
						case "New":
							newingObj = true
						}
					case vblexer.FUNCTION:
						switch t {
						case "Eval":
							messages = append(messages, fmt.Sprintf("%d: Function [%s] is not recommended", lex.Line, t))
						}
					case vblexer.IDENTIFIER:
						switch strings.ToLower(v) {
						case "createobject", "server.createobject", "wscript.createobject":
							creatingObj = true
						default:
							if creatingObj && obj {
								messages = append(messages, fmt.Sprintf("%d: Using object [%s]", lex.Line, v))
							}
							if newingObj && objNew {
								messages = append(messages, fmt.Sprintf("%d: New object [%s]", lex.Line, v))
							}
							creatingObj = false
							newingObj = false
						}
					case vblexer.STRING:
						if creatingObj && obj {
							messages = append(messages, fmt.Sprintf("%d: Using object [%s]", lex.Line, v))
						}
						creatingObj = false
						newingObj = false
					case vblexer.CHAR:
						if !strings.ContainsAny(v, "(),.@!") {
							messages = append(messages, fmt.Sprintf("%d: Unrecognized character [%s]", lex.Line, v))
						}
					}
				}
				if len(messages) > 0 {
					fmt.Println("*** ", f, " ***")
					for _, m := range messages {
						fmt.Println(m)
					}
					fmt.Println()
				}
			}(fil, path)
		}
		return nil
	})
}
