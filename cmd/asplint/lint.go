package main

import (
	"flag"
	"fmt"
	"github.com/ancientlore/vbscribble/vblexer"
	"github.com/ancientlore/vbscribble/vbscanner"
	"io"
	"os"
	"path/filepath"
)

func main() {
	var root string
	flag.StringVar(&root, "root", ".", "Root folder to search")
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
				for k, t, _ := lex.Lex(); k != vblexer.EOF; k, t, _ = lex.Lex() {
					switch k {
					case vblexer.STATEMENT:
						switch t {
						case "Stop":
							messages = append(messages, fmt.Sprintf("%d: Stop should not be used in production code", lex.Line))
						case "Execute", "Executeglobal":
							messages = append(messages, fmt.Sprintf("%d: %s is not recommended", lex.Line, t))
						}
					case vblexer.FUNCTION:
						switch t {
						case "Eval":
							messages = append(messages, fmt.Sprintf("%d: %s is not recommended", lex.Line, t))
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
