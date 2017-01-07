package main

import (
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"

	"github.com/ancientlore/vbscribble/vblexer"
	"github.com/ancientlore/vbscribble/vbscanner"
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
				fmt.Println("\n*** ", f, " ***")
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
					for k, t, _ := lex.Lex(); k != vblexer.EOF; k, t, _ = lex.Lex() {
						fmt.Printf("%s%v", aft, t)
						if k == vblexer.EOL {
							aft = ""
						} else {
							aft = " "
						}
					}
				}(fil, f)
				fil.Close()
			}
		}
	}
}
