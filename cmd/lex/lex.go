package main

import (
	"flag"
	"fmt"
	"github.com/ancientlore/vbscribble/vblexer"
	"github.com/ancientlore/vbscribble/vbscanner"
	"io"
	"log"
	"os"
	"path/filepath"
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
					for k, t, v := lex.Lex(); k != vblexer.EOF; k, t, v = lex.Lex() {
						fmt.Printf("%-10s %v %#v\n", k, t, v)
					}
				}(fil, f)
				fil.Close()
			}
		}
	}
}
