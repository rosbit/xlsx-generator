// excel表cell列编号生成器
// A, B, C, ..., Z, AA, AB, AC, ..., AZ, BA, BB, BC, ..., BZ, CA, ......, ZA, ZB, ZC, ..., ZZ
// 最多26*27 = 702列，暂不支持3个字母以上的列编号

package toxlsx

import (
	"fmt"
)

type columnGenerator struct {
	g chan string
	exit bool
	lastCol string
}

func NewColumnGenerator() *columnGenerator {
	g := &columnGenerator{
		g: make(chan string),
		exit: false,
	}
	go g.generate()
	return g
}

func (g *columnGenerator) Next() string {
	return <-g.g
}

func (g *columnGenerator) Stop() {
	g.exit = true
	<-g.g
}

func (g *columnGenerator) Last() string {
	return g.lastCol
}

func (g *columnGenerator) generate() {
	var c1, c2 byte
	c2 = 'A'

	for !g.exit {
		g.lastCol = output(c1, c2)
		g.g <- g.lastCol

		incByte(&c2)
		if c2 == 'A' {
			incByte(&c1)
		}
	}

	close(g.g)
}

func output(c1, c2 byte) string {
	if c1 == 0 {
		return fmt.Sprintf("%c", c2)
	}
	return fmt.Sprintf("%c%c", c1, c2)
}

func incByte(c *byte) {
	switch {
	case *c == 0 || *c == 'Z':
		*c = 'A'
	default:
		*c += 1
	}
}
