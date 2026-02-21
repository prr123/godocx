// changes prr
// 1. add getter for numbering field of number manager
//

package docx

import (

	"encoding/xml"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"sync"
)

type DocxList struct {
	Ord bool
	AbId int
	Mark [9]string
	Start [9]int
}

// NumInstance represents a w:num element that creates an instance of abstract numbering
type NumInstance struct {
	XMLName       xml.Name `xml:"num"`
	NumId         int      `xml:"numId,attr"`
	AbstractNumId abstNum   `xml:"abstractNumId"`
}

// Numbering represents the numbering part of a Word document
type Numbering struct {
	XMLName   xml.Name       `xml:"numbering"`
	Instances []*NumInstance `xml:"num"`
    List []*NumList  `xml:"abstractNum"`
	NMap map[int]int
}

type abstNum struct {
    XMLName xml.Name `xml:"abstractNumId"`
    Val int `xml:"val,attr"`
}

type NumList struct {
    XMLName xml.Name `xml:"abstractNum"`
    AbstNumId int `xml:"abstractNumId,attr"`
    NsId nsid `xml:"nsid"`
    ML ml `xml:"multiLevelType"`
    Lvl []level `xml:"lvl"`
}

type nsid struct {
    XMLName xml.Name `xml:"nsid"`
    Val string `xml:"val,attr"`
}
type ml struct {
    XMLName xml.Name `xml:"multiLevelType"`
    Val string `xml:"val,attr"`
}

type level struct {
    XMLName xml.Name `xml:"lvl"`
    Ilvl string `xml:"ilvl,attr"`
    Start start `xml:"start"`
    NumFmt numFmt `xml:"numFmt"`
    LvlText lvlText `xml:"lvlText"`
}

type start struct {
    XMLName xml.Name `xml:"start"`
    Val int `xml:"val,attr"`
}

type numFmt struct {
    XMLName xml.Name `xml:"numFmt"`
    Val string `xml:"val,attr"`
}

type lvlText struct {
    XMLName xml.Name `xml:"lvlText"`
    Val string `xml:"val,attr"`
}

// NumberingManager manages numbering instances for the document
type NumberingManager struct {
	mu        sync.Mutex
	numbering *Numbering
	nextNumId int
	rootDoc   *RootDoc
}

// get numbering field
func (nmgr *NumberingManager) GetNumbering ()(nb *Numbering){
	nb = nmgr.numbering
	return nb
}

func (nmgr *NumberingManager) Fill(nb *Numbering, rd *RootDoc){
	nmgr.numbering = nb
	nmgr.nextNumId = len(nb.NMap)+1
	nmgr.rootDoc = rd
	return
}
//              nmgr.numbering = numbObj
//              nmgr.nextNumId = len(numbObj.NMap)+1
//              nmgr.rootDoc = root


// NewNumberingManager creates a new numbering manager
func NewNumberingManager(root *RootDoc) *NumberingManager {
	return &NumberingManager{
		numbering: &Numbering{
			XMLName:   xml.Name{Local: "w:numbering", Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
			Instances: make([]*NumInstance, 0),
		},
		nextNumId: 1,
		rootDoc:   root,
	}
}

// NewNumberingManager creates a new numbering manager
func GetNumberingManager(root *RootDoc, numbObj *Numbering) (nmgr *NumberingManager) {

	nmgr = &NumberingManager{}
	if numbObj != nil {
		nmgr.numbering = numbObj
		nmgr.nextNumId = len(numbObj.NMap)+1
		nmgr.rootDoc = root
		return nmgr
	}
	numbObj = &Numbering{
            XMLName:   xml.Name{Local: "w:numbering", Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
            Instances: make([]*NumInstance, 0),
			}

	nmgr.numbering = numbObj
	nmgr.nextNumId = 1
	nmgr.rootDoc = root
	return nmgr
}

// NewListInstance creates a new numbering instance for the given abstract numbering ID
// Returns the numId that can be used with paragraph.Numbering()
func (nm *NumberingManager) NewListInstance(abstractNumId int) int {
	nm.mu.Lock()
	defer nm.mu.Unlock()

	// Avoid numId collisions with any template-defined w:num entries
	nm.ensureNextNumIdFromTemplate()

	numId := nm.nextNumId
	nm.nextNumId++

	instance := &NumInstance{
		XMLName:       xml.Name{Local: "w:num"},
		NumId:         numId,
		AbstractNumId: abstNum{Val:nm.normalizeAbstract(abstractNumId)},
	}

	nm.numbering.Instances = append(nm.numbering.Instances, instance)
	return numId
}

// GetNumberingXML returns the XML representation of the numbering part
func (nm *NumberingManager) GetNumberingXML() ([]byte, error) {
	nm.mu.Lock()
	defer nm.mu.Unlock()

	if len(nm.numbering.Instances) == 0 {
		return []byte{}, nil
	}

	return xml.Marshal(nm.numbering)
}

// MarshalXML implements xml.Marshaler for NumInstance
func (ni NumInstance) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name.Local = "w:num"
	start.Attr = []xml.Attr{
		{Name: xml.Name{Local: "w:numId"}, Value: fmt.Sprintf("%d", ni.NumId)},
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Write w:abstractNumId element
	abstractNumElem := xml.StartElement{
		Name: xml.Name{Local: "w:abstractNumId"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "w:val"}, Value: fmt.Sprintf("%d", ni.AbstractNumId)},
		},
	}
	if err := e.EncodeElement("", abstractNumElem); err != nil {
		return err
	}

	// Ensure numbering restarts at 1 for level 0 in this instance
	lvlOverride := xml.StartElement{
		Name: xml.Name{Local: "w:lvlOverride"},
		Attr: []xml.Attr{{Name: xml.Name{Local: "w:ilvl"}, Value: "0"}},
	}
	if err := e.EncodeToken(lvlOverride); err != nil {
		return err
	}
	startOverride := xml.StartElement{
		Name: xml.Name{Local: "w:startOverride"},
		Attr: []xml.Attr{{Name: xml.Name{Local: "w:val"}, Value: "1"}},
	}
	if err := e.EncodeElement("", startOverride); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.EndElement{Name: lvlOverride.Name}); err != nil {
		return err
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// applyToFileMap injects generated numbering instances into the existing numbering.xml
// within the root document's file map. It preserves existing abstract numbering definitions
// from the template and appends only the new w:num instance elements.
func (nm *NumberingManager) applyToFileMap() error {
	nm.mu.Lock()
	defer nm.mu.Unlock()

	if len(nm.numbering.Instances) == 0 || nm.rootDoc == nil {
		return nil
	}

	// Determine which instances are already present to avoid duplicates
	// Collect existing numIds from current numbering.xml content if present
	existingIDs := make(map[int]struct{})
	const numberingPath = "word/numbering.xml"
	if existing, ok := nm.rootDoc.FileMap.Load(numberingPath); ok {
		content := string(existing.([]byte))
		idRe := regexp.MustCompile(`w:numId=\"(\d+)\"`)
		for _, m := range idRe.FindAllStringSubmatch(content, -1) {
			if len(m) == 2 {
				if v, err := strconv.Atoi(m[1]); err == nil {
					existingIDs[v] = struct{}{}
				}
			}
		}
	}

	// Build the XML snippet only for missing instances to keep operation idempotent
	var sb strings.Builder
	for _, inst := range nm.numbering.Instances {
		if _, found := existingIDs[inst.NumId]; found {
			continue
		}
		// <w:num w:numId="X"><w:abstractNumId w:val="Y"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
		sb.WriteString("<w:num w:numId=\"")
		sb.WriteString(fmt.Sprintf("%d", inst.NumId))
		sb.WriteString("\"><w:abstractNumId w:val=\"")
		sb.WriteString(fmt.Sprintf("%d", inst.AbstractNumId))
		sb.WriteString("\"/>")
		sb.WriteString("<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride>")
		sb.WriteString("</w:num>")
	}
	instancesXML := sb.String()

	if existing, ok := nm.rootDoc.FileMap.Load(numberingPath); ok {
		// Insert before closing tag of w:numbering
		content := string(existing.([]byte))
		// Ensure our multilevel abstract definitions exist
		content = nm.ensureMultilevelAbstracts(content)
		if instancesXML == "" {
			// Nothing new to add
			nm.rootDoc.FileMap.Store(numberingPath, []byte(content))
			return nil
		}
		if strings.Contains(content, "</w:numbering>") {
			updated := strings.Replace(content, "</w:numbering>", instancesXML+"</w:numbering>", 1)
			nm.rootDoc.FileMap.Store(numberingPath, []byte(updated))
			return nil
		}
		// Fallback: if unexpected structure, append instances at end
		nm.rootDoc.FileMap.Store(numberingPath, []byte(content+instancesXML))
		return nil
	}

	// If numbering.xml doesn't exist (unlikely with the default template), create a minimal one
	minimal := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		nm.multilevelAbstractsXML() + instancesXML + `</w:numbering>`
	nm.rootDoc.FileMap.Store(numberingPath, []byte(minimal))
	return nil
}

// normalizeAbstract maps simple ids used by API to internal multilevel abstract ids.
// 1 -> decimal multilevel, 2 -> bullet multilevel; others remain unchanged.
func (nm *NumberingManager) normalizeAbstract(abstractNumId int) int {
	switch abstractNumId {
	case 1:
		return 201
	case 2:
		return 202
	default:
		return abstractNumId
	}
}

// ensureNextNumIdFromTemplate raises nextNumId above any existing numIds in the template
func (nm *NumberingManager) ensureNextNumIdFromTemplate() {
	const numberingPath = "word/numbering.xml"
	existing, ok := nm.rootDoc.FileMap.Load(numberingPath)
	if !ok {
		return
	}
	content := string(existing.([]byte))
	re := regexp.MustCompile(`w:numId=\"(\d+)\"`)
	maxID := 0
	for _, m := range re.FindAllStringSubmatch(content, -1) {
		if len(m) == 2 {
			if v, err := strconv.Atoi(m[1]); err == nil && v > maxID {
				maxID = v
			}
		}
	}
	if nm.nextNumId <= maxID {
		nm.nextNumId = maxID + 1
	}
}

// ensureMultilevelAbstracts injects multilevel abstract definitions used by normalizeAbstract
func (nm *NumberingManager) ensureMultilevelAbstracts(content string) string {
	if strings.Contains(content, `w:abstractNumId="201"`) && strings.Contains(content, `w:abstractNumId="202"`) {
		return content
	}

	insert := nm.multilevelAbstractsXML()
	// Place right after opening <w:numbering ...>
	start := strings.Index(content, "<w:numbering")
	if start >= 0 {
		end := strings.Index(content[start:], ">")
		if end >= 0 {
			pos := start + end + 1
			return content[:pos] + insert + content[pos:]
		}
	}
	// Fallback: append
	return content + insert
}

func (nm *NumberingManager) multilevelAbstractsXML() string {
	// Build 9 levels for decimal and bullet schemes
	var dec, bul strings.Builder
	dec.WriteString(`<w:abstractNum w:abstractNumId="201"><w:multiLevelType w:val="hybridMultilevel"/>`)
	bul.WriteString(`<w:abstractNum w:abstractNumId="202"><w:multiLevelType w:val="hybridMultilevel"/>`)
	for lvl := 0; lvl < 9; lvl++ {
		pos := 360 * (lvl + 1)
		// Ordered: cycle formats per level: decimal, lowerLetter, lowerRoman
		fmtVal := orderedNumFmtForLevel(lvl)
		dec.WriteString(`<w:lvl w:ilvl="` + strconv.Itoa(lvl) + `"><w:start w:val="1"/><w:numFmt w:val="` + fmtVal + `"/><w:lvlText w:val="%` + strconv.Itoa(lvl+1) + `."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="` + strconv.Itoa(pos) + `"/></w:tabs><w:ind w:left="` + strconv.Itoa(pos) + `" w:hanging="360"/></w:pPr></w:lvl>`)
		// Bullets: cycle glyphs (•, o, square) and fonts (Symbol, Symbol, Wingdings)
		glyph, font := bulletGlyphForLevel(lvl)
		bul.WriteString(`<w:lvl w:ilvl="` + strconv.Itoa(lvl) + `"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="` + glyph + `"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="` + strconv.Itoa(pos) + `"/></w:tabs><w:ind w:left="` + strconv.Itoa(pos) + `" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="` + font + `" w:hAnsi="` + font + `" w:hint="default"/></w:rPr></w:lvl>`)
	}
	dec.WriteString(`</w:abstractNum>`)
	bul.WriteString(`</w:abstractNum>`)
	return dec.String() + bul.String()
}

// orderedNumFmtForLevel returns the WordprocessingML numFmt for a given level,
// cycling through decimal, lowerLetter, and lowerRoman.
func orderedNumFmtForLevel(level int) string {
	switch level % 4 {
	case 0:
		return "decimal"
	case 1:
		return "lowerLetter"
	case 2:
		return "lowerRoman"
	default:
		return "upperLetter"
	}
}

// bulletGlyphForLevel returns bullet glyph and font for each level cycling.
// - 0: • (Symbol "")
// - 1: o (Symbol)
// - 2: ■ (Wingdings "")
// Then repeats.
func bulletGlyphForLevel(level int) (glyph string, font string) {
	switch level % 4 {
	case 0:
		return "", "Symbol" // disc
	case 1:
		return "○", "Symbol" // hollow circle
	case 2:
		return "■", "Wingdings" // square
	default:
		return "♦", "Symbol" // diamond
	}
}


func (nm *NumberingManager)PrintNumObj() {

	num := nm.numbering
    fmt.Println("*** numbering ****")
    fmt.Printf("Name: %s\n",num.XMLName.Local)



    fmt.Println("*** List ****")

    for il:=0; il<len(num.List); il++ {
        nL:= num.List[il]

        fmt.Printf("Name: %s Abs Num: %d\n",nL.XMLName.Local, nL.AbstNumId)
//  fmt.Printf("Abst Num: Id: %d\n", nL.AbstNumId)

        nsid := nL.NsId
        fmt.Println("  *** nsid ****")
        fmt.Printf("  Name: %s Value: %s\n",nsid.XMLName.Local, nsid.Val)
//  fmt.Printf("Value: %s\n",nsid.Val)

        ml := nL.ML
        fmt.Printf("  *** multiLevel ****")
        fmt.Printf("  Name: %s Value:%s\n",ml.XMLName.Local, ml.Val)
//  fmt.Printf("value: %s\n", ml.Val)

        fmt.Printf("  *** levels: %d ****\n", len(nL.Lvl))
        for i:=0; i< len(nL.Lvl); i++ {
            level := nL.Lvl[i]
            fmt.Printf("    *** level %d ***\n", i)
            fmt.Printf("      Name: %s Ilvl: %s\n",level.XMLName.Local, level.Ilvl)
//      fmt.Printf("  ilevel: %s\n", level.Ilvl)
            fmt.Printf("      start name: %s val: %d\n", level.Start.XMLName.Local, level.Start.Val)
            fmt.Printf("      numFmt name: %s val: %s\n", level.NumFmt.XMLName.Local, level.NumFmt.Val)
            fmt.Printf("      lvlTxt name: %s val: %v\n", level.LvlText.XMLName.Local, level.LvlText.Val)
        }
    }

    fmt.Println("*** Num ****")

	for abs, nm := range num.NMap {
		fmt.Printf(" abs: %d numid: %d\n", abs, nm)
	}

	fmt.Println("*** end of PrintList ***")
}

func PrintDocxList(rd *RootDoc) {
// DL *DocxLists
	Dl:= rd.DLists
	fmt.Printf("**** DocxLists: %d ****\n", len(Dl))

	for i:=0; i< len(Dl); i++ {
		dl := Dl[i]
		fmt.Printf("  *** DL: %d ***\n",i+1)
		fmt.Printf("   order: %t\n", dl.Ord)
		fmt.Printf("   Abs Id: %d\n", dl.AbId)

		for il:=0; il< 9; il++ {
			fmt.Printf("    level: %d\n", il)
			fmt.Printf("      mark:  %s\n", dl.Mark[il])
			fmt.Printf("      start: %d\n", dl.Start[il])

		}
	}
}
