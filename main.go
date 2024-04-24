package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var CardNames []string
var DeckNames []string

var lettermapping = map[int]string{
	0:  "A",
	1:  "B",
	2:  "C",
	3:  "D",
	4:  "E",
	5:  "F",
	6:  "G",
	7:  "H",
	8:  "I",
	9:  "J",
	10: "K",
	11: "L",
	12: "M",
	13: "N",
	14: "O",
	15: "P",
	16: "Q",
	17: "R",
	18: "S",
	19: "T",
	20: "U",
	21: "V",
	22: "W",
	23: "X",
	24: "Y",
	25: "Z",
	26: "AA",
	27: "AB",
	28: "AC",
	29: "AD",
	30: "AE",
	31: "AF",
	32: "AG",
	33: "AH",
	34: "AI",
	35: "AJ",
	36: "AK",
	37: "AL",
	38: "AM",
	39: "AN",
	40: "AO",
	41: "AP",
}

//---------------------------------------------------------------------------------
// SETUP FUNCTIONS
//---------------------------------------------------------------------------------

/**
* Setup Raw Score section
* Params: *excelize.File
* Returns: Row that it ends on
 */
func RawScoreSection(wrkbook *excelize.File, cursheet string) {
	//setup section with repeatables
	Headings(wrkbook, 2, true, cursheet)
	Matchups(wrkbook, 2, cursheet)
	WeightedSum(wrkbook, 2, false, cursheet)
	ScoreVsMean(wrkbook, 2, false, cursheet)
	//setup other info
	//setup boldstyle for bold headings
	validboldstyle := true
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		validboldstyle = false
	}
	//standard headings and key
	wrkbook.SetCellValue(cursheet, "A1", "<Put Name of Sheet Here>")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "A1", "A1", boldstyle)
	}
	wrkbook.SetCellValue(cursheet, "A2", "Blowout = 4pts")
	wrkbook.SetCellValue(cursheet, "A3", "High Impact = 4pts")
	wrkbook.SetCellValue(cursheet, "A4", "Mid Impact = 2 pts")
	wrkbook.SetCellValue(cursheet, "A5", "Low Impact = 1pt")
	wrkbook.SetCellValue(cursheet, "A6", "Non-Factor = 0pts")
	wrkbook.SetCellValue(cursheet, "C2", "# represented")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "C2", "C2", boldstyle)
	}
	wrkbook.SetCellValue(cursheet, "D2", "% frequency")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "D2", "D2", boldstyle)
	}
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(3+len(DeckNames)), "Raw Score")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "B"+strconv.Itoa(3+len(DeckNames)), "B"+strconv.Itoa(3+len(DeckNames)), boldstyle)
	}
}

func HistoricalFrequency(wrkbook *excelize.File, cursheet string, row int) {
	//setup repeatable sections
	Headings(wrkbook, row, false, cursheet)
	Matchups(wrkbook, row, cursheet)
	WeightedSum(wrkbook, row, true, cursheet)
	ScoreVsMean(wrkbook, row, true, cursheet)
	Frequency(wrkbook, row, cursheet)
	SetStandardFormulas(wrkbook, row, cursheet)
	//set up info
	//setup boldstyle for bold headings
	validboldstyle := true
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		validboldstyle = false
	}
	curCell := "A" + strconv.Itoa(row)
	wrkbook.SetCellValue(cursheet, curCell, "Weighted by Historical Frequency")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, curCell, curCell, boldstyle)
	}
	curCell = "A" + strconv.Itoa(row+2)
	wrkbook.SetCellValue(cursheet, curCell, "Round 1 Deck Breakdown")
}

func ProjectedFrequency(wrkbook *excelize.File, cursheet string, row int) {
	//setup repeatable sections
	Headings(wrkbook, row, false, cursheet)
	Matchups(wrkbook, row, cursheet)
	WeightedSum(wrkbook, row, true, cursheet)
	ScoreVsMean(wrkbook, row, true, cursheet)
	SetStandardFormulas(wrkbook, row, cursheet)
	Frequency(wrkbook, row, cursheet)
	//set up info
	//setup boldstyle for bold headings
	validboldstyle := true
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		validboldstyle = false
	}
	curCell := "A" + strconv.Itoa(row)
	wrkbook.SetCellValue(cursheet, curCell, "Weighted by Projected Frequency")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, curCell, curCell, boldstyle)
	}
	//TODO: add more info
}

func MatchupDifficulty(wrkbook *excelize.File, cursheet string, row int) {
	//setup repeatable sections
	Headings(wrkbook, row, false, cursheet)
	Matchups(wrkbook, row, cursheet)
	WeightedSum(wrkbook, row, false, cursheet)
	ScoreVsMean(wrkbook, row, false, cursheet)
	SetStandardFormulas(wrkbook, row, cursheet)
	//setup info
	validboldstyle := true
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		validboldstyle = false
	}
	curCell := "A" + strconv.Itoa(row)
	wrkbook.SetCellValue(cursheet, curCell, "Weighted by Matchup")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, curCell, curCell, boldstyle)
	}
	curCell = "D" + strconv.Itoa(row)
	wrkbook.SetCellValue(cursheet, curCell, "Difficulty")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, curCell, curCell, boldstyle)
	}
	//setup key
	row++
	wrkbook.SetCellValue(cursheet, "A"+strconv.Itoa(row), "Unfavored = 3pts")
	row++
	wrkbook.SetCellValue(cursheet, "A"+strconv.Itoa(row), "Even = 2pts")
	row++
	wrkbook.SetCellValue(cursheet, "A"+strconv.Itoa(row), "Favored = 1pt")
	row++
	wrkbook.SetCellValue(cursheet, "A"+strconv.Itoa(row), "Free = 0pts")

}

func RankCards(wrkbook *excelize.File, cursheet string, row int) {
	//hard coded cells
	rowstr := strconv.Itoa(row)
	validboldstyle := true
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		validboldstyle = false
	}
	//Final Ranking section headings
	wrkbook.SetCellValue(cursheet, "B"+rowstr, "Name")
	wrkbook.SetCellValue(cursheet, "C"+rowstr, "Total")
	wrkbook.SetCellValue(cursheet, "D"+rowstr, "Rank")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "B"+rowstr, "D"+rowstr, boldstyle)
	}
	//Parse ranking section headings
	wrkbook.SetCellValue(cursheet, "I"+rowstr, "Rank")
	wrkbook.SetCellValue(cursheet, "J"+rowstr, "Name")
	wrkbook.SetCellValue(cursheet, "K"+rowstr, "Hist. Freq.")
	wrkbook.SetCellValue(cursheet, "L"+rowstr, "Pred. Freq.")
	wrkbook.SetCellValue(cursheet, "M"+rowstr, "M/U Diff")
	wrkbook.SetCellValue(cursheet, "N"+rowstr, "Total")
	if validboldstyle {
		wrkbook.SetCellStyle(cursheet, "I"+rowstr, "N"+rowstr, boldstyle)
	}
	//move to next row in preparation for looping
	row++

	//set up parse ranking section
	i := row
	rowstr = strconv.Itoa(row)
	curCell := ""
	formula := ""
	//I column (ranking)
	for i = row; i < row+len(CardNames); i++ {
		curCell = "N" + strconv.Itoa(i)
		formula = "RANK(" + curCell + ",N" + rowstr + ":N" + strconv.Itoa(row+len(CardNames)-1) + ",0)"
		curCell = "I" + strconv.Itoa(i)
		wrkbook.SetCellFormula(cursheet, curCell, formula)
	}
	//J column (name)
	index := 4
	for i = row; i < row+len(CardNames); i++ {
		formula = lettermapping[index] + "2"
		wrkbook.SetCellFormula(cursheet, "J"+strconv.Itoa(i), formula)
		index++
	}
	//K column (Historical Frequency)
	index = 4
	//Historical Frequency row calculation
	histrowint := 2 + len(DeckNames) + 4 + len(DeckNames) + 3
	histrow := strconv.Itoa(histrowint)
	for i = row; i < row+len(CardNames); i++ {
		formula = lettermapping[index] + histrow
		wrkbook.SetCellFormula(cursheet, "K"+strconv.Itoa(i), formula)
		index++
	}
	//L column (Projected Frequency)
	index = 4
	//Projected Frequency row calculation
	projrowint := histrowint + len(DeckNames) + 5
	projrow := strconv.Itoa(projrowint)
	for i = row; i < row+len(CardNames); i++ {
		formula = lettermapping[index] + projrow
		wrkbook.SetCellFormula(cursheet, "L"+strconv.Itoa(i), formula)
		index++
	}
	//M column (Weighted Matchup (Matchup Difficulty))
	index = 4
	//Weighted Matchup row calculation
	matchrowint := projrowint + len(DeckNames) + 4
	matchrow := strconv.Itoa(matchrowint)
	for i = row; i < row+len(CardNames); i++ {
		formula = lettermapping[index] + matchrow
		wrkbook.SetCellFormula(cursheet, "M"+strconv.Itoa(i), formula)
		index++
	}
	//N column (Total)
	for i = row; i < row+len(CardNames); i++ {
		curint := strconv.Itoa(i)
		formula = "SUM(K" + curint + ":M" + curint + ")"
		wrkbook.SetCellFormula(cursheet, "N"+curint, formula)
	}

	//Final rankings section
	//B column (Rank)
	index = 1
	for i = row; i < row+len(CardNames); i++ {
		curint := strconv.Itoa(i)
		formula = "VLOOKUP(" + strconv.Itoa(index) + ",I" + strconv.Itoa(row) + ":J" + strconv.Itoa(row+len(CardNames)-1) + ",2,FALSE )"
		wrkbook.SetCellFormula(cursheet, "B"+curint, formula)
		index++
	}
	//C column (Total) and D column (rank)
	index = 1
	for i = row; i < row+len(CardNames); i++ {
		curint := strconv.Itoa(i)
		formula = "VLOOKUP(" + strconv.Itoa(index) + ",I" + strconv.Itoa(row) + ":N" + strconv.Itoa(row+len(CardNames)-1) + ",6,FALSE )"
		wrkbook.SetCellFormula(cursheet, "C"+curint, formula)
		wrkbook.SetCellValue(cursheet, "D"+curint, strconv.Itoa(index))
		index++
	}

}

//---------------------------------------------------------------------------------
// HELPER FUNCTIONS
//---------------------------------------------------------------------------------

func Headings(wrkbook *excelize.File, row int, start bool, cursheet string) {
	//start from E, go to end of thing
	curcell := ""
	i := 0
	if start {
		for i = 0; i < len(CardNames); i++ {
			curcell = lettermapping[4+i] + strconv.Itoa(row)
			wrkbook.SetCellValue(cursheet, curcell, CardNames[i])
		}
	} else {
		for i = 0; i < len(CardNames); i++ {
			curcell = lettermapping[4+i] + strconv.Itoa(row)
			formula := lettermapping[4+i] + "2"
			wrkbook.SetCellFormula(cursheet, curcell, formula)
		}
	}
}

func Matchups(wrkbook *excelize.File, row int, cursheet string) {
	//set up matchups in B column
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(row), "Matchup")
	//setup bold style
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err == nil {
		wrkbook.SetCellStyle(cursheet, "B"+strconv.Itoa(row), "B"+strconv.Itoa(row), boldstyle)
	}
	row += 1
	index := row
	for index < row+len(DeckNames) {
		wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(index), DeckNames[index-row])
		index++
	}

}

func SetStandardFormulas(wrkbook *excelize.File, row int, cursheet string) {
	//setup number format
	numstyle, err := wrkbook.NewStyle(`{"custom_number_format": "0.0000_ "}`)
	//separate counter for looping through formula
	index := 3
	//set row to next row
	row++
	//first loop: row
	for i := row; i < len(DeckNames)+row; i++ {
		//second loop: column
		for j := 4; j < len(CardNames)+4; j++ {
			curCell := lettermapping[j] + strconv.Itoa(i)
			formula := "D" + strconv.Itoa(i) + "*" + lettermapping[j] + strconv.Itoa(index)
			wrkbook.SetCellFormula(cursheet, curCell, formula)
			if err == nil {
				wrkbook.SetCellStyle(cursheet, curCell, curCell, numstyle)
			}
		}
		index++
	}
}

func WeightedSum(wrkbook *excelize.File, row int, space bool, cursheet string) {
	thisrow := 0
	if !space {
		thisrow = row + len(DeckNames) + 1
	} else {
		thisrow = row + len(DeckNames) + 2
	}
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(thisrow), "Weighted Sum")
	//setup bold style
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err == nil {
		wrkbook.SetCellStyle(cursheet, "B"+strconv.Itoa(thisrow), "B"+strconv.Itoa(thisrow), boldstyle)
	}
	for j := 4; j < len(CardNames)+4; j++ {
		curCell := lettermapping[j] + strconv.Itoa(thisrow)
		formula := "SUM(" + lettermapping[j] + strconv.Itoa(row+1) + ":"
		if !space {
			formula = formula + lettermapping[j] + strconv.Itoa(thisrow-1) + ")"
		} else {
			formula = formula + lettermapping[j] + strconv.Itoa(thisrow-2) + ")"
		}
		wrkbook.SetCellFormula(cursheet, curCell, formula)
	}
}

func ScoreVsMean(wrkbook *excelize.File, row int, space bool, cursheet string) {
	//setup number format
	numstyle, err := wrkbook.NewStyle(`{"custom_number_format": "0.0000_ "}`)
	//recalculate row to correct place
	if space {
		row += len(DeckNames) + 3
	} else {
		row += len(DeckNames) + 2
	}
	//set name of row
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(row), "Score Vs Mean")
	//setup bold style
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err == nil {
		wrkbook.SetCellStyle(cursheet, "B"+strconv.Itoa(row), "B"+strconv.Itoa(row), boldstyle)
	}
	//set row string
	formstr := strconv.Itoa(row - 1)
	//grab weighted sums and conver to formula
	formula := "SUM(E" + formstr + ":" + lettermapping[len(CardNames)+3] + formstr + ")/" + strconv.Itoa(len(CardNames))
	//save formula cell to save work later
	formcell := "C" + strconv.Itoa(row)
	wrkbook.SetCellFormula(cursheet, formcell, formula)
	//loop through and set score vs mean formula
	for j := 4; j < len(CardNames)+4; j++ {
		formula = lettermapping[j] + formstr + "/" + formcell
		wrkbook.SetCellFormula(cursheet, lettermapping[j]+strconv.Itoa(row), formula)
		if err == nil {
			wrkbook.SetCellStyle(cursheet, lettermapping[j]+strconv.Itoa(row), lettermapping[j]+strconv.Itoa(row), numstyle)
		}
	}

}

func Frequency(wrkbook *excelize.File, row int, cursheet string) {
	//setup number format
	numstyle, err := wrkbook.NewStyle(`{"custom_number_format": "0.0000_ "}`)
	//set Total Number of Decks
	thisrow := row + len(DeckNames) + 1
	formula := "SUM(C" + strconv.Itoa(row+1) + ":C" + strconv.Itoa(thisrow-1) + ")"
	wrkbook.SetCellValue(cursheet, "B"+strconv.Itoa(thisrow), "Total # of Decks")
	//setup bold style
	boldstyle, err := wrkbook.NewStyle(`{"font":{"bold":true}}`)
	if err == nil {
		wrkbook.SetCellStyle(cursheet, "B"+strconv.Itoa(thisrow), "B"+strconv.Itoa(thisrow), boldstyle)
	}
	wrkbook.SetCellFormula(cursheet, "C"+strconv.Itoa(thisrow), formula)
	//set new frequency formulas
	row += 1
	curCell := ""
	for row < thisrow {
		curCell = "D" + strconv.Itoa(row)
		formula = "C" + strconv.Itoa(row) + "/C" + strconv.Itoa(thisrow)
		wrkbook.SetCellFormula(cursheet, curCell, formula)
		if err == nil {
			wrkbook.SetCellStyle(cursheet, curCell, curCell, numstyle)
		}
		row++
	}
}

func main() {

	reader := bufio.NewReader(os.Stdin)
	fmt.Print("Enter path to save to (google 'find path of file ' + the type of operating system you're using if you're not sure what this means): ")
	savetothisfile, _ := reader.ReadString('\n')
	savetothisfile = "C:/Users/johnn/Documents/projects-yugioh/datasheets/blanksheet.xlsx"

	//open defensive cards text file
	file, err := os.Open("C:/Users/johnn/Documents/projects-yugioh/datasheets/defensivecards.txt")

	if err != nil {
		panic(err)
	}

	scanner := bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)

	//initialize CardNames slice

	for scanner.Scan() {
		if CardNames == nil {
			CardNames = []string{scanner.Text()}
		} else {
			CardNames = append(CardNames, scanner.Text())
		}
	}

	file.Close()

	//open decks text file
	file, err = os.Open("C:/Users/johnn/Documents/projects-yugioh/datasheets/metadecks.txt")

	if err != nil {
		panic(err)
	}

	scanner = bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)

	//initialize CardNames slice

	for scanner.Scan() {
		if DeckNames == nil {
			DeckNames = []string{scanner.Text()}
		} else {
			DeckNames = append(DeckNames, scanner.Text())
		}
	}

	file.Close()

	wrkbook := excelize.NewFile()
	cursheet := "Sheet1"
	wrkbook.NewSheet(cursheet)
	index := 2
	RawScoreSection(wrkbook, cursheet)
	index += len(DeckNames) + 4
	HistoricalFrequency(wrkbook, cursheet, index)
	index += len(DeckNames) + 5
	ProjectedFrequency(wrkbook, cursheet, index)
	index += len(DeckNames) + 5
	MatchupDifficulty(wrkbook, cursheet, index)
	index += len(DeckNames) + 4
	RankCards(wrkbook, cursheet, index)
	err = wrkbook.SaveAs(savetothisfile)
	if err != nil {
		panic(err)
	}
	fmt.Println("check sheet")
}
