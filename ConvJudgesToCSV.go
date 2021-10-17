package main

import (
	"fmt"
	"github.com/droundy/goopt"
	"github.com/pterm/pterm"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"
)

type Judge struct {
	CertNumber int
	LastName string
	FirstName string
	SecondName string
	LastNameEn string
	FirstNameEn string
	Region string
	CityEn string
	City string
	Tel string
	Email string
	Rang string
	CacAbroad bool
}
type Judges []Judge

func (judges *Judges) setJudges(filename string) {
	spinnerJudges, _ := pterm.DefaultSpinner.Start("Loading judges ...")
	var xlsx, err = excelize.OpenFile(filename)
	if err != nil {
		spinnerJudges.Fail()
		panic(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := xlsx.GetRows("Лист1")
	if err != nil {
		spinnerJudges.Fail()
		panic(err)
		return
	}
	// We begin to sort through the lines and add the data in the judges struct
	judge := Judge{}
	for i := 0; i < len(rows); i++ {
		row := rows[i]
		if i > 0 {
			for x := 0; x < len(row); x++ {
				switch x {
				case 0:
					if CertNumber, err := strconv.Atoi(row[x]); err == nil {
						judge.CertNumber = CertNumber
					} else {
						spinnerJudges.Fail()
						panic(err)
						return
					}
				case 1:
					judge.LastName = row[x]
				case 2:
					judge.FirstName = row[x]
				case 3:
					judge.SecondName = row[x]
				case 4:
					judge.LastNameEn = row[x]
				case 5:
					judge.FirstNameEn = row[x]
				case 6:
					judge.Region = row[x]
				case 7:
					judge.CityEn = row[x]
				case 8:
					judge.City = row[x]
				case 9:
					judge.Tel = row[x]
				case 10:
					judge.Email = row[x]
				case 11:
					judge.Rang = row[x]
				case 12:
					if CacAbroad, err := strconv.ParseBool(row[x]); err == nil {
						judge.CacAbroad = CacAbroad
					} else {
						spinnerJudges.Fail()
						panic(err)
						return
					}
				}
			}
			*judges = append(*judges, judge)
		}
	}
	spinnerJudges.Success()
}

type Breed struct {
	CertNumber int
	AllBreeder bool
	GroupId int
	FciNumber int
	NameRus string
	NameEn string
}

type Breeds []Breed

func (breeds *Breeds) setBreeds(filename string) {
	spinnerBreeds, _ := pterm.DefaultSpinner.Start("Loading breeds ...")
	var xlsx, err = excelize.OpenFile(filename)
	if err != nil {
		spinnerBreeds.Fail()
		panic(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := xlsx.GetRows("Лист2")
	if err != nil {
		spinnerBreeds.Fail()
		panic(err)
		return
	}
	// We begin to sort through the lines and add the data in the breeds struct
	breed := Breed{}
	for i := 0; i < len(rows); i++ {
		row := rows[i]
		if i > 0 {
			for x := 0; x < len(row); x++ {
				switch x {
				case 0:
					if CertNumber, err := strconv.Atoi(row[x]); err == nil {
						breed.CertNumber = CertNumber
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 1:
					if AllBreeder, err := strconv.ParseBool(row[x]); err == nil {
						breed.AllBreeder = AllBreeder
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 2:
					if GroupId, err := strconv.Atoi(row[x]); err == nil {
						breed.GroupId = GroupId
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 3:
					if FciNumber, err := strconv.Atoi(row[x]); err == nil {
						breed.FciNumber = FciNumber
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 4:
					breed.NameRus = row[x]
				case 5:
					breed.NameEn = row[x]
				}
			}
			*breeds = append(*breeds, breed)
		}
	}
	spinnerBreeds.Success()
}

func (breeds *Breeds) filter(CertNumber int) []Breed {
	var out []Breed
	for _, rows := range *breeds {
		if rows.CertNumber == CertNumber {
			out = append(out, rows)
		}
	}
	return out
}

type Group struct {
	CertNumber int
	GroupId int
	FciNumber int
	Name string
	NameEng string
}

type Groups []Group

func (groups *Groups) setGroups(filename string) {
	spinnerBreeds, _ := pterm.DefaultSpinner.Start("Loading groups ...")
	var xlsx, err = excelize.OpenFile(filename)
	if err != nil {
		spinnerBreeds.Fail()
		panic(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := xlsx.GetRows("Лист3")
	if err != nil {
		spinnerBreeds.Fail()
		panic(err)
		return
	}
	// We begin to sort through the lines and add the data in the group struct
	group := Group{}
	for i := 0; i < len(rows); i++ {
		row := rows[i]
		if i > 0 {
			for x := 0; x < len(row); x++ {
				switch x {
				case 0:
					if CertNumber, err := strconv.Atoi(row[x]); err == nil {
						group.CertNumber = CertNumber
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 1:
					if GroupId, err := strconv.Atoi(row[x]); err == nil {
						group.GroupId = GroupId
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 2:
					if FciNumber, err := strconv.Atoi(row[x]); err == nil {
						group.FciNumber = FciNumber
					} else {
						spinnerBreeds.Fail()
						panic(err)
						return
					}
				case 3:
					group.Name = row[x]
				case 4:
					group.NameEng = row[x]
				}
			}
			*groups = append(*groups, group)
		}
	}
	spinnerBreeds.Success()
}

func (groups *Groups) filter(CertNumber int) []Group {
	var out []Group
	for _, rows := range *groups {
		if rows.CertNumber == CertNumber {
			out = append(out, rows)
		}
	}
	return out
}

type CSV struct {
	Judge string
	Rank string
	Group string
}

type CSVs []CSV

func ConvXlsxToCsv(FileXlsx, FileCsv string) {
	// Import of data on judges from the first sheet
	judges := Judges{}
	judges.setJudges(FileXlsx)
	sort.SliceStable(judges, func(i, j int) bool {
		return judges[i].LastName < judges[j].LastName
	})
	// Import of rock data from the second sheet
	breeds := Breeds{}
	breeds.setBreeds(FileXlsx)
	// Import of rock data from the second sheet
	groups := Groups{}
	groups.setGroups(FileXlsx)
	spinnerExec, _ := pterm.DefaultSpinner.Start("Starting conversion process ...")
	time.Sleep(time.Second *2)
	// Import of rock data from the second sheet
	spinnerExec.UpdateText("Processing of information ...")
	time.Sleep(time.Second *2)
	var csv []CSVs
	csv = append(csv, []CSV{{"\"Краткая информация о судье\"", "\"Ранг\"", "\"Группа, номер стандарта, название породы, конкурсы\""}})
	for _, rows := range judges {
		var j = fmt.Sprintf("<div>%s %s<br/><br/>(%v) %s %s %s<br/>%s<br/>%s, t. %s</div>",
			strings.ToUpper(rows.LastNameEn),
			strings.ToUpper(rows.FirstNameEn),
			rows.CertNumber,
			strings.ToUpper(rows.LastName),
			strings.ToUpper(rows.FirstName),
			strings.ToUpper(rows.SecondName),
			rows.City,
			rows.Email,
			rows.Tel)
		var r = rows.Rang
		// Selecting and sorting the necessary groups by judge
		var gr = groups.filter(rows.CertNumber)
		sort.SliceStable(gr, func(i, j int) bool {
			return gr[i].Name < gr[j].Name
		})
		var g string
		for _, rec := range gr {
			if g == "" {
				g = fmt.Sprintf("%v,%v,%s/%s", rec.GroupId, rec.FciNumber, rec.Name, rec.NameEng)
			} else {
				g = g + "\n" + fmt.Sprintf("%v,%v,%s/%s", rec.GroupId, rec.FciNumber, rec.Name, rec.NameEng)
			}
		}
		// Select and sort the necessary breeds according to the judge
		var br = breeds.filter(rows.CertNumber)
		sort.SliceStable(br, func(i, j int) bool {
			return br[i].NameRus < br[j].NameRus
		})
		for _, rec := range br {
			if g == "" {
				g = fmt.Sprintf("%v,%v,%s/%s", rec.GroupId, rec.FciNumber, rec.NameRus, rec.NameEn)
			} else {
				g = g + "\n" + fmt.Sprintf("%v,%v,%s/%s", rec.GroupId, rec.FciNumber, rec.NameRus, rec.NameEn)
			}
		}
		csv = append(csv, []CSV{{"\"" + j + "\"", "\"" + r + "\"", "\"" + g + "\""}})
	}
	spinnerExec.UpdateText("Saving the received data to disk ...")
	time.Sleep(time.Second *2)
	// If the array with data is not empty, then export it to a csv file
	if len(csv) > 0 {
		f, err := os.Create(FileCsv)
		if err != nil {
			spinnerExec.Fail()
			panic(err)
		}
		// Close fo on exit and check for its returned error
		defer func() {
			if err := f.Close(); err != nil {
				spinnerExec.Fail()
				panic(err)
			}
		}()
		// Save the data to a file
		for _, rec := range csv {
			// write a chunk
			if _, err := f.WriteString(rec[0].Judge + ";" + rec[0].Rank + ";" + rec[0].Group + "\n"); err != nil {
				spinnerExec.Fail()
				panic(err)
			}
		}
	} else
	{
		spinnerExec.Fail()
		panic("No data to export, buffer is empty")
	}
	spinnerExec.Success("Finally!")
}

var License = `License GPLv3+: GNU GPL version 3 or later <http://gnu.org/licenses/gpl.html>
This is free software: you are free to change and redistribute it.
There is NO WARRANTY, to the extent permitted by law`

func Version() error {
	fmt.Printf("ConvJudgeToCSV 0.1 %s\n\nCopyright (C) 2021 %s\n%s\n", goopt.Version, goopt.Author, License)
	os.Exit(0)
	return nil
}

func PrintUsage() {
	fmt.Fprintf(os.Stderr, goopt.Usage())
	os.Exit(1)
}

func main() {
	goopt.Author = "Dmitry Fofanov"
	goopt.Version = "Link"
	goopt.Summary = "Converting Excel file with data on judges to Csv format"
	goopt.Usage = func() string {
		return fmt.Sprintf("Usage:\t%s FileXlsx FileCsv\n or:\t%s OPTION\n", os.Args[0], os.Args[0]) + goopt.Summary + "\n\n" + goopt.Help()
	}
	goopt.Description = func() string {
		return goopt.Summary + "\n\nUnless an option is passed to it."
	}
	goopt.NoArg([]string{"-v", "--version"}, "outputs version information and exits", Version)
	goopt.Parse(nil)
	if len(goopt.Args) != 2 {
		PrintUsage()
	}
	FileXlsx := goopt.Args[0]
	FileCsv := goopt.Args[1]
	ConvXlsxToCsv(FileXlsx, FileCsv)
}