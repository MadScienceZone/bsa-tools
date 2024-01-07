package main

import (
	"bufio"
	"fmt"
	"os"
	"sort"
	"time"
)

type MeritBadge struct {
	Name     string
	Earned   time.Time
	Required bool
	Counted  bool
}

type ByDateEarned []MeritBadge

func (a ByDateEarned) Len() int           { return len(a) }
func (a ByDateEarned) Less(i, j int) bool { return a[i].Earned.Before(a[i].Earned) }
func (a ByDateEarned) Swap(i, j int)      { a[i], a[j] = a[j], a[i] }

func main() {
	fmt.Println(`
--------------------------------------------------------------------------------
This is a basic, quick-and-dirty tool to verify the dates entered on the BSA
Eagle Scout application form. Enter each date in MMDDYY form (e.g., enter 
010224 for January 2, 2024). If you just enter a period (".") in place of a
date, that means to repeat the previously entered date for this badge as well.

As such a simple program that I wrote in a couple of minutes of spare time,
it's not fancy enough to do sophisticated data entry or allow you to edit any-
thing; if you make a mistake, type "Q" instead of a date to stop early, and
run the program again (when entering merit badge dates, typing "Q" will go
back to the start of the merit badge list, so you can change mistakes in that
list).
--------------------------------------------------------------------------------`)

	input := bufio.NewScanner(os.Stdin)
	badges := []MeritBadge{
		{Name: "01 Camping", Required: true, Counted: false},
		{Name: "02 Citizenship in the Community", Required: true, Counted: false},
		{Name: "03 Citizenship in the Nation", Required: true, Counted: false},
		{Name: "04 Citizenship in Society", Required: true, Counted: false},
		{Name: "05 Citizenship in the World", Required: true, Counted: false},
		{Name: "06 Communication", Required: true, Counted: false},
		{Name: "07 Cooking", Required: true, Counted: false},
		{Name: "08 Emergency Preparedness/Lifesaving", Required: true, Counted: false},
		{Name: "09 Environmental Science/Sustainability", Required: true, Counted: false},
		{Name: "10 First Aid", Required: true, Counted: false},
		{Name: "11 Swimming/Hiking/Cycling", Required: true, Counted: false},
		{Name: "12 Personal Management", Required: true, Counted: false},
		{Name: "13 Personal Fitness", Required: true, Counted: false},
		{Name: "14 Family Life", Required: true, Counted: false},
		{Name: "15 elective"},
		{Name: "16 elective"},
		{Name: "17 elective"},
		{Name: "18 elective"},
		{Name: "19 elective"},
		{Name: "20 elective"},
		{Name: "21 elective"},
	}

	getDateHeader()
	first, err := getDate("Date of First Class", input)
	if err != nil {
		stop()
	}
	star, err := getDate("Date of Star", input)
	if err != nil {
		stop()
	}
	bday, err := getDate("Date of Birth", input)
	if err != nil {
		stop()
	}
	life, err := getDate("Date of Life", input)
	if err != nil {
		stop()
	}

	bday18 := bday.AddDate(18, 0, 0)
	endDate := time.Now()

	if endDate.After(bday18) {
		endDate = bday18
	}

	if star.Before(first.AddDate(0, 4, 0)) {
		fmt.Printf("ERROR: Required 4 months between First Class and Star.\n")
		fmt.Printf("    First Class: %s\n", first.Format(time.UnixDate))
		fmt.Printf("    Star:        %s\n", star.Format(time.UnixDate))
		fmt.Printf("    Earliest:    %s\n", first.AddDate(0, 4, 0).Format(time.UnixDate))
	}

	if life.Before(star.AddDate(0, 6, 0)) {
		fmt.Printf("ERROR: Required 6 months between Star and Life.\n")
		fmt.Printf("    Star:        %s\n", star.Format(time.UnixDate))
		fmt.Printf("    Life:        %s\n", life.Format(time.UnixDate))
		fmt.Printf("    Earliest:    %s\n", star.AddDate(0, 6, 0).Format(time.UnixDate))
	}

	if endDate.Before(life.AddDate(0, 6, 0)) {
		fmt.Printf("ERROR: Required 6 months at Life rank.\n")
		fmt.Printf("    Life:        %s\n", life.Format(time.UnixDate))
		fmt.Printf("    Today/Bday:  %s\n", endDate.Format(time.UnixDate))
		fmt.Printf("    Earliest:    %s\n", life.AddDate(0, 6, 0).Format(time.UnixDate))
	}

	fmt.Println("")
	fmt.Println("Enter merit badge dates. Enter Q to go back to the first MB to correct what you entered.")
	getDateHeader()

	for i, b := range badges {
		badges[i].Earned, err = getDate(b.Name, input)
		if err != nil {
			break
		}
	}

	for err != nil {
		fmt.Println("EDITING BADGES: Re-enter dates: type . to keep the date as-is:\n")

		err = nil
		for i, b := range badges {
			if badges[i].Earned.IsZero() {
				if lastDate.IsZero() {
					fmt.Print("(no date entered yet; don't type .)     ")
				} else {
					fmt.Printf("last date: %s (type . for this)   ", lastDate.Format("01/02/06"))
				}
			} else {
				fmt.Printf("you entered: %s (type . for this) ", badges[i].Earned.Format("01/02/06"))
				lastDate = badges[i].Earned
			}

			badges[i].Earned, err = getDate(b.Name, input)
			lastDate = badges[i].Earned
			if err != nil {
				break
			}
		}
	}

	fmt.Println()
	sort.Sort(ByDateEarned(badges))
	fmt.Printf("For Star Rank:\n")
	need_required := 4
	need_elective := 2
	for i, b := range badges {
		if !b.Counted {
			if need_required > 0 && b.Required && !b.Earned.After(star) {
				badges[i].Counted = true
				need_required--
				fmt.Printf(" Required: %-40s %s\n", b.Name, b.Earned.Format("01/02/06"))
			} else if need_elective > 0 && !b.Required && !b.Earned.After(star) {
				badges[i].Counted = true
				need_elective--
				fmt.Printf(" Elective: %-40s %s\n", b.Name, b.Earned.Format("01/02/06"))
			}
		}
	}
	if need_required > 0 || need_elective > 0 {
		fmt.Printf("ERROR! Still needed: %d required, %d elective\n", need_required, need_elective)
	}

	fmt.Printf("For Life Rank:\n")
	need_required = 3
	need_elective = 2
	for i, b := range badges {
		if !b.Counted {
			if need_required > 0 && b.Required && !b.Earned.After(life) {
				badges[i].Counted = true
				need_required--
				fmt.Printf(" Required: %-40s %s\n", b.Name, b.Earned.Format("01/02/06"))
			} else if need_elective > 0 && !b.Required && !b.Earned.After(life) {
				badges[i].Counted = true
				need_elective--
				fmt.Printf(" Elective: %-40s %s\n", b.Name, b.Earned.Format("01/02/06"))
			}
		}
	}
	if need_required > 0 || need_elective > 0 {
		fmt.Printf("ERROR! Still needed: %d required, %d elective\n", need_required, need_elective)
	}
}

var lastDate time.Time

func getDateHeader() {
	fmt.Printf("%-40s MMDDYY (enter . to repeat the previous date)\n", "")
}

func getDate(prompt string, input *bufio.Scanner) (time.Time, error) {
	for {
		fmt.Printf("%-40s ", prompt)
		input.Scan()
		if input.Text() == "." {
			return lastDate, nil
		}
		if input.Text() == "Q" {
			return lastDate, fmt.Errorf("user abort")
		}
		date, err := time.Parse("010206", input.Text())
		if err == nil {
			lastDate = date
			return date, nil
		}
		fmt.Printf("Error: %v; try again or type Q to abandon entry and start over, or . to repeat the previous date.\n", err)
	}
}

func stop() {
	fmt.Println("Stopping program at user request.")
	os.Exit(1)
}
