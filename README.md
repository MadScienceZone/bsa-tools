# bsatools
A few miscellaneous tools for managing BSA district advancement activities.

None of these are particularly well-polished or "ready for prime time". They are just quick hacks that we've used to automate our processes to manage Eagle boards of review.

# Programs Included
## mb
This is a tiny Go program which is run from the command line. It prompts for the dates on the Eagle rank application in the order they appear on the form. Each date is entered in `MMDDYY` format (e.g., `010224` for January 2, 2024). It uses a crude interface which allows a small degree of error correction but otherwise it's a quick and dirty hack to just get the dates entered and checked.
It will alert if the dates don't meet the rank requirements for tenure between ranks and merit badges earned at each level.

### Prerequisites
You need a Go compiler on your system, or a pre-compiled binary (such as `mb.exe` on a Windows system) compiled for you by someone else. As always, never run a compiled program on your system unless you trust the person who compiled it.

### Usage
Execute the program directly from here using the command `go run mb/main.go`. 
(Alternatively, run the pre-compiled program you compiled yourself or got from a trusted source. Remember this needs to be run in a command shell window since it is not a GUI program.)
Type in each date as prompted by the program. 
You can type `.` to repeat the same date again if multiple consecutive badges were earned on the same date.
You can also type `Q` to abandon the entry of ranks and terminate the program, or to re-do entry of merit badges (in the latter case, it goes back to the first merit badge and prompts for all of their dates again, allowing you to enter a new date or type `.` to keep the previous one you had entered for that one before).

## boardprep
This is a Python script that more or less "grew in the telling" as J.R.R would say.
It takes the information from our spreadsheet of Eagle candidates, and prepares the paperwork for the board itself as well as sending out email to the affected parties at each step of the way.

### Prerequisites
You need the following to be installed on your system:
   * A Python interpreter
   * A copy of `ps2pdf` or some equivalent application which converts PostScript to PDF, or the ability to print PostScript files.
   * An SMTP email-sending program (this script connects to that SMTP TCP port to send outbound email).

You need to set up:
   * Edit the script
   * Prepare your standard board volunteer contacts in the `contacts.csv` file.
   * Make sure the `boardprep` directory is in your `PYTHONPATH` environment variable or you run the script out of the `boardprep` directory itself.

### Usage
   * Download a copy of the boards of review tab of the main spreadsheet to a CSV file (we'll call that `boards.csv` for these instructions)
   * Run `./boardprep` to perform the various operations needed. These are documented fully in the included `boardprep.pdf` document.

# DISCLAIMER
The software in this repository are unreleased works created as an educational exercise and as a matter of personal interest for personal purposes, shared privately with a few friends and associates.
It is not to be considered an industrial-grade product. It is assumed that the user has the necessary understanding and skill to use it appropriately.
The author makes NO representation as to suitability or fitness for any purpose whatsoever, and disclaims any and all liability or warranty to the full extent permitted by applicable law. It is explicitly not designed for use where the safety of persons, animals, property, or anything of real value depends on the correct operation of the software.
