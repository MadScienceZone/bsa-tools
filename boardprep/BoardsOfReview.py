# Yeah, I know this isn't going to win any trophies for clean or well-organized
# coding. It's just a quick hack that kept growing over 10+ years.
#
# Reads exported CSV data for active candidates as the only named input file,
# then outputs:
#    boardsheets.ps -- cover sheets for boards for candidates in "Pending" status
#    envelopes.ps   -- packet covers for candidates in "Council" status
#

import sys, os
import csv
import datetime
import time
import smtplib

class Venue:
    def __init__(self, code):
        code = code.strip()
        self.instructions = None
        self.addr = None
        self.location = None
        if code.startswith('0-'):
            self.desc = 'virtually using Zoom or other teleconferencing application'
            self.instructions = 'Your board chairperson will send instructions about joining the teleconference.'
            self.location = 'via Teleconference'
        elif code.startswith('1-'):
            self.desc = 'at the Helvetia Community Church'
            self.addr = '11295 NW Helvetia Rd<br/>Hillsboro, OR 97124'
            self.location = 'Helvetia Community Church'
        elif code.startswith('2-'):
            self.desc = 'in the Cathy Stanton meeting room at the Beaverton City Library'
            self.addr = '12375 SW 5th St<br/>Beaverton, OR 97005'
            self.instructions = 'The meeting room is immediately to the left as you enter the library\'s lobby, <i>before</i> entering the library itself.'
            self.location = 'Beaverton Library'
        elif code.startswith('3-'):
            self.desc = 'in Meeting Room A at the Beaverton City Library'
            self.addr = '12375 SW 5th St<br/>Beaverton, OR 97005'
            self.location = 'Beaverton Library'
        elif code.startswith('4-'):
            self.desc = 'at Office Evolution'
            self.addr = '9620 NE Tanasbourne Dr, 300<br/>Hillsboro, OR 97124'
            self.location = 'Office Evolution'
        elif code.startswith('5-'):
            self.desc = 'at Our Redeemer Lutheran Church'
            self.addr = '13401 SW Benish St, Tigard, OR 97223'
            self.location = 'Our Redeemer Lutheran Church'
        else:
            raise KeyError(f"Venue code '{code}' not valid")

    def address(self):
        if self.addr is not None:
            return 'The address of the board of review location is:<br/>' + self.addr + '\n'
        return ''

    def extra_instructions(self):
        if self.instructions is None:
            return ''
        return '<blockquote>\n' + self.instructions + '\n</blockquote>\n'


def timestamp():
    return time.strftime('%a, %d %b %Y %H:%M:%S %z')

CFIRST=1
CLAST=2
L_RANK=UNIT=3
CEMAIL=4
CPHONE=5
LFIRST=10
LLAST=11
LPHONE=12
YEARS=19
MONTHS=20
DAYS=21
STAT=28
BORTIME=29
REFS=32
D_RECV=33
D_SENT=34
CHAIR=42
MEMBER=43
L_MEMBER2=NOTES=22
PRJ_COST=23
PRJ_TITLE=24
PRJ_DATE=26
SM_MAIL=16
SM_MAIL2=17
C_ADDR=9
BOR3RD=44
BOR3RDPHONE=45
VENUECODE=48
def yorn(prompt):
    while True:
        answer = input(prompt)
        if answer == 'y':
            return True
        if answer == 'n':
            return False

class BoardsOfReview (object):
    def __init__(self, csvfile, contacts=None):
        self.file_name = csvfile
        self.contacts_file = contacts
        self.envelope_list = []
        self.board_roster = []
        self.pending_roster = []
        self.council_roster = []
        self.all_roster = []
        self.contacts = {}

        if self.file_name:
            self.reformat_file(self.file_name)

        if self.contacts_file:
            self.reformat_file(self.contacts_file)
            self.load_contacts(self.contacts_file)



    def load_contacts(self, file_name):
        TYPE=4
        FN=5
        LN=6
        PHONE=7
        EMAIL=12
        CLEN=EMAIL+1
        with open(file_name) as csv_file:
            csv_reader = csv.reader(csv_file)
            self.contacts = {}
            self.contact_phone = {}
            for contact in csv_reader:
                contact = list(map(str.strip, contact))
                if len(contact) < CLEN:
                    print("length", len(contact), "<", CLEN, contact)
                    continue
                if contact[0].startswith('#'):
                    continue
                if 'MEMBER' in contact[TYPE] or 'CHAIR' in contact[TYPE]:
                    self.contacts[contact[FN] + ' ' + contact[LN]] = contact[EMAIL]
                    self.contact_phone[contact[FN] + ' ' + contact[LN]] = contact[PHONE]
                    #print "Added", contact[TYPE], contact[FN] + ' ' + contact[LN], "as", contact[EMAIL]
                else:
                    print("Rejected", contact)

    def reformat_file(self, file_name=None):
        csvfile = file_name or self.file_name
        
        os.rename(csvfile, csvfile + '.bak')
        raw_file = open(csvfile + '.bak')
        cooked_file = open(csvfile, 'w')

        for line in raw_file:
            cooked_file.write(line.replace('\r', '\n'))

        cooked_file.close()
        raw_file.close()

    def write_envelopes(self, output_file='envelopes.ps'):
        envelopes = open(output_file, 'w')
        envelopes.write(r'''%!PS
% vi:set ai sm nu ts=4 sw=4 expandtab:
%
% To create new envelopes, go to the bottom of this file...
%                   ||
%                   ||
%                   ||
%                  _||_
%                 \    /
%                  \  /
%                   \/
%
/ssfont {
    /Helvetica findfont 10 scalefont setfont 12 setlead
} def
/italfont {
    /Times-Italic findfont 12 scalefont setfont 14 setlead
} def
/normalfont {
    /Times-Roman findfont 12 scalefont setfont 14 setlead
} def
/boldfont {
    /Times-Bold findfont 12 scalefont setfont 14 setlead
} def
/bigfont {
    /Times-Bold findfont 18 scalefont setfont 20 setlead
} def
/setlead {
    /leading exch def
} def
/startpage {
    normalfont
    /Xmargin 36 def
    /Ymargin 36 def
    /Xwidth  8.5 72 mul def
    /Yheight 11  72 mul def
    /X Xmargin def
    /Y Yheight Ymargin sub leading sub def
} def
/down {
    Y exch sub /Y exch def
} def
/newline {
    leading down
    /X Xmargin def
} def
/center {
    dup stringwidth pop 2 div Xwidth 2 div exch sub /X exch def
    X Y moveto show newline
} def
/underlinecenter {
    % s 
    dup stringwidth pop dup 2 div Xwidth 2 div exch sub /X exch def
    X Y 2 sub moveto 0 rlineto stroke
    center
} def
/s {
    X Y moveto show currentpoint /Y exch def /X exch def
} def
/bigskip {
    leading 3 mul down
} def
/medskip {
    leading 2 mul down
} def
/smallskip {
    leading down
} def
/rulefill {
    %0 -2 rmoveto
    %rulelimit Y lineto 
    %0 +2 rmoveto
    %stroke
    X Y 2 sub moveto rulelimit Y 2 sub lineto stroke
    /X rulelimit def
} def
%
% string width showfield -
%
/showfieldcore {
    /sfw exch def
    dup
    X Y 2 sub moveto sfw 0 rlineto stroke
    X sfw 2 div add exch
    stringwidth pop 2 div sub Y moveto show
    /X X sfw add def
} def
/showbigfield {
    bigfont
    showfieldcore
    normalfont
} def
/showfield {
    boldfont
    showfieldcore
    normalfont
} def
/rule {
    0 -2 rmoveto
    56 0 rlineto 
    0 +2 rmoveto stroke
    /X X 56 add def
} def
/showdate {
    boldfont
    dup stringwidth pop neg X add /X exch def s
    normalfont
} def
    
%
% forwarded reviewed received envelope -
%
/envelope {
    startpage
    /rulelimit Xwidth .66 mul def

    (CASCADE PACIFIC COUNCIL) center
    (BOY SCOUTS OF AMERICA) center
    (2145 SW NAITO PARKWAY) center
    (PORTLAND, OR 97201) center
    (503-226-3423) center

    bigskip

    boldfont(EAGLE SCOUT AWARD APPLICATION) underlinecenter
    normalfont

    bigskip

    italfont (\(District Advancement/Eagle Board Chairman\)) s newline
    normalfont (DATE APPLICATION RECEIVED:) s rulefill showdate newline
    smallskip
    (DATE APPLICATION REVIEWED:) s rulefill showdate newline
    smallskip
    (DATE FORWARDED TO COUNCIL OFFICE:) s rulefill showdate newline

    bigskip

    italfont(\(Council Eagle Application Processor\)) s newline
    normalfont(RANKS & MERIT BADGES VERIFIED:) s rulefill newline

    smallskip

    (REGISTRATION VERIFIED:) s rulefill newline

    bigskip

    italfont(\(Return to District\)) s newline
    normalfont(BOARD OF REVIEW DATE:) s rulefill newline
    smallskip
    (RECOMMENDED:) s rule (YES) s rule (NO) s rule (DEFERRED) s newline

    bigskip

    italfont(\(Council Eagle Application Processor\)) s newline
    normalfont(DATE APPLICATION SENT TO NATIONAL:) s rulefill newline
    smallskip
    (DATE APPROVAL RECEIVED FROM NATIONAL:) s rulefill newline
    smallskip
    (DATE UNIT LEADER NOTIFIED:) s rulefill newline

    bigskip

    (NOTES FOR REVIEW BOARD:) s newline
   newline 
    bigfont s normalfont

    gsave
    -90 rotate
    -11 72 mul 0 translate
    /X 36 def
    /Y Xwidth 36 sub def
    ssfont(UNIT TYPE/NUMBER:) s 80 showbigfield
    ssfont(DISTRICT:) s (Pacific Trail) 100 showfield
    ssfont(EAGLE CANDIDATE'S NAME:) s 200 showbigfield newline
    smallskip

    ssfont(UNIT LEADER'S) s newline
    ssfont(NAME:) s 280 showbigfield /ColX X 85 add def
    ColX Y leading add moveto ssfont(UNIT LEADER'S) show
    /X ColX def ssfont(DAYTIME PHONE NUMBER:) s 140 showbigfield 
    newline
    /X X 100 add def ssfont(\(For Notification\)) s 
    grestore
    showpage
} def
/note {
    pop
} def
''')
        for date_sent, date_recv, notes, candidate in self.envelope_list:
            envelopes.write('({0})({1} {2})({3} {4})({5})({6})({7})({8})({8})envelope\n'.format(
                self.ps(candidate[LPHONE]), 
                self.ps(candidate[LFIRST]), 
                self.ps(candidate[LLAST]), 
                self.ps(candidate[CFIRST]),
                self.ps(candidate[CLAST]), 
                self.ps(candidate[UNIT]), 
                self.ps('; '.join(notes)),
                date_sent.strftime('%d-%b-%Y') if date_sent is not None else '',
                date_recv.strftime('%d-%b-%Y') if date_recv is not None else ''))
        envelopes.close()

    def write_cover_sheets(self, output_file='boardsheets.ps'):
        bsheet = open(output_file, 'w')
        bsheet.write(r'''%!PS
% vim:set syntax=postscr ai sm nu ts=4 sw=4 expandtab:
%
% This is my standard set of form-building elements for PostScript forms
% Copyright (c) 2003, 2004, 2009, 2010, 2013 Steve Willoughby, Aloha, Oregon, 
% USA.  All Rights Reserved.
%
% Version 1.4
% (This file used to be part of one monolithic file
% implemented in the "ps.tcl" module of my "gma" software; broken
% into separate files as of version 1.3 to make the form-building
% code easier to use in other places).
%------------------------------------------------------------------------
% Shortcuts for general PostScript operations
% Note that applications may override some of these.
%
% np cp mv ln rln rmv ... ok     basic line drawing
% {wide|med|thin}LineWidth       standard line sizes
% eject                          ship out page border and eject page

/np  { newpath   } def 
/mv  { moveto    } def
/ln  { lineto    } def
/rln { rlineto   } def
/rmv { rmoveto   } def
/cp  { closepath } def
/ok  { stroke    } def

/SetLine_wide  { 3.0 setlinewidth } def
/SetLine_med   { 1.0 setlinewidth } def
/SetLine_thin  { 0.5 setlinewidth } def
%
% w R -
%   move X right w points
%
/R {
    X add /X exch def
} def
%
% h D -
%   move Y down h points
%
/D {
    Y exch sub /Y exch def
} def
%
% Reserve some space for what's coming next, and
% start a new page if we can't do that.
%
/RequiredVerticalSpace {
    Y exch sub PageBottomMargin lt {
        eject
    } if
} def
%
% x y w h BoxFrame -
%   Draw empty, unfilled box with (x,y) at top left
%
/BoxPath {
    4 2 roll        % w h x y
    np mv       % w h     @(x,y)
    dup neg     % w h -h
    0 exch      % w h 0 -h
    rln     % w h     down h
    exch 0          % h w 0
    rln     % h       right w
    0 exch      % 0 h
    rln     %         up h
    cp      % 
} def
/BoxFrame { BoxPath stroke} def
/BoxFill  { BoxPath fill  } def

/appEject {
    % Applications should override this
    % (called at end of each page)
} def
/appStart {
    % Applications should override this
    % (called at start of each page)
    % WARNING: an extra appStart may be 
    % called at end of job, so
    % don't emit anything here!
} def
/eject {
    appEject
    showpage
    appStart
} def
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% THEME MANAGEMENT
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% The form theme should redefine the following commands which
% set colors of elements:
/BlankHue       { 1.00 setgray } def    % blank cell background
/GreyHue    { 0.80 setgray } def    % greyed out elements
/AltGreyHue { 0.60 setgray } def    % for bonus
/DarkHue    { 0.40 setgray } def    % disabled (heavy grey)
/FormHue    { 0.20 setgray } def    % borders, etc.
/Highlight  { 1 1 0 setrgbcolor } def
/SetColor_form  { /FormColor { FormHue   } def FormColor } def
/SetColor_data  { /FormColor { 0 setgray } def FormColor } def

% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% FONT MANAGEMENT
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
/FontSize_footnote 6 def
/FontSize_data_ttl 6 def
/FontSize_body 8 def
/FontLead_body 9 def

/FontFace_footnote /Helvetica def
/FontFace_data_ttl /Helvetica def
/FontFace_body /Palatino-Roman def

/SelectFootnoteFont  { FontFace_footnote findfont FontSize_footnote scalefont setfont } def
/SelectDataTitleFont { FontFace_data_ttl findfont FontSize_data_ttl scalefont setfont } def
/SelectBodyFont      { FontFace_body     findfont FontSize_body     scalefont setfont } def

%
% Formatted output with line wrapping
%
% [ chunk0 chunk1 ... chunkN ] width showproc nlproc WrapAndFormat -
%
% prints each chunk in order, calling showproc for each in turn.  Calls
% nlproc as needed to start a new line to avoid going over width.
%
% Each chunk is an array of 3 elements: [ eproc [s0 s1 ... sN] sproc ]
% sproc is called to setup the chunk (font selection, etc), then each
% string s0...sN are displayed as will fit on the line.  Unlike the
% BreakIntoLines procedure, we don't do the breaking ourselves here,
% the generating application does that by assembling the arrays used.
% eproc is called to clean up after the chunk is complete.  
%
% It is legal for sproc or eproc to call showproc and/or nlproc themselves.

/WrapAndFormat {
    /WaF__nlproc exch def
    /WaF__showproc exch def
    /WaF__width exch def
    /WaF__curwidth 0 def
    {
        aload pop exec  % run chunk's sproc
        {
            dup stringwidth pop dup WaF__curwidth add % strN len len+cw 
            WaF__width ge {               % strN len 
                WaF__nlproc
                /WaF__curwidth exch def
            } {
                WaF__curwidth add /WaF__curwidth exch def
            } ifelse
            WaF__showproc
        } forall
        exec    % run chunk's eproc
    } forall
} def

% Generic nlproc and showproc (and friends) you can use with WrapAndFormat
% Call WaF_init first to set starting location for rendering
% x y lead WaF_init -
/WaF_init {
    /WaF__Y exch def
    /WaF__X exch def
    /WaF__lead exch def
    WaF__X WaF__Y moveto
} def
/WaF_nl {
    /WaF__Y WaF__Y WaF__lead sub def
    WaF__X WaF__Y moveto
} def
/WaF_show {
    show
} def

%
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Adapted from the Blue Book -- line breaks in running text
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
/wordbreak ( ) def
% string width proc BreakIntoLines -
% Calls:    line proc    for every line broken out
/BreakIntoLines {
    /proc exch def
    /linewidth exch def
    /textstring exch def

    /breakwidth wordbreak stringwidth pop def
    /curwidth 0 def
    /lastwordbreak 0 def
    /startchar 0 def
    /restoftext textstring def

    {
        restoftext wordbreak search {
            /nextword exch def pop
            /restoftext exch def
            /wordwidth nextword stringwidth pop def
            curwidth wordwidth add linewidth gt {
                textstring startchar
                lastwordbreak startchar sub
                getinterval proc
                /startchar lastwordbreak def
                /curwidth wordwidth breakwidth add def
            } {
                /curwidth curwidth wordwidth add
                breakwidth add def
            } ifelse
            /lastwordbreak lastwordbreak nextword length add
            1 add def
        } {
            pop exit
        } ifelse
    } loop
    /lastchar textstring length def
    textstring startchar lastchar startchar sub
    getinterval proc
} def
%
% text x y RenderText w h
% Draw text in current font at (x,y) baseline
% Return text dimensions
%
/RenderText {
    mv dup stringwidth  % text w h
    3 -1 roll       % w h text
    show            % w h
} def
%
% text font size w x y FitText -
% Draw text in specified font at (x,y) baseline, scaled down
% if necessary to fit in maximum width w
%
/FitTextCommon {
    /FormFT_y   exch def 
    /FormFT_x   exch def
    /FormFT_w   exch def
    /FormFT_sz  exch def
    /FormFT_fnt exch def
    /FormFT_txt exch def

    {
        FormFT_fnt findfont FormFT_sz scalefont setfont
        FormFT_txt stringwidth pop FormFT_w le {exit} if
        /FormFT_sz FormFT_sz 0.1 sub def
        FormFT_sz 3 le {exit} if
    } loop
} def

/FitText {
    FitTextCommon
    FormFT_x FormFT_y mv FormFT_txt show
} def
/FitTextCtr {
    FitTextCommon
    FormFT_x FormFT_w 2 div add
    FormFT_txt stringwidth pop 2 div sub
    FormFT_y mv
    FormFT_txt show
} def
/FitTextR {
    FitTextCommon
    FormFT_x FormFT_w add
    FormFT_txt stringwidth pop sub
    FormFT_y mv
    FormFT_txt show
} def

%
% Row of check-boxes, shaded in blocks 
% qty used checked half-checked boxesperrow box-x y w h gap-x y shadeinterval CheckBoxMatrix -
%
/CheckBoxMatrix {
    /CBM_shadeinter exch def
    /CBM_gap_y exch def
    /CBM_gap_x exch def
    /CBM_bx_h exch def
    /CBM_bx_w exch def
    dup /CBM_bx_y exch def /CBM__y exch def
    dup /CBM_bx_x exch def /CBM__x exch def
    /CBM_bpr  exch def      % q u c h
    3 1 roll            % q h u c
    exch                % q h c u
    /CBM_qty_used exch def      % q h c
    dup             % q h c c
    /CBM_qty_chkd exch def      % q h c
    add             % q h+c
    /CBM_qty_half exch def      % q
    /CBM_qty_total exch def     % this adjusts _half and _chkd to be the position for each:
    %  X X X X X / / [] [] [] _ _ _
    % |<-chkd-->|
    % |<---half---->|
    % |<-------used--------->|
    % |<------------qty----------->|
    % 
    % (__x,__y) give the position for the next box to draw
    /CBM_shademode false def
    1 1 CBM_qty_total {
        CBM__boxpath CBM_shademode { AltGreyHue } { BlankHue } ifelse fill
        SetColor_form
        CBM__boxpath                % i         (path set)
        dup CBM_qty_used le {           % i i<=u?
            stroke              % i         used block: solid box
            SetColor_data
            dup CBM_qty_half le {       % i i<=h?
                np          % i     within chkd/half zone, make 1/2 check
                    CBM__x CBM__y CBM_bx_h sub mv
                    CBM_bx_w CBM_bx_h rln
                stroke
            } if
            dup CBM_qty_chkd le {       % i i<=c?
                np          %       within chks zone, make other 1/2 check
                    CBM__x CBM__y mv
                    CBM_bx_w CBM_bx_h neg rln
                stroke
            } if
        } {
                            % outsize used zone; draw dashed outline of box
            SetLine_thin
            [1 1] .5 setdash
            stroke
        } ifelse
        % i still on stack here, regardless
        dup CBM_shadeinter mod 0 eq {       % i i%int?
            /CBM_shademode CBM_shademode not def    % flip shade mode
        } if
        /CBM__x CBM__x CBM_gap_x CBM_bx_w add add def   % advance to next horizontal position
        CBM_bpr mod 0 eq {          % box number at end of row?
            /CBM__x CBM_bx_x def
            /CBM__y CBM__y CBM_gap_y CBM_bx_h add sub def
        } if
    } for
    [] 0 setdash
} def

/CBM__boxpath {
    np 
        CBM__x CBM__y mv
        CBM_bx_w 0 rln
        0 CBM_bx_h neg rln
        CBM_bx_w neg 0 rln
    cp
} def


%
% Row of data blocks:
%
% textfont textsize boxheight BeginDataBlock 
% [set-hue-command] data title w DataBlock
% [set-hue-command] data w TitleBlockCtr      centered reversed colors, no title
%   .
%   .
%   .
% EndDataBlock
%
/BeginDataBlock {
    dup RequiredVerticalSpace
    /FormDBboxht exch def
    ChangeDataBlockFont
    /FormDB__X__ X def
} def
/ChangeDataBlockFont {
    /FormDBtsize exch def
    /FormDBtfont exch def
} def
/DataBlockCommon {
    /FormDBbox_w exch def
    SetLine_med
    X Y FormDBbox_w FormDBboxht BoxFill
    FormHue  X Y FormDBbox_w FormDBboxht BoxFrame
    SelectDataTitleFont
    X 1 add Y FontSize_data_ttl sub RenderText pop pop
    FormColor 
    FormDBtfont FormDBtsize FormDBbox_w 6 sub X 3 add Y FormDBboxht .80 mul sub 
    FormDBbox_w R
} def
/DataBlock {
    DataBlockCommon
    FitText
} def
/DataBlockR {
    DataBlockCommon
    FitTextR
} def
/DataBlockCtr {
    DataBlockCommon
    FitTextCtr
} def
/DataBlockCR {  % take 2 data args; first on stack is centered in white, then second right-justified over it
    DataBlockCommon
    1 setgray
    /DBcr__y exch def
    /DBcr__x exch def
    /DBcr__w exch def
    /DBcr__s exch def
    /DBcr__f exch def
    DBcr__f DBcr__s DBcr__w DBcr__x DBcr__y FitTextCtr
    FormColor
    DBcr__f DBcr__s DBcr__w DBcr__x DBcr__y FitTextR
} def

/DataBlockLR {  % take 2 data args; first on stack is centered in white, then second right-justified over it
    DataBlockCommon
    1 setgray
    /DBcr__y exch def
    /DBcr__x exch def
    /DBcr__w exch def
    /DBcr__s exch def
    /DBcr__f exch def
    DBcr__f DBcr__s DBcr__w DBcr__x DBcr__y FitText
    FormColor
    DBcr__f DBcr__s DBcr__w DBcr__x DBcr__y FitTextR
} def

/TitleBlockCtr {
    /FormDBbox_w exch def
    FormHue X Y FormDBbox_w FormDBboxht BoxFill
    SetLine_med X Y FormDBbox_w FormDBboxht BoxFrame
    1 setgray FormDBtfont FormDBtsize FormDBbox_w X Y FormDBboxht .80 mul sub FitTextCtr
    FormDBbox_w R
} def

/EndDataBlock {
    /X FormDB__X__ def
    FormDBboxht D
} def
%
% (End of form-preamble.ps)
%

%
% Board sheet macros
%

% roster -> nil
% date BeginRoster -> nil
% EndRoster -> nil
%
/BeginRoster {
    SetColor_form
    50 D
    /X 40 def
    /Palatino-Roman 12 20 BeginDataBlock
        BlankHue (DATE) 100 DataBlock
    EndDataBlock
    20 D
} def
/EndRoster {
    eject
} def
/roster {
    /Palatino-Roman 12 20 BeginDataBlock
        GreyHue  (TIME)         35 DataBlock
        BlankHue (APPLICANT)   145 DataBlock
        BlankHue (PHONE)        75 DataBlock
        BlankHue (UNIT)         40 DataBlock
        GreyHue  (UNIT LEADER)  85 DataBlock
        GreyHue  (PHONE)        75 DataBlock
        BlankHue (BOARD CHAIR) 100 DataBlock
        BlankHue (BOARD MEMBER)100 DataBlock
        dup () eq {
         BlankHue 
        } { 
         Highlight 
        } 
        ifelse      (LTR)       15 DataBlock
        BlankHue () (ROOM)      45 DataBlock
        %1 1 0 0 1 X 5 add Y 5 sub 10 10 0 0 1 CheckBoxMatrix
    EndDataBlock
    12 D
} def

% chairlast chairfirst date time r g b board -> nil
/board {
    /boxw 300 def
    /boxh 214 def
    /boxx 30 def
    /boxy 612 30 sub def

    /bB exch def
    /bG exch def
    /bR exch def
    %bR bG bB setrgbcolor boxx boxy boxw boxh BoxFill
    /boxstep 360 def
    /boxSliceWidth boxw 2 mul boxstep div def
    /bxx boxx def
    /bsR 1 bR sub boxstep div def
    /bsG 1 bG sub boxstep div def
    /bsB 1 bB sub boxstep div def
    1 -1 boxstep div 0 {
        /shadeFactor exch def
        bR bG bB setrgbcolor
        /bR bR bsR add def
        /bG bG bsG add def
        /bB bB bsB add def
        bxx boxy boxSliceWidth boxh BoxFill
        /bxx bxx boxSliceWidth add def
    } for
    /bortime exch def
    1 setgray bortime /Helvetica 100 boxw boxx boxh 1.5 div boxy exch sub FitTextCtr
    /Helvetica 12 boxw boxx 2 add boxy boxh sub 2 add FitText
    0 setgray
    /X 400 def
    /Y boxy def
    SetColor_form
    50 D
    /Helvetica-Bold 24 360 X Y FitText
    /Helvetica-Bold 200 360 X Y FitTextCommon
    FormFT_sz D X Y mv FormFT_txt show
%    /Helvetica 24 32 BeginDataBlock
%        BlankHue (BOARD CHAIR) 360 DataBlock
%    EndDataBlock
%    /Helvetica 200 150 BeginDataBlock
%        BlankHue () 360 DataBlock
%    EndDataBlock
%    /Helvetica 24 32 BeginDataBlock
%        BlankHue (2ND BOARD MEMBER) 360 DataBlock
%    EndDataBlock
    50 D
    /Helvetica 24 360 X Y FitText

    /X 100 def
    64 D
    /Helvetica 24 32 BeginDataBlock
        BlankHue (APPLICANT) 300 DataBlock
        BlankHue (PHONE) 150 DataBlock
        BlankHue (UNIT) 100 DataBlock
    EndDataBlock
    /Helvetica 24 32 BeginDataBlock
        BlankHue (UNIT LEADER) 300 DataBlock
        BlankHue (PHONE) 150 DataBlock
        BlankHue () () 100 DataBlock
    EndDataBlock
    eject
} def
% board-date project -> nil
%  --- This is a terrible kludge because I don't have time to do it right.
%  --- REFACTOR THIS LATER!
/project {
    48 D
    /Tab PageWidth 2 div def
    (EAGLE PROJECT INFORMATION) /Helvetica-Bold 14 PageWidth 0 Y FitTextCtr 28 D
    (Pacific Trail District)     /Helvetica-Bold 14 PageWidth 0 Y FitTextCtr 48 D
    (TODAY'S DATE:)             /Helvetica      12 PageWidth 48 Y FitText
                                /Palatino-Bold  14 PageWidth 150 Y FitText
    (UNIT TYPE & NUMBER:)       /Helvetica      12 PageWidth Tab Y FitText
                                /Palatino-Bold  14 PageWidth Tab 150 add Y FitText 20 D
    (SCOUT'S NAME:)             /Helvetica      12 PageWidth 48 Y FitText
                                /Palatino-Bold  14 PageWidth 150 Y FitText 20 D
    (PROJECT COMPLETION DATE:)  /Helvetica      12 PageWidth 48 Y FitText
                                /Palatino-Bold  14 PageWidth 250 Y FitText 20 D
    (PROJECT FINAL PRICE/FUNDRAISING AMOUNT:)  
                                /Helvetica      12 PageWidth 48 Y FitText
                                /Palatino-Bold  14 PageWidth 350 Y FitText 48 D
    (Please give a brief explanation of what the Eagle Candidate's Service Project was \(2-5 sentences)
                                /Helvetica      12 PageWidth 48 2 mul sub 48 Y FitText 14 D
    (only\):)                   /Helvetica      12 PageWidth 48 Y FitText 48 D
    /Palatino-Bold 14 PageWidth 0 Y FitTextCtr 36 D
    8 {
        newpath 48 Y moveto PageWidth 48 2 mul sub 0 rlineto stroke
        36 D
    } repeat
    eject
} def
/signatures {
    72 D
    /Tab PageWidth 2 div def
    /TabL0 72 def
    /TabL1 PageWidth 2 div 50 sub def
    /TabLw TabL1 TabL0 sub def
    /TabR0 PageWidth 2 div 50 add def
    /TabR1 PageWidth 72 sub def
    /TabRw TabR1 TabR0 sub def
    /TabC0 TabL1 5 add def
    /TabC1 TabR0 5 sub def
    /TabCw TabC1 TabC0 sub def

    /TabA0 72 def
    /TabA1 TabL1 def
    /TabAw TabA1 TabA0 sub def
    /TabB0 Tab def
    /TabB1 TabR1 def
    /TabBw TabB1 TabB0 sub def

    (Cascade Pacific Council) /Palatino 10 PageWidth 48 Y FitText 
    (Boy Scouts of America)   /Palatino 10 PageWidth 72 sub 48 Y FitTextR 
    72 D
    /Palatino-Bold 14 TabLw TabL0 Y FitTextCtr 
    /Palatino-Bold 14 TabCw TabC0 Y FitTextCtr 
    (Pacific Trail) /Palatino-Bold 14 TabRw TabR0 Y FitTextCtr 
    3 D
    newpath TabL0 Y moveto TabL1 Y lineto stroke
    newpath TabC0 Y moveto TabC1 Y lineto stroke
    newpath TabR0 Y moveto TabR1 Y lineto stroke
    10 D
    (Applicant) /Helvetica 10 TabLw TabL0 Y FitTextCtr
    (Unit)      /Helvetica 10 TabCw TabC0 Y FitTextCtr
    (District)  /Helvetica 10 TabRw TabR0 Y FitTextCtr
    20 D
    /Palatino-Bold 14 TabLw TabL0 Y FitTextCtr 
    /Palatino-Bold 14 TabCw TabC0 Y FitTextCtr 
    /Palatino-Bold 14 TabRw TabR0 Y FitTextCtr 
    3 D
    newpath TabL0 Y moveto TabL1 Y lineto stroke
    newpath TabC0 Y moveto TabC1 Y lineto stroke
    newpath TabR0 Y moveto TabR1 Y lineto stroke
    10 D
    (Address) /Helvetica 10 TabLw TabL0 Y FitTextCtr
    (Phone)   /Helvetica 10 TabCw TabC0 Y FitTextCtr
    (Date)    /Helvetica 10 TabRw TabR0 Y FitTextCtr
    36 D
    (We, the undersigned, representing the Advancement Committee of the Cascade Pacific Council)
    /Helvetica 12 PageWidth 48 Y FitText 14 D
    (of the Boy Scouts of America, have reviewed the above named applicant of the requirements for)
    /Helvetica 12 PageWidth 48 Y FitText 14 D
    (the Eagle Scout rank. We find him worthy of advancement.)
    /Helvetica 12 PageWidth 48 Y FitText 14 D
    36 D
    (Date of Review) /Helvetica 10 TabAw TabA0 Y FitText
    newpath TabA0 70 add Y moveto TabA1 Y lineto stroke
    -3 D /Palatino-Bold 14 TabAw 70 sub TabA0 70 add Y FitTextCtr 3 D
    (Place) /Helvetica 10 TabBw TabB0 Y FitText
    newpath TabB0 30 add Y moveto TabB1 Y lineto stroke
    -3 D /Palatino-Bold 14 TabBw 30 sub TabB0 30 add Y FitTextCtr 3 D
    72 D

    (Signature)        /Helvetica 10 TabLw TabL0 Y FitTextCtr
    (Printed Name)     /Helvetica 10 TabCw TabC0 Y FitTextCtr
    (Telephone Number) /Helvetica 10 TabRw TabR0 Y FitTextCtr

    /o 10 def
    2 D
    /YY Y def
    6 {
        newpath
            72 Y moveto
            PageWidth 72 sub Y lineto
        stroke
        /Palatino-Bold 14 TabR0 TabL1 sub o 2 mul add TabL1 o sub Y 20 sub FitTextCtr
        /Palatino-Bold 14 TabRw o sub TabR0 o add Y 20 sub FitTextCtr
        36 D
    } repeat
    -36 D
    newpath TabL1 o sub YY moveto TabL1 o sub Y lineto stroke
    newpath TabR0 o add YY moveto TabR0 o add Y lineto stroke

%
%
%
%
%    (EAGLE PROJECT INFORMATION) /Helvetica-Bold 14 PageWidth 0 Y FitTextCtr 28 D
%    (District 2)     /Helvetica-Bold 14 PageWidth 0 Y FitTextCtr 48 D
%    (TODAY'S DATE:)             /Helvetica      12 PageWidth 48 Y FitText
%                                /Palatino-Bold  14 PageWidth 150 Y FitText
%    (UNIT TYPE & NUMBER:)       /Helvetica      12 PageWidth Tab Y FitText
%                                /Palatino-Bold  14 PageWidth Tab 150 add Y FitText 20 D
%    (SCOUT'S NAME:)             /Helvetica      12 PageWidth 48 Y FitText
%                                /Palatino-Bold  14 PageWidth 150 Y FitText 20 D
%    (PROJECT COMPLETION DATE:)  /Helvetica      12 PageWidth 48 Y FitText
%                                /Palatino-Bold  14 PageWidth 250 Y FitText 20 D
%    (PROJECT FINAL PRICE/FUNDRAISING AMOUNT:)  
%                                /Helvetica      12 PageWidth 48 Y FitText
%                                /Palatino-Bold  14 PageWidth 350 Y FitText 48 D
%    (Please give a brief explanation of what the Eagle Candidate's Service Project was \(2-5 sentences)
%                                /Helvetica      12 PageWidth 48 2 mul sub 48 Y FitText 14 D
%    (only\):)                   /Helvetica      12 PageWidth 48 Y FitText 48 D
%    /Palatino-Bold 14 PageWidth 0 Y FitTextCtr 36 D
%    8 {
%        newpath 48 Y moveto PageWidth 48 2 mul sub 0 rlineto stroke
%        36 D
%    } repeat
    eject
} def
/appStart {
    gsave
    pageSetup
} def
/landscapePageSetup {
    90 rotate
    0 -612 translate
    /PageLength 612 def
    /PageWidth 792 def
    /PageTopMargin 612 def
    /PageBottomMargin 0 def
    /PageRightMargin 792 def
    /PageLeftMargin 0 def
    /PageTextWidth 792 def
    /InterBlockGap 5 def
    /X 30 def
    /Y PageTopMargin def
} def
/portraitPageSetup {
    /PageWidth 612 def
    /PageLength 792 def
    /PageTopMargin 792 def
    /PageBottomMargin 0 def
    /PageRightMargin 612 def
    /PageLeftMargin 0 def
    /PageTextWidth 612 def
    /InterBlockGap 5 def
    /X 30 def
    /Y PageTopMargin def
} def
/setPortrait  { /pageSetup { portraitPageSetup } def } def
/setLandscape { /pageSetup { landscapePageSetup} def } def
/appEject {
    grestore
} def
setLandscape
appStart
''')
        for board_time, candidate in self.board_roster:
            bsheet.write('({12})({10} {11})({8})({9})({6} {7})({5})({3})({4})({0})({1}){2} board\n'.format(
                board_time.strftime('%d-%b-%Y'),
                board_time.strftime('%I:%M'),
                '.6 0 0' if board_time.hour == 19 else '0 .6 0' if board_time.minute == 15 else '0 0 1',
                self.ps(candidate[CHAIR].upper().split()[1]),
                self.ps(candidate[CHAIR].upper().split()[0]),
                self.ps(candidate[MEMBER].upper()),
                self.ps(candidate[CFIRST]),
                self.ps(candidate[CLAST]),
                self.ps(candidate[UNIT]),
                self.ps(candidate[CPHONE]),
                self.ps(candidate[LFIRST]),
                self.ps(candidate[LLAST]),
                self.ps(candidate[LPHONE]),
            ))

        bsheet.write('({0})BeginRoster\n'.format(
            self.board_roster[0][0].strftime('%d-%b-%Y')
        ))
        for entry in sorted(self.board_roster):
            try: 
                irefs = int(entry[1][REFS])
            except ValueError:
                if entry[1][REFS]:
                    raise ValueError("Invalid number of references: {0}".format(entry[1][REFS]))
                irefs = 0

            bsheet.write('({10})({0})({1})({2})({3} {4})({5})({6})({7} {8})({9})roster\n'.format(
                entry[1][MEMBER],
                entry[1][CHAIR],
                entry[1][LPHONE],
                entry[1][LFIRST],
                entry[1][LLAST],
                entry[1][UNIT],
                entry[1][CPHONE],
                entry[1][CFIRST],
                entry[1][CLAST],
                entry[0].strftime('%I:%M'),
                3 - irefs if irefs < 3 else '',
            ))
        bsheet.write('setPortrait\nEndRoster\n')
        for board_time, candidate in sorted(self.board_roster):
            try:
                venue = Venue(candidate[VENUECODE])
            except KeyError as e:
                print(f"ERROR: {e}")
                continue

                #()()()()()()({14})({13})({11})({9})({10})({8})(Via Teleconference)({0})({0})({7})({12})({1})({2} {3}) signatures
            bsheet.write('''
                ({6})({5})({4})({2} {3})({1})({0}) project
                ()()()()()()({14})({13})({11})({9})({10})({8})({15})({0})({0})({7})({12})({1})({2} {3}) signatures
            '''.format(
                board_time.strftime('%d-%b-%Y'),
                self.ps(candidate[UNIT]),
                self.ps(candidate[CFIRST]),
                self.ps(candidate[CLAST]),
                self.ps(candidate[PRJ_DATE]),
                self.ps(candidate[PRJ_COST]),
                self.ps(candidate[PRJ_TITLE]),
                self.ps(candidate[CPHONE]),
                self.ps(candidate[CHAIR]),
                self.ps(candidate[MEMBER]),
                self.ps(self.contact_phone.get(candidate[CHAIR], '')),
                self.ps(self.contact_phone.get(candidate[MEMBER], '')),
                self.ps(candidate[C_ADDR]),
                self.ps(candidate[BOR3RD]),
                self.ps(candidate[BOR3RDPHONE]),
                self.ps(venue.location),
            ))
        bsheet.close()

    def ps(self, str):
        return str.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')

    def date_or_blank(self, datestr):
        if not datestr:
            return None

        try:
            return datetime.datetime.strptime(datestr, '%m/%d/%Y')
        except ValueError:
            return datetime.datetime.strptime(datestr, '%m/%d/%y')

    def date_time_str(self, dt):
        return dt.strftime('%d-%b-%Y, %l:%M PM')

    def check(self, csv_file=None):
        "process input, report on column alignment"

        with open(csv_file or self.file_name) as csv_file:
            csv_reader = csv.reader(csv_file)

            for row, candidate in enumerate(csv_reader):
                if row == 4:
                    print('candidate first name {0}'.format(candidate[CFIRST].replace('\n',' // ')))
                    print('candidate last name  {0}'.format(candidate[CLAST].replace('\n',' // ')))
                    print('candidate email      {0}'.format(candidate[CEMAIL].replace('\n',' // ')))
                    print('candidate phone      {0}'.format(candidate[CPHONE].replace('\n',' // ')))
                    print('candidate address    {0}'.format(candidate[C_ADDR].replace('\n',' // ')))
                    print('candidate age years  {0}'.format(candidate[YEARS].replace('\n',' // ')))
                    print('candidate age months {0}'.format(candidate[MONTHS].replace('\n',' // ')))
                    print('candidate age days   {0}'.format(candidate[DAYS].replace('\n',' // ')))
                    print('leader    first name {0}'.format(candidate[LFIRST].replace('\n',' // ')))
                    print('leader    last name  {0}'.format(candidate[LLAST].replace('\n',' // ')))
                    print('leader    email #1   {0}'.format(candidate[SM_MAIL].replace('\n',' // ')))
                    print('leader    email #2   {0}'.format(candidate[SM_MAIL2].replace('\n',' // ')))
                    print('leader    phone      {0}'.format(candidate[LPHONE].replace('\n',' // ')))
                    print('application status   {0}'.format(candidate[STAT].replace('\n',' // ')))
                    print('board time           {0}'.format(candidate[BORTIME].replace('\n',' // ')))
                    print('letters received     {0}'.format(candidate[REFS].replace('\n',' // ')))
                    print('date received        {0}'.format(candidate[D_RECV].replace('\n',' // ')))
                    print('date sent            {0}'.format(candidate[D_SENT].replace('\n',' // ')))
                    print('board chairperson    {0}'.format(candidate[CHAIR].replace('\n',' // ')))
                    print('board dist member    {0}'.format(candidate[MEMBER].replace('\n',' // ')))
                    print('board unit member    {0}'.format(candidate[BOR3RD].replace('\n',' // ')))
                    print('board unit member Ph {0}'.format(candidate[BOR3RDPHONE].replace('\n',' // ')))
                    print('project cost         {0}'.format(candidate[PRJ_COST].replace('\n',' // ')))
                    print('project title        {0}'.format(candidate[PRJ_TITLE].replace('\n',' // ')))
                    print('project date         {0}'.format(candidate[PRJ_DATE].replace('\n',' // ')))
                    return
                    
    def read_input(self, csv_file=None):
        "process input, store internally for output and other operations.  Name defaults to what was given the constructor."

        with open(csv_file or self.file_name) as csv_file:
            csv_reader = csv.reader(csv_file)
            self.board_roster = []
            for candidate in csv_reader:
                candidate = list(map(str.strip, candidate))

                if len(candidate) == 0:
                    continue

                if candidate[0] == '#':
                    continue

                if len(candidate) < STAT-1:
                    print("Short record", candidate[0])
                    continue

                if candidate[STAT].lower().strip() == 'scheduled':
                    for pattern in '%m/%d/%Y %H:%M:%S', '%m/%d/%y %H:%M:%S', '%m/%d/%Y %H:%M', '%m/%d/%y %H:%M', '%m/%d/%Y', '%m/%d/%y':
                        try:
                            board_time = datetime.datetime.strptime(candidate[BORTIME], pattern)
                        except ValueError:
                            pass
                        else:
                            break
                    else:
                        #raise ValueError('Date/time not in expected formats: '+candidate[BORTIME])
                        sys.stderr.write("WARNING: Candidate {0} {1} date/time not in expected formats: {2}\n".format(candidate[CFIRST], candidate[CLAST], candidate[BORTIME]))
                        continue

                    if len(candidate) <= CHAIR:
                        sys.stderr.write("WARNING: Candidate {0} {1} scheduled for board but without assigned chair; ignored.\n".format(candidate[CFIRST], candidate[CLAST]))
                        continue

                    if len(candidate[CHAIR].split()) < 2:
                        sys.stderr.write("WARNING: Candidate {0} {1} scheduled for board but can't handle assigned chair {2}; ignored.\n".format(candidate[CFIRST], candidate[CLAST], candidate[CHAIR]))
                        continue

                    self.board_roster.append((board_time, candidate[:]))
                    #print "Added {0} to board roster for {1}".format(candidate, board_time)

                elif candidate[STAT].lower().strip() == 'pending':
                    self.pending_roster.append(candidate[:])

                elif candidate[STAT].lower().strip() == 'council':
                    self.council_roster.append(candidate[:])
                    date_sent = self.date_or_blank(candidate[D_SENT])
                    date_recv = self.date_or_blank(candidate[D_RECV])
                    notes = []
                    if int(candidate[YEARS]) >= 18:
                        notes.append("18+")
                    elif int(candidate[YEARS]) >= 17:
                        if int(candidate[MONTHS]) >= 11:
                            notes.append("17++")
                        elif int(candidate[MONTHS]) >= 10:
                            notes.append("17+")
                        else:
                            notes.append("17")

                    if not candidate[REFS] or int(candidate[REFS]) < 3:
                        notes.append("Will bring letters of recommendation to board")

                    if candidate[NOTES]:
                        notes.append(candidate[NOTES])

                    self.envelope_list.append((date_sent, date_recv, notes, candidate[:]))
                    #print "Added {0} to envelope list".format(candidate)

                elif candidate[STAT].lower().strip() not in ('council approval required','returning','missing','msg','status','future','pending', '', 'verification', 'completed', 'wcb', 'msg', 'incomplete','in progress'):
                    sys.stderr.write("WARNING: Candidate {0} {1} status '{2}' not understood; ignored.\n".format(
                        candidate[CFIRST], candidate[CLAST], candidate[STAT]
                    ))
                    continue

                if candidate[STAT].lower().strip() in ('wcb','msg','pending'):
                    self.all_roster.append(candidate[:])

    def reminder_mail(self, dry_run=False, interactive=True, extra_msg=None, no_message=False):
        staff = {}
        for board_time, candidate in self.board_roster:
            chair = candidate[CHAIR]
            member= candidate[MEMBER]
            scout = candidate[CFIRST] + ' ' + candidate[CLAST]
            if chair not in staff:
                staff[chair] = []
            staff[chair].append((board_time, 'chair', candidate[:]))
            if member not in staff:
                staff[member] = []
            staff[member].append((board_time, 'member', candidate[:]))
            if scout in staff:
                print("Warning: scout {0} appears multiple times!".format(scout))
            else:
                staff[scout] = ((board_time, 'scout', candidate[:]),)

#        print "=== {0} ===".format(scout)

        mail = smtplib.SMTP('localhost', 1025)
        for person in staff:
            cc = None
            if staff[person][0][1] == 'scout':
                s_emails = staff[person][0][2][CEMAIL].split(',')
                n_refs = int(staff[person][0][2][REFS])
                print("Sending to scout {0} at {1}: {2} letter(s) needed".format(person, s_emails,
                 'NO' if n_refs >= 3 else 3 - n_refs))

                if interactive and not yorn("Send this mail?"):
                    continue

                recips = ['pacifictrail.advancement@cpcbsa.org']
                cc = staff[person][0][2][SM_MAIL]
                cc2 = staff[person][0][2][SM_MAIL2]
                if not dry_run:
                    recips.append('PacificTrail.Advancement@cpcbsa.org')
                    recips.extend(s_emails)
                if cc:
                    print("CC: Scout leader {0}".format(cc))
                    if not dry_run:
                        recips.append(cc)
                if cc2:
                    print("CC: Scout leader {0}".format(cc2))
                    if not dry_run:
                        recips.append(cc2)
                    cc += ', ' + cc2

                try:
                    venue = Venue(staff[person][0][2][VENUECODE])
                except KeyError as e:
                    print(f"ERROR: {e}")
                    continue

                mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, 
'''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {2}{6}
Subject: Communication in Regard to Eagle Board of Review
Date: {7}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<p>Dear {0},</P>

<p>{4}</p>

<P>Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(person, self.date_time_str(staff[person][0][0]), ', '.join(s_emails),
#bring them with you to the board.  The letter{3}
#must be in {4} sealed envelope{3}.</p>'''.format(
'''<p><b>*** NOTE: According to our records, we have received {0} letter{1} of recommendation
for your application.  We need 3&ndash;5 letters for the board to review.</b>  Please arrange to have
the remaining {2} letter{3} sent to us in advance of your board, or if that is not possible,
bring them with you to the board.  The letter{3}
must be emailed <i>by the references</i> to <a href="mailto:pacifictrail.advancement@cpcbsa.org">pacifictrail.advancement@cpcbsa.org</a>. Do not have them sent to you first.</p>'''.format(
    n_refs, '' if n_refs==1 else 's',
    3-n_refs, '' if n_refs==2 else 's', 'a' if n_refs==2 else '',
    ) if n_refs < 3 else '',
    extra_msg or '',
    '',
    '\nCC: '+cc if cc else '',
#<p>Please arrive 15 minutes early along with your party, which
    timestamp()) if no_message else '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {2}{6}
Subject: Reminder of Eagle Board of Review
Date: {7}
Content-type: text/html

<p>Dear {0},</P>

<B>{4}</B>

<p>This is a reminder of your Eagle Board of Review
scheduled for <b>{1}</b>.</p>

<p>Please arrive 15 minutes early along with your party, which
should incude as many as possible of:
<ul>
  <li> You</li>
  <li> Your parents/guardians</li>
  <li> Your scout leader (if available)</li>
  <li> A member of your troop committee* to serve on your board of review</li>
  <li> A physical copy of your project workbook along with any photographs or other materials you wish the board to see about your project</li>
</ul></p>
<small><em>*It is the responsibility of your unit to assign this person. You
as the eagle candidate may not request a particular person.  You may wish to
double-check with your unit committee chair or your scoutmaster to remind them
about this.</em></small>

<P>The board will be held {8}</p>
<p>{9}</p>
{10}

{3}

<P>Please dress in your full Class A uniform.  If that is not possible, please
dress as formally as possible (e.g., dress shirt, tie, and slacks) to the
best of your ability.</p>

<P>Good luck!</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(person, self.date_time_str(staff[person][0][0]), ', '.join(s_emails),
'''<p><b>*** NOTE: According to our records, we have received {0} letter{1} of recommendation
for your application.  We need 3&ndash;5 letters for the board to review.</b>  Please arrange to have
the remaining {2} letter{3} sent to us in advance of your board via e-mail to <a href="mailto:pacifictrail.advancement@cpcbsa.org">pacifictrail.advancement@cpcbsa.org</a>.  The letter{3} must be sent in by the reference{3}, not sent to you first.</p>'''.format(
    n_refs, '' if n_refs==1 else 's',
    3-n_refs, '' if n_refs==2 else 's', 'a' if n_refs==2 else '',
    ) if n_refs < 3 else '',
    extra_msg or '',
    '',
    '\nCC: '+cc if cc else '',
    timestamp(),
    venue.desc,
    venue.address(),
    venue.extra_instructions())
)
            else:
                if person not in self.contacts:
                    print("Unable to contact", person)

                recips = ['PacificTrail.Advancement@cpcbsa.org']
                if not dry_run:
                    recips.append('PacificTrail.Advancement@cpcbsa.org')
                    for addr in self.contacts[person].split(','):
                        recips.append(addr.strip())
                        print("leader email: <{0}>".format(addr.strip()))

                print("Sending to leader {0} at {1}".format(person, self.contacts[person]))
                if interactive and not yorn("Send this mail?"):
                    continue
                try:
                    venue = Venue(staff[person][0][2][VENUECODE])
                except KeyError as e:
                    print(f"ERROR: {e}")
                    continue

                mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: {2}
Cc: pacifictrail.advancement@cpcbsa.org
Subject: Reminder of Eagle Boards of Review participation
Date: {4}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<P>Dear {0},</p>

<B>{3}</B>

<P>This is a reminder that you are on the schedule for our upcoming boards of
review:</p>

<pre>
    ROLE-- DATE-------- TIME---- UNIT--- SCOUT
{1}
</pre>

<P>The board will be held {5}</p>
<p>{6}</p>
{7}

<P>Thank you for your service to the district and to the scouts you will be
working with.  </p>

<P>Yours in Scouting,<br/>
Pacific Trail District<br/>
Eagle Board Co-coordinator</p>
'''.format(person, 
    '\n'.join(['    {0:6s} {1:21s} {2:7s} {3} {4}'.format(
        role, self.date_time_str(boardtime), candidate[UNIT], candidate[CFIRST], candidate[CLAST])
        for boardtime, role, candidate in staff[person]]
    ),
    self.contacts[person],
    extra_msg or '',
    timestamp(),
    venue.desc,
    venue.address(),
    venue.extra_instructions()))

    def board_chair_mail(self, dry_run=False, interactive=True, extra_msg=None, no_message=False, chair_name=None):
        staff = {}
        for board_time, candidate in self.board_roster:
            chair = candidate[CHAIR]
            member= candidate[MEMBER]
            scout = candidate[CFIRST] + ' ' + candidate[CLAST]
            if chair not in staff:
                staff[chair] = []
            staff[chair].append((board_time, 'chair', candidate[:]))
            if member not in staff:
                staff[member] = []
            staff[member].append((board_time, 'member', candidate[:]))
            if scout in staff:
                print("Warning: scout {0} appears multiple times!".format(scout))
            else:
                staff[scout] = ((board_time, 'scout', candidate[:]),)

        mail = smtplib.SMTP('localhost', 1025)
        for person in staff:
            cc = None
            if staff[person][0][1] == 'scout':
                s_emails = staff[person][0][2][CEMAIL].split(',')
                n_refs = int(staff[person][0][2][REFS])
                print("Sending to scout {0} at {1}: {2} letter(s) needed".format(person, s_emails,
                 'NO' if n_refs >= 3 else 3 - n_refs))

                if interactive and not yorn("Send this mail?"):
                    continue

                recips = ['PacificTrail.Advancement@cpcbsa.org']
                cc = staff[person][0][2][SM_MAIL]
                cc2 = staff[person][0][2][SM_MAIL2]
                if not dry_run:
                    recips.append('PacificTrail.Advancement@cpcbsa.org')
                    recips.extend(s_emails)
                if cc:
                    print("CC: Scout leader {0}".format(cc))
                    if not dry_run:
                        recips.append(cc)
                if cc2:
                    print("CC: Scout leader {0}".format(cc2))
                    if not dry_run:
                        recips.append(cc2)
                    cc += ', ' + cc2


                with open('zoom-{}'.format(self.date_time_str(staff[person][0][0]).replace(' ',''))) as zoom:
                    meeting_info = zoom.read()

                mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, 
'''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {2}{6}
Subject: Instructions for On-Line Eagle Board of Review
Date: {7}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<p>Dear {0},</P>

<p>{4}</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(person, self.date_time_str(staff[person][0][0]), ', '.join(s_emails),
'''<p><b>*** NOTE: According to our records, we have received {0} letter{1} of recommendation
for your application.  We need 3&ndash;5 letters for the board to review.</b>  Please arrange to have
the remaining {2} letter{3} sent to us in advance of your board.  The letter{3}
must be emailed <i>by the references</i> to <a href="mailto:pacifictrail.advancement@cpcbsa.org">pacifictrail.advancement@cpcbsa.org</a>. Do not have them sent to you first.</p>'''.format(
    n_refs, '' if n_refs==1 else 's',
    3-n_refs, '' if n_refs==2 else 's', 'a' if n_refs==2 else '',
    ) if n_refs < 3 else '',
    extra_msg or '',
    '',
    '\nCC: '+cc if cc else '',
#<p>Please arrive 15 minutes early along with your party, which
    timestamp()) if no_message else '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {2}{6}
Subject: Instructions for On-Line Eagle Board of Review for {0}
Date: {7}
Content-type: text/html

<p>Dear {0},</P>

<B>{4}</B>

<P>{9} will be the chairperson of your Eagle Board of Review
scheduled for <b>{1}</b>.</p>

{3}

<P>The following message provides instructions for how to participte in
this board of review. Please review the entire message and let me know if
you still have any questions about what to do.</P>

<P>We will be conducting the board using the Zoom video conferencing 
application. You can go ahead of time to the <a href="zoom.us">Zoom.US website</a>
to install the Zoom application to your computer and ensure it's all working. If you
do not wish to install the app, you can follow the link below to join the meeting
via web browser. Especially if you haven't used Zoom before, please try the software out
ahead of time to make sure your system, including microphone, camera, and speakers, are
working properly and you'll be able to participate when the time comes for your
board.</P>

<p>To keep to the current social distancing recommendations, we expect that each
of these persons will be joining the board of review separately from their own
locations if they are not already living in the same home together.</P>

<H2>Notes For All Participants (Including the Scout, Family, and Unit Leaders)</H2>
<P>Please arrange for a quiet, private location where you can join the meeting without anyone
not immediately participating in the board being in a position to overhear the conversation.</P>

<P>Start the Zoom application and type in the meeting ID and password shown below, or click on
the link provided to connect to the meeting 10 minutes before the scheduled time for your board
of review. You'll be in a waiting room while the Board has time to consult with one another.
When the board is ready to meet with you, you will be moved from the waiting room to the main
meeting.</P>

<p><b>Please be patient if the meeting does not start on time. There may be a scout
ahead of you still completing his or her board of review, or the board may need a little
extra time to consult before meeting with you.</b></p>

<p>Please also be aware of the following requirements from the BSA:</p>
<ul>
 <li>Everyone must be visible at all times, on camera, in view of all other participants. 
 <i>There must not be anyone at any location who is listening to the meeting but off-camera.</i></li>
 <li>The meeting <b>may not be recorded</b> by anyone.</li>
 <li>A parent or guardian of the Scout, or two registered adult leaders, must be directly present with the 
 Scout at the beginning of the conference. The Scouters may be from the council, district or unit. 
 Their role is to verify that the Scout is in a safe environment and that the board of review appears 
 to be in compliance with these requirements. Once all the members of the board of review are present 
 on their end of the call and introductions are completed, and the review is about to begin, anyone present 
 with the Scout must leave the room or move out of hearing distance unless they have specifically been 
 approved to remain as observers.</li>
 <li>Once the review process has been concluded, if the Scout is under age 18, the Scout's parent or guardian, 
 or two registered adult leaders, must rejoin the Scout. Their purpose is to be available to answer any questions 
 that may arise, to join in the celebration of the Scout's accomplishment, or to be party to any instructions 
 or arrangements regarding the appeals process or the reconvening of an incomplete review. Once this is done, 
 the board members end the call and sign off.</li>
</ul>
<P>Please dress in your full Class A uniform.  If that is not possible, please
dress as formally as possible (e.g., dress shirt, tie, and slacks) to the
best of your ability.</p>

<H2>Notes for Unit Representative on the Board of Review</H2>
<P>In addition to everything already written above, please note:</P>
<P>Please be ready to join the board of review 15 minutes prior to the scheduled start time. You will
be admitted immediately to confer with the other board members before being joined by the Scout and his/her
party.</P>
<P>As the unit representative, you will receive electronic copies of the documents involved in the board of review
in advance of the board. These documents must be kept strictly confidential and are only to be viewed
by members of the board of review. Once the board has concluded, please destroy all copies of them in your
possession.</P>
<P>Part of that paperwork will include a signature form for you to complete and e-mail back to the board
chair.</P>

<H2>Meeting Information:</H2>
{8}

<P>Again, if you have any questions about this process, please contact us at <a href="mailto:PacificTrail.Advancement@cpcbsa.org">PacificTrail.Advancement@cpcbsa.org</a>.</p>
<P>Good luck!</p>

<P>
Pacific Trail District<br/>
Eagle Board Co-coordinator</p>
'''.format(person, self.date_time_str(staff[person][0][0]), ', '.join(s_emails),
'''<p><b>*** NOTE: According to our records, we have received {0} letter{1} of recommendation
for your application.  We need 3&ndash;5 letters for the board to review.</b>  Please arrange to have
the remaining {2} letter{3} sent to us in advance of your board.  The letter{3}
must be emailed <i>by the references</i> to <a href="mailto:pacifictrail.advancement@cpcbsa.org">pacifictrail.advancement@cpcbsa.org</a>. Do not have them sent to you first.</p>'''.format(
    n_refs, '' if n_refs==1 else 's',
    3-n_refs, '' if n_refs==2 else 's', 'a' if n_refs==2 else '',
    ) if n_refs < 3 else '',
    extra_msg or '',
    '',
    '\nCC: '+cc if cc else '',
    timestamp(), meeting_info, 'I' if chair_name is None else chair_name)
)
            else:
                print("Skipping", person)


    def invitation_mail(self, dry_run=False, interactive=True, extra_msg=None, all_scouts=False, no_message=False):
        for candidate in self.all_roster if all_scouts else self.pending_roster:
            scout = candidate[CFIRST] + ' ' + candidate[CLAST]
            sm = candidate[SM_MAIL]
            sm2 = candidate[SM_MAIL2]
            scms = candidate[CEMAIL].split(',')

            mail = smtplib.SMTP('localhost', 1025)
            n_refs = int(candidate[REFS])

            print("Sending to leader {0} for scout {1}: {2} letter(s) needed".format(
                sm, scout, 'NO' if n_refs >= 3 else 3 - n_refs))

            if interactive and not yorn("Send this mail?"):
                continue

            recips = ['PacificTrail.Advancement@cpcbsa.org']
            cclist = []
            if not dry_run:
                recips.append('PacificTrail.Advancement@cpcbsa.org')
                for addr in sm.split(','):
                    recips.append(addr)
                if sm2: 
                    recips.append(sm2)
                    cclist.append(sm2)
                    print("CC: {0} (scout leader #2)".format(sm2))
                if scms:
                    recips.extend(scms)
                    cclist.extend(scms)
                    print("CC: {0} (scout)".format(scms))

            mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {5}{4}
Subject: Communication in Regard to Eagle Board of Review for {2}
Date: {6}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<p>Dear {0} Leaders,</P>

<P>{1}</P>

<P>Yours in Scouting,</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(
    candidate[UNIT],
    extra_msg or '',
    scout,
    '''In this case, we received {0} letter{1}.  Please arrange to have the remaining
{2} letter{3} sent to me in advance of the board or contact me to make other arrangements.  
The letter{3} must be in {4} sealed envelope{3} or emailed directly <i>by the references</i> themselves to pacifictrail.advancement@cpcbsa.org. Do not have them sent via the scout or unit.'''.format(n_refs, '' if n_refs==1 else 's', 3-n_refs, '' if n_refs==2 else 's',
    'a' if n_refs==2 else '') if n_refs < 3 else '''We have received the required letters already.
No further action is required on this point.''',
    '\nCc: {0}'.format(', '.join(cclist)) if cclist else '',
    sm,
    timestamp(),
) if no_message else '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {5}{4}
Subject: Invitation for Eagle Board of Review for {2}
Date: {6}
Content-type: text/html

<p>Dear {0} Leaders,</P>

<B>{1}</B>

<p>This is to inform you that {2} from your unit has submitted an application to be advanced
to the rank of Eagle Scout.  This application has been reviewed by the Council and approved to
move forward.  The next step is to schedule a board of review.</p>

<p>Eagle boards of review are handled at the District level.  As the unit leader, we are reaching
out to you to make contact with the scout and other unit personnel to coordinate a time and date
which would be convenient for everyone.</p>

<p><b>Boards are normally scheduled on the 3rd Wednesday and 4th Thursday of each month,</b>
except when those dates conflict with Thanksgiving and Christmas holidays.</p>

<P>In addition to the scout, we need a member of your unit to be on the board of review itself.
This person does not need to be registered with the BSA, but they should <i>not</i> be the scout's
current direct leader nor a relative (e.g., scoutmaster, assistant scoutmaster, crew advisor, etc.)  
The individual
needs to be at least 21 years of age and familiar with what the rank of Eagle Scout entails.
<B>The scout and his or her family may not have any role in selecting the individual to serve on their
board of review.</B></P>

<p>Please notify me of the name and phone number of the unit representative who will be on the board
if possible.</p>

<P>While not required, we would like to have the scout's parent(s) or guardian(s) and direct 
scout leader in attendance, as most boards of review prefer to spend a few minutes talking
with these individuals as well.</p>

<P>The scout is expected to appear in his or her full Class A BSA Uniform if at all possible.  Otherwise,
he or she should be dressed respectfully to be best of his or her ability (pressed shirt, tie, and slacks, for example).</P>

<P>{3}</P>

<p>Please contact me at <a href="mailto:PacificTrail.Advancement@cpcbsa.org">PacificTrail.Advancement@cpcbsa.org</a> with your preference of time and
date after conferring with the people listed above.  Have a 2nd or 3rd choice in mind as well
in case the spot you wish is not available.</P>

<P>Yours in Scouting,</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(
#<P><B>We hold boards of review at 7:00 and 8:15 P.M. on the 2nd and 4th Thursdays of every month
#(except for Thanksgiving and Christmas holiday weeks).</B></P>
#<p>The board will be held in the following location:</p>
#
#<P>Valley Community Presbyterian Church<br/>
#8060 SW Brentwood St<br/>
#Portland, OR 97225</P>
    candidate[UNIT],
    extra_msg or '',
    scout,
    '''We request that all scouts request letters of recommendation from 3&ndash;5 adults who know
them well. These are not necessarily the same as those listed as references on the application, although
that is best. If not, they should represent the same mix of areas of life (education, religious life,
parents, general personal references, employment) if possible.
    In this case, we received {0} letter{1}.  Please arrange to have the remaining
{2} letter{3} sent to me in advance of the board or contact me to make other arrangements.  
The letter{3} must be in {4} sealed envelope{3} or emailed by the references themselves (not via the scout or unit) to pacifictrail.advancement@cpcbsa.org'''.format(n_refs, '' if n_refs==1 else 's', 3-n_refs, '' if n_refs==2 else 's',
    'a' if n_refs==2 else '') if n_refs < 3 else '''We have received the required letters of recommendation already.
No further action is required on this point.''',
    '\nCc: {0}'.format(', '.join(cclist)) if cclist else '',
    sm,
    timestamp(),
))

    def received_mail(self, dry_run=False, interactive=True, extra_msg=None, all_scouts=False, no_message=False):
        "Thank candidate for submitting their application"

        for candidate in self.all_roster if all_scouts else self.council_roster:
            scout = candidate[CFIRST] + ' ' + candidate[CLAST]
            sm = candidate[SM_MAIL]
            sm2 = candidate[SM_MAIL2]
            scms = candidate[CEMAIL].split(',')

            mail = smtplib.SMTP('localhost', 1025)
            n_refs = int(candidate[REFS])

            print("Sending to leader {0} for scout {1}: {2} letter(s) needed".format(
                sm, scout, 'NO' if n_refs >= 3 else 3 - n_refs))

            if interactive and not yorn("Send this mail?"):
                continue

            recips = ['PacificTrail.Advancement@cpcbsa.org']
            cclist = []
            if not dry_run:
                recips.append('PacificTrail.Advancement@cpcbsa.org')
                for addr in sm.split(','):
                    recips.append(addr)
                if sm2: 
                    recips.append(sm2)
                    cclist.append(sm2)
                    print("CC: {0} (scout leader #2)".format(sm2))
                if scms:
                    recips.extend(scms)
                    cclist.extend(scms)
                    print("CC: {0} (scout)".format(scms))

            mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, '''From: "PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {5}{4}
Subject: Receipt of Eagle Application for {2}
Date: {6}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<p>Dear {0} Leaders,</P>

<P>{1}</P>

<P>Yours in Scouting,</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(
    candidate[UNIT],
    extra_msg or '',
    scout,
    '''In this case, we received {0} letter{1}.  Please arrange to have the remaining
{2} letter{3} sent to me in advance of the board or contact me to make other arrangements.  
The letter{3} must be in {4} sealed envelope{3} or emailed by the references themselves (not via the scout or unit) to pacifictrail.advancement@cpcbsa.org'''.format(n_refs, '' if n_refs==1 else 's', 3-n_refs, '' if n_refs==2 else 's',
    'a' if n_refs==2 else '') if n_refs < 3 else '''We have received the required letters already.
No further action is required on this point.''',
    '\nCc: {0}'.format(', '.join(cclist)) if cclist else '',
    sm,
    timestamp(),
) if no_message else '''From: "PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: pacifictrail.advancement@cpcbsa.org, {5}{4}
Subject: Receipt of Eagle Application for {2}
Date: {6}
Content-type: text/html

<p>Dear {0} Leaders,</P>

<B>{1}</B>

<p>This is to inform you that {2} from your unit has submitted an application to be advanced
to the rank of Eagle Scout.  We have checked over the paperwork and have now passed it on to
the Council office for their review. Once they have granted approval for us to move forward
with a board of review, we will be in touch with you to work out the arrangements for that
to take place.
</p>

<p>{3}</p>

<P>Yours in Scouting,</p>

<P>
Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(
#<P><B>We hold boards of review at 7:00 and 8:15 P.M. on the 2nd and 4th Thursdays of every month
#(except for Thanksgiving and Christmas holiday weeks).</B></P>
#<p>The board will be held in the following location:</p>
#
#<P>Valley Community Presbyterian Church<br/>
#8060 SW Brentwood St<br/>
#Portland, OR 97225</P>
    candidate[UNIT],
    extra_msg or '',
    scout,
    '''We request that all scouts request letters of recommendation from 3&ndash;5 adults who know
them well. These are not necessarily the same as those listed as references on the application, although
that is best. If not, they should represent the same mix of areas of life (education, religious life,
parents, general personal references, employment) if possible.
    In this case, we received {0} letter{1}.  Please arrange to have the remaining
{2} letter{3} sent to me in advance of the board or contact me to make other arrangements.  
The letter{3} must be in {4} sealed envelope{3} or emailed by the references themselves (not via the scout or unit) to pacifictrail.advancement@cpcbsa.org'''.format(n_refs, '' if n_refs==1 else 's', 3-n_refs, '' if n_refs==2 else 's',
    'a' if n_refs==2 else '') if n_refs < 3 else '''We have received the required letters of recommendation already.
No further action is required on this point.''',
    '\nCc: {0}'.format(', '.join(cclist)) if cclist else '',
    sm,
    timestamp(),
))
            
    def query_mail(self, dry_run=False, interactive=True, extra_msg=None, board_date='the 2nd or 4th Thursday', no_message=False, subject=None):
        mail = smtplib.SMTP('localhost', 1025)
        for person in self.contacts:
            recips = ['PacificTrail.Advancement@cpcbsa.org']
            if not dry_run:
                recips.append('PacificTrail.Advancement@cpcbsa.org')
                for addr in self.contacts[person].split(','):
                    recips.append(addr.strip())
                    print("leader email: <{0}>".format(addr.strip()))

            print("Sending to leader {0} at {1}".format(person, self.contacts[person]))
            if interactive and not yorn("Send this mail?"):
                continue

            mail.sendmail('PacificTrail.Advancement@cpcbsa.org', recips, '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: {2}
Cc: pacifictrail.advancement@cpcbsa.org
Subject: {5}
Date: {4}
Reply-To: PacificTrail.Advancement@cpcbsa.org
Content-type: text/html

<P>Dear {0},</p>

<p>{3}</p>

<P>Yours in Scouting,<br/>

Eagle Board of Review Co-Coordinator<br/>
Pacific Trail District, Cascade Pacific Council (Volunteer)<br/>
PacificTrail.Advancement@cpcbsa.org</p>
'''.format(person, 
    board_date,
    self.contacts[person],
    extra_msg or '',
    timestamp(),
    subject or 'Communication in Regard to Upcoming Eagle Boards of Review',
    ) if no_message else '''From: "BSA CPC PACIFIC TRAIL DISTRICT Eagle Boards of Review" <PacificTrail.Advancement@cpcbsa.org>
To: {2}
Cc: pacifictrail.advancement@cpcbsa.org
Subject: {5}
Date: {4}
Content-type: text/html

<P>Dear {0},</p>

<B>{3}</B>

<P>Boards of review are coming up. Would you be available and willing to help us
by participating on board(s) for the next set on {1}?
</p>

<P>Thank you for your service to the district and to the scouts you will be
working with.  </p>

<P>Yours in Scouting,<br/>

Pacific Trail District<br/>
Eagle Board Co-coordinator</p>
'''.format(person, 
    board_date,
    self.contacts[person],
    extra_msg or '',
    timestamp(),
    subject or 'Help with Upcoming Eagle Boards of Review'
    )
)
            

        # 0  1     2    3    4     5     6         7     8    9     10
        # ID,First,Last,Unit,Email,Phone,Alt Phone,First,Last,Phone,Alt Phone,
        # 11       12    13     14   15     16    17      18     19         20
        # Birthday,Years,Months,Days,Status,Board,Project,Packet,References,App Rec'd,
        # 21         22       23       24           25          26        27             
        # To Council,Verified,BoR Date,Recommended?,To National,BoR Chair,BoR 2nd Member,
        # 28           29
        # Packet Notes,NOTES,,,,,,
