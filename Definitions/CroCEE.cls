%% =====================================================================
%%  CroCEE.cls -- LaTeX class for full papers and extended abstracts of
%%                the Croatian Conference on Earthquake Engineering (CroCEE)
%%
%%  The class is conference-independent: edition, name, dates, place,
%%  copyright line and logo are read from Definitions/conference.cfg
%%  (or set with \croceeconference{key=value,...} in the preamble).
%%  Additional packages are read from Definitions/packages.tex.
%%
%%  Authors: Marin Grubišić and Vedran Jagodnik, with the assistance of
%%           Claude (Anthropic) and Gemini (Google).
%%           Version 6.2, October 2026.
%%
%%  LaTeX re-implementation of the layout of the official MS Word templates
%%  of the conference (full paper and extended abstract; the extended
%%  abstract has no Abstract section and is limited to 2 pages). The page
%%  geometry, fonts and spacings follow the Word templates; see README.md
%%  ("Validation") for what has and has not been compared with Word.
%%
%%  Usage:  \documentclass{CroCEE}            % the file CroCEE.cls is in the
%%          \documentclass[extended]{CroCEE}  % same folder as the .tex file
%%
%%  ARCHITECTURE NOTE:
%%  All layout, font, metric, and spacing parameters are centralized in
%%  Section 1 (MASTER CONFIGURATION PARAMETERS). All dependent quantities
%%  (\baselineskip values, \crocee@ndesc, \crocee@xb, \topskip, section
%%  skips, float compensations, etc.) are dynamically computed via
%%  \crocee@recalculate. Parameters can be edited in Section 1 or
%%  overridden in the document preamble using \croceesetup{key=value}.
%% =====================================================================
%% Required: LaTeX kernel 2021-06-01 or newer (TeX Live 2021+, MiKTeX 21.6+,
%% Overleaf) for the hook system (\AddToHook with shipout/background,
%% para/begin, env/.../after). \fpeval is in the kernel since 2022-06; on
%% older kernels the xfp package provides it (harmless wrapper on newer ones).
\NeedsTeXFormat{LaTeX2e}[2021/06/01]
\ProvidesClass{Definitions/CroCEE}[2026/10/29 v6.2 CroCEE paper class (Palatino, exact line spacing)]
\RequirePackage{xfp}

%% =====================================================================
%%  1. MASTER CONFIGURATION PARAMETERS (SINGLE SOURCE OF TRUTH)
%% =====================================================================

%% --- 1.1 Line Metrics
%% The Word template uses "Exactly" line spacing of 1.26 x font size for
%% every style (11 pt -> 13.86 pt, 10 pt -> 12.6 pt, 9 pt -> 11.34 pt,
%% 13 pt -> 16.38 pt, 14 pt -> 17.64 pt). Word keeps the font descent
%% (Palatino Linotype: 0.285 x size) at the bottom of the line box and takes
%% the rest from the top, so the ascent of the line box is pitch - descent.
%% (lh = 1.26 with mult = 1.0 is the "exact" model; lh = 1.05 with mult = 1.2
%% would give the same pitch but the extra-space-below model of Word's
%% "multiple" spacing, which does not match "Exactly".)
\def\crocee@lh{1.26}               % line pitch / font size (Word "Exactly")
\def\crocee@desc{0.285}            % descent of the line box / font size
\def\crocee@asc{0.975}             % ascent of the line box = lh - desc

%% --- 1.2 Line Spacing Multipliers (1.0 = the exact pitch above)
\def\crocee@mult{1.0}               % body text (\normalsize) line spacing multiple
\def\crocee@mult@abstract{1.0}       % abstract body line spacing multiple
\def\crocee@mult@caption{1.0}        % figure/table captions line spacing multiple
\def\crocee@mult@table{1.0}          % table cells line spacing multiple
\def\crocee@mult@bib{1.0}            % references list line spacing multiple

%% --- 1.3 Base Font Size Scale (in Word points: 1 bp = 1.00375 pt)
\def\crocee@sz@tiny{6}
\def\crocee@sz@scriptsize{8}
\def\crocee@sz@footnotesize{9}
\def\crocee@sz@small{10}
\def\crocee@sz@normalsize{11}
\def\crocee@sz@large{12}
\def\crocee@sz@Large{13}
\def\crocee@sz@LARGE{14}
\def\crocee@sz@huge{16}
\def\crocee@sz@Huge{20}

%% --- 1.4 Semantic Element Font Sizes (in bp, mapped to base scale)
\def\crocee@sz@title{\crocee@sz@LARGE}             % 14 bp: paper title
\def\crocee@sz@authors{\crocee@sz@normalsize}      % 11 bp: author names
\def\crocee@sz@affil{\crocee@sz@footnotesize}      %  9 bp: affiliations & header
\def\crocee@sz@titlespacer{\crocee@sz@large}       % 12 bp: empty spacer lines in title block
\def\crocee@sz@abshead{\crocee@sz@normalsize}      % 11 bp: "Abstract" heading
\def\crocee@sz@abstext{\crocee@sz@small}           % 10 bp: abstract text
\def\crocee@sz@keywords{\crocee@sz@small}          % 10 bp: keywords text
\def\crocee@sz@sec{\crocee@sz@Large}               % 13 bp: \section heading
\def\crocee@sz@subsec{\crocee@sz@normalsize}       % 11 bp: \subsection heading
\def\crocee@sz@subsubsec{\crocee@sz@normalsize}    % 11 bp: \subsubsection heading
\def\crocee@sz@caption{\crocee@sz@small}           % 10 bp: captions
\def\crocee@sz@table{\crocee@sz@small}             % 10 bp: table content
\def\crocee@sz@bib{\crocee@sz@small}               % 10 bp: bibliography items

%% --- 1.5 Page Geometry (in twips: 1 twip = 1/20 bp; 1440 tw = 2.54 cm)
\def\crocee@tw@top{1843}             % top margin: 3.25 cm
\def\crocee@tw@bottom{1440}          % bottom margin: 2.54 cm
\def\crocee@tw@left{1440}            % left margin: 2.54 cm
\def\crocee@tw@right{1440}           % right margin: 2.54 cm

%% --- 1.6 Paragraph, Equation, Section & Float Spacing
\def\crocee@par@skip{6bp}            % space after each body paragraph
\def\crocee@par@indent{1cm}          % paragraph indentation (when enabled)
%% Display-equation skip (\abovedisplayskip, \belowdisplayskip, ...).
%% The Word template specifies 8 pt before and after an equation, measured
%% between line boxes of "exact" height. TeX measures the display skip between
%% the real boxes of the neighbouring lines, so the same number would give a
%% different distance. 6 bp is an EMPIRICAL value, calibrated against a PDF
%% exported from Word (not derived from font metrics); re-check it with a new
%% Word export whenever the Word template or the font metrics change.
\def\crocee@eq@skip{6bp}
\def\crocee@title@before{24bp}       % space before paper title
\def\crocee@title@after{6bp}         % space directly after paper title
\def\crocee@abs@before{0bp}         % space before Abstract heading
\def\crocee@sec@before{12bp}         % space before \section
\def\crocee@sec@after{6bp}           % space after \section
\def\crocee@subsec@before{12bp}      % space before \subsection & \subsubsection
\def\crocee@subsec@after{6bp}        % space after \subsection & \subsubsection
\def\crocee@cap@tableskip{6bp}       % space below table caption
\def\crocee@cap@figureskip{12bp}     % space below figure caption
\def\crocee@float@gap{6bp}           % base gap above/below in-text floats
\def\crocee@float@sep{12bp}          % space between two adjacent floats
%% Fine correction (EMPIRICAL, calibrated against the Word export, like
%% \crocee@eq@skip) of the space before a \section that directly follows a
%% float; it is subtracted in \crocee@floatafter (Section 11). Set to 0bp to
%% see the uncorrected, purely metric value.
\def\crocee@float@trim{0.75bp}
\def\crocee@table@colsep{5.4bp}      % horizontal cell padding in tables (0.19 cm)
\def\crocee@table@noteskip{4pt}      % space above \tablenotes
\def\crocee@bib@itemsep{6bp}         % space between bibliography items

%% --- 1.7 Indentation & Tab-Stop Parameters (in twips)
\def\crocee@tw@affilhang{180}        % affiliation hanging indent
\def\crocee@tw@listleft{720}         % list left margin (1.27 cm)
\def\crocee@tw@listlabel{360}        % list label indent/width (0.635 cm)
\def\crocee@tw@secstopA{340}         % section 1st tab stop
\def\crocee@tw@secstopB{360}         % section 2nd tab stop / hangindent
\def\crocee@tw@subsecstopB{454}      % subsection 2nd tab stop / hangindent
\def\crocee@tw@subsubsecstopA{357}   % subsubsection 1st tab stop / hangindent
\def\crocee@tw@bibindent{357}        % bibliography hanging indent (0.63 cm)
\def\crocee@tw@tabgrid{720}          % default Word tab stop grid (0.5 in)

%% --- 1.8 Superscript Scaling
\def\crocee@sup@scale{0.65}          % superscript font size ratio
\def\crocee@sup@raise{0.33}          % superscript baseline raise ratio

%% --- 1.9 Page Count Limits (for automatic compilation warnings)
\def\crocee@pages@extmax{2}          % max pages for extended abstract
\def\crocee@pages@fullmin{6}         % min pages for full paper
\def\crocee@pages@fullmax{12}        % max pages for full paper

%% --- 1.10 External Directories & Bibliography Style
\def\crocee@figdir{Figures/}                 % default folder for figures and logo
\def\crocee@defdir{Definitions/}             % default folder for .bst / definitions
\def\crocee@bst{Elsevier}                   % default BibTeX style: Definitions/Elsevier.bst

%% =====================================================================
%%  2. CLASS OPTIONS
%% =====================================================================
\newif\ifcrocee@extended \crocee@extendedfalse
\DeclareOption{fullpaper}{\crocee@extendedfalse}
\DeclareOption{extended}{\crocee@extendedtrue}
\newif\ifcrocee@parindent \crocee@parindenttrue
\newif\ifcrocee@hyphenation \crocee@hyphenationfalse
\newif\ifcrocee@ulurl \crocee@ulurltrue
\DeclareOption{noparindent}{\crocee@parindentfalse}
\DeclareOption{hyphenation}{\crocee@hyphenationtrue}
\DeclareOption{nourlunderline}{\crocee@ulurlfalse}
\DeclareOption{a4paper}{}
\DeclareOption{11pt}{}
\DeclareOption{draft}{\PassOptionsToClass{draft}{article}}
\DeclareOption{final}{\PassOptionsToClass{final}{article}}
\DeclareOption*{\ClassWarning{CroCEE}{Unknown class option `\CurrentOption' ignored
  (valid: fullpaper, extended, noparindent, hyphenation, nourlunderline)}}
\ProcessOptions\relax

\LoadClass[a4paper,11pt]{article}

%% =====================================================================
%%  3. CORE METRIC HELPERS & AUTOMATIC RECALCULATION ENGINE
%% =====================================================================
%% 1 twip = 1/20 bp (Word's internal unit)
\newcommand*\crocee@tw[1]{\dimexpr #1bp/20\relax}

%% Derived metric macros (populated by \crocee@recalculate)
\def\crocee@ndesc{0.29700}
\def\crocee@xb{2.2308bp}
\def\crocee@eqskip{6bp}

%% Vertical correction when font size changes between two single-spaced
%% paragraphs: gap' = gap + desc*size_prev - (pitch - asc)*size_next
\newcommand*\crocee@vgap[3]{%
  \dimexpr #2 + \crocee@desc\dimexpr #1bp\relax
            - \crocee@ndesc\dimexpr #3bp\relax\relax}

%% A strut with Word's line metrics for the current font size
\newcommand*\crocee@strut{%
  \vrule\@width\z@\@height\crocee@asc\dimexpr\f@size\p@\relax
  \@depth\crocee@desc\dimexpr\f@size\p@\relax}

%% Font selector for a given size in Word points (bp) and optional line
%% spacing multiplier (#1 defaults to 1 = single spacing)
\newcommand*\crocee@fs[2][1]{%
  \fontsize{\fpeval{round((#2)*1.00375, 5)}pt}%
           {\fpeval{round((#2)*1.00375*\crocee@lh*(#1), 5)}pt}\selectfont}

%% Internal helper to define LaTeX size commands with exact bp-to-pt conversion
\newcommand*\crocee@defsizecmd[3]{% #1 = cmd, #2 = size (bp), #3 = mult
  \edef\crocee@tmp@sz{\fpeval{round((#2)*1.00375, 5)}}%
  \edef\crocee@tmp@bs{\fpeval{round((#2)*1.00375*\crocee@lh*(#3), 5)}}%
  \expandafter\edef\csname crocee@set@\string#1\endcsname{%
    \noexpand\@setfontsize\noexpand#1{\crocee@tmp@sz}{\crocee@tmp@bs}}%
}

%% Allocate dependent length registers
\newlength\crocee@secbefore
\newlength\crocee@secafter
\newlength\crocee@subbefore
\newlength\crocee@subafter
\newskip\crocee@secbeforeskip
\newlength\crocee@floatcomp

%% \crocee@floatcomp -- correction for the change of font size at the end of a
%% float. The last line of a float is a 10 bp caption or table line, the next
%% line is an 11 bp body line. With Word's line model (descent kept at the
%% bottom of the line box, see Section 1.1) the baseline distance changes by
%% desc*(size_prev - size_next) when the font size changes, which is exactly
%% \crocee@vgap{10}{0pt}{11} = -0.285 bp. \crocee@floatend subtracts
%% \crocee@floatcomp (= +0.285 bp for the default sizes) to apply it; the
%% formula below is this difference written with the master parameters, so it
%% follows automatically if font sizes or line pitch are changed.
%%
%% Master recalculation routine: updates all derived quantities from Section 1
\newcommand*\crocee@recalculate{%
  %% 1. Derived line metrics
  \edef\crocee@ndesc{\fpeval{round(\crocee@lh - \crocee@asc, 5)}}%
  \edef\crocee@xb{\fpeval{round((\crocee@mult - 1) * \crocee@lh * (\crocee@sz@normalsize), 5)}bp}%
  \edef\crocee@eqskip{\crocee@eq@skip}%
  %% 2. Pre-calculate font size definitions (1 bp = 1.00375 pt)
  \crocee@defsizecmd\normalsize{\crocee@sz@normalsize}{\crocee@mult}%
  \crocee@defsizecmd\small{\crocee@sz@small}{1}%
  \crocee@defsizecmd\footnotesize{\crocee@sz@footnotesize}{1}%
  \crocee@defsizecmd\scriptsize{\crocee@sz@scriptsize}{1}%
  \crocee@defsizecmd\tiny{\crocee@sz@tiny}{1}%
  \crocee@defsizecmd\large{\crocee@sz@large}{1}%
  \crocee@defsizecmd\Large{\crocee@sz@Large}{1}%
  \crocee@defsizecmd\LARGE{\crocee@sz@LARGE}{1}%
  \crocee@defsizecmd\huge{\crocee@sz@huge}{1}%
  \crocee@defsizecmd\Huge{\crocee@sz@Huge}{1}%
  %% 3. Paragraph and top-of-page parameters
  \setlength\topskip{\fpeval{round(\crocee@asc * (\crocee@sz@normalsize) * 1.00375, 5)}pt}%
  \setlength\parskip{\crocee@par@skip}%
  \ifcrocee@parindent
    \setlength\parindent{\crocee@par@indent}%
  \else
    \setlength\parindent{0pt}%
  \fi
  %% 4. Sectioning vertical spacings
  \setlength\crocee@secbefore{%
    \dimexpr\crocee@vgap{\crocee@sz@normalsize}{\crocee@sec@before}{\crocee@sz@sec}+\crocee@xb-\parskip\relax}%
  \global\crocee@secbeforeskip\crocee@secbefore
  \setlength\crocee@secafter{%
    \dimexpr\crocee@vgap{\crocee@sz@sec}{\crocee@sec@after}{\crocee@sz@normalsize}-\crocee@xb-\parskip\relax}%
  \setlength\crocee@subbefore{%
    \dimexpr\crocee@vgap{\crocee@sz@normalsize}{\crocee@subsec@before}{\crocee@sz@subsec}+\crocee@xb-\parskip\relax}%
  \setlength\crocee@subafter{%
    \dimexpr\crocee@vgap{\crocee@sz@subsec}{\crocee@subsec@after}{\crocee@sz@normalsize}-\crocee@xb-\parskip\relax}%
  %% 5. Float spacings and post-float baseline compensation
  \setlength\intextsep{\dimexpr\crocee@float@gap+\crocee@xb\relax}%
  \setlength\textfloatsep{\dimexpr\crocee@float@gap+\crocee@xb\relax}%
  \setlength\crocee@floatcomp{\dimexpr\crocee@xb
    +(\crocee@mult\dimexpr\crocee@lh\dimexpr\crocee@sz@normalsize bp\relax\relax)
    -\crocee@asc\dimexpr\crocee@sz@normalsize bp\relax
    -\crocee@desc\dimexpr\crocee@sz@caption bp\relax\relax}%
  \setlength\floatsep{\crocee@float@sep}%
  %% 6. Update caption skips if caption package is already loaded
  \@ifpackageloaded{caption}{%
    \captionsetup[table]{position=above, skip=\crocee@cap@tableskip}%
    \captionsetup[figure]{position=below, skip=\crocee@cap@figureskip}%
  }{}%
}

%% Bind standard LaTeX font size macros to the dynamically computed definitions
\renewcommand\normalsize{%
  \csname crocee@set@\string\normalsize\endcsname
  \abovedisplayskip\crocee@eqskip \abovedisplayshortskip\crocee@eqskip
  \belowdisplayskip\crocee@eqskip \belowdisplayshortskip\crocee@eqskip
  \let\@listi\@listI}
\renewcommand\small{%
  \csname crocee@set@\string\small\endcsname
  \abovedisplayskip\crocee@eqskip \abovedisplayshortskip\crocee@eqskip
  \belowdisplayskip\crocee@eqskip \belowdisplayshortskip\crocee@eqskip}
\renewcommand\footnotesize{\csname crocee@set@\string\footnotesize\endcsname}
\renewcommand\scriptsize{\csname crocee@set@\string\scriptsize\endcsname}
\renewcommand\tiny{\csname crocee@set@\string\tiny\endcsname}
\renewcommand\large{\csname crocee@set@\string\large\endcsname}
\renewcommand\Large{\csname crocee@set@\string\Large\endcsname}
\renewcommand\LARGE{\csname crocee@set@\string\LARGE\endcsname}
\renewcommand\huge{\csname crocee@set@\string\huge\endcsname}
\renewcommand\Huge{\csname crocee@set@\string\Huge\endcsname}

%% Run initial calculation and activate \normalsize
\crocee@recalculate
\normalsize

%% Public key-value interface for adjusting parameters in the preamble
\ExplSyntaxOn
\keys_define:nn { crocee }
  {
    linespacing     .tl_set:N = \crocee@mult,
    abstractspacing .tl_set:N = \crocee@mult@abstract,
    captionspacing  .tl_set:N = \crocee@mult@caption,
    tablespacing    .tl_set:N = \crocee@mult@table,
    bibspacing      .tl_set:N = \crocee@mult@bib,
    linepitch       .tl_set:N = \crocee@lh,
    ascent          .tl_set:N = \crocee@asc,
    descent         .tl_set:N = \crocee@desc,
    bodyfontsize    .tl_set:N = \crocee@sz@normalsize,
    smallfontsize   .tl_set:N = \crocee@sz@small,
    footnotefontsize.tl_set:N = \crocee@sz@footnotesize,
    secfontsize     .tl_set:N = \crocee@sz@sec,
    subsecfontsize  .tl_set:N = \crocee@sz@subsec,
    titlefontsize   .tl_set:N = \crocee@sz@title,
    parskip         .tl_set:N = \crocee@par@skip,
    parindent       .tl_set:N = \crocee@par@indent,
    eqskip          .tl_set:N = \crocee@eq@skip,
    secbefore       .tl_set:N = \crocee@sec@before,
    secafter        .tl_set:N = \crocee@sec@after,
    subsecbefore    .tl_set:N = \crocee@subsec@before,
    subsecafter     .tl_set:N = \crocee@subsec@after,
    tablecapskip    .tl_set:N = \crocee@cap@tableskip,
    figurecapskip   .tl_set:N = \crocee@cap@figureskip,
    floatgap        .tl_set:N = \crocee@float@gap,
    floatsep        .tl_set:N = \crocee@float@sep,
    floattrim       .tl_set:N = \crocee@float@trim,
    bibitemsep      .tl_set:N = \crocee@bib@itemsep,
  }
\NewDocumentCommand \croceesetup { m }
  {
    \keys_set:nn { crocee } { #1 }
    \crocee@recalculate
    \normalsize
  }
\NewDocumentCommand \croceerecalculate { }
  {
    \crocee@recalculate
    \normalsize
  }
\ExplSyntaxOff

%% =====================================================================
%%  4. PAGE LAYOUT (A4, Word margins), NO PAGE NUMBERS
%% =====================================================================
\RequirePackage[a4paper,
  top=\crocee@tw{\crocee@tw@top}, bottom=\crocee@tw{\crocee@tw@bottom},
  left=\crocee@tw{\crocee@tw@left}, right=\crocee@tw{\crocee@tw@right},
  nohead, nofoot]{geometry}
\raggedbottom
\pagestyle{empty}
\let\ps@plain\ps@empty
\let\ps@headings\ps@empty
\let\ps@myheadings\ps@empty

%% =====================================================================
%%  5. FONTS (Palatino Linotype / TeX Gyre Pagella)
%% =====================================================================
\RequirePackage{iftex}
\RequirePackage{amsmath}
\ifPDFTeX
  \RequirePackage[T1]{fontenc}
  \IfFileExists{newpxtext.sty}{%
    \RequirePackage{newpxtext}%
    \RequirePackage{newpxmath}%
  }{%
    \RequirePackage{mathpazo}%
  }%
\else
  \IfFileExists{newpxmath.sty}{\RequirePackage{newpxmath}}{\RequirePackage{mathpazo}}%
  \RequirePackage[no-math]{fontspec}%
  \IfFontExistsTF{Palatino Linotype}{%
    \setmainfont{Palatino Linotype}%
  }{%
    \ClassWarningNoLine{CroCEE}{Palatino Linotype not installed; using TeX Gyre
      Pagella (same glyph widths)}%
    \setmainfont{TeX Gyre Pagella}%
  }%
\fi
\IfFileExists{t1lmtt.fd}{\renewcommand\ttdefault{lmtt}}{}
\RequirePackage{bm}

%% Superscripts scaled and raised according to Section 1.8
\def\@textsuperscript#1{{\m@th
  \edef\crocee@supsize{\strip@pt\dimexpr\crocee@sup@scale\dimexpr\f@size\p@\relax\relax}%
  \raise\crocee@sup@raise\dimexpr\f@size\p@\relax
  \hbox{\fontsize{\crocee@supsize}{\z@}\selectfont#1}}}

%% =====================================================================
%%  6. PARAGRAPHS & LISTS
%% =====================================================================
\setlength\lineskiplimit{-1bp}       % Word does not open up lines for superscripts
\setlength\emergencystretch{3em}
\frenchspacing                       % single space after a full stop
\ifcrocee@hyphenation\else
  \hyphenpenalty=10000 \exhyphenpenalty=10000
  \uchyph=0
\fi
\widowpenalty=10000 \clubpenalty=10000

%% Footnotes: 9 bp text (\footnotesize), no paragraph indent, hanging mark.
%% The Word template does not specify footnotes, so this is a conservative
%% choice that stays within the page geometry.
\renewcommand\footnoterule{\kern-3\p@\hrule\@width.3\columnwidth\@height.5bp\kern2.5bp}
\skip\footins=8bp plus 2bp minus 2bp\relax
\setlength\footnotesep{6bp}
\long\def\@makefntext#1{%
  \parindent\z@\noindent\hangindent\crocee@tw{\crocee@tw@bibindent}\hangafter\@ne
  \hb@xt@\crocee@tw{\crocee@tw@bibindent}{\@makefnmark\hfil}#1}

%% Endnotes: \endnote{text} in the text, \printendnotes where the list should
%% appear (usually before the References). Same 9 bp style as footnotes.
\RequirePackage{endnotes}
\renewcommand\notesname{Notes}
\renewcommand\theendnote{\roman{endnote}}   % i, ii, ... (footnotes: 1, 2, ...)
\renewcommand\makeenmark{\textsuperscript{\theenmark}}
\def\enotesize{\footnotesize}
\def\enoteheading{\section*{\notesname}}
\def\enoteformat{%
  \parindent\z@\leftskip\crocee@tw{\crocee@tw@bibindent}\rightskip\z@
  \leavevmode\llap{\theenmark.\hskip.3em}}
\newcommand\printendnotes{\theendnotes}

\IfFileExists{enumitem.sty}{%
  \RequirePackage{enumitem}%
  \setlist{topsep=\z@, partopsep=\z@, itemsep=\z@, parsep=\z@,
    labelsep=\z@, align=left}%
  \setlist[1]{leftmargin=\crocee@tw{\crocee@tw@listleft},
    labelindent=\crocee@tw{\crocee@tw@listlabel},
    labelwidth=\crocee@tw{\crocee@tw@listlabel}}%
  \setlist[2]{leftmargin=\crocee@tw{\crocee@tw@listleft},
    labelindent=\crocee@tw{\crocee@tw@listlabel},
    labelwidth=\crocee@tw{\crocee@tw@listlabel}}%
  \setlist[itemize,1]{label=\textbullet}%
  \setlist[enumerate,1]{label=\arabic*.}%
}{}

%% =====================================================================
%%  7. HEADER & FOOTER
%% =====================================================================
\RequirePackage{xcolor}
\definecolor{croceelink}{RGB}{5, 99, 193} % Word hyperlink blue (used for URLs, see Section 14)

\RequirePackage{graphicx}
%% ---- conference data (keys; defaults = 4CroCEE 2027) ----------------
%%   \croceeconference{edition=4, ordinal=th, name={Croatian Conference on
%%     Earthquake Engineering}, acronym=CroCEE, short=4CroCEE,
%%     dates={22--25 March 2027}, place={Opatija, Croatia}, year=2027,
%%     copyright=CroCEE, logo=Figures/4CroCEE_logo}
%% Definitions/conference.cfg (if present) is read after the defaults.
%% Text macros for the examples: \ConfShort, \ConfYear, \ConfName,
%% \ConfEdition ("4th"), \ConfFullName ("4th Croatian Conference ...")
\ExplSyntaxOn
\keys_define:nn { crocee/conf }
  {
    edition   .tl_set:N = \crocee@conf@edition,
    ordinal   .tl_set:N = \crocee@conf@ordinal,
    name      .tl_set:N = \crocee@conf@name,
    acronym   .tl_set:N = \crocee@conf@acronym,
    short     .tl_set:N = \crocee@conf@short,
    dates     .tl_set:N = \crocee@conf@dates,
    place     .tl_set:N = \crocee@conf@place,
    year      .tl_set:N = \crocee@conf@year,
    copyright .tl_set:N = \crocee@conf@copyright,
    logo      .tl_set:N = \crocee@logo,
  }
\NewDocumentCommand \croceeconference { m } { \keys_set:nn { crocee/conf } { #1 } }
\ExplSyntaxOff
\croceeconference{edition=4, ordinal=th,
  name={Croatian Conference on Earthquake Engineering}, acronym=CroCEE,
  short=4CroCEE, dates={22--25 March 2027}, place={Opatija, Croatia},
  year=2027, copyright=CroCEE, logo=Figures/4CroCEE_logo}
\InputIfFileExists{Definitions/conference.cfg}{}{}
\newcommand\ConfEdition{\crocee@conf@edition\textsuperscript{\crocee@conf@ordinal}}
\newcommand\ConfName{\crocee@conf@name}
\newcommand\ConfFullName{\ConfEdition\ \crocee@conf@name}
\newcommand\ConfShort{\crocee@conf@short}
\newcommand\ConfAcronym{\crocee@conf@acronym}
\newcommand\ConfYear{\crocee@conf@year}
\newcommand\ConfDates{\crocee@conf@dates}
\newcommand\ConfPlace{\crocee@conf@place}
%% header lines, built from the keys; \croceeheader{..}{..}{..} overrides
\let\crocee@hdr@custom\relax
\newcommand\croceeheader[3]{%
  \gdef\crocee@hdr@a{#1}\gdef\crocee@hdr@b{#2}\gdef\crocee@hdr@c{#3}%
  \let\crocee@hdr@custom\@empty}
\newcommand\crocee@sethdr{%
  \ifx\crocee@hdr@custom\relax
    \gdef\crocee@hdr@a{Proceedings of the \ConfFullName\ -- \ConfShort}%
    \gdef\crocee@hdr@b{\ConfDates, \ConfPlace}%
    \gdef\crocee@hdr@c{Copyright \textcopyright{} \ConfYear\ \crocee@conf@copyright}%
  \fi}
\newcommand\croceelogo[1]{\gdef\crocee@logo{#1}}
\newcommand\crocee@logobox{%
  \IfFileExists{\crocee@logo.png}{%
    \includegraphics[width=68.47bp,height=41.67bp]{\crocee@logo.png}%
  }{\IfFileExists{\crocee@logo.pdf}{%
      \includegraphics[width=68.47bp,height=41.67bp]{\crocee@logo.pdf}%
    }{%
      \ClassWarningNoLine{CroCEE}{Logo file \crocee@logo.png not found}%
      \fbox{\parbox[b][37pt][c]{64pt}{\centering\tiny CroCEE logo}}%
    }}}
\let\crocee@doi\@empty
\newcommand\paperdoi[1]{\gdef\crocee@doi{#1}}
\let\crocee@firstpage\@empty
\newcommand\firstpage[1]{\gdef\crocee@firstpage{#1}}
\AtBeginDocument{\ifx\crocee@firstpage\@empty\else
  \setcounter{page}{\crocee@firstpage}\fi}
%% header: 709 tw (35.45 bp) from the top edge; first baseline at the
%% line-box ascent of the header font, then one line pitch per line
\def\crocee@hdr@top{35.45}
\newcommand\crocee@hdrline[1]{%
  \fpeval{round(\crocee@hdr@top + \crocee@asc*\crocee@sz@affil
    + (#1)*\crocee@lh*\crocee@sz@affil, 2)}}
\newcommand\crocee@drawheader{%
  \begingroup
  \crocee@sethdr
  \setlength\unitlength{1bp}%
  \crocee@fs{\crocee@sz@affil}\itshape
  \put(75.43,-28.90){\raisebox{-41.67bp}{\crocee@logobox}}%
  \put(523.28,-\crocee@hdrline{0}){\llap{\crocee@hdr@a}}%
  \put(523.28,-\crocee@hdrline{1}){\llap{\crocee@hdr@b}}%
  \put(523.28,-\crocee@hdrline{2}){\llap{\crocee@hdr@c}}%
  \ifx\crocee@doi\@empty\else
    \put(523.28,-\crocee@hdrline{3}){\llap{DOI: \url{https://doi.org/\crocee@doi}}}%
  \fi
  \ifx\crocee@firstpage\@empty\else
    % footer: 709 tw (35.45 bp) from the bottom edge, right-aligned
    % with the text block
    \put(523.28,-806.44){\llap{\upshape\thepage}}%
  \fi
  \endgroup}
\AddToHook{shipout/background}{\crocee@drawheader}

%% =====================================================================
%%  8. TITLE, AUTHORS, AFFILIATIONS
%% =====================================================================
\let\crocee@authors\@empty
\let\crocee@affils\@empty
\let\crocee@pdfauthor\@empty      % plain list of names for the PDF metadata
\renewcommand\author[2][]{%
  \ifx\crocee@pdfauthor\@empty\else\g@addto@macro\crocee@pdfauthor{, }\fi
  \g@addto@macro\crocee@pdfauthor{#2}%
  \ifx\crocee@authors\@empty\else\g@addto@macro\crocee@authors{, }\fi
  \if\relax\detokenize{#1}\relax
    \g@addto@macro\crocee@authors{#2}%
  \else
    \g@addto@macro\crocee@authors{#2~\textsuperscript{(#1)}}%
  \fi}
\newcommand\affiliation[2][]{%
  \g@addto@macro\crocee@affils{\crocee@affil{#1}{#2}}}
\newcommand\crocee@affil[2]{%
  \par\noindent\hangindent\crocee@tw{\crocee@tw@affilhang}\hangafter1
  \if\relax\detokenize{#1}\relax\else\textsuperscript{(#1)} \fi#2\par}
\newcommand\crocee@titlefont{\crocee@fs{\crocee@sz@title}\bfseries}
%% The title is printed in capital letters (as required by the template), so
%% the author may type it in any case. Math is not changed; protect text that
%% must keep its case (units, symbols) with \NoCaseChange{...}, e.g.
%% \title{Piers loaded with 10 \NoCaseChange{kN}}. The PDF title keeps the
%% original case.
\newcommand\crocee@titletext{\MakeUppercase{\@title}}

\renewcommand\maketitle{\par
  \begingroup
    \parskip\z@ \parindent\z@ \lineskiplimit-\maxdimen
    \setbox\@tempboxa\vbox{\crocee@titlefont\centering
      \crocee@titletext\par\global\@tempcnta\prevgraf}%
    \ifnum\@tempcnta>2
      \ClassWarningNoLine{CroCEE}{The title has more than two lines}%
    \fi
    % Space before the title
    \vspace*{\dimexpr\crocee@title@before-\topskip\relax}%
    \prevdepth-1000pt
    {\crocee@titlefont\centering\crocee@strut\crocee@titletext\par}%
    % Space after title + empty spacer line
    \vskip\crocee@vgap{\crocee@sz@title}%
      {\dimexpr\crocee@title@after+\crocee@lh\dimexpr\crocee@sz@titlespacer bp\relax\relax}%
      {\crocee@sz@authors}%
    {\crocee@fs{\crocee@sz@authors}\centering\crocee@authors\par}%
    % Empty spacer line before affiliations
    \vskip\crocee@vgap{\crocee@sz@authors}%
      {\crocee@lh\dimexpr\crocee@sz@titlespacer bp\relax}%
      {\crocee@sz@affil}%
    {\crocee@fs{\crocee@sz@affil}\itshape\raggedright\crocee@affils}%
    \ifcrocee@extended
      % Two empty affiliation-sized lines; keywords follow directly
      \vskip\dimexpr\crocee@vgap{\crocee@sz@affil}%
        {2\dimexpr\crocee@lh\dimexpr\crocee@sz@affil bp\relax\relax}%
        {\crocee@sz@keywords}-\parskip\relax
    \else
      % Two empty affiliation-sized lines + space before Abstract heading
      \vskip\dimexpr\crocee@vgap{\crocee@sz@affil}%
        {\dimexpr2\dimexpr\crocee@lh\dimexpr\crocee@sz@affil bp\relax\relax+\crocee@abs@before\relax}%
        {\crocee@sz@abshead}-\parskip\relax
    \fi
  \endgroup
  \global\let\crocee@authors\@empty
  \global\let\crocee@affils\@empty
  \global\let\@title\@empty
  \global\let\title\relax
  \global\let\maketitle\relax}
\newcommand\crocee@pdftitle{}
\renewcommand\title[1]{\gdef\@title{#1}\gdef\crocee@pdftitle{#1}}

%% =====================================================================
%%  9. ABSTRACT AND KEYWORDS
%% =====================================================================
\renewcommand\abstractname{Abstract}
\renewenvironment{abstract}{%
  \par
  {\noindent\crocee@fs{\crocee@sz@abshead}\bfseries\itshape\abstractname\par}%
  \crocee@fs[\crocee@mult@abstract]{\crocee@sz@abstext}%
  \vskip\crocee@vgap{\crocee@sz@abshead}{0bp}{\crocee@sz@abstext}%
  \noindent\ignorespaces
}{\par}

\newcommand\keywordsname{Keywords}
\newcommand\keywords[1]{%
  \par{\noindent\crocee@fs{\crocee@sz@keywords}\itshape\keywordsname: #1\par}%
  \global\crocee@secbeforeskip\dimexpr\crocee@secbefore-\crocee@xb\relax
  \vskip\crocee@vgap{\crocee@sz@keywords}{0bp}{\crocee@sz@normalsize}}

%% =====================================================================
%%  10. SECTIONING
%% =====================================================================
\newcommand\crocee@tabto[3]{%
  \setbox\z@\hbox{#1}%
  \ifdim\wd\z@<#2\relax
    \hb@xt@#2{\unhbox\z@\hfil}%
  \else\ifdim\wd\z@<#3\relax
    \hb@xt@#3{\unhbox\z@\hfil}%
  \else
    \count@\wd\z@ \divide\count@ \number\crocee@tw{\crocee@tw@tabgrid}\relax
    \advance\count@\@ne
    \hb@xt@\count@\crocee@tw{\crocee@tw@tabgrid}{\unhbox\z@\hfil}%
  \fi\fi}
\newcommand\crocee@sechead[1]{%
  \hangindent\crocee@tw{\crocee@tw@secstopB}\hangafter\@ne
  \crocee@tabto{\thesection.}{\crocee@tw{\crocee@tw@secstopA}}{\crocee@tw{\crocee@tw@secstopB}}#1}
\newcommand\crocee@subsechead[1]{%
  \hangindent\crocee@tw{\crocee@tw@subsecstopB}\hangafter\@ne
  \crocee@tabto{\thesubsection.}{\crocee@tw{\crocee@tw@secstopA}}{\crocee@tw{\crocee@tw@subsecstopB}}\ #1}
\newcommand\crocee@subsubsechead[1]{%
  \hangindent\crocee@tw{\crocee@tw@subsubsecstopA}\hangafter\@ne
  \crocee@tabto{\thesubsubsection.}{\crocee@tw{\crocee@tw@subsubsecstopA}}{\crocee@tw{\crocee@tw@tabgrid}}#1}

\RequirePackage{titlesec}
\setcounter{secnumdepth}{3}
\titleformat{\section}[block]
  {\crocee@fs{\crocee@sz@sec}\bfseries\raggedright}{}{0pt}{\crocee@sechead}
\titleformat{name=\section,numberless}[block]
  {\crocee@fs{\crocee@sz@sec}\bfseries\raggedright}{}{0pt}{}
\titleformat{\subsection}[block]
  {\crocee@fs{\crocee@sz@subsec}\bfseries\raggedright}{}{0pt}{\crocee@subsechead}
\titleformat{name=\subsection,numberless}[block]
  {\crocee@fs{\crocee@sz@subsec}\bfseries\raggedright}{}{0pt}{}
\titleformat{\subsubsection}[block]
  {\crocee@fs{\crocee@sz@subsubsec}\bfseries\raggedright}{}{0pt}{\crocee@subsubsechead}
\titleformat{name=\subsubsection,numberless}[block]
  {\crocee@fs{\crocee@sz@subsubsec}\bfseries\raggedright}{}{0pt}{}

%% One-shot override: \keywords (Section 9) and the float hooks (Section 11)
%% set \crocee@secbeforeskip for the NEXT \section only; it is restored at the
%% start of every following paragraph.
\newcommand\crocee@kwreset{\global\crocee@secbeforeskip\crocee@secbefore}
\AddToHook{para/begin}{\crocee@kwreset}

\def\ttl@strut{\crocee@strut}
\titlespacing*{\section}{0pt}{\crocee@secbeforeskip}{\crocee@secafter}
\titlespacing*{\subsection}{0pt}{\crocee@subbefore}{\crocee@subafter}
\titlespacing*{\subsubsection}{0pt}{\crocee@subbefore}{\crocee@subafter}
\titleformat{\paragraph}[runin]
  {\crocee@fs{\crocee@sz@subsec}\bfseries}{}{0pt}{}
\titlespacing*{\paragraph}{0pt}{\crocee@subbefore}{1em}

%% =====================================================================
%%  11. FLOATS, CAPTIONS AND TABLES
%% =====================================================================
\RequirePackage{caption}
\DeclareCaptionFont{croceestrut}{%
  \crocee@fs[\crocee@mult@caption]{\crocee@sz@caption}%
  \setbox\strutbox\hbox{\crocee@strut}}
\captionsetup{format=plain, font={croceestrut}, labelfont=bf, %{}
  textfont={}, labelsep=period, justification=centering,
  singlelinecheck=false, strut=on}
\captionsetup[table]{position=above, skip=\crocee@cap@tableskip}
\captionsetup[figure]{position=below, skip=\crocee@cap@figureskip}
\renewcommand\figurename{Figure}
\renewcommand\tablename{Table}
\def\fps@figure{!htbp}
\def\fps@table{!htbp}

%% After a float: \crocee@floatend removes the 0.285 bp font-size difference
%% (see \crocee@floatcomp), and \crocee@floatafter shortens the space before
%% the next \section by what has already been inserted: the extra line-box
%% term \crocee@floatcomp, the float gap below the float (\textfloatsep =
%% \intextsep) and \parskip, which \titlespacing adds again. The remaining
%% \crocee@float@trim is empirical (Section 1.6).
\newcommand\crocee@floatend{\par\vskip-\crocee@floatcomp}
\newcommand\crocee@floatafter{%
  \global\crocee@secbeforeskip\dimexpr\crocee@sec@before+\crocee@floatcomp-\textfloatsep-\parskip-\crocee@float@trim\relax}
\AddToHook{env/table/end}{\crocee@floatend}
\AddToHook{env/figure/end}{\crocee@floatend}
\AddToHook{env/table/after}{\crocee@floatafter}
\AddToHook{env/figure/after}{\crocee@floatafter}

\renewcommand\topfraction{.9}
\renewcommand\bottomfraction{.9}
\renewcommand\textfraction{.07}
\renewcommand\floatpagefraction{.8}

\RequirePackage{array}
\RequirePackage{xcolor}
\RequirePackage{booktabs}
\RequirePackage{tabularx}
\RequirePackage{changepage}

\newlength\extralength
\setlength\extralength{0cm}
%% \fulllength is a macro (not a length register), so it follows a change of
%% \extralength made anywhere in the document, e.g. before a wide table. The
%% table is extended to the left only (adjustwidth{-\extralength}{0cm}).
\newcommand\fulllength{\dimexpr\textwidth+\extralength\relax}
\newcolumntype{C}{>{\centering\arraybackslash}X}
\setlength\heavyrulewidth{1.0bp}
\setlength\lightrulewidth{0.5bp}
\setlength\aboverulesep{2pt}
\setlength\belowrulesep{2pt}
\newcommand\tablenotes[1]{\par\vskip\crocee@table@noteskip\noindent{\footnotesize #1\par}}

\def\crocee@tabletype{table}
\newcommand\crocee@floatsetup{%
  \ifx\@captype\crocee@tabletype
    \crocee@fs[\crocee@mult@table]{\crocee@sz@table}%
    \setlength\tabcolsep{\crocee@table@colsep}%
    \setbox\strutbox\hbox{\crocee@strut}%
  \fi
  \lineskip\z@}
\g@addto@macro\@floatboxreset{\crocee@floatsetup}

%% =====================================================================
%%  12. EQUATIONS AND CROSS-REFERENCES
%% =====================================================================
\newcommand\vect[1]{\bm{#1}}
\newcommand\matr[1]{\bm{#1}}
\newcommand\figref[1]{Fig.~\ref{#1}}
\newcommand\tabref[1]{Table~\ref{#1}}
\newcommand\Eqref[1]{Eq.~\eqref{#1}}

\newcommand\crocee@afterdisplay{%
  \ifhmode\ifx\par\crocee@displaypar\else
    \let\crocee@savedpar\par
    \let\par\crocee@displaypar
  \fi\fi}
\newcommand\crocee@displaypar{%
  \let\par\crocee@savedpar\crocee@savedpar\crocee@displayfix}
\newcommand\crocee@displayfix{%
  \ifvmode\ifdim\lastskip>\z@
    \vskip-\lastskip\vskip\dimexpr\crocee@eqskip-\parskip\relax
  \fi\fi}
\@for\crocee@env:={equation,equation*,align,align*,gather,gather*,%
    multline,multline*,flalign,flalign*,alignat,alignat*,displaymath}\do{%
  \edef\crocee@tmp{\noexpand\AddToHook{env/\crocee@env/after}}%
  \crocee@tmp{\crocee@afterdisplay}}

%% =====================================================================
%%  13. REFERENCES (Elsevier numeric with titles)
%% =====================================================================
\RequirePackage[nospace]{cite}
\renewcommand\refname{References}
\let\crocee@bibpreamble\@empty
\newcommand\setbibpreamble[1]{\gdef\crocee@bibpreamble{#1\par}}
\renewcommand\@biblabel[1]{[#1]}
\renewenvironment{thebibliography}[1]{%
  \section*{\refname}%
  \ifx\crocee@bibpreamble\@empty
    \kern\crocee@vgap{\crocee@sz@normalsize}{0bp}{\crocee@sz@bib}%
  \else
    \crocee@bibpreamble
  \fi
  \crocee@fs[\crocee@mult@bib]{\crocee@sz@bib}%
  \list{\@biblabel{\@arabic\c@enumiv}}{%
    \setlength\labelwidth{\crocee@tw{\crocee@tw@bibindent}}%
    \setlength\leftmargin{\crocee@tw{\crocee@tw@bibindent}}%
    \setlength\labelsep{\z@}%
    \setlength\itemindent{\z@}%
    \setlength\listparindent{\z@}%
    \setlength\itemsep{\crocee@bib@itemsep}%
    \setlength\parsep{\z@}%
    \setlength\topsep{\dimexpr\crocee@vgap{\crocee@sz@normalsize}{0bp}{\crocee@sz@bib}+\crocee@xb\relax}%
    \setlength\partopsep{\z@}%
    \def\makelabel##1{%
      \crocee@tabto{##1}{\crocee@tw{\crocee@tw@bibindent}}{\crocee@tw{\crocee@tw@secstopB}}}%
    \usecounter{enumiv}%
    \let\p@enumiv\@empty
    \renewcommand\theenumiv{\@arabic\c@enumiv}}%
  \sloppy
  \clubpenalty4000 \@clubpenalty\clubpenalty \widowpenalty4000
  \sfcode`\.\@m}
 {\def\@noitemerr{\@latex@warning{Empty `thebibliography' environment}}%
  \endlist}

%% Automatically locate .bst in \crocee@defdir (Definitions/) or root folder,
%% and apply default \crocee@bst if the author omits \bibliographystyle
\newif\ifcrocee@bstset \crocee@bstsetfalse
\let\crocee@orig@bibliographystyle\bibliographystyle
\renewcommand\bibliographystyle[1]{%
  \crocee@bstsettrue
  \IfFileExists{\crocee@defdir#1.bst}{%
    \crocee@orig@bibliographystyle{\crocee@defdir#1}%
  }{%
    \crocee@orig@bibliographystyle{#1}%
  }%
}

\AtBeginDocument{%
  \let\crocee@bibliography\bibliography
  \renewcommand\bibliography[1]{%
    \ifcrocee@bstset\else
      \bibliographystyle{\crocee@bst}%
    \fi
    \IfFileExists{\jobname.bbl}{}{%
      \section*{\refname}\crocee@bibpreamble}%
    \crocee@bibliography{#1}}}

%% =====================================================================
%%  14. URLs AND HYPERREF
%% =====================================================================
\RequirePackage{url}
\urlstyle{same}
\mathchardef\UrlBreakPenalty=9000
\mathchardef\UrlBigBreakPenalty=9000
\ifcrocee@ulurl
  \newcommand\crocee@ulbox[1]{\hbox{\setbox\z@\hbox{\char#1}%
    \rlap{\vrule\@width\wd\z@\@height-.1089em\@depth.1577em}\box\z@}}
  \def\crocee@ulatom#1#2{%
    \ifcase#1\relax\mathord\or\mathord\or\mathbin\or\mathrel\or
      \mathopen\or\mathclose\or\mathpunct\else\mathord\fi
    {\crocee@ulbox{#2}}}
  \def\crocee@uldef#1#2{%
    \begingroup\lccode`\~=#1\relax
    \lowercase{\endgroup\edef~}{\noexpand\crocee@ulatom{#2}{#1}}%
    \mathcode#1="8000\relax}
  \def\crocee@ulall{%
    \count@=33
    \loop
      \ifnum\mathcode\count@<32768
        \@tempcnta\mathcode\count@ \divide\@tempcnta 4096
        \crocee@uldef{\the\count@}{\the\@tempcnta}%
      \fi
      \advance\count@\@ne
    \ifnum\count@<127 \repeat
    \def\Url@percent{\crocee@ulatom{2}{37}}}
  \let\crocee@Url@MathSetup\Url@MathSetup
  \def\Url@MathSetup{%
    \let\crocee@nolig\verbatim@nolig@list
    \let\verbatim@nolig@list\@empty
    \crocee@Url@MathSetup
    \let\verbatim@nolig@list\crocee@nolig}
  \def\UrlLeft{\crocee@ulall}
\fi

%% hyperref is loaded at \begin{document} unless the author has already loaded
%% it (with his or her own options). Links are not coloured or framed
%% (hidelinks) when the class loads it. The PDF metadata (title, author,
%% subject, language) are filled in from \title and \author; to override them
%% use \hypersetup after \begin{document}.
%% Additional packages: Definitions/packages.tex is read here if it exists
%% (e.g. tikz, pgfplots) -- before hyperref, which the class loads last.
\InputIfFileExists{Definitions/packages.tex}{}{}

\AddToHook{begindocument/before}{%
  \@ifpackageloaded{hyperref}{}{%
    \RequirePackage{hyperref}%
    % URLs (including the DOI line of the header) in Word's hyperlink blue;
    % internal links and citations black, no frames
    \hypersetup{colorlinks=true, linkcolor=black, citecolor=black,
      filecolor=black, urlcolor=croceelink}}%
  \hypersetup{pdfstartview=FitH,
    pdftitle={\crocee@pdftitle},
    pdfauthor={\crocee@pdfauthor},
    pdfsubject={\crocee@conf@edition\crocee@conf@ordinal\space\crocee@conf@name\space-\space\crocee@conf@short},
    pdflang={en-GB}}%
  \urlstyle{same}}

%% =====================================================================
%%  15. PAGE LENGTH CHECK
%% =====================================================================
\AddToHook{enddocument/afterlastpage}{%
  \@tempcnta\numexpr\c@page-1\relax
  \ifx\crocee@firstpage\@empty\else
    \advance\@tempcnta-\crocee@firstpage\relax \advance\@tempcnta\@ne
  \fi
  \ifcrocee@extended
    \ifnum\@tempcnta>\crocee@pages@extmax\relax
      \ClassWarningNoLine{CroCEE}{The extended abstract has
        \the\@tempcnta\space pages; the maximum is \crocee@pages@extmax}%
    \fi
  \else
    \ifnum\@tempcnta>\crocee@pages@fullmax\relax
      \ClassWarningNoLine{CroCEE}{The paper has \the\@tempcnta\space pages;
        the maximum is \crocee@pages@fullmax}%
    \fi
    \ifnum\@tempcnta<\crocee@pages@fullmin\relax
      \ClassWarningNoLine{CroCEE}{The paper has \the\@tempcnta\space pages;
        the minimum is \crocee@pages@fullmin}%
    \fi
  \fi}

\endinput