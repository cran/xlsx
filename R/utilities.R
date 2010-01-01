.onLoad <- function(libname, pkgname)
{
  #.jpackage(pkgname)  # not necessary anymore
  
  # what's your java  version?  Need > 1.5.0.
##   jversion <- .jcall('java.lang.System','S','getProperty','java.version')
##   if (jversion <= "1.5.0")
##     stop(paste("Your java version is ", jversion,
##                  ".  Need 1.5.0 or higher.", sep=""))

  #wb <- createWorkbook()   # load/initialize jars here as it takes 
  #rm(wb)                   # a few seconds ...
}

########################################################################
# take an R color and return an XSSFColor object
# where Rcolor is a string from "colors()"
.xssfcolor <- function(Rcolor)
{
  rgb  <- as.integer(col2rgb(Rcolor))
  jcol <- .jnew("java.awt.Color", rgb[1], rgb[2], rgb[3])
  .jnew("org.apache.poi.xssf.usermodel.XSSFColor", jcol)
}


  .CELL_STYLES <- c(2,6,4,0,5,1,3,3,9,9,11,3,7,6,4,2,10,12,8,0,13,5,1,
    10,16,2,18,17,0,1,4,15,7,8,5,6,13,14,11,12,2,1,3,0)
  names(.CELL_STYLES) <- c('ALIGN_CENTER' ,'ALIGN_CENTER_SELECTION'
    ,'ALIGN_FILL' ,'ALIGN_GENERAL' ,'ALIGN_JUSTIFY' ,'ALIGN_LEFT'
    ,'ALIGN_RIGHT' ,'ALT_BARS' ,'BIG_SPOTS' ,'BORDER_DASH_DOT'
    ,'BORDER_DASH_DOT_DOT1' ,'BORDER_DASHED' ,'BORDER_DOTTED'
    ,'BORDER_DOUBLE' ,'BORDER_HAIR' ,'BORDER_MEDIUM'
    ,'BORDER_MEDIUM_DASH_DOT1' ,'BORDER_MEDIUM_DASH_DOT_DOT1'
    ,'BORDER_MEDIUM_DASHED' ,'BORDER_NONE' ,'BORDER_SLANTED_DASH_DOT1'
    ,'BORDER_THICK' ,'BORDER_THIN' ,'BRICKS1' ,'DIAMONDS1'
    ,'FINE_DOTS' ,'LEAST_DOTS1' ,'LESS_DOTS1' ,'NO_FILL'
    ,'SOLID_FOREGROUND' ,'SPARSE_DOTS' ,'SQUARES1'
    ,'THICK_BACKWARD_DIAG' ,'THICK_FORWARD_DIAG' ,'THICK_HORZ_BANDS'
    ,'THICK_VERT_BANDS' ,'THIN_BACKWARD_DIAG1' ,'THIN_FORWARD_DIAG1'
    ,'THIN_HORZ_BANDS1' ,'THIN_VERT_BANDS1' ,'VERTICAL_BOTTOM'
    ,'VERTICAL_CENTER' ,'VERTICAL_JUSTIFY' ,'VERTICAL_TOP')

  .INDEXED_COLORS <- c(8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,
    25,26,28,29,30,31,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,
    56,57,58,59,60,61,62,63,64)
  names(.INDEXED_COLORS) <- c('BLACK' ,'WHITE' ,'RED' ,'BRIGHT_GREEN'
    ,'BLUE' ,'YELLOW' ,'PINK' ,'TURQUOISE' ,'DARK_RED' ,'GREEN'
    ,'DARK_BLUE' ,'DARK_YELLOW' ,'VIOLET' ,'TEAL' ,'GREY_25_PERCENT'
    ,'GREY_50_PERCENT' ,'CORNFLOWER_BLUE' ,'MAROON' ,'LEMON_CHIFFON'
    ,'ORCHID' ,'CORAL' ,'ROYAL_BLUE' ,'LIGHT_CORNFLOWER_BLUE'
    ,'SKY_BLUE' ,'LIGHT_TURQUOISE' ,'LIGHT_GREEN' ,'LIGHT_YELLOW'
    ,'PALE_BLUE' ,'ROSE' ,'LAVENDER' ,'TAN' ,'LIGHT_BLUE' ,'AQUA'
    ,'LIME' ,'GOLD' ,'LIGHT_ORANGE' ,'ORANGE' ,'BLUE_GREY'
    ,'GREY_40_PERCENT' ,'DARK_TEAL' ,'SEA_GREEN' ,'DARK_GREEN'
    ,'OLIVE_GREEN' ,'BROWN' ,'PLUM' ,'INDIGO' ,'GREY_80_PERCENT'
    ,'AUTOMATIC')
  
                   





## aux <- "
## 	ALIGN_CENTER	2
## 	ALIGN_CENTER_SELECTION	6
## 	ALIGN_FILL	4
## 	ALIGN_GENERAL	0
## 	ALIGN_JUSTIFY	5
## 	ALIGN_LEFT	1
## 	ALIGN_RIGHT	3
## 	ALT_BARS	3
## 	BIG_SPOTS	9
## 	BORDER_DASH_DOT	9
## 	BORDER_DASH_DOT_DOT	11
## 	BORDER_DASHED	3
## 	BORDER_DOTTED	7
## 	BORDER_DOUBLE	6
## 	BORDER_HAIR	4
## 	BORDER_MEDIUM	2
## 	BORDER_MEDIUM_DASH_DOT	10
## 	BORDER_MEDIUM_DASH_DOT_DOT	12
## 	BORDER_MEDIUM_DASHED	8
## 	BORDER_NONE	0
## 	BORDER_SLANTED_DASH_DOT	13
## 	BORDER_THICK	5
## 	BORDER_THIN	1
## 	BRICKS	10
## 	DIAMONDS	16
## 	FINE_DOTS	2
## 	LEAST_DOTS	18
## 	LESS_DOTS	17
## 	NO_FILL	0
## 	SOLID_FOREGROUND	1
## 	SPARSE_DOTS	4
## 	SQUARES	15
## 	THICK_BACKWARD_DIAG	7
## 	THICK_FORWARD_DIAG	8
## 	THICK_HORZ_BANDS	5
## 	THICK_VERT_BANDS	6
## 	THIN_BACKWARD_DIAG	13
## 	THIN_FORWARD_DIAG	14
## 	THIN_HORZ_BANDS	11
## 	THIN_VERT_BANDS	12
## 	VERTICAL_BOTTOM	2
## 	VERTICAL_CENTER	1
## 	VERTICAL_JUSTIFY	3
## 	VERTICAL_TOP	0"
  
## bux <- strsplit(aux, "\n")[[1]][-1]
## fields <- paste(gsub("(.*)[[:digit:]]+", "\\1", bux), sep="",
##   collapse="' ,'")
## values <- paste(as.numeric(gsub("[[:upper:]]|_", "", bux)), sep="",
##                 collapse=",")

## gsub(".*([[:digit:]]+)$", "\\1", bux, perl=TRUE)



##   aux <- 
##    "BLACK(8),
##     WHITE(9),
##     RED(10),
##     BRIGHT_GREEN(11),
##     BLUE(12),
##     YELLOW(13),
##     PINK(14),
##     TURQUOISE(15),
##     DARK_RED(16),
##     GREEN(17),
##     DARK_BLUE(18),
##     DARK_YELLOW(19),
##     VIOLET(20),
##     TEAL(21),
##     GREY_25_PERCENT(22),
##     GREY_50_PERCENT(23),
##     CORNFLOWER_BLUE(24),
##     MAROON(25),
##     LEMON_CHIFFON(26),
##     ORCHID(28),
##     CORAL(29),
##     ROYAL_BLUE(30),
##     LIGHT_CORNFLOWER_BLUE(31),
##     SKY_BLUE(40),
##     LIGHT_TURQUOISE(41),
##     LIGHT_GREEN(42),
##     LIGHT_YELLOW(43),
##     PALE_BLUE(44),
##     ROSE(45),
##     LAVENDER(46),
##     TAN(47),
##     LIGHT_BLUE(48),
##     AQUA(49),
##     LIME(50),
##     GOLD(51),
##     LIGHT_ORANGE(52),
##     ORANGE(53),
##     BLUE_GREY(54),
##     GREY_40_PERCENT(55),
##     DARK_TEAL(56),
##     SEA_GREEN(57),
##     DARK_GREEN(58),
##     OLIVE_GREEN(59),
##     BROWN(60),
##     PLUM(61),
##     INDIGO(62),
##     GREY_80_PERCENT(63),
##     AUTOMATIC(64)"
## bux <- gsub("^ *", "", strsplit(aux, ",\n")[[1]])
## fields <- paste(gsub("(.*)\\([[:digit:]]+\\)", "\\1", bux), sep="",
##   collapse="' ,'")
## values <- paste(as.numeric(gsub(".*\\(([[:digit:]]+)\\)", "\\1", bux)),
##   sep="", collapse=",")


## aux <- "
## public static final byte	ANSI_CHARSET	0
## public static final short	BOLDWEIGHT_BOLD	700
## public static final short	BOLDWEIGHT_NORMAL	400
## public static final short	COLOR_NORMAL	32767
## public static final short	COLOR_RED	10
## public static final byte	DEFAULT_CHARSET	1
## public static final short	SS_NONE	0
## public static final short	SS_SUB	2
## public static final short	SS_SUPER	1
## public static final byte	SYMBOL_CHARSET	2
## public static final byte	U_DOUBLE	2
## public static final byte	U_DOUBLE_ACCOUNTING	34
## public static final byte	U_NONE	0
## public static final byte	U_SINGLE	1
## public static final byte	U_SINGLE_ACCOUNTING	33"
## bux <- strsplit(aux, " ")[[1]])
## fields <- paste(gsub("(.*)\\([[:digit:]]+\\)", "\\1", bux), sep="",
##   collapse="' ,'")
## values <- paste(as.numeric(gsub(".*\\(([[:digit:]]+)\\)", "\\1", bux)),
##   sep="", collapse=",")

