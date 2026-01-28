# Excel to Python Implementation Guide

**Source File:** `content-calendar.xlsx`

**Generated:** analyze_excel_structure.py


---

## Overview

- **Total Sheets:** 6

- **Total Functions Used:** 15

- **Total Number Formats:** 6


### Sheet List

- **Content**: 27 rows × 90 columns
  - Formulas: 1653

- **Calendar**: 96 rows × 17 columns
  - Formulas: 941

- **Hashtags**: 29 rows × 9 columns
  - Formulas: 23

- **Settings**: 94 rows × 5 columns
  - Formulas: 44

- **Help**: 46 rows × 4 columns
  - Formulas: 5

- **©**: 19 rows × 3 columns


---

## Sheet: Content

- **Dimensions:** 27 rows × 90 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 1.20 |

| B | 1.60 |

| BC | 2.50 |

| C | 8.00 |

| D | 23.10 |

| E | 8.70 |

| F | 30.00 |

| G | 6.50 |

| H | 9.00 |

| I | 10.50 |

| J | 10.50 |

| L | 8.10 |

| M | 2.50 |



### Merged Cells

- `BJ2:BP2`

- `CE2:CK2`

- `BC2:BI2`

- `BQ2:BW2`

- `B24:F24`

- `M2:S2`

- `AH2:AN2`

- `T2:Z2`

- `AO2:AU2`

- `AA2:AG2`

- `AV2:BB2`

- `BX2:CD2`



### Formulas

**Total:** 1653


| Cell | Formula |
|------|---------|

| M2 | `=M3` |

| T2 | `=T3` |

| AA2 | `=AA3` |

| AH2 | `=AH3` |

| AO2 | `=AO3` |

| AV2 | `=AV3` |

| BC2 | `=BC3` |

| BJ2 | `=BJ3` |

| BQ2 | `=BQ3` |

| BX2 | `=BX3` |

| CE2 | `=CE3` |

| M3 | `=J2-WEEKDAY(J2,1)+2+7*(J3-1)` |

| N3 | `=M3+1` |

| O3 | `=N3+1` |

| P3 | `=O3+1` |

| Q3 | `=P3+1` |

| R3 | `=Q3+1` |

| S3 | `=R3+1` |

| T3 | `=S3+1` |

| U3 | `=T3+1` |

| V3 | `=U3+1` |

| W3 | `=V3+1` |

| X3 | `=W3+1` |

| Y3 | `=X3+1` |

| Z3 | `=Y3+1` |

| AA3 | `=Z3+1` |

| AB3 | `=AA3+1` |

| AC3 | `=AB3+1` |

| AD3 | `=AC3+1` |

| AE3 | `=AD3+1` |

| AF3 | `=AE3+1` |

| AG3 | `=AF3+1` |

| AH3 | `=AG3+1` |

| AI3 | `=AH3+1` |

| AJ3 | `=AI3+1` |

| AK3 | `=AJ3+1` |

| AL3 | `=AK3+1` |

| AM3 | `=AL3+1` |

| AN3 | `=AM3+1` |

| AO3 | `=AN3+1` |

| AP3 | `=AO3+1` |

| AQ3 | `=AP3+1` |

| AR3 | `=AQ3+1` |

| AS3 | `=AR3+1` |

| AT3 | `=AS3+1` |

| AU3 | `=AT3+1` |

| AV3 | `=AU3+1` |

| AW3 | `=AV3+1` |

| AX3 | `=AW3+1` |

| AY3 | `=AX3+1` |

| ... | ... (1603 more) |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | None | n |  | Bold:False |  | None |

| B1 | Content Calendar | s |  | Bold:False |  | None |

| C1 | None | n |  | Bold:False |  | None |

| D1 | None | n |  | Bold:False |  | None |

| E1 | None | n |  | Bold:False |  | None |

| F1 | None | n |  | Bold:False |  | None |

| G1 | None | n |  | Bold:False |  | right |

| H1 | None | n |  | Bold:False |  | right |

| I1 | None | n |  | Bold:False |  | None |

| J1 | None | n |  | Bold:False |  | None |



---

## Sheet: Calendar

- **Dimensions:** 96 rows × 17 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 4.20 |

| B | 13.00 |

| C | 4.20 |

| D | 13.00 |

| E | 4.20 |

| F | 13.00 |

| G | 4.20 |

| H | 13.00 |

| I | 4.20 |

| J | 13.00 |

| K | 4.20 |

| L | 13.00 |

| M | 4.20 |

| N | 13.00 |

| O | 6.70 |

| P | 15.00 |

| Q | 9.60 |

| R | 9.00 |



### Merged Cells

- `I1:N1`

- `M3:N3`

- `C3:D3`

- `A3:B3`

- `G3:H3`

- `E3:F3`

- `K3:L3`

- `I3:J3`



### Formulas

**Total:** 941


| Cell | Formula |
|------|---------|

| A1 | `=UPPER(TEXT(B2,"mmmm yyyy"))` |

| B2 | `=DATE(P6,Q8,1)` |

| A3 | `=A15` |

| C3 | `=C15` |

| E3 | `=E15` |

| G3 | `=G15` |

| I3 | `=I15` |

| K3 | `=K15` |

| M3 | `=M15` |

| A4 | `=DATE(P6,Q8,1)-(WEEKDAY(DATE(P6,Q8,1),1)-(P10-1))-IF((WEEKDAY(DATE(P6,Q8,1),1)-(P10-1))<=0,7,0)+1` |

| B4 | `=IFERROR(VLOOKUP(A4,Settings!$A$45:$C$94,3,FALSE),"")` |

| C4 | `=A4+1` |

| D4 | `=IFERROR(VLOOKUP(C4,Settings!$A$45:$C$94,3,FALSE),"")` |

| E4 | `=C4+1` |

| F4 | `=IFERROR(VLOOKUP(E4,Settings!$A$45:$C$94,3,FALSE),"")` |

| G4 | `=E4+1` |

| H4 | `=IFERROR(VLOOKUP(G4,Settings!$A$45:$C$94,3,FALSE),"")` |

| I4 | `=G4+1` |

| J4 | `=IFERROR(VLOOKUP(I4,Settings!$A$45:$C$94,3,FALSE),"")` |

| K4 | `=I4+1` |

| L4 | `=IFERROR(VLOOKUP(K4,Settings!$A$45:$C$94,3,FALSE),"")` |

| M4 | `=K4+1` |

| N4 | `=IFERROR(VLOOKUP(M4,Settings!$A$45:$C$94,3,FALSE),"")` |

| P4 | `=HYPERLINK("https://www.vertex42.com/calendars/content-calendar.html","Content Calendar Template")` |

| A5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(A4=Content!$J$4:$J$23) + 1*((1*(A4>=Content!$J$4:$J$2` |

| B5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(A4=Content!$J$4:$J$23) + 1*((1*(A4>=Content!$J$4:$J$2` |

| C5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(C4=Content!$J$4:$J$23) + 1*((1*(C4>=Content!$J$4:$J$2` |

| D5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(C4=Content!$J$4:$J$23) + 1*((1*(C4>=Content!$J$4:$J$2` |

| E5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(E4=Content!$J$4:$J$23) + 1*((1*(E4>=Content!$J$4:$J$2` |

| F5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(E4=Content!$J$4:$J$23) + 1*((1*(E4>=Content!$J$4:$J$2` |

| G5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(G4=Content!$J$4:$J$23) + 1*((1*(G4>=Content!$J$4:$J$2` |

| H5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(G4=Content!$J$4:$J$23) + 1*((1*(G4>=Content!$J$4:$J$2` |

| I5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(I4=Content!$J$4:$J$23) + 1*((1*(I4>=Content!$J$4:$J$2` |

| J5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(I4=Content!$J$4:$J$23) + 1*((1*(I4>=Content!$J$4:$J$2` |

| K5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(K4=Content!$J$4:$J$23) + 1*((1*(K4>=Content!$J$4:$J$2` |

| L5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(K4=Content!$J$4:$J$23) + 1*((1*(K4>=Content!$J$4:$J$2` |

| M5 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(M4=Content!$J$4:$J$23) + 1*((1*(M4>=Content!$J$4:$J$2` |

| N5 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(M4=Content!$J$4:$J$23) + 1*((1*(M4>=Content!$J$4:$J$2` |

| A6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(A4=Content!$J$4:$J$23) + 1*((1*(A4>=Content!$J$4:$J$2` |

| B6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(A4=Content!$J$4:$J$23) + 1*((1*(A4>=Content!$J$4:$J$2` |

| C6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(C4=Content!$J$4:$J$23) + 1*((1*(C4>=Content!$J$4:$J$2` |

| D6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(C4=Content!$J$4:$J$23) + 1*((1*(C4>=Content!$J$4:$J$2` |

| E6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(E4=Content!$J$4:$J$23) + 1*((1*(E4>=Content!$J$4:$J$2` |

| F6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(E4=Content!$J$4:$J$23) + 1*((1*(E4>=Content!$J$4:$J$2` |

| G6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(G4=Content!$J$4:$J$23) + 1*((1*(G4>=Content!$J$4:$J$2` |

| H6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(G4=Content!$J$4:$J$23) + 1*((1*(G4>=Content!$J$4:$J$2` |

| I6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(I4=Content!$J$4:$J$23) + 1*((1*(I4>=Content!$J$4:$J$2` |

| J6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(I4=Content!$J$4:$J$23) + 1*((1*(I4>=Content!$J$4:$J$2` |

| K6 | `=IFERROR(INDEX(Content!$C$4:$C$23,SMALL(IF((1*(K4=Content!$J$4:$J$23) + 1*((1*(K4>=Content!$J$4:$J$2` |

| L6 | `=IFERROR(INDEX(Content!$D$4:$D$23,SMALL(IF((1*(K4=Content!$J$4:$J$23) + 1*((1*(K4>=Content!$J$4:$J$2` |

| ... | ... (891 more) |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | =UPPER(TEXT(B2,"mmmm yyyy")) | f | =UPPER(TEXT(B2,"mmmm yyyy")) | Bold:False |  | None |

| B1 | None | n |  | Bold:False |  | None |

| C1 | None | n |  | Bold:False |  | None |

| D1 | None | n |  | Bold:False |  | None |

| E1 | None | n |  | Bold:False |  | None |

| F1 | None | n |  | Bold:False |  | None |

| G1 | None | n |  | Bold:False |  | None |

| H1 | None | n |  | Bold:False |  | None |

| I1 | Content Calendar | s |  | Bold:False |  | right |

| J1 | None | n |  | Bold:False |  | None |



---

## Sheet: Hashtags

- **Dimensions:** 29 rows × 9 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 1.50 |

| B | 23.20 |

| C | 9.10 |

| D | 8.20 |

| E | 16.40 |

| F | 13.00 |

| G | 10.00 |

| H | 3.40 |

| I | 9.00 |



### Formulas

**Total:** 23


| Cell | Formula |
|------|---------|

| G1 | `=HYPERLINK("https://www.instagram.com/vertex42/","@vertex42")` |

| D3 | `=COUNTIF(Table1[INCLUDE],"<>")` |

| C4 | `="First "&D3&" Hashtags:"` |

| D4 | `=_xlfn.TEXTJOIN(" ",TRUE,OFFSET(Table1[[#Headers],[HASHTAG]],1,0,$D$3,1))` |

| G7 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B7,"#","")&"/","View")` |

| G8 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B8,"#","")&"/","View")` |

| G9 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B9,"#","")&"/","View")` |

| G10 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B10,"#","")&"/","View")` |

| G11 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B11,"#","")&"/","View")` |

| G12 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B12,"#","")&"/","View")` |

| G13 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B13,"#","")&"/","View")` |

| G14 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B14,"#","")&"/","View")` |

| G15 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B15,"#","")&"/","View")` |

| G16 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B16,"#","")&"/","View")` |

| G17 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B17,"#","")&"/","View")` |

| G18 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B18,"#","")&"/","View")` |

| G19 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B19,"#","")&"/","View")` |

| G20 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B20,"#","")&"/","View")` |

| G21 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B21,"#","")&"/","View")` |

| G22 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B22,"#","")&"/","View")` |

| G23 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B23,"#","")&"/","View")` |

| G24 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B24,"#","")&"/","View")` |

| G25 | `=HYPERLINK("https://www.instagram.com/explore/tags/"&SUBSTITUTE(B25,"#","")&"/","View")` |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | None | n |  | Bold:False |  | None |

| B1 | Hashtag Organizer | s |  | Bold:False |  | None |

| C1 | None | n |  | Bold:True |  | None |

| D1 | None | n |  | Bold:True |  | None |

| E1 | None | n |  | Bold:False |  | None |

| F1 | None | n |  | Bold:False |  | None |

| G1 | =HYPERLINK("https://www.instag | f | =HYPERLINK("https://www.instag | Bold:False |  | right |

| H1 | None | n |  | Bold:False |  | None |

| I1 | None | n |  | Bold:False |  | None |

| A2 | None | n |  | Bold:False |  | None |



---

## Sheet: Settings

- **Dimensions:** 94 rows × 5 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 18.50 |

| B | 6.00 |

| C | 33.70 |

| D | 9.00 |

| E | 34.90 |

| F | 9.00 |



### Formulas

**Total:** 44


| Cell | Formula |
|------|---------|

| C32 | `=_xlfn.CONCAT(B34:B40)` |

| B43 | `=Calendar!P6` |

| A47 | `=TODAY()` |

| A48 | `=DATE(B43,1,1)` |

| A49 | `=(DATE($B$43,1,1)+(3-1)*7)+2-WEEKDAY(DATE($B$43,1,1))+IF(2<WEEKDAY(DATE($B$43,1,1)),7,0)` |

| A50 | `=DATE(B43,5,4)` |

| A51 | `=DATE($B$43,2,2)` |

| A52 | `=(DATE($B$43,2,1)+(3-1)*7)+2-WEEKDAY(DATE($B$43,2,1))+IF(2<WEEKDAY(DATE($B$43,2,1)),7,0)` |

| A53 | `=DATE($B$43,2,14)` |

| A54 | `=DATE($B$43,4,1)` |

| A55 | `=DATE($B$43,4,22)` |

| A56 | `=(DATE($B$43,6,1)+(0-1)*7)+2-WEEKDAY(DATE($B$43,6,1))+IF(2<WEEKDAY(DATE($B$43,6,1)),7,0)` |

| A57 | `=DATE($B$43,5,5)` |

| A58 | `=(DATE($B$43,5,1)+(2-1)*7)+1-WEEKDAY(DATE($B$43,5,1))+IF(1<WEEKDAY(DATE($B$43,5,1)),7,0)` |

| A59 | `=DATE($B$43,6,14)` |

| A60 | `=(DATE($B$43,6,1)+(3-1)*7)+1-WEEKDAY(DATE($B$43,6,1))+IF(1<WEEKDAY(DATE($B$43,6,1)),7,0)` |

| A61 | `=DATE($B$43,7,4)` |

| A62 | `=(DATE($B$43,9,1)+(1-1)*7)+2-WEEKDAY(DATE($B$43,9,1))+IF(2<WEEKDAY(DATE($B$43,9,1)),7,0)` |

| A63 | `=DATE($B$43,9,11)` |

| A64 | `=DATE($B$43,9,17)` |

| A65 | `=DATE($B$43,10,16)` |

| A66 | `=DATE($B$43,10,24)` |

| A67 | `=DATE($B$43,10,31)` |

| A68 | `=DATE($B$43,11,11)` |

| A69 | `=(DATE($B$43,11,1)+(4-1)*7)+5-WEEKDAY(DATE($B$43,11,1))+IF(5<WEEKDAY(DATE($B$43,11,1)),7,0)` |

| A70 | `=DATE($B$43,12,7)` |

| A71 | `=DATE($B$43,12,24)` |

| A72 | `=DATE(B43,12,25)` |

| A73 | `=DATE($B$43,12,26)` |

| A74 | `=DATE($B$43,12,31)` |

| A75 | `=(DATE($B$43,5,1)+(1-1)*7)+2-WEEKDAY(DATE($B$43,5,1))+IF(2<WEEKDAY(DATE($B$43,5,1)),7,0)` |

| A76 | `=(DATE($B$43,8,1)+(1-1)*7)+2-WEEKDAY(DATE($B$43,8,1))+IF(2<WEEKDAY(DATE($B$43,8,1)),7,0)` |

| A77 | `=(DATE($B$43,9,1)+(0-1)*7)+2-WEEKDAY(DATE($B$43,9,1))+IF(2<WEEKDAY(DATE($B$43,9,1)),7,0)` |

| A78 | `=DATE($B$43,5,24)-MOD(WEEKDAY(DATE($B$43,5,24),1)-2,7)` |

| A79 | `=IF(WEEKDAY(DATE($B$43,4+1,0),1)=7,DATE($B$43,4+1,0)-(7-4),(DATE($B$43,4+1,0)-WEEKDAY(DATE($B$43,4+1` |

| A80 | `=IF(AND($B$43>1900,$B$43<2199),ROUND(DATE($B$43,4,1)/7+MOD(19*MOD($B$43,19)-7,30)*0.14,0)*7-6,"")` |

| A81 | `=IF(AND($B$43>=2020,$B$43<=2030),DATEVALUE(INDEX({"2020-01-25";"2021-02-12";"2022-02-01";"2023-01-22` |

| A82 | `=IF(AND($B$43>=2020,$B$43<=2030),DATEVALUE(INDEX({"2020-04-24";"2021-04-13";"2022-04-03";"2023-03-23` |

| A83 | `=IF(AND($B$43>=2020,$B$43<=2030),DATEVALUE(INDEX({"2020-09-19";"2021-09-07";"2022-09-26";"2023-09-16` |

| A84 | `=IF(AND($B$43>=2020,$B$43<=2030),DATEVALUE(INDEX({"2020-12-10";"2021-11-28";"2022-12-18";"2023-12-07` |

| A85 | `=ROUNDDOWN((DATE(2000,3,20)+TIME(7,29,0))+($B$43-2000)*365.24238,0)` |

| A86 | `=ROUNDDOWN((DATE(2000,6,21)+TIME(1,36,0))+($B$43-2000)*365.24163,0)` |

| A87 | `=ROUNDDOWN((DATE(2000,9,22)+TIME(17,17,0))+($B$43-2000)*365.24205,0)` |

| A88 | `=ROUNDDOWN((DATE(2000,12,21)+TIME(13,30,0))+($B$43-2000)*365.242743,0)` |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | SETTINGS | s |  | Bold:False | FF3464AB | left |

| B1 | None | n |  | Bold:True | FF3464AB | left |

| C1 | None | n |  | Bold:True | FF3464AB | left |

| D1 | None | n |  | Bold:True | FF3464AB | left |

| E1 | None | n |  | Bold:True | FF3464AB | left |

| A2 | None | n |  | Bold:False |  | None |

| B2 | None | n |  | Bold:False |  | None |

| C2 | None | n |  | Bold:False |  | None |

| D2 | None | n |  | Bold:False |  | None |

| E2 | Content Calendar Template © 20 | s |  | Bold:False |  | right |



---

## Sheet: Help

- **Dimensions:** 46 rows × 4 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 9.10 |

| B | 63.60 |

| C | 16.70 |



### Merged Cells

- `B2:C2`



### Formulas

**Total:** 5


| Cell | Formula |
|------|---------|

| B38 | `=HYPERLINK("https://www.vertex42.com/blog/excel-tips/how-to-use-conditional-formatting-in-excel.html` |

| B40 | `=HYPERLINK("https://www.vertex42.com/ExcelTips/how-to-make-a-gantt-chart-in-excel.html","► How to Ma` |

| B42 | `=HYPERLINK("https://www.vertex42.com/ExcelTemplates/excel-project-management.html","► More Project M` |

| B44 | `=HYPERLINK("https://www.vertex42.com/ExcelTemplates/business-templates.html","► More Business Templa` |

| B46 | `=HYPERLINK("https://www.vertex42.com/calendars/","► More Calendars")` |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | HELP | s |  | Bold:False | FF3464AB | left |

| B1 | None | n |  | Bold:True | FF3464AB | left |

| C1 | None | n |  | Bold:True | FF3464AB | left |

| D1 | None | n |  | Bold:False |  | None |

| A2 | None | n |  | Bold:False |  | None |

| B2 | https://www.vertex42.com/calen | s |  | Bold:False |  | right |

| C2 | None | n |  | Bold:False |  | None |

| D2 | None | n |  | Bold:False |  | None |

| A3 | None | n |  | Bold:False |  | None |

| B3 | None | n |  | Bold:False |  | None |



---

## Sheet: ©

- **Dimensions:** 19 rows × 3 columns

### Column Widths

| Column | Width |
|--------|-------|

| A | 2.50 |

| B | 62.60 |

| C | 19.50 |

| D | 9.00 |



### Sample Cell Formatting

| Cell | Value | Type | Formula | Font | Fill | Alignment |

|------|-------|------|---------|------|------|-----------|

| A1 | None | n |  | Bold:True | FF3464AB | left |

| B1 | Content Calendar Template | s |  | Bold:False | FF3464AB | left |

| C1 | None | n |  | Bold:False | FF3464AB | None |

| A2 | None | n |  | Bold:False | Values mus | None |

| B2 | None | n |  | Bold:False | Values mus | left |

| C2 | None | n |  | Bold:False | Values mus | None |

| A3 | None | n |  | Bold:False | Values mus | None |

| B3 | By Vertex42.com | s |  | Bold:False | Values mus | None |

| C3 | None | n |  | Bold:False | Values mus | None |

| A4 | None | n |  | Bold:False | Values mus | None |



---

## Excel Functions Used

| Function | Count | Locations |
|----------|-------|-----------|

| `INDEX` | 56 | Calendar!A5, Calendar!A5, Calendar!B5, Calendar!B5, Calendar!C5 ... (51 more) |

| `ROW` | 56 | Calendar!A5, Calendar!A5, Calendar!B5, Calendar!B5, Calendar!C5 ... (51 more) |

| `IFERROR` | 35 | Calendar!B4, Calendar!D4, Calendar!F4, Calendar!H4, Calendar!J4 ... (30 more) |

| `IF` | 29 | Calendar!A4, Calendar!A5, Calendar!B5, Calendar!C5, Calendar!D5 ... (24 more) |

| `SMALL` | 28 | Calendar!A5, Calendar!B5, Calendar!C5, Calendar!D5, Calendar!E5 ... (23 more) |

| `VLOOKUP` | 7 | Calendar!B4, Calendar!D4, Calendar!F4, Calendar!H4, Calendar!J4 ... (2 more) |

| `HYPERLINK` | 7 | Calendar!P4, Hashtags!G1, Hashtags!G7, Hashtags!G8, Hashtags!G9 ... (2 more) |

| `SUBSTITUTE` | 5 | Hashtags!G7, Hashtags!G8, Hashtags!G9, Hashtags!G10, Hashtags!G11 |

| `DATE` | 4 | Calendar!B2, Calendar!A4, Calendar!A4, Calendar!A4 |

| `WEEKDAY` | 2 | Calendar!A4, Calendar!A4 |

| `UPPER` | 1 | Calendar!A1 |

| `TEXT` | 1 | Calendar!A1 |

| `COUNTIF` | 1 | Hashtags!D3 |

| `TEXTJOIN` | 1 | Hashtags!D4 |

| `OFFSET` | 1 | Hashtags!D4 |



## Number Formats Used

| Format | Count |
|--------|-------|

| `General` | 527 |

| `mmmm\ yyyy` | 9 |

| `dddd` | 7 |

| `d` | 7 |

| `#,##0` | 5 |

| `mm-dd-yy` | 2 |



---

## Python Implementation Guide


### 1. Data Structure

```python

# Recommended data structure

from dataclasses import dataclass

from typing import Dict, List, Any


@dataclass

class SheetData:

    name: str

    rows: int

    cols: int

    data: List[List[Any]]

    formulas: Dict[str, str]  # cell_coordinate -> formula

    formats: Dict[str, Dict]  # cell_coordinate -> format_info


@dataclass

class ExcelWorkbook:

    sheets: Dict[str, SheetData]

```


### 2. Function Mapping

| Excel Function | Python Equivalent | Notes |

|----------------|-------------------|-------|

| `COUNTIF` | `TBD - needs implementation` | Used in 1 cells |

| `DATE` | `datetime.date()` | Used in 4 cells |

| `HYPERLINK` | `TBD - needs implementation` | Used in 7 cells |

| `IF` | `if/else or ternary operator` | Used in 29 cells |

| `IFERROR` | `TBD - needs implementation` | Used in 35 cells |

| `INDEX` | `list/dict indexing` | Used in 56 cells |

| `OFFSET` | `TBD - needs implementation` | Used in 1 cells |

| `ROW` | `TBD - needs implementation` | Used in 56 cells |

| `SMALL` | `TBD - needs implementation` | Used in 28 cells |

| `SUBSTITUTE` | `TBD - needs implementation` | Used in 5 cells |

| `TEXT` | `str.format() or f-strings` | Used in 1 cells |

| `TEXTJOIN` | `TBD - needs implementation` | Used in 1 cells |

| `UPPER` | `TBD - needs implementation` | Used in 1 cells |

| `VLOOKUP` | `dict lookup or pandas merge` | Used in 7 cells |

| `WEEKDAY` | `TBD - needs implementation` | Used in 2 cells |


### 3. Implementation Steps


1. **Load Data Structure**

   - Read all sheets

   - Extract cell values and formulas

   - Store formatting information


2. **Implement Functions**

   - Map Excel functions to Python equivalents

   - Handle cell references (e.g., A1, $B$2)

   - Implement calculation engine


3. **Apply Formatting**

   - Recreate cell styles (font, fill, alignment)

   - Apply number formats

   - Handle merged cells


4. **Output Generation**

   - Generate Excel file (openpyxl)

   - Or generate other formats (CSV, JSON, etc.)

---

## Python Implementation Example

### Basic Structure

```python
from dataclasses import dataclass
from typing import Dict, List, Any, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

@dataclass
class CellData:
    """셀 데이터 구조"""
    value: Any
    formula: Optional[str] = None
    font: Optional[Dict] = None
    fill: Optional[Dict] = None
    alignment: Optional[Dict] = None
    number_format: Optional[str] = None

@dataclass
class SheetData:
    """시트 데이터 구조"""
    name: str
    rows: int
    cols: int
    cells: Dict[str, CellData]  # "A1" -> CellData
    column_widths: Dict[str, float]
    merged_cells: List[str]

class ExcelToPythonConverter:
    """Excel을 Python으로 변환하는 클래스"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.sheets: Dict[str, SheetData] = {}
    
    def load_excel(self):
        """Excel 파일 로드"""
        from openpyxl import load_workbook
        wb = load_workbook(self.excel_path, data_only=False)
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_data = SheetData(
                name=sheet_name,
                rows=ws.max_row,
                cols=ws.max_column,
                cells={},
                column_widths={},
                merged_cells=[]
            )
            
            # 셀 데이터 추출
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None or cell.data_type == 'f':
                        cell_data = CellData(
                            value=cell.value if cell.data_type != 'f' else None,
                            formula=cell.value if cell.data_type == 'f' else None
                        )
                        sheet_data.cells[cell.coordinate] = cell_data
            
            # 열 너비
            for col_letter in ws.column_dimensions:
                if ws.column_dimensions[col_letter].width:
                    sheet_data.column_widths[col_letter] = ws.column_dimensions[col_letter].width
            
            # 병합된 셀
            for merged in ws.merged_cells.ranges:
                sheet_data.merged_cells.append(str(merged))
            
            self.sheets[sheet_name] = sheet_data
    
    def evaluate_formula(self, formula: str, sheet_data: SheetData) -> Any:
        """함수 평가 (기본 구현)"""
        # 셀 참조 추출 (예: A1, $B$2)
        cell_refs = re.findall(r'([A-Z]+\$?\d+)', formula)
        
        # 간단한 함수 처리
        if formula.startswith('=IF('):
            # IF 함수 처리
            pass
        elif formula.startswith('=VLOOKUP('):
            # VLOOKUP 함수 처리
            pass
        elif formula.startswith('=INDEX('):
            # INDEX 함수 처리
            pass
        
        return None
    
    def generate_python_code(self) -> str:
        """Python 코드 생성"""
        code = []
        code.append("#!/usr/bin/env python3")
        code.append("# -*- coding: utf-8 -*-")
        code.append('"""')
        code.append("Generated Python implementation of Excel file")
        code.append(f"Source: {self.excel_path}")
        code.append('"""')
        code.append("")
        code.append("from dataclasses import dataclass")
        code.append("from typing import Dict, List, Any")
        code.append("")
        
        # 각 시트에 대한 클래스 생성
        for sheet_name, sheet_data in self.sheets.items():
            class_name = sheet_name.replace(' ', '_').replace('-', '_')
            code.append(f"@dataclass")
            code.append(f"class {class_name}Sheet:")
            code.append(f'    """{sheet_name} 시트 데이터"""')
            code.append(f"    name: str = '{sheet_name}'")
            code.append(f"    rows: int = {sheet_data.rows}")
            code.append(f"    cols: int = {sheet_data.cols}")
            code.append("")
        
        return "\n".join(code)
    
    def export_to_excel(self, output_path: str):
        """Excel 파일로 내보내기 (포맷 유지)"""
        wb = Workbook()
        wb.remove(wb.active)
        
        for sheet_name, sheet_data in self.sheets.items():
            ws = wb.create_sheet(sheet_name)
            
            # 열 너비 설정
            for col_letter, width in sheet_data.column_widths.items():
                ws.column_dimensions[col_letter].width = width
            
            # 셀 데이터 설정
            for cell_coord, cell_data in sheet_data.cells.items():
                cell = ws[cell_coord]
                
                if cell_data.formula:
                    cell.value = cell_data.formula
                else:
                    cell.value = cell_data.value
                
                # 포맷 적용
                if cell_data.font:
                    cell.font = Font(**cell_data.font)
                if cell_data.fill:
                    cell.fill = PatternFill(**cell_data.fill)
                if cell_data.alignment:
                    cell.alignment = Alignment(**cell_data.alignment)
            
            # 병합된 셀
            for merged_range in sheet_data.merged_cells:
                ws.merge_cells(merged_range)
        
        wb.save(output_path)
```

### Usage Example

```python
# Excel 파일 로드 및 분석
converter = ExcelToPythonConverter("content-calendar.xlsx")
converter.load_excel()

# Python 코드 생성
python_code = converter.generate_python_code()
with open("content_calendar_impl.py", "w", encoding="utf-8") as f:
    f.write(python_code)

# Excel로 내보내기 (포맷 유지)
converter.export_to_excel("content-calendar_recreated.xlsx")
```

---

## Detailed Function Implementation Guide

### 1. IF Function

**Excel:** `=IF(condition, value_if_true, value_if_false)`

**Python:**
```python
def excel_if(condition, value_if_true, value_if_false):
    return value_if_true if condition else value_if_false
```

### 2. VLOOKUP Function

**Excel:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

**Python:**
```python
def excel_vlookup(lookup_value, table_array, col_index, exact_match=True):
    for row in table_array:
        if exact_match:
            if row[0] == lookup_value:
                return row[col_index - 1]
        else:
            if row[0] >= lookup_value:
                return row[col_index - 1]
    return None
```

### 3. INDEX Function

**Excel:** `=INDEX(array, row_num, [column_num])`

**Python:**
```python
def excel_index(array, row_num, col_num=None):
    if col_num is None:
        return array[row_num - 1]
    return array[row_num - 1][col_num - 1]
```

### 4. IFERROR Function

**Excel:** `=IFERROR(value, value_if_error)`

**Python:**
```python
def excel_iferror(value, value_if_error):
    try:
        return value
    except:
        return value_if_error
```

### 5. ROW Function

**Excel:** `=ROW([reference])`

**Python:**
```python
def excel_row(cell_reference=None):
    if cell_reference is None:
        return current_row  # 현재 행 번호
    # 셀 참조에서 행 번호 추출
    match = re.match(r'[A-Z]+(\d+)', cell_reference)
    return int(match.group(1)) if match else None
```

---

## Next Steps

1. **함수 구현 우선순위 결정**
   - 가장 많이 사용된 함수부터 구현
   - 현재: INDEX (56회), ROW (56회), IFERROR (35회), IF (29회)

2. **셀 참조 처리**
   - 상대 참조 (A1) vs 절대 참조 ($A$1)
   - 다른 시트 참조 (Sheet1!A1)

3. **의존성 그래프 생성**
   - 함수 간 의존성 분석
   - 계산 순서 결정

4. **테스트 케이스 작성**
   - 각 함수별 단위 테스트
   - 통합 테스트

5. **성능 최적화**
   - 대용량 데이터 처리
   - 캐싱 전략

